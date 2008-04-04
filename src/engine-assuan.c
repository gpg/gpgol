/* engine-assuan.c - Crypto engine using an Assuan server
 *	Copyright (C) 2007, 2008 g10 Code GmbH
 *
 * This file is part of GpgOL.
 *
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1 
 * of the License, or (at your option) any later version.
 *  
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#include <config.h>

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <errno.h>
#include <assert.h>
#define WIN32_LEAN_AND_MEAN 
#define WINVER 0x0500  /* Required for AllowSetForegroundWindow.  */
#include <windows.h>

#include <assuan.h>
#include "common.h"
#include "engine.h"
#include "engine-assuan.h"


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                       SRCNAME, __func__, __LINE__); \
                        } while (0)

/* Debug macros.  */
#define debug_ioworker        (opt.enable_debug & DBG_IOWORKER)
#define debug_ioworker_extra  (opt.enable_debug & DBG_IOWORKER_EXTRA)


/* How many times we will try to connect to a server after we have
   started him.  */
#define FIREUP_RETRIES 10


/* This is the buffer object used for the asynchronous reading of the
   status channel.  */
struct status_buffer_s
{
  int eof;
  int linelen;  /* Used length of LINE. */
  char line[ASSUAN_LINELENGTH];
};
typedef struct status_buffer_s *status_buffer_t;


/* We operate in an asynchronous mode and thus need to run code for
   final cleanup.  Thus all functions need to implement a closure
   function and setup an closure_data_t object.  */ 
struct closure_data_s;
typedef struct closure_data_s *closure_data_t;
struct closure_data_s
{
  void (*closure)(closure_data_t);
  gpg_error_t final_err;         /* Final error code.  */
  engine_filter_t filter;
  assuan_context_t assctx;
  ULONG cmdid;
  assuan_fd_t status_read_fd;
  struct gpgme_data_cbs status_cbs;
  gpgme_data_t status_data;
  status_buffer_t status_buffer; /* Allocated on demand.  */
  int status_ready;
  gpgme_data_t sigdata;   /* Used by verify_closure.  */
  gpg_error_t last_err;
};


/* The object used by our I/O worker thread.  */
struct work_item_s;
typedef struct work_item_s *work_item_t;
struct work_item_s
{
  work_item_t next;
  int used;          /* If not set this object may be reused.  */
  int waiting;       /* Helper for async_worker_thread.  */

  const char *name;  /* Description used for debugging.  */
  ULONG cmdid;       /* Used to group work items of one command.  */
  closure_data_t cld;/* NULL or the closure.  */
  int wait_on_success; /* This work item needs to be ready before
                          invoking a closure for this command.  */
  gpgme_data_t data; /* The data object we write to or read from.  */
  int writing;       /* If true we are going to write to HD.  */
  HANDLE hd;         /* The handle we read from or write to.  */
  int inactive;      /* If set, the handle is not yet active.  */
  int io_pending;    /* I/O is still pending.  The value is the number
                        of bytes to be written or the size of the
                        buffer given to ReadFile. */
  int got_ready;     /* Operation finished.  */
  int delayed_ready; /* Ready but delayed due to a missing prerequesite.  */
  int got_error;     /* An error as been encountered.  */
  int aborting;      /* Set to true after a CancelIO has been issued.  */
  void (*finalize)(work_item_t); /* Function called immediately before
                                    the item is removed from the
                                    queue.  */
  OVERLAPPED ov;     /* The overlapped info structure.  */
  char buffer[1024]; /* The buffer used by ReadFile or WriteFile.  */

  ULONG switch_counter; /* Used by switch_threads.  */
};


/* A helper context used to convey information from op_assuan_encrypt
   to op_assuan_encrypt_bottom.  */
struct engine_assuan_encstate_s
{
  engine_filter_t filter;
  const char *protocol_name;
  HANDLE inpipe[2];
  HANDLE outpipe[2];
  closure_data_t cld;
  assuan_context_t ctx;
  ULONG cmdid;
};


/* The queue of all outstandig I/O operations.  Protected by the
   work_queue_lock.  */
static work_item_t work_queue;

/* The big lock used to protect the work queue.  */
static CRITICAL_SECTION work_queue_lock;

/* An auto-reset event which will be signaled to get the
   async_worker_thread out of its WFMO and to inspect the work
   queue.  */
static HANDLE work_queue_event;


/*-- prototypes --*/
static DWORD WINAPI async_worker_thread (void *dummy);




/* Return the next command id.  Command Ids are used to group
   resources of one command. */
static ULONG
create_command_id (void)
{
  static ULONG command_id;
  ULONG cmdid;

  while (!(cmdid = InterlockedIncrement (&command_id)))
    ;
  return cmdid;
}


static void
close_pipe (HANDLE apipe[2])
{
  int i;

  for (i=0; i < 2; i++)
    if (apipe[i] != INVALID_HANDLE_VALUE)
      {
        CloseHandle (apipe[i]);
        apipe[i] = INVALID_HANDLE_VALUE;
      }
}


/* Duplicate HANDLE into the server's process and close HANDLE.  Note
   that HANDLE is closed even if the function fails.  Returns the
   duplicated handle on success or INVALID_HANDLE_VALUE on error.  */
static HANDLE
dup_to_server (HANDLE handle, pid_t serverpid)
{
  HANDLE prochandle, newhandle;

  prochandle = OpenProcess (PROCESS_DUP_HANDLE, FALSE, serverpid);
  if (!prochandle)
    {
      log_error_w32 (-1, "%s:%s: OpenProcess(%lu) failed", 
                     SRCNAME, __func__, (unsigned long)serverpid);
      CloseHandle (handle);
      return INVALID_HANDLE_VALUE;
    }

  if (!DuplicateHandle (GetCurrentProcess(), handle,
                        prochandle, &newhandle, 0,
                        TRUE, DUPLICATE_SAME_ACCESS ))
    {
      log_error_w32 (-1, "%s:%s: DuplicateHandle to pid %lu failed", 
                     SRCNAME, __func__, (unsigned long)serverpid);
      CloseHandle (prochandle);
      CloseHandle (handle);
      return INVALID_HANDLE_VALUE;
    }
  CloseHandle (prochandle);
  CloseHandle (handle);
  return newhandle;
}


/* Create pipe with one end being inheritable and prepared for
   overlapped I/O.

     FILEDES[0] := read handle. 
     FILEDES[1] := write handle. 

   SERVERPID is the PID of the server.  FOR_WRITE is seen out of our
   perspective; if it is set, the read handle is created in the server
   process and the write handle is overlapped.  If it is not set the
   write handle is created in the server process and the read handle
   is overlapped.
*/
static gpg_error_t
create_io_pipe (HANDLE filedes[2], pid_t serverpid, int for_write)
{
  static ULONG pipenumber;
  ULONG pipeno;
  char pipename[100];
  HANDLE r, w;
  SECURITY_ATTRIBUTES sec_attr;

  memset (&sec_attr, 0, sizeof sec_attr );
  sec_attr.nLength = sizeof sec_attr;

  /* CreatePipe is in reality implemented using a Named Pipe.  We do
     it the same but use a name which is in our name space.  We allow
     only one instance, use the standard timeout of 120 seconds and
     buffers of 4k. */
  pipeno = InterlockedIncrement (&pipenumber);
  snprintf (pipename, sizeof pipename, "\\\\.\\pipe\\GpgOL_anon.%08lx.%08lx",
            (unsigned long)GetCurrentProcessId(), pipeno);
  sec_attr.bInheritHandle = /*for_write? TRUE :*/FALSE;
  r = CreateNamedPipe (pipename, (PIPE_ACCESS_INBOUND
                                  | (for_write? 0:FILE_FLAG_OVERLAPPED)),
                       PIPE_TYPE_BYTE | PIPE_WAIT,
                       1, 4096, 4096, 120000, &sec_attr);
  if (r == INVALID_HANDLE_VALUE)
    {
      log_error_w32 (-1, "%s:%s: CreateNamedPipe failed for `%s'",
                     SRCNAME, __func__, pipename);
      return gpg_error (GPG_ERR_GENERAL);
    }
  if (for_write)
    {
      r = dup_to_server (r, serverpid);
      if (r == INVALID_HANDLE_VALUE)
        {
          log_error_w32 (-1, "%s:%s: dup_for_server(r) failed for `%s'",
                         SRCNAME, __func__, pipename);
          return gpg_error (GPG_ERR_GENERAL);
        }
    }

  /* Now open the other side of the named pipe.  Because we have not
     called ConnectNamedPipe another process should not be able to
     open the pipe in the meantime.  This is an educated guess by
     looking at REACTOS and WINE - they implement an anonymous pipe
     this way.  */
  sec_attr.bInheritHandle = /*for_write?*/ FALSE /*: TRUE*/;
  w = CreateFile (pipename, GENERIC_WRITE, 0, &sec_attr,
                  OPEN_EXISTING, (FILE_ATTRIBUTE_NORMAL
                                  | (for_write? FILE_FLAG_OVERLAPPED:0)),
                  NULL);
  if (w == INVALID_HANDLE_VALUE)
    {
      log_error_w32 (-1, "%s:%s: CreateFile failed for `%s'",
                     SRCNAME, __func__, pipename);
      CloseHandle (r);
      return gpg_error (GPG_ERR_GENERAL);
    }
  if (!for_write)
    {
      w = dup_to_server (w, serverpid);
      if (w == INVALID_HANDLE_VALUE)
        {
          log_error_w32 (-1, "%s:%s: dup_for_server(w) failed for `%s'",
                         SRCNAME, __func__, pipename);
          CloseHandle (r);
          return gpg_error (GPG_ERR_GENERAL);
        }
    }

  filedes[0] = r;
  filedes[1] = w;
  if (debug_ioworker)
    log_debug ("%s:%s: new pipe created: r=%p%s w=%p%s",  SRCNAME, __func__,
               r, for_write? " (server)":"",
               w, !for_write?" (server)":"");
  return 0;
}


/* Return the socket name of the UI Server.  */
static const char *
get_socket_name (void)
{
  static char *name;

  if (!name)
    {
      const char *dir = default_homedir ();
      name = xmalloc (strlen (dir) + 11 + 1);
      strcpy (stpcpy (name, dir), "\\S.uiserver");
    }

  return name;
}


/* Return the name of the default UI server.  This name is used to
   auto start an UI server if an initial connect failed.  */
static char *
get_uiserver_name (void)
{
  char *name = NULL;
  char *dir, *uiserver, *p;
  int extra_arglen = 0;

  dir = read_w32_registry_string ("HKEY_LOCAL_MACHINE", GNUPG_REGKEY,
                                  "Install Directory");
  if (dir)
    {
      uiserver = read_w32_registry_string (NULL, GNUPG_REGKEY, 
                                           "UI Server");
      if (!uiserver)
        {
          uiserver = xstrdup ("kleopatra.exe");
          extra_arglen = 9; /* Space required for " --daemon".  */
        }

      name = xmalloc (strlen (dir) + strlen (uiserver) + extra_arglen + 2);
      strcpy (stpcpy (stpcpy (name, dir), "\\"), uiserver);
      for (p = name; *p; p++)
        if (*p == '/')
          *p = '\\';
      xfree (uiserver);
      if (extra_arglen && access (name, F_OK))
        {
          /* Kleopatra iis not nstalled: Try GPA instead but if it is
             also not available still return the Kleopatra
             filename.  */
          const char gpaserver[] = "gpa.exe";
          char *name2;
          
          name2 = xmalloc (strlen (dir) + strlen (gpaserver) + extra_arglen+2);
          strcpy (stpcpy (stpcpy (name2, dir), "\\"), gpaserver);
          for (p = name2; *p; p++)
            if (*p == '/')
              *p = '\\';
          if (access (name2, F_OK ))
            xfree (name2);
          else
            {
              xfree (name);
              name = name2;
            }
        }
      xfree (dir);

      /* Append the arg for Kleopatra.  */
      if (name && extra_arglen)
        strcat (name, " --daemon");
    }
  
  return name;
}



static gpg_error_t
send_one_option (assuan_context_t ctx, const char *name, const char *value)
{
  gpg_error_t err;
  char buffer[1024];

  if (!value || !*value)
    err = 0;  /* Avoid sending empty strings.  */
  else 
    {
      snprintf (buffer, sizeof buffer, "OPTION %s=%s", name, value);
      err = assuan_transact (ctx, buffer, NULL, NULL, NULL, NULL, NULL, NULL);
    }

  return err;
}


static int
getinfo_pid_cb (void *opaque, const void *buffer, size_t length)
{
  pid_t *pid = opaque;
  char pidbuf[50];

  /* There is only the pid in the server's response.  */
  if (length >= sizeof pidbuf)
    length = sizeof pidbuf -1;
  if (length)
    {
      strncpy (pidbuf, buffer, length);
      pidbuf[length] = 0;
      *pid = (pid_t)strtoul (pidbuf, NULL, 10);
    }
  return 0;
}


/* Send options to the UI server and return the server's PID.  */
static gpg_error_t
send_options (assuan_context_t ctx, void *hwnd, pid_t *r_pid)
{
  gpg_error_t err = 0;
  char numbuf[50];

  *r_pid = (pid_t)(-1);
  err = assuan_transact (ctx, "GETINFO pid", getinfo_pid_cb, r_pid,
                         NULL, NULL, NULL, NULL);
  if (!err && *r_pid == (pid_t)(-1))
    {
      log_debug ("%s:%s: server did not return a PID", SRCNAME, __func__);
      err = gpg_error (GPG_ERR_ASSUAN_SERVER_FAULT);
    }

  if (*r_pid != (pid_t)(-1) && !AllowSetForegroundWindow (*r_pid))
    log_error_w32 (-1, "AllowSetForegroundWindow(%u) failed", *r_pid);

  if (!err && hwnd)
    {
      snprintf (numbuf, sizeof numbuf, "%lx", (unsigned long)hwnd);
      err = send_one_option (ctx, "window-id", numbuf);
    }

  return err;
}


/* Connect to the UI server and setup the connection.  */
static gpg_error_t
connect_uiserver (assuan_context_t *r_ctx, pid_t *r_pid, ULONG *r_cmdid,
                  void *hwnd)
{
  static ULONG retry_counter;
  ULONG retry_count;
  gpg_error_t err;
  assuan_context_t ctx;

  *r_ctx = NULL;
  *r_pid = (pid_t)(-1);
  *r_cmdid = 0;
 retry:
  err = assuan_socket_connect (&ctx, get_socket_name (), -1);
  if (err)
    {
      /* Let only one thread start an UI server but all allow threads
         to check for a connection.  Note that this is not really
         correct as the maximum waiting time decreases with the number
         of threads.  However, it is unlikely that we have more than 2
         or 3 threads here - if at all more than one.  */
      retry_count = InterlockedExchangeAdd (&retry_counter, 1);
      if (retry_count < FIREUP_RETRIES)
        {
          if (!retry_count)
            {
              char *uiserver = get_uiserver_name ();
              if (!uiserver)
                {
                  log_error ("%s:%s: UI server not installed",
                             SRCNAME, __func__);
                  InterlockedExchange (&retry_counter, FIREUP_RETRIES);
                  retry_count = FIREUP_RETRIES;
                }
              else
                {
                  log_debug ("%s:%s: UI server not running, starting `%s'",
                             SRCNAME, __func__, uiserver);
                  if (gpgol_spawn_detached (uiserver))
                    {
                      /* Error; try again to connect in case the
                         server has been started in the meantime.
                         Make sure that we don't get here a second
                         time.  */
                      InterlockedExchange (&retry_counter, FIREUP_RETRIES);
                    }
                  xfree (uiserver);
                }
            }
          if (retry_count < FIREUP_RETRIES)
            {
              log_debug ("%s:%s: waiting for UI server to come up",
                         SRCNAME, __func__);
              Sleep (1000);
              goto retry;
            }
        }
      else
        {
          /* Avoid a retry counter overflow by limiting to the limit.  */
          InterlockedExchange (&retry_counter, FIREUP_RETRIES);
        }

      log_error ("%s:%s: error connecting `%s': %s\n", SRCNAME, __func__,
                 get_socket_name (), gpg_strerror (err));
    }
  else if ((err = send_options (ctx, hwnd, r_pid)))
    {
      assuan_disconnect (ctx);
    }
  else
    {
      *r_cmdid = create_command_id ();
      *r_ctx = ctx;
    }
  return err;
}





static void
cleanup (void)
{
  /* Fixme: We should stop the worker thread.  */
}


/* Cleanup static resources. */
void
op_assuan_deinit (void)
{
  cleanup ();
}


/* Initialize this system. */
int
op_assuan_init (void)
{
  static int init_done;
  gpgme_error_t err;
  assuan_context_t ctx;
  pid_t pid;
  ULONG cmdid;

  if (init_done)
    return 0;
  
  /* Run a test connection to see whether the UI server is available.  */
  err = connect_uiserver (&ctx, &pid, &cmdid, NULL);
  if (!err)
    {
      err = assuan_transact (ctx, "NOP", NULL, NULL, NULL, NULL, NULL, NULL);
      assuan_disconnect (ctx);
    }
  if (err)
    return err;
  
  /* Fire up the pipe worker thread. */
  {
    HANDLE th;
    DWORD  mytid, tid;

    InitializeCriticalSection (&work_queue_lock);
    work_queue_event = CreateEvent (NULL, FALSE, FALSE, NULL);
    if (!work_queue_event)
      {
        log_error_w32 (-1, "%s:%s: CreateEvent failed", SRCNAME, __func__);
        return gpg_error (GPG_ERR_GENERAL);
      }
    mytid = GetCurrentThreadId ();
    th = CreateThread (NULL, 256*1024, async_worker_thread, (void*)mytid,
                       0, &tid);
    if (th == INVALID_HANDLE_VALUE)
      log_error ("failed to launch the async_worker_thread");
    else
      CloseHandle (th);
  }

  init_done = 1; 
  return 0;
}

#if 0 /* Not used. */
/* Dummy window procedure.  */
static LRESULT CALLBACK 
attach_thread_input_wndw_proc (HWND hwnd, UINT msg, 
                               WPARAM wparam, LPARAM lparam)
{		
  return DefWindowProc (hwnd, msg, wparam, lparam);
}


/* Our helper thread needs to attach its input events to the main
   message queue.  To do this we need to create a the message queue
   first by creating an hidden window within this thread.  */
static void
attach_thread_input (DWORD other_tid)
{
  WNDCLASS wndwclass = {0, attach_thread_input_wndw_proc, 0, 0, glob_hinst,
                        0, 0, 0, 0, "gpgol-assuan-engine"};
  HWND hwnd;

  /* First create a window to make sure that a message queue exists
     for this thread.  */
  if (!RegisterClass (&wndwclass))
    {
      log_error_w32 (-1, "%s:%s: error registering window class",
                     SRCNAME, __func__);
      return;
    }
  hwnd = CreateWindow ("gpgol-assuan-engine", "gpgol-assuan-engine",
                       0, 0, 0, 0, 0, NULL, NULL, glob_hinst, NULL);
  if (!hwnd)
    {
      log_error_w32 (-1, "%s:%s: error creating main window",
                     SRCNAME, __func__);
      return;
    }

  /* Now attach it to the main thread.  */
  if (!AttachThreadInput (GetCurrentThreadId (), other_tid,  TRUE))
    log_error_w32 (-1, "%s:%s: AttachThreadInput failed",
                   SRCNAME, __func__);
  log_debug ("%s:%s: attached thread %lu to %lu", SRCNAME, __func__,
             GetCurrentThreadId (), other_tid);
}
#endif /* not used.  */


/* This is a wraper around SwitchToThread, a syscall we unfortunately
   need due to the lack of an sophisticated event system.  The wrapper
   calls SwitchToThread but after a couple of immediate folliwing
   switches, it introduces a short delays.  */
static void
switch_threads (work_item_t item)
{
  ULONG count;

  count = InterlockedExchangeAdd (&item->switch_counter, 1);
  if (count > 5)
    {
      /* Tried too often without success.  Use Sleep until
         clear_switch_threads has been called.  */
      InterlockedExchange (&item->switch_counter, 5);
      SleepEx (60, TRUE); 
    }
  else if (!SwitchToThread ())
    {
      /* No runable other thread: Fall asleep. */
      SleepEx (8, TRUE);
    }
}

/* Call this fucntion if some action has been done.  */
static void
clear_switch_threads (work_item_t item)
{
  InterlockedExchange (&item->switch_counter, 0);
}




/* Helper to write to the callback.  */
static void
write_to_callback (work_item_t item, DWORD nbytes)
{
  int nwritten;

  if (!nbytes)
    {
      /* (With overlapped, EOF is not indicated by NBYTES==0.)  */
      log_error ("%s:%s: [%s:%p] short read (0 bytes)",
                 SRCNAME, __func__, item->name, item->hd);
    }
  else
    {
      assert (nbytes > 0);
      nwritten = gpgme_data_write (item->data, item->buffer, nbytes);
      if (nwritten < 0)
        {
          log_error ("%s:%s: [%s:%p] error writing to callback: %s",
                     SRCNAME, __func__, item->name, item->hd,strerror (errno));
          item->got_error = 1;
        }
      else if (nwritten < nbytes)
        {
          log_error ("%s:%s: [%s:%p] short write to callback (%d of %lu)",
                     SRCNAME, __func__, item->name, item->hd, nwritten,nbytes);
          item->got_error = 1;
        }
      else
        {
          if (debug_ioworker)
            log_debug ("%s:%s: [%s:%p] wrote %d bytes to callback",
                       SRCNAME, __func__, item->name, item->hd, nwritten);
        }
    }
}

/* Helper for async_worker_thread.  Returns true if the item's handle
   needs to be put on the wait list.  This is called with the worker
   mutex hold. */
static int
worker_start_read (work_item_t item)
{
  DWORD nbytes;
  int retval = 0;

  /* Read from the handle and write to the callback.  The gpgme
     callback is expected to never block.  */
  if (ReadFile (item->hd, item->buffer, sizeof item->buffer,
                &nbytes, &item->ov) )
    {
      write_to_callback (item, nbytes);
      retval = 1;
    }
  else 
    {
      int syserr = GetLastError ();

      if (syserr == ERROR_IO_PENDING)
        {
          if (debug_ioworker)
            log_debug ("%s:%s: [%s:%p] io(read) pending",
                       SRCNAME, __func__, item->name, item->hd);
          item->io_pending = sizeof item->buffer;
          retval = 1;
        }
      else if (syserr == ERROR_HANDLE_EOF || syserr == ERROR_BROKEN_PIPE)
        {
          if (debug_ioworker)
            log_debug ("%s:%s: [%s:%p] EOF%s seen",
                       SRCNAME, __func__, item->name, item->hd,
                       syserr == ERROR_BROKEN_PIPE? " (broken pipe)":"");
          item->got_ready = 1;
        }
      else
        {
          log_error_w32 (syserr, "%s:%s: [%s:%p] read error",
                         SRCNAME, __func__, item->name, item->hd);
          item->got_error = 1;
        }
    }

  return retval;
}

/* Result checking helper for async_worker_thread.  This is called with
   the worker mutex hold.  */
static void
worker_check_read (work_item_t item, DWORD nbytes)
{
  write_to_callback (item, nbytes);
}



/* Helper for async_worker_thread.  Returns true if the item's handle
   needs to be put on the wait list.  This is called with the worker
   mutex hold.  */
static int
worker_start_write (work_item_t item)
{
  int nread;
  DWORD nbytes;
  int retval = 0;

  assert (!item->io_pending);

  /* Read from the callback and the write to the handle.  The gpgme
     callback is expected to never block.  */
  nread = gpgme_data_read (item->data, item->buffer, sizeof item->buffer);
  if (nread < 0)
    {
      if (errno == EAGAIN)
        {
          /* EAGAIN from the callback.  That means that data is
             currently not available.  */
          if (debug_ioworker_extra)
            log_debug ("%s:%s: [%s:%p] EAGAIN received from callback",
                       SRCNAME, __func__, item->name, item->hd);
          switch_threads (item);
          retval = 1;
        }
      else
        {
          clear_switch_threads (item);
          log_error ("%s:%s: [%s:%p] error reading from callback: %s",
                     SRCNAME, __func__, item->name, item->hd,strerror (errno));
          item->got_error = 1;
        }
    }
  else if (!nread)
    {
      clear_switch_threads (item);
      if (debug_ioworker)
        log_debug ("%s:%s: [%s:%p] EOF received from callback",
                   SRCNAME, __func__, item->name, item->hd);
      item->got_ready = 1;
    }
  else 
    {                  
      clear_switch_threads (item);
      if (WriteFile (item->hd, item->buffer, nread, &nbytes, &item->ov))
        {
          if (nbytes < nread)
            {
              log_error ("%s:%s: [%s:%p] short write (%lu of %d)", 
                         SRCNAME, __func__, item->name,item->hd,nbytes, nread);
              item->got_error = 1;
            }
          else
            {
              if (debug_ioworker)
                log_debug ("%s:%s: [%s:%p] wrote %lu bytes", 
                           SRCNAME, __func__, item->name, item->hd, nbytes);
            }
          retval = 1; /* Keep on waiting for space in the pipe.  */
        }
      else 
        {
          int syserr = GetLastError ();

          if (syserr == ERROR_IO_PENDING)
            {
              /* This is the common case.  Start the async I/O.  */
              if (debug_ioworker)
                log_debug ("%s:%s: [%s:%p] io(write) pending (%d bytes)",
                           SRCNAME, __func__, item->name, item->hd, nread);
              item->io_pending = nread;
              retval = 1; /* Need to wait for the I/O to complete.  */
            }
          else
            {
              log_error_w32 (syserr, "%s:%s: [%s:%p] write error",
                             SRCNAME, __func__, item->name, item->hd);
              item->got_error = 1;
            }
        }
    }

  return retval;
}


/* Result checking helper for async_worker_thread.  This is called with
   the worker mutex hold.  */
static void
worker_check_write (work_item_t item, DWORD nbytes)
{
  if (nbytes < item->io_pending)
    {
      log_error ("%s:%s: [%s:%p] short write (%lu of %d)",
                 SRCNAME,__func__, item->name, item->hd, nbytes,
                 item->io_pending);
      item->got_error = 1;
    }
  else
    {
      if (debug_ioworker)
        log_debug ("%s:%s: [%s:%p] write finished (%lu bytes)", 
                   SRCNAME, __func__, item->name, item->hd, nbytes);
    }
}



/* The worker thread which feeds the pipes.  */
static DWORD WINAPI
async_worker_thread (void *dummy)
{
  work_item_t item;
  int n;
  DWORD nbytes;
  HANDLE hdarray[MAXIMUM_WAIT_OBJECTS];
  int count, addit, any_ready, hdarraylen;
  
/*   attach_thread_input ( (DWORD)dummy ); */
  (void)dummy;

  for (;;)
    {
      /* 
         Step 1: Walk our queue and fire up async I/O requests.  
       */
      if (debug_ioworker_extra)
        log_debug ("%s:%s: step 1 - scanning work queue", SRCNAME, __func__);
      EnterCriticalSection (&work_queue_lock);
      
      /* We always need to wait on the the work queue event.  */
      hdarraylen = 0;
      hdarray[hdarraylen++] = work_queue_event;

      count = 0;
      any_ready = 0;
      for (item = work_queue; item; item = item->next)
        {
          item->waiting = 0;
          if (!item->used)
            continue;
          assert (item->hd != INVALID_HANDLE_VALUE);
          count++;
          if (item->got_error)
            {
              if (!item->delayed_ready)
                any_ready = 1;
              continue; 
            }
          assert (item->data);
          if (hdarraylen == DIM (hdarray))
            {
              if (debug_ioworker)
                log_debug ("%s:%s: [%s:%p] wait array full - ignored for now",
                           SRCNAME, __func__, item->name, item->hd);
              continue;
            }

          /* Decide whether we need to wait for this item.  This is
             the case if the previous WriteFile or ReadFile indicated
             that I/O is pending or if the worker_start_foo function
             indicated that we should wait.  Put handles we want to
             wait upon into HDARRAY. */ 
          if (item->inactive)
            addit = 0;
          else if (item->io_pending)
            addit = 1;
          else if (item->writing)
            addit = worker_start_write (item);
          else 
            addit = worker_start_read (item);

          if (addit)
            {
              hdarray[hdarraylen++] = item->hd;
              item->waiting = 1; /* Only required for debugging.  */
            }

          /* Set a flag if this work item is ready or got an error.  */
          if (!item->delayed_ready && (item->got_error || item->got_ready))
            any_ready = 1;
        }
      LeaveCriticalSection (&work_queue_lock);

      /*
         Step 2: Wait for events or handle activitity.
       */
      if (any_ready)
        {
          /* There is at least one work item which is ready or got an
             error.  Skip the wait step so that we process it
             immediately.  */
          if (debug_ioworker_extra)
            log_debug ("%s:%s: step 2 - %d items in queue; skipping wait", 
                       SRCNAME, __func__, count);
        }
      else
        {
          if (debug_ioworker_extra)
            {
              log_debug ("%s:%s: step 2 - "
                         "%d items in queue; waiting for %d items:",
                         SRCNAME, __func__, count, hdarraylen-1);
              for (item = work_queue; item; item = item->next)
                {
                  if (item->waiting)
                    log_debug ("%s:%s: [%s:%p]",
                               SRCNAME, __func__, item->name, item->hd);
                }
            }
          /* [Currently not used]
             First process any window messages of this thread.  Do
             this before wating so that the message queue is cleared
             before waiting and we don't get stucked due to messages
             not removed.  We need to process the message queue also
             after the wait because we will only get to here if there
             is actual ui-server work to be done but some messages
             might still be in the queue.  */
/*           { */
/*             MSG msg; */

/*             while (PeekMessage (&msg, NULL, 0, 0, PM_REMOVE)) */
/*               { */
/*                 TranslateMessage (&msg); */
/*                 DispatchMessage (&msg); */
/*               } */
/*           } */

          n = WaitForMultipleObjects (hdarraylen, hdarray, FALSE, INFINITE);
/*           n = MsgWaitForMultipleObjects (hdarraylen, hdarray, FALSE, */
/*                                          INFINITE, QS_ALLEVENTS); */
          if (n == WAIT_FAILED)
            {
              /* The WFMO failed.  This is an error; to help debugging
                 we now print information about all the handles we
                 wanted to wait upon.  */
              int i;
              DWORD hdinfo;
                  
              log_error_w32 (-1, "%s:%s: WFMO failed", SRCNAME, __func__);
              for (i=0; i < hdarraylen; i++)
                {
                  hdinfo = 0;
                  if (!GetHandleInformation (hdarray[i], &hdinfo))
                    log_debug_w32 (-1, "%s:%s: WFMO GetHandleInfo(%p) failed", 
                                   SRCNAME, __func__, hdarray[i]);
                  else
                    log_debug ("%s:%s: WFMO GetHandleInfo(%p)=0x%lu", 
                               SRCNAME, __func__, hdarray[i],
                               (unsigned long)hdinfo);
                }
              Sleep (1000);
            }
          else if (n >= 0 && n < hdarraylen)
            {
              if (debug_ioworker_extra)
                log_debug ("%s:%s: WFMO succeeded (res=%d, hd=%p)",
                           SRCNAME, __func__, n, hdarray[n]);
            }
          else if (n == hdarraylen)
            {
              if (debug_ioworker_extra)
                log_debug ("%s:%s: WFMO succeeded (res=%d, msgevent)",
                           SRCNAME, __func__, n);
            }
          else
            {
              log_error ("%s:%s: WFMO returned: %d", SRCNAME, __func__, n);
              Sleep (1000);
            }

          /* [Currently not used] 
             Try to process the message queue.  */
/*           { */
/*             MSG msg; */
            
/*             while (PeekMessage (&msg, NULL, 0, 0, PM_REMOVE)) */
/*               { */
/*                 TranslateMessage (&msg); */
/*                 DispatchMessage (&msg); */
/*               } */
/*           } */
        }

      /*
         Step 3: Handle I/O completion status.
       */
      EnterCriticalSection (&work_queue_lock);
      if (debug_ioworker_extra)
        log_debug ("%s:%s: step 3 - checking completion states", 
                   SRCNAME, __func__);
      for (item = work_queue; item; item = item->next)
        {
          if (!item->io_pending || item->inactive)
            {
              /* No I/O is pending for that item, thus there is no
                 need to check a completion status.  */
            }
          else if (GetOverlappedResult (item->hd, &item->ov, &nbytes, FALSE))
            {
              /* An overlapped I/O result is available.  Check that
                 the returned number of bytes are plausible and clear
                 the I/O pending flag of this work item.  For a a read
                 item worker_check_read forwards the received data to
                 the caller.  */
              if (item->writing)
                worker_check_write (item, nbytes);
              else
                worker_check_read (item, nbytes);
              item->io_pending = 0;
            }
          else 
            {
              /* Some kind of error occured: Set appropriate
                 flags.  */
              int syserr = GetLastError ();
              if (syserr == ERROR_IO_INCOMPLETE)
                {
                  /* This is a common case, the I/O has not yet
                     completed for this work item.  No need for any
                     action. */
                }
              else if (!item->writing && (syserr == ERROR_HANDLE_EOF
                                          || syserr == ERROR_BROKEN_PIPE) )
                {
                  /* Got EOF.  */
                  if (debug_ioworker)
                    log_debug ("%s:%s: [%s:%p] EOF%s received",
                               SRCNAME, __func__, item->name, item->hd,
                               syserr==ERROR_BROKEN_PIPE?" (broken pipe)":"");
                  item->io_pending = 0;
                  item->got_ready = 1;
                }
              else
                {
                  /* Something went wrong.  We better cancel the I/O. */
                  log_error_w32 (syserr,
                                 "%s:%s: [%s:%p] GetOverlappedResult failed",
                                 SRCNAME, __func__, item->name, item->hd);
                  item->io_pending = 0;
                  item->got_error = 1;
                  if (!item->aborting)
                    {
                      item->aborting = 1;
                      if (!CancelIo (item->hd))
                        log_error_w32 (-1, "%s:%s: [%s:%p] CancelIo failed",
                                       SRCNAME,__func__, item->name, item->hd);
                    }
                  else 
                    item->got_ready = 1;
                }
            }
        }
      LeaveCriticalSection (&work_queue_lock);

      /* 
         Step 4: Act on the flags set in step 3.
       */

      /* Give the system a chance to process cancel I/O etc. */
      SwitchToThread ();

      EnterCriticalSection (&work_queue_lock);
      if (debug_ioworker_extra)
        log_debug ("%s:%s: step 4 - cleaning up work queue",
                   SRCNAME, __func__); 
      for (item = work_queue; item; item = item->next)
        {
          if (item->used && (item->got_ready || item->got_error))
            {
              /* This is a work item either flagged as ready or as in
                 error state. */
              
              if (item->cld)
                {
                  /* Perform the closure.  */
                  work_item_t itm2;

                  /* If the Assuan state did not return an ERR but the
                     I/O machinery set this work item int the error
                     state, we set the final error to reflect this.  */
                  if (!item->cld->final_err && item->got_error)
                    item->cld->final_err = gpg_error (GPG_ERR_EIO);

                  /* Check whether there are other work items in this
                     group we need to wait for before invoking the
                     closure. */
                  for (itm2=work_queue; itm2; itm2 = itm2->next)
                    if (itm2->used && itm2 != item 
                        && itm2->cmdid == item->cmdid
                        && itm2->wait_on_success
                        && !(itm2->got_ready || itm2->got_error))
                      break;
                  if (itm2)
                    {
                      /* We need to delay the closure until all work
                         items we depend on are ready.  */
                      if (debug_ioworker)
                        log_debug ("%s:%s: [%s:%p] delaying closure "
                                   "due to [%s:%p]", SRCNAME, __func__,
                                   item->name, item->hd, 
                                   itm2->name, itm2->hd);
                      item->delayed_ready = 1;
                      if (item->cld->final_err)
                        {
                          /* However, if we received an error we better
                             do not assume that the server has properly
                             closed all I/O channels.  Send a cancel
                             to the work item we are waiting for. */
                          if (!itm2->aborting)
                            {
                              if (debug_ioworker)
                                log_debug ("%s:%s: [%s:%p] calling CancelIO",
                                           SRCNAME, __func__,
                                           itm2->name, itm2->hd);
                              itm2->aborting = 1;
                              if (!CancelIo (itm2->hd))
                                log_error_w32 (-1, "%s:%s: [%s:%p] "
                                               "CancelIo failed",
                                               SRCNAME,__func__, 
                                               itm2->name, itm2->hd);
                            }
                          else
                            {
                              /* Ooops second time here: Clear the
                                 wait flag so that we won't get to
                                 here anymore.  */
                              if (debug_ioworker)
                                log_debug ("%s:%s: [%s:%p] clearing "
                                           "wait on success flag",
                                           SRCNAME, __func__,
                                           itm2->name, itm2->hd);
                              itm2->wait_on_success = 0;
                            }
                        }
                      goto step4_cont;
                    }
                  
                  item->delayed_ready = 0;
                  if (debug_ioworker)
                    log_debug ("%s:%s: [%s:%p] invoking closure",
                               SRCNAME,__func__, item->name, item->hd);
                  
                  item->cld->closure (item->cld);
                  xfree (item->cld);
                  item->cld = NULL;
                } /* End closure processing.  */

              item->got_ready = 0;
              item->finalize (item);
              item->used = 0;
            step4_cont:
              ;
            }
        }

      LeaveCriticalSection (&work_queue_lock);
    }
}


void
engine_assuan_cancel (void *cancel_data)
{
  /* FIXME */
}




/* Standard finalize handler.  Called right before the item is removed
   from the queue.  Called while the work_queue_lock is hold.  */
static void
finalize_handler (work_item_t item)
{
  if (debug_ioworker)
    log_debug ("%s:%s: [%s:%p] closing handle", 
               SRCNAME, __func__, item->name, item->hd);
  CloseHandle (item->hd);
  item->hd = INVALID_HANDLE_VALUE;
}

/* A finalize handler which does not close the handle.  */
static void
noclose_finalize_handler (work_item_t item)
{
  if (debug_ioworker)
    log_debug ("%s:%s: [%s:%p] called",
               SRCNAME, __func__, item->name, item->hd);
  item->hd = INVALID_HANDLE_VALUE;
}


/* Add a data callback and a handle to the work queue.  This should
   only be called once per handle.  Caller gives up ownership of
   CLD. */
static void
enqueue_callback (const char *name, assuan_context_t ctx, 
                  gpgme_data_t data, HANDLE hd,
                  int for_write, void (*fin_handler)(work_item_t),
                  ULONG cmdid, closure_data_t cld, 
                  int wait_on_success, int inactive)
{
  work_item_t item;
  int created = 0;

  EnterCriticalSection (&work_queue_lock);
  for (item = work_queue; item; item = item->next)
    if (!item->used)
      break;
  if (!item)
    {
      item = xmalloc (sizeof *item);
      item->next = work_queue;
      work_queue = item;
      created = 1;
    }
  item->used = 1;
  item->name = name;
  item->cmdid = cmdid;
  item->cld = cld;
  item->wait_on_success = wait_on_success;
  item->data = data;
  item->writing = for_write;
  item->hd = hd;
  item->inactive = inactive;
  item->io_pending = 0;
  item->got_ready = 0;
  item->delayed_ready = 0;
  item->got_error = 0;
  item->aborting = 0;
  item->finalize = fin_handler;
  memset (&item->ov, 0, sizeof item->ov);
  if (debug_ioworker)
    log_debug ("%s:%s: [%s:%p] created%s",
               SRCNAME, __func__, item->name, item->hd,
               created?"":" (reusing)");
  LeaveCriticalSection (&work_queue_lock);
}


/* Set all items of CMDID into the active state.  */
static void
set_items_active (ULONG cmdid)
{
  work_item_t item;

  EnterCriticalSection (&work_queue_lock);
  for (item = work_queue; item; item = item->next)
    if (item->used && item->cmdid == cmdid)
      item->inactive = 0;
  LeaveCriticalSection (&work_queue_lock);
}


static void
destroy_command (ULONG cmdid, int force)
{
  work_item_t item;

  EnterCriticalSection (&work_queue_lock);
  for (item = work_queue; item; item = item->next)
    if (item->used && item->cmdid == cmdid 
        && (!item->wait_on_success || force))
      {
        if (debug_ioworker)
          log_debug ("%s:%s: [%s:%p] cmdid=%lu registered for destroy",
                     SRCNAME, __func__, item->name, item->hd, item->cmdid);
        /* First send an I/O cancel in case the last
           GetOverlappedResult returned only a partial result.  This
           works because we are always running within the
           async_worker_thread.  */
/*         if (!CancelIo (item->hd)) */
/*           log_error_w32 (-1, "%s:%s: [%s:%p] CancelIo failed", */
/*                          SRCNAME, __func__, item->name, item->hd); */
        item->got_ready = 1;
      }
  LeaveCriticalSection (&work_queue_lock);
}


/* Process a status line.  */
static int
status_handler (closure_data_t cld, const char *line)
{
  gpg_error_t err;
  int retval = 0;

  if (debug_ioworker)
    log_debug ("%s:%s: cld %p, line `%s'", SRCNAME, __func__, cld, line);

  if (*line == '#' || !*line)
    ;
  else if (line[0] == 'D' && line[1] == ' ')
    {
      line += 2;
    }
  else if (line[0] == 'S' && (!line[1] || line[1] == ' '))
    {
      for (line += 1; *line == ' '; line++)
        ;
    }  
  else if (line[0] == 'O' && line[1] == 'K' && (!line[2] || line[2] == ' '))
    {
      for (line += 2; *line == ' '; line++)
        ;
      cld->final_err = 0;
      retval = 1;
    }
  else if (!strncmp (line, "ERR", 3) && (!line[3] || line[3] == ' '))
    {
      for (line += 3; *line == ' '; line++)
        ;
      err = strtoul (line, NULL, 10);
      if (!err)
        err = gpg_error (GPG_ERR_ASS_INV_RESPONSE);
      cld->final_err = err;
      retval = 1;
    }  
  else if (!strncmp (line, "INQUIRE", 7) && (!line[7] || line[7] == ' '))
    {
      for (line += 7; *line == ' '; line++)
        ;
      /* We have no inquire handler thus get out of it immediately.  */
      err = assuan_write_line (cld->assctx, "END");
      if (err)
        cld->last_err = err;
    }
  else if (!strncmp (line, "END", 3) && (!line[3] || line[3] == ' '))
    {
      for (line += 3; *line == ' '; line++)
        ;
    }
  else
    retval = -1; /* Invalid response.  */

  return retval;
}


/* This write callback is used by GPGME to push data to our status
   line handler.  The function should return the number of bytes
   written, and -1 on error.  If an error occurs, ERRNO should be set
   to describe the type of the error.  */
static ssize_t
status_in_cb (void *opaque, const void *buffer, size_t size)
{
  size_t orig_size = size;
  closure_data_t cld = opaque;
  status_buffer_t sb;
  size_t nleft, nbytes;
  char *p;

  assert (cld);
  if (!size)
    return 0;

  if (!(sb=cld->status_buffer))
    {
      cld->status_buffer = sb = xmalloc (sizeof *cld->status_buffer);
      sb->eof = 0;
      sb->linelen = 0;
    }

  do
    {
      assert (sb->linelen < ASSUAN_LINELENGTH);
      nleft = ASSUAN_LINELENGTH - sb->linelen;
      nbytes = size < nleft? size : nleft;
      memcpy (sb->line+sb->linelen, buffer, nbytes);
      sb->linelen += nbytes;
      size -= nbytes;
      while ((p = memchr (sb->line, '\n', sb->linelen)) && !cld->status_ready)
        {
          *p = 0;
          if (p > sb->line && p[-1] == '\r')
            p[-1] = 0;
          switch (status_handler (cld, sb->line))
            {
            case 0: 
              break;
            case 1: /* Ready. */
              cld->status_ready = 1;
              destroy_command (cld->cmdid, 0);
              break;
            default:
              log_error ("%s:%s: invalid line from server", SRCNAME, __func__);
              errno = EINVAL;
              return -1;
            }
          sb->linelen -= (p+1 - sb->line);
          memmove (sb->line, p+1, sb->linelen);
        }
      if (sb->linelen >= ASSUAN_LINELENGTH)
        {
          log_error ("%s:%s: line from server too long", SRCNAME, __func__);
          errno = ERANGE;
          return -1;
        }
    }
  while (size);
  
  return orig_size;
}



/* Start an asynchronous command.  Caller gives up ownership of
   CLD.  */
static gpg_error_t
start_command (assuan_context_t ctx, closure_data_t cld,
               ULONG cmdid, const char *line)
{
  gpg_error_t err;
  assuan_fd_t fds[5];
  int nfds;

  if (debug_ioworker)
    log_debug ("%s:%s: sending `%s'", SRCNAME, __func__, line);

  /* Get the fd used by assuan for status channel reads.  This is the
     first fd returned by assuan_get_active_fds for read fds.  */
  nfds = assuan_get_active_fds (ctx, 0, fds, DIM (fds));
  if (nfds < 1)
    return gpg_error (GPG_ERR_GENERAL);	/* Ooops.  */

  cld->cmdid = cmdid;
  cld->status_cbs.write = status_in_cb;
  cld->assctx = ctx;
  /* Fixme: We might want to have reference counting for CLD to cope
     with thye problem that the gpgme data object uses CLD which might
     get invalidated at any time.  */
  err = gpgme_data_new_from_cbs (&cld->status_data, &cld->status_cbs, cld);
  if (err)
    {
      xfree (cld);
      return err;
    }

  set_items_active (cmdid);
  enqueue_callback ("status", ctx, cld->status_data, fds[0], 0,
                    noclose_finalize_handler, cmdid, cld, 0, 0);
  cld = NULL; /* Now belongs to the status work item.  */

  /* Process the work queue.  */
  if (!SetEvent (work_queue_event))
    log_error_w32 (-1, "%s:%s: SetEvent failed", SRCNAME, __func__);
  /* Send the command. */
  return assuan_write_line (ctx, line);
}


static const char *
get_protocol_name (protocol_t protocol)
{
  switch (protocol)
    {
    case PROTOCOL_OPENPGP: return "OpenPGP"; break;
    case PROTOCOL_SMIME:   return "CMS"; break;
    default: return NULL;
    }
}


/* Callback used to get the protocool status line form a PREP_ENCRYPT
   or SENDER command.  */
static assuan_error_t
prep_foo_status_cb (void *opaque, const char *line)
{
  protocol_t *protocol = opaque;

  if (!strncmp (line, "PROTOCOL", 8) && (line[8]==' ' || !line[8]))
    {
      for (line += 8; *line == ' '; line++)
        ;
      if (!strncmp (line, "OpenPGP", 7) && (line[7]==' '||!line[7]))
        *protocol = PROTOCOL_OPENPGP;
      else if (!strncmp (line, "CMS", 3) && (line[3]==' '||!line[3]))
        *protocol = PROTOCOL_SMIME;
    }
  return 0;
}




/* Note that this closure is called in the context of the
   async_worker_thread.  */
static void
encrypt_closure (closure_data_t cld)
{
  engine_private_finished (cld->filter, cld->final_err);
}


/* Encrypt the data from INDATA to the OUTDATA object for all
   recpients given in the NULL terminated array RECIPIENTS.  This
   function terminates with success and then expects the caller to
   wait for the result of the encryption using engine_wait.  FILTER is
   used for asynchronous commnication with the engine module.  HWND is
   the window handle of the current window and used to maintain the
   correct relationship between a popups and the active window.  If
   this function returns success, the data objects may only be
   destroyed after an engine_wait or engine_cancel.  On success the
   function returns a poiunter to the encryption state and thus
   requires that op_assuan_encrypt_bottom will be run later. */
int
op_assuan_encrypt (protocol_t protocol, 
                   gpgme_data_t indata, gpgme_data_t outdata,
                   engine_filter_t filter, void *hwnd,
                   char **recipients, protocol_t *r_used_protocol,
                   struct engine_assuan_encstate_s **r_encstate)
{
  gpg_error_t err;
  closure_data_t cld;
  assuan_context_t ctx;
  char line[1024];
  HANDLE inpipe[2], outpipe[2];
  ULONG cmdid;
  pid_t pid;
  int i;
  char *p;
  int detect_protocol;
  const char *protocol_name;
  struct engine_assuan_encstate_s *encstate;

  *r_encstate = NULL;

  detect_protocol = !(protocol_name = get_protocol_name (protocol));

  err = connect_uiserver (&ctx, &pid, &cmdid, hwnd);
  if (err)
    return err;

  if ((err = create_io_pipe (inpipe, pid, 1)))
    return err;
  if ((err = create_io_pipe (outpipe, pid, 0)))
    {
      close_pipe (inpipe);
      return err;
    }

  cld = xcalloc (1, sizeof *cld);
  cld->closure = encrypt_closure;
  cld->filter = filter;

  err = assuan_transact (ctx, "RESET", NULL, NULL, NULL, NULL, NULL, NULL);
  if (err)
    goto leave;
  for (i=0; recipients && recipients[i]; i++)
    {
      snprintf (line, sizeof line, "RECIPIENT %s", recipients[i]);
      for (p=line; *p; p++)
        if (*p == '\n' || *p =='\r' )
          *p = ' ';
      err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
      if (err)
        goto leave;
    }

  /* If the protocol has not been given, let the UI server tell us the
     protocol to use. */
  if (detect_protocol)
    {
      protocol = PROTOCOL_UNKNOWN;
      err = assuan_transact (ctx, "PREP_ENCRYPT", NULL, NULL, NULL, NULL,
                             prep_foo_status_cb, &protocol);
      if (err)
        {
          if (gpg_err_code (err) == GPG_ERR_ASS_UNKNOWN_CMD)
            err = gpg_error (GPG_ERR_INV_VALUE);
          goto leave;
        }
      if ( !(protocol_name = get_protocol_name (protocol)) )
        {
          err = gpg_error (GPG_ERR_INV_VALUE);
          goto leave;
        }
    }

  *r_used_protocol = protocol;

  /* Note: We don't use real descriptor passing but a hack: We
     duplicate the handle into the server process and the server then
     uses this handle.  Eventually we should put this code into
     assuan_sendfd.  */
  snprintf (line, sizeof line, "INPUT FD=%ld", (unsigned long int)inpipe[0]);
  err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
  if (err)
    goto leave;
  snprintf (line, sizeof line, "OUTPUT FD=%ld", (unsigned long int)outpipe[1]);
  err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
  if (err)
    goto leave;

  /* The work items are created as inactive, so that the worker thread
     ignores them.  They are set to active by with the start_command
     function called by op_assuan_encrypt_bottom.  */
  enqueue_callback (" input", ctx, indata, inpipe[1], 1, finalize_handler,
                    cmdid, NULL, 0, 1); 
  enqueue_callback ("output", ctx, outdata, outpipe[0], 0, finalize_handler, 
                    cmdid, NULL, 1 /* Wait on success */, 1); 

  encstate = xcalloc (1, sizeof *encstate);
  encstate->filter = filter;
  encstate->protocol_name = protocol_name;
  encstate->inpipe[0] = inpipe[0];
  encstate->inpipe[1] = inpipe[1];
  encstate->outpipe[0] = outpipe[0];
  encstate->outpipe[1] = outpipe[1];
  encstate->cld = cld;
  encstate->ctx = ctx;
  encstate->cmdid = cmdid;
  *r_encstate = encstate;
  return 0;

 leave:
  if (err)
    {
      /* Fixme: Cancel stuff in the work_queue. */
      close_pipe (inpipe);
      close_pipe (outpipe);
      xfree (cld);
      assuan_disconnect (ctx);
    }
  else
    engine_private_set_cancel (filter, ctx);
  return err;
}

/* Continue and actually start the encryption or cancel it with CANCEL
   set to TRUE.  The fucntion takes ownvership of ENCSTATE.  */
int
op_assuan_encrypt_bottom (struct engine_assuan_encstate_s *encstate,
                          int cancel)
{
  char line[1024];
  gpg_error_t err;

  if (!encstate)
    return 0;
  if (cancel)
    err = gpg_error (GPG_ERR_CANCELED);
  else
    {
      snprintf (line, sizeof line, "ENCRYPT --protocol=%s",
                encstate->protocol_name);
      err = start_command (encstate->ctx, encstate->cld, 
                           encstate->cmdid, line);
      encstate->cld = NULL; /* Now owned by start_command.  */
    }

  if (err)
    {
      xfree (encstate->cld);
      encstate->cld = NULL;
      engine_private_set_cancel (encstate->filter, NULL);
      close_pipe (encstate->inpipe);
      close_pipe (encstate->outpipe);
      if (cancel)
        destroy_command (encstate->cmdid, 1);
      assuan_disconnect (encstate->ctx);
      encstate->ctx = NULL;
    }
  else
    engine_private_set_cancel (encstate->filter, encstate->ctx);
  xfree (encstate);
  return err;
}




/* Note that this closure is called in the context of the
   async_worker_thread.  */
static void
sign_closure (closure_data_t cld)
{
  engine_private_finished (cld->filter, cld->final_err);
}


/* Created a detached signature for INDATA and write it to OUTDATA.
   On termination of the signing command engine_private_finished() is
   called with FILTER as the first argument.  SENDER is the sender's
   mail address (a mailbox).  The used protocol wioll be stored at
   R_PROTOCOL. */
int 
op_assuan_sign (protocol_t protocol, 
                gpgme_data_t indata, gpgme_data_t outdata,
                engine_filter_t filter, void *hwnd,
                const char *sender, protocol_t *r_used_protocol)
{
  gpg_error_t err;
  closure_data_t cld;
  assuan_context_t ctx;
  char line[1024];
  HANDLE inpipe[2], outpipe[2];
  ULONG cmdid;
  pid_t pid;
  int detect_protocol;
  const char *protocol_name;
  protocol_t suggested_protocol;

  detect_protocol = !(protocol_name = get_protocol_name (protocol));

  err = connect_uiserver (&ctx, &pid, &cmdid, hwnd);
  if (err)
    return err;

  if ((err = create_io_pipe (inpipe, pid, 1)))
    return err;
  if ((err = create_io_pipe (outpipe, pid, 0)))
    {
      close_pipe (inpipe);
      return err;
    }

  cld = xcalloc (1, sizeof *cld);
  cld->closure = sign_closure;
  cld->filter = filter;

  err = assuan_transact (ctx, "RESET", NULL, NULL, NULL, NULL, NULL, NULL);
  if (err)
    goto leave;

  /* We always send the SENDER command becuase it allows us to figure
     out the protocol to use.  In case the UI server faisl to send the
     protocol we fall back to OpenPGP.  */
  suggested_protocol = PROTOCOL_UNKNOWN;
  if (!sender)
    sender = "<kleopatra-does-not-allow-an-empty-arg@example.net>";
  snprintf (line, sizeof line, "SENDER%s%s", sender? " ":"", sender?sender:"");
  err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL,
                         prep_foo_status_cb, &suggested_protocol);
  if (err)
    {
      if (gpg_err_code (err) == GPG_ERR_ASS_UNKNOWN_CMD)
        err = gpg_error (GPG_ERR_INV_VALUE);
      goto leave;
    }
  if (detect_protocol)
    {
      log_debug ("%s:%s: suggested protocol is %d", 
                 SRCNAME, __func__, suggested_protocol);
      protocol = (suggested_protocol == PROTOCOL_UNKNOWN?
                  PROTOCOL_OPENPGP : suggested_protocol);
      if ( !(protocol_name = get_protocol_name (protocol)) )
        {
          err = gpg_error (GPG_ERR_INV_VALUE);
          goto leave;
        }
    }
  *r_used_protocol = protocol;
  log_debug ("%s:%s: using protocol %s", SRCNAME, __func__, protocol_name);

  snprintf (line, sizeof line, "INPUT FD=%ld", (unsigned long int)inpipe[0]);
  err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
  if (err)
    goto leave;
  snprintf (line, sizeof line, "OUTPUT FD=%ld", (unsigned long int)outpipe[1]);
  err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
  if (err)
    goto leave;

  enqueue_callback (" input", ctx, indata, inpipe[1], 1, finalize_handler,
                    cmdid, NULL, 0, 0); 
  enqueue_callback ("output", ctx, outdata, outpipe[0], 0, finalize_handler, 
                    cmdid, NULL, 1 /* Wait on success */, 0); 

  snprintf (line, sizeof line, "SIGN --protocol=%s --detached",
            protocol_name);
  err = start_command (ctx, cld, cmdid, line);
  cld = NULL; /* Now owned by start_command.  */
  if (err)
    goto leave;


 leave:
  if (err)
    {
      /* Fixme: Cancel stuff in the work_queue. */
      close_pipe (inpipe);
      close_pipe (outpipe);
      xfree (cld);
      assuan_disconnect (ctx);
    }
  else
    engine_private_set_cancel (filter, ctx);
  return err;
}




/* Note that this closure is called in the context of the
   async_worker_thread.  */
static void
decrypt_closure (closure_data_t cld)
{
  engine_private_finished (cld->filter, cld->final_err);
}


/* Decrypt data from INDATA to OUTDATE.  If WITH_VERIFY is set, the
   signature of a PGP/MIME combined message is also verified the same
   way as with op_assuan_verify.  */
int 
op_assuan_decrypt (protocol_t protocol,
                   gpgme_data_t indata, gpgme_data_t outdata, 
                   engine_filter_t filter, void *hwnd,
                   int with_verify)
{
  gpg_error_t err;
  closure_data_t cld;
  assuan_context_t ctx;
  char line[1024];
  HANDLE inpipe[2], outpipe[2];
  ULONG cmdid;
  pid_t pid;
  const char *protocol_name;

  if (!(protocol_name = get_protocol_name (protocol)))
    return gpg_error(GPG_ERR_INV_VALUE);

  err = connect_uiserver (&ctx, &pid, &cmdid, hwnd);
  if (err)
    return err;

  if ((err = create_io_pipe (inpipe, pid, 1)))
    return err;
  if ((err = create_io_pipe (outpipe, pid, 0)))
    {
      close_pipe (inpipe);
      return err;
    }

  cld = xcalloc (1, sizeof *cld);
  cld->closure = decrypt_closure;
  cld->filter = filter;

  err = assuan_transact (ctx, "RESET", NULL, NULL, NULL, NULL, NULL, NULL);
  if (err)
    goto leave;

  snprintf (line, sizeof line, "INPUT FD=%ld", (unsigned long int)inpipe[0]);
  err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
  if (err)
    goto leave;
  snprintf (line, sizeof line, "OUTPUT FD=%ld", (unsigned long int)outpipe[1]);
  err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
  if (err)
    goto leave;

  enqueue_callback (" input", ctx, indata, inpipe[1], 1, finalize_handler,
                    cmdid, NULL, 0, 0); 
  enqueue_callback ("output", ctx, outdata, outpipe[0], 0, finalize_handler, 
                    cmdid, NULL, 1 /* Wait on success */, 0); 

  snprintf (line, sizeof line, "DECRYPT --protocol=%s%s",
            protocol_name, with_verify? "":" --no-verify");
  err = start_command (ctx, cld, cmdid, line);
  cld = NULL; /* Now owned by start_command.  */
  if (err)
    goto leave;


 leave:
  if (err)
    {
      /* Fixme: Cancel stuff in the work_queue. */
      close_pipe (inpipe);
      close_pipe (outpipe);
      xfree (cld);
      assuan_disconnect (ctx);
    }
  else
    engine_private_set_cancel (filter, ctx);
  return err;
}



/* Note that this closure is called in the context of the
   async_worker_thread.  */
static void
verify_closure (closure_data_t cld)
{
  gpgme_data_release (cld->sigdata);
  cld->sigdata = NULL;
  engine_private_finished (cld->filter, cld->final_err);
}


/* With MSGDATA, SIGNATURE and SIGLEN given: 

      Verify a detached message where the data is in the gpgme object
      MSGDATA and the signature given as the string SIGNATURE. 

   With MSGDATA and OUTDATA given:

      Verify an opaque signature from MSGDATA and write the decoded
      plaintext to OUTDATA.

*/
int 
op_assuan_verify (gpgme_protocol_t protocol, 
                  gpgme_data_t msgdata, const char *signature, size_t sig_len,
                  gpgme_data_t outdata,
                  engine_filter_t filter, void *hwnd)
{
  gpg_error_t err;
  closure_data_t cld = NULL;
  assuan_context_t ctx;
  char line[1024];
  HANDLE msgpipe[2], sigpipe[2], outpipe[2];
  ULONG cmdid;
  pid_t pid;
  gpgme_data_t sigdata = NULL;
  const char *protocol_name;
  int opaque_mode;

  msgpipe[0] = INVALID_HANDLE_VALUE;
  msgpipe[1] = INVALID_HANDLE_VALUE;
  sigpipe[0] = INVALID_HANDLE_VALUE;
  sigpipe[1] = INVALID_HANDLE_VALUE;
  outpipe[0] = INVALID_HANDLE_VALUE;
  outpipe[1] = INVALID_HANDLE_VALUE;

  if (!(protocol_name = get_protocol_name (protocol)))
    return gpg_error(GPG_ERR_INV_VALUE);

  if (signature && sig_len && !outdata)
    opaque_mode = 0;
  else if (!signature && !sig_len && outdata)
    opaque_mode = 1;
  else
    return gpg_error(GPG_ERR_INV_VALUE);

  if (!opaque_mode)
    {
      err = gpgme_data_new_from_mem (&sigdata, signature, sig_len, 0);
      if (err)
        goto leave;
    }

  err = connect_uiserver (&ctx, &pid, &cmdid, hwnd);
  if (err)
    goto leave;

  if (!opaque_mode)
    {
      if ((err = create_io_pipe (msgpipe, pid, 1)))
        goto leave;
      if ((err = create_io_pipe (sigpipe, pid, 1)))
        goto leave;
    }
  else
    {
      if ((err = create_io_pipe (msgpipe, pid, 1)))
        goto leave;
      if ((err = create_io_pipe (outpipe, pid, 0)))
        goto leave;
    }

  cld = xcalloc (1, sizeof *cld);
  cld->closure = verify_closure;
  cld->filter = filter;
  cld->sigdata = sigdata;

  err = assuan_transact (ctx, "RESET", NULL, NULL, NULL, NULL, NULL, NULL);
  if (err)
    goto leave;

  if (!opaque_mode)
    {
      snprintf (line, sizeof line, "MESSAGE FD=%ld",
                (unsigned long int)msgpipe[0]);
      err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
      if (err)
        goto leave;
      snprintf (line, sizeof line, "INPUT FD=%ld",
                (unsigned long int)sigpipe[0]);
      err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
      if (err)
        goto leave;
      enqueue_callback ("   msg", ctx, msgdata, msgpipe[1], 1,
                        finalize_handler, cmdid, NULL, 0, 0); 
      enqueue_callback ("   sig", ctx, sigdata, sigpipe[1], 1, 
                        finalize_handler, cmdid, NULL, 0, 0); 
    }
  else 
    {
      snprintf (line, sizeof line, "INPUT FD=%ld",
                (unsigned long int)msgpipe[0]);
      err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
      if (err)
        goto leave;
      snprintf (line, sizeof line, "OUTPUT FD=%ld",
                (unsigned long int)outpipe[1]);
      err = assuan_transact (ctx, line, NULL, NULL, NULL, NULL, NULL, NULL);
      if (err)
        goto leave;
      enqueue_callback ("   msg", ctx, msgdata, msgpipe[1], 1,
                        finalize_handler, cmdid, NULL, 0, 0); 
      enqueue_callback ("   out", ctx, outdata, outpipe[0], 0,
                        finalize_handler, cmdid, NULL, 1, 0); 
    }

  snprintf (line, sizeof line, "VERIFY --protocol=%s",  protocol_name);
  err = start_command (ctx, cld, cmdid, line);
  cld = NULL;     /* Now owned by start_command.  */
  sigdata = NULL; /* Ditto.  */
  if (err)
    goto leave;


 leave:
  if (err)
    {
      /* Fixme: Cancel stuff in the work_queue. */
      close_pipe (msgpipe);
      close_pipe (sigpipe);
      close_pipe (outpipe);
      gpgme_data_release (sigdata);
      xfree (cld);
      assuan_disconnect (ctx);
    }
  else
    engine_private_set_cancel (filter, ctx);
  return err;
}



/* Ask the server to fire up the key manager.  */
int 
op_assuan_start_keymanager (void *hwnd)
{
  gpg_error_t err;
  assuan_context_t ctx;
  ULONG cmdid;
  pid_t pid;

  err = connect_uiserver (&ctx, &pid, &cmdid, hwnd);
  if (!err)
    {
      err = assuan_transact (ctx, "START_KEYMANAGER",
                             NULL, NULL, NULL, NULL, NULL, NULL);
      assuan_disconnect (ctx);
    }
  return err;
}
