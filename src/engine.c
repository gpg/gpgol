/* engine.c - Crypto engine dispatcher
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
#include <windows.h>

#include "common.h"
#include "engine.h"
#include "engine-gpgme.h"
#include "engine-assuan.h"


#define FILTER_BUFFER_SIZE 4096


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                       SRCNAME, __func__, __LINE__); \
                        } while (0)

#define debug_filter        (opt.enable_debug & DBG_FILTER)
#define debug_filter_extra  (opt.enable_debug & DBG_FILTER_EXTRA)

/* This variable indicates whether the assuan engine is used.  */
static int use_assuan;


/* Definition of the key object.  */
struct engine_keyinfo_s
{
  struct {
    gpgme_key_t key;
  } gpgme;

};


/* Definition of the filter object.  This object shall only be
   accessed by one thread. */
struct engine_filter_s
{
  int use_assuan;          /* The same as the global USE_ASSUAN.  */

  struct {
    CRITICAL_SECTION lock; /* The lock for the this object. */
    HANDLE condvar;        /* Manual reset event signaled if LENGTH > 0.  */
    int nonblock;          /* Put gpgme data cb in non blocking mode.  */
    size_t length;         /* Number of bytes in BUFFER waiting to be
                              send down the pipe.  */
    char buffer[FILTER_BUFFER_SIZE];
    int got_eof;         /* End of file has been indicated.  */
    /* These objects are only in this structure because we
       use this structure's lock to protect them.  */
    int ready;           /* Set to true if the gpgme process has finished.  */
    HANDLE ready_event;  /* And the corresponding event.  */
    gpg_error_t status;  /* Status of the gpgme process.  */
  } in;

  struct {
    CRITICAL_SECTION lock; /* The lock for the this object. */
    HANDLE condvar;        /* Manual reset event signaled if LENGTH == 0.  */
    int nonblock;          /* Put gpgme data cb in non blocking mode.  */
    size_t length;         /* Number of bytes in BUFFER waiting to be
                              send back to the caller.  */
    char buffer[FILTER_BUFFER_SIZE];
  } out;

  /* Flag to push an extra LF out.  */
  int add_extra_lf;

  /* The data sink as set by engine_create_filter.  */
  int (*outfnc) (void *, const void *, size_t);
  void *outfncdata;

  /* Objects to be released by engine_wait/cancel.  */
  struct gpgme_data_cbs cb_inbound;  /* Callback structure for gpgme.  */
  struct gpgme_data_cbs cb_outbound; /* Ditto.  */
  gpgme_data_t indata;               /* Input data.  */
  gpgme_data_t outdata;              /* Output data.  */
  void *cancel_data;                 /* Used by engine_cancel.  */

  /* A pointer used convey information from engine_encrypt_prepare to
     engine_encrypt_start.  */
  struct engine_assuan_encstate_s *encstate;

  /* Counter used to optimize voluntary thread switching. */
  ULONG switch_counter;

  /* Optional data to be passed to the backend to help the server
     display a useful title and to associate operations with one OL
     task.  */
  unsigned int session_number;
  char *session_title;     /* (malloced) */
  char *sender_address;    /* (malloced) */
};


static void
take_in_lock (engine_filter_t filter, const char *func)
{
  EnterCriticalSection (&filter->in.lock);
  if (debug_filter_extra)
    log_debug ("%s:%s: in.lock taken\n", SRCNAME, func);
}

static void
release_in_lock (engine_filter_t filter, const char *func)
{
  LeaveCriticalSection (&filter->in.lock);
  if (debug_filter_extra)
    log_debug ("%s:%s: in.lock released\n", SRCNAME, func);
}

static void
take_out_lock (engine_filter_t filter, const char *func)
{
  EnterCriticalSection (&filter->out.lock);
  if (debug_filter_extra)
    log_debug ("%s:%s: out.lock taken\n", SRCNAME, func);
}

static void
release_out_lock (engine_filter_t filter, const char *func)
{
  LeaveCriticalSection (&filter->out.lock);
  if (debug_filter_extra)
    log_debug ("%s:%s: out.lock released\n", SRCNAME, func);
}





/* Create a new filter object. */
static engine_filter_t 
create_filter (void)
{
  engine_filter_t filter;

  filter = xcalloc (1, sizeof *filter);

  InitializeCriticalSection (&filter->in.lock);
  filter->in.condvar = CreateEvent (NULL, TRUE, 0, NULL);
  if (!filter->in.condvar)
    log_error_w32 (-1, "%s:%s: create in.condvar failed", SRCNAME, __func__);

  InitializeCriticalSection (&filter->out.lock);
  filter->out.condvar = CreateEvent (NULL, TRUE, 0, NULL);
  if (!filter->out.condvar)
    log_error_w32 (-1, "%s:%s: create out.condvar failed", SRCNAME, __func__);

  /* Create an automatic event (it only used one time so the type is
     actually not important).  */
  filter->in.ready_event = CreateEvent (NULL, 0, 0, NULL);
  if (!filter->in.ready_event)
    log_error_w32 (-1, "%s:%s: CreateEvent failed", SRCNAME, __func__);

  /* If we are using the assuan engine we need to make the gpgme read
     callback non blocking.  */
  if (use_assuan)
    {
      filter->use_assuan = 1;
      filter->in.nonblock = 1;
    }

  return filter;
}


static void
release_filter (engine_filter_t filter)
{
  if (filter)
    {
      if (filter->in.condvar)
        CloseHandle (filter->in.condvar);
      if (filter->out.condvar)
        CloseHandle (filter->out.condvar);
      if (filter->in.ready_event)
        CloseHandle (filter->in.ready_event);
      gpgme_data_release (filter->indata);
      gpgme_data_release (filter->outdata);
      xfree (filter->session_title);
      xfree (filter->sender_address);
      xfree (filter);
    }
}


/* This is a wraper around SwitchToThread, a syscall we unfortunately
   need due to the lack of an sophisticated event system.  The wrapper
   calls SwitchToThread but after a couple of immediate folliwing
   switches, it introduces a short delays.  */
static void
switch_threads (engine_filter_t filter)
{
  ULONG count;

  count = InterlockedExchangeAdd (&filter->switch_counter, 1);
  if (count > 5)
    {
      InterlockedExchange (&filter->switch_counter, 5);
      SleepEx (50, TRUE); 
    }
  else if (!SwitchToThread ())
    {
      /* No runable other thread: Fall asleep. */
      SleepEx (5, TRUE);
    }
}

/* Call this fucntion if some action has been done.  */
static void
clear_switch_threads (engine_filter_t filter)
{
  InterlockedExchange (&filter->switch_counter, 0);
}


/* This read callback is used by GPGME to read data from a filter
   object.  The function should return the number of bytes read, 0 on
   EOF, and -1 on error.  If an error occurs, ERRNO should be set to
   describe the type of the error.  */
static ssize_t
filter_gpgme_read_cb (void *handle, void *buffer, size_t size)
{
  engine_filter_t filter = handle;
  int nbytes;

  if (!filter || !buffer || !size)
    {
      errno = EINVAL;
      return (ssize_t)(-1);
    }

  if (debug_filter)
    log_debug ("%s:%s: enter\n",  SRCNAME, __func__);
  take_in_lock (filter, __func__);
  while (!filter->in.length)
    {
      if (filter->in.got_eof || filter->in.ready)
        {
          release_in_lock (filter, __func__);
          if (debug_filter)
            log_debug ("%s:%s: returning EOF\n", SRCNAME, __func__);
          return 0; /* Return EOF. */
        }
      release_in_lock (filter, __func__);
      if (filter->in.nonblock)
        {
          errno = EAGAIN;
          if (debug_filter_extra)
            log_debug ("%s:%s: leave; result=EAGAIN\n", SRCNAME, __func__);
          switch_threads (filter);
          return -1;
        }
      else
        clear_switch_threads (filter);
      if (debug_filter)
        log_debug ("%s:%s: waiting for in.condvar\n", SRCNAME, __func__);
      WaitForSingleObject (filter->in.condvar, 500);
      take_in_lock (filter, __func__);
      if (debug_filter)
        log_debug ("%s:%s: continuing\n", SRCNAME, __func__);
    }
     
  if (debug_filter)
    log_debug ("%s:%s: requested read size=%d (filter.in.length=%d)\n",
               SRCNAME, __func__, (int)size, (int)filter->in.length);
  nbytes = size < filter->in.length ? size : filter->in.length;
  memcpy (buffer, filter->in.buffer, nbytes);
  if (filter->in.length > nbytes)
    memmove (filter->in.buffer, filter->in.buffer + nbytes, 
             filter->in.length - nbytes);
  filter->in.length -= nbytes;
  release_in_lock (filter, __func__);

  if (debug_filter)
    log_debug ("%s:%s: leave; result=%d\n",
               SRCNAME, __func__, (int)nbytes);
  
  return nbytes;
}


/* This write callback is used by GPGME to write data to the filter
   object.  The function should return the number of bytes written,
   and -1 on error.  If an error occurs, ERRNO should be set to
   describe the type of the error.  */
static ssize_t
filter_gpgme_write_cb (void *handle, const void *buffer, size_t size)
{
  engine_filter_t filter = handle;
  int nbytes;

  if (!filter || !buffer || !size)
    {
      errno = EINVAL;
      return (ssize_t)(-1);
    }

  if (debug_filter)
    log_debug ("%s:%s: enter\n",  SRCNAME, __func__);
  take_out_lock (filter, __func__);
  while (filter->out.length)
    {
      release_out_lock (filter, __func__);
      if (filter->out.nonblock)
        {
          errno = EAGAIN;
          if (debug_filter_extra)
            log_debug ("%s:%s: leave; result=EAGAIN\n", SRCNAME, __func__);
          return -1;
        }
      if (debug_filter)
        log_debug ("%s:%s: waiting for out.condvar\n", SRCNAME, __func__);
      WaitForSingleObject (filter->out.condvar, 500);
      take_out_lock (filter, __func__);
      if (debug_filter)
        log_debug ("%s:%s: continuing\n", SRCNAME, __func__);
    }

  if (debug_filter)
    log_debug ("%s:%s: requested write size=%d\n",
               SRCNAME, __func__, (int)size);
  nbytes = size < FILTER_BUFFER_SIZE ? size : FILTER_BUFFER_SIZE;
  memcpy (filter->out.buffer, buffer, nbytes);
  filter->out.length = nbytes;
  release_out_lock (filter, __func__);

  if (debug_filter)
    log_debug ("%s:%s: leave; result=%d\n", SRCNAME, __func__, (int)nbytes);
  return nbytes;
}


/* Store a cancel parameter into FILTER.  Only use by the engine backends. */
void
engine_private_set_cancel (engine_filter_t filter, void *cancel_data)
{
  filter->cancel_data = cancel_data;
}


/* This function is called by the gpgme backend to notify a filter
   object about the final status of an operation.  It may only be
   called by the engine-gpgme.c module. */
void
engine_private_finished (engine_filter_t filter, gpg_error_t status)
{
  if (!filter)
    {
      log_debug ("%s:%s: called without argument\n", SRCNAME, __func__);
      return;
    }
  if (debug_filter)
    log_debug ("%s:%s: filter %p: process terminated: %s <%s>\n", 
               SRCNAME, __func__, filter, 
               gpg_strerror (status), gpg_strsource (status));
  
  take_in_lock (filter, __func__);
  filter->in.ready = 1;
  filter->in.status = status;
  filter->cancel_data = NULL;
  if (!SetEvent (filter->in.ready_event))
    log_error_w32 (-1, "%s:%s: SetEvent failed", SRCNAME, __func__);
  release_in_lock (filter, __func__);
  if (debug_filter)
    log_debug ("%s:%s: leaving\n", SRCNAME, __func__);
}


unsigned int
engine_private_get_session_number (engine_filter_t filter)
{
  return filter? filter->session_number : 0;
}

const char *
engine_private_get_session_title (engine_filter_t filter)
{
  return filter? filter->session_title : NULL;
}



/* Initialize the engine dispatcher.  */
int
engine_init (void)
{
  gpg_error_t err;

  err = op_gpgme_basic_init ();
  if (err)
    return err;

  err = op_assuan_init ();
  if (err)
    {
      use_assuan = 0;
      MessageBox (NULL,
                  _("The user interface server is not available or does "
                    "not work.  Using an internal user interface.\n\n"
                    "This is limited to the PGP/MIME protocol and "
                    "thus S/MIME protected message are not readable."),
                  _("GpgOL"), MB_ICONWARNING|MB_OK);
      err = op_gpgme_init ();
    }
  else
    use_assuan = 1;

  return err;
}


/* Shutdown the engine dispatcher.  */
void
engine_deinit (void)
{
  op_assuan_deinit ();
  op_gpgme_deinit ();
}

/* Helper function to return a new session number.  */
unsigned int
engine_new_session_number (void)
{
  static ULONG number;
  ULONG n;

  /* The protocol says the session number is a 32 bit value, thus we
     make that sure even in the case that ULONG is changed to a 64 bit
     type (Not an issue under Windows, though). */
  while (!((n = InterlockedIncrement (&number))&0xffffffff) )
    ;
  return n;
}



/* Filter the INDATA of length INDATA and write the output using
   OUTFNC.  OUTFNCDATA is passed as first argument to OUTFNC, followed
   by the data to be written and its length.  FILTER is an object
   returned for example by engine_encrypt_start.  The function returns
   0 on success or an error code on error.

   Passing INDATA as NULL and INDATALEN as 0, the filter will be
   flushed, that is all remaining stuff will be written to OUTFNC.
   This indicates EOF and the filter won't accept anymore input.  */
int
engine_filter (engine_filter_t filter, const void *indata, size_t indatalen)
{
  gpg_error_t err;
  int nbytes;

  if (debug_filter)
    log_debug ("%s:%s: enter; filter=%p\n", SRCNAME, __func__, filter); 
  /* Our implementation is for now straightforward without any
     additional buffer filling etc.  */
  if (!filter || !filter->outfnc)
    return gpg_error (GPG_ERR_INV_VALUE);
  if (filter->in.length > FILTER_BUFFER_SIZE
      || filter->out.length > FILTER_BUFFER_SIZE)
    return gpg_error (GPG_ERR_BUG);
  if (filter->in.got_eof)
    return gpg_error (GPG_ERR_CONFLICT); /* EOF has already been indicated.  */

  if (debug_filter)
    log_debug ("%s:%s: indata=%p indatalen=%d outfnc=%p\n",
               SRCNAME, __func__, indata, (int)indatalen, filter->outfnc); 
  for (;;)
    {
      int any;

      /* If there is something to write out, do this now to make space
         for more data.  */
      any = 0;
      take_out_lock (filter, __func__);
      while (filter->out.length)
        {
          if (debug_filter)
            log_debug ("%s:%s: pushing %d bytes to the outfnc\n",
                       SRCNAME, __func__, filter->out.length); 
          nbytes = filter->outfnc (filter->outfncdata, 
                                   filter->out.buffer, filter->out.length);
          if (nbytes == -1)
            {
              if (debug_filter)
                log_debug ("%s:%s: error writing data\n", SRCNAME, __func__);
              release_out_lock (filter, __func__);
              return gpg_error (GPG_ERR_EIO);
            }
          assert (nbytes <= filter->out.length && nbytes >= 0);
          if (nbytes < filter->out.length)
            memmove (filter->out.buffer, filter->out.buffer + nbytes,
                     filter->out.length - nbytes); 
          filter->out.length -= nbytes;
          any = 1;
        }
      if (!PulseEvent (filter->out.condvar))
        log_error_w32 (-1, "%s:%s: PulseEvent(%p)[out] failed", 
                       SRCNAME, __func__, filter->out.condvar);
      release_out_lock (filter, __func__);

      if (any)
        clear_switch_threads (filter);
      else
        switch_threads (filter);

      any = 0;
      take_in_lock (filter, __func__);
      if (!indata && !indatalen)
        {
          filter->in.got_eof = 1;
          /* Flush requested.  Tell the output function to also flush.  */
          nbytes = filter->outfnc (filter->outfncdata, NULL, 0);
          if (nbytes == -1)
            {
              log_debug ("%s:%s: error flushing data\n", SRCNAME, __func__);
              err = gpg_error (GPG_ERR_EIO);
            }
          else
            err = 0;
          release_in_lock (filter, __func__);
          return err; 
        }

      /* Fill the input buffer, relinquish control to the callback
         processor and loop until all input data has been
         processed.  */
      if (!filter->in.length && indatalen)
        {
          filter->in.length = (indatalen > FILTER_BUFFER_SIZE
                               ? FILTER_BUFFER_SIZE : indatalen);
          memcpy (filter->in.buffer, indata, filter->in.length);
          indata    += filter->in.length;
          indatalen -= filter->in.length;
          any = 1;
        }
      /* Terminate the loop if the filter queue is empty OR the filter
         is ready and there is nothing left for output.  */
      if (!filter->in.length || (filter->in.ready && !filter->out.length))
        {
          release_in_lock (filter, __func__);
          err = filter->in.status;
          break;  /* the loop.  */
        }
      if (!PulseEvent (filter->in.condvar))
        log_error_w32 (-1, "%s:%s: PulseEvent(%p)[in] failed", 
                       SRCNAME, __func__, filter->in.condvar);
      release_in_lock (filter, __func__);
      if (any)
        clear_switch_threads (filter);
      else
        switch_threads (filter);
    }

  if (debug_filter)
    log_debug ("%s:%s: leave; err=%d\n", SRCNAME, __func__, err); 
  return err;
}


/* Dummy data sink used if caller does not need an output
   function.  */
static int
dummy_outfnc (void *opaque, const void *data, size_t datalen)
{
  (void)opaque;
  (void)data;
  return (int)datalen;
}


/* Create a new filter object which uses OUTFNC as its data sink.  If
   OUTFNC is called with NULL/0 for the data to be written, the
   function should do a flush.  OUTFNC is expected to return the
   number of bytes actually written or -1 on error.  It may return 0
   to indicate that no data has been written and the caller shall try
   again.  OUTFNC and OUTFNCDATA are internally used by the engine
   even after the call to this function.  There lifetime only ends
   after an engine_wait or engine_cancel. */
int
engine_create_filter (engine_filter_t *r_filter,
                      int (*outfnc) (void *, const void *, size_t),
                      void *outfncdata)
{
  gpg_error_t err;
  engine_filter_t filter;

  filter = create_filter ();
  filter->cb_inbound.read = filter_gpgme_read_cb;
  filter->cb_outbound.write = filter_gpgme_write_cb;
  filter->outfnc = outfnc? outfnc : dummy_outfnc;
  filter->outfncdata = outfncdata;

  err = gpgme_data_new_from_cbs (&filter->indata, 
                                 &filter->cb_inbound, filter);
  if (err)
    goto failure;

  err = gpgme_data_new_from_cbs (&filter->outdata,
                                 &filter->cb_outbound, filter);
  if (err)
    goto failure;

  *r_filter = filter;
  return 0;

 failure:
  release_filter (filter);
  return err;
}


/* Set the FILTER in a mode which pushes an extra linefeed out.  */
void
engine_request_exra_lf (engine_filter_t filter)
{
  filter->add_extra_lf = 1;
}


/* Set the session number for FILTER to NUMBER.  */
void
engine_set_session_number (engine_filter_t filter, unsigned int value)
{
  if (filter)
    filter->session_number = value;
}

/* Set the session title for FILTER to TITLE.  */
void
engine_set_session_title (engine_filter_t filter, const char *title)
{
  if (filter)
    {
      xfree (filter->session_title);
      filter->session_title = title? xstrdup (title) : NULL;
    }
}

/* Set the sender address for FILTER to ADDR.  Setting the sender
   address to NULL is allowed.  This sender address may be used by the
   server to check whether the used key matches the sender address
   while verifying a message.  Note that this sender address is not
   used for selecting a key during encryption or verification. */
void
engine_set_sender_address (engine_filter_t filter, const char *addr)
{
  if (filter)
    {
      xfree (filter->sender_address);
      filter->sender_address = addr? xstrdup (addr) : NULL;
    }
}


/* Wait for FILTER to finish.  Returns 0 on success.  FILTER is not
   valid after the function has returned success.  */
int
engine_wait (engine_filter_t filter)
{
  gpg_error_t err;
  int more;
  int nbytes; 

  if (!filter || !filter->outfnc)
    return gpg_error (GPG_ERR_INV_VALUE);

  /* If we are here, engine_filter is not anymore called but there is
     likely stuff in the output buffer which needs to be written
     out.  */
  /* Argh, Busy waiting.  As soon as we change fromCritical Sections
     to a kernel based objects we should use WaitOnMultipleObjects to
     wait for the out.lock as well as for the ready_event.  */
  do 
    {
      more = 0;
      take_out_lock (filter, __func__);
      if (filter->out.length)
        {
          nbytes = filter->outfnc (filter->outfncdata, 
                                   filter->out.buffer, filter->out.length);
          if (nbytes < 0)
            {
              log_error ("%s:%s: error writing data\n", SRCNAME, __func__);
              release_out_lock (filter, __func__);
              return gpg_error (GPG_ERR_EIO);
            }
         
          assert (nbytes <= filter->out.length && nbytes >= 0);
          if (nbytes < filter->out.length)
            memmove (filter->out.buffer, filter->out.buffer + nbytes,
                     filter->out.length - nbytes); 
          filter->out.length -= nbytes;
          if (filter->out.length)
            {
              if (debug_filter_extra)
                log_debug ("%s:%s: still %d pending bytes for outfnc\n",
                           SRCNAME, __func__, filter->out.length);
              more = 1;
            }
        }
      if (!PulseEvent (filter->out.condvar))
        log_error_w32 (-1, "%s:%s: PulseEvent(out) failed", SRCNAME, __func__);
      release_out_lock (filter, __func__);
      take_in_lock (filter, __func__);
      if (!filter->in.ready)
        more = 1;
      release_in_lock (filter, __func__);
      if (more)
        SwitchToThread ();
    }
  while (more);

  /* If requested write an extra LF, so that the MIME parser sees a
     complete line.  */
  if (filter->add_extra_lf)
    {
      int extra_written = 0;
      do 
        {
          more = 0;
          take_out_lock (filter, __func__);
          if (!extra_written)
            {
              nbytes = filter->outfnc (filter->outfncdata, "\n", 1);
              if (nbytes < 0)
                {
                  log_error ("%s:%s: error writing extra lf\n", 
                             SRCNAME, __func__);
                  release_out_lock (filter, __func__);
                  return gpg_error (GPG_ERR_EIO);
                }
              if (!nbytes)
                {
                  if (debug_filter_extra)
                    log_debug ("%s:%s: extra lf still pending for outfnc\n",
                               SRCNAME, __func__);
                  more = 1;
                }
              else
                extra_written = 1;
            }
          if (!PulseEvent (filter->out.condvar))
            log_error_w32 (-1, "%s:%s: PulseEvent(out) failed",
                           SRCNAME, __func__);
          release_out_lock (filter, __func__);
          take_in_lock (filter, __func__);
          if (!filter->in.ready)
            more = 1;
          release_in_lock (filter, __func__);
          if (more)
            SwitchToThread ();
        }
      while (more);
    }

  if (WaitForSingleObject (filter->in.ready_event, INFINITE) != WAIT_OBJECT_0)
    {
      log_error_w32 (-1, "%s:%s: WFSO failed", SRCNAME, __func__);
      return gpg_error (GPG_ERR_GENERAL);
    }
  err = filter->in.status;
  log_debug ("%s:%s: filter %p ready: %s", SRCNAME, __func__, 
             filter, gpg_strerror (err));

  if (!err)
    release_filter (filter);
  return err;
}


/* Cancel FILTER. */
void
engine_cancel (engine_filter_t filter)
{
  void *cancel_data;

  if (!filter)
    return;
  
  /* First we need to cancel a possible prepared encrypt operation. */
  engine_encrypt_start (filter, 1);

  take_in_lock (filter, __func__);
  cancel_data = filter->cancel_data;
  filter->cancel_data = NULL;
  filter->in.ready = 1;
  release_in_lock (filter, __func__);
  if (cancel_data)
    {
      log_debug ("%s:%s: filter %p: sending cancel command to backend",
                 SRCNAME, __func__, filter);
      if (filter->use_assuan)
        engine_assuan_cancel (cancel_data);
      else
        engine_gpgme_cancel (cancel_data);
      if (WaitForSingleObject (filter->in.ready_event, INFINITE)
          != WAIT_OBJECT_0)
        log_error_w32 (-1, "%s:%s: WFSO failed", SRCNAME, __func__);
      else
        log_debug ("%s:%s: filter %p: backend has been canceled", 
                   SRCNAME, __func__,  filter);
    }
  log_debug ("%s:%s: filter %p: canceled", SRCNAME, __func__, filter);
  release_filter (filter);
}



/* Start an encryption operation to all RECIPIENTS using PROTOCOL
   RECIPIENTS is a NULL terminated array of rfc2822 addresses.  FILTER
   is an object created by engine_create_filter.  The caller needs to
   call engine_wait to finish the operation.  A filter object may not
   be reused after having been used through this function.  However,
   the lifetime of the filter object lasts until the final engine_wait
   or engine_cancel.  On return the protocol to be used is stored at
   R_PROTOCOL.  This is a two part fucntion.  engine_encrypt_prepare
   needs to be called first followed by engine_encrypt_start.  The
   latter command has just one argument CANCEL which can be set to
   true to cancel the prepared command.  */
int
engine_encrypt_prepare (engine_filter_t filter, HWND hwnd,
                        protocol_t req_protocol, char **recipients,
                        protocol_t *r_protocol)
{
  gpg_error_t err;
  protocol_t used_protocol;

  *r_protocol = req_protocol;
  if (filter->use_assuan)
    {
      err = op_assuan_encrypt (req_protocol, filter->indata, filter->outdata,
                               filter, hwnd, recipients, &used_protocol,
                               &filter->encstate);
      if (!err)
        *r_protocol = used_protocol;
    }
  else
    err = op_gpgme_encrypt (req_protocol, filter->indata, filter->outdata,
                            filter, hwnd, recipients);
      
  return err;
}

/* See engine_encrypt_prepare.  */
int
engine_encrypt_start (engine_filter_t filter, int cancel)
{
  gpg_error_t err;

  if (filter->use_assuan)
    {
      err = op_assuan_encrypt_bottom (filter->encstate, cancel);
      filter->encstate = NULL;
    }
  else
    err = 0; /* This is a dummy here.  */
      
  return err;
}


/* Start an detached signing operation.  FILTER is an object created
   by engine_create_filter.  The caller needs to call engine_wait to
   finish the operation.  A filter object may not be reused after
   having been used through this function.  However, the lifetime of
   the filter object lasts until the final engine_wait or
   engine_cancel.  SENDER is the sender's mailbox or NULL.  On return
   the protocol to be used is stored at R_PROTOCOL. */
int
engine_sign_start (engine_filter_t filter, HWND hwnd, protocol_t protocol,
                   const char *sender, protocol_t *r_protocol)
{
  gpg_error_t err;
  protocol_t used_protocol;

  if (filter->use_assuan)
    {
      err = op_assuan_sign (protocol, filter->indata, filter->outdata,
                            filter, hwnd, sender, &used_protocol);
      if (!err)
        *r_protocol = used_protocol;
    }
  else
    {
      err = op_gpgme_sign (protocol, filter->indata, filter->outdata,
                           filter, hwnd);
      if (!err)
        *r_protocol = (protocol == GPGME_PROTOCOL_UNKNOWN?
                       GPGME_PROTOCOL_OpenPGP : protocol);
    }
  return err;
}


/* Start an decrypt operation.  FILTER is an object created by
   engine_create_filter.  The caller needs to call engine_wait to
   finish the operation.  A filter object may not be reused after
   having been used through this function.  However, the lifetime of
   the filter object lasts until the final engine_wait or
   engine_cancel.  */
int
engine_decrypt_start (engine_filter_t filter, HWND hwnd, protocol_t protocol,
                      int with_verify)
{
  gpg_error_t err;

  if (filter->use_assuan)
    err = op_assuan_decrypt (protocol, filter->indata, filter->outdata,
                            filter, hwnd, with_verify);
  else
    err = op_gpgme_decrypt (protocol, filter->indata, filter->outdata,
                            filter, hwnd, with_verify);
  return err;
}


/* Start a verify operation.  FILTER is an object created by
   engine_create_filter; an output function is not required. SIGNATURE
   is the detached signature or NULL if FILTER delivers an opaque
   signature.  The caller needs to call engine_wait to finish the
   operation.  A filter object may not be reused after having been
   used through this function.  However, the lifetime of the filter
   object lasts until the final engine_wait or engine_cancel.  */
int
engine_verify_start (engine_filter_t filter, HWND hwnd, const char *signature,
		     size_t sig_len, protocol_t protocol)
{
  gpg_error_t err;

  if (!signature && !filter->use_assuan)
    {
      log_error ("%s:%s: opaque signatures are not supported "
                 "by the internal backend\n",
                 SRCNAME, __func__);
      return gpg_error (GPG_ERR_NOT_SUPPORTED);
    }

  if (filter->use_assuan && !signature)
    err = op_assuan_verify (protocol, filter->indata, NULL, 0,
			    filter->outdata, filter, hwnd);
  else if (filter->use_assuan)
    err = op_assuan_verify (protocol, filter->indata, signature, sig_len,
			    NULL, filter, hwnd);
  else
    err = op_gpgme_verify (protocol, filter->indata, signature, sig_len,
                           filter, hwnd);
  return err;
}


/* Fire up the key manager.  Returns 0 on success.  */
int
engine_start_keymanager (HWND hwnd)
{
  if (use_assuan)
    return op_assuan_start_keymanager (hwnd);
  else
    return gpg_error (GPG_ERR_NOT_SUPPORTED);
}
