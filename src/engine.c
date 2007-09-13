/* engine.c - Crypto engine dispatcher
 *	Copyright (C) 2007 g10 Code GmbH
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
 * You should have received a copy of the GNU Lesser General Public
 * License along with GpgOL; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#include <config.h>

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <errno.h>
#include <assert.h>
/*#define WIN32_LEAN_AND_MEAN  uncomment it after remove LPSTREAM*/
#include <windows.h>
#include <objidl.h> /* For LPSTREAM in engine-gpgme.h   FIXME: Remove it.  */

#include "common.h"
#include "engine.h"
#include "engine-gpgme.h"
#include "engine-assuan.h"

#define FILTER_BUFFER_SIZE 128  /* FIXME: Increase it after testing  */


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                       SRCNAME, __func__, __LINE__); \
                        } while (0)


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
  struct {
    CRITICAL_SECTION lock; /* The lock for the this object. */
    HANDLE condvar;        /* Manual reset event signaled if LENGTH > 0.  */
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
    size_t length;         /* Number of bytes in BUFFER waiting to be
                              send back to the caller.  */
    char buffer[FILTER_BUFFER_SIZE];
  } out;

  /* The data sink as set by engine_create_filter.  */
  int (*outfnc) (void *, const void *, size_t);
  void *outfncdata;

  /* Objects to be released by engine_wait/cancel.  */
  struct gpgme_data_cbs cb_inbound;  /* Callback structure for gpgme.  */
  struct gpgme_data_cbs cb_outbound; /* Ditto.  */
  gpgme_data_t indata;               /* Input data.  */
  gpgme_data_t outdata;              /* Output data.  */
};




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
  TRACEPOINT ();

  return filter;
}


static void
release_filter (engine_filter_t filter)
{
  if (filter)
    {
      TRACEPOINT ();
      if (filter->in.condvar)
        CloseHandle (filter->in.condvar);
      if (filter->out.condvar)
        CloseHandle (filter->out.condvar);
      if (filter->in.ready_event)
        CloseHandle (filter->in.ready_event);
      gpgme_data_release (filter->indata);
      gpgme_data_release (filter->outdata);
      xfree (filter);
    }
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

  log_debug ("%s:%s: enter\n",  SRCNAME, __func__);
  EnterCriticalSection (&filter->in.lock);
  while (!filter->in.length)
    {
      if (filter->in.got_eof || filter->in.ready)
        {
          LeaveCriticalSection (&filter->in.lock);
          log_debug ("%s:%s: returning EOF\n", SRCNAME, __func__);
          return 0; /* Return EOF. */
        }
      LeaveCriticalSection (&filter->in.lock);
      log_debug ("%s:%s: waiting for in.condvar\n", SRCNAME, __func__);
      WaitForSingleObject (filter->in.condvar, 500);
      EnterCriticalSection (&filter->in.lock);
      log_debug ("%s:%s: continuing\n", SRCNAME, __func__);
    }
     
  log_debug ("%s:%s: requested read size=%d (filter.in.length=%d)\n",
             SRCNAME, __func__, (int)size, (int)filter->in.length);
  nbytes = size < filter->in.length ? size : filter->in.length;
  memcpy (buffer, filter->in.buffer, nbytes);
  if (filter->in.length > nbytes)
    memmove (filter->in.buffer, filter->in.buffer + nbytes, 
             filter->in.length - nbytes);
  filter->in.length -= nbytes;
  LeaveCriticalSection (&filter->in.lock);

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

  log_debug ("%s:%s: enter\n",  SRCNAME, __func__);
  EnterCriticalSection (&filter->out.lock);
  while (filter->out.length)
    {
      LeaveCriticalSection (&filter->out.lock);
      log_debug ("%s:%s: waiting for out.condvar\n", SRCNAME, __func__);
      WaitForSingleObject (filter->out.condvar, 500);
      EnterCriticalSection (&filter->out.lock);
      log_debug ("%s:%s: continuing\n", SRCNAME, __func__);
    }

  log_debug ("%s:%s: requested write size=%d\n",
             SRCNAME, __func__, (int)size);
  nbytes = size < FILTER_BUFFER_SIZE ? size : FILTER_BUFFER_SIZE;
  memcpy (filter->out.buffer, buffer, nbytes);
  filter->out.length = nbytes;
  LeaveCriticalSection (&filter->out.lock);

  log_debug ("%s:%s: write; result=%d\n", SRCNAME, __func__, (int)nbytes);
  return nbytes;
}

/* This function is called by the gpgme backend to notify a filter
   object about the final status of an operation.  It may not be
   called by the engine-gpgme.c module. */
void
engine_gpgme_finished (engine_filter_t filter, gpg_error_t status)
{
  if (!filter)
    {
      log_debug ("%s:%s: called without argument\n", SRCNAME, __func__);
      return;
    }
  log_debug ("%s:%s: filter %p: process terminated: %s <%s>\n", 
             SRCNAME, __func__, filter, 
             gpg_strerror (status), gpg_strsource (status));

  EnterCriticalSection (&filter->in.lock);
  if (filter->in.ready)
    log_debug ("%s:%s: filter %p: Oops: already flagged as finished\n", 
               SRCNAME, __func__, filter);
  filter->in.ready = 1;
  filter->in.status = status;
  if (!SetEvent (filter->in.ready_event))
    log_error_w32 (-1, "%s:%s: SetEvent failed", SRCNAME, __func__);
  LeaveCriticalSection (&filter->in.lock);
  log_debug ("%s:%s: leaving\n", SRCNAME, __func__);
}





/* Initialize the engine dispatcher.  */
int
engine_init (void)
{
  op_gpgme_init ();
  return 0;
}


/* Shutdown the engine dispatcher.  */
void
engine_deinit (void)
{
  op_gpgme_deinit ();

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
  size_t nbytes;

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

  log_debug ("%s:%s: indata=%p indatalen=%d outfnc=%p\n",
             SRCNAME, __func__, indata, (int)indatalen, filter->outfnc); 
  for (;;)
    {
      /* If there is something to write out, do this now to make space
         for more data.  */
      EnterCriticalSection (&filter->out.lock);
      while (filter->out.length)
        {
          TRACEPOINT ();
          nbytes = filter->outfnc (filter->outfncdata, 
                                   filter->out.buffer, filter->out.length);
          if (nbytes == -1)
            {
              log_debug ("%s:%s: error writing data\n", SRCNAME, __func__);
              LeaveCriticalSection (&filter->out.lock);
              return gpg_error (GPG_ERR_EIO);
            }
          assert (nbytes < filter->out.length && nbytes >= 0);
          if (nbytes < filter->out.length)
            memmove (filter->out.buffer, filter->out.buffer + nbytes,
                     filter->out.length - nbytes); 
          filter->out.length =- nbytes;
        }
      if (!PulseEvent (filter->out.condvar))
        log_error_w32 (-1, "%s:%s: PulseEvent(out) failed", SRCNAME, __func__);
      LeaveCriticalSection (&filter->out.lock);
      
      EnterCriticalSection (&filter->in.lock);
      if (!indata && !indatalen)
        {
          TRACEPOINT ();
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
          LeaveCriticalSection (&filter->in.lock);
          return err; 
        }

      /* Fill the input buffer, relinquish control to the callback
         processor and loop until all input data has been
         processed.  */
      if (!filter->in.length && indatalen)
        {
          TRACEPOINT ();
          filter->in.length = (indatalen > FILTER_BUFFER_SIZE
                               ? FILTER_BUFFER_SIZE : indatalen);
          memcpy (filter->in.buffer, indata, filter->in.length);
          indata    += filter->in.length;
          indatalen -= filter->in.length;
        }
      if (!filter->in.length || (filter->in.ready && !filter->out.length))
        {
          LeaveCriticalSection (&filter->in.lock);
          err = 0;
          break;  /* the loop.  */
        }
      if (!PulseEvent (filter->in.condvar))
        log_error_w32 (-1, "%s:%s: PulseEvent(in) failed", SRCNAME, __func__);
      LeaveCriticalSection (&filter->in.lock);
      Sleep (0);
    }

  log_debug ("%s:%s: leave; err=%d\n", SRCNAME, __func__, err); 
  return err;
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
  filter->outfnc = outfnc;
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



/* Wait for FILTER to finish.  Returns 0 on success.  FILTER is not
   valid after the function has returned success.  */
int
engine_wait (engine_filter_t filter)
{
  gpg_error_t err;
  int more;

  TRACEPOINT ();
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
      EnterCriticalSection (&filter->out.lock);
      while (filter->out.length)
        {
          int nbytes; 
          TRACEPOINT ();
          nbytes = filter->outfnc (filter->outfncdata, 
                                   filter->out.buffer, filter->out.length);
          if (nbytes == -1)
            {
              log_debug ("%s:%s: error writing data\n", SRCNAME, __func__);
              LeaveCriticalSection (&filter->out.lock);
              break;
            }
          assert (nbytes < filter->out.length && nbytes >= 0);
          if (nbytes < filter->out.length)
            memmove (filter->out.buffer, filter->out.buffer + nbytes,
                     filter->out.length - nbytes); 
          filter->out.length =- nbytes;
        }
      if (!PulseEvent (filter->out.condvar))
        log_error_w32 (-1, "%s:%s: PulseEvent(out) failed", SRCNAME, __func__);
      LeaveCriticalSection (&filter->out.lock);
      EnterCriticalSection (&filter->in.lock);
      more = !filter->in.ready;
      LeaveCriticalSection (&filter->in.lock);
      if (more)
        Sleep (0);
    }
  while (more);
  
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
  if (!filter)
    return;
  
  EnterCriticalSection (&filter->in.lock);
  filter->in.ready = 1;
  LeaveCriticalSection (&filter->in.lock);
  log_debug ("%s:%s: filter %p canceled", SRCNAME, __func__, filter);
  /* FIXME:  Here we need to kill the underlying gpgme process. */
  release_filter (filter);
}



/* Start an encryption operation to all RECIPEINTS using PROTOCOL
   RECIPIENTS is a NULL terminated array of rfc2822 addresses.  FILTER
   is an object created by engine_create_filter.  The caller needs to
   call engine_wait to finish the operation.  A filter object may not
   be reused after having been used through this function.  However,
   the lifetime of the filter object lasts until the final engine_wait
   or engine_cancel.  */
int
engine_encrypt_start (engine_filter_t filter,
                      protocol_t protocol, char **recipients)
{
  gpg_error_t err;

  err = op_gpgme_encrypt_data (protocol, filter->indata, filter->outdata,
                               filter, recipients, NULL, 0);
  return err;
}


/* Start an detached signing operation.
   FILTER
   is an object created by engine_create_filter.  The caller needs to
   call engine_wait to finish the operation.  A filter object may not
   be reused after having been used through this function.  However,
   the lifetime of the filter object lasts until the final engine_wait
   or engine_cancel.  */
int
engine_sign_start (engine_filter_t filter, protocol_t protocol)
{
  gpg_error_t err;

  err = op_gpgme_sign_data (protocol, filter->indata, filter->outdata,
                            filter);
  return err;
}





