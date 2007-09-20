/* engine-assuan.c - Crypto engine using an Assuan server
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
#define WIN32_LEAN_AND_MEAN 
#include <windows.h>

#include "common.h"
#include "engine.h"
#include "engine-assuan.h"


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                       SRCNAME, __func__, __LINE__); \
                        } while (0)


/* Because we are using asynchronous gpgme commands, we need to have a
   closure to cleanup allocated resources and run the code required
   adfter gpgme finished the command (e.g. getting the signature
   verification result.  Thus all functions need to implement a
   closure function and pass it using a closure_data_t object via the
   gpgme_progress_cb hack.  */
struct closure_data_s;
typedef struct closure_data_s *closure_data_t;
struct closure_data_s
{
  void (*closure)(closure_data_t, gpgme_ctx_t, gpg_error_t);
  engine_filter_t filter;
  struct passphrase_cb_s pw_cb;  /* Passphrase callback info.  */
};


static int init_done = 0;


static DWORD WINAPI pipe_worker_thread (void *dummy);


static void
cleanup (void)
{
  /* Fixme: We should stop the thread.  */
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
  gpgme_error_t err;

  if (init_done)
    return 0;

  /* FIXME: Connect to the server and return failure if it is not
     possible.  */
  return gpg_error (GPG_ERR_NOT_IMPLEMENTED);

  /* Fire up the pipe worker thread. */
  {
    HANDLE th;
    DWORD tid;

    th = CreateThread (NULL, 128*1024, pipe_worker_thread, NULL, 0, &tid);
    if (th == INVALID_HANDLE_VALUE)
      log_error ("failed to start the piper worker thread\n");
    else
      CloseHandle (th);
  }

  init_done = 1; 
  return 0;
}


/* The worker thread which feeds the pipes.  */
static DWORD WINAPI
pipe_worker_thread (void *dummy)
{
  gpgme_ctx_t ctx;
  gpg_error_t err;
  void *a_voidptr;
  closure_data_t closure_data;

  (void)dummy;

  for (;;)
    {
      Sleep (1000);
    }
}








/* Not that this closure is called in the context of the
   waiter_thread.  */
static void
encrypt_closure (closure_data_t cld, gpg_error_t err)
{
  engine_private_finished (cld->filter, err);
}


/* Encrypt the data from INDATA to the OUTDATA object for all
   recpients given in the NULL terminated array KEYS.  If SIGN_KEY is
   not NULL the message will also be signed.  On termination of the
   encryption command engine_gpgme_finished() is called with
   NOTIFY_DATA as the first argument.  

   This global function is used to avoid allocating an extra context
   just for this notification.  We abuse the gpgme_set_progress_cb
   value for storing the pointer with the gpgme context.  */
int
op_assuan_encrypt (protocol_t protocol, 
                   gpgme_data_t indata, gpgme_data_t outdata,
                   void *notify_data, /* FIXME: Add hwnd */
                   char **recipients)
{
  gpg_error_t err;
  closure_data_t cld;

  cld = xcalloc (1, sizeof *cld);
  cld->closure = encrypt_closure;
  cld->filter = notify_data;


  /* FIXME:  We should not hardcode always trust. */
/*   if (sign_key) */
/*     { */
/*       gpgme_set_passphrase_cb (ctx, passphrase_callback_box, &cld->pw_cb); */
/*       cld->pw_cb.ctx = ctx; */
/*       cld->pw_cb.ttl = ttl; */
/*       err = gpgme_signers_add (ctx, sign_key); */
/*       if (!err) */
/*         err = gpgme_op_encrypt_sign_start (ctx, keys, */
/*                                            GPGME_ENCRYPT_ALWAYS_TRUST, */
/*                                            indata, outdata); */
/*     } */
/*   else */
/*     err = gpgme_op_encrypt_start (ctx, keys, GPGME_ENCRYPT_ALWAYS_TRUST,  */
/*                                   indata, outdata); */

  err = -1;

 leave:
  if (err)
    {
      xfree (cld);
    }
  return err;
}



