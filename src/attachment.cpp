/* attachment.cpp - Functions for attachment handling
 *    Copyright (C) 2005, 2007 g10 Code GmbH
 *    Copyright (C) 2015 Intevation GmbH
 *
 * This file is part of GpgOL.
 *
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#include "config.h"
#include "common.h"
#include "attachment.h"
#include "serpent.h"
#include "oomhelp.h"
#include "mymapitags.h"
#include "mapihelp.h"

#include <gpg-error.h>

#define COPYBUFFERSIZE 4096

#define IV_DEFAULT_LEN 16

/** Decrypt the first 16 bytes of stream and check that it contains
  our header. Return 0 on success. */
static int
check_header (LPSTREAM stream, symenc_t symenc)
{
  HRESULT hr;
  char tmpbuf[16];
  ULONG nread;
  hr = stream->Read (tmpbuf, 16, &nread);
  if (hr || nread != 16)
    {
      log_error ("%s:%s: Read failed: hr=%#lx", SRCNAME, __func__, hr);
      return -1;
    }
  symenc_cfb_decrypt (symenc, tmpbuf, tmpbuf, 16);
  if (memcmp (tmpbuf, "GpgOL attachment", 16))
    {
      log_error ("%s:%s: Invalid header.",
                 SRCNAME, __func__);
      char buf2 [17];
      snprintf (buf2, 17, "%s", tmpbuf);
      log_error("Buf2: %s", buf2);
      return -1;
    }
  return 0;
}

/** Encrypts or decrypts a stream in place using the symenc context.
  Returns 0 on success. */
static int
do_crypt_stream (LPSTREAM stream, symenc_t symenc, bool encrypt)
{
  char *buf = NULL;
  HRESULT hr;
  ULONG nread;
  bool fixed_str_written = false;
  int rc = -1;
  ULONG written = 0;
  /* The original intention was to use IStream::Clone to have
     an independent read / write stream. But the MAPI attachment
     stream returns E_NOT_IMPLMENTED for that :-)
     So we manually track the read and writepos. Read is offset
     at 16 because of the GpgOL message. */
  LARGE_INTEGER readpos = {0, 0},
                writepos = {0, 0};
  ULARGE_INTEGER new_size = {0, 0};

  if (!encrypt)
    {
      readpos.QuadPart = 16;
    }

  buf = (char*)xmalloc (COPYBUFFERSIZE);
  do
    {
      hr = stream->Read (buf, COPYBUFFERSIZE, &nread);
      if (hr)
        {
          log_error ("%s:%s: Read failed: hr=%#lx", SRCNAME, __func__, hr);
          goto done;
        }
      if (!nread)
        {
          break;
        }
      readpos.QuadPart += nread;
      stream->Seek(writepos, STREAM_SEEK_SET, NULL);
      if (nread && encrypt && !fixed_str_written)
        {
          char tmpbuf[16];
          /* Write an encrypted fixed 16 byte string which we need to
             check at decryption time to see whether we have actually
             encrypted it using this session key.  */
          symenc_cfb_encrypt (symenc, tmpbuf, "GpgOL attachment", 16);
          stream->Write (tmpbuf, 16, NULL);
          fixed_str_written = true;
          writepos.QuadPart = 16;
        }
      if (encrypt)
        {
          symenc_cfb_encrypt (symenc, buf, buf, nread);
        }
      else
        {
          symenc_cfb_decrypt (symenc, buf, buf, nread);
        }

        hr = stream->Write (buf, nread, &written);
        if (FAILED (hr) || written != nread)
          {
            log_error ("%s:%s: Write failed: %i", SRCNAME, __func__, __LINE__);
            goto done;
          }
        writepos.QuadPart += written;
        stream->Seek(readpos, STREAM_SEEK_SET, NULL);
      }
    while (nread == COPYBUFFERSIZE);

  new_size.QuadPart = writepos.QuadPart;
  hr = stream->SetSize (new_size);
  if (FAILED (hr))
    {
      log_error ("%s:%s: Failed to update size", SRCNAME, __func__);
      goto done;
    }
  rc = 0;

done:
  xfree (buf);

  if (rc)
    {
      stream->Revert ();
    }
  else
    {
      stream->Commit (0);
    }

  return rc;
}

/** If encrypt is set to true this will encrypt the attachment
  data with serpent otherwiese it will decrypt.
  This function handles the mapi side of things.
  */
static int
do_crypt_mapi (LPATTACH att, bool encrypt)
{
  char *iv;
  ULONG tag;
  size_t ivlen = IV_DEFAULT_LEN;
  symenc_t symenc = NULL;
  HRESULT hr;
  LPSTREAM stream = NULL;
  int rc = -1;

  if (!att)
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      return -1;
    }

  /* Get or create a new IV */
  if (get_gpgolprotectiv_tag ((LPMESSAGE)att, &tag) )
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      return -1;
    }
  if (encrypt)
    {
      iv = (char*)create_initialization_vector (IV_DEFAULT_LEN);
    }
  else
    {
      iv = mapi_get_binary_prop ((LPMESSAGE)att, tag, &ivlen);
    }
  if (!iv)
    {
      log_error ("%s:%s: Error creating / getting IV: %i", SRCNAME,
                 __func__, __LINE__);
      goto done;
    }

  symenc = symenc_open (get_128bit_session_key (), 16, iv, ivlen);
  xfree (iv);
  if (!symenc)
    {
      log_error ("%s:%s: can't open encryption context", SRCNAME, __func__);
      goto done;
    }

  hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream,
                          0, MAPI_MODIFY, (LPUNKNOWN*) &stream);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open data stream of attachment: hr=%#lx",
                 SRCNAME, __func__, hr);
      goto done;
    }

  /* When decrypting check the first 16 bytes for the header */
  if (!encrypt && check_header (stream, symenc))
    {
      goto done;
    }

  if (FAILED (hr))
    {
      log_error ("%s:%s: can't create temp file: hr=%#lx",
                 SRCNAME, __func__, hr);
      goto done;
    }

  if (do_crypt_stream (stream, symenc, encrypt))
    {
      log_error ("%s:%s: stream handling failed",
                 SRCNAME, __func__);
      goto done;
    }
  rc = 0;

done:
  if (symenc)
    symenc_close (symenc);
  gpgol_release (stream);

  return rc;
}

/** Protect or unprotect attachments.*/
static int
do_crypt (LPDISPATCH mailitem, bool protect)
{
  LPDISPATCH attachments = get_oom_object (mailitem, "Attachments");
  LPMESSAGE message = get_oom_base_message (mailitem);
  int count = 0;
  int err = -1;
  char *item_str;
  int i;
  ULONG tag_id;

  if (!attachments || !message)
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      return -1;
    }
  count = get_oom_int (attachments, "Count");

  if (get_gpgolattachtype_tag (message, &tag_id))
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      goto done;
    }

  if (count < 1)
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      goto done;
    }

  /* Yes the items start at 1! */
  for (i = 1; i <= count; i++)
    {
      LPDISPATCH attachment;
      LPATTACH mapi_attachment;
      attachtype_t att_type;

      if (gpgrt_asprintf (&item_str, "Item(%i)", i) == -1)
        {
          log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
          goto done;
        }

      attachment = get_oom_object (attachments, item_str);
      xfree (item_str);
      if (!attachment)
        {
          log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
        }
      mapi_attachment = (LPATTACH) get_oom_iunknown (attachment,
                                                     "MapiObject");
      if (!mapi_attachment)
        {
          log_debug ("%s:%s: Failed to get MapiObject of attachment: %p",
                     SRCNAME, __func__, attachment);
          attachment->Release ();
          continue;
        }

      att_type = get_gpgolattachtype (mapi_attachment, tag_id);
      if ((protect && att_type == ATTACHTYPE_FROMMOSS_DEC) ||
          (!protect && att_type == ATTACHTYPE_FROMMOSS))
        {
          if (do_crypt_mapi (mapi_attachment, protect))
            {
              log_error ("%s:%s: Error: Session crypto failed.",
                         SRCNAME, __func__);
              mapi_attachment->Release ();
              attachment->Release ();
              goto done;
            }
        }
      mapi_attachment->Release ();
      attachment->Release ();
    }
  err = 0;

done:

  gpgol_release (message);
  gpgol_release (attachments);
  return err;
}

int
protect_attachments (LPDISPATCH mailitem)
{
  return do_crypt (mailitem, true);
}

int
unprotect_attachments (LPDISPATCH mailitem)
{
  return do_crypt (mailitem, false);
}
