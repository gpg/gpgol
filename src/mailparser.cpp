/* @file mailparser.cpp
 * @brief Parse a mail and decrypt / verify accordingly
 *
 *    Copyright (C) 2016 Intevation GmbH
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

#include "mailparser.h"
#include "attachment.h"

MailParser::MailParser(LPSTREAM instream, msgtype_t type):
    m_body (std::shared_ptr<std::string>(new std::string())),
    m_htmlbody (std::shared_ptr<std::string>(new std::string())),
    m_stream (instream),
    m_type (type),
    m_error (false),
    m_in_data (false),
    signed_data (nullptr),
    sig_data (nullptr)
{
  log_debug ("%s:%s: Creating parser for stream: %p",
             SRCNAME, __func__, instream);
  instream->AddRef();
}

MailParser::~MailParser()
{
  gpgol_release(m_stream);
}

std::string
MailParser::parse()
{
  m_body = std::shared_ptr<std::string>(new std::string("Hello world"));
  Attachment *att = new Attachment ();
  att->write ("Hello attachment", strlen ("Hello attachment"));
  att->set_display_name ("The Attachment.txt");
  m_attachments.push_back (std::shared_ptr<Attachment>(att));
  return std::string();
}

std::shared_ptr<std::string> MailParser::get_utf8_html_body()
{
  return m_htmlbody;
}

std::shared_ptr<std::string> MailParser::get_utf8_text_body()
{
  return m_body;
}

std::vector<std::shared_ptr<Attachment> > MailParser::get_attachments()
{
  return m_attachments;
}
