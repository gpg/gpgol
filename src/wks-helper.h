/* @file wks-helper.cpp
 * @brief Helper to work with a web-key-service
 *
 * Copyright (C) 2018 Intevation GmbH
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

#include <string>
#include "oomhelp.h"

#include <utility>

class Mail;
namespace GpgME
{
  class Data;
} // namespace GpgME

/** @brief Helper for web key services.
 *
 * Everything is public to make it easy to access data
 * members from another windows thread. Don't mess with them.
 *
 * This is all a bit weird, don't look at it too much as it works ;-)
 */
class WKSHelper
{
protected:
    /** Loads the list of checked keys */
    explicit WKSHelper ();
public:
    enum WKSState
      {
        NotChecked, /*<-- Supported state was not checked */
        NotSupported, /* <-- WKS is not supported for this address */
        NeedsPublish, /* <-- There was no key published for this address */
        IsPublished, /* <-- WKS is supported for this address and published */
        ConfirmationSeen, /* A confirmation request was seen for this mail addres. */
        NeedsUpdate, /* <-- Not yet implemeted. */
        PublishInProgress, /* <-- Publishing is currently in progress. */
        RequestSent, /* <-- A publishing request has been sent. */
        PublishDenied, /* <-- A user denied publishing. */
        ConfirmationSent, /* <-- The confirmation response was sent. */
      };

    ~WKSHelper ();

    /** Get the WKSHelper

        On the initial request:
        Ensure that the OOM is available.
        Will load all account addresses from OOM and then return.

        Starts a background thread to load info from a file
        and run checks if necessary.

        When the thread is finished initialized will be true.
    */
    static const WKSHelper* instance ();

    /** If the key for the address @address should be published */
    WKSState get_state (const std::string &mbox) const;

    /** Start a supported check for a given mbox.

        If force is true the check will be run. Otherwise
        the state will only be updated if the last check
        was more then 7 days ago.

        Returns immediately as the check is run in a background
        thread.
    */
    void start_check (const std::string &mbox, bool force = false) const;

    /** Starts gpg-wks-client --create */
    void start_publish (const std::string &mbox) const;

    /** Allow queueing a notification after a sleepTime */
    void allow_notify (int sleepTimeMS = 0) const;

    /** Send a notification and start publishing accordingly */
    void notify (const char *mbox) const;

    /** Store the current static maps. */
    void save () const;

    /** Update or insert a state in the static maps. */
    void update_state (const std::string &mbox, WKSState state, bool save = true) const;

    /** Update or insert last_checked in the static maps. */
    void update_last_checked (const std::string &mbox, time_t last_checked,
                              bool save = true) const;

    /** Create / Build / Send Mail
      returns 0 on success.
    */
    int send_mail (const std::string &mimeData) const;

    /** Handle a confirmation mail read event */
    void handle_confirmation_read (Mail *mail, LPSTREAM msgstream) const;

    /** Handle the notifcation following the read. */
    void handle_confirmation_notify (const std::string &mbox) const;

    /** Get the cached confirmation data. Caller takes ownership of
      the data object and has to delete it. It is removed from the cache. */
    std::pair <GpgME::Data *, Mail *> get_cached_confirmation (const std::string &mbox) const;
private:
    time_t get_check_time (const std::string &mbox) const;

    void load() const;
};
