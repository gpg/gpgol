#ifndef CATEGORYMANAGER_H
#define CATEGORYMANAGER_H

/* @file categorymanager.h
 * @brief Handles category management
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

#include <memory>
#include <string>

class Mail;
class GpgolAddin;

/* The category manager is supposed to be only accessed from
   the main thread and is not guarded by locks. */
class CategoryManager
{
    friend class GpgolAddin;

protected:
    /** Internal ctor */
    explicit CategoryManager ();

public:
    /** Get the CategoryManager */
    static std::shared_ptr<CategoryManager> instance ();

    /** Get the Category seperator from the registry. */
    static const std::string& getSeperator ();

    /** Add a category to a mail.

      @returns the storeID of the mail / category.
    */
    std::string addCategoryToMail (Mail *mail, const std::string &category,
                                   int color);

    /** Remove the category @category */
    void removeCategory (Mail * mail,
                         const std::string &category);

    /** Remove all GpgOL categories from all stores. */
    static void removeAllGpgOLCategories ();

    /** Get the name of the encryption category. */
    static const std::string & getEncMailCategory ();

    /** Get the name of the junk category. */
    static const std::string & getJunkMailCategory ();
private:
    class Private;
    std::shared_ptr<Private> d;
};

#endif
