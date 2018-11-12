/* @file categorymanager.cpp
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

#include "categorymanager.h"
#include "common.h"
#include "mail.h"
#include "gpgoladdin.h"
#include "oomhelp.h"

#include <unordered_map>

#undef _
#define _(a) utf8_gettext (a)

class CategoryManager::Private
{
public:
  Private()
  {
  }

  void createCategory (shared_disp_t store,
                       const std::string &category, int color)
    {
      TSTART;
      LPDISPATCH categories = get_oom_object (store.get(), "Categories");
      if (!categories)
        {
          STRANGEPOINT;
          TRETURN;
        }
      if (create_category (categories, category.c_str (), color))
        {
          log_debug ("%s:%s: Failed to create category %s",
                     SRCNAME, __func__, anonstr (category.c_str()));
          gpgol_release (categories);
          TRETURN;
        }
      gpgol_release (categories);
      TRETURN;
    }

  void registerCategory (const std::string &storeID,
                         const std::string &category)
    {
      TSTART;
      auto storeIt = mCategoryStoreMap.find (storeID);
      if (storeIt == mCategoryStoreMap.end())
        {
          /* First category for this store. Create a new
             category ref map. */
          std::unordered_map <std::string, int> categoryMap;
          categoryMap.insert (std::make_pair (category, 1));
          mCategoryStoreMap.insert (std::make_pair (storeID, categoryMap));
          log_debug ("%s:%s: Register category %s in new store %s ref now 1",
                     SRCNAME, __func__, anonstr (category.c_str()),
                     anonstr (storeID.c_str()));
          TRETURN;
        }
      auto categoryIt = storeIt->second.find (category);
      if (categoryIt == storeIt->second.end ())
        {
          storeIt->second.insert (std::make_pair (category, 1));
          log_debug ("%s:%s: Register category %s in store %s ref now 1",
                     SRCNAME, __func__, anonstr (category.c_str()),
                     anonstr (storeID.c_str()));
        }
      else
        {
          categoryIt->second++;
          log_debug ("%s:%s: Register category %s in store %s ref now %i",
                     SRCNAME, __func__, anonstr (category.c_str()),
                     anonstr (storeID.c_str()), categoryIt->second);
        }
      TRETURN;
    }

  void unregisterCategory (const std::string &storeID,
                           const std::string &category)
    {
      TSTART;
      auto storeIt = mCategoryStoreMap.find (storeID);
      if (storeIt == mCategoryStoreMap.end ())
        {
          log_error ("%s:%s: Unregister called for unregistered store %s",
                     SRCNAME, __func__, anonstr (storeID.c_str()));
          TRETURN;
        }
      auto categoryIt = storeIt->second.find (category);
      if (categoryIt == storeIt->second.end ())
        {
          log_debug ("%s:%s: Unregister %s not found for store %s",
                     SRCNAME, __func__, anonstr (category.c_str()),
                     anonstr (storeID.c_str()));
          TRETURN;
        }
      categoryIt->second--;
      log_debug ("%s:%s: Unregister category %s in store %s ref now %i",
                 SRCNAME, __func__, anonstr (category.c_str()),
                 anonstr (storeID.c_str()), categoryIt->second);
      if (categoryIt->second < 0)
        {
          log_debug ("%s:%s: Unregister %s negative for store %s",
                     SRCNAME, __func__, anonstr (category.c_str()),
                     anonstr (storeID.c_str()));
          TRETURN;
        }
      if (categoryIt->second == 0)
        {
          log_debug ("%s:%s: Deleting %s for store %s",
                     SRCNAME, __func__, anonstr (category.c_str()),
                     anonstr (storeID.c_str()));

          LPDISPATCH store = get_store_for_id (storeID.c_str());
          if (!store)
            {
              STRANGEPOINT;
              TRETURN;
            }
          delete_category (store, category.c_str ());
          gpgol_release (store);
          storeIt->second.erase (categoryIt);
        }
      TRETURN;
    }

  bool categoryExistsInMap (const std::string &storeID,
                            const std::string &category)
    {
      const auto it = mCategoryStoreMap.find (storeID);
      if (it == mCategoryStoreMap.end ())
        {
          return false;
        }
      return it->second.find (category) != it->second.end();
    }

private:
  /* Map from: store to map of category -> refs. */
  std::unordered_map<std::string,
    std::unordered_map<std::string, int> > mCategoryStoreMap;
};

/* static */
std::shared_ptr <CategoryManager>
CategoryManager::instance ()
{
  return GpgolAddin::get_instance ()->get_category_mngr ();
}

CategoryManager::CategoryManager():
  d(new Private)
{
}

std::string
CategoryManager::addCategoryToMail (Mail *mail, const std::string &category, int color)
{
  TSTART;
  std::string ret;
  if (!mail || category.empty())
    {
      TRETURN ret;
    }

  auto store = MAKE_SHARED (get_oom_object (mail->item (), "Parent.Store"));
  if (!store)
    {
      log_error ("%s:%s Failed to obtain store",
                 SRCNAME, __func__);
      TRETURN std::string ();
    }
  char *storeID = get_oom_string (store.get (), "StoreID");
  if (!storeID)
    {
      log_error ("%s:%s Failed to obtain storeID",
                 SRCNAME, __func__);
      TRETURN std::string ();
    }
  ret = storeID;
  xfree (storeID);

  if (!d->categoryExistsInMap (ret, category))
    {
      d->createCategory (store, category, color);
    }
  d->registerCategory (ret, category);

  if (add_category (mail->item (), category.c_str()))
    {
      /* Probably the category already existed there
         so it is not super worrysome. */
      log_debug ("%s:%s Failed to add category.",
                 SRCNAME, __func__);
    }
  return ret;
}

void
CategoryManager::removeCategory (Mail *mail, const std::string &category)
{
  TSTART;
  if (!mail || category.empty())
    {
      STRANGEPOINT;
      TRETURN;
    }
  if (remove_category (mail->item (), category.c_str (), true))
    {
      log_debug ("%s:%s Failed to remvoe category.",
                 SRCNAME, __func__);
    }
  d->unregisterCategory (mail->storeID (), category.c_str ());

  TRETURN;
}

/* static */
void
CategoryManager::removeAllGpgOLCategories ()
{
  TSTART;
  delete_all_categories_starting_with ("GpgOL: ");
  TRETURN;
}

/* static */
const std::string &
CategoryManager::getEncMailCategory ()
{
  static std::string decStr;
  if (decStr.empty())
    {
      decStr = std::string ("GpgOL: ") +
                            std::string (_("Encrypted Message"));
    }
  return decStr;
}

/* static */
const std::string &
CategoryManager::getJunkMailCategory ()
{
  static std::string decStr;
  if (decStr.empty())
    {
      decStr = std::string ("GpgOL: ") +
                            std::string (_("Junk Email cannot be processed"));
    }
  return decStr;
}
