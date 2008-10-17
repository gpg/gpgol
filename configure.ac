# configure.ac - for GpgOL
# Copyright (C) 2005, 2006, 2007, 2008 g10 Code GmbH
#
# This file is free software; as a special exception the author gives
# unlimited permission to copy and/or distribute it, with or without
# modifications, as long as this notice is preserved.
#
# This program is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY, to the extent permitted by law; without even the
# implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

# Process this file with autoconf to produce a configure script.
AC_PREREQ(2.59)
min_automake_version="1.9.4"

# Remember to change the version number immediately *after* a release.
# Set my_issvn to "yes" for non-released code.  Remember to run an
# "svn up" and "autogen.sh" right before creating a distribution.
m4_define([my_version], [0.10.16])
m4_define([my_issvn], [yes])

m4_define([svn_revision], m4_esyscmd([echo -n $( (svn info 2>/dev/null \
            || echo 'Revision: 0')|sed -n '/^Revision:/ {s/[^0-9]//gp;q;}')]))
AC_INIT([gpgol], 
        [my_version[]m4_if(my_issvn,[yes],[-svn[]svn_revision])],
        [bug-gpgol@g10code.com])

NEED_GPG_ERROR_VERSION=1.4
NEED_GPGME_API=1
NEED_GPGME_VERSION=1.1.0
NEED_LIBASSUAN_API=1
NEED_LIBASSUAN_VERSION=1.0.4


SVN_REVISION=svn_revision
PACKAGE=$PACKAGE_NAME
PACKAGE_GT=${PACKAGE_NAME}
VERSION=$PACKAGE_VERSION

AC_CONFIG_SRCDIR(src/gpgol.def)
AM_CONFIG_HEADER(config.h)
AC_CANONICAL_TARGET()
AM_INIT_AUTOMAKE

AC_GNU_SOURCE

have_gpg_error=no
have_libassuan=no

AC_SUBST(PACKAGE)
AC_SUBST(PACKAGE_GT)
AC_SUBST(VERSION)
AC_DEFINE_UNQUOTED(PACKAGE, "$PACKAGE", [Name of this package])
AC_DEFINE_UNQUOTED(PACKAGE_GT, "$PACKAGE_GT",
                                [Name of this package for gettext])
AC_DEFINE_UNQUOTED(VERSION, "$VERSION", [Version of this package])
AC_DEFINE_UNQUOTED(PACKAGE_BUGREPORT, "$PACKAGE_BUGREPORT",
                                        [Bug report address])
AC_DEFINE_UNQUOTED(NEED_GPGME_VERSION, "$NEED_GPGME_VERSION",
                                       [Required version of GPGME])
AC_DEFINE_UNQUOTED(SVN_REVISION, ${SVN_REVISION}, [Current SVN revision])


BUILD_TIMESTAMP=`date --iso-8601=minutes`
AC_SUBST(BUILD_TIMESTAMP)
changequote(,)dnl
BUILD_FILEVERSION=`echo "$VERSION" | sed 's/\([0-9.]*\).*/\1./;s/\./,/g'`
changequote([,])dnl
BUILD_FILEVERSION="${BUILD_FILEVERSION}${SVN_REVISION}"
AC_SUBST(BUILD_FILEVERSION)


AH_BOTTOM([
/* Some global constants. */

/* Force using of NLS for W32 even if no libintl has been found.  This is 
   okay because we have our own gettext implementation for W32.  */
#if defined(HAVE_W32_SYSTEM) && !defined(ENABLE_NLS)
#define ENABLE_NLS 1
#endif
])

AM_MAINTAINER_MODE

# Checks for programs.
AC_PROG_MAKE_SET
AM_SANITY_CHECK
missing_dir=`cd $ac_aux_dir && pwd`
AM_MISSING_PROG(ACLOCAL, aclocal, $missing_dir)
AM_MISSING_PROG(AUTOCONF, autoconf, $missing_dir)
AM_MISSING_PROG(AUTOMAKE, automake, $missing_dir)
AM_MISSING_PROG(AUTOHEADER, autoheader, $missing_dir)
AM_MISSING_PROG(MAKEINFO, makeinfo, $missing_dir)
AC_PROG_AWK
AC_PROG_CC
AC_PROG_CPP
AC_PROG_CXX
AC_PROG_INSTALL
AC_PROG_LN_S
AC_PROG_MAKE_SET
AC_PROG_RANLIB
AC_CHECK_TOOL(AR, ar, :)
AC_CHECK_TOOL(DLLTOOL, dlltool, :)
AC_CHECK_TOOL(WINDRES, windres, :)


have_dosish_system=no
have_w32_system=no
case "${host}" in
    *-mingw32*)
        # special stuff for Windoze NT
        ac_cv_have_dev_random=no
        AC_DEFINE(USE_ONLY_8DOT3,1,
                  [set this to limit filenames to the 8.3 format])
        AC_DEFINE(HAVE_DRIVE_LETTERS,1,
                  [defined if we must run on a stupid file system])
        have_dosish_system=yes
        have_w32_system=yes
        ;;
    *)
    AC_MSG_ERROR([[
***
*** This software my only be build for W32 systems.  Use
***     ./autogen.sh --build-w32
*** to prepare it for such a build.
***]])
       ;;
esac

if test "$have_dosish_system" = yes; then
   AC_DEFINE(HAVE_DOSISH_SYSTEM,1,
             [Defined if we run on some of the PCDOS like systems
              (DOS, Windoze. OS/2) with special properties like
              no file modes])
fi
AM_CONDITIONAL(HAVE_DOSISH_SYSTEM, test "$have_dosish_system" = yes)

if test "$have_w32_system" = yes; then
   AC_DEFINE(HAVE_W32_SYSTEM,1, [Defined if we run on a W32 API based system])
fi
AM_CONDITIONAL(HAVE_W32_SYSTEM, test "$have_w32_system" = yes)



#
# Checks for libraries.
#


AM_PATH_GPGME("$NEED_GPGME_API:$NEED_GPGME_VERSION",
               have_gpgme=yes,have_gpgme=no)
AM_PATH_GPG_ERROR("$NEED_GPG_ERROR_VERSION",
                  have_gpg_error=yes,have_gpg_error=no)
AC_DEFINE(GPG_ERR_SOURCE_DEFAULT, GPG_ERR_SOURCE_USER_2,
          [The default error source for GpgOL.])
AM_PATH_LIBASSUAN("$NEED_LIBASSUAN_API:$NEED_LIBASSUAN_VERSION",
                  have_libassuan=yes,have_libassuan=no)

AH_BOTTOM([
/* We don't want the old assuan codes anymore. */
#define _ASSUAN_ONLY_GPG_ERRORS 1
])



# Note, that autogen.sh greps for the next line.
AM_GNU_GETTEXT_VERSION(0.12.1)
AM_GNU_GETTEXT([external])
# We always want NLs, despite what the above macro figures out.
USE_NLS=yes

# Checks for header files.
AC_HEADER_STDC
AC_CHECK_HEADERS(string.h unistd.h langinfo.h termio.h locale.h)

AC_CHECK_FUNCS(stpcpy)

# Checks for typedefs, structures, and compiler characteristics.
AC_C_CONST
AC_C_INLINE
AC_TYPE_SIZE_T
AC_TYPE_SIGNAL

AM_CONDITIONAL(CROSS_COMPILING, test x$cross_compiling = xyes)

# Add some extra libs here so that previous tests don't fail for
# mysterious reasons - the final link step should bail out.
if test "$have_w32_system" = yes; then
   W32LIBS="-lwsock32"
fi

if test "$GCC" = yes; then
    CFLAGS="$CFLAGS -Wall -mms-bitfields"
    CXXFLAGS="$CXXFLAGS -Wall -mms-bitfields"
    if test "$USE_MAINTAINER_MODE" = "yes"; then
        CFLAGS="$CFLAGS -Wcast-align -Wshadow -Wstrict-prototypes"
        CFLAGS="$CFLAGS -Wno-format-y2k -Wformat-security"
        CXXFLAGS="$CXXFLAGS -Wcast-align -Wshadow"
        CXXFLAGS="$CXXFLAGS -Wno-format-y2k -Wformat-security"
    fi
fi

AC_SUBST(W32LIBS)



#
# Print errors here so that they are visible all
# together and the user can acquire them all together.
#
die=no
if test "$have_gpg_error" = "no"; then
   die=yes
   AC_MSG_NOTICE([[
***  
*** You need libgpg-error to build this program.
**  This library is for example available at
***   ftp://ftp.gnupg.org/pub/gcrypt/libgpg-error
*** (at least version $NEED_GPG_ERROR_VERSION is required.)
***]])
fi
if test "$have_gpgme" = "no"; then
   die=yes
   AC_MSG_NOTICE([[
***  
*** You need gpgme to build this program.
**  This library is for example available at
***   ftp://ftp.gnupg.org/gcrypt/gpgme/
*** (at least version $NEED_GPGME_VERSION is required.)
***]])
fi
if test "$have_libassuan" = "no"; then
   die=yes
   AC_MSG_NOTICE([[
***
*** You need libassuan to build this program.
*** This library is for example available at
***   ftp://ftp.gnupg.org/pub/gcrypt/libassuan/
*** (at least version $NEED_LIBASSUAN_VERSION is required).
***]])
fi
if test "$die" = "yes"; then
    AC_MSG_ERROR([[
***
*** Required libraries not found. Please consult the above messages
*** and install them before running configure again.
***]])
fi


AC_CONFIG_FILES([ Makefile
src/Makefile
src/versioninfo.rc        
forms/Makefile
doc/Makefile
po/Makefile.in
m4/Makefile
])
AC_OUTPUT
