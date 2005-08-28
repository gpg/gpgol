# configure.ac - for Outlgpg
# Copyright (C) 2005 g10 Code GmbH
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

# Version number: Remember to change it immediately *after* a release.
#                 Add a "-cvs" prefix for non-released code.
AC_INIT(outlgpg, 0.99.5-cvs, bug-outlgpg@g10code.com)

NEED_GPGME_API=1
NEED_GPGME_VERSION=1.1.0


PACKAGE=$PACKAGE_NAME
PACKAGE_GT=${PACKAGE_NAME}
VERSION=$PACKAGE_VERSION

AC_CONFIG_SRCDIR(src/olgpgcore.def)
AM_CONFIG_HEADER(config.h)
AC_CANONICAL_TARGET()
AM_INIT_AUTOMAKE($PACKAGE, $VERSION)

AC_GNU_SOURCE

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

BUILD_TIMESTAMP=`date --iso-8601=minutes`
AC_SUBST(BUILD_TIMESTAMP)
changequote(,)dnl
BUILD_FILEVERSION=`echo "$VERSION" | sed 's/\([0-9.]*\).*/\1.0/;s/\./,/g'`
changequote([,])dnl
AC_SUBST(BUILD_FILEVERSION)


AH_BOTTOM([
/* Some global constants. */


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


try_gettext=yes
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
        AC_DEFINE(USE_SIMPLE_GETTEXT,1,
                  [because the Unix gettext has too much overhead on
                   MingW32 systems and these systems lack Posix functions,
                   we use a simplified version of gettext])
        have_dosish_system=yes
        have_w32_system=yes
        try_gettext="no"
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
    if test "$USE_MAINTAINER_MODE" = "yes"; then
        CFLAGS="$CFLAGS -Wall -Wcast-align -Wshadow -Wstrict-prototypes"
        CFLAGS="$CFLAGS -Wno-format-y2k -Wformat-security"
    else
        CFLAGS="$CFLAGS -Wall"
    fi
fi

AC_SUBST(W32LIBS)



# AC_CONFIG_LINKS([
# olflange/gpgexch.cpp:olflange/GPGExch.cpp
# olflange/gpgexch.def:olflange/GPGExch.def
# olflange/gpgexch.dsp:olflange/GPGExch.dsp
# olflange/gpgexch.dsw:olflange/GPGExch.dsw
# olflange/gpgexch.h:olflange/GPGExch.h
# olflange/gpgexch.plg:olflange/GPGExch.plg
# olflange/gpgexch.rc:olflange/GPGExch.rc
# olflange/gpgexchange.h:olflange/GPGExchange.h
# olflange/gpgoptions.cpp:olflange/GPGOptions.cpp
# olflange/gpgoptionsdlg.cpp:olflange/GPGOptionsDlg.cpp
# olflange/stdafx.cpp:olflange/StdAfx.cpp
# olflange/stdafx.h:olflange/StdAfx.h
# ])
# 

#
# Print errors here so that they are visible all
# together and the user can acquire them all together.
#
die=no
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
])
AC_OUTPUT