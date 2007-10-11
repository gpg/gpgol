# Makefile.am - main makefile for GPGol
# Copyright (C) 2005 g10 Code GmbH
# 
# This file is free software; as a special exception the author gives
# unlimited permission to copy and/or distribute it, with or without
# modifications, as long as this notice is preserved.
#
# This program is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY, to the extent permitted by law; without even the
# implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

## Process this file with automake to produce Makefile.in

ACLOCAL_AMFLAGS = -I m4
AUTOMAKE_OPTIONS = dist-bzip2
# Because we can only build the w32 version e need to help automake here a bit.
DISTCHECK_CONFIGURE_FLAGS = --host=i586-mingw32msvc --build=i686-pc-linux-gnu \
                            --prefix=@prefix@ \
                            --with-gpg-error-prefix=@prefix@ \
                            --with-gpgme-prefix=@prefix@ \
                            --with-libassuan-prefix=@prefix@

EXTRA_DIST = autogen.sh


SUBDIRS = src po m4 doc

dist-hook:
	echo "$(VERSION)" > $(distdir)/VERSION




