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

# Note: Do not put ChangeLog-2011 here - it exists only in the repo.
EXTRA_DIST = autogen.sh autogen.rc


SUBDIRS = src forms po m4 doc

dist-hook: gen-ChangeLog
	echo "$(VERSION)" > $(distdir)/VERSION

gen_start_date = 2011-12-01T00:00:00
.PHONY: gen-ChangeLog
gen-ChangeLog:
	if test -d $(top_srcdir)/.git; then				\
	  (cd $(top_srcdir) &&                      			\
	    ./build-aux/gitlog-to-changelog				\
	    --amend=build-aux/git-log-fix  --tear-off			\
	    --since=$(gen_start_date) ) > $(distdir)/cl-t;		\
          cat $(top_srcdir)/ChangeLog-2011 >> $(distdir)/cl-t;          \
	  rm -f $(distdir)/ChangeLog;					\
	  mv $(distdir)/cl-t $(distdir)/ChangeLog;			\
	fi




