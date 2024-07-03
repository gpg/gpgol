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
# Because we can only build the w32 version we need to help automake here a bit.
DISTCHECK_CONFIGURE_FLAGS = --host=@host@  --build=@build@ \
			    GPGRT_CONFIG=@SYSROOT@/bin/gpgrt-config \
                            SYSROOT=@SYSROOT@ \
                            PKG_CONFIG_LIBDIR=@SYSROOT@/lib/pkgconfig

# Note: Do not put ChangeLog-2011 here - it exists only in the repo.
EXTRA_DIST = autogen.sh autogen.rc

# Fixme: Running the tests fails in in distcheck.
tests=

if HAVE_W32_SYSTEM
SUBDIRS = src ${tests} forms po m4 doc
else
SUBDIRS = tests
endif

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

distcheck-hook:
	set -e; ( \
	pref="#+macro: gpgol_" ;\
	reldate="$$(date -u +%Y-%m-%d)" ;\
        echo "$${pref}ver  $(PACKAGE_VERSION)"  ;\
        echo "$${pref}date $${reldate}" ;\
        list='$(DIST_ARCHIVES)'; for i in $$list; do \
	  case "$$i" in *.tar.bz2) \
            echo "$${pref}size $$(wc -c <$$i|awk '{print int($$1/1024)}')k" ;\
	    echo "$${pref}sha1 $$(sha1sum <$$i|cut -d' ' -f1)" ;\
	    echo "$${pref}sha2 $$(sha256sum <$$i|cut -d' ' -f1)" ;;\
	  esac;\
	done ) | tee $(distdir).swdb
