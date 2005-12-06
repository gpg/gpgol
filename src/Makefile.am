# Makefile.am - main makefile for dialogs part of GPGol
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

bin_PROGRAMS = gpgol 
EXTRA_DIST = versioninfo.rc.in gpgol-rsrcs.rc olflange-rsrcs.rc \
	     logo.bmp decrypt.bmp encrypt.bmp sign.bmp key_mana.bmp \
	     mapi32.def
EXEEXT = .dll

AM_CFLAGS = $(GPGME_CFLAGS) -shared
AM_CXXFLAGS = $(GPGME_CFLAGS) -shared

gpgol_SOURCES = \
	main.c  gpgol.def           \
	resource.rc	            \
        gpgol-ids.h                 \
	olflange.cpp olflange.h     \
	olflange-def.h              \
	olflange-dlgs.cpp           \
	olflange-ids.h              \
	myexchext.h                 \
	display.cpp display.h       \
	gpgmsg.cpp gpgmsg.hh        \
	pgpmime.c pgpmime.h         \
	msgcache.c msgcache.h       \
        engine-gpgme.c	engine.h    \
	rfc822parse.c rfc822parse.h \
        common.c util.h intern.h    \
	passcache.c passcache.h     \
        config-dialog.c	            \
        passphrase-dialog.c         \
        recipient-dialog.c          \
        verify-dialog.c             \
	mymapi.h  mymapitags.h      \
        vasprintf.c                 \
	watcher.cpp \
	attach.cpp attach.h \
	w32-gettext.c w32-gettext.h 


# W32API 3.2 comes with an unusable libmapi32.a.  We build our own
# version.  Note the omission of -k (--kill-at) from the DLLTOOL
# command line.  We also create our own virtual copies to the _static_
# versions of GPGME and gpg-error, because we want to link to them
# statically, and not dynamically (otherwise Outlook would not find
# them).
gpgol_DEPENDENCIES = libmapi32.a libgpgme.a libgpg-error.a

libmapi32.a: mapi32.def
	$(DLLTOOL) --output-lib $@ --def $<

libgpgme.a:
	ln -s $(shell $(GPGME_CONFIG) --prefix)/lib/libgpgme.a

libgpg-error.a:
	ln -s $(shell $(GPG_ERROR_CONFIG) --prefix)/lib/libgpg-error.a

clean-local:
	rm -f libmapi32.a libgpgme.a libgpg-error.a

gpgol_LDADD = $(srcdir)/gpgol.def  \
	-L . -lgpgme -lgpg-error -lmapi32 -lshell32 -lgdi32 -lcomdlg32 \
        -loleaut32

resource.o: resource.rc versioninfo.rc gpgol-rsrcs.rc olflange-rsrcs.rc

.rc.o:
	$(WINDRES) -I $(srcdir) -I . `test -f '$<' || echo '$(srcdir)/'`$< $@
