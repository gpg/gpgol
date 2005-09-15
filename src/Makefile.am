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
	     logo.bmp decrypt.bmp encrypt.bmp sign.bmp key_mana.bmp    
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
        common.c intern.h	    \
	passcache.c passcache.h     \
        keycache.c keycache.h	    \
        config-dialog.c	            \
        passphrase-dialog.c         \
        recipient-dialog.c          \
        verify-dialog.c             \
	mymapi.h  mymapitags.h      \
        vasprintf.c                 \
	watcher.cpp

gpgol_LDADD = gpgol.def  \
		$(GPGME_LIBS) mapi32.dll -lshell32 -lgdi32 -lcomdlg32

resource.o: resource.rc versioninfo.rc gpgol-rsrcs.rc olflange-rsrcs.rc

.rc.o:
	$(WINDRES) `test -f '$<' || echo '$(srcdir)/'`$< $@







