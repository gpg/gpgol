# Makefile.am - main makefile for dialogs part of OutlGPG
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

bin_PROGRAMS = outlgpg 
EXTRA_DIST = versioninfo.rc.in outlgpg-rsrcs.rc olflange-rsrcs.rc \
	     logo.bmp decrypt.bmp encrypt.bmp sign.bmp key_mana.bmp    
EXEEXT = .dll

AM_CFLAGS = $(GPGME_CFLAGS) -shared
AM_CXXFLAGS = $(GPGME_CFLAGS) -shared

outlgpg_SOURCES = \
	main.c  outlgpg.def         \
	resource.rc	            \
        outlgpg-ids.h               \
	olflange.cpp olflange.h     \
	olflange-def.h              \
	olflange-dlgs.cpp           \
	olflange-ids.h              \
	myexchext.h                 \
	display.cpp display.h       \
	gpgmsg.cpp gpgmsg.hh        \
	msgcache.c msgcache.h       \
        engine-gpgme.c	engine.h    \
        common.c intern.h	    \
	passcache.c passcache.h     \
        HashTable.cpp HashTable.h   \
        keycache.c keycache.h	    \
	usermap.c  usermap.h        \
        config-dialog.c	            \
        passphrase-dialog.c         \
        recipient-dialog.c          \
        verify-dialog.c             \
	mymapi.h  mymapitags.h      \
        vasprintf.c          

outlgpg_LDADD = outlgpg.def  \
		$(GPGME_LIBS) mapi32.dll -lshell32 -lgdi32 -lcomdlg32

resource.o: resource.rc versioninfo.rc outlgpg-rsrcs.rc olflange-rsrcs.rc

.rc.o:
	$(WINDRES) `test -f '$<' || echo '$(srcdir)/'`$< $@







