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

bin_PROGRAMS = olgpgcore 
EXEEXT = .dll

AM_CFLAGS = $(GPGME_CFLAGS) -shared
AM_CXXFLAGS = $(GPGME_CFLAGS) -shared

olgpgcore_SOURCES = \
         main.c  olgpgcore.def	      \
	 MapiGPGME.cpp  MapiGPGME.h   \
         engine-gpgme.c	engine.h      \
         common.c intern.h	      \
         HashTable.cpp HashTable.h    \
         keycache.c keycache.h	      \
	 usermap.c  usermap.h         \
         olgpgcoredlgs.rc             \
         olgpgcoredlgs.h              \
         config-dialog.c	      \
         passphrase-dialog.c          \
         recipient-dialog.c           \
         verify-dialog.c              \
	 mymapi.h  mymapitags.h       \
         vasprintf.c          

olgpgcore_LDADD = olgpgcore.def olgpgcoredlgs.o mapi32.dll $(GPGME_LIBS) \
                  -lshell32 -lgdi32 -lcomdlg32


.rc.o:
	$(WINDRES) `test -f '$<' || echo '$(srcdir)/'`$< $@







