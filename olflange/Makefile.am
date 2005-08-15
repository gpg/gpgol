# Makefile.am - main makefile for the Outlook Flange part of OutlGPG
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

EXTRA_DIST = \
             gpgexch.cpp gpgoptionsdlg.cpp stdafx.cpp gpgexch.h \
             gpgexchange.h resource.h stdafx.h bitmap1.bmp \
	     decrypt.bmp encrypt.bmp gpgexch.def gpgexch.rc \
             key_mana.bmp sign.bmp


bin_PROGRAMS = olgpgmain
EXEEXT = .dll

AM_CFLAGS = $(GPGME_CFLAGS) -shared
AM_CXXFLAGS = $(GPGME_CFLAGS) -shared

olgpgmain_SOURCES = \
        GPGExch.cpp  \
	GPGOptionsDlg.cpp \
	olgpgmaindlgs.rc


olgpgmain_LDADD = GPGExch.def olgpgmaindlgs.o \
	mapi32.dll ../src/olgpgcore.dll $(GPGME_LIBS) \
        -lshell32 -lgdi32 -lcomdlg32


.rc.o:
	$(WINDRES) `test -f '$<' || echo '$(srcdir)/'`$< $@





