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

bin_PROGRAMS = libgpgmedlgs 

AM_CFLAGS = $(GPGME_CFLAGS) -shared
AM_CXXFLAGS = $(GPGME_CFLAGS) -shared

libgpgmedlgs_SOURCES = \
              MapiGPGME.cpp        \
              MapiGPGME.h	   \
	      mymapi.h             \
              mymapitags.h         \
              gpgmedlgs.rc	   \
              libgpgmedlgs.def	   \
              common.c		   \
              config-dialog.c	   \
              engine-gpgme.c	   \
              engine.h		   \
              intern.h		   \
	HashTable.h HashTable.cpp  \
              keycache.c	   \
              keycache.h	   \
	      usermap.c            \
	      usermap.h            \
              main.c		   \
              passphrase-dialog.c  \
              recipient-dialog.c   \
              resource.h           \
              verify-dialog.c      \
	vasprintf.c 

libgpgmedlgs_LDADD = $(GPGME_LIBS) -lmapi32 -lshell32 -lgdi32 -lcomdlg32






