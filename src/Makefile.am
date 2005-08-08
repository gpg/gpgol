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

AM_CFLAGS = $(GPGME_CFLAGS)
AM_CXXFLAGS = $(GPGME_CFLAGS)

outlgpg_SOURCES = \
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
              keycache.c	   \
              keycache.h	   \
	      usermap.c            \
	      usermap.h            \
              main.c		   \
              passphrase-dialog.c  \
              recipient-dialog.c   \
              resource.h           \
              verify-dialog.c

outlgpg_LDADD = $(GPGME_LIBS)






