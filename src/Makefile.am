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

unused_sources = item-events.cpp 

bin_PROGRAMS = gpgol 
EXTRA_DIST = versioninfo.rc.in mapi32.def $(unused_sources) Outlook.gpl \
	     logo.bmp decrypt.bmp encrypt.bmp sign.bmp key_mana.bmp \
             proto-auto.bmp proto-pgpmime.bmp proto-smime.bmp \
             cryptostate.bmp

EXEEXT = .dll

AM_CFLAGS = $(GPGME_CFLAGS) $(LIBASSUAN_CFLAGS) -shared
AM_CXXFLAGS = $(GPGME_CFLAGS) $(LIBASSUAN_CFLAGS) -shared

gpgol_SOURCES = \
	main.c  gpgol.def           \
	resource.rc	            \
        gpgol-ids.h                 \
	olflange.cpp olflange.h     \
	olflange-def.h              \
	olflange-dlgs.cpp           \
	dialogs.rc dialogs.h        \
	myexchext.h                 \
	display.cpp display.h       \
	message.cpp message.h       \
	revert.cpp revert.h         \
	mimeparser.c mimeparser.h   \
	mimemaker.c mimemaker.h     \
	msgcache.c msgcache.h       \
	engine.c engine.h           \
        engine-assuan.c engine-assuan.h \
        engine-gpgme.c engine-gpgme.h   \
	rfc822parse.c rfc822parse.h \
        common.h common.c util.h    \
	xmalloc.h		    \
	passcache.c passcache.h     \
        config-dialog.c	            \
        passphrase-dialog.c         \
        recipient-dialog.c          \
        verify-dialog.c             \
	mapihelp.cpp mapihelp.h     \
	mymapi.h  mymapitags.h      \
	serpent.c serpent.h         \
        vasprintf.c                 \
	ext-commands.cpp ext-commands.h       \
	user-events.cpp     user-events.h     \
	session-events.cpp  session-events.h  \
	message-events.cpp  message-events.h  \
	attached-file-events.cpp attached-file-events.h \
	property-sheets.cpp property-sheets.h \
	item-events.h         \
	ol-ext-callback.cpp ol-ext-callback.h \
	w32-gettext.c w32-gettext.h 


# W32API 3.2 comes with an unusable libmapi32.a.  We build our own
# version.  Note the omission of -k (--kill-at) from the DLLTOOL
# command line.  We also create our own virtual copies to the _static_
# versions of GPGME and gpg-error, because we want to link to them
# statically, and not dynamically (otherwise Outlook would not find
# them).
gpgol_DEPENDENCIES = libmapi32.a libgpg-error.a libgpgme.a libassuan.a

libmapi32.a: mapi32.def
	$(DLLTOOL) --output-lib $@ --def $<

libgpg-error.a:
	ln -s $$($(GPG_ERROR_CONFIG) --prefix)/lib/libgpg-error.a .

libgpgme.a:
	ln -s $$($(GPGME_CONFIG) --prefix)/lib/libgpgme.a .

libassuan.a:
	ln -s $$($(LIBASSUAN_CONFIG) --prefix)/lib/libassuan.a .

clean-local:
	rm -f libmapi32.a libgpg-error.a libgpgme.a libassuan.a

gpgol_LDADD = $(srcdir)/gpgol.def  \
	-L . -lgpgme -lassuan -lgpg-error \
	-lmapi32 -lshell32 -lgdi32 -lcomdlg32 \
        -lole32 -loleaut32 -lws2_32 -ladvapi32

resource.o: resource.rc versioninfo.rc dialogs.rc 

.rc.o:
	$(WINDRES) -I $(srcdir) -I . `test -f '$<' || echo '$(srcdir)/'`$< $@
