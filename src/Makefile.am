# Makefile.am - makefile for GPGol
# Copyright (C) 2005 g10 Code GmbH
# Copyright (C) 2016 Bundesamt für Sicherheit in der Informationstechnik
# Software engineering by Intevation GmbH
# Copyright (C) 2018 Intevation GmbH
#
# This file is free software; as a special exception the author gives
# unlimited permission to copy and/or distribute it, with or without
# modifications, as long as this notice is preserved.
#
# This program is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY, to the extent permitted by law; without even the
# implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

## Process this file with automake to produce Makefile.in
SUBDIRS = icons

bin_PROGRAMS = gpgol 
EXTRA_DIST = \
	versioninfo.rc.in mapi32.def Outlook.gpl \
	dialogs.rc

EXEEXT = .dll

AM_CFLAGS = $(GPGME_CFLAGS) -shared
AM_CXXFLAGS = $(GPGME_CFLAGS) $(LIBASSUAN_CFLAGS) -shared -std=c++11
AM_CXXFLAGS += $(GPGMEPP_CXXFLAGS) -D_FILE_OFFSET_BITS=64

gpgol_SOURCES = \
    addin-options.cpp addin-options.h \
    application-events.cpp \
    attachment.h attachment.cpp \
    common.h common.cpp \
    common_indep.h common_indep.c \
    cpphelp.cpp cpphelp.h \
    cryptcontroller.cpp cryptcontroller.h \
    dialogs.h \
    eventsink.h \
    eventsinks.h \
    exechelp.c exechelp.h \
    explorer-events.cpp \
    explorers-events.cpp \
    filetype.c filetype.h \
    folder-events.cpp \
    gmime-table-private.h \
    gpgoladdin.cpp gpgoladdin.h \
    gpgol.def \
    gpgol-ids.h \
    keycache.cpp keycache.h \
    mail.h mail.cpp \
    mailitem-events.cpp \
    main.c \
    mapihelp.cpp mapihelp.h \
    mimedataprovider.cpp mimedataprovider.h \
    mimemaker.cpp mimemaker.h \
    mlang-charset.cpp mlang-charset.h \
    mymapi.h \
    mymapitags.h \
    olflange.cpp olflange.h \
    oomhelp.cpp oomhelp.h \
    overlay.cpp overlay.h \
    parsecontroller.cpp parsecontroller.h \
    parsetlv.h parsetlv.c \
    resource.rc \
    revert.cpp revert.h \
    rfc2047parse.h rfc2047parse.c \
    rfc822parse.c rfc822parse.h \
    ribbon-callbacks.cpp ribbon-callbacks.h \
    w32-gettext.cpp w32-gettext.h \
    windowmessages.h windowmessages.cpp \
    wks-helper.cpp wks-helper.h \
    xmalloc.h

#treeview_SOURCES = treeview.c

# W32API 3.2 comes with an unusable libmapi32.a.  We build our own
# version.  Note the omission of -k (--kill-at) from the DLLTOOL
# command line.  We also create our own virtual copies to the _static_
# versions of GPGME and gpg-error, because we want to link to them
# statically, and not dynamically (otherwise Outlook would not find
# them).
gpgol_DEPENDENCIES = libmapi32.a libgpg-error.a libgpgme.a libassuan.a libgpgmepp.a

if BUILD_W64
  DLLTOOLFLAGS64=--as-flags=--64 -m i386:x86-64
endif

libmapi32.a: mapi32.def
	$(DLLTOOL) $(DLLTOOLFLAGS64) --output-lib $@ --def $<

libgpg-error.a:
	ln -s $$($(GPG_ERROR_CONFIG) --prefix)/lib/libgpg-error.a .

libgpgme.a:
	ln -s $$($(GPGME_CONFIG) --prefix)/lib/libgpgme.a .

libassuan.a:
	ln -s $$($(LIBASSUAN_CONFIG) --prefix)/lib/libassuan.a .

libgpgmepp.a:
	ln -s $$($(GPGME_CONFIG) --prefix)/lib/libgpgmepp.a .

clean-local:
	rm -f libmapi32.a libgpg-error.a libgpgme.a libassuan.a libgpgmepp.a

gpgol_LDFLAGS = -static-libgcc -static-libstdc++

gpgol_LDADD = $(srcdir)/gpgol.def  \
	-L . -lgpgmepp -lgpgme -lassuan -lgpg-error \
	-lmapi32 -lshell32 -lgdi32 -lcomdlg32 \
	-lole32 -loleaut32 -lws2_32 -ladvapi32 \
	-luuid -lgdiplus -lrpcrt4

resource.o: resource.rc versioninfo.rc dialogs.rc dialogs.h

.rc.o:
	$(WINDRES) -I $(srcdir) -I . -I .. `test -f '$<' || echo '$(srcdir)/'`$< $@
