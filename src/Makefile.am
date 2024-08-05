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

# CFLAGS of used libraries
libs_cflags = $(GPGME_CFLAGS) $(LIBASSUAN_CFLAGS) $(GPG_ERROR_CFLAGS)
libs_cxxflags = $(libs_cflags) $(GPGMEPP_CXXFLAGS)

AM_CFLAGS = $(libs_cflags) -shared
AM_CXXFLAGS = $(libs_cxxflags) -D_FILE_OFFSET_BITS=64 -shared -std=c++14

gpgol_SOURCES = \
    addin-options.cpp addin-options.h \
    addressbook.cpp addressbook.h \
    application-events.cpp \
    attachment.h attachment.cpp \
    categorymanager.h categorymanager.cpp \
    common.h common.cpp \
    common_indep.h common_indep.c \
    cpphelp.cpp cpphelp.h \
    cryptcontroller.cpp cryptcontroller.h \
    debug.h debug.cpp \
    dialogs.h \
    dispcache.h dispcache.cpp \
    eventsink.h \
    eventsinks.h \
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
    mapierr.cpp mapierr.h \
    memdbg.cpp memdbg.h \
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
    recipient.h recipient.cpp \
    recipientmanager.h recipientmanager.cpp \
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

if BUILD_W64
libgpg-error.a:
	if [ -e "$$($(GPG_ERROR_CONFIG) --prefix)/lib64/libgpg-error.a" ]; then \
		ln -sf $$($(GPG_ERROR_CONFIG) --prefix)/lib64/libgpg-error.a .; \
	else \
		ln -sf $$($(GPG_ERROR_CONFIG) --prefix)/lib/libgpg-error.a .; \
	fi
else
libgpg-error.a:
	ln -sf $$($(GPG_ERROR_CONFIG) --prefix)/lib/libgpg-error.a .
endif

if BUILD_W64
libassuan.a:
	if [ -e "$$($(LIBASSUAN_CONFIG) --prefix)/lib64/libassuan.a" ]; then \
		ln -sf $$($(LIBASSUAN_CONFIG) --prefix)/lib64/libassuan.a .; \
	else \
		ln -sf $$($(LIBASSUAN_CONFIG) --prefix)/lib/libassuan.a .; \
	fi
else
libassuan.a:
	ln -sf $$($(LIBASSUAN_CONFIG) --prefix)/lib/libassuan.a .
endif

if BUILD_W64
libgpgme.a:
	if [ -e "$$($(GPGME_CONFIG) --prefix)/lib64/libgpgme.a" ]; then \
		ln -sf $$($(GPGME_CONFIG) --prefix)/lib64/libgpgme.a .; \
	else \
		ln -sf $$($(GPGME_CONFIG) --prefix)/lib/libgpgme.a .; \
	fi
else
libgpgme.a:
	ln -sf $$($(GPGME_CONFIG) --prefix)/lib/libgpgme.a .
endif

if BUILD_W64
libgpgmepp.a:
	if [ -e "$$($(GPGME_CONFIG) --prefix)/lib64/libgpgmepp.a" ]; then \
		ln -sf $$($(GPGME_CONFIG) --prefix)/lib64/libgpgmepp.a .; \
	else \
		ln -sf $$($(GPGME_CONFIG) --prefix)/lib/libgpgmepp.a .; \
	fi
else
libgpgmepp.a:
	ln -sf $$($(GPGME_CONFIG) --prefix)/lib/libgpgmepp.a .
endif

clean-local:
	rm -f libmapi32.a libgpg-error.a libgpgme.a libassuan.a libgpgmepp.a

gpgol_LDFLAGS = -static-libgcc -static-libstdc++ -static -lpthread


gpgol_LDADD = $(srcdir)/gpgol.def  \
	-L . -lgpgmepp -lgpgme -lassuan -lgpg-error \
	-lmapi32 -lshell32 -lgdi32 -lcomdlg32 \
	-lole32 -loleaut32 -lws2_32 -ladvapi32 \
	-luuid -lgdiplus -lrpcrt4

resource.o: resource.rc versioninfo.rc dialogs.rc dialogs.h

.rc.o:
	$(WINDRES) -I $(srcdir) -I . -I .. `test -f '$<' || echo '$(srcdir)/'`$< $@
