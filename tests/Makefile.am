# Makefile.am - Makefile for GpgOL tests.
# Copyright (C) 2016 Bundesamt für Sicherheit in der Informationstechnik
# Software engineering by Intevation GmbH
#
# This file is part of GPGOL.
#
# GpgOL is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# GpgOL is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with this program; if not, see <http://www.gnu.org/licenses/>.

## Process this file with automake to produce Makefile.in

GPG = gpg

if HAVE_W32_SYSTEM
TESTS = t-parser
endif

AM_LDFLAGS = @GPGME_LIBS@ -lgpgmepp

AM_CFLAGS = -I$(top_srcdir)/src $(GPGME_CFLAGS) $(LIBASSUAN_CFLAGS) -DBUILD_TESTS
AM_CXXFLAGS = -I$(top_srcdir)/src $(GPGME_CFLAGS) $(GPGME_CFLAGS)/gpgme++ \
			  $(LIBASSUAN_CFLAGS) -std=c++11  -D_FILE_OFFSET_BITS=64 \
			  -DBUILD_TESTS
LDADD = ../src/gpgol.la @GPG_ERROR_LIBS@

if HAVE_W32_SYSTEM
run_parser_LDADD =  \
	-L ../src -lgpgmepp -lgpgme -lassuan -lgpg-error \
	-lmapi32 -lshell32 -lgdi32 -lcomdlg32 \
	-lole32 -loleaut32 -lws2_32 -ladvapi32 \
	-luuid -lgdiplus -lrpcrt4
endif

parser_SRC= ../src/parsecontroller.cpp \
			../src/parsecontroller.h \
			../src/attachment.cpp ../src/attachment.h \
			../src/mimedataprovider.h ../src/mimedataprovider.cpp \
			../src/rfc822parse.c ../src/rfc822parse.h \
			../src/rfc2047parse.c ../src/rfc2047parse.h \
			../src/common_indep.c ../src/common_indep.h \
			../src/debug.cpp ../src/debug.h \
			../src/memdbg.cpp ../src/memdbg.h \
			../src/cpphelp.cpp ../src/cpphelp.h \
			../src/w32-gettext.cpp ../src/w32-gettext.h \
			../src/xmalloc.h

if !HAVE_W32_SYSTEM
t_parser_SOURCES = t-parser.cpp $(parser_SRC)
endif
run_parser_SOURCES = run-parser.cpp $(parser_SRC)

if !HAVE_W32_SYSTEM
noinst_PROGRAMS = t-parser run-parser
else
noinst_PROGRAMS = run-parser
endif
