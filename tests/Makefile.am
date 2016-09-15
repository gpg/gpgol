# Makefile.am - Makefile for GpgOL tests.
# Copyright (C) 2016 Intevation GmbH
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

TESTS_ENVIRONMENT = GNUPGHOME=$(abs_srcdir)/tests/gnupg_home

TESTS = initial.test t-parser

AM_LDFLAGS = @GPGME_LIBS@ -lgpgmepp

AM_CFLAGS = -I$(top_srcdir)/src $(GPGME_CFLAGS) $(LIBASSUAN_CFLAGS)
AM_CXXFLAGS = -I$(top_srcdir)/src $(GPGME_CFLAGS) $(GPGME_CFLAGS)/gpgme++ \
			  $(LIBASSUAN_CFLAGS) -std=c++11

parser_SRC= ../src/parsecontroller.cpp \
			../src/parsecontroller.h \
			../src/attachment.cpp ../src/attachment.h \
			../src/mimedataprovider.h ../src/mimedataprovider.cpp \
			../src/rfc822parse.c ../src/rfc822parse.h \
			../src/rfc2047parse.c ../src/rfc2047parse.h \
			../src/common_indep.c ../src/common_indep.h \
			../src/xmalloc.h

t_parser_SOURCES = t-parser.cpp $(parser_SRC)
run_parser_SOURCES = run-parser.cpp $(parser_SRC)

noinst_PROGRAMS = t-parser run-parser

export GNUPGHOME := $(abs_srcdir)/tests/gnupg_home
