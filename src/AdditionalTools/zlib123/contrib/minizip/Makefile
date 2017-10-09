top_srcdir=../../../../..

-include $(top_srcdir)/build/common.mk

ifeq ($(OS),Cygwin)
CC=gcc -mno-cygwin
CFLAGS=-Wall -I ../..
else
LD=ld
CC=cc
CFLAGS=-Wall
endif

ifeq ($(OS),Darwin)
DLLEXT=dylib
else
DLLEXT=so
endif

DEBUG=-g -O0

SOURCE = zip.c unzip.c ioapi.c
HEADER = $(wildcard *.h) $(wildcard win32-zlib/*.h)
DISTFILES=$(SOURCE) $(HEADER)

ifeq ($(OS),Cygwin)
all-local: libzlibwapi.a
else
all-local: libzlibwapi.$(DLLEXT) libzlibwapi.a
endif

libzlibwapi.$(DLLEXT) : $(SOURCE) Makefile
ifeq ($(OS),Darwin)
	$(CC) -dynamiclib -single_module -install_name '@executable_path/$@' -lz -fPIC $(DEBUG) $(CFLAGS) \
	    -o libzlibwapi.$(DLLEXT) $(SOURCE)
else
	$(CC) -lz -shared -fPIC $(DEBUG) $(CFLAGS) \
	    -o libzlibwapi.$(DLLEXT) $(SOURCE)
endif

libzlibwapi.a : $(SOURCE) Makefile
	$(CC) -c $(DEBUG) $(CFLAGS) $(SOURCE)
	ar rc libzlibwapi.a zip.o unzip.o ioapi.o # FIXME - do better
ifeq ($(OS),Darwin)
	ranlib libzlibwapi.a
endif

clean-local:
	rm -f *.o libzlibwapi.$(DLLEXT) libzlibwapi.a

dist-local: dist-common
