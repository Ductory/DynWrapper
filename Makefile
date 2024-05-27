PROJ   := dynwrap
TARGET := $(PROJ)
FILES  := $(notdir $(wildcard ./*.c))
OBJS   := $(FILES:.c=.o)
LIBS   :=

CC := tcc
LD := tcc

CFLAGS  := -DNDEBUG
LDFLAGS := -loleaut32 -lmyguid -ladvapi32 -lshlwapi

.PHONY : all clean

all : $(TARGET).dll

$(TARGET).dll : $(OBJS)
	./symmod $(TARGET).o
	$(LD) $^ $(LDFLAGS) -shared -o $@

$(OBJS) : %.o : %.c
	$(CC) $(CFLAGS) -c $< -o $@

clean :
	@rm -fv *.o

dynwrapper.o : dynwrapper.c
	tcc -c dynwrapper.c -DNDEBUG
