# DynWrapper

现在只是一个demo，可能存在大量bug和内存泄漏问题。文档也没写完。

# Build

Some functions are a hybrid of C and Assembly (see `wrapper.c`), so they depend heavily on the compiler implementation. Hence, **only TCC (aka. Tiny C Compiler) is supported.**

Besides, due to tcc will decorates the function names of the `__stdcall` convention (see `tccgen.c`), so the exported functions cannot be recognized. To solve this problem, You can either modify the source code of tcc, or just implement `symmod` yourself to change the symbol names in the elf file (`dynwrap.o`).

If you choose the latter, you can simply use the Makefile to build it. Otherwise, just remove the symmod in Makefile.

# Register DLL

Run `regsvr32` as administrator. "Register for current user" is not implemented.

```
regsvr32 dynwrap.dll
```

To unregister it, just run

```
regsvr32 /u dynwrap.dll
```

# How to use

懒得写，到时候再补上

Note that we only implemented the 32-bit version, so you cannot just click the vbs file and run it.

```
%WINDIR%\SysWOW64\cscript.exe your_vbs.vbs
```

# Methods List

## Version

## Register

## RegisterCallback

## VarPtr

## StrPtr

# Example

```vbscript
set dw = CreateObject("DynWrapper")
' StrPtr
dw.Register "user32", "MessageBoxW", "=uupSu"
dw.MessageBoxW 0, dw.StrPtr("Prompt"), "Title", 0
' BSTR <-> ANSI, alias
dw.Register "user32:MessageBoxA", "MessageBox", "=uuaau"
dw.MessageBox 0, dw.Version, "Version", 0
' cdecl, va args
dw.Register "msvcrt", "sprintf", "_=iaa."
buf = Space(256)
dw.sprintf buf, "decimal: %d, hex: %x", CLng(100), CLng(100)
MsgBox buf
' callback
dw.Register "msvcrt", "qsort", "_pzzp"
arr = Array(5, 0, 2, 7, 9, 0, 4)
pcomp = dw.RegisterCallback(GetRef("comp"), "_=iVV")
dw.qsort dw.VarPtr(arr(0)), 7, 16, pcomp
MsgBox Join(arr)

function comp(a, b)
	comp = a - b
end function
```

