#ifndef _DYNWRAP_H
#define _DYNWRAP_H

#ifndef __TINYC__
#error tiny c compiler only.
#endif

#include <oaidl.h>
#include <windows.h>

#ifdef NDEBUG
#define con_log(x...) ((void)0)
#define CON_PROLOG
#define CON_EPILOG
#else
#define CON_PROLOG con_log("enter `%s`\n", __func__);
#define CON_EPILOG con_log("leave `%s`\n", __func__);
void con_log(const char *fmt, ...);
#endif

#define GET_STRUCT(p,typ,of) ((typ*)((void*)(p) - offsetof(typ, of)))
#define ELEMCOUNT(arr) (sizeof(arr) / sizeof(*arr))
#define DEF_METHDATA(name,i,vt) {L ## # name, name ## ParamData, i, i, CC_STDCALL, ELEMCOUNT(name ## ParamData), DISPATCH_METHOD, vt}
#define DEF_METHDATA_0(name,i,vt) {L ## # name, NULL, i, i, CC_STDCALL, 0, DISPATCH_METHOD, vt}
#define CALL_ORIGIN(vtbl,typ,meth) ((typ*)((void**)(vtbl) + (sizeof(typ) - offsetof(typ, meth)) / sizeof(void*)))->meth


// thiscall is not in CALLCONV enum.
// we create it manually.
#define CC_THISCALL (CC_MAX + 1)
// to support variant arguments
#define VT_VAARGS 256


extern LONG ObjectCount;
extern LONG LockCount;

#endif