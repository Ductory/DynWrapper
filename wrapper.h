#ifndef _WRAPPER_H
#define _WRAPPER_H

#include <oaidl.h>
#include <windows.h>

typedef struct DynWrapper {
	void **vtbl;
	INTERFACEDATA idata;
	UINT size;
} DynWrapper;

HRESULT STDMETHODCALLTYPE DynWrapper_Invoke(ITypeInfo *this, PVOID pvInstance, MEMBERID memid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
ULONG STDMETHODCALLTYPE DynWrapper_Release(IUnknown *this);
HRESULT STDMETHODCALLTYPE DynWrapper_CreateInstance(IUnknown *pUnkOuter, IUnknown **ppunk);

#endif