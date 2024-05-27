#ifndef _STMETHOD_H
#define _STMETHOD_H

#include "wrapper.h"

BSTR STDMETHODCALLTYPE DynWrapper_Version(DynWrapper *this);
HRESULT STDMETHODCALLTYPE DynWrapper_Register(DynWrapper *this, const BSTR libname, const BSTR procname, const BSTR procdesc);
LPVOID STDMETHODCALLTYPE DynWrapper_RegisterCallback(DynWrapper *this, IDispatch *pdisp, const BSTR procdesc);
HRESULT STDMETHODCALLTYPE DynWrapper_VarPtr(VARIANT *vret, VARIANTARG *var);
HRESULT STDMETHODCALLTYPE DynWrapper_StrPtr(VARIANT *vret, VARIANTARG *var);

extern METHODDATA WrapperStaticMethData[];
extern const void *WrapperStaticVtbl[];
extern const size_t StaticMethCount;

#endif