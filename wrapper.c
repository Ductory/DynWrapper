#include "dynwrap.h"
#include "stmethod.h"
#include "hook.h"
#include "util.h"
#include <stdio.h>
#include <stdlib.h>


/**
 * get METHODDATA by DISPID.
 */

static HRESULT STDMETHODCALLTYPE
mdata_of_dispid(DynWrapper *this, MEMBERID memid, WORD wFlags, METHODDATA **ppmdata)
{
	INTERFACEDATA *idata = &this->idata;
	METHODDATA *mdata = idata->pmethdata + StaticMethCount;
	METHODDATA *end = idata->pmethdata + idata->cMembers;
	for (METHODDATA *p = mdata; p < end; ++p)
		if (p->dispid == memid && (p->wFlags & wFlags)) {
			*ppmdata = p;
			return S_OK;
		}
	return DISP_E_MEMBERNOTFOUND;
}

/**
 * for convenience, we convert DISPPARAMS to INVOKEARGS first.
 */

typedef struct {
	UINT cArgs;
	VARIANT *var;
} INVOKEARGS;

static HRESULT STDMETHODCALLTYPE
build_invoke_args(DISPPARAMS *dparams, PARAMDATA *ppdata, INVOKEARGS **ppiargs)
{
	HRESULT hr;
	// initialize INVOKEARGS
	INVOKEARGS *iargs = malloc(sizeof(INVOKEARGS));
	if (!iargs) {
		hr = E_OUTOFMEMORY;
		goto alloc_iargs_fail;
	}
	iargs->var = calloc(dparams->cArgs, sizeof(VARIANT));
	if (!iargs->var) {
		hr = E_OUTOFMEMORY;
		goto alloc_var_fail;
	}
	// parse args
	iargs->cArgs = 0;
	VARIANT *begin = dparams->rgvarg, *end;
	for (VARIANT *src = dparams->rgvarg + dparams->cArgs - 1, *dst = iargs->var + dparams->cArgs - 1; src >= begin; --src, --dst, ++ppdata) {
		if (ppdata->vt == VT_VAARGS) {
			*dst = *src;
			--ppdata;
		} else if (ppdata->vt & VT_BYREF) {
			if (src->vt != ppdata->vt) {
				hr = DISP_E_TYPEMISMATCH;
				goto change_type_fail;
			}
			*dst = *src;
		} else if (ppdata->vt == VT_LPSTR) { // BSTR -> ANSI
			hr = VariantChangeTypeEx(dst, src, LOCALE_USER_DEFAULT, 0, VT_BSTR);
			if (FAILED(hr))
				goto change_type_fail;
			BSTR s = dst->bstrVal;
			hr = str2ansi(&dst->bstrVal, s);
			SysFreeString(s);
			if (FAILED(hr)) {
				VariantClear(dst);
				goto change_type_fail;
			}
		} else {
			hr = VariantChangeTypeEx(dst, src, LOCALE_USER_DEFAULT, 0, ppdata->vt);
			if (FAILED(hr))
				goto change_type_fail;
		}
		++iargs->cArgs;
	}
	*ppiargs = iargs;
	return S_OK;
change_type_fail:
	end = iargs->var + dparams->cArgs;
	for (VARIANT *p = end - iargs->cArgs; p < end; ++p)
		VariantClear(p);
	free(iargs->var);
alloc_var_fail:
	free(iargs);
alloc_iargs_fail:
	return hr;
}

static HRESULT STDMETHODCALLTYPE
release_invoke_args(INVOKEARGS *iargs)
{
	VARIANT *end = iargs->var + iargs->cArgs;
	for (VARIANT *p = iargs->var; p < end; ++p)
		VariantClear(p);
	free(iargs->var);
	free(iargs);
	return S_OK;
}

/**
 * note:
 * these function are a hybrid of C and ASM,
 * so they depend heavily on the compiler implementation.
 * (or you should rewrite them entirely with ASM :P)
 */

static HRESULT STDMETHODCALLTYPE
call_func(DynWrapper *this, METHODDATA *pmdata, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	HRESULT hr;
	INVOKEARGS *iargs;
	hr = build_invoke_args(pDispParams, pmdata->ppdata, &iargs);
	if (FAILED(hr))
		return hr;
	void *meth = this->vtbl[pmdata->iMeth];
	int regs = (pmdata->cc == CC_THISCALL) | (pmdata->cc == CC_FASTCALL) << 1;
	if (pDispParams->cArgs < regs)
		regs = pDispParams->cArgs;
	int stack_size = regs * -4;
	void *res;
	for (VARIANT *p = iargs->var, *end = iargs->var + pDispParams->cArgs; p < end; ++p) {
		if (p->vt | VT_BYREF)
			goto push4;
		switch (p->vt) {
			case VT_I8: case VT_UI8: case VT_R8: case VT_DATE: case VT_CY:
				asm volatile ("pushl %0; pushl %1": : "r"(p->cyVal.Lo), "r"(p->cyVal.Hi));
				stack_size += 8;
				break;
		}
		continue;
	push4:
		asm volatile ("pushl %0": : "r"(p->lVal));
		stack_size += 4;
	}
	if (regs == 2)
		asm volatile ("popl %edx");
	if (regs >= 1)
		asm volatile ("popl %ecx");
	asm volatile ("call *%0": "=a"(res) : "r"(meth));
	if (pmdata->cc != CC_STDCALL)
		asm volatile ("add %0, %%esp": : "r"(stack_size));
	PARAMDATA *end = pmdata->ppdata + pmdata->cArgs;
	VARIANT *darg = pDispParams->rgvarg + pDispParams->cArgs - 1;
	VARIANT *iarg = iargs->var + iargs->cArgs - 1;
	for (PARAMDATA *p = pmdata->ppdata; p < end; ++p, --darg, --iarg) { // ANSI -> BSTR
		if (p->vt != VT_LPSTR)
			continue;
		if (darg->vt != (VT_VARIANT | VT_BYREF) || darg->pvarVal->vt != VT_BSTR)
			continue;
		hr = str2unicode(&darg->pvarVal->bstrVal, iarg->bstrVal);
		if (FAILED(hr)) {
			VariantClear(pVarResult);
			goto clean;
		}
	}
	if (pVarResult) {
		if (pmdata->vtReturn == VT_LPSTR) { // ANSI -> BSTR
			pVarResult->vt = VT_BSTR;
			BSTR s = pVarResult->bstrVal;
			hr = str2unicode(&pVarResult->bstrVal, s);
			SysFreeString(s);
			if (FAILED(hr)) {
				VariantClear(pVarResult);
				goto clean;
			}
		} else {
			pVarResult->vt = pmdata->vtReturn;
			pVarResult->byref = res;
		}
	}
clean:
	return release_invoke_args(iargs);
}

/**
 * we wrapped `CDispTypeInfo::Invoke`.
 * if the static method is called, we simply forward it to `CDispTypeInfo::Invoke`.
 * otherwise (with a dynamically registered method called), we call `call_func`.
 */

HRESULT STDMETHODCALLTYPE
DynWrapper_Invoke(ITypeInfo *this, PVOID pvInstance, MEMBERID memid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	CON_PROLOG
	for (UINT i = 0; i < pDispParams->cArgs; ++i)
		con_log("\tvt: %x\n", pDispParams->rgvarg[i].vt);

	HRESULT hr;
	// VarPtr, StrPtr
	if (memid >= 3 && memid <= 4) {
		// we don't support named arguments
		if (pDispParams->cArgs != 1 || pDispParams->cNamedArgs)
			return DISP_E_BADPARAMCOUNT;
		return ((HRESULT STDMETHODCALLTYPE(*)(VARIANT *vret, VARIANTARG *var))WrapperStaticVtbl[memid])(pVarResult, pDispParams->rgvarg);
	}
	// static method
	if (memid < StaticMethCount)
		return CALL_ORIGIN(this->lpVtbl, ITypeInfoVtbl, Invoke)(this, pvInstance, memid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	// dynamic method
	DynWrapper *wrapper = pvInstance;
	METHODDATA *mdata;
	hr = mdata_of_dispid(wrapper, memid, wFlags, &mdata);
	if (FAILED(hr))
		return hr;
	// we ignore redundant parameters (to support variant arguments)
	if (pDispParams->cArgs < mdata->cArgs || pDispParams->cNamedArgs)
		return DISP_E_BADPARAMCOUNT;
	return call_func(wrapper, mdata, pDispParams, pVarResult, pExcepInfo, puArgErr);
}

static VOID STDMETHODCALLTYPE
release_instance(DynWrapper *this)
{
	free(this->vtbl);
	METHODDATA *mdata = this->idata.pmethdata;
	METHODDATA *end = mdata + this->idata.cMembers;
	for (METHODDATA *p = mdata + StaticMethCount; p < end; ++p) {
		int c = *p->szName;
		p->szName = (void*)p->szName - sizeof(HMODULE);
		if (c) {
			FreeLibrary(*(HMODULE*)p->szName);
		} else {
			LPVOID code = *(LPVOID*)p->szName;
			IDispatch *pdisp = *(IDispatch**)(code + 7);
			pdisp->lpVtbl->Release(pdisp);
			VirtualFree(code, 0, MEM_RELEASE);
		}
		free(p->szName);
	}
	free(mdata);
	free(this);
}

/**
 * we wrapped `CStdDispUnkImpl::Release` to release wrapper and decrease `ObjectCount` when _cRef becomes zero.
 * 
 * we don't release vtable because all wrappers share the same vatble.
 */

/* This structure is reversed from `CStdDisp::Create`. */

typedef struct StdDisp {
	IDispatchVtbl *dispVtbl;
	IUnknownVtbl *unkVtbl;
	struct StdDisp *self;
	SIZE_T _cRef;
	IUnknown *unkOuter;
	void *pubObject;
	ITypeInfo *tinfo;
} StdDisp;

ULONG STDMETHODCALLTYPE
DynWrapper_Release(IUnknown *this)
{
	StdDisp *stddisp = GET_STRUCT(this, StdDisp, unkVtbl);
	DynWrapper *wrapper = stddisp->pubObject;
	ULONG cref = CALL_ORIGIN(this->lpVtbl, IUnknownVtbl, Release)(this);
	if (!cref) {
		release_instance(wrapper);
		InterlockedDecrement(&ObjectCount);
	}
	return cref;
}

/**
 * this is the most important part of the program.
 * we use `CreateDispTypeInfo` and `CreateStdDispatch` to dynamically register the method,
 * and no longer have to implement the `ITypeInfo` and `IDispatch` interfaces ourselves.
 * the drawback is that it relies on the `StdDisp` structure obtained by reversing,
 * so its compatibility is not guaranteed.
 */

#define DYNWRAPPER_INITSIZE 16

HRESULT STDMETHODCALLTYPE
DynWrapper_CreateInstance(IUnknown *pUnkOuter, IUnknown **ppunk)
{
	HRESULT hr;
	DynWrapper *wrapper = malloc(sizeof(DynWrapper));
	if (!wrapper) {
		hr = E_OUTOFMEMORY;
		goto create_fail;
	}
	wrapper->idata.cMembers = StaticMethCount;
	wrapper->size = DYNWRAPPER_INITSIZE;
	wrapper->vtbl = malloc(wrapper->size * sizeof(void*));
	if (!wrapper->vtbl) {
		hr = E_OUTOFMEMORY;
		goto alloc_vtbl_fail;
	}
	memcpy(wrapper->vtbl, WrapperStaticVtbl, StaticMethCount * sizeof(void*));
	wrapper->idata.pmethdata = malloc(wrapper->size * sizeof(METHODDATA));
	if (!wrapper->idata.pmethdata) {
		hr = E_OUTOFMEMORY;
		goto alloc_methdata_fail;
	}
	memcpy(wrapper->idata.pmethdata, WrapperStaticMethData, StaticMethCount * sizeof(METHODDATA));
	ITypeInfo *ptinfo;
	hr = CreateDispTypeInfo(&wrapper->idata, LOCALE_USER_DEFAULT, &ptinfo);
	if (FAILED(hr))
		goto create_itinfo_fail;
	hr = hook_tinfo(ptinfo);
	if (FAILED(hr))
		goto hook_tinfo_fail;
	IUnknown *punk;
	hr = CreateStdDispatch(pUnkOuter, wrapper, ptinfo, &punk);
	if (FAILED(hr))
		goto create_stddisp_fail;
	hr = hook_stddisp(punk);
	if (FAILED(hr))
		goto hook_stddisp_fail;
	ptinfo->lpVtbl->Release(ptinfo);
	*ppunk = punk;
	return S_OK;
hook_stddisp_fail:
	punk->lpVtbl->Release(punk);
create_stddisp_fail:
hook_tinfo_fail:
	ptinfo->lpVtbl->Release(ptinfo);
create_itinfo_fail:
	free(wrapper->idata.pmethdata);
alloc_methdata_fail:
	free(wrapper->vtbl);
alloc_vtbl_fail:
	free(wrapper);
create_fail:
	*ppunk = NULL;
	return hr;
}