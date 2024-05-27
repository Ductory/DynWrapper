#include "dynwrap.h"
#include "wrapper.h"


static IUnknownVtbl *unkVtbl;
static ITypeInfoVtbl *tinfoVtbl;

HRESULT STDMETHODCALLTYPE
hook_tinfo(ITypeInfo *this)
{
	if (!tinfoVtbl) {
		// we allocate an extra entry to store the original method
		tinfoVtbl = malloc(sizeof(ITypeInfoVtbl) + sizeof(void*));
		if (!tinfoVtbl)
			return E_OUTOFMEMORY;
		memcpy(tinfoVtbl, this->lpVtbl, sizeof(ITypeInfoVtbl));
		CALL_ORIGIN(tinfoVtbl, ITypeInfoVtbl, Invoke) = tinfoVtbl->Invoke;
		tinfoVtbl->Invoke = DynWrapper_Invoke;
	}
	this->lpVtbl = tinfoVtbl;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE
hook_stddisp(IUnknown *this)
{
	if (!unkVtbl) {
		// we allocate an extra entry to store the original method
		unkVtbl = malloc(sizeof(IUnknownVtbl) + sizeof(void*));
		if (!unkVtbl)
			return E_OUTOFMEMORY;
		memcpy(unkVtbl, this->lpVtbl, sizeof(IUnknownVtbl));
		CALL_ORIGIN(unkVtbl, IUnknownVtbl, Release) = unkVtbl->Release;
		unkVtbl->Release = DynWrapper_Release;
	}
	this->lpVtbl = unkVtbl;
	return S_OK;
}