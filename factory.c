#include "dynwrap.h"
#include "wrapper.h"


static HRESULT STDMETHODCALLTYPE
ClassFactory_QueryInterface(IClassFactory *this, REFIID riid, void **ppvObject)
{
	if (!IsEqualIID(&IID_IUnknown, riid)
	 && !IsEqualIID(&IID_IClassFactory, riid)) {
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}
	*ppvObject = this;
	this->lpVtbl->AddRef(this);
	return S_OK;
}

static ULONG STDMETHODCALLTYPE
ClassFactory_AddRef(IClassFactory *this)
{
	return 1;
}

static ULONG STDMETHODCALLTYPE
ClassFactory_Release(IClassFactory *this)
{
	return 1;
}

static HRESULT STDMETHODCALLTYPE
ClassFactory_CreateInstance(IClassFactory *this, IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
	HRESULT hr;
	IUnknown *punk;
	hr = DynWrapper_CreateInstance(pUnkOuter, &punk);
	if (FAILED(hr))
		return hr;
	// `CreateObject` (aka. msvbvm60@rtcCreateObject2) requests IUnknown
	hr = punk->lpVtbl->QueryInterface(punk, riid, ppvObject);
	if (SUCCEEDED(hr))
		InterlockedIncrement(&ObjectCount);
	punk->lpVtbl->Release(punk);
	return hr;
}

static HRESULT STDMETHODCALLTYPE
ClassFactory_LockServer(IClassFactory *this, BOOL flock)
{
	if (flock)
		InterlockedIncrement(&LockCount);
	else
		InterlockedDecrement(&LockCount);
	return S_OK;
}

static IClassFactoryVtbl IClassFactory_Vtbl = {
	ClassFactory_QueryInterface,
	ClassFactory_AddRef,
	ClassFactory_Release,
	ClassFactory_CreateInstance,
	ClassFactory_LockServer,
};

IClassFactory MyClassFactory = {.lpVtbl = &IClassFactory_Vtbl};