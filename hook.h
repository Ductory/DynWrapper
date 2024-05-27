#ifndef _HOOK_H
#define _HOOK_H

HRESULT STDMETHODCALLTYPE hook_tinfo(ITypeInfo *this);
HRESULT STDMETHODCALLTYPE hook_stddisp(IUnknown *this);

#endif