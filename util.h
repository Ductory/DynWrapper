#ifndef _UTIL_H
#define _UTIL_H

HRESULT STDMETHODCALLTYPE str2ansi(BSTR *pdst, BSTR src);
HRESULT STDMETHODCALLTYPE str2unicode(BSTR *pdst, BSTR src);

#endif