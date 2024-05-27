#include <oaidl.h>
#include <windows.h>

HRESULT STDMETHODCALLTYPE
str2ansi(BSTR *pdst, BSTR src)
{
	if (!src) {
		*pdst = NULL;
		return S_OK;
	}
	UINT slen = SysStringLen(src);
	UINT dlen = WideCharToMultiByte(0, 0, src, slen, NULL, 0, NULL, NULL);
	BSTR dst = SysAllocStringByteLen(NULL, dlen);
	if (!dst)
		return E_OUTOFMEMORY;
	WideCharToMultiByte(0, 0, src, slen + 1, (LPSTR)dst, dlen + 1, NULL, NULL);
	*pdst = dst;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE
str2unicode(BSTR *pdst, BSTR src)
{
	if (!src) {
		*pdst = NULL;
		return S_OK;
	}
	UINT slen = SysStringByteLen(src);
	UINT dlen = MultiByteToWideChar(0, 0, (LPSTR)src, slen, NULL, 0) * 2;
	BSTR dst = SysAllocStringByteLen(NULL, dlen);
	if (!dst)
		return E_OUTOFMEMORY;
	MultiByteToWideChar(0, 0, (LPSTR)src, slen + 1, dst, dlen + 1);
	*pdst = dst;
	return S_OK;
}