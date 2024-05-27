#include "dynwrap.h"
#include "factory.h"
#include <winreg.h>
#include <shlwapi.h>
#include <olectl.h>

#define _DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) const GUID DECLSPEC_SELECTANY name = { l,w1,w2,{ b1,b2,b3,b4,b5,b6,b7,b8 } }
_DEFINE_GUID(CLSID_DynWrapper, 0xf9383d93, 0x24b9, 0x4ec7, 0x97,0x3f, 0xb8,0x93,0xb5,0xa9,0x6c,0x30);
#define CLSID_STR "{f9383d93-24b9-4ec7-973f-b893b5a96c30}"
#define CLASS_STR "DynWrapper"


#ifndef NDEBUG

#include <stdio.h>
#include <stdarg.h>

void
con_log(const char *fmt, ...)
{
	fputs("<con_log>: ", stderr);
	va_list ap;
	va_start(ap, fmt);
	vfprintf(stderr, fmt, ap);
	va_end(ap);
}
#endif

static HINSTANCE hInst;

BOOL WINAPI DllMain(HINSTANCE hInstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	if (fdwReason == DLL_PROCESS_ATTACH)
		hInst = hInstDLL;
	return TRUE;
}

LONG ObjectCount;
LONG LockCount;

__declspec(dllexport) HRESULT WINAPI
DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	if (!IsEqualCLSID(&CLSID_DynWrapper, rclsid)) {
		*ppv = NULL;
		return CLASS_E_CLASSNOTAVAILABLE;
	}
	return MyClassFactory.lpVtbl->QueryInterface(&MyClassFactory, riid, ppv);
}

__declspec(dllexport) HRESULT WINAPI
DllCanUnloadNow(void)
{
	return ObjectCount | LockCount ? S_FALSE : S_OK;
}

static HKEY hKey = HKEY_LOCAL_MACHINE;
static const char ThreadingModel[] = "Both";

#define SUBKEY_CLSID "Software\\Classes\\CLSID\\" CLSID_STR
#define SUBKEY_CLASS "Software\\Classes\\" CLASS_STR

__declspec(dllexport) HRESULT WINAPI
DllRegisterServer(void)
{
	char name[512];
	DWORD len = GetModuleFileName(hInst, name, sizeof(name));
	HKEY hkr;
	if (len >= sizeof(name))
		return E_OUTOFMEMORY;
	if (RegCreateKeyEx(hKey, SUBKEY_CLSID "\\InProcServer32", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hkr, NULL))
		return SELFREG_E_CLASS;
	if (RegSetValueEx(hkr, NULL, 0, REG_SZ, name, len + 1))
		return SELFREG_E_CLASS;
	if (RegSetValueEx(hkr, "ThreadingModel", 0, REG_SZ, ThreadingModel, sizeof(ThreadingModel)))
		return SELFREG_E_CLASS;
	RegCloseKey(hkr);
	if (RegCreateKeyEx(hKey, SUBKEY_CLASS "\\CLSID", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hkr, NULL))
		return SELFREG_E_CLASS;
	if (RegSetValueEx(hkr, NULL, 0, REG_SZ, CLSID_STR, sizeof(CLSID_STR)))
		return SELFREG_E_CLASS;
	RegCloseKey(hkr);
	return S_OK;
}

__declspec(dllexport) HRESULT WINAPI
DllUnregisterServer(void)
{
	if (SHDeleteKey(hKey, SUBKEY_CLSID))
		return SELFREG_E_CLASS;
	if (SHDeleteKey(hKey, SUBKEY_CLASS))
		return SELFREG_E_CLASS;
	return S_OK;
}

__declspec(dllexport) HRESULT WINAPI
DllInstall(BOOL bInstall, PCWSTR pszCmdLine)
{
	return S_OK;
}