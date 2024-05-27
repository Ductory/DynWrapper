#include "dynwrap.h"
#include "wrapper.h"
#include <stdio.h>

BSTR STDMETHODCALLTYPE
DynWrapper_Version(DynWrapper *this)
{
	return SysAllocString(L"1.0.0");
}

/**
 * check and expand wrapper (if necessary).
 */

static HRESULT STDMETHODCALLTYPE
check_size(DynWrapper *this)
{
	// expand
	if (this->idata.cMembers >= this->size) {
		this->size <<= 1;
		void *ptr;
		ptr = realloc(this->vtbl, this->size * sizeof(void*));
		if (!ptr)
			return E_OUTOFMEMORY;
		this->vtbl = ptr;
		ptr = realloc(this->idata.pmethdata, this->size * sizeof(METHODDATA));
		if (!ptr)
			return E_OUTOFMEMORY;
		this->idata.pmethdata = ptr;
	}
	return S_OK;
}

/** 
 * symbol to VARTYPE
 * 
 * note:
 * due to the implementation of CStdDisp, many VARTYPEs are unsupported.
 * here we allocate a symbol for every VARTYPE CStdDisp supported.
 */

static VARTYPE STDMETHODCALLTYPE
sym2vartype(const wchar_t **p)
{
	VARTYPE vt;
	// basic type
	switch (*(*p)++) {
		case L'v': vt = VT_EMPTY; break;
		case L'c': vt = VT_I1; break;
		case L's': vt = VT_I2; break;
		case L'i': vt = VT_I4; break;
		case L'L': vt = VT_I8; break;
		case L'b': vt = VT_UI1; break;
		case L'w': vt = VT_UI2; break;
		case L'd': case L'u': vt = VT_UI4; break;
		case L'q': vt = VT_UI8; break;
		case L'l': vt = VT_INT; break;
		case L'z': case L'p': vt = VT_UINT; break;
		case L'f': vt = VT_R4; break;
		case L'F': vt = VT_R8; break;
		case L'B': vt = VT_BOOL; break;
		case L'S': vt = VT_BSTR; break;
		case L'V': return VT_VARIANT | VT_BYREF; // !!!
		case L'U': vt = VT_UNKNOWN; break;
		case L'D': vt = VT_DISPATCH; break;
		case L'C': vt = VT_CY; break;
		case L'T': vt = VT_DATE; break;
		case L'E': vt = VT_ERROR; break;
		case L'a': vt = VT_LPSTR; break; // BSTR <-> ANSI
		case L'.': vt = VT_VAARGS; break;
		default: return VT_ILLEGAL;
	}
	// is pointer
	// VT_ARRAY is unnecessary
	if (**p == L'*') {
		++*p;
		vt |= VT_BYREF;
	}
	return vt;
}

/**
 * register the procedure into wrapper's METHODDATA.
 * 
 * procdesc:
 * format:
 *   cc vtret vtargs
 * where
 *   cc: calling convention
 *   vtret: VARTYPE of retval
 *   vtargs: VARTYPE of arguments
 * e.g. `=uuBBu` means `uint32_t (__stdcall *)(uint32_t, BSTR, BSTR, uint32_t)`
 * 
 * note:
 * procname should be allocated by caller
 */

#define DYN_METHOD_BASE 100

static HRESULT STDMETHODCALLTYPE
register_proc(DynWrapper *this, const FARPROC proc, const BSTR procname, const BSTR procdesc)
{
	HRESULT hr;
	UINT idx = this->idata.cMembers++;
	METHODDATA *mdata = this->idata.pmethdata + idx;
	mdata->szName = procname;
	mdata->wFlags = DISPATCH_METHOD;
	mdata->dispid = DYN_METHOD_BASE + idx;
	mdata->iMeth = idx;
	this->vtbl[idx] = proc;
	const wchar_t *p = procdesc;
	// calling convention
	switch (*p) {
		case L'?': // thiscall
			mdata->cc = CC_THISCALL; ++p;
			break;
		case L'_': // cdecl
			mdata->cc = CC_CDECL; ++p;
			break;
		case L'@': // fastcall
			mdata->cc = CC_FASTCALL; ++p;
			break;
		default:
			mdata->cc = CC_STDCALL;
			break;
	}
	VARTYPE vt;
	// return type
	mdata->vtReturn = VT_EMPTY;
	if (*p == L'=') {
		++p;
		vt = sym2vartype(&p);
		if (vt == VT_ILLEGAL) {
			hr = E_INVALIDARG;
			goto parse_vtret_fail;
		}
		mdata->vtReturn = vt;
	}
	// empty parameter list
	if (!*p) {
		mdata->cArgs = 0;
		mdata->ppdata = NULL;
		return S_OK;
	}
	UINT count = 0;
	// cArgs will no more that the length of vtargs
	PARAMDATA *pdata = calloc(wcslen(p), sizeof(PARAMDATA));
	if (!pdata) {
		hr = E_OUTOFMEMORY;
		goto parse_vtargs_fail;
	}
	while (*p) {
		vt = sym2vartype(&p);
		if (vt == VT_ILLEGAL) {
			hr = E_INVALIDARG;
			goto parse_vtargs_fail;
		}
		pdata[count++].vt = vt;
	}
	pdata = realloc(pdata, count * sizeof(PARAMDATA));
	if (!pdata) {
		hr = E_OUTOFMEMORY;
		goto parse_vtargs_fail;
	}
	mdata->cArgs = count;
	mdata->ppdata = pdata;
	return S_OK;
parse_vtargs_fail:
	free(pdata);
parse_vtret_fail:
	free(mdata->szName);
	--this->idata.cMembers;
	return hr;
}

/**
 * register procedure
 * 
 * libname:
 * format:
 *   name
 *   name:id ; register procedure by id
 *   name:procname ; the procname is invalid in the script syntax, use alias instead
 * 
 * procdesc:
 * see `register_proc`
 */

HRESULT STDMETHODCALLTYPE
DynWrapper_Register(DynWrapper *this, const BSTR libname, const BSTR procname, const BSTR procdesc)
{
	CON_PROLOG
	HRESULT hr = check_size(this);
	if (FAILED(hr))
		return hr;
	// parse libname
	wchar_t *lib = wcsdup(libname);
	wchar_t *idname = wcschr(lib, L':');
	if (idname)
		*idname++ = L'\0';
	// load module
	HMODULE hMod = LoadLibraryW(lib);
	if (!hMod) {
		hr = E_INVALIDARG;
		goto loadlib_fail;
	}
	char buf[256], *p = buf;
	wcstombs(buf, idname ? idname : procname, sizeof(buf));
	// register by id
	if (isdigit(buf[0])) {
		int id;
		sscanf(buf, "%d", &id);
		p = (void*)(size_t)id;
	}
	// get address of procedure
	FARPROC proc = GetProcAddress(hMod, p);
	size_t len = wcslen(procname) + 1;
	wchar_t *name = malloc(len * 2 + sizeof(HMODULE));
	if (!name)
		goto dup_fail;
	con_log("hMod: %x\n", hMod);
	*(HMODULE*)name = hMod;
	name = (void*)name + sizeof(HMODULE);
	wcscpy(name, procname);
	hr = register_proc(this, proc, name, procdesc);
	if (FAILED(hr))
		goto register_fail;
	goto clean;
register_fail:
	free((void*)name - sizeof(HMODULE));
dup_fail:
	FreeLibrary(hMod);
loadlib_fail:
clean:
	free(lib);
	CON_EPILOG
	return hr;
}

/**
 * note:
 * if exit abnormally, it may cause errors
 */

static VOID
call_callback(IDispatch *pdisp, METHODDATA *pmdata, ...)
{
	static unsigned char *shell_code;
	if (pmdata->cc != CC_STDCALL && pmdata->cc != CC_CDECL)
		return;
	VARIANT *args = malloc(pmdata->cArgs * sizeof(VARIANT));
	if (!args)
		return;
	DISPPARAMS dparams = {args, NULL, pmdata->cArgs, 0};
	int stack_size = 0;
	va_list ap;
	va_start(ap, pmdata);
	VARIANT *darg = args + dparams.cArgs - 1;
	for (PARAMDATA *p = pmdata->ppdata, *end = p + pmdata->cArgs; p < end; ++p, --darg) {
		darg->vt = p->vt;
		switch (p->vt) {
			case VT_I8: case VT_UI8: case VT_R8: case VT_DATE: case VT_CY:
				darg->llVal = va_arg(ap, long long);
				stack_size += 8;
				break;
			default:
				darg->lVal = va_arg(ap, int);
				stack_size += 4;
				break;
		}
	}
	VARIANT vret = {.vt = VT_EMPTY};
	HRESULT hr;
	hr = pdisp->lpVtbl->Invoke(pdisp, 0, &IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dparams, &vret, NULL, NULL);
	if (FAILED(hr))
		goto clean;
	hr = VariantChangeTypeEx(&vret, &vret, LOCALE_USER_DEFAULT, 0, pmdata->vtReturn);
clean:
	free(args);
	if (!shell_code) {
		shell_code = VirtualAlloc(0, 9, MEM_COMMIT | MEM_RESERVE, PAGE_EXECUTE_READWRITE);
		if (!shell_code)
			return;
		shell_code[0] = 0xB8; // mov eax, xxx
		shell_code[5] = 0xC9; // leave
		shell_code[6] = 0xC2; // ret xxx
	}
	*(LPVOID*)(shell_code + 1) = vret.byref;
	VariantClear(&vret);
	*(unsigned short*)(shell_code + 7) = 8 + (pmdata->cc == CC_STDCALL ? stack_size : 0);
	goto *shell_code;
}

LPVOID STDMETHODCALLTYPE
DynWrapper_RegisterCallback(DynWrapper *this, IDispatch *pdisp, const BSTR procdesc)
{
	static const unsigned char shell_code[] = {
		0x58, // pop eax
		0x68, 0,0,0,0, // push xxx
		0x68, 0,0,0,0, // push xxx
		0x50, // push eax
		0xE9, 0,0,0,0, // jmp xxx
	};
	HRESULT hr = check_size(this);
	if (FAILED(hr))
		return NULL;
	unsigned char *code = VirtualAlloc(NULL, sizeof(shell_code), MEM_COMMIT | MEM_RESERVE, PAGE_EXECUTE_READWRITE);
	if (!code)
		return NULL;
	void *name = malloc(2 + sizeof(HMODULE));
	if (!name)
		goto alloc_name_fail;
	*(LPVOID*)name = code;
	name += sizeof(HMODULE);
	*(wchar_t*)name = L'\0';
	memcpy(code, shell_code, sizeof(shell_code));
	pdisp->lpVtbl->AddRef(pdisp);
	*(LPVOID*)(code + 2) = this->idata.pmethdata + this->idata.cMembers;
	*(LPVOID*)(code + 7) = pdisp;
	*(UINT*)(code + 13) = (void*)call_callback - (void*)(code + 17);
	hr = register_proc(this, (FARPROC)code, name, procdesc);
	if (FAILED(hr))
		goto register_fail;
	return code;
register_fail:
	free(name - sizeof(HMODULE));
alloc_name_fail:
	pdisp->lpVtbl->Release(pdisp);
	VirtualFree(code, 0, MEM_RELEASE);
	return NULL;
}

HRESULT STDMETHODCALLTYPE
DynWrapper_VarPtr(VARIANT *vret, VARIANTARG *var)
{
	if (!vret)
		return S_OK;
	vret->vt = VT_UINT;
	if (var->vt & VT_BYREF) {
		vret->byref = var->byref;
		return S_OK;
	}
	switch (var->vt) {
		case VT_BSTR: case VT_UNKNOWN: case VT_DISPATCH:
			vret->byref = var->byref;
			return S_OK;
		default: // note: return a temporary address. it's meaningless
			vret->byref = var;
			return S_OK;
	}
}

HRESULT STDMETHODCALLTYPE
DynWrapper_StrPtr(VARIANT *vret, VARIANTARG *var)
{
	if (!vret)
		return S_OK;
	vret->vt = VT_UINT;
	if (var->vt == (VT_VARIANT | VT_BYREF) && var->pvarVal->vt == VT_BSTR) {
		vret->byref = var->pvarVal->bstrVal;
		return S_OK;
	}
	if (var->vt == VT_BSTR) {
		vret->byref = var->bstrVal;
		return S_OK;
	}
	return DISP_E_BADVARTYPE;
}


static PARAMDATA RegisterParamData[] = {{NULL, VT_BSTR}, {NULL, VT_BSTR}, {NULL, VT_BSTR}};
static PARAMDATA RegisterCallbackParamData[] = {{NULL, VT_DISPATCH}, {NULL, VT_BSTR}};
static PARAMDATA VarPtrParamData[] = {{NULL, VT_TYPEMASK}};
static PARAMDATA StrPtrParamData[] = {{NULL, VT_BSTR}};

METHODDATA WrapperStaticMethData[] = {
	DEF_METHDATA_0(Version, 0, VT_BSTR),
	DEF_METHDATA(Register, 1, VT_ERROR),
	DEF_METHDATA(RegisterCallback, 2, VT_UINT),
	DEF_METHDATA(VarPtr, 3, VT_UINT),
	DEF_METHDATA(StrPtr, 4, VT_UINT),
};

const void *WrapperStaticVtbl[] = {
	DynWrapper_Version,
	DynWrapper_Register,
	DynWrapper_RegisterCallback,
	DynWrapper_VarPtr,
	DynWrapper_StrPtr,
};

const size_t StaticMethCount = ELEMCOUNT(WrapperStaticMethData);