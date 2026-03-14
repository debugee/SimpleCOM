#include <windows.h>
#include <objbase.h>
#include <strsafe.h>
#include "xdlldata.h"
#include "SimpleCOM.h"
#include "dupatlbase.h"

CAtlModule  _dllModule;
static HMODULE g_hModule = nullptr;
// GUID转字符串的辅助函数
HRESULT GuidToString(const GUID& guid, WCHAR* pszGuid, size_t cchGuid)
{
    return StringCchPrintfW(pszGuid, cchGuid,
        L"{%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
        guid.Data1, guid.Data2, guid.Data3,
        guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
        guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);
}

// 注册表操作辅助函数
HRESULT SetRegistryKeyValue(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, LPCWSTR pszValue)
{
    HKEY hKey;
    LONG lResult = RegCreateKeyExW(hKeyRoot, pszKeyName, 0, nullptr, 
        REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr);
    
    if (lResult != ERROR_SUCCESS)
        return HRESULT_FROM_WIN32(lResult);

    if (pszValue)
    {
        lResult = RegSetValueExW(hKey, pszValueName, 0, REG_SZ, 
            (LPBYTE)pszValue, (DWORD)((wcslen(pszValue) + 1) * sizeof(WCHAR)));
    }

    RegCloseKey(hKey);
    return HRESULT_FROM_WIN32(lResult);
}

HRESULT DeleteRegistryKey(HKEY hKeyRoot, LPCWSTR pszKeyName)
{
    LONG lResult = RegDeleteKeyW(hKeyRoot, pszKeyName);
    return HRESULT_FROM_WIN32(lResult);
}

// DLL入口点
BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID lpReserved)
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(hModule, dwReason, lpReserved))
		return FALSE;
#endif

    switch (dwReason)
    {
    case DLL_PROCESS_ATTACH:
        g_hModule = hModule;
        DisableThreadLibraryCalls(hModule);
        break;
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

// DLL导出函数
STDAPI DllCanUnloadNow(void)
{
#ifdef _MERGE_PROXYSTUB
	HRESULT hr = PrxDllCanUnloadNow();
	if (hr != S_OK)
		return hr;
#endif
    return _pAtlModule->GetLockCount() == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
#ifdef _MERGE_PROXYSTUB
	HRESULT hr = PrxDllGetClassObject(rclsid, riid, ppv);
	if (hr != CLASS_E_CLASSNOTAVAILABLE)
		return hr;
#endif

    if (ppv == nullptr)
        return E_POINTER;

    *ppv = nullptr;

    if (IsEqualCLSID(rclsid, CLSID_DualInterfaceComponent))
    {
        //
        HRESULT hr = CComCreator< CComObject< CComClassFactory > >::CreateInstance((void *)&CDualInterfaceComponent::_CreatorClass::CreateInstance, riid, ppv);
        //we not complete CAtlModule Term() to Release the class factory, but ATL do it.so we use CComObject instead of CComObjectCached
        //HRESULT hr = CDualInterfaceComponent::_ClassFactoryCreatorClass::CreateInstance((void *)&CDualInterfaceComponent::_CreatorClass::CreateInstance, riid, ppv);
        return hr;
    }

    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllRegisterServer(void)
{
    HRESULT hr = S_OK;
    WCHAR szModulePath[MAX_PATH];
    WCHAR szCLSID[50];
    WCHAR szKeyName[256];
    
    // 获取当前DLL路径
    if (GetModuleFileNameW(g_hModule, szModulePath, ARRAYSIZE(szModulePath)) == 0)
        return HRESULT_FROM_WIN32(GetLastError());

    // 将CLSID转换为字符串
    if (FAILED(GuidToString(CLSID_DualInterfaceComponent, szCLSID, ARRAYSIZE(szCLSID))))
        return E_FAIL;

    // 注册CLSID
    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"CLSID\\%s", szCLSID);
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, nullptr, L"DualInterfaceComponent Class");
    if (FAILED(hr)) return hr;
    
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, L"AppID", szCLSID);//CLSID==AppID
    if (FAILED(hr)) return hr;

    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"CLSID\\%s\\ProgID", szCLSID);
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, nullptr, L"DualInterfaceComponent.1");
    if (FAILED(hr)) return hr;

    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"CLSID\\%s\\VersionIndependentProgID", szCLSID);
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, nullptr, L"DualInterfaceComponent");
    if (FAILED(hr)) return hr;

    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"CLSID\\%s\\InprocServer32", szCLSID);
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, nullptr, szModulePath);
    if (FAILED(hr)) return hr;

    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, L"ThreadingModel", L"Apartment");
    if (FAILED(hr)) return hr;

    // 注册ProgID
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, L"DualInterfaceComponent.1", nullptr, L"DualInterfaceComponent Class");
    if (FAILED(hr)) return hr;

    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"DualInterfaceComponent.1\\CLSID");
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, nullptr, szCLSID);
    if (FAILED(hr)) return hr;

    // 注册版本无关的ProgID
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, L"DualInterfaceComponent", nullptr, L"DualInterfaceComponent Class");
    if (FAILED(hr)) return hr;

    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"DualInterfaceComponent\\CurVer");
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, nullptr, L"DualInterfaceComponent.1");
    if (FAILED(hr)) return hr;

    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"DualInterfaceComponent\\CLSID");
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, nullptr, szCLSID);
    if (FAILED(hr)) return hr;

    //注册AppID CLSID_appid == CLSID, all proxy_stub clsid == defalut interface id
    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"AppID\\%s", szCLSID);
    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, nullptr, L"SimpleCOM Application");
    if (FAILED(hr)) return hr;

    hr = SetRegistryKeyValue(HKEY_CLASSES_ROOT, szKeyName, L"DllSurrogate", L"");
    if (FAILED(hr)) return hr;

    HKEY hKey = nullptr;
    LONG lResult = RegCreateKeyW(HKEY_CLASSES_ROOT, szKeyName, &hKey);
    if (lResult != ERROR_SUCCESS)
        return HRESULT_FROM_WIN32(lResult);
    DWORD dwAuthLevel = RPC_C_AUTHN_LEVEL_NONE;
    RegSetValueExW(hKey, L"AuthenticationLevel", 0, REG_DWORD, 
                   (BYTE*)&dwAuthLevel, sizeof(DWORD));
    RegCloseKey(hKey);

#ifdef _MERGE_PROXYSTUB
	hr = PrxDllRegisterServer();
    if (FAILED(hr))
		return hr;
#endif

    return S_OK;
}

STDAPI DllUnregisterServer(void)
{
    HRESULT hr = S_OK;

    WCHAR szCLSID[50];
    WCHAR szKeyName[256];
    
    // 将CLSID转换为字符串
    if (FAILED(GuidToString(CLSID_DualInterfaceComponent, szCLSID, ARRAYSIZE(szCLSID))))
        return E_FAIL;

    // 删除CLSID注册表项
    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"CLSID\\%s\\InprocServer32", szCLSID);
    DeleteRegistryKey(HKEY_CLASSES_ROOT, szKeyName);

    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"CLSID\\%s\\ProgID", szCLSID);
    DeleteRegistryKey(HKEY_CLASSES_ROOT, szKeyName);

    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"CLSID\\%s\\VersionIndependentProgID", szCLSID);
    DeleteRegistryKey(HKEY_CLASSES_ROOT, szKeyName);

    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"CLSID\\%s", szCLSID);
    DeleteRegistryKey(HKEY_CLASSES_ROOT, szKeyName);

    // 删除ProgID
    DeleteRegistryKey(HKEY_CLASSES_ROOT, L"DualInterfaceComponent.1\\CLSID");
    DeleteRegistryKey(HKEY_CLASSES_ROOT, L"DualInterfaceComponent.1");

    DeleteRegistryKey(HKEY_CLASSES_ROOT, L"DualInterfaceComponent\\CurVer");
    DeleteRegistryKey(HKEY_CLASSES_ROOT, L"DualInterfaceComponent\\CLSID");
    DeleteRegistryKey(HKEY_CLASSES_ROOT, L"DualInterfaceComponent");

    // 删除AppID
    StringCchPrintfW(szKeyName, ARRAYSIZE(szKeyName), L"AppID\\%s", szCLSID);
    DeleteRegistryKey(HKEY_CLASSES_ROOT, szKeyName);

#ifdef _MERGE_PROXYSTUB
    if (FAILED(hr))
        return hr;
    hr = PrxDllRegisterServer();//unregister success every time
    if (FAILED(hr))
        return hr;
    hr = PrxDllUnregisterServer();
#endif

    return S_OK;
}
