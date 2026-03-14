#include "SimpleCOM.h"
#include <algorithm>
#include <cwctype>
#include <strsafe.h>
#include <shlwapi.h>

// IDispatch接口实现
STDMETHODIMP CDualInterfaceComponent::GetTypeInfoCount(UINT* pctinfo)
{
    if (!pctinfo)
        return E_POINTER;
    
    *pctinfo = 1;
    return S_OK;
}

STDMETHODIMP CDualInterfaceComponent::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    if (!ppTInfo)
        return E_POINTER;
    
    *ppTInfo = nullptr;
    
    if (iTInfo != 0)
        return DISP_E_BADINDEX;
    
    HRESULT hr = LoadTypeInfo();
    if (SUCCEEDED(hr))
    {
        *ppTInfo = m_pTypeInfo;
        m_pTypeInfo->AddRef();
    }
    
    return hr;
}

STDMETHODIMP CDualInterfaceComponent::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
    if (!IsEqualIID(riid, IID_NULL))
        return DISP_E_UNKNOWNINTERFACE;
    
    if (!rgszNames || !rgDispId || cNames == 0)
        return E_INVALIDARG;
    
    // 手动映射方法名到 DISPID
    // ICalculator methods: DISPID 1-4
    // IStringProcessor methods: DISPID 101-104
    LPOLESTR szName = rgszNames[0];
    
    // ICalculator methods
    if (_wcsicmp(szName, L"Add") == 0)
    {
        rgDispId[0] = 1;
        return S_OK;
    }
    else if (_wcsicmp(szName, L"Subtract") == 0)
    {
        rgDispId[0] = 2;
        return S_OK;
    }
    else if (_wcsicmp(szName, L"Multiply") == 0)
    {
        rgDispId[0] = 3;
        return S_OK;
    }
    else if (_wcsicmp(szName, L"Divide") == 0)
    {
        rgDispId[0] = 4;
        return S_OK;
    }
    // IStringProcessor methods
    else if (_wcsicmp(szName, L"GetLength") == 0)
    {
        rgDispId[0] = 101;
        return S_OK;
    }
    else if (_wcsicmp(szName, L"ToUpper") == 0)
    {
        rgDispId[0] = 102;
        return S_OK;
    }
    else if (_wcsicmp(szName, L"ToLower") == 0)
    {
        rgDispId[0] = 103;
        return S_OK;
    }
    else if (_wcsicmp(szName, L"Reverse") == 0)
    {
        rgDispId[0] = 104;
        return S_OK;
    }
    
    return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP CDualInterfaceComponent::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
    if (!IsEqualIID(riid, IID_NULL))
        return DISP_E_UNKNOWNINTERFACE;
    
    // ICalculator methods (DISPID 1-4)
    if (dispIdMember >= 1 && dispIdMember <= 4)
    {
        if (wFlags & DISPATCH_METHOD)
        {
            if (pDispParams->cArgs != 2)
                return DISP_E_BADPARAMCOUNT;
            
            VARIANT* pArgs = pDispParams->rgvarg;
            long a = 0, b = 0, result = 0;
            
            // COM 中参数是逆序传递的
            if (V_VT(&pArgs[1]) == VT_I4) a = V_I4(&pArgs[1]);
            else if (V_VT(&pArgs[1]) == VT_I2) a = V_I2(&pArgs[1]);
            else return DISP_E_TYPEMISMATCH;
            
            if (V_VT(&pArgs[0]) == VT_I4) b = V_I4(&pArgs[0]);
            else if (V_VT(&pArgs[0]) == VT_I2) b = V_I2(&pArgs[0]);
            else return DISP_E_TYPEMISMATCH;
            
            HRESULT hr = S_OK;
            switch (dispIdMember)
            {
                case 1: hr = Add(a, b, &result); break;
                case 2: hr = Subtract(a, b, &result); break;
                case 3: hr = Multiply(a, b, &result); break;
                case 4: hr = Divide(a, b, &result); break;
                default: return DISP_E_MEMBERNOTFOUND;
            }
            
            if (SUCCEEDED(hr) && pVarResult)
            {
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = result;
            }
            
            return hr;
        }
    }
    // IStringProcessor methods (DISPID 101-104)
    else if (dispIdMember >= 101 && dispIdMember <= 104)
    {
        if (wFlags & DISPATCH_METHOD)
        {
            VARIANT* pArgs = pDispParams->rgvarg;
            
            switch (dispIdMember)
            {
                case 101: // GetLength
                {
                    if (pDispParams->cArgs != 1)
                        return DISP_E_BADPARAMCOUNT;
                    
                    if (V_VT(&pArgs[0]) != VT_BSTR)
                        return DISP_E_TYPEMISMATCH;
                    
                    long length = 0;
                    HRESULT hr = GetLength(V_BSTR(&pArgs[0]), &length);
                    
                    if (SUCCEEDED(hr) && pVarResult)
                    {
                        V_VT(pVarResult) = VT_I4;
                        V_I4(pVarResult) = length;
                    }
                    
                    return hr;
                }
                case 102: // ToUpper
                case 103: // ToLower
                case 104: // Reverse
                {
                    if (pDispParams->cArgs != 1)
                        return DISP_E_BADPARAMCOUNT;
                    
                    if (V_VT(&pArgs[0]) != VT_BSTR)
                        return DISP_E_TYPEMISMATCH;
                    
                    BSTR output = nullptr;
                    HRESULT hr = S_OK;
                    
                    switch (dispIdMember)
                    {
                        case 102: hr = ToUpper(V_BSTR(&pArgs[0]), &output); break;
                        case 103: hr = ToLower(V_BSTR(&pArgs[0]), &output); break;
                        case 104: hr = Reverse(V_BSTR(&pArgs[0]), &output); break;
                    }
                    
                    if (SUCCEEDED(hr) && pVarResult)
                    {
                        V_VT(pVarResult) = VT_BSTR;
                        V_BSTR(pVarResult) = output;
                    }
                    else if (output)
                    {
                        SysFreeString(output);
                    }
                    
                    return hr;
                }
                default:
                    return DISP_E_MEMBERNOTFOUND;
            }
        }
    }
    
    return DISP_E_MEMBERNOTFOUND;
}

HRESULT CDualInterfaceComponent::LoadTypeInfo()
{
    if (m_pTypeInfo)
        return S_OK;
    
    ITypeLib* pTypeLib = nullptr;
    HRESULT hr = LoadRegTypeLib(LIBID_DualInterfaceCOMLib, 1, 0, LOCALE_SYSTEM_DEFAULT, &pTypeLib);
    
    if (FAILED(hr))
    {
        // 尝试从当前模块加载类型库
        WCHAR szPath[MAX_PATH];
        GetModuleFileNameW(GetModuleHandleW(L"SimpleCOM.dll"), szPath, MAX_PATH);
        hr = LoadTypeLib(szPath, &pTypeLib);
    }
    
    if (SUCCEEDED(hr))
    {
        // 首先尝试加载 ICalculator 的类型信息
        hr = pTypeLib->GetTypeInfoOfGuid(IID_ICalculator, &m_pTypeInfo);
        pTypeLib->Release();
    }
    
    return hr;
}

// ICalculator接口实现
STDMETHODIMP CDualInterfaceComponent::Add(long a, long b, long* result)
{
    if (!result)
        return E_POINTER;
    
    *result = a + b;
    return S_OK;
}

STDMETHODIMP  CDualInterfaceComponent::Subtract(long a, long b, long* result)
{
    if (!result)
        return E_POINTER;
    
    *result = a - b;
    return S_OK;
}

STDMETHODIMP CDualInterfaceComponent::Multiply(long a, long b, long* result)
{
    if (!result)
        return E_POINTER;
    
    *result = a * b;
    return S_OK;
}

STDMETHODIMP CDualInterfaceComponent::Divide(long a, long b, long* result)
{
    if (!result)
        return E_POINTER;
    
    if (b == 0)
        return E_INVALIDARG;
    
    *result = a / b;
    return S_OK;
}

// IStringProcessor接口实现
STDMETHODIMP CDualInterfaceComponent::GetLength(BSTR input, long* length)
{
    if (!length)
        return E_POINTER;
    
    if (!input)
    {
        *length = 0;
        return S_OK;
    }
    
    *length = SysStringLen(input);
    return S_OK;
}

STDMETHODIMP CDualInterfaceComponent::ToUpper(BSTR input, BSTR* output)
{
    if (!output)
        return E_POINTER;
    
    *output = nullptr;
    
    if (!input)
        return S_OK;
    
    UINT len = SysStringLen(input);
    *output = SysAllocStringLen(nullptr, len);
    
    if (!*output)
        return E_OUTOFMEMORY;
    
    for (UINT i = 0; i < len; ++i)
    {
        (*output)[i] = towupper(input[i]);
    }
    
    return S_OK;
}

STDMETHODIMP CDualInterfaceComponent::ToLower(BSTR input, BSTR* output)
{
    if (!output)
        return E_POINTER;
    
    *output = nullptr;
    
    if (!input)
        return S_OK;
    
    UINT len = SysStringLen(input);
    *output = SysAllocStringLen(nullptr, len);
    
    if (!*output)
        return E_OUTOFMEMORY;
    
    for (UINT i = 0; i < len; ++i)
    {
        (*output)[i] = towlower(input[i]);
    }
    
    return S_OK;
}

STDMETHODIMP CDualInterfaceComponent::Reverse(BSTR input, BSTR* output)
{
    if (!output)
        return E_POINTER;
    
    *output = nullptr;
    
    if (!input)
        return S_OK;
    
    UINT len = SysStringLen(input);
    *output = SysAllocStringLen(nullptr, len);
    
    if (!*output)
        return E_OUTOFMEMORY;
    
    for (UINT i = 0; i < len; ++i)
    {
        (*output)[i] = input[len - 1 - i];
    }
    
    return S_OK;
}