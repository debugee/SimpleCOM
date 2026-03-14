#pragma once

#include <windows.h>
#include <oleauto.h>
#include "dupatlbase.h"
#include <SimpleCOM_i.h>

class ATL_NO_VTABLE CDualInterfaceComponent :
    public CComObjectRootEx<CComObjectThreadModel>,
    public CComCoClass<CDualInterfaceComponent, &CLSID_DualInterfaceComponent>,
    public ICalculator,
    public IStringProcessor
{ 
public:
    CDualInterfaceComponent()
    {
        m_pTypeInfo = nullptr;
    }

    ~CDualInterfaceComponent()
    {
        if (m_pTypeInfo)
        {
            m_pTypeInfo->Release();
        }
    }
       
    BEGIN_COM_MAP(CDualInterfaceComponent)
        COM_INTERFACE_ENTRY(ICalculator)
        COM_INTERFACE_ENTRY(IStringProcessor)
        COM_INTERFACE_ENTRY2(IDispatch, ICalculator) // 通过 ICalculator 提供 IDispatch
    END_COM_MAP()

    // IDispatch 实现
    STDMETHOD(GetTypeInfoCount)(UINT* pctinfo);
    STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
    STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
    STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

    // ICalculator接口实现
    STDMETHOD(Add)(long a, long b, long* result);
    STDMETHOD(Subtract)(long a, long b, long* result);
    STDMETHOD(Multiply)(long a, long b, long* result);
    STDMETHOD(Divide)(long a, long b, long* result);

    // IStringProcessor接口实现
    STDMETHOD(GetLength)(BSTR input, long* length);
    STDMETHOD(ToUpper)(BSTR input, BSTR* output);
    STDMETHOD(ToLower)(BSTR input, BSTR* output);
    STDMETHOD(Reverse)(BSTR input, BSTR* output);

private:
    ITypeInfo* m_pTypeInfo;
    HRESULT LoadTypeInfo();
};

