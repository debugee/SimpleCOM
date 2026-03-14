#include <windows.h>
#include <iostream>
#include "dupatlbase.h"

#define INITGUID
#include <guiddef.h>
#undef INITGUID

CAtlModule _Module;

// 被聚合的组件接口
DEFINE_GUID(IID_IInnerComponent,0x12345678, 0x0000, 0x0000, 0xc0,0x00, 0x00,0x00,0x00,0x00,0x00,0x46);
MIDL_INTERFACE("12345678-0000-0000-c000-000000000046")
IInnerComponent : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE DoInnerWork() = 0;
    virtual HRESULT STDMETHODCALLTYPE GetInnerValue(LONG* pValue) = 0;
};
#ifndef _MSC_VER
__CRT_UUID_DECL(IInnerComponent, 0x12345678, 0x0000, 0x0000, 0xc0,0x00, 0x00,0x00,0x00,0x00,0x00,0x46);
#endif

// 外部组件接口
DEFINE_GUID(IID_IOuterComponent,0x12345679, 0x0000, 0x0000, 0xc0,0x00, 0x00,0x00,0x00,0x00,0x00,0x46);
MIDL_INTERFACE("12345679-0000-0000-C000-000000000046")
IOuterComponent : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE DoOuterWork() = 0;
    virtual HRESULT STDMETHODCALLTYPE GetCombinedResult(BSTR* pResult) = 0;
};
#ifndef _MSC_VER
__CRT_UUID_DECL(IOuterComponent, 0x12345679, 0x0000, 0x0000, 0xc0,0x00, 0x00,0x00,0x00,0x00,0x00,0x46);
#endif

class CInnerComponent : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public IInnerComponent
{
public:
    CInnerComponent() : m_nValue(42) 
    {
    }
    
    virtual ~CInnerComponent()
    {
    }

    BEGIN_COM_MAP(CInnerComponent)
        COM_INTERFACE_ENTRY(IInnerComponent)
    END_COM_MAP()

    STDMETHOD(DoInnerWork)() override
    {
        m_nValue++;
        return S_OK;
    }

    STDMETHOD(GetInnerValue)(LONG* pValue) override
    {
        if (!pValue) return E_POINTER;
        *pValue = m_nValue;
        return S_OK;
    }

    DECLARE_AGGREGATABLE(CInnerComponent)

private:
    LONG m_nValue;
};

// 外部组件实现（聚合内部组件）
class COuterComponent : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public IOuterComponent
{
public:
    COuterComponent() : m_pInnerUnk(nullptr)
    {
    }

    virtual ~COuterComponent()
    {
    }

    HRESULT FinalConstruct()
    {
        // 外部组件创建后再创建被聚合的内部组件
        HRESULT hr = CInnerComponent::_CreatorClass::CreateInstance(
            GetUnknown(),    // 传递外部IUnknown,给内部转发
            IID_IUnknown, 
            (void**)&m_pInnerUnk
        );
        return hr;
    }

    void FinalRelease()
    {
        if (m_pInnerUnk)
        {
            m_pInnerUnk->Release();
            m_pInnerUnk = nullptr;
        }
    }

    BEGIN_COM_MAP(COuterComponent)
        COM_INTERFACE_ENTRY(IOuterComponent)
        COM_INTERFACE_ENTRY_AGGREGATE(IID_IInnerComponent, m_pInnerUnk) // 聚合内部组件
    END_COM_MAP()

    STDMETHOD(DoOuterWork)() override
    {
        IInnerComponent* pInner = nullptr;
        HRESULT hr = m_pInnerUnk->QueryInterface(__uuidof(IInnerComponent), (void**)&pInner);
        if (SUCCEEDED(hr))
        {
            pInner->DoInnerWork();
            pInner->Release();
        }
        
        return S_OK;
    }

    STDMETHOD(GetCombinedResult)(BSTR* pResult) override
    {
        if (!pResult) return E_POINTER;
        
        // 获取内部组件的值
        LONG innerValue = 0;
        IInnerComponent* pInner = nullptr;
        HRESULT hr = m_pInnerUnk->QueryInterface(__uuidof(IInnerComponent), (void**)&pInner);
        if (SUCCEEDED(hr))
        {
            pInner->GetInnerValue(&innerValue);
            pInner->Release();
        }
        
        WCHAR buffer[256];
        swprintf_s(buffer, L"外部组件结果，包含内部值: %d", innerValue);
        *pResult = SysAllocString(buffer);
        
        return S_OK;
    }

    DECLARE_NOT_AGGREGATABLE(COuterComponent)

private:
    IUnknown* m_pInnerUnk;//控制内部时使用,不提供给外部
};

int main()
{
    CoInitialize(NULL);
    CComObject<COuterComponent>* pOuter = nullptr;
    HRESULT hr = CComObject<COuterComponent>::CreateInstance(&pOuter);
    if (FAILED(hr))
    {
        std::wcout << L"创建外部组件失败" << std::endl;
        return -1;
    }
    pOuter->AddRef();

    IInnerComponent *pInnerInterface = nullptr;
    hr = pOuter->QueryInterface(&pInnerInterface);
    pInnerInterface->Release();

    IOuterComponent *pOuterInterface = nullptr;
    hr = pOuter->QueryInterface(&pOuterInterface);
    pOuterInterface->Release();

    pOuter->Release();
    CoUninitialize();

    return 0;
}