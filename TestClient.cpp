#include <windows.h>
#include <stdio.h>
#include <SimpleCOM_i.h>
#include <thread>
#include <condition_variable>
#include <mutex>
#include <iostream>
std::mutex mtx;
std::condition_variable cv;
bool threadFinished(false);

HRESULT testAdd(IDispatch *pDispatch)
{
    HRESULT hr;
    DISPID dispid;
    OLECHAR methodName[] = L"Add"; // 假设插件有一个名为 "Add" 的方法
    LPOLESTR methodNames[1] = {methodName};
    hr = pDispatch->GetIDsOfNames(IID_NULL, methodNames, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr))
    {
        return hr;
    }

    VARIANTARG args[2];
    args[0].vt = VT_I4; // 第二个参数
    args[0].intVal = 10;
    args[1].vt = VT_I4; // 第一个参数
    args[1].intVal = 20;

    DISPPARAMS dispParams = {args, NULL, 2, 0};
    VARIANT result;
    VariantInit(&result);

    hr = pDispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispParams, &result, NULL, NULL);
    if (SUCCEEDED(hr))
    {
        // 输出结果
        if (result.vt == VT_I4)
        {
            std::cout << "Result: " << result.intVal << std::endl;
        }
    }

    VariantClear(&result);

    return hr;
}

HRESULT testMarshalInterface()
{
    CoInitialize(NULL);

    IDispatch *pDispatch = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_DualInterfaceComponent, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void **)&pDispatch);
    if (FAILED(hr))
    {
        CoUninitialize();
        return hr;
    }

    ICalculator *pCalc = nullptr;
    hr = pDispatch->QueryInterface(IID_ICalculator, (void **)&pCalc);
    if (FAILED(hr))
    {
        pDispatch->Release();
        CoUninitialize();
        return hr;
    }

    IStream *pStream = nullptr;
    hr = CreateStreamOnHGlobal(NULL, TRUE, &pStream);
    if (FAILED(hr))
    {
        pCalc->Release();
        pDispatch->Release();
        CoUninitialize();
        return hr;
    }

    hr = CoMarshalInterface(
        pStream,
        IID_ICalculator,
        pCalc,
        MSHCTX_INPROC,   // 封送上下文（本地进程）
        NULL,
        MSHLFLAGS_NORMAL
    );
    if (FAILED(hr))
    {
        pStream->Release();
        pCalc->Release();
        pDispatch->Release();
        CoUninitialize();
        return hr;
    }

    std::thread t1([&pStream, &hr]()
                   {
        CoInitialize(NULL);
        pStream->Seek({0}, STREAM_SEEK_SET, NULL);
        ICalculator *pCalcThread1 = nullptr;
        hr = CoUnmarshalInterface(
            pStream,
            IID_ICalculator,
            (void **)&pCalcThread1);
        if (SUCCEEDED(hr))
        {
            long result = 0;
            hr = pCalcThread1->Add(30, 40, &result);
            if (SUCCEEDED(hr))
            {
                std::cout << "Thread 1 Add Result: " << result << std::endl;
            }
            IDispatch * pDispatch = nullptr;
            hr = pCalcThread1->QueryInterface(IID_IDispatch, (void **)&pDispatch);
            if (SUCCEEDED(hr))
            {
                hr = testAdd(pDispatch);
            }
            IStringProcessor * pStringProcessor = nullptr;
            hr = pCalcThread1->QueryInterface(IID_IStringProcessor, (void **)&pStringProcessor);
            if (SUCCEEDED(hr))
            {
                BSTR input = SysAllocString(L"Hello COM");
                BSTR output = nullptr;
                hr = pStringProcessor->ToUpper(input, &output);
                if (SUCCEEDED(hr))
                {
                    std::wcout << L"ToUpper Result: " << output << std::endl;
                    SysFreeString(output);
                }
                SysFreeString(input);
            }
            if(pStringProcessor)
                pStringProcessor->Release();
            if (pDispatch)
                pDispatch->Release();
            if (pCalcThread1)
                pCalcThread1->Release();
        }

        CoUninitialize(); 
        {
            std::lock_guard<std::mutex> lock(mtx);
            threadFinished = true;
        }
    });

    while (true)
    {
        MSG msg;
        if (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        {
            if (msg.message == WM_QUIT)
                break;
            DispatchMessage(&msg);
            TranslateMessage(&msg);
        }

        {
            std::unique_lock<std::mutex> lock(mtx);
            if (cv.wait_for(lock, std::chrono::milliseconds(0), [] { return threadFinished; })) {
                break;
            }
        }
    };
    t1.join();

    hr = testAdd(pDispatch);
    
    pStream->Release();
    pCalc->Release();
    pDispatch->Release();
    CoUninitialize();

    return hr;
}

HRESULT testDispatch()
{
    CoInitialize(NULL);

    IDispatch *pDispatch = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_DualInterfaceComponent, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void **)&pDispatch);
    if (SUCCEEDED(hr))
    {
        hr = testAdd(pDispatch);
        pDispatch->Release();
    }
    CoUninitialize();

    return hr;
}

HRESULT testOutOfProcess()
{
    CoInitialize(NULL);
    IDispatch *pDispatch = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_DualInterfaceComponent, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pDispatch);
    if (SUCCEEDED(hr))
    {
        testAdd(pDispatch);
        ICalculator *pCalc = nullptr;
        hr = pDispatch->QueryInterface(IID_ICalculator, (void **)&pCalc);
        if (SUCCEEDED(hr))
        {
            long result = 0;
            hr = pCalc->Multiply(6, 7, &result);
            if (SUCCEEDED(hr))
            {
                std::cout << "Multiply Result: " << result << std::endl;
            }
            pCalc->Release();
        }
        pDispatch->Release();
    }

    CoUninitialize();
    return hr;
}

int main()
{
    testDispatch();
    testMarshalInterface();
    testOutOfProcess();
    return 0;
}
