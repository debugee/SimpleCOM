#ifndef DUPATLBASE_H
#define DUPATLBASE_H

#include <windows.h>
#include <assert.h>

#ifndef _ATL_DISABLE_NOTHROW_NEW
#include <new.h>
#define _ATL_NEW		new(std::nothrow)
#else
#define _ATL_NEW		new
#endif

#ifndef ATL_NO_VTABLE
#define ATL_NO_VTABLE
#endif

inline BOOL WINAPI InlineIsEqualUnknown(REFGUID rguid1)
{
   return (
	  ((PLONG) &rguid1)[0] == 0 &&
	  ((PLONG) &rguid1)[1] == 0 &&
#ifdef _ATL_BYTESWAP
	  ((PLONG) &rguid1)[2] == 0xC0000000 &&
	  ((PLONG) &rguid1)[3] == 0x00000046);
#else
	  ((PLONG) &rguid1)[2] == 0x000000C0 &&
	  ((PLONG) &rguid1)[3] == 0x46000000);
#endif
}

class CComCriticalSection
{
public:
	CComCriticalSection() throw()
	{
		memset(&m_sec, 0, sizeof(CRITICAL_SECTION));
	}

	~CComCriticalSection()
	{
	}

	HRESULT Lock() throw()
	{
		EnterCriticalSection(&m_sec);
		return S_OK;
	}
	HRESULT Unlock() throw()
	{
		LeaveCriticalSection(&m_sec);
		return S_OK;
	}
	HRESULT Init() throw()
	{
		HRESULT hRes = S_OK;
		if (!InitializeCriticalSectionEx(&m_sec, 0, 0))
		{
			hRes = HRESULT_FROM_WIN32(GetLastError());
		}

		return hRes;
	}

	HRESULT Term() throw()
	{
		DeleteCriticalSection(&m_sec);
		return S_OK;
	}
	CRITICAL_SECTION m_sec;
};

class CComAutoCriticalSection :
	public CComCriticalSection
{
public:
	CComAutoCriticalSection()
	{
		HRESULT hr = CComCriticalSection::Init();
		// if (FAILED(hr))
		// 	AtlThrow(hr);
	}
	~CComAutoCriticalSection() throw()
	{
		CComCriticalSection::Term();
	}
private :
	HRESULT Init(); // Not implemented. CComAutoCriticalSection::Init should never be called
	HRESULT Term(); // Not implemented. CComAutoCriticalSection::Term should never be called
};

class CComSafeDeleteCriticalSection :
	public CComCriticalSection
{
public:
	CComSafeDeleteCriticalSection(): m_bInitialized(false)
	{
	}

	~CComSafeDeleteCriticalSection() throw()
	{
		if (!m_bInitialized)
		{
			return;
		}
		m_bInitialized = false;
		CComCriticalSection::Term();
	}

	HRESULT Init() throw()
	{
		HRESULT hr = CComCriticalSection::Init();
		if (SUCCEEDED(hr))
		{
			m_bInitialized = true;
		}
		return hr;
	}

	HRESULT Term() throw()
	{
		if (!m_bInitialized)
		{
			return S_OK;
		}
		m_bInitialized = false;
		return CComCriticalSection::Term();
	}

	HRESULT Lock()
	{
		return CComCriticalSection::Lock();
	}

private:
	bool m_bInitialized;
};

class CComAutoDeleteCriticalSection :
	public CComSafeDeleteCriticalSection
{
private:
	// CComAutoDeleteCriticalSection::Term should never be called
	HRESULT Term() throw();
};

class CComFakeCriticalSection
{
public:
	HRESULT Lock() throw()
	{
		return S_OK;
	}
	HRESULT Unlock() throw()
	{
		return S_OK;
	}
	HRESULT Init() throw()
	{
		return S_OK;
	}
	HRESULT Term() throw()
	{
		return S_OK;
	}
};

class CComMultiThreadModelNoCS
{
public:
	static ULONG WINAPI Increment(_Inout_ LPLONG p) throw()
	{
		return ::InterlockedIncrement(p);
	}
	static ULONG WINAPI Decrement(_Inout_ LPLONG p) throw()
	{
		return ::InterlockedDecrement(p);
	}
	typedef CComFakeCriticalSection AutoCriticalSection;
	typedef CComFakeCriticalSection AutoDeleteCriticalSection;
	typedef CComFakeCriticalSection CriticalSection;
	typedef CComMultiThreadModelNoCS ThreadModelNoCS;
};

class CComMultiThreadModel
{
public:
	static ULONG WINAPI Increment(_Inout_ LPLONG p) throw()
	{
		return ::InterlockedIncrement(p);
	}
	static ULONG WINAPI Decrement(_Inout_ LPLONG p) throw()
	{
		return ::InterlockedDecrement(p);
	}
	typedef CComAutoCriticalSection AutoCriticalSection;
	typedef CComAutoDeleteCriticalSection AutoDeleteCriticalSection;
	typedef CComCriticalSection CriticalSection;
	typedef CComMultiThreadModelNoCS ThreadModelNoCS;
};

class CComSingleThreadModel
{
public:
	static ULONG WINAPI Increment(_Inout_ LPLONG p) throw()
	{
		return ++(*p);
	}
	static ULONG WINAPI Decrement(_Inout_ LPLONG p) throw()
	{
		return --(*p);
	}
	typedef CComFakeCriticalSection AutoCriticalSection;
	typedef CComFakeCriticalSection AutoDeleteCriticalSection;
	typedef CComFakeCriticalSection CriticalSection;
	typedef CComSingleThreadModel ThreadModelNoCS;
};

#if defined(_ATL_SINGLE_THREADED)

#if defined(_ATL_APARTMENT_THREADED) || defined(_ATL_FREE_THREADED)
#pragma message ("More than one global threading model defined.")
#endif

	typedef CComSingleThreadModel CComObjectThreadModel;
	typedef CComSingleThreadModel CComGlobalsThreadModel;

#elif defined(_ATL_APARTMENT_THREADED)

#if defined(_ATL_SINGLE_THREADED) || defined(_ATL_FREE_THREADED)
#pragma message ("More than one global threading model defined.")
#endif

	typedef CComSingleThreadModel CComObjectThreadModel;
	typedef CComMultiThreadModel CComGlobalsThreadModel;

#elif defined(_ATL_FREE_THREADED)

#if defined(_ATL_SINGLE_THREADED) || defined(_ATL_APARTMENT_THREADED)
#pragma message ("More than one global threading model defined.")
#endif

	typedef CComMultiThreadModel CComObjectThreadModel;
	typedef CComMultiThreadModel CComGlobalsThreadModel;

#else
#pragma message ("No global threading model defined")
#endif

#ifndef _ATL_STATIC_LIB_IMPL

class CAtlModule;
__declspec(selectany) CAtlModule* _pAtlModule = NULL;

class CAtlModule
{
public:
	CAtlModule() throw()
	{
		// Should have only one instance of a class
		// derived from CAtlModule in a project.
		//ATLASSERT(_pAtlModule == NULL);
		m_nLockCnt = 0;
		_pAtlModule = this;
	}

	void Term() throw()
	{

	}

	virtual ~CAtlModule() throw()
	{
		Term();
	}

	virtual LONG Lock() throw()
	{
		return CComGlobalsThreadModel::Increment(&m_nLockCnt);
	}

	virtual LONG Unlock() throw()
	{
		return CComGlobalsThreadModel::Decrement(&m_nLockCnt);
	}

	virtual LONG GetLockCount() throw()
	{
		return m_nLockCnt;
	}
private:
    LONG m_nLockCnt;
};

#endif // _ATL_STATIC_LIB_IMPL

typedef HRESULT (WINAPI _ATL_CREATORARGFUNC)(
	void* pv,
	REFIID riid,
	LPVOID* ppv,
	DWORD_PTR dw);

#define _ATL_SIMPLEMAPENTRY ((_ATL_CREATORARGFUNC*)1)

struct _ATL_INTMAP_ENTRY
{
	const IID* piid;       // the interface id (IID)
	DWORD_PTR dw;
	_ATL_CREATORARGFUNC* pFunc; //NULL:end, 1:offset, n:ptr
};

#define ATLAPI __declspec(nothrow) HRESULT __stdcall

ATLAPI AtlInternalQueryInterface(
	void* pThis,
	const _ATL_INTMAP_ENTRY* pEntries,
	 REFIID iid,
	void** ppvObject);

class CComObjectRootBase
{
public:
    CComObjectRootBase(){
		m_dwRef = 0L;
	}
    ~CComObjectRootBase(){};

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	HRESULT _AtlFinalConstruct()
	{
		return S_OK;
	}
	void FinalRelease()
	{

	}

	void _AtlFinalRelease()
	{

	}

	static HRESULT WINAPI InternalQueryInterface(
		void* pThis,
		const _ATL_INTMAP_ENTRY* pEntries,
		REFIID iid,
		void** ppvObject)
	{
		//ATLASSERT(pThis != NULL);
		// First entry in the com map should be a simple map entry
		//ATLASSERT(pEntries->pFunc == _ATL_SIMPLEMAPENTRY);
	
		HRESULT hRes = AtlInternalQueryInterface(pThis, pEntries, iid, ppvObject);
		return hRes;
	}

	ULONG OuterAddRef()
	{
		return m_pOuterUnknown->AddRef();
	}
	ULONG OuterRelease()
	{
		return m_pOuterUnknown->Release();
	}
	HRESULT OuterQueryInterface(
		REFIID iid,
		void** ppvObject)
	{
		return m_pOuterUnknown->QueryInterface(iid, ppvObject);
	}
	void SetVoid(void*)
	{
	}
	void InternalFinalConstructAddRef()
	{
	}
	void InternalFinalConstructRelease()
	{
		//ATLASSUME(m_dwRef == 0);
	}

	static HRESULT WINAPI _Delegate(
	void* pv,
	REFIID iid,
	void** ppvObject,
	DWORD_PTR dw)
	{
		HRESULT hRes = E_NOINTERFACE;
		IUnknown* p = *(IUnknown**)((DWORD_PTR)pv + dw);
		*ppvObject = NULL;
		if (p != NULL)
			hRes = p->QueryInterface(iid, ppvObject);
		return hRes;
	}
protected:
    union
	{
		long m_dwRef;
		IUnknown* m_pOuterUnknown;
	};
};

template <class ThreadModel>
class CComObjectRootEx;

template <class ThreadModel>
class CComObjectLockT
{
public:
	CComObjectLockT(CComObjectRootEx<ThreadModel>* p)
	{
		if (p)
			p->Lock();
		m_p = p;
	}

	~CComObjectLockT()
	{
		if (m_p)
			m_p->Unlock();
	}
	CComObjectRootEx<ThreadModel>* m_p;
};

template<class ThreadModel>
class CComObjectRootEx : public CComObjectRootBase
{
public:
	typedef ThreadModel _ThreadModel;
	typedef typename _ThreadModel::AutoCriticalSection _CritSec;
	typedef typename _ThreadModel::AutoDeleteCriticalSection _AutoDelCritSec;
	typedef CComObjectLockT<_ThreadModel> ObjectLock;

	~CComObjectRootEx()
	{
	}

	ULONG InternalAddRef()
	{
		//ATLASSUME(m_dwRef != -1L);
		return _ThreadModel::Increment(&m_dwRef);
	}
	ULONG InternalRelease()
	{
		return _ThreadModel::Decrement(&m_dwRef);
	}

	HRESULT _AtlInitialConstruct()
	{
		return m_critsec.Init();
	}
	void Lock()
	{
		m_critsec.Lock();
	}
	void Unlock()
	{
		m_critsec.Unlock();
	}
private:
	_AutoDelCritSec m_critsec;
};

template <> class CComObjectLockT<CComSingleThreadModel>;

template <>
class CComObjectLockT<CComSingleThreadModel>
{
public:
	CComObjectLockT(CComObjectRootEx<CComSingleThreadModel>*)
	{
	}
	~CComObjectLockT()
	{
	}
};

template <>
class CComObjectRootEx<CComSingleThreadModel> :
	public CComObjectRootBase
{
public:
	typedef CComSingleThreadModel _ThreadModel;
	typedef _ThreadModel::AutoCriticalSection _CritSec;
	typedef _ThreadModel::AutoDeleteCriticalSection _AutoDelCritSec;
	typedef CComObjectLockT<_ThreadModel> ObjectLock;

	~CComObjectRootEx() {}

	ULONG InternalAddRef()
	{
		return _ThreadModel::Increment(&m_dwRef);
	}
	ULONG InternalRelease()
	{
		return _ThreadModel::Decrement(&m_dwRef);
	}

	HRESULT _AtlInitialConstruct()
	{
		return S_OK;
	}

	void Lock()
	{
	}
	void Unlock()
	{
	}
};

typedef CComObjectRootEx<CComObjectThreadModel> CComObjectRoot;

typedef HRESULT (WINAPI _ATL_CREATORFUNC)(
	void* pv,
	REFIID riid,
	LPVOID* ppv);

#ifndef _ATL_STATIC_LIB_IMPL

class CComClassFactory :
	public IClassFactory,
	public CComObjectRootEx<CComGlobalsThreadModel>
{
public:
	// BEGIN_COM_MAP(CComClassFactory)
	// 	COM_INTERFACE_ENTRY(IClassFactory)
	// END_COM_MAP()
    HRESULT _InternalQueryInterface(REFIID iid, void** ppvObject)
    {
        if (ppvObject == nullptr)
            return E_POINTER;

        *ppvObject = nullptr;

        if (IsEqualIID(iid, IID_IUnknown))
        {
            IUnknown* p = static_cast<IUnknown*>(this);
            p->AddRef();
            *ppvObject = p;
        }
        else if (IsEqualIID(iid, IID_IClassFactory))
        {
            IClassFactory *p = static_cast<IClassFactory*>(this);
            p->AddRef();
            *ppvObject = p;
        }
        else
        {
            return E_NOINTERFACE;
        }

        return S_OK;
    }

	virtual ~CComClassFactory()
	{
	}

	// IClassFactory
	STDMETHOD(CreateInstance)(
		LPUNKNOWN pUnkOuter,
		REFIID riid,
		void** ppvObj)
	{
		//ATLASSUME(m_pfnCreateInstance != NULL);
		HRESULT hRes = E_POINTER;
		if (ppvObj != NULL)
		{
			*ppvObj = NULL;
			// can't ask for anything other than IUnknown when aggregating

			if ((pUnkOuter != NULL) && !InlineIsEqualUnknown(riid))
			{
				//ATLTRACE(atlTraceCOM, 0, _T("CComClassFactory: asked for non IUnknown interface while creating an aggregated object"));
				hRes = CLASS_E_NOAGGREGATION;
			}
			else
				hRes = m_pfnCreateInstance(pUnkOuter, riid, ppvObj);
		}
		return hRes;
	}

	STDMETHOD(LockServer)(BOOL fLock)
	{
		if (fLock)
			_pAtlModule->Lock();
		else
			_pAtlModule->Unlock();
		return S_OK;
	}
	// helper
	void SetVoid(void* pv)
	{
		m_pfnCreateInstance = (_ATL_CREATORFUNC*)pv;
	}
	_ATL_CREATORFUNC* m_pfnCreateInstance;
};

#endif // _ATL_STATIC_LIB_IMPL

#define DECLARE_PROTECT_FINAL_CONSTRUCT()\
	void InternalFinalConstructAddRef() {InternalAddRef();}\
	void InternalFinalConstructRelease() {InternalRelease();}

template <class T1>
class CComCreator
{
public:
	static HRESULT WINAPI CreateInstance(
		void* pv,
		REFIID riid,
		LPVOID* ppv)
	{
		//ATLASSERT(ppv != NULL);
		if (ppv == NULL)
			return E_POINTER;
		*ppv = NULL;

		HRESULT hRes = E_OUTOFMEMORY;
		T1* p = NULL;

		p = _ATL_NEW T1(pv);

		if (p != NULL)
		{
			p->SetVoid(pv);
			p->InternalFinalConstructAddRef();
			hRes = p->_AtlInitialConstruct();
			if (SUCCEEDED(hRes))
				hRes = p->FinalConstruct();
			if (SUCCEEDED(hRes))
				hRes = p->_AtlFinalConstruct();
			p->InternalFinalConstructRelease();
			if (hRes == S_OK)
			{
				hRes = p->QueryInterface(riid, ppv);
			}
			if (hRes != S_OK)
				delete p;
		}
		return hRes;
	}
};

template <class T1>
class CComInternalCreator
{
public:
	static HRESULT WINAPI CreateInstance(
		 void* pv,
		REFIID riid,
		LPVOID* ppv)
	{
		//ATLASSERT(ppv != NULL);
		if (ppv == NULL)
			return E_POINTER;
		*ppv = NULL;

		HRESULT hRes = E_OUTOFMEMORY;
		T1* p = NULL;

		p = _ATL_NEW T1(pv);

		if (p != NULL)
		{
			p->SetVoid(pv);
			p->InternalFinalConstructAddRef();
			hRes = p->_AtlInitialConstruct();
			if (SUCCEEDED(hRes))
				hRes = p->FinalConstruct();
			if (SUCCEEDED(hRes))
				hRes = p->_AtlFinalConstruct();
			p->InternalFinalConstructRelease();
			if (hRes == S_OK)
				hRes = p->_InternalQueryInterface(riid, ppv);
			if (hRes != S_OK)
				delete p;
		}
		return hRes;
	}
};

template <HRESULT hr>
class CComFailCreator
{
public:
	static 
	HRESULT WINAPI CreateInstance(
		void*,
		REFIID,
		LPVOID* ppv)
	{
		if (ppv == NULL)
			return E_POINTER;
		*ppv = NULL;

		return hr;
	}
};

template <class T1, class T2>
class CComCreator2
{
public:
	static HRESULT WINAPI CreateInstance(
		void* pv,
		REFIID riid,
		LPVOID* ppv)
	{
		//ATLASSERT(ppv != NULL);

		return (pv == NULL) ?
			T1::CreateInstance(NULL, riid, ppv) :
			T2::CreateInstance(pv, riid, ppv);
	}
};

#define DECLARE_NOT_AGGREGATABLE(x) public:\
	typedef CComCreator2< CComCreator< CComObject< x > >, CComFailCreator<CLASS_E_NOAGGREGATION> > _CreatorClass;
#define DECLARE_AGGREGATABLE(x) public:\
	typedef CComCreator2< CComCreator< CComObject< x > >, CComCreator< CComAggObject< x > > > _CreatorClass;
#define DECLARE_ONLY_AGGREGATABLE(x) public:\
	typedef CComCreator2< CComFailCreator<E_FAIL>, CComCreator< CComAggObject< x > > > _CreatorClass;
#define DECLARE_POLY_AGGREGATABLE(x) public:\
	typedef CComCreator< CComPolyObject< x > > _CreatorClass;


// ModuleLockHelper handles unlocking module while going out of scope
class ModuleLockHelper
{
public:
	ModuleLockHelper()
	{
#ifndef _ATL_STATIC_LIB_IMPL
		_pAtlModule->Lock();
#endif
	}

	~ModuleLockHelper()
	{
#ifndef _ATL_STATIC_LIB_IMPL
		_pAtlModule->Unlock();
#endif
	}
};

//Base is the user's class that derives from CComObjectRoot and whatever
//interfaces the user wants to support on the object
template <class Base>
class CComObject :
	public Base
{
public:
	typedef Base _BaseClass;
	CComObject(void* = NULL)
	{
		_pAtlModule->Lock();
	}
	// Set refcount to -(LONG_MAX/2) to protect destruction and
	// also catch mismatched Release in debug builds
	virtual ~CComObject()
	{
		this->m_dwRef = -(LONG_MAX/2);
		this->FinalRelease();
		_pAtlModule->Unlock();
	}
	//If InternalAddRef or InternalRelease is undefined then your class
	//doesn't derive from CComObjectRoot
	STDMETHOD_(ULONG, AddRef)()
	{
		return this->InternalAddRef();
	}
	STDMETHOD_(ULONG, Release)()
	{
		ULONG l = this->InternalRelease();
		if (l == 0)
		{
			// Lock the module to avoid DLL unload when destruction of member variables take a long time
			ModuleLockHelper lock;
			delete this;
		}
		return l;
	}
	//if _InternalQueryInterface is undefined then you forgot BEGIN_COM_MAP
	STDMETHOD(QueryInterface)(
		REFIID iid,
		void** ppvObject) throw()
	{
		return this->_InternalQueryInterface(iid, ppvObject);
	}
	template <class Q>
	HRESULT STDMETHODCALLTYPE QueryInterface(
		Q** pp) throw()
	{
		return QueryInterface(__uuidof(Q), (void**)pp);
	}

	static HRESULT WINAPI CreateInstance(CComObject<Base>** pp) throw();
};

template <class Base>
HRESULT WINAPI CComObject<Base>::CreateInstance(
	CComObject<Base>** pp) throw()
{
	//ATLASSERT(pp != NULL);
	if (pp == NULL)
		return E_POINTER;
	*pp = NULL;

	HRESULT hRes = E_OUTOFMEMORY;
	CComObject<Base>* p = NULL;
	p = _ATL_NEW CComObject<Base>();
	if (p != NULL)
	{
		p->SetVoid(NULL);
		p->InternalFinalConstructAddRef();
		hRes = p->_AtlInitialConstruct();
		if (SUCCEEDED(hRes))
			hRes = p->FinalConstruct();
		if (SUCCEEDED(hRes))
			hRes = p->_AtlFinalConstruct();
		p->InternalFinalConstructRelease();
		if (hRes != S_OK)
		{
			delete p;
			p = NULL;
		}
	}
	*pp = p;
	return hRes;
}


//Base is the user's class that derives from CComObjectRoot and whatever
//interfaces the user wants to support on the object
// CComObjectCached is used primarily for class factories in DLL's
// but it is useful anytime you want to cache an object
template <class Base>
class CComObjectCached :
	public Base
{
public:
	typedef Base _BaseClass;
	CComObjectCached(void* = NULL)
	{
	}
	// Set refcount to -(LONG_MAX/2) to protect destruction and
	// also catch mismatched Release in debug builds
	virtual ~CComObjectCached()
	{
		this->m_dwRef = -(LONG_MAX/2);
		this->FinalRelease();
	}
	//If InternalAddRef or InternalRelease is undefined then your class
	//doesn't derive from CComObjectRoot
	STDMETHOD_(ULONG, AddRef)() throw()
	{
		ULONG l = this->InternalAddRef();
		if (l == 2)
			_pAtlModule->Lock();
		return l;
	}
	STDMETHOD_(ULONG, Release)() throw()
	{
		ULONG l = this->InternalRelease();
		if (l == 0)
			delete this;
		else if (l == 1)
			_pAtlModule->Unlock();
		return l;
	}
	//if _InternalQueryInterface is undefined then you forgot BEGIN_COM_MAP
	STDMETHOD(QueryInterface)(
		REFIID iid,
		void** ppvObject) throw()
	{
		return this->_InternalQueryInterface(iid, ppvObject);
	}
	static HRESULT WINAPI CreateInstance(
		CComObjectCached<Base>** pp) throw();
};

template <class Base>
HRESULT WINAPI CComObjectCached<Base>::CreateInstance(
	CComObjectCached<Base>** pp) throw()
{
	if (pp == NULL)
		return E_POINTER;
	*pp = NULL;

	HRESULT hRes = E_OUTOFMEMORY;
	CComObjectCached<Base>* p = NULL;
	p = _ATL_NEW CComObjectCached<Base>();
	if (p != NULL)
	{
		p->SetVoid(NULL);
		p->InternalFinalConstructAddRef();
		hRes = p->_AtlInitialConstruct();
		if (SUCCEEDED(hRes))
			hRes = p->FinalConstruct();
		if (SUCCEEDED(hRes))
			hRes = p->_AtlFinalConstruct();
		p->InternalFinalConstructRelease();
		if (hRes != S_OK)
		{
			delete p;
			p = NULL;
		}
	}
	*pp = p;
	return hRes;
}


//Base is the user's class that derives from CComObjectRoot and whatever
//interfaces the user wants to support on the object
template <class Base>
class CComObjectNoLock :
	public Base
{
public:
	typedef Base _BaseClass;
	CComObjectNoLock(void* = NULL)
	{
	}
	// Set refcount to -(LONG_MAX/2) to protect destruction and
	// also catch mismatched Release in debug builds

	virtual ~CComObjectNoLock()
	{
		this->m_dwRef = -(LONG_MAX/2);
		this->FinalRelease();
	}

	//If InternalAddRef or InternalRelease is undefined then your class
	//doesn't derive from CComObjectRoot
	STDMETHOD_(ULONG, AddRef)() throw()
	{
		return this->InternalAddRef();
	}
	STDMETHOD_(ULONG, Release)() throw()
	{
		ULONG l = this->InternalRelease();
		if (l == 0)
			delete this;
		return l;
	}
	//if _InternalQueryInterface is undefined then you forgot BEGIN_COM_MAP
	STDMETHOD(QueryInterface)(
		REFIID iid,
		void** ppvObject) throw()
	{
		return this->_InternalQueryInterface(iid, ppvObject);
	}
};


// It is possible for Base not to derive from CComObjectRoot
// However, you will need to provide _InternalQueryInterface
template <class Base>
class CComObjectGlobal :
	public Base
{
public:
	typedef Base _BaseClass;
	CComObjectGlobal(void* = NULL)
	{
		m_hResFinalConstruct = S_OK;
		this->InternalFinalConstructAddRef();
		m_hResFinalConstruct = this->_AtlInitialConstruct();
		if (SUCCEEDED(m_hResFinalConstruct))
			m_hResFinalConstruct = this->FinalConstruct();
		this->InternalFinalConstructRelease();
	}
	virtual ~CComObjectGlobal()
	{
		this->FinalRelease();
	}

	STDMETHOD_(ULONG, AddRef)() throw()
	{
		return _pAtlModule->Lock();
	}
	STDMETHOD_(ULONG, Release)() throw()
	{
		return _pAtlModule->Unlock();
	}
	STDMETHOD(QueryInterface)(
		REFIID iid,
		void** ppvObject) throw()
	{
		return this->_InternalQueryInterface(iid, ppvObject);
	}
	HRESULT m_hResFinalConstruct;
};

// It is possible for Base not to derive from CComObjectRoot
// However, you will need to provide FinalConstruct and InternalQueryInterface
template <class Base>
class CComObjectStack :
	public Base
{
public:
	typedef Base _BaseClass;
	CComObjectStack(void* = NULL)
	{
		m_hResFinalConstruct = this->_AtlInitialConstruct();
		if (SUCCEEDED(m_hResFinalConstruct))
			m_hResFinalConstruct = this->FinalConstruct();
	}
	virtual ~CComObjectStack()
	{
		this->FinalRelease();
	}

	STDMETHOD_(ULONG, AddRef)() throw()
	{
		//ATLASSERT(FALSE);
		return 0;
	}
	STDMETHOD_(ULONG, Release)() throw()
	{
		//ATLASSERT(FALSE);
		return 0;
	}
	STDMETHOD(QueryInterface)(
		REFIID,
		void** ppvObject) throw()
	{
		//ATLASSERT(FALSE);
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}
	HRESULT m_hResFinalConstruct;
};

// Base must be derived from CComObjectRoot
template <class Base>
class CComObjectStackEx :
	public Base
{
public:
	typedef Base _BaseClass;

	CComObjectStackEx(void* = NULL)
	{
#ifndef NDEBUG
		this->m_dwRef = 0;
#endif
		m_hResFinalConstruct = this->_AtlInitialConstruct();
		if (SUCCEEDED(m_hResFinalConstruct))
			m_hResFinalConstruct = this->FinalConstruct();
	}

	virtual ~CComObjectStackEx()
	{
		// This assert indicates mismatched ref counts.
		//
		// The ref count has no control over the
		// lifetime of this object, so you must ensure
		// by some other means that the object remains
		// alive while clients have references to its interfaces.
		//ATLASSUME(this->m_dwRef == 0);
		this->FinalRelease();
	}

	STDMETHOD_(ULONG, AddRef)() throw()
	{
#ifndef NDEBUG
		return this->InternalAddRef();
#else
		return 0;
#endif
	}

	STDMETHOD_(ULONG, Release)() throw()
	{
#ifndef NDEBUG
		return this->InternalRelease();
#else
		return 0;
#endif
	}

	STDMETHOD(QueryInterface)(
		REFIID iid,
		void** ppvObject) throw()
	{
		return this->_InternalQueryInterface(iid, ppvObject);
	}

	HRESULT m_hResFinalConstruct;
};

template <class Base> //Base must be derived from CComObjectRoot
class CComContainedObject :
	public Base
{
public:
	typedef Base _BaseClass;
	CComContainedObject(void* pv)
	{
		this->m_pOuterUnknown = (IUnknown*)pv;
	}
	STDMETHOD_(ULONG, AddRef)() throw()
	{
		return this->OuterAddRef();
	}
	STDMETHOD_(ULONG, Release)() throw()
	{
		return this->OuterRelease();
	}
	STDMETHOD(QueryInterface)(
		REFIID iid,
		void** ppvObject) throw()
	{
		return this->OuterQueryInterface(iid, ppvObject);
	}
	template <class Q>
	HRESULT STDMETHODCALLTYPE QueryInterface(
		Q** pp)
	{
		return QueryInterface(__uuidof(Q), (void**)pp);
	}
	//GetControllingUnknown may be virtual if the Base class has declared
	//DECLARE_GET_CONTROLLING_UNKNOWN()
	IUnknown* GetControllingUnknown() throw()
	{
		return this->m_pOuterUnknown;
	}
};

//contained is the user's class that derives from CComObjectRoot and whatever
//interfaces the user wants to support on the object
template <class contained>
class CComAggObject :
	public IUnknown,
	public CComObjectRootEx<typename contained::_ThreadModel::ThreadModelNoCS>
{
private:
	typedef CComObjectRootEx<typename contained::_ThreadModel::ThreadModelNoCS> _MyCComObjectRootEx;

public:
	typedef contained _BaseClass;
	CComAggObject(void* pv) :
		m_contained(pv)
	{
		_pAtlModule->Lock();
	}
	HRESULT _AtlInitialConstruct()
	{
		HRESULT hr = m_contained._AtlInitialConstruct();
		if (SUCCEEDED(hr))
		{
			hr = _MyCComObjectRootEx::_AtlInitialConstruct();
		}
		return hr;
	}
	//If you get a message that this call is ambiguous then you need to
	// override it in your class and call each base class' version of this
	HRESULT FinalConstruct()
	{
		_MyCComObjectRootEx::FinalConstruct();
		return m_contained.FinalConstruct();
	}
	void FinalRelease()
	{
		_MyCComObjectRootEx::FinalRelease();
		m_contained.FinalRelease();
	}
	// Set refcount to -(LONG_MAX/2) to protect destruction and
	// also catch mismatched Release in debug builds
	virtual ~CComAggObject()
	{
		this->m_dwRef = -(LONG_MAX/2);
		FinalRelease();
		_pAtlModule->Unlock();
	}

	STDMETHOD_(ULONG, AddRef)()
	{
		return this->InternalAddRef();
	}
	STDMETHOD_(ULONG, Release)()
	{
		ULONG l = this->InternalRelease();
		if (l == 0)
		{
			ModuleLockHelper lock;
			delete this;
		}
		return l;
	}
	STDMETHOD(QueryInterface)(
		REFIID iid,
		void** ppvObject)
	{
		//ATLASSERT(ppvObject != NULL);
		if (ppvObject == NULL)
			return E_POINTER;
		*ppvObject = NULL;

		HRESULT hRes = S_OK;
		if (InlineIsEqualUnknown(iid))
		{
			*ppvObject = (void*)(IUnknown*)this;
			AddRef();
		}
		else
			hRes = m_contained._InternalQueryInterface(iid, ppvObject);
		return hRes;
	}
	template <class Q>
	HRESULT STDMETHODCALLTYPE QueryInterface(Q** pp)
	{
		return QueryInterface(__uuidof(Q), (void**)pp);
	}
	static HRESULT WINAPI CreateInstance(
		LPUNKNOWN pUnkOuter,
		CComAggObject<contained>** pp)
	{
		//ATLASSERT(pp != NULL);
		if (pp == NULL)
			return E_POINTER;
		*pp = NULL;

		HRESULT hRes = E_OUTOFMEMORY;
		CComAggObject<contained>* p = NULL;
		p = _ATL_NEW CComAggObject<contained>(pUnkOuter);
		if (p != NULL)
		{
			p->SetVoid(NULL);
			p->InternalFinalConstructAddRef();
			hRes = p->_AtlInitialConstruct();
			if (SUCCEEDED(hRes))
				hRes = p->FinalConstruct();
			if (SUCCEEDED(hRes))
				hRes = p->_AtlFinalConstruct();
			p->InternalFinalConstructRelease();
			if (hRes != S_OK)
			{
				delete p;
				p = NULL;
			}
		}
		*pp = p;
		return hRes;
	}

	CComContainedObject<contained> m_contained;
};

#if defined(_WINDLL) | defined(_USRDLL)
#define DECLARE_CLASSFACTORY_EX(cf) typedef CComCreator< CComObjectCached< cf > > _ClassFactoryCreatorClass;
#else
// don't let class factory refcount influence lock count
#define DECLARE_CLASSFACTORY_EX(cf) typedef CComCreator< CComObjectNoLock< cf > > _ClassFactoryCreatorClass;
#endif
#define DECLARE_CLASSFACTORY() DECLARE_CLASSFACTORY_EX(CComClassFactory)

#ifndef _ATL_STATIC_LIB_IMPL

template <class T, const CLSID* pclsid = &CLSID_NULL>
class CComCoClass
{
public:
	DECLARE_CLASSFACTORY()
	DECLARE_AGGREGATABLE(T)
	typedef T _CoClass;
	static const CLSID& WINAPI GetObjectCLSID()
	{
		return *pclsid;
	}
	static LPCTSTR WINAPI GetObjectDescription()
	{
		return NULL;
	}
	template <class Q>
	static HRESULT CreateInstance(
		IUnknown* punkOuter,
		Q** pp)
	{
		return T::_CreatorClass::CreateInstance(punkOuter, __uuidof(Q), (void**) pp);
	}
	template <class Q>
	static HRESULT CreateInstance(Q** pp)
	{
		return T::_CreatorClass::CreateInstance(NULL, __uuidof(Q), (void**) pp);
	}
};

#endif // _ATL_STATIC_LIB_IMPL

#ifndef _ATL_NO_UUIDOF
#define _ATL_IIDOF(x) __uuidof(x)
#else
#define _ATL_IIDOF(x) IID_##x
#endif

#define _ATL_PACKING 8

#define offsetofclass(base, derived) ((DWORD_PTR)(static_cast<base*>((derived*)_ATL_PACKING))-_ATL_PACKING)

#define _ATL_DECLARE_GET_UNKNOWN(x) IUnknown* GetUnknown() throw() {return _GetRawUnknown();}

inline ATLAPI AtlInternalQueryInterface(
	void* pThis,
	const _ATL_INTMAP_ENTRY* pEntries,
	REFIID iid,
	void** ppvObject)
{
	//ATLASSERT(pThis != NULL);
	//ATLASSERT(pEntries!= NULL);

	if(pThis == NULL || pEntries == NULL)
		return E_INVALIDARG;

	// First entry in the com map should be a simple map entry
	//ATLASSERT(pEntries->pFunc == _ATL_SIMPLEMAPENTRY);

	if (ppvObject == NULL)
		return E_POINTER;

	if (InlineIsEqualUnknown(iid)) // use first interface
	{
		IUnknown* pUnk = (IUnknown*)((INT_PTR)pThis+pEntries->dw);
		pUnk->AddRef();
		*ppvObject = pUnk;
		return S_OK;
	}

	HRESULT hRes;

	for (;; pEntries++)
	{
		if (pEntries->pFunc == NULL)
		{
			hRes = E_NOINTERFACE;
			break;
		}

		BOOL bBlind = (pEntries->piid == NULL);
		if (bBlind || InlineIsEqualGUID(*(pEntries->piid), iid))
		{
			if (pEntries->pFunc == _ATL_SIMPLEMAPENTRY) //offset
			{
				//ATLASSERT(!bBlind);
				IUnknown* pUnk = (IUnknown*)((INT_PTR)pThis+pEntries->dw);
				pUnk->AddRef();
				*ppvObject = pUnk;
				return S_OK;
			}

			// Actual function call

			hRes = pEntries->pFunc(pThis,
				iid, ppvObject, pEntries->dw);
			if (hRes == S_OK)
				return S_OK;
			if (!bBlind && FAILED(hRes))
				break;
		}
	}

	*ppvObject = NULL;

	return hRes;
}

#define BEGIN_COM_MAP(x) public: \
	typedef x _ComMapClass; \
	IUnknown* _GetRawUnknown() \
	{ assert(_GetEntries()[0].pFunc == _ATL_SIMPLEMAPENTRY); return (IUnknown*)((INT_PTR)this+_GetEntries()->dw); } \
	_ATL_DECLARE_GET_UNKNOWN(x)\
	HRESULT _InternalQueryInterface( \
		REFIID iid, \
		void** ppvObject) \
	{ \
		return this->InternalQueryInterface(this, _GetEntries(), iid, ppvObject); \
	} \
	const static _ATL_INTMAP_ENTRY* WINAPI _GetEntries() { \
	static const _ATL_INTMAP_ENTRY _entries[] = {

#define COM_INTERFACE_ENTRY(x)\
	{&_ATL_IIDOF(x), \
	offsetofclass(x, _ComMapClass), \
	_ATL_SIMPLEMAPENTRY},

#define COM_INTERFACE_ENTRY_IID(iid, x)\
	{&iid,\
	offsetofclass(x, _ComMapClass),\
	_ATL_SIMPLEMAPENTRY},

template<typename Base, typename Derived, bool doInherit = __is_base_of(Base, Derived)>
struct AtlVerifyInheritance
{
	static_assert(doInherit,	"'Derived' class must inherit from 'Base'");
	static const DWORD_PTR value = 0; // Needed to be able to instantiate template
};

#define COM_INTERFACE_ENTRY2(x, x2)\
	{&_ATL_IIDOF(x),\
	AtlVerifyInheritance<x, x2>::value + (reinterpret_cast<DWORD_PTR>(static_cast<x*>(static_cast<x2*>(reinterpret_cast<_ComMapClass*>(_ATL_PACKING))))-_ATL_PACKING),\
	_ATL_SIMPLEMAPENTRY},

#define COM_INTERFACE_ENTRY2_IID(iid, x, x2)\
	{&iid,\
	ATL::AtlVerifyInheritance<x, x2>::value + (reinterpret_cast<DWORD_PTR>(static_cast<x*>(static_cast<x2*>(reinterpret_cast<_ComMapClass*>(_ATL_PACKING))))-_ATL_PACKING), \
	_ATL_SIMPLEMAPENTRY},

#define COM_INTERFACE_ENTRY_AGGREGATE(iid, punk)\
	{&iid,\
	(DWORD_PTR)offsetof(_ComMapClass, punk),\
	&_ComMapClass::_Delegate},

#define END_COM_MAP() \
	{NULL, 0, 0}}; return _entries;}

#endif // DUPATLBASE_H