// PropBee.h : Declaration of the CPropBee

#pragma once
#include "resource.h"       // main symbols
#include <shlguid.h>     // IID_IWebBrowser2, DIID_DWebBrowserEvents2, etc.
#include <exdispid.h> // DISPID_DOCUMENTCOMPLETE, etc.
#include <mshtml.h>         // DOM interfaces
#include "PropBeeIE.h"
#include "PropBeeDB.h"
#include "AppearancePageDlg.h"
#include "DataPageDlg.h"
#include "GeneralPageDlg.h"

#define REG_BASE		_T("Software\\PropertySpy\\Settings")
#define REG_IEBASE		_T("Software\\Microsoft\\Internet Explorer\\Extensions\\{ADCC389F-B74A-4881-B5D5-ABDAF58E646F}")

#define RMURL_RENT_SUMMARY	_T("http://www.rightmove.co.uk/property-to-rent/")
#define RMURL_RENT_DETAILS	_T("http://www.rightmove.co.uk/property-to-rent/property-")
#define RMURL_SALE_SUMMARY	_T("http://www.rightmove.co.uk/property-for-sale/")
#define RMURL_SALE_DETAILS	_T("http://www.rightmove.co.uk/property-for-sale/property-")

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

#define TAG_REGULAR				_T("regular")
#define TAG_PREMIUM				_T("premium")
#define TAG_SUMMARY_PREFIX		_T("summary")
#define TAG_DETAILS				_T("details")	//class name when in summary listing mode

#define TAG_STATUS_SOLD					_T("propertystatus sold")
#define TAGID_PROPERTY_DETAILS			_T("propertydetails")		//class name when viewing a particular property
#define TAGID_PAGE_HEADER				_T("pageheader")

#define TAG_PRICE						_T("price")
#define TAG_DESC						_T("description")
#define TAG_DETAIL_DESC					_T("propertyDetailDescription")
#define TAG_AGENT						_T("branchblurb")
#define TAG_BRANCH						_T("branch")
#define TAG_TITLE						_T("displayaddress")
#define TAG_SUBTITLE					_T("address bedrooms")

// CPropBee

class ATL_NO_VTABLE CPropBee :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CPropBee, &CLSID_PropBee>,
	public IObjectWithSiteImpl<CPropBee>,
	public IDispatchImpl<IPropBee, &IID_IPropBee, &LIBID_PropBeeIELib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispEventImpl<1, CPropBee, &DIID_DWebBrowserEvents2, &LIBID_SHDocVw, 1, 1>,
	public IOleCommandTarget
{
public:
	CPropBee();
	virtual ~CPropBee();

DECLARE_REGISTRY_RESOURCEID(IDR_PROPBEE)

DECLARE_NOT_AGGREGATABLE(CPropBee)

BEGIN_COM_MAP(CPropBee)
	COM_INTERFACE_ENTRY(IPropBee)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IObjectWithSite)
	COM_INTERFACE_ENTRY(IOleCommandTarget)
END_COM_MAP()

BEGIN_SINK_MAP(CPropBee)
	SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_DOCUMENTCOMPLETE, OnDocumentComplete)
END_SINK_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

	STDMETHOD(SetSite)(IUnknown *pUnkSite);
	// DWebBrowserEvents2
	void STDMETHODCALLTYPE OnDocumentComplete(IDispatch *pDisp, VARIANT *pvarURL);
	STDMETHOD(Exec)(const GUID *pguidCmdGroup, DWORD nCmdID,DWORD nCmdExecOpt, VARIANTARG *pvaIn, VARIANTARG *pvaOut);
	STDMETHOD(QueryStatus)(const GUID *pguidCmdGroup, ULONG cCmds,OLECMD *prgCmds, OLECMDTEXT *pCmdText);

public:
	CString GetInstallPath(){return m_szInstallPath;}

private:
	HRESULT Walk(CComPtr<IHTMLDocument2> pDoc);
	HRESULT ExtractDetails(CComPtr<IHTMLElement> pElement, CHouseRecord *pHouseRecord);
	void BuildHistoryTable(CHouseRecord *pHouseRecord, CComBSTR& bstrTableHTML);
	void ReadSettingsFromRegistry();
	void SetFirstTimeFlagValue(BOOL bVal=FALSE);
	void SetSQLServerRegistryValue(CString szSQLServerName);
	BOOL RunConfigurationWizard(void);
	BOOL WriteHistoryToCache(long nKey, CComBSTR& bstrHistory);
	BOOL LoadHistoryFromCache(long nKey, CComBSTR& bstrHistory);

private:
	 CComPtr<IWebBrowser2>  m_spWebBrowser;
	 CComQIPtr<IOleCommandTarget,&IID_IOleCommandTarget> m_spTarget;
	 BOOL m_fAdvised;
	 CAtlList<CHouseRecord *> m_HouseRecords;
	 CAtlList<CHouseRecord *> m_HistoryRecords;
	 CPropBeeDB m_PropBeeDB;	//our database connection
	 CString m_szURL;
	 BOOL m_bDetailMode;
	 CString m_szSQLServer;
	 BOOL m_bRunningForFirstTime;
	 CString m_szInstallPath;
	 CString m_szCacheFilePrefix;
};

OBJECT_ENTRY_AUTO(__uuidof(PropBee), CPropBee)
