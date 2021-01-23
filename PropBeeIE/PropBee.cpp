// PropBee.cpp : Implementation of CPropBee

#include "stdafx.h"
#include "PropBee.h"
#include "ImportFromPBDB.h"

CPropBee::CPropBee()
{
	m_bRunningForFirstTime = FALSE;
	m_szCacheFilePrefix = _T("PropBeePlus");
	ReadSettingsFromRegistry();		//read the user settings from registry
}

CPropBee::~CPropBee()
{
	m_PropBeeDB.FreeRecordMemory(m_HouseRecords);
}

// CPropBee
STDMETHODIMP CPropBee::SetSite(IUnknown* pUnkSite)
{
	HRX hrx;

	if (pUnkSite != NULL)
	{
		// Cache the pointer to IWebBrowser2.
		hrx.Set(pUnkSite->QueryInterface(IID_IWebBrowser2, (void **)&m_spWebBrowser));
		if (SUCCEEDED(hrx))
		{
			//get the IOleCommand interface
			hrx = pUnkSite->QueryInterface(IID_IOleCommandTarget,(void**)&m_spTarget);
			hrx = DispEventAdvise(m_spWebBrowser);
			m_fAdvised = TRUE;
		}
	}
	else
	{
		// Unregister event sink.
		if (m_fAdvised)
		{
			DispEventUnadvise(m_spWebBrowser);
			m_fAdvised = FALSE;
		}

		// Release cached pointers and other resources here.
		m_spWebBrowser.Release();
		m_spTarget.Release();
	}

	// Call base class implementation.
	return IObjectWithSiteImpl<CPropBee>::SetSite(pUnkSite);

}

void STDMETHODCALLTYPE CPropBee::OnDocumentComplete(IDispatch *pDisp, VARIANT *pvarURL)
{
	HRESULT hr = S_OK;

	// Query for the IWebBrowser2 interface.
	CComQIPtr<IWebBrowser2> spTempWebBrowser = pDisp;

	// Is this event associated with the top-level browser?
	if (spTempWebBrowser && m_spWebBrowser &&
		m_spWebBrowser.IsEqualObject(spTempWebBrowser))
	{
		// Get the current document object from browser...
		CComPtr<IDispatch> spDispDoc;
		hr = m_spWebBrowser->get_Document(&spDispDoc);
		if (SUCCEEDED(hr))
		{
			// ...and query for an HTML document.
			CComQIPtr<IHTMLDocument2> spHTMLDoc = spDispDoc;
			if (spHTMLDoc != NULL)
			{
				CComBSTR bstrURL;
				spHTMLDoc->get_URL(&bstrURL);
				m_szURL = bstrURL;m_szURL.MakeLower();

				if( m_bRunningForFirstTime )		//are we running for the first time?
				{
					RunConfigurationWizard();
				}
				
				if( m_szURL.Find(RMURL_RENT_SUMMARY) != -1 || m_szURL.Find(RMURL_SALE_SUMMARY) != -1 )
				{
					if( !m_PropBeeDB.IsConnected() )
					{
						m_PropBeeDB.ConnectToJetDatabase(m_szSQLServer);		//connect to database
						if( !m_PropBeeDB.IsConnected() )
						{
							AfxMessageBox(_T("Unable to connect to Property Spy database!!!"));
							return;
						}
					}
	
					//m_PropBeeDB.UploadDataFromPropertyBeeDB(_T("c:\\property-bee.sqlite"));
					Walk(spHTMLDoc);
				}
			}
		}
	}
}

HRESULT CPropBee::Walk(CComPtr<IHTMLDocument2> pDocument)
{
	HRX hrx;
	CComPtr<IHTMLElementCollection> pColl;
	CComBSTR bstrTagName, bstrText, bstrHistoryTable, bstrHTML;

	try
	{
		// retrieve a reference to the ALL collection
		hrx = pDocument->get_all(&pColl);
		long nElems=0;
		ASSERT(pColl);

		// retrieve the count of elements in the collection
		hrx = pColl->get_length(&nElems);
		// for each element retrieve properties such as TAGNAME and HREF
		for( int i=0; i < nElems; i++ )
		{
			static CString szDetailPrice, szBedrooms, szAddress;
			CComVariant svarItemIndex(i);
			CComVariant svarEmpty;
			CComPtr<IDispatch> pDisp;

			hrx.Set(pColl->item(svarItemIndex, svarEmpty, &pDisp));
			if( SUCCEEDED(hrx) )
			{
				CComPtr<IHTMLElement> pElem;
				hrx.Set(pDisp->QueryInterface(IID_IHTMLElement, (LPVOID*)&pElem));

				if( SUCCEEDED(hrx) )
				{
					ASSERT(pElem);
					bstrTagName.Empty();bstrText.Empty();
					hrx = pElem->get_className(&bstrTagName);
					hrx = pElem->get_id(&bstrText);
					CString szClassName = bstrTagName; szClassName.MakeLower();
					CString szID = bstrText; szID.MakeLower();

					m_bDetailMode = (m_szURL.Find(RMURL_SALE_DETAILS) != -1 || m_szURL.Find(RMURL_RENT_DETAILS) != -1);

					if( m_bDetailMode && szID.Find(TAGID_PAGE_HEADER) != -1 )
					{
						CHouseRecord TmpRecord;
						ExtractDetails(pElem, &TmpRecord);
						szDetailPrice = TmpRecord.m_szPrice;
					}

					if( ((szClassName.Find(TAG_REGULAR) != -1 || szClassName.Find(TAG_PREMIUM) != -1) && szID.Find(TAG_SUMMARY_PREFIX) != -1) ||
						(szID.Find(TAGID_PROPERTY_DETAILS) != -1) )
					{
						if( m_bDetailMode )
						{
							CString szTemp = m_szURL;
							szTemp.Replace(RMURL_SALE_DETAILS,_T(""));
							szTemp.Replace(RMURL_RENT_DETAILS,_T(""));
							int nDashPos = szTemp.Find('.');
							if( nDashPos != -1 )
								szID = szTemp.Left(nDashPos);
						}
						else
						{
							szID.Replace(TAG_SUMMARY_PREFIX,_T(""));
						}

						if( szID.IsEmpty() )	//if RM changes their data then stop
							break;

						CHouseRecord *pHouseRecord = new CHouseRecord();
						m_HouseRecords.AddTail(pHouseRecord);
						pHouseRecord->m_nKey = _ttoi(szID);
						
						ExtractDetails(pElem, pHouseRecord);
						bstrHistoryTable.Empty();

						if( m_bDetailMode )
						{
							pHouseRecord->m_szPrice = szDetailPrice;
						}

						BuildHistoryTable(pHouseRecord,bstrHistoryTable);

						//now display our history
						bstrText.Empty();
						hrx = pElem->get_innerHTML(&bstrText);

						if( m_bDetailMode )	//in detail mode we will display our history at top
						{
							bstrHTML = bstrHistoryTable;
							bstrHTML.Append(bstrText);
						}
						else		//in summary mode we will display our history after the house details
						{
							bstrHTML = bstrText;
							bstrHTML.Append(bstrHistoryTable);
						}
						hrx = pElem->put_innerHTML(bstrHTML);
					}
				}
			}
		} 
		//DumpRecords();		//useful for debugging
	}
	catch (_com_error &e)
	{
		hrx.Set(e.Error());
	}
	catch(...)
	{
		hrx.Set(E_UNEXPECTED);
	}

	return hrx;
}

HRESULT CPropBee::ExtractDetails(CComPtr<IHTMLElement> pElement, CHouseRecord *pHouseRecord)
{
	HRX hrx;

	try
	{
		CComPtr<IDispatch> pDisp;
		CComPtr<IHTMLElementCollection> pCollection;
		long nElems = 0;
		CComBSTR bstrTagName, bstrText;

		hrx = pElement->get_children(&pDisp);
		hrx = pDisp->QueryInterface(IID_IHTMLElementCollection, (LPVOID*)&pCollection);
		hrx = pCollection->get_length(&nElems);

		for( int i=0; i < nElems; i++ )
		{
			CComVariant svarItemIndex(i);
			CComVariant svarEmpty;
			CComPtr<IDispatch> pDisp2;

			hrx.Set(pCollection->item(svarItemIndex, svarEmpty, &pDisp2));
			if( SUCCEEDED(hrx) )
			{
				CComPtr<IHTMLElement> pElem;
				hrx.Set(pDisp2->QueryInterface(IID_IHTMLElement, (LPVOID*)&pElem));
				if( SUCCEEDED(hrx) )
				{
					ASSERT(pElem);
					bstrTagName.Empty();bstrText.Empty();
					hrx = pElem->get_className(&bstrTagName);
					hrx = pElem->get_innerText(&bstrText);
					CString szClassName = bstrTagName; szClassName.MakeLower();szClassName.Trim();
					CString szValue = bstrText;szValue.Trim();
				
					if( szClassName.CompareNoCase(TAG_PRICE) == 0 )
					{
						pHouseRecord->m_szPrice = szValue;
						
						int nPoundChar = 163;
						int nPoundPos = szValue.Find(nPoundChar);
						if( nPoundPos != -1 )
						{
							CString szTmpPrice = szValue.Mid(nPoundPos+1);
							szTmpPrice.Replace(_T(","),_T(""));
							pHouseRecord->m_nPrice = _ttoi(szTmpPrice);
						}
					}
					if( szClassName.CompareNoCase(TAG_STATUS_SOLD) == 0 )
					{
						pHouseRecord->m_szStatus = _T("Sold SSTC");
					}
					else if( szClassName.CompareNoCase(TAG_AGENT) == 0 )
					{
						szValue.Replace(_T("Marketed by"),_T(""));
						szValue.Trim();

						int nCommaPos = szValue.Find(_T(','));
						int nDotPos = szValue.Find(_T('.'));

						if( nCommaPos > -1 && nDotPos > -1 )
						{
							pHouseRecord->m_szAgent = szValue.Left(nCommaPos);
							int nLocLength = (nDotPos) - (nCommaPos+2);
							pHouseRecord->m_szAgentLocation = szValue.Mid(nCommaPos+2,nLocLength);
							pHouseRecord->m_szAgentLocation.Replace(_T("'"),_T("&#39;"));
							pHouseRecord->m_szAgentLocation.Replace(_T("\""),_T("&quot;"));
						}

						int nTelPos = szValue.Find(_T(':'));
						int nBrackPos = szValue.Find(_T('('));
						if( nTelPos > -1 && nBrackPos > -1 )
						{
							pHouseRecord->m_szAgentTel = szValue.Mid(nTelPos+2,nBrackPos-(nTelPos+2));
							pHouseRecord->m_szAgentTel.Trim();
							pHouseRecord->m_szAgentTel.Replace(_T(" "),_T(""));
							pHouseRecord->m_szAgentTel.Replace(_T("'"),_T("&#39;"));
							pHouseRecord->m_szAgentTel.Replace(_T("\""),_T("&quot;"));
						}
					}
					else if( szClassName.CompareNoCase(TAG_DESC) == 0 )
					{
						pHouseRecord->m_szDescription = szValue;
						pHouseRecord->m_szDescription.Replace(_T("'"),_T("&#39;"));
						pHouseRecord->m_szDescription.Replace(_T("\""),_T("&quot;"));
					}
					else if( szClassName.CompareNoCase(TAG_DETAIL_DESC) == 0 )
					{
						pHouseRecord->m_szDetailDescription = szValue;
						pHouseRecord->m_szDetailDescription.Replace(_T("'"),_T("&#39;"));
						pHouseRecord->m_szDetailDescription.Replace(_T("\""),_T("&quot;"));
					}
					else if( szClassName.CompareNoCase(TAG_SUBTITLE) == 0 )
					{
						pHouseRecord->m_szSubTitle = szValue;pHouseRecord->m_szSubTitle.MakeLower();
						int nPos = pHouseRecord->m_szSubTitle.Find(_T("for sale"));
						if( nPos != -1 )
						{
							pHouseRecord->m_szSubTitle = szValue.Left(nPos);
							pHouseRecord->m_szSubTitle.Trim();
							pHouseRecord->m_szSubTitle.Replace(_T("'"),_T("&#39;"));
							pHouseRecord->m_szSubTitle.Replace(_T("\""),_T("&quot;"));
						}
					}
					else if( szClassName.CompareNoCase(TAG_TITLE) == 0 )
					{
						pHouseRecord->m_szTitle = szValue;
						pHouseRecord->m_szTitle.Trim();
						pHouseRecord->m_szTitle.Replace(_T("'"),_T("&#39;"));
						pHouseRecord->m_szTitle.Replace(_T("\""),_T("&quot;"));
					}
					ExtractDetails(pElem,pHouseRecord);
				}
			}
		}
	}
	catch (_com_error &e)
	{
		hrx.Set(e.Error());
	}
	catch(...)
	{
		hrx.Set(E_UNEXPECTED);
	}

	return hrx;
}

void CPropBee::BuildHistoryTable(CHouseRecord *pHouseRecord, CComBSTR& bstrTableHTML)
{
	CString szContent, szAttractColor=_T("#FFCC00"), szNormalColor=_T("#CCCCCC"), szWhiteColor=_T("#FFFFFF"), szHeaderColor=_T("#9999FF");
	BOOL bNewEntryFound=FALSE;
	int nCurPrice=0,nDiff=0;
	BOOL bPriceNotAdded=TRUE;
	BOOL bAtleastOneHighlighted = FALSE;
	
	int nCount = m_PropBeeDB.ManageHistory(pHouseRecord,m_HistoryRecords,m_bDetailMode,bNewEntryFound);
	
	szContent = _T("<table border=\"1\" cellspacing=\"0\" cellpadding=\"5\" width=\"90%\" >");
	szContent.AppendFormat(_T("<tr bgcolor=\"%s\"><th align=\"left\">date</th><th align=\"left\">event</th></tr>"), m_bDetailMode?szHeaderColor:szWhiteColor);
		
	if( nCount <= 1 && bNewEntryFound )	//use the current record
	{
		szContent.AppendFormat(_T("<tr><td bgcolor=\"%s\" width=\"30%%\" valign=\"top\">%s</td>"), (bNewEntryFound)?szAttractColor:szNormalColor, pHouseRecord->m_oleDate.Format(_T("%d %b %Y %I:%M:%S %p")));
		szContent.AppendFormat(_T("<td bgcolor=\"%s\" valign=\"top\">%s</td></tr>"),(bNewEntryFound)?szAttractColor:szNormalColor, pHouseRecord->m_szHistory);
	}
	else  
	{
		//traverse the record in reverse order to print the latest history first
		POSITION pos = m_HistoryRecords.GetTailPosition();
		while( pos != NULL )
		{
			COleDateTimeSpan timeDifference;
			CHouseRecord *pRecord = m_HistoryRecords.GetPrev(pos);
			if( !pRecord )
				continue;

			if( !pRecord->m_bShowRecord )
				continue;

			if( bPriceNotAdded )
			{
				nCurPrice = pRecord->m_nPrice;
				bPriceNotAdded = FALSE;
			}

			COleDateTime timeNow = COleDateTime::GetCurrentTime();
			timeDifference = timeNow - pRecord->m_oleDate;
			if( timeDifference.GetTotalDays() <= 7 )
			{
				bAtleastOneHighlighted = TRUE;
				pRecord->m_bHighlight = TRUE;
			}

			szContent.AppendFormat(_T("<tr><td bgcolor=\"%s\" width=\"30%%\" valign=\"top\">%s</td>"), (pRecord->m_bHighlight)?szAttractColor:szNormalColor,pRecord->m_oleDate.Format(_T("%d %b %Y %I:%M:%S %p")));
			szContent.AppendFormat(_T("<td bgcolor=\"%s\" valign=\"top\">%s</td></tr>"),(pRecord->m_bHighlight)?szAttractColor:szNormalColor, pRecord->m_szHistory);
		}
	}
	szContent.Append(_T("</table><br>"));
	//bstrTableHTML.Append(_T("<b>Property History:</b>"));

	/*
	if( m_HistoryRecords.GetCount() > 0 )
	{
		CHouseRecord *pBase = m_HistoryRecords.GetHead();
		if( pBase && pBase->m_nPrice > 0 && nCurPrice > 0 )
			nDiff=pBase->m_nPrice - nCurPrice;

		if( nDiff > 0 || nDiff < 0 )
		{
			CString strImage;
			strImage.AppendFormat(_T("<img src=\"res://PropBeeIE.dll/PNG/#%s\" ALIGN=right title=\"%s\">"),
				(nDiff > 0)?_T("707"):_T("706"),(nDiff > 0)?_T("price of this property has gone down"):_T("price of this property has gone up"));
			bstrTableHTML.Append(strImage);
		}
	}
	*/
	CString strImage;
	strImage.AppendFormat(_T("<img src=\"res://PropBeeIE.dll/PNG/#%s\" title=\"Property Spy analysis\">"),(bAtleastOneHighlighted ||bNewEntryFound)? _T("708"):_T("709"));
	bstrTableHTML.Append(strImage);

	bstrTableHTML.Append(szContent);
	
	m_PropBeeDB.FreeRecordMemory(m_HistoryRecords);	//free up the memory
	m_HistoryRecords.RemoveAll();
}

STDMETHODIMP CPropBee::Exec(const GUID *pguidCmdGroup, DWORD nCmdID,DWORD nCmdExecOpt, VARIANTARG *pvaIn, VARIANTARG *pvaOut)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CPropertySheet settingsSheet(_T("Property Spy"));
	CAppearancePageDlg pageAppearance;
	CDataPageDlg pageData;

	settingsSheet.AddPage(&pageAppearance);
	settingsSheet.AddPage(&pageData);
	pageData.SetInstallPath(m_szInstallPath);

	settingsSheet.m_psh.dwFlags |= PSH_NOAPPLYNOW;
	settingsSheet.m_psh.dwFlags &= ~PSH_HASHELP;

	settingsSheet.DoModal();
	return S_OK;
}

STDMETHODIMP CPropBee::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds,OLECMD *prgCmds, OLECMDTEXT *pCmdText)
{
	return E_NOTIMPL;
}

void CPropBee::ReadSettingsFromRegistry()
{
	HRX hrx;
	try
	{
		CRegKey regKey;

		int nRetVal = regKey.Open(HKEY_LOCAL_MACHINE,REG_BASE);
		if( nRetVal == ERROR_SUCCESS )
		{
			TCHAR szBuf[1024];
			ULONG nBufSize = sizeof(szBuf);
			nRetVal = regKey.QueryStringValue(_T("SQLServer"),szBuf,&nBufSize);

			if( nRetVal == ERROR_SUCCESS )
				m_szSQLServer = szBuf;

			DWORD dwValue=0;
			nRetVal = regKey.QueryDWORDValue(_T("FirstTime"),dwValue);
			if( nRetVal == ERROR_SUCCESS )
				m_bRunningForFirstTime = dwValue;

			nBufSize = sizeof(szBuf);
			nRetVal = regKey.QueryStringValue(_T("InstallPath"),szBuf,&nBufSize);
			if( nRetVal == ERROR_SUCCESS )
				m_szInstallPath = szBuf;
		}
	}
	catch (_com_error &e)
	{
		hrx.Set(e.Error());
	}
	catch(...)
	{
		hrx.Set(E_UNEXPECTED);
	}
}

void CPropBee::SetFirstTimeFlagValue(BOOL bVal)
{
	HRX hrx;
	try
	{
		CRegKey regKey;

		int nRetVal = regKey.Open(HKEY_LOCAL_MACHINE,REG_BASE);
		if( nRetVal == ERROR_SUCCESS )
		{
			DWORD dwValue=bVal;
			nRetVal = regKey.SetDWORDValue(_T("FirstTime"),dwValue);
		}
	}
	catch (_com_error &e)
	{
		hrx.Set(e.Error());
	}
	catch(...)
	{
		hrx.Set(E_UNEXPECTED);
	}
}

void CPropBee::SetSQLServerRegistryValue(CString szSQLServerName)
{
	HRX hrx;
	try
	{
		CRegKey regKey;

		int nRetVal = regKey.Open(HKEY_LOCAL_MACHINE,REG_BASE);
		if( nRetVal == ERROR_SUCCESS )
		{
			nRetVal = regKey.SetStringValue(_T("SQLServer"),szSQLServerName);
		}
	}
	catch (_com_error &e)
	{
		hrx.Set(e.Error());
	}
	catch(...)
	{
		hrx.Set(E_UNEXPECTED);
	}
}

BOOL CPropBee::RunConfigurationWizard(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	BOOL bSuccess = TRUE;
	CImportPBDBDlg importDlg;
	importDlg.SetInstallPath(m_szInstallPath);
	bSuccess = (importDlg.DoModal() == IDOK)?TRUE:FALSE;

	if( bSuccess )	//if successful flag that we have completed our configuration successfully
	{
		SetSQLServerRegistryValue(importDlg.GetSQLServerName());
		SetFirstTimeFlagValue(FALSE);	
		ReadSettingsFromRegistry();	//refresh our variables
	}

	return bSuccess;
}

BOOL CPropBee::WriteHistoryToCache(long nKey, CComBSTR& bstrHistory)
{
	BOOL bHistoryWritten = FALSE;
	CString szHistoryFile;

	CString szFilePath;
	szFilePath.GetEnvironmentVariable(_T("TEMP"));
	if( szFilePath.IsEmpty() )
		return bHistoryWritten;

	szHistoryFile.Format(_T("%s\\%s-%lld%s"),szFilePath,m_szCacheFilePrefix,nKey,_T(".txt"));

	BOOL bFilexists = PathFileExists(szHistoryFile);
	if( bFilexists )
		return bHistoryWritten;

	CStdioFile cacheFile;
	cacheFile.Open(szHistoryFile,CFile::modeCreate|CFile::modeWrite|CFile::typeText);
	cacheFile.Write(bstrHistory.m_str,bstrHistory.Length());
	cacheFile.Close();
	bHistoryWritten = TRUE;

	return bHistoryWritten;
}

BOOL CPropBee::LoadHistoryFromCache(long nKey, CComBSTR& bstrHistory)
{
	BOOL bHistoryRead = FALSE;
	CString szHistoryFile;
	CString szFilePath;
	TCHAR szBuf[1024];
	szFilePath.GetEnvironmentVariable(_T("TEMP"));
	if( szFilePath.IsEmpty() )
		return bHistoryRead;

	szHistoryFile.Format(_T("%s\\%s-%lld%s"),szFilePath,m_szCacheFilePrefix,nKey,_T(".txt"));

	BOOL bFilexists = PathFileExists(szHistoryFile);
	if( !bFilexists )
		return bHistoryRead;

	CStdioFile cacheFile;
	cacheFile.Open(szHistoryFile,CFile::modeRead|CFile::shareDenyNone);
	ULONGLONG dwLength = cacheFile.GetLength();
	cacheFile.Read(szBuf,sizeof(szBuf));
	cacheFile.Close();
	bstrHistory = szBuf;

	bHistoryRead = TRUE;
	return bHistoryRead;
}