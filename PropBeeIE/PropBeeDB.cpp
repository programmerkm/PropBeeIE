// PropBeeDB.cpp : implementation file
//

#include "stdafx.h"
#include "PropBeeDB.h"


// CPropBeeDB

CPropBeeDB::CPropBeeDB()
{
	m_bConnected = FALSE;
	m_bDetailsMode = FALSE;
	m_szSQLServerDBName = _T("PropertyBeeDB");
	m_szTableName = _T("RMData");
	m_nDatabaseType = CPropBeeDB::DATABASE_SQLITE;
	m_pSQLiteConn = new CppSQLite3DB();
}

CPropBeeDB::~CPropBeeDB()
{
	if( m_pConn && m_bConnected )
		m_pConn->Close();

	if( m_pSQLiteConn )
		delete m_pSQLiteConn;

	//m_File.Close();
}

// CPropBeeDB member functions
BOOL CPropBeeDB::ConnectToSQLServerDatabase(CString szServer,CString szDatabase)
{
	BOOL bRetVal = FALSE;
	HRX hrx;
	CComBSTR bstrConnection;
	CString szMsg;

	//m_File.Open(_T("c:\\rmDB.txt"), CFile::modeCreate|CFile::modeWrite|CFile::typeText);

	if( m_bConnected && m_pConn )	//already connected?
	{
		m_pConn->Close();
		m_pConn.Release();
		m_bConnected = FALSE;
	}

	try
	{
		//if server name not given assume the local computer name
		if( szServer.IsEmpty() )
		{
			TCHAR szComputerName[MAX_COMPUTERNAME_LENGTH*20];
			DWORD dwBufSize = sizeof(szComputerName);
			GetComputerNameEx(ComputerNamePhysicalDnsFullyQualified,szComputerName,&dwBufSize);
			szServer = szComputerName;
			//szMsg.Format(_T("Connecting to database..%s\n"),szServer);
			//m_File.WriteString(szMsg);
		}

		CString szConnection;
		szConnection.Format(_T("Provider=SQLOLEDB;Server=%s;Database=%s;Integrated Security=SSPI;"),szServer, !szDatabase.IsEmpty()?szDatabase:m_szSQLServerDBName);
		bstrConnection = szConnection;
		hrx = m_pConn.CreateInstance((__uuidof(Connection)));
		hrx = m_pConn->Open(bstrConnection.m_str,"","",0);
		bRetVal = TRUE;
		m_bConnected = TRUE;
		m_nDatabaseType = CPropBeeDB::DATABASE_SQLSERVER;
		//szMsg.Format(_T("Connection string:%s   Status:Connected\n"),szConnection);
		//m_File.WriteString(szMsg);
	}
	catch(_com_error & e)
	{
		hrx.Set(e.Error());
		bRetVal = FALSE;
	}

	return bRetVal;
}

// CPropBeeDB member functions
BOOL CPropBeeDB::ConnectToJetDatabase(CString szJetDBPath)
{
	BOOL bRetVal = FALSE;
	HRX hrx;
	CComBSTR bstrConnection;
	CString szMsg;

	//m_File.Open(_T("c:\\rmDB.txt"), CFile::modeCreate|CFile::modeWrite|CFile::typeText);

	if( m_bConnected )	//already connected?
		return m_bConnected;

	try
	{
		CString szConnection;
		szConnection.Format(_T("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=%s;Jet OLEDB:Database Password=pr0pertyspy;"),szJetDBPath);
		bstrConnection = szConnection;
		hrx = m_pConn.CreateInstance((__uuidof(Connection)));
		hrx = m_pConn->Open(bstrConnection.m_str,"","",0);
		bRetVal = TRUE;
		m_bConnected = TRUE;
		m_szJetDBPath = szJetDBPath;
		m_nDatabaseType = CPropBeeDB::DATABASE_JETOLDEB;
		//szMsg.Format(_T("Connection string:%s   Status:Connected\n"),szConnection);
		//m_File.WriteString(szMsg);
	}
	catch(_com_error & e)
	{
		hrx.Set(e.Error());
		bRetVal = FALSE;
	}

	return bRetVal;
}


BOOL CPropBeeDB::ConnectToSQLiteDatabase(CString szSqliteDBPath)
{
	BOOL bRetVal = FALSE;

	//m_File.Open(_T("c:\\rmDB.txt"), CFile::modeCreate|CFile::modeWrite|CFile::typeText);

	if( m_bConnected )	//already connected?
		return m_bConnected;

	try
	{
		m_pSQLiteConn->open(szSqliteDBPath);
		bRetVal = TRUE;
		m_bConnected = TRUE;
		m_szSQLiteDBPath = szSqliteDBPath;
		m_nDatabaseType = CPropBeeDB::DATABASE_SQLITE;
	}
	catch (CppSQLite3Exception& e)
	{
		bRetVal = FALSE;
	}
	catch(...)
	{
		bRetVal = FALSE;
	}

	return bRetVal;
}



BOOL CPropBeeDB::ExecuteSQLQuery(CComBSTR bstrQuery, CComVariant& varResult, CString& szErrorMsg)
{
	CString csResult;
	BOOL bRetVal = FALSE;
	szErrorMsg = _T("");

	if( !m_bConnected )
		return bRetVal;

	try
	{
		_CommandPtr pCommand;
		pCommand.CreateInstance(__uuidof(Command));
		pCommand->ActiveConnection = m_pConn;
		pCommand->CommandType = adCmdText;
		pCommand->CommandText = bstrQuery.m_str;
		_RecordsetPtr pRstSet = NULL;

		pRstSet = pCommand->Execute(NULL,NULL,adCmdText); 
		varResult.ClearToZero();

		if( pRstSet->GetADO_EOF() != VARIANT_TRUE )
			varResult = pRstSet->Fields->GetItem(_variant_t((long)0))->Value;

		pRstSet->Close();
		bRetVal = TRUE;
	}
	catch(_com_error& e)
	{
		bRetVal = FALSE;
		szErrorMsg = (LPWSTR)e.Description();
	}

	return bRetVal;
}

BOOL CPropBeeDB::ExecuteSQLQuery(CComBSTR bstrQuery, _RecordsetPtr& pRecordSets, CString& szErrorMsg)
{
	CString csResult;
	BOOL bRetVal = FALSE;
	szErrorMsg = _T("");

	if( !m_bConnected )
		return bRetVal;

	try
	{
		_CommandPtr pCommand;
		pCommand.CreateInstance(__uuidof(Command));
		pCommand->ActiveConnection = m_pConn;
		pCommand->CommandType = adCmdText;
		pCommand->CommandText = bstrQuery.m_str;
		pRecordSets = pCommand->Execute(NULL,NULL,adCmdText); 
		bRetVal = TRUE;
	}
	catch(_com_error& e)
	{
		bRetVal = FALSE;
		szErrorMsg = (LPWSTR)e.Description();
	}

	return bRetVal;
}

BOOL CPropBeeDB::CreateSQLServerTables(void)
{
	CString szQuery, szError;
	CComBSTR bstrQuery;
	CComVariant varResult;
	BOOL bRetVal = FALSE;

	if( !m_bConnected )
		return FALSE;

	_RecordsetPtr pRecordSets;

	//first the main RMData table
	szQuery.AppendFormat(_T("CREATE TABLE [%s]([%s] [int] NOT NULL,"),m_szTableName,COL_PROPID);
	szQuery.AppendFormat(_T("[%s] [datetime] NOT NULL,[%s] [nvarchar](50) NULL,"),COL_DATE,COL_PRICE);
	szQuery.AppendFormat(_T("[%s] [int] NULL,[%s] [nvarchar](50) NULL,"),COL_NUMPRICE,COL_STATUS); 
	szQuery.AppendFormat(_T("[%s] [nvarchar](4000) NULL,"),COL_DESC);
	szQuery.AppendFormat(_T("[%s] [nvarchar](MAX) NULL,"),COL_DETAIL_DESC); 
	szQuery.AppendFormat(_T("[%s] [nvarchar](150) NULL,"),COL_AGENT);
	szQuery.AppendFormat(_T("[%s] [nvarchar](150) NULL,"),COL_AGENT_ADDRESS);
	szQuery.AppendFormat(_T("[%s] [nvarchar](100) NULL,"),COL_AGENT_LOCATION);
	szQuery.AppendFormat(_T("[%s] [nvarchar](30) NULL,"),COL_AGENT_TELEPHONE);
	szQuery.AppendFormat(_T("[%s] [nvarchar](150) NULL,"),COL_TITLE);
	szQuery.AppendFormat(_T("[%s] [nvarchar](200) NULL) ON [PRIMARY]"),COL_SUBTITLE);

	bstrQuery = szQuery;
	bRetVal = ExecuteSQLQuery(bstrQuery,pRecordSets,szError);

	return bRetVal;
}

//this method will add the current record if necessary and return back the complete history for the current record (based on property id)
int CPropBeeDB::ManageHistory(CHouseRecord *pCurHouseRecord, CAtlList<CHouseRecord *>& HistoryRecords, BOOL bDetailsMode, BOOL& bNewEntryFound)
{
	int nCount = 0;
	CString szMsg;
	HRX hrx;
	CString szQuery, szError;
	CComBSTR bstrQuery;
	COleDateTime dateLastEntry;
			
	try
	{
		bNewEntryFound = FALSE;
		m_bDetailsMode = bDetailsMode;
		
		pCurHouseRecord->m_oleDate = COleDateTime::GetCurrentTime();
		pCurHouseRecord->m_szSQLDate = pCurHouseRecord->m_oleDate.Format(_T("%Y-%m-%d %H:%M:%S"));

		BOOL bEntryFound = GetLastEntryDate(pCurHouseRecord->m_nKey,dateLastEntry);
		if( !bEntryFound )	//there were no entries for our property?
		{
			pCurHouseRecord->m_szHistory = _T("• initial entry found.");
			InsertRecord(pCurHouseRecord);	//this is the first entry so insert everything
			bNewEntryFound = TRUE;
			return 1;
		}

		nCount = BuildRecordList(pCurHouseRecord->m_nKey,TEST_MASK_ALL,HistoryRecords,pCurHouseRecord);
	}
	catch(_com_error& e)
	{
		hrx.Set(e.Error());
		nCount = 0;
	}
	
	return nCount;
}

BOOL CPropBeeDB::InsertRecord(CHouseRecord *pCurHouseRecord,int *pnMask)
{
	CString szQuery, szMsg, szError;
	CComBSTR bstrQuery;
	HRX hrx;
	BOOL bRetVal = FALSE;

	try
	{
		_RecordsetPtr pRecordSets;
		int nFieldsToInsert=0;
		
		if( !pnMask )
			nFieldsToInsert = GetInsertMask(pCurHouseRecord);
		else
			nFieldsToInsert = *pnMask;
		
		if( IsUsingSQLServer() )
			szQuery.AppendFormat(_T("INSERT INTO [%s].[dbo].[%s] ([%s],[%s]"),m_szSQLServerDBName,m_szTableName,COL_PROPID,COL_DATE);
		else if( IsUsingJetOLEDB() || IsUsingSQLITE() )
			szQuery.AppendFormat(_T("INSERT INTO [%s]([%s],[%s] "),m_szTableName,COL_PROPID,COL_DATE);

		if( (nFieldsToInsert & TEST_MASK_PRICE) || (nFieldsToInsert == TEST_MASK_ALL) )
		{
			szQuery.AppendFormat(_T(",[%s]"),COL_PRICE);
			szQuery.AppendFormat(_T(",[%s]"),COL_NUMPRICE);
		}

		if( nFieldsToInsert & TEST_MASK_STATUS || (nFieldsToInsert == TEST_MASK_ALL) )
			szQuery.AppendFormat(_T(",[%s]"),COL_STATUS);

		if( nFieldsToInsert & TEST_MASK_DESC || (nFieldsToInsert == TEST_MASK_ALL) )
			szQuery.AppendFormat(_T(",[%s]"),COL_DESC);

		if( nFieldsToInsert & TEST_MASK_DETAIL_DESC || (nFieldsToInsert == TEST_MASK_ALL) )
			szQuery.AppendFormat(_T(",[%s]"),COL_DETAIL_DESC);

		if( nFieldsToInsert & TEST_MASK_AGENT || (nFieldsToInsert == TEST_MASK_ALL) )
		{
			szQuery.AppendFormat(_T(",[%s]"),COL_AGENT);
			szQuery.AppendFormat(_T(",[%s]"),COL_AGENT_ADDRESS);
			szQuery.AppendFormat(_T(",[%s]"),COL_AGENT_LOCATION);
			szQuery.AppendFormat(_T(",[%s]"),COL_AGENT_TELEPHONE);
		}

		if( nFieldsToInsert & TEST_MASK_TITLE || (nFieldsToInsert == TEST_MASK_ALL) )
			szQuery.AppendFormat(_T(",[%s]"),COL_TITLE);

		if( nFieldsToInsert & TEST_MASK_SUBTITLE || (nFieldsToInsert == TEST_MASK_ALL) )
			szQuery.AppendFormat(_T(",[%s]"),COL_SUBTITLE);

		szQuery.Append(_T(") VALUES ("));
		szQuery.AppendFormat(_T("%d,'%s' "),pCurHouseRecord->m_nKey,pCurHouseRecord->m_szSQLDate);
		
		if( nFieldsToInsert & TEST_MASK_PRICE || (nFieldsToInsert == TEST_MASK_ALL) )
		{
			szQuery.AppendFormat(_T(",'%s'"),pCurHouseRecord->m_szPrice);
			szQuery.AppendFormat(_T(",%d"),pCurHouseRecord->m_nPrice);
		}

		if( nFieldsToInsert & TEST_MASK_STATUS || (nFieldsToInsert == TEST_MASK_ALL) )
			szQuery.AppendFormat(_T(",'%s'"),pCurHouseRecord->m_szStatus);

		if( nFieldsToInsert & TEST_MASK_DESC || (nFieldsToInsert == TEST_MASK_ALL) )
			szQuery.AppendFormat(_T(",'%s'"),pCurHouseRecord->m_szDescription);

		if( nFieldsToInsert & TEST_MASK_DETAIL_DESC || (nFieldsToInsert == TEST_MASK_ALL) )
			szQuery.AppendFormat(_T(",'%s'"),pCurHouseRecord->m_szDetailDescription);

		if( nFieldsToInsert & TEST_MASK_AGENT || (nFieldsToInsert == TEST_MASK_ALL) )
		{
			szQuery.AppendFormat(_T(",'%s'"),pCurHouseRecord->m_szAgent);
			szQuery.AppendFormat(_T(",'%s'"),pCurHouseRecord->m_szAgentAddress);
			szQuery.AppendFormat(_T(",'%s'"),pCurHouseRecord->m_szAgentLocation);
			szQuery.AppendFormat(_T(",'%s'"),pCurHouseRecord->m_szAgentTel);
		}

		if( nFieldsToInsert & TEST_MASK_TITLE || (nFieldsToInsert == TEST_MASK_ALL) )
			szQuery.AppendFormat(_T(",'%s'"),pCurHouseRecord->m_szTitle);

		if( nFieldsToInsert & TEST_MASK_SUBTITLE || (nFieldsToInsert == TEST_MASK_ALL) )
			szQuery.AppendFormat(_T(",'%s'"),pCurHouseRecord->m_szSubTitle);

		szQuery.Append(_T(" );"));

		bstrQuery = szQuery;
		bRetVal = ExecuteSQLQuery(bstrQuery,pRecordSets,szError);

	}
	catch(_com_error& e)
	{
		hrx.Set(e.Error());
		bRetVal = FALSE;
	}

	return bRetVal;
}

int CPropBeeDB::BuildRecordList(long nKey, int nHistoryMask,CAtlList<CHouseRecord *>& RecordsList, CHouseRecord *pNewRecord)
{
	CString szQuery, szMsg, szError;
	CComBSTR bstrQuery;
	HRX hrx;
	int nCount = 0;

	try
	{
		FreeRecordMemory(RecordsList);
		RecordsList.RemoveAll();
		CHouseRecord *pRecordCopy = NULL;
		_RecordsetPtr pRecordSets;

		szQuery.Format(_T("SELECT %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s FROM %s where %s = %d order by %s ASC"), 
			COL_DATE,COL_PRICE,COL_NUMPRICE,COL_STATUS,COL_DESC,COL_DETAIL_DESC,COL_AGENT,
			COL_AGENT_ADDRESS,COL_AGENT_LOCATION,COL_AGENT_TELEPHONE,COL_TITLE,COL_SUBTITLE,
			m_szTableName,COL_PROPID,nKey,COL_DATE);

		bstrQuery = szQuery;

		//szMsg.Format(_T("About to execute query:%s\n"), szQuery);
		//m_File.WriteString(szMsg);

		ExecuteSQLQuery(bstrQuery,pRecordSets,szError);
		while( pRecordSets->GetADO_EOF() != VARIANT_TRUE )
		{
			CComVariant varValue;
			CHouseRecord *pHouseRecord = new CHouseRecord();

			pHouseRecord->m_nKey = nKey;

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(_T("EntryDate"))->Value;
			if( varValue.vt != VT_NULL && varValue.vt == VT_DATE )
			{
				pHouseRecord->m_oleDate = varValue;
				pHouseRecord->m_szSQLDate = pHouseRecord->m_oleDate.Format(_T("%d-%m-%Y %H:%M:%S"));
			}

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_PRICE)->Value;
			pHouseRecord->m_szPrice = (varValue.vt != VT_NULL)?varValue.bstrVal:_T("");

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_NUMPRICE)->Value;
			pHouseRecord->m_nPrice = (varValue.vt != VT_NULL)?varValue.intVal:0;

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_STATUS)->Value;
			pHouseRecord->m_szStatus = (varValue.vt != VT_NULL)?varValue.bstrVal:_T("");

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_DESC)->Value;
			pHouseRecord->m_szDescription = (varValue.vt != VT_NULL)?varValue.bstrVal:_T("");

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_DETAIL_DESC)->Value;
			pHouseRecord->m_szDetailDescription = (varValue.vt != VT_NULL)?varValue.bstrVal:_T("");

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_AGENT)->Value;
			pHouseRecord->m_szAgent = (varValue.vt != VT_NULL)?varValue.bstrVal:_T("");

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_AGENT_ADDRESS)->Value;
			pHouseRecord->m_szAgentAddress = (varValue.vt != VT_NULL)?varValue.bstrVal:_T("");

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_AGENT_LOCATION)->Value;
			pHouseRecord->m_szAgentLocation = (varValue.vt != VT_NULL)?varValue.bstrVal:_T("");

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_AGENT_TELEPHONE)->Value;
			pHouseRecord->m_szAgentTel = (varValue.vt != VT_NULL)?varValue.bstrVal:_T("");

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_TITLE)->Value;
			pHouseRecord->m_szTitle = (varValue.vt != VT_NULL)?varValue.bstrVal:_T("");

			varValue.ClearToZero();
			varValue = pRecordSets->Fields->GetItem(COL_SUBTITLE)->Value;
			pHouseRecord->m_szSubTitle = (varValue.vt != VT_NULL)?varValue.bstrVal:_T("");

			RecordsList.AddTail(pHouseRecord);
			hrx = pRecordSets->MoveNext();
			nCount++;
		}

		if( pNewRecord )	//do we want to check whether a new record needs to be added?
		{
			pRecordCopy = new CHouseRecord();
			CopyRecord(pNewRecord,pRecordCopy);
			RecordsList.AddTail(pRecordCopy);
		}

		//now go thru the list and prepare the history
		int nStart=0;
		int nLastItemChangeFlag = 0;
		while( nStart <= (RecordsList.GetCount()-1) )
		{
			CHouseRecord valuesRecord;
			POSITION pos = RecordsList.FindIndex(nStart);
			if( !pos )
				break;

			nHistoryMask = TEST_MASK_NOCHANGE;
			int nItemsChanged = TEST_MASK_NOCHANGE;
			CString szHistory;
			CHouseRecord *pRecord = (CHouseRecord *)RecordsList.GetAt(pos);
			if( !pRecord )
				break;

			BOOL bLessThanEqual = (nStart == (RecordsList.GetCount()-1)) && (pRecordCopy != NULL);

			if( nStart == 0 )
				pRecord->m_szHistory = _T("• initial entry found.");
			else
			{
				if( !pRecord->m_szPrice.IsEmpty() )
				{
					GetLastValue(pRecord->m_nKey,COL_PRICE,pRecord->m_oleDate,valuesRecord.m_szPrice,bLessThanEqual);
					nHistoryMask |= TEST_MASK_PRICE;
				}

				if( !pRecord->m_szStatus.IsEmpty() )
				{
					GetLastValue(pRecord->m_nKey,COL_STATUS,pRecord->m_oleDate,valuesRecord.m_szStatus,bLessThanEqual);
					nHistoryMask |= TEST_MASK_STATUS;
				}

				if( !pRecord->m_szDescription.IsEmpty() )
				{
					GetLastValue(pRecord->m_nKey,COL_DESC,pRecord->m_oleDate,valuesRecord.m_szDescription,bLessThanEqual);
					nHistoryMask |= TEST_MASK_DESC;
				}

				if( !pRecord->m_szDetailDescription.IsEmpty() )
				{
					GetLastValue(pRecord->m_nKey,COL_DETAIL_DESC,pRecord->m_oleDate,valuesRecord.m_szDetailDescription,bLessThanEqual);
					nHistoryMask |= TEST_MASK_DETAIL_DESC;
				}

				if( !pRecord->m_szAgent.IsEmpty() )
				{
					GetLastValue(pRecord->m_nKey,COL_AGENT,pRecord->m_oleDate,valuesRecord.m_szAgent,bLessThanEqual);
					nHistoryMask |= TEST_MASK_AGENT;
				}

				if( !pRecord->m_szAgentAddress.IsEmpty() )
				{
					GetLastValue(pRecord->m_nKey,COL_AGENT_ADDRESS,pRecord->m_oleDate,valuesRecord.m_szAgentAddress,bLessThanEqual);
					nHistoryMask |= TEST_MASK_AGENTADDRESS;
				}

				if( !pRecord->m_szAgentLocation.IsEmpty() )
				{
					GetLastValue(pRecord->m_nKey,COL_AGENT_LOCATION,pRecord->m_oleDate,valuesRecord.m_szAgentLocation,bLessThanEqual);
					nHistoryMask |= TEST_MASK_AGENTLOCATION;
				}

				if( !pRecord->m_szAgentTel.IsEmpty() )
				{
					GetLastValue(pRecord->m_nKey,COL_AGENT_TELEPHONE,pRecord->m_oleDate,valuesRecord.m_szAgentTel,bLessThanEqual);
					nHistoryMask |= TEST_MASK_AGENTTEL;
				}

				if( !pRecord->m_szTitle.IsEmpty() )
				{
					GetLastValue(pRecord->m_nKey,COL_TITLE,pRecord->m_oleDate,valuesRecord.m_szTitle,bLessThanEqual);
					nHistoryMask |= TEST_MASK_TITLE;
				}

				if( !pRecord->m_szSubTitle.IsEmpty() )
				{
					GetLastValue(pRecord->m_nKey,COL_SUBTITLE,pRecord->m_oleDate,valuesRecord.m_szSubTitle,bLessThanEqual);
					nHistoryMask |= TEST_MASK_SUBTITLE;
				}

				if( nHistoryMask != TEST_MASK_NOCHANGE )
					nItemsChanged = IsChanged(pRecord,&valuesRecord,nHistoryMask,szHistory);

				if( nItemsChanged == TEST_MASK_NOCHANGE || szHistory.IsEmpty() )
					pRecord->m_bShowRecord = FALSE;
				else
					pRecord->m_szHistory = szHistory;

				if( nStart == (RecordsList.GetCount()-1) && pRecordCopy && pNewRecord )
				{
					POSITION posPrev = RecordsList.FindIndex((RecordsList.GetCount()-2));
					CHouseRecord *pPrevRecord = NULL;
					if( posPrev )
					{
						pPrevRecord = RecordsList.GetAt(posPrev);
						if( pPrevRecord )
						{
							CString szTmp;
							nLastItemChangeFlag = IsChanged(pRecordCopy,pPrevRecord,TEST_MASK_ALL,szTmp);
							if( nLastItemChangeFlag == TEST_MASK_NOCHANGE )
								pRecordCopy->m_bShowRecord = FALSE;
						}
					}
				}
			}

			++nStart;
		}

		if( pRecordCopy && pRecordCopy->m_bShowRecord && nLastItemChangeFlag !=TEST_MASK_NOCHANGE )
			InsertRecord(pRecordCopy,&nLastItemChangeFlag);

	}
	catch(_com_error& e)
	{
		hrx.Set(e.Error());
		nCount = 0;
	}

	return nCount;

}

BOOL CPropBeeDB::GetLastValue(long nKey, CString szColumn, COleDateTime& timeBefore, CString& szResult, BOOL bLessAndEqualTo)
{
	CString szQuery, szMsg, szError;
	CString szTableName;
	CComBSTR bstrQuery;
	HRX hrx;
	BOOL bRetVal = FALSE;

	try
	{
		_RecordsetPtr pRecordSets;
		CString szTmpDate = timeBefore.Format(_T("%Y-%m-%d %H:%M:%S"));

		if( IsUsingSQLServer() )
		{
			szQuery.Format(_T("SELECT TOP 1 %s FROM %s where %s= %d and %s IS NOT NULL and %s %s CONVERT(datetime,'%s', 20) order by %s desc"),
				szColumn,m_szTableName,COL_PROPID,nKey,szColumn,COL_DATE,bLessAndEqualTo?_T("<="):_T("<"),szTmpDate,COL_DATE);
		}
		else if( IsUsingJetOLEDB() )
		{
			szQuery.Format(_T("SELECT TOP 1 %s FROM %s where %s= %d and %s IS NOT NULL and %s %s #%s# order by %s desc"),
				szColumn,m_szTableName,COL_PROPID,nKey,szColumn,COL_DATE,bLessAndEqualTo?_T("<="):_T("<"),szTmpDate,COL_DATE);
		}

		bstrQuery = szQuery;

		ExecuteSQLQuery(bstrQuery,pRecordSets,szError);
		if( pRecordSets->GetADO_EOF() != VARIANT_TRUE )
		{
			CComVariant varValue, varColumn;
			varColumn.ChangeType(VT_BSTR);
			varColumn.bstrVal = CComBSTR(szColumn);
			varValue = pRecordSets->Fields->GetItem(varColumn)->Value;
			if( varValue.vt != VT_NULL )
			{
				if( varValue.vt == VT_DATE )
				{
					COleDateTime DateInfo(varValue);
					szResult = DateInfo.Format(_T("%Y-%m-%d %H:%M:%S"));
				}
				else if( varValue.vt == VT_BSTR )
				{
					szResult = varValue.bstrVal;
				}
			}
		}
	}
	catch(_com_error& e)
	{
		hrx.Set(e.Error());
		bRetVal = FALSE;
	}

	return bRetVal;

}

BOOL CPropBeeDB::GetLastEntryDate(long nKey,COleDateTime& oleLastDate)
{
	CString szQuery, szMsg, szError;
	CComBSTR bstrQuery;
	HRX hrx;
	BOOL bRetVal = FALSE;

	try
	{
		_RecordsetPtr pRecordSets;

		COleDateTime timeNow = COleDateTime::GetCurrentTime();
		CString szTmpDate = timeNow.Format(_T("%Y-%m-%d %H:%M:%S"));

		if( m_bUsingSQLServer )
		{
			szQuery.Format(_T("SELECT TOP 1 %s FROM %s where %s= %d and %s <= CONVERT(datetime,'%s', 20) order by %s desc"), 
				COL_DATE,m_szTableName,COL_PROPID,nKey,COL_DATE,szTmpDate,COL_DATE);
		}
		else
		{
			szQuery.Format(_T("SELECT TOP 1 %s FROM %s where %s= %d and %s <= #%s# order by %s desc"), 
				COL_DATE,m_szTableName,COL_PROPID,nKey,COL_DATE,szTmpDate,COL_DATE);
		}

		bstrQuery = szQuery;

		ExecuteSQLQuery(bstrQuery,pRecordSets,szError);
		if( pRecordSets->GetADO_EOF() != VARIANT_TRUE )
		{
			CComVariant varValue;
			CComVariant varColumn(CComBSTR(COL_DATE));
			
			varValue = pRecordSets->Fields->GetItem(varColumn)->Value;
			if( varValue.vt != VT_NULL && varValue.vt == VT_DATE )
			{
				oleLastDate = varValue.date;
				bRetVal = TRUE;
			}
		}
	}
	catch(_com_error& e)
	{
		hrx.Set(e.Error());
	}

	return bRetVal;

}

void CPropBeeDB::FreeRecordMemory(CAtlList<CHouseRecord *>& ListRecords)
{
	POSITION pos = ListRecords.GetHeadPosition();
	while(pos != NULL)
	{
		CHouseRecord *pHouseRecord = ListRecords.GetNext(pos);
		if( pHouseRecord )
		{
			delete pHouseRecord;pHouseRecord=NULL;
		}
	}

	ListRecords.RemoveAll();
}


int CPropBeeDB::IsChanged(CHouseRecord *pCurHouseRecord, CHouseRecord *pLastHouseRecord, int nItemsToCheck, CString& szChangeHistory)
{
	int nRetVal = TEST_MASK_NOCHANGE;
	szChangeHistory = _T("");

	if( ((nItemsToCheck & TEST_MASK_PRICE) || (nItemsToCheck == TEST_MASK_ALL)) )
	{
		if( pCurHouseRecord->m_szPrice.CompareNoCase(pLastHouseRecord->m_szPrice) != 0 )	//price change?
		{
			nRetVal |= TEST_MASK_PRICE;
			if( !pLastHouseRecord->m_szPrice.IsEmpty() )
				szChangeHistory.AppendFormat(_T("• price changed: from <b>%s</b> to <b>%s</b>"),pLastHouseRecord->m_szPrice, pCurHouseRecord->m_szPrice);
			else
				szChangeHistory.AppendFormat(_T("• price changed to: <b>%s</b>"),pCurHouseRecord->m_szPrice);
		}
		else
			nRetVal &=~ TEST_MASK_PRICE;
	}

	if( ((nItemsToCheck & TEST_MASK_STATUS) || (nItemsToCheck == TEST_MASK_ALL)))
	{
		if( pCurHouseRecord->m_szStatus.CompareNoCase(pLastHouseRecord->m_szStatus) != 0 )	//status change?
		{
			nRetVal |= TEST_MASK_STATUS;
			if( !szChangeHistory.IsEmpty() )
				szChangeHistory.AppendFormat(_T("<br>"));

			if( !pLastHouseRecord->m_szStatus.IsEmpty() )
				szChangeHistory.AppendFormat(_T("• status changed from <b>%s</b> to <b>%s</b>"),pLastHouseRecord->m_szStatus, pCurHouseRecord->m_szStatus);
			else
				szChangeHistory.AppendFormat(_T("• status changed to: <b>%s</b>"),pCurHouseRecord->m_szStatus);
		}
		else
			nRetVal &=~ TEST_MASK_STATUS;
	}

	if( ((nItemsToCheck & TEST_MASK_TITLE) || (nItemsToCheck == TEST_MASK_ALL)))
	{
		if( pCurHouseRecord->m_szTitle.CompareNoCase(pLastHouseRecord->m_szTitle) != 0 )	//status change?
		{
			nRetVal |= TEST_MASK_TITLE;
			if( !szChangeHistory.IsEmpty() )
				szChangeHistory.AppendFormat(_T("<br>"));

			if( !pLastHouseRecord->m_szTitle.IsEmpty() )
				szChangeHistory.AppendFormat(_T("• title changed from %s to <b>%s</b>"),pLastHouseRecord->m_szTitle, pCurHouseRecord->m_szTitle);
		}
		else
			nRetVal &=~ TEST_MASK_TITLE;
	}

	if( ((nItemsToCheck & TEST_MASK_SUBTITLE) || (nItemsToCheck == TEST_MASK_ALL)) )
	{
		if( pCurHouseRecord->m_szSubTitle.CompareNoCase(pLastHouseRecord->m_szSubTitle) != 0 )	//status change?
		{
			nRetVal |= TEST_MASK_SUBTITLE;
			if( !szChangeHistory.IsEmpty() )
				szChangeHistory.AppendFormat(_T("<br>"));

			if( !pLastHouseRecord->m_szSubTitle.IsEmpty() )
				szChangeHistory.AppendFormat(_T("• subtitle changed from %s to <b>%s</b>"),pLastHouseRecord->m_szSubTitle, pCurHouseRecord->m_szSubTitle);
		}
		else
			nRetVal &=~ TEST_MASK_SUBTITLE;
	}

	if( ((nItemsToCheck & TEST_MASK_DESC) || (nItemsToCheck == TEST_MASK_ALL)) )
	{
		CString szTmp;
		if( FindAndFormatDifferences(pCurHouseRecord->m_szDescription,pLastHouseRecord->m_szDescription,szTmp) )
		{
			nRetVal |= TEST_MASK_DESC;
			if( !szChangeHistory.IsEmpty() )
				szChangeHistory.AppendFormat(_T("<br>"));

			szChangeHistory.AppendFormat(_T("• brief description changed:<br>%s<br>"),szTmp);
		}
		else
			nRetVal &=~ TEST_MASK_DESC;
	}

	if( ((nItemsToCheck & TEST_MASK_DETAIL_DESC) || (nItemsToCheck == TEST_MASK_ALL)) )
	{
		CString szTmp;
		if( FindAndFormatDifferences(pCurHouseRecord->m_szDetailDescription,pLastHouseRecord->m_szDetailDescription,szTmp) )
		{
			nRetVal |= TEST_MASK_DETAIL_DESC;
			if( !szChangeHistory.IsEmpty() )
				szChangeHistory.AppendFormat(_T("<br>"));

			szChangeHistory.AppendFormat(_T("• detailed description changed:<br>%s<br>"),szTmp);
		}
		else
			nRetVal &=~ TEST_MASK_DETAIL_DESC;
	}

	if( ((nItemsToCheck & TEST_MASK_AGENT) || (nItemsToCheck == TEST_MASK_ALL)))
	{
		if( pCurHouseRecord->m_szAgent.CompareNoCase(pLastHouseRecord->m_szAgent) != 0 )	//agent change?
		{
			nRetVal |= TEST_MASK_AGENT;
			if( !szChangeHistory.IsEmpty() )
				szChangeHistory.AppendFormat(_T("<br>"));

			if( !pLastHouseRecord->m_szAgent.IsEmpty() )
				szChangeHistory.AppendFormat(_T("• estate agent changed: from %s to %s"),pLastHouseRecord->m_szAgent, pCurHouseRecord->m_szAgent);
		}
		else
			nRetVal &=~ TEST_MASK_AGENT;
	}

	if( ((nItemsToCheck & TEST_MASK_AGENTADDRESS) || (nItemsToCheck == TEST_MASK_ALL)))
	{
		if( pCurHouseRecord->m_szAgentAddress.CompareNoCase(pLastHouseRecord->m_szAgentAddress) != 0 )	//agent change?
		{
			nRetVal |= TEST_MASK_AGENTADDRESS;
			if( !szChangeHistory.IsEmpty() )
				szChangeHistory.AppendFormat(_T("<br>"));

			if( !pLastHouseRecord->m_szAgentAddress.IsEmpty() )
				szChangeHistory.AppendFormat(_T("• estate agent address changed: from %s to %s"),pLastHouseRecord->m_szAgentAddress, pCurHouseRecord->m_szAgentAddress);
		}
		else
			nRetVal &=~ TEST_MASK_AGENTADDRESS;
	}

	if( ((nItemsToCheck & TEST_MASK_AGENTLOCATION) || (nItemsToCheck == TEST_MASK_ALL)))
	{
		if( pCurHouseRecord->m_szAgentLocation.CompareNoCase(pLastHouseRecord->m_szAgentLocation) != 0 )	//agent change?
		{
			nRetVal |= TEST_MASK_AGENTLOCATION;
			if( !szChangeHistory.IsEmpty() )
				szChangeHistory.AppendFormat(_T("<br>"));

			if( !pLastHouseRecord->m_szAgentLocation.IsEmpty() )
				szChangeHistory.AppendFormat(_T("• estate agent location changed: from %s to %s"),pLastHouseRecord->m_szAgentLocation, pCurHouseRecord->m_szAgentLocation);
		}
		else
			nRetVal &=~ TEST_MASK_AGENTLOCATION;
	}

	if( ((nItemsToCheck & TEST_MASK_AGENTTEL) || (nItemsToCheck == TEST_MASK_ALL)))
	{
		if( pCurHouseRecord->m_szAgentTel.CompareNoCase(pLastHouseRecord->m_szAgentTel) != 0 )	//agent change?
		{
			nRetVal |= TEST_MASK_AGENTTEL;
			if( !szChangeHistory.IsEmpty() )
				szChangeHistory.AppendFormat(_T("<br>"));

			if( !pLastHouseRecord->m_szAgentTel.IsEmpty() )
				szChangeHistory.AppendFormat(_T("• estate agent telephone changed: from %s to %s"),pLastHouseRecord->m_szAgentTel, pCurHouseRecord->m_szAgentTel);
		}
		else
			nRetVal &=~ TEST_MASK_AGENTTEL;
	}

	return nRetVal;
}

//this method will import the data from CSV generated by old property bee but make sure all the fields are ticked while generating the CSV
//Reference,Sampled On,Agent,Agents Address,Agents Location,Agents Telephone,Brief Description,Detailed Description,Currency,Price,Term,Status,Subtitle,Title,First Seen,Last Seen

int CPropBeeDB::UploadDataFromPropertyBeeCSV(CString szSourceCSV)
{
	int nCount=0;
	CStdioFile fileCSV;

	if( !fileCSV.Open(szSourceCSV,CFile::modeRead|CFile::typeText|CFile::shareDenyNone) )
		return nCount;

	CAtlList<CHouseRecord *> RecordsList;
	CMapStringToString mapKeysProcessed;
	int nLine=0;

	while(TRUE)
	{
		CString szRecord;
		if( !fileCSV.ReadString(szRecord) )
			break;

		if( nLine == 0 )	//header row? then ignore it
		{
			nLine++;
			continue;
		}

		std::vector<std::wstring> row;
		ParseRecord(row, szRecord, ',');	//parse the record
		CHouseRecord *pRecord = new CHouseRecord();

		pRecord->m_szKey = row.at(0).c_str();	//property id
		pRecord->m_szSQLDate = row.at(1).c_str();	//01,03,2008 21:39
		pRecord->m_szSQLDate.Replace(_T(","),_T("-"));
		pRecord->m_szSQLDate.Append(_T(":00"));
		pRecord->m_szAgent.Format(_T("%s %s %s %s"),row.at(2).c_str(),row.at(3).c_str(),row.at(4).c_str(),row.at(5).c_str());
		pRecord->m_szDescription = row.at(6).c_str();	//brief description
		pRecord->m_szDetailDescription = row.at(7).c_str();	//detailed description
		pRecord->m_szPrice = row.at(9).c_str();	//price
		pRecord->m_szStatus = row.at(11).c_str();	//status
		pRecord->m_szSubTitle = row.at(13).c_str();	//address

		//we will skip those entries which are blank
		if( pRecord->m_szAgent.IsEmpty() && pRecord->m_szDescription.IsEmpty() && pRecord->m_szDetailDescription.IsEmpty() &&
			pRecord->m_szPrice.IsEmpty() &&  pRecord->m_szStatus.IsEmpty() )
			continue;


		//now let us format it the way we want
		pRecord->m_szKey.Replace(_T("rm"),_T(""));
		pRecord->m_nKey = _ttoi(pRecord->m_szKey);
		mapKeysProcessed.SetAt(pRecord->m_szKey,_T("done"));

		RecordsList.AddTail(pRecord);
		nLine++;

		//lets try it out
		InsertRecord(pRecord);
	}

	fileCSV.Close();
	FreeRecordMemory(RecordsList);

	return nCount;
}

void CPropBeeDB::ParseRecord(std::vector<std::wstring> &record, const CString& line, TCHAR delimiter)
{
	int linepos=0;
	int inquotes=false;
	TCHAR c;
	int linemax=line.GetLength();
	std::wstring curstring;
	record.clear();

	while(line[linepos]!=0 && linepos < linemax)
	{
		c = line[linepos];
		if (!inquotes && curstring.length()==0 && c=='"')
		{
			//beginquotechar
			inquotes=true;
		}
		else if (inquotes && c==_T('"') )
		{
			//quotechar
			if ( (linepos+1 <linemax) && (line[linepos+1]=='"') )
			{
				//encountered 2 double quotes in a row (resolves to 1 double quote)
				curstring.push_back(c);
				linepos++;
			}
			else
			{
				//endquotechar
				inquotes=false;
			}
		}
		else if (!inquotes && c==delimiter)
		{
			//end of field
			record.push_back( curstring );
			curstring.clear();
		}
		else if (!inquotes && (c==_T('\r') || c==_T('\n')) )
		{
			record.push_back( curstring );
			return;
		}
		else
		{
			curstring.push_back(c);
		}
		linepos++;
	}
	record.push_back( curstring );

	return;
}

void CPropBeeDB::DumpRecords(CString szFileName, CAtlList<CHouseRecord *>& RecordList)
{
	CStdioFile File;
	File.Open(szFileName, CFile::modeCreate|CFile::modeWrite|CFile::typeText);

	POSITION pos = RecordList.GetHeadPosition();
	CString szContent;

	while(pos != NULL)
	{
		CHouseRecord *pHouseRecord = RecordList.GetNext(pos);
		if( pHouseRecord )
		{
			szContent.Format(_T("ID:%s\t ID(n):%d\n\tDate:%s\n\tStatus:%s\n\tPrice:%s\n\tAgent:%s\n\tDescription:%s\n\tHistory:%s\n\n\n"),
				pHouseRecord->m_szKey,
				pHouseRecord->m_nKey,
				pHouseRecord->m_szSQLDate,
				pHouseRecord->m_szStatus,
				pHouseRecord->m_szPrice,
				pHouseRecord->m_szAgent,
				pHouseRecord->m_szDescription,
				pHouseRecord->m_szHistory);

			File.WriteString(szContent);
		}
	}

	File.Close();
}

void CPropBeeDB::CopyRecord(CHouseRecord *pSource, CHouseRecord *pTarget)
{
	if( !pSource || !pTarget )
		return;

	pTarget->m_nKey						= pSource->m_nKey;			
	pTarget->m_szKey					= pSource->m_szKey;		
	pTarget->m_szPrice					= pSource->m_szPrice;
	pTarget->m_nPrice					= pSource->m_nPrice;
	pTarget->m_szDescription			= pSource->m_szDescription;
	pTarget->m_szDetailDescription		= pSource->m_szDetailDescription;
	pTarget->m_szAgent					= pSource->m_szAgent;	
	pTarget->m_szAgentAddress			= pSource->m_szAgentAddress;	
	pTarget->m_szAgentLocation			= pSource->m_szAgentLocation;
	pTarget->m_szAgentTel				= pSource->m_szAgentTel;
	pTarget->m_szStatus					= pSource->m_szStatus;	
	pTarget->m_oleDate					= pSource->m_oleDate;	
	pTarget->m_szSQLDate				= pSource->m_szSQLDate;
	pTarget->m_szHistory				= pSource->m_szHistory;
	pTarget->m_bHighlight				= pSource->m_bHighlight;
	pTarget->m_bShowRecord				= pSource->m_bShowRecord;
	pTarget->m_szTitle					= pSource->m_szTitle;
	pTarget->m_szSubTitle				= pSource->m_szSubTitle;
}

BOOL CPropBeeDB::FindAndFormatDifferences(CString& szOld, CString& szNew, CString& szDifference)
{
	CStringArray arrayDiff;
	CString szOldToken, szNewToken, szDiffToken;
	int nOldPos=0, nNewPos=0;
	BOOL bRetVal = FALSE;

	szOldToken = szOld.Tokenize(_T(" "), nOldPos);
	szNewToken = szNew.Tokenize(_T(" "), nNewPos);

	while( TRUE )
	{
		if( szOldToken.IsEmpty() && szNewToken.IsEmpty() )    //if both the tokens are empty it means we have exhausted our words
			break;

		if( !szOldToken.IsEmpty() && !szNewToken.IsEmpty() && szOldToken.CompareNoCase(szNewToken) != 0 )           //different
		{
			szDiffToken.Format(_T("<del>%s</del>"), szOldToken);
			arrayDiff.Add(szDiffToken);
			arrayDiff.Add(szNewToken);
			bRetVal = TRUE;	//flag that there are differences
		}
		else if( !szOldToken.IsEmpty() && !szNewToken.IsEmpty() && szOldToken.CompareNoCase(szNewToken) == 0 )           //same?
		{
			arrayDiff.Add(szNewToken);    //we will just add the new token
		}
		else if( !szOldToken.IsEmpty() && szNewToken.IsEmpty() )
		{
			szDiffToken.Format(_T("<del>%s</del>"), szOldToken);
			arrayDiff.Add(szDiffToken);
		}
		else if( szOldToken.IsEmpty() && !szNewToken.IsEmpty() )
			arrayDiff.Add(szNewToken);

		if( nOldPos > 0 )
			szOldToken = szOld.Tokenize(_T(" "), nOldPos);
		else
			szOldToken = _T("");

		if( nNewPos > 0 )
			szNewToken = szNew.Tokenize(_T(" "), nNewPos);
		else
			szNewToken = _T("");
	}

	for( int nCount=0; nCount < arrayDiff.GetCount(); nCount++)
	{
		szDifference.Append(arrayDiff.GetAt(nCount) + _T(" "));
	}

	return bRetVal;
}

int CPropBeeDB::UploadDataFromPropertyBeeDB(CString szSourceDB, CProgressCtrl *pProgressCtrl)
{
	int nCount=0;

	CString szQuery = _T("SELECT Property.Reference as Ref,Sample.Sampled_On as Date,Agent.Data as Agent,Agents_Address.Data as AgentAddress,Agents_Location.Data as AgentLocation,");
	szQuery.Append(_T("Agents_Telephone.Data as AgentTelephone,Price.Data as Price,Price.Value as NumPrice,Status.Data as Status,Title.Data as Title ,Subtitle.Data as SubTitle,"));
	szQuery.Append(_T("Brief_Description.Data as BriefDesc,Detailed_Description.Data as DetailDesc FROM Sample "));
	szQuery.Append(_T("LEFT OUTER JOIN Property ON (Sample.Property_ID = Property.Property_ID) ")); 
	szQuery.Append(_T("	LEFT OUTER JOIN Agent ON (Sample.Agent_ID = Agent.Agent_ID) ")); 
	szQuery.Append(_T("	LEFT OUTER JOIN Agents_Address ON (Sample.Agents_Address_ID = Agents_Address.Agents_Address_ID) ")); 
	szQuery.Append(_T("	LEFT OUTER JOIN Agents_Location ON (Sample.Agents_Location_ID = Agents_Location.Agents_Location_ID) ")); 
	szQuery.Append(_T("	LEFT OUTER JOIN Agents_Telephone ON (Sample.Agents_Telephone_ID = Agents_Telephone.Agents_Telephone_ID) ")); 
	szQuery.Append(_T("	LEFT OUTER JOIN Brief_Description ON (Sample.Brief_Description_ID = Brief_Description.Brief_Description_ID) ")); 
	szQuery.Append(_T("	LEFT OUTER JOIN Price ON (Sample.Price_ID = Price.Price_ID) ")); 
	szQuery.Append(_T("	LEFT OUTER JOIN Status ON (Sample.Status_ID = Status.Status_ID) ")); 
	szQuery.Append(_T("	LEFT OUTER JOIN Title ON (Sample.Title_ID = Title.Title_ID) ")); 
	szQuery.Append(_T("	LEFT OUTER JOIN Subtitle ON (Sample.SubTitle_ID = SubTitle.SubTitle_ID) ")); 
	szQuery.Append(_T("	LEFT OUTER JOIN Detailed_Description ON (Sample.Detailed_Description_ID = Detailed_Description.Detailed_Description_ID) "));       
	szQuery.Append(_T("	ORDER by Sample.Property_Id, Sample.Sampled_On ASC;"));

	try
	{
		CHouseRecord HouseRecord;
		CppSQLite3DB sqlLiteDB;
		sqlLiteDB.open(szSourceDB);
		CppSQLite3Query sqlQuery = sqlLiteDB.execQuery(szQuery);


		while( !sqlQuery.eof() )
		{
			int nMask = TEST_MASK_NOCHANGE;

			HouseRecord.m_szKey = sqlQuery.fieldValue(_T("Ref"));
			HouseRecord.m_szKey.Replace(_T("rm"),_T(""));
			HouseRecord.m_nKey = _ttoi(HouseRecord.m_szKey);
			HouseRecord.m_szSQLDate = sqlQuery.fieldValue(_T("Date"));
			HouseRecord.m_szAgent = sqlQuery.fieldValue(_T("Agent"));HouseRecord.m_szAgent.Trim();
			HouseRecord.m_szAgent.Replace(_T("'"),_T("&#39;"));
			HouseRecord.m_szAgent.Replace(_T("\""),_T("&quot;"));

			HouseRecord.m_szAgentAddress = sqlQuery.fieldValue(_T("AgentAddress"));HouseRecord.m_szAgentAddress.Trim();
			HouseRecord.m_szAgentAddress.Replace(_T("'"),_T("&#39;"));
			HouseRecord.m_szAgentAddress.Replace(_T("\""),_T("&quot;"));

			HouseRecord.m_szAgentLocation = sqlQuery.fieldValue(_T("AgentLocation"));HouseRecord.m_szAgentLocation.Trim();
			HouseRecord.m_szAgentLocation.Replace(_T("'"),_T("&#39;"));
			HouseRecord.m_szAgentLocation.Replace(_T("\""),_T("&quot;"));

			HouseRecord.m_szAgentTel = sqlQuery.fieldValue(_T("AgentTelephone"));HouseRecord.m_szAgentTel.Trim();
			HouseRecord.m_szAgentTel.Replace(_T("'"),_T("&#39;"));
			HouseRecord.m_szAgentTel.Replace(_T("\""),_T("&quot;"));

			HouseRecord.m_szPrice = sqlQuery.fieldValue(_T("Price"));HouseRecord.m_szPrice.Trim();
			CString szNumPrice = sqlQuery.fieldValue(_T("NumPrice"));
			HouseRecord.m_nPrice = _ttoi(szNumPrice);
			HouseRecord.m_szStatus = sqlQuery.fieldValue(_T("Status"));HouseRecord.m_szStatus.Trim();
			HouseRecord.m_szStatus.Replace(_T("'"),_T("&#39;"));
			HouseRecord.m_szStatus.Replace(_T("\""),_T("&quot;"));

			HouseRecord.m_szTitle = sqlQuery.fieldValue(_T("Title"));HouseRecord.m_szTitle.Trim();
			HouseRecord.m_szTitle.Replace(_T("'"),_T("&#39;"));
			HouseRecord.m_szTitle.Replace(_T("\""),_T("&quot;"));

			HouseRecord.m_szSubTitle = sqlQuery.fieldValue(_T("SubTitle"));HouseRecord.m_szSubTitle.Trim();
			HouseRecord.m_szSubTitle.Replace(_T("'"),_T("&#39;"));
			HouseRecord.m_szSubTitle.Replace(_T("\""),_T("&quot;"));

			HouseRecord.m_szDescription = sqlQuery.fieldValue(_T("BriefDesc"));HouseRecord.m_szDescription.Trim();
			HouseRecord.m_szDescription.Replace(_T("'"),_T("&#39;"));
			HouseRecord.m_szDescription.Replace(_T("\""),_T("&quot;"));

			HouseRecord.m_szDetailDescription = sqlQuery.fieldValue(_T("DetailDesc"));HouseRecord.m_szDetailDescription.Trim();
			HouseRecord.m_szDetailDescription.Replace(_T("'"),_T("&#39;"));
			HouseRecord.m_szDetailDescription.Replace(_T("\""),_T("&quot;"));
			

			//lets try it out
			BOOL bRecordInserted = InsertRecord(&HouseRecord);
			++nCount;
			sqlQuery.nextRow();
			if( pProgressCtrl )
				pProgressCtrl->StepIt();
		}
		sqlLiteDB.close();
	}
	catch (CppSQLite3Exception& e)
	{
	}
	catch(...)
	{
	}

	//DumpRecords(_T("c:\\records.txt"),RecordsList);

	return nCount;
}

int CPropBeeDB::GetInsertMask(CHouseRecord* pHouseRecord)
{
	int nMask = TEST_MASK_NOCHANGE;

	if( !pHouseRecord->m_szPrice.IsEmpty() )
		nMask = nMask|TEST_MASK_PRICE;

	if( !pHouseRecord->m_szStatus.IsEmpty() )
		nMask = nMask|TEST_MASK_STATUS;

	if( !pHouseRecord->m_szAgent.IsEmpty() )
		nMask = nMask|TEST_MASK_AGENT;

	if( !pHouseRecord->m_szAgentAddress.IsEmpty() )
		nMask = nMask|TEST_MASK_AGENTADDRESS;

	if( !pHouseRecord->m_szAgentLocation.IsEmpty() )
		nMask = nMask|TEST_MASK_AGENTLOCATION;

	if( !pHouseRecord->m_szAgentTel.IsEmpty() )
		nMask = nMask|TEST_MASK_AGENTTEL;

	if( !pHouseRecord->m_szTitle.IsEmpty() )
		nMask = nMask|TEST_MASK_TITLE;

	if( !pHouseRecord->m_szSubTitle.IsEmpty() )
		nMask = nMask|TEST_MASK_SUBTITLE;

	if( !pHouseRecord->m_szDescription.IsEmpty() )
		nMask = nMask|TEST_MASK_DESC;

	if( !pHouseRecord->m_szDetailDescription.IsEmpty() )
		nMask = nMask|TEST_MASK_DETAIL_DESC;

	return nMask;
}

BOOL CPropBeeDB::CreateSQLServerDatabase(CString szServer,CString szPath, CString szDBName)
{
	BOOL bDatabaseCreated = FALSE;
	HRX hrx;
	
	if( !szDBName.IsEmpty() )
		m_szSQLServerDBName = szDBName;

	try
	{
		CString szQuery,szErrorMsg;
		CString szDBFileName, szLogFileName;

		szDBFileName.Format(_T("%s%s.mdf"),szPath,m_szSQLServerDBName);
		szLogFileName.Format(_T("%s%s_log.ldf"),szPath,m_szSQLServerDBName);

		szQuery.Format(_T("CREATE DATABASE [%s] ON  PRIMARY "),m_szSQLServerDBName);
		szQuery.AppendFormat(_T("( NAME = N'%s', FILENAME = N'%s' , SIZE = 5120KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )"),m_szSQLServerDBName,szDBFileName);
		szQuery.Append(_T("LOG ON "));
		szQuery.AppendFormat(_T("( NAME = N'%s_log', FILENAME = N'%s' , SIZE = 1024KB , MAXSIZE = 2048GB , FILEGROWTH = 10%%)"),m_szSQLServerDBName,szLogFileName);
		CComBSTR bstrQuery = szQuery;

		_RecordsetPtr pRecordSets;
		ConnectToSQLServerDatabase(szServer,_T("master"));
		if( IsConnected() )
			bDatabaseCreated = ExecuteSQLQuery(bstrQuery,pRecordSets, szErrorMsg);	//dont check for connection

		//now lets create the tables we need
		if( bDatabaseCreated )
		{
			ConnectToSQLServerDatabase(szServer,m_szSQLServerDBName);
			if( IsConnected() )
			{
				CreateSQLServerTables();
			}
		}
		
		return bDatabaseCreated;
	}
	catch(_com_error& e)
	{
		hrx.Set(e.Error());
	}

	return bDatabaseCreated;
}

