#pragma once

#include <string>
#include <vector>
#include "CppSQLite3U.h"
// CPropBeeDB command target

#define PROPBEE_TABLEPREFIX			_T("Data")

//column names in the data table
#define COL_PROPID				_T("PropertyID")
#define COL_DATE				_T("EntryDate")
#define COL_STATUS				_T("Status")
#define COL_PRICE				_T("Price") 
#define COL_NUMPRICE			_T("NumPrice") 
#define COL_DESC				_T("Description") 
#define COL_DETAIL_DESC			_T("DetailDescription") 
#define COL_AGENT				_T("EstateAgent") 
#define COL_AGENT_ADDRESS		_T("AgentAddress") 
#define COL_AGENT_LOCATION		_T("AgentLocation") 
#define COL_AGENT_TELEPHONE		_T("AgentTel") 
#define COL_TITLE				_T("Title") 
#define COL_SUBTITLE			_T("SubTitle") 

#define TEST_MASK_NOCHANGE			0
#define TEST_MASK_ALL				1
#define TEST_MASK_PRICE				2
#define TEST_MASK_STATUS			4
#define TEST_MASK_AGENT				8
#define TEST_MASK_DESC				16
#define TEST_MASK_DETAIL_DESC		32
#define TEST_MASK_TITLE				64
#define TEST_MASK_SUBTITLE			128
#define TEST_MASK_AGENTTEL			256
#define TEST_MASK_AGENTADDRESS		512
#define TEST_MASK_AGENTLOCATION		1024

class CHouseRecord: public CObject
{
public:
	CHouseRecord()
	{
		m_bHighlight = FALSE;
		m_szStatus = _T("Available");
		m_nPrice = 0;
		m_bShowRecord = TRUE;
	}

public:
	long m_nKey;			//unique reference key
	CString m_szKey;			//for compatibility with old property bee
	CString m_szPrice;
	int m_nPrice;
	CString m_szDescription;
	CString m_szDetailDescription;
	CString m_szAgent;
	CString m_szAgentLocation;
	CString m_szAgentAddress;
	CString m_szAgentTel;
	CString m_szStatus;
	CString m_szTitle;
	CString m_szSubTitle;

	//the below field come from the tables
	CString m_szSQLDate;
	COleDateTime m_oleDate;
	CString m_szHistory;
	BOOL m_bHighlight;
	BOOL m_bShowRecord;
};

class CPropBeeDB : public CObject
{
public:
	CPropBeeDB();
	virtual ~CPropBeeDB();
	enum {DATABASE_SQLSERVER=0,DATABASE_JETOLDEB,DATABASE_SQLITE};
	BOOL CreateSQLServerDatabase(CString szServer,CString szPath, CString szDBName=_T(""));
	BOOL ConnectToSQLServerDatabase(CString szServer=_T(""),CString szDatabase=_T(""));
	BOOL ConnectToJetDatabase(CString szJetDBPath);
	BOOL ConnectToSQLiteDatabase(CString szSqliteDBPath);
	BOOL IsConnected(){return m_bConnected;}
	int UploadDataFromPropertyBeeDB(CString szSourceDB,CProgressCtrl *pProgressCtrl=NULL);
	BOOL ExecuteSQLQuery(CComBSTR bstrQuery, CComVariant& varResult, CString& szErrorMsg);
	BOOL ExecuteSQLQuery(CComBSTR bstrQuery, _RecordsetPtr& pRecordSets, CString& szErrorMsg);
	void FreeRecordMemory(CAtlList<CHouseRecord *>& ListRecords);
	void DumpRecords(CString szFileName, CAtlList<CHouseRecord *>& RecordList);
	int ManageHistory(CHouseRecord *pCurHouseRecord, CAtlList<CHouseRecord *>& HistoryRecords, BOOL bDetailsMode, BOOL& bNewEntryFound);
	BOOL IsUsingSQLServer(){return (m_nDatabaseType==DATABASE_SQLSERVER) ? TRUE:FALSE;}
	BOOL IsUsingJetOLEDB(){return (m_nDatabaseType==DATABASE_JETOLDEB) ? TRUE:FALSE;}
	BOOL IsUsingSQLITE(){return (m_nDatabaseType==DATABASE_SQLITE) ? TRUE:FALSE;}

private:
	BOOL CreateSQLServerTables(void);
	BOOL GetLastEntryDate(long nKey,COleDateTime& oleLastDate);
	BOOL InsertRecord(CHouseRecord *pCurHouseRecord,int *pnMask=NULL);
	int IsChanged(CHouseRecord *pCurHouseRecord, CHouseRecord *pLastHouseRecord, int nItemsToCheck, CString& szChangeHistory);
	int BuildRecordList(long nKey, int nHistoryMask,CAtlList<CHouseRecord *>& RecordsList,CHouseRecord *pNewRecord=NULL);
	int UploadDataFromPropertyBeeCSV(CString szSourceCSV);
	void ParseRecord(std::vector<std::wstring> &record, const CString& line, TCHAR delimiter);
	void CopyRecord(CHouseRecord *pSource, CHouseRecord *pTarget);
	BOOL FindAndFormatDifferences(CString& szOld, CString& szNew, CString& szChanged);
	BOOL GetLastValue(long nKey, CString szColumn, COleDateTime& timeBefore, CString& szResult,BOOL bLessAndEqualTo=FALSE);
	int GetInsertMask(CHouseRecord* pHouseRecord);
	int GetDifferenceMask(CHouseRecord* pRecord);	

private:
	CString m_szSQLServer;
	CString m_szSQLServerDBName;
	CString m_szJetDBPath;
	CString m_szSQLiteDBPath;
	BOOL m_bConnected;
	_ConnectionPtr m_pConn;
	CppSQLite3DB *m_pSQLiteConn;
	CStdioFile m_File;
	BOOL m_bDetailsMode;
	CString m_szTableName;
	int m_nDatabaseType;
};


