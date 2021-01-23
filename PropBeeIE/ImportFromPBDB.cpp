// ModelessDemoDlg.cpp : implementation file
//

#include "stdafx.h"
#include "PropBee.h"
#include "ImportFromPBDB.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CImportPBDBDlg dialog
CImportPBDBDlg::CImportPBDBDlg(CWnd* pParent /*=NULL*/)
: CDialog(CImportPBDBDlg::IDD, pParent)
{
}

void CImportPBDBDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDOK, m_btnImport);
	DDX_Control(pDX, IDC_UPDATE_PROGRESS, m_ctlProgress);
	DDX_Control(pDX, IDC_STATIC_UPDATE, m_ctlStaticText);
}

BEGIN_MESSAGE_MAP(CImportPBDBDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDOK, &CImportPBDBDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CImportPBDBDlg message handlers

BOOL CImportPBDBDlg::OnInitDialog()
{
	CDialog::OnInitDialog();
	return TRUE;  // return TRUE  unless you set the focus to a control
}


void CImportPBDBDlg::OnBnClickedOk()
{
	int nRecImported = 0;
	CString szMsg;

	m_ctlStaticText.ShowWindow(SW_SHOW);
	m_ctlProgress.SetRange32(1,10000);
	m_ctlProgress.SetStep(20);

	CWaitCursor waitCursor;
	CPropBeeDB importDB;

	m_ctlStaticText.SetWindowText(_T("searching Property Bee database..."));
	RecursiveFileFind(_T("C:\\Documents and Settings\\Administrator\\Application Data\\Mozilla\\Firefox\\Profiles"),_T("property-bee.sqlite"),m_szSourcePDBFile);

	m_ctlStaticText.SetWindowText(_T("creating Property Spy database..."));
	m_szSQLServerName = m_szInstallPath + _T("PropertySpy.accdb");
	importDB.ConnectToJetDatabase(m_szSQLServerName);
	if( importDB.IsConnected() && !m_szSourcePDBFile.IsEmpty() )
	{
		m_ctlStaticText.SetWindowText(_T("importing history from Property Bee..."));
		nRecImported = importDB.UploadDataFromPropertyBeeDB(m_szSourcePDBFile,&m_ctlProgress);
		szMsg.Format(_T("%d records successfully imported.Please restart Internet Explorer to see the Property Spy toolbar."),nRecImported);
		AfxMessageBox(szMsg);
	}

	OnOK();
}


void CImportPBDBDlg::RecursiveFileFind(CString szFolder, CString szFileToFind, CString& szPathFound)
{
	//RecursiveFileFind(_T("C:\\Users\\kuldeep_mann\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\*.*"));
	szFolder.Append(_T("\\*.*"));
	CFileFind findFile;
	BOOL bWorking = findFile.FindFile(szFolder);

	while( bWorking )
	{
		m_ctlProgress.StepIt();
		bWorking = findFile.FindNextFile();
		if( findFile.IsDots() )
			continue;
		else if( findFile.IsDirectory() )
			RecursiveFileFind(findFile.GetFilePath(),szFileToFind,szPathFound);

		m_ctlStaticText.SetWindowText(findFile.GetFilePath());
		if( szFileToFind.CompareNoCase( findFile.GetFileName()) == 0 )
		{
			szPathFound = findFile.GetFilePath();
			break;
		}
	}
}
