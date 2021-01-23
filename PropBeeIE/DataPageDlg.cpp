// DataPageDlg.cpp : implementation file
//

#include "stdafx.h"
#include "DataPageDlg.h"


// CDataPageDlg dialog

IMPLEMENT_DYNAMIC(CDataPageDlg, CPropertyPage)

CDataPageDlg::CDataPageDlg()
	: CPropertyPage(CDataPageDlg::IDD)
{
	m_psp.dwFlags &= ~PSP_HASHELP;
}

CDataPageDlg::~CDataPageDlg()
{
}

void CDataPageDlg::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_STATIC_DBNAME, m_ctlDBName);
}


BEGIN_MESSAGE_MAP(CDataPageDlg, CPropertyPage)
END_MESSAGE_MAP()


// CDataPageDlg message handlers
BOOL CDataPageDlg::OnInitDialog()
{
	CPropertyPage::OnInitDialog();

	m_ctlDBName.SetWindowText(m_szInstallPath+_T("PropertySpy.accdb"));

	return TRUE;  // return TRUE  unless you set the focus to a control
}