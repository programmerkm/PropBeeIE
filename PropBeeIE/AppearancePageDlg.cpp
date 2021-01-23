// AppearancePageDlg.cpp : implementation file
//

#include "stdafx.h"
#include "AppearancePageDlg.h"


// CAppearancePageDlg dialog

IMPLEMENT_DYNAMIC(CAppearancePageDlg, CPropertyPage)

CAppearancePageDlg::CAppearancePageDlg()
	: CPropertyPage(CAppearancePageDlg::IDD)
{
	m_psp.dwFlags &= ~PSP_HASHELP;
}

CAppearancePageDlg::~CAppearancePageDlg()
{
}

void CAppearancePageDlg::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CAppearancePageDlg, CPropertyPage)
END_MESSAGE_MAP()


// CAppearancePageDlg message handlers
