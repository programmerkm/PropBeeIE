// GeneralPageDlg.cpp : implementation file
//

#include "stdafx.h"
#include "GeneralPageDlg.h"


// CGeneralPageDlg dialog

IMPLEMENT_DYNAMIC(CGeneralPageDlg, CPropertyPage)

CGeneralPageDlg::CGeneralPageDlg()
	: CPropertyPage(CGeneralPageDlg::IDD)
{
	m_psp.dwFlags &= ~PSP_HASHELP;
}

CGeneralPageDlg::~CGeneralPageDlg()
{
}

void CGeneralPageDlg::DoDataExchange(CDataExchange* pDX)
{
	CPropertyPage::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CGeneralPageDlg, CPropertyPage)
END_MESSAGE_MAP()


// CGeneralPageDlg message handlers
