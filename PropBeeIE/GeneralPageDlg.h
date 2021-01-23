#pragma once


// CGeneralPageDlg dialog

class CGeneralPageDlg : public CPropertyPage
{
	DECLARE_DYNAMIC(CGeneralPageDlg)

public:
	CGeneralPageDlg();
	virtual ~CGeneralPageDlg();

// Dialog Data
	enum { IDD = IDD_DIALOG_GENERAL };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
};
