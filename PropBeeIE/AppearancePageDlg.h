#pragma once


// CAppearancePageDlg dialog

class CAppearancePageDlg : public CPropertyPage
{
	DECLARE_DYNAMIC(CAppearancePageDlg)

public:
	CAppearancePageDlg();
	virtual ~CAppearancePageDlg();

// Dialog Data
	enum { IDD = IDD_DIALOG_APPEARANCE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
};
