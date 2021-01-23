#pragma once


// CDataPageDlg dialog

class CDataPageDlg : public CPropertyPage
{
	DECLARE_DYNAMIC(CDataPageDlg)

public:
	CDataPageDlg();
	virtual ~CDataPageDlg();
	void SetInstallPath(CString szPath){m_szInstallPath = szPath;}

// Dialog Data
	enum { IDD = IDD_DIALOG_DATA };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();

	DECLARE_MESSAGE_MAP()
public:
	CStatic m_ctlDBName;
	CString m_szInstallPath;
};
