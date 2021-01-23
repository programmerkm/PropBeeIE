// CModelessDemoDlg dialog
#include "stdafx.h"
#include "PropBeeDB.h"

class CImportPBDBDlg : public CDialog
{
	// Construction
public:
	CImportPBDBDlg(CWnd* pParent = NULL);	// standard constructor
	CString GetSQLServerName(){return m_szSQLServerName; }
	void SetInstallPath(CString szPath){m_szInstallPath=szPath;}

	// Dialog Data
	enum { IDD = IDD_IMPORTDB_DIALOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	// Generated message map functions
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()	

public:
	afx_msg void OnBnClickedOk();
	
private:
	void RecursiveFileFind(CString szFolder, CString szFileToFind, CString& szPathFound);

private:
	CString m_szSourcePDBFile;
	CString m_szSQLServerName;
	CButton m_btnImport;
	CString m_szInstallPath;
	CProgressCtrl m_ctlProgress;
	CStatic m_ctlStaticText;
};