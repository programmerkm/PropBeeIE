
#pragma once

// PicStatic.h : header file
//

#include <atlimage.h>

/////////////////////////////////////////////////////////////////////////////
// CPicStatic window

class CPicStatic : public CStatic
{
// Construction
public:
	CPicStatic();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPicStatic)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CPicStatic();
	void SetBitmapResourceID(const UINT nID);

	// Generated message map functions
protected:
	HBITMAP LoadResourceBitmap(HINSTANCE hInstance, LPTSTR lpString, HPALETTE FAR* lphPalette);
	virtual void DisplayBitmap(UINT nResID, CDC *dc,CRect* prctDest = NULL);
	HPALETTE CreateDIBPalette (LPBITMAPINFO lpbmi, LPINT lpiNumColors);
	BOOL LoadPNGFromResource(CImage& picImage, int nResID);

	UINT m_nBitmapResourceID;

	//{{AFX_MSG(CPicStatic)
	afx_msg void OnPaint();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

private:
	CImage m_picImage;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.
