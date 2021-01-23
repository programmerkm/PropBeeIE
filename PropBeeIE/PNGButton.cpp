
#include "stdafx.h"
#include "PNGButton.h"

#include "CGdiPlusBitmap.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

COLORREF defaultGrayPalette[256];
BOOL bGrayPaletteSet = FALSE;
/////////////////////////////////////////////////////////////////////////////
// CPNGButton

CPNGButton::CPNGButton()
{
	m_pStdImage = NULL;
	m_pAltImage = NULL;

	m_bHaveBitmaps = FALSE;
	m_bHaveAltImage = FALSE;

	m_pCurBtn = NULL;

	m_bIsDisabled = FALSE;
	m_bIsToggle = FALSE;

	m_bIsHovering = FALSE;
	m_bIsTracking = FALSE;

	m_nCurType = STD_TYPE;

	m_pToolTip = NULL;

}

CPNGButton::~CPNGButton()
{
	if(m_pStdImage) delete m_pStdImage;
	if(m_pAltImage) delete m_pAltImage;

	if(m_pToolTip)	delete m_pToolTip;
}


BEGIN_MESSAGE_MAP(CPNGButton, CButton)
	//{{AFX_MSG_MAP(CPNGButton)
	ON_WM_DRAWITEM()
	ON_WM_ERASEBKGND()
	ON_WM_CTLCOLOR_REFLECT()
	ON_WM_MOUSEMOVE()
	ON_MESSAGE(WM_MOUSELEAVE, OnMouseLeave)
	ON_MESSAGE(WM_MOUSEHOVER, OnMouseHover)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()



//=============================================================================
BOOL CPNGButton::LoadStdImage(UINT id, LPCTSTR pType)
{
	m_pStdImage = new CGdiPlusBitmapResource;
	return m_pStdImage->Load(id, pType);
}


BOOL CPNGButton::LoadAltImage(UINT id, LPCTSTR pType)
{
	m_bHaveAltImage = TRUE;
	m_pAltImage = new CGdiPlusBitmapResource;
	return (m_pAltImage->Load(id, pType));
}

//=============================================================================
//
//	The framework calls this member function when a child control is about to 
//	be drawn.  All the bitmaps are created here on the first call. Every thing
//	is done with a memory DC except the background, which get's it's information 
//	from the parent. The background is needed for transparent portions of PNG 
//	images. An always on top app (such as Task Manager) that is in the way can 
//	cause it to get an incorrect background.  To avoid this, the parent should 
//	call the SetBkGnd function with a memory DC when it creates the background.
//				
//=============================================================================
HBRUSH CPNGButton::CtlColor(CDC* pScreenDC, UINT nCtlColor) 
{
	if(!m_bHaveBitmaps)
	{
		if(!m_pStdImage)
		{
			return NULL; // Load the standard image with LoadStdImage()
		}

		CBitmap bmp, *pOldBitmap;

		CRect rect;
		GetClientRect(rect);

		// do everything with mem dc
		CMemDC pDC(pScreenDC, rect);

		Gdiplus::Graphics graphics(pDC->m_hDC);

		// background
		if (m_dcBk.m_hDC == NULL)
		{

			CRect rect1;
			CClientDC clDC(GetParent());
			GetWindowRect(rect1);
			GetParent()->ScreenToClient(rect1);

			m_dcBk.CreateCompatibleDC(&clDC);
			bmp.CreateCompatibleBitmap(&clDC, rect.Width(), rect.Height());
			pOldBitmap = m_dcBk.SelectObject(&bmp);
			m_dcBk.BitBlt(0, 0, rect.Width(), rect.Height(), &clDC, rect1.left, rect1.top, SRCCOPY);
			bmp.DeleteObject();
		}

		// standard image
		if (m_dcStd.m_hDC == NULL)
		{
			PaintBk(pDC);

			graphics.DrawImage(*m_pStdImage, 0, 0);
		
			m_dcStd.CreateCompatibleDC(pDC);
			bmp.CreateCompatibleBitmap(pDC, rect.Width(), rect.Height());
			pOldBitmap = m_dcStd.SelectObject(&bmp);
			m_dcStd.BitBlt(0, 0, rect.Width(), rect.Height(), pDC, 0, 0, SRCCOPY);
			bmp.DeleteObject();

			// standard image pressed
			if (m_dcStdP.m_hDC == NULL)
			{
				PaintBk(pDC);

				graphics.DrawImage(*m_pStdImage, 1, 1);

				m_dcStdP.CreateCompatibleDC(pDC);
				bmp.CreateCompatibleBitmap(pDC, rect.Width(), rect.Height());
				pOldBitmap = m_dcStdP.SelectObject(&bmp);
				m_dcStdP.BitBlt(0, 0, rect.Width(), rect.Height(), pDC, 0, 0, SRCCOPY);
				bmp.DeleteObject();
			}

			// standard image hot
			if(m_dcStdH.m_hDC == NULL)
			{
				PaintBk(pDC);

				ColorMatrix HotMat = {	1.05f, 0.00f, 0.00f, 0.00f, 0.00f,
										0.00f, 1.05f, 0.00f, 0.00f, 0.00f,
										0.00f, 0.00f, 1.05f, 0.00f, 0.00f,
										0.00f, 0.00f, 0.00f, 1.00f, 0.00f,
										0.05f, 0.05f, 0.05f, 0.00f, 1.00f	};

				ImageAttributes ia;
				ia.SetColorMatrix(&HotMat);

				float width = (float)m_pStdImage->m_pBitmap->GetWidth();
				float height = (float)m_pStdImage->m_pBitmap->GetHeight();

				RectF grect; grect.X=0, grect.Y=0; grect.Width = width; grect.Height = height;

				graphics.DrawImage(*m_pStdImage, grect, 0, 0, width, height, UnitPixel, &ia);

				m_dcStdH.CreateCompatibleDC(pDC);
				bmp.CreateCompatibleBitmap(pDC, rect.Width(), rect.Height());
				pOldBitmap = m_dcStdH.SelectObject(&bmp);
				m_dcStdH.BitBlt(0, 0, rect.Width(), rect.Height(), pDC, 0, 0, SRCCOPY);
				bmp.DeleteObject();
			}

			// grayscale image
			if(m_dcGS.m_hDC == NULL)
			{
				PaintBk(pDC);

				ColorMatrix GrayMat = {	0.30f, 0.30f, 0.30f, 0.00f, 0.00f,
										0.59f, 0.59f, 0.59f, 0.00f, 0.00f,
										0.11f, 0.11f, 0.11f, 0.00f, 0.00f,
										0.00f, 0.00f, 0.00f, 1.00f, 0.00f,
										0.00f, 0.00f, 0.00f, 0.00f, 1.00f	};

				ImageAttributes ia;
				ia.SetColorMatrix(&GrayMat);

				float width = (float)m_pStdImage->m_pBitmap->GetWidth();
				float height = (float)m_pStdImage->m_pBitmap->GetHeight();

				RectF grect; grect.X=0, grect.Y=0; grect.Width = width; grect.Height = height;

				graphics.DrawImage(*m_pStdImage, grect, 0, 0, width, height, UnitPixel, &ia);

				m_dcGS.CreateCompatibleDC(pDC);
				bmp.CreateCompatibleBitmap(pDC, rect.Width(), rect.Height());
				pOldBitmap = m_dcGS.SelectObject(&bmp);
				m_dcGS.BitBlt(0, 0, rect.Width(), rect.Height(), pDC, 0, 0, SRCCOPY);
				bmp.DeleteObject();
			}
		}

		// alternate image
		if( (m_dcAlt.m_hDC == NULL) && m_bHaveAltImage )
		{
			PaintBk(pDC);

			graphics.DrawImage(*m_pAltImage, 0, 0);
		
			m_dcAlt.CreateCompatibleDC(pDC);
			bmp.CreateCompatibleBitmap(pDC, rect.Width(), rect.Height());
			pOldBitmap = m_dcAlt.SelectObject(&bmp);
			m_dcAlt.BitBlt(0, 0, rect.Width(), rect.Height(), pDC, 0, 0, SRCCOPY);
			bmp.DeleteObject();

			// alternate image pressed
			if( (m_dcAltP.m_hDC == NULL) && m_bHaveAltImage )
			{
				PaintBk(pDC);

				graphics.DrawImage(*m_pAltImage, 1, 1);
			
				m_dcAltP.CreateCompatibleDC(pDC);
				bmp.CreateCompatibleBitmap(pDC, rect.Width(), rect.Height());
				pOldBitmap = m_dcAltP.SelectObject(&bmp);
				m_dcAltP.BitBlt(0, 0, rect.Width(), rect.Height(), pDC, 0, 0, SRCCOPY);
				bmp.DeleteObject();
			}

			// alternate image hot
			if(m_dcAltH.m_hDC == NULL)
			{
				PaintBk(pDC);

				ColorMatrix HotMat = {	1.05f, 0.00f, 0.00f, 0.00f, 0.00f,
										0.00f, 1.05f, 0.00f, 0.00f, 0.00f,
										0.00f, 0.00f, 1.05f, 0.00f, 0.00f,
										0.00f, 0.00f, 0.00f, 1.00f, 0.00f,
										0.05f, 0.05f, 0.05f, 0.00f, 1.00f	};

				ImageAttributes ia;
				ia.SetColorMatrix(&HotMat);

				float width = (float)m_pStdImage->m_pBitmap->GetWidth();
				float height = (float)m_pStdImage->m_pBitmap->GetHeight();

				RectF grect; grect.X=0, grect.Y=0; grect.Width = width; grect.Height = height;

				graphics.DrawImage(*m_pAltImage, grect, 0, 0, width, height, UnitPixel, &ia);

				m_dcAltH.CreateCompatibleDC(pDC);
				bmp.CreateCompatibleBitmap(pDC, rect.Width(), rect.Height());
				pOldBitmap = m_dcAltH.SelectObject(&bmp);
				m_dcAltH.BitBlt(0, 0, rect.Width(), rect.Height(), pDC, 0, 0, SRCCOPY);
				bmp.DeleteObject();
			}
		}

		if(m_pCurBtn == NULL)
		{
			m_pCurBtn = &m_dcStd;
		}

		m_bHaveBitmaps = TRUE;
	}

	return NULL;
}

//=============================================================================
// paint the background
//=============================================================================
void CPNGButton::PaintBk(CDC *pDC)
{
	CRect rect;
	GetClientRect(rect);
	pDC->BitBlt(0, 0, rect.Width(), rect.Height(), &m_dcBk, 0, 0, SRCCOPY);
}

//=============================================================================
// paint the bitmap currently pointed to with m_pCurBtn
//=============================================================================
void CPNGButton::PaintBtn(CDC *pDC)
{
	CRect rect;
	GetClientRect(rect);
	pDC->BitBlt(0, 0, rect.Width(), rect.Height(), m_pCurBtn, 0, 0, SRCCOPY);
}

//=============================================================================
// enables the toggle mode
// returns if it doesn't have the alternate image
//=============================================================================
void CPNGButton::EnableToggle(BOOL bEnable)
{
	if(!m_bHaveAltImage) return;

	m_bIsToggle = bEnable; 

	// this actually makes it start in the std state since toggle is called before paint
	if(bEnable)	m_pCurBtn = &m_dcAlt;
	else		m_pCurBtn = &m_dcStd;

}

//=============================================================================
// sets the image type and disabled state then repaints
//=============================================================================
void CPNGButton::SetImage(int type)
{
	m_nCurType = type;

	(type == DIS_TYPE) ? m_bIsDisabled = TRUE : m_bIsDisabled = FALSE;

	Invalidate();
}

//=============================================================================
// set the control to owner draw
//=============================================================================
void CPNGButton::PreSubclassWindow()
{
	// Set control to owner draw
	ModifyStyle(0, BS_OWNERDRAW, SWP_FRAMECHANGED);

	CButton::PreSubclassWindow();
}

//=============================================================================
// disable double click 
//=============================================================================
BOOL CPNGButton::PreTranslateMessage(MSG* pMsg) 
{
	if (pMsg->message == WM_LBUTTONDBLCLK)
		pMsg->message = WM_LBUTTONDOWN;

	if (m_pToolTip != NULL)
	{
		if (::IsWindow(m_pToolTip->m_hWnd))
		{
			m_pToolTip->RelayEvent(pMsg);		
		}
	}

	return CButton::PreTranslateMessage(pMsg);
}


//=============================================================================
// overide the erase function
//=============================================================================
BOOL CPNGButton::OnEraseBkgnd(CDC* pDC) 
{
	return TRUE;
}

//=============================================================================
// Paint the button depending on the state of the mouse
//=============================================================================
void CPNGButton::DrawItem(LPDRAWITEMSTRUCT lpDIS) 
{
	CDC* pDC = CDC::FromHandle(lpDIS->hDC);

	// handle disabled state
	if(m_bIsDisabled)
	{
		m_pCurBtn = &m_dcGS;
		PaintBtn(pDC);
		return;
	}

	BOOL bIsPressed = (lpDIS->itemState & ODS_SELECTED);

	// handle toggle button
	if(m_bIsToggle && bIsPressed)
	{
		(m_nCurType == STD_TYPE) ? m_nCurType = ALT_TYPE : m_nCurType = STD_TYPE;
	}

	if(bIsPressed)
	{
		if(m_nCurType == STD_TYPE)
			m_pCurBtn = &m_dcStdP;
		else
			m_pCurBtn = &m_dcAltP;
	}
	else if(m_bIsHovering)
	{

		if(m_nCurType == STD_TYPE)
			m_pCurBtn = &m_dcStdH;
		else
			m_pCurBtn = &m_dcAltH;
	}
	else
	{
		if(m_nCurType == STD_TYPE)
			m_pCurBtn = &m_dcStd;
		else
			m_pCurBtn = &m_dcAlt;
	}

	// paint the button
	PaintBtn(pDC);
}

//=============================================================================
LRESULT CPNGButton::OnMouseHover(WPARAM wparam, LPARAM lparam) 
//=============================================================================
{
	m_bIsHovering = TRUE;
	Invalidate();
	DeleteToolTip();

	// Create a new Tooltip with new Button Size and Location
	SetToolTipText(m_tooltext);

	if (m_pToolTip != NULL)
	{
		if (::IsWindow(m_pToolTip->m_hWnd))
		{
			//Display ToolTip
			m_pToolTip->Update();
		}
	}

	return 0;
}


//=============================================================================
LRESULT CPNGButton::OnMouseLeave(WPARAM wparam, LPARAM lparam)
//=============================================================================
{
	m_bIsTracking = FALSE;
	m_bIsHovering = FALSE;
	Invalidate();
	return 0;
}

//=============================================================================
void CPNGButton::OnMouseMove(UINT nFlags, CPoint point) 
//=============================================================================
{
	if (!m_bIsTracking)
	{
		TRACKMOUSEEVENT tme;
		tme.cbSize = sizeof(tme);
		tme.hwndTrack = m_hWnd;
		tme.dwFlags = TME_LEAVE|TME_HOVER;
		tme.dwHoverTime = 1;
		m_bIsTracking = _TrackMouseEvent(&tme);
	}
	
	CButton::OnMouseMove(nFlags, point);
}

//=============================================================================
//	
//	Call this member function with a memory DC from the code that paints 
//	the parents background.  Passing the screen DC defeats the purpose of 
//  using this function.
//
//=============================================================================
void CPNGButton::SetBkGnd(CDC* pDC)
{
	CRect rect, rectS;
	CBitmap bmp, *pOldBitmap;

	GetClientRect(rect);
	GetWindowRect(rectS);
	GetParent()->ScreenToClient(rectS);

	m_dcBk.DeleteDC();

	m_dcBk.CreateCompatibleDC(pDC);
	bmp.CreateCompatibleBitmap(pDC, rect.Width(), rect.Height());
	pOldBitmap = m_dcBk.SelectObject(&bmp);
	m_dcBk.BitBlt(0, 0, rect.Width(), rect.Height(), pDC, rectS.left, rectS.top, SRCCOPY);
	bmp.DeleteObject();
}


//=============================================================================
// Set the tooltip with a string resource
//=============================================================================
void CPNGButton::SetToolTipText(UINT nId, BOOL bActivate)
{
	// load string resource
	m_tooltext.LoadString(nId);

	// If string resource is not empty
	if (m_tooltext.IsEmpty() == FALSE)
	{
		SetToolTipText(m_tooltext, bActivate);
	}

}

//=============================================================================
// Set the tooltip with a CString
//=============================================================================
void CPNGButton::SetToolTipText(CString spText, BOOL bActivate)
{
	// We cannot accept NULL pointer
	if (spText.IsEmpty()) return;

	// Initialize ToolTip
	InitToolTip();
	m_tooltext = spText;

	// If there is no tooltip defined then add it
	if (m_pToolTip->GetToolCount() == 0)
	{
		CRect rectBtn; 
		GetClientRect(rectBtn);
		m_pToolTip->AddTool(this, m_tooltext, rectBtn, 1);
	}

	// Set text for tooltip
	m_pToolTip->UpdateTipText(m_tooltext, this, 1);
	m_pToolTip->SetDelayTime(2000);
	m_pToolTip->Activate(bActivate);
}

//=============================================================================
void CPNGButton::InitToolTip()
//=============================================================================
{
	if (m_pToolTip == NULL)
	{
		m_pToolTip = new CToolTipCtrl;
		// Create ToolTip control
		m_pToolTip->Create(this);
		m_pToolTip->Activate(TRUE);
	}
} 

//=============================================================================
void CPNGButton::DeleteToolTip()
//=============================================================================
{
	// Destroy Tooltip incase the size of the button has changed.
	if (m_pToolTip != NULL)
	{
		delete m_pToolTip;
		m_pToolTip = NULL;
	}
}


BOOL CPNGButton::LoadPNGFromResource(CImage *pImage, int nResID)
{
	HRESULT hr = S_OK;
	HMODULE hModule = ::GetModuleHandle(_T(""));

	HRSRC hRes = ::FindResource(hModule,MAKEINTRESOURCE(nResID),_T("PNG"));
	if( !hRes )
		return FALSE;

	DWORD dwImageSize = ::SizeofResource(hModule, hRes);
	HGLOBAL hMem = ::LoadResource(hModule,hRes);

	if( !hMem )
		return FALSE;

	LPVOID lpData = ::LockResource(hMem);
	HGLOBAL hBuffer  = ::GlobalAlloc(GMEM_MOVEABLE, dwImageSize);
	void* pBuffer = ::GlobalLock(hBuffer);

	if( !pBuffer )
		return FALSE;

	CopyMemory(pBuffer, lpData, dwImageSize);

	IStream *pStream=NULL;
	hr = CreateStreamOnHGlobal(hBuffer,TRUE,&pStream);
	if( FAILED(hr) )
		return FALSE;

	hr = pImage->Load(pStream);

	pStream->Release();
	::GlobalUnlock(hBuffer);
	::GlobalFree(hBuffer);
	::UnlockResource(hMem);

	if( FAILED(hr) )
		return FALSE;

	return TRUE;
}

HICON CPNGButton::CreateGrayscaleIcon( HICON hIcon, COLORREF* pPalette )
{
	if (hIcon == NULL)
	{
		return NULL;
	}

	HDC hdc = ::GetDC(NULL);

	HICON      hGrayIcon      = NULL;
	ICONINFO   icInfo         = { 0 };
	ICONINFO   icGrayInfo     = { 0 };
	LPDWORD    lpBits         = NULL;
	LPBYTE     lpBitsPtr      = NULL;
	SIZE sz;
	DWORD c1 = 0;
	BITMAPINFO bmpInfo        = { 0 };
	bmpInfo.bmiHeader.biSize  = sizeof(BITMAPINFOHEADER);

	if (::GetIconInfo(hIcon, &icInfo))
	{
		if (::GetDIBits(hdc, icInfo.hbmColor, 0, 0, NULL, &bmpInfo, DIB_RGB_COLORS) != 0)
		{
			bmpInfo.bmiHeader.biCompression = BI_RGB;

			sz.cx = bmpInfo.bmiHeader.biWidth;
			sz.cy = bmpInfo.bmiHeader.biHeight;
			c1 = sz.cx * sz.cy;

			lpBits = (LPDWORD)::GlobalAlloc(GMEM_FIXED, (c1) * 4);

			if (lpBits && ::GetDIBits(hdc, icInfo.hbmColor, 0, sz.cy, lpBits, &bmpInfo, DIB_RGB_COLORS) != 0)
			{
				lpBitsPtr     = (LPBYTE)lpBits;
				UINT off      = 0;

				for (UINT i = 0; i < c1; i++)
				{
					off = (UINT)( 255 - (( lpBitsPtr[0] + lpBitsPtr[1] + lpBitsPtr[2] ) / 3) );

					if (lpBitsPtr[3] != 0 || off != 255)
					{
						if (off == 0)
						{
							off = 1;
						}

						lpBits[i] = pPalette[off] | ( lpBitsPtr[3] << 24 );
					}

					lpBitsPtr += 4;
				}

				icGrayInfo.hbmColor = ::CreateCompatibleBitmap(hdc, sz.cx, sz.cy);

				if (icGrayInfo.hbmColor != NULL)
				{
					::SetDIBits(hdc, icGrayInfo.hbmColor, 0, sz.cy, lpBits, &bmpInfo, DIB_RGB_COLORS);

					icGrayInfo.hbmMask = icInfo.hbmMask;
					icGrayInfo.fIcon   = TRUE;

					hGrayIcon = ::CreateIconIndirect(&icGrayInfo);

					::DeleteObject(icGrayInfo.hbmColor);
				}

				::GlobalFree(lpBits);
				lpBits = NULL;
			}
		}

		::DeleteObject(icInfo.hbmColor);
		::DeleteObject(icInfo.hbmMask);
	}

	::ReleaseDC(NULL,hdc);

	return hGrayIcon;
}

HICON CPNGButton::CreateGrayscaleIcon( HICON hIcon )
{
	if (hIcon == NULL)
	{
		return NULL;
	}

	if (!bGrayPaletteSet)
	{
		for(int i = 0; i < 256; i++)
		{
			defaultGrayPalette[i] = RGB(255-i, 255-i, 255-i);
		}

		bGrayPaletteSet = TRUE;
	}

	return CreateGrayscaleIcon(hIcon, defaultGrayPalette);
}