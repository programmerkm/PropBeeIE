
// PicStatic.cpp : implementation file
//

#include "stdafx.h"
#include "PicStatic.h"
#include "resource.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPicStatic

CPicStatic::CPicStatic()
{
	m_nBitmapResourceID = 0;
}

CPicStatic::~CPicStatic()
{
}


BEGIN_MESSAGE_MAP(CPicStatic, CStatic)
	//{{AFX_MSG_MAP(CPicStatic)
	ON_WM_PAINT()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Draw the specified bitmap into the static control that is linked
// to this class.

void CPicStatic::OnPaint() 
{
	CPaintDC dc(this); // device context for painting

	ASSERT(m_nBitmapResourceID);	// Caller forgot to set the resource ID
	if(!m_nBitmapResourceID)
		return;

	CRect rect;
	this->GetClientRect(&rect);
	DisplayBitmap(m_nBitmapResourceID, &dc, &rect);
}

/////////////////////////////////////////////////////////////////////////////
// PUBLIC
// Specifies the resource ID of the bitmap to be displayed.

void CPicStatic::SetBitmapResourceID(const UINT nID)
{
	m_nBitmapResourceID = nID;

	if (!m_picImage.IsNull())
		m_picImage.Detach();

	if (!LoadPNGFromResource(m_picImage, nID))
		m_picImage.LoadFromResource(AfxGetInstanceHandle(), MAKEINTRESOURCE(nID));
}

/////////////////////////////////////////////////////////////////////////////
// This function is used to display a resource which can be a PNG resource file
// Called from the OnPaint() routine of a dialog.

void CPicStatic::DisplayBitmap(UINT nResID, CDC *dc,CRect* prctDest)
{
	// Try and first load it as a PNG file, and if it isn't a PNG then load it from resource
	//
	if (prctDest == NULL)
		m_picImage.BitBlt(dc->GetSafeHdc(), 0, 0, SRCCOPY);
	else
		m_picImage.StretchBlt(dc->GetSafeHdc(), 
				   prctDest->top,
				   prctDest->left,
				   prctDest->Width(),
				   prctDest->Height(),
				   SRCCOPY);
	return;
}

/////////////////////////////////////////////////////////////////////////////

HBITMAP CPicStatic::LoadResourceBitmap(HINSTANCE hInstance, LPTSTR lpString,
                           HPALETTE FAR* lphPalette)
{
    HRSRC  hRsrc;
    HGLOBAL hGlobal;
    HBITMAP hBitmapFinal = NULL;
    LPBITMAPINFOHEADER  lpbi;
    HDC hdc;
    int iNumColors;
	
    if (hRsrc = FindResource(hInstance, lpString, RT_BITMAP))
	{
		hGlobal = LoadResource(hInstance, hRsrc);
		lpbi = (LPBITMAPINFOHEADER)LockResource(hGlobal);
		
		hdc = ::GetDC(NULL);
		*lphPalette =  CreateDIBPalette ((LPBITMAPINFO)lpbi, &iNumColors);
		if (*lphPalette)
		{
			SelectPalette(hdc,*lphPalette,FALSE);
			RealizePalette(hdc);
		}
		
		hBitmapFinal = CreateDIBitmap(hdc,
			(LPBITMAPINFOHEADER)lpbi,
			(LONG)CBM_INIT,
			(LPSTR)lpbi + lpbi->biSize + iNumColors *
			sizeof(RGBQUAD),
			
			(LPBITMAPINFO)lpbi,
			DIB_RGB_COLORS );
		
		::ReleaseDC(NULL,hdc);
		UnlockResource(hGlobal);
		FreeResource(hGlobal);
	}
    return (hBitmapFinal);
}

/////////////////////////////////////////////////////////////////////////////

HPALETTE CPicStatic::CreateDIBPalette (LPBITMAPINFO lpbmi, LPINT lpiNumColors)
{
	LPBITMAPINFOHEADER  lpbi;
	LPLOGPALETTE     lpPal;
	HANDLE           hLogPal;
	HPALETTE         hPal = NULL;
	int              i;
	
	lpbi = (LPBITMAPINFOHEADER)lpbmi;

	//ASSERT(lpbi->biBitCount <= 8);	// we cannot handle >256 colors

	if (lpbi->biBitCount <= 8)
		*lpiNumColors = (1 << lpbi->biBitCount);
	else
		*lpiNumColors = 0;  // No palette needed for 24 BPP DIB
	
	if (lpbi->biClrUsed > 0)
		*lpiNumColors = lpbi->biClrUsed;  // Use biClrUsed
		
		if (*lpiNumColors)
		{
			hLogPal = GlobalAlloc (GHND, sizeof (LOGPALETTE) +
				sizeof (PALETTEENTRY) *(*lpiNumColors));
			lpPal = (LPLOGPALETTE) GlobalLock (hLogPal);
			lpPal->palVersion    = 0x300;
			lpPal->palNumEntries = (short)*lpiNumColors;
			
			for (i = 0;  i < *lpiNumColors;  i++)
			{
				lpPal->palPalEntry[i].peRed = lpbmi->bmiColors[i].rgbRed;
				lpPal->palPalEntry[i].peGreen = lpbmi->bmiColors[i].rgbGreen;
				lpPal->palPalEntry[i].peBlue = lpbmi->bmiColors[i].rgbBlue;
				lpPal->palPalEntry[i].peFlags = 0;
			}
			hPal = CreatePalette (lpPal);
			GlobalUnlock (hLogPal);
			GlobalFree   (hLogPal);
		}
		return hPal;
}

BOOL CPicStatic::LoadPNGFromResource(CImage& picImage, int nResID)
{
	// Check that it is a PNG file resource
	//
	HRSRC hRes = ::FindResource(AfxGetInstanceHandle(), MAKEINTRESOURCE(nResID), _T("PNG"));
	if (!hRes)
		return FALSE;

	// Load the resource into memory
	//
	DWORD dwImageSize = ::SizeofResource(AfxGetInstanceHandle(), hRes);
	HGLOBAL hMem      = ::LoadResource(AfxGetInstanceHandle(), hRes);
	if (!hMem)
		return FALSE;

	LPVOID lpData = ::LockResource(hMem);
	if (!lpData)
		return FALSE;

	HGLOBAL hBuffer = ::GlobalAlloc(GMEM_MOVEABLE, dwImageSize);
	if (!hBuffer)
	{
		::UnlockResource(hMem);
		return FALSE;
	}

	void* pBuffer = ::GlobalLock(hBuffer);
	if (!pBuffer)
	{
		::UnlockResource(hMem);
		::GlobalFree(hBuffer);
		return FALSE;
	}

	CopyMemory(pBuffer, lpData, dwImageSize);

	// Create a stream on the memory and load it into the CImage
	//
	IStream *pStream = NULL;
	HRESULT hr = CreateStreamOnHGlobal(hBuffer, FALSE, &pStream);

	if (SUCCEEDED(hr))
		hr = picImage.Load(pStream);

	// Clean up
	//
	if (pStream)
		pStream->Release();

	::GlobalUnlock(hBuffer);
	::GlobalFree(hBuffer);
	::UnlockResource(hMem);

	if (FAILED(hr))
		return FALSE;
	
	return TRUE;
}