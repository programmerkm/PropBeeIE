#ifndef HRX_H_INCLUDED
#define HRX_H_INCLUDED

#pragma once

#include <comdef.h>

#pragma warning( disable : 4290 ) // Ignore 'C++ Exception Specification ignored'

class HRX
{
private:
	HRESULT m_hr;

public:
	// Defualt constructor sets HRESULT to success
	HRX() { m_hr = S_OK; };

	// Initialise the HRESULT, but use the assignment operator to ensure
	// that a failed result causes an exception to be thrown
	//
	// This correctly throws exceptions for errors in this situation:
	//
	//    HRX hrx = FunctionReturningHResult();
	//
	// And throws exceptions when constructing a "failed" object e.g.
	//
	//    HRX hrx(E_FAIL);
	//
	HRX(HRESULT hr) throw(_com_error) { *this = hr; }

	// Dito for copy constructor
	HRX(const HRX &hrx) throw(_com_error) { *this = hrx; }

	// Set the HRESULT and immediately throw a _com_error on failure
	//
	HRESULT operator=(HRESULT hr) 
	{ 
			m_hr = hr;
			if (FAILED(hr)) 
			{
				//_ASSERTE(SUCCEEDED(hr));

				throw _com_error(m_hr);
			}
			return m_hr;
	}

	// Override the default assignment operator
	HRESULT operator=(const HRX &hrx)
	{
			return (*this = (HRESULT)hrx);
	}

	// Enables the use of HRX as an HRESULT
	//
	operator HRESULT() const 
	{
		return m_hr;
	}

	// Enables the use of HRX as a pointer to an HRESULT
	//
	HRESULT* operator&()
	{
		return &m_hr;
	}

	// Set the HRESULT but do not throw on failure
	//
	HRESULT Set(HRESULT hr) { return m_hr = hr; } 

	void RaiseIfError()
	{
		if (FAILED(m_hr)) throw _com_error(m_hr);
	}

};

#endif