
// CleanVSDlg.cpp : implementation file
//

#include "stdafx.h"
#include "CleanVS.h"
#include "CleanVSDlg.h"
#include "afxdialogex.h"

#include <string>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CCleanVSDlg dialog




CCleanVSDlg::CCleanVSDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CCleanVSDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CCleanVSDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_COMBO_TYPE, m_ComboSelect );
	DDX_Control(pDX, IDC_EDIT_INFO, m_EditInfo );
}

BEGIN_MESSAGE_MAP(CCleanVSDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_RUN, &CCleanVSDlg::OnBnClickedButtonRun)
END_MESSAGE_MAP()


// CCleanVSDlg message handlers

BOOL CCleanVSDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	std::wstring strType[EMCLEAN_TYPE_SIZE] = 
	{
		_T("VS2010 ALL CleanUp"),
		_T("VS2010 RecentProject CleanUp"),
		_T("VS2010 RecentFile CleanUp"),
		_T("VS2010 Search CleanUp"),
		_T("VS2010 LastLoadedSolution CleanUp"),
	};
	
	m_ComboSelect.ResetContent();
	for ( int i=0; i<EMCLEAN_TYPE_SIZE; ++i )
	{
		int nIndex = m_ComboSelect.AddString ( strType[i].c_str() );
		m_ComboSelect.SetItemData ( nIndex, i );
	}

	m_ComboSelect.SetCurSel(0);

	AddInfo( _T("....") );

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CCleanVSDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CCleanVSDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CCleanVSDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CCleanVSDlg::OnBnClickedButtonRun()
{
	int nRet = MessageBoxW((L"Proceed with cleanup?"), ( L"Question"), MB_YESNO );
	if ( nRet == IDYES )
	{
		DoCleanUp();
	}
}

void CCleanVSDlg::DoCleanUp()
{
	int nSelectedIndex = m_ComboSelect.GetCurSel ();
	if ( nSelectedIndex == LB_ERR )
	{
		AddInfo( _T("Selected == LB_ERR") );
		return;

	}

	EMCLEAN_TYPE emType = static_cast<EMCLEAN_TYPE>( nSelectedIndex );

	switch( emType )
	{
	case EMCLEAN_TYPE_VS2010_ALL_CLEAN:
		{
			VS2010CleanMRURecentProject();
			VS2010CleanMRURecentFile();
			VS2010CleanSearchFind();
			VS2010CleanLastLoadedSolution();
		}break;

	case EMCLEAN_TYPE_VS2010_MRU_RECENT_PROJECT:
		{
			VS2010CleanMRURecentProject();
		}break;

	case EMCLEAN_TYPE_VS2010_MRU_RECENT_FILE:
		{
			VS2010CleanMRURecentFile();
		}break;

	case EMCLEAN_TYPE_VS2010_SEARCH_FIND:
		{
			VS2010CleanSearchFind();
		}break;

	case EMCLEAN_TYPE_VS2010_LAST_LOADED_SOLUTION:
		{
			VS2010CleanLastLoadedSolution();
		}break;

	default:
		{
			AddInfo( _T("Selected function not implemented.") );
			break;
		};
	};


}

void CCleanVSDlg::AddInfo( const wchar_t * _szFormat, ... )
{
	wchar_t szBuffer[2048];
	va_list argList;
	va_start( argList, _szFormat );
	vswprintf( szBuffer, 2048, _szFormat, argList );
	va_end( argList );

	int len = m_EditInfo.GetWindowTextLength();

	if ( m_EditInfo.GetWindowTextLength() >= m_EditInfo.GetLimitText() )
	{
		m_EditInfo.SetSel( 0, m_EditInfo.GetWindowTextLength() );
	}else{
		m_EditInfo.SetSel( len, len );
	}

	m_EditInfo.ReplaceSel( szBuffer );	

	len = m_EditInfo.GetWindowTextLength();
	m_EditInfo.SetSel( len, len );
	m_EditInfo.ReplaceSel( (L"\r\n") );
}


void CCleanVSDlg::VS2010CleanMRURecentProject()
{
	AddInfo( _T("Starting VS2010 Recent Project CleanUp") );

	std::wstring strSubKey = _T("Software\\Microsoft\\VisualStudio\\10.0\\ProjectMRUList");
	AddInfo( _T("Opening Registry: %s"), strSubKey.c_str() );

	HKEY hKey = NULL;
	long lReturn = RegOpenKeyEx( HKEY_CURRENT_USER, strSubKey.c_str(), 0L, KEY_SET_VALUE, &hKey );
	if ( lReturn == ERROR_SUCCESS )
	{
		for ( int i=0; i<VS2010_MAX_PROJECT_MRU; ++i )
		{
			wchar_t szBuffer[128];
			swprintf( szBuffer, 128, _T("File%d"), i );

			lReturn = RegDeleteValue( hKey, szBuffer);
			if ( lReturn == ERROR_SUCCESS )
			{
				AddInfo( _T("Delete: %s"), szBuffer );
			}
		}

		RegCloseKey(hKey);
		AddInfo( _T("Registry Closed"));
	}
	else
	{
		if ( lReturn == ERROR_FILE_NOT_FOUND )
			AddInfo( _T("Registry not found"));
		else
			AddInfo( _T("Registry open failed"));
	}
	
}

void CCleanVSDlg::VS2010CleanMRURecentFile()
{
	AddInfo( _T("Starting VS2010 Recent File CleanUp") );

	std::wstring strSubKey = _T("Software\\Microsoft\\VisualStudio\\10.0\\FileMRUList");
	AddInfo( _T("Opening Registry: %s"), strSubKey.c_str() );

	HKEY hKey = NULL;
	long lReturn = RegOpenKeyEx( HKEY_CURRENT_USER, strSubKey.c_str(), 0L, KEY_SET_VALUE, &hKey );
	if ( lReturn == ERROR_SUCCESS )
	{
		for ( int i=0; i<VS2010_MAX_FILE_MRU; ++i )
		{
			wchar_t szBuffer[128];
			swprintf( szBuffer, 128, _T("File%d"), i );

			lReturn = RegDeleteValue( hKey, szBuffer);
			if ( lReturn == ERROR_SUCCESS )
			{
				AddInfo( _T("Delete: %s"), szBuffer );
			}
		}

		RegCloseKey(hKey);
		AddInfo( _T("Registry Closed"));
	}
	else
	{
		if ( lReturn == ERROR_FILE_NOT_FOUND )
			AddInfo( _T("Registry not found"));
		else
			AddInfo( _T("Registry open failed"));
	}

}

void CCleanVSDlg::VS2010CleanSearchFind()
{
	AddInfo( _T("Starting VS2010 Search Find CleanUp") );

	std::wstring strSubKey = _T("Software\\Microsoft\\VisualStudio\\10.0\\Find");
	AddInfo( _T("Opening Registry: %s"), strSubKey.c_str() );

	HKEY hKey = NULL;
	long lReturn = RegOpenKeyEx( HKEY_CURRENT_USER, strSubKey.c_str(), 0L, KEY_SET_VALUE, &hKey );
	if ( lReturn == ERROR_SUCCESS )
	{
		for ( int i=0; i<VS2010_MAX_FILE_MRU; ++i )
		{
			wchar_t szBuffer[128];
			swprintf( szBuffer, 128, _T("Find %d"), i );

			lReturn = RegDeleteValue( hKey, szBuffer);
			if ( lReturn == ERROR_SUCCESS )
			{
				AddInfo( _T("Delete: %s"), szBuffer );
			}
		}

		lReturn = RegDeleteValue( hKey, _T("Find"));
		if ( lReturn == ERROR_SUCCESS )
		{
			AddInfo( _T("Delete: %s"), _T("Find"));
		}

		for ( int i=0; i<VS2010_MAX_FILE_MRU; ++i )
		{
			wchar_t szBuffer[128];
			swprintf( szBuffer, 128, _T("Replace %d"), i );

			lReturn = RegDeleteValue( hKey, szBuffer);
			if ( lReturn == ERROR_SUCCESS )
			{
				AddInfo( _T("Delete: %s"), szBuffer );
			}
		}

		lReturn = RegDeleteValue( hKey, _T("Replace"));
		if ( lReturn == ERROR_SUCCESS )
		{
			AddInfo( _T("Delete: %s"), _T("Replace"));
		}

		RegCloseKey(hKey);
		AddInfo( _T("Registry Closed"));
	}
	else
	{
		if ( lReturn == ERROR_FILE_NOT_FOUND )
			AddInfo( _T("Registry not found"));
		else
			AddInfo( _T("Registry open failed"));
	}

}

void CCleanVSDlg::VS2010CleanLastLoadedSolution()
{
	AddInfo( _T("Starting VS2010 LastLoadedSolution CleanUp") );

	std::wstring strSubKey = _T("Software\\Microsoft\\VisualStudio\\10.0");
	AddInfo( _T("Opening Registry: %s"), strSubKey.c_str() );

	HKEY hKey = NULL;
	long lReturn = RegOpenKeyEx( HKEY_CURRENT_USER, strSubKey.c_str(), 0L, KEY_SET_VALUE, &hKey );
	if ( lReturn == ERROR_SUCCESS )
	{
		lReturn = RegDeleteValue( hKey, _T("LastLoadedSolution"));
		if ( lReturn == ERROR_SUCCESS )
		{
			AddInfo( _T("Delete: %s"), _T("LastLoadedSolution"));
		}

		RegCloseKey(hKey);
		AddInfo( _T("Registry Closed"));
	}
	else
	{
		if ( lReturn == ERROR_FILE_NOT_FOUND )
			AddInfo( _T("Registry not found"));
		else
			AddInfo( _T("Registry open failed"));
	}

}