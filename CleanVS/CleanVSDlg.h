
// CleanVSDlg.h : header file
//

#pragma once

enum
{
	VS2010_MAX_PROJECT_MRU = 99,
	VS2010_MAX_FILE_MRU = 99,
	VS2010_MAX_FIND_SEARCH = 99,
};

enum EMCLEAN_TYPE
{
	EMCLEAN_TYPE_VS2010_ALL_CLEAN = 0,
	EMCLEAN_TYPE_VS2010_MRU_RECENT_PROJECT,
	EMCLEAN_TYPE_VS2010_MRU_RECENT_FILE,
	EMCLEAN_TYPE_VS2010_SEARCH_FIND,
	EMCLEAN_TYPE_VS2010_LAST_LOADED_SOLUTION,

	EMCLEAN_TYPE_SIZE,
};



// CCleanVSDlg dialog
class CCleanVSDlg : public CDialogEx
{
// Construction
public:
	CCleanVSDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_CLEANVS_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


public:
	CComboBox	m_ComboSelect;
	CEdit		m_EditInfo;

public:
	void DoCleanUp();
	void AddInfo( const wchar_t * _szFormat, ... );

	void VS2010CleanMRURecentProject();
	void VS2010CleanMRURecentFile();
	void VS2010CleanSearchFind();
	void VS2010CleanLastLoadedSolution();

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonRun();
};
