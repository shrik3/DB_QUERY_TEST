
// QUERY_TESTDlg.h : ͷ�ļ�
//

#pragma once


// CQUERY_TESTDlg �Ի���
class CQUERY_TESTDlg : public CDialogEx
{
// ����
public:
	CQUERY_TESTDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_QUERY_TEST_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedAbort2();
	_ConnectionPtr m_pConnection;
	_RecordsetPtr  m_pRecordset;//���ݼ�����
	_variant_t vUsername, vFeature;	//���ȡ���������ݵı���


	afx_msg void OnBnClickedConnect();
	afx_msg void OnBnClickedSearch();
	afx_msg void OnBnClickedInsert();
	afx_msg void OnBnClickedDelete();
	afx_msg void OnBnClickedUpdate();
	afx_msg void OnEnChangeEdit3();
};
