VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "ͼ��ݹ���ϵͳ"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu m_system 
      Caption         =   "ϵͳ����"
      Begin VB.Menu m_rigister 
         Caption         =   "ע�����û�"
      End
      Begin VB.Menu m_exit 
         Caption         =   "�˳�ϵͳ"
      End
   End
   Begin VB.Menu m_basic 
      Caption         =   "������Ϣ"
      Begin VB.Menu m_reader 
         Caption         =   "������Ϣ����"
      End
      Begin VB.Menu m_readerscan 
         Caption         =   "������Ϣ���"
      End
      Begin VB.Menu m_readerquery 
         Caption         =   "������Ϣ��ѯ"
      End
      Begin VB.Menu m_book 
         Caption         =   "ͼ����Ϣ����"
      End
      Begin VB.Menu m_bookquery 
         Caption         =   "ͼ����Ϣ��ѯ"
      End
   End
   Begin VB.Menu m_flow 
      Caption         =   "ͼ����ͨ"
      Begin VB.Menu m_lend 
         Caption         =   "���Ĺ���"
      End
      Begin VB.Menu m_punish 
         Caption         =   "�������"
      End
   End
   Begin VB.Menu m_report 
      Caption         =   "����"
      Begin VB.Menu m_lendp 
         Caption         =   "ͼ���������"
      End
      Begin VB.Menu m_lend_report 
         Caption         =   "���ı���"
      End
      Begin VB.Menu m_punish_report 
         Caption         =   "�����"
      End
   End
   Begin VB.Menu m_about 
      Caption         =   "����"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub m_about_Click()
AboutForm.Show
End Sub

Private Sub m_book_Click()
BookForm.Show
End Sub

Private Sub m_bookquery_Click()
BookQueryForm.Show
End Sub

Private Sub m_exit_Click()
End
End Sub

Private Sub m_lend_report_Click()
LendReport.Show
End Sub

Private Sub m_lendp_Click()
LendpReport.Show
End Sub

Private Sub m_punish_report_Click()
PunishReport.Show
End Sub

Private Sub m_reader_Click()
ReaderForm.Show
End Sub

Private Sub m_readerquery_Click()
ReaderQueryForm.Show
End Sub

Private Sub m_lend_Click()
LendForm.Show
End Sub

Private Sub m_punish_Click()
PunishForm.Show
End Sub

Private Sub m_readerscan_Click()
ReaderScanForm.Show
End Sub

Private Sub m_rigister_Click()
RigisterForm.Show
End Sub
