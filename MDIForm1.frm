VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "图书馆管理系统"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu m_system 
      Caption         =   "系统管理"
      Begin VB.Menu m_rigister 
         Caption         =   "注册新用户"
      End
      Begin VB.Menu m_exit 
         Caption         =   "退出系统"
      End
   End
   Begin VB.Menu m_basic 
      Caption         =   "基本信息"
      Begin VB.Menu m_reader 
         Caption         =   "读者信息管理"
      End
      Begin VB.Menu m_readerscan 
         Caption         =   "读者信息浏览"
      End
      Begin VB.Menu m_readerquery 
         Caption         =   "读者信息查询"
      End
      Begin VB.Menu m_book 
         Caption         =   "图书信息管理"
      End
      Begin VB.Menu m_bookquery 
         Caption         =   "图书信息查询"
      End
   End
   Begin VB.Menu m_flow 
      Caption         =   "图书流通"
      Begin VB.Menu m_lend 
         Caption         =   "借阅管理"
      End
      Begin VB.Menu m_punish 
         Caption         =   "罚款管理"
      End
   End
   Begin VB.Menu m_report 
      Caption         =   "报表"
      Begin VB.Menu m_lendp 
         Caption         =   "图书借阅排行"
      End
      Begin VB.Menu m_lend_report 
         Caption         =   "借阅报表"
      End
      Begin VB.Menu m_punish_report 
         Caption         =   "罚款报表"
      End
   End
   Begin VB.Menu m_about 
      Caption         =   "关于"
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
