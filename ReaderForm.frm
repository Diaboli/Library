VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ReaderForm 
   Caption         =   "读者管理"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   5310
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "删除"
      Height          =   345
      Left            =   3000
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "确定"
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "添加"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   "末记录"
      Height          =   420
      Left            =   4440
      TabIndex        =   11
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "后移"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton CmdPre 
      Caption         =   "前移"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "首记录"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   360
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=图书管理"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "图书管理"
      OtherAttributes =   ""
      UserName        =   "Devil"
      Password        =   "1137Thornsthrone"
      RecordSource    =   "Select * From 读者表"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      DataField       =   "是否有超期"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      DataField       =   "性别"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      DataField       =   "读者姓名"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataField       =   "借书证号"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "是否有超期"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "性别"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "读者姓名"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "借书证号"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "ReaderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CmdFirst.Enabled = False
CmdPre.Enabled = False
CmdNext.Enabled = True
CmdLast.Enabled = True
CmdAdd.Enabled = True
CmdDel.Enabled = True
CmdOk.Enabled = False
CmdCancel.Enabled = False
End Sub

Private Sub CmdAdd_Click()
Adodc1.Recordset.AddNew
CmdAdd.Enabled = False
CmdDel.Enabled = False
CmdOk.Enabled = True
CmdCancel.Enabled = True
End Sub

Private Sub CmdOk_Click()
Adodc1.Recordset.Update
CmdAdd.Enabled = True
CmdDel.Enabled = True
CmdOk.Enabled = False
CmdCancel.Enabled = False
End Sub

Private Sub CmdDel_Click()
x = MsgBox("确实要删除当前记录吗?", vbYesNo + vbQuestion)
If x = vbYes Then
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.EOF = True Then
        Adodc1.Recordset.MoveLast
    End If
Else
    Adodc1.Refresh
End If
End Sub

Private Sub CmdCancel_Click()
Adodc1.Refresh
CmdAdd.Enabled = True
CmdDel.Enabled = True
CmdOk.Enabled = False
CmdCancel.Enabled = False
End Sub

Private Sub CmdFirst_Click()
Adodc1.Recordset.MoveFirst
CmdFirst.Enabled = False
CmdPre.Enabled = False
CmdNext.Enabled = True
CmdLast.Enabled = True
End Sub

Private Sub CmdPre_Click()
Adodc1.Recordset.MovePrevious
CmdNext.Enabled = True
CmdLast.Enabled = True
If Adodc1.Recordset.BOF = True Then
    Adodc1.Recordset.MoveFirst
    CmdFirst.Enabled = False
    CmdPre.Enabled = False
End If
End Sub

Private Sub CmdNext_Click()
Adodc1.Recordset.MoveNext
CmdFirst.Enabled = True
CmdPre.Enabled = True
If Adodc1.Recordset.EOF = True Then
    Adodc1.Recordset.MoveLast
    CmdLast.Enabled = False
    CmdNext.Enabled = False
End If
End Sub

Private Sub CmdLast_Click()
Adodc1.Recordset.MoveLast
CmdFirst.Enabled = True
CmdPre.Enabled = True
CmdNext.Enabled = False
CmdLast.Enabled = False
End Sub

Private Sub Label3_Click()

End Sub
