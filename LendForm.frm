VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LendForm 
   Caption         =   "借阅管理"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5760
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   3960
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      RecordSource    =   "Select * From 图书表"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2040
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "Select * From 借阅表"
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
   Begin VB.CommandButton CmdExit 
      Caption         =   "返回"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CmdReturn 
      Caption         =   "还书"
      Height          =   540
      Left            =   2400
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CmdLend 
      Caption         =   "借书"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "条形码号"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "借书证号"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "LendForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdLend_Click()
Dim st1 As String
Dim st2 As String
Dim st3 As String
Dim i As Integer
st2 = "select 读者类别, 是否有超期 from 读者表 where 借书证号 = '" & Trim(Text1) & "'"
Adodc2.RecordSource = st2
Adodc2.Refresh
' 有借书证号
If Adodc2.Recordset.BOF = False Then
    ' 读者是教师
    If Adodc2.Recordset.Fields("读者类别") = "t" And Adodc2.Recordset.Fields("是否有超期") = "n" Then
    ' 是否超过最多的10本
    st1 = "select * from 借阅表 where 借书证号 = '" & Trim(Text1) & "'"
    Adodc1.RecordSource = st1
    Adodc1.Refresh
        If Adodc1.Recordset.RecordCount >= 10 Then
            MsgBox "教师借书数量达到最多, 不能再借!"
        Else    ' 写入借阅记录
            Adodc1.Recordset.AddNew
            Adodc1.Recordset.Fields("借书证号") = Text1
            Adodc1.Recordset.Fields("条形码号") = Text2
            Adodc1.Recordset.Fields("借出日期") = str(Date) ' 系统的日期
            Adodc1.Recordset.Fields("归还日期") = Null
            Adodc1.Recordset.Update
            ' 在图书表中写入"借阅状态"和"借阅次数"信息
            st3 = "select * from 图书表 where 条形码号 = '" & Trim(Text2) & "'"
            Adodc3.RecordSource = st3
            Adodc3.Refresh
            Adodc3.Recordset.Fields("借阅状态") = "借出"
            Adodc3.Recordset.Fields("借阅次数") = Adodc3.Recordset.Fields("借阅次数") + 1
            Adodc3.Recordset.Update
            MsgBox "借阅书籍已记录!"
            Text1.Text = ""
            Text2.Text = ""
        End If
    ' 读者是学生
    ElseIf Adodc2.Recordset.Fields("读者类别") = "s" And Adodc2.Recordset.Fields("是否有超期") = "n" Then
        ' 是否超过最多的3本
        st1 = "select * from 借阅表 where 借书证号 = '" & Trim(Text1) & "'"
        Adodc1.RecordSource = st1
        Adodc1.Refresh
        If Adodc1.Recordset.RecordCount >= 3 Then
            MsgBox "学生借书数量达到最多, 不能再借!"
        Else
            Adodc1.Recordset.AddNew
            Adodc1.Recordset.Fields("借书证号") = Text1
            Adodc1.Recordset.Fields("条形码号") = Text2
            Adodc1.Recordset.Fields("借出日期") = str(Date)
            Adodc1.Recordset.Update
            ' 在图书表中写入"借阅状态"和"借阅次数"信息
            st3 = "select * from 图书表 where 条形码号 = '" & Trim(Text2) & "'"
            Adodc3.RecordSource = st3
            Adodc3.Refresh
            Adodc3.Recordset.Fields("借阅状态") = "借出"
            Adodc3.Recordset.Fields("借阅次数") = Adodc3.Recordset.Fields("借阅次数") + 1
            Adodc3.Recordset.Update
            MsgBox "借阅书籍已记录!"
            Text1.Text = ""
            Text2.Text = ""
        End If
    ElseIf Adodc2.Recordset.Fields("是否有超期") = "y" Then
        MsgBox "该读者有超期书, 不能再借书"
    End If
End If

End Sub

Private Sub CmdReturn_Click()      ' 单击"还书"按钮代码
Dim sst As String
Dim sstt As String
Dim sst3 As String
sst = "select * from 借阅表 where 条形码号 = '" & Text2 & "'"
Adodc1.RecordSource = sst
Adodc1.Refresh
sstt = "select * from 读者表 where 借书证号 = '" & Adodc1.Recordset.Fields("借书证号") & "'"
Adodc2.RecordSource = sstt
Adodc2.Refresh
' 先还书, 再罚款
Adodc1.Recordset.Fields("归还日期") = Date
Adodc1.Recordset.Update
MsgBox "还书成功"
' 在图书表写入该书的"在库"信息
sst3 = "select * from 图书表 where 条形码号 = '" & Text2 & "'"
Adodc3.RecordSource = sst3
Adodc3.Refresh
Adodc3.Recordset.Fields("借阅状态") = "在库"
Adodc3.Recordset.Update
If Adodc2.Recordset.Fields("读者类别") = "t" Then
    If Date - Adodc1.Recordset.Fields("借出日期") > 90 Then
        MsgBox "有超期罚款"
        PunishForm.Show
    End If
ElseIf Adodc2.Recordset.Fields("读者类别") = "s" Then
    If Date - Adodc1.Recordset.Fields("借出日期") > 30 Then
        MsgBox "有超期罚款"
        PunishForm.Show
    End If
End If
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
