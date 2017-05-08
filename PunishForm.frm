VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PunishForm 
   Caption         =   "罚款管理"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   6435
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4320
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   $"PunishForm.frx":0000
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
      Left            =   2280
      Top             =   4560
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
      Left            =   240
      Top             =   4560
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
      RecordSource    =   "Select * From 罚款表"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton CmdPunish 
      Caption         =   "处罚"
      Height          =   465
      Left            =   2280
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "超期查询"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "罚款总额"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "借书证号"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "PunishForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdQuery_Click()       ' 单击"超期查询"按钮代码
Dim str3 As String
str3 = "select 读者表.借书证号, 条形码号, 读者姓名, 借出日期, (month(getdate()) - month(借出日期) - 3) * 30 + day(getdate() - day(借出日期)) As 超期天数"
str3 = str3 & " from 读者表 inner join 借阅表 on 读者表.借书证号 = 借阅表.借书证号"
str3 = str3 & " where 借阅表.借书证号 = '" & Text1 & "' and 读者表.读者类别 = 't' And month(归还日期) - month(借出日期) > 3"
str3 = str3 & " or 读者类别 = 's' and month(归还日期) - month(借出日期) > 1"
Adodc3.RecordSource = str3
Adodc3.Refresh
Text2.Text = ""
End Sub

Private Sub CmdPunish_Click()      ' 单击"处罚"按钮代码
Dim pcount As Integer              ' 超期记录的条数
Dim i As Integer
Dim sum As Single
sum = 0
If Adodc3.Recordset.BOF = False Then     ' 有超期的记录
pcount = Adodc3.Recordset.RecordCount
End If
For i = 1 To pcount
sum = sum + Adodc3.Recordset.Fields("超期天数") * 0.1
' 向罚款表中添加记录
Adodc1.Recordset.Fields("借书证号") = Adodc3.Recordset.Fields("借书证号")
Adodc1.Recordset.Fields("条形码号") = Adodc3.Recordset.Fields("条形码号")
Adodc1.Recordset.Fields("处罚日期") = Date
Adodc1.Recordset.Fields("超期天数") = Adodc3.Recordset.Fields("超期天数")
Adodc1.Recordset.Fields("罚款金额") = Adodc3.Recordset.Fields("超期天数") * 0.1
Adodc1.Recordset.Update
Adodc3.Recordset.MoveNext
Next i
Text2 = sum & "元"
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
