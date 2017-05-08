VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BookQueryForm 
   Caption         =   "图书信息查询"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   15225
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   390
      Left            =   1680
      Top             =   2880
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * From 图书表"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "BookQueryForm.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   3201
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
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   300
      Left            =   4920
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   300
      Left            =   3840
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "输入查询信息"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "选择查询条件"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "BookQueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Combo1.AddItem "条形码号"
Combo1.AddItem "图书编号"
Combo1.AddItem "书名"
Combo1.AddItem "作者"
Combo1.AddItem "出版社"
Combo1.AddItem "类别"
Combo1.AddItem "出版日期"
Combo1.AddItem "借阅状态"
Combo1.AddItem "借阅次数"
Combo1.Text = "条形码号"
Text1.Text = ""
End Sub

Private Sub Command1_Click()        ' 单击"查询"按钮代码
Dim str As String
str = "select * from 图书表 where 图书表." & Combo1.Text & " like '" & Text1.Text & "%'"
Adodc1.RecordSource = str
Adodc1.Refresh
End Sub

' "返回"按钮代码
Private Sub Command2_Click()
Unload Me
End Sub

