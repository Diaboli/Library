VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LoginForm 
   Caption         =   "��¼"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   5160
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   1440
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   794
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
      Connect         =   "DSN=ͼ�����"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ͼ�����"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * From �����"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton CmdLogin 
      Caption         =   "��¼"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "��������"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "�����û���"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
    Unload Me
End Sub
Private Sub CmdLogin_Click()
Dim miCount As Integer
Dim str As String
str = "Select ���� From ����� Where �û��� = '" & Text1.Text & "'"
Adodc1.RecordSource = str
Adodc1.Refresh
If Adodc1.Recordset.EOF = True Then
MsgBox "�û�������, ����������", vbExclamation, "����"
Text1.Text = ""
Text1.SetFocus
Else
    str = "Select * From ����� Where �û��� = '" & Text1.Text & "' and ���� = '" & Text2.Text & "'"
    Adodc1.RecordSource = str
    Adodc1.Refresh
    If Adodc1.Recordset.EOF = True Then
        MsgBox "�������, ����������", vbExclamation, "����"
        Text2.Text = ""
        Text2.SetFocus
    Else
    '    MsgBox Adodc1.Recordset.Fields("����"), vbExclamation, "����"
    '    MsgBox Text2.Text, vbExclamation, "����"
        MDIForm1.Show
        Unload Me
        
    End If
End If
miCount = miCount + 1
If miCount >= 3 Then
    Unload Me
End If
End Sub
