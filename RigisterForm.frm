VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RigisterForm 
   Caption         =   "ע��"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1440
      Top             =   2640
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
      Left            =   2760
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CmdRegister 
      Caption         =   "ע��"
      Height          =   420
      Left            =   720
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   390
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "ȷ������"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "����(6���ַ�����)"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "�û���(10���ַ�����)"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "RigisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdRegister_Click()
Dim str As String
' ���ȷ�������Ƿ�һ��
If Text2.Text <> Text3.Text Then
    MsgBox "ȷ�����벻һ��, ����������", vbExlamation, "����"
    Text3.Text = ""
    Text3.SetFocus
Else
    ' ����û����Ƿ��ѱ�ע��
    str = "select * from ����� where �û��� = '" & Text1.Text & "'"
    Adodc1.RecordSource = str
    Adodc1.Refresh
    If Adodc1.Recordset.EOF <> True Then
        MsgBox "�û����ѱ�ע��, ����������", vbExlamation, "����"
        Text1.Text = ""
        Text1.SetFocus
    Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields("�û���") = Text1
        Adodc1.Recordset.Fields("����") = Text2
        Adodc1.Recordset.Update
        MsgBox "�û���ע��"
        MDIForm1.Show
        Unload Me
    End If
End If
End Sub
