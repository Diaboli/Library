VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BookForm 
   Caption         =   "ͼ�����"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   8430
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "�׼�¼"
      Height          =   495
      Left            =   3960
      TabIndex        =   24
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton CmdPre 
      Caption         =   "ǰ��"
      Height          =   495
      Left            =   4080
      TabIndex        =   23
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "����"
      Height          =   420
      Left            =   4080
      TabIndex        =   22
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   "ĩ��¼"
      Height          =   420
      Left            =   4080
      TabIndex        =   21
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "ɾ��"
      Height          =   420
      Left            =   2760
      TabIndex        =   20
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "���"
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   5400
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "Select * From ͼ���"
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
   Begin VB.TextBox Text9 
      DataField       =   "���Ĵ���"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      DataField       =   "����״̬"
      DataSource      =   "Adodc1"
      Height          =   270
      Left            =   2040
      TabIndex        =   16
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      DataField       =   "���"
      DataSource      =   "Adodc1"
      Height          =   270
      Left            =   2040
      TabIndex        =   15
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      DataField       =   "��������"
      DataSource      =   "Adodc1"
      Height          =   270
      Left            =   2040
      TabIndex        =   14
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      DataField       =   "������"
      DataSource      =   "Adodc1"
      Height          =   270
      Left            =   2040
      TabIndex        =   13
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      DataField       =   "����"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "����"
      DataSource      =   "Adodc1"
      Height          =   270
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "ͼ����"
      DataSource      =   "Adodc1"
      Height          =   270
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "�������"
      DataSource      =   "Adodc1"
      Height          =   270
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "���Ĵ���"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "����״̬"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "���"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "��������"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "������"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "����"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "����"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "ͼ����"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�������"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "BookForm"
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

Private Sub cmdOK_Click()
Adodc1.Recordset.Update
CmdAdd.Enabled = True
CmdDel.Enabled = True
CmdOk.Enabled = False
CmdCancel.Enabled = False
End Sub

Private Sub CmdDel_Click()
x = MsgBox("ȷʵҪɾ����ǰ��¼��?", vbYesNo + vbQuestion)
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


