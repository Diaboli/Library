VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LendForm 
   Caption         =   "���Ĺ���"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5760
   StartUpPosition =   3  '����ȱʡ
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
      Connect         =   "DSN=ͼ�����"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ͼ�����"
      OtherAttributes =   ""
      UserName        =   "Devil"
      Password        =   "1137Thornsthrone"
      RecordSource    =   "Select * From ͼ���"
      Caption         =   "Adodc3"
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
      Connect         =   "DSN=ͼ�����"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ͼ�����"
      OtherAttributes =   ""
      UserName        =   "Devil"
      Password        =   "1137Thornsthrone"
      RecordSource    =   "Select * From ���߱�"
      Caption         =   "Adodc2"
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
      Connect         =   "DSN=ͼ�����"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ͼ�����"
      OtherAttributes =   ""
      UserName        =   "Devil"
      Password        =   "1137Thornsthrone"
      RecordSource    =   "Select * From ���ı�"
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
   Begin VB.CommandButton CmdExit 
      Caption         =   "����"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CmdReturn 
      Caption         =   "����"
      Height          =   540
      Left            =   2400
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CmdLend 
      Caption         =   "����"
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
      Caption         =   "�������"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "����֤��"
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
st2 = "select �������, �Ƿ��г��� from ���߱� where ����֤�� = '" & Trim(Text1) & "'"
Adodc2.RecordSource = st2
Adodc2.Refresh
' �н���֤��
If Adodc2.Recordset.BOF = False Then
    ' �����ǽ�ʦ
    If Adodc2.Recordset.Fields("�������") = "t" And Adodc2.Recordset.Fields("�Ƿ��г���") = "n" Then
    ' �Ƿ񳬹�����10��
    st1 = "select * from ���ı� where ����֤�� = '" & Trim(Text1) & "'"
    Adodc1.RecordSource = st1
    Adodc1.Refresh
        If Adodc1.Recordset.RecordCount >= 10 Then
            MsgBox "��ʦ���������ﵽ���, �����ٽ�!"
        Else    ' д����ļ�¼
            Adodc1.Recordset.AddNew
            Adodc1.Recordset.Fields("����֤��") = Text1
            Adodc1.Recordset.Fields("�������") = Text2
            Adodc1.Recordset.Fields("�������") = str(Date) ' ϵͳ������
            Adodc1.Recordset.Fields("�黹����") = Null
            Adodc1.Recordset.Update
            ' ��ͼ�����д��"����״̬"��"���Ĵ���"��Ϣ
            st3 = "select * from ͼ��� where ������� = '" & Trim(Text2) & "'"
            Adodc3.RecordSource = st3
            Adodc3.Refresh
            Adodc3.Recordset.Fields("����״̬") = "���"
            Adodc3.Recordset.Fields("���Ĵ���") = Adodc3.Recordset.Fields("���Ĵ���") + 1
            Adodc3.Recordset.Update
            MsgBox "�����鼮�Ѽ�¼!"
            Text1.Text = ""
            Text2.Text = ""
        End If
    ' ������ѧ��
    ElseIf Adodc2.Recordset.Fields("�������") = "s" And Adodc2.Recordset.Fields("�Ƿ��г���") = "n" Then
        ' �Ƿ񳬹�����3��
        st1 = "select * from ���ı� where ����֤�� = '" & Trim(Text1) & "'"
        Adodc1.RecordSource = st1
        Adodc1.Refresh
        If Adodc1.Recordset.RecordCount >= 3 Then
            MsgBox "ѧ�����������ﵽ���, �����ٽ�!"
        Else
            Adodc1.Recordset.AddNew
            Adodc1.Recordset.Fields("����֤��") = Text1
            Adodc1.Recordset.Fields("�������") = Text2
            Adodc1.Recordset.Fields("�������") = str(Date)
            Adodc1.Recordset.Update
            ' ��ͼ�����д��"����״̬"��"���Ĵ���"��Ϣ
            st3 = "select * from ͼ��� where ������� = '" & Trim(Text2) & "'"
            Adodc3.RecordSource = st3
            Adodc3.Refresh
            Adodc3.Recordset.Fields("����״̬") = "���"
            Adodc3.Recordset.Fields("���Ĵ���") = Adodc3.Recordset.Fields("���Ĵ���") + 1
            Adodc3.Recordset.Update
            MsgBox "�����鼮�Ѽ�¼!"
            Text1.Text = ""
            Text2.Text = ""
        End If
    ElseIf Adodc2.Recordset.Fields("�Ƿ��г���") = "y" Then
        MsgBox "�ö����г�����, �����ٽ���"
    End If
End If

End Sub

Private Sub CmdReturn_Click()      ' ����"����"��ť����
Dim sst As String
Dim sstt As String
Dim sst3 As String
sst = "select * from ���ı� where ������� = '" & Text2 & "'"
Adodc1.RecordSource = sst
Adodc1.Refresh
sstt = "select * from ���߱� where ����֤�� = '" & Adodc1.Recordset.Fields("����֤��") & "'"
Adodc2.RecordSource = sstt
Adodc2.Refresh
' �Ȼ���, �ٷ���
Adodc1.Recordset.Fields("�黹����") = Date
Adodc1.Recordset.Update
MsgBox "����ɹ�"
' ��ͼ���д������"�ڿ�"��Ϣ
sst3 = "select * from ͼ��� where ������� = '" & Text2 & "'"
Adodc3.RecordSource = sst3
Adodc3.Refresh
Adodc3.Recordset.Fields("����״̬") = "�ڿ�"
Adodc3.Recordset.Update
If Adodc2.Recordset.Fields("�������") = "t" Then
    If Date - Adodc1.Recordset.Fields("�������") > 90 Then
        MsgBox "�г��ڷ���"
        PunishForm.Show
    End If
ElseIf Adodc2.Recordset.Fields("�������") = "s" Then
    If Date - Adodc1.Recordset.Fields("�������") > 30 Then
        MsgBox "�г��ڷ���"
        PunishForm.Show
    End If
End If
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
