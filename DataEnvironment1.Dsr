VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DataEnvironment1 
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13350
   _ExtentX        =   23548
   _ExtentY        =   12330
   FolderFlags     =   1
   TypeLibGuid     =   "{B825BB30-6611-4681-95C6-C408BB8D012D}"
   TypeInfoGuid    =   "{240445D9-120E-4BA9-8E70-70FBE30C995A}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "图书管理"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=图书管理"
      Expanded        =   -1  'True
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   3
   BeginProperty Recordset1 
      CommandName     =   "借阅管理"
      CommDispId      =   1003
      RsDispId        =   1011
      CommandText     =   $"DataEnvironment1.dsx":0000
      ActiveConnectionName=   "图书管理"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "借书证号"
         Caption         =   "借书证号"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "读者姓名"
         Caption         =   "读者姓名"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "书名"
         Caption         =   "书名"
      EndProperty
      BeginProperty Field4 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "借出日期"
         Caption         =   "借出日期"
      EndProperty
      BeginProperty Field5 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "归还日期"
         Caption         =   "归还日期"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "罚款管理"
      CommDispId      =   1005
      RsDispId        =   1014
      CommandText     =   $"DataEnvironment1.dsx":00D4
      ActiveConnectionName=   "图书管理"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   6
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "借书证号"
         Caption         =   "借书证号"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "读者姓名"
         Caption         =   "读者姓名"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "书名"
         Caption         =   "书名"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "超期天数"
         Caption         =   "超期天数"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "罚款金额"
         Caption         =   "罚款金额"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "处罚日期"
         Caption         =   "处罚日期"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "图书借阅排行"
      CommDispId      =   1007
      RsDispId        =   1017
      CommandText     =   "SELECT 图书编号, 书名, SUM(借阅次数) AS 次数 FROM 图书表 GROUP BY 图书编号, 书名 ORDER BY 次数 DESC"
      ActiveConnectionName=   "图书管理"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "图书编号"
         Caption         =   "图书编号"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "书名"
         Caption         =   "书名"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "次数"
         Caption         =   "次数"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DataEnvironment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
