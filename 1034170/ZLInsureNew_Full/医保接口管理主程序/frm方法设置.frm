VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm方法设置 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置各个功能模块所使用到的方法"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frm方法设置.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList img权限 
      Left            =   930
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm方法设置.frx":628A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5550
      TabIndex        =   3
      Top             =   5100
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6780
      TabIndex        =   4
      Top             =   5100
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img方法 
      Left            =   3240
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm方法设置.frx":C524
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img菜单 
      Left            =   870
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm方法设置.frx":D7A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw方法 
      Height          =   4365
      Left            =   3150
      TabIndex        =   2
      Top             =   630
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   7699
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img方法"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvw模块 
      Height          =   2025
      Left            =   30
      TabIndex        =   0
      Top             =   630
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   3572
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img菜单"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "模块"
         Object.Width           =   4233
      EndProperty
   End
   Begin MSComctlLib.ListView lvw权限 
      Height          =   2295
      Left            =   30
      TabIndex        =   1
      Top             =   2700
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4048
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img权限"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请仔细设置各模块所需使用到的方法："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   690
      TabIndex        =   5
      Top             =   390
      Width           =   6900
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frm方法设置.frx":E5F8
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frm方法设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSQL As String
Private mblnFirst As Boolean            '启动
Private mblnReturn As Boolean
Private mrs模块 As New ADODB.Recordset
Private mrs方法 As New ADODB.Recordset

Private mlngModul As Long           '上次选择的模块
Private mstrPrivs As String         '上次选择的权限串

Public Function ShowEditor(rs模块 As ADODB.Recordset, ByVal rs方法 As ADODB.Recordset) As Boolean
    mblnReturn = False
    Set mrs模块 = rs模块
    Set mrs方法 = rs方法
    
    Me.Show 1
    
    If mblnReturn Then Set rs模块 = mrs模块
    ShowEditor = mblnReturn
End Function

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Call SavePrivs
    
    mblnReturn = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mblnFirst = True
    
    '装入医保接口所涉及到的模块（门诊挂号1111、门诊收费1121、病人入院管理1131、病人入出管理1132、住院记帐(医嘱)1133、住院结算1137、病人费用查询1139）
    mstrSQL = "Select 序号,标题 From zlPrograms " & _
        " Where 系统=100 And 序号 IN (1111,1121,1131,1132,1133,1137,1139,1203,1205,1206) Order By 序号"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "获取医保接口所涉及到的模块")
    With rsTemp
        Do While Not .EOF
            lvw模块.ListItems.Add , "K_" & !序号, !标题, , 1
            .MoveNext
        Loop
    End With
    
    '加入可供选择的方法（也就是接口所支持的方法列表）
    With mrs方法
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lvw方法.ListItems.Add , "K_" & !序号, !方法, , 1
            .MoveNext
        Loop
    End With
    
    Call lvw模块_ItemClick(lvw模块.ListItems(1))
    
    mblnFirst = False
End Sub

Private Sub lvw模块_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTemp As New ADODB.Recordset
    '读取出该模块拥有的权限，根据模块记录集恢复已选择的方法
    lvw权限.ListItems.Clear
    If lvw模块.SelectedItem Is Nothing Then Exit Sub
    
    Call SavePrivs
    
    mstrSQL = "Select 功能 From zlProgfuncs Where 系统=100 And 序号=" & Mid(Item.Key, 3)
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "获取指定模块的权限")
    With rsTemp
        Do While Not .EOF
            lvw权限.ListItems.Add , "K_" & .AbsolutePosition, !功能, , 1
            .MoveNext
        Loop
    End With
    
    If lvw权限.ListItems.Count <> 0 Then Call lvw权限_ItemClick(lvw权限.ListItems(1))
End Sub

Private Sub lvw权限_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '提取指定模块，指定权限所使用的方法
    Call SavePrivs
    
    Call ShowPrivs
End Sub

Private Sub SavePrivs()
    Dim intItem As Integer, intCount As Integer
    Dim strMethod As String
    Dim strField As String, strValue As String
    
    '删除选择模块指定权限的所有功能
    If lvw模块.SelectedItem Is Nothing Then Exit Sub
    If lvw权限.SelectedItem Is Nothing Then Exit Sub
    If mblnFirst Then
        mlngModul = Mid(lvw模块.SelectedItem.Key, 3)
        mstrPrivs = lvw权限.SelectedItem.Text
        Exit Sub
    End If
    
    With mrs模块
        .Filter = "模块=" & mlngModul & " And 权限串='" & mstrPrivs & "'"
        Do While Not .EOF
            .Delete
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    '保存上次选择的模块的指定权限
    strField = "模块|权限串|方法"
    strValue = mlngModul & "|" & mstrPrivs & "|"
    intCount = lvw方法.ListItems.Count
    For intItem = 1 To intCount
        If lvw方法.ListItems(intItem).Checked Then
            strMethod = lvw方法.ListItems(intItem).Text
            Call Record_Add(mrs模块, strField, strValue & strMethod)
        End If
    Next
    
    mlngModul = Mid(lvw模块.SelectedItem.Key, 3)
    mstrPrivs = lvw权限.SelectedItem.Text
End Sub

Private Sub ShowPrivs()
    Dim lng模块 As Long, str权限 As String
    Dim intItem As Integer, intCount As Integer
    
    lng模块 = Mid(lvw模块.SelectedItem.Key, 3)
    str权限 = lvw权限.SelectedItem.Text
    
    intCount = lvw方法.ListItems.Count
    For intItem = 1 To intCount
        lvw方法.ListItems(intItem).Checked = False
    Next
    
    With mrs模块
        .Filter = "模块=" & lng模块 & " And 权限串='" & str权限 & "'"
        Do While Not .EOF
            If mrs方法.RecordCount <> 0 Then mrs方法.MoveFirst
            mrs方法.Find "方法='" & !方法 & "'"
            If Not mrs方法.EOF Then
                lvw方法.ListItems("K_" & mrs方法!序号).Checked = True
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
End Sub
