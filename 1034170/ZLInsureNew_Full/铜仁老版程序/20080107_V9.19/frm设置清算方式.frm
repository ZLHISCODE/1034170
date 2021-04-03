VERSION 5.00
Begin VB.Form frm设置清算方式 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置清算方式"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frm设置清算方式.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   12
      Top             =   3000
      Width           =   5085
   End
   Begin VB.CommandButton cmd恢复 
      Caption         =   "还原(&R)"
      Height          =   350
      Left            =   180
      TabIndex        =   9
      Top             =   2310
      Width           =   1100
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "下载(&W)"
      Height          =   350
      Left            =   180
      TabIndex        =   10
      Top             =   3180
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2340
      TabIndex        =   7
      Top             =   3180
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3600
      TabIndex        =   8
      Top             =   3180
      Width           =   1100
   End
   Begin VB.TextBox txt清算方式 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1290
      TabIndex        =   6
      Top             =   1710
      Width           =   3525
   End
   Begin VB.TextBox txt疾病信息 
      Height          =   300
      Left            =   1290
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CommandButton cmd疾病信息 
      Caption         =   "…"
      Height          =   300
      Left            =   4530
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label lblNote 
      Caption         =   "    如果选择错误，通过“还原”按钮可以恢复默认的病种及清算方式，然后点击确定提交本次修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   2
      Left            =   1290
      TabIndex        =   11
      Top             =   2220
      Width           =   3615
   End
   Begin VB.Label lblNote 
      Caption         =   "    请选择一个单病种，本次住院将按该病种对应的清算方式对费用进行结算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   405
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   750
      Width           =   3615
   End
   Begin VB.Label lbl清算方式 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "清算方式(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   210
      TabIndex        =   5
      Top             =   1770
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frm设置清算方式.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    如果是第一次使用或医院的单病种数据发生变化，请使用下载功能，将单病种清算数据下载到本地。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   150
      Width           =   3645
   End
   Begin VB.Label lbl疾病信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "单病种(&J)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   390
      TabIndex        =   2
      Top             =   1380
      Width           =   810
   End
End
Attribute VB_Name = "frm设置清算方式"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Private mblnOK As Boolean
Private mint险类 As Integer
Private mlng病人ID As Long
Private mstr卡号 As String
Private mstr医保号 As String
Private mstr分中心编号 As String
Private mstr密码 As String
Private mrs病种 As New ADODB.Recordset

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", "0101")
    If Not CommServer("GETHOSPSINGLEILLNESS") Then Exit Sub
    MsgBox "下载成功！", vbIbeam, gstrSysName
End Sub

Private Sub cmdOK_Click()
    If txt疾病信息.Tag = "" Then
        MsgBox "请选择一个单病种！", vbInformation, gstrSysName
        txt疾病信息.SetFocus
        Exit Sub
    End If
    
    '将选择的清算方式上传到医保中心
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstr医保号)
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", mstr分中心编号)
    Call InsertChild(mdomInput.documentElement, "RECKONINGTYPE", txt清算方式.Tag)
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", txt疾病信息.Tag)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    If CommServer("SETRECKONINGTYPE") = False Then Exit Sub
    
    On Error Resume Next
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & mint险类 & ",'单病种','''" & txt疾病信息.Tag & "|" & txt清算方式.Tag & "''')"
    Call ExecuteProcedure("保存单病种编码")
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd恢复_Click()
    txt疾病信息.Text = ""
    txt疾病信息.Tag = "00000001"
    txt清算方式.Text = "控制线清算方式"
    txt清算方式.Tag = 1
End Sub

Private Sub cmd疾病信息_Click()
    Dim blnReturn As Boolean
    blnReturn = frmListSel.ShowSelect(mrs病种, "ID", "单病种选择", "请选择单病种：")
    If Not blnReturn Then mrs病种.Filter = 0: Exit Sub
    
    txt疾病信息.Text = "(" & mrs病种!编码 & ")" & mrs病种!名称
    txt疾病信息.Tag = mrs病种!编码
    txt清算方式.Tag = mrs病种!清算方式
    Select Case mrs病种!清算方式
    Case 4
        txt清算方式.Text = "单病种按时间包干清算方式"
    Case 3
        txt清算方式.Text = "单病种按人次定额清算方式"
    Case Else
        txt清算方式.Text = "控制线清算方式"
    End Select
    mrs病种.Filter = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    '读取该病人的医保信息
    gstrSQL = "Select 单病种 From 保险帐户 Where 病人ID=" & mlng病人ID & " And 险类=" & mint险类
    Call OpenRecordset(rsTemp, "读取该病人的医保信息")
    txt疾病信息.Text = NVL(rsTemp!单病种)
    If InStr(1, txt疾病信息.Text, "|") <> 0 Then txt疾病信息.Text = Mid(txt疾病信息.Text, 1, InStr(1, txt疾病信息.Text, "|") - 1)
    txt疾病信息.Tag = txt疾病信息.Text
    
    Call Get验证_贵阳(mstr卡号, mstr医保号, mstr分中心编号, mstr密码, mlng病人ID)
    
    Call 获取单病种
    Call 显示病种信息
End Sub

Public Function ShowSelect(ByVal lng病人ID As Long, ByVal int险类 As Integer, ByVal frmParent As Object) As Boolean
    mblnOK = False
    mlng病人ID = lng病人ID
    mint险类 = int险类
    Me.Show 1, frmParent
    ShowSelect = mblnOK
End Function

Private Function 获取单病种() As Boolean
    Dim strFields As String, strValues As String
    Dim str编码 As String, str名称 As String, str简码 As String, str清算方式 As String, str清算标准 As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Set mrs病种 = New ADODB.Recordset
    strFields = "ID," & adVarChar & ",30|" & _
                "编码," & adLongVarChar & ",30|" & _
                "名称," & adLongVarChar & ",200|" & _
                "简码," & adLongVarChar & ",30|" & _
                "清算方式," & adLongVarChar & ",10|" & _
                "清算标准," & adLongVarChar & ",500"
    Call Record_Init(mrs病种, strFields)
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", "0101")   '分中心编码固定
    If CommServer("QUERYHOSPSINGLEILLNESS") = False Then Exit Function
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then Exit Function
    '根据编码得到险种名称
    strFields = "ID|编码|名称|简码|清算方式|清算标准"
    For Each nodRow In nodRowset.childNodes
        str编码 = GetAttributeValue(nodRow, "SINGLEILLNESSCODE")
        str名称 = GetAttributeValue(nodRow, "SINGLEILLNESSNAME")
        str清算方式 = GetAttributeValue(nodRow, "RECKONINGTYPE")
        str清算标准 = GetAttributeValue(nodRow, "PAYSTD")
        str简码 = zlCommFun.SpellCode(str名称)
        strValues = str编码 & "|" & str编码 & "|" & str名称 & "|" & str简码 & "|" & str清算方式 & "|" & str清算标准
        Call Record_Add(mrs病种, strFields, strValues)
    Next
    获取单病种 = True
End Function

Private Function 显示病种信息(Optional ByVal bln任意匹配 As Boolean = False) As Boolean
    Dim blnReturn As Boolean
    Dim strInput As String, strFilter As String
    
    If Trim(txt疾病信息.Text) = "" Then Exit Function
    'bln任意匹配:如果不是任意匹配，表明是从数据库里读上次已选择的病种，因此采取从左匹配，怕有编码存在相似的，而操作通过输入来查病种时需要任意匹配
    If bln任意匹配 Then
        strInput = UCase("'" & txt疾病信息.Text & "*'")
        strFilter = "编码 Like " & strInput & " Or 名称 Like " & strInput & " Or 简码 Like " & strInput
    Else
        strInput = UCase("'" & txt疾病信息.Text & "'")
        strFilter = "编码=" & strInput
    End If
    
    With mrs病种
        .Filter = strFilter
        If .RecordCount = 0 Then
            If bln任意匹配 Then
                MsgBox "没有找到指定的单病种！[病种编码为:" & UCase(txt疾病信息.Text) & "]", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txt疾病信息)
            txt疾病信息.Text = ""
            txt疾病信息.Tag = ""
            txt清算方式.Text = ""
            txt清算方式.Tag = 1
            .Filter = 0
            Exit Function
        Else
            If mrs病种.RecordCount > 1 Then
                blnReturn = frmListSel.ShowSelect(mrs病种, "ID", "单病种选择", "请选择单病种：")
            Else
                blnReturn = True
            End If
            If blnReturn = False Then
                txt疾病信息.Text = ""
                txt疾病信息.Tag = ""
                txt清算方式.Text = ""
                txt清算方式.Tag = 1
                Call zlControl.TxtSelAll(txt疾病信息)
            Else
                txt疾病信息.Text = "(" & mrs病种!编码 & ")" & mrs病种!名称
                txt疾病信息.Tag = mrs病种!编码
                txt清算方式.Tag = mrs病种!清算方式
                Select Case mrs病种!清算方式
                Case 4
                    txt清算方式.Text = "单病种按时间包干清算方式"
                Case 3
                    txt清算方式.Text = "单病种按人次定额清算方式"
                Case Else
                    txt清算方式.Text = "控制线清算方式"
                End Select
                显示病种信息 = True
            End If
        End If
        .Filter = 0
    End With
End Function

Private Sub txt疾病信息_GotFocus()
    Call zlControl.TxtSelAll(txt疾病信息)
End Sub

Private Sub txt疾病信息_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt疾病信息.Text) = "" Then Exit Sub
    
    If Not 显示病种信息(True) Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
