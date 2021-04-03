Attribute VB_Name = "mdlRelease"
Option Explicit
Public Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Public Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Public Const madDbDateDefault As Integer = 20               '日期型字段缺省长度
Public Enum 方法
    身份验证
    身份验证_自助挂号
    帐户余额
    门诊挂号
    门诊挂号作废
    门诊虚拟结算
    门诊结算
    门诊结算作废
    个人帐户转预交
    预交退个人帐户
    住院虚拟结算
    住院结算
    住院结算作废
    入院登记
    入院登记撤销
    出院登记
    出院登记撤销
    费用明细上传
    住院信息变动
    获取医保项目信息
    获取医保项目相关信息
    病种选择
    取消就诊登记
    总数
End Enum

Sub Main()
    frmUserLogin.Show 1
    If gcnOracle.State = 0 Then Exit Sub
    
    Call InitCommon(gcnOracle)
    frm医保部件发布.Show
End Sub

Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
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

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
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

Public Function CopyNewRec(ByVal SourceRec As ADODB.Recordset) As ADODB.Recordset
    Dim RecTarget As New ADODB.Recordset
    Dim intFields As Integer, LngLocate As Long
    '编制人:朱玉宝
    '编制日期:2000-11-02
    '该记录集与凭证控件对应
    '也使用于保存
    
    LngLocate = -1
    Set RecTarget = New ADODB.Recordset
    With RecTarget
        If .State = 1 Then .Close
        If SourceRec.RecordCount <> 0 Then
            On Error Resume Next
            Err = 0
            LngLocate = SourceRec.AbsolutePosition
            If Err <> 0 Then LngLocate = -1
            SourceRec.MoveFirst
        End If
        For intFields = 0 To SourceRec.Fields.Count - 1
            .Fields.Append SourceRec.Fields(intFields).Name, SourceRec.Fields(intFields).Type, SourceRec.Fields(intFields).DefinedSize, adFldIsNullable     '0:表示新增
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
        Do While Not SourceRec.EOF
            .AddNew
            For intFields = 0 To SourceRec.Fields.Count - 1
                .Fields(intFields) = SourceRec.Fields(intFields).Value
            Next
            .Update
            SourceRec.MoveNext
        Loop
    End With
    
    If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
    If LngLocate > 0 Then SourceRec.Move LngLocate - 1
    Set CopyNewRec = RecTarget
End Function
