VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Key As String
Private mrsDataSource As New ADODB.Recordset

Public Property Set DataSource(vData As ADODB.Recordset)
    Set mrsDataSource = vData
End Property

Public Property Get DataSource() As ADODB.Recordset
    Set DataSource = mrsDataSource
End Property

Public Property Get FieldValue(ByVal strFieldName As String) As String
    
    If IsNull(mrsDataSource(strFieldName).Value) Then
        FieldValue = ""
    Else
        FieldValue = mrsDataSource(strFieldName).Value
    End If
    
End Property

Public Sub Sort(ByVal strSort As String)
    
    mrsDataSource.Sort = strSort
    
End Sub

Public Sub Filter(ByVal strFilter As String)
    
    mrsDataSource.Filter = ""
    mrsDataSource.Filter = strFilter
        
End Sub

Public Function FilterValue(ByVal strFilter As String, ByVal strFieldValue As String) As String
    
    mrsDataSource.Filter = ""
    mrsDataSource.Filter = strFilter
    If mrsDataSource.RecordCount > 0 Then
        mrsDataSource.MoveFirst
        FilterValue = zlCommFun.NVL(mrsDataSource(strFieldValue).Value)
    End If

End Function

Public Function FilterDataSource(ByVal strFilter As String) As ADODB.Recordset

    mrsDataSource.Filter = ""
    mrsDataSource.Filter = strFilter

    Set FilterDataSource = mrsDataSource
End Function

Public Property Get FieldValueJoin(ByVal strKeyName As String, ByVal strKeyValue As String, ByVal strFieldName As String, Optional ByVal strJoinChar As String = "、") As String
    Dim strTmp As String
    Dim rsValue As ADODB.Recordset
    
    mrsDataSource.Filter = ""
    
    If strKeyName <> "" Then
        mrsDataSource.Filter = strKeyName & "='" & strKeyValue & "'"
    End If
    If mrsDataSource.RecordCount > 0 Then
        
        Set rsValue = New ADODB.Recordset
        With rsValue
            .Fields.Append "结果", adVarChar, 255
            .Open
        End With
        
        mrsDataSource.MoveFirst
        Do While Not mrsDataSource.EOF
            
            If IsNull(mrsDataSource(strFieldName).Value) = False Then
                
                '重复的自动不加
                rsValue.Filter = ""
                rsValue.Filter = "结果='" & mrsDataSource(strFieldName).Value & "'"
                If rsValue.RecordCount = 0 Then
                    rsValue.AddNew
                    rsValue("结果").Value = mrsDataSource(strFieldName).Value
                End If
            End If
    
            mrsDataSource.MoveNext
        Loop
        
        rsValue.Filter = ""
        If rsValue.RecordCount > 0 Then
            rsValue.MoveFirst
            Do While Not rsValue.EOF
                strTmp = strTmp & strJoinChar & rsValue("结果").Value
                rsValue.MoveNext
            Loop
            If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        End If
    End If
    
    FieldValueJoin = strTmp
    
End Property

Public Function FieldSum(ByVal strSumField As String, ByVal strFilter As String, ByVal strSumKeyField As String, ByVal strShowField As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim strSumKeyValue As String
    Dim sglSum As Single
    Dim rsData As ADODB.Recordset
    Dim varShowField As Variant
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    varShowField = Split(strShowField, ",")
    
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append strSumField, adVarChar, 50
        .Fields.Append strSumKeyField, adVarChar, 50
        For intLoop = 0 To UBound(varShowField)
            .Fields.Append CStr(varShowField(intLoop)), adVarChar, 50
        Next
        .Open
    End With
        
    mrsDataSource.Filter = ""
    mrsDataSource.Filter = strFilter
    If mrsDataSource.RecordCount > 0 Then
        mrsDataSource.MoveFirst
        
        
        Do While Not mrsDataSource.EOF
            
            If IsNull(mrsDataSource(strSumField).Value) = False Then
                sglSum = Val(mrsDataSource(strSumField).Value)
            Else
                sglSum = 0
            End If
            
            strSumKeyValue = mrsDataSource(strSumKeyField).Value
            
            'Add
            rsData.Filter = ""
            rsData.Filter = strSumKeyField & "=" & strSumKeyValue
            If rsData.RecordCount = 0 Then
                rsData.AddNew
                rsData(strSumKeyField).Value = strSumKeyValue
                For intLoop = 0 To UBound(varShowField)
                    If IsNull(mrsDataSource(CStr(varShowField(intLoop))).Value) Then
                        rsData(CStr(varShowField(intLoop))).Value = ""
                    Else
                        rsData(CStr(varShowField(intLoop))).Value = mrsDataSource(CStr(varShowField(intLoop))).Value
                    End If
                Next
                rsData(strSumField).Value = 0
            End If
            
            If IsNull(mrsDataSource(strSumField).Value) Then
                rsData(strSumField).Value = Val(rsData(strSumField).Value)
            Else
                rsData(strSumField).Value = Val(rsData(strSumField).Value) + Val(mrsDataSource(strSumField).Value)
            End If
            
'            rsData(strSumField).Value = Val(rsData(strSumField).Value) + Val(mrsDataSource(strSumField).Value)

            mrsDataSource.MoveNext
        Loop
    End If
    
    
    If Not (rsData Is Nothing) Then
        If rsData.State = adStateOpen Then
            rsData.Filter = ""
            If rsData.RecordCount > 0 Then
                rsData.MoveFirst
            End If
        End If
    End If
    
    Set FieldSum = rsData
    
    Exit Function
    
    '
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

