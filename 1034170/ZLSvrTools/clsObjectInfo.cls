VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'��˵����
'���÷�Χ�����ڽ����ڱ��Լ��м�顣����������ʱ��δ֧��
'ʹ�÷�����
'               ����һ����ʱ��������ȡ�������ʵ��
'               Set objTmp=New clsObjectInfo --������ȡ�������ʵ��
'               Set objParent=objTmp.GetObject(....)--��ȡ��Ҫʵ���ĸ���ʵ��
'               Set Objcurrent=objTmp.GetObject(....,objParent)--��ȡ��ǰʵ����������ʵ����ϵ
'               ʹ�ö�����ʱ����鲢�޸��Ӽ����Զ�����޸�������

Public Enum ObjectType
    OT_Table = 0
    OT_Column
    OT_Constraint
    OT_Index
    OT_Sequence
    OT_FildLength
End Enum
Private mstrObjectName As String
Private motObjectType As ObjectType
Private mobjParent As clsObjectInfo
Private mstrOwner As String
Private mstrAdjustSQL As String
Private mblnChecked As Boolean '�����Ѿ����ɹ�
Private mcnOwner As ADODB.Connection '�����ߵ�����
Private mlngLength As Long  '�ֶγ���

'ObjectName:��������
Public Property Get ObjectName() As String
    ObjectName = mstrObjectName
End Property
Public Property Let ObjectName(ByVal strObjectName As String)
    mstrObjectName = strObjectName
End Property
'ObjectType:��������
Public Property Get ObjectType() As ObjectType
    ObjectType = motObjectType
End Property

Public Property Let ObjectType(ByVal otObjectType As ObjectType)
    motObjectType = otObjectType
End Property
'ObjectParent:������
Public Property Get ObjectParent() As clsObjectInfo
    Set ObjectParent = mobjParent
End Property

Public Property Set ObjectParent(ByVal objParent As clsObjectInfo)
    Set mobjParent = objParent
End Property
'ObjectOwner:���������û�
Public Property Get ObjectOwner() As String
    ObjectOwner = mstrOwner
End Property

Public Property Let ObjectOwner(ByVal strOwner As String)
    mstrOwner = strOwner
End Property
'AdjustSQL:����SQL
Public Property Get AdjustSQL() As String
    AdjustSQL = mstrAdjustSQL
End Property

Public Property Let AdjustSQL(ByVal strAdjustSQL As String)
    mstrAdjustSQL = strAdjustSQL
End Property
'Checked:�Ƿ��Ѿ�ͨ�����
Public Property Get Checked() As Boolean
    Checked = mblnChecked
End Property

Public Property Let Checked(ByVal blnChecked As Boolean)
    mblnChecked = blnChecked
End Property
'OwnerConn:����������
Public Property Get OwnerConn() As ADODB.Connection
    Set OwnerConn = mcnOwner
End Property

Public Property Set OwnerConn(ByVal cnOwner As ADODB.Connection)
    Set mcnOwner = cnOwner
End Property

'mlngLength:�ֶγ���
Public Property Get FildLength() As Long
    FildLength = mlngLength
End Property

Public Property Let FildLength(ByVal lngLength As Long)
    mlngLength = lngLength
End Property

Public Function GetObject(ByVal strObjectName As String, ByVal otObjectType As ObjectType, Optional ByVal strAdjustSQL As String, _
                                        Optional ByVal strObjectOwner As String = "ZLTOOLS", Optional ByVal objParent As clsObjectInfo, Optional ByVal cnOwner As ADODB.Connection, _
                                        Optional ByVal lngFildLen As Long) As clsObjectInfo
'����:�������󣬸÷���ʹ�þ���鿴����˵�����ڸ��࿪ͷ
'������strObjectName=������
'         otObjectType=��������
'         strAdjustSQL=����SQL
'         strObjectOwner=���������û�
'         objParent=��ǰ����ĸ�������
' ���أ���ǰ����
    Dim objCurent As clsObjectInfo
    Set objCurent = New clsObjectInfo
    objCurent.ObjectName = strObjectName
    objCurent.ObjectType = otObjectType
    objCurent.AdjustSQL = strAdjustSQL
    objCurent.ObjectOwner = strObjectOwner
    Set objCurent.ObjectParent = objParent
    Set objCurent.OwnerConn = cnOwner
    objCurent.FildLength = lngFildLen
    Set GetObject = objCurent
End Function

Public Function ObjectCheck(Optional ByVal blnMsg As Boolean = True) As Boolean
'���ܣ��Ե�ǰ������м���Լ��޸������Զ��Ե��ڶ���ĸ������м���޸���
'������blnMsg=�Ƿ������Ϣ��ʾ
'���أ��Ƿ���ɹ�������������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strMsg As String
    Dim arrAdjustSQL As Variant, strOtherSQL As String
    Dim strTableName As String
    Dim blnOk As Boolean, i As Integer
    Dim cnTmp As ADODB.Connection
    
    On Error GoTo errH
    '���ڸ��������鸸�����������ʧ�ܣ����˳�
    If Not Me.ObjectParent Is Nothing Then
        If Not Me.ObjectParent.Checked Then '������δͨ�����
            If Not Me.ObjectParent.ObjectCheck(blnMsg) Then
                Exit Function
            Else
                Me.ObjectParent.Checked = True
            End If
        End If
    End If
    If Not Me.Checked Then '�Ƿ��Ѿ�ͨ�����
        '���ж�����
        Select Case Me.ObjectType
            Case OT_Table
                If Me.OwnerConn Is Nothing Then
                    strSQL = "Select 1 FROM ALL_TABLES WHERE TABLE_NAME=[1] AND OWNER=[2]"
                    strMsg = "��""" & Me.ObjectOwner & "." & Me.ObjectName & """�����ڣ����ֹ������ñ�" & IIf(UCase(Me.ObjectOwner) = "ZLTOOLS", "��ͬʱ����ͬ��ʲ���Ȩ��", "��")
                Else
                    strSQL = "Select 1 FROM User_TABLES WHERE TABLE_NAME=[1]"
                End If
                If UCase(Me.ObjectOwner) = "ZLTOOLS" Then
                    strOtherSQL = "create public synonym " & Me.ObjectName & "  for  ZLTOOLS." & Me.ObjectName
                    strOtherSQL = strOtherSQL & "|Grant Select On ZLTOOLS." & Me.ObjectName & " to PuBLIC"
                    strOtherSQL = strOtherSQL & "|Grant Select,Insert,Update,Delete On ZLTOOLS." & Me.ObjectName & " to " & gstrUserName & "With Grant Option"
                End If
                strTableName = Me.ObjectName
            Case OT_Column
                If Me.OwnerConn Is Nothing Then
                    strSQL = "Select 1 From All_Tab_Columns Where Owner =[2] And Table_Name = [1] And Column_Name = [3]"
                Else
                    strSQL = "Select 1 From User_Tab_Columns Where Table_Name = [1] And Column_Name = [3]"
                End If
                strMsg = "��""" & Me.ObjectOwner & "." & Me.ObjectParent.ObjectName & """����""" & Me.ObjectName & """ �����ڣ����ֹ����Ӹ��С�"
                strTableName = Me.ObjectParent.ObjectName
            Case OT_FildLength
                If Me.OwnerConn Is Nothing Then
                    strSQL = "Select 1 From All_Tab_Columns Where Owner =[2] And Table_Name = [1] And Column_Name = [3] And DATA_LENGTH<[4]"
                Else
                    strSQL = "Select 1 From User_Tab_Columns Where Table_Name = [1] And Column_Name = [3] And DATA_LENGTH<[4]"
                End If
                strTableName = Me.ObjectParent.ObjectName
        End Select
        Set rsTmp = gclsBase.OpenSQLRecord(IIf(Me.OwnerConn Is Nothing, gcnOracle, Me.OwnerConn), strSQL, gstrSysName, UCase(strTableName), UCase(Me.ObjectOwner), UCase(Me.ObjectName), Me.FildLength)
        blnOk = Not rsTmp.EOF
        If blnOk Then
            If Me.ObjectType = OT_FildLength Then
                arrAdjustSQL = Split(Me.AdjustSQL, "|")
                On Error Resume Next
                Set cnTmp = GetConnection("ZLTOOLS")
                If Not cnTmp Is Nothing Then
                    For i = LBound(arrAdjustSQL) To UBound(arrAdjustSQL)
                        cnTmp.Execute arrAdjustSQL(i)
                    Next
                End If
            End If
        Else
            If Me.ObjectType <> OT_FildLength Then
                If Me.AdjustSQL = "" Then
                    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                Else
                    If UCase(Me.ObjectOwner) = "ZLTOOLS" Then
                        Set cnTmp = GetConnection("ZLTOOLS")
                        If cnTmp Is Nothing Then
                            Set cnTmp = gcnOracle
                        End If
                    ElseIf mcnOwner Is Nothing Then
                        Set cnTmp = gcnOracle
                    Else
                        Set cnTmp = mcnOwner
                    End If
                    arrAdjustSQL = Split(Me.AdjustSQL, "|")
                    On Error Resume Next
                    For i = LBound(arrAdjustSQL) To UBound(arrAdjustSQL)
                        cnTmp.Execute arrAdjustSQL(i)
                        If err.Number <> 0 Then
                            err.Clear
                            '�޸���ʱ�����������������Զ��˳�
                            '���޸�ʧ�ܣ����Զ��˳�
                            If i = LBound(arrAdjustSQL) Then
                                If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
                                Exit For
                            End If
                        End If
                    Next
                    If strOtherSQL <> "" Then
                        arrAdjustSQL = Split(strOtherSQL, "|")
                        For i = LBound(arrAdjustSQL) To UBound(arrAdjustSQL)
                            cnTmp.Execute arrAdjustSQL(i)
                            If err.Number <> 0 Then err.Clear
                        Next
                    End If
                    On Error GoTo errH
                    Set rsTmp = gclsBase.OpenSQLRecord(IIf(Me.OwnerConn Is Nothing, gcnOracle, Me.OwnerConn), strSQL, gstrSysName, UCase(strTableName), UCase(Me.ObjectOwner), UCase(Me.ObjectName))
                    blnOk = Not rsTmp.EOF
                End If
            End If
        End If
        ObjectCheck = blnOk
        Me.Checked = blnOk
    Else
        ObjectCheck = Me.Checked
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function
