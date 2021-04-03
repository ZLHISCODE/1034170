VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIDKind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Private mintIDKind As Integer
Private mrsIDKind As New ADODB.Recordset
Private mblnCard As Boolean
Private mblnICCard As Boolean
Private mobjICCard As Object
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mcnOracle As ADODB.Connection
Private mbytApplyMode As Byte                                   '1-�����˵�;2-�����˵�
Private mobjFindKey As Object
Private mobjSquareCard As Object
Public Event AfterSelectSearchKey(ByVal Caption As String)
Public Event AfterInputComplete(ByVal CardNo As String)
Public Event AfterReadICCard(ByVal CardNo As String)
Public Event Location()

Public Event ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)

'�ӿ�
'######################################################################################################################
Public Function InitIDKind(ByVal bytApplyMode As Byte, ByVal cnOracle As ADODB.Connection, ByVal lngSysNo As Long, ByVal strUser As String, ByVal frmMain As Object, ByVal lngMoudal As Long, ByVal strIDKind As String) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    mbytApplyMode = bytApplyMode
    Set mcnOracle = cnOracle
    
    On Error GoTo errHand
    
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    End If
        
    'ȡȱʡ��ˢ����ʽ
    '����|�����|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    '��7λ��,��ֻ��������,��Ȼȡ������

    If mobjSquareCard.zlInitComponents(frmMain, lngMoudal, lngSysNo, strUser, mcnOracle) Then
        strIDKind = mobjSquareCard.zlGetIDKindStr(strIDKind)
        Set mrsIDKind = AnsyIDKind(strIDKind)
    End If
    
    InitIDKind = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IsCard() As Boolean
    
    IsCard = (Val(mrsIDKind("�����ID").Value) > 0)
    
End Function

Public Function IsReadCard() As Boolean
    
    IsReadCard = (Val(mrsIDKind("�����ID").Value) > 0 And Val(mrsIDKind("ˢ����־").Value) = 1)
    
End Function


Public Function zlGetPatiIDByCardNo(ByVal strInputCardNo As String, ByRef lngPatientKey As Long, Optional ByRef strPassword As String, Optional ByRef strErrMsg As String) As Boolean
    
'    strCardPasword = zlCommFun.zlStringEncode(strCardPassword)
    If IsCard Then
        zlGetPatiIDByCardNo = mobjSquareCard.zlGetPatiID(Val(mrsIDKind("�����ID").Value), strInputCardNo, False, lngPatientKey, strPassword, strErrMsg)
    Else
        zlGetPatiIDByCardNo = mobjSquareCard.zlGetPatiID(GetValue("ȫ��"), strInputCardNo, False, lngPatientKey, strPassword, strErrMsg)
    End If
    
End Function

Public Function GetPatientKey() As Long
    GetPatientKey = mobjSquareCard.zlGetPatiID
End Function

Public Property Get ShortName() As String
    '******************************************************************************************************************
    '����:���ض�����
    '����:
    '����:
    '******************************************************************************************************************
    
    If mrsIDKind.EOF = False Then ShortName = mrsIDKind("����").Value
    
End Property

Public Property Let ShortName(vData As String)
    '******************************************************************************************************************
    '����:���ض�����
    '����:
    '����:
    '******************************************************************************************************************
    
    mrsIDKind.Filter = ""
    mrsIDKind.Filter = "����='" & vData & "'"
    If mrsIDKind.RecordCount > 0 Then mintIDKind = mrsIDKind("���").Value
    
End Property

Public Property Get IsPasswordCheck() As Boolean
    '******************************************************************************************************************
    '����:���ض�����
    '����:
    '����:
    '******************************************************************************************************************
    
    If mrsIDKind.EOF = False Then LongName = mrsIDKind("ȫ��").Value
    
End Property

Public Property Get LongName() As String
    '******************************************************************************************************************
    '����:���ض�����
    '����:
    '����:
    '******************************************************************************************************************
    
    If mrsIDKind.EOF = False Then LongName = mrsIDKind("ȫ��").Value
    
End Property

Public Property Let LongName(vData As String)
    '******************************************************************************************************************
    '����:���ض�����
    '����:
    '����:
    '******************************************************************************************************************
    If IsOpen Then
        mrsIDKind.Filter = ""
        mrsIDKind.Filter = "ȫ��='" & vData & "'"
        If mrsIDKind.RecordCount > 0 Then mintIDKind = mrsIDKind("���").Value
    End If
    
End Property


Public Property Get CardLength() As Long
    If mrsIDKind.EOF = False Then CardLength = Val(mrsIDKind("���ų���").Value)
End Property

Private Function IsOpen() As Boolean
    
    If Not (mrsIDKind Is Nothing) Then
        If mrsIDKind.State = adStateOpen Then
            IsOpen = True
        End If
    End If
    
End Function

Public Function GetCard(ByRef rs As ADODB.Recordset) As Boolean

    If IsOpen Then
        mrsIDKind.Filter = ""
        mrsIDKind.Filter = "�����id>0"
        If mrsIDKind.RecordCount > 0 Then
            Set rs = mrsIDKind
            GetCard = True
        End If
    Else
        GetCard = False
    End If
    
End Function

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)

    RaiseEvent ShowIDCardInfo(strID, strName, strSex, strNation, datBirthDay, strAddress)
    
End Sub


Public Sub EnterInputBox(Optional ByVal strIDKind As String)
    '******************************************************************************************************************
    '����:���������
    '����:
    '����:
    '******************************************************************************************************************
    If strIDKind = "" Then strIDKind = GetValue("ȫ��")

    Select Case strIDKind
    Case "����"

        Call zlCommFun.OpenIme(True)
                    
    Case "�����"
        
        Call zlCommFun.OpenIme(False)
                
    Case "סԺ��"
        
        Call zlCommFun.OpenIme(False)
        
    Case "������"
        
        Call zlCommFun.OpenIme(False)
        
    Case "����֤��", "����֤"
        
        If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard
        If Not (mobjIDCard Is Nothing) Then Call mobjIDCard.SetEnabled(True)
        
    End Select
            
End Sub

Public Sub LeaveInputBox()
    If Not (mobjIDCard Is Nothing) Then Call mobjIDCard.SetEnabled(False)
End Sub


Public Sub InputKeyPress(KeyAscii As Integer, objInputTxt As Object)
    '******************************************************************************************************************
    '����:���������ֵ
    '����:
    '����:
    '******************************************************************************************************************
    Dim blnCard As Boolean
    Dim strKind As String
    
    strKind = GetValue("ȫ��")
    
    objInputTxt.PasswordChar = IIf(GetValue("��������") <> "", "*", "")
    Select Case strKind
    Case "����"
        
        blnCard = zlCommFun.InputIsCard(objInputTxt, KeyAscii, (GetValue("��������") <> ""))
        
    Case "�����"
        
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        
    Case "סԺ��"
        
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        
    Case "������"
        
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        
    Case "����֤��"
        
        If InStr("0123456789X" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        
    End Select
    
     'ˢ����ϻ���������س�
    If blnCard And Len(objInputTxt.Text) = Val(GetValue("���ų���")) - 1 And KeyAscii <> vbKeyBack Or KeyAscii = vbKeyReturn And Trim(objInputTxt.Text) <> "" Then
        
        If KeyAscii <> vbKeyReturn Then
            objInputTxt.Text = objInputTxt.Text & Chr(KeyAscii)
            objInputTxt.SelStart = Len(objInputTxt.Text)
            KeyAscii = 0
        End If

        RaiseEvent AfterInputComplete(Trim(objInputTxt.Text))
        
        zlControl.TxtSelAll objInputTxt
    End If
End Sub


Public Function ReadICCard() As String
    Dim strICCard As String
    
    If mobjICCard Is Nothing Then
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Set mobjICCard.gcnOracle = mcnOracle
    End If
    If Not mobjICCard Is Nothing Then
        strICCard = mobjICCard.Read_Card()
        If strICCard <> "" Then mblnICCard = True
        
        ReadICCard = strICCard
    End If
End Function

Public Sub ReadIDCard(strID As String, strName As String, strSex As String, strNation As String, datBirthDay As Date, strAddress As String, picPhoto As StdPicture)
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard
    Call mobjIDCard.ReadIDCard(strID, strName, strSex, strNation, datBirthDay, strAddress, picPhoto)
End Sub

Public Function GetPhotoAsStdPicture(picPhoto As StdPicture)
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard
    Call mobjIDCard.GetPhotoAsStdPicture(picPhoto)
End Function

'˽�й���
'######################################################################################################################
Private Function GetValue(ByVal strName As String) As String
    GetValue = mrsIDKind(strName).Value
End Function

Private Function AnsyIDKind(ByVal strIDKind As String) As ADODB.Recordset
    '******************************************************************************************************************
    '����:�ֽ�IDKind��
    '����:
    '����:
    '******************************************************************************************************************
    Dim rsIDKind As ADODB.Recordset
    Dim aryRow As Variant
    Dim aryCol As Variant
    Dim intRow As Integer
    Dim IntCount As Integer
    
    Set rsIDKind = New ADODB.Recordset
    With rsIDKind
        .Fields.Append "���", adInteger
        .Fields.Append "����", adVarChar, 10
        .Fields.Append "ȫ��", adVarChar, 50
        .Fields.Append "ˢ����־", adInteger                '1-����;0-ˢ��
        .Fields.Append "�����ID", adInteger
        .Fields.Append "���ų���", adInteger
        .Fields.Append "��������", adVarChar, 50
        .Fields.Append "����", adVarChar, 50
        .Open
        .ActiveConnection = Nothing
    End With
    
    If strIDKind <> "" Then
        aryRow = Split(strIDKind, ";")
        IntCount = 0
        For intRow = 0 To UBound(aryRow)
            
            If Trim(aryRow(intRow)) <> "" Then
                aryCol = Split(aryRow(intRow), "|")
                
                IntCount = IntCount + 1
                rsIDKind.AddNew
                
                rsIDKind("���").Value = IntCount
                rsIDKind("����").Value = Trim(aryCol(0))
                
                rsIDKind("ȫ��").Value = Trim(aryCol(1))
                If UBound(aryCol) >= 2 Then rsIDKind("ˢ����־").Value = Val(aryCol(2))
                If UBound(aryCol) >= 3 Then rsIDKind("�����ID").Value = Val(aryCol(3))
                If UBound(aryCol) >= 4 Then rsIDKind("���ų���").Value = Val(aryCol(4))
                If UBound(aryCol) >= 7 Then rsIDKind("��������").Value = Trim(aryCol(7))
                
                rsIDKind("����").Value = Trim(aryRow(intRow))
                
            End If
            
        Next
    End If
    
    If rsIDKind.RecordCount > 0 Then rsIDKind.MoveFirst
    
    Set AnsyIDKind = rsIDKind
    
End Function

Private Sub Class_Terminate()
    If Not (mobjIDCard Is Nothing) Then Call mobjIDCard.SetEnabled(False)
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
    
    If Not (mobjSquareCard Is Nothing) Then Set mobjSquareCard = Nothing
End Sub