Attribute VB_Name = "mdlCard"
Option Explicit

'���ѿ����������
Public gstrCardsAndProperty As String

'���ѿ���ʽ
Public Enum gCardFormat
    ���� = 0
    ȫ�� = 1
    ˢ����־ = 2
    �����ID = 3
    ���ų��� = 4
    ȱʡ��־ = 5
    �Ƿ�����ʻ� = 6
    �������� = 7
End Enum

Public Function zlfuncCard_Confirm(ByRef objSquareCard As Object, ByVal FrmMain As Form, ByVal lngModule As Long, _
    ByVal strPrivs As String, ByVal lng����id As Long, _
    ByVal lngCardID As Long, ByVal intType As Integer, _
    ByVal strNos As String) As Boolean
    
    If objSquareCard.zlSquareAffirm(FrmMain, lngModule, strPrivs, lng����id, lngCardID, False, intType, strNos) = False Then
        Exit Function
    End If
    zlfuncCard_Confirm = True
End Function

Public Function zlfuncCard_GetPatiName(ByRef objSquareCard As Object, ByVal lngCardID As Long, ByVal strCardNo As String) As String
    'һ��ͨ���ܣ�ͨ������ȡ��������
    Dim lng����id As Long
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    If Not objSquareCard Is Nothing Then
        'ͨ����ID�Ϳ��Ų��Ҳ���ID
        objSquareCard.zlGetPatiID CStr(lngCardID), strCardNo, False, lng����id
        If lng����id > 0 Then
            gstrSQL = "Select ���� From ������Ϣ Where ����id = [1] "
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "FindSpecialRow", lng����id)
            If Not rsData.EOF Then
                zlfuncCard_GetPatiName = UCase(rsData!����)
            End If
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlfuncCard_GetPatiID(ByRef objSquareCard As Object, ByVal lngCardID As Long, ByVal strCardNo As String) As Long
    'һ��ͨ���ܣ�ͨ������ȡ����ID
    Dim lng����id As Long
    
    On Error GoTo errHandle
    If Not objSquareCard Is Nothing Then
        'ͨ����ID�Ϳ��Ų��Ҳ���ID
        objSquareCard.zlGetPatiID CStr(lngCardID), strCardNo, False, lng����id
        
        If lng����id > 0 Then
            zlfuncCard_GetPatiID = lng����id
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlfuncCard_Ini(ByRef objSquareCard As Object, ByVal FrmMain As Form, ByVal lngModule As Long) As String
    'һ��ͨ�ӿڳ�ʼ�����������ѿ����������
    Dim strCards As String
    
    On Error Resume Next
    
    Set objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Not objSquareCard Is Nothing Then
        If objSquareCard.zlInitComponents(FrmMain, lngModule, glngSys, gstrDbUser, gcnOracle) = False Then
            Set objSquareCard = Nothing
        Else
            strCards = objSquareCard.zlGetIDKindStr
            
            '�������￨������Ժ��Ϊ���ѿ�
            zlfuncCard_Ini = Mid(strCards, InStr(1, strCards, "��|���￨"))
        End If
    End If
End Function

Public Sub zlfuncCard_SetCardMenu(ByVal lngModule As Long, ByVal objMenu As Object, ByVal strCards As String)
    '�������ѿ��˵���
    
End Sub

Public Sub zlfuncCard_SetText(ByVal objTxt As TextBox, ByVal strCardProperty As String)
    '�������������
    '���п���𣬸�ʽ������|ȫ��|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������);��
    objTxt.Text = ""
    objTxt.Tag = ""
    objTxt.MaxLength = 0
    
    objTxt.Tag = strCardProperty
    objTxt.MaxLength = Val(Split(strCardProperty, "|")(gCardFormat.���ų���))
    objTxt.PasswordChar = IIf(Trim(Split(strCardProperty, "|")(gCardFormat.��������)) <> "", "*", "")
End Sub

Public Sub zlfuncCard_Unload(ByRef objSquareCard As Object)
    'ж��һ��ͨ�ӿ�
    Set objSquareCard = Nothing
End Sub
