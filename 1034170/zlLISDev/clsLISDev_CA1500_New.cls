VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CA1500_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:CA-1500
'����:�ձ����� sysmex
'�û�:��������ҽԺ
'����:�¶�
Private mstrCode As String
Private mstr�ܺ�  As String
Private mstr��Ŀ As String

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    
    Dim lng_S As Long, lng_E As Long, str_Code As String, str_Line As String, i As Integer
    Dim str_���� As String, str_�ܺ� As String, str_�ʿ� As String, str_���� As String, str_���� As String
    Dim str_��� As String, str_��Ŀ As String, strDateType As String, intDataStart As Integer
    Dim intNoStart As Integer, intNoLen As Integer
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    '-- �����ļ���ȡ��������
    
    strDateType = Trim(ReadIni("CA1500New", "���ڸ�ʽ", App.Path & "\CA1500New.ini"))
    If strDateType = "" Then Call WriteIni("CA1500New", "���ڸ�ʽ", "MM-DD-YY", App.Path & "\CA1500New.ini")

    intDataStart = Val(ReadIni("CA1500New", "���λ��", App.Path & "\CA1500New.ini"))
    If intDataStart <= 0 Then
        intDataStart = 60
        Call WriteIni("CA1500New", "���λ��", CStr(intDataStart), App.Path & "\CA1500New.ini")
    End If
    
    intNoStart = Val(ReadIni("CA1500New", "�걾λ��", App.Path & "\CA1500New.ini"))
    If intNoStart <= 0 Then
        intNoStart = 29
        Call WriteIni("CA1500New", "�걾λ��", CStr(intNoStart), App.Path & "\CA1500New.ini")
    End If
    
    intNoLen = Val(ReadIni("CA1500New", "�걾����", App.Path & "\CA1500New.ini"))
    If intNoLen <= 0 Then
        intNoLen = 15
        Call WriteIni("CA1500New", "�걾����", CStr(intNoLen), App.Path & "\CA1500New.ini")
    End If
    '------
    If InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    End If
    
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------------------------
        If Mid(str_Code, 2, 1) = "R" Then       '˫������ɨ����󷢵� ָ��
            str_���� = "20" & Mid(str_Code, 11, 2) & "-" & Mid(str_Code, 13, 2) & "-" & Mid(str_Code, 15, 2) & " " & Mid(str_Code, 17, 2) & ":" & Mid(str_Code, 19, 2)
            
            str_���� = Trim(Mid(str_Code, 27, 13))
            str_���� = Mid(str_Code, 40, 1) ' M"�ֶ�����,"A"������ֵ,"B"�������ȡ����ȡ,"C"���������ָ����Ϣ�趨
            
            mstr�ܺ� = Mid(str_Code, 21, 6)
            mstr��Ŀ = Mid(str_Code, 52)
            mstr��Ŀ = Replace(mstr��Ŀ, Space(6) & Chr(&H3), "")
            mstr��Ŀ = Replace(mstr��Ŀ, Space(6), ",")
            If Len(str_����) >= 13 Then
                strResult = strResult & "||" & str_���� & "|^^" & str_����
                strCmd = "1|"
            End If
            
        ElseIf Mid(str_Code, 2, 1) = "D" Then   '      �������صļ�����
            str_�ʿ� = Mid(str_Code, 10, 1)
            If str_�ʿ� = "C" Then
                str_�ʿ� = "1"
            Else
                str_�ʿ� = "0"
            End If
            If strDateType = "YY-MM-DD" Then
                
                str_���� = "20" & Mid(str_Code, 11, 2) & "-" & Mid(str_Code, 13, 2) & "-" & Mid(str_Code, 15, 2) & " " & Mid(str_Code, 17, 2) & ":" & Mid(str_Code, 19, 2)
            ElseIf strDateType = "DD-MM-YY" Then
                str_���� = "20" & Mid(str_Code, 15, 2) & "-" & Mid(str_Code, 13, 2) & "-" & Mid(str_Code, 11, 2) & " " & Mid(str_Code, 17, 2) & ":" & Mid(str_Code, 19, 2)
            Else
                '"MM-DD-YY"
                str_���� = "20" & Mid(str_Code, 15, 2) & "-" & Mid(str_Code, 11, 2) & "-" & Mid(str_Code, 13, 2) & " " & Mid(str_Code, 17, 2) & ":" & Mid(str_Code, 19, 2)
            End If
            
            str_�ܺ� = Mid(str_Code, 21, 6)
            str_���� = Trim(Mid(str_Code, intNoStart, intNoLen))   '27,13
            
            str_���� = Mid(str_Code, 40, 1) ' M"�ֶ�����,"A"������ֵ,"B"�������ȡ����ȡ,"C"���������ָ����Ϣ�趨
            str_��Ŀ = ""
            str_Code = Mid(str_Code, intDataStart) '52 ��  60
            'If str_���� = 48 Then Stop
            For i = 0 To (Len(str_Code) - 1) / 9
                If Len(str_Code) < 9 Then Exit For
                str_��� = Trim(Mid(str_Code, 4, 6))
                If str_��� <> "" Then
                    str_��Ŀ = str_��Ŀ & "|" & Mid(str_Code, 1, 3) & "|" & Trim(Mid(str_Code, 4, 5))
                End If
                str_Code = Mid(str_Code, 10)
            Next
            If Len(str_����) >= 13 Then
                strResult = strResult & "||" & str_���� & "|^^" & str_���� & "| |Ѫ��|" & str_�ʿ� & str_��Ŀ
            ElseIf Val(str_����) > 0 Then
                strResult = strResult & "||" & str_���� & "|" & str_���� & "| |Ѫ��|" & str_�ʿ� & str_��Ŀ
            End If
        End If
        
        '--------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
    End If
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    
    Dim var_�걾 As Variant, i_Count As Integer, i_Loop As Integer, str_��¼ As String, str_�걾�� As String
    Dim i_���� As Integer, str_���� As String, str_������Ŀ As String, str_�걾���� As String, str_tmp As String
    Dim str_SendItem As String, str_Line As String
    Dim str��Ŀ As String
    
    '-------------------
    'str��Ŀ = "040,050,060,080,090,120,150,170,180,190,200,210,220,300,310,320,330,510,600,610"
    str��Ŀ = mstr��Ŀ
    '----------------------------
    WriteLog "��ʼ", strSampleInfo, ""
    
    var_�걾 = Split(strSampleInfo, "||")
    SendSample = ""
    For i_Count = LBound(var_�걾) To UBound(var_�걾)
        str_��¼ = var_�걾(i_Count)
        
        If Val(Split(str_��¼, "|")(1)) = 0 Then
            str_�걾�� = Trim(Split(str_��¼, "|")(5))
        Else
            str_�걾�� = Val(Split(str_��¼, "|")(1))
        End If
        str_�걾���� = Split(str_��¼, "|")(3)
        i_���� = Val(Split(str_��¼, "|")(4))
        str_���� = Trim(Split(str_��¼, "|")(5))
        str_������Ŀ = ""
        str_tmp = ""
        For i_Loop = 10 To UBound(Split(str_��¼, "|"))
            If Val(Split(str_��¼, "|")(i_Loop)) > 0 Then
                str_tmp = str_tmp & "," & Mid(Split(str_��¼, "|")(i_Loop), 1, 2)
            End If
        Next
        
        For i_Loop = 0 To UBound(Split(str��Ŀ, ","))
            If InStr(str_tmp, Mid(Split(str��Ŀ, ",")(i_Loop), 1, 2)) > 0 Then
                str_������Ŀ = str_������Ŀ & Split(str��Ŀ, ",")(i_Loop) & Space(6)
            Else
                str_������Ŀ = str_������Ŀ & "000" & Space(6)
            End If
        Next
        If str_������Ŀ <> "" And str_���� <> "" Then
            
            str_Line = Chr(&H2) & "S2210101U" & Format(Now, "yyMMddHHmm") & mstr�ܺ� & str_���� & "B" & Space(11) & str_������Ŀ & Chr(&H3)

        End If
    Next
    iSendStep = 0
    WriteLog "����", "���͵���������Ϣ", SendSample
    
End Function

