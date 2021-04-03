VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_Bc5300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrCode As String  '����������
Const mstrItem  As String = ",WBC,BAS#,BAS%,NEU#,NEU%,EOS#,EOS%,LYM#,LYM%,MON#,MON%,*ALY#,*ALY%,*LIC#,*LIC%,RBC,HGB,MCV,MCH,MCHC,RDW-CV,RDW-SD,HCT,PLT,MPV,PDW,PCT,"
' base 64 encoder string
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

'--------------------------------------------------------------------------------------------------------------------
'�����Ǳ���Ľӿں���

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
    Dim str_Code As String '1����������Ϣ
    Dim var_Code As Variant
    Dim lng_count As Long
    
    Dim str_�ֶηָ��� As String
    Dim str_����ָ��� As String
    Dim str_��������� As String
    Dim str_�ظ��ָ��� As String
    Dim str_ת���ַ� As String
    
    Dim str_Segment As String 'һ����Ϣ��(��chr(&HD)�ָ�)
    Dim var_Filed As Variant   '�ֶ�
    Dim str_Ψһ��ʶ As String '���ڱ�ʶһ����Ϣ
    Dim str_�汾 As String
    
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_�걾�� As String, str_���� As String
    Dim str_�ʿ� As String
    Dim str_��Ŀ As String
    Dim str_��� As String
    Dim str_Return As String
    
    Dim str_Base64 As String, str_WBC As String, str_RBC As String, str_PLT As String, str_img As String
    Dim str_WBCDIFF As String
    Dim str_��Ϣ���� As String '����5380��R01Ӧ��
    Dim str_�����ͺ� As String
    Dim str_��Ϣ��ʶ  As String
    
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
                 
        var_Code = Split(str_Code, Chr(&HD))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If Mid(str_Segment, 2, 3) = "MSH" Then
                '��Ϣͷ
                str_�ֶηָ��� = Mid(str_Segment, 5, 1)
                str_����ָ��� = Mid(str_Segment, 6, 1)
                str_��������� = Mid(str_Segment, 7, 1)
                str_�ظ��ָ��� = Mid(str_Segment, 8, 1)
                str_ת���ַ� = Mid(str_Segment, 9, 1)
                
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_Ψһ��ʶ = var_Filed(9)
                str_�ʿ� = var_Filed(10)
                str_�汾 = var_Filed(11)
                
                str_�����ͺ� = var_Filed(3)
                str_��Ϣ��ʶ = var_Filed(9)
                str_��Ϣ���� = var_Filed(10)
                
                If str_�汾 <> "2.3.1" Then
                    WriteLog "�汾����", "", ""
                End If
                str_Return = "": str_�걾�� = "": str_���� = ""
                
                If str_�����ͺ� = "BC-5380" And var_Filed(8) = "ORU^R01" Then
                    strCmd = Chr(&HB) & "MSH|^~\&|LIS||||" & Format(Now, "yyyyMMddHHmmss") & "||ORU^R02|" & str_��Ϣ��ʶ & "|" & str_��Ϣ���� & "|2.3.1||||||UNICODE" & Chr(&HD) _
                           & "MSA|AA|" & str_��Ϣ��ʶ & Chr(&H1C) & Chr(&HD)
                End If
            ElseIf Left(str_Segment, 3) = "OBR" Then
                '�걾�ţ�����
                If Val(str_�걾��) > 0 And str_Return <> "" Then
                    strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & IIf(str_�ʿ� = "Q", 1, 0) & str_Return
                    str_img = ""
                    If str_WBC <> "" Then str_img = str_img & "^WBC;103;1;" & str_WBC
                    If str_RBC <> "" Then str_img = str_img & "^RBC;103;1;" & str_RBC
                    If str_PLT <> "" Then str_img = str_img & "^PLT;103;1;" & str_PLT
                    If str_WBCDIFF <> "" Then str_img = str_img & "^DIFF;103;1;" & str_WBCDIFF
                    If str_img <> "" Then strResult = strResult & vbCrLf & Mid(str_img, 2)
                End If
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_�걾�� = var_Filed(3)
                str_���� = var_Filed(7)
                str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2) & ":" & Mid(str_����, 13, 2)
                str_Return = ""
            ElseIf Left(str_Segment, 3) = "OBX" Then
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_��Ŀ = var_Filed(3)
                If InStr(mstrItem, "," & Split(str_��Ŀ, str_����ָ���)(1) & ",") > 0 Then
                    str_Return = str_Return & "|" & Split(str_��Ŀ, str_����ָ���)(1) & "|" & Trim(var_Filed(5))
                ElseIf InStr(",15008,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'wbc
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_WBC = SaveBase64Bmp("WBC", str_�걾��, str_Base64)
                    
                ElseIf InStr(",15056,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'RBC
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_RBC = SaveBase64Bmp("RBC", str_�걾��, str_Base64)
                ElseIf InStr(",15116,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'PLT
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_PLT = SaveBase64Bmp("PLT", str_�걾��, str_Base64)
                ElseIf InStr(",15200,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'WBC DIFF
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_WBCDIFF = SaveBase64Bmp("DIFF", str_�걾��, str_Base64)
                End If
            End If
        Next
        If Val(str_�걾��) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & IIf(str_�ʿ� = "Q", 1, 0) & str_Return
            str_img = ""
            If str_WBC <> "" Then str_img = str_img & "^WBC;113;1;" & str_WBC
            If str_RBC <> "" Then str_img = str_img & "^RBC;113;1;" & str_RBC
            If str_PLT <> "" Then str_img = str_img & "^PLT;113;1;" & str_PLT
            If str_WBCDIFF <> "" Then str_img = str_img & "^DIFF;113;1;" & str_WBCDIFF
            If str_img <> "" Then strResult = strResult & vbCrLf & Mid(str_img, 2)
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "Bc5300_analyse", strSource, Err.Number & " " & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function

'---------------------------------------------------------------------------------------------------------------
'�������ڲ�����
Private Function SaveBase64Bmp(ByVal strType As String, ByVal str�걾�� As String, ByVal str2Decode As String) As String

' ******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
' ascii character equivalent. Stop converting at the end of the input string
' or when the first '=' (equal sign) is encountered.
'
' ******************************************************************************

    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim bits(1 To 4)    As Byte
    Dim frmPic As New frmGraph
    Dim ByteData() As Byte, lngCount As Long, strFileBmp As String, lngFileNum
    Dim strFileJpg As String
    
    lngCount = Len(str2Decode)
    ReDim ByteData(lngCount / 4 * 3)
    lngCount = 0
    ' for each 4 character group....
    For lPtr = 1 To Len(str2Decode) Step 4
        iLen = 4
        For iCtr = 0 To 3
            ' retrive the base 64 value, 4 at a time
            iValue = InStr(1, BASE64CHR, Mid$(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
                Case 1 To 64: bits(iCtr + 1) = iValue - 1
                ' =
                Case 65
                    iLen = iCtr
                    Exit For
                ' not found
                Case 0: Exit Function
            End Select
        Next

        ' convert the 4, 6 bit values into 3, 8 bit values
        bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
        bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
        bits(3) = (bits(3) And &H3) * &H40 + bits(4)

        ' add the three new characters to the output string
        For iCtr = 1 To iLen - 1
            ByteData(lngCount) = bits(iCtr)
            lngCount = lngCount + 1
        Next
    Next
    
    strFileBmp = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_" & strType & ".bmp"
    strFileJpg = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFileBmp For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    '-- 2012-05-08 ͼƬ��ɫ����
    frmPic.Picture1.AutoSize = True
    frmPic.Picture1.Picture = LoadPicture(strFileBmp)
    Call PicInvertAndSave(frmPic.Picture1, strFileJpg, "bmp")
    SaveBase64Bmp = strFileJpg

End Function