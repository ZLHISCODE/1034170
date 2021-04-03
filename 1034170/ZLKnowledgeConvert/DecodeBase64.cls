VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecodeBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BASE64CHR         As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private psBase64Chr(0 To 63)    As String

'��һ������Base64���ַ����н��뵽Դ�ַ���
Public Function DecodeBase64String(str2Decode As String) As String
    DecodeBase64String = StrConv(DecodeBase64Byte(str2Decode), vbUnicode)
End Function
 
'��һ������Base64���ַ����н��뵽Դ�ֽ�����
Public Function DecodeBase64Byte(str2Decode As String) As Byte()
 
    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim bits(1 To 4)    As Byte
    Dim strDecode       As String
    Dim str             As String
    Dim Output()        As Byte
    
    Dim iIndex          As Long

    Dim lFrom As Long
    Dim lTo As Long
    
    InitBase
    
    '//��ȥ�س�
    str = Replace(str2Decode, vbCrLf, "")
 
    '//ÿ4���ַ�һ�飨4���ַ���ʾ3���֣�
    For lPtr = 1 To Len(str) Step 4
        iLen = 4
        For iCtr = 0 To 3
            '//�����ַ���BASE64�ַ����е�λ��
            iValue = InStr(1, BASE64CHR, Mid$(str, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue  'A~Za~z0~9+/
                Case 1 To 64:
                    bits(iCtr + 1) = iValue - 1
                Case 65         '=
                    iLen = iCtr
                    Exit For
                    '//û�з���
                Case 0: Exit Function
            End Select
        Next
 
        '//ת��4��6��������Ϊ3��8������
        bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
        bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
        bits(3) = (bits(3) And &H3) * &H40 + bits(4)
 
        '//�����������ʼλ��
        lFrom = lTo
        lTo = lTo + (iLen - 1) - 1
                
        '//���¶����������
        ReDim Preserve Output(0 To lTo)
        
        For iIndex = lFrom To lTo
            Output(iIndex) = bits(iIndex - lFrom + 1)
        Next
 
        lTo = lTo + 1
        
    Next
    DecodeBase64Byte = Output
End Function


 
 
'��һ���ֽ��������Base64���룬�������ַ���
Public Function EncodeBase64Byte(sValue() As Byte) As String
    Dim lCtr                As Long
    Dim lPtr                As Long
    Dim lLen                As Long
    Dim sEncoded            As String
    Dim Bits8(1 To 3)       As Byte
    Dim Bits6(1 To 4)       As Byte
    
    Dim I As Integer
    
    InitBase
 
    For lCtr = 1 To UBound(sValue) + 1 Step 3
        For I = 1 To 3
            If lCtr + I - 2 <= UBound(sValue) Then
                Bits8(I) = sValue(lCtr + I - 2)
                lLen = 3
            Else
                Bits8(I) = 0
                lLen = lLen - 1
            End If
        Next
 
        '//ת���ַ���Ϊ���飬Ȼ��ת��Ϊ4��6λ(0-63)
        Bits6(1) = (Bits8(1) And &HFC) \ 4
        Bits6(2) = (Bits8(1) And &H3) * &H10 + (Bits8(2) And &HF0) \ &H10
        Bits6(3) = (Bits8(2) And &HF) * 4 + (Bits8(3) And &HC0) \ &H40
        Bits6(4) = Bits8(3) And &H3F
 
        '//����4�����ַ�
        For lPtr = 1 To lLen + 1
            sEncoded = sEncoded & psBase64Chr(Bits6(lPtr))
        Next
    Next
 
    '//����4λ����=���
    Select Case lLen + 1
        Case 2: sEncoded = sEncoded & "=="
        Case 3: sEncoded = sEncoded & "="
        Case 4:
    End Select
 
    EncodeBase64Byte = sEncoded
End Function
 

'���ַ�������Base64���벢�����ַ���
Public Function EncodeBase64String(str2Encode As String) As String
    Dim sValue()            As Byte
    sValue = StrConv(str2Encode, vbFromUnicode)
    EncodeBase64String = EncodeBase64Byte(sValue)
End Function




Private Sub InitBase()
    Dim iPtr    As Integer
    '��ʼ�� BASE64����
    For iPtr = 0 To 63
        psBase64Chr(iPtr) = Mid$(BASE64CHR, iPtr + 1, 1)
    Next
End Sub


