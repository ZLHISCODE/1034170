Attribute VB_Name = "mdlJSCA"


Public Enum DataType
    responseText = 1
    ResponseBody = 2
End Enum

Public Function GetData(ByVal Url As String, ByVal DataStic As DataType) As Variant
     
    On Error GoTo errH:

    Dim objXMLHTTP As Object
    Dim strData As String
    Dim bytData() As Byte
     
    Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
     
    objXMLHTTP.Open "GET", Url, True
    objXMLHTTP.Send
     
    Do While objXMLHTTP.ReadyState <> 4
        DoEvents
    Loop
    '--------------------------------------函数返回
    Select Case DataStic
    Case responseText
        '--------------------------------直接返回字符串
        DataS = objXMLHTTP.responseText
        GetData = DataS
    Case ResponseBody
        '--------------------------------直接返回二进制
        DataB = objXMLHTTP.ResponseBody
        GetData = DataB
    Case ResponseBody + responseText
        '------------------------------二进制转字符串[直接返回字串出现乱码时尝试]
        DataS = BytesToStr(objXMLHTTP.ResponseBody)
        GetData = DataS
    Case Else
        '--------------------------------无效的返回
        GetData = ""
    End Select
    '--------------------------------------释放空间
    Set objXMLHTTP = Nothing
    Exit Function
errH:
    MsgBox "错误号:" & Err.Number & "错误信息:" & Err.Description
End Function
  
Public Function PostData(ByVal strURL As String, ByVal strSendData As String, ByVal DataStic As DataType) As Variant
    On Error GoTo errH:
     
    Dim objXMLHTTP As Object
    Dim strData As String
    Dim bytData() As Byte
     
    Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
     
    objXMLHTTP.Open "POST", strURL, True
    objXMLHTTP.setRequestHeader "Content-Length", Len(PostData)
    objXMLHTTP.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded;charset=GB2312"
    objXMLHTTP.Send (strSendData)
     
    Do Until objXMLHTTP.ReadyState = 4
        DoEvents
    Loop
    '-----------------------------函数返回
    Select Case DataStic
    Case responseText
        '--------------------------------直接返回字符串
        strData = objXMLHTTP.responseText
        PostData = strData
    Case ResponseBody
        '--------------------------------直接返回二进制
        bytData = objXMLHTTP.ResponseBody
        PostData = bytData
    Case ResponseBody + responseText
        '---------------------------二进制转字符串[直接返回字串出现乱码时尝试]
        strData = BytesToStr(objXMLHTTP.ResponseBody)
        PostData = strData
    Case Else
        '--------------------------------无效的返回
        PostData = ""
    End Select
    '------------------------------------释放空间
    Set objXMLHTTP = Nothing
    
    Exit Function
errH:
    MsgBox "错误号:" & Err.Number & "错误信息:" & Err.Description
End Function
  
Function BytesToStr(ByVal varInput As Variant) As String
    Dim strReturn As String
    Dim strCharCode As String
    For i = 1 To LenB(varInput)
        strCharCode = AscB(MidB(varInput, i, 1))
        If strCharCode < &H80 Then
            strReturn = strReturn & Chr(strCharCode)
        Else
            NextCharCode = AscB(MidB(varInput, i + 1, 1))
            strReturn = strReturn & Chr(CLng(strCharCode) * &H100 + CInt(NextCharCode))
            i = i + 1
        End If
    Next
    BytesToStr = strReturn
End Function

Public Function SaveBase64ToFile(ByVal strType As String, ByVal strSN As String, ByVal str2Decode As String) As String

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
    
    Dim ByteData() As Byte, lngCount As Long, strFileName As String, lngFileNum
    
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
    
    strFileName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & strSN & "." & strType
    lngFileNum = FreeFile
    Open strFileName For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    SaveBase64ToFile = strFileName

End Function

