Attribute VB_Name = "mdlServer"
Option Explicit
Public Const G_NOT_ZIP_FILES = ";ZLHISCRUSTCOM.DLL;ZLHISCRUST.EXE;7Z.EXE;7Z.DLL;AAMD532.DLL;ZLRUNAS.EXE;REGCOM.DLL;GACUTIL.EXE;GACUTIL.EXE.CONFIG;ZL7Z.DLL;"
'密码加密程序
Public Function Cipher(ByVal strText As String) As String
    Const MIN_ASC = 32    '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intLen As Integer, intSeedLen As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    Dim strSeed As String
    
    If strText = "" Then Exit Function
    '获取随机种子
    '随机种子的随机数为999
    Rnd (-1)
    Randomize (999)
    strSeed = "456"
    intSeedLen = Len(strSeed)
    strDeText = Chr(intSeedLen + MIN_ASC)
    For i = 1 To intSeedLen
        intChr = Asc(Mid(strSeed, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr + lngOffset) Mod NUM_ASC)
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        End If
    Next
    Rnd (-1)
    Randomize (Val(strSeed))
    intLen = Len(strText)
    For i = 1 To intLen
        intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr + lngOffset) Mod NUM_ASC)
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        ElseIf intChr < 0 Then '非ASCII字符的处理,如中文，中文不加密
            strDeText = strDeText & Mid(strText, i, 1)
        End If
    Next
    Cipher = strDeText
End Function

Public Function Decipher(ByVal strText As String) As String
'密码解密程序
    Const MIN_ASC = 32    '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intLen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '随机种子长度
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intLen = Len(strText)
    '采用旧的随机算法
    If intSeedLen > 0 And intSeedLen < intLen - 3 And intSeedLen < 5 Then
        '获取随机种子
        '随机种子的随机数为999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
            If intChr >= MIN_ASC And intChr <= MAX_ASC Then
                intChr = intChr - MIN_ASC
                lngOffset = Int((NUM_ASC + 1) * Rnd())
                intChr = ((intChr - lngOffset) Mod NUM_ASC)
                If intChr < 0 Then
                    intChr = intChr + NUM_ASC
                End If
                intChr = intChr + MIN_ASC
                strDeText = strDeText & Chr(intChr)
            End If
        Next
        If Not IsNumeric(strDeText) Then
            strDeText = "123"
            intStart = 1
        Else
            intStart = 2 + intSeedLen
        End If
    Else
        strDeText = "123"
        intStart = 1
    End If
        
    '内容解密的种子
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intLen
        intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr - lngOffset) Mod NUM_ASC)
            If intChr < 0 Then
                intChr = intChr + NUM_ASC
            End If
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        Else '非ASCII字符的处理,如中文，中文不加密
            strDeText = strDeText & Mid(strText, i, 1)
        End If
    Next
    Decipher = strDeText
End Function
