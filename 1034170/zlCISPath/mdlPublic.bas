Attribute VB_Name = "mdlPublic"
Option Explicit
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long

Public Const ETO_OPAQUE = 2
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

'Shell------------------------------------------------
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function OpenAs_RunDLL Lib "shell32.dll" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpFile As String, ByVal nShowCmd As Long) As Long

'Icon------------------------------------------------
Public Enum ICON_SIZE
    ICON_SMALL = 16
    ICON_LARGE = 32
End Enum
Private Const MAX_PATH = 260
Private Const DI_NORMAL = &H3
Private Const SHGFI_LARGEICON = &H0                      '获取大图标
Private Const SHGFI_SMALLICON = &H1                      '获取小图标
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_ICON = &H100                         '获取图标句柄
Private Type SHFILEINFO
        hIcon As Long                       '  out: icon
        iIcon As Long                       '  out: icon index
        dwAttributes As Long                '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH  '  out: display name (or path)
        szTypeName As String * 80           '  out: type name
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Declare Function GetScrollRange Lib "user32" (ByVal Hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
    Public Const SB_HORZ = &H0
    Public Const SB_VERT = &H1

'目录选择对话框函数=================================================================
Private mstrAPIPath As String

Private Type BrowseInfo
  hWndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As String
  lpszTitle      As String
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type

Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'===================================================================================

Public Function VsScroll(vsf As VSFlexGrid) As Boolean '判断水平滚动条的可见性
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    VsScroll = False
    i = GetScrollRange(vsf.Hwnd, SB_HORZ, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos Then VsScroll = True
    End Function
    
Public Function HeScroll(vsf As VSFlexGrid) As Boolean '判断垂直滚动条的可见性
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    HeScroll = False
    i = GetScrollRange(vsf.Hwnd, SB_VERT, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos Then HeScroll = True
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'功能：将VB的系统颜色转换为RGB色
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

Public Function GetFileIcon(ByVal strFile As String, ByVal intSize As ICON_SIZE, objPicTmp As PictureBox) As StdPicture
'功能：返回指定文件的大图标或小图标
'参数：strFile，包含后缀的文件名；当文件真实文件时，应该包含完整的路径名
'      intSize，图标的大小
'      objPicTmp：临时用PictureBox控件，无边框，AutoRedraw = True
    Dim fInfo As SHFILEINFO
    Dim objText As TextStream
    Dim blnExists As Boolean, lngRetu As Long
    
    blnExists = gobjFile.FileExists(strFile)
    If Not blnExists Then
        strFile = gobjFile.GetSpecialFolder(TemporaryFolder).Path & "\" & App.hInstance & Int(Timer) & strFile
        Set objText = gobjFile.CreateTextFile(strFile)
    End If

    If intSize = ICON_LARGE Then
        lngRetu = SHGetFileInfo(strFile, 0, fInfo, Len(fInfo), SHGFI_SHELLICONSIZE Or SHGFI_ICON Or SHGFI_LARGEICON)
    Else
        lngRetu = SHGetFileInfo(strFile, 0, fInfo, Len(fInfo), SHGFI_SHELLICONSIZE Or SHGFI_ICON Or SHGFI_SMALLICON)
    End If
    If Not blnExists Then
        objText.Close
        gobjFile.DeleteFile strFile, True
    End If
    
    objPicTmp.Width = intSize * Screen.TwipsPerPixelX
    objPicTmp.Height = intSize * Screen.TwipsPerPixelY
    objPicTmp.Cls
    If lngRetu <> 0 Then
        DrawIconEx objPicTmp.hDC, 0, 0, fInfo.hIcon, intSize, intSize, 0, 0, DI_NORMAL
        DestroyIcon fInfo.hIcon
    End If
    Set GetFileIcon = objPicTmp.Image
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
'参数：blnForceNum=当为Null时，是否强制表示为数字型
    ZVal = IIf(Val(varValue) = 0, IIf(blnForceNum, "-NULL", "NULL"), Val(varValue))
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function Between(X, a, b) As Boolean
'功能：判断x是否在a和b之间
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Public Function SeekCboIndex(objCbo As Object, varData As Variant) As Integer
'功能：由ItemData或Text查找ComboBox的索引值
    Dim strType As String, i As Integer
    
    SeekCboIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If IsType(varData.Type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '先精确查找
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    SeekCboIndex = i: Exit Function
                ElseIf NeedName(objCbo.List(i)) = varData And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
            '再模糊查找
            For i = 0 To objCbo.ListCount - 1
                If InStr(objCbo.List(i), varData) > 0 And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    Else
        If varData <> 0 Then
            For i = 0 To objCbo.ListCount - 1
                If objCbo.ItemData(i) = varData Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    End If
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'功能：判断某个ADO字段数据类型是否与指定字段类型是同一类(如数字,日期,字符,二进制)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    
    IsType = intA = intB
End Function

Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Public Function BrowseForFolder(ByVal Hwnd As Long, ByVal Title As String, ByVal InitDir As String) As String
    Dim lpIDList As Long
    Dim szTitle As String
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    
    mstrAPIPath = InitDir & Chr(0)
    
    szTitle = Title
    
    With tBrowseInfo
        .hWndOwner = Hwnd
        .lpszTitle = szTitle
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_STATUSTEXT
        .lpfnCallback = AddressOfFunction(AddressOf BrowseCallbackProc)
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If lpIDList <> 0 Then
        sBuffer = Space(512)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
        BrowseForFolder = sBuffer
    End If
End Function

Private Function BrowseCallbackProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
  
    On Error Resume Next
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(Hwnd, BFFM_SETSELECTION, 1, ByVal mstrAPIPath)
        Case BFFM_SELCHANGED
            sBuffer = Space(512)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                Call SendMessage(Hwnd, BFFM_SETSTATUSTEXT, 0, ByVal sBuffer)
            End If
    End Select
    
    BrowseCallbackProc = 0
End Function

Private Function AddressOfFunction(Address As Long) As Long
    AddressOfFunction = Address
End Function



Public Function GetColumnLength(strTable As String, strColumn As String) As Long
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select Nvl(Data_Precision, Data_Length) collen From All_Tab_Columns Where Table_Name = [1] And Column_Name = [2]"
    'strSQL = "Select " & strColumn & " From " & strTable & " Where Rownum<1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, strTable, strColumn)
    GetColumnLength = Val("" & rsTmp!collen)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function
