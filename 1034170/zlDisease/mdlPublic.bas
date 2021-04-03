Attribute VB_Name = "mdlPublic"
Option Explicit

Public gstrUnitName As String       '当前用户单位名称
Public gfrmMain As Object           '导航台窗体
Public gcnOracle As ADODB.Connection  '数据库连接
Public gstrSysName As String                '系统名称，例如：中联软件
Public gstrProductName As String            '产品简称，例如：中联
Public glngModul As Long                    '模块编号
Public glngSys As Long                      '系统编号，例如：100
Public gstrDBUser As String
Public gstrPrivs As String                     '用户在该模块下面的权限
Public gblnShowInTaskBar As Boolean         '是否显示窗体在任务条上
Public UserInfo As TYPE_USER_INFO            '用户信息
Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Public gMainPrivs As String
Public gstrNodeNo As String          '当前站点编号；如果未设置启用站点，则为"-"
Private mclsZip As New cZip
Private mclsUnzip As New cUnzip
Public gclsMipModule As zl9ComLib.clsMipModule
Public gstrLike As String  '项目匹配方法,%或空
Public gbytCode As Byte '简码输入方式
Public gstrDBOwer As String
Public gobjComlib As zl9ComLib.clsComLib
Public glngPreHWnd As Long '用于支持鼠标滚轮功能
Public glngOpenedID As Long '医生站处理时打开的反馈单ID
Public gObjRichEPR As zlRichEPR.cRichEPR

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'改变窗体位置、Zorder、尺寸等
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = -4
Public Const WM_MOUSEWHEEL = &H20A

Public Type TYPE_USER_INFO
    ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    性质 As String
    部门ID As Long
    部门码 As String
    部门名 As String
    用药级别 As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Type POINTAPI
        X As Long
        Y As Long
End Type

'Public Enum zlEnumDClick
'    cprEmDClickApplyTo = 1         '双击适用科室
'    cprEmDClickRequest = 2         '双击时限要求
'End Enum

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim strTmp As String
    gstrLike = IIf(gobjComlib.zlDatabase.GetPara("输入匹配") = "0", "%", "")
    gbytCode = Val(gobjComlib.zlDatabase.GetPara("简码方式"))
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.用户名 = rsTmp!User
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = NVL(rsTmp!简码)
            UserInfo.姓名 = NVL(rsTmp!姓名)
            UserInfo.部门ID = NVL(rsTmp!部门ID, 0)
            UserInfo.部门码 = NVL(rsTmp!部门码)
            UserInfo.部门名 = NVL(rsTmp!部门名)
            UserInfo.性质 = Get人员性质
            GetUserInfo = True
        End If
    End If
    Exit Function
errH:
   If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get人员性质(Optional ByVal str姓名 As String) As String
'功能：读取当前登录人员或指定人员的人员性质
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    If str姓名 <> "" Then
        strSQL = "Select B.人员性质 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID And A.姓名=[1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str姓名)
    Else
        strSQL = "Select 人员性质 From 人员性质说明 Where 人员ID = [1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get人员性质 = Get人员性质 & "," & rsTmp!人员性质
        rsTmp.MoveNext
    Loop
    Get人员性质 = Mid(Get人员性质, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## 功能：  将数据从一个XtremeReportControl控件复制到VSFlexGrid，以便进行打印
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow

    Dim lngCol As Long, lngRow As Long

    On Error GoTo errHand:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        '标题行复制
        .Row = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = rptCol.Width * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol

        '数据行复制
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function
errHand:
    zlReportToVSFlexGrid = False
End Function
'
Public Function DynamicCreate(ByVal strClass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strClass)
    If Err <> 0 Then
        If blnMsg Then MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function MovedByDate(ByVal vDate As Date) As Boolean
'功能：判断指定日期之前的是否可能已经执行了数据转出
'参数：vDate=时间点或时间段的开始时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "Select 上次日期 From zlDataMove Where 系统=[1] And 组号=1 And 上次日期 is Not Null"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '上次日期没有时点,"<"判断与转出过程中一致
        If vDate < rsTmp!上次日期 Then
            MovedByDate = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetDbOwner(ByVal lngSys As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String

    GetDbOwner = ""
    On Error GoTo errHand
    strSQL = "Select 所有者 From Zlsystems Where 编号 = [1]"
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetDbOwner", lngSys)
    If rsTemp.RecordCount <> 0 Then GetDbOwner = "" & rsTemp!所有者
    rsTemp.Close
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'################################################################################################################
'## 功能：  将指定的LOB字段复制为临时文件
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  存放内容的文件名，失败则返回零长度""
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    Dim strSQL As String

    Err = 0: On Error GoTo errHand

    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum

    strSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value

        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next

        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile
    If ErrCenter = 1 Then
        Resume
    End If
    zlBlobRead = ""
End Function


'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String

    On Error GoTo errHand

    If Not gobjFSO.FileExists(strZipFile) Then zlFileUnzip = "": Exit Function

    strZipPath = gobjFSO.GetSpecialFolder(2) '取临时目录
    strZipPathTmp = strZipPath & "\" & Format(Now, "yyMMdd") & CStr(100 * Timer)
    If Not gobjFSO.FolderExists(strZipPathTmp) Then Call gobjFSO.CreateFolder(strZipPathTmp)

    strZipFileTmp = strZipPathTmp & "\TMP.RTF"
    If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp

    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FileExists(strZipFileTmp) Then

        strZipFileName = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer) & ".RTF"
        If gobjFSO.FileExists(strZipFileName) Then gobjFSO.DeleteFile strZipFileName

        Call gobjFSO.CopyFile(strZipFileTmp, strZipFileName)
        If gobjFSO.FileExists(strZipFileTmp) Then gobjFSO.DeleteFile strZipFileTmp, True
        On Error Resume Next
        If gobjFSO.FolderExists(strZipPathTmp) Then gobjFSO.DeleteFolder strZipPathTmp, True

        zlFileUnzip = strZipFileName
    Else
        zlFileUnzip = ""
    End If
    Exit Function
errHand:
    Call SaveErrLog
End Function


'################################################################################################################
'## 功能：  替换诊治要素的处理
'##
'## 参数：  ElementName     :替换项目的名称
'##         sPatientID      :病人ID
'##         sPageID         :主页ID或挂号id
'##         iPatientType    :0=门诊、1=住院
'##         lng医嘱ID       :医嘱ID
'##
'## 返回：  返回替换结果
'################################################################################################################
Public Function GetReplaceEleValue(ByVal ElementName As String, _
    ByVal sPatientID As String, _
    ByVal sPageID As String, _
    ByVal iPatientType As PatiFromEnum, _
    ByVal lng医嘱id As Long, Optional lngBabyNum As Long) As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset

    strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5],[6]) From Dual"
    Err = 0: On Error GoTo DBError
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "读取替换项", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lng医嘱id, lngBabyNum)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value))
    End If
    Exit Function
DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function


'################################################################################################################
'## 功能：  搜索整个文本给出指定关键字区域的定位信息
'##
'## 参数：  edtThis         :   IN  ，编辑控件
'##         strKeyType      :   IN  ，给定关键字名称。取值为："O"、"P"、"T"、"E"、"U"
'##         lngKey           :   IN  ，给定欲查找的关键字ID号。
'##         lngKSS、lngKSE  :   OUT ，分别表示起始关键字的开始位置和结束位置；
'##         lngKES、lngKEE  :   OUT ，分别表示终止关键字的开始位置和结束位置；
'##         blnNeeded:      :   OUT ，是否是保留对象
'##
'## 返回：  如果找到该关键字具体位置，则返回True，否则返回False
'################################################################################################################
Public Function FindKey(ByRef edtThis As Object, _
        ByRef strKeyType As String, _
        ByRef lngKey As Long, _
        ByRef lngKSS As Long, _
        ByRef lngKSE As Long, _
        ByRef lngKES As Long, _
        ByRef lngKEE As Long, _
        ByRef blnNeeded As Boolean) As Boolean

    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！

    sTMP = strKeyType & "S(" & Format(lngKey, "00000000")
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = 1
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字

            '查找结束关键字
            j = i + 16
LL2:
            sTMP = strKeyType & "E(" & Format(lngKey, "00000000")
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = strKeyType
                lngKSS = i - 1 '转换为0开始的坐标位置。
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindKey = True
            End If
        End If
    End With
End Function


Public Function GetControlRect(ByVal lngHwnd As Long, Optional ByVal blnTwip As Boolean = True) As RECT
'功能：获取指定控件在屏幕中的位置(Twip/Pixel)
'返回：blnTwip=True-返回Twip单位，False-返回像素单位
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    If blnTwip Then
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    End If
    GetControlRect = vRect
End Function

Public Function RPAD(ByVal strText As String, ByVal intCount As Integer, Optional ByVal StrPAD As String = " ", Optional ByVal blnAutoSub As Boolean) As String
'功能：等同Oracle的RPAD函数
'功能:按指定长度填制空格
 '参数：
 '       strText:填充字符串
 '       intCount:填充后的长度
 '       StrPAD:填充的字符
 '       blnAutoSub:字符串超长后自动截取
'返回:返回字串
   
    Dim lngTmp As Long, lngFill As Long
    If StrPAD = "" Then
        StrPAD = " "
    Else
        StrPAD = Mid(StrPAD, 1, 1)
    End If
    
    lngFill = LenB(StrConv(StrPAD, vbFromUnicode))
    lngTmp = LenB(StrConv(strText, vbFromUnicode))
    If lngTmp <= intCount - lngFill Then
        RPAD = strText & String((intCount - lngTmp) \ lngFill, StrPAD)
    ElseIf lngTmp > intCount And blnAutoSub Then
        RPAD = SubB(strText, 1, intCount)
    Else
        RPAD = strText
    End If
End Function

Public Function SubB(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
'功能:读取指定字串的值,字串中可以包含汉字
 '入参:strInfor-原串
 '         lngStart-直始位置
'         lngLen-长度
'返回:子串
    Dim strTmp As String, i As Long
    Err = 0: On Error GoTo errH:
    SubB = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    SubB = Replace(SubB, Chr(0), "")
    Exit Function
errH:
    Err.Clear
    SubB = ""
End Function


Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun存在该函数
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function


Public Sub FormLock(Optional ByVal lngHwnd As Long)
'功能：锁定对象区域不刷新。或者解除锁定
'参数：lngHwnd=0-解除锁定,<>0要锁定对象的Hwnd
    LockWindowUpdate (lngHwnd)
End Sub


Public Function RowValue(ByVal strTable As String, Optional ByVal arrValues As Variant, Optional ByVal strGetFields As String, Optional ByVal strWhereField As String = "ID") As Variant
'功能：获取指定表指定字段信息
'参数：strTable=读取数据的表
'          arrValues=过滤值，可以传数组，也可以传单个值，也可以不传，不传读取全表
'          strGetField=获取的字段,多个字段以逗号分割，同SQL书写获取字段一致
'          strWhereField=过滤的字段，该字段为简单的数值或字符类型或日期类型，其余类型无法支持
'返回：
'只有以下一种情况返回特定的一个值（未处理NULL值）：
'      strGetField=单个字段
'      arrValues=为单个值，或不高于一个元素的数组
'其余情况均返回记录集

    Dim rsTmp As New ADODB.Recordset, blnReturnRec As Boolean
    Dim strSQL As String
    Dim strWhere As String
    Dim arrPars As Variant
    Dim i As Long, strPars As String
    
    On Error GoTo errH
    blnReturnRec = True
    If TypeName(arrValues) = "Variant()" Then
        arrPars = arrValues
        For i = LBound(arrValues) To UBound(arrValues)
            strPars = strPars & ",[" & i + 1 & "]"
        Next
        If strGetFields <> "" Then '数组顾虑元素不超过一个,且获取单个元素，则不返回记录集
            If UBound(arrValues) - LBound(arrValues) + 1 <= 1 And Not strGetFields Like "*,*" Then blnReturnRec = False
        End If
        If strPars <> "" Then
            strWhere = " Where " & strWhereField & " In (" & strPars & ")"
        End If
    ElseIf TypeName(arrValues) <> "Error" Then
        '单个值时，若获取单个字段，则不反悔记录集
         If strGetFields <> "" And Not strGetFields Like "*,*" Then blnReturnRec = False
        arrPars = Array(arrValues)
        strWhere = " Where " & strWhereField & "=[1]"
    Else
        strWhere = ""
    End If
    
    If strGetFields = "" Then strGetFields = "*"
    strSQL = "Select " & strGetFields & " From " & strTable & strWhere
    If strWhere <> "" Then
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "RowValue", arrPars)
    Else
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "RowValue")
    End If
    If blnReturnRec Then
        Set RowValue = rsTmp
    Else
        If Not rsTmp.EOF Then
            RowValue = rsTmp.Fields(strGetFields).Value
        Else '获取单值时，没有获取到数值，则返回默认值
            If IsType(rsTmp.Fields(strGetFields).Type, adVarChar) Then
                RowValue = ""
            ElseIf IsType(rsTmp.Fields(strGetFields).Type, adInteger) Then
                RowValue = 0
            Else
                RowValue = Null
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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

Public Function GetDeptID(ByVal strDeptCode As String) As Long
'功能：根据部门编码获取部门ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
On Error GoTo errH
    strSQL = "Select a.Id, a.名称 From 部门表 A Where A.编码 = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "部门查询", strDeptCode)
    If rsTmp.RecordCount > 0 Then
        GetDeptID = rsTmp!ID
    Else
        GetDeptID = 0
        MsgBox "没有查询到编码为“" & strDeptCode & "”的部门科室，请联系管理员对码！", vbInformation, gstrSysName
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'功能：支持滚轮的滚动
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '向下滚
            gobjComlib.ZLCommFun.PressKey vbKeyPageDown
        Case 7864320   '向上滚
            gobjComlib.ZLCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(glngPreHWnd, hwnd, wMsg, wParam, lParam)
End Function

Public Function CheckOperateState(ByVal lngID As Long, ByRef intCode As Integer) As Boolean
'功能: 查询是否能够处理该反馈单（删除或者修改）
'参数: lngID - 反馈单ID ；intCode - 不能操作的原因 ；1-未查找到；2-他人的反馈单；3-医生已经处理
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '读取反馈单相关信息
    On Error GoTo errH
    strSQL = "Select a.Id, a.记录状态, a.登记人 From 疾病阳性记录 A  Where a.Id = [1] "

    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询阳性结果反馈单", lngID)
    
    If rsTmp.RecordCount > 0 Then
        If UserInfo.姓名 <> NVL(rsTmp!登记人) Then
            intCode = 2
            Exit Function
        ElseIf rsTmp!记录状态 > 1 Then
            intCode = 3
            Exit Function
        End If
    Else
        intCode = 1
        Exit Function
    End If
    CheckOperateState = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub PrintDiseaseRegist(ByVal intType As Integer, ByVal lngID As Long, ByRef frmParent As Object)
'功能: 打印阳性结果反馈单
'参数：lngID : 反馈单ID；intType:1-预览，2-打印
    Dim objReport As clsReport
    
    On Error GoTo errH
  
    If objReport Is Nothing Then Set objReport = New clsReport
    Call objReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1278_1", frmParent, "反馈单ID=" & lngID, intType)
    If Not objReport Is Nothing Then Set objReport = Nothing
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckDisNum(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByRef rsDisease As ADODB.Recordset, Optional ByVal lngID As Long) As Boolean
'功能: 检查该病人有多少没有填写报告卡的反馈单
'lngPatFrom: 2-住院, 1-门诊
    Dim strSQL As String
    On Error GoTo errH
    If lngID <> 0 Then
        strSQL = " and a.ID = " & lngID
    End If
    If lngPatFrom = 1 Then
        strSQL = "select rowNum as NO,a.ID,c.名称 as 科室, a.登记时间, a.记录状态, a.处理情况说明 from  疾病阳性记录 A ,病人挂号记录 B ,部门表 C where A.文件ID is NULL  and A.挂号单 = B.NO and A.病人ID = B.病人ID and A.记录状态 <> 3 and A.登记科室ID = C.ID  and A.病人ID = [1] and B.ID = [2]" & strSQL
    ElseIf lngPatFrom = 2 Then
        strSQL = "select rowNum as NO,a.ID ,c.名称 as 科室,a.登记时间, a.记录状态, a.处理情况说明 from  疾病阳性记录 A ,部门表 C  where A.文件ID is NULL  and A.记录状态 <> 3  and A.登记科室ID = C.ID and A.病人ID = [1] and A.主页ID = [2] " & strSQL
    End If
    Set rsDisease = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "查询阳性结果反馈单", lngPatiID, lngPageId)
    
    If rsDisease.RecordCount > 0 Then
        CheckDisNum = True
    Else
        CheckDisNum = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SaveReason(ByVal strReason As String, ByVal lngID As Long, ByVal lngState As Long) As Boolean
'功能: 存储不填写报告卡的原因
'参数：strReason-原因；lngID-反馈单ID ；lngState-反馈单当前的记录状态
    Dim strSQL As String
    Dim str处理时间 As String
    Dim str处理医生 As String
    Dim str处理情况 As String, strTmp As String

    On Error GoTo errH
    str处理时间 = "to_date('" & Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
    str处理医生 = "'" & UserInfo.姓名 & "'"
    str处理情况 = "'" & strReason & "'"
    lngState = IIf(lngState = 1, 2, lngState)

    strSQL = "Zl_疾病阳性检测记录_update(1," & lngID & "," & "NULL" & "," & CStr(lngState) & "," & str处理医生 & "," & str处理时间 & "," & str处理情况 & ")"
    Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, "保存反馈单的处理情况")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
