VERSION 5.00
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.Form frmReport 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10845
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   ScaleHeight     =   10845
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin zlSubclass.Subclass Subclass1 
      Left            =   900
      Top             =   3495
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.HScrollBar hsbReport 
      Height          =   255
      LargeChange     =   500
      Left            =   0
      Max             =   100
      SmallChange     =   10
      TabIndex        =   4
      Top             =   0
      Width           =   8535
   End
   Begin VB.VScrollBar vsbReport 
      Height          =   7335
      LargeChange     =   50
      Left            =   0
      Max             =   100
      SmallChange     =   10
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   16845
      Left            =   1785
      ScaleHeight     =   16815
      ScaleWidth      =   11865
      TabIndex        =   6
      Top             =   45
      Width           =   11895
      Begin zlDisReportCard.PaneFour PaneFour 
         Height          =   4515
         Left            =   1050
         TabIndex        =   3
         Top             =   11295
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   7964
      End
      Begin zlDisReportCard.PaneThree PaneThree 
         Height          =   4260
         Left            =   1050
         TabIndex        =   2
         Top             =   7020
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   7514
      End
      Begin zlDisReportCard.PaneOne PaneOne 
         Height          =   1065
         Left            =   1020
         TabIndex        =   0
         Top             =   1005
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   1879
      End
      Begin zlDisReportCard.PaneTwo PaneTwo 
         Height          =   4848
         Left            =   1050
         TabIndex        =   1
         Top             =   2145
         Width           =   9825
         _ExtentX        =   17383
         _ExtentY        =   8705
      End
      Begin VB.Line Line2 
         X1              =   1065
         X2              =   10875
         Y1              =   11280
         Y2              =   11280
      End
      Begin VB.Line Line1 
         X1              =   1050
         X2              =   10875
         Y1              =   7005
         Y2              =   7005
      End
      Begin VB.Shape Shape1 
         Height          =   13935
         Left            =   1035
         Top             =   2110
         Width           =   9855
      End
   End
   Begin VB.PictureBox picShadow 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   750
      ScaleHeight     =   1770
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   660
      Width           =   1140
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private marrSql() As Variant        '保存数据时候执行的SQL
Private mColCls As New Collection   '需要保存到数据库的数据
Private mColData As New Collection  '保存从数据库读取到的数据
Public Event HaveSavedSQL()     '执行保存SQL时触发,每执行一条出发一次
Public blnHaveStatus As Boolean  '是否存在状态栏
Private blnFirstGot As Boolean  '第一次获得焦点

Private mlngPatiID As Long '病人id
Private mlngPageID As Long '主页ID（门诊传挂号ID）
Private mbytType As Byte   '编辑方式0-新增　1-修改，用于区别提取数据
Private mbytFrom As Byte   '病人来源1-门诊 2-住院
Private mlngDeptID As Long '当前科室ID
Private mlngFileId As Long   '文件ID,来源于电子病历记录.ID
Private mbytBabyNo As Long '婴儿ID
Private mbln身份证必填 As Boolean '身份证信息必填 参数：传染病报告身份证号码必填

Private mstrChkType As String '数据格式是："[男][艾滋病][AIDS][...]......"

Private Type POINTAPI
        x As Long
        y As Long
End Type

Public Sub SetMyFocus()
    If picReport.Enabled = True Then
        Call picReport.SetFocus
    End If
End Sub

Public Function HaveChanged() As Boolean
'功能：判断四个自定义控件里面的显示信息是否发生改变
    HaveChanged = False
    If PaneOne.HaveChanged = True Then
        HaveChanged = True
    ElseIf PaneTwo.HaveChanged = True Then
        HaveChanged = True
    ElseIf PaneThree.HaveChanged = True Then
        HaveChanged = True
    ElseIf PaneFour.HaveChanged = True Then
        HaveChanged = True
    End If
    
End Function
Public Sub CanWrite()
'功能：是界面可以编辑
    picReport.Enabled = True
End Sub
Public Sub PrintReport(ByVal frmParent As Object, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal strPrintDeviceName As String)
'功能：打印报告
    Dim strSql As String
    Dim strPos As String
    Dim strPosInfo() As String
    Dim strPosTmp() As String
    Dim i As Integer
    
    On Error GoTo errHand
    
    Call zlRefresh(lngPatiID, lngPageID, lngFileId, False)

    If Trim(strPrintDeviceName) <> "" Then
        For i = 0 To Printers.Count - 1
            If Trim(Printers(i).DeviceName) = Trim(strPrintDeviceName) Then
                Set Printer = Printers(i)
                Exit For
            End If
            If i = Printers.Count - 1 Then
                MsgBox "没有找到相应的打印机，请核对打印机名称！", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            End If
        Next
    End If
    Printer.PaperSize = vbPRPSA4 'A4纸
    Printer.ScaleMode = vbPixels

    glngOffsetX = -GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX) '可打印左边缘
    glngOffsetY = -GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY) '可打印上边缘

    Call PaneOne.PrintOne
    Call PaneTwo.PrintTwo
    Call PaneThree.PrintThree
    Call PaneFour.PrintFour

    strPos = "69,142,725,142|69,142,69,1069|69,1069,725,1069|725,142,725,1069|" & _
             "69,466,725,466|69,514,725,514|69,678,725,678|69,749,725,749|" & _
             "69,793,725,793|69,934,725,934|69,1025,725,1025"
    
    strPosInfo = Split(strPos, "|")
    For i = 0 To UBound(strPosInfo)
        strPosTmp = (Split(strPosInfo(i), ","))
        Printer.Line (glngOffsetX + PScaleX(val(strPosTmp(0))), glngOffsetY + PScaleY(val(strPosTmp(1))))-(glngOffsetX + PScaleX(val(strPosTmp(2))), glngOffsetY + PScaleY(val(strPosTmp(3)))), &H0&, B
    Next
    
    Printer.EndDoc
    
    strSql = "Zl_电子病历打印_Insert(" & mlngFileId & ",20," & mlngPatiID & "," & mlngPageID & ",'" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "")
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub zlRefresh(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal blnMoved As Boolean)
    mlngPatiID = lngPatiID
    mlngPageID = lngPageID
    mlngFileId = lngFileId
 
    Call PaneOne.ClearMe
    Call PaneTwo.ClearMe
    Call PaneThree.ClearMe
    Call PaneFour.ClearMe
    Call InitReport(mbytType, mlngPatiID, mlngPageID, mbytFrom, 0, mlngDeptID, mlngFileId)
    If lngPatiID <> 0 Then
        Call LoadData(1, blnMoved)
    End If
End Sub

Public Sub InitReport(ByVal bytType As Byte, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal bytFrom As Byte, ByVal bytBabyNo As Byte, ByVal lngDeptID As Long, ByVal lngFileId As Long)
    mbytType = bytType
    mlngPatiID = lngPatiID
    mlngPageID = lngPageID
    mbytFrom = bytFrom
    mlngDeptID = lngDeptID
    mlngFileId = lngFileId
    mbytBabyNo = bytBabyNo
End Sub

Public Function SaveData(ByVal blnFinish As Boolean) As Boolean
    Dim i As Integer
    Dim strSql As String
    Dim blnBegin As Boolean
    Dim SLevel As SignLevel
    Dim lngFileId As Long       '文件ID 来源于病历文件列表
    Dim strFileName As String   '文件名称 来源于病历文件列表
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHand
    
    SaveData = False
    
    '新增需要提取新的文件ID
    If mbytType = 0 Then
        mlngFileId = zlDatabase.GetNextId("电子病历记录")
        mbytType = 1
    End If
    
    SLevel = GetUserSignLevel(UserInfo.ID, mlngPatiID, mlngPageID)
    
    strSql = "select t.id,t.名称 from 病历文件列表 t where t.种类=5 and t.编号='000'"
    Set rsTemp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTemp, strSql, "")
    lngFileId = Nvl(rsTemp!ID, 0)
    strFileName = Nvl(rsTemp!名称, "")
    strSql = "Zl_传染病报告卡记录_Update(" & mlngFileId & "," & mbytFrom & "," & mlngPatiID & "," & _
              mlngPageID & "," & mlngDeptID & ",'" & UserInfo.姓名 & "'," & lngFileId & ",'" & strFileName & _
               "','" & UserInfo.姓名 & "'," & IIf(blnFinish, 1, 0) & "," & IIf(blnFinish, SLevel, "Null") & "," & mbytBabyNo & ")"
    
    Call MakeSaveSql(marrSql, mColCls, mlngFileId)

    gcnOracle.BeginTrans
    blnBegin = True
    Call zlDatabase.ExecuteProcedure(strSql, "")
    For i = 0 To mColCls.Count - 1
        Call zlDatabase.ExecuteProcedure(CStr(marrSql(i)), "")
        RaiseEvent HaveSavedSQL
    Next
    gcnOracle.CommitTrans
    blnBegin = False
    SaveData = True
    If blnFinish = True Then
        picReport.Enabled = False
    End If
    Exit Function
errHand:
    If blnBegin Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Sub LoadData(ByVal bytType As Byte, Optional blnMoved As Boolean)
    Dim strSql As String
    Dim strKey As String
    Dim strNo As String
    Dim strID As String
    Dim strTmp As String
    Dim strInfo() As String
    Dim objCls As clsReport
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHand
    Set mColCls = New Collection
    mstrChkType = ""
    If bytType = 1 Then
        
        Set mColData = New Collection
        strSql = "select t.id,t.对象序号,t.内容文本,t.要素名称 from 电子病历内容 t where t.文件id=[1]"
        If blnMoved = True Then
            strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "电子病历内容", mlngFileId)
        
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp.EOF = False Then
                strID = Nvl(rsTemp!ID)
                strNo = Nvl(rsTemp!对象序号)
                strTmp = Nvl(rsTemp!内容文本)
                strKey = "K" & Trim(strNo)
                mColData.Add strTmp, strKey
                If InStr(GSTR_OBJNO, "," & strNo & ",") > 0 Then
                    mstrChkType = mstrChkType & "[" & strNo & "," & Trim(strTmp) & "]"
                End If
                Set objCls = New clsReport
                objCls.ID = strID
                mColCls.Add objCls, strKey
                rsTemp.MoveNext
            End If
        Next
        
    ElseIf bytType = 0 Then
        For i = 1 To 44
            Set objCls = New clsReport
            strKey = "K" & i
            objCls.ID = 0
            mColCls.Add objCls, strKey
        Next
        Set mColData = New Collection
        strTmp = "姓名|身份证号|性别|出生日期|年龄|工作单位|联系人电话|家庭电话|单位电话|婚姻状况|学历|单位名称|当前日期|家庭地址"
        strInfo = Split(strTmp, "|")
        
        For i = 0 To UBound(strInfo)
            If mbytBabyNo <> 0 And Trim(strInfo(i)) = "姓名" Then
                strSql = "select Zl_Replace_Element_Value([1],[2],[3],[4],null,[5]) as 信息 from dual"
            Else
                strSql = "select Zl_Replace_Element_Value([1],[2],[3],[4]) as 信息 from dual"
            End If
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "数据读取", strInfo(i), mlngPatiID, mlngPageID, mbytFrom, mbytBabyNo)
            strNo = i
            strTmp = Nvl(rsTemp!信息)
            mColData.Add strTmp, "K" & Trim(strNo)
        Next
        '家长姓名
        If mbytBabyNo <> 0 Then
            strSql = "select Zl_Replace_Element_Value([1],[2],[3],[4],null,[5]) as 信息 from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "数据读取", "家长姓名", mlngPatiID, mlngPageID, mbytFrom, mbytBabyNo)
            strTmp = Nvl(rsTemp!信息)
            mColData.Add strTmp, "KParent"
        Else
            mColData.Add "", "KParent"
        End If
        '发病日期
        If mbytFrom = 1 Then
            strSql = "select t.登记时间 as 发病日期 from 病人挂号记录 t where t.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "数据读取", mlngPageID)
        Else
            strSql = "select t.入院日期 as 发病日期 from 病案主页 t where t.病人id=[1] and t.主页id=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "数据读取", mlngPatiID, mlngPageID)
        End If
        
        If rsTemp.RecordCount <> 0 Then
            mColData.Add Format(Nvl(rsTemp!发病日期), "yyyy-mm-dd"), "K14"
        Else
            mColData.Add "--", "K14"
        End If
        '诊断日期
        strSql = "select decode(t.发病时间,null,t.记录日期,t.发病时间) as 诊断日期 from 病人诊断记录 t where t.病人id=[1] and t.主页id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "数据读取", mlngPatiID, mlngPageID)
    
        If rsTemp.RecordCount <> 0 Then
            mColData.Add Format(Nvl(rsTemp!诊断日期), "yyyy-mm-dd-hh"), "K15"
        Else
            mColData.Add "---", "K15"
        End If
        '死亡日期
        strSql = " Select a.开始执行时间 as 死亡日期 " & _
                 " From 病人医嘱记录 A, 诊疗项目目录 B " & _
                 " Where a.诊疗项目id = b.Id And b.类别 = 'Z' And " & _
                 " b.操作类型 = '11'  And a.病人来源 = [1] And a.病人id=[2] and a.主页id=[3] "
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "数据读取", mbytFrom, mlngPatiID, mlngPageID)
        If rsTemp.RecordCount <> 0 Then
            mColData.Add Format(Nvl(rsTemp!死亡日期), "yyyy-mm-dd"), "K17"
        Else
            mColData.Add "--", "K17"
        End If
        '病种
        strSql = "Select a.Id, b.文件id, b.报告病种, a.病人id, a.主页id, a.医嘱id, a.诊断类型, a.疾病id, a.诊断id" & _
                 " From 病人诊断记录 A, 疾病报告前提 B " & _
                 " Where (a.疾病id = b.疾病id Or " & _
                 " a.诊断id = b.诊断id Or " & _
                 " b.诊断id = (Select c.诊断id From 疾病诊断对照 c Where c.疾病id =a.疾病id) or " & _
                 " b.疾病id = (select d.疾病id from 疾病诊断对照 d where d.诊断id=a.诊断id)) And " & _
                 " b.文件id =(select e.id from 病历文件列表 e where e.种类=5  and e.名称='中华人民共和国传染病报告卡' and e.保留=4 ) and " & _
                 " a.记录来源=3 and a.病人id=[1] and a.主页id=[2]"
        
        strTmp = ""
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "数据读取", mlngPatiID, mlngPageID)
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp.EOF = False Then
                strTmp = strTmp & Nvl(rsTemp!报告病种) & "|"
                rsTemp.MoveNext
            End If
        Next
        
        mColData.Add strTmp, "K16"
    End If
    '修改时候加载数据是44条如果少于44条说明病历文件不完整
    '新增时候加载数据是19条，如果少于19条说明信息来源破坏
    If (mColData.Count = 44 And bytType = 1) Or (mColData.Count = 19 And bytType = 0) Then
        Call PaneOne.LoadData(mColData, bytType, mstrChkType)
        Call PaneTwo.LoadData(mColData, bytType, mstrChkType)
        Call PaneThree.LoadData(mColData, bytType, mstrChkType)
        Call PaneFour.LoadData(mColData, bytType, mstrChkType)
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub SetCaption身份证()
    mbln身份证必填 = val(zlDatabase.GetPara("传染病报告身份证号码必填", glngSys, 1277, 0)) = 1
    Call PaneTwo.SetCaption身份证(mbln身份证必填)
End Sub

Private Sub Form_Load()
        
    blnFirstGot = True
    gbytDiseaseType = 5
    gbytAcute = 2
    
    picReport.ScaleHeight = Me.ScaleY(297, 6, 3)
    picReport.ScaleWidth = Me.ScaleX(210, 6, 3)
    picReport.Top = Me.ScaleTop + 200
    marrSql = Array()
    Subclass1.hWnd = Me.hWnd
    Subclass1.Messages(WM_MOUSEWHEEL) = True
    mbln身份证必填 = val(zlDatabase.GetPara("传染病报告身份证号码必填", glngSys, 1277, 0)) = 1
    Call PaneTwo.SetCaption身份证(mbln身份证必填)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picReport.Left = Me.ScaleLeft + (Me.ScaleWidth / 2) - (picReport.Width / 2)
    
    If Me.ScaleWidth < picReport.Width Then
        hsbReport.Visible = True
    Else
        hsbReport.Visible = False
    End If
    
    vsbReport.Top = Me.ScaleTop
    vsbReport.Left = Me.ScaleLeft + Me.ScaleWidth - vsbReport.Width
    vsbReport.Height = Me.ScaleHeight - IIf(hsbReport.Visible = True, hsbReport.Height, 0) - IIf(blnHaveStatus = True, 375, 0)
    vsbReport.LargeChange = 100 / ((picReport.Height + 800) / Me.ScaleHeight)
    vsbReport.SmallChange = vsbReport.LargeChange
    
    hsbReport.Top = vsbReport.Top + vsbReport.Height
    hsbReport.Left = Me.ScaleLeft
    hsbReport.Width = Me.ScaleLeft + Me.ScaleWidth
    hsbReport.LargeChange = 100 / (picReport.Width / Me.ScaleWidth)
    hsbReport.SmallChange = hsbReport.LargeChange
    
    picShadow.Move picReport.Left + 50, picReport.Top + 50, picReport.Width, picReport.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mColCls = Nothing
    Set mColData = Nothing
    Erase marrSql
    gstrKey = ""
    Subclass1.Messages(WM_MOUSEWHEEL) = False
End Sub

Private Sub hsbReport_Change()
    picReport.Left = -((picReport.Width - Me.Width) * (hsbReport.Value / 100))
    picShadow.Left = picReport.Left + 50
End Sub

Private Sub PaneTwo_ClickPositives(blnSelected As Boolean)
    Call PaneThree.SetAIDS(blnSelected)
End Sub

Private Sub picReport_GotFocus()
    If blnFirstGot = True And picReport.Enabled = True Then
        Call PaneOne.SetMyFoucs
    End If
    blnFirstGot = False
End Sub

Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, lParam As Long, result As Long)
    '自定义的消息处理函数
    Dim tP As POINTAPI
    Dim sngX As Single, sngY As Single   '鼠标坐标
    Dim intShift As Integer              '鼠标按键
    Dim bWay As Boolean                  '鼠标方向
    Dim bMouseFlag As Boolean            '鼠标事件激活标志
    Dim wzDelta, wKeys As Integer
    Select Case Msg
        Case WM_MOUSEWHEEL   '滚动
            wzDelta = HIWORD(wParam)
            If wzDelta > 0 Then
                vsbReport.Value = IIf(vsbReport.Value > 10, vsbReport.Value - 10, 0)
            Else
                vsbReport.Value = IIf(vsbReport.Value < 90, vsbReport.Value + 10, 100)
            End If
    End Select
End Sub

Private Sub vsbReport_Change()
    picReport.Top = 200 - ((picReport.Height + 800 - Me.Height) * (vsbReport.Value / 100))
    picShadow.Top = picReport.Top + 50
End Sub

Public Function MakeSaveSql(arrSql() As Variant, colCls As Collection, ByVal strFileId As String) As Boolean
    Call PaneOne.MakeSaveSql(arrSql, colCls, strFileId)
    Call PaneTwo.MakeSaveSql(arrSql, colCls, strFileId)
    Call PaneThree.MakeSaveSql(arrSql, colCls, strFileId)
    Call PaneFour.MakeSaveSql(arrSql, colCls, strFileId)
End Function

Public Sub ClearEnterInfo()
    Call PaneFour.ClearEnterInfo
End Sub
Public Sub SetEnterInfo()
    Dim strDate As String
    If mColData.Count < 44 Then
        strDate = Trim(CStr(mColData("K12")))
    Else
        strDate = Trim(CStr(mColData("K43")))
    End If
    If strDate = "" Or strDate = "--" Then
        strDate = zlDatabase.Currentdate
    End If
    Call PaneFour.SetEnterInfo(UserInfo.姓名, strDate)
End Sub
Public Function CheckValidity() As Boolean
    Dim strMsg As String
    Dim strTmp As String
    Dim strMsgInfo() As String
    Dim i As Integer
    On Error GoTo errHand
    
    strMsg = ""
    strTmp = ""
    Call PaneTwo.CheckValidity(strMsg)
    Call PaneThree.CheckValidity(strMsg)
    Call PaneFour.CheckValidity(strMsg)
    If Trim(strMsg) = "" Then
        CheckValidity = True
    Else
        strMsgInfo = Split(strMsg, "$")
        For i = 0 To UBound(strMsgInfo) - 1
            strTmp = strTmp & i + 1 & ". " & strMsgInfo(i) & vbCrLf
        Next
        Call ShowMsg(strTmp)
        CheckValidity = False
    End If

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

