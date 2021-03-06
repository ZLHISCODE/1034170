VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm保险项目选择重庆 
   AutoRedraw      =   -1  'True
   Caption         =   "医保项目选择"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm保险项目选择重庆.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7845
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7845
      TabIndex        =   4
      Top             =   4350
      Width           =   7845
      Begin VB.CommandButton cmdRequery 
         Caption         =   "更新明细"
         Height          =   350
         Left            =   3900
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打印列表"
         Height          =   350
         Left            =   2790
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         Top             =   175
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   7
         Top             =   150
         Width           =   1100
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "明细查找(&F)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2340
      MousePointer    =   9  'Size W E
      ScaleHeight     =   930
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   4050
      Left            =   3060
      TabIndex        =   2
      Top             =   270
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   7144
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   45
      Top             =   3915
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目选择重庆.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目选择重庆.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4050
      Left            =   0
      TabIndex        =   10
      Top             =   270
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7144
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "项目大类(&K)"
      Height          =   240
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2970
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "项目明细(&D)"
      Height          =   240
      Left            =   3060
      TabIndex        =   1
      Top             =   30
      Width           =   4710
   End
End
Attribute VB_Name = "frm保险项目选择重庆"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Private mstrCode As String
Private mstrName As String
Private mdbl医院价格 As Double
Private mobjStream As TextStream
Private mobjFileSystem As New FileSystemObject
Private mblnOK As Boolean
Private mcnYB As New ADODB.Connection   '医保前置服务器连接
'调试重庆医保银海版 204-04-07 主要是加了函数，并修改了药品、诊疗及病种，解决名称含单引号的问题

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim str标准单价 As String, str限价 As String, str特批项目 As String, str特批价 As String
    
    If lvwDetail.SelectedItem Is Nothing Then
        MsgBox "没有选择项目！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '对价格进行判断
    Call GetValueByCol("限价", str限价)
    If str限价 <> "" And mdbl医院价格 > 0 Then
        Call GetValueByCol("标准单价", str标准单价)
        Call GetValueByCol("特批项目标志", str特批项目)
        Call GetValueByCol("特批价", str特批价)
        
        If mint险类 = TYPE_重庆市 Then
            If 价格判断_重庆(mdbl医院价格, Val(str标准单价), str限价, str特批项目 = "是", Val(str特批价)) = False Then
                Exit Sub
            End If
        Else
            If 价格判断_重庆银海版(mdbl医院价格, Val(str标准单价), str限价, str特批项目 = "是", Val(str特批价)) = False Then
                Exit Sub
            End If
        End If
    End If
    
    '返回选择项目编码
    mstrCode = lvwDetail.SelectedItem.Text
    '商品名与项目名称只可能有一个有效
    Call GetValueByCol("商品名", mstrName)
    Call GetValueByCol("项目名称", mstrName)
    
    mblnOK = True
    Unload Me
End Sub

Private Sub GetValueByCol(ByVal strColumnName As String, strValue As String)
    Dim lngCount As Long, lngIndex As Long

    For lngCount = 1 To lvwDetail.ColumnHeaders.Count
        If lvwDetail.ColumnHeaders(lngCount).Text = strColumnName Then
            lngIndex = lngCount
            Exit For
        End If
    Next
    
    If lngIndex > 0 Then
        strValue = lvwDetail.SelectedItem.SubItems(lngIndex - 1)
    End If
End Sub

Public Function GetCode(strCode As String, strName As String, ByVal dbl医院价格 As Double, ByVal int险类 As Integer) As Boolean
'功能：获得一个收费项目的医保编码
'参数：strCode 既作为输入参数，又输出
'返回：成功返回True
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim nod As Node
    
    mblnOK = False
    mstrCode = strCode
    mdbl医院价格 = dbl医院价格
    mint险类 = int险类
    
    On Error GoTo errH
    
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & int险类
    Call OpenRecordset(rsTemp, Me.Caption)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保服务器"
                strServer = strTemp
            Case "医保用户名"
                strUser = strTemp
            Case "医保用户密码"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
        Exit Function
    End If
    
    If int险类 = TYPE_重庆银海版 Then
        '调试重庆医保银海版 204-03-29
        On Error Resume Next
        If Not 医保初始化_重庆银海版 Then
            Unload Me
            Exit Function
        End If
    End If
    
    '调试重庆医保银海版 204-03-29
    If int险类 = TYPE_重庆市 Then
        '显示药品类别
        gstrSQL = "select BH id,FXBH 上级ID,LBDM 编码,LBMC 名称,'Y' 类别,level 级数 from YPML_LBDM start with FXBH=0 connect by prior BH=FXBH " & _
                 " Union All " & _
                 " select LBDM as id,'0' 上级ID,LBDM 编码,LBMC 名称,'Z' 类别,1 级数 from zlxm_lbdm2 " & _
                 " order by 类别 Desc,级数,编码"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
    Else
        gstrSQL = "Select ID,编码,名称 From 保险支付大类 Where 险类=" & TYPE_重庆银海版
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    End If
    If rsTemp.EOF = True Then
        MsgBox "医保前置服务器中没有药品类别数据，无法选择。", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    
    tvwClass.Nodes.Clear
    Do Until rsTemp.EOF
        If int险类 = TYPE_重庆市 Then
            If rsTemp("上级id") = 0 Then
                Set nod = tvwClass.Nodes.Add(, , rsTemp("类别") & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "Class", "Class")
            Else
                Set nod = tvwClass.Nodes.Add(rsTemp("类别") & rsTemp("上级id"), tvwChild, rsTemp("类别") & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "Class", "Class")
            End If
        Else
            Set nod = tvwClass.Nodes.Add(, , "K" & rsTemp("id"), rsTemp("名称"), "Class", "Class")
        End If
        nod.Sorted = True
        rsTemp.MoveNext
    Loop
    
    tvwClass.Nodes(1).Selected = True
    Call FillList
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    
    frm保险项目选择重庆.Show vbModal, frm保险项目
    '返回值
    If mblnOK = True Then
        strCode = mstrCode
        strName = mstrName
    End If
    GetCode = mblnOK
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FillList()
'功能：显示当前类别下的医保明细
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, fld As ADODB.Field
    Dim str类别代码 As String, blnColSet As Boolean
    Dim lngCol  As Long
    Dim varValue As Variant
    
    Me.MousePointer = vbHourglass
    
    On Error GoTo errHandle
    If mint险类 = TYPE_重庆市 Then
        With tvwClass.SelectedItem
            str类别代码 = Mid(.Text, 2, InStr(.Text, "】") - 2)
        End With
    Else
        Select Case tvwClass.SelectedItem
        Case "甲类药"
            str类别代码 = 1
        Case "乙类药"
            str类别代码 = 2
        Case "自费药"
            str类别代码 = 3
        End Select
    End If
    
    rsTemp.CursorLocation = adUseClient
    '暂时让列表不能刷新
    LockWindowUpdate lvwDetail.hwnd
    lvwDetail.ListItems.Clear
    
    If mint险类 = TYPE_重庆市 Then
        If Left(tvwClass.SelectedItem.Key, 1) = "Y" Then
            '当前选择是的一个药品类别
            gstrSQL = "select YPLSH  医保编码,YPBM 药品编码,TYM 通用名称,SPM 商品名,SPMZJM 商品名简码,YCMC 药厂名称,decode(FYDJ,1,'甲类',2,'乙类','自费') 费用等级 " & _
                      "      ,PFJ 批发价,BZDJ 标准单价,ZFBL 自付比例,JX 剂型,BZSL 包装数量,BZDW 包装单位,HL 含量,HLDW 含量单位,RL 容量,RLDW 容量单位 " & _
                      "      ,DECODE(CFYBZ,1,'是') 处方药标志,decode(GMP,1,'是') GMP标志,decode(YPXJFS,1,'限价') 限价,TQFYDJ 特群项目等级,TQZFBL 特群自付比例,TQBZDJ 特群标准单价 " & _
                      "  FROM YPML WHERE LBDM='" & str类别代码 & "'"
        Else
            '当前选择是的一个诊疗类别
            gstrSQL = "Select XMLSH 医保编码,XMBM 诊疗编码,XMMC 项目名称,ZJM 简码,decode(FYDJ,1,'甲类',2,'乙类','自费') 费用等级,DW 单位 " & _
                     "       ,TPJ 特批价,BZJ 标准单价,ZZBL 在职自付比例,TXBL 退休自付比例,decode(XJFS,1,'统一限价',2,'按医院等级定价',3,'按二级医院标准浮动比例') 限价 " & _
                     "       ,TQFYDJ 特群项目等级,TQZFBL 特群自付比例,TQBZDJ 特群标准单价,decode(TPXMBZ,1,'是') 特批项目标志,BZ 备注 " & _
                     "   FROM ZLXM WHERE LBDM2='" & str类别代码 & "'"
        End If
    Else
        If tvwClass.SelectedItem Like "*药*" Then
            '当前选择是的一个药品类别
            gstrSQL = "select 流水号 医保编码,编码 药品编码,通用名 通用名称,商品名 商品名,商品名助记码 商品名简码,药厂名称,decode(项目等级,1,'甲类',2,'乙类','自费') 费用等级 " & _
                      "      ,批发价,标准单价,自付比例,剂型,包装数量,包装单位,含量,含量单位,容量,容量单位 " & _
                      "      ,DECODE(处方药标志,1,'是') 处方药标志,decode(GMP标志,1,'是') GMP标志,decode(限价方式,1,'限价') 限价,特群项目等级,特群自付比例,特群标准单价 " & _
                      "  FROM 中间库_药品目录 WHERE 项目等级='" & str类别代码 & "'"
        Else
            '当前选择是的一个诊疗类别
            gstrSQL = "Select 流水号 医保编码,项目编码 诊疗编码,项目名称,助记码 简码,decode(项目等级,1,'甲类',2,'乙类','自费') 费用等级,单位 " & _
                     "       ,特批价,标准单价,在职比例 在职自付比例,退休比例 退休自付比例,decode(限价方式,1,'统一限价',2,'按医院等级定价',3,'按二级医院标准浮动比例') 限价 " & _
                     "       ,特群项目等级,特群自付比例,特群项目单价,decode(特批项目标志,1,'是') 特批项目标志,备注 " & _
                     "   FROM 中间库_诊疗项目 "
        End If
    End If
    rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly
    
    If lvwDetail.ColumnHeaders.Count <> rsTemp.Fields.Count Then
        '重新处理表头
        blnColSet = True
        lvwDetail.ColumnHeaders.Clear
        For Each fld In rsTemp.Fields
            lvwDetail.ColumnHeaders.Add , , fld.Name, 1000
        Next
    End If
        
    Do Until rsTemp.EOF
        Set lst = lvwDetail.ListItems.Add(, "K" & rsTemp("医保编码"), rsTemp("医保编码"), "Detail", "Detail")
        
        '根据ListView的列名从数据库取数
        For lngCol = 2 To lvwDetail.ColumnHeaders.Count
            varValue = rsTemp(lvwDetail.ColumnHeaders(lngCol).Text).Value
            lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
        Next
        rsTemp.MoveNext
    Loop
    If blnColSet = True Then
        '重新对列进行了处理
        If lvwDetail.ListItems.Count > 0 Then Call zlControl.LvwSetColWidth(lvwDetail)
    End If
    LockWindowUpdate 0
    Me.MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    LockWindowUpdate 0
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    
    objPrint.Title.Text = "保险项目"
    Set objPrint.Body.objData = lvwDetail
    objPrint.UnderAppItems.Add "医保大类：" & tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "打印人：" & gstrUserName
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    Select Case zlPrintAsk(objPrint)
        Case 1
             zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
    End Select

End Sub

Private Sub cmdRequery_Click()
    Dim strInput As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln全量 As Boolean
    
    If MsgBox("本操作可能会花比较长的时间，是否继续？" & vbCrLf & vbCrLf & "另外注意，本操作只更新医保项目明细，而不包括对应关系。", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    Me.Caption = "医保项目选择（正在读取从文件或网络读取保险项目明细，请稍候......）"
    
    '调试重庆医保银海版 204-04-07
    '检查本次是全量更新还是增量更新(修改量大)
    gstrSQL = "Select 1 From zlcq.中间库_药品目录 where rownum<2"
    Call OpenRecordset(rsTemp, "检查本次是全量更新还是增量更新")
    bln全量 = (rsTemp.RecordCount = 0)
    
    mcnYB.BeginTrans
    gcnOracle.BeginTrans
    If Not AnalyFile_YPML(bln全量) Then
        MousePointer = vbDefault
        mcnYB.RollbackTrans
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    If Not AnalyFile_ZLML(bln全量) Then
        MousePointer = vbDefault
        mcnYB.RollbackTrans
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    If Not AnalyFile_BZML(bln全量) Then
        MousePointer = vbDefault
        mcnYB.RollbackTrans
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    mcnYB.CommitTrans
    gcnOracle.CommitTrans
    
    '重新装入明细
    Call FillList
    MousePointer = vbDefault
    Me.Caption = "医保项目选择"
End Sub

Private Sub Form_Load()
    cmdRequery.Visible = (mint险类 = TYPE_重庆银海版)
End Sub

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = tvwClass.Width
    
    On Error Resume Next
    
    tvwClass.Left = 0: tvwClass.Top = lblClass.Top + lblClass.Height
    tvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = tvwClass.Top
    picSplit.Left = tvwClass.Left + tvwClass.Width
    picSplit.Height = tvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If tvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
    
    lvwDetail.Top = tvwClass.Top
    lvwDetail.Left = lblDetail.Left
    lvwDetail.Width = lblDetail.Width
    lvwDetail.Height = tvwClass.Height
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
    cmdPrint.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwDetail_DblClick()
    cmdOK_Click
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvwClass.Width + x < 1000 Or lvwDetail.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        tvwClass.Width = tvwClass.Width + x
        
        lblDetail.Left = lblDetail.Left + x
        lblDetail.Width = lblDetail.Width - x
        
        lvwDetail.Left = lvwDetail.Left + x
        lvwDetail.Width = lvwDetail.Width - x
    End If
End Sub

Private Sub lvwdetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwDetail.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwDetail.SortOrder = lvwDescending
    Else
        lvwDetail.SortOrder = lvwAscending
    End If
    lvwDetail.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwDetail.SelectedItem Is Nothing Then lvwDetail.SelectedItem.EnsureVisible
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillList
End Sub

Private Sub txtFind_Change()
'功能：根据用户输入的内容查找匹配的内容
    Dim lst As ListItem, lngIndex As Long, lngSubItems As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    If lvwDetail.ListItems.Count = 0 Then Exit Sub

    Set lst = lvwDetail.FindItem(strFind, lvwText, , lvwPartial)
    If Not lst Is Nothing Then
        lst.Selected = True
        lst.EnsureVisible
    Else
        '非文本不能做到部分匹配
        lngSubItems = lvwDetail.ColumnHeaders.Count - 1
        For Each lst In lvwDetail.ListItems
            For lngIndex = 1 To lngSubItems
                If lst.SubItems(lngIndex) Like strFind & "*" Then
                    lst.Selected = True
                    lst.EnsureVisible
                    Exit Sub
                End If
            Next
            
        Next
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Function AnalyFile_YPML(Optional ByVal bln全量 As Boolean = True) As Boolean
    '分析接口返回的药品目录文件，并保存到中间库
    Dim lngRow As Long, lngCol As Long
    Dim lngRows As Long, lngCols As Long
    Dim strData As String, strDate As String
    Dim strDeal As String, strInput As String
    Dim arrCol
'    变量类型——1：新增2：修改3：删除
    Const int变更时间 As Integer = 23
    On Error GoTo ErrHand
    
    strInput = IIf(bln全量, "", Format(zlDatabase.Currentdate(), "yyyyMMdd HH:mm:ss")) & "|C:\CQYB_YH\YPML.txt"
    Call 调用接口_准备_重庆银海版("02", strInput)
    If Not 调用接口_重庆银海版() Then Exit Function
    
    If Not mobjFileSystem.FileExists("C:\CQYB_YH\YPML.txt") Then Exit Function
    Set mobjStream = mobjFileSystem.OpenTextFile("C:\CQYB_YH\YPML.txt", ForReading, False, TristateMixed)
    If bln全量 Then mcnYB.Execute "ZL_中间库_药品目录_DELETEALL()", , adCmdStoredProc
    
    strInput = "ZL_中间库_药品目录_Insert("
    Do While Not mobjStream.AtEndOfStream
        strData = mobjStream.ReadLine
        arrCol = Split(strData, vbTab)
        lngCols = UBound(arrCol)
        strDeal = ""
        For lngCol = 0 To lngCols
            '如果不是全量,且是最后一个字段,则判断操作类型
            If Not bln全量 And lngCol = lngCols Then
                Select Case Val(arrCol(lngCol))
                Case 1
                    strDeal = "ZL_中间库_药品目录_Insert(" & Mid(strDeal, 2) & ")"
                Case 2
                    strDeal = "ZL_中间库_药品目录_Update(" & Mid(strDeal, 2) & ")"
                Case Else
                    strDeal = "ZL_中间库_药品目录_Delete('" & arrCol(0) & "')"
                End Select
            Else
                Select Case lngCol
                Case int变更时间
                    '由于日期格式不同，需要转换
                    strDate = ReplaceStr(arrCol(lngCol))
                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                    strDate = ",to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                    strDeal = strDeal & strDate
                Case Else
                    If Trim(arrCol(lngCol)) = "" Then
                        strDeal = strDeal & ",NULL"
                    Else
                        strDeal = strDeal & ",'" & ReplaceStr(arrCol(lngCol)) & "'"
                    End If
                End Select
            End If
        Next
        If bln全量 Then strDeal = strInput & Mid(strDeal, 2) & ")"
        mcnYB.Execute strDeal, , adCmdStoredProc
    Loop
    mobjStream.Close
    
    AnalyFile_YPML = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyFile_ZLML(Optional ByVal bln全量 As Boolean = True) As Boolean
    '分析接口返回的药品目录文件，并保存到中间库
    Dim lngRow As Long, lngCol As Long
    Dim lngRows As Long, lngCols As Long
    Dim strData As String, strDate As String
    Dim strDeal As String, strInput As String
    Dim arrRow, arrCol
    Dim strHosCode As String, bln特批项目 As Boolean
    
    Const int变更时间 As Integer = 14
    Const int诊疗项目分类 As Integer = 17
    Const int医疗机构编码 As Integer = 16
    On Error GoTo ErrHand
    
    strInput = IIf(bln全量, "", Format(zlDatabase.Currentdate(), "yyyyMMdd HH:mm:ss")) & "|C:\CQYB_YH\ZLML.txt"
    Call 调用接口_准备_重庆银海版("03", strInput)
    If Not 调用接口_重庆银海版() Then Exit Function
    
    If Not mobjFileSystem.FileExists("C:\CQYB_YH\ZLML.txt") Then Exit Function
    Set mobjStream = mobjFileSystem.OpenTextFile("C:\CQYB_YH\ZLML.txt", ForReading, False, TristateMixed)
    If bln全量 Then mcnYB.Execute "ZL_中间库_诊疗项目_DELETEALL()", , adCmdStoredProc
    
    strInput = "ZL_中间库_诊疗项目_Insert("
    Do While Not mobjStream.AtEndOfStream
        strData = mobjStream.ReadLine
        arrCol = Split(strData, vbTab)
        lngCols = UBound(arrCol)
        strDeal = ""
        For lngCol = 0 To lngCols
            If Not bln全量 And lngCol = lngCols Then
                Select Case Val(arrCol(lngCol))
                Case 1
                    strDeal = "ZL_中间库_诊疗项目_Insert(" & Mid(strDeal, 2) & IIf(bln特批项目, ",1", "") & ")"
                Case 2
                    strDeal = "ZL_中间库_诊疗项目_Update(" & Mid(strDeal, 2) & IIf(bln特批项目, ",1", "") & ")"
                Case Else
                    strDeal = "ZL_中间库_诊疗项目_Delete('" & arrCol(0) & "')"
                End Select
            Else
                Select Case lngCol
                Case int变更时间
                    '由于日期格式不同，需要转换
                    strDate = ReplaceStr(arrCol(lngCol))
                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                    strDate = ",to_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                    strDeal = strDeal & strDate
                Case int医疗机构编码
                    strHosCode = ReplaceStr(arrCol(lngCol))
                    strDeal = strDeal & ",'" & Trim(arrCol(lngCol)) & "'"
                Case int诊疗项目分类
                    bln特批项目 = False
                    If strHosCode = gComInfo_重庆银海版.医院编码 Then
                        If Val(arrCol(lngCol)) = 3 Then
                            '特批项目
                            bln特批项目 = True
                        End If
                    End If
                    strDeal = strDeal & ",'" & ReplaceStr(arrCol(lngCol)) & "'"
                Case Else
                    strDeal = strDeal & ",'" & ReplaceStr(arrCol(lngCol)) & "'"
                End Select
            End If
        Next
        If bln全量 Then strDeal = strInput & Mid(strDeal, 2) & IIf(bln特批项目, ",1", "") & ")"
        mcnYB.Execute strDeal, , adCmdStoredProc
    Loop
    mobjStream.Close
    
    AnalyFile_ZLML = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyFile_BZML(Optional ByVal bln全量 As Boolean = True) As Boolean
    '分析接口返回的药品目录文件，并保存到中间库
    Dim lngRow As Long, lngCol As Long
    Dim lngRows As Long, lngCols As Long
    Dim str编码 As String, str名称 As String, str简码 As String, str类别 As String
    Dim strDeal As String, strInput As String, strData As String
    Dim arrRow, arrCol
    Dim lngNextID As Long
    Dim intType As Integer          '1-新增;2-修改;3-删除
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    strInput = "C:\CQYB_YH\BZML.txt"
    Call 调用接口_准备_重庆银海版("04", strInput)
    If Not 调用接口_重庆银海版() Then Exit Function
    
    If Not mobjFileSystem.FileExists("C:\CQYB_YH\BZML.txt") Then Exit Function
    Set mobjStream = mobjFileSystem.OpenTextFile("C:\CQYB_YH\BZML.txt", ForReading, False, TristateMixed)
    
    '打开现有病种
    gstrSQL = "Select ID,编码 From 保险病种 Where 险类=" & TYPE_重庆银海版
    Call OpenRecordset(rsTemp, "读取现有病种目录")
    
    Do While Not mobjStream.AtEndOfStream
        strData = mobjStream.ReadLine
        arrCol = Split(strData, vbTab)
        
        str编码 = ReplaceStr(arrCol(0))
        str名称 = ReplaceStr(arrCol(1))
        str简码 = ReplaceStr(arrCol(4))
        str类别 = Val(arrCol(2)) - 1
        If Val(str类别) < 0 Then str类别 = 0
        
        With rsTemp
            .Filter = "编码='" & str编码 & "'"
            intType = IIf(.RecordCount = 0, 1, 2)
        End With
        
        If Not bln全量 Then
            intType = Val(arrCol(UBound(arrCol)))
        End If
        
        '更新保险疾病
        Select Case intType
        Case 1
            lngNextID = zlDatabase.GetNextId("保险病种")
            gstrSQL = "zl_保险病种_INSERT(" & lngNextID & "," & TYPE_重庆银海版 & ",'" & str编码 & _
                        "','" & str名称 & "','" & str简码 & "'," & str类别 & ",NULL,NULL)"
        Case 2
            lngNextID = rsTemp!ID
            gstrSQL = "zl_保险病种_UPDATE(" & lngNextID & ",'" & str编码 & _
                        "','" & str名称 & "','" & str简码 & "'," & str类别 & ",NULL,NULL)"
        Case Else
            lngNextID = rsTemp!ID
            gstrSQL = "zl_保险病种_DELETE(" & lngNextID & ")"
        End Select
        Call ExecuteProcedure(Me.Caption)
    Loop
    mobjStream.Close
    
    AnalyFile_BZML = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReplaceStr(ByVal strInput As String) As String
    ReplaceStr = Trim(Replace(strInput, "'", ""))
    ReplaceStr = Replace(ReplaceStr, """", "")
End Function
