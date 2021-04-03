VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function EditDiseaseDoc(ByVal frmPar As Object, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long, ByVal strDiseaseIDs As String, ByVal strDiagnoseIDs As String, Optional ByRef strReason As String) As Boolean
    Dim rsTemp As ADODB.Recordset, lngFileID As Long, strMsg As String
    Dim objDoc As cEPRDocument, objInfection As Object
    Dim strSQL As String
    Dim rsDis As ADODB.Recordset
    On Error GoTo errHand
    If strDiseaseIDs = "" And strDiagnoseIDs = "" Then Exit Function
    strSQL = ""
    If strDiseaseIDs <> "" Then
        strSQL = " Union Select distinct 文件ID From 疾病报告前提 Where 疾病ID IN (Select Column_Value From Table(f_Num2list([3])))"
    End If
    If strDiagnoseIDs <> "" Then
        strSQL = strSQL & " Union Select distinct 文件ID From 疾病报告前提 Where 诊断ID IN (Select Column_Value From Table(f_Num2list([4])))"
    End If
    '之所以去掉后面的去重，是因为允许填写两张报告卡。
    strSQL = Mid(strSQL, 8)
'    strSQL = "(" & Mid(strSQL, 8) & ") Minus Select 文件ID From 电子病历记录 Where 病人ID=[1] And 主页ID=[2] And 病历种类=5"
    strSQL = "Select /*+ Rule*/ distinct " & vbNewLine & _
            " a.Id, a.种类, a.编号, a.名称, a.保留, a.说明" & vbNewLine & _
            "From 病历文件列表 A ,(" & strSQL & ") B Where A.ID=B.文件ID  And" & vbNewLine & _
            "(a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 病历应用科室 C Where c.文件id = a.Id And c.科室id = [5]))" & vbNewLine & _
            "Order By a.编号"
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "CheckDiseaseFile", lngPatiID, lngPageId, strDiseaseIDs, strDiagnoseIDs, lngDeptId)

    If rsTemp.RecordCount = 0 Then Exit Function

    Do Until rsTemp.EOF
        strMsg = strMsg & vbCrLf & "【" & rsTemp!名称 & "】"
        rsTemp.MoveNext
    Loop
    rsTemp.MoveFirst

    If MsgBox("根据病人的诊断信息，以下疾病证明报告还没有填写：" & vbCrLf & vbCrLf & Mid(strMsg, 3) & vbCrLf & vbCrLf & "是否立即填写？", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        If CheckDisNum(lngPatiID, lngPageId, lngPatFrom, rsDis, glngOpenedID) Then
            strReason = frmReason.ShowMe(frmPar, NVL(rsDis!处理情况说明))
            Call SaveReason(strReason, rsDis!ID, rsDis!记录状态)
        End If
        Exit Function
    End If
                    
    If rsTemp.RecordCount = 1 Then
            lngFileID = Val(rsTemp!ID & "")
            Set objDoc = New cEPRDocument
            Call objDoc.InitEPRDoc(cprEM_新增, cprET_单病历编辑, lngFileID, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, False)
            Call objDoc.ShowEPREditor(frmPar, , vbModal)
    ElseIf rsTemp.RecordCount > 1 Then
        If frmDiseaseFileList.ShowMe(frmPar, rsTemp, lngFileID) Then
            rsTemp.Filter = "ID=" & lngFileID
            lngFileID = Val(rsTemp!ID & "")
            Set objDoc = New cEPRDocument
            Call objDoc.InitEPRDoc(cprEM_新增, cprET_单病历编辑, lngFileID, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, False)
            Call objDoc.ShowEPREditor(frmPar, , vbModal)
        End If
    End If
   
    EditDiseaseDoc = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function ModifyDiseaseDoc(ByVal frmPar As Object, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long, ByVal lngDeptId As Long) As Boolean
'功能： 修改传染病报告
'参数：
'   lngFileID:报告ID
'   lngPatiID:病人ID
'   lngPageId:病人主页ID或者挂号ID
'   lngPatFrom:病人来源，1-门诊，2-住院
'   lngDeptId:病人科室ID
    Dim objDoc As cEPRDocument
    Set objDoc = New cEPRDocument
    Call objDoc.InitEPRDoc(cprEM_修改, cprET_单病历编辑, lngFileID, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, False)
    Call objDoc.ShowEPREditor(frmPar, InStr(1, gstrPrivs, "病历打印") > 0, 1)
    Set objDoc = Nothing
    ModifyDiseaseDoc = True
End Function

