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
        strSQL = " Union Select distinct �ļ�ID From ��������ǰ�� Where ����ID IN (Select Column_Value From Table(f_Num2list([3])))"
    End If
    If strDiagnoseIDs <> "" Then
        strSQL = strSQL & " Union Select distinct �ļ�ID From ��������ǰ�� Where ���ID IN (Select Column_Value From Table(f_Num2list([4])))"
    End If
    '֮����ȥ�������ȥ�أ�����Ϊ������д���ű��濨��
    strSQL = Mid(strSQL, 8)
'    strSQL = "(" & Mid(strSQL, 8) & ") Minus Select �ļ�ID From ���Ӳ�����¼ Where ����ID=[1] And ��ҳID=[2] And ��������=5"
    strSQL = "Select /*+ Rule*/ distinct " & vbNewLine & _
            " a.Id, a.����, a.���, a.����, a.����, a.˵��" & vbNewLine & _
            "From �����ļ��б� A ,(" & strSQL & ") B Where A.ID=B.�ļ�ID  And" & vbNewLine & _
            "(a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From ����Ӧ�ÿ��� C Where c.�ļ�id = a.Id And c.����id = [5]))" & vbNewLine & _
            "Order By a.���"
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "CheckDiseaseFile", lngPatiID, lngPageId, strDiseaseIDs, strDiagnoseIDs, lngDeptId)

    If rsTemp.RecordCount = 0 Then Exit Function

    Do Until rsTemp.EOF
        strMsg = strMsg & vbCrLf & "��" & rsTemp!���� & "��"
        rsTemp.MoveNext
    Loop
    rsTemp.MoveFirst

    If MsgBox("���ݲ��˵������Ϣ�����¼���֤�����滹û����д��" & vbCrLf & vbCrLf & Mid(strMsg, 3) & vbCrLf & vbCrLf & "�Ƿ�������д��", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        If CheckDisNum(lngPatiID, lngPageId, lngPatFrom, rsDis, glngOpenedID) Then
            strReason = frmReason.ShowMe(frmPar, NVL(rsDis!�������˵��))
            Call SaveReason(strReason, rsDis!ID, rsDis!��¼״̬)
        End If
        Exit Function
    End If
                    
    If rsTemp.RecordCount = 1 Then
            lngFileID = Val(rsTemp!ID & "")
            Set objDoc = New cEPRDocument
            Call objDoc.InitEPRDoc(cprEM_����, cprET_�������༭, lngFileID, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, False)
            Call objDoc.ShowEPREditor(frmPar, , vbModal)
    ElseIf rsTemp.RecordCount > 1 Then
        If frmDiseaseFileList.ShowMe(frmPar, rsTemp, lngFileID) Then
            rsTemp.Filter = "ID=" & lngFileID
            lngFileID = Val(rsTemp!ID & "")
            Set objDoc = New cEPRDocument
            Call objDoc.InitEPRDoc(cprEM_����, cprET_�������༭, lngFileID, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, False)
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
'���ܣ� �޸Ĵ�Ⱦ������
'������
'   lngFileID:����ID
'   lngPatiID:����ID
'   lngPageId:������ҳID���߹Һ�ID
'   lngPatFrom:������Դ��1-���2-סԺ
'   lngDeptId:���˿���ID
    Dim objDoc As cEPRDocument
    Set objDoc = New cEPRDocument
    Call objDoc.InitEPRDoc(cprEM_�޸�, cprET_�������༭, lngFileID, lngPatFrom, lngPatiID, lngPageId, 0, lngDeptId, 0, False)
    Call objDoc.ShowEPREditor(frmPar, InStr(1, gstrPrivs, "������ӡ") > 0, 1)
    Set objDoc = Nothing
    ModifyDiseaseDoc = True
End Function
