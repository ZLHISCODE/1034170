Attribute VB_Name = "mdlPublicDefine"
Option Explicit

Public gstrSysName As String
Public gbytDiseaseType As Byte '0��ʾ���Ʋ��� 1�ٴ���ϲ��� 2ʵ����ȷ�ﲡ�� 3��ԭЯ���� 4���Լ��������ѪԱ��5δѡ��
Public gbytAcute As Byte       '0��ʾ���� 1��ʾ���� 2��ʾδѡ��
Public gstrKey As String       '��ʾ�༯�Ϲؼ���
Public gstrSql As String
Public glngSys As Long  'ϵͳ��

'��Ҫ��UCheckNorm��ʾ��Ԫ���У������1,�����2,...
Public Const GSTR_OBJNO = ",2,6,9,12,14,15,16,20,21,22,23,24,25,26,27,28,29,30,32,33,34,35,36,37,"
'Ҫ������

Public Const GSTR_ELENAME = "��Ƭ���$�������$����$�ҳ�����$����֤��$�Ա�$��������$����$���䵥λ" & _
                    "$������λ$��ϵ�绰$��������$סַ$����ְҵ$��������1$��������2$��������$�������" & _
                    "$��������$���ഫȾ��$���ഫȾ��$���̲�$�����Ը���$̿��$����$�ν��$�˺�" & _
                    "$�ܲ�$ű��$���ഫȾ��$������Ⱦ��$����Բ�$����״��$ѧ��$��Ⱦ;��$���Դ���" & _
                    "$ѪҺ����$��������$�˿�ԭ��$���浥λ$��ϵ�绰$�ҽ��$�����$��ע"
'�滻��
Public Const GSTR_REPLACE = "0$0$1$0$1$1$1$1$0$1$" & _
                            "1$0$0$0$0$0$0$0$0$0$" & _
                            "0$0$0$0$0$0$0$0$0$0$" & _
                            "0$0$1$1$0$0$0$0$0$1$" & _
                            "0$0$1$0"
'Ҫ������
Public Const GSTR_ELETYPE = "1$1$1$1$1$1$2$1$1$1$" & _
                            "1$1$1$1$1$1$2$2$2$1$" & _
                            "1$1$1$1$1$1$1$1$1$1$" & _
                            "1$1$1$1$1$1$1$1$1$1$" & _
                            "1$1$2$1"

'Ҫ�ر�ʾ
Public Const GSTR_ELEIDT = "0$2$0$0$0$0$0$0$2$0$" & _
                           "0$2$0$2$2$2$0$0$0$3$" & _
                           "3$2$2$2$2$2$2$2$2$3$" & _
                           "0$3$2$2$2$2$2$0$0$0$" & _
                           "0$0$0$0"

Public gcnOracle As ADODB.Connection
Public Const conMenu_Manage_Save = 2601     '�ݴ�
Public Const conMenu_Manage_Finish = 2602   '���
Public Const conMenu_Manage_Cancel = 2603   'ȡ�����
Public Const conMenu_Manage_Exit = 2604     '�˳�
Public Const M_STR_MODULE_MENU_TAG = 26     'ϵͳ��
Public Const FCONTROL = 8
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ��ҩ���� As Long
End Type

Public UserInfo As TYPE_USER_INFO   '�û���Ϣ

Public Enum SignLevel
    cprSL_�հ� = 0              'δǩ��
    cprSL_���� = 1              '����ҽʦǩ��
    cprSL_���� = 2              '����ҽʦǩ��
    cprSL_���� = 3              '����ҽʦǩ��
    cprSL_���� = 4              '���ߣ�ǩ�����𲻰�����ֻ��ʾ��Ա��������ְ�ƣ��Ա���������ҽʦ
End Enum

Public Const PHYSICALOFFSETX = 112  '���ڴ�ӡ�豸���ԣ���ʾ������ҳ�����Ե���ɴ�ӡ��������Ե�ľ��룬�����豸��λ��
Public Const PHYSICALOFFSETY = 113  '���ڴ�ӡ�豸���ԣ���ʾ������ҳ���ϱ�Ե���ɴ�ӡ������ϱ�Ե�ľ��룬�����豸��λ��
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Const WM_MOUSEWHEEL = &H20A          '������

Public glngOffsetX As Long, glngOffsetY As Long

'*************************************************************************
'**�� �� ����HIWORD
'**��    �룺LongIn(Long) - 32λֵ
'**��    ����(Integer) - 32λֵ�ĵ�16λ
'**����������ȡ��32λֵ�ĸ�16λ
'*************************************************************************
Public Function HIWORD(LongIn As Long) As Integer
   ' ȡ��32λֵ�ĸ�16λ
     HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

'*************************************************************************
'**�� �� ����LOWORD
'**��    �룺LongIn(Long) - 32λֵ
'**��    ����(Integer) - 32λֵ�ĵ�16λ
'**����������ȡ��32λֵ�ĵ�16λ
'*************************************************************************
Public Function LOWORD(LongIn As Long) As Integer
   ' ȡ��32λֵ�ĵ�16λ
     LOWORD = LongIn And &HFFFF&
End Function

Public Sub ClearInfo(objCtl As Control)
    On Error GoTo errHand
    
    Select Case TypeName(objCtl)
        Case "uCheckNorm"
            objCtl.Checked = False
        Case "TextBox"
            objCtl.Text = ""
    End Select
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub
Public Sub PrintInfo(ByVal objCtl As Control)

    Dim x As Integer
    Dim y As Integer
    Dim strXY() As String
    Dim intOffset As Integer
    
    On Error GoTo errHand
    intOffset = 0   '�������������ƫ����
    If objCtl.Tag <> "" Then
        strXY = Split(objCtl.Tag, ",")
        x = strXY(0) - intOffset
        y = strXY(1) - intOffset
    Else
        Exit Sub
    End If
    
    Select Case TypeName(objCtl)
        Case "uCheckNorm"
            If objCtl.BoxVisible = True Then
                Printer.Line (glngOffsetX + PScaleX(x), glngOffsetY + PScaleY(y + 2))-(glngOffsetX + PScaleX(x + 13), glngOffsetY + PScaleY(y + 16)), &H0&, B
            End If
            
            If objCtl.Checked = True Then
                Printer.CurrentX = glngOffsetX + PScaleX(x + 1): Printer.CurrentY = glngOffsetY + PScaleY(y + 4)
                Printer.FontName = "����": Printer.FontSize = 8
                Printer.Print "��"
            End If
            
            Printer.FontName = "����_GB2312": Printer.FontSize = 9 'С���
            If objCtl.BoxVisible = True Or objCtl.Name = "ucCheckType" Then
                Printer.CurrentX = glngOffsetX + PScaleX(x + 14)
                Printer.CurrentY = glngOffsetY + PScaleY(y + 3)
            Else
                Printer.CurrentX = glngOffsetX + PScaleX(x)
                Printer.CurrentY = glngOffsetY + PScaleY(y + 3)
            End If

            Printer.Print Trim(objCtl.Caption)
            
        Case "Label"
            Printer.FontName = "����_GB2312": Printer.FontSize = IIf(objCtl.Name = "lblTitle", 18, 9)  'С���
            Printer.FontBold = IIf(objCtl.Name = "lblTitle", True, False)
            Printer.CurrentX = glngOffsetX + PScaleX(x)
            Printer.CurrentY = glngOffsetY + PScaleY(y)
            Printer.Print Trim(objCtl.Caption)
            Printer.FontBold = False
        Case "TextBox"
            If objCtl.Name = "txtIDCard" Then
                Printer.Line (glngOffsetX + PScaleX(x), glngOffsetY + PScaleY(y + 2))-(glngOffsetX + PScaleX(x + 14), glngOffsetY + PScaleY(y + 17)), &H0&, B
                Printer.FontName = "����_GB2312": Printer.FontSize = 9 'С���
                Printer.CurrentX = glngOffsetX + PScaleX(x + 3)
                Printer.CurrentY = glngOffsetY + PScaleY(y + 3)
                Printer.Print Trim(objCtl.Text)
                Exit Sub
            End If
            Printer.FontName = "����_GB2312": Printer.FontSize = 9  'С���
            Printer.CurrentX = glngOffsetX + PScaleX(x + 2)
            Printer.CurrentY = glngOffsetY + PScaleY(y)
            Printer.Print Trim(objCtl.Text)
        Case "Line"
            Printer.Line (glngOffsetX + PScaleX(x), glngOffsetY + PScaleY(y + 2))-(glngOffsetX + PScaleX(strXY(2)), glngOffsetY + PScaleY(y + 2)), &H0&, B
            
    End Select
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub
Public Function PScaleX(ByVal x As Single) As Single
'��ӡ������������Ļ�����ز�һ����ͬ����210���ף���ӡ��������4960.625,��Ļ��793.7
    PScaleX = Printer.ScaleX(Screen.TwipsPerPixelX * x, vbTwips, vbPixels)
End Function

Public Function PScaleY(ByVal y As Single) As Single
    PScaleY = Printer.ScaleY(Screen.TwipsPerPixelY * y, vbTwips, vbPixels)
End Function

Public Sub GetUserInfo()
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
        
    Set rsTemp = zlDatabase.GetUserInfo
    With rsTemp
        If .RecordCount <> 0 Then
            UserInfo.�û��� = .Fields("�û���").Value
            UserInfo.ID = .Fields("ID").Value                 '��ǰ�û�id
            UserInfo.��� = .Fields("���").Value             '��ǰ�û�����
            UserInfo.���� = .Fields("����").Value             '��ǰ�û�����
            UserInfo.���� = Nvl(.Fields("����").Value, "")   '��ǰ�û�����
            UserInfo.����ID = .Fields("����id").Value             '��ǰ�û�����id
        Else
            UserInfo.�û��� = ""
            UserInfo.ID = 0
            UserInfo.��� = ""
            UserInfo.���� = ""
            UserInfo.���� = ""
            UserInfo.����ID = 0
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Function AddStrKey(ByVal strKey As String) As Boolean
'���ܣ����ӹؼ���
'���أ�TRUE��ʾ���ӳɹ���False��ʾ����ʧ��
    On Error GoTo errHand
    
    If InStr(gstrKey, strKey) = 0 Then
        gstrKey = gstrKey & "," & Trim(strKey)
        AddStrKey = True
    Else
        AddStrKey = False
    End If
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Function CheckVal(ByRef intVal As Integer) As Boolean

    On Error GoTo errHand
    
    If InStr("0,1,2,3,4,5,6,7,8,9", Chr(intVal)) = 0 And intVal <> 8 Then
        intVal = 0
        CheckVal = False
    Else
        CheckVal = True
    End If
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Sub ShowMsg(ByVal strMsg As String)
    MsgBox strMsg, vbOKOnly + vbInformation, gstrSysName
End Sub

Public Function GetSaveSql(arrSql() As Variant, colCls As Collection, ByVal strFileId As String, strReportInfo) As Boolean
'���ܣ���֯�����Sql���
'������arrSql:����Sql����
'      colcls:���󼯺�
'      strFile:�ļ�ID
'      strReport:������Ϣ
    Dim objCls As clsReport
    Dim strAllInfo() As String  '���б�����Ϣ��ʽ���������|�����ı�
    Dim strObjNo() As String    '���������Ϣ��ʽ���������1$�������2$�������3.......
    Dim strContent() As String
    Dim strReplace() As String  '�滻����Ϣ��ʽ���滻��1$�滻��2$�滻��3.......
    Dim strEleName() As String  'Ҫ��������Ϣ��ʽ��Ҫ������1$Ҫ������2$Ҫ������3.......
    Dim strEleType() As String  'Ҫ��������Ϣ��ʽ��Ҫ������1$Ҫ������2$Ҫ������3.......
    Dim strEleIdt() As String   'Ҫ�ر�ʾ��Ϣ��ʽ��Ҫ�ر�ʾ1$Ҫ�ر�ʾ2$Ҫ�ر�ʾ3.......
    Dim blnAddCol As Boolean    '�Ƿ���Ҫ�����µĶ��󵽼���
    Dim strKey As String        '���󼯺ϵĹؼ���
    Dim i As Integer
    Dim intNo As Integer
    Dim strTmp As String
    On Error GoTo errHand
    
    GetSaveSql = False
    strAllInfo = Split(strReportInfo, "|")
    
    
    strObjNo = Split(strAllInfo(0), "$")
    strContent = Split(strAllInfo(1), "$")
    
    strReplace = Split(GSTR_REPLACE, "$")
    strEleName = Split(GSTR_ELENAME, "$")
    strEleType = Split(GSTR_ELETYPE, "$")
    strEleIdt = Split(GSTR_ELEIDT, "$")
    
    For i = 0 To UBound(strContent) - 1
        strKey = "K" & Trim(strObjNo(i))
        intNo = val(Trim(strObjNo(i))) - 1
        blnAddCol = AddStrKey(strKey)
        Set objCls = colCls(strKey)
        objCls.FileID = Trim(strFileId)
        objCls.StartR = 1
        objCls.StopR = 0
        objCls.ObjNo = Trim(strObjNo(i))
        objCls.ObjType = IIf(val(objCls.ObjNo) = 42, 8, 4)
        strTmp = Replace(Trim(strContent(i)), "��", "")
        strTmp = Replace(strTmp, "(", "")
        strTmp = Replace(strTmp, ")", "")
        objCls.Txt = strTmp
        objCls.Replace = Trim(strReplace(intNo))
        objCls.EleName = Trim(strEleName(intNo))
        objCls.EleType = Trim(strEleType(intNo))
        objCls.EleIdt = Trim(strEleIdt(intNo))
        objCls.EleRange = ""
        Call objCls.GetSaveSql(arrSql, blnAddCol)
    Next
    
    GetSaveSql = True
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
Public Function GetUserSignLevel(ByVal lngUserID As Long, ByVal lngPatiID As Long, lngPatiPageID As Long) As SignLevel
'## ˵����  ���ݡ���Ա�����еġ�Ƹ�μ���ְ���ֶ�ȷ��ҽ����������סԺҽʦ������ҽʦ������ҽʦ��
    Dim rs As New ADODB.Recordset, lngR As Long, lngLevel1 As Long, lngLevel2 As Long
    Err = 0: On Error GoTo errHand
    
    gstrSql = "select Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        lngR = Nvl(rs("Ƹ�μ���ְ��"), 0)
    End If
    Select Case lngR    '1 ����  2 ����  3 �м�  4 ����/ʦ��  5 Ա/ʿ  9 ��Ƹ
    Case 1: lngLevel1 = cprSL_����
    Case 2: lngLevel1 = cprSL_����
    Case 3: lngLevel1 = cprSL_����
    Case Else: lngLevel1 = cprSL_����
    End Select
    If lngLevel1 = cprSL_���� Then lngLevel1 = cprSL_���� '���ߣ�ǩ�����𲻰�����ֻ��ʾ��Ա��������ְ�ƣ��Ա���������ҽʦ;�ڱ������в�ʹ�� ����
    rs.Close
    
    If lngPatiID > 0 Then
        gstrSql = "Select ����ҽʦ, ����ҽʦ, ����ҽʦ " & _
            " From ���˱䶯��¼ " & _
            " Where ����ID = [1] And ��ҳID = [2] And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1) " & _
            "       And ��ʼʱ�� Is Not Null And Nvl(���Ӵ�λ, 0) = 0"
        Set rs = zlDatabase.OpenSQLRecord(gstrSql, "cEPRDocument", lngPatiID, lngPatiPageID)
        If rs.EOF Then
            lngLevel2 = cprSL_����
        Else
            If rs.Fields("����ҽʦ") = UserInfo.���� Then
                lngLevel2 = cprSL_����
            ElseIf rs.Fields("����ҽʦ") = UserInfo.���� Then
                lngLevel2 = cprSL_����
            Else
                lngLevel2 = cprSL_����
            End If
        End If
    End If
    GetUserSignLevel = IIf(lngLevel1 >= lngLevel2, lngLevel1, lngLevel2)
    Exit Function

errHand:
    GetUserSignLevel = cprSL_�հ�
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
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

Public Function GetNextDoubleId(strTable As String) As Double
    '------------------------------------------------------------------------------------
    '���ܣ���ȡָ��������Ӧ������(���淶������������Ϊ��������_id��)����һ��ֵ
    '������
    '   strTable��������
    '���أ�
    '------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strtab As String
    
    '�����ô��������,ԭ��������ʧЧ��û������ʱ,Ӧ�÷��ش���,��Ȼ������,��������!
    '31730
    'On Error GoTo errH
    strtab = Trim(strTable)
    If strtab = "������ü�¼" Or strtab = "סԺ���ü�¼" Then strtab = "���˷��ü�¼"
    
    strSQL = "Select " & strtab & "_ID.Nextval From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetNextDoubleId")
    GetNextDoubleId = rsTmp.Fields(0).Value
'    Exit Function
'errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
End Function