VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private arrForms() As Variant
Private mblnDataIsEmpty As Boolean

'�¼�����
Public Event CommitCondition(ByVal strNO As String, strCond As String, frmParent As Object) '��������ȷ��

Public Event ResetCondition(ByVal strNO As String, strCond As String, blnInhere As Boolean, frmParent As Object) '������������

Public Event EnterCell(ByVal strNO As String, Row As Integer, Col As Integer, Text As String) '����ĳ�������һ����Ԫ
Public Event LevelCell(ByVal strNO As String, Row As Integer, Col As Integer, Text As String) '�뿪ĳ�������һ����Ԫ
Public Event EnterRow(ByVal strNO As String, Row As Integer, RowText As String, Sheet As Object) '����ĳ�������һ��
Public Event EnterCol(ByVal strNO As String, Col As Integer, Sheet As Object) '����ĳ�������һ��
Public Event ColResize(ByVal strNO As String, Col As Integer, NewWidth As Single, OldWidth As Single) '�ı�ĳ��������п�

Public Event BeforeReportLoad(ByVal strNO As String, Form As Object) '��ʼ��ȡ����
Public Event AfterReportLoad(ByVal strNO As String, Form As Object) '��ȡ��������
Public Event ReportActive(ByVal strNO As String, Form As Object) '�������弤��

Public Event BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant) '������ӡ֮ǰ
Public Event PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long) '��ӡ����������ʱ
Public Event AfterPrint(ByVal ReportNum As String) '������ӡ֮��

Public Event PaperMouseDown(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, blnPop As Boolean) '����ֽ������갴��
Public Event PaperMouseMove(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single) '����ֽ��������ƶ�
Public Event PaperMouseUp(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single) '����ֽ��������ɿ�

Public Event SheetDblClick(ByVal strNO As String, Sheet As Object, frmParent As Object)

Public Event SheetMouseDown(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object) '������������갴��
Public Event SheetMouseMove(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object) '��������������ƶ�
Public Event SheetMouseUp(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object) '��������������ɿ�

Public Event PaperScroll(ByVal strNO As String, Way As Byte, Percent As Single) '�������������
Public Event SheetScroll(ByVal strNO As String, Sheet As Object) 'ĳ�������ڲ�����

Public Event ReportUnload(ByVal strNO As String, frmReport As Object) 'ĳ�������ڲ�����

Public Property Get DataIsEmpty() As Boolean
'���Թ��ܣ�ֱ�ӵ��ñ�����ӡʱ���жϱ��������Ƿ�Ϊ�ա�
'ע��������ֻ��ֱ�ӵ��ñ�����ӡʱ���á�
    DataIsEmpty = mblnDataIsEmpty
End Property

Friend Property Let DataIsEmpty(ByVal blnDataIsEmpty As Boolean)
'˵���������ɶ���ֻ�����ԣ�ʵ�ʲ��Ժ�����Ч��
    mblnDataIsEmpty = blnDataIsEmpty
End Property

'==�¼��������(�ⲿ����)=====================================================================================
Friend Sub Act_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant)
'���ܣ�����ֱ���������ӡ��ʱ����ҳ��������ϴ�ӡ֮ǰ����
'������ReportNum=�������
'      TotalPages=ȫ����ӡʱ����ҳ����-1��ʾֻ��ӡ��ǰҳ��-2��ʾ��ӡָ����Χ��ҳ
'      Cancel=�Ƿ�ȡ����ӡ
'      arrBill=ʵ��Ҫ��ӡ��Ʊ�ݺ����顰Split("A,B,C",",")����ȱʡΪ������/Empty
    RaiseEvent BeforePrint(ReportNum, TotalPages, Cancel, arrBill)
End Sub

Friend Sub Act_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long)
'���ܣ���ӡ����������ʱ����
'������Page=ҳ��(1-n),Row=�к�(1-n),ID=����Դ�е�ID��ʶ(��һ���󶨵��˱�����)
'˵����������е�Ԫ��һ����û�����ݣ�����Ϊ����ӡ�����в�����¼�
    RaiseEvent PrintSheetRow(ReportNum, Sheet, Page, Row, ID)
End Sub

Friend Sub Act_AfterPrint(ByVal ReportNum As String)
'���ܣ�����ֱ���������ӡ��֮�󼤻�
'������ReportNum=�������
    RaiseEvent AfterPrint(ReportNum)
End Sub

Friend Sub Act_ColResize(ByVal strNO As String, Col As Integer, NewWidth As Single, OldWidth As Single)
'���ܣ��ı�ĳ��������п�ʱ����
'���������������µĸı���п���
    RaiseEvent ColResize(strNO, Col, NewWidth, OldWidth)
End Sub

Friend Sub Act_SheetScroll(ByVal strNO As String, Sheet As Object)
'���ܣ�ĳ�������ڲ�����ʱ����
'��������������Sheet������
    RaiseEvent SheetScroll(strNO, Sheet)
End Sub

Friend Sub Act_PaperScroll(ByVal strNO As String, Way As Byte, Percent As Single)
'���ܣ��������������ʱ����
'������Way=0:����1:����,Percent:��ǰ����λ�ðٷֱ�,���Ը���
    RaiseEvent PaperScroll(strNO, Way, Percent)
End Sub

Friend Sub Act_PaperMouseUp(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
'���ܣ�����ֽ��������ɿ�ʱ����
    RaiseEvent PaperMouseUp(strNO, Button, Shift, X, Y)
End Sub

Friend Sub Act_PaperMouseMove(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
'���ܣ�����ֽ��������ƶ�ʱ����
    RaiseEvent PaperMouseMove(strNO, Button, Shift, X, Y)
End Sub

Friend Sub Act_PaperMouseDown(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, blnPop As Boolean)
'���ܣ�����ֽ������갴��ʱ����
    RaiseEvent PaperMouseDown(strNO, Button, Shift, X, Y, blnPop)
End Sub

Friend Sub Act_SheetDblClick(ByVal strNO As String, Sheet As Object, frmParent As Object)
'���ܣ�������������갴��ʱ����
    RaiseEvent SheetDblClick(strNO, Sheet, frmParent)
End Sub

Friend Sub Act_SheetMouseUp(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object)
'���ܣ���������������ɿ�ʱ����
    RaiseEvent SheetMouseUp(strNO, Button, Shift, X, Y, Sheet, frmParent)
End Sub

Friend Sub Act_SheetMouseMove(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object)
'���ܣ���������������ƶ�ʱ����
    RaiseEvent SheetMouseMove(strNO, Button, Shift, X, Y, Sheet, frmParent)
End Sub

Friend Sub Act_SheetMouseDown(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object)
'���ܣ�������������갴��ʱ����
    RaiseEvent SheetMouseDown(strNO, Button, Shift, X, Y, Sheet, frmParent)
End Sub

Friend Sub Act_ReportActive(ByVal strNO As String, Form As Object)
'���ܣ��ڱ������弤��ʱ����
'���������Ըı�Formֵ���ñ���
    RaiseEvent ReportActive(strNO, Form)
End Sub

Friend Sub Act_AfterReportLoad(ByVal strNO As String, Form As Object)
'���ܣ���ȡ��������ʱ����
'���������Ըı�Formֵ���ñ���
    RaiseEvent AfterReportLoad(strNO, Form)
End Sub

Friend Sub Act_BeforeReportLoad(ByVal strNO As String, Form As Object)
'���ܣ���ʼ��ȡ����ʱ����
'���������Ըı�Formֵ���ñ���
    RaiseEvent BeforeReportLoad(strNO, Form)
End Sub

Friend Sub Act_EnterCol(ByVal strNO As String, Col As Integer, Sheet As Object)
'���ܣ�����ĳ��������ʱ����
'���������Ըı�Sheetֵ���ñ���
    RaiseEvent EnterCol(strNO, Col, Sheet)
End Sub

Friend Sub Act_EnterRow(ByVal strNO As String, Row As Integer, RowText As String, Sheet As Object)
'���ܣ�����ĳ��������ʱ����,RowText="����|����|....."
'���������Ըı�Sheetֵ���ñ���
    RaiseEvent EnterRow(strNO, Row, RowText, Sheet)
End Sub

Friend Sub Act_LevelCell(ByVal strNO As String, Row As Integer, Col As Integer, Text As String)
'���ܣ��뿪ĳ������Ԫ��ʱ����
    RaiseEvent LevelCell(strNO, Row, Col, Text)
End Sub

Friend Sub Act_EnterCell(ByVal strNO As String, Row As Integer, Col As Integer, Text As String)
'���ܣ�����ĳ������Ԫ��ʱ����
'���������Ըı�Row,Col,Textֵ���ñ���
    RaiseEvent EnterCell(strNO, Row, Col, Text)
End Sub

Friend Sub Act_ReportUnload(ByVal strNO As String, Form As Object)
'���ܣ���ȡ��������ʱ����
'���������Ըı�Formֵ���ñ���
    RaiseEvent ReportUnload(strNO, Form)
End Sub

Friend Sub Act_CommitCondition(ByVal strNO As String, strCond As String, frmParent As Object)
'���ܣ�����ȷ�����ѯǰ����
'������
'   strNO=��ǰ�򿪵ı������
'   strCond=����DLL����ó��������Ļ���ר��,��ʽΪ"������=����ֵ|������=����ֵ...."
'           ����������ж����ʽ,�����"ReportFormat=x",����ʱҲӦ�ô���,��Ȼ�ͱ��ֵ�ǰ��ʽ����
'   frmParent=������(��������)
    RaiseEvent CommitCondition(strNO, strCond, frmParent)
End Sub

Friend Sub Act_ResetCondition(ByVal strNO As String, strCond As String, blnInhere As Boolean, frmParent As Object)
'���ܣ���������ʱ����
'������
'   strNO=��ǰ�򿪵ı������
'   strCond=����DLL����ó��������Ļ���ר��,��ʽΪ"������=����ֵ|������=����ֵ...."
'           ����������ж����ʽ,�����"ReportFormat=x",����ʱҲӦ�ô���,��Ȼ�ͱ��ֵ�ǰ��ʽ����
'   blnInHere=�Ƿ����DLL��ԭ�е��������봰��
'   frmParent=������(��������)
    RaiseEvent ResetCondition(strNO, strCond, blnInhere, frmParent)
End Sub
'===================================================================================================

Public Function ReportMan(ByVal cnMain As ADODB.Connection, Optional frmParent As Object _
    , Optional strUser As String, Optional ByRef frmDBConnect As Object) As Boolean
'���ܣ��Զ��屨������
'������
'     cnMain:����������ݿ�����
'     frmParent:������
'     strUser:�������߽���ĵ�¼�û���
'     frmDBConnect���������ӹ���
'���أ������Ƿ�ɹ�

    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    If Not ChartInstall Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    gstrLoginUser = strUser
    gstrLoginUserName = GetTheUserName(strUser)
    gstrComputerName = ComputerName
    glngSessionID = GetOracleSessionID()
    Set gfrmMain = frmParent
        
    Call InitPar
    gblnAutoConnect = GetAutoConnect()
    
    '��ȡ�û�Ȩ�޶���(ֻ����һ��)
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Function
    If grsObject.State = adStateClosed Then
        Set grsObject = Nothing
        Set grsObject = UserObject
        If grsObject Is Nothing Then Exit Function
    End If
    
    '��ʼ���������ӵļ�¼������
    gblnManagementTool = True
    Call mdlPublic.SetControlDBConnect(grsConnect)
    
    '�������ӹ�������
    Set gfrmDBConnect = frmDBConnect
    
    If Not frmParent Is Nothing Then
        Load frmMain
        Err.Clear
        On Error Resume Next
        frmMain.Show , frmParent
        If Err.Number <> 0 Then
            Err.Clear
            frmMain.Show 1
        End If
    Else
        frmMain.Show
    End If
    
    ReportMan = True
End Function

Public Function ReportImport(ByVal strFile As String, ByVal cnMain As ADODB.Connection, _
    Optional ByVal strNum As String, Optional ByVal blnOnlyData As Boolean) As Boolean
'���ܣ�����һ�������ļ�
'������
'      strFile:Ҫ����ı����ļ���
'      cnMain:���ó�������ݿ�����
'      strNum:���ǹ̶�����ʱ��,Ҫ���ǵı������
'      blnOnlyData:���ǹ̶�����ʱ��,�Ƿ�ֻ��������Դ
'���أ������Ƿ�ɹ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    
    If strNum <> "" Then
        strSQL = "Select ID From zlReports Where ���=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "ReportImport", strNum)
        If rsTmp.EOF Then Exit Function
        ReportImport = ImportReport(strFile, rsTmp!ID, blnOnlyData) <> ""
    Else
        ReportImport = ImportReport(strFile) <> ""
    End If
End Function

Public Function ReportOpen(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, ParamArray arrPar() As Variant) As Boolean
'���ܣ������������ĳ���Ż򱨱���Ŵ򿪱�����
'������
'     lngSys=����������ϵͳ,��Ϊ����ϵͳ����ʱ,����Ϊ0
'     varReport:��Ҫִ�еı�����Ż��Ӧ�ĳ���ID(������Ϊ�����˵�����ʱ��ʹ�ó���ID)
'     cnMain:����������ݿ�����
'     frmParent:������
'     arrPar():1.��������,��ʽΪ"������1=����ֵ1"��"��������Դ1"...,���Ƿ��еȺ������Ǳ����������Ǹ��ӱ������ݡ�
'              2.����ѡ�������½ӿ����͵Ĳ���,���Ҫ��ֱ��ָ����ʾֵ/��ֵ,������"������=��ʾֵ|��ֵ"�ĸ�ʽ����
'                Ҳ��ֻ����"������=��ֵ"�ķ�ʽ��������Զ�����ȥѰ�Ҷ�Ӧ����ʾֵ��
'              3.�����ж����ʽ�ı���,�����ڲ����д���"ReportFormat=x",��ȷ��Ҫ�򿪵ı�����ʽ,�Ա�������Ч
'              4.��ֱ�Ӵ�ӡʱ,�����ڲ����д���"PrintEmpty=1/0",������ǰ��ʽ�е����б���������ʱ,�Ƿ���Դ�ӡ
'              5.�����ȷ��ӡ��Ԥ��ʱ��Ҫ�����״����ڲ����д���"PressWork=1"����"PressWorkFirst=1".
'                ��ʾȫ������ҳ�״�������ֻ��Ա����е����ݽ���������Ա�ͷ�������߼�����Ԫ�ز��������
'              6.�����ӡҳ��ʱҪȷ����ʼ�����׼ҳ���ɴ���"StartPageNum=x"(x>=1)������[ҳ��]��ǩ���ڴ˻����ϼ��㡣
'              7.��󸽼Ӳ���:0=ȱʡֵ,�ɲ���,��ʾ����(��������Ԥ��),1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel,4-�����PDF
'              8.�����Excelʱ������ͨ��"ExcelFile=***.xls"������ָ������ļ���ָ��������ļ�ʱ�����Զ���
'              9.���򿪵��Ǳ�����ʱ,���Խ����б����Ĳ������δ���
'���أ������Ƿ�ɹ�
'˵����1.��ʹ�ó���Ŵ򿪱���ʱ,���������Ѿ�����������δ�����ı���ʱ,�����ֹ������ݿ�������Ȩ�ޡ�
'      2.����������ɱ�������ʱ����,ÿ����ʽΪ"������=����ֵ",��:"��������='��һ��','������'","��ʼʱ��=2000-01-02","����=����","����ID=23"��
'      3.���ַ�ʽ����������������Ʋ������봰��,�Լ�ǿʹ�����������������Լ�������(���������Ա����ʱ,���Լ�һ����Աѡ����)��
'      4.��Ϊ�����ж������Դ,ÿ������Դ���ж������,��Щ��������Ψһ,���ȴ���ҵ�,���Դ�����ʱҪ������Ϊ����
'      5.������û�д���û�д���(����������ʽ����),��ϵͳ���Զ������������봰���δ�������������롣
    Dim bytStyle As Byte, blnCache As Boolean, intTmp As Integer
    
    On Error GoTo errH
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    If Not ChartInstall Then Exit Function
    
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    Set gfrmMain = frmParent
    glngSys = lngSys
    glngSessionID = GetOracleSessionID()
    mblnDataIsEmpty = False
    
    Call InitPar
    
    '��������ʾ��ʽ
    garrPars = arrPar
    If UBound(arrPar) <> -1 Then
        If IsNumeric(arrPar(UBound(arrPar))) Then
            bytStyle = CByte(arrPar(UBound(arrPar)))
        End If
    End If
    If bytStyle = 3 Then
        If zlRegInfo("��Ȩ����") <> "1" Then
            MsgBox "���û���԰汾����ʹ�øù��ܡ�", vbInformation, App.Title
            Exit Function
        End If
    End If
    
    '��ʼ���������ӵļ�¼������
    If grsConnect Is Nothing Then
        Call mdlPublic.SetControlDBConnect(grsConnect)
    Else
        If grsConnect.State <> adStateOpen Then
            Call mdlPublic.SetControlDBConnect(grsConnect)
        End If
    End If
    
    gblnError = False
    
    If isGroup(lngSys, varReport, glngGroup) Then
        '�򿪱�����
        If glngGroup = 0 Then Exit Function
        Set gobjReport = Nothing
        ReportOpen = ShowReport(frmParent, Me, bytStyle)
    Else
        '�Ϸ��Լ��
        If Not CheckPass(glngGroup) Then
            MsgBox "����:[" & grsReport!��� & "]" & grsReport!���� & " ���ݴ��󣬲���ִ�иñ�����", vbInformation, App.Title
            Exit Function
        End If
        'Ȩ���ж�
        intTmp = ReportCanQuery(glngGroup)
        If intTmp <> 0 Then
            If intTmp = 1 Then '����������ʾ,Ʊ�ݲ���ʾ
                MsgBox "��û��Ȩ��ִ�б���:[" & grsReport!��� & "]" & grsReport!���� & "����ȷ��ϵͳ�Ƿ���ȷ��Ȩ��", vbInformation, App.Title
            End If
            Exit Function
        End If
        
        '��ȡ����
        If Not gobjReport Is Nothing And ReportReaded(glngGroup) Then
            If gobjReport.��� = grsReport!��� And gobjReport.ϵͳ = Nvl(grsReport!ϵͳ, 0) _
                And Format(gobjReport.�޸�ʱ��, "yyyy-MM-dd HH:mm:ss") = Format(grsReport!�޸�ʱ��, "yyyy-MM-dd HH:mm:ss") Then
                blnCache = True
            End If
        End If
        If Not blnCache Then
            Set gobjReport = Nothing
            Set gobjReport = ReadReport(glngGroup)
        End If
        
        '�򿪱���
        glngGroup = 0
        If gobjReport Is Nothing Then Exit Function
        ReportOpen = ShowReport(frmParent, Me, bytStyle)
    End If
    
    ReportOpen = ReportOpen And Not gblnError
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ReportTaskBegin()
'���ܣ���ʼ��ӡ�����ڶ౨���ڵ�һ�����д�ӡ
'˵������Ҫ���ȵ��øú���������ReportOpen���ҵ���ReportOpenʱһ��Ҫ�Դ�ӡ��ʽ
    If Not gblnSingleTask Then
        gblnSingleTask = True
    End If
End Sub

Public Sub ReportTaskEnd()
'���ܣ�������ӡ�����ڶ౨���ڵ�һ�����д�ӡ
'˵����������ReportOpen֮�����ŵ��øú���
    If gblnSingleTask Then
        Printer.EndDoc
        gblnSingleTask = False
    End If
End Sub

Public Function ReportPrintSet(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, Optional frmParent As Object) As Boolean
'���ܣ����ñ��ر�����ӡ��(��ӡʱ��ӡ��˳�򣺱���-������-��ǰȱʡ)
'������
'     varReport:��Ҫ���õı����������(����ID)�򱨱����
'     cnMain:����������ݿ�����
'     frmParent:������
'���أ������Ƿ�ɹ�
    Dim lngGroup As Long, intTmp As Integer
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    glngSessionID = GetOracleSessionID()
    
    If Not isGroup(lngSys, varReport, lngGroup) Then
        'Ȩ���ж�
        intTmp = ReportCanQuery(lngGroup)
        If intTmp <> 0 Then
            If intTmp <> 3 Then
                MsgBox "��û���㹻��Ȩ��ִ�и����������ȷ���Ƿ���ȷ��Ȩ��", vbInformation, App.Title
            End If
            Exit Function
        End If
    End If
    
    ReportPrintSet = ReportLocalSet(lngSys, varReport, True, , frmParent)
    If Not frmParent Is Nothing Then frmParent.Refresh
End Function

Public Function ReportDesign(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal varReport As Variant, frmParent As Object, Optional blnModiData As Boolean) As Boolean
'���ܣ���ָ����������ʾ��������
'������lngSys=ϵͳ���,����������0
'      varReport:�������
'      frmParent=���ø�����
'      blnModiData=�Ƿ������޸�����Դ(ȱʡΪ��)
'���أ������Ƿ��޸�
    Dim rsTmp As New ADODB.Recordset
    Dim frmNew As frmDesign
    Dim strSQL As String, i As Integer
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    If Not ChartInstall Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    glngSys = lngSys
    
    Call InitPar
        
    '��ȡ�û�Ȩ�޶���(ֻ����һ��)
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Function
    If grsObject.State = adStateClosed Then
        Set grsObject = Nothing
        Set grsObject = UserObject
        If grsObject Is Nothing Then Exit Function
    End If
        
    '��ȡ������Ϣ
    If TypeName(varReport) = "String" Then
        strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ�� From zlReports Where Nvl(ϵͳ,0)=[3] And ���=[1]"
    Else
        strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ�� From zlReports Where Nvl(ϵͳ,0)=[3] And ����ID=[2]"
    End If
    'Set rsTmp = New ADODB.Recordset
    Set rsTmp = OpenSQLRecord(strSQL, "ReportDesign", UCase(varReport), Val(varReport), lngSys)
    If rsTmp.EOF Then Exit Function
    
    '--
    Set frmNew = New frmDesign
    frmNew.lngRPTID = rsTmp!ID
    frmNew.mblnNotModiData = Not blnModiData
    
    On Error Resume Next
    frmNew.Show 1, frmParent
    On Error GoTo 0
    
    ReportDesign = gblnModi
End Function

Public Sub ReportShow(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal varReport As Variant, objParent As Object)
'���ܣ���ָ����������ʾ��������
'������lngSys=ϵͳ���,����������0
'      varReport:�������
'      objParent=Ҫ��ʾ����������(һ��ΪPictureBox��Form)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    If cnMain Is Nothing Then Exit Sub
    If cnMain.State = adStateClosed Then Exit Sub
    If Not ChartInstall Then Exit Sub
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    glngSessionID = GetOracleSessionID()
    
    '��ȡ������Ϣ
    If TypeName(varReport) = "String" Then
        strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ�� From zlReports Where Nvl(ϵͳ,0)=[3] And ���=[1]"
    Else
        strSQL = "Select ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ�� From zlReports Where Nvl(ϵͳ,0)=[3] And ����ID=[2]"
    End If
    'Set rsTmp = New ADODB.Recordset
    Set rsTmp = OpenSQLRecord(strSQL, "ReportDesign", UCase(varReport), Val(varReport), lngSys)
    If rsTmp.EOF Then Exit Sub
    
    'ȷ��һ��δʹ�õ�λ��
    For i = 0 To UBound(arrForms)
        If arrForms(i) Is Nothing Then Exit For
    Next
    If i > UBound(arrForms) Then
        ReDim Preserve arrForms(UBound(arrForms) + 1)
        i = UBound(arrForms)
    End If
    
    '��ʾ����
    Set arrForms(i) = New frmShowDesign
    arrForms(i).mlngRPTID = rsTmp!ID
    Load arrForms(i)
    
    SetParent arrForms(i).hwnd, objParent.hwnd
    
    arrForms(i).Tag = varReport
    
    '������API��λ
    MoveWindow arrForms(i).hwnd, 0, 0, arrForms(i).ScaleX(objParent.ScaleWidth, vbTwips, vbPixels), arrForms(i).ScaleY(objParent.ScaleHeight, vbTwips, vbPixels), 0
    
    arrForms(i).Show
End Sub

Public Sub ReportWindowResize(ByVal varReport As Variant, ByVal W As Long, ByVal H As Long)
'���ܣ��������ߴ緢���仯ʱ,ͬ�����ı�����ʾ�ĳߴ�
'������varReport:�������
'      W,H=�µĸ߶�,����(vbTwips Units)
    Dim i As Integer
    For i = 0 To UBound(arrForms)
        If Not arrForms(i) Is Nothing Then
            If arrForms(i).Tag = varReport Then
                MoveWindow arrForms(i).hwnd, 0, 0, arrForms(i).ScaleX(W, vbTwips, vbPixels), arrForms(i).ScaleY(H, vbTwips, vbPixels), 0
            End If
        End If
    Next
End Sub

Public Sub ReportWindowUnload(ByVal varReport As Variant)
'���ܣ�ж��ָ������ʾ�����Ĵ���
'������varReport:�������
    Dim i As Integer
    For i = 0 To UBound(arrForms)
        If Not arrForms(i) Is Nothing Then
            If arrForms(i).Tag = varReport Then
                Unload arrForms(i)
                Set arrForms(i) = Nothing
            End If
        End If
    Next
End Sub

Public Function ReportNew(cnMain As ADODB.Connection, lngSys As Long, strNum As String, strName As String, strSQL As String, ParamArray arrPar() As Variant) As Boolean
'���ܣ�����һ�Ź̶�δ��������
'������lngSys=ϵͳ���
'      strNum=�������
'      strName=��������
'      strSQL=����Դ,���п��ܰ�����������д
'      arrPar=�������Դ�а�������,�򴫾������������ʽΪ"������;����;ȱʡֵ",����ȱʡֵ��Ϊ��,˳����������Դ�ж���Ĳ������˳��
'���أ��������Դ������ȷ����,��᷵��ʧ�ܡ�
'˵����1.ȱʡֽ��ΪA4,����,��ӡ���ڴ�ӡǰ������غ������á�
'      2.����ȱʡ�ı���Ϊ��������
'      3.���ݲ�������˳��,������������Ϊ������Ŀ��֯
'      4.��������Դ������֯���������
    Dim blnTran As Boolean, lngNewID As Long
    Dim objReport As Report, objPars As RPTPars
    Dim strObject As String, strField As String
    Dim i As Integer, j As Integer, k As Integer
    Dim strTmp As String, strR As String
    
    On Error GoTo errH
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    
    '��ȡ�û�Ȩ�޶���(ֻ����һ��)
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Function
    
    '���ɱ�������
    Set objReport = New Report
    objReport.��� = strNum
    objReport.���� = strName
    objReport.��ֽ = 15
    
    '����Դ
    '����
    Set objPars = New RPTPars
    If UBound(arrPar) <> -1 Then
        For i = 0 To UBound(arrPar)
            strTmp = CStr(arrPar(i))
            objPars.Add "", CByte(i), CStr(Split(strTmp, ";")(0)), CByte(Split(strTmp, ";")(1)), CStr(Split(strTmp, ";")(2)), 0, "", "", "", "", "", "", "_" & i
        Next
    End If
    
    '��������
    strObject = SQLObject(strSQL)
    If strObject = "" Then Exit Function
    strObject = ObjectOwner(strObject)
    If strObject = "ȡ��" Then Exit Function 'ȡ������
    
    '�����ֶ�
    strTmp = SQLOwner(RemoveNote(strSQL), strObject)
    If GetParCount(strTmp) = 0 Then
        strField = CheckSQL(strTmp, strR)
    Else
        strField = CheckSQL(strTmp, strR, ReplaceParSysNo(objPars, lngSys))
    End If
    If strField = "" Then Exit Function
    
    objReport.Datas.Add "��������", 0, strSQL, strField, strObject, 0, "", objPars, "_��������"
    
    '��ʽ
    objReport.Fmts.Add 1, strName, INIT_WIDTH, INIT_HEIGHT, 9, 1, False, 0, "_1"
    
    'Ԫ��
    '������
    j = 1
    objReport.Items.Add j, 1, "����", 0, 4, 0, "", 0, "��������", "", Twip_mm * 10, Twip_mm * 30, INIT_WIDTH - Twip_mm * 20, INIT_HEIGHT - Twip_mm * 40, 285, 0, False, "����", 9, False, False, False, 0, 0, vbWhite, False, 1, "", "", "", False, , , , , , "_" & j
    For i = 0 To UBound(Split(strField, "|"))
        j = j + 1
        Select Case CLng(Split(Split(strField, "|")(i), ",")(1))
            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                k = 2
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                k = 0
            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                k = 0
            Case Else
                k = 0
        End Select
        objReport.Items.Add j, 1, "", 1, 6, i, "", 0, "[��������." & Split(Split(strField, "|")(i), ",")(0) & "]", "4^285^" & Split(Split(strField, "|")(i), ",")(0), 0, 0, 1000, 0, 0, CByte(k), False, "", 0, False, False, False, 0, 0, 0, False, 0, "", "", "", False, , , , , , "_" & j
        objReport.Items("_1").SubIDs.Add j, "_" & j
    Next
    '����
    j = j + 1
    objReport.Items.Add j, 1, "����", 0, 2, 0, "����", 12, "[��λ����]" & strName, "", Twip_mm * 10, Twip_mm * 10, Twip_mm * 10, Twip_mm * 10, 0, 1, True, "����", 20, False, False, False, 0, 0, vbWhite, False, 0, 0, 0, 0, False, , True, , , , "_" & j
    
    lngNewID = GetNextID("zlReports")
    
    '���뱨��
    blnTran = True
    gcnOracle.BeginTrans
    gcnOracle.Execute "Insert Into zlReports(ID,���,����,ϵͳ,����) Values(" & lngNewID & ",'" & strNum & "','" & strName & "'," & lngSys & "," & AdjustStr(GetPass(strNum, strName)) & ")"
    gcnOracle.CommitTrans
    blnTran = False
    
    '����汨������
    ReportNew = SaveReport(lngNewID, objReport)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Public Function ReportDel(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strNum As String) As Boolean
'���ܣ�ɾ��һ�ű���
'������cnMain=���ݿ�����,strNum=�������
    Dim strSQL As String
    
    On Error GoTo errH
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
        
    strSQL = "Delete From zlReports Where Upper(���)='" & UCase(strNum) & "' And Nvl(ϵͳ,0)=" & lngSys
    cnMain.BeginTrans
    cnMain.Execute strSQL
    cnMain.CommitTrans
    
    ReportDel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    cnMain.RollbackTrans
    Call SaveErrLog
End Function

Public Sub ReportAbout(Optional frmParent As Object)
'���ܣ���ʾ�Զ��屨�����ڴ���
'˵�������øýӿ�ǰһ��ע���ȵ���InitOracle
    Call ShowAbout(frmParent)
End Sub

Public Function GenReportPass(ByVal strCode As String, ByVal strName As String) As String
'���ܣ����ݱ����ı�ż����Ʒ��غϷ�������
'˵�������øýӿ�ǰһ��ע���ȵ���InitOracle
    GenReportPass = GetPass(strCode, strName)
End Function

Public Sub InitOracle(ByVal cnMain As ADODB.Connection)
'���ܣ����Զ��屨�������ݿ����ӽ��г�ʼ��
    Set gcnOracle = cnMain
End Sub

Private Sub Class_Initialize()
    mblnDataIsEmpty = False
    
    gstrFind = MakeFile("Find", "AVI")
    Call RegReportFile
    
    Call InitEnv    '��ʼ�����ܴ�
    Call GetCompare '��ԭ���ܴ�
    
    '��ʼ��������ʾ����
    arrForms = Array()
    
    Set gclsCNs = New RPTDBCNs
    Set grsConnect = New ADODB.Recordset
    Set gfrmDBConnect = Nothing
    gblnManagementTool = False
End Sub

Private Sub Class_Terminate()
    Dim i As Integer
    
    '�Զ�ж�ر�����ʾ����
    For i = 0 To UBound(arrForms)
        If Not arrForms(i) Is Nothing Then
            Unload arrForms(i)
            Set arrForms(i) = Nothing
        End If
    Next
    Unload frmFlash
    
    Set grsConnect = Nothing
    Set gclsCNs = Nothing
    
    'Set gcnOracle = Nothing
    'Set grsObject = Nothing
    
    On Error Resume Next
    Kill gstrFind
    Err.Clear
    On Error GoTo 0
    
    mblnDataIsEmpty = False
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.count = 0)
End Function

Public Function LoadReport(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, objForm As Object, LibDatas As Object, ParamArray arrPar() As Variant) As Boolean
'���ܣ������������ĳ���Ż򱨱���ż��ر�����
'������
'     lngSys=����������ϵͳ,��Ϊ����ϵͳ����ʱ,����Ϊ0
'     varReport:��Ҫִ�еı�����Ż��Ӧ�ĳ���ID(������Ϊ�����˵�����ʱ��ʹ�ó���ID)
'     cnMain:����������ݿ�����
'     frmParent:������
'     objForm�����صı���Ԥ�����壬����Ƕ��ʽԤ��
'     arrPar():1.��������,��ʽΪ"������1=����ֵ1"��"��������Դ1"...,���Ƿ��еȺ������Ǳ����������Ǹ��ӱ������ݡ�
'              2.����ѡ�������½ӿ����͵Ĳ���,���Ҫ��ֱ��ָ����ʾֵ/��ֵ,������"������=��ʾֵ|��ֵ"�ĸ�ʽ����
'                Ҳ��ֻ����"������=��ֵ"�ķ�ʽ��������Զ�����ȥѰ�Ҷ�Ӧ����ʾֵ��
'              3.�����ж����ʽ�ı���,�����ڲ����д���"ReportFormat=x",��ȷ��Ҫ�򿪵ı�����ʽ,�Ա�������Ч
'              4.��ֱ�Ӵ�ӡʱ,�����ڲ����д���"PrintEmpty=1/0",������ǰ��ʽ�е����б���������ʱ,�Ƿ���Դ�ӡ
'              5.�����ȷ��ӡ��Ԥ��ʱ��Ҫ�����״����ڲ����д���"PressWork=1"����"PressWorkFirst=1".
'                ��ʾȫ������ҳ�״�������ֻ��Ա����е����ݽ���������Ա�ͷ�������߼�����Ԫ�ز��������
'              6.�����ӡҳ��ʱҪȷ����ʼ�����׼ҳ���ɴ���"StartPageNum=x"(x>=1)������[ҳ��]��ǩ���ڴ˻����ϼ��㡣
'              7.��󸽼Ӳ���:0=ȱʡֵ,�ɲ���,��ʾ����(��������Ԥ��),1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3-�����Excel,4-�����PDF
'              8.�����Excelʱ������ͨ��"ExcelFile=***.xls"������ָ������ļ���ָ��������ļ�ʱ�����Զ���
'              9.���򿪵��Ǳ�����ʱ,���Խ����б����Ĳ������δ���
'���أ������Ƿ�ɹ�
'˵����1.��ʹ�ó���Ŵ򿪱���ʱ,���������Ѿ�����������δ�����ı���ʱ,�����ֹ������ݿ�������Ȩ�ޡ�
'      2.����������ɱ�������ʱ����,ÿ����ʽΪ"������=����ֵ",��:"��������='��һ��','������'","��ʼʱ��=2000-01-02","����=����","����ID=23"��
'      3.���ַ�ʽ����������������Ʋ������봰��,�Լ�ǿʹ�����������������Լ�������(���������Ա����ʱ,���Լ�һ����Աѡ����)��
'      4.��Ϊ�����ж������Դ,ÿ������Դ���ж������,��Щ��������Ψһ,���ȴ���ҵ�,���Դ�����ʱҪ������Ϊ����
'      5.������û�д���û�д���(����������ʽ����),��ϵͳ���Զ������������봰���δ�������������롣
    Dim bytStyle As Byte
    Dim blnCache As Boolean, blnDo As Boolean
    Dim intTmp As Integer, intState As Integer
    Dim strInfo As String, strGroup As String
    Dim objfrmShow   As Object
    
    On Error GoTo errH
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    If Not ChartInstall Then Exit Function
    
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    
    Set gfrmMain = frmParent
    glngSys = lngSys
    
    mblnDataIsEmpty = False
    
    Call InitPar
    
    '��������ʾ��ʽ
    garrPars = arrPar
    If UBound(arrPar) <> -1 Then
        If IsNumeric(arrPar(UBound(arrPar))) Then
            bytStyle = CByte(arrPar(UBound(arrPar)))
        End If
    End If
    If bytStyle = 3 Then
        If zlRegInfo("��Ȩ����") <> "1" Then
            MsgBox "���û���԰汾����ʹ�øù��ܡ�", vbInformation, App.Title
            Exit Function
        End If
    End If
    
    '��ʼ���������ӵļ�¼������
    If grsConnect Is Nothing Then
        Call mdlPublic.SetControlDBConnect(grsConnect)
    Else
        If grsConnect.State <> adStateOpen Then
            Call mdlPublic.SetControlDBConnect(grsConnect)
        End If
    End If
    
    gblnError = False
    
    If isGroup(lngSys, varReport, glngGroup) Then
        '�򿪱�����
        If glngGroup = 0 Then Exit Function
        Set gobjReport = Nothing
        LoadReport = ShowReport(frmParent, Me, bytStyle)
    Else
        '�Ϸ��Լ��
        If Not CheckPass(glngGroup) Then
            MsgBox "����:[" & grsReport!��� & "]" & grsReport!���� & " ���ݴ��󣬲���ִ�иñ�����", vbInformation, App.Title
            Exit Function
        End If
        'Ȩ���ж�
        intTmp = ReportCanQuery(glngGroup)
        If intTmp <> 0 Then
            If intTmp = 1 Then '����������ʾ,Ʊ�ݲ���ʾ
                MsgBox "��û��Ȩ��ִ�б���:[" & grsReport!��� & "]" & grsReport!���� & "����ȷ��ϵͳ�Ƿ���ȷ��Ȩ��", vbInformation, App.Title
            End If
            Exit Function
        End If
        
        '��ȡ����
        If Not gobjReport Is Nothing And ReportReaded(glngGroup) Then
            If gobjReport.��� = grsReport!��� And gobjReport.ϵͳ = Nvl(grsReport!ϵͳ, 0) _
                And Format(gobjReport.�޸�ʱ��, "yyyy-MM-dd HH:mm:ss") = Format(grsReport!�޸�ʱ��, "yyyy-MM-dd HH:mm:ss") Then
                blnCache = True
            End If
        End If
        If Not blnCache Then
            Set gobjReport = Nothing
            Set gobjReport = ReadReport(glngGroup)
        End If
        
        '�򿪱���
        glngGroup = 0
        If gobjReport Is Nothing Then Exit Function
        LoadReport = GetReportFrom(frmParent, Me, bytStyle, objForm, LibDatas)
    End If
    
    LoadReport = LoadReport And Not gblnError
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReportOpenForRec(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, LibDatas As Object, ParamArray arrPar() As Variant) As Boolean
'���ܣ������������ĳ���Ż򱨱���Ŵ򿪱�����
'������
'     lngSys=����������ϵͳ,��Ϊ����ϵͳ����ʱ,����Ϊ0
'     varReport:��Ҫִ�еı�����Ż��Ӧ�ĳ���ID(������Ϊ�����˵�����ʱ��ʹ�ó���ID)
'     cnMain:����������ݿ�����
'     frmParent:������
'     LibDatas������ļ�¼�������ڴ�ӡ
'���أ������Ƿ�ɹ�
'˵����1.��ʹ�ó���Ŵ򿪱���ʱ,���������Ѿ�����������δ�����ı���ʱ,�����ֹ������ݿ�������Ȩ�ޡ�
'      2.����������ɱ�������ʱ����,ÿ����ʽΪ"������=����ֵ",��:"��������='��һ��','������'","��ʼʱ��=2000-01-02","����=����","����ID=23"��
'      3.���ַ�ʽ����������������Ʋ������봰��,�Լ�ǿʹ�����������������Լ�������(���������Ա����ʱ,���Լ�һ����Աѡ����)��
'      4.��Ϊ�����ж������Դ,ÿ������Դ���ж������,��Щ��������Ψһ,���ȴ���ҵ�,���Դ�����ʱҪ������Ϊ����
'      5.������û�д���û�д���(����������ʽ����),��ϵͳ���Զ������������봰���δ�������������롣
    Dim blnCache As Boolean, blnDo As Boolean
    Dim intTmp As Integer, intState As Integer
    Dim strInfo As String, strGroup As String
    Dim bytStyle As Byte
    
    On Error GoTo errH
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    If Not ChartInstall Then Exit Function
    
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    
    Set gfrmMain = frmParent
    glngSys = lngSys
    glngSessionID = GetOracleSessionID()
    mblnDataIsEmpty = False
    
    Call InitPar
    '��������ʾ��ʽ
    garrPars = arrPar
    If UBound(arrPar) <> -1 Then
        If IsNumeric(arrPar(UBound(arrPar))) Then
            bytStyle = CByte(arrPar(UBound(arrPar)))
        End If
    End If
    
    If bytStyle = 3 Then
        If zlRegInfo("��Ȩ����") <> "1" Then
            MsgBox "���û���԰汾����ʹ�øù��ܡ�", vbInformation, App.Title
            Exit Function
        End If
    End If
    
    '��ʼ���������ӵļ�¼������
    If grsConnect Is Nothing Then
        Call mdlPublic.SetControlDBConnect(grsConnect)
    Else
        If grsConnect.State <> adStateOpen Then
            Call mdlPublic.SetControlDBConnect(grsConnect)
        End If
    End If
    
    gblnError = False
    
    If isGroup(lngSys, varReport, glngGroup) Then
        '�򿪱�����
        If glngGroup = 0 Then Exit Function
        Set gobjReport = Nothing
        ReportOpenForRec = ShowReport(frmParent, Me, bytStyle)
    Else
        '�Ϸ��Լ��
        If Not CheckPass(glngGroup) Then
            MsgBox "����:[" & grsReport!��� & "]" & grsReport!���� & " ���ݴ��󣬲���ִ�иñ�����", vbInformation, App.Title
            Exit Function
        End If
        'Ȩ���ж�
        intTmp = ReportCanQuery(glngGroup)
        If intTmp <> 0 Then
            If intTmp = 1 Then '����������ʾ,Ʊ�ݲ���ʾ
                MsgBox "��û��Ȩ��ִ�б���:[" & grsReport!��� & "]" & grsReport!���� & "����ȷ��ϵͳ�Ƿ���ȷ��Ȩ��", vbInformation, App.Title
            End If
            Exit Function
        End If
        
        '��ȡ����
        If Not gobjReport Is Nothing And ReportReaded(glngGroup) Then
            If gobjReport.��� = grsReport!��� And gobjReport.ϵͳ = Nvl(grsReport!ϵͳ, 0) _
                And Format(gobjReport.�޸�ʱ��, "yyyy-MM-dd HH:mm:ss") = Format(grsReport!�޸�ʱ��, "yyyy-MM-dd HH:mm:ss") Then
                blnCache = True
            End If
        End If
        If Not blnCache Then
            Set gobjReport = Nothing
            Set gobjReport = ReadReport(glngGroup)
        End If
        
        '�򿪱���
        glngGroup = 0
        If gobjReport Is Nothing Then Exit Function
        ReportOpenForRec = ShowReportForRec(frmParent, Me, LibDatas, bytStyle)
    End If
    
    ReportOpenForRec = ReportOpenForRec And Not gblnError
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function