VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'Ҫ���������
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ��¼���ݿ��û���
'���أ�
    Set gfrmMain = frmMain: glngMain = 0
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    gstrDBUser = strDbUser
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If

    '-------------------------------------------------
    '1110       �ҺŰ��Ź���
    '1111       ����ҺŹ���
    '1112       �Һ�Ա�ձ�
    '1114       �ٴ����ﰲ��
    '1115       ���߷�������
    Call InitLocPar(glngModul)
    
    Select Case glngModul
        Case 1110                 '�ҺŰ���
            Call InitSysPar
            frmRegistPlan.Show , frmMain
        Case 1111                 '����ҺŹ���
            Call InitSysPar
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(gfrmMain, lngModul)
            If gfrmMain Is Nothing Then
                If gbytRegistMode = 0 Then
                    frmRegist.Show
                Else
                    frmRegistNew.Show
                End If
            Else
                If gbytRegistMode = 0 Then
                    frmRegist.Show , frmMain
                Else
                    frmRegistNew.Show , frmMain
                End If
            End If
        Case 1113
            Call InitSysPar
            Call CreateSquareCardObject(gfrmMain, lngModul)
            frmDistRoomManager.Show , frmMain     '����������
        Case 1114
            If gfrmMain Is Nothing Then
                frmClinicPlanMainManage.Show
            Else
                frmClinicPlanMainManage.Show , frmMain
            End If
        Case 1115
            Call InitSysPar
            Call CreateSquareCardObject(gfrmMain, lngModul)
            If gfrmMain Is Nothing Then
                frmServiceCenter.Show
            Else
                frmServiceCenter.Show , frmMain
            End If
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ��¼���ݿ��û���
'���أ�
    Set gfrmMain = Nothing: glngMain = lngMain
    If glngMain = 0 Then glngMain = -1
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    gstrDBUser = strDbUser
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If

    '-------------------------------------------------
    '1110       �ҺŰ��Ź���
    '1111       ����ҺŹ���
    '1112       �Һ�Ա�ձ�
    '1114       �ٴ����ﰲ��
    '1115       ���߷�������
    Call InitLocPar(glngModul)
    Select Case glngModul
        Case 1110                 '�ҺŰ���
            Call InitSysPar
            zlCommFun.ShowChildWindow frmRegistPlan.Hwnd, lngMain
        Case 1111                 '����ҺŹ���
            Call InitSysPar
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(gfrmMain, lngModul)
            If gbytRegistMode = 0 Then
                zlCommFun.ShowChildWindow frmRegist.Hwnd, lngMain
            Else
                zlCommFun.ShowChildWindow frmRegistNew.Hwnd, lngMain
            End If
        Case 1113                 '����������
            Call InitSysPar
            Call CreateSquareCardObject(gfrmMain, lngModul)
            zlCommFun.ShowChildWindow frmDistRoomManager.Hwnd, lngMain
        Case 1114                 '�ٴ����ﰲ��
            zlCommFun.ShowChildWindow frmClinicPlanMainManage.Hwnd, lngMain
        Case 1115                 '���߷�������
            Call InitSysPar
            Call CreateSquareCardObject(gfrmMain, lngModul)
            zlCommFun.ShowChildWindow frmServiceCenter.Hwnd, lngMain
    End Select
End Sub

Public Function PrintRegistBill(frmMain As Object, cnMain As ADODB.Connection, _
    ByVal lngSys As Long, ByVal strDbUser As String, _
    ByVal strNO As String, ByVal lng����ID As Long, _
    Optional ByVal blnFirstNO As Boolean) As Boolean
    '���ܣ�����Һ�Ʊ�ݣ������շѹ����еİ����˲���Ʊ�ݵ��ã����ŵ���ʱ��ֱ����
    '��Σ�
    '   blnFirstNO:�Ƿ��һ�ŵ���
    '���أ���ӡ�ɹ�����True����ӡʧ�ܷ���False
    Dim blnStartFactUseType  As Boolean, strUseType As String
    Dim intInsure As Integer, blnVirtualPrint As Boolean
    
    If blnFirstNO Then
        Set gfrmMain = frmMain
        If Not CheckValid Then Exit Function
        
        Set gcnOracle = cnMain
        glngSys = lngSys
        gstrDBUser = strDbUser
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        gstrUnitName = GetUnitName
        If GetUserInfo() = False Then Exit Function
        
        gstrPrivs = ";" & GetPrivFunc(lngSys, 1111) & ";"
        If zlStr.IsHavePrivs(gstrPrivs, "����Ʊ��") = False Then
            MsgBox "��û�б�����ʹ�ùҺŵĲ���Ʊ�ݹ��ܵ�Ȩ��,����ʹ��,����ϵͳ����Ա��ϵ!", vbInformation, gstrSysName
            Exit Function
        End If
        
        Call InitLocPar(1111)
        Call InitSysPar
        Call gclsInsure.InitOracle(gcnOracle)
    End If
    
    intInsure = ExistInsure(strNO)
    If intInsure <> 0 Then
        blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure)
    End If
    
    If gblnSharedInvoice Then
        '�Һ�������Ʊ��:42703
        blnStartFactUseType = zlStartFactUseType("1")
        If blnStartFactUseType Then
            strUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
        End If
    End If
    '��һ�ŵ���ʱ��Ҫȷ�Ϸ�Ʊ�ţ���������ȷ�Ϸ�Ʊ��
    If Not RePrintBill(frmMain, 4, strNO, 0, 0, blnVirtualPrint, strUseType, True, blnFirstNO) Then Exit Function
    PrintRegistBill = True
End Function

Public Function Register(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, _
    ByVal strDbUser As String, ByVal blnSharedInvoice As Boolean, ByVal lngInvoiceGroupID As Long) As Boolean
'���ܣ��շ�ֱ�ӵ��ùҺŹ���
'���أ��Ƿ�ҹ���
    Static blnNoFirst As Boolean  '�Ƿ��һ�ε���,��ʼ��ȫ�ֱ���
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, 1111) & ";"
    If Not blnNoFirst Then
        Set gfrmMain = frmMain: glngMain = 0
        If Not CheckValid Then Exit Function
        
        Set gcnOracle = cnMain
        glngSys = lngSys
        gstrDBUser = strDbUser
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        gstrUnitName = GetUnitName
        
        If InStr(";" & gstrPrivs & ";", ";���շѺ�;") = 0 And InStr(";" & gstrPrivs & ";", ";����Ѻ�;") = 0 Then
            MsgBox "��û�б�����ʹ�ùҺŹ��ܵ�Ȩ��,����ʹ��,����ϵͳ����Ա��ϵ!", vbInformation, gstrSysName
            Exit Function
        End If
        
        'ҽ���ӿ�
        Call gclsInsure.InitOracle(gcnOracle)
        '�û���Ϣ
        Call GetUserInfo
        '��ʼ������
        Call InitLocPar(1111)
        Call InitSysPar
        Call CreateSquareCardObject(gfrmMain, 1111)
        blnNoFirst = True
    End If
    gblnSharedInvoice = blnSharedInvoice
    If gblnSharedInvoice Then glng�Һ�ID = lngInvoiceGroupID
    
    '����������Ʊ�ݴ�ӡ����
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, 1111, UserInfo.���, UserInfo.����)
    End If
    
    '���ùҺ�
    Err.Clear
    On Error Resume Next
    If gbytRegistMode = 0 Then
        frmRegistEdit.mblnCharge = True
        frmRegistEdit.mlngModul = 1111
        frmRegistEdit.mstrPrivs = gstrPrivs
        frmRegistEdit.mbytMode = 0
        frmRegistEdit.mbytInState = 0
        If gfrmMain Is Nothing Then
            frmRegistEdit.Show
        Else
            frmRegistEdit.Show 1, gfrmMain
        End If
    Else
        If Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
            frmRegistEdit.mblnCharge = True
            frmRegistEdit.mlngModul = 1111
            frmRegistEdit.mstrPrivs = gstrPrivs
            frmRegistEdit.mbytMode = 0
            frmRegistEdit.mbytInState = 0
            If gfrmMain Is Nothing Then
                frmRegistEdit.Show
            Else
                frmRegistEdit.Show 1, gfrmMain
            End If
        Else
            frmRegistEditNew.mblnCharge = True
            frmRegistEditNew.mlngModul = 1111
            frmRegistEditNew.mstrPrivs = gstrPrivs
            frmRegistEditNew.mbytMode = 0
            frmRegistEditNew.mbytInState = 0
            If gfrmMain Is Nothing Then
                frmRegistEditNew.Show
            Else
                frmRegistEditNew.Show 1, gfrmMain
            End If
        End If
    End If
    Register = gblnOk
    
    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
End Function

Public Function StationRegister(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, ByVal strRoom As String, _
    Optional ByVal blnNoneCut As Boolean, Optional ByVal blnPrice As Boolean = True, _
    Optional lng����ID As Long = 0, Optional strDbUser As String) As String
'���ܣ�����ҽ��վ���ùҺ�,��������Ȩ
'������strRoom=��������
'      blnNoneCut=�Ƿ�����ʹ�ô��۷ѱ�
'      blnPrice=�Ƿ������Һŷ������ɻ��۵��շ�
'���أ�����Һųɹ����򷵻عҺŵ��ţ����򷵻ؿ�
    
    Static blnNFStationRegister As Boolean  '�Ƿ��һ�ε���,��ʼ��ȫ�ֱ���
   
    gstrPrivsStation = ";" & GetPrivFunc(lngSys, 1260) & ";" '56274
    If Not blnNFStationRegister Then
        blnNFStationRegister = True
    
        If Not CheckValid Then Exit Function
    
        glngSys = lngSys
        Set gcnOracle = cnMain
        Set gfrmMain = frmMain: glngMain = 0
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        gstrUnitName = GetUnitName
        gstrDBUser = strDbUser  '52689
        
        '�û���Ϣ
        If Not GetUserInfo Then
            MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
            Exit Function
        End If
        'ҽ���ӿ�
        Call gclsInsure.InitOracle(gcnOracle)
        '��ʼ������
        Call InitLocPar(1111)
        Call InitSysPar
        Call CreateSquareCardObject(gfrmMain, 1111)
    End If
    '��Ҫÿ�θ�ֵ,�������������ģ����ı�
    If InStr(gstrPrivsStation, ";�����޸������;") > 0 Then
        gstrPrivs = ";���շѺ�;����Ѻ�;��������;�˺�;��ʱ�Һ�;�����޸������;" '����Ȩ��,���ó���ű�����Ȩ
    Else
        gstrPrivs = ";���շѺ�;����Ѻ�;��������;�˺�;��ʱ�Һ�;" '����Ȩ��,���ó���ű�����Ȩ
    End If
    
    '���ùҺ�
    Err.Clear
    On Error Resume Next
    If gbytRegistMode = 0 Then
        frmRegistEdit.mblnStation = True
        frmRegistEdit.mstrRoom = strRoom
        frmRegistEdit.mstrRegNo = ""
        frmRegistEdit.mblnNoneCut = blnNoneCut
        frmRegistEdit.mblnStationPrice = blnPrice
        frmRegistEdit.mlngModul = 1111
        frmRegistEdit.mstrPrivs = gstrPrivs
        frmRegistEdit.mbytMode = 0
        frmRegistEdit.mbytInState = 0
        If lng����ID <> 0 Then
            frmRegistEdit.zl_StationInitPatient (lng����ID)
        End If
        If gfrmMain Is Nothing Then
            frmRegistEdit.Show
        Else
            frmRegistEdit.Show 1, frmMain
        End If
        If gblnOk Then
            StationRegister = frmRegistEdit.mstrRegNo
        End If
    Else
        If Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
            frmRegistEdit.mblnStation = True
            frmRegistEdit.mstrRoom = strRoom
            frmRegistEdit.mstrRegNo = ""
            frmRegistEdit.mblnNoneCut = blnNoneCut
            frmRegistEdit.mblnStationPrice = blnPrice
            frmRegistEdit.mlngModul = 1111
            frmRegistEdit.mstrPrivs = gstrPrivs
            frmRegistEdit.mbytMode = 0
            frmRegistEdit.mbytInState = 0
            If lng����ID <> 0 Then
                frmRegistEdit.zl_StationInitPatient (lng����ID)
            End If
            If gfrmMain Is Nothing Then
                frmRegistEdit.Show
            Else
                frmRegistEdit.Show 1, frmMain
            End If
            If gblnOk Then
                StationRegister = frmRegistEdit.mstrRegNo
            End If
        Else
            frmRegistEditNew.mblnStation = True
            frmRegistEditNew.mstrRoom = strRoom
            frmRegistEditNew.mstrRegNo = ""
            frmRegistEditNew.mblnNoneCut = blnNoneCut
            frmRegistEditNew.mblnStationPrice = blnPrice
            frmRegistEditNew.mlngModul = 1111
            frmRegistEditNew.mstrPrivs = gstrPrivs
            frmRegistEditNew.mbytMode = 0
            frmRegistEditNew.mbytInState = 0
            If lng����ID <> 0 Then
                frmRegistEditNew.zl_StationInitPatient (lng����ID)
            End If
            If gfrmMain Is Nothing Then
                frmRegistEditNew.Show
            Else
                frmRegistEditNew.Show 1, frmMain
            End If
            If gblnOk Then
                StationRegister = frmRegistEditNew.mstrRegNo
            End If
        End If
    End If
End Function

Public Function StationBespeak(frmMain As Object, cnMain As ADODB.Connection, _
    lngSys As Long, ByVal strRoom As String, Optional ByVal blnNoneCut As Boolean, _
    Optional ByVal lng����ID As Long = 0, Optional strDbUser As String) As String
    '���ܣ�����ҽ��վ����ԤԼ�Һ�,��������Ȩ
    '������strRoom=��������
    '      blnNoneCut=�Ƿ�����ʹ�ô��۷ѱ�
    '���أ�����Һųɹ����򷵻عҺŵ��ţ����򷵻ؿ�
    Static blnNFStationBespeak As Boolean  '�Ƿ��һ�ε���,��ʼ��ȫ�ֱ���
   gstrPrivsStation = ";" & GetPrivFunc(lngSys, 1260) & ";" '56274
    If Not blnNFStationBespeak Then
        blnNFStationBespeak = True
        If Not CheckValid Then Exit Function
        
        glngSys = lngSys
        Set gcnOracle = cnMain
        Set gfrmMain = frmMain: glngMain = 0
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        gstrUnitName = GetUnitName
        
        gstrDBUser = strDbUser  '52689
        '�û���Ϣ
        If Not GetUserInfo Then
            MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
            Exit Function
        End If
        
        'ҽ���ӿ�
        Call gclsInsure.InitOracle(gcnOracle)
            
        '��ʼ������
        Call InitLocPar(1111)
        Call InitSysPar
        
        Call CreateSquareCardObject(gfrmMain, 1111)
        
    End If
    '��Ҫÿ�θ�ֵ,�������������ģ����ı�
    If InStr(gstrPrivsStation, ";�����޸������;") > 0 Then
        gstrPrivs = ";ԤԼ�Һ�;���շѺ�;����Ѻ�;��������;�����޸������;" '����Ȩ��,���ó���ű�����Ȩ
    Else
        gstrPrivs = ";ԤԼ�Һ�;���շѺ�;����Ѻ�;��������;" '����Ȩ��,���ó���ű�����Ȩ
    End If
    '���ùҺ�
    Err.Clear
    On Error Resume Next
    If gbytRegistMode = 0 Then
        frmRegistEdit.mblnStation = True
        frmRegistEdit.mstrRoom = strRoom
        frmRegistEdit.mstrRegNo = ""
        frmRegistEdit.mblnNoneCut = blnNoneCut
        frmRegistEdit.mlngModul = 1111
        frmRegistEdit.mstrPrivs = gstrPrivs
        frmRegistEdit.mbytMode = 1
        frmRegistEdit.mbytInState = 0
        If lng����ID <> 0 Then
            frmRegistEdit.zl_StationInitPatient (lng����ID)
        End If
        If gfrmMain Is Nothing Then
            frmRegistEdit.Show
        Else
            frmRegistEdit.Show 1, frmMain
        End If
        If gblnOk Then
            StationBespeak = frmRegistEdit.mstrRegNo
        End If
    Else
        If Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
            frmRegistEdit.mblnStation = True
            frmRegistEdit.mstrRoom = strRoom
            frmRegistEdit.mstrRegNo = ""
            frmRegistEdit.mblnNoneCut = blnNoneCut
            frmRegistEdit.mlngModul = 1111
            frmRegistEdit.mstrPrivs = gstrPrivs
            frmRegistEdit.mbytMode = 1
            frmRegistEdit.mbytInState = 0
            If lng����ID <> 0 Then
                frmRegistEdit.zl_StationInitPatient (lng����ID)
            End If
            If gfrmMain Is Nothing Then
                frmRegistEdit.Show
            Else
                frmRegistEdit.Show 1, frmMain
            End If
            If gblnOk Then
                StationBespeak = frmRegistEdit.mstrRegNo
            End If
        Else
            frmRegistEditNew.mblnStation = True
            frmRegistEditNew.mstrRoom = strRoom
            frmRegistEditNew.mstrRegNo = ""
            frmRegistEditNew.mblnNoneCut = blnNoneCut
            frmRegistEditNew.mlngModul = 1111
            frmRegistEditNew.mstrPrivs = gstrPrivs
            frmRegistEditNew.mbytMode = 1
            frmRegistEditNew.mbytInState = 0
            If lng����ID <> 0 Then
                frmRegistEditNew.zl_StationInitPatient (lng����ID)
            End If
            If gfrmMain Is Nothing Then
                frmRegistEditNew.Show
            Else
                frmRegistEditNew.Show 1, frmMain
            End If
            If gblnOk Then
                StationBespeak = frmRegistEditNew.mstrRegNo
            End If
        End If
    End If
    
End Function
Public Sub zlPrintBespeak(ByVal frmMain As Object, ByVal cnMain As ADODB.Connection, _
    ByVal lngSys As Long, ByVal strDbUser As String, ByVal strPrivs As String, ByVal strNO As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡԤԼ�Һŵ�
    '���:strNO=ԤԼ�Һŵ���
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2012-12-24 10:34:14
    '˵��:
    '����:56274
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If InStr(1, strPrivs, ";ԤԼ�Һŵ�;") = 0 Then Exit Sub
    Call ReportOpen(cnMain, lngSys, "ZL" & lngSys \ 100 & "_BILL_1111_1", frmMain, "NO=" & strNO, 2)
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
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
    CloseSquareCardObject
End Sub