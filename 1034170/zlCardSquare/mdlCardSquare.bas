Attribute VB_Name = "mdlCardSquare"
Option Explicit
Public Enum gС������
    g_���� = 0
    g_�ɱ���
    g_�ۼ�
    g_���
    g_�ۿ���
End Enum
Private Type m_С��λ
    ����С�� As Integer
    �ɱ���С�� As Integer
    ���ۼ�С�� As Integer
    ���С�� As Integer
    �ۿ��� As Integer
End Type
Public g_С��λ�� As m_С��λ

'С����ʽ����
Public Type g_FmtString
    FM_���� As String
    FM_�ɱ��� As String
    FM_���ۼ� As String
    FM_��� As String
    FM_�ۿ��� As String
End Type
Public Enum gCardEditType   '���༭����
    gEd_���� = 0
    gEd_�������� = 1
    gEd_�޸� = 2
    gEd_ɾ�� = 3
    gEd_��ѯ = 4
    gEd_��ֵ = 5
    gEd_���� = 6
    gEd_���� = 7
    gEd_ȡ������ = 8
    gEd_�˿� = 9
    gEd_ȡ���˿� = 10
End Enum
Public Type zlTyCustumRecordset
    rs�շ���� As ADODB.Recordset
    rs���ѿ��ӿ� As ADODB.Recordset
    rs�շ������� As ADODB.Recordset
    rs�ֵ������� As ADODB.Recordset
    dbl�����ܶ� As Double
    dblHIS������Ѷ� As Double
    dbl��ˢ�ۼƶ� As Double
End Type
Public gblnShowCard As Boolean  '���￨����ʾ(true,��ʾ����,false,������ʾ)
Public gObjXFCards As clsCards  'ר��������ѿ���(Ҫ������)
Public gobjSquare As SquareCard

Public grsStatic As zlTyCustumRecordset
Public gVbFmtString As g_FmtString
Public gOraFmtString As g_FmtString
Public gbln�Զ���ȡ As Boolean '��ǰ�Ƿ�Ϊ��Ƶ��
Public gblnCardNoSHowPW As Boolean  '������ʾ����
Public gDebug As Boolean '���Կ���
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjControl As Object
Public gstrLike As String  '��Ŀƥ�䷽��,%���
Private Type Ty_TestDebug
    blndebug As Boolean
    objSquareCard As clsCard
    BytType  As Byte  '1-�����������,2-��ȡ����
    strStartNo As String    '��ʼ����
    bln�������� As Boolean
End Type
Public gTy_TestBug As Ty_TestDebug
Public gobjStartCards As Collection  '������ˢ������
 
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"
Public gintFeePrecision As Integer    '����С������
Public gstrFeePrecisionFmt As String '����С����ʽ:0.00000
Public gblnOK As Boolean
'LED�������ۿ���
Public gblnLED As Boolean '�Ƿ�ʹ��Led��ʾ

Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    'support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��

    support���������շ� = 22        '�����������֤�󣬿ɽ��ж���շѲ���
    support�����շ���ɺ���֤ = 23  '�������շ���ɣ��Ƿ��ٴε��������֤
    
    supportҽ���ϴ� = 24            'ҽ����������ʱ�Ƿ�ʵʱ����
    support�ֱҴ��� = 25            'ҽ�������Ƿ���ֱ�
    support��;������������ϴ����� = 26 '�ṩ�����ϴ��������ݵĽ��㹦��
    support��������ѽ��ʵļ��ʵ��� = 27 '�Ƿ�����������ʵ��ݣ�����õ����Ѿ�����
    
    support�����ݳ������� = 28
    support��Ժ��ʵ�ʽ��� = 29      '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
    support�൥���շ� = 30          '�Ƿ�֧�ֶ൥���շ�
    
    support�����շѴ�Ϊ���۵� = 31  '�������շѵ�תΪ���۵����棬�޸���ǰ�̶��ж�ĳ��ҽ���ķ�ʽ
    
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧��ֻ�и������ʻ�ԭ����,�����ҽ�����㷽ʽ��Ϊ�ֽ�,֧�ֵ����ж�ÿһ�ֽ��㷽ʽ�Ƿ������˻�
    support�൥���շѱ���ȫ�� = 39  '�൥���շѱ���ȫ��
    
    supportҽ���ӿڴ�ӡƱ�� = 46    'HIS��ֻ��Ʊ�ݺŵ�������ӡ��ҽ���ӿ�(����)�д�ӡ
    support�൥��һ�ν��� = 47      '�൥��Ԥ����ʱ��ҽ���ӿڽ������һ�ε���ʱ���ؽ�������HIS���ٷ�̯��ÿ�ŵ�����
    
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
    supportҽ��ȷ���������� = 48
    supportʵʱ��� = 60             '�Ƿ����÷���ʵʱ���
    
    '���˺�:27536 20100119
    support�����ѽɿ���� = 64            '���շ�ʱ,����շѲ�����"�����нɿ�������ۼƿ���"Ϊtrueʱ,ͬʱ��ҽ������ʱû������ɿ���ʱ�������û�
    support�˷Ѻ��ӡ�ص� = 65   'ҽ�������Ƿ��˷Ѻ��ӡ�ص�:����
End Enum

Public Sub zlinitSystemPara(Optional cnOracle As ADODB.Connection)
    '------------------------------------------------------------------------------
    '------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ����
    '����:���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objDataBase As Object, objTemp As clsDataBase
    
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDataBase = objTemp
    Else
         Set objDataBase = zlDatabase
    End If
    
    '����:52913
    strSQL = "Select �������� From ҽ�ƿ���� Where ����='���￨' and nvl(�Ƿ�̶�,0)=1"
    Set rsTemp = objDataBase.OpenSQLRecord(strSQL, "��ȡԭ���￨����������ʾ����")
    gblnShowCard = False
    If Not rsTemp.EOF Then
        gblnShowCard = Nvl(rsTemp!��������) = ""
    End If
    '78773:���ϴ�,2014-10-29,LED��ʾһ��֧ͨ����Ϣ
    gblnLED = Val(GetSetting("ZLSOFT", "����ȫ��", "ʹ��", 0)) <> 0
    gstrLike = IIf(Val(objDataBase.GetPara("����ƥ��")) = 0, "%", "")
    With gSystemPara
        '0-ƴ����,1-�����,2-����
        .int���뷽ʽ = Val(objDataBase.GetPara("���뷽ʽ"))
        .bln���Ի���� = objDataBase.GetPara("ʹ�ø��Ի����") = "1"
        
        '��1λ1-ȫ����ֻ�����,��2λ1-ȫ��ĸֻ�����,��HIS��������������
        strTemp = objDataBase.GetPara(44, glngSys)
        If strTemp = "" Then strTemp = "00"
        If Len(strTemp) = 1 Then strTemp = strTemp & "0"
        .blnȫ���ְ������ = Val(Left(strTemp, 1)) = 1
        .blnȫ��ĸ������� = Val(Mid(strTemp, 2, 1)) = 1
        '���ý��С����λ��
        gbytDec = Val(objDataBase.GetPara(9, glngSys, , 2))
        gstrDec = "0." & String(gbytDec, "0")
        '���˺� ����:????    ����:2010-12-06 23:38:53
        '���õ��۱���λ��
        gintFeePrecision = Val(objDataBase.GetPara(157, glngSys, , "5"))
        gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
     End With
     gintDebug = -1
     '���绯վ����Ϣ
     Call Initվ����Ϣ: Call ��ʼС��λ��
     Call zlInitColorSet
     Set objDataBase = Nothing
     Set objTemp = Nothing
End Sub
Public Sub ��ʼС��λ��()
    '------------------------------------------------------------------------------------------------------
    '����:��ʼС��λ��
    '���:
    '����:
    '����:7
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/6
    '------------------------------------------------------------------------------------------------------
    With g_С��λ��
        .�ɱ���С�� = 7
        .���ۼ�С�� = 7
        .���С�� = 2
        .����С�� = 3
        .�ۿ��� = 2
    End With
    With gVbFmtString
        .FM_�ɱ��� = GetFmtString(g_�ɱ���, False)
        .FM_��� = GetFmtString(g_���, False)
        .FM_���ۼ� = GetFmtString(g_�ۼ�, False)
        .FM_���� = GetFmtString(g_����, False)
        .FM_�ۿ��� = GetFmtString(g_�ۿ���, False)
    End With
    With gOraFmtString
        .FM_�ɱ��� = GetFmtString(g_�ɱ���, True)
        .FM_��� = GetFmtString(g_���, True)
        .FM_���ۼ� = GetFmtString(g_�ۼ�, True)
        .FM_���� = GetFmtString(g_����, True)
        .FM_�ۿ��� = GetFmtString(g_�ۿ���, True)
    End With
End Sub

Public Function GetFmtString(ByVal С������ As gС������, Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '����:����ָ����С����ʽ��
    '���: lngС��λ��-С��λ��
    '     blnOracle-������oracle�ĸ�ʽ������Vb�ĸ�ʽ��
    '����:
    '����:����ָ���ĸ�ʽ��
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strFmt As String
    Dim intλ�� As Integer
    Select Case С������
    Case g_����
         intλ�� = g_С��λ��.����С��
    Case g_���
         intλ�� = g_С��λ��.���С��
    Case g_�ɱ���
         intλ�� = g_С��λ��.�ɱ���С��
    Case g_�ۼ�
         intλ�� = g_С��λ��.���ۼ�С��
    Case Else
        intλ�� = 0
    End Select
    If blnOracle Then
       GetFmtString = "'999999999990." & String(intλ��, "9") & "'"
    Else
       GetFmtString = "#0." & String(intλ��, "0") & ";-#0." & String(intλ��, "0") & "; ;"
    End If
End Function

Public Function zlCheckPrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ����Ȩ���Ƿ����
    '����:strPrivs-Ȩ�޴�
    '     strMyPriv-����Ȩ��
    '����,����Ȩ��,����true,���򷵻�False
    '����:���˺�
    '����:2009-11-19 14:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlCheckPrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function
Public Function zlGet�շ����() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շ����
    '����:���˺�
    '����:2009-12-09 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '�Ȼ��浽����
    
    On Error GoTo errHandle
    
    gstrSQL = "Select  ����,���� From �շ���Ŀ���"
    If grsStatic.rs�շ���� Is Nothing Then
        Set grsStatic.rs�շ���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ����")
    ElseIf grsStatic.rs�շ����.State <> 1 Then
        Set grsStatic.rs�շ���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ����")
    End If
    Set zlGet�շ���� = grsStatic.rs�շ����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet���ѿ��ӿ�(Optional cnOracle As ADODB.Connection) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ѿ��ӿ�
    '����:���˺�
    '����:2009-12-09 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '�Ȼ��浽����
    Dim objDataBase  As Object, objTemp As clsDataBase
    On Error GoTo errHandle
    Set objDataBase = zlDatabase
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDataBase = objTemp
    End If
    '56615
    gstrSQL = "" & _
    "   Select ���,����,���㷽ʽ,nvl(���ƿ�,0)  as ���ƿ�,ǰ׺�ı�,���ų���, " & _
    "           nvl(�Ƿ�����,0) as �Ƿ�����,nvl(�Ƿ�ȫ��,0) as �Ƿ�ȫ��,nvl(�Ƿ�ˢ��,0) as �Ƿ�ˢ��," & _
    "           nvl(���볤��,10) as ���볤��,nvl(���볤������,0) as ���볤������,nvl(�������,0) as �������," & _
    "           ����,ϵͳ,�Ƿ�����,0 as ������������,0 as �Ƿ�ȱʡ����," & _
    "           0 as �Ƿ�ģ������,0 as �Ƿ��ƿ�, 1 as �Ƿ񷢿�, 0 as �Ƿ�д��" & _
    "   From �����ѽӿ�Ŀ¼ where nvl(����,0)=1 "
    If grsStatic.rs���ѿ��ӿ� Is Nothing Then
        Set grsStatic.rs���ѿ��ӿ� = objDataBase.OpenSQLRecord(gstrSQL, "��ȡ���ѿ��ӿ� ")
    ElseIf grsStatic.rs���ѿ��ӿ�.State <> 1 Then
        Set grsStatic.rs���ѿ��ӿ� = objDataBase.OpenSQLRecord(gstrSQL, "��ȡ���ѿ��ӿ� ")
    End If
     

    grsStatic.rs���ѿ��ӿ�.Filter = 0
    Set zlGet���ѿ��ӿ� = grsStatic.rs���ѿ��ӿ�
    Exit Function
errHandle:
    If Not cnOracle Is Nothing And Not objTemp Is Nothing Then
        If objTemp.ErrCenter = 1 Then Resume
        Set objTemp = Nothing: Set objDataBase = Nothing
        Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Set objTemp = Nothing: Set objDataBase = Nothing
End Function

Public Function zlIsCardNoShowPW(ByRef lng��� As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�����Ƿ�������ʾ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2010-10-25 10:31:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = zlGet���ѿ��ӿ�
    If rsTemp.EOF Then Exit Function
    rsTemp.Filter = "���=" & lng���
    If rsTemp.EOF Then
        zlIsCardNoShowPW = False
    Else
         zlIsCardNoShowPW = Val(Nvl(rsTemp!�Ƿ�����)) = 1
    End If
    rsTemp.Filter = 0
End Function
Public Function zlCreateBrushObjects(ByVal objCard As clsCard, ByRef objBrhushCardObject As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '���:clsCard-������
    '����:
    '����:
    '����:���˺�
    '����:2009-12-31 14:46:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCommpentName As String
    If objCard.���� Then
        '����豸�Ƿ�����
        If objCard.�ӿڳ����� = "" Then
            '���ѿ�
            Set objBrhushCardObject = New clsSimulateSquareCard: zlCreateBrushObjects = True
        Else
            strCommpentName = objCard.�ӿڳ����� & "." & "cls" & Replace(Replace(UCase(objCard.�ӿڳ�����), "ZL9", ""), "ZL", "")
            Err = 0: On Error Resume Next
            Set objBrhushCardObject = CreateObject(strCommpentName)
            If Err <> 0 Then
                ShowMsgbox "����:" & objCard.�ӿڱ��� & "-" & objCard.���� & "( " & strCommpentName & ")����ʧ��!" & vbCrLf & "��ϸ����ϢΪ:" & Err.Description
                Call WritLog("mdlCardSquare.zlCreateBrushObjects", "", "����:" & objCard.�ӿڱ��� & "-" & objCard.���� & "����ʧ��!��ϸ����ϢΪ:" & Err.Description)
                Exit Function
            End If
            zlCreateBrushObjects = True
        End If
    Else
        Set objBrhushCardObject = Nothing
    End If
End Function
Public Function zlGetCardObject(ByVal lng�ӿڱ�� As Long, ByRef objBrushCard As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����ָ�����㿨��Ż�ȡ���㿨����
    '��Σ�lng�ӿڱ��-���㿨�����
    '���Σ�objCard-���ؽ��㿨����
    '���أ���ȡ�ɹ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 11:58:54
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objCardTemp As Object
    If gobjStartCards Is Nothing Then Exit Function
    
    If gobjStartCards.count = 0 Then Exit Function
    For i = 1 To gobjStartCards.count
         Err = 0: On Error Resume Next
         Set objCardTemp = gobjStartCards(i)(0)
         If Err = 0 Then
            If gobjStartCards(i)(2) = lng�ӿڱ�� Then
                Set objBrushCard = objCardTemp
                zlGetCardObject = True: Exit Function
            End If
        End If
        On Error GoTo 0
    Next
    zlGetCardObject = False
End Function

Public Function zlInitCards() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:�ɹ�!����true,���򷵻�False
    '����:���˺�
    '����:2009-12-15 14:31:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, int�Զ���ȡ As Integer, bln���� As Boolean, str���� As String, objCard As clsCard
    Dim objBrushCards As Object, int�Զ���� As Integer
    
    Err = 0: On Error GoTo Errhand:
    Set gObjXFCards = New clsCards
    Set gobjStartCards = New Collection '��ʽ;array(��������,���ƿ�,�ӿڱ��)
    Set rsTemp = zlGet���ѿ��ӿ�
    With rsTemp
        '���ƿ�(�����ѿ�)
        .Filter = "���ƿ�=1"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            ' "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ"
            int�Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\" & Nvl(!���), "�Զ���ȡ", "0"))
            bln���� = Val(GetSetting("ZLSOFT", "����ģ��\zlSquareCard\" & Nvl(!���), "����", "1")) = 1
            int�Զ���� = Val(GetSetting("ZLSOFT", "����ģ��\zlSquareCard\" & Nvl(!���), "�Զ���ȡ���", "1"))
                
            str���� = Nvl(rsTemp!����)
            Set objCard = gObjXFCards.AddItem(EM_CardType_Consume, Val(Nvl(!���)), Nvl(!���), Nvl(rsTemp!����), Left(Nvl(rsTemp!����), 1), bln����, True, str����, True, 1, int�Զ���ȡ, int�Զ����, Val(Nvl(rsTemp!ϵͳ)) = 1, Nvl(rsTemp!���㷽ʽ), Nvl(rsTemp!ǰ׺�ı�), Val(Nvl(rsTemp!���ų���)), True, Val(Nvl(rsTemp!�Ƿ�ˢ��)) = 1, False, Val(Nvl(rsTemp!�Ƿ�ȫ��)) = 1, "", "", True, Val(Nvl(rsTemp!�Ƿ�����)), Val(Nvl(rsTemp!�Ƿ�����)) = 1, Val(Nvl(rsTemp!���볤��)), Val(Nvl(rsTemp!���볤������)), Val(Nvl(rsTemp!�������)), "K" & Nvl(rsTemp!���))
            If zlCreateBrushObjects(objCard, objBrushCards) Then
                gobjStartCards.Add Array(objBrushCards, "1", CStr(Nvl(!���))), "K" & Nvl(!���)
            End If
            .MoveNext
        Loop
        '������
        .Filter = "���ƿ�<>1"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            int�Զ���ȡ = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard\" & Nvl(!���), "�Զ���ȡ", 0))
            int�Զ���� = Val(GetSetting("ZLSOFT", "����ģ��\zlSquareCard\" & Nvl(!���), "�Զ���ȡ���", "1"))
            bln���� = Val(GetSetting("ZLSOFT", "����ģ��\zlSquareCard\" & Nvl(!���), "����", "1")) = 1
            str���� = Nvl(rsTemp!����)
             Set objCard = gObjXFCards.AddItem(EM_CardType_Consume, Val(Nvl(!���)), Nvl(!���), Nvl(rsTemp!����), Left(Nvl(rsTemp!����), 1), bln����, True, str����, False, 1, int�Զ���ȡ, int�Զ����, Val(Nvl(rsTemp!ϵͳ)) = 1, Nvl(rsTemp!���㷽ʽ), Nvl(rsTemp!ǰ׺�ı�), Val(Nvl(rsTemp!���ų���)), True, Val(Nvl(rsTemp!�Ƿ�ˢ��)) = 1, True, Val(Nvl(rsTemp!�Ƿ�ȫ��)) = 1, "", "", True, Val(Nvl(rsTemp!�Ƿ�����)), Val(Nvl(rsTemp!�Ƿ�����)) = 1, Val(Nvl(rsTemp!���볤��)), Val(Nvl(rsTemp!���볤������)), Val(Nvl(rsTemp!�������)), "K" & Nvl(rsTemp!���))
            If zlCreateBrushObjects(objCard, objBrushCards) Then
                gobjStartCards.Add Array(objBrushCards, 0, CStr(Nvl(!���))), "K" & Nvl(!���)
            End If
            .MoveNext
        Loop
    End With
    zlInitCards = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub WritLog(ByVal strDev As String, strInput As String, strOutPut As String)
    On Error GoTo errHandle
    If gDebug Then
        Open App.Path & "\SquareCard" & Format(Now(), "yyyyMMdd") & ".log" For Append As #1
        Write #1, Now
        Write #1, strDev; strInput; strOutPut
        Write #1, "======================================================================="
        Close #1
    End If
    Exit Sub
errHandle:
    MsgBox "д��־���ִ���" & vbNewLine & Err.Description, vbExclamation, "IC���ӿ�"
End Sub

Public Function Readģ�⿨��(ByVal strFile As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ѿ������Ŀ����ж�ȡһ������־�Ŀ���(����ж��,�����һ��Ϊ׼)
    '����:���˺�
    '����:2009-12-17 10:35:51
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim objFile As New FileSystemObject, objText As TextStream, varData As Variant
    Dim strText As String, strCardNo As String
    strCardNo = ""
    Set objText = objFile.OpenTextFile(strFile)
    Do While Not objText.AtEndOfStream
        strText = Trim(objText.ReadLine)
        varData = Split(strText, vbTab)
        If Val(varData(0)) = 1 Then
            strCardNo = varData(1)
        End If
    Loop
    objText.Close
    Readģ�⿨�� = strCardNo
    Exit Function
Errhand:
End Function
Public Sub zlInitBrushCardRec(ByRef rsTemp As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ؼ�¼��
    '����:���ر��ؽ���ĳ�����¼��
    '����:���˺�
    '����:2009-12-23 11:22:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        If .State = adStateOpen Then .Close
        .Fields.Append "�ӿڱ��", adDouble, 18, adFldIsNullable
        .Fields.Append "���ѿ�ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "���㷽ʽ", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���", adDouble, 16, adFldIsNullable
        .Fields.Append "������", adDouble, 16, adFldIsNullable
        .Fields.Append "����ʱ��", adDate, 50, adFldIsNullable
        .Fields.Append "������ˮ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��ע", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "�����־", adNumeric, 2, adFldIsNullable
        .Fields.Append "��̯ҳ��", adLongVarChar, 600, adFldIsNullable  '�൥����Ч,��HIS������Զ�����:�ö��ŷ���,��,2,3��ʾ,����ˢ�����ѷ����ڵڶ��ŵ��ݺ͵����ŵ���
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
End Sub
Public Sub zlInit�շ����Struc()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ؼ�¼��
    '����:���ر��ؽ���ĳ�����¼��
    '����:���˺�
    '����:2009-12-23 11:22:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set grsStatic.rs�շ������� = New ADODB.Recordset
    Set grsStatic.rs�ֵ������� = New ADODB.Recordset
    
    grsStatic.dbl�����ܶ� = 0: grsStatic.dbl��ˢ�ۼƶ� = 0
    With grsStatic.rs�շ�������
        If .State = adStateOpen Then .Close
        .Fields.Append "�շ����", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ʵ�ս��", adDouble, 16, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    
    With grsStatic.rs�ֵ�������
        If .State = adStateOpen Then .Close
        .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "�շ����", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ʵ�ս��", adDouble, 16, adFldIsNullable
        .Fields.Append "��̯���", adDouble, 16, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
End Sub
Public Function zlInit�շ��������(ByVal rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷ��ü�¼������ȡ��ǰ���������ѵ������
    '���:rsFeeList-��ϸ����:
    '    �ֶ�: �ѱ�,NO,ʵ��Ʊ�š�����ʱ�䡢����ID���շ�����վݷ�Ŀ�����㵥λ�������ˡ��շ�ϸĿID�����������ۡ�ʵ�ս��Ƿ����������ID��ִ�в���ID
    '����:
    '����:���˺�
    '����:2009-12-23 16:11:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl������Ѷ� As Double, str�շ���� As String, lng��� As Long
    Err = 0: On Error GoTo Errhand:
    Call zlInit�շ����Struc
    lng��� = 0
    With rsFeeList
        .Sort = "�շ����"
        Do While Not rsFeeList.EOF
            If str�շ���� <> Nvl(!�շ����) Then
                grsStatic.rs�շ�������.AddNew
                grsStatic.rs�շ�������!�շ���� = Nvl(!�շ����)
                str�շ���� = Nvl(!�շ����)
            End If
            grsStatic.rs�շ�������!ʵ�ս�� = Val(Nvl(grsStatic.rs�շ�������!ʵ�ս��)) + Val(Nvl(!ʵ�ս��))
            grsStatic.rs�շ�������.Update
            grsStatic.dbl�����ܶ� = grsStatic.dbl�����ܶ� + Val(Nvl(!ʵ�ս��))
            
            grsStatic.rs�ֵ�������.Find "����='" & Nvl(rsFeeList!�������) & "_" & Nvl(!�շ����) & "'", , , 1
            If grsStatic.rs�ֵ�������.EOF Then
                grsStatic.rs�ֵ�������.AddNew
                grsStatic.rs�ֵ�������!�շ���� = Nvl(!�շ����)
                
            End If
            grsStatic.rs�ֵ�������!������� = Val(Nvl(!�������))
            grsStatic.rs�ֵ�������!ʵ�ս�� = Val(Nvl(grsStatic.rs�ֵ�������!ʵ�ս��)) + Val(Nvl(!ʵ�ս��))
            grsStatic.rs�ֵ�������.Update
            rsFeeList.MoveNext
        Loop
    End With
    zlInit�շ�������� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zl��ȡ������Ѷ�(ByVal str������� As String, ByVal dbl������Ѷ� As Double, ByVal dbl��ˢ�ۼ� As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ѷ�
    '    dbl������Ѷ�=-1��ʾδ����������Ѷ�
    '����:���˺�
    '����:2009-12-24 10:24:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�޶���� As Double, dbl������ As Double
    Err = 0: On Error GoTo Errhand:
    
    If str������� <> "" Then
        str������� = zlGet��ȡ�������FromNameToCode(str�������)
    End If
    dbl�޶���� = 0
    If str������� <> "" Then
        With grsStatic.rs�շ�������
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                If InStr(1, str�������, "," & Nvl(!�շ����) & ",") > 0 Then
                    dbl�޶���� = dbl�޶���� + Val(Nvl(!ʵ�ս��))
                End If
                .MoveNext
            Loop
        End With
    End If
    '���㹫ʽ:
    '�������Ѷ�= �ܷ���-��Ԥ��-�����Ѷ�-�޶����
    dbl������ = dbl������Ѷ� - dbl�޶���� - dbl��ˢ�ۼ�
    zl��ȡ������Ѷ� = IIf(dbl������ < 0, 0, dbl������)
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlGetʧЧ���(ByVal lng���ѿ�ID As Long, ByVal lng�ӿڱ�� As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡʧЧ���
    '����:ʧЧ���
    '����:���˺�
    '����:2009-12-23 15:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, dblTemp As Double
    Err = 0: On Error GoTo Errhand:
    gstrSQL = " " & _
    "  Select Sum(Nvl(ʧЧ���, 0)) As ʧЧ��� " & _
    "  From (Select ������ As ʧЧ��� " & _
    "         From ���ѿ�Ŀ¼ A " & _
    "         Where ID =  [1]  " & _
    "         Union All " & _
    "         Select -1 * Sum(Nvl(A.������, 0)) As ʧЧ��� " & _
    "         From ���˿������¼ A, ���ѿ�Ŀ¼ B " & _
    "         Where A.���ѿ�id = B.ID And A.���ѿ�id =  [1]  And A.�ӿڱ�� =  [2]  And " & _
    "               A.����ʱ�� <= Nvl(B.��Ч��, To_Date('3000-01-01', 'yyyy-mm-dd')))"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡʧЧ��", lng���ѿ�ID, lng�ӿڱ��)
    dblTemp = Val(Nvl(rsTemp!ʧЧ���))
    If dblTemp < 0 Then dblTemp = 0
    zlGetʧЧ��� = dblTemp
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlGet��ȡ�������FromNameToCode(ByVal str������� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������������ȡ��صı���
    '����:
    '����:���˺�
    '����:2009-12-23 16:31:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = zlGet�շ����
    rsTemp.Filter = 0
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    If str������� = "" Then zlGet��ȡ�������FromNameToCode = "": Exit Function
    str������� = "," & str������� & ","
    With rsTemp
        Do While Not .EOF
            str������� = Replace(str�������, "," & Nvl(rsTemp!����) & ",", "," & Nvl(rsTemp!����) & ",")
            .MoveNext
        Loop
    End With
    zlGet��ȡ�������FromNameToCode = str�������
 End Function
Public Function zl��̯��������(ByRef rsRquare As ADODB.Recordset, ByRef rs��̯ As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ˢ�������̯�������ݸ�ÿ�ŵ�����ϸ
    '����� rsRquare-(�ӿڱ�� ���ѿ�ID,����,���㷽ʽ,������,���,������ ����ʱ��,��ע,�����־)
    '       rs��̯-��ʾÿ�ŵ��ݷ�̯���
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-06 10:13:43
    '����˵��:
    '   1.�ȷ�̯��������
    '   2.�ٷ�̯����������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strTemp As String, str������� As String, dbl��� As Double
    Dim dbl�ܶ� As Double
    Set rs��̯ = New ADODB.Recordset
    With rs��̯
        If .State = adStateOpen Then .Close
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "���ѿ�ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "���㷽ʽ", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "��̯��", adDouble, 16, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With

    Set rsTemp = zlDatabase.CopyNewRec(rsRquare)
    Err = 0: On Error GoTo Errhand:
    
    '��ȷ����������Щ�������
    rsTemp.Filter = "���ѿ�ID >0"
    str������� = ""
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        strTemp = zlFromCardGet�������(Val(Nvl(rsTemp!���ѿ�ID)), False)
        If InStr(1, str�������, strTemp) <= 0 Then
            str������� = str������� & "," & strTemp
        End If
        rsTemp.MoveNext
    Loop
    
    rsTemp.Filter = 0
    If str������� <> "" Then
        str������� = zlGet��ȡ�������FromNameToCode(str�������) & ","
    End If
    
    rsTemp.Filter = 0
    With grsStatic.rs�ֵ�������
        '�Ƚ��������Ľ��з�̯
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '��Ҫ����
            If InStr(1, str�������, "," & Nvl(!�շ����) & ",") > 0 Then
                '�����������,�Ƚ��ⲿ�ַ�̯��
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                   strTemp = zlFromCardGet�������(Val(Nvl(rsTemp!���ѿ�ID)), True)
                   If InStr(1, strTemp, "," & Nvl(!�շ����) & ",") <= 0 And Val(Nvl(rsTemp!������)) > 0 Then
                      'ֻ���ò��޶������ķ�̯
                       dbl��� = Val(Nvl(!ʵ�ս��))
                      If dbl��� >= Val(Nvl(rsTemp!������)) Then
                        dbl��� = Val(Nvl(rsTemp!������))
                        rsTemp!������ = 0
                        rsTemp.Update
                        !��̯��� = Val(Nvl(!��̯���)) + dbl���
                        .Update
                      Else
                        'С�Ļ�
                        rsTemp!������ = Val(Nvl(rsTemp!������)) - dbl���
                        rsTemp.Update
                        !��̯��� = Val(Nvl(!��̯���)) + dbl���
                      End If
                      rs��̯.Filter = "�������=" & Val(Nvl(rsTemp!�������)) & " And ���ѿ�ID=" & Val(Nvl(rsTemp!���ѿ�ID)) & " And ����='" & Nvl(rsTemp!����) & "'"
                      If rs��̯.EOF Then
                          rs��̯.AddNew
                      End If
                      rs��̯!������� = Val(Nvl(rsTemp!�������))
                      rs��̯!���ѿ�ID = Val(Nvl(rsTemp!���ѿ�ID))
                      rs��̯!���� = Nvl(rsTemp!����)
                      rs��̯!���㷽ʽ = Trim(Nvl(rsTemp!���㷽ʽ))
                      rs��̯!��̯�� = Val(Nvl(rs��̯!��̯��)) + dbl���
                      rs��̯.Update
                   End If
                   If !��̯��� = !ʵ�ս�� Then Exit Do
                   rsTemp.MoveNext
                Loop
            End If
            .MoveNext
        Loop
        '�ٷ�̯���޶���
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
             If Val(Nvl(!��̯���)) <= Val(Nvl(!ʵ�ս��)) Then
                
                rsTemp.Filter = 0
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                   strTemp = zlFromCardGet�������(Val(Nvl(rsTemp!���ѿ�ID)), True)
                   If InStr(1, strTemp, "," & Nvl(!�շ����) & ",") <= 0 And Val(Nvl(rsTemp!������)) > 0 Then
                      dbl��� = Val(Nvl(!ʵ�ս��))
                      If dbl��� >= Val(Nvl(rsTemp!������)) Then
                        dbl��� = Val(Nvl(rsTemp!������))
                        rsTemp!������ = 0
                        rsTemp.Update
                        !��̯��� = Val(Nvl(!��̯���)) + dbl���
                        .Update
                      Else
                        'С�Ļ�
                        rsTemp!������ = Val(Nvl(rsTemp!������)) - dbl���
                        rsTemp.Update
                        !��̯��� = Val(Nvl(!��̯���)) + dbl���
                      End If
                      rs��̯.Filter = "�������=" & Val(Nvl(!�������)) & " And ���ѿ�ID=" & Val(Nvl(rsTemp!���ѿ�ID)) & " And ����='" & Nvl(rsTemp!����) & "'"
                      If rs��̯.EOF Then
                          rs��̯.AddNew
                      End If
                      rs��̯!������� = Val(Nvl(!�������))
                      rs��̯!���ѿ�ID = Val(Nvl(rsTemp!���ѿ�ID))
                      rs��̯!���� = Nvl(rsTemp!����)
                      rs��̯!���㷽ʽ = Trim(Nvl(rsTemp!���㷽ʽ))
                      rs��̯!��̯�� = Val(Nvl(rs��̯!��̯��)) + dbl���
                      rs��̯.Update
                   End If
                   If !��̯��� = !ʵ�ս�� Then Exit Do
                   rsTemp.MoveNext
                Loop
             End If
             .MoveNext
        Loop
    End With
    
    With rs��̯
        .Filter = 0
        If .RecordCount > 0 Then .MoveFirst
        dbl��� = 0
        Do While Not .EOF
            dbl��� = dbl��� + Val(Nvl(!��̯��))
            .MoveNext
        Loop
    End With
    dbl�ܶ� = 0
    With rsRquare
        .Filter = 0
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            dbl�ܶ� = dbl�ܶ� + Val(Nvl(!������))
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    
    If Round(dbl�ܶ�, 4) <> Round(dbl���, 4) Then
        ShowMsgbox "�൥�ݷ�̯ʱ�������˲������,������ˢ��!"
        Exit Function
    End If
    '����������ϸ��̯�����ܵ��Ƿ�һ��
    zl��̯�������� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlFromCardGet�������(ByVal lng���ѿ�ID As Long, ByVal blnCode As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ѿ�,��ȡ��ص��޶�����
    '���:lng���ѿ�ID-���ѿ�ID
    '     blnCode-����
    '����:
    '����:�����������
    '����:���˺�
    '����:2010-01-06 11:18:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, str������� As String
    Err = 0: On Error GoTo Errhand:
    gstrSQL = "Select ������� From ���ѿ�Ŀ¼ Where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ѿ�Ŀ¼���������", lng���ѿ�ID)
    If rsTemp.EOF Then Exit Function
    str������� = Nvl(rsTemp!�������)
    If blnCode Then
        zlFromCardGet������� = zlGet��ȡ�������FromNameToCode(str�������)
    Else
        zlFromCardGet������� = str�������
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlGetRquare(ByVal str����ID_IN As String, ByRef rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����㽻��ʱ�����Ԥ������
    '���:str����ID_IN-ָ���Ľ���ID
    '����:rsSquare-��������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-15 11:08:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, lngID As Long
    
    On Error GoTo errHandle
    
    Call zlInitBrushCardRec(rsSquare)
    If str����ID_IN = "" Then str����ID_IN = "0"
    
    strSQL = "Select  /*+ rule */ Distinct A.ID, �ӿڱ��, A.���ѿ�id, A.���, A.��¼״̬, A.���㷽ʽ, A.������, A.����, A.������ˮ��, " & _
             "                   A.����ʱ��, A.��ע, A.�����־, C.����id " & _
             "   From ����Ԥ����¼ C, ���˿������¼ A, ���˿�������� B, " & _
             "        (Select Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) J " & _
             "   Where B.Ԥ��id = C.ID And B.������id = A.ID And C.����id = J.Column_Value And A.�����־ = 0 And C.��¼״̬ = 1" & _
             " Order by ID,����ID"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ID�����ˢ����Ϣ", str����ID_IN)
    gTy_TestBug.bln�������� = True
    With rsSquare
        Do While Not rsTemp.EOF
            If lngID <> Val(Nvl(rsTemp!id)) Then
                .AddNew
                !�ӿڱ�� = Val(Nvl(rsTemp!�ӿڱ��))
                !���ѿ�ID = Val(Nvl(rsTemp!���ѿ�ID))
                !���� = Nvl(rsTemp!����)
                !���㷽ʽ = Nvl(rsTemp!���㷽ʽ)
                !������ = zlGet�ӿ�����(Val(Nvl(rsTemp!�ӿڱ��)))
                !��� = 0
                !������ = Val(Nvl(rsTemp!������))
                !����ʱ�� = rsTemp!����ʱ��
                !������ˮ�� = IIf(Val(Nvl(rsTemp!���ѿ�ID)) = 0, Nvl(rsTemp!������ˮ��), Nvl(rsTemp!id))     '���ڣ����ѿ��Ĵ���û���ر�Ĵ����ڲ�������ʱ��ֻ��ģ�����á��򵥵ĸ�����صı�ʶ
                !��ע = Nvl(rsTemp!��ע)
                !�����־ = 0
            End If
            !��̯ҳ�� = Nvl(!��̯ҳ��) & "," & Val(Nvl(rsTemp!����ID))
            .Update
            rsTemp.MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    zlGetRquare = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlGet�ӿ�����(ByVal lng�ӿڱ�� As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ӿ�����
    '����:�ӿ�����
    '����:���˺�
    '����:2010-01-15 11:23:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp  As ADODB.Recordset
    Set rsTemp = zlGet���ѿ��ӿ�
    rsTemp.Filter = "���=" & lng�ӿڱ��
    If rsTemp.EOF Then
        zlGet�ӿ����� = ""
    Else
        zlGet�ӿ����� = Nvl(rsTemp!����)
    End If
End Function
Public Function zlGet�ӿڱ��(ByVal lngԤ��ID As Long) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����Ԥ��ID,��ȡ��Ӧ�Ľӿڱ��
    '����:���㿨�Ľӿ�ID
    '���ƣ����˺�
    '���ڣ�2010-06-18 14:05:08
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select  distinct A.�ӿڱ�� " & _
    "   From  ���˿������¼ A,���˿�������� C" & _
    "   Where  C.������ID=A.ID  and  C.Ԥ��ID=[1]    "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�˵��Ľӿڱ��", lngԤ��ID)
    If rsTemp.RecordCount = 0 Then zlGet�ӿڱ�� = 0: Exit Function
    zlGet�ӿڱ�� = Val(Nvl(rsTemp!�ӿڱ��))
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlSave�������¼(ByVal lngԤ��ID As Long, ByVal strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������صĽ�������
    '           ��||�ָ�: �ӿڱ��||���ѿ�ID(�ɴ�'')||���㷽ʽ||������||����||������ˮ��||����ʱ��(yyyy-mm-dd hh24:mi:ss)||��ע
    '���ƣ����˺�
    '���ڣ�2010-06-18 16:07:05
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, strSQL As String, strTemp As String
    If strBlanceInfor = "" Then Exit Function
    varData = Split(strBlanceInfor, "||")
    If UBound(varData) < 7 Then Exit Function
    
    'Zl_���˿������¼_Insert
    strSQL = "Zl_���˿������¼_Insert("
    '  �ӿڱ��_In   In ���˿������¼.�ӿڱ��%Type,
    strSQL = strSQL & "" & Val(varData(0)) & ","
    '  ���ѿ�id_In   In ���˿������¼.���ѿ�id%Type,
    strSQL = strSQL & "" & IIf(Val(varData(1)) = 0, "NULL", Val(varData(1))) & ","
    '  ���㷽ʽ_In   In ���˿������¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & Trim(varData(2)) & "',"
    '  ������_In   In ���˿������¼.������%Type,
    strSQL = strSQL & "" & Val(varData(3)) & ","
    '  ����_In       In ���˿������¼.����%Type,
    strSQL = strSQL & "'" & Trim(varData(4)) & "',"
    '  ������ˮ��_In In ���˿������¼.������ˮ��%Type,
    strSQL = strSQL & "'" & Trim(varData(5)) & "',"
    '  ����ʱ��_In   In ���˿������¼.����ʱ��%Type,
    If Trim(varData(6)) = "" Or IsDate(varData(6)) = False Then
        strTemp = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        strTemp = Trim(varData(6))
    End If
    If strTemp = "" Then
        strSQL = strSQL & "NULL,"
    Else
        strSQL = strSQL & "to_date('" & strTemp & "','yyyy-mm-dd hh24:mi:ss'),"
    End If
    '  ��ע_In       In ���˿������¼.��ע%Type,
    strSQL = strSQL & "'" & Trim(varData(7)) & "',"
    '  ����id_In     In Varchar2
    strSQL = strSQL & "NULL,"
    '   Ԥ��id_In     In ����Ԥ����¼.ID%Type := -1
    strSQL = strSQL & "" & lngԤ��ID & ")"
    zlDatabase.ExecuteProcedure strSQL, "���濨�����¼"
    zlSave�������¼ = True
End Function

Public Function zlInputIsCard(ByRef txtInput As Object, ByVal KeyAscii As Integer, ByVal lngSys As Long, Optional ByVal blnPassWd As Boolean = False) As Boolean
'���ܣ��ж�ָ���ı����е�ǰ�����Ƿ���ˢ��(�Ƿ�ﵽ���ų��ȣ��ڵ��ó������ж�),������ϵͳ���������Ƿ�������ʾ
'������KeyAscii=��KeyPress�¼��е��õĲ���
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, strText As String
    
     'ˢ��ʱ����������ŵ��ɵ��÷�ȡ������
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then Exit Function
    
    '����ǰ�������ʾ������(��δ��ʾ����)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    '�ж��Ƿ���ˢ��
    If KeyAscii > 32 Then
        sngNow = timer
        If txtInput.Text = "" Or strText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True   '��һ̨�ʼǱ����ԣ�һ����0.014����
        End If
    End If
    'ˢ��ʱ�����Ƿ�������ʾ
    If blnCard Then
        txtInput.PasswordChar = IIf(Not blnPassWd, "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    zlInputIsCard = blnCard
End Function

Public Function zl_GetԤԼ��ʽByNo(strNO As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݹҺŵ��ݺŻ�ȡ����ԤԼ��ʽ
    '���:strNo-�Һŵ��ݺ�
    '����:ԤԼ��ʽ
    '����:����
    '����:2012-07-03
    '�����:48350
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strԤԼ��ʽ As String
    Dim rsTemp As Recordset
    strSQL = "" & _
        "Select ԤԼ��ʽ From ���˹Һż�¼ Where ��¼״̬=1 And No=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡԤԼ��ʽ", strNO)
    If rsTemp Is Nothing Then zl_GetԤԼ��ʽByNo = "": Exit Function
    If rsTemp.RecordCount = 0 Then zl_GetԤԼ��ʽByNo = "": Exit Function
    While rsTemp.EOF = False
        strԤԼ��ʽ = rsTemp!ԤԼ��ʽ
        rsTemp.MoveNext
    Wend
    zl_GetԤԼ��ʽByNo = strԤԼ��ʽ
End Function
Public Sub CreateSquareCardObject(ByRef frmMain As Object, _
    ByVal lngModule As Long, Optional cnOracle As ADODB.Connection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, IIf(cnOracle Is Nothing, gcnOracle, cnOracle), False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
End Sub
Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: �رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         'Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub

Public Function ZVal(ByVal varValue As Variant, Optional ByVal varDefault As Variant = 0) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
    Dim varTmp As Variant
    varTmp = IIf(Val(varValue) = 0, varDefault, varValue)
    ZVal = IIf(Val(varTmp) = 0, "NULL", varTmp)
End Function

Public Function zlGet֧����ʽ(ByVal lng�����ID As Long, ByVal str���㷽ʽ As String) As String
    '���ݽ��㷽ʽ����֧����ʽ
    Dim strSQL As String, rsTemp As Recordset
    '����|���㷽ʽ|�Ƿ�����|�Ƿ�ȫ��|��������
    zlGet֧����ʽ = str���㷽ʽ & "|" & str���㷽ʽ & "|1|0"
    On Error GoTo Errhand
    strSQL = "" & _
            " Select A.����,A.�Ƿ�����,A.�Ƿ�ȫ��,B.���� from ҽ�ƿ���� A,���㷽ʽ B where A.���㷽ʽ = B.���� And A.ID = [1] And A.���㷽ʽ=[2]" & _
            " Union All " & _
            " Select A.����,A.�Ƿ�����,A.�Ƿ�ȫ��,B.���� from �����ѽӿ�Ŀ¼ A,���㷽ʽ B where A.���㷽ʽ = B.���� And A.���=[1] And A.���㷽ʽ=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ֧�������㷽ʽ", lng�����ID, str���㷽ʽ)
    If Not rsTemp.EOF Then
        zlGet֧����ʽ = Nvl(rsTemp!����, str���㷽ʽ) & "|" & str���㷽ʽ & "|" & Nvl(rsTemp!�Ƿ�����, 1) & "|" & Nvl(rsTemp!�Ƿ�ȫ��, 0) & "|" & Nvl(rsTemp!����, 0)
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatiIDFromProcedure(ByVal lngModel As Long, ByVal frmParent As Object, _
                                                                ByVal strIDnumber As String, ByVal objControl As Object, _
                                                                Optional ByVal strPatiName As String = "", _
                                                                Optional ByVal strPatiSex As String = "", _
                                                                Optional ByRef blnCancel As Boolean = False) As Long
    '����:���ݲ������֤��(����,�Ա�)��ȡ����id,����id�п����Ƕ��
    '���:lngModel-ģ���
    '       frmParent-��ʾ�ĸ�����
    '       vRect-�ؼ�����Ļ�е�λ��
    '       objControl-�������֤��ˢ���֤�Ŀؼ�
    '       strIDnumber-���֤��
    '       strPatiName-��������
    '       strPatiSex-�����Ա�
    Dim strSQL As String, strPatiIDs As String
    Dim rsTmp  As ADODB.Recordset
    Dim vRect As RECT
    On Error GoTo Errhand
    strSQL = "Select zl_Custom_PatiIDs_Get([1],[2],[3],[4]) As ����IDs From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lngModel, strIDnumber, strPatiName, strPatiSex)
    If rsTmp.EOF Then
        GetPatiIDFromProcedure = 0: Exit Function
    End If
    strPatiIDs = Nvl(rsTmp!����IDs)
    If InStr(strPatiIDs, ",") > 0 Then
        strSQL = _
                    " Select /*+cardinality(B,10)*/ distinct A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����,A.�����,A.��������,A.���֤��,A.��ͥ��ַ,A.������λ " & _
                    " From ������Ϣ A, Table(f_Str2List([1])) B " & _
                    " Where a.����ID=b.Column_Value" & _
                    " Order by ����,�Ա�,����"
        strSQL = "Select  *  From (" & strSQL & ") Where Rownum < 101"
        
        vRect = zlControl.GetControlRect(objControl.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, objControl.Height, blnCancel, False, True, strPatiIDs)
        If Not rsTmp Is Nothing Then
            If Val(rsTmp!id) <> 0 Then GetPatiIDFromProcedure = Val(rsTmp!id)
        End If
    Else
        GetPatiIDFromProcedure = strPatiIDs
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


