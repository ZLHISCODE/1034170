Attribute VB_Name = "mdlMain"
Option Explicit


Public Sub Main()
'------------------------------------------------
'���ܣ������򣬸���������Ƭ����
'       ����������������õĽӿڷ���������C#����ʾ��Ƭվ������ɱ��ػ��棬�����ȴӱ��ػ����ж�ȡͼ��
'       ����lngOrderID��strImages�����ݿ��в���ͼ���ļ�������һ������ͼ�񣬵���һ�α��ӿ�
'       ����ʹ�������ݿ����Ӵ�����֧�� 10.35.10֮���HIS�汾
'������
'���أ���
'-----------------------------------------------
'����Ĳ������壬���������ӷ��������ַ���{+}��
    '������ʽ��strImages{+}lngOrderID{+}strDBConnection{+}blnMoved{+}bAdd{+}intImageInterval{+}lngSys{+}blnReconnectDB
    '�������ͣ� strImages --- ͼ���,�����ǡ�����UID1|1-3;5-27;33-100+����UID2|ȫ����,ȫ����ʾ��ȫ��ͼ��
    '           lngOrderID --- ҽ��ID
    '           strDBConnection --- ���ݿ����Ӵ���������������[+]�û���[+]����[+]�����Ƿ�ת���������ӷ��������ַ���[+]��
    '                          �������롱���û���¼����ʱ���������Ƿ�ת����=1���������롱�����ݿ��¼����ʱ���������Ƿ�ת����=0
    '           blnMoved --- �����Ƿ�ת��
    '           bAdd --- ��ѡ������Ĭ��ֵFalse����ͼ�������ӽ���Ƭվ�������滻ԭ��Ƭվ����ͼ��TrueΪ���ӣ�FasleΪ�滻
    '           intImageInterval --- ��ѡ������Ĭ��ֵ0����ͼ��ļ����ֻ�Դ�ȫ������,��������ͼ������>100ʱ��Ч
    '           lngSys --- ��ѡ������Ĭ��,100��ϵͳ���
    '           blnReconnectDB --- ��ѡ������Ĭ��ֵFalse���Ƿ������������ݿ⡣��һ�δ򿪹�Ƭʱ�Զ��������ݿ⣬֮���ٴ򿪹�Ƭ��
    '                           ��blnReconnectDB���������Ƿ������������ݿ⡣
    '                           =True��ʹ��strDBConnection���������������ݿ⣻=False�����������������ݿ⣬ʹ�ù�Ƭ�������ڵ����ݿ�����
    '
    
    Dim strMsgs As String
    Dim lngWinHandle As Long
    Dim strDBConnection As String
    
    On Error GoTo err
    
    '�ȴ�����־�ļ�Ŀ¼
    gstrLogPath = GetLogDir()
    
    '��exe����Ĳ���������strMsgs �������ȴ������������޲�����exe���ã�ֱ���˳�
    strMsgs = Command
    If Trim(strMsgs) = "" Then Exit Sub
    
    '����������Ѿ�������һ�Σ�����������ֱ��ˢ�½������ݺ��˳�
    If App.PrevInstance Then
        Call WriteCommLog("zlSoftViewImage.Sub Main", "��ｫ��Ϣ���͸��Ѵ��ڵ�zlSoftViewImage����ǰ�����˳����汾Ϊ��" & App.Major & "." & App.Minor & "." & App.Revision, "����Ϊ��strMsgs = " & strMsgs, ltDebug)
        lngWinHandle = FindWindow(vbNullString, HIS_CAPTION)
        If lngWinHandle = 0 Then
            Call WriteCommLog("zlSoftViewImage.Sub Main", "�����Ҵ��ھ��=0", "�Զ�ɱ��ԭ���Ľ��̣������½��̡�", ltDebug)
            Shell "taskkill /f /im zlSoftViewImage.exe"
        Else
            Call SendMsg(strMsgs)
            Exit Sub
        End If
    Else
        Call WriteCommLog("zlSoftViewImage.Sub Main", "����һ������zlSoftViewImage.�汾Ϊ��" & App.Major & "." & App.Minor & "." & App.Revision, "����Ϊ��strMsgs = " & strMsgs, ltDebug)
    End If
    
    '���յ�QUIT��ֱ���˳�
    If UCase(Trim(strMsgs)) = "QUIT" Then Exit Sub
    
    '��ʼ��comlib�����ݿ�����
    If UBound(Split(strMsgs, MSG_SPLIT)) >= 3 Then
        strDBConnection = Split(strMsgs, MSG_SPLIT)(2)
        
        If UBound(Split(strDBConnection, "[+]")) = 3 Then
            gstrZLHIS�����ַ��� = Split(strDBConnection, "[+]")(0)
            gstr�û��� = Split(strDBConnection, "[+]")(1)
            gstr���� = Split(strDBConnection, "[+]")(2)
            gbln�Ƿ�ת������ = Val(Split(strDBConnection, "[+]")(3)) = 1
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    Call InitInterface(gstr�û���)
    
    '������Ϣ�����壬������Ϣhook��Ȼ�����ش���
    If gfrmViewImage Is Nothing Then Set gfrmViewImage = New frmViewImage
    Call gfrmViewImage.ShowMe(True)
    gfrmViewImage.Hide
    
    '������Ϣ
    Call ProcessMessage(strMsgs)
    
    Exit Sub
err:
    If errHandle("exe Main", "��ʾ��Ƭ���ڳ��ִ���") = 1 Then Resume
End Sub

Private Sub SendMsg(ByVal strMsg As String)
'------------------------------------------------
'���ܣ�����Ϣ���͸���Ϣѭ��������
'������strMsg -- ����exeʱ����Ĳ�����
'���أ���
'-----------------------------------------------
    Dim lngWinHandle As Long        '��Ҫ������Ϣ�ġ�zlSoftViewImage.exe������Ĵ��ھ��
    Dim wParam As Long
    Dim lResult As Long
    Dim strTemp As String
    Dim buf(1 To 1024) As Byte
    
    '��Ϣ���壺wParam = 223��dss��dwData = 33 ������Ϣ��dwData = 32 �˳�
    wParam = 223
   
    Call CopyMemory(buf(1), ByVal strMsg, LenB(StrConv(strMsg, vbFromUnicode)))
    
    'dss.dwData�����Ϣ���ã�ֻ��˫�������һ����Ƕ���
    If UCase(Trim(strMsg)) = "QUIT" Then
        dss.dwData = 32 '���Ϊ�ر����д���
    Else
        dss.dwData = 33 '���Ϊˢ�´��ڻ��ߴ��´���
    End If
    
    dss.cbData = LenB(StrConv(strMsg, vbFromUnicode)) + 1
    
    'ʹ��buf���ͣ����Կ�����Ϣ��1024֮��
    dss.lpData = VarPtr(buf(1))
    
    '������Ϣѭ��������
    lngWinHandle = FindWindow(vbNullString, HIS_CAPTION)
    
    Call WriteCommLog("zlSoftViewImage.SendMsg", "����Ϣ���͸���Ϣѭ��������", "��ϢΪ��" & strMsg & "�����ھ��Ϊ��" & lngWinHandle, ltDebug)
    
    If lngWinHandle <> 0 Then
        lResult = SendMessage(lngWinHandle, WM_COPYDATA, wParam, dss)
    End If
End Sub



    
    
