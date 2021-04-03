VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_SLE4428"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

Private Declare Function com_open Lib "hnic32.dll" (ByVal fi_port As Integer) As Long
Private Declare Sub com_close Lib "hnic32.dll" (ByVal fl_hcom As Long)
Private Declare Function com_test Lib "hnic32.dll" (ByVal fl_hcom As Long) As Integer
Private Declare Function icc_power_on Lib "hnic32.dll" (ByVal fl_hcom As Long) As Integer
Private Declare Function icc_power_off Lib "hnic32.dll" (ByVal fl_hcom As Long) As Integer
Private Declare Function icc_type_set Lib "hnic32.dll" (ByVal fl_hcom As Long, ByVal fi_type As Integer) As Integer
Private Declare Function icc_read Lib "hnic32.dll" (ByVal fl_hcom As Long, ByVal fs_buf As String, ByVal fi_addr As Integer, ByVal fi_len As Integer) As Integer
Private Declare Function icc_write Lib "hnic32.dll" (ByVal fl_hcom As Long, ByVal fs_buf As String, ByVal fi_addr As Integer, ByVal fi_len As Integer) As Integer
Private Declare Function icc_password Lib "hnic32.dll" (ByVal fl_hcom As Long, ByVal fs_inbuf As String, ByVal fb_len As Byte) As Integer
Private Declare Function icc_testcard Lib "hnic32.dll" (ByVal fl_hcom As Long) As Integer
Private Declare Function icc_changc Lib "hnic32.dll" (ByVal fl_hcom As Long, ByVal fs_buf As String, ByVal fb_len As Byte) As Integer
Private Declare Function icc_rdpass Lib "hnic32.dll" (ByVal fl_hcom As Long, ByVal fs_buf As String, ByVal fb_len As Byte) As Integer
Private Declare Function ic_rdbaohu Lib "hnic32.dll" (ByVal fl_hcom As Long, ByVal fs_buf As String, ByVal fi_addr As Integer, ByVal fi_len As Integer) As Integer
Private Declare Function ic_wdbaohu Lib "hnic32.dll" (ByVal fl_hcom As Long, ByVal fs_buf As String, ByVal fi_addr As Integer, ByVal fi_len As Integer) As Integer
Private Declare Function icc_rdbaohu Lib "hnic32.dll" (ByVal fl_hcom As Long, ByVal fs_buf As String, ByVal fi_addr As Integer, ByVal fi_len As Integer) As Integer
Private Declare Function icc_wdbaohu Lib "hnic32.dll" (ByVal fl_hcom As Long, ByVal fs_buf As String, ByVal fi_addr As Integer, ByVal fi_len As Integer) As Integer

Dim pi_com As Integer
Dim hcom As Long    'ͨѶ�˿ھ��
Dim li_type As Integer  '������
Dim li_addr As Integer  '��ʼ��ַ
Dim li_len As Integer   '���ݳ���
Dim ls_pass As String   '������
Dim ls_outbuf As String '�������
Dim ls_inbuf As String  '��������
Dim li_ret As Integer

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    Call com_close(hcom)
End Function

Private Function clsCardDev_GetSpare() As Currency

End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean

    pi_com = CInt(GetSetting("ZLSOFT", "����ȫ��\ICCard\10", "�˿�", 0)) + 1
    li_addr = CInt(GetSetting("ZLSOFT", "����ȫ��\ICCard\10", "��ʼ��ַ", 32))
    li_len = CInt(GetSetting("ZLSOFT", "����ȫ��\ICCard\10", "����", 10))
    li_type = 3
    Call WritLog("clsCardDev_initCard", "", "���� com_open")
    hcom = com_open(pi_com)   '���ô���ΪCOM1
    Call WritLog("clsCardDev_initCard", "", "" & hcom)
    li_ret = com_test(hcom)   '����д������
    Call WritLog("clsCardDev_initCard", "", IIf(li_ret = 0, "��ʼ���ɹ�!", "��ʼ��ʧ��!"))
    If li_ret <> 0 Then Exit Function
    clsCardDev_initCard = True
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean

End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim i As Integer
    ls_outbuf = String(100, 0)
    li_ret = icc_testcard(hcom)
    Call WritLog("clsCardDev_initCard", "", IIf(li_ret = 0, "�鿨�ɹ���", "δ�忨��忨����"))
    If li_ret <> 0 Then Exit Function
    li_ret = icc_power_on(hcom)
    Call WritLog("clsCardDev_initCard", "", IIf(li_ret = 0, "IC���ϵ�ɹ���", "IC���ϵ������"))
    If li_ret <> 0 Then Exit Function
    li_ret = icc_type_set(hcom, li_type)
    Call WritLog("clsCardDev_initCard", "", IIf(li_ret = 0, "IC���������óɹ���", "IC���������ó�����"))
    If li_ret <> 0 Then Exit Function
     Call WritLog("clsCardDev_initCard", "", "li_addr:" & li_addr & "    li_len:" & li_len)
    li_ret = icc_read(hcom, ls_outbuf, li_addr, li_len)
    Call WritLog("clsCardDev_initCard", "", IIf(li_ret = 0, "�����ɹ���", "����������"))
    If li_ret <> 0 Then Exit Function
    clsCardDev_ReadCard = ls_outbuf
    Beep
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean
    
End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmParameterSet.ShowMe(10)
End Sub

Private Function clsCardDev_UploadSwap(ByVal lng����ID As Long, strOut As String) As Boolean

End Function