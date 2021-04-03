VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_SS728MQ1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

'���� �򿪴����豸
'����ԭ��  long ss_reader_open(char *szDevCom);
'����      �����������Ӧ���ն��豸�˿ڣ�������óɹ��򷵻��豸�����
'����      <1> szDevCom [in] �豸���ƣ�ȡֵ��AUTO ������COMn �������С�n ����ȡֵΪ1~9��
'����ֵ      ������������ֵΪ��С��0 ���豸��ʶ����
Private Declare Function ss_reader_open Lib "SS728M01.dll" (ByVal szDevCom As String) As Long

'���� �رմ����豸
'����ԭ��  long ss_reader_close(long icdev);
'����      �ر���ss_reader_open �������򿪵��ն��豸�˿ڡ�
'����      <1> icdev [in] �豸��ʶ����
'����ֵ ִ�гɹ�����ֵΪ0?
Private Declare Function ss_reader_close Lib "SS728M01.dll" (ByVal icdev As Long) As Long

'���� ���Ʒ�����
'����ԭ��
'long ss_dev_beep( long icdev,unsigned short _Amount,unsigned short _Msec);
'���� ���ƶ����ն˵ķ���������?
'����
'<1> icdev [in] �豸��ʶ����
'<2> _Amount [in] ���������
'<3> _Msec [in] ����ʱ�䣬ʱ����100msΪ��λ��
'����ֵ ִ�гɹ�����ֵΪ0?
Private Declare Function ss_dev_beep Lib "SS728M01.dll" (ByVal icdev As Long, ByVal amount As Integer, ByVal msec As Integer) As Long

'���� ����ָʾ��
'����ԭ��
'long ss_dev_led(long icdev,
'unsigned char _Color,
'unsigned char _Status,
'unsigned short _Amount,
'unsigned short _Msec);
'����  ���ƶ����ն˵�ָʾ�ƶ�����Ŀǰ�����ơ��������ơ�
'����
'<1> icdev [in] �豸��ʶ����
'<2> _Color [in] ��ɫ����ֵ��Χ0~2�����У�0��ʾ��ɫ��1��ʾ��ɫ��2��ʾ��ɫ��
'<3> _Status [in] ״̬����ֵ��Χ1~3�����У�1��ʾ����2��ʾ��3��ʾ��˸��
'<4> _Amount [in] ��˸������
'<5> _Msec [in] ��˸ʱ�䣬ʱ����100msΪ��λ��
'����ֵ ִ�гɹ�����ֵΪ0?
Private Declare Function ss_dev_led Lib "SS728M01.dll" (ByVal icdev As Long, ByVal color As Integer, ByVal status As Integer, ByVal amount As Integer, ByVal msec As Integer) As Long

'���� ��ȡ�ſ�����
'����ԭ��
'long ss_readmagcard(long icdev,
'char *pszTrack1,
'char *pszTrack2,
'char *pszTrack3,
'unsigned char WaitTime );
'���÷�Χ �ſ�
'���� ��ָ����ʱ���ڶ�ȡ�ſ���3���ŵ�������?
'����
'<1> icdev [in] �豸��ʶ����
'<2> pszTrack1 [out] �ŵ�1���ݣ�255���ֽ����ڵ��ַ�������'\0'������
'<3> pszTrack2 [out] �ŵ�2���ݣ�255���ֽ����ڵ��ַ�������'\0'������
'<4> pszTrack3 [out] �ŵ�3���ݣ�255���ֽ����ڵ��ַ�������'\0'������
'<5> WaitTime [in] �ȴ�ʱ�䣬��λΪ�룬��ֵ��Χ10~60��
'����ֵ ִ�гɹ�����ֵΪ0?
Private Declare Function ss_readmagcard Lib "SS728M01.dll" (ByVal icdev As Long, ByVal pszTrack1 As String, ByVal pszTrack2 As String, ByVal pszTrack3 As String, ByVal WaitTime As Byte) As Long



Private Declare Function ss_id_find_card Lib "SS728M01.dll" (ByVal icdev As Long) As Long
Private Declare Function ss_id_read_card Lib "SS728M01.dll" (ByVal icdev As Long) As Long
Private Declare Function ss_id_query_name Lib "SS728M01.dll" (ByVal icdev As Long, ByVal szName As String) As Long
Private Declare Function ss_id_query_address Lib "SS728M01.dll" (ByVal icdev As Long, ByVal szAddr As String) As Long
Private Declare Function ss_id_free_card Lib "SS728M01.dll" (ByVal icdev As Long) As Long

Private micdev As Long
Private lngReturn As Long
Private mintStart As Integer
Private mintLength As Integer

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    lngReturn = ss_reader_close(micdev)
End Function

Private Function clsCardDev_GetSpare() As Currency

End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    micdev = ss_reader_open("AUTO")
    clsCardDev_initCard = micdev > 0
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean

End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim pszTrack1 As String * 255, pszTrack2 As String * 255, pszTrack3 As String * 255
    Dim WaitTime As Byte
    WaitTime = 10
    
    mintStart = CInt(GetSetting("ZLSOFT", "����ȫ��\ICCard\13", "��ʼ��ַ", 2))
    mintStart = IIf(mintStart = 0, 2, mintStart)
    mintLength = CInt(GetSetting("ZLSOFT", "����ȫ��\ICCard\14", "����", 16))
    
    lngReturn = ss_readmagcard(micdev, pszTrack1, pszTrack2, pszTrack3, WaitTime)
    If lngReturn = 0 Then
        clsCardDev_ReadCard = Trim(Mid(pszTrack2, mintStart, mintLength))
    Else
        clsCardDev_ReadCard = ""
    End If
'    Dim rtn As Long
'    Dim szName As String * 30
'    Dim szAddr As String * 70
'
'    rtn = ss_id_find_card(micdev)
'    If rtn = 0 Then
'        rtn = ss_id_read_card(micdev)
'        If rtn = 0 Then
'            rtn = ss_id_query_name(micdev, szName)
'            rtn = ss_id_query_address(micdev, szAddr)
'            Debug.Print szName
'            Debug.Print szAddr
'        End If
'    End If
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean

End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmParameterSet.ShowMe(13)
End Sub

Private Function clsCardDev_UploadSwap(ByVal lng����ID As Long, strOut As String) As Boolean
    
End Function