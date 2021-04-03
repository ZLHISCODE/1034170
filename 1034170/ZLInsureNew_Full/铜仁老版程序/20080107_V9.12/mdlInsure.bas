Attribute VB_Name = "mdlInsure"
Option Explicit

Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As String
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_BROWSEFORCOMPUTER = &H1000  'Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000   'Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000 'Browsing for Everything
Private Const CSIDL_NETWORK As Long = &H12

Private Const MAX_PATH = 260
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2
Private Const LVM_SETCOLUMNWIDTH = &H101E

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'���뷨����API----------------------------------------------------------------------------------------------
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Public Const KLF_REORDER = &H8

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

'���ı������м��ܻ���ܵĺ���
Public Declare Function EncryptStr Lib "FTP_Trans.dll" (ByVal SourceStr As String, ByVal Key As String, ByVal IsEncrypt As Boolean) As String

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    վ�� As String
End Type
Public UserInfo As TYPE_USER_INFO

Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public glngSys As Long                      'ϵͳ��Ų���
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSQL As String                    '������Ϊ������ʱSQL���

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public gstrUserName As String               '��ǰ�û�����
Public gstr��λ���� As String
Public gbln�������� As Boolean              '����ר��,���ڷ����Ƿ�Ϊ��������
Public gstr���ⲡ�� As String               '���ⲡ����

Public gstrMatchMethod As String    'ƥ�䷽ʽ:0��ʾ˫��ƥ��

Public gintInsure As Integer
Public gstrҽԺ���� As String * 10               'ҽԺ���
Public gstrҽ���������� As String

Public Type T��������
    ����ID       As Long
    ���         As Long
    סԺ����     As Long
    �ʻ��ۼ�����   As Currency
    �ʻ��ۼ�֧��   As Currency
    �ۼƽ���ͳ��   As Currency
    �ۼ�ͳ�ﱨ��   As Currency
    ����         As Currency
    �ⶥ��         As Currency
    ʵ������     As Currency
    �������ý��   As Currency
    ȫ�Էѽ��   As Currency
    �����Ը����   As Currency
    ����ͳ����   As Currency
    ͳ�ﱨ�����   As Currency
    �����Ը����   As Currency
    �����ʻ�֧��   As Currency
    ֧��˳���     As String
    ��ҳID         As Long
    ��;����       As Long
    סԺ����       As Long
End Type
Public g�������� As T��������           '����Ԥ����֮�����Ľ������������д���ս����¼
Public gcol������� As New Collection   '����Ԥ����֮�����Ľ������������д���ս������
                                        'ÿ����ԱΪһ�����飬����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������

Public Enum ҽ��Enum
    TYPE_������ = 10
    TYPE_�������� = 11                  'Modified by ZYB ##2002-10-28
    TYPE_�����ɽ = 12
    TYPE_��������ɽ = 13
    TYPE_���������� = 14
    TYPE_�ɶ��� = 20
    TYPE_�ɶ����� = 21
    TYPE_�ɶ����� = 22
    TYPE_�ɶ��ϳ� = 23
    type_���� = 24
    TYPE_�Ĵ�üɽ = 25
    TYPE_��ɽ = 26
    TYPE_����ʡ = 30
    TYPE_������ = 31
    TYPE_���Ͻ�ˮ = 32
    TYPE_�Թ��� = 40
    TYPE_������ = 43
    TYPE_������ = 50
    TYPE_������ = 61
    TYPE_�������� = 70
    'Modified by ���� 20031218 ����������
    TYPE_����ʡ = 71
    TYPE_������ = 72
    TYPE_��ƽ�� = 73
    TYPE_������ = 80
    TYPE_ͭ�� = 81
    TYPE_������ = 82
    TYPE_���������� = 83
    TYPE_���� = 84
    TYPE_���� = 85
    TYPE_�ش�У԰�� = 86    '���˺�:200403
    TYPE_���� = 90
    TYPE_��Ԫ = 87
End Enum

Public Enum ���Enum
    balan���� = 10
    balan��Ժ = 20
    balanԤ�� = 30
    balan���� = 40
End Enum

Public Enum �����֤Enum
    id�����շ� = 0
    id��Ժ�Ǽ� = 1
    id�ʻ����� = 2
    id�Һ� = 3
    id���� = 4
    id����ȷ�� = 5
End Enum

Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
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
End Enum

Public Function GetErrInfo(strCode As String) As String
'���ܣ����ݴ�����뷵�ش�����Ϣ
'������bytType=�������,strCode=�������
    Dim rsTmpErr As New ADODB.Recordset
    
    Select Case gintInsure
        Case TYPE_����ʡ, TYPE_������, TYPE_���Ͻ�ˮ
            Select Case strCode
                Case "0000":      GetErrInfo = "����"
                Case "0001":      GetErrInfo = "�޷���ȡ�����ļ�����رձ�������������ļ���"
                Case "0002":      GetErrInfo = "��Ӧ�ó������������ʧ��(�޷��ҵ�Ӧ�ó��������),��ȷ��Socket Server�Ƿ���������!"
                Case "0003":      GetErrInfo = "Ӧ�ó�������������޷���ɽ���!"
                Case "0004":      GetErrInfo = "�޷��õ�ϵͳ������Ϣ!"
                Case "0005":      GetErrInfo = "�Ҳ����α����������ĵĳ������������������������!"
                Case "0009":      GetErrInfo = "���������ÿ��Ŷ�Ӧ�ķ����ı��"
                Case "1":         GetErrInfo = "�ն��豸��֧�ִ˹���"
                Case "10":        GetErrInfo = "��֤����,����ĸ����������"
                Case "1001":      GetErrInfo = "˳��ų��ȷǷ�"
                Case "1002":      GetErrInfo = "�շ���Ŀ�������Ƿ�"
                Case "1003":      GetErrInfo = "�շ���Ŀ����Ƿ�"
                Case "1004":      GetErrInfo = "������۸���Ϊ��"
                Case "1005":      GetErrInfo = "������۸���С�ڵ���0"
                Case "1006":      GetErrInfo = "����������Ƿ�"
                Case "11":        GetErrInfo = "֧������,����"
                Case "1101":      GetErrInfo = "˳��Ŵ���"
                Case "1102":      GetErrInfo = "�����ѽ��㲻���ٴ��ݷ�����ϸ"
                Case "1103":      GetErrInfo = "û�м�������Ҫ�޸ĵķ�����ϸ����!�����������������ȷ!"
                Case "1104":      GetErrInfo = "�޸ķ�����ϸ���ϳ���!"
                Case "1105":      GetErrInfo = "�ò���סԺ�����ѽ����!������Ŀ����ͬ���ദ��"
                Case "12":        GetErrInfo = "֧������,�û�����ʼ��ʧ��"
                Case "13":        GetErrInfo = "֧������,SAM����ʼ��ʧ��"
                Case "14":        GetErrInfo = "֧������,�û�����֤MAC1ʧ��"
                Case "15":        GetErrInfo = "֧������,SAM����֤MAC2ʧ��"
                Case "16":        GetErrInfo = "������,��ȡ���ʧ��"
                Case "17":        GetErrInfo = "���¶�̬��Ϣ,�û������´���"
                Case "18":        GetErrInfo = "δ֪�����"
                Case "19":        GetErrInfo = "���¶�̬��Ϣ,PSAM����ȡ����"
                Case "2":         GetErrInfo = "���׳�ʼ��,��ⲻ���ն��豸+���豸����"
                Case "20":        GetErrInfo = "�޴�������"
                Case "2001":      GetErrInfo = "�����˻�������ƷǷ�,�����ٽ��н���!"
                Case "21":        GetErrInfo = "֧������,TACУ�����"
                Case "2101":      GetErrInfo = "�����Ѱ����Ժ����,�����ٽ��н���"
                Case "2102":      GetErrInfo = "��������δͨ������Ժ�ڼ�ķ���Ϊȫ�Է�"
                Case "2103":      GetErrInfo = "���ý���ʱ��⵽�洢���̵����������˳��š�λ������ȷ!"
                Case "22":        GetErrInfo = "Ȧ�潻��,MAC1У�����"
                Case "2201":      GetErrInfo = "������ȡ���˵�֧������޷����з��ý���<bnzxx>��"
                Case "2202":      GetErrInfo = "������ȡ���ⲡ�˵�֧������޷����з��ý���<by21bzxx>��"
                Case "2203":      GetErrInfo = "���ý���ʱ��By10cyjsb��д���ݳ���!"
                Case "2204":      GetErrInfo = "���û���ʧ��!"
                Case "2205":      GetErrInfo = "Ԥ����ʧ��!"
                Case "2206":      GetErrInfo = "����Ա�洢����ִ�г���!"
                Case "2207":      GetErrInfo = "�����Ƚ��з��õ������֮���ٽ��н������!"
                Case "2208":      GetErrInfo = "û����Ч��Ԥ�����¼���޷�����!"
                Case "2209":      GetErrInfo = "סԺ�����Ѿ������·����ƣ�ϵͳֻ���屾�µķ���!"
                Case "2210":      GetErrInfo = "δ��ѯ������Ǽǻ�������ڳ���������˵�������ޣ�"
                Case "2211":      GetErrInfo = "��ǰ���˽����¼�������һ�ν��㣬����ѵ�ǰ�����¼֮������н����¼����֮����ܽ��л���ҵ�����!"
                Case "23":        GetErrInfo = "Ȧ�潻��,TACУ�����"
                Case "24":        GetErrInfo = "Ȧ�潻��,�û�����ʼ��ʧ��"
                Case "25":        GetErrInfo = "Ȧ�潻��,�û�����֤MAC2ʧ��"
                Case "29":        GetErrInfo = "��������ʧ��"
                Case "3":         GetErrInfo = "���׳�ʼ��,��ⲻ��PSAM��"
                Case "30":        GetErrInfo = "�޴˽��״���"
                Case "3001":      GetErrInfo = "����ҽԺ����Ƿ������������ļ������ã�"
                Case "3002":      GetErrInfo = "�����Ͳ�����δ�忨���루�أ��忨��"
                Case "3003":      GetErrInfo = "�޷���ȡ������Ϣ�������ԣ�"
                Case "31":        GetErrInfo = "���ͽ�������ʧ��,����ͨѶ�˿�"
                Case "3100":      GetErrInfo = "�����Ѱ���סԺ�Ǽǣ����ܽ��и���ҵ��"
                Case "3101":      GetErrInfo = "������ʹ�õĿ��Ƿ�������ƾ������ҽ��������"
                Case "3102":      GetErrInfo = "����������ҽԺδ���н��㣬�޷����и�ҵ��"
                Case "3103":      GetErrInfo = "δ�����˿��Ļ������ϣ�����������!"
                Case "3104":      GetErrInfo = "���ܻ�ȡ���˵ĳ������ڣ����ܽ���Ժ�Ǽǣ�"
                Case "3105":      GetErrInfo = "�ò�����/סԺ���Ѿ���ռ�ã���������סԺ��/�����ţ�"
                Case "3106":      GetErrInfo = "���ܼ�������ҽԺ�ȼ������ݣ����ʵҽԺ���룡"
                Case "3107":      GetErrInfo = "������鲻��ȷ,����������!"
                Case "3108":      GetErrInfo = "�޷�����������Ⱥ��ı�������!"
                Case "3109":      GetErrInfo = "ִ��������Ⱥ�Ĵ洢���̡������ִ���!"
                Case "3110":      GetErrInfo = "ҽ�����ı�����IC��ʵ�ʼ�¼��ҽ�����ı��벻һ��!"
                Case "3111":      GetErrInfo = "����Ա�������ʱ����!"
                Case "3128":      GetErrInfo = "��ǰ���˵ľ�������Ǽ������ִ�м���תסԺ������"
                Case "32":        GetErrInfo = "������Ӧ���ݳ�ʱ���ױ�ȡ��,����ͨѶ�˿�"
                Case "33":        GetErrInfo = "У����Ӧ����(ETX)����"
                Case "34":        GetErrInfo = "У����Ӧ����(LRC)����"
                Case "35":        GetErrInfo = "У����Ӧ����(STX)����"
                Case "36":        GetErrInfo = "У����Ӧ���ݴ�����Կ����"
                Case "37":        GetErrInfo = "������Ӧ����ʧ��,����ͨѶ�˿�"
                Case "38":        GetErrInfo = "δ֪����,������ȡ��"
                Case "4":         GetErrInfo = "���׳�ʼ��,PSAM����ȡ����"
                Case "40":        GetErrInfo = "ͨѶ����"
                Case "4001":      GetErrInfo = "�ļ����Ƿ�"
                Case "4002":      GetErrInfo = "д�ļ����̳���"
                Case "41":        GetErrInfo = "Ȧ�潻����֤ʧ�ܣ��뽫������ҽ�����Ĵ���"
                Case "4101":      GetErrInfo = "�޷�����ϸ��Ϣ"
                Case "42":        GetErrInfo = "�ſ�����"
                Case "5":         GetErrInfo = "���׳�ʼ��,��ⲻ���û���"
                Case "5001":      GetErrInfo = "֧��ԭ�򳤶ȷǷ�,֧��ʧ��"
                Case "5002":      GetErrInfo = "֧�����Ӧ�ô���0,֧��ʧ��"
                Case "5003":      GetErrInfo = "֧������ڿ������,֧��ʧ��"
                Case "5004":      GetErrInfo = "д��ʧ��"
                Case "5101":      GetErrInfo = "�����ѳ�Ժ��˳��Ŵ���"
                Case "5102":      GetErrInfo = "������Ǽ�ʱ��ʹ�õĿ�����"
                Case "5103":      GetErrInfo = "�޷��õ���Ч�ĸ����˻�֧�����ֽ�֧����"
                Case "5104":      GetErrInfo = "��Ա��δ���з��ô���/���߼��������ò�Ա��֧�����ݡ�"
                Case "6":         GetErrInfo = "���׳�ʼ��,�Ǳ�ϵͳ��"
                Case "6101":      GetErrInfo = "�޷����ܴ���"
                Case "6102":      GetErrInfo = "��֧������ڷ����ܶ�޷�֧��"
                Case "7":         GetErrInfo = "�û�����ȡ����"
                Case "7101":      GetErrInfo = "���������ݿ�����ʧ��,��ȷ�����糩ͨ�Լ�NET8����������ȷ!"
                Case "7102":      GetErrInfo = "��ǰ�û����ݿ�����ʧ��,��ȷ��Socket Server�Ƿ���������!"
                Case "7103":      GetErrInfo = "ʡҽ��������������ʧ��"
                Case "7104":      GetErrInfo = "�����Ѿ���ȡ����"
                Case "7106":      GetErrInfo = "��ҽ��������������ʧ��"
                Case "7107":      GetErrInfo = "ʡ��ҽ��������������ʧ��!"
                Case "8":         GetErrInfo = "��֤���Ų���"
                Case "8001":      GetErrInfo = "��ȡ���˻�����Ϣ����"
                Case "8002":      GetErrInfo = "������Ч��������¼δ���㣬����������"
                Case "8003":      GetErrInfo = "�Ѿ�������Ժ�Ǽ�"
                Case "8004":      GetErrInfo = "����Ա�޷�����ҽ�ƴ���"
                Case "8005":      GetErrInfo = "���ҽ�ƴ��������ʸ�ʱ��ϵͳ����"
                Case "8006":      GetErrInfo = "���ȵ�ҽ�����Ľ����ʸ�����"
                Case "8007":      GetErrInfo = "�������ʱ����"
                Case "8008":      GetErrInfo = "�����ڷ���״̬"
                Case "8009":      GetErrInfo = "���ý���ʱ��ϵͳ����"
                Case "8010":      GetErrInfo = "û����Ч��������¼"
                Case "8011":      GetErrInfo = "��δ�����������¼�����ʵ"
                Case "8012":      GetErrInfo = "ȫ�ԷѲ��ּӹҹ��ԷѲ��ִ��ڷ����ܶ���ʵ"
                Case "8013":      GetErrInfo = "��������¼����δͨ������Ϊ�����ܴ������㣬ȫ���Է�"
                Case "8014":      GetErrInfo = "����Ա����˷����޷�����ҽ�ƴ���"
                Case "8015":      GetErrInfo = "����Ա��ҽ�Ʊ����չ���Ա"
                Case "8016":      GetErrInfo = "����ԱΪҽ�Ʊ����չ���Ա"
                Case "8017":      GetErrInfo = "��ǰ�����ѽ���"
                Case "8080":      GetErrInfo = "ҽ��������δ�ڹ�ҽԺ��ͨ����ҽ��ҵ��,����ҽ��������ϵ"
                Case "9":         GetErrInfo = "��֤����,�û����������뱻��"
                Case "9001":      GetErrInfo = "Ӧ�÷�����ִ�д洢����/�������������"
                Case "9002":      GetErrInfo = "�������ӵ��������ݿ�(hisint/hisintkm),�޷����н��״���!"
                Case "9003":      GetErrInfo = "�򱾵����ݿ����ύ�����޸ĳ���,�޷������ݽ����ύ���߻ع�!"
                Case "9004":      GetErrInfo = "�ò��˵Ļ������ϻ�û�еǼǻ����Ѿ��ύ�ɹ�,�޷��ع�����!"
                Case "9005":      GetErrInfo = "���ݿ���δ�������ò��˵�δ�ύ����������,�޷��ύ����!"
                Case "9006":      GetErrInfo = "�ⲿӦ�ó������������ƺŵ�λ����Ϊ18λ!"
                Case "9201":      GetErrInfo = "��ѯ�ֶη�����ϸ��¼������!"
                Case "9202":      GetErrInfo = "��ѯסԺ���˴��������¼������!"
                Case "9203":      GetErrInfo = "��Ч�Ĳ�ѯ���!"
                Case "9204":      GetErrInfo = "���������������ʧ��,�޷��������µı����Ϣ!"
                Case "9205":      GetErrInfo = "���������Ϣ��ѯ����!"
                Case "9301":      GetErrInfo = "�޷���λ���˵�ҽ���������޷�����!"
                Case "9996":      GetErrInfo = "ʡ�������ݴ���ʧ��!"
                Case "9997":      GetErrInfo = "���������ݴ���ʧ��!"
                Case "9998":      GetErrInfo = "ʡ/���������ݴ���ʧ��!"
                Case "9999":      GetErrInfo = "Ӧ�÷����������쳣����"
                Case Else
                    GetErrInfo = "ҽ��֧�ֲ��ֳ��ִ���"
            End Select
            GetErrInfo = GetErrInfo & "[������롪" & strCode & "]"
        Case TYPE_�ɶ���
            gstrSQL = "select errtext from errcode where code='" & strCode & "'"
            rsTmpErr.CursorLocation = adUseClient
            rsTmpErr.Open gstrSQL, gcnSybase, adOpenKeyset
            If Not rsTmpErr.EOF Then
                GetErrInfo = IIf(IsNull(rsTmpErr!errtext), "δ֪ԭ��Ĵ���", rsTmpErr!errtext)
            Else
                GetErrInfo = "δ֪ԭ��Ĵ���"
            End If
        Case TYPE_����������, TYPE_������
               Select Case Val(strCode)
                    Case 0: GetErrInfo = "����"
                    Case -2:      GetErrInfo = "�����ڴ��ϵͳ������������ϵͳ���ܽ��"
                Case -1001, -1003, -1004, -1005 - 1006 - 1007:
                        GetErrInfo = "ϵͳ�������������ȷ�������Ƿ���������!"
                Case -1002:
                        GetErrInfo = "���������Ӵ���,������ԭ�������:" & vbCrLf & _
                                     "    ��1�����粻ͨ" & vbCrLf & _
                                     "    ��2�������������ʧ��" & vbCrLf & _
                                     "    ��3���ͻ������ô���" & vbCrLf & _
                                     "    ��4���ͻ�������ҽԺ�������" & vbCrLf & _
                                     "����취Ϊ:ȷ�������Ƿ�������ȷ�Ϸ����Ƿ���������"
                Case -5555
                    GetErrInfo = "��������������IC���Ƿ�����������Ͳ�ƥ��!"
                '�ܺ�ȫ���� 2003-12-17
                Case -5556
                    GetErrInfo = "���Ų�һ�£�"
                Case -6001, -6002, -6003, -6004, -6005, -6007, -6008
                    GetErrInfo = "ϵͳ�������ݽ���ʱ����,����ϵͳ�ļ�package.dat" & vbCrLf & _
                                 "�ļ��⵽�ƻ��򴫵ݵĲ���ֵ����!"
                Case -6009
                    GetErrInfo = "�����е�ҽԺ��ź�ע���ҽԺ��Ų�һ��!"
                Case -6006, -7001
                    GetErrInfo = "ϵͳ���кϷ�����֤���󣬿�������ϵͳ��������" & vbCrLf & _
                                 "��ע������Ƿ�ʹ�ã�ҽ�����Ľ��������ҽԺ�Ľ���������!"
                Case 1001
                    GetErrInfo = "�����ڸñ���!"
                Case 1002
                    GetErrInfo = "���Ŵ��鿨��!"
                Case 1003
                    GetErrInfo = "ֹ�������֧�������鿨!"
                Case 1004
                    GetErrInfo = "��ҽԺ�ֵ䣬ҽԺ��Ŵ�!"
                Case 1005
                    GetErrInfo = "ҽԺ�ѱ�����!"
                Case 1007
                    GetErrInfo = "����ʱ����ڵ�ǰϵͳʱ�䣬Ӧ�ô���"
                Case 1008
                    GetErrInfo = "��������ظ���IC���ϵ����ݴ������鿨��"
                '�ܺ�ȫ���� 2003-12-17
                Case 1009: GetErrInfo = "У�����ݴ������η�Ϊ��������"
                Case 1011
                    GetErrInfo = "������Ϣ�вα��˵Ļ�����Ϣ�������鿨��"
                Case 1016
                    GetErrInfo = "���ķ�����ͣ�����ڽ��и��£��������Ӻ����ԣ�"
                Case 1020: GetErrInfo = "�Ƿ�����ҵ�����ʹ���Ӧ�ô���"
                Case 1022: GetErrInfo = "���������"
                Case 1023: GetErrInfo = "������סԺ��"
                Case 1024: GetErrInfo = "�������ز���"
                Case 1025: GetErrInfo = "���ķ�������뼰ʱ�����ķ������Ա��ϵ��"
                Case 1026: GetErrInfo = "�������кŴ���ҽԺϵͳ���Ƿ��������⵽�ƻ���"
                Case 1027: GetErrInfo = "�����˻����ⲻһ�£����鿨��"
                Case 1028: GetErrInfo = "�Ƿ�����ʱ�䣬���ķ�����ͣ��"
                Case 1030: GetErrInfo = "ҽԺ��������㷨����"
                Case 1031: GetErrInfo = "ҽԺ������˽����㷨����"
                Case 1032: GetErrInfo = "����������"
                Case 1033: GetErrInfo = "ҽԺסԺ�����㷨����"
                Case 1034: GetErrInfo = "ҽԺסԺ���˽����㷨����"
                Case 1035: GetErrInfo = "ͳ���ۼƴ���"
                Case 1036: GetErrInfo = "д����Ŵ�"
                Case 1037: GetErrInfo = "���Ŵ���"
                Case 1038: GetErrInfo = "�˻�������"
                Case 1039: GetErrInfo = "סԺû��סԺ������"
                Case 1041: GetErrInfo = "ת��û��ת�ﵥ��"
                Case 1042: GetErrInfo = "����������"
                Case 1043: GetErrInfo = "ҽԺ��ͳ�ڱ���Ϊ�㣬������סԺ��"
                Case 1044: GetErrInfo = "�˻���������Ĳ�����"
                Case 1045: GetErrInfo = "�˻����ͬ��"
                Case 1046: GetErrInfo = "ҽԺ����󲡽����㷨����"
                Case 1047: GetErrInfo = "ҽԺ����󲡳��˽����㷨����"
                Case 1048: GetErrInfo = "�����������"
                Case 1049: GetErrInfo = "ҽԺ������˽������ݴ���"
                Case 1050: GetErrInfo = "����ʱ�����������"
                Case 1052: GetErrInfo = "סԺ�����ѵǼ���Ժ��"
                Case 1053: GetErrInfo = "�ò���δ�Ǽǲ���Ժ��"
                Case 1054: GetErrInfo = "�ô������ƣ����Ѿ����"
                Case 1058: GetErrInfo = "�󲡵Ǽ�ʱ�󲡱��벻��Ϊ�գ�"
                Case 1059: GetErrInfo = "��ҽ�޶�ҽԺ����"
                Case 1062: GetErrInfo = "ת�ﵥ00000E�Ļ�������С��70��"
                Case 1063: GetErrInfo = "ת�ﵥ00000E�Ļ������֤�Ŵ���"
                Case 1064: GetErrInfo = "��Ժ�ĵǼ����ں���Ժ���ڴ���"
                Case 1301: GetErrInfo = "�󲡱����Ѵ��ڣ�"
                Case 1302: GetErrInfo = "�޴˴󲡱��룡"
                Case 1303: GetErrInfo = "�����ƴ󲡱����Ѵ��ڣ�"
                Case 1304: GetErrInfo = "�޴������ƴ󲡱��� , û�ö�Ӧ�������ʻ���"
                Case 1305: GetErrInfo = "��ת�ﵥ���Ѵ��ڣ�"
                Case 1306: GetErrInfo = "�޴�ת�ﵥ�ţ�"
                Case 1307: GetErrInfo = "���޶�ҽԺ�Ѵ��ڣ�"
                Case 1308: GetErrInfo = "�޴��޶�ҽԺ��"
                Case 7001, 7002, 7003, 7004, 7005
                    GetErrInfo = "���������Ӵ���ϵͳʹ�õĶ�����̬���ӿ����IC������"
                Case 7006: GetErrInfo = "ϵͳ���󣬵��붯̬���ӿ����"
                Case 7007: GetErrInfo = "д��ʱУ�鿨�������� (���ܿ�������)��"
                Case -8001: GetErrInfo = "������У�����"
                Case -8002: GetErrInfo = "ҽԺ��Ŵ���ϵͳ���ô���"
                Case -8003: GetErrInfo = "ϵͳ�汾������Ҫ���¿ͻ��ĳ���"
                Case -8004: GetErrInfo = "ϵͳ���ڴ�����Ҫ���Ŀͻ������ڣ�"
                Case 1401: GetErrInfo = "������ҽ�����㣡"
                Case 1402: GetErrInfo = "�����ﲻ���ڣ�"
                Case 1403: GetErrInfo = "�Ҵ����Ѵ����"
                Case 1404: GetErrInfo = "�α�����1����2 (ҽ�Ʊ��ղ�����)��"
                Case 1405: GetErrInfo = "����������������գ�"
                Case 1406: GetErrInfo = "��������й��˱��գ�"
                Case 1407: GetErrInfo = "��Ϊ��Ժ״̬���������������"
                Case 1408: GetErrInfo = "��Ч����"
                Case 1409: GetErrInfo = "���ʽ���ͳ���ۼƳ��ָ�ֵ��"
                Case 1410: GetErrInfo = "�Ƿ�ҽԺ , ҽԺ�����ڣ�"
                Case 1411: GetErrInfo = "������ת�ﵥ�ѱ����ã�"
                Case 1412: GetErrInfo = "�Ƿ����ڸ�ʽ Ӧ��'YYYYMMDD'"
                Case 1413: GetErrInfo = "�Ƿ�����ʱ���ʽ Ӧ��'YYYYMMDDhhmmss'��"
                Case 1414: GetErrInfo = "סԺ�Ǽǵ�ʱ��д����Ų���ȣ�"
                Case 1415: GetErrInfo = "���ҽԺ����������סԺ��"
                Case 1416: GetErrInfo = "���ҽԺ�����Թ���סԺ��"
                Case 1417: GetErrInfo = "�������˽������"
                Case 1418: GetErrInfo = "���˳��˽������"
                Case 1419: GetErrInfo = "�������˽������"
                '�ܺ�ȫ���� 2003-12-17
                '�������´������
                Case 1424: GetErrInfo = "�����ܶ���ָ�����"
                Case 1425: GetErrInfo = "���㷽ʽ����Ժ�ǼǷ�ʽ��һ�£�"
                Case 1427: GetErrInfo = "��Ժ���ڴ��ڳ�Ժ���ڻ��߳�Ժ���ڴ��ڽ������ڣ�"
                Case Else
                    GetErrInfo = "ҽ��֧�ֲ��ֳ��ִ���"
                End Select
                '�ܺ�ȫ���� 2003-12-17
                'ͬʱ���ϴ����ţ��Է�����
                GetErrInfo = "�����ţ�" & strCode & vbCr & "����������" & GetErrInfo
        Case TYPE_�ش�У԰��
                Select Case Val(strCode)
                Case 0: GetErrInfo = " �ɹ�"
                Case -1: GetErrInfo = "�򿪴���ʧ��"
                Case -2: GetErrInfo = "��д������ʧ��"
                Case -3: GetErrInfo = "��������"
                Case -4: GetErrInfo = "��ʱ����"
                Case -5: GetErrInfo = "�޿�"
                Case -6: GetErrInfo = "�û�������"
                Case -7: GetErrInfo = "��������"
                Case -8: GetErrInfo = "д������"
                Case -9: GetErrInfo = "��ֵʧ��"
                Case -10: GetErrInfo = "��ֵʧ��"
                Case -11: GetErrInfo = "����Licence����"
                Case -12: GetErrInfo = "Licence����"
                Case -13: GetErrInfo = "ϵͳ������"
                Case -14: GetErrInfo = "����"
                Case -15: GetErrInfo = "����δ����"
                Case -16: GetErrInfo = "����ͨѶ����"
                Case -17: GetErrInfo = "�����ļ�����"
                Case -18: GetErrInfo = "Ӧ�����"
                Case -19: GetErrInfo = "�������еĿ�"
                Case -20: GetErrInfo = "���ѵ���"
                Case -21: GetErrInfo = "���ݿ����ʧ��"
                Case -22: GetErrInfo = "��������ʧ��"
                Case -23: GetErrInfo = "�������"
                Case -24: GetErrInfo = "���ſ�����"
                Case -25: GetErrInfo = "�����������޶�"
                Case -100: GetErrInfo = "�޷�ʶ��Ŀ�"
                Case Else
                    GetErrInfo = "У԰��֧�ֲ��ֳ��ִ���"
                End Select
                GetErrInfo = "�����ţ�" & strCode & vbCr & "����������" & GetErrInfo
               Case Else
    End Select
End Function

Public Function OraDataOpen(cnOracle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional blnMessage As Boolean = True) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strError As String
    
    On Error Resume Next
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
    End With
    If Err <> 0 Then
        If blnMessage = True Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            Else
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
            End If
        End If
        
        Err.Clear
        OraDataOpen = False
        Exit Function
    End If
    OraDataOpen = True
End Function

Public Sub GetUserInfo()
 '���ܣ���ȡ��½�û���Ϣ
    Dim rsUser As New ADODB.Recordset
    Dim strSql As String
    
    Set rsUser = New ADODB.Recordset
    rsUser.CursorLocation = adUseClient
    'rsUser.Open "Select A.ID,A.����ID,A.���,A.����,A.����,B.�û���,C.���� as ���� from ��Ա�� A,�ϻ���Ա�� B,���ű� C Where A.����ID=C.ID And  B.��ԱID=A.ID AND Upper(B.�û���)=Upper(User)", gcnOracle, adOpenKeyset
    
    strSql = "select P.*,D.���� as ���ű���,D.���� as ��������,M.����ID,u.�û��� " & _
                " from �ϻ���Ա�� U,��Ա�� P,���ű� D,������Ա M " & _
                " Where U.��Աid = P.id And P.ID=M.��ԱID and  M.ȱʡ=1 and M.����id = D.id and U.�û���=user"
    rsUser.Open strSql, gcnOracle, adOpenKeyset
    
    If rsUser.RecordCount <> 0 Then
        UserInfo.ID = rsUser!ID
        UserInfo.��� = rsUser!���
        UserInfo.����ID = IIf(IsNull(rsUser!����ID), 0, rsUser!����ID)
        UserInfo.���� = IIf(IsNull(rsUser!����), "", rsUser!����)
        UserInfo.���� = IIf(IsNull(rsUser!����), "", rsUser!����)
        UserInfo.���� = rsUser!��������
        UserInfo.�û��� = rsUser!�û���
        UserInfo.վ�� = rsUser!�û���
        
        'Ϊ�˲������������ظ�������һ������
        gstrUserName = UserInfo.����
    End If
End Sub

Public Function DateStr() As String
    Dim rsTmp As New ADODB.Recordset

    rsTmp.Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    DateStr = Format(rsTmp.Fields(0).Value, "yyyy-MM-dd HH:mm:ss")
End Function

Public Function TrimStr(ByVal str As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Public Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Function NextNo(intBillId As Integer) As Variant
'------------------------------------------------------------------------------------
'���ܣ������ض���������µĺ���,�������£�
'   һ�����ԭ��
'   1   ����ID         ����    ��Զ������� �Զ���ȱ��
'   �������λȷ��ԭ��:
'       ��1990Ϊ���������������������0��9/A��Z��˳����Ϊ��ȱ���
'������
'   intBillId:�ɡ�������Ʊ�ָ���ĵ��ݱ�ʶ
'���أ�
'------------------------------------------------------------------------------------
    Dim rsCtrl As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim vntNo As Variant        '��ȡ�ĺ�����м����
    Dim blnNext As Boolean      '��õĺ����Ƿ�Ϊ�������룬����Ϊ��ȱ����
    Dim intYear, strYear As String      '��ȱ�־λ
    
    Dim blnByDate As Boolean, curDate As Date
RESTART:
    Err = 0
    On Error GoTo errHand
    
    If intBillId = 1 Then
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSql = "select * from ������Ʊ� where ��Ŀ���=" & intBillId
                Call SQLTest(App.ProductName, "mdlInPatient", strSql) 'SQLTest
                .Open strSql, gcnOracle, adOpenKeyset, adLockOptimistic
                Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!������), 0, !������)
            strSql = "select nvl(max(����ID),0)+1 from ������Ϣ where ����ID>=" & vntNo & ""
            
            With rsTmp
                If .State = adStateOpen Then .Close
                Call SQLTest(App.ProductName, "mdlInsure", strSql) 'SQLTest
                .Open strSql, gcnOracle
                Call SQLTest
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            On Error Resume Next
            .Update "������", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            If Err <> 0 Then
                .CancelUpdate
                GoTo RESTART
            End If
            NextNo = vntNo
        End With
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    NextNo = Null
End Function

Public Function Get��Ժ���(lng����ID As Long, lng��ҳID As Long, _
Optional ByVal bln����� As Boolean = True, Optional ByVal bln�������� As Boolean = False) As String
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.������Ϣ as ��Ժ���,B.���� �������� " & _
             " From ������ A,��������Ŀ¼ B " & _
             " Where A.����ID=" & lng����ID & " And A.����ID=B.ID(+) And A.��ҳID=" & lng��ҳID & " And A.�������=2"
    rsInNote.CursorLocation = adUseClient
    Call OpenRecordset(rsInNote, "ҽ���ӿ�", strTmp)
    
    If Not rsInNote.EOF Then
        Get��Ժ��� = IIf(IsNull(rsInNote!��Ժ���), "", rsInNote!��Ժ���)
    End If
    If Not bln����� Then
        Get��Ժ��� = Trim(Get��Ժ���)
        If Get��Ժ��� = "" Then Get��Ժ��� = "��"
    End If
    If bln�������� Then
        If Not rsInNote.EOF Then
            Get��Ժ��� = Get��Ժ��� & "|" & NVL(rsInNote!��������)
        Else
            Get��Ժ��� = Get��Ժ��� & "|"
        End If
    End If
End Function

Public Function BuildPatiInfo(ByVal bytType As Byte, ByVal strInfo As String, ByVal lng����ID As Long) As Long
'���ܣ����������ʻ���Ϣ
'������bytType=0-����,1-סԺ
'      strInfo='0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
'      8����;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(1,2,3);15����֤��;16�����;17�Ҷȼ�
'      18�ʻ������ۼ�;19�ʻ�֧���ۼ�;20����ͳ���ۼ�;21ͳ�ﱨ���ۼ�;22סԺ�����ۼ�;23�������
'      24��������;25�����ۼ�;26����ͳ���޶�
'���أ�����ID
    Const MAX_BOUND = 26 'Ҫ�������Ϣ����
    
    Dim rsPati As ADODB.Recordset, str��λ���� As String, lng���� As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String, curDate As Date
    Dim lng���� As Long, array��Ϣ As Variant
    Dim lngTemp As Long
    
    gcnOracle.BeginTrans
    On Error GoTo ErrHandle
    
    If Len(Trim(strInfo)) <> 0 Then
        curDate = zlDatabase.Currentdate
        
        '200308z012:��֤�������Ϣ������
        If UBound(Split(strInfo, ";")) < MAX_BOUND Then
            strInfo = strInfo & String(MAX_BOUND - UBound(Split(strInfo, ";")), ";")
        End If
        array��Ϣ = Split(strInfo, ";")
        
        '�ӵ�7��������ȡ����λ����
        If array��Ϣ(7) Like "*(*" Then
            str��λ���� = Split(array��Ϣ(7), "(")(UBound(Split(array��Ϣ(7), "(")))
            str��λ���� = Mid(str��λ����, 1, Len(str��λ����) - 1)
        End If
        'ȡ����
        If IsDate(array��Ϣ(5)) Then
            lng���� = Int(curDate - CDate(array��Ϣ(5))) / 365
        End If
        
        lng���� = Val(array��Ϣ(8))
        
        If lng����ID > 0 Then
            '�ò����Ѿ�����
            gstrSQL = "Select nvl(����ID,0) ����ID from �����ʻ� where ҽ����='" & CStr(array��Ϣ(1)) & "' and ����=" & lng���� & " and ����=" & gintInsure
            Call OpenRecordset(rsTemp, "�����ʻ�")
            If rsTemp.EOF = False Then
                If rsTemp("����ID") <> lng����ID Then
                    If gintInsure = TYPE_�ɶ��� Then
                        If MsgBox("�Ѿ�������ͬҽ���ŵ�����һλ���ˣ�����Ҫ������λ���˺ϲ���", vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then
                            gcnOracle.RollbackTrans
                            Exit Function
                        End If
                        '�����������˽��кϲ�
                        lngTemp = MergePatient(lng����ID, rsTemp!����ID)
                        If lngTemp = 0 Then
                            gcnOracle.RollbackTrans
                            Exit Function
                        End If
                        lng����ID = lngTemp
                    Else
                        MsgBox "�Ѿ�������ͬҽ���ŵ�����һλ���ˣ������ڲ��˹����н�����λ���˺ϲ�", vbInformation, gstrSysName
                        gcnOracle.RollbackTrans
                        Exit Function
                    End If
                End If
            End If
        End If
        
        '�ʻ�Ψһ������,����,ҽ����
        strSql = "Select A.*,B.ҽ���� From ������Ϣ A," & _
            " (Select * From �����ʻ�" & _
            " Where ����=" & gintInsure & _
            " And ҽ����='" & CStr(array��Ϣ(1)) & "'" & _
            " And ����=" & lng���� & ") B" & _
            " Where " & IIf(lng����ID = 0, "A.����ID=B.����ID", "A.����ID=B.����ID(+) and A.����ID=" & lng����ID) '���ܲ���ID�Ѿ�ȷ��
        Set rsPati = New ADODB.Recordset
        rsPati.CursorLocation = adUseClient
        Call OpenRecordset(rsPati, "ҽ���ӿ�", strSql)
        
        If rsPati.EOF Then
            '�ޱ����ʻ�����Ϊû�в�����Ϣ
            If lng����ID = 0 Then lng����ID = NextNo(1)
            strSql = "zl_������Ϣ_Insert(" & lng����ID & ",NULL,NULL,'������ҽ�Ʊ���'," & _
                "'" & array��Ϣ(3) & "','" & array��Ϣ(4) & "'," & IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
                "To_Date('" & Format(array��Ϣ(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "NULL,'" & array��Ϣ(6) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
                "NULL,NULL,NULL,NULL,NULL,NULL,'" & array��Ϣ(7) & "',NULL,NULL,NULL," & _
                "NULL,NULL,NULL," & gintInsure & "," & _
                "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call SQLTest(App.ProductName, "ҽ���ӿ�", strSql)
            gcnOracle.Execute strSql, , adCmdStoredProc
            Call SQLTest
        Else
            '�в�����Ϣ�ͱ����ʻ���Ϣ
            If rsPati("����") <> array��Ϣ(3) Then
                If MsgBox("����ԭ�еǼǵ������� " & rsPati("����") & " ����ˢ���õ������� " & array��Ϣ(3) & " ������" & vbCrLf & _
                          "��������²���ԭ�еĵǼ���Ϣ���Ƿ�ȷ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
            End If
            If lng����ID = 0 Then lng����ID = rsPati!����ID
            strSql = "zl_������Ϣ_Update(" & _
                lng����ID & "," & IIf(IsNull(rsPati!�����), "NULL", rsPati!�����) & "," & _
                IIf(IsNull(rsPati!סԺ��), "NULL", rsPati!סԺ��) & ",'" & IIf(IsNull(rsPati!�ѱ�), "", rsPati!�ѱ�) & "'," & _
                "'" & IIf(IsNull(rsPati!ҽ�Ƹ��ʽ), "", rsPati!ҽ�Ƹ��ʽ) & "'," & _
                "'" & array��Ϣ(3) & "','" & array��Ϣ(4) & "'," & IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
                "To_Date('" & Format(array��Ϣ(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "'" & IIf(IsNull(rsPati!�����ص�), "", rsPati!�����ص�) & "','" & array��Ϣ(6) & "'," & _
                "'" & IIf(IsNull(rsPati!���), "", rsPati!���) & "','" & IIf(IsNull(rsPati!ְҵ), "", rsPati!ְҵ) & "'," & _
                "'" & IIf(IsNull(rsPati!����), "", rsPati!����) & "','" & IIf(IsNull(rsPati!����), "", rsPati!����) & "'," & _
                "'" & IIf(IsNull(rsPati!ѧ��), "", rsPati!ѧ��) & "','" & IIf(IsNull(rsPati!����״��), "", rsPati!����״��) & "'," & _
                "'" & IIf(IsNull(rsPati!��ͥ��ַ), "", rsPati!��ͥ��ַ) & "','" & IIf(IsNull(rsPati!��ͥ�绰), "", rsPati!��ͥ�绰) & "'," & _
                "'" & IIf(IsNull(rsPati!�����ʱ�), "", rsPati!�����ʱ�) & "','" & IIf(IsNull(rsPati!��ϵ������), "", rsPati!��ϵ������) & "'," & _
                "'" & IIf(IsNull(rsPati!��ϵ�˹�ϵ), "", rsPati!��ϵ�˹�ϵ) & "','" & IIf(IsNull(rsPati!��ϵ�˵�ַ), "", rsPati!��ϵ�˵�ַ) & "'," & _
                "'" & IIf(IsNull(rsPati!��ϵ�˵绰), "", rsPati!��ϵ�˵绰) & "'," & IIf(IsNull(rsPati!��ͬ��λID), "NULL", rsPati!��ͬ��λID) & "," & _
                "'" & array��Ϣ(7) & "','" & IIf(IsNull(rsPati!��λ�绰), "", rsPati!��λ�绰) & "'," & _
                "'" & IIf(IsNull(rsPati!��λ�ʱ�), "", rsPati!��λ�ʱ�) & "','" & IIf(IsNull(rsPati!��λ������), "", rsPati!��λ������) & "'," & _
                "'" & IIf(IsNull(rsPati!��λ�ʺ�), "", rsPati!��λ�ʺ�) & "','" & IIf(IsNull(rsPati!������), "", rsPati!������) & "'," & _
                "" & IIf(IsNull(rsPati!������), "NULL", rsPati!������) & "," & gintInsure & ")"
            Call SQLTest(App.ProductName, "ҽ���ӿ�", strSql)
            gcnOracle.Execute strSql, , adCmdStoredProc
            Call SQLTest
        End If
        
        '�������±����ʻ���Ϣ(�Զ�)
        strSql = "zl_�����ʻ�_insert(" & lng����ID & "," & gintInsure & "," & _
            lng���� & "," & _
            "'" & IIf(array��Ϣ(0) = "-1", array��Ϣ(1), array��Ϣ(0)) & "'," & _
            "'" & array��Ϣ(1) & "'," & _
            "'" & array��Ϣ(2) & "'," & _
            "'" & array��Ϣ(9) & "'," & _
            "'" & array��Ϣ(15) & "'," & _
            "'" & array��Ϣ(10) & "'," & _
            "'" & str��λ���� & "'," & _
            Val(array��Ϣ(11)) & "," & _
            Val(array��Ϣ(12)) & "," & _
            IIf(Val(array��Ϣ(13)) = 0, "NULL", Val(array��Ϣ(13))) & "," & _
            IIf(Val(array��Ϣ(14)) = 0, 1, Val(array��Ϣ(14))) & "," & _
            IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
            "'" & array��Ϣ(17) & "'," & _
            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call SQLTest(App.ProductName, "ҽ���ӿ�", strSql)
        gcnOracle.Execute strSql, , adCmdStoredProc
        Call SQLTest
        
        '���������ʻ������Ϣ(�Զ�)
        '200308z012:�ɶ�:����"24��������=zyjs,25�����ۼ�=tcbxbl,26����ͳ���޶�=zyxe"
        strSql = "zl_�ʻ������Ϣ_Insert(" & lng����ID & "," & gintInsure & "," & Year(curDate) & "," & _
            Val(array��Ϣ(18)) & "," & Val(array��Ϣ(19)) & "," & _
            Val(array��Ϣ(20)) & "," & Val(array��Ϣ(21)) & "," & _
            Val(array��Ϣ(22)) & "," & Val(array��Ϣ(24)) & "," & Val(array��Ϣ(25)) & "," & Val(array��Ϣ(26)) & ")"
        Call SQLTest(App.ProductName, "ҽ���ӿ�", strSql)
        gcnOracle.Execute strSql, , adCmdStoredProc
        Call SQLTest
    End If
    gcnOracle.CommitTrans
    BuildPatiInfo = lng����ID
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
End Function

Public Function GetTextFromCombo(cmbTemp As ComboBox, ByVal blnAfter As Boolean, Optional strSplit As String = ".") As String
'������cmbTemp  ׼����ȡ���ݵ�ComboBox�ؼ�
'      blnAfter ��ʾ��.֮ǰ��֮��ȡֵ
    Dim lngPos As Long
    
    lngPos = InStr(cmbTemp.Text, strSplit)
    If lngPos = 0 Then
        'ֱ�ӷ��������ַ���
        GetTextFromCombo = "'" & cmbTemp.Text & "'"
    Else
        If blnAfter = False Then
            'Բ��֮ǰ
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, 1, lngPos - 1) & "'"
        Else
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, lngPos + 1) & "'"
        End If
    End If
End Function

Public Sub SetComboByText(cmbTemp As ComboBox, ByVal strText As String, ByVal blnAfter As Boolean, Optional strSplit As String = ".")
'������cmbTemp  ׼�����õ�ComboBox�ؼ�
'      blnAfter ��ʾ��.֮ǰ��֮��ȡֵ
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cmbTemp.ListCount - 1
        strTemp = cmbTemp.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            'ֱ�ӷ��������ַ���
            If strText = strTemp Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                'Բ��֮ǰ
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '�Ѿ��ҵ�
        cmbTemp.ListIndex = lngCount
    Else
        cmbTemp.ListIndex = -1
        If blnAfter = True Then
            cmbTemp.AddItem strText
        End If
    End If
End Sub

Public Function MidUni(ByVal strTemp As String, ByVal Start As Long, ByVal Length As Long) As String
'���ܣ������ݿ����õ��ַ������Ӽ���Ҳ���Ǻ��ְ������ַ��㣬����ĸ����һ��
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    'ȥ�����ܳ��ֵİ���ַ�
    MidUni = Replace(MidUni, Chr(0), "")
End Function

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
'���ܣ����ı���Varchar2�ĳ��ȼ��㷽�����нض�
    Dim strText As String
    
    strText = IIf(IsNull(varText), "", varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    'ȥ�����ܳ��ֵİ���ַ�
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function

Public Function GetComputer(frmParant As Form, Optional ByVal strCaption As String = "ѡ������") As String
'���ܣ����ؼ������
   Dim BI As BrowseInfo
   Dim pidl As Long
   Dim sPath As String
   Dim pos As Integer
   
  'obtain the pidl to the special folder 'network'
   If SHGetSpecialFolderLocation(frmParant.hwnd, CSIDL_NETWORK, pidl) = 0 Then
     'fill in the required members, limiting the
     'Browse to the network by specifying the
     'returned pidl as pidlRoot
      With BI
         .hwndOwner = frmParant.hwnd
         .pIDLRoot = pidl
         .pszDisplayName = Space$(MAX_PATH)
         .lpszTitle = lstrcat(strCaption, "")
         .ulFlags = BIF_BROWSEFORCOMPUTER
      End With
         
     'show the browse dialog. We don't need
     'a pidl, so it can be used in the If..then directly.
      If SHBrowseForFolder(BI) <> 0 Then
               
         'a server was selected. Although a valid pidl
         'is returned, SHGetPathFromIDList only return
         'paths to valid file system objects, of which
         'a networked machine is not. However, the
         'BROWSEINFO displayname member does contain
         'the selected item, which we return
          GetComputer = TrimStr(BI.pszDisplayName)
            
      End If  'If SHBrowseForFolder
      
      Call CoTaskMemFree(pidl)
               
   End If  'If SHGetSpecialFolderLocation
   
End Function

Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSql As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSql = "", gstrSQL, strSql))
    rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Sub ExecuteProcedure(ByVal strCaption As String)
'���ܣ�ִ��SQL���
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Public Sub CenterTableCaption(mshTemp As Object)
'���ܣ����ñ�����ͷ���ж���
    With mshTemp
        .Col = 0
        .Row = .FixedRows - 1
        .ColSel = .Cols - 1
        .RowSel = .Row
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = .FixedRows: .Col = .FixedCols
    End With
End Sub

Public Function GetסԺ����(lng����ID As Long) As Integer
'���ܣ���ȡָ�����˱����סԺ����
'˵��������סԺ��������궼����һ��סԺ��
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select Count(*) as ���� From ������ҳ Where Nvl(��Ժ����,Sysdate)=To_Date(To_Char(Sysdate,'YYYY')||'-01-01','YYYY-MM-DD') And ����ID=" & lng����ID
    rsTmp.CursorLocation = adUseClient
    Call OpenRecordset(rsTmp, "ҽ���ӿ�")
    
    If Not rsTmp.EOF Then GetסԺ���� = IIf(IsNull(rsTmp!����), 0, rsTmp!����)
End Function

Public Function Get�ʻ���Ϣ(ByVal lng����ID As Long, ByVal str��� As String, intסԺ�����ۼ� As Integer, _
    cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, _
    curͳ�ﱨ���ۼ� As Currency, Optional cur�������� As Currency, Optional cur�����ۼ� As Currency, _
    Optional cur����ͳ���޶� As Currency) As Boolean
'���ܣ��õ��ʻ������Ϣ
'200308z012:�����������ز���
    Dim rsTemp As New ADODB.Recordset
    
    cur�ʻ������ۼ� = 0
    cur�ʻ�֧���ۼ� = 0
    cur����ͳ���ۼ� = 0
    curͳ�ﱨ���ۼ� = 0
    intסԺ�����ۼ� = 0
    cur�������� = 0
    cur�����ۼ� = 0
    cur����ͳ���޶� = 0
    
    '�ʻ������Ϣ
    gstrSQL = "Select * From �ʻ������Ϣ Where ����ID=" & lng����ID & " And ����=" & gintInsure & " And ���=" & str���
    Call OpenRecordset(rsTemp, "ҽ���ӿ�")
    
    If rsTemp.EOF = False Then
        cur�ʻ������ۼ� = IIf(IsNull(rsTemp("�ʻ������ۼ�")), 0, rsTemp("�ʻ������ۼ�"))
        cur�ʻ�֧���ۼ� = IIf(IsNull(rsTemp("�ʻ�֧���ۼ�")), 0, rsTemp("�ʻ�֧���ۼ�"))
        cur����ͳ���ۼ� = IIf(IsNull(rsTemp("����ͳ���ۼ�")), 0, rsTemp("����ͳ���ۼ�"))
        curͳ�ﱨ���ۼ� = IIf(IsNull(rsTemp("ͳ�ﱨ���ۼ�")), 0, rsTemp("ͳ�ﱨ���ۼ�"))
        intסԺ�����ۼ� = IIf(IsNull(rsTemp("סԺ�����ۼ�")), 0, rsTemp("סԺ�����ۼ�"))
        cur�������� = IIf(IsNull(rsTemp("��������")), 0, rsTemp("��������"))
        cur�����ۼ� = IIf(IsNull(rsTemp("�����ۼ�")), 0, rsTemp("�����ۼ�"))
        cur����ͳ���޶� = IIf(IsNull(rsTemp("����ͳ���޶�")), 0, rsTemp("����ͳ���޶�"))
    End If

End Function

Public Function �����������(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim clsҽ�� As New clsInsure
    Dim dblȫ�Է� As Currency, dbl�����Ը� As Currency, dbl����ͳ�� As Currency
    Dim dbl�����ʻ� As Double
    Dim lng����ID As Long
    Dim rs��׼��Ŀ As New ADODB.Recordset
    
    If rs��ϸ.RecordCount > 0 Then
        rs��ϸ.MoveFirst
        lng����ID = rs��ϸ("����ID")
    End If
    
    gstrSQL = "select A.�շ�ϸĿID from ������׼��Ŀ A,�����ʻ� B " & _
            "where A.����ID=B.����ID and B.����ID=" & lng����ID & " and ����=" & gintInsure
    Call OpenRecordset(rs��׼��Ŀ, "�������")
    
    Do Until rs��ϸ.EOF
        rs��׼��Ŀ.Filter = "�շ�ϸĿID = " & rs��ϸ("�շ�ϸĿID")
        
        If rs��ϸ("�Ƿ�ҽ��") = 1 Or rs��׼��Ŀ.EOF = False Then
            '�������׼��Ŀ��ǿ�н���ͳ��
            dbl����ͳ�� = dbl����ͳ�� + rs��ϸ("ͳ����")
            dbl�����Ը� = dbl�����Ը� + rs��ϸ("ʵ�ս��") - rs��ϸ("ͳ����")
        Else
            dblȫ�Է� = dblȫ�Է� + rs��ϸ("ʵ�ս��")
        End If
            
        rs��ϸ.MoveNext
    Loop
    If clsҽ��.GetCapability(support�շ��ʻ�ȫ�Է�) = True Then
        dbl�����ʻ� = dbl�����ʻ� + dblȫ�Է�
    End If
    
    If Isȫ��ͳ��(lng����ID) = True Then
        '�����Ը�Ҳ����ҽ������֧��
        str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";0|ҽ������;" & dbl����ͳ�� + dbl�����Ը� & ";0"
    Else
        If clsҽ��.GetCapability(support�շ��ʻ������Ը�) = True Then
            dbl�����ʻ� = dbl�����ʻ� + dbl�����Ը�
        End If
        
        str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";0|ҽ������;" & dbl����ͳ�� & ";0"
    End If
    
    ����������� = True
End Function

Public Function Isȫ��ͳ��(ByVal ����ID As Long) As Boolean
'���ܣ��ж��Ƿ�ȫ��ͳ�ﲡ��(ע�⣺���Ĳ���ID���ܷ�ҽ�����˵�)
    Dim rsTemp As New ADODB.Recordset
    
    If gintInsure = TYPE_�Թ��� Then
        '�����Թ�ҽ����ֻҪ������������Ա���Ǿ���ȫ��ͳ��
        gstrSQL = "select ��ְ from �����ʻ� where ����ID=" & ����ID & " and ����=" & TYPE_�Թ���
        Call OpenRecordset(rsTemp, "ҽ���ӿ�")
        If rsTemp.EOF = False Then
            Isȫ��ͳ�� = IIf(rsTemp("��ְ") = 3, True, False)
        End If
    Else
        gstrSQL = _
            "Select Nvl(B.ȫ��ͳ��,0) as ȫ��ͳ��" & _
            " From �����ʻ� A,��������� B" & _
            " Where A.���� = B.���� And Nvl(A.����, 0) = Nvl(B.����, 0)" & _
            " And Nvl(A.��ְ,0)=Nvl(B.��ְ,0)" & _
            " And B.����<=Nvl(A.�����,0) And (A.�����<=B.���� Or B.����=0)" & _
            " And A.����ID=" & ����ID & " And A.����=" & gintInsure
        Set rsTemp = New ADODB.Recordset
        rsTemp.CursorLocation = adUseClient
        Call OpenRecordset(rsTemp, "ҽ���ӿ�")
        
        If Not rsTemp.EOF Then Isȫ��ͳ�� = (rsTemp!ȫ��ͳ�� = 1)
    End If
End Function

Public Function AddDate(ByVal strOrin As String, Optional ByVal blnʱ As Boolean = False) As String
'���ܣ�Ϊ��ȫ��������Ϣ��������
    Dim strTemp As String
    Dim intPos As Integer
    
    strTemp = Trim(strOrin)
    
    If strTemp = "" Then
        AddDate = ""
        Exit Function
    End If
    
    intPos = InStr(strTemp, "-")
    If intPos = 0 Then
        intPos = InStr(strTemp, ".")
        If intPos <> 0 Then
            'ʹ�� . ��
            strTemp = Replace(strTemp, ".", "-")
        End If
    End If
    
    If intPos = 0 Then
        'û��"-",�ֹ�����
        intPos = Len(strTemp)
        If intPos <= 8 Then
            If intPos = 8 Then
                strTemp = Mid(strTemp, 1, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7, 2)
            ElseIf intPos > 4 Then
                strTemp = Left(strTemp, intPos - 4) & "-" & Mid(Right(strTemp, 4), 1, 2) & "-" & Right(strTemp, 2)
            ElseIf intPos > 2 Then
                strTemp = Format(Date, "yyyy") & "-" & Left(strTemp, intPos - 2) & "-" & Right(strTemp, 2)
            Else
                strTemp = Format(Date, "yyyy") & "-" & Format(Date, "MM") & "-" & strTemp
            End If
        End If
    Else
        If blnʱ = False Then
            If IsDate(strTemp) Then
                strTemp = Format(CDate(strTemp), "yyyy-MM-dd")
            End If
        Else
            '����Сʱ
            If InStr(strTemp, " ") > 0 Then
                '������Сʱ
                If IsDate(strTemp & ":00") Then
                    strTemp = Format(CDate(strTemp & ":00"), "yyyy-MM-dd HH:ss")
                End If
            Else
                If IsDate(strTemp) Then
                    strTemp = Format(CDate(strTemp), "yyyy-MM-dd HH:ss")
                End If
            End If
        End If
    End If
    
    AddDate = strTemp
End Function

Public Function Insert�����������(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str���㷽ʽ As String) As Boolean
'���ܣ��������������ݱ�������
'���������㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    Dim cnTemp As New ADODB.Connection
    Dim strDate As String
    Dim lngCount As Long, arr���㷽ʽ As Variant, arr��� As Variant
    
    cnTemp.Open gcnOracle.ConnectionString 'Ϊ�˷�ֹһ�����Ӵ���ν�������
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    cnTemp.BeginTrans
    On Error GoTo ErrHandle
    
    gstrSQL = "zl_����ģ�����_Clear(" & lng����ID & "," & lng��ҳID & ")"
    cnTemp.Execute gstrSQL, , adCmdStoredProc
    
    arr���㷽ʽ = Split(str���㷽ʽ, "|")
    For lngCount = 0 To UBound(arr���㷽ʽ)
        If arr���㷽ʽ(lngCount) <> "" Then
            arr��� = Split(arr���㷽ʽ(lngCount), ";")
            If UBound(arr���) > 1 Then
                If Val(arr���(1)) <> 0 Then
                    gstrSQL = "zl_����ģ�����_Insert(" & lng����ID & "," & IIf(lng��ҳID = 0, "null", lng��ҳID) & _
                        ",'" & arr���(0) & "'," & Val(arr���(1)) & "," & strDate & ")"
                    cnTemp.Execute gstrSQL, , adCmdStoredProc
                End If
            End If
        End If
    Next
    
    cnTemp.CommitTrans
    Insert����������� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    cnTemp.RollbackTrans
End Function

Public Function Clear�����������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��ڽ���֮�󣬽����������������
    
    gstrSQL = "zl_����ģ�����_Clear(" & lng����ID & "," & lng��ҳID & ")"
    Call ExecuteProcedure("�������")
    
    Clear����������� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get��������(ByVal str���֤ As String, ByVal lng���� As Long) As String
'���ܣ��������֤���������õ���������
    Dim strDate As String
    If Len(str���֤) = 15 Then
        '��ʽ�����֤��
        strDate = AddDate(Mid(str���֤, 7, 6))
        strDate = "19" & strDate
    ElseIf Len(str���֤) = 18 Then
        '��ʽ�����֤��
        strDate = AddDate(Mid(str���֤, 7, 8))
    Else
        'û�����֤��
        strDate = Format(DateAdd("yyyy", lng���� * -1, Date), "yyyy-MM-dd")
    End If
    
    If IsDate(strDate) = True Then
        Get�������� = Format(CDate(strDate), "yyyy-MM-dd")
    End If
End Function

Public Function GetOracleFormat(ByVal dat���� As Date)
    GetOracleFormat = "To_Date('" & Format(dat����, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Public Function NVL(ByVal varValue As Variant, Optional varDefalut As Variant = "") As Variant
'���ܣ�ģ��Oracle�ĺ���
    NVL = IIf(IsNull(varValue) = True, varDefalut, varValue)
End Function

Public Sub RemoveSelect(lvw As ListView)
'���ܣ�ɾ����ǰѡ����
    Dim lngIndex  As Long
    
    With lvw
        If .SelectedItem Is Nothing Then Exit Sub
        
        lngIndex = .SelectedItem.Index
        .ListItems.Remove lngIndex
        
        If .ListItems.Count > 0 Then
            '��������б��������һ��ѡ��
            lngIndex = IIf(.ListItems.Count > lngIndex, lngIndex, .ListItems.Count)
            .ListItems(lngIndex).Selected = True
            .ListItems(lngIndex).EnsureVisible
        End If
    End With

End Sub

Public Function CanסԺ�������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��жϲ��˵�סԺ���������Ƿ��������ϡ��жϱ�׼�Ǽ�鲡�����µ�סԺ��¼������У��Ͳ��ܽ�����
'������lng����ID     ����ID
'      lng��ҳID     �ý��ʼ�¼���ڵ�סԺ����
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle

    gstrSQL = "SELECT COUNT(*) as סԺ���� FROM ������ҳ WHERE ����ID=" & lng����ID & " AND ��ҳID>" & lng��ҳID
    Call OpenRecordset(rsTemp, "��������")
    If rsTemp("סԺ����") > 0 Then
        MsgBox "�ò����Ѿ����µ�סԺ��¼������������ǰסԺ�Ľ������ݡ�", vbInformation, gstrSysName
        Exit Function
    End If

    CanסԺ������� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ҽ�������Ѿ���Ժ(ByVal lng����ID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select Nvl(��ǰ״̬,0) ״̬ From �����ʻ� Where ����ID=" & lng����ID
    Call OpenRecordset(rsTmp, "�ж�ҽ�������Ƿ��Ժ")
    
    ҽ�������Ѿ���Ժ = (rsTmp!״̬ = 0)
End Function

Public Function ����δ�����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim rs���� As New ADODB.Recordset
    '���ô�סԺ�Ƿ��з���δ����
    gstrSQL = "Select nvl(�������,0) as ���  from ������� where ����ID=" & lng����ID & " and ����=1"
    Call OpenRecordset(rs����, "�Ƿ����δ�����")
    If rs����.EOF = True Then
        ����δ����� = False
    Else
        ����δ����� = (rs����("���") <> 0)
    End If
End Function

Public Function ��ȡ���Ժ���(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
Optional ByVal bln��Ժ��� As Boolean = True, Optional ByVal bln����� As Boolean = True, _
Optional ByVal bln�������� As Boolean = False) As String
    
    '1-�������;2-��Ժ���;3-��Ժ���
    Dim rs��� As New ADODB.Recordset
    If bln�������� = False Then
        gstrSQL = " Select A.������Ϣ" & _
                  " From ������ A" & _
                  " Where A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & _
                  " And A.�������=" & IIf(bln��Ժ���, "1", "3") & " And ��ϴ���=1"
    Else
        gstrSQL = " Select A.������Ϣ,B.���� ��������" & _
                  " From ������ A,��������Ŀ¼ B" & _
                  " Where A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & _
                  " And A.����ID=B.ID(+) And A.�������=" & IIf(bln��Ժ���, "1", "3")
    End If
    Call OpenRecordset(rs���, "��ȡ���Ժ���")
    
    ��ȡ���Ժ��� = ""
    If Not rs���.EOF Then
        ��ȡ���Ժ��� = IIf(IsNull(rs���!������Ϣ), "", rs���!������Ϣ)
    End If
    
    ��ȡ���Ժ��� = Trim(��ȡ���Ժ���)
    If Not bln����� And ��ȡ���Ժ��� = "" Then
        ��ȡ���Ժ��� = "��"
    End If
    If bln�������� Then
        If Not rs���.EOF Then
            ��ȡ���Ժ��� = ��ȡ���Ժ��� & "|" & NVL(rs���!��������, " ")
        Else
            ��ȡ���Ժ��� = ��ȡ���Ժ��� & "| "
        End If
    End If
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function ��������(ByVal int���� As Integer) As Boolean
    Dim rs���� As New ADODB.Recordset
    
    �������� = False
    gstrSQL = "Select Nvl(��������,0) ���� From ������� Where ���=" & int����
    Call OpenRecordset(rs����, "�Ƿ�������")
    If Not rs����.EOF Then
        �������� = (rs����!���� = 1)
    End If
End Function

Private Function GetPatiInfo(lngID As Long) As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select * From ������Ϣ A,������ҳ B Where A.����ID=B.����ID(+) And A.����ID=" & lngID & " Order by ��ҳID"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnOracle, adOpenKeyset
    If Not rsTmp.EOF Then Set GetPatiInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
        
Private Function MergePatient(ByVal lngOld As Long, ByVal lngInsure As Long) As Long
    Dim i As Integer, j As Integer
    Dim lngNew As Long
    Dim curDate As Date
    Dim strSql As String
    Dim rsPatiS As New ADODB.Recordset
    Dim rsPatiO As New ADODB.Recordset
    Set rsPatiS = GetPatiInfo(lngOld)
    Set rsPatiO = GetPatiInfo(lngInsure)
        
    'AB��ס��Ժ
    If Not IsNull(rsPatiS!��ҳID) And Not IsNull(rsPatiO!��ҳID) Then
        '1.��סԺ����Ժ,������(�Ⱥ�סԺ����Ϊ����Ժ-��Ժ,��Ժ-��Ժ����������Ժ-��Ժ,��Ժ-��Ժ)
        '��Ϊ�����˺ϲ���,���򲻶��⴦���Զ���Ժ������Ժ
        rsPatiS.MoveLast
        rsPatiO.MoveLast
        If rsPatiS!��Ժʱ�� <= rsPatiO!��Ժʱ�� Then
            If IsNull(rsPatiS!��Ժʱ��) Then
                MsgBox "����:" & rsPatiS!סԺ�� & " ���һ��סԺ����Ժ,����ǰδ��Ժ,����ִ�кϲ�������", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If IsNull(rsPatiO!��Ժʱ��) Then
                MsgBox "����:" & rsPatiO!סԺ�� & " ���һ��סԺ����Ժ,����ǰδ��Ժ,����ִ�кϲ�������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '2.ʱ�佻����ʾ�Ƿ����
        curDate = zlDatabase.Currentdate
        rsPatiS.MoveFirst
        For i = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For j = 1 To rsPatiO.RecordCount
                If Not (rsPatiO!��Ժ���� >= IIf(IsNull(rsPatiS!��Ժ����), curDate, rsPatiS!��Ժ����) Or _
                    IIf(IsNull(rsPatiO!��Ժ����), curDate, rsPatiO!��Ժ����) <= rsPatiS!��Ժ����) Then
                    If MsgBox("���ֲ���:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]�� " & rsPatiS!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiS!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiS!��Ժ����), curDate, rsPatiS!��Ժ����), "yyyy-MM-dd") & vbCrLf & _
                        "�벡��:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]�ĵ� " & rsPatiO!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiO!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiO!��Ժ����), curDate, rsPatiO!��Ժ����), "yyyy-MM-dd") & _
                        vbCrLf & "���ཻ�棬Ӧ�ò���ͬһ�����ˣ�ȷʵҪ�ϲ���", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
                rsPatiO.MoveNext
            Next
            rsPatiS.MoveNext
        Next
        
        lngNew = NextNo(1)
    End If
    
    strSql = "zl_������Ϣ_MERGE(" & lngOld & "," & lngInsure & IIf(lngNew <> 0, "," & lngNew, "") & ")"
    Screen.MousePointer = 11
    DoEvents
    
    gcnOracle.Execute strSql, , adCmdStoredProc
    Screen.MousePointer = 0
    
    If lngNew <> 0 Then
        If glngSys Like "8??" Then
            MsgBox "�ͻ��ϲ��ɹ�,�ϲ���Ŀͻ�IDΪ""" & lngNew & """��", vbInformation, gstrSysName
        Else
            MsgBox "���˺ϲ��ɹ�,�ϲ���Ĳ���IDΪ""" & lngNew & """��", vbInformation, gstrSysName
        End If
        MergePatient = lngNew
    Else
        If glngSys Like "8??" Then
            MsgBox "�ͻ��ϲ��ɹ���", vbInformation, gstrSysName
        Else
            MsgBox "���˺ϲ��ɹ���", vbInformation, gstrSysName
        End If
        MergePatient = lngInsure
    End If
End Function

Public Sub DebugTool(ByVal strInfo As String)
    Dim intDebug As Integer
    '�ж��Ƿ��ǵ���״̬��������ʾ��ʾ��
    intDebug = GetSetting("ZLSOFT", "ҽ��", "����", 0)
    If intDebug = 0 Then Exit Sub
    MsgBox strInfo
End Sub

Public Function SystemImes() As Variant
'���ܣ���ϵͳ�������뷨���Ʒ��ص�һ���ַ���������
'���أ�����������������뷨,�򷵻ؿմ�
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), strName, Len(strName))
            arrName(j) = Mid(strName, 1, InStr(strName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function OpenIme(Optional strIme As String) As Boolean
'����:�����ƴ��������뷨,��ָ������ʱ�ر��������뷨��֧�ֲ������ơ�
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    
    If strIme = "���Զ�����" Then OpenIme = True: Exit Function
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function
