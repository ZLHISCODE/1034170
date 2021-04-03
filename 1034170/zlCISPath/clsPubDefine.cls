VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPubDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'·���汾
Public Type TYPE_PATH_VERSION
    �汾�� As Integer
    ��׼סԺ�� As String
    ��׼���� As String
    �汾˵�� As String
    ������ As String
    ����ʱ�� As Date
    ����� As String
    ���ʱ�� As Date
    ͣ���� As String
    ͣ��ʱ�� As Date
    ҩ��������� As String
    ҩ�������ʱ�� As Date
End Type

'��֧·��
Public Type TYPE_PATH_BRANCH
    ��֧ID As Long
    ��֧���� As String
    �汾�� As Integer
    ǰһ�׶�ID As Long
    ǰһ�׶����� As String
    ��׼סԺ�� As String
    ��׼���� As String
    ˵�� As String
    ������ As String
    ����ʱ�� As Date
End Type

'����ָ��
Public Type TYPE_PATH_EvalMark
    ID As Long
    ��� As Integer
    ����ָ�� As String
    ָ������ As Byte
    ָ���� As String
End Type

'��������
Public Type TYPE_PATH_EvalCond
    ָ��ID As Long
    ��ĿID As Long
    ��ϵʽ As String
    ����ֵ As String
    ������� As Byte
End Type

'·������
Public Type TYPE_PATH_EVAL
    ָ�꼯 As Collection
    ������ As Collection
End Type

'ʱ��׶�
Public Type TYPE_PATH_STEP
    ID As Long
    ��ID As Long
    ��� As Integer
    ���� As String
    ��ʼ���� As Integer
    �������� As Integer
    ��־ As String
    ���� As String
    ˵�� As String
    ���� As TYPE_PATH_EVAL
    Edit As Integer '0-ԭʼ,1-����,2-�޸�
    ���ڷ�֧ As Boolean
End Type

'·����Ŀ
Public Type TYPE_PATH_ITEM
    ID As Long
    ��Ŀ��� As Integer
    ��Ŀ���� As String
    ִ�з�ʽ As Integer
    ִ���� As Integer
    ������ As Integer
    ��Ŀ��� As String
    ͼ��ID As Long
    ����IDs As String '(�ϰ�)ID1,ID2,...������
    �°没��IDs As String  ' ID1,ID2,....�ַ��ͺ�
    �������� As String  '�ļ�ID1,ԭ��ID1,����1,���1;�ļ�ID2,ԭ��ID2,����2,���2;...
    ҽ��IDs As String 'ID1,ID2,...
    ԭҽ��IDs As String 'ID1,ID2,...  ����˰汾��·���Ŵ��˸�ֵ
    �����ҽ��IDs As String 'ID1,ID2,...  ����˰汾��·���Ŵ��˸�ֵ
    Edit As Integer '0-ԭʼ,1-����,2-�޸�
    Tip As String '������ʾ��ʱʹ��
    ����Ҫ�� As Integer '0-ȫ�����ɣ�1-ѡ������
    ����ԭ�� As String
    ����ο� As String
    ������ As Integer
    ǰһ�汾��ĿID As Long '��ʾ�����ʱ��ֵ
End Type


'����·��
Public Type TYPE_PATH_Pati
    ·��ID As Long
    �汾�� As Long
    ����·��ID As Long
    ����·��״̬ As Long    '-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
    ��ǰ�׶�ID As Long
    �׶θ�ID As Long
    ��ǰ���� As Long
    ��ǰ���� As String
    δ����ԭ�� As String
    ԭ·��ID As Long        '·����תʱ��֮ǰ��·��ID
    ��ǰ�׶η�֧ID As Long
    ����·������ As Long
    �ϲ�·������ As Long
    ����ʱ��  As Date
End Type

'������Ϣ
Public Type TYPE_Pati
    ����ID As Long
    ��ҳID As Long
    ����ID As Long
    ����ID As Long
    ����״̬ As Integer  '������ҳ�е�״̬ 0-����סԺ��1-��δ��ƣ�2-����ת�ƣ�3-��Ԥ��Ժ
End Type