VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTypeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����ģ��ص��¼�����
Public Enum TVideoEventType
    vetDelAllImg = 0        'ɾ������ͼ��
    vetGetImg = 1           '��ȡͼ��

    vetLockStudy = 2        '�������
    vetUnLockStudy = 3      '�������

    vetCaptureFirstImg = 4  '�ɼ���һ��ͼ��
    vetUpdateImg = 5        '����ͼ��
End Enum