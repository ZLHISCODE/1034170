VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub Show(ByVal lngWindowNo As Long)
'���ܣ�ʵ���Ŷ�������ʾ
'lngWindowNo:���ڱ�ţ����ݴ��ڱ�Ŷ�ȡ������Ϣ����������ʾ

End Sub


Public Function ShowCfg(ByVal lngWindowNo As Long, objOwner As Object) As Boolean
'���ܣ��򿪶�Ӧ����ʽ���ô���
'lngWindowNo:���ڱ��
'objOwner:����ӵ����

End Function



Public Function MsgProcess(ByVal lngWindowNo As Long, _
    ByVal strMsgKey As String, ByVal strXmlContext As String, rsData As ADODB.Recordset) As Boolean
'��Ϣ���մ���
    
    
End Function


Public Function WindNo() As Long
'��ȡ��ǰ��ʽ���ڱ��

End Function


Public Sub RefreshQueueData(Optional ByVal lngQueueId As Long)
'ˢ�½����ŶӺ�����Ϣ

End Sub