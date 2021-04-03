VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_EHK40_Bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:EHK-40
'����:������������ӿƼ���չ���޹�˾
'����:�¶�
'�û�:�����һ����ҽԺ
Dim mStrInput  As String
Dim mlngStep As Long '���ղ���
Dim mlng_ͨ���� As Long '�ܵ�ͨ����
Dim mlng_�������� As Long
Dim mlng_���ʱ�� As Long
Dim mlng_Total As Long '�ܵ�������

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function
Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim strCode As String
    Dim lngS As Long, lngE As Long, i As Integer, str�걾�� As String, strEsr As String
    Dim strRecord As String
    
    If strSource <> "" Then
        strCmd = strSource
        mStrInput = mStrInput & strSource
    Else
        Exit Sub
    End If
    
    lngS = InStr(mStrInput, ",69,00")
    If lngS > 0 And mlngStep < 2 Then
        mlngStep = 2 '��ʼ
        
        mStrInput = Mid(mStrInput, lngS)
    Else
        
        If mlngStep = 2 Then
            mlng_ͨ���� = Val("&H" & Mid(Mid(mStrInput, 7, 3), 2))
            mlngStep = 3
        ElseIf mlngStep = 3 Then
            mlng_�������� = Val("&H" & Mid(Mid(mStrInput, 10, 3), 2))
            mlngStep = 4
            
            mlng_Total = (6 + mlng_��������) * mlng_ͨ����
        ElseIf mlngStep = 4 Then
            mlng_���ʱ�� = Val("&H" & Mid(Mid(mStrInput, 13, 3), 2))
            mlngStep = 5
        ElseIf mlngStep = 5 Then
            '�ռ�¼����
            If Len(Mid(mStrInput, 16)) / 3 >= mlng_Total Then
                '������
                mStrInput = Mid(mStrInput, 16)
                For i = 1 To mlng_ͨ����
                    strRecord = Mid(mStrInput, 1, (6 + mlng_��������) * 3)
                    str�걾�� = Val("&H" & Mid(strRecord, 2, 2))
                    strEsr = "ESR|" & Val("&H" & Mid(strRecord, 14, 2))
                    
                    If Val(str�걾��) <> 0 Then
                        strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "|" & Val(str�걾��) & "| |ѪҺ|0|" & strEsr
                    End If
                    mStrInput = Mid(mStrInput, (6 + mlng_��������) * 3 + 1)
                Next
                mlngStep = 0
                mlng_Total = 0
                mlng_ͨ���� = 0
                mlng_�������� = 0
                mlng_���ʱ�� = 0
                
                str�걾�� = ""
                strEsr = ""
            End If
            
        End If
        
    End If
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function