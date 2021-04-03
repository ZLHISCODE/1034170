VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_RT2100C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�׶�ø���ǣ�RT-2100C��
'����:Rayto(�׶�)
'����:����
'�û�:����ҽԺ

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
    
    Dim strLine As String, varLine As Variant, var��Ŀ As Variant, strData As String
    Dim str_�걾�� As String, str_���� As String, str_��Ŀ As String
    Dim intRow As Integer, strTmp As String, strR As String
    Dim intLoop As Integer
    Dim aItem() As String
    Dim intCount As Integer
    Dim strҽԺ As String
    
    On Error GoTo errH
    aItem = Split(strSource, ",")
    
    'û�д���һ��ʱ�˳�
    If UBound(aItem) < 15 Then
        strReserved = strSource
        Exit Sub
    End If
    
    strҽԺ = ReadIni("RT2100C", "ҽԺ����", App.Path & "\RT2100C.ini")
    
    If strҽԺ = "" Then
        strҽԺ = "Ĭ��"
    End If
    Call WriteIni("RT2100C", "ҽԺ����", strҽԺ, App.Path & "\RT2100C.ini")
    
    For intLoop = 0 To UBound(aItem)
        If InStr(aItem(intLoop), "R") > 0 Then
            If UBound(aItem) - intLoop > 12 Then
                If strҽԺ = "Ĭ��" Or strҽԺ = "" Then
                    'д��һ���걾
                    strData = strData & "||" & aItem(intLoop + 5) & "|" & aItem(intLoop + 6) & "| |Ѫ��|0|" & aItem(intLoop + 7) & _
                    "|" & aItem(intLoop + 11) & "^" & aItem(intLoop + 12) & "^^"
                ElseIf strҽԺ = "����" Then
                    'д��һ���걾
                    strData = strData & "||" & aItem(intLoop + 5) & "|" & aItem(intLoop + 6) & "| |Ѫ��|0|" & aItem(intLoop + 7) & _
                    "|" & aItem(intLoop + 12)
                End If
            Else
                strReserved = strReserved & "," & aItem(intLoop)
            End If
        Else
            If UBound(aItem) - intLoop < 11 And strReserved <> "" Then
                strReserved = strReserved & "," & aItem(intLoop)
            End If
        End If
    Next
    
    strResult = Mid(strData, 3)
    strReserved = Mid(strReserved, 2)
    Exit Sub
errH:
    WriteLog "Rt2100c", Err.Description, ""
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    blnSuccess = True
    iSendStep = 0
    SendSample = ""
End Function

