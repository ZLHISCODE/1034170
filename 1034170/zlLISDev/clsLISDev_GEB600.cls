VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_geb600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:GEB600
'����:ϣɭ����
'����:
'�û�:տ����Ժ

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
    Dim lngS As Long, lngE As Long, i As Integer
    Dim str�걾�� As String, str��Ŀ As String, str��� As String, str���� As String, str�ʿ� As String
    Dim strItems As String
    Dim varline As Variant
    Dim iRow As Integer
    
    strReserved = strSource
    
    If strReserved Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        lngS = InStr(strReserved, Chr(2))
        lngE = InStr(strReserved, Chr(3))
        strCode = Mid(strReserved, lngS, lngE - lngS + 1)
        strReserved = Mid(strReserved, lngE + 1)
    
        varline = Split(strCode, vbCrLf)
        
        str���� = Replace(Mid(varline(1), 6), "/", "-")
        If Not IsDate(str����) Then
            str���� = Format(Now, "yyyy-MM-dd hh:mm:ss")
        Else
            str���� = Format(CDate(str����), "yyyy-MM-dd hh:mm:ss")
        End If
        'str�ʿ� = Mid(varline(2), 7, 1)
        str�ʿ� = "0"
        str�걾�� = Val(Mid(varline(1), 2, 7))
        strItems = ""
        For iRow = 3 To UBound(varline)
            str��Ŀ = Mid(varline(iRow), 1, 5)
            str��Ŀ = Replace(str��Ŀ, "*", "")
            str��Ŀ = UCase(Replace(str��Ŀ, " ", ""))
            If str��Ŀ = "LEU" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str��Ŀ = "NIT" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str��Ŀ = "URO" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Replace(Mid(varline(iRow), 6, 12), " ", "")
            End If
            If str��Ŀ = "BIL" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str��Ŀ = "VC" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str��Ŀ = "PRO" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str��Ŀ = "BLD" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str��Ŀ = "PH" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Val(Mid(varline(iRow), 6))
            End If
            If str��Ŀ = "SG" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Val(Mid(varline(iRow), 6))
            End If
            If str��Ŀ = "GLU" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str��Ŀ = "KET" Then
                strItems = strItems & "|" & str��Ŀ & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
        Next
        

        If Val(str�걾��) > 0 And strItems <> "" Then
            strResult = str���� & "|" & Val(str�걾��) & "| |ѪҺ|" & str�ʿ� & strItems
        End If
    End If
    
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



