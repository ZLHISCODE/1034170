VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_K4500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:K-4500
'����:ϣɭ����
'����:�¶�
'�û�:�Ͼ���������ҽԺ

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = Chr(6)
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String, lngLength As Long, strRecord As String, lngByte As Long
    Dim str_Code As String, str_Day As String, StrWbc As String, StrRbc As String, StrPLt As String, strBit As String
    
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    strCmd = ",06" 'Ӧ��Ϊ������
    '----  ȡһ�������Ĵ�
    If Len_Bin(strSource) < 456 Then
        strReserved = strSource
        Exit Sub
    End If
    
    For lngByte = 1 To Len_Bin(strSource) - 4
        str_Code = Mid_bin(strSource, lngByte, 1)
        str_Day = Mid_bin(strSource, lngByte + 3, 1)
    
        If str_Code = "" And str_Day = "/" Then
    
            If Len_Bin(strSource) - lngByte < 456 Then
                strReserved = strSource
                Exit Sub
            End If
    
            strLine = Mid_bin(strSource, lngByte, 457, False)
            strReserved = Mid_bin(strSource, lngByte + 457, False)
            Exit For
        End If
    Next
    
    If lngByte = Len_Bin(strSource) - 3 Then
        Exit Sub
    End If
    '-- ��ʼ���� strLine
    str�걾�� = Mid_bin(strLine, 15, 13)
    str���� = Mid_bin(strLine, 2, 2)
    str���� = str���� & "-" & Mid_bin(strLine, 5, 2)
    str���� = str���� & "-" & Mid_bin(strLine, 8, 2)
    If str�걾�� <> 0 And IsDate(str����) Then
        strResult = "||" & Format(str����, "yyyy-MM-dd") & "|" & Val(str�걾��) & "| |ѪҺ|0"
        For lngByte = 1 To 21
            strItem = Mid_bin(strLine, 78 + lngByte * 5, 5)
        
            If ((((lngByte = 8) Or (lngByte = 9)) Or (lngByte = 10)) Or (lngByte = 14)) Then
                strItem = Format(Val(strItem) * 100, "##0.0")
            End If
        
            strResult = strResult & "|" & lngByte & "|" & strItem
        Next
        
        'WBC
        StrWbc = ""
        For lngByte = 243 To 243 + 50
            strBit = Mid_bin(strLine, lngByte, 1, False)
            StrWbc = StrWbc & ";" & Format(CByte("&H" & Replace(strBit, ",", "")), "000")
        Next
        'RBC
        For lngByte = 293 To 293 + 50
            strBit = Mid_bin(strLine, lngByte, 1, False)
            StrRbc = StrRbc & ";" & Format(CByte("&H" & Replace(strBit, ",", "")), "000")
        Next

        'PLT
        For lngByte = 343 To 343 + 40
            strBit = Mid_bin(strLine, lngByte, 1, False)
            StrPLt = StrPLt & ";" & Format(CByte("&H" & Replace(strBit, ",", "")), "000")
        Next
    End If
    
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3) & _
           Chr(13) & Chr(10) & _
          "WBC;0;100,50,8-,17-100,25-,33-200,42-,50-300" & StrWbc & "^" & _
          "RBC;0;100,50,10-,20-100,30-,40-200" & StrRbc & "^" & _
          "PLT;0;100,40,10-10,20-20,30-30" & StrPLt
        
    End If
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function


