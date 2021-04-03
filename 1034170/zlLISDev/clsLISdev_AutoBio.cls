VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_AutoBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: ��ͼø����
'����:��ͼ
'����:�¶�
'�û�:������ҽԺ
'2012-10-31
Private mstrCode As String
 
Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    'GetStartCmd = Chr(6)
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
    Dim lngS As Long, lngE As Long, strCode As String
    Dim varData As Variant, strLine As String, iRow As Integer
    Dim strSampleNo As String, strDate As String, strChannel As String, strValue As String
    Dim varTmp As Variant
    On Error GoTo hErr
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    lngS = InStr(mstrCode, Chr(&H5))
    lngE = 0
    lngE = InStr(lngS, mstrCode, Chr(&H4))
    
    strResult = ""
    
    Do While lngS > 0 And lngE > 0
        strCode = Mid$(mstrCode, lngS, lngE - lngS + 1)
        mstrCode = Mid$(mstrCode, lngE + 1)
        '-----------------------------------------------------------------
        varData = Split(strCode, Chr(&HD) & Chr(&HA))
        For iRow = LBound(varData) To UBound(varData)
            strLine = varData(iRow)
            varTmp = Split(strLine, Chr(&H9))
            
            If UBound(varTmp) >= 3 Then
                strSampleNo = varTmp(0)
                If IsNumeric(strSampleNo) Then
                    strDate = Left(strSampleNo, 6)
                    strDate = "20" & Mid$(strDate, 1, 2) & "-" & Mid$(strDate, 3, 2) & "-" & Mid$(strDate, 5, 2)
                    strSampleNo = Mid(strSampleNo, 7)
                    strChannel = varTmp(1)
                    strChannel = Replace$(strChannel, "��", "")
                    strValue = varTmp(2) ' & "^" & varTmp(14) & "^" & varTmp(16)
                    
                    If IsDate(strDate) And strSampleNo <> "" And strChannel <> "" And strValue <> "" Then
                       strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Boold|0|" & strChannel & "|" & strValue
                    Else
                        Call SaveLog(iRow & "�����ݲ���ȷ���걾��" & strSampleNo & ",����" & strDate & "��ͨ��" & strChannel & ",���" & strValue)
                    End If
                End If
            End If
        Next
        '-----------------------------------------------------------------
        lngS = InStr(mstrCode, Chr(&H5))
        lngE = 0
        lngE = InStr(lngS, mstrCode, Chr(&H4))
    Loop
    If strResult <> "" Then strResult = Mid$(strResult, 3)
    Exit Sub
hErr:
    Call SaveLog("��" & CStr(Erl()) & "�У�" & Err.Description)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    
End Function

Private Sub SaveLog(ByVal strOutput As String)
    '------------------------------------------------------
    '--  ����:���ݵ��Ա�־,д��־����ǰĿ¼
    '------------------------------------------------------
    
    '���±������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    If Dir(App.Path & "\����.TXT") = "" Then Exit Sub
    strFileName = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"
    
    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    strOutput = Replace(strOutput, Chr(&HD), "<CR>")
    strOutput = Replace(strOutput, Chr(&HA), "<LF>")
    objStream.WriteLine (strDate & ":   " & strOutput)
    objStream.Close
    Set objStream = Nothing
End Sub



