VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_IQ200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ȫ�Զ������������ IQ200
'����:����Iris
'����:���
'�û�:�Ͼ��н�������ҽԺ(����ǿ)

Private mstrCode As String

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
    
    Dim varline As Variant, varItem As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim varTmp As Variant, intLoop As Integer
    Dim iRow As Integer, strTmp As String, strLine As String, str��Ŀ As String
    Dim lng_S As Long, lng_E As Long
    Dim strData As String, strNode As String
    Dim strValues As String, lngLength As Long
    Dim str�걾���� As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)

    mstrCode = mstrCode & strSource
    
    If mstrCode Like "***" Then
        blnEnd = False
'        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            lngLength = InStr(mstrCode, "") - InStr(mstrCode, "")
            strData = Mid(mstrCode, InStr(mstrCode, ""), lngLength)
            mstrCode = Mid(mstrCode, InStr(mstrCode, "") + 1)
            
            str���� = "": str�걾�� = "":  strItem = "": strTmp = ""
            '����
            '��һ�������յ���16��������תΪ XML ��ʽ
            varline = Split(strData, vbCrLf)
            strItem = "": strData = ""
            For iRow = LBound(varline) To UBound(varline) - 1
                lng_S = InStr(varline(iRow), "") + 2
                lng_E = IIf(InStr(varline(iRow), "") > 0, InStr(varline(iRow), ""), InStr(varline(iRow), ""))
                
                strLine = Mid(varline(iRow), lng_S, lng_E - lng_S)
                strData = strData & GetStr(strLine)
            Next
            
            '�ڶ��������� XML �ļ�
            strData = Replace(Replace(strData, Chr(&HD), ""), Chr(&HA), "")     '�滻�س����з�
            
            '�����걾�źͼ�������
            strData = Mid(strData, InStr(strData, "<SA"))
            strNode = Mid(strData, 1, InStr(strData, ">"))
                
            varTmp = Split(strNode, " ")
            For intLoop = LBound(varTmp) To UBound(varTmp)
                varItem = Split(varTmp(intLoop), "=")
                If varItem(0) = "ID" Then str�걾�� = Replace(varItem(1), """", "")
                If varItem(0) = "REDTS" Then str���� = Format(Replace(varItem(1), """", ""), "yyyy-MM-dd")
            Next
            
            '����������Ŀ�ͼ�����
            strData = Mid(strData, InStr(strData, "<AR "))
            lng_S = InStr(strData, "<AR ")
            lng_E = InStr(strData, "</AR>") + 4
            
            While lng_S > 0 And lng_E > 0
                strData = Mid(strData, lng_S)
                strNode = Mid(strData, 1, lng_E)
                strData = Mid(strData, lng_E + 1)
                
                '�������ִ���
                '��Ŀ
                strTmp = Mid(strNode, InStr(strNode, "<"), InStr(strNode, ">") - InStr(strNode, "<") + 1)
                strNode = Mid(strNode, InStr(strNode, ">") + 1)
                varTmp = Split(strTmp, " ")
                For intLoop = LBound(varTmp) To UBound(varTmp)
                    varItem = Split(varTmp(intLoop), "=")
                    If varItem(0) = "Key" Then strItem = strItem & "|" & Replace(varItem(1), """", "")
                Next
                '���
                strTmp = Mid(strNode, 1, InStr(strNode, "<") - 1)
                If strTmp = "[none]" Then
                    strItem = strItem & "|��"
                ElseIf InStr(strTmp, "/") > 0 Then
                    strItem = strItem & "|" & Trim(Replace(Replace(Split(strTmp, "/")(0), "&gt;", ">"), "&lt;", "<"))
                Else
                    strItem = strItem & "|����û��"
                End If
                
                lng_S = InStr(strData, "<AR ")
                lng_E = InStr(strData, "</AR>") + 4
            Wend
            
            If strItem <> "" And Val(str�걾��) <> 0 Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "| |��Һ|0" & strItem
            End If
            If Not mstrCode Like "" Then
                blnEnd = True
            End If
        Loop

    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
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

'��16�����ַ���ת��Ϊ�ַ���ʽ
Public Function GetStr(ByVal strBin As String) As String
    Dim strTmp As String
    Dim strResult As String
    Dim intLoop As Integer
    
    For intLoop = 1 To Len(strBin) Step 4
        strTmp = Mid(strBin, intLoop, 2)
        strResult = strResult & Chr("&H" & strTmp)
    Next
    GetStr = strResult
End Function

