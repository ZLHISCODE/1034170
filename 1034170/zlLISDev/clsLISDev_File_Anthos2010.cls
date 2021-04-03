VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Anthos2010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ø����Anthos2010
'����:֣�ݰ�ͼ�������޹�˾
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'��׼���������ļ���ȡ������
'   strFile������·�����ļ���
'   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
'   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
'   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
'   ���صļ���������(������������밴���±�׼��֯���)
'   Ԫ��֮����|�ָ�
'   ��0��Ԫ�أ�����ʱ��
'   ��1��Ԫ�أ��������
'   ��2��Ԫ�أ�������
'   ��3��Ԫ�أ��걾
'   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
'   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strChenel As String, str����Ա As String, str���� As String
    Dim strFilePath As String
    Dim strfilename As String
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim varStr As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strData As String           '��������
    Dim intCount As Integer         '����
    Dim blnBegin As Boolean         '�Ƿ�ʼ����
    Dim blnItem As Boolean          '�Ƿ�ʼһ����Ŀ
    Dim intloop As Integer
    Dim aitem() As String           '��Ŀ
    Dim strDate As String           'ʱ��
    
    On Error GoTo errHandle
    'û���ҵ��ļ�ʱ�˳�
    If Dir(strFile) = "" Then Exit Function
    ReDim Preserve aResults(0)
    Open strFile For Input As #1 ' ���ļ���
    Do While Not EOF(1) ' ѭ�����ļ�β��
        Line Input #1, strLine ' ����һ�����ݲ����丳��ĳ������
        
        'ȥ���м����Ŀո�
        
        For intloop = 20 To 2 Step -1
            strLine = Replace(strLine, Space(intloop), Space(1))
        Next
        strLine = Replace(strLine, Chr(9), " ")
        aitem = Split(strLine, Space(1))
        
        If UBound(aitem) > 2 And aitem(0) <> "��Ʒ���" Then
            strData = aitem(10) & "|" & aitem(0) & "| |Ѫ��|0|" & Replace(aitem(12), "���ƻ���", "") & "|" & _
            Replace(Replace(aitem(15), "-", "����(-)"), "+", "����(+)") & "^" & Format(aitem(14), "#0.00#") & "^^" & aitem(16)
            Debug.Print strData
            aResults(UBound(aResults)) = strData
            ReDim Preserve aResults(UBound(aResults) + 1)
        End If
        
    Loop
    Close #1 ' �ر��ļ���
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_MP_2.ResultFromFile", strFile, Err.Number & " " & Err.Description
    Close #1
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    'ת�����ĵ��·�Ϊ����
    Select Case UCase(strMonth)
        Case "һ��": TrunMonth = "01"
        Case "����": TrunMonth = "02"
        Case "����": TrunMonth = "03"
        Case "����": TrunMonth = "04"
        Case "����": TrunMonth = "05"
        Case "����": TrunMonth = "06"
        Case "����": TrunMonth = "07"
        Case "����": TrunMonth = "08"
        Case "����": TrunMonth = "09"
        Case "ʮ��": TrunMonth = "10"
        Case "ʮһ": TrunMonth = "11"
        Case "ʮ����": TrunMonth = "12"
        
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select
End Function





