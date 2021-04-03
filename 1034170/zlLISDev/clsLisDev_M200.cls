VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clslisdev_M200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:M200 Ѫ����
'����:�������ۿ�ѧ������˾
'����:
'�û�:
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
        Dim strCode As String, lngS As Long, lngE As Long
        Dim varData As Variant, strSampleNo As String, strDate As String, strItem As String, strValue As String
        Dim iRow As Integer
        On Error GoTo hErr
100     mstrCode = mstrCode & strSource
102     lngS = InStr(mstrCode, Chr(&H2))
104     lngE = 0
106     If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H3))
    
108     Do While lngS > 0 And lngE > 0
110         strCode = Mid$(mstrCode, lngS, lngE - lngS + 1)
112         mstrCode = Mid$(mstrCode, lngE + 1)
114         strItem = ""
            '--------------------------------------------------------------------
116         varData = Split(strCode, ";")
118         If UBound(varData) > 23 Then
120             strSampleNo = varData(20)
122             strDate = varData(5)
124             For iRow = 21 To UBound(varData) Step 4
126                 If iRow + 1 > UBound(varData) Then Exit For
128                 If varData(iRow + 1) = "INR" Then
130                     If Trim(varData(iRow)) <> "" Then strItem = strItem & "|INR|" & varData(iRow)
132                 ElseIf varData(iRow - 1) = "D-DIMER" Then
134                     If Trim(varData(iRow)) <> "" Then strItem = strItem & "|D-DIMER|" & varData(iRow)
                    Else
136                     If Trim$(varData(iRow)) <> "" And Trim(varData(iRow + 1)) <> "" Then strItem = strItem & "|" & varData(iRow) & "|" & varData(iRow + 1)
                    End If
                Next
138             If strItem <> "" And strSampleNo <> "" And strDate <> "" Then
140                 strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Blood|0" & strItem
                End If
            End If
            '--------------------------------------------------------------------
142         lngS = InStr(mstrCode, Chr(&H2))
144         lngE = 0
146         If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H3))
        Loop
148     If strResult <> "" Then strResult = Mid$(strResult, 3)
        Exit Sub
hErr:
150     WriteLog "M200", strSource, CStr(Erl()) & "��," & Err.Description
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


