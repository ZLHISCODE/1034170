VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Medica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�����
'����:??
'����:����
'�û�:�����ж�Ժ
Private mstrLastrTRAY As String
Private mstrCode As String

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
    mstrLastrTRAY = ""
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
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim strLine As String, lngS As Long, lngE As Long, strSampleNo As String, strQC As String, strItem As String, strValue As String
    mstrCode = mstrCode & strSource
    
    Do While InStr(mstrCode, Chr(&HA)) > 0
        lngE = InStr(mstrCode, Chr(&HA))
        strLine = Mid$(mstrCode, 1, lngE)
        mstrCode = Mid$(mstrCode, lngE + 1)
        If InStr(strLine, "TRAY =") > 0 Then
            mstrLastrTRAY = Val(Split(strLine, "TRAY =")(1))
        ElseIf strLine Like "SAMPLE*Cl*" And mstrLastrTRAY <> "" Then
            strSampleNo = Val(mstrLastrTRAY & Mid(strLine, 8, 3))
            strItem = "|Na|" & Val(Split(strLine, "Na")(1)) & _
                      "|K|" & Val(Split(strLine, "K")(1)) & _
                      "|Cl|" & Val(Split(strLine, "Cl")(1))
            
            If Val(strSampleNo) <> 0 And strItem <> "" Then
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd hh:MM:ss") & "|" & strSampleNo & "| |Ѫ��|0" & strItem
            End If
        ElseIf strLine Like "NORM CNTRL*Cl*" Then
            strSampleNo = "1"
            strItem = "|Na|" & Val(Split(strLine, "Na")(1)) & _
                      "|K|" & Val(Split(strLine, "K")(1)) & _
                      "|Cl|" & Val(Split(strLine, "Cl")(1))
            If Val(strSampleNo) <> 0 And strItem <> "" Then
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd hh:MM:ss") & "|" & strSampleNo & "| |Ѫ��|1" & strItem
            End If
            mstrLastrTRAY = ""
        ElseIf strLine Like "ABNRM CNTRL*Cl*" Then
            strSampleNo = "2"
            strItem = "|Na|" & Val(Split(strLine, "Na")(1)) & _
                      "|K|" & Val(Split(strLine, "K")(1)) & _
                      "|Cl|" & Val(Split(strLine, "Cl")(1))
            mstrLastrTRAY = ""
            If Val(strSampleNo) <> 0 And strItem <> "" Then
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd hh:MM:ss") & "|" & strSampleNo & "| |Ѫ��|1" & strItem
            End If
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    
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
