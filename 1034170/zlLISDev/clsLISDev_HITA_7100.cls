VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA7100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  ����7100
'ҽԺ:  ɽ��
'����:  �¶�
'ʱ��:  2011/10/24

Private mstrCode As String  '�������ַ���

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
        Dim lng_S As Long, lng_E As Long, str_Code As String, strDate As String
        Dim strSampleNo As String, strItem As String, strTmp  As String
        Dim intQ As Integer, intE As Integer, intLoop As Integer
        On Error GoTo hErr
    
100     If InStr(strSource, Chr(&H3)) > 0 Then strCmd = Chr(&H6)
102     mstrCode = mstrCode & strSource
        
104     lng_S = InStr(mstrCode, Chr(&H2))
106     If lng_S > 0 Then
108         lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
110         lng_E = 0
        End If
112     strResult = ""
114     Do While lng_S > 0 And lng_E > 0
116         str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
118         mstrCode = Mid(mstrCode, lng_E + 1)
            '-------------------------------------------------------
120         If InStr(",:,1,2,", Mid(str_Code, 2, 1)) > 0 Then
            
122             intE = 0
124             If UCase(Mid(str_Code, 3, 1)) = "Q" Then intE = "1"
            
126             intQ = 0
128             If Mid(str_Code, 3, 1) = "f" Then intQ = "1"
            
130             strSampleNo = ""
132             If Val(Mid(str_Code, 5, 5)) = 0 Then 'ȡ����
134                 If Val(Mid(str_Code, 15, 2)) <> 0 Then
                        '����ǰ��λ��Ϊ0 ����������
136                     strSampleNo = "^" & IIf(intE = 1, "1", "") & "^" & Val(Mid(str_Code, 15, 13)) & "| |Ѫ��|" & intQ
                    Else
                        '����ǰ��λΪ0,�Ͱ�������Ϊ�걾�Ŵ���
138                     strSampleNo = Val(Mid(str_Code, 15, 13)) & IIf(intE = 1, "^1^", "") & "| |Ѫ��|" & intQ
                    End If
                Else
140                 strSampleNo = Val(Mid(str_Code, 5, 5)) & IIf(intE = 1, "^1^", "") & "| |Ѫ��|" & intQ
                End If
        
                '��ʼ�������
142             str_Code = Replace(Mid(str_Code, 52), "", "")
144             strItem = ""
146             For intLoop = 0 To Len(str_Code) / 10
148                 If Len(str_Code) < 10 Then Exit For
150                 strTmp = Mid(str_Code, 1, 10)
152                 strItem = strItem & "|" & Trim(Mid(strTmp, 1, 3)) & "|" & Val(Trim(Mid(strTmp, 4)))
154                 str_Code = Mid(str_Code, 11)
                Next
        
156             If strItem <> "" And strSampleNo <> "" Then
158                 strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "|" & strSampleNo & strItem
                End If
            End If
            '-------------------------------------------------------
160         lng_S = InStr(mstrCode, Chr(&H2))
162         If lng_S > 0 Then
164             lng_E = InStr(lng_S, mstrCode, Chr(&H3))
            Else
166             lng_E = 0
            End If
        Loop
168     If strResult <> "" Then strResult = Mid(strResult, 3)
        Exit Sub
hErr:
170     WriteLog "H7100", strSource, CStr(Erl()) & " " & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    blnSuccess = True
    strSample = ""
    iSendStep = 0
End Function

