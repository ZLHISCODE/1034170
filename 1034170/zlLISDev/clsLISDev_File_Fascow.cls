VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Fascow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
    
        Dim aResults() As String
        Dim cnAccess As New ADODB.Connection
        Dim strFilePath As String, objFile As New Scripting.FileSystemObject
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim lngResultCount As Long, strBh As String, strLine_Man As String, strLine_Wman As String, str��Ŀ��λ As String
        Dim str�� As String, str�� As String, str�� As String, strimg As String
        Dim blnUnion As Boolean
        Dim str���� As String
        Dim str�����ͺ� As String, strItems As String, strQC As String
        Dim arrLow(2) As Double, arrMid(2) As Double, arrHigh(2) As Double, arrNianDu(2) As Double
    
        On Error GoTo hErr
100     strFilePath = objFile.GetParentFolderName(strFile)
102     If Len(strFilePath) = 0 Then strFilePath = App.Path
    
104     str�����ͺ� = Trim(ReadIni("FASCOW", "�����ͺ�", App.Path & "\FASCOW.ini"))
106     If str�����ͺ� = "" Then str�����ͺ� = "Ĭ��": Call WriteIni("FASCOW", "�����ͺ�", str�����ͺ�, App.Path & "\FASCOW.ini")
    
108     With cnAccess
110         .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
112         .Open
114         If .State = adStateClosed Then Exit Function
116         strSQL = "Select ��Ŀ����,��Ŀ��λ,��Ů��,��Ů��,������,������ From Item Where ��Ŀ����='ȫѪ����' "
118         Set rsTmp = cnAccess.Execute(strSQL)
120         strLine_Man = "": strLine_Wman = "": str��Ŀ��λ = ""
122         Do Until rsTmp.EOF
124             str��Ŀ��λ = "" & rsTmp!��Ŀ��λ
126             str��Ŀ��λ = Mid(str��Ŀ��λ, InStr(str��Ŀ��λ, "(") + 1, InStr(str��Ŀ��λ, "/") - InStr(str��Ŀ��λ, "(") - 1)
128             str�� = str��Ŀ��λ
130             strLine_Wman = "" & rsTmp!��Ů�� & "," & str��Ŀ��λ & "|" & rsTmp!��Ů�� & "," & str��Ŀ��λ
132             strLine_Man = "" & rsTmp!������ & "," & str��Ŀ��λ & "|" & rsTmp!������ & "," & str��Ŀ��λ
134             rsTmp.MoveNext
            Loop
        
136         strSQL = "Select ��Ŀ����,��Ŀ��λ,��Ů��,��Ů��,������,������ From Item Where ��Ŀ����='ȫѪ����' "
138         Set rsTmp = cnAccess.Execute(strSQL)
140         Do Until rsTmp.EOF
142             str��Ŀ��λ = "" & rsTmp!��Ŀ��λ
144             str��Ŀ��λ = Mid(str��Ŀ��λ, InStr(str��Ŀ��λ, "(") + 1, InStr(str��Ŀ��λ, "/") - InStr(str��Ŀ��λ, "(") - 1)
146             str�� = str��Ŀ��λ
148             strLine_Wman = Split(strLine_Wman, "|")(0) & "," & rsTmp!��Ů�� & "," & str��Ŀ��λ & "|" & Split(strLine_Wman, "|")(1) & "," & rsTmp!��Ů�� & "," & str��Ŀ��λ
150             strLine_Man = Split(strLine_Man, "|")(0) & "," & rsTmp!������ & "," & str��Ŀ��λ & "|" & Split(strLine_Man, "|")(1) & "," & rsTmp!������ & "," & str��Ŀ��λ
152             rsTmp.MoveNext
            Loop
        
154         strSQL = "Select ��Ŀ����,��Ŀ��λ,��Ů��,��Ů��,������,������ From Item Where ��Ŀ����='ȫѪ����' "
156         Set rsTmp = cnAccess.Execute(strSQL)
158         Do Until rsTmp.EOF
160             str��Ŀ��λ = "" & rsTmp!��Ŀ��λ
162             str��Ŀ��λ = Mid(str��Ŀ��λ, InStr(str��Ŀ��λ, "(") + 1, InStr(str��Ŀ��λ, "/") - InStr(str��Ŀ��λ, "(") - 1)
164             str�� = str��Ŀ��λ
166             strLine_Wman = Split(strLine_Wman, "|")(0) & "," & rsTmp!��Ů�� & "," & str��Ŀ��λ & "|" & Split(strLine_Wman, "|")(1) & "," & rsTmp!��Ů�� & "," & str��Ŀ��λ
168             strLine_Man = Split(strLine_Man, "|")(0) & "," & rsTmp!������ & "," & str��Ŀ��λ & "|" & Split(strLine_Man, "|")(1) & "," & rsTmp!������ & "," & str��Ŀ��λ
170             rsTmp.MoveNext
            Loop
        
172         If str�����ͺ� = "2010D" Then
174             strSQL = "Select ��Ŀ����,��Ŀ��λ,��Ů��,��Ů��,������,������ From Item Where ��Ŀ����='Ѫ��ճ��' "
176             Set rsTmp = cnAccess.Execute(strSQL)
178             Do Until rsTmp.EOF
180                 str��Ŀ��λ = "" & rsTmp!��Ŀ��λ
182                 str��Ŀ��λ = Mid(str��Ŀ��λ, InStr(str��Ŀ��λ, "(") + 1, InStr(str��Ŀ��λ, ".") - InStr(str��Ŀ��λ, "(") - 1)
184                 str�� = str��Ŀ��λ
186                 strLine_Wman = Split(strLine_Wman, "|")(0) & "," & rsTmp!��Ů�� & "," & str��Ŀ��λ & "|" & Split(strLine_Wman, "|")(1) & "," & rsTmp!��Ů�� & "," & str��Ŀ��λ
188                 strLine_Man = Split(strLine_Man, "|")(0) & "," & rsTmp!������ & "," & str��Ŀ��λ & "|" & Split(strLine_Man, "|")(1) & "," & rsTmp!������ & "," & str��Ŀ��λ
190                 rsTmp.MoveNext
                Loop
            End If
        
192         strSQL = "": blnUnion = False
194         If dtStart < date Then 'Ҫ��ѯ����
196             strSQL = "SELECT a.���,A.�Ա�, Format(A.����,'yyyy-mm-dd') As ����, a.����ҽʦ, 'ѪҺ' As �걾, 0 As �ʿ�,a.ȫѪ����,a.ȫѪ����," & _
                "a.ȫѪ����,Ѫ��ճ��,��ϸ��ѹ��,���л�ԭճ��,���л�ԭճ��,���л�ԭճ��,��ϸ���ۼ�ָ��,��ϸ������ָ��,Ѫ��,Ѫ������Kֵ,��ϸ������ָ��,��ϸ������,��ϸ����Ӿʱ��,����ճ��,��������Ӧ�� " & _
                    " FROM record a WHERE  A.ȫѪ����<>0 and A.ȫѪ����<>0 and A.ȫѪ����<>0 And " & _
                    "  Format(A.����,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                        " And Format(A.����,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'" & _
                        IIf(lngSampleNo = -1, "", " And a.���=" & lngSampleNo)
            End If
198         If dtEnd >= date Then 'Ҫ�鵱��
200             If Len(strSQL) > 0 Then blnUnion = True 'ҪUnion All
            
202             strSQL = IIf(Len(strSQL) = 0, "", strSQL & " Union All ") & _
                    "SELECT a.���,A.�Ա�, Format(A.����,'yyyy-mm-dd') As ����, a.����ҽʦ, 'ѪҺ' As �걾, 0 As �ʿ�,a.ȫѪ����,a.ȫѪ����," & _
                    "a.ȫѪ����,Ѫ��ճ��,��ϸ��ѹ��,���л�ԭճ��,���л�ԭճ��,���л�ԭճ��,��ϸ���ۼ�ָ��,��ϸ������ָ��,Ѫ��,Ѫ������Kֵ,��ϸ������ָ��,��ϸ������,��ϸ����Ӿʱ��,����ճ��,��������Ӧ�� " & _
                    " FROM record a WHERE  A.ȫѪ����<>0 and A.ȫѪ����<>0 and A.ȫѪ����<>0 And " & _
                    "  Format(A.����,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                        " And Format(A.����,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'" & _
                    IIf(lngSampleNo = -1, "", " And a.���=" & lngSampleNo)
            End If
204         strSQL = strSQL & IIf(blnUnion, " Order By ���� Desc , ���", " Order By Format(A.����,'yyyy-mm-dd') Desc , ���")
206         Set rsTmp = cnAccess.Execute(strSQL)
208         If rsTmp.EOF Then
210             .Close
                Exit Function
            End If
        
212         lngResultCount = -1
214         strBh = ""
216         ReDim aResults(0)
            'rsTmp.MoveFirst
218         Do Until rsTmp.EOF
                    
220                 strBh = rsTmp("���")
222                 str���� = Format(rsTmp("����"), "yyyy-MM-dd HH:mm:ss")
224                 strQC = Val("" & rsTmp("�ʿ�"))

226                 strItems = ""
228                 If Trim("" & rsTmp("ȫѪ����")) <> "" Then strItems = strItems & "|01|" & Trim("" & rsTmp("ȫѪ����"))
230                 If Trim("" & rsTmp("ȫѪ����")) <> "" Then strItems = strItems & "|02|" & Trim("" & rsTmp("ȫѪ����"))
232                 If Trim("" & rsTmp("ȫѪ����")) <> "" Then strItems = strItems & "|03|" & Trim("" & rsTmp("ȫѪ����"))
234                 If Trim("" & rsTmp("Ѫ��ճ��")) <> "" Then strItems = strItems & "|04|" & Trim("" & rsTmp("Ѫ��ճ��"))
236                 If Trim("" & rsTmp("��ϸ��ѹ��")) <> "" Then strItems = strItems & "|05|" & Trim("" & rsTmp("��ϸ��ѹ��"))
238                 If Trim("" & rsTmp("���л�ԭճ��")) <> "" Then strItems = strItems & "|06|" & Trim("" & rsTmp("���л�ԭճ��"))
240                 If Trim("" & rsTmp("���л�ԭճ��")) <> "" Then strItems = strItems & "|07|" & Trim("" & rsTmp("���л�ԭճ��"))
242                 If Trim("" & rsTmp("���л�ԭճ��")) <> "" Then strItems = strItems & "|08|" & Trim("" & rsTmp("���л�ԭճ��"))
244                 If Trim("" & rsTmp("��ϸ���ۼ�ָ��")) <> "" Then strItems = strItems & "|09|" & Trim("" & rsTmp("��ϸ���ۼ�ָ��"))
246                 If Trim("" & rsTmp("��ϸ������ָ��")) <> "" Then strItems = strItems & "|10|" & Trim("" & rsTmp("��ϸ������ָ��"))
248                 If Trim("" & rsTmp("Ѫ��")) <> "" Then strItems = strItems & "|11|" & Trim("" & rsTmp("Ѫ��"))
250                 If Trim("" & rsTmp("Ѫ������Kֵ")) <> "" Then strItems = strItems & "|12|" & Trim("" & rsTmp("Ѫ������Kֵ"))
252                 If Trim("" & rsTmp("��ϸ������ָ��")) <> "" Then strItems = strItems & "|13|" & Trim("" & rsTmp("��ϸ������ָ��"))
254                 If Trim("" & rsTmp("��ϸ������")) <> "" Then strItems = strItems & "|14|" & Trim("" & rsTmp("��ϸ������"))
256                 If Trim("" & rsTmp("��ϸ����Ӿʱ��")) <> "" Then strItems = strItems & "|15|" & Trim("" & rsTmp("��ϸ����Ӿʱ��"))
258                 If Trim("" & rsTmp("����ճ��")) <> "" Then strItems = strItems & "|16|" & Trim("" & rsTmp("����ճ��"))
260                 If Trim("" & rsTmp("��������Ӧ��")) <> "" Then strItems = strItems & "|17|" & Trim("" & rsTmp("��������Ӧ��"))
                    
262                 If strItems <> "" And strBh <> "" And IsDate(str����) Then
264                     If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
266                     aResults(UBound(aResults)) = str���� & "|" & strBh & "| |XY|" & strQC & strItems
                    End If
268                 If str�����ͺ� <> "2010D" Then
270                     strimg = vbCrLf & "ճ����������;2;20,210|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|3-3,6-6,9-9,12-12,15-15;"
272                     If "" & rsTmp!�Ա� = "��" Then
274                         strimg = strimg & strLine_Man
                        Else
276                         strimg = strimg & strLine_Wman
                        End If
278                     strimg = strimg & "~" & str�� & "-" & rsTmp!ȫѪ���� & "," & str�� & "-" & rsTmp!ȫѪ���� & "," & str�� & "-" & rsTmp!ȫѪ���� '���
280                     strimg = strimg & ";(mPa.s),15,20~(1/S),205,1" '������
                    Else
                        Dim varTmp() As String
                    
282                     If "" & rsTmp!�Ա� = "��" Then
                            '����
284                         varTmp = Split(Split(strLine_Man, "|")(0), ",")
286                         arrLow(0) = varTmp(0): arrMid(0) = varTmp(4): arrHigh(0) = varTmp(2): arrNianDu(0) = varTmp(6)
                            '����ֵ
288                         arrLow(1) = Val(rsTmp!ȫѪ����): arrMid(1) = Val(rsTmp!ȫѪ����): arrHigh(1) = Val(rsTmp!ȫѪ����): arrNianDu(1) = Val(rsTmp!Ѫ��ճ��)
                            '����
290                         varTmp = Split(Split(strLine_Man, "|")(1), ",")
292                         arrLow(2) = varTmp(0): arrMid(2) = varTmp(4): arrHigh(2) = varTmp(2): arrNianDu(2) = varTmp(6)
                        Else
                             '����
294                         varTmp = Split(Split(strLine_Wman, "|")(0), ",")
296                         arrLow(0) = varTmp(0): arrMid(0) = varTmp(4): arrHigh(0) = varTmp(2): arrNianDu(0) = varTmp(6)
                            '����ֵ
298                         arrLow(1) = Val(rsTmp!ȫѪ����): arrMid(1) = Val(rsTmp!ȫѪ����): arrHigh(1) = Val(rsTmp!ȫѪ����): arrNianDu(1) = Val(rsTmp!Ѫ��ճ��)
                            '����
300                         varTmp = Split(Split(strLine_Wman, "|")(1), ",")
302                         arrLow(2) = varTmp(0): arrMid(2) = varTmp(4): arrHigh(2) = varTmp(2): arrNianDu(2) = varTmp(6)
                        End If
                    
304                     strimg = "ճ����������;110;1;" & frmGraph.Draw_2010D(arrHigh, arrMid, arrLow, arrNianDu, App.Path, rsTmp!���)
                    End If
306                 If strimg <> "" Then
308                         aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & strimg
                    End If
310             rsTmp.MoveNext
            Loop
        End With
312     ResultFromFile = aResults
hErr:
314     WriteLog "Fascow", "��" & CStr(Erl()) & "��", Err.Description
End Function