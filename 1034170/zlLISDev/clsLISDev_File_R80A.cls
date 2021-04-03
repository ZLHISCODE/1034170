VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_R80A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ����
'���ң������������ۿ�ѧ������˾
'ҽԺ��������ʡũ����ҽԺ,�Ͼ���ʮ��ҽԺ,������ʯҽԺ
'2009-03-08 ZHQ
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
    '���صļ���������(������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������(΢��������"^"��Ϊ����,�ֱ���:��ҩ^����^���)
    
    '    2) ճ����������:ͼ������;ͼ�񻭷�;��������;���߼��������;�������������
    '                  ����  �������ݣ�Y����,X����|X����-X������ʾ������,....|Y����-Y������ʾ������,....
    '                  ���߼��������:ճ������1�ĸߵ�͵͵�����|ճ������2�ĸߵ�͵͵�����~���е�����,���е�����,���е�����
    '                  �������������:Y�����������,X����,Y����~X�����������,X����,Y����
    '       ��:ճ����������;2;20,200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;
    '          9.25,10,4.4,150|6.5,10,3.65,150~10-8.989,60-4.803,150-4.05;
    '          VIS(mPa.s),25,20~SHR(1/S),195,1

    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long, lngBh As Long
    Dim blnUnion As Boolean
    Dim i As Integer
    Dim strItems As String '�ֶ�
    Dim strimg As String
    
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = " Select * From Checks Where CheckDate between cdate('" & Format(dtStart, "yyyy-mm-dd") & "') and cdate('" & Format(dtEnd, "yyyy-mm-dd") & "')"
        
        If lngSampleNo <> -1 Then
           strSQL = strSQL & " and val(CheckNum) = " & lngSampleNo
        End If
        strSQL = strSQL & " Order By CheckDate,CheckNum"
        
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        '---  ���Ӻ�����ʯҽԺ�����Դ�������ʯҽԺ�ޡ�LowAlbumen���ֶΡ�
        '---  �����¼�������ֶΣ����ں����ж�,�д��ֶΣ���ִ�������䣬�޴��ֶΣ���ִ�С�
        strItems = ""
        For i = 0 To rsTmp.Fields.count - 1
            strItems = strItems & "," & rsTmp.Fields(i).Name
        Next
        
        lngResultCount = -1
        Do While Not rsTmp.EOF
            If InStr(";123456789;", Mid(rsTmp!checknum, 1, 1)) <= 0 Or Not IsNumeric(rsTmp!checknum) Then
                '��λ�Ǿ������־�Ĭ��Ϊ�ʿر걾
                lngBh = 999
            Else
                lngBh = rsTmp!checknum
            End If
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(rsTmp!CheckDate, "yyyy-MM-DD") & "|" & lngBh & "| " & "|ѪҺ|" & IIf(lngBh = 999, 1, 0)
            '���濪ʼ���������Ŀ��ֵ
            If Nvl(rsTmp!BloodValueHighHigh, 0) <> 0 Then
                'ȫѪճ��ֵ(mPa/s): ͨ����200
                aResults(lngResultCount) = aResults(lngResultCount) & "|1|" & rsTmp!BloodValueHighHigh
            End If
            If Nvl(rsTmp!BloodValueHigh, 0) <> 0 Then
                'ȫѪճ��ֵ(mPa/s): ͨ����100
                aResults(lngResultCount) = aResults(lngResultCount) & "|2|" & rsTmp!BloodValueHigh
            End If
            If Nvl(rsTmp!BloodValueMiddle, 0) <> 0 Then
                'ȫѪճ��ֵ(mPa/s): ͨ����5
                aResults(lngResultCount) = aResults(lngResultCount) & "|3|" & rsTmp!BloodValueMiddle
            End If
            If Nvl(rsTmp!BloodValueLow, 0) <> 0 Then
                'ȫѪճ��ֵ(mPa/s): ͨ����1
                aResults(lngResultCount) = aResults(lngResultCount) & "|4|" & rsTmp!BloodValueLow
            End If
            If Nvl(rsTmp!PlasmaValue, 0) <> 0 Then
                'Ѫ��ճ��ֵ(mPa/s): ͨ����100
                aResults(lngResultCount) = aResults(lngResultCount) & "|5|" & rsTmp!PlasmaValue
            End If
            If Nvl(rsTmp!ESR, 0) <> 0 Then
                'ESRѪ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|6|" & rsTmp!ESR
            End If
            If Nvl(rsTmp!Press, 0) <> 0 Then
                '��ϸ��ѹ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|7|" & rsTmp!Press
            End If
            If Nvl(rsTmp!HighShearVisc, 0) <> 0 Then
                'ȫѪ���л�ԭճ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|8|" & rsTmp!HighShearVisc
            End If
            If Nvl(rsTmp!LowShearVisc, 0) <> 0 Then
                'ȫѪ���л�ԭճ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|9|" & rsTmp!LowShearVisc
            End If
            If Nvl(rsTmp!Rigidity, 0) <> 0 Then
                '��ϸ������ָ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|10|" & rsTmp!Rigidity
            End If
            If Nvl(rsTmp!Assemble, 0) <> 0 Then
                '��ϸ���ۼ�ָ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|11|" & rsTmp!Assemble
            End If
            If Nvl(rsTmp!EquationK, 0) <> 0 Then
                'Ѫ������Kֵ
                aResults(lngResultCount) = aResults(lngResultCount) & "|12|" & rsTmp!EquationK
            End If
            If Nvl(rsTmp!HighShearOppositeVisc, 0) <> 0 Then
                'ȫѪ�������ճ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|13|" & rsTmp!HighShearOppositeVisc
            End If
            If Nvl(rsTmp!LowShearOppositeVisc, 0) <> 0 Then
                'ȫѪ�������ճ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|14|" & rsTmp!LowShearOppositeVisc
            End If
            If Nvl(rsTmp!TransformTK, 0) <> 0 Then
                '��ϸ������ָ��TK
                aResults(lngResultCount) = aResults(lngResultCount) & "|15|" & rsTmp!TransformTK
            End If
            If Nvl(rsTmp!Proteid, 0) <> 0 Then
                '��ά����ԭ
                aResults(lngResultCount) = aResults(lngResultCount) & "|16|" & rsTmp!Proteid
            End If
            If Nvl(rsTmp!Cholesterin, 0) <> 0 Then
                'Ѫ���̴�
                aResults(lngResultCount) = aResults(lngResultCount) & "|17|" & rsTmp!Cholesterin
            End If
            If Nvl(rsTmp!Glyceride, 0) <> 0 Then
                '������֬
                aResults(lngResultCount) = aResults(lngResultCount) & "|18|" & rsTmp!Glyceride
            End If
            If Nvl(rsTmp!Albumen, 0) <> 0 Then
                '����֬����
                aResults(lngResultCount) = aResults(lngResultCount) & "|19|" & rsTmp!Albumen
            End If
            If Nvl(rsTmp!Bloodsugar, 0) <> 0 Then
                'Ѫ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|20|" & rsTmp!Bloodsugar
            End If
            If Nvl(rsTmp!ThrombusLen, 0) <> 0 Then
                '����Ѫ˨����
                aResults(lngResultCount) = aResults(lngResultCount) & "|21|" & rsTmp!ThrombusLen
            End If
            If Nvl(rsTmp!ThrombusdampWeight, 0) <> 0 Then
                '����Ѫ˨ʪ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|22|" & rsTmp!ThrombusdampWeight
            End If
            If Nvl(rsTmp!ThrombusdryWeight, 0) <> 0 Then
                '����Ѫ˨����
                aResults(lngResultCount) = aResults(lngResultCount) & "|23|" & rsTmp!ThrombusdryWeight
            End If
            If Nvl(rsTmp!BeforePlateletNum, 0) <> 0 Then
                'ճ��ǰѪС����
                aResults(lngResultCount) = aResults(lngResultCount) & "|24|" & rsTmp!BeforePlateletNum
            End If
            If Nvl(rsTmp!AfterPlateletNum, 0) <> 0 Then
                'ճ����ѪС����
                aResults(lngResultCount) = aResults(lngResultCount) & "|25|" & rsTmp!AfterPlateletNum
            End If
            If Nvl(rsTmp!PlateletAdhere, 0) <> 0 Then
                'ѪС��ճ����
                aResults(lngResultCount) = aResults(lngResultCount) & "|26|" & rsTmp!PlateletAdhere
            End If
            If Nvl(rsTmp!Electrophoresis, 0) <> 0 Then
                '��ϸ����Ӿ
                aResults(lngResultCount) = aResults(lngResultCount) & "|27|" & rsTmp!Electrophoresis
            End If
            If Nvl(rsTmp!PlateletAssemble, 0) <> 0 Then
                'ѪС��ۼ���
                aResults(lngResultCount) = aResults(lngResultCount) & "|28|" & rsTmp!PlateletAssemble
            End If
            If InStr(strItems & ",", ",LowAlbumen" & ",") > 0 Then
                If Nvl(rsTmp!LowAlbumen, 0) <> 0 Then
                    '����֬����
                    aResults(lngResultCount) = aResults(lngResultCount) & "|29|" & rsTmp!LowAlbumen
                End If
            End If
            '��������ٴ��ο���   Result(����HIS�޷��ش�����ʱȡ��)
            If Nvl(rsTmp!Result, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|30|" & rsTmp!Result
            End If
           
            '����ͼ��(��ʽ��ͼ������;ͼ�񻭷�;��������;���߼��������;�������������)
            '�����Ա��жϸı仮������ֵ
            If rsTmp!sex = False Then
               strimg = "9.25,10,4.4,150|6.5,10,3.65,150"
               'strimg = "21.35,1,4.4,150|6.5,10,3.65,150"
            Else
               strimg = "9.58,10,4.6,150|6.8,10,3.73,150"
            End If

'            '����ͼ��(��ʽ��ͼ������;ͼ�񻭷�;��������;���߼��������;�������������)
'            aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & "ճ����������;2;" & _
'                "30,200|3-3,10-10,30-30,100-100,200-200|5-5,10-10,15-15,20-20,25-25,30-30;" & _
'                Nvl(rsTmp!BloodValueHighHigh, 0) & "," & Nvl(rsTmp!BloodValueHigh, 0) & "," & Nvl(rsTmp!BloodValueMiddle, 0) & "," & Nvl(rsTmp!BloodValueLow, 0) & ";" & _
'                "ճ��ֵ(mPa.s),1,35~�б���(1/s),210,1"
            
            aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & "ճ����������;2;" & _
                "20,230|3-3,10-10,30-30,100-100,200-200|5-5,10-10,15-15,20-20,25-25,30-30;" & _
                strimg & _
                "~5-" & Nvl(rsTmp!BloodValueLow, 0) & ",30-" & Nvl(rsTmp!BloodValueMiddle, 0) & ",200-" & Nvl(rsTmp!BloodValueHighHigh, 0) & _
                ";ճ��ֵ(mPa.s),35,20~�б���(1/s),210,1"
  
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function