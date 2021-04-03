VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SH212C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�������� SH212C
Private maccess As New ADODB.Connection
Private mstrsql As String
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, ByVal dtStart As Date, _
                    Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    
    Dim access As New ADODB.Recordset, Rsym As New ADODB.Recordset
    Dim i As Integer, strPass As String
    Dim str��� As String
    Dim strֵ As String
    Dim aResults() As String, strFilePath As String, objFile As New Scripting.FileSystemObject, strFileErs As String
    i = UBound(Split(strFile, "/"))
    
    If i >= 1 Then
       strPass = Split(strFile, "/")(1)
       strFilePath = objFile.GetParentFolderName(Split(strFile, "/")(0)) 'ȡ·��
       strFile = Split(strFile, "/")(0)
    Else
       strFilePath = objFile.GetParentFolderName(strFile) 'ȡ·��
    End If
    
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    strFileErs = strFile

    If i >= 1 Then
       If Trim(Dir(strFileErs)) <> "" Then
          maccess.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & strFile & _
                                     ";Jet OLEDB:Database Password=" & strPass
          maccess.Open
          If maccess.State = adStateClosed Then strFileErs = ""
       Else
          strFileErs = ""
       End If
    Else
       If Trim(Dir(strFileErs)) <> "" Then
          maccess.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & strFile
          maccess.Open
          If maccess.State = adStateClosed Then strFileErs = ""
       Else
          strFileErs = ""
       End If
    End If
    ReDim aResults(0)
        
    If strFileErs <> "" Then
    
    
        mstrsql = "SELECT �ͼ����� AS ����ʱ��,���ID AS �걾��,[ȫѪ����ճ��(200/S)] as a11,[ȫѪ����ճ��(40/S)] as a12,[ȫѪ����ճ��(30/S)] as a13,[ȫѪ����ճ��(10/S)] as a14,[ȫѪ����ճ��(1/S)] as a15,[��������(e9SI)] as a16,[��������(e9SI)] as a17,[��������(e9SI)] as a18,[ȫѪ��ɭճ��(mPas)] as a19,[ȫѪ��ɭӦ��(dyn/cm^2)] as a20," & _
        "[Ѫ��ճ��(mPas)] as a21,[��ϸ��ѹ��(Hct)] as a22,[ȫѪ��ԭճ��(200/S)] as a23,[ȫѪ��ԭճ��(40/S)] as a24,[ȫѪ��ԭճ��(30/S)] as a25,[ȫѪ��ԭճ��(10/S)] as a26,[ȫѪ��ԭճ��(1/S)] as a27,[ȫѪ�������ճ��(BR)] as a28,[ȫѪ�������ճ��(Br)] as a29,[��ϸ���ۼ�ָ��(AI)] as a30,[��ϸ���ۼ�ϵ��(VAI)] as a31,[��ϸ��������(TK)] as a32,[��ϸ������ָ��(IR)] as a33,[��ϸ������(e12��/L)] as a34,[��ϸ����ճ��(mPas)] as a35,[��ϸ����Ӿʱ��(s)] as a36,[Ѫ�쵰��Ũ��(g/L)] as a37," & _
        "[��ϸ��������(mm/h)] as a38,[ѪС����ճ��ǰ] as a39,[ѪС����ճ����] as a40,[ѪС��ճ����(PAd)] as a41,[Ѫ������Kֵ(ESRK)] as a42,[����Ѫ˨����(mm)] as a43,[����Ѫ˨����(mg)] as a44,[����Ѫ˨ʪ��(mg)] as a45,[��ά����ԭ(g/L)] as a46,[���̴�(mmol/L)] as a47,[������֬(mmol/L)] as a48,[Ѫ��(mmol/L)] as a49,[���ܶ�֬����(mmol/L)] as a50,[���ܶ�֬����(mmol/L)] as a51 " & _
        " FROM �������ݱ�  where "
        
        mstrsql = mstrsql & "Format(�ͼ�����,'yyyy-MM-dd')>='" & _
        Format(dtStart, "yyyy-MM-dd") & "'"
        mstrsql = mstrsql & " And Format(�ͼ�����,'yyyy-MM-dd')<='" & _
        Format(dtEnd, "yyyy-MM-dd") & "'"
        
        Call openaccess(Rsym, "ȡ��ϸ")
              
        Do While Not Rsym.EOF
            str��� = Rsym("����ʱ��") & "|" & Rsym("�걾��") & "|" & " |ѪҺ|0"
            
            For i = 11 To 51
               
                strֵ = "a" & i
                str��� = str��� & "|" & i & "|" & Rsym(strֵ)
              
            Next
        
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                     aResults(UBound(aResults)) = str���
             Rsym.MoveNext
        Loop

        ResultFromFile = aResults
        str��� = ""
        maccess.Close
    End If
End Function

Private Sub openaccess(rsTemp As ADODB.Recordset, ByVal strCaption As String)
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open mstrsql, maccess, adOpenStatic, adLockReadOnly
End Sub



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



