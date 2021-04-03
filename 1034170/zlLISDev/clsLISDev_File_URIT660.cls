VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_URIT660"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'ҽԺ���ƣ������о���ҽԺ
'�������ͣ�ø����
'�����ͺţ�URIT-660
'�������ң�����������
'��    �ߣ��µ���
'��    �ڣ�2012-06-01
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

    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String, ingCount As Long, ingBh As Long
    Dim blnUnion As Boolean
    Dim strResult As String, strTmp As String
    Dim strOD As String
    Dim strCutOff As String
    Dim strCOV As String
    
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    With cnAccess
        If .State = adStateOpen Then .Close
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        
        If .State = adStateClosed Then
            'MsgBox "�����ݿ�����ʧ��,��������Դ�Ƿ���ȷ���ã�"
            Exit Function
        End If
        
        strSQL = "": strResult = "": strTmp = ""
        strSQL = "Select * From (SELECT Sample.PlateID,Sample.SampleNO, Format(Plate.TestDate+Plate.TestTime,'yyyy-mm-dd hh:mm:ss') AS Drq," & _
            "Item.Code AS ItemCode,Format(CDbl(Sample.Result),'###0.0###') As OD,Sample.Label AS Result,Sample.WellTypeID " & _
            "FROM ((Sample LEFT JOIN Plate ON Sample.PlateID=Plate.PlateID) LEFT JOIN Item ON Sample.ItemID=Item.ID) LEFT JOIN WellType ON Sample.WellTypeID=WellType.WellTypeID) A " & _
            "Where WellTypeID=2 And Drq >='" & Format(dtStart, "YYYY-MM-DD 00:00:00") & "'" & " And Drq <='" & Format(dtEnd, "YYYY-MM-DD 23:59:59") & "'" & _
            IIf(lngSampleNO = -1, "", " And val(PlateID)=" & lngSampleNO) & "  Order by SampleNO"
        
        Open "C:\OutSql.txt" For Output As #1
        Print #1, strSQL
        Close #1
        
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            lngResultCount = 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = ""
            ResultFromFile = aResults
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        
        ingCount = 0
                  
        Do While Not rsTmp.EOF
            ingBh = Val(Mid("" & rsTmp("SampleNO"), 7))
            
            If strBh <> Str(ingBh) Then '�±걾
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(IIf(IsNull(rsTmp("Drq")), Now, rsTmp("Drq")), "yyyy-mm-dd hh:mm:ss") '����ʱ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Val(ingBh)  '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "| "     '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|Ѫ��"  '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"     '�ʿ�
                strBh = Mid(Trim("" & rsTmp("SampleNO")), 7)
            End If
                
            If rsTmp("OD") > 0 Then
                strOD = IIf(Trim(rsTmp("OD")) = "", "0.000", Format(rsTmp("OD"), "###0.0###"))
                strCutOff = "0.000"
                strCOV = "0.000"
            Else
                strOD = "0.000"
                strCutOff = "0.000"
                strCOV = "0.000"
            End If
            strTmp = Trim("" & rsTmp("Result"))
            Select Case strTmp
                Case "1"
                  strResult = "����(+)"
                Case "2"
                  strResult = "����(-)"
            End Select
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                 rsTmp("ItemCode") & "|" & strResult & "^" & strOD & "^" & strCutOff & "^" & strCOV '���
            rsTmp.MoveNext
        Loop
    End With
    For ingCount = LBound(aResults) To UBound(aResults)
      Debug.Print aResults(ingCount)
    Next
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_URIT660.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function
