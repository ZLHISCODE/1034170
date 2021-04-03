VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_PCRKH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:PCR
'����:�Ϻ��ƻ�
'����:�¶�
'�û�:3201ҽԺ
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
    
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim iRow As Integer '�к�
    Dim varline As Variant
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    iRow = 0
    Do Until objStream.AtEndOfStream
        str������ = "": str����Ա = "": str���� = "": str�걾�� = ""
        
        strLine = objStream.ReadLine
        '�걾�ţ�������
        
        If UBound(Split(strLine, Chr(&H1E))) = 3 Then
            strChenel = Trim(Split(strLine, Chr(&H1E))(2))
            If strChenel <> "" Then
                str����Ա = Trim(Split(strLine, Chr(&H1E))(0))
                
                str������ = Trim(Split(strLine, Chr(&H1E))(1))
                
                If UBound(Split(str������, " ")) > 5 Then
                    str���� = Split(str������, " ")(2) & "-" & Split(str������, " ")(3) & "-" & Split(str������, " ")(4)
                End If
                
                str������ = Split(str������, " ")(5) * 100
                If Val(str������) = 0 Then
                    str������ = "0.000E+0"
                Else
                    str������ = Split(str������, ".")(0)
                    str������ = Format(Val(str������) / 10 ^ (Len(str������) - 1), "0.000") & "E+" & Format(Len(str������) - 3, "00")
                End If
                str�걾�� = ""
                Do While str����Ա <> ""
                    If IsNumeric(Right(str����Ա, 1)) Then
                        str�걾�� = Right(str����Ա, 1) & str�걾��
                        str����Ա = Left(str����Ա, Len(str����Ա) - 1)
                    Else
                        str����Ա = ""
                        Exit Do
                    End If
                Loop
                If Not IsDate(str����) Then str���� = Format(Now, "yyyy-MM-dd")
                
                If Val(str�걾��) > 0 And str������ <> "" And strChenel <> "" Then
                    If CDate(str����) >= dtStart And CDate(str����) <= dtEnd Then
                        If lngSampleNO = -1 Or lngSampleNO = Val(str�걾��) Then
                            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                            aResults(UBound(aResults)) = Format(CDate(str����), "yyyy-MM-dd") & "|" & str�걾�� & "| |ѪҺ|0|" & strChenel & "|" & str������
                        End If
                    End If
                End If
            End If
        
        End If
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_Anytest.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function