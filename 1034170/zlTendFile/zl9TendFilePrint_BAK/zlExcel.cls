VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlExcel"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'EXCEL��VBA�е�һЩ����
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const xlCenter As Integer = -4108
Private Const xlEdgeLeft As Integer = 7
Private Const xlEdgeTop As Integer = 8
Private Const xlEdgeBottom As Integer = 9
Private Const xlEdgeRight As Integer = 10
Private Const xlContinuous As Integer = 1
Private Const xlInsideVertical As Integer = 11
Private Const xlInsideHorizontal As Integer = 12


Public ExcelApp As Object        '��ƷʱҪ��ΪObject����
Public frmTempExcel As Object

Private mlngWideTable As Long
Private mintChar As Integer            'һ���ַ��Ŀ���
Private mintProgress As Integer
Private mstrTempFile As String

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'���ܣ� �������������EXCEEL�������
'
'zlExcelFile                �����������תվ

'OutTextFile                ����������ݵ��ı��ļ���
'TendsToFile                ��msFlexGrid�����������ı��ļ�
'TendsToExcel               ��msFlexGrid����������Excel
'strRow                     ���һ�б��ϻ������Ŀ��EXCEL�����ʱ������
'RowInttoStr                �������к�ת��ΪEXCEL�еı�ʾ����
'chgAlignment               ��VB�Ķ��뷽ʽת��ΪEXCEL�еĶ��뷽ʽ
'GrdAlignment               ��FlexGrid�Ķ��뷽ʽת��ΪEXCEL�еĶ��뷽ʽ
'SetProgress                ��ʾ����
'

Public Sub zlExcelFile()
     '------------------------------------------------
    '���ܣ� �����������תվ
    '������
    '���أ� �ɹ�����true �����󷵻�false
    '������̷�:40%���EXCEL,30%����ı��ļ�;10%����ʱ�ļ�;2%����;4%ҳü;4%ҳ��;����10%
    '------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim intPercent As Integer, intPagePer As Integer
    Dim lngCurPage As Long, lngMaxPage As Long, lngStartPage As Long, lngPages As Long
    
    On Error GoTo errHandle
    
'    If gblnIsWps Then
'        '���˺�:2007/4/20
'        '��WPSΪ׼
'        Set ExcelApp = CreateObject("ET.Application") '��һ��WPS�е�ET����
'    Else
        '���ӵ�EXCEL��
        Set ExcelApp = CreateObject("Excel.Application") '��һ��EXCEL����
        ExcelApp.Workbooks.Add
'    End If
    
    intPercent = 5
    SetProgress intPercent
    
    '��ҳ��ѭ��,ÿ��ҳ�����һ��SHEET
    lngMaxPage = frmTendFileReader.GetPages
    lngStartPage = frmTendFileReader.GetStartPage
    lngPages = lngMaxPage - lngStartPage
    intPagePer = 100
    
    gfrmTemp.Font.Name = ExcelApp.StandardFont
    gfrmTemp.Font.Size = CCur(ExcelApp.StandardFontSize)
'    If gblnIsWps Then
'        '���˺�:2007/4/20
'        '�ó�1�����ٸ��
'        mintChar = gfrmTemp.ScaleX(1, vbPoints, vbTwips)
'    Else
    mintChar = CInt(gfrmTemp.TextWidth("0"))   '�õ�һ�����ϡ�������Ŀ���ַ�����
'    End If
    
    gfrmTemp.Font.Name = "����"
    gfrmTemp.Font.Size = 9
    intPercent = intPercent + (intPagePer * 0.03)
    SetProgress intPercent
    
    Call setHeaderFooter
    intPercent = intPercent + (intPagePer * 0.04)
    SetProgress intPercent
    
    Call TendsToExcel(intPercent, (intPagePer * 0.4))
    intPercent = intPercent + (intPagePer * 0.4)
    SetProgress intPercent
    
    '��������ı��ļ���
    Call OutTextFile 'ͬʱ�������������
    intPercent = intPercent + (intPagePer * 0.3)
    SetProgress intPercent
    
'    If gblnIsWps Then
'        '���˺�:2007/4/20
'        ExcelApp.Workbooks.Add
'        ExcelApp.Cells.Select
'        ExcelApp.Selection.NumberFormatLocal = "@"
'        ExcelApp.Range("A1").Select
'        SetProgress 95
'        'ճ����ʽ
'        Clipboard.Clear
'        With ExcelApp
'            .Windows(1).Activate
'            .Cells.Select
'            .Selection.Copy
'            .Windows(2).Activate
'            .Cells.Select
'            .Selection.PasteSpecial Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False
'        End With
'        Set objText = objFile.OpenTextFile(mstrTempFile)
'        Clipboard.Clear
'        Clipboard.SetText objText.ReadAll
'        Call ExcelApp.Sheets(1).Paste
'        Clipboard.Clear
'        SetProgress 96
'    Else
        Call ExcelApp.Workbooks.OpenText(mstrTempFile, , 1, 1, 1, False, True, False, False, False, False, , arrFormat)
        intPercent = intPercent + (intPagePer * 0.1)
        SetProgress intPercent
        'ճ����ʽ
        With ExcelApp
            .Windows(2).Activate
            .Cells.Select
            .Selection.Copy
            .Windows(2).Activate
            .Cells.Select
            .Selection.PasteSpecial Paste:=-4122, Operation:=-4142, SkipBlanks:=False, Transpose:=False
        End With
'    End If
    
    Clipboard.Clear
'    If gblnIsWps Then
'        '���˺�:2007/4/20:�رյ�һ������
'        ExcelApp.Windows(1).Close False
'    Else
    ExcelApp.Windows(2).Close False
'    End If
    intPercent = intPercent + (intPagePer * 0.03)
    SetProgress intPercent
    
    ExcelApp.Range("A1").Select
    
    SetProgress 97
    ExcelApp.Visible = True
    SetProgress 100
    Unload frmTempExcel
    Set frmTempExcel = Nothing
    Exit Sub
    
errHandle:
    MsgBox "�����Excelʱ�������ش������������ֹ��" & vbCrLf & Err.Description, vbCritical, gstrSysName
    Unload frmTempExcel
    Set ExcelApp = Nothing
    Set frmTempExcel = Nothing
End Sub

Private Function OutTextFile() As Boolean
    '------------------------------------------------
    '���ܣ��Ȱ�Ҫ�������������һ���ı��ļ�
    '------------------------------------------------
    Dim strFile As String, strPath As String
    Dim intFileNum As Integer
    '������ʱ�ļ�
    strPath = Space(256): strFile = Space(256)
    GetTempPath 256, strPath
    strPath = Left$(strPath, InStr(strPath, Chr(0)) - 1)
    
    GetTempFileName strPath, "excel", 0, strFile
    mstrTempFile = Left$(strFile, InStr(strFile, Chr(0)) - 1)
    '���ļ�׼�����
    intFileNum = FreeFile()
    Open mstrTempFile For Binary Access Write As intFileNum
    
    TendsToFile intFileNum
    Close intFileNum
End Function

Private Sub setHeaderFooter()
    '------------------------------------------------
    '���ܣ� ����ҳü��ҳ��
    '------------------------------------------------
    Dim strLeft As String, strMiddle As String, strRight As String
    Dim intPos As Long
    Dim intPos1 As Long
    Dim strHeader As String, strFooter As String
    
    On Error Resume Next
    strHeader = ConvHFtoExcel(gobjSend.Header)
    intPos = InStr(strHeader, ";")
    intPos1 = intPos + 1
    strLeft = Mid(strHeader, 1, intPos - 1)
    intPos = InStr(intPos1, strHeader, ";")
    strMiddle = Mid(strHeader, intPos1, intPos - intPos1)
    intPos1 = intPos + 1
    strRight = Mid(strHeader, intPos1)

    With ExcelApp.ActiveSheet.PageSetup
        .LeftHeader = strLeft
        .CenterHeader = strMiddle
        .RightHeader = strRight
    End With
    
    strFooter = ConvHFtoExcel(gobjSend.Footer)
    intPos = InStr(strFooter, ";")
    intPos1 = intPos + 1
    strLeft = Mid(strFooter, 1, intPos - 1)
    intPos = InStr(intPos1, strFooter, ";")
    strMiddle = Mid(strFooter, intPos1, intPos - intPos1)
    intPos1 = intPos + 1
    strRight = Mid(strFooter, intPos1)

    With ExcelApp.ActiveSheet.PageSetup
        .LeftFooter = strLeft
        .CenterFooter = strMiddle
        .RightFooter = strRight
    End With

End Sub

Private Function ConvHFtoExcel(ByVal strSource As String) As String
    '------------------------------------------------
    '���ܣ���ҳü��ҳ��ת����ʵ�ʴ�ӡ������
    '������strSource    ҳü��ҳ��
    '���أ�ʵ�ʴ�ӡ������
    '------------------------------------------------
    Dim strTemp As String
    
    strTemp = Replace(strSource, "[ҳ��]", "&P")
    strTemp = Replace(strTemp, "[ҳ��]", "&N")
    strTemp = Replace(strTemp, "[ʱ��]", "&T")
    strTemp = Replace(strTemp, "[����]", "&D")
    
    Dim strKeyValue As String       '��ֵ
    Dim lngKey As Long
    Dim lngKeySize As Long
    Dim strRegPath As String
    If IsWindows95 Then
        strRegPath = "Software\MicroSoft\Windows\CurrentVersion"
    Else
        strRegPath = "Software\MicroSoft\Windows NT\CurrentVersion"
    End If
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, strRegPath, 0, KEY_READ, lngKey) = 0 Then
        strKeyValue = Space(256)
        lngKeySize = 256
        If RegQueryValueEx(lngKey, "RegisteredOrganization", 0, 1, strKeyValue, lngKeySize) = 0 Then
            gstrUnitName = StripTerminator(strKeyValue)
        End If
        strKeyValue = Space(256)
        lngKeySize = 256
        If RegQueryValueEx(lngKey, "RegisteredOwner", 0, 1, strKeyValue, lngKeySize) = 0 Then
            gstrUserName = StripTerminator(strKeyValue)
        End If
    End If
    RegCloseKey lngKey

    strTemp = Replace(strTemp, "[�û���]", gstrUserName)
    strTemp = Replace(strTemp, "[��λ��]", gstrUnitName)
    ConvHFtoExcel = strTemp
End Function

Private Sub ExcelGrid(GrdTemp As Object, intRow As Long, ByVal LeftCol As Long, ByVal RightCol As Long, ByVal TopRow As Long, ByVal BottomRow As Long, intBegin As Long, intPrecent As Integer)
'���ܣ�����ָ������ı������ö�ӦExcel����ĸ�ʽ�������ϲ���Ԫ��
'������GrdTemp   ����ؼ�
'      intRow    Excel�е���ʼ��
'      LeftCol   �����е���ʼ��
'      RightCol  �����е���ֹ��
'      TopRow    �����е���ʼ��
'      BottomRow �����е���ֹ��
'      intBegin      ��ʼ����ֵ
'      intPrecent    ������ռ�õĽ���ֵ

    Dim intY As Long
    Dim strTemp As String
    Dim iCount As Long, intColcnt As Long
    Dim intBlankǰ As Long, intBlank�� As Long '�ϲ���Ԫ��ǰ���������
    Dim iTemp As Long, intRowUp As Long
    Dim blnMerge As Boolean
    
    
    If GrdTemp.MergeCells = 0 Then Exit Sub
    
    '�ϲ���Ԫ��
    '����ϲ�
    intY = intRow
    For iCount = TopRow To BottomRow - 1
        If GrdTemp.MergeRow(iCount) = True And iCount < GrdTemp.FixedRows Then
            intBlankǰ = 0
            For intColcnt = LeftCol To RightCol - 2
                If GrdTemp.ColWidth(intColcnt) <> 0 And Not GrdTemp.ColHidden(intColcnt) Then
                    'ֻ�з������вŵõ�����
                    If GrdTemp.TextMatrix(iCount, intColcnt) <> "" Then '�հ�����Ͳ��ϲ�
                        strTemp = RowInttoStr(intColcnt + 1 - intBlankǰ) & Trim(str(intY))
                        ExcelApp.Range(strTemp).Select
                        If Not ExcelApp.Selection.MergeCells Then   '�Ѻϲ���
                            '��ʼ��ѭ���ڲ�ʹ�õı���
                            blnMerge = False
                            intBlank�� = 0
                            For iTemp = intColcnt + 1 To RightCol - 1
                                If GrdTemp.ColWidth(iTemp) <> 0 Then
                                    '��ȻҪ�ж��п��������в���ʾ
                                    '��֪��һ����Ԫ��ͬ���˳�
                                    If GrdTemp.TextMatrix(iCount, intColcnt) <> GrdTemp.TextMatrix(iCount, iTemp) Then Exit For
                                    If GrdTemp.MergeCells = 3 Or GrdTemp.MergeCells = 4 Then  '��������
                                        intRowUp = iCount - 1
                                        Do While intRowUp >= TopRow
                                            '����û�ϲ����˳�
                                            If GrdTemp.TextMatrix(intRowUp, intColcnt) <> GrdTemp.TextMatrix(intRowUp, iTemp) Then Exit For
                                            intRowUp = intRowUp - 1
                                        Loop
                                    End If
                                    
                                    blnMerge = True
                                Else
                                    intBlank�� = intBlank�� + 1
                                End If
                            Next
                            If blnMerge = True Then
                                strTemp = RowInttoStr(intColcnt + 1 - intBlankǰ) & Trim(str(intY)) & ":" & RowInttoStr(iTemp - intBlankǰ - intBlank��) & Trim(str(intY))
                                ExcelApp.Range(strTemp).Select
                                ExcelApp.Selection.MergeCells = True
                                intColcnt = iTemp - 1 '�����Ѻϲ�����
                            End If
                        End If
                    End If
                Else
                    intBlankǰ = intBlankǰ + 1
                End If
            Next
         End If
        intY = intY + 1
        'SetProgress intBegin + (iCount - TopRow) / (BottomRow - TopRow) * intPrecent / 2
    Next
    '����ϲ�
    intY = intRow
    intBlankǰ = 0
    For intColcnt = LeftCol To RightCol - 1
        If GrdTemp.ColWidth(intColcnt) <> 0 And Not GrdTemp.ColHidden(intColcnt) Then
            'ֻ���������ص���
            If GrdTemp.MergeCol(intColcnt) = True Then
                For iCount = TopRow To BottomRow - 2
                    If GrdTemp.TextMatrix(iCount, intColcnt) <> "" Then '�հ�����Ͳ��ϲ�
                        strTemp = RowInttoStr(intColcnt + 1 - intBlankǰ) & Trim(str(intY + iCount))
                        ExcelApp.Range(strTemp).Select
                        If ExcelApp.Selection.MergeCells = False Then  '�Ѻϲ���
                            blnMerge = False
                            For iTemp = iCount + 1 To BottomRow - 1
                                '��֪��һ����Ԫ��ͬ���˳�
                                If GrdTemp.TextMatrix(iCount, intColcnt) <> GrdTemp.TextMatrix(iTemp, intColcnt) Then Exit For
                                If GrdTemp.MergeCells = 2 Or GrdTemp.MergeCells = 4 Then  '��������
                                    intRowUp = intColcnt - 1
                                    Do While intRowUp >= LeftCol
                                        '����û�ϲ����˳�
                                        If GrdTemp.TextMatrix(iCount, intRowUp) <> GrdTemp.TextMatrix(iTemp, intRowUp) Then Exit For
                                        intRowUp = intRowUp - 1
                                    Loop
                                End If
                                
                                blnMerge = True
                            Next
                            If blnMerge = True Then
                                strTemp = RowInttoStr(intColcnt + 1 - intBlankǰ) & Trim(str(intY + iCount - TopRow)) & ":" & RowInttoStr(intColcnt + 1 - intBlankǰ) & Trim(str(intY + iTemp - 1 - TopRow))
                                ExcelApp.Range(strTemp).Select
                                ExcelApp.Selection.MergeCells = True
                                iCount = iTemp - 1 '�����Ѻϲ�����
                            End If
                        End If
                    End If
                Next
             End If
        Else
            intBlankǰ = intBlankǰ + 1
        End If
        'SetProgress intBegin + intPrecent / 2 + (intColcnt - LeftCol) / (RightCol - LeftCol) * intPrecent / 2
    Next
End Sub

Private Function TendsToExcel(ByVal intCurPer As Integer, ByVal intMax As Integer)
    '------------------------------------------------
    '���ܣ� ��msFlexGrid��������
    '������
    '���أ� �ɹ�����true �����󷵻�false
    '------------------------------------------------
    Dim intY As Long
    Dim GrdTemp As Object
    Dim strTemp As String
    Dim iCount As Long, intColcnt As Long
    
    Dim intBlankCol As Long '�հ��е�����
    
    Set GrdTemp = gobjSend.Body
    mlngWideTable = 0
    For intColcnt = 0 To GrdTemp.Cols - 1
        If GrdTemp.ColWidth(intColcnt) <> 0 And Not GrdTemp.ColHidden(intColcnt) Then
            mlngWideTable = mlngWideTable + GrdTemp.ColWidth(intColcnt)
            ExcelApp.Columns(RowInttoStr(intColcnt - intBlankCol + 1) & ":" & RowInttoStr(intColcnt - intBlankCol + 1)).Select
            If GrdTemp.ColWidth(intColcnt) / mintChar > 0 Then ExcelApp.Selection.ColumnWidth = GrdTemp.ColWidth(intColcnt) / mintChar
        Else
            intBlankCol = intBlankCol + 1
        End If
    Next
    intCurPer = intCurPer + intMax * 0.2
    SetProgress intCurPer
    
    '�������ݸ�ʽ
    intBlankCol = 0
    For intColcnt = 0 To GrdTemp.Cols - 1
        GrdTemp.Col = intColcnt
        If intColcnt = 5 Then
                Debug.Print GrdTemp.Text
            
        End If
        If GrdTemp.ColWidth(intColcnt) = 0 Or GrdTemp.ColHidden(intColcnt) Then
            intBlankCol = intBlankCol + 1
        Else
            For iCount = GrdTemp.FixedRows To GrdTemp.Rows - 1
                GrdTemp.Row = iCount
                '���ǿհ׾��и�ʽ
                If Trim(GrdTemp.Text) <> "" Then Exit For
            Next
            strTemp = RowInttoStr(intColcnt - intBlankCol + 1)
            strTemp = strTemp & ":" & strTemp
            ExcelApp.Columns(strTemp).Select
            With ExcelApp.Selection
                Debug.Print GrdTemp.Text
                If IsNumeric(GrdTemp.Text) And Len(GrdTemp.Text) <= 15 Then
                    iCount = InStr(GrdTemp.Text, ".")
                    If iCount = 0 Then
                        .NumberFormatLocal = "0_ "
                    Else
                        .NumberFormatLocal = "0." & String(Len(Mid(GrdTemp.Text, iCount + 1)), "0") & "_ "
                    End If
                ElseIf IsDate(GrdTemp.Text) Then
                    If InStr(GrdTemp.Text, ":") > 0 Or InStr(GrdTemp.Text, "��") > 0 Then
                        If InStr(GrdTemp.Text, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd hh:mm:ss"
                        Else
                            .NumberFormatLocal = "yyyy""��""mm""��""dd""��"" hh""ʱ""mm""��""ss""��"""
                        End If
                    Else
                        If InStr(GrdTemp.Text, "-") > 0 Then
                            .NumberFormatLocal = "yyyy-mm-dd"
                        Else
                            .NumberFormatLocal = "yyyy""��""mm""��""dd""��"""
                        End If
                    End If
                Else
                    .NumberFormatLocal = "@"
                End If
                .Font.Name = GrdTemp.CellFontName
                .Font.Size = GrdTemp.CellFontSize
                If ExcelApp.Selection.MergeCells = True Then
                    .HorizontalAlignment = -4108
                Else
                    .HorizontalAlignment = GrdAlignment(GrdTemp)
                End If
                .VerticalAlignment = -4108
            End With
        End If
    Next
    intCurPer = intCurPer + intMax * 0.4
    SetProgress intCurPer

    ExcelGrid GrdTemp, gobjSend.UnderAppRows.Count + 2, 0, GrdTemp.Cols, 0, GrdTemp.FixedRows, 16, 9
    'ExcelGrid GrdTemp, gobjSend.UnderAppRows.Count + GrdTemp.FixedRows + 2, 0, GrdTemp.Cols, GrdTemp.FixedRows, GrdTemp.Rows, 25, 10
        
    
    '����������
    intY = gobjSend.UnderAppRows.Count + GrdTemp.Rows + 1
    ExcelApp.Range("A" & Trim(str(gobjSend.UnderAppRows.Count + 2)) & ":" & RowInttoStr(GrdTemp.Cols - intBlankCol) & Trim(str(intY))).Select
    With ExcelApp.Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        If GrdTemp.Cols - intBlankCol > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
        If intY <> gobjSend.UnderAppRows.Count + 2 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        '2007/04/20:���˺����,��Ҫ�ǵ���wpsʱ,������,��Ҫ���������ߵľ�ϸ�̶�
        If gblnIsWps Then
            .Borders.Weight = 2
        End If
        
    End With
    intCurPer = intCurPer + intMax * 0.1
    SetProgress intCurPer
    
    '��ͷ��ʽ
    intBlankCol = 0
    For intColcnt = 0 To GrdTemp.Cols - 1
        GrdTemp.Col = intColcnt
        If GrdTemp.ColWidth(intColcnt) = 0 Then
            intBlankCol = intBlankCol + 1
        Else
            For iCount = 0 To GrdTemp.FixedRows - 1
                GrdTemp.Row = iCount
                strTemp = RowInttoStr(intColcnt - intBlankCol + 1) & Trim(CStr(iCount + gobjSend.UnderAppRows.Count + 2))
                ExcelApp.Range(strTemp).Select
                With ExcelApp.Selection
                    .NumberFormatLocal = "@"
                    .Font.Name = GrdTemp.CellFontName
                    .Font.Size = GrdTemp.CellFontSize
                    .HorizontalAlignment = -4108
                    .VerticalAlignment = -4108
                End With
            Next
        End If
    Next
    intCurPer = intCurPer + intMax * 0.1
    SetProgress intCurPer
    '����ĸ�ʽ
    ExcelApp.Range("A1:" & RowInttoStr(GrdTemp.Cols - intBlankCol) & "1").Select
    With ExcelApp.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = -4108
        .MergeCells = True
        .Font.Name = gobjSend.Title.Font.Name
        .Font.Size = gobjSend.Title.Font.Size
        .RowHeight = GetstrHeight(.Font.Name, .Font.Size, "��") * 2 / 20
    End With
    intCurPer = intCurPer + intMax * 0.1
    SetProgress intCurPer
    '������Ŀ���
    intY = 2
    For iCount = 1 To gobjSend.UnderAppRows.Count
        ExcelApp.Range("A" & Trim(str(intY)) & ":" & RowInttoStr(GrdTemp.Cols - intBlankCol) & Trim(str(intY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            
            .NumberFormatLocal = "G/ͨ�ø�ʽ"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
        End With
        intY = intY + 1
    Next
    intCurPer = intCurPer + intMax * 0.1
    SetProgress intCurPer
    '������Ŀ���
    intY = gobjSend.UnderAppRows.Count + GrdTemp.Rows + 2
    For iCount = 1 To gobjSend.BelowAppRows.Count
        ExcelApp.Range("A" & Trim(str(intY)) & ":" & RowInttoStr(GrdTemp.Cols - intBlankCol) & Trim(str(intY))).Select
        With ExcelApp.Selection
            .Font.Name = gobjSend.AppFont.Name
            .Font.Size = gobjSend.AppFont.Size
            .NumberFormatLocal = "G/ͨ�ø�ʽ"
            .HorizontalAlignment = -4131
            .VerticalAlignment = -4108
            .MergeCells = True
            
        End With
        intY = intY + 1
    Next
End Function

Private Function TendsToFile(ByVal intFileNumber As Integer)
    '------------------------------------------------
    '���ܣ� ��msFlexGrid��������
    '������
    '���أ� �ɹ�����true �����󷵻�false
    '------------------------------------------------
    Dim GrdTemp As Object, i As Long
    Dim iCount As Long, intColcnt As Long
    Dim intSum As Long, strText As String
    
    Set GrdTemp = gobjSend.Body
  
    '�������
    Put intFileNumber, , Replace(gobjSend.Title.Text, vbCrLf, "")
    Put intFileNumber, , vbCrLf
    'SetProgress 52
    
    '������Ŀ���
    Dim AppRowTemp As zlTabAppRow
    For Each AppRowTemp In gobjSend.UnderAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    'SetProgress 53
    
   '��ͷ���
    intSum = GrdTemp.FixedRows
    For iCount = 0 To GrdTemp.FixedRows - 1
        GrdTemp.Row = iCount
        For intColcnt = 0 To GrdTemp.Cols - 1
            GrdTemp.Col = intColcnt
            If GrdTemp.ColWidth(intColcnt) <> 0 And Not GrdTemp.ColHidden(intColcnt) Then
                Put intFileNumber, , Replace(GrdTemp.Text, vbCrLf, "")
                Put intFileNumber, , vbTab
            End If
        Next
        Put intFileNumber, , vbCrLf
        'SetProgress 53 + (iCount / intSum * 7)
    Next
    
    'Excel���Ԥ���ʽ
    iCount = 0
    For i = 0 To GrdTemp.Cols - 1
        If GrdTemp.ColWidth(i) <> 0 And Not GrdTemp.ColHidden(i) Then iCount = iCount + 1
    Next
    If iCount <> 0 Then
        ReDim arrFormat(iCount - 1, 1) As Integer
        For i = 0 To iCount - 1
            arrFormat(i, 0) = i + 1 '�к�
            arrFormat(i, 1) = 1 '���������ͣ�xlGeneralFormat = 1,xlTextFormat = 2
        Next
    End If
    
     '�����������
    intSum = GrdTemp.Rows - GrdTemp.FixedRows
    For iCount = GrdTemp.FixedRows To GrdTemp.Rows - 1
        i = 0: GrdTemp.Row = iCount
        'If Not GrdTemp.RowHidden(iCount) Then
            For intColcnt = 0 To GrdTemp.Cols - 1
                GrdTemp.Col = intColcnt
                If GrdTemp.ColWidth(intColcnt) <> 0 And Not GrdTemp.ColHidden(intColcnt) Then
                    strText = Replace(GrdTemp.Text, vbCrLf, "")
                   '���ո������ǿ�д���Ϊ�ı���ʽ
                    If IsNumeric(strText) And Left(strText, 2) = "  " Then
                        arrFormat(i, 1) = 2
                    End If
                
                    Put intFileNumber, , strText
                    Put intFileNumber, , vbTab
                    
                    '�������:ֻҪ��һ����������,�򳤶ȴ���15,��Ϊ�ı�
                    strText = Trim(strText)
                    If strText <> "" And (Not IsNumeric(strText) Or Len(strText) > 15) _
                        And InStr(",�ϼ�,С��,", "," & Replace(strText, " ", "") & ",") = 0 Then
                        arrFormat(i, 1) = 2
                    End If
                    i = i + 1
                End If
            Next
            Put intFileNumber, , vbCrLf
            'SetProgress 60 + (iCount / intSum * 34)
        'End If
    Next
    
    '������Ŀ���
    For Each AppRowTemp In gobjSend.BelowAppRows
        Put intFileNumber, , strRow(AppRowTemp)
        Put intFileNumber, , vbCrLf
    Next
    'SetProgress 94
    
End Function

Private Function strRow(zlrow As zlTabAppRow) As String
    '------------------------------------------------
    '���ܣ����һ�б��ϻ������Ŀ��EXCEL�����ʱ������
    '������
    '------------------------------------------------

    Dim varTempstr As Variant
    Dim intCurrentWidth As Long
    Dim strBlankSum As String
    
    gfrmTemp.Font.Name = gobjSend.AppFont.Name
    gfrmTemp.Font.Size = CCur(gobjSend.AppFont.Size)
    
    For Each varTempstr In zlrow
        intCurrentWidth = intCurrentWidth + CInt(gfrmTemp.TextWidth(varTempstr.Text)) '��ʱ�õ�������Ч�ַ�����
    Next
    intCurrentWidth = (mlngWideTable - intCurrentWidth) / CInt(gfrmTemp.TextWidth("0"))  '��ʱ�õ����ǿո���
    If intCurrentWidth < 0 Then intCurrentWidth = 0
    strBlankSum = String(intCurrentWidth, " ")
    
    gfrmTemp.Font.Name = "����"
    gfrmTemp.Font.Size = 9
    
    For Each varTempstr In zlrow
        If zlrow.Count = 1 Then
            strRow = strRow & varTempstr.Text
        Else
            strRow = strRow & varTempstr.Text & String(intCurrentWidth / (zlrow.Count - 1), " ")
        End If
    Next
    strRow = RTrim(strRow)
End Function

Private Function RowInttoStr(ByVal intRow As Long) As String
    '------------------------------------------------
    '���ܣ��������к�ת��ΪEXCEL�еı�ʾ����
    '������
    '------------------------------------------------

    Dim int1 As Long, int2 As Long
    If intRow < 27 Then
        RowInttoStr = Chr(Asc("A") + intRow - 1)
        Exit Function
    End If
    
    int1 = intRow \ 26
    int2 = intRow Mod 26
    If int2 = 0 Then
        int1 = int1 - 1
        int2 = 26
    End If
    RowInttoStr = Chr(Asc("A") + int1 - 1) & Chr(Asc("A") + int2 - 1)
End Function
Private Function chgAlignment(ByVal GrdAlignment As Integer) As Long
    '------------------------------------------------
    '���ܣ���VB�Ķ��뷽ʽת��ΪEXCEL�еĶ��뷽ʽ
    '������GrdAlignment     FlexGrid�Ķ��뷽ʽ
    '���أ�EXCEL�еĶ��뷽ʽ
    '------------------------------------------------

    Select Case GrdAlignment
         Case 0       '�����
            chgAlignment = -4131
         Case 2        '����
            chgAlignment = -4108
         Case 1        '�Ҷ���
            chgAlignment = -4152
        Case Else
            chgAlignment = -4131
    End Select
End Function
Private Function GrdAlignment(objGrid As Object) As Long
    '------------------------------------------------
    '���ܣ���FlexGrid�Ķ��뷽ʽת��ΪEXCEL�еĶ��뷽ʽ
    '������objGrid      FlexGrid����
    '���أ�EXCEL�еĶ��뷽ʽ
    '------------------------------------------------
    Dim Alignment As Integer
        
    '��ȡ�������ԣ�
    If objGrid.CellAlignment <> 0 Then
        Alignment = objGrid.CellAlignment           '���յ�Ԫ
    Else
        If objGrid.Col < objGrid.FixedCols Or objGrid.Row < objGrid.FixedRows Then
            Alignment = objGrid.FixedAlignment(objGrid.Col) '���չ̶���Ԫ
        Else
            Alignment = objGrid.ColAlignment(objGrid.Col)   '������
        End If
    End If
    Select Case Alignment
    Case 0, 1, 2        '�����
        GrdAlignment = -4131
    Case 3, 4, 5        '����
        GrdAlignment = -4108
    Case 6, 7, 8        '�Ҷ���
        GrdAlignment = -4152
    Case 9
        If IsNumeric(Trim(objGrid.Text)) Then
            GrdAlignment = -4152
        Else
            GrdAlignment = -4131
        End If
    Case Else
            GrdAlignment = -4131
    End Select

End Function

Private Sub SetProgress(ByVal i As Integer)
    '------------------------------------------------
    '���ܣ���ʾ����
    '������i        ����ֵ
    '���أ�
    '------------------------------------------------

    mintProgress = i
    If mintProgress > 100 Then mintProgress = 100
    frmTempExcel.lblnum = Trim(str(mintProgress))
    frmTempExcel.ProgressBar1.Value = mintProgress
    frmTempExcel.Refresh
End Sub


Private Function GetstrHeight(strFontName As String, curFontSize As Currency, str As String) As Long
'���ܣ������������÷���һ���ַ����߶�
    Dim objFrm As frmAbout
    Set objFrm = New frmAbout
    objFrm.Font.Name = strFontName
    objFrm.Font.Size = curFontSize
    GetstrHeight = objFrm.TextHeight(str)
    Unload objFrm
    Set objFrm = Nothing
End Function
