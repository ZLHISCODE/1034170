VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFText"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'��������ֵ�ľֲ�����
Private mvarText As String
Private mvarBold As Boolean
Private mvarItalic As Boolean
Private mvarForeColor As ColorEnum
Private mvarFontSizeFixup As FontSizeEnum
Private mvarFontStyle As FontStyleEnum
Private mvarCharStyle As CharStyleEnum
Private mvarAnimType As AnimTypeEnum
Private mvarScaleX As Long
Private mvarEmbo As EmboEnum
Private mvarExpand As Long
Private mvarOutLine As Boolean
Private mvarShadow As Boolean
Private mvarStrike As Boolean
Private mvarStrikeD As Boolean
Private mvarSubChar As Boolean
Private mvarSuperChar As Boolean
Private mvarUnderLineType As UnderLineTypeEnum
Private mvarUnderLineColor As ColorEnum
Private mvarHided As Boolean
Private mvarWebHidden As Boolean
Private mvarCharBorder As Boolean
Private mvarCharShadow As Long
Private mvarCharPatForeColor As ColorEnum
Private mvarCharPatBackColor As ColorEnum
Private mvarPatType As PatTypeEnum
Private mvarHighlight As Boolean
Private mvarHighlightColor As ColorEnum
Private mvarAlignment As AlignTypeEnum
Private mvarListType As ListTypeEnum
Private mvarIsInTable As Boolean
Private mvarFontName As String              'Ŀǰֻ֧�֡����塱�͡����塱
Private mvarSpaceAfter As Long
Private mvarSpaceBefore As Long
Private mvarSpaceLine As Long
Private mvarSpaceFirstLine As Long
Private mvarFontSizeDefault As Long
Private mvarFontSize As Long

Public Property Let FontSize(ByVal vData As Long)
Attribute FontSize.VB_Description = "�����С������ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FontSize = 5
    mvarFontSize = vData
End Property


Public Property Get FontSize() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FontSize
    FontSize = mvarFontSize
End Property




Public Property Let FontSizeDefault(ByVal vData As Long)
Attribute FontSizeDefault.VB_Description = "Ĭ������ߴ硣һ����18��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FontSizeDefault = 5
    mvarFontSizeDefault = vData
End Property


Public Property Get FontSizeDefault() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FontSizeDefault
    FontSizeDefault = mvarFontSizeDefault
End Property




Public Property Let SpaceFirstLine(ByVal vData As Long)
Attribute SpaceFirstLine.VB_Description = "�����������ַ�����"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceFirstLine = 5
    mvarSpaceFirstLine = vData
End Property


Public Property Get SpaceFirstLine() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceFirstLine
    SpaceFirstLine = mvarSpaceFirstLine
End Property




Public Property Let SpaceLine(ByVal vData As Long)
Attribute SpaceLine.VB_Description = "�м�ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceLine = 5
    mvarSpaceLine = vData
End Property


Public Property Get SpaceLine() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceLine
    SpaceLine = mvarSpaceLine
End Property



Public Property Let SpaceBefore(ByVal vData As Long)
Attribute SpaceBefore.VB_Description = "�κ��ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceBefore = 5
    mvarSpaceBefore = vData
End Property


Public Property Get SpaceBefore() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceBefore
    SpaceBefore = mvarSpaceBefore
End Property



Public Property Let SpaceAfter(ByVal vData As Long)
Attribute SpaceAfter.VB_Description = "��ǰ��ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceAfter = 5
    mvarSpaceAfter = vData
End Property


Public Property Get SpaceAfter() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceAfter
    SpaceAfter = mvarSpaceAfter
End Property





Public Property Let FontName(ByVal vData As String)
Attribute FontName.VB_Description = "�������ơ�Ŀǰֻ֧������ͺ��塣"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FontName = 5
    mvarFontName = vData
End Property


Public Property Get FontName() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FontName
    FontName = mvarFontName
End Property




Public Property Let IsInTable(ByVal vData As Boolean)
Attribute IsInTable.VB_Description = "�Ƿ����������С�Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.IsInTable = 5
    mvarIsInTable = vData
End Property


Public Property Get IsInTable() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.IsInTable
    IsInTable = mvarIsInTable
End Property




Public Property Let ListType(ByVal vData As ListTypeEnum)
Attribute ListType.VB_Description = "��Ŀ�������͡�1��15��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ListType = 5
    mvarListType = vData
End Property


Public Property Get ListType() As ListTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ListType
    ListType = mvarListType
End Property


Public Property Let Alignment(ByVal vData As AlignTypeEnum)
Attribute Alignment.VB_Description = "�����ı����뷽ʽ��1��4��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Alignment = 5
    mvarAlignment = vData
End Property


Public Property Get Alignment() As AlignTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Alignment
    Alignment = mvarAlignment
End Property


Public Function GetTextRTF() As String
'��ȡ����RTF�ı�
    Dim strR As String
    
    Dim strAlignment As String
    Select Case mvarAlignment
    Case alm2_���ж���
        strAlignment = "\qc"
    Case alm3_���Ҷ���
        strAlignment = "\qr"
    Case alm4_���˶���
        strAlignment = "\qj"
    Case alm5_��ɢ����
        strAlignment = "\qd"
    Case Else
        strAlignment = "\ql"
    End Select
    
    Dim strIsInTable As String
    If mvarIsInTable Then strIsInTable = "\intbl"
    
    'Ĭ�������С����ദ��
    Dim strSB As String, strSA As String, strSL As String, strFL As String
    Dim strFontSize As String
    
    If mvarFontSize <> 0 Then   '�����С����Ȩ��FontSize > FontSizeFixup > FontSizeDefault
        strFontSize = "\fs" & mvarFontSize * 2
    Else
        If mvarFontSizeFixup <> fs00_Ĭ�� Then
            strFontSize = "\fs" & mvarFontSizeFixup * 2
        Else
            If mvarFontSizeDefault <> 0 Then
                strFontSize = "\fs" & mvarFontSizeDefault * 2
            Else
                strFontSize = "\fs18"
            End If
        End If
    End If
    If mvarSpaceBefore <> 0 Then
        strSB = "\sb" & mvarSpaceBefore
    Else    'Ĭ�϶�ǰ���Ϊ0
        If mvarIsInTable Then
            strSB = "\sb0"
        Else
            strSB = "\sb0"
        End If
    End If
    If mvarSpaceAfter <> 0 Then
        strSA = "\sa" & mvarSpaceAfter
    Else    'Ĭ�϶κ���Ϊ0
        If mvarIsInTable Then
            strSA = "\sa0"
        Else
            strSA = "\sa0"
        End If
    End If
    If mvarSpaceLine <> 0 Then
        strSL = "\sl" & mvarSpaceLine & "\slmult0"
    Else    'Ĭ���м��Ϊ30
        If mvarIsInTable Then
            strSL = "\sl0\slmult0"
        Else
            strSL = "\sl30\slmult0"
        End If
    End If
    If mvarSpaceFirstLine <> 0 Then
        strFL = "\fi" & mvarSpaceFirstLine * 210
    Else    'Ĭ����������Ϊ0
        If mvarIsInTable Then
            strFL = "\fi0"
        Else
            strFL = "\fi0"
        End If
    End If
    
    If ListType <> ltp00_�� Then
        Dim strListString As String
        Select Case ListType
        Case ltp09_����1
            strListString = "\'6c"
        Case ltp10_����2
            strListString = "\'6e"
        Case ltp11_����3
            strListString = "\'75"
        Case ltp12_����4
            strListString = "\'a8"
        Case ltp13_����5
            strListString = "\'fc"
        Case ltp14_����6
            strListString = "\'d8"
        Case ltp15_����7
            strListString = "\'b2"
        End Select
        
        If ListType <= ltp08_�ı�8 Then
            strR = "\pard\plain " & strAlignment & strIsInTable & " \fi-420\li3360\ri0\nowidctlpar\jclisttab\tx260\aspalpha\aspnum\faauto\ls1\ilvl" & CStr(mvarListType - 1) & "\adjustright\rin0\lin240\itap0 "
        Else
            strR = "\pard\plain " & strAlignment & strIsInTable & " \li0\ri0\nowidctlpar\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 " & _
                    "{\listtext\pard\plain\f10\fs18\kerning2 \loch\af10\dbch\af13\hich\f10 " & strListString & " \tab}" & _
                    "\pard\plain " & strAlignment & strIsInTable & " \fi-420\li3360\ri0\nowidctlpar\jclisttab\tx260\aspalpha\aspnum\faauto\ls2\ilvl" & CStr(mvarListType - 9) & "\adjustright\rin0\lin240\itap0"
        End If
    Else
        If mvarFontStyle = ft0_���� Then
            '�������
            
            strR = strR & "\pard\plain " & strAlignment & strIsInTable & strFL & "\li0\ri0" & strSB & strSA & strSL & "\nowidctlpar\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs18\lang1033\langfe2052\kerning2\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 "
        ElseIf mvarFontStyle = ft1_����1 Then
            strR = strR & "\pard\plain \s1" & strAlignment & strIsInTable & " \li0\ri0\sb340\sa330\sl578\slmult1\keep\keepn\nowidctlpar\aspalpha\aspnum\faauto\outlinelevel0\adjustright\rin0\lin0\itap0 \b\fs44\lang1033\langfe2052\kerning44\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 "
        ElseIf mvarFontStyle = ft2_����2 Then
            strR = strR & "\pard\plain \s2" & strAlignment & strIsInTable & " \li0\ri0\sb260\sa260\sl416\slmult1\keep\keepn\nowidctlpar\aspalpha\aspnum\faauto\outlinelevel1\adjustright\rin0\lin0\itap0 \b\fs32\lang1033\langfe2052\kerning2\loch\af1\hich\af1\dbch\af17\cgrid\langnp1033\langfenp2052 "
        ElseIf mvarFontStyle = ft3_����3 Then
            strR = strR & "\pard\plain \s3" & strAlignment & strIsInTable & " \li0\ri0\sb260\sa260\sl416\slmult1\keep\keepn\nowidctlpar\aspalpha\aspnum\faauto\outlinelevel2\adjustright\rin0\lin0\itap0 \b\fs32\lang1033\langfe2052\kerning2\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 "
        End If
    End If
    
    If mvarFontStyle = ft0_���� Then
        '���壺\fN�������������
        If mvarFontName = "����" Then strR = strR & "\f17"
        
        '����ߴ�
        strR = strR & strFontSize
        
        If mvarBold Then
            strR = strR & "\b"
        Else
            strR = strR & "\b0"
        End If
        If mvarItalic Then
            strR = strR & "\i"
        Else
            strR = strR & "\i0"
        End If
        If mvarForeColor <> clr00_Ĭ�� Then
            strR = strR & "\cf" & mvarForeColor
        End If
        If mvarCharStyle = 0 Then
            strR = strR & "\cs10"
        Else
            strR = strR & "\cs" & mvarCharStyle
        End If
        If mvarAnimType <> ani0_�� Then
            strR = strR & "\animtext" & mvarAnimType
        End If
        If mvarScaleX <> 100 Then
            strR = strR & "\charscalex" & mvarScaleX
        End If
        If mvarEmbo = emb1_����Ч�� Then
            strR = strR & "\embo"
        ElseIf mvarEmbo = emb2_���Ч�� Then
            strR = strR & "\impr"
        End If
        If mvarExpand <> 0 Then
            strR = strR & "\expnd" & mvarExpand & "\expndtw" & mvarExpand
        End If
        
        If mvarOutLine Then
            strR = strR & "\outl"
        Else
            strR = strR & "\outl0"
        End If
        If mvarShadow Then
            strR = strR & "\shad"
        Else
            strR = strR & "\shad0"
        End If
        If mvarStrike Then
            strR = strR & "\strike"
        Else
            strR = strR & "\strike0"
        End If
        If mvarStrikeD Then
            strR = strR & "\striked1"
        Else
            strR = strR & "\striked0"
        End If
        If mvarSubChar Then strR = strR & "\sub"
        If mvarSuperChar Then strR = strR & "\super"
        
        If mvarUnderLineType <> udl00_�� Then
            Select Case mvarUnderLineType
                Case udl01_������:
                    strR = strR & "\ul"
                Case udl02_����:
                    strR = strR & "\uld"
                Case udl03_�̻���:
                    strR = strR & "\uldash"
                Case udl04_�㻮��:
                    strR = strR & "\uldashd"
                Case udl05_˫�㻮��:
                    strR = strR & "\uldashdd"
                Case udl06_˫��:
                    strR = strR & "\uldb"
                Case udl07_���ز�����:
                    strR = strR & "\ulhwave"
                Case udl08_������:
                    strR = strR & "\ulldash"
                Case udl09_����:
                    strR = strR & "\ulth"
                Case udl10_�ֵ���:
                    strR = strR & "\ulthd"
                Case udl11_�ֶ̻���:
                    strR = strR & "\ulthdash"
                Case udl12_�ֵ㻮��:
                    strR = strR & "\ulthdashd"
                Case udl13_��˫�㻮��:
                    strR = strR & "\ulthdashdd"
                Case udl14_�ֳ�����:
                    strR = strR & "\ulthldash"
                Case udl15_˫������:
                    strR = strR & "\ululdbwave"
                Case udl16_���¼���:
                    strR = strR & "\ulw"
                Case udl17_������:
                    strR = strR & "\ulwave"
                Case Else
                    strR = strR & "\ulnone"
            End Select
            If mvarUnderLineColor <> clr00_Ĭ�� Then
                strR = strR & "\ulc" & mvarUnderLineColor
            Else
                strR = strR & "\ulc" & mvarUnderLineColor
            End If
        End If
        If mvarHided Then strR = strR & "\v"
        If mvarWebHidden Then strR = strR & "\webhidden"
        If mvarCharBorder Then strR = strR & "\chbrdr"
        If mvarPatType <> ptt00_�� Then
            Select Case mvarPatType
            Case ptt01_����:
                strR = strR & "\chbghoriz"
            Case ptt02_����:
                strR = strR & "\chbgvert"
            Case ptt03_��б��:
                strR = strR & "\chbgfdiag"
            Case ptt04_��б��:
                strR = strR & "\chbgbdiag"
            Case ptt05_������:
                strR = strR & "\chbgcross"
            Case ptt06_б������:
                strR = strR & "\chbgdcross"
            Case ptt07_�ֺ���:
                strR = strR & "\chbgdkhoriz"
            Case ptt08_������:
                strR = strR & "\chbgdkvert"
            Case ptt09_����б��:
                strR = strR & "\chbgdkfdiag"
            Case ptt10_����б��:
                strR = strR & "\chbgdkbdiag"
            Case ptt11_��������:
                strR = strR & "\chbgdkcross"
            Case ptt12_��б������:
                strR = strR & "\chbgdkdcross"
            End Select
            If mvarCharShadow Then strR = strR & "\chshdng" & mvarCharShadow '�ַ���Ӱ�ٷֱ�
            If mvarCharPatForeColor <> clr00_Ĭ�� Then
                strR = strR & "\chcfpat" & mvarCharPatForeColor
            End If
            If mvarCharPatBackColor <> clr00_Ĭ�� Then
                strR = strR & "\chcbpat" & mvarCharPatBackColor
            End If
        End If
        If mvarHighlight Then
            strR = strR & "\highlight" & mvarHighlightColor
        End If
'    Else
'        If mvarFontSizeFixup <> 0 Then strR = strR & "\fs" & mvarFontSizeFixup
    End If
    strR = strR & vbCrLf & StrToASC(mvarText)
    GetTextRTF = strR
End Function


Public Property Let HighlightColor(ByVal vData As ColorEnum)
Attribute HighlightColor.VB_Description = "������ʾ������ɫ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.HighlightColor = 5
    mvarHighlightColor = vData
End Property


Public Property Get HighlightColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.HighlightColor
    HighlightColor = mvarHighlightColor
End Property



Public Property Let Highlight(ByVal vData As Boolean)
Attribute Highlight.VB_Description = "������ʾ��Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Highlight = 5
    mvarHighlight = vData
End Property


Public Property Get Highlight() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Highlight
    Highlight = mvarHighlight
End Property



Public Property Let PatType(ByVal vData As PatTypeEnum)
Attribute PatType.VB_Description = "����ͼ�����ͣ�1��12��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PatType = 5
    mvarPatType = vData
End Property


Public Property Get PatType() As PatTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PatType
    PatType = mvarPatType
End Property



Public Property Let CharPatBackColor(ByVal vData As ColorEnum)
Attribute CharPatBackColor.VB_Description = "����ͼ��ǰ��ɫ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CharPatBackColor = 5
    mvarCharPatBackColor = vData
End Property


Public Property Get CharPatBackColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CharPatBackColor
    CharPatBackColor = mvarCharPatBackColor
End Property



Public Property Let CharPatForeColor(ByVal vData As ColorEnum)
Attribute CharPatForeColor.VB_Description = "����ͼ�����ɫ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CharPatForeColor = 5
    mvarCharPatForeColor = vData
End Property


Public Property Get CharPatForeColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CharPatForeColor
    CharPatForeColor = mvarCharPatForeColor
End Property



Public Property Let CharShadow(ByVal vData As Boolean)
Attribute CharShadow.VB_Description = "�ַ���Ӱ��Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CharShadow = 5
    mvarCharShadow = vData
End Property


Public Property Get CharShadow() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CharShadow
    CharShadow = mvarCharShadow
End Property



Public Property Let CharBorder(ByVal vData As Boolean)
Attribute CharBorder.VB_Description = "�ַ��߿�Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CharBorder = 5
    mvarCharBorder = vData
End Property


Public Property Get CharBorder() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CharBorder
    CharBorder = mvarCharBorder
End Property



Public Property Let WebHidden(ByVal vData As Boolean)
Attribute WebHidden.VB_Description = "Web��ʽ����ȫ���أ������ڱ���ʱ��Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.WebHidden = 5
    mvarWebHidden = vData
End Property


Public Property Get WebHidden() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.WebHidden
    WebHidden = mvarWebHidden
End Property



Public Property Let Hided(ByVal vData As Boolean)
Attribute Hided.VB_Description = "�����ı���Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Hided = 5
    mvarHided = vData
End Property


Public Property Get Hided() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Hided
    Hided = mvarHided
End Property



Public Property Let UnderLineColor(ByVal vData As ColorEnum)
Attribute UnderLineColor.VB_Description = "�»�����ɫ����ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.UnderLineColor = 5
    mvarUnderLineColor = vData
End Property


Public Property Get UnderLineColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.UnderLineColor
    UnderLineColor = mvarUnderLineColor
End Property



Public Property Let UnderLineType(ByVal vData As UnderLineTypeEnum)
Attribute UnderLineType.VB_Description = "�»������ͣ�1��17��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.UnderLineType = 5
    mvarUnderLineType = vData
End Property


Public Property Get UnderLineType() As UnderLineTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.UnderLineType
    UnderLineType = mvarUnderLineType
End Property



Public Property Let SuperChar(ByVal vData As Boolean)
Attribute SuperChar.VB_Description = "�ϱꡣĬ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SuperChar = 5
    mvarSuperChar = vData
End Property


Public Property Get SuperChar() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SuperChar
    SuperChar = mvarSuperChar
End Property



Public Property Let SubChar(ByVal vData As Boolean)
Attribute SubChar.VB_Description = "�±ꡣĬ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SubChar = 5
    mvarSubChar = vData
End Property


Public Property Get SubChar() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SubChar
    SubChar = mvarSubChar
End Property



Public Property Let StrikeD(ByVal vData As Boolean)
Attribute StrikeD.VB_Description = "˫ɾ���ߣ�Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.StrikeD = 5
    mvarStrikeD = vData
End Property


Public Property Get StrikeD() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.StrikeD
    StrikeD = mvarStrikeD
End Property



Public Property Let Strike(ByVal vData As Boolean)
Attribute Strike.VB_Description = "ɾ���ߣ�Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Strike = 5
    mvarStrike = vData
End Property


Public Property Get Strike() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Strike
    Strike = mvarStrike
End Property



Public Property Let Shadow(ByVal vData As Boolean)
Attribute Shadow.VB_Description = "��Ӱ��Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Shadow = 5
    mvarShadow = vData
End Property


Public Property Get Shadow() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Shadow
    Shadow = mvarShadow
End Property



Public Property Let OutLine(ByVal vData As Boolean)
Attribute OutLine.VB_Description = "�߿�Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.OutLine = 5
    mvarOutLine = vData
End Property


Public Property Get OutLine() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.OutLine
    OutLine = mvarOutLine
End Property



Public Property Let Expand(ByVal vData As Long)
Attribute Expand.VB_Description = "�ַ���࣬��ֵ��ʾѹ����Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Expand = 5
    mvarExpand = vData
End Property


Public Property Get Expand() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Expand
    Expand = mvarExpand
End Property



Public Property Let Embo(ByVal vData As EmboEnum)
Attribute Embo.VB_Description = "������Ч����0:��;1:����;2:��̡�Ĭ��:0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Embo = 5
    mvarEmbo = vData
End Property


Public Property Get Embo() As EmboEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Embo
    Embo = mvarEmbo
End Property



Public Property Let ScaleX(ByVal vData As Long)
Attribute ScaleX.VB_Description = "�������ű�����Ĭ�ϣ�100��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ScaleX = 5
    mvarScaleX = vData
End Property


Public Property Get ScaleX() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ScaleX
    ScaleX = mvarScaleX
End Property



Public Property Let AnimType(ByVal vData As AnimTypeEnum)
Attribute AnimType.VB_Description = "��̬�ı����ͣ�1��6��Ĭ�ϣ�0��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.AnimType = 5
    mvarAnimType = vData
End Property


Public Property Get AnimType() As AnimTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.AnimType
    AnimType = mvarAnimType
End Property



Public Property Let CharStyle(ByVal vData As CharStyleEnum)
Attribute CharStyle.VB_Description = "�ַ���ʽ0��1��2...��Ĭ�ϣ�-1��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.CharStyle = 5
    mvarCharStyle = vData
End Property


Public Property Get CharStyle() As CharStyleEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.CharStyle
    CharStyle = mvarCharStyle
End Property



Public Property Let FontStyle(ByVal vData As FontStyleEnum)
Attribute FontStyle.VB_Description = "������ʽ0��1��2��...��Ĭ�ϣ�-1"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FontStyle = 5
    mvarFontStyle = vData
End Property


Public Property Get FontStyle() As FontStyleEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FontStyle
    FontStyle = mvarFontStyle
End Property


Public Property Let FontSizeFixup(ByVal vData As FontSizeEnum)
Attribute FontSizeFixup.VB_Description = "����ߴ硣Ĭ�ϣ�21"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FontSizeFixup = 5
    mvarFontSizeFixup = vData
End Property


Public Property Get FontSizeFixup() As FontSizeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FontSizeFixup
    FontSizeFixup = mvarFontSizeFixup
End Property



Public Property Let ForeColor(ByVal vData As ColorEnum)
Attribute ForeColor.VB_Description = "ǰ��ɫ��Ĭ�ϣ�0"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ForeColor = 5
    mvarForeColor = vData
End Property


Public Property Get ForeColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ForeColor
    ForeColor = mvarForeColor
End Property



Public Property Let Italic(ByVal vData As Boolean)
Attribute Italic.VB_Description = "б�塣Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Italic = 5
    mvarItalic = vData
End Property


Public Property Get Italic() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Italic
    Italic = mvarItalic
End Property



Public Property Let Bold(ByVal vData As Boolean)
Attribute Bold.VB_Description = "���塣Ĭ�ϣ�False"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Bold = 5
    mvarBold = vData
End Property


Public Property Get Bold() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Bold
    Bold = mvarBold
End Property



Public Property Let Text(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Text = 5
    mvarText = vData
End Property


Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Text
    Text = mvarText
End Property


