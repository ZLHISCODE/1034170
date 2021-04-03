VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABL5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ�������� ABL5
'����:�����׶�
'����:�¶�
'�û�:���콭��ҽԺ,������ɽҽԺ,�����ж�Ժ
Private mblnStart As Boolean '��ʼ
Private mblnEnd As Boolean   '����
Private mstrData As String '����������

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
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim lng_Stx As String, lng_Etx As Long, str_Code As String
    Dim varIn As Variant '�����������
    Dim strLine As String
    Dim str�걾�� As String, str���� As String, str�����˱걾�ʿ� As String, strָ�� As String, strֵ As String
    Dim strҪ��ȡָ�� As String, varҪ��ȡָ�� As Variant
    Dim iStrPoint As Integer
    
    str�����˱걾�ʿ� = " |ѪҺ|0"
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    '-- ����Ҫ��ȡ��ָ��
    strҪ��ȡָ�� = "Temp;FIO2;tHb;pH(T);pCO2(T);pO2(T);AaDpO2,T;AaDpO2;pH;pCO2;pO2;HCO3-;sO2;tCO2;ABE;SBC;SBE"
    varҪ��ȡָ�� = Split(strҪ��ȡָ��, ";")
    
    
    If strSource Like "*Pc40R*Pc40R*Analyzer [#]*" Then
        strReserved = strSource
        lng_Stx = InStr(strReserved, "Pc40R")
        lng_Etx = InStr(strReserved, "Analyzer #") + 10
        str_Code = Mid(strReserved, lng_Stx, lng_Etx - lng_Stx)
        strReserved = Mid(strReserved, lng_Etx + 1)
        If strReserved = Space(Len(strReserved)) Then strReserved = ""
        
        varIn = Split(str_Code, vbNewLine)
        If UBound(varIn) <= 0 Then
            varIn = Split(str_Code, Chr(10))
        End If

        strLine = ""
        For iStrPoint = LBound(varIn) To UBound(varIn)

            strLine = varIn(iStrPoint)

            '--- �걾��
            If strLine Like "*W*#*" Then
                str�걾�� = Trim(Mid(strLine, InStr(strLine, "#") + 1))
                If str�걾�� = "" And mblnEnd Then Exit Sub
            End If

            '--- ����
            If strLine Like "Date*.*.*" Then
                str���� = Format(CDate(Format(Replace(Replace(strLine, "Date                   ", ""), ".", "-"), "dd-MM-yy")), "yyyy-MM-dd")
            End If


            '-- �˳�
            If InStr(strLine, "Analyzer #") > 0 And strResult = "" Then
                If strResult = "" And str�걾�� <> "" And strָ�� <> "" Then
                    If str���� = "" Then str���� = Format(Now(), "yyyy-MM-dd")
                    strResult = str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strָ��
                End If
                Call WriteLog("Abl5.Analyse ������ ", mstrData, strResult)
                mblnEnd = False: mblnStart = False: mstrData = ""
                Exit Sub
            End If

            '-- ������ָ�� "
            Dim i As Integer
            For i = LBound(varҪ��ȡָ��) To UBound(varҪ��ȡָ��)
                If InStr(strLine, varҪ��ȡָ��(i)) > 0 Then
                    If varҪ��ȡָ��(i) = "[CPc40R" Then
                        strָ�� = strָ�� & "|" & "tHb" & "|" & Trim(Replace(Replace(Mid(strLine, InStr(strLine, varҪ��ȡָ��(i)) + 20), varҪ��ȡָ��(i), ""), "g/L", ""))
                    Else
                        strֵ = Trim(Replace(Mid(strLine, 1, 28), varҪ��ȡָ��(i), ""))
                        strֵ = Trim(Replace(strֵ, "(T)", ""))
                        strָ�� = strָ�� & "|" & varҪ��ȡָ��(i) & "|" & IIf(strֵ = "", "����", strֵ)
                    End If
                    Exit For
                End If
            Next
        Next
        
    Else
        strReserved = strSource
    End If
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function


'------------------------------------------------------------------------------------------------------------
'    Dim iStrPoint As Integer '�ַ�ָ��
'    Dim strBuffer As String, strData As String
'    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
'
'    Dim varIn As Variant '�����������
'    Dim strTime As String
'
'    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
'    Dim strLine As String
'    Dim str�걾�� As String, str���� As String, str�����˱걾�ʿ� As String, strָ�� As String, strֵ As String
'    Dim strҪ��ȡָ�� As String, varҪ��ȡָ�� As Variant
'
'    On Error GoTo errHandle
'    str�����˱걾�ʿ� = " |ѪҺ|0"
'    '��ʼ������ֵ
'    strReserved = "": strResult = "": strCmd = ""
'    '-- ����Ҫ��ȡ��ָ��
'    strҪ��ȡָ�� = "Temp;FIO2;pH;pCO2;pO2;HCO3-;sO2;tCO2;ABE;SBC;SBE"
'    varҪ��ȡָ�� = Split(strҪ��ȡָ��, ";")
'
'    If strSource = "" Then Exit Sub
'    Call WriteLog("Abl5.Analyse ���� ", strSource, mstrData)
'    varIn = Split(strSource, vbNewLine)
'    If UBound(varIn) <= 0 Then
'        varIn = Split(strSource, Chr(10))
'    End If
'
'    If mblnStart = True And mblnEnd = False Then
'            'strReserved = strSource
'        mstrData = mstrData & strSource
'        Call WriteLog("Abl5.Analyse ��ȡ���� ", strSource, mstrData)
'        If mstrData Like "*Analyzer *" Then
'            mblnEnd = True
'            Call WriteLog("Abl5.Analyse �ҵ������� ", strSource, mstrData)
'        End If
'    End If
'
'
'    If strSource Like "*ABL5*" Then
'        mblnStart = True
'        'strReserved = strSource
'        Call WriteLog("Abl5.Analyse �ҵ���ʼ�� ", strSource, mstrData)
'        If strSource Like "*Analyzer *" Then
'            mblnEnd = True
'            Call WriteLog("Abl5.Analyse ȫ������,�ҵ������� ", strSource, mstrData)
'        End If
'        mstrData = strSource
'    End If
'
'    If mblnStart And mblnEnd And mstrData <> "" Then
'        Call WriteLog("Abl5.Analyse ��ʼ���� ", strSource, mstrData)
'    End If
'    Exit Sub
'errHandle:
'    Call WriteLog("Abl5.Analyse �������", strSource, strResult & vbNewLine & Err.Description)