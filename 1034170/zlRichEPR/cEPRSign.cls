VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'###############################################################################################
'##ģ �� ����cEPRSign.cls
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��11��24��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ��������ǩ�����󣨶������ͣ�8��
'##��    ����
'###############################################################################################

Option Explicit

Private mvarKey As Long                     '�ؼ��֣��洢�ڡ������ǡ���
Private mvarID As Double                      '���ڱ����IDֵ
Private mvar�ļ�ID As Long                  '�ļ�ID
Private mvar��ID As Double                    '��ID
Private mvar��ʼ�� As Long                  '��ʼ��
Private mvar��ֹ�� As Long                  '��ֹ��
Private mvar������� As Double                'ÿ��ǩ�������ı�˳���и�Ψһ��ţ�ȫ��Ψһ��
Private mvar���� As String                  'ǩ���û����������洢�ڡ������ı�����
Private mvarǩ����Ϣ As String              '����ǩ���ַ������洢�ڡ�Ҫ��ֵ����
Private mvarǰ������ As String              '�ƺ����߶ٺŵ�ǰ�����֣��洢�ڡ�Ҫ�����ơ���
Private mvarǩ������ As EPRSignLevelEnum    '0,1,2,3���洢�ڡ�Ҫ�ر�ʾ����
Private mvarǩ����ID As Long                '��¼ǩ����ID
Private mblnǩ��ͼƬ As Boolean             '�Ƿ�ʹ��ǩ��ͼƬ,�洢��"������̬"��

'����5��������ɡ��������ԡ����洢�ڡ��������ԡ��ֶ���
Private mvarǩ����ʽ As Long                '1������ǩ����2������ǩ��
Private mvarǩ������ As Long                '��1��ʼ��ţ��ɳ����Զ����亭�壨��ʾ��ͬ�汾��ǩ��������
                                            'ǩ������=1 ʹ��RTF.Text��Ϊ����ǩ��ԭ�ģ�������ԭ�ĵķ�ʽ�����⣬ǩ������λ�û���ʱ�����
                                            'ǩ������=2 ʹ��RTF.Text��Ϊ����ǩ��ԭ�� 10.28.170�汾�޸ļ� frmman.Get
                                            'δ��BY ZXH =3 ʹ�ñ������ݿ��������ı�������ǩ��Ҫ�أ�Ϊ����ǩ��ԭ��
Private mvar֤��ID As Long                  '����֤��ID��
Private mvar��ʾ��ǩ As Boolean             '�Ƿ���ʾ��ǩ��־
Private mvarǩ��ʱ�� As String              'ǩ��ʱ�䣬���ڻ���ʱ�ָ����ʱ��
Private mvar��ʾʱ�� As String              '�Ƿ���ʾǩ��ʱ�估��ʾ�ĸ�ʽ��,�մ���ʾ����ʾ
Private mvarǩ��Ҫ�� As String             '�Ƿ�Ϊǩ��Ҫ��Keyֵ
Private mvarʱ��� As String                'ʱ������洢�ڡ�Ҫ�ص�λ����
Private mvarʱ�����Ϣ As String            '����ǩ�����ص�ʱ�������֤������Ϣ(Base64)

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Double)
    mvarID = vData
End Property

Public Property Get ID() As Double
    ID = mvarID
End Property

Public Property Let �ļ�ID(ByVal vData As Long)
    mvar�ļ�ID = vData
End Property

Public Property Get �ļ�ID() As Long
    �ļ�ID = mvar�ļ�ID
End Property

Public Property Let ��ID(ByVal vData As Double)
    mvar��ID = vData
End Property

Public Property Get ��ID() As Double
    ��ID = mvar��ID
End Property

Public Property Let ��ʼ��(ByVal vData As Long)
    mvar��ʼ�� = vData
End Property

Public Property Get ��ʼ��() As Long
    ��ʼ�� = mvar��ʼ��
End Property

Public Property Let ��ֹ��(ByVal vData As Long)
    mvar��ֹ�� = vData
End Property

Public Property Get ��ֹ��() As Long
    ��ֹ�� = mvar��ֹ��
End Property

Public Property Let �������(ByVal vData As Double)
    mvar������� = vData
End Property

Public Property Get �������() As Double
    ������� = mvar�������
End Property

Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property

Public Property Get ����() As String
Attribute ����.VB_UserMemId = 0
    ���� = mvar����
End Property
Public Property Let ǩ����ID(ByVal vData As Long)
    mvarǩ����ID = vData
End Property

Public Property Get ǩ����ID() As Long
    ǩ����ID = mvarǩ����ID
End Property

Public Property Let ǩ����Ϣ(ByVal vData As String)
    mvarǩ����Ϣ = vData
End Property

Public Property Get ǩ����Ϣ() As String
    ǩ����Ϣ = mvarǩ����Ϣ
End Property

Public Property Let ǩ��Ҫ��(ByVal vData As String)
    mvarǩ��Ҫ�� = vData
End Property

Public Property Get ǩ��Ҫ��() As String
    ǩ��Ҫ�� = mvarǩ��Ҫ��
End Property

Public Property Let ǰ������(ByVal vData As String)
    mvarǰ������ = vData
End Property

Public Property Get ǰ������() As String
    ǰ������ = mvarǰ������
End Property

Public Property Let ǩ������(ByVal vData As EPRSignLevelEnum)
    mvarǩ������ = vData
End Property

Public Property Get ǩ������() As EPRSignLevelEnum
    ǩ������ = mvarǩ������
End Property

Public Property Let ǩ����ʽ(ByVal vData As Long)
    mvarǩ����ʽ = vData
End Property

Public Property Get ǩ����ʽ() As Long
    ǩ����ʽ = mvarǩ����ʽ
End Property

Public Property Let ǩ������(ByVal vData As Long)
    mvarǩ������ = vData
End Property

Public Property Get ǩ������() As Long
    ǩ������ = mvarǩ������
End Property

Public Property Let ֤��ID(ByVal vData As Long)
    mvar֤��ID = vData
End Property

Public Property Get ֤��ID() As Long
    ֤��ID = mvar֤��ID
End Property

Public Property Let ǩ��ʱ��(ByVal vData As String)
    mvarǩ��ʱ�� = vData
End Property

Public Property Get ǩ��ʱ��() As String
    ǩ��ʱ�� = mvarǩ��ʱ��
End Property

Public Property Let ��ʾʱ��(ByVal vData As String)
    mvar��ʾʱ�� = vData
End Property

Public Property Get ��ʾʱ��() As String
    ��ʾʱ�� = mvar��ʾʱ��
End Property

Public Property Let ʱ���(ByVal vData As String)
    mvarʱ��� = vData
End Property

Public Property Get ʱ���() As String
    ʱ��� = mvarʱ���
End Property

Public Property Let ʱ�����Ϣ(ByVal vData As String)
    mvarʱ�����Ϣ = vData
End Property

Public Property Get ʱ�����Ϣ() As String
    ʱ�����Ϣ = mvarʱ�����Ϣ
End Property

Public Property Let ��ʾ��ǩ(ByVal vData As Boolean)
    mvar��ʾ��ǩ = vData
End Property

Public Property Get ��ʾ��ǩ() As Boolean
    ��ʾ��ǩ = mvar��ʾ��ǩ
End Property

Public Property Let ��������(ByVal vData As String)
    Dim T As Variant
    T = Split(vData, ";")
    On Error Resume Next
    mvarǩ����ʽ = Val(T(0))
    mvarǩ������ = Val(T(1))
    mvar֤��ID = Val(T(2))
    mvar��ʾ��ǩ = (Val(T(3)) = 1)
    mvarǩ��ʱ�� = Format(T(4), "yyyy-mm-dd hh:mm:ss")
    mvar��ʾʱ�� = T(5)
    If UBound(T) >= 6 Then
        mvarǩ��Ҫ�� = CStr(T(6))
    Else
        mvarǩ��Ҫ�� = ""
    End If
    
    If UBound(T) >= 7 Then
        mvarʱ�����Ϣ = CStr(T(7))
    Else
        mvarʱ�����Ϣ = ""
    End If
End Property

Public Property Get ��������() As String
    �������� = CStr(mvarǩ����ʽ) & ";" & CStr(mvarǩ������) & ";" & _
        CStr(mvar֤��ID) & ";" & IIf(mvar��ʾ��ǩ, 1, 0) & ";" & _
        Format(mvarǩ��ʱ��, "yyyy-mm-dd hh:mm:ss") & ";" & mvar��ʾʱ�� & ";" & mvarǩ��Ҫ�� & ";" & mvarʱ�����Ϣ
End Property

'################################################################################################################
'## ���ܣ�  ����ǩ������
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cEPRSign
    '���ڿ������ݣ�����һ���������
    Dim cNode As New cEPRSign
    If blnWithKey Then cNode.Key = Me.Key
    cNode.ID = Me.ID
    cNode.�ļ�ID = Me.�ļ�ID
    cNode.��ID = Me.��ID
    cNode.��ʼ�� = Me.��ʼ��
    cNode.��ֹ�� = Me.��ֹ��
    cNode.������� = Me.�������
    cNode.���� = Me.����
    cNode.ǩ����Ϣ = Me.ǩ����Ϣ
    cNode.ǰ������ = Me.ǰ������
    cNode.ǩ������ = Me.ǩ������
    cNode.�������� = Me.��������
    cNode.ʱ��� = Me.ʱ���
    cNode.ǩ��Ҫ�� = Me.ǩ��Ҫ��
    cNode.ǩ����ID = Me.ǩ����ID
    cNode.ǩ��ͼƬ = Me.ǩ��ͼƬ
    Set Clone = cNode
End Function
Public Function FillElementMember(ByVal rs As ADODB.Recordset, ByVal strTableName As String) As Boolean
'����:�����ݼ�¼������Ա
    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    
    If NVL(rs!������, 0) <> 0 Then Me.Key = NVL(rs!������, 0)
    Me.ID = rs!ID
    Me.�ļ�ID = NVL(rs!�ļ�ID, 0)
    Me.��ID = NVL(rs!��ID, 0)
    Me.������� = NVL(rs!�������, 0)
    Me.���� = Split(NVL(rs!�����ı�, ";"), ";")(0)
    Me.ǩ����Ϣ = NVL(rs!Ҫ��ֵ��)
    Me.ǰ������ = NVL(rs!Ҫ������)
    Me.ǩ������ = NVL(rs!Ҫ�ر�ʾ)
    Me.�������� = NVL(rs!��������)
    Me.��ʼ�� = NVL(rs!��ʼ��, 1)
    Me.��ֹ�� = NVL(rs!��ֹ��, 0)
    Me.ʱ��� = NVL(rs!Ҫ�ص�λ)
    Me.ǩ��ͼƬ = NVL(rs!������̬, 0) = 1
    If UBound(Split(NVL(rs!�����ı�), ";")) > 0 Then
        Me.ǩ����ID = Val(Split(NVL(rs!�����ı�), ";")(1))
    End If
    
    FillElementMember = True
End Function
'################################################################################################################
'## ���ܣ�  �����ݿ�ֱ�ӻ�ȡǩ������
'##
'## ������  eEdtType    :��ǰ�༭��ʽ
'##         lngID       :���ݿ��еĶ���ID��
'##         blnWithKey  :�Ƿ����Keyֵһ��ָ�
'################################################################################################################
Public Function GetSignFromDB(ByVal lngID As Double, _
    Optional blnWithKey As Boolean = False, _
    Optional ByVal strTableName As String = "") As Boolean

    Dim rs As New ADODB.Recordset
    If strTableName = "" Then strTableName = "���Ӳ�������"
    
    gstrSQL = "Select ID, ��id, �ļ�id, ������, �������, ��������, �����ı�, Ҫ��ֵ��, Ҫ������, Ҫ�ر�ʾ, Ҫ�ص�λ, ������̬, ��ʼ��, ��ֹ��" & vbNewLine & _
                "From " & strTableName & vbNewLine & _
                "Where �������� = 8 And ID = [1]"

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRSign", lngID)
    If Not rs.EOF Then
        If blnWithKey And NVL(rs("������"), 0) <> 0 Then Me.Key = NVL(rs("������"), 0)
        Me.ID = rs("ID")
        Me.�ļ�ID = NVL(rs("�ļ�ID"), 0)
        Me.��ID = NVL(rs("��ID"), 0)
        Me.������� = NVL(rs("�������"), 0)
        Me.���� = Split(NVL(rs("�����ı�"), ";"), ";")(0)
        Me.ǩ����Ϣ = NVL(rs("Ҫ��ֵ��"))
        Me.ǰ������ = NVL(rs("Ҫ������"))
        Me.ǩ������ = NVL(rs("Ҫ�ر�ʾ"))
        Me.�������� = NVL(rs("��������"))
        Me.��ʼ�� = NVL(rs("��ʼ��"), 1)
        Me.��ֹ�� = NVL(rs("��ֹ��"), 0)
        Me.ʱ��� = NVL(rs("Ҫ�ص�λ"))
        Me.ǩ��ͼƬ = NVL(rs("������̬"), 0) = 1
        If UBound(Split(NVL(rs!�����ı�), ";")) > 0 Then
            Me.ǩ����ID = Val(Split(NVL(rs("�����ı�")), ";")(1))
        End If
        GetSignFromDB = True
    Else
        GetSignFromDB = False
    End If
End Function
Public Function InsertIntoEditor(ByRef edtThis As Object, Optional ByVal lngStartPos As Long = -1, Optional ByVal blnForceInsert As Boolean = False, Optional ByRef objDocument As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ� ����ָ������ǩ����Editor��
    '������ edtThis         :��ǰ�ı༭���ؼ�
    '       lngStartPos     :��ǰλ��
    '       blnForceInsert  :�Ƿ�ǿ�Ʋ��룬���޸�Ҫ��ʱ���ǰ�󶼱�������ʱ��Ҫǿ�Ʋ���
    '���أ�
    '******************************************************************************************************************
    
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim blnForce As Boolean
    Dim lngLen As Long
    Dim strTmp As String
    Dim StrText As String
    Dim intLoop As Integer
    Dim lESS As Long, lESE As Long, lEES As Long, lEEE As Long
    Dim blnNeeded As Boolean
    Dim blnFinded As Boolean
    
    '���ǩ������Ϊ�գ�Ҳ�˳�
    If Trim(Me.����) = "" Then Exit Function
    
    '�γ�ǩ����ʾ����
    strTmp = Format(Me.Key, "00000000") & ",0,0)"
    StrText = mvarǰ������ & Me.���� & IIf(mvar��ʾ��ǩ, "����ǩ��_____________", "")
    StrText = StrText & IIf(Trim(mvar��ʾʱ��) = "", "", "��" & Format(mvarǩ��ʱ��, mvar��ʾʱ��))
    lngLen = Len(StrText)
        
    '�ȼ���Ƿ��ж�Ӧ��ǩ��Ҫ�أ�����У���ǩ��ǩ��Ҫ����
    '------------------------------------------------------------------------------------------------------------------
    If Not (objDocument Is Nothing) Then
        For intLoop = 1 To objDocument.Elements.Count
            If objDocument.Elements(intLoop).�滻�� = 1 And Trim(objDocument.Elements(intLoop).�����ı�) = "" Then
                Select Case objDocument.Elements(intLoop).Ҫ������
                Case "����ҽʦǩ��"
                    If Me.ǩ������ = cprSL_���� Then Exit For
                Case "����ҽʦǩ��"
                    If Me.ǩ������ = cprSL_���� Then Exit For
                Case "����ҽʦǩ��"
                    If Me.ǩ������ = cprSL_���� Or Me.ǩ������ = cprSL_���� Then Exit For
                End Select
            End If
        Next

        If intLoop <= objDocument.Elements.Count And objDocument.Elements.Count > 0 Then
            blnFinded = FindKey(edtThis, "E", objDocument.Elements(intLoop).Key, lESS, lESE, lEES, lEEE, blnNeeded)
            If blnFinded Then
                lngStartPos = lESS
                objDocument.Elements(intLoop).�����ı� = StrText
                objDocument.Elements(intLoop).Refresh edtThis
                Me.ǩ��Ҫ�� = objDocument.Elements(intLoop).Key
            End If
        End If
    End If
    
    
    '���û�ж�Ӧ��ǩ��Ҫ�أ���ǩ����ǰ���λ�ô�
    '------------------------------------------------------------------------------------------------------------------
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    
    '��֤���ܲ���ؼ����ڲ�
    If bInKeys Then
        InsertIntoEditor = False
        Exit Function
    End If
        
    If blnForceInsert = False And edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then
        Exit Function
    End If
    
    With edtThis
        .Freeze
        blnForce = .ForceEdit
        .ForceEdit = True
        
        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        
        .Range(lngStartPos, lngStartPos).Text = "SS(" & strTmp & StrText & "SE(" & strTmp
            
        .Range(lngStartPos, lngStartPos + 32 + lngLen).Font.Protected = True
        .Range(lngStartPos, lngStartPos + 16).Font.Hidden = True
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Hidden = False
        'ɾ����
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Strikethrough = (Me.��ֹ�� > 0)
        '�»���
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = IIf(Me.���� = "", ELE_UNDERLINE, cprNone)
        'ǰ��ɫ
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.ForeColor = GetCharColor(Me.��ʼ��, Me.��ֹ�� + 1)
        '����ɫ
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.BackColor = ELE_BACKCOLOR
        .Range(lngStartPos + 16 + lngLen, lngStartPos + 32 + lngLen).Font.Hidden = True
        lngStartPos = lngStartPos + 32 + lngLen

                If Me.ǩ��Ҫ�� <> "" Then '��ǩ��Ҫ��ʱ��궨λ��Ҫ�������ΪҪ�ر�����
            lngStartPos = lngStartPos + 32 + lngLen
        End If

        .Range(lngStartPos, lngStartPos).Selected
        
        .ForceEdit = blnForce
        .UnFreeze
    End With

    InsertIntoEditor = True
End Function

'################################################################################################################
'## ���ܣ�  ˢ��ǩ�������ʾ
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'################################################################################################################
Public Function Refresh(ByRef edtThis As Object) As Boolean
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "S", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            Dim ParaFmt As New cParaFormat
            Set ParaFmt = edtThis.Range(lKSE, lKES).Para.GetParaFmt

            .Range(lKSS, lKEE) = ""
            Me.InsertIntoEditor edtThis, lKSS, True
            .Range(lKSE, lKSE + Len(Me.ǰ������ & Me.����)).Selected
            .Range(lKSE, lKES).Para.SetParaFmt ParaFmt
            .ForceEdit = blnForce
            .UnFreeze
        End With
        Refresh = True
    Else
        Refresh = False
    End If
End Function

'################################################################################################################
'## ���ܣ�  �ӱ༭����ɾ��ǩ����
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'################################################################################################################
Public Function DeleteFromEditor(ByRef edtThis As Object, Optional ByRef objDocument As Object) As Boolean
    '��������Ҫ�ر༭���
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "S", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True

            If Not (objDocument Is Nothing) Then
                If Me.ǩ��Ҫ�� <> "" Then
                    '��ǩ������ʾ�������صĶ�Ӧ�����ǩ��Ҫ�أ��Ա��ٴ�ǩ��ʱ�Զ���λ
                    
                    Dim lESS As Long
                    Dim lESE As Long
                    Dim lEES As Long
                    Dim lEEE As Long
                    
                    Dim blnNeeded As Boolean
                    Dim blnFinded As Boolean
        
                    blnFinded = FindKey(edtThis, "E", Me.ǩ��Ҫ��, lESS, lESE, lEES, lEEE, blnNeeded)
                
                    If blnFinded Then
                        .Range(lESE, lEES).Font.Hidden = False
                    End If
                End If
            End If
            
            .Range(lKSS, lKEE) = ""
            
            .ForceEdit = blnForce
            .UnFreeze
        End With
        
        If blnFinded Then
            objDocument.Elements("K" & Me.ǩ��Ҫ��).�����ı� = ""
            objDocument.Elements("K" & Me.ǩ��Ҫ��).Refresh edtThis
        End If
                    
        DeleteFromEditor = True
    Else
        DeleteFromEditor = False
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ȡָ����ǩ�����ϱ��浽���ݿ��SQL���
'##
'## ������  ֻ���ڡ����Ӳ�����¼���������������������ǩ��
'##         ArraySQL()          :IN/OUT��   SQL����
'##         lng���             :IN��       ������ţ�ͬ����ǩ������ͬһ��ţ�Ȼ���������д�����
'################################################################################################################
Public Function GetSaveSQL(ByRef ArraySQL() As String, _
    ByVal lng��� As Long) As Boolean
    
    Dim lngCount As Long
    Call CheckValidID     '���ID��Ч�ԣ�
    
    Dim i As Long, lng�д� As Long, lngEleKey As Long
    
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String
    
    gstrSQL = "ZL_���Ӳ�������_Update(" & Me.ID & "," & Me.�ļ�ID & "," & Me.��ʼ�� & "," & Me.��ֹ�� & "," & IIf(Me.��ID = 0, "NULL", Me.��ID) & "," & _
        lng��� & ",8," & Me.Key & ",1,'" & Me.�������� & "',NULL,'" & Me.���� & ";" & Me.ǩ����ID & "',0,NULL,NULL,NULL,NULL,NULL,Null," & _
        "'" & Me.ǰ������ & "',1,50,0,'" & Me.ʱ��� & "'," & Me.ǩ������ & "," & IIf(Me.ǩ��ͼƬ, 1, 0) & ",'" & Me.ǩ����Ϣ & "')"
    
    ArraySQL(lngCount) = gstrSQL
    GetSaveSQL = True
End Function
'################################################################################################################
'## ���ܣ�  �������ID�Ƿ���Ч
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š�
'################################################################################################################
Public Sub CheckValidID()
    If Me.ID = 0 Then Me.ID = GetEPRContentNextId
End Sub

Public Function IsFinished() As Boolean
    IsFinished = (Trim(Me.����) <> "")
End Function

Public Property Let ǩ��ͼƬ(ByRef vData As Boolean)
    mblnǩ��ͼƬ = vData
End Property

Public Property Get ǩ��ͼƬ() As Boolean
     ǩ��ͼƬ = mblnǩ��ͼƬ
End Property