VERSION 5.00
Begin VB.Form frmEPRSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��дǩ��"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6330
   Icon            =   "frmEPRSign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -270
      TabIndex        =   15
      Top             =   1785
      Width           =   6555
   End
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   3630
      TabIndex        =   25
      Top             =   1710
      Width           =   30
   End
   Begin VB.CheckBox chkOrgPic 
      Caption         =   "ǩ��ԭͼ"
      Height          =   195
      Left            =   3735
      TabIndex        =   22
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.TextBox txtHeight 
      Height          =   270
      Left            =   4965
      TabIndex        =   20
      Text            =   "50"
      Top             =   2625
      Width           =   390
   End
   Begin VB.CheckBox chkSignPic 
      Caption         =   "ǩ��ʹ��ͼƬ"
      Height          =   195
      Left            =   3735
      TabIndex        =   19
      Top             =   1965
      Width           =   1395
   End
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2610
      Width           =   2310
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -270
      TabIndex        =   18
      Top             =   510
      Width           =   6555
   End
   Begin VB.CheckBox chkPreText 
      Caption         =   "��ǩ��������Ϊǰ׺����(&P)"
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   1950
      Width           =   2565
   End
   Begin VB.CheckBox chkHandSign 
      Caption         =   "��ʾ��ǩλ��(&H)"
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   2257
      Width           =   1695
   End
   Begin VB.CheckBox chkEsign 
      Caption         =   "����ǩ��(&E)"
      Height          =   195
      Left            =   3105
      TabIndex        =   7
      Top             =   1013
      Width           =   1365
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1605
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1387
      Width           =   1365
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1605
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   1365
   End
   Begin VB.OptionButton optName 
      Caption         =   "ָ���û�(&U)"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1020
      Width           =   1320
   End
   Begin VB.OptionButton optName 
      Caption         =   "��ǰ�û�(&C)"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5130
      TabIndex        =   13
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3885
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   4110
   End
   Begin VB.PictureBox picǩ��ͼƬ 
      AutoRedraw      =   -1  'True
      Height          =   810
      Left            =   5415
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   21
      Top             =   2325
      Width           =   810
   End
   Begin VB.Label lblH 
      Caption         =   "��"
      Height          =   225
      Left            =   5055
      TabIndex        =   24
      Top             =   2415
      Width           =   180
   End
   Begin VB.Label lblWH 
      Height          =   225
      Left            =   3720
      TabIndex        =   23
      Top             =   3015
      Width           =   1605
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "ǩ��ʱ��(&T)"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   2670
      Width           =   990
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   240
      TabIndex        =   17
      Top             =   3255
      Width           =   5970
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ǩ��Ч��Ԥ��:"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   3030
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û�����(&P)"
      Height          =   180
      Left            =   510
      TabIndex        =   5
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1605
      TabIndex        =   14
      Top             =   660
      Width           =   360
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ǩ������(&L)"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmEPRSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object                 '������
Private Sign As cEPRSign                    'ǩ������

Private mlngPassType As Long                 '������֤����ϵͳ������ 0-���룻1�����֣�2�����߽Կ�
Private mblnOk As Boolean
Private mintSign As Integer                 '��ʾ����������ʾǩ�� ����ͬ�����
Private msSource As String                 '����ǩ����Դ�ַ���
Private mpicSign  As StdPicture
Private morgSign  As StdPicture             'ǩ��ԭʼͼ(��Ա��.ǩ��ͼƬ)


'################################################################################################################
'## ���ܣ�  ��ʾ������
'##
'## ������  edtThis     :IN     �༭���ؼ�
'##         fParent     :IN     ������
'##         strSource   :IN     ����ǩ����Դ�ַ��������ı�����ȡ��ȥ��ǩ����٣�
'################################################################################################################
Public Function ShowMe(ByRef edtThis As Editor, ByRef fParent As Object, _
                        ByVal sSource As String, ByRef picSign As StdPicture) As cEPRSign
    
    Dim bytFileKind As Byte    '�Ƿ�����
    bytFileKind = fParent.Document.EPRPatiRecInfo.��������
    Set mpicSign = Nothing
    Set morgSign = Nothing
    
    Dim lngStart As Long, strPreText As String
    mintSign = zlDatabase.GetPara("SignShow", glngSys, 1070, 0)
    
    Me.cboTime.Clear
    Me.cboTime.AddItem "����ʾ"
    Me.cboTime.AddItem Format(Now(), "yyyy-MM-dd hh:mm")
    Me.cboTime.AddItem Format(Now(), "yyyy��MM��dd�� hh:mm")
    
    lngStart = edtThis.Selection.StartPos
    strPreText = edtThis.Range(lngStart - 1, lngStart)
    If strPreText = ":" Or strPreText = "��" Then
        Me.chkPreText.Value = vbUnchecked
    Else
        Me.chkPreText.Value = vbChecked
    End If
    
    Set Sign = New cEPRSign
    Set frmParent = fParent
    msSource = sSource
    
    '����ǩ����������ʼ����ǩ������
    Select Case bytFileKind
    Case cpr������
        cmbLevel.AddItem "1 - ��ʿ"
        cmbLevel.AddItem "3 - ��ʿ��"
        cmbLevel.ListIndex = 0
        If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 1
    Case cpr���Ʊ���
        cmbLevel.AddItem "1 - ҽ��"
        cmbLevel.AddItem "2 - ����"
        cmbLevel.AddItem "3 - ����"
        cmbLevel.ListIndex = 0
        If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 1
        If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 2
    Case Else
        cmbLevel.AddItem "1 - ����ҽʦ"
        cmbLevel.AddItem "2 - ����ҽʦ"
        cmbLevel.AddItem "3 - ������ҽʦ"
        cmbLevel.AddItem "4 - ����ҽʦ"
        cmbLevel.ListIndex = 0
        If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 1
        If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 2
        If frmParent.Document.�û�ǩ������ >= cprSL_���� Then cmbLevel.ListIndex = 3
    End Select
    
    '��ȡ��ǰǩ����ʽ��ϵͳ����26��
    Dim lS As Long
    Select Case fParent.Document.EPRPatiRecInfo.��������
    Case cpr���ﲡ��
        lS = 1
    Case cprסԺ����
        lS = 2
    Case cpr���Ʊ���
        Select Case fParent.Document.EPRFileInfo.lngModule
            Case 1290, 1291, 1294
                lS = 7
            Case Else
                lS = 3
        End Select
        
    Case cpr������
        lS = 4
    Case Else
        Select Case fParent.Document.EPRPatiRecInfo.������Դ
        Case cprPF_����
            lS = 1
        Case cprPF_סԺ
            lS = 2
        Case Else
            lS = 2  '������סԺΪ׼
        End Select
    End Select
    
    mlngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), lS, 1)) '����,סԺ,ҽ��,����,ҩƷ,LIS,PACS (1111111),Ϊ��Ĭ�ϲ�������ģʽ
    If mlngPassType = 1 Then
        If gstrESign = "" Or (lS = 3 And gstrESign = "0") Then 'ҽ������վ��д����û�е���clsDockxx��,�����ˢ��"סԺ����"ҳ�棬����д�������clsDockInEPR�в���gstrESign = "0"
            gstrESign = getPassESign(3, fParent.Document.EPRPatiRecInfo.����ID)
        End If
        mlngPassType = Val(gstrESign)
    End If
    
    lblUserName.Caption = gstrUserName
    chkEsign.Value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", vbUnchecked)
    
    txtHeight.Text = zlDatabase.GetPara("ǩ��ͼƬ�߶�", glngSys, 1070, "50", Array(txtHeight), InStr(gstrPrivsEpr, "��������") > 0)
    txtHeight.ToolTipText = txtHeight.Text: txtHeight.Tag = txtHeight.Text
    chkOrgPic.Value = zlDatabase.GetPara("ǩ��ʹ��ԭͼ", glngSys, 1070, "1", Array(chkOrgPic, lblH), InStr(gstrPrivsEpr, "��������") > 0)
    chkOrgPic.Tag = chkOrgPic.Value
    
    chkSignPic.Value = zlDatabase.GetPara("ǩ��ʹ��ͼƬ", glngSys, 1070, "0", Array(chkSignPic), InStr(gstrPrivsEpr, "��������") > 0)
    chkSignPic.Tag = chkSignPic.Value
    
    chkHandSign.Value = zlDatabase.GetPara("��ʾ��ǩλ��", glngSys, 1070, "0", Array(chkHandSign), InStr(gstrPrivsEpr, "��������") > 0)
    chkHandSign.Tag = chkHandSign.Value
    
    chkPreText.Value = zlDatabase.GetPara("��ǩ��������Ϊǰ׺����", glngSys, 1070, "0", Array(chkPreText), InStr(gstrPrivsEpr, "��������") > 0)
    chkPreText.Tag = chkPreText.Value

    cboTime.ListIndex = zlDatabase.GetPara("ǩ��ʱ��", glngSys, 1070, "0", Array(cboTime), InStr(gstrPrivsEpr, "��������") > 0)
    cboTime.Tag = cboTime.ListIndex
   
    Call RefControls
    
    Me.Show vbModal, frmParent
    If mblnOk Then
        Set ShowMe = Sign
        If Sign.ǩ��ͼƬ Then
            Set picSign = mpicSign
        Else
            Set picSign = Nothing
        End If
    Else
        Set picSign = Nothing
        Set ShowMe = Nothing
    End If
    Set mpicSign = Nothing
    Set morgSign = Nothing
End Function

'################################################################################################################
'## ���ܣ�  ����ǩ�����ڲ�ǩ���鲢ˢ����ʾ����֤�����������ǩ����
'################################################################################################################
Private Function Validation() As Boolean
    Dim blnSpecify As Boolean, strSpecifySign, lngSpecifyId As Long, lngSpecifyLevel As Long
    Dim lngCertID As Long, strSign As String, strʱ��� As String, objSignPic As Object, strʱ��Base64 As String
    Dim rsTemp As ADODB.Recordset, l As Long
    
    On Error GoTo errHand
    Dim lngPatiId As Long, lngPageId As Long, bFileType As Byte
    lngPatiId = frmParent.Document.EPRPatiRecInfo.����ID
    lngPageId = frmParent.Document.EPRPatiRecInfo.��ҳID
    bFileType = frmParent.Document.EPRPatiRecInfo.��������

    If optName(1).Value Then  'ָ���ʺ�ǩ��
        blnSpecify = True
        txtName = Trim(txtName)
        txtPass = Trim(txtPass)
        
        If frmParent.Document.EPRPatiRecInfo.�������� = cprסԺ���� Or frmParent.Document.EPRPatiRecInfo.�������� = cpr���ﲡ�� Or frmParent.Document.EPRPatiRecInfo.�������� = cpr������ Then
            gstrSQL = "Select 1 From �ϻ���Ա�� A, ������Ա B Where a.�û��� = [1] And a.��Աid = b.��Աid And b.����id = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ǩ���û��뵱ǰ�û��Ƿ�ͬ����", UCase(txtName.Text), frmParent.Document.EPRPatiRecInfo.����ID)
            If rsTemp.EOF Then
                MsgBox "ָ��ǩ���û��뵱ǰ������Ա������ͬһ���ң���ֹ�����ÿ��Ҳ��˲�����", vbExclamation, gstrSysName: Exit Function
            End If
        End If
        
        If chkEsign.Value = vbUnchecked Then '����ǩ��
            If Trim(txtPass) = "" Then MsgBox "ָ���ʺ����벻��Ϊ�գ����飡", vbExclamation: Exit Function
            If Not OraDataOpen(txtName, IIf(UCase(txtName) = "SYS" Or UCase(txtName) = "SYSTEM", txtPass, TranPasswd(txtPass))) Then
                MsgBox "ָ���ʺ�/�������,�����������¼�ʺź����룡", vbInformation + vbOKOnly, gstrSysName: Exit Function
            End If
        End If
        
        gstrSQL = "Select ID,����,ǩ�� From ��Ա�� p Where ID=(Select ��ԱID From �ϻ���Ա�� Where �û���='" & UCase(txtName) & "')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Sign-GetUserInfo")
        If rsTemp.EOF Then MsgBox "ָ���ʺŲ����ڣ������������¼�ʺź�����!", vbInformation, gstrSysName: Exit Function
        
        If mintSign = 0 Then
            strSpecifySign = rsTemp!����
        Else
            strSpecifySign = NVL(rsTemp!ǩ��, rsTemp!����)         '��ʾǩ��
        End If
        lngSpecifyId = rsTemp.Fields("ID")   '�û�ID
        
        lngSpecifyLevel = GetUserSignLevel(lngSpecifyId, rsTemp!����, frmParent.Document.EPRPatiRecInfo.����ID, frmParent.Document.EPRPatiRecInfo.��ҳID) '��ȡָ���û���ǩ������
        If lngSpecifyLevel = cprSL_�հ� Then MsgBox "ָ���ʺ���δ����ǩ������������Ա�����е���Ƹ��ְ��", vbInformation, gstrSysName: Exit Function
        For l = 1 To frmParent.Document.Signs.Count
            If frmParent.Document.Signs(l).ǩ������ > lngSpecifyLevel Then
                MsgBox "��ǰ�������и��߼����ǩ��,��ǰǩ��������Ȩ��ǩ������", vbInformation, gstrSysName: Exit Function
            End If
        Next
    End If
    
    If Not (IIf(blnSpecify, lngSpecifyLevel, frmParent.Document.�û�ǩ������) >= Val(cmbLevel.Text)) Then '
        MsgBox "�û�ӵ�е�ǩ���������ѡ����ǩ������,������ѡ��ǩ������", vbInformation, gstrSysName: Exit Function
    End If

    If chkEsign.Value = vbChecked Then '����ǩ��,�ڴ˴����ж�ǩ��������г�ʼ�����˴��ڹرպ����ݱ��棬��ȡ��������Դ���ݽ���ǩ������ǩ�������ʼ��ʧ���򲻱���
        If gobjESign Is Nothing Then
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            If gobjESign.Initialize(gcnOracle, glngSys) = False Then Exit Function
        End If
        
        If gobjESign.CheckCertificate(IIf(blnSpecify, UCase(txtName), gstrDBUser)) = False Then Exit Function
        
        'ͣ�õģ�ֻ��������ǩ��
        If Not gobjESign.CertificateStoped(IIf(blnSpecify, strSpecifySign, gstrUserName)) Then
            strSign = gobjESign.signature(msSource, IIf(chkEsign.Value = vbChecked, IIf(blnSpecify, UCase(txtName), gstrDBUser), ""), lngCertID, strʱ���, objSignPic, strʱ��Base64, False, lngPatiId, IIf(bFileType = cpr���ﲡ��, 0, lngPageId), IIf(bFileType <> cpr���ﲡ��, 0, lngPageId)) '����ǩ����Ϣ,lngCertID����ǩ��ʹ�õ�֤���¼ID
            If strSign = "" Then MsgBox "����ǩ��ʧ�ܣ����ٴ�ǩ����", vbInformation + vbOKOnly, gstrSysName: Exit Function
        Else
            chkEsign.Value = vbUnchecked
        End If
    End If
    
    Sign.���� = IIf(blnSpecify, strSpecifySign, IIf(mintSign = 0, gstrUserName, gstrSignName))
    Sign.ǩ����ID = IIf(blnSpecify, lngSpecifyId, glngUserId)
    Sign.ǩ������ = Val(cmbLevel.Text)
    If Sign.ǩ������ > cprSL_���� Then Sign.ǩ������ = cprSL_����
    
    If Me.chkPreText.Value = vbChecked Then
        Sign.ǰ������ = Trim(Mid(Me.cmbLevel.Text, 4)) & "��"
    Else
        Sign.ǰ������ = ""
    End If
    Sign.��ʾ��ǩ = (chkHandSign.Value = vbChecked)
    Sign.ǩ����ʽ = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.ǩ��ʱ�� = zlDatabase.Currentdate()
    Select Case Me.cboTime.ListIndex
    Case 1: Sign.��ʾʱ�� = "yyyy-MM-dd hh:mm"
    Case 2: Sign.��ʾʱ�� = "yyyy��MM��dd�� hh:mm"
    Case Else: Sign.��ʾʱ�� = ""
    End Select
    
    'ǩ������=2 ʹ��RTF.Text��Ϊ����ǩ��ԭ�� ��cEPRSignע��
    Sign.ǩ������ = 2
    Sign.ǩ����Ϣ = strSign
    Sign.֤��ID = lngCertID
    Sign.ʱ��� = strʱ���
    Sign.ʱ�����Ϣ = strʱ��Base64
'    'ǩ������=3 ʹ�ñ������ݿ��������ı�������ǩ��Ҫ�أ�ǩ������,ͼƬ������Ӷ���Ϊ����ǩ��ԭ��
'    '����ǩ����Ϣ�ڱ�����������ǩ���󷵻ز���������
'    Sign.ǩ������ = 3
'    Sign.ǩ����Ϣ = IIf(chkEsign.Value = vbChecked, IIf(blnSpecify, UCase(txtName), gstrDBUser), "") '�������ǩ�����ȴ�ǩ���ʺţ���������ǩ���������,ǩ����ɺ����
'    Sign.֤��ID = 0
'    Sign.ʱ��� = ""
    
    If chkSignPic.Value = 1 And picǩ��ͼƬ.Picture.Handle <> 0 And chkSignPic.Visible Then
        Sign.ǩ��ͼƬ = True
        Set mpicSign = picǩ��ͼƬ.Picture
    ElseIf chkSignPic.Value = 1 And picǩ��ͼƬ.Picture.Handle = 0 And chkSignPic.Visible Then
        MsgBox IIf(optName(0).Value, "��ǰ", "ָ��") & "�ʺ�û�п��õ�ǩ��ͼ������ʹ��ͼƬǩ�����ܣ�����ϵ����Ա��", vbExclamation, gstrSysName
        Exit Function
    Else
        Sign.ǩ��ͼƬ = False
        Set mpicSign = Nothing
    End If
    
    If picǩ��ͼƬ.Tag <> "" Then 'ɾ����ʱͼƬ
        Kill picǩ��ͼƬ.Tag
        picǩ��ͼƬ.Tag = ""
    End If
    
    Validation = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��֤�û��������Ƿ���ȷ
'################################################################################################################
Private Function OraDataOpen(ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    Dim strSQL As String
    Dim strError As String
    Dim Cn As New ADODB.Connection
    
    On Error Resume Next
    Err = 0
    With Cn
        If .State = adStateOpen Then .Close
'        .Provider = "MSDataShape"
        .Open gcnOracle.ConnectionString, strUserName, strUserPwd
        If Err <> 0 Then
            OraDataOpen = False
            Exit Function
        End If
        .Close
    End With
    Set Cn = Nothing
    OraDataOpen = True
    Exit Function
errHand:
    Set Cn = Nothing
    OraDataOpen = False
    Err = 0
End Function

'################################################################################################################
'## ���ܣ�  ����ת������
'##
'## ������  strOld  :ԭ����
'##
'## ���أ�  �������ɵ�����
'################################################################################################################
Public Function TranPasswd(strOld As String) As String
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew
End Function

'################################################################################################################
'## ���ܣ�  ˢ�¿ؼ�
'################################################################################################################
Private Sub RefControls()
    If optName(0).Value Then
        txtName.Enabled = False
        txtPass.Enabled = False
        Select Case mlngPassType
        Case 0
            '����ǩ��
            chkEsign.Value = vbUnchecked
            chkEsign.Visible = False
        Case 1
            '1������
            chkEsign.Value = vbChecked
            chkEsign.Move txtPass.Left, txtPass.Top
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        End Select
    Else
        chkEsign.Enabled = True
        txtPass.Enabled = True
        txtName.Enabled = True
        Select Case mlngPassType
        Case 0
            '����ǩ��
            chkEsign.Value = vbUnchecked
            txtPass.Enabled = True
        Case 1
            '1������
            chkEsign.Value = vbChecked
            chkEsign.Move txtPass.Left, txtPass.Top
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        End Select
    End If
End Sub

Private Sub cboTime_Click()
     Call Preview
End Sub

Private Sub cboTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkEsign_Click()
    txtPass.Enabled = (chkEsign.Value = vbUnchecked)
    txtPass.Enabled = IIf(optName(0).Value, False, txtPass.Enabled)
    If txtPass.Enabled And txtPass.Visible Then
        txtPass.SetFocus
    End If
End Sub

Private Sub chkEsign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkHandSign_Click()
     Call Preview
End Sub

Private Sub chkHandSign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub chkOrgPic_Click()
    If chkOrgPic.Value = vbUnchecked Then
        txtHeight.Visible = True
        lblH.Visible = True
    Else
        txtHeight.Visible = False
        lblH.Visible = False
    End If
    DrawSignPicture
End Sub

Private Sub chkPreText_Click()
    Call Preview
End Sub

Private Sub chkPreText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub chkSignPic_Click()
Dim strFile As String, rsTemp As ADODB.Recordset
    If chkSignPic.Value = 1 Then
        picǩ��ͼƬ.Tag = ""
        picǩ��ͼƬ.ToolTipText = ""
        lblWH.Caption = "": lblWH.Visible = True: lblH.Visible = True
        chkOrgPic.Visible = True
        txtHeight.Visible = True
        picǩ��ͼƬ.Cls
        Set picǩ��ͼƬ.Picture = Nothing
        If optName(1).Value And Trim(txtName) = "" Then Exit Sub '���"ָ���ʺ�"
        gstrSQL = "Select b.ǩ��ͼƬ From �ϻ���Ա�� A, ��Ա�� B Where a.�û��� = '" & IIf(optName(0).Value, gstrDBUser, UCase(txtName)) & "' And a.��Աid = b.id"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ǩ��ͼƬ")
        If Not rsTemp.EOF Then
            strFile = zlDatabase.ReadPicture(rsTemp, "ǩ��ͼƬ")
            picǩ��ͼƬ.Tag = strFile
        End If
        strFile = picǩ��ͼƬ.Tag
        
        If strFile <> "" Then
            Set morgSign = LoadPicture(strFile)
            picǩ��ͼƬ.ToolTipText = "ԭʼ��С:" & CLng(picǩ��ͼƬ.ScaleX(morgSign.Width, vbHimetric, vbPixels)) & " X " & CLng(picǩ��ͼƬ.ScaleY(morgSign.Height, vbHimetric, vbPixels))
            DrawSignPicture
        End If
        chkPreText.Value = vbUnchecked: chkPreText.Enabled = False
        chkHandSign.Value = vbUnchecked: chkHandSign.Enabled = False
        cboTime.ListIndex = 0:          cboTime.Enabled = False
    Else
        Set morgSign = Nothing
        Set picǩ��ͼƬ.Picture = Nothing
        picǩ��ͼƬ.ToolTipText = ""
        lblWH.Caption = "": lblWH.Visible = False: lblH.Visible = False
        picǩ��ͼƬ.Cls
        chkOrgPic.Visible = False
        txtHeight.Visible = False: picǩ��ͼƬ.Move picǩ��ͼƬ.Left, picǩ��ͼƬ.Top, 810, 810
        Call DrawSignPicture
        chkPreText.Enabled = True
        chkHandSign.Enabled = True
        cboTime.Enabled = True
    End If
End Sub

Private Sub cmbLevel_Click()
    Call Preview
End Sub

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    If picǩ��ͼƬ.Tag <> "" Then 'ɾ����ʱͼƬ
        Kill picǩ��ͼƬ.Tag
        picǩ��ͼƬ.Tag = ""
    End If
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If InStr(gstrPrivsEpr, "��������") > 0 Then
        If chkHandSign.Tag <> chkHandSign.Value Then Call zlDatabase.SetPara("��ʾ��ǩλ��", chkHandSign.Value, glngSys, 1070)
        If chkPreText.Tag <> chkPreText.Value Then Call zlDatabase.SetPara("��ǩ��������Ϊǰ׺����", chkPreText.Value, glngSys, 1070)
        If cboTime.Tag <> cboTime.ListIndex Then Call zlDatabase.SetPara("ǩ��ʱ��", cboTime.ListIndex, glngSys, 1070)
        If chkSignPic.Tag <> chkSignPic.Value Then Call zlDatabase.SetPara("ǩ��ʹ��ͼƬ", chkSignPic.Value, glngSys, 1070)
        If chkOrgPic.Tag <> chkOrgPic.Value Then Call zlDatabase.SetPara("ǩ��ʹ��ԭͼ", chkOrgPic.Value, glngSys, 1070)
        If txtHeight.Tag <> txtHeight.Text Then Call zlDatabase.SetPara("ǩ��ͼƬ�߶�", txtHeight.Text, glngSys, 1070)
    End If
    If Validation Then
        mblnOk = True
        Unload Me
    End If
End Sub

Private Sub Preview()
    Dim StrText As String, bln��ǩ As Boolean, strǰ������ As String
    
    If Me.chkPreText.Value = vbChecked Then
        strǰ������ = Trim(Mid(Me.cmbLevel.Text, 4)) & "��"
    Else
        strǰ������ = ""
    End If
    bln��ǩ = (chkHandSign.Value = vbChecked)
    StrText = strǰ������ & IIf(mintSign = 0, gstrUserName, gstrSignName) & IIf(bln��ǩ, "����ǩ��_____________", "")
    If Me.cboTime.ListIndex > 0 Then
        StrText = StrText & "��" & Me.cboTime.Text
    End If
    lblPreview.Caption = StrText
    
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mlngPassType = 2 Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", chkEsign.Value
    End If
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cmbLevel", cmbLevel.ListIndex
    Set frmParent = Nothing
End Sub

Private Sub Label1_Click()

End Sub

Private Sub optName_Click(Index As Integer)
    Call RefControls
    Call chkSignPic_Click
    If Index = 1 Then
        If txtName.Enabled And txtName.Visible Then txtName.SetFocus
    End If
End Sub

Private Sub optName_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub optPassType_Click(Index As Integer)
    If Index = 1 Then
        txtPass.Enabled = True
        If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus
    Else
        txtPass.Enabled = False
    End If
End Sub

Private Sub optPassType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtHeight_Change()
    On Error Resume Next
    DrawSignPicture
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkEsign.Value = vbUnchecked Then
            If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus: Call Preview: Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
        End If
    End If
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtNames_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtName_LostFocus()
    Call chkSignPic_Click
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < 32 Or KeyAscii > 126 Then KeyAscii = 0
    If InStr("""@\ ", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Function getPassType(ByVal lngFileKind As Long, ByVal lngPatientSource As Long, ByVal lngDeptId As Long, ByVal lngModule As Long) As Long
Dim rsTemp As New ADODB.Recordset, lS As Long
    On Error GoTo errHand
    '0-����ҽ���Ͳ�����1-סԺҽ��ҽ���Ͳ�����2-סԺ��ʿҽ����3-ҽ��ҽ���ͱ��棻4-�����¼�ͻ�������5-ҩƷ��ҩ��6-LIS;7-PACS
    Select Case lngFileKind
        Case cpr���ﲡ��
            lS = 0
        Case cprסԺ����
            lS = 1
        Case cpr���Ʊ���
            Select Case lngModule
                Case 1290, 1291, 1294
                    lS = 7
                Case Else
                    lS = 3
            End Select
        Case cpr������
            lS = 4
        Case Else
            Select Case lngModule
            Case cprPF_����
                lS = 0
            Case cprPF_סԺ
                lS = 1
            Case Else
                lS = 1  '������סԺΪ׼
            End Select
    End Select
    
    gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) as ���� From Dual "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ǩ�����Ʋ���", lS, lngDeptId)
    If rsTemp.EOF Then
        getPassType = 1
    Else
        getPassType = rsTemp!����
    End If

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub DrawSignPicture()
    On Error Resume Next
    If Not morgSign Is Nothing Then
        If chkOrgPic.Value = vbChecked Then
            Set picǩ��ͼƬ.Picture = morgSign
            picǩ��ͼƬ.Appearance = 0: picǩ��ͼƬ.BorderStyle = 0
            If picǩ��ͼƬ.Width <> 810 Then picǩ��ͼƬ.Move picǩ��ͼƬ.Left, picǩ��ͼƬ.Top, 810, 810
            picǩ��ͼƬ.PaintPicture picǩ��ͼƬ.Picture, 0, 0, picǩ��ͼƬ.ScaleX(picǩ��ͼƬ.Width, vbTwips, vbPixels), picǩ��ͼƬ.ScaleY(picǩ��ͼƬ.Height, vbTwips, vbPixels)
            lblWH.Caption = CLng(picǩ��ͼƬ.ScaleX(morgSign.Width, vbHimetric, vbPixels)) & " X " & CLng(picǩ��ͼƬ.ScaleY(morgSign.Height, vbHimetric, vbPixels)) & " Pixels"
        Else
            Dim lngWidth As Long
            lngWidth = CLng(txtHeight.Text * (morgSign.Width / morgSign.Height))
            picǩ��ͼƬ.Appearance = 0: picǩ��ͼƬ.BorderStyle = 0
            picǩ��ͼƬ.Move picǩ��ͼƬ.Left, picǩ��ͼƬ.Top, picǩ��ͼƬ.ScaleX(lngWidth, vbPixels, vbTwips), picǩ��ͼƬ.ScaleY(txtHeight.Text, vbPixels, vbTwips)
            picǩ��ͼƬ.PaintPicture morgSign, 0, 0, lngWidth, txtHeight.Text
            Set picǩ��ͼƬ.Picture = picǩ��ͼƬ.Image
            lblWH.Caption = lngWidth & " X " & txtHeight.Text & " Pixels"
        End If
    End If
    Err.Clear
End Sub
