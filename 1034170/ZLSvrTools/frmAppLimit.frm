VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAppLimit 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   16545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAppLimit.frx":0000
   ScaleHeight     =   10590
   ScaleWidth      =   16545
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   0
      ScaleHeight     =   9615
      ScaleWidth      =   15735
      TabIndex        =   10
      Top             =   600
      Width           =   15735
      Begin MSComctlLib.ImageList img16 
         Left            =   5280
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppLimit.frx":803A
               Key             =   "unCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAppLimit.frx":85D4
               Key             =   "Check"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox pctOpt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3495
         ScaleWidth      =   14175
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   6000
         Width           =   14175
         Begin VB.TextBox txtTip 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "????"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2775
            Left            =   1680
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "frmAppLimit.frx":8B6E
            Top             =   480
            Width           =   5535
         End
         Begin VB.PictureBox pctOption 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   8160
            ScaleHeight     =   2385
            ScaleWidth      =   5505
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   360
            Width           =   5535
            Begin VB.CommandButton cmdCancel 
               Caption         =   "????(&C)"
               Height          =   350
               Left            =   3900
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   1920
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.ComboBox txtApp 
               Height          =   300
               Left            =   1020
               TabIndex        =   2
               Top             =   630
               Width           =   2775
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   3
               Left            =   2640
               MaxLength       =   3
               TabIndex        =   6
               Tag             =   "IP????"
               Top             =   1125
               Width           =   315
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   0
               Left            =   1080
               MaxLength       =   3
               TabIndex        =   3
               Tag             =   "IP????"
               Top             =   1125
               Width           =   315
            End
            Begin VB.TextBox txtIP 
               Height          =   300
               Index           =   2
               Left            =   3420
               MaxLength       =   3
               TabIndex        =   7
               Tag             =   "IP"
               Top             =   1095
               Width           =   390
            End
            Begin VB.TextBox txtUser 
               Height          =   350
               Left            =   1020
               TabIndex        =   1
               Top             =   120
               Width           =   2415
            End
            Begin VB.CommandButton cmdStop 
               Caption         =   "????(&S)"
               Height          =   350
               Left            =   3900
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   1560
               Width           =   1455
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "????????(&M)"
               Height          =   350
               Left            =   3900
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   600
               Width           =   1455
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "????????(&A)"
               Height          =   350
               Left            =   3900
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   120
               Width           =   1455
            End
            Begin VB.CommandButton cmdDel 
               Caption         =   "????????(&D)"
               Height          =   350
               Left            =   3900
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1455
            End
            Begin VB.CommandButton cmdMore 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "????"
                  Size            =   6.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   3420
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   2160
               MaxLength       =   3
               TabIndex        =   5
               Tag             =   "IP????"
               Top             =   1125
               Width           =   315
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   1560
               MaxLength       =   3
               TabIndex        =   4
               Tag             =   "IP????"
               Top             =   1125
               Width           =   315
            End
            Begin VB.TextBox txtDesc 
               Height          =   495
               Left            =   1020
               MaxLength       =   99
               MultiLine       =   -1  'True
               TabIndex        =   8
               Top             =   1560
               Width           =   2775
            End
            Begin VB.TextBox txtIP 
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   1020
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   18
               TabStop         =   0   'False
               Tag             =   "IP"
               Text            =   "    ??    ??    ??"
               Top             =   1080
               Width           =   1965
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "????(&O)"
               Height          =   350
               Left            =   3900
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   1560
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label lblIP 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "IP??"
               Height          =   180
               Left            =   480
               TabIndex        =   27
               Top             =   1170
               Width           =   360
            End
            Begin VB.Label lblUser 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "??????"
               Height          =   180
               Left            =   300
               TabIndex        =   26
               Top             =   210
               Width           =   540
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "-"
               Height          =   180
               Index           =   11
               Left            =   3180
               TabIndex        =   25
               Top             =   1155
               Width           =   90
            End
            Begin VB.Label lblDesc 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "????"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   480
               TabIndex        =   24
               Top             =   1560
               Width           =   360
            End
            Begin VB.Label lblApp 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "????????"
               Height          =   180
               Left            =   120
               TabIndex        =   23
               Top             =   690
               Width           =   720
            End
         End
         Begin VB.Image imgIcon 
            Appearance      =   0  'Flat
            Height          =   1155
            Left            =   480
            Picture         =   "frmAppLimit.frx":8C24
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1125
         End
      End
      Begin VB.PictureBox pctPer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   0
         ScaleHeight     =   4695
         ScaleWidth      =   15135
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   15135
         Begin VB.TextBox txtFind 
            ForeColor       =   &H80000010&
            Height          =   350
            Left            =   960
            TabIndex        =   9
            Text            =   "??????????????????????????????????????"
            Top             =   80
            Width           =   3855
         End
         Begin VB.TextBox txtStop 
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   120
            Width           =   90
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfPer 
            Height          =   3255
            Left            =   120
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   600
            Width           =   7215
            _cx             =   12726
            _cy             =   5741
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "????"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   1
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblFind 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "????"
            Height          =   180
            Left            =   120
            TabIndex        =   14
            Top             =   165
            Width           =   360
         End
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "????????????"
      BeginProperty Font 
         Name            =   "????"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmAppLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsUsers As ADODB.Recordset  '????????????????,??????????????????????????????????
Private mrsApps As ADODB.Recordset  '????????????????,??????????????????????????????????
Private mstrApp As String   '??????????????
Private mblnKeypress As Boolean

Private Enum Color
    tipColor = &H80000010
    txtColor = &H80000012
End Enum
Private Const conCol = "????,250,1;????????,1200,1;??????,1200,1;????,1200,1;????IP,1200,1;????IP,1200,1;????,500,1;????,1200,1"

Private Sub ChangeCmdVisiable(ByVal blnIsAdd)
    '??????????????
    cmdAdd.Visible = Not blnIsAdd
    cmdDel.Visible = Not blnIsAdd
    cmdEdit.Visible = Not blnIsAdd
    cmdStop.Visible = Not blnIsAdd
    cmdSave.Visible = blnIsAdd
    cmdCancel.Visible = blnIsAdd
    
    '??????
    If blnIsAdd Then
        txtUser.Text = ""
        txtApp.Text = ""
        txtbeforeIp(0).Text = ""
        txtbeforeIp(1).Text = ""
        txtbeforeIp(2).Text = ""
        txtbeforeIp(3).Text = ""
        txtIP(2).Text = ""
        txtDesc.Text = ""
    Else
        With vsfPer
            vsfPer_AfterRowColChange 0, 0, .Row, .Col
        End With
    End If
    
    '????????????
    cmdMore.Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtUser.Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtApp.Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(0).Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(1).Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(2).Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(3).Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtIP(2).Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
    txtDesc.Enabled = Val(vsfPer.Row) > 0 Or cmdSave.Visible
End Sub

Private Sub cmdAdd_Click()
    ChangeCmdVisiable True
End Sub

Private Sub cmdCancel_Click()
    ChangeCmdVisiable False
End Sub
Private Sub cmdSave_Click()
    Dim strTmp As String, i As Integer
    Dim strStartIP As String, strEndIp As String
    Dim strUser As String, strApp As String, strDesc As String
    Dim varUsers As Variant
    
    On Error GoTo errh
    '????????
    Call GetDataFromCard(strUser, strApp, strStartIP, strEndIp, strDesc)

    If mrsUsers Is Nothing Then
        Set mrsUsers = LoadUsers
    End If
    
    strTmp = CheckExist("??????", strUser, mrsUsers)
    If strTmp <> "" Then
        MsgBox "??????????:" & strTmp & "??????,??????????????????", , "????"
        Exit Sub
    End If
    
    '??????????
    strTmp = CheckPerOnly(strApp, strUser)
    If strTmp <> "" Then
        MsgBox "????????????????????????????????????,??????????????" & vbNewLine & strTmp, , "????"
        Exit Sub
    End If
    
    strTmp = ValidateTxt
    If strTmp <> "" Then
        frmMDIMain.stbThis.Panels(2).Text = strTmp
        Exit Sub
    End If
    
    '????????
    gcnOracle.BeginTrans
    Screen.MousePointer = vbArrowHourglass
    If Len(strUser) < 2000 Then
        Call ExecuteProcedure("zltools.Zl_Zlapppermission_Edit(1,'" & strApp & "','" & strUser & "','" & strStartIP & "','" & strEndIp & "',1,'" & strDesc & "','','')", Me.Caption)
    Else
        varUsers = Str2Array(strUser, ",", 2000)
        For i = 0 To UBound(varUsers)
            Call ExecuteProcedure("zltools.Zl_Zlapppermission_Edit(1,'" & strApp & "','" & varUsers(i) & "','" & strStartIP & "','" & strEndIp & "',1,'" & strDesc & "','','')", Me.Caption)
        Next
    End If
    gcnOracle.CommitTrans
    Screen.MousePointer = vbDefault
    
    With vsfPer
        .Redraw = flexRDNone
        Call LoadAppPermission
    End With
    frmMDIMain.stbThis.Panels(2).Text = "??????????????"
    Exit Sub
errh:
    Screen.MousePointer = vbDefault
    frmMDIMain.stbThis.Panels(2).Text = ""
    
    If InStr(1, UCase(err.Description), "ORA") Then '??????????,??????????,????????,????????????
        MsgBox "????????????????????" & vbNewLine & err.Description
        gcnOracle.RollbackTrans
    Else
        frmMDIMain.stbThis.Panels(2).Text = "????????????????????" & vbNewLine & err.Description
    End If
End Sub

Private Sub cmdDel_Click()
    Dim varApps As Variant
    Dim i As Integer, intSRow As Integer
    
    mstrApp = GetSelectData
    '??????????????2000??,????????????,????2000??,??????????????????
    Screen.MousePointer = vbArrowHourglass
    gcnOracle.BeginTrans
    If Len(mstrApp) < 2000 Then
        Call ExecuteProcedure("Zl_ZlApppermission_Delete('" & mstrApp & "')", Me.Caption)
    Else
        varApps = Str2Array(mstrApp, ",", 2000)
        For i = 0 To UBound(varApps)
            Call ExecuteProcedure("Zl_ZlApppermission_Delete('" & varApps(i) & "')", Me.Caption)
        Next
    End If
    gcnOracle.CommitTrans
    Screen.MousePointer = vbDefault
    
    With vsfPer
        intSRow = .Row
        .Redraw = flexRDNone
        
        '??????????????
        For i = .FixedRows To .Rows - .FixedRows
            If i > .Rows - .FixedRows Or .Rows = .FixedRows Then
                Exit For
            End If
            If InstrEx(mstrApp, .TextMatrix(i, .ColIndex("????????")) & ":" & .TextMatrix(i, .ColIndex("??????"))) Then
                .RemoveItem (i)
                i = i - 1
            End If
        Next
        .Redraw = flexRDDirect
        
        '??????????
        If intSRow > .Rows - .FixedRows Then
            .Select .Rows - .FixedRows, 0
        Else
            .Select intSRow, 0
        End If
        .TopRow = .Row
    End With
    mstrApp = GetSelectData
    frmMDIMain.stbThis.Panels(2).Text = "??????????????"
    Exit Sub
    
errh:
    Screen.MousePointer = vbDefault
    
    If InStr(1, UCase(err.Description), "ORA") Then '??????????,??????????,????????,????????????
        MsgBox "????????????????????" & vbNewLine & err.Description
        gcnOracle.RollbackTrans
    Else
        frmMDIMain.stbThis.Panels(2).Text = "????????????????????" & vbNewLine & err.Description
    End If
End Sub

Private Sub cmdEdit_Click()
    EditPermission
End Sub

Private Sub CmdStop_Click()
    EditPermission ("????")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 And Shift = 2 Then
        cmdMore_Click
    End If
End Sub

Private Sub cmdMore_Click()
    Dim strUsers As String
    Dim p As POINTAPI
    Dim rstmp As ADODB.Recordset
    Dim strTmp() As String, i As Integer
    
    p.x = (pctOption.Left + cmdMore.Left + cmdMore.Width - FindUserWidth) / Screen.TwipsPerPixelX
    p.y = (pctOpt.Top + pctContainer.Top - FindUserHeight) / Screen.TwipsPerPixelY
    ClientToScreen Me.hwnd, p
    
    If mrsUsers Is Nothing Then
        Set mrsUsers = LoadUsers
    End If
    
    strUsers = frmFindUser.ShowMe(Me, mrsUsers, Trim(txtUser.Text), p.x * Screen.TwipsPerPixelX, p.y * Screen.TwipsPerPixelY)
    txtUser.Text = strUsers
    
End Sub

Private Sub Form_Load()
    Call InitTable(vsfPer, conCol)
    Call LoadAppPermission
    Call ChangeCmdVisiable(False)
    '????????????????
    With vsfPer
        .ColSort(-1) = flexSortCustom
        .ColSort(0) = flexSortNone
        .ColDataType(0) = flexDTBoolean
        .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        .Cell(flexcpText, 0, 0) = ""
        .Cell(flexcpPictureAlignment, 0, 0) = flexPicAlignCenterCenter
        .Editable = flexEDKbdMouse
    End With
    
    FindApp
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    pctContainer.Width = Me.ScaleWidth
    pctContainer.Height = Me.ScaleHeight - pctContainer.Top
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set mrsApps = Nothing
    Set mrsUsers = Nothing
End Sub

Private Sub pctContainer_Resize()
    On Error Resume Next
    
    pctPer.Width = pctContainer.Width
    pctPer.Height = pctContainer.Height - pctOpt.Height
    
    pctOpt.Width = pctContainer.Width
    pctOpt.Top = pctPer.Top + pctPer.Height
End Sub

Private Sub pctOpt_Resize()
    On Error Resume Next
    
    pctOption.Left = pctOpt.Width - pctOption.Width - 120
End Sub

Private Sub pctPer_Resize()
    On Error Resume Next
    
    vsfPer.Width = pctPer.ScaleWidth - 240
    vsfPer.Height = pctPer.ScaleHeight - vsfPer.Top - 30
    
    lblFind.Left = vsfPer.Left
    txtFind.Left = lblFind.Left + lblFind.Width + 45
End Sub

Private Sub txtApp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = "??????????????????????????????????????" Then
        txtFind.Text = ""
        txtFind.ForeColor = txtColor
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intRow As Integer
    
    If KeyAscii = 13 Then
        If Trim(txtFind.Text) = "" Then
            '??????????????????,??????
            LoadAppPermission
        Else
            Call GetRowPos(vsfPer, txtFind.Text, "??????,????,????????")
        End If
    End If
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "??????????????????????????????????????"
        txtFind.ForeColor = tipColor
    End If
End Sub

Private Sub txtIp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtUser_LostFocus()
    Call txtUser_KeyPress(13)
End Sub

Private Sub txtUser_Validate(Cancel As Boolean)
     If mblnKeypress Then
        mblnKeypress = False
    Else
        Call txtUser_KeyPress(13)
    End If
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
    Dim strTmp As String, intPos As Integer
    
    If KeyAscii = 13 Then    '????????
        strTmp = Trim(txtUser.Text)
        intPos = InStrRev(strTmp, ",")
        strTmp = UCase(Mid(strTmp, intPos + 1))
        If strTmp = "" Then Exit Sub
        strTmp = Left(Trim(txtUser.Text), intPos) & FindUser(strTmp)
        
        txtUser.Text = strTmp
        txtUser.SelStart = Len(strTmp)
    End If
    
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub LoadAppPermission()
'????:????????????????????
    Dim strSQL As String, rstmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errh
    
    strSQL = "Select a.????????, a.??????, c.????, a.????ip, a.????ip, decode(a.????,1,'??????','??????') ????, a.????" & vbNewLine & _
                    "From Zlapppermission A, ?????????? B, ?????? C" & vbNewLine & _
                    "Where a.?????? = b.??????(+) And b.????id = c.Id(+)"
                    
    Set rstmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "LoadAppLimit")
    Set mrsApps = rstmp
                            
    With vsfPer
        If rstmp.RecordCount = 0 Then
             .Rows = .FixedRows
            Exit Sub
        End If

        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = rstmp.RecordCount + .FixedRows
        
        i = .FixedRows
        Do While Not rstmp.EOF
            .RowData(i) = "" & rstmp!???????? & ":" & rstmp!??????
            .TextMatrix(i, 0) = "0"
            .TextMatrix(i, .ColIndex("????????")) = rstmp!???????? & ""
            .TextMatrix(i, .ColIndex("??????")) = rstmp!?????? & ""
            .TextMatrix(i, .ColIndex("????")) = rstmp!???? & ""
            .TextMatrix(i, .ColIndex("????ip")) = rstmp!????IP & ""
            .TextMatrix(i, .ColIndex("????ip")) = rstmp!????IP & ""
            .TextMatrix(i, .ColIndex("????")) = rstmp!???? & ""
            .TextMatrix(i, .ColIndex("????")) = rstmp!???? & ""
            i = i + 1: rstmp.MoveNext
        Loop
        
        .AutoResize = True: .AutoSize 0, .Cols - 1
        
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Select .FixedRows, 0
        End If
    End With
    
    Exit Sub
errh:
    MsgBox err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub vsfPer_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
        Dim strTmp() As String
        
        With vsfPer
            If .Redraw = flexRDNone Then Exit Sub
            If .Row = 0 Then Exit Sub
            
            txtUser.Text = .TextMatrix(NewRow, .ColIndex("??????"))
            txtApp.Text = .TextMatrix(NewRow, .ColIndex("????????"))
            txtDesc.Text = .TextMatrix(NewRow, .ColIndex("????"))
            cmdStop.Caption = IIf(.TextMatrix(NewRow, .ColIndex("????")) = "??????", "????", "????")
            
            If .TextMatrix(NewRow, .ColIndex("????IP")) <> "" Then
                strTmp = Split(.TextMatrix(NewRow, .ColIndex("????IP")), ".")
                txtbeforeIp(0).Text = strTmp(0)
                txtbeforeIp(1).Text = strTmp(1)
                txtbeforeIp(2).Text = strTmp(2)
                txtbeforeIp(3).Text = strTmp(3)
                txtIP(2).Text = Split(.TextMatrix(NewRow, .ColIndex("????IP")), ".")(3)
            Else
                txtbeforeIp(0).Text = ""
                txtbeforeIp(1).Text = ""
                txtbeforeIp(2).Text = ""
                txtbeforeIp(3).Text = ""
                txtIP(2).Text = ""
            End If
        End With
End Sub

Private Sub vsfper_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsfper_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer
    
    With vsfPer
        If .Rows = .FixedRows Then Exit Sub
        If Col = 0 Then
            If .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture Then
                .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
                For i = .FixedRows To .Rows - .FixedRows
                    .TextMatrix(i, 0) = "-1"
                Next
            Else
                .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
                For i = .FixedRows To .Rows - .FixedRows
                    .TextMatrix(i, 0) = "0"
                Next
            End If
        End If
    End With
End Sub

Private Sub vsfper_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer, blnAllSelectd As Boolean
    
    blnAllSelectd = True
    With vsfPer
        If .Redraw = flexRDNone Then Exit Sub
        
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 0) = "0" Then
                blnAllSelectd = False
                Exit For
            End If
        Next

        
        If blnAllSelectd Then
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
        Else
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        End If
    End With
End Sub


Private Function ValidateTxt() As String
'????:????????????????,??????????????,????????????.
    Dim strStartIP As String, strEndIp As String
    Dim strErr As String
    
    If txtUser.Text = "" Or txtApp.Text = "" Then
        ValidateTxt = "??????????????????????????"
        Exit Function
    End If
    
    strStartIP = txtbeforeIp(0) & "." & txtbeforeIp(1) & "." & txtbeforeIp(2) & "." & txtbeforeIp(3)
    strEndIp = txtbeforeIp(0) & "." & txtbeforeIp(1) & "." & txtbeforeIp(2) & "." & txtIP(2)
    
    If strStartIP <> "..." Or strEndIp <> "..." Then
        Call CheckIpValidate(strStartIP, strEndIp, strErr)
    End If

    If strErr <> "" Then
        ValidateTxt = strErr
        Exit Function
    End If
  
End Function


Private Sub GetDataFromCard(ByRef strUser As String, ByRef strApp As String, ByRef strStartIP As String, ByRef strEndIp As String, ByRef strDesc As String)
'????:????????????????????
    
    strUser = Trim(txtUser.Text)
    strApp = Trim(txtApp.Text)
    
    If txtbeforeIp(0).Text = "" Then
        strStartIP = "": strEndIp = ""
    Else
        strStartIP = txtbeforeIp(0).Text & "." & txtbeforeIp(1).Text & "." & txtbeforeIp(2).Text & "." & txtbeforeIp(3).Text
        strEndIp = txtbeforeIp(0).Text & "." & txtbeforeIp(1).Text & "." & txtbeforeIp(2).Text & "." & IIf(txtIP(2).Text = "", txtbeforeIp(3).Text, txtIP(2).Text)
    End If
    
    strDesc = txtDesc.Text
    
End Sub


Private Sub EditPermission(Optional ByVal strStop As String)
'????:????????
    Dim strTmp As String, i As Integer
    Dim strStartIP As String, strEndIp As String
    Dim strDesc As String, strUser As String, strApp As String
    Dim strNewUser As String, strNewApp As String
    
    On Error GoTo errh
    
    With vsfPer
        strApp = .TextMatrix(.Row, .ColIndex("????????"))
        strUser = .TextMatrix(.Row, .ColIndex("??????"))
    End With
    '????????
    Call GetDataFromCard(strNewUser, strNewApp, strStartIP, strEndIp, strDesc)
    
    strTmp = ValidateTxt
    If strTmp <> "" Then
        frmMDIMain.stbThis.Panels(2).Text = strTmp
        Exit Sub
    End If
    
    '????????
    If strStop = "" Then
        '????????????????????,????????,????????????????????
        strStop = IIf(vsfPer.TextMatrix(vsfPer.Row, vsfPer.ColIndex("????")) = "??????", 1, 0)
    Else
        strStop = IIf(vsfPer.TextMatrix(vsfPer.Row, vsfPer.ColIndex("????")) = "??????", 0, 1)
    End If

    Screen.MousePointer = vbArrowHourglass
    Call ExecuteProcedure("zltools.Zl_Zlapppermission_Edit(2,'" & strApp & "','" & strUser & "','" & strStartIP & "','" & strEndIp & "'," & strStop & ",'" & strDesc & "','" & strNewApp & "','" & strNewUser & "' )", Me.Caption)
    Screen.MousePointer = vbDefault
    
    cmdStop.Caption = IIf(strStop = 0, "????", "????")
    With vsfPer
        .Redraw = flexRDNone
        Call LoadAppPermission
    End With
    frmMDIMain.stbThis.Panels(2).Text = "??????????????"
    Exit Sub
errh:
    Screen.MousePointer = vbDefault
    frmMDIMain.stbThis.Panels(2).Text = ""
    
    If InStr(1, UCase(err.Description), "ORA") Then '??????????,??????????,????????,????????????
        MsgBox "????????????????????" & vbNewLine & err.Description
    Else
        frmMDIMain.stbThis.Panels(2).Text = "????????????????????" & vbNewLine & err.Description
    End If
End Sub

Private Function GetSelectData() As String
'????:??????????????????????,????????ID,????????????
    Dim i As Integer, strTmp As String
    
    With vsfPer
        If .Rows = .FixedRows Then Exit Function
        
        '??????????????????
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 0) = "-1" Then
                If strTmp = "" Then
                    strTmp = .TextMatrix(i, .ColIndex("????????")) & ":" & .TextMatrix(i, .ColIndex("??????"))
                Else
                    strTmp = strTmp & "," & .TextMatrix(i, .ColIndex("????????")) & ":" & .TextMatrix(i, .ColIndex("??????"))
                End If
            End If
        Next
        
        If strTmp = "" Then
            '????????,??????????????????
            GetSelectData = .TextMatrix(.Row, .ColIndex("????????")) & ":" & .TextMatrix(.Row, .ColIndex("??????"))
        Else
            GetSelectData = strTmp
        End If
    End With
End Function

Private Sub txtbeforeIp_Change(Index As Integer)
    Dim lngLineNo As Long '????
    Dim lngColNo  As Long '????
    err = 0
    On Error Resume Next
    If Trim(txtbeforeIp(0).Text) <> "" And Trim(txtbeforeIp(1).Text) <> "" And Trim(txtbeforeIp(2).Text) <> "" And Trim(txtbeforeIp(3).Text) <> "" And Trim(txtIP(2).Text) <> "" Then
        cmdAdd.Enabled = True
    End If
    Call GetCursorPos(Me.txtbeforeIp(Index).hwnd, lngLineNo, lngColNo)
    If lngColNo > 3 Then
        If Index < 3 Then
            If txtbeforeIp(Index + 1).Enabled Then txtbeforeIp(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtbeforeIp_GotFocus(Index As Integer)
    txtbeforeIp(Index).SelStart = 0
    txtbeforeIp(Index).SelLength = Len(txtbeforeIp(Index).Text)
End Sub

Private Sub txtbeforeIp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lngLineNo As Long '????
    Dim lngColNo  As Long '????
    err = 0
    Call GetCursorPos(Me.txtbeforeIp(Index).hwnd, lngLineNo, lngColNo)
    
    Select Case KeyCode
    Case 37    '<-
        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtbeforeIp(Index - 1).Enabled Then
                txtbeforeIp(Index - 1).SelStart = Len(txtbeforeIp(Index - 1))
                txtbeforeIp(Index - 1).SetFocus
            End If
        End If
    Case 39    '->
        If Index < 3 Then
            If lngColNo <= Len(txtbeforeIp(Index)) Then Exit Sub
            If txtbeforeIp(Index + 1).Enabled Then
                txtbeforeIp(Index + 1).SelStart = 0
                txtbeforeIp(Index + 1).SetFocus
            End If
        End If
    Case 8     'BACKSPACE
        If Index > 0 Then
            If lngColNo > 1 Then Exit Sub
            If txtbeforeIp(Index - 1).Enabled Then
                txtbeforeIp(Index - 1).SelStart = Len(txtbeforeIp(Index - 1))
                txtbeforeIp(Index - 1).SetFocus
            End If
        End If
    End Select
    
    If InStr(1, "1234567890", Chr(KeyCode)) = 0 Then
        KeyCode = 0
    End If
    
End Sub

Private Sub txtbeforeIp_KeyPress(Index As Integer, KeyAscii As Integer)
    err = 0
    On Error Resume Next
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> 13 Then
            If KeyAscii <> 8 Then
                If KeyAscii = Asc(".") Then
                    If Index < 3 And Index >= 0 And Trim(txtbeforeIp(Index)) <> "" Then
                        If txtbeforeIp(Index + 1).Enabled Then txtbeforeIp(Index + 1).SetFocus
                    End If
                End If
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Function CheckPerOnly(ByVal strApp As String, ByVal strUser As String) As String
'????:????????????????????????,????????True,??????????False
    
    Dim strSQL As String, rstmp As ADODB.Recordset
    Dim arrUser() As String, strResult As String, i As Integer
    
    On Error GoTo errh
    
    strSQL = "Select ????????,??????" & vbNewLine & _
                "From Zlapppermission" & vbNewLine & _
                "Where ?????? In (Select /*+ cardinality(A,10) */" & vbNewLine & _
                "               Column_Value" & vbNewLine & _
                "              From Table(f_Str2list([1])) A) And ???????? = [2]"
    '??????????????????????????Oracle??????????,????????????????
    If Len(strUser) > 2000 Then
        arrUser = Str2Array(strUser, ",", 2000)
        For i = 0 To UBound(arrUser)
            Set rstmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "CheckPerOnly", arrUser(i), strApp)
            
            Do While Not rstmp.EOF
                If strResult = "" Then
                    strResult = rstmp!??????
                Else
                    strResult = strResult & "," & rstmp!??????
                End If
                rstmp.MoveNext
            Loop
        Next
    Else
        Set rstmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "CheckPerOnly", strUser, strApp)
        Do While Not rstmp.EOF
            If strResult = "" Then
                strResult = rstmp!??????
            Else
                strResult = strResult & "," & rstmp!??????
            End If
            rstmp.MoveNext
        Loop
    End If
    
    CheckPerOnly = strResult
    Exit Function
errh:
    MsgBox err.Description
    If 0 = 1 Then
        Resume
    End If
End Function


Private Sub GetCursorPos(ByVal hwnd5 As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long
    Dim lParam As Long, wParam As Long
    Dim K As Long
    
    i = SendMessage(hwnd5, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16 '??????????????????????????????Byte
    LineNo = SendMessage(hwnd5, EM_LINEFROMCHAR, j, 0) '????????????????????
    LineNo = LineNo + 1
    K = SendMessage(hwnd5, EM_LINEINDEX, -1, 0)
    '??????????????????????????????Byte
    ColNo = j - K + 1
End Sub

Public Function SupportPrint() As Boolean
'????????????????????????????????????
End Function


Private Sub FindApp()
'????:??????????????????????????
    
    Dim strSQL As String, rstmp As ADODB.Recordset
    Dim strResult As String
    
    On Error GoTo errh
    strSQL = "Select Distinct Program From V$session where Program not like 'ORACLE.EXE%' order by Program"
    
    Set rstmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "FindApp")
    
    Do While Not rstmp.EOF
        txtApp.AddItem rstmp!Program
        rstmp.MoveNext
    Loop
    Exit Sub
errh:
    MsgBox err.Description
End Sub

