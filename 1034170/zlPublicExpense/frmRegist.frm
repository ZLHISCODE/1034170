VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmRegist 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeCommandBars.CommandBars cbsTemp 
      Left            =   1200
      Top             =   990
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object, _
                            ByVal blnAddInTool As Boolean, MenuControlBefore As CommandBarControl, ToolControlBefore As CommandBarControl)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    Set mfrmParent = frmParent
    If cbsMain Is Nothing Then Exit Sub
    If frmParent.Name = "frmDistRoomManager" And glngModul <> 1113 Then Exit Sub
    If frmParent.Name = "frmOutDoctorStation" And glngModul <> 1260 Then Exit Sub
    If frmParent.Name = "frmInDoctorStation" And glngModul <> 1261 Then Exit Sub
    
    If glngModul = 1113 Then
        '�������
        If MenuControlBefore Is Nothing Then
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_EditPopup)
            Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Regist, "�Һ�", 1, False)
        Else
            Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, MenuControlBefore.ID)
            Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Manage_Regist, "�Һ�", objControl.Index, False)
        End If
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        Set objControl = objMenu.CommandBar.Controls.Find(, conMenu_File_Exit)
        Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Option, "�Һ�ѡ������", objControl.Index, False)
    
        '����������
        '-----------------------------------------------------
        If blnAddInTool Then
            Set objBar = cbsMain(2)
            If ToolControlBefore Is Nothing Then
                With objBar.Controls
                    Set objControl = .Find(, conMenu_File_Preview)
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Һ�", objControl.Index, False)
                    objControl.BeginGroup = True
                End With
            Else
                With objBar.Controls
                    Set objControl = .Find(, ToolControlBefore.ID)
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Һ�", objControl.Index, False)
                    objControl.BeginGroup = True
                End With
            End If
        End If
        
        '����Ŀ����
        '-----------------------------------------------------
        With cbsMain.KeyBindings
            .Add 0, vbKeyF3, conMenu_Manage_Regist
        End With
    
        '���ò���������
        '-----------------------------------------------------
        With cbsMain.Options
        End With
        For Each objControl In objBar.Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    End If
    
    If glngModul = 1260 Then
        '����ҽ������վ
        If MenuControlBefore Is Nothing Then
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
            Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Regist, "�Һ�", 1, False)
            Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ", 2, False)
            Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, 3564, "ԤԼ�Ǽ�", 3, False)
        Else
            Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, MenuControlBefore.ID)
            cbsMain.ActiveMenuBar.Controls.Add xtpControlButton, conMenu_Manage_Regist, "�Һ�", objControl.Index, False
            cbsMain.ActiveMenuBar.Controls.Add xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ", objControl.Index, False
            cbsMain.ActiveMenuBar.Controls.Add xtpControlButton, 3564, "ԤԼ�Ǽ�", objControl.Index, False
        End If
        
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        Set objControl = objMenu.CommandBar.Controls.Find(, conMenu_File_Exit)
        Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Option, "�Һ�ѡ������", objControl.Index, False)
    
'        '����������
'        '-----------------------------------------------------
        If blnAddInTool Then
            Set objBar = cbsMain(2)
            With objBar.Controls
                If ToolControlBefore Is Nothing Then
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Һ�", 1, False)
                Else
                    Set objControl = .Find(, ToolControlBefore.ID)
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Һ�", objControl.Index, False)
                End If
            End With
            
            For Each objControl In objBar.Controls
                objControl.Style = xtpButtonIconAndCaption
            Next
        End If
    End If
    
    If glngModul = 1261 Then
        'סԺҽ������վ
        If MenuControlBefore Is Nothing Then
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
            Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ", 1, False)
            Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, 3564, "ԤԼ�Ǽ�", 1, False)
        Else
            Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, MenuControlBefore.ID)
            cbsMain.ActiveMenuBar.Controls.Add xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ", objControl.Index, False
            cbsMain.ActiveMenuBar.Controls.Add xtpControlButton, 3564, "ԤԼ�Ǽ�", objControl.Index, False
        End If
        
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        Set objControl = objMenu.CommandBar.Controls.Find(, conMenu_File_Exit)
        Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Option, "�Һ�ѡ������", objControl.Index, False)
    End If
    
    If glngModul = 1115 Then
        '���߷�������
        If MenuControlBefore Is Nothing Then
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit)
            Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ�Һ�", 1, False)
        Else
            Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, MenuControlBefore.ID)
            cbsMain.ActiveMenuBar.Controls.Add xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ�Һ�", objControl.Index, False
        End If
        
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        Set objControl = objMenu.CommandBar.Controls.Find(, conMenu_File_Exit)
        Set objControl = objMenu.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Option, "�Һ�ѡ������", objControl.Index, False)
        If blnAddInTool Then
            Set objBar = cbsMain(2)
            With objBar.Controls
                If ToolControlBefore Is Nothing Then
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ�Һ�", 1, False)
                Else
                    Set objControl = .Find(, ToolControlBefore.ID)
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_Bespeak, "ԤԼ�Һ�", objControl.Index, False)
                End If
            End With
            
            For Each objControl In objBar.Controls
                objControl.Style = xtpButtonIconAndCaption
            Next
        End If
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    If Control.ID = conMenu_Manage_Regist Then
        Control.Enabled = zlCheckPrivs(gstrPrivs, "�Һ�")
        Control.Visible = zlCheckPrivs(gstrPrivs, "�Һ�")
    End If
    If Control.ID = conMenu_Manage_Bespeak Then
        Control.Enabled = zlCheckPrivs(gstrPrivs, "ԤԼ")
        Control.Visible = zlCheckPrivs(gstrPrivs, "ԤԼ")
    End If
    If Control.ID = conMenu_View_Option Then
        Control.Enabled = zlCheckPrivs(gstrPrivs, "�Һ�ѡ������")
        Control.Visible = zlCheckPrivs(gstrPrivs, "�Һ�ѡ������")
    End If
    If Control.ID = 3564 Then
        Control.Enabled = zlCheckPrivs(gstrPrivs, "ԤԼ�Ǽ�")
        Control.Visible = zlCheckPrivs(gstrPrivs, "ԤԼ�Ǽ�")
    End If
End Sub

Public Sub zlExecuteCommandBars(ByVal frmMain As Object, ByVal Control As CommandBarControl, _
                                ByRef strOutNO As String, Optional ByVal lngPatiID As Long)
    Dim strSQL As String, rsTmp As ADODB.Recordset, datNow As Date
    Select Case Control.ID
        Case conMenu_Manage_Regist
            If glngModul = 1113 Then
                If gbytRegistMode = 0 Then
                    frmDistRoomRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, strOutNO, False
                Else
                    datNow = gobjDatabase.CurrentDate
                    If Format(datNow, "yyyy-mm-dd") < Format(gdatRegistTime, "yyyy-mm-dd") Then
                        frmDistRoomRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, strOutNO, False
                    Else
                        frmDistRoomRegistNew.zlShowMe frmMain, glngModul, gstrDeptIDs, strOutNO, False
                    End If
                End If
            End If
            If glngModul = 1260 Then
                If gstrDeptIDs = "" Then
                    strSQL = "Select Distinct a.����id" & vbNewLine & _
                            " From ������Ա A, ��������˵�� B" & vbNewLine & _
                            " Where a.��Աid = [1] And a.����id = b.����id And b.������� In (1, 3)"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
                    Do While Not rsTmp.EOF
                        gstrDeptIDs = gstrDeptIDs & "," & Nvl(rsTmp!����ID)
                        rsTmp.MoveNext
                    Loop
                    If gstrDeptIDs <> "" Then gstrDeptIDs = Mid(gstrDeptIDs, 2)
                End If
                gstrRooms = gobjDatabase.GetPara("��������", glngSys, 1260, "")
                If UCase(gstrRooms) = "NONE" Then gstrRooms = ""
                If gbytRegistMode = 0 Then
                    frmStationRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, False, lngPatiID, strOutNO
                Else
                    datNow = gobjDatabase.CurrentDate
                    If Format(datNow, "yyyy-mm-dd") < Format(gdatRegistTime, "yyyy-mm-dd") Then
                        frmStationRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, False, lngPatiID, strOutNO
                    Else
                        frmStationRegistNew.zlShowMe frmMain, glngModul, gstrDeptIDs, False, lngPatiID, strOutNO
                    End If
                End If
            End If
            If glngModul = 1261 Then
                If gstrDeptIDs = "" Then
                    strSQL = "Select Distinct a.����id" & vbNewLine & _
                            " From ������Ա A, ��������˵�� B" & vbNewLine & _
                            " Where a.��Աid = [1] And a.����id = b.����id And b.������� In (1, 3)"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
                    Do While Not rsTmp.EOF
                        gstrDeptIDs = gstrDeptIDs & "," & Nvl(rsTmp!����ID)
                        rsTmp.MoveNext
                    Loop
                    If gstrDeptIDs <> "" Then gstrDeptIDs = Mid(gstrDeptIDs, 2)
                End If
                If gbytRegistMode = 0 Then
                    frmStationRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, False, lngPatiID, strOutNO
                Else
                    datNow = gobjDatabase.CurrentDate
                    If Format(datNow, "yyyy-mm-dd") < Format(gdatRegistTime, "yyyy-mm-dd") Then
                        frmStationRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, False, lngPatiID, strOutNO
                    Else
                        frmStationRegistNew.zlShowMe frmMain, glngModul, gstrDeptIDs, False, lngPatiID, strOutNO
                    End If
                End If
            End If
        Case conMenu_Manage_Bespeak
            If glngModul = 1113 Then
                If gbytRegistMode = 0 Then
                    frmDistRoomRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, strOutNO, True
                Else
                    datNow = gobjDatabase.CurrentDate
                    If Format(datNow, "yyyy-mm-dd") < Format(gdatRegistTime, "yyyy-mm-dd") Then
                        frmDistRoomRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, strOutNO, True
                    Else
                        frmDistRoomRegistNew.zlShowMe frmMain, glngModul, gstrDeptIDs, strOutNO, True
                    End If
                End If
            End If
            If glngModul = 1260 Then
                If gstrDeptIDs = "" Then
                    strSQL = "Select Distinct a.����id" & vbNewLine & _
                            " From ������Ա A, ��������˵�� B" & vbNewLine & _
                            " Where a.��Աid = [1] And a.����id = b.����id And b.������� In (1, 3)"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
                    Do While Not rsTmp.EOF
                        gstrDeptIDs = gstrDeptIDs & "," & Nvl(rsTmp!����ID)
                        rsTmp.MoveNext
                    Loop
                    If gstrDeptIDs <> "" Then gstrDeptIDs = Mid(gstrDeptIDs, 2)
                End If
                gstrRooms = gobjDatabase.GetPara("��������", glngSys, 1260, "")
                If UCase(gstrRooms) = "NONE" Then gstrRooms = ""
                If gbytRegistMode = 0 Then
                    frmStationRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, True, lngPatiID, strOutNO
                Else
                    datNow = gobjDatabase.CurrentDate
                    If Format(datNow, "yyyy-mm-dd") < Format(gdatRegistTime, "yyyy-mm-dd") Then
                        frmStationRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, True, lngPatiID, strOutNO
                    Else
                        frmStationRegistNew.zlShowMe frmMain, glngModul, gstrDeptIDs, True, lngPatiID, strOutNO
                    End If
                End If
            End If
            If glngModul = 1261 Then
                If gstrDeptIDs = "" Then
                    strSQL = "Select Distinct a.����id" & vbNewLine & _
                            " From ������Ա A, ��������˵�� B" & vbNewLine & _
                            " Where a.��Աid = [1] And a.����id = b.����id And b.������� In (1, 3)"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
                    Do While Not rsTmp.EOF
                        gstrDeptIDs = gstrDeptIDs & "," & Nvl(rsTmp!����ID)
                        rsTmp.MoveNext
                    Loop
                    If gstrDeptIDs <> "" Then gstrDeptIDs = Mid(gstrDeptIDs, 2)
                End If
                If gbytRegistMode = 0 Then
                    frmStationRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, True, lngPatiID, strOutNO
                Else
                    datNow = gobjDatabase.CurrentDate
                    If Format(datNow, "yyyy-mm-dd") < Format(gdatRegistTime, "yyyy-mm-dd") Then
                        frmStationRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, True, lngPatiID, strOutNO
                    Else
                        frmStationRegistNew.zlShowMe frmMain, glngModul, gstrDeptIDs, True, lngPatiID, strOutNO
                    End If
                End If
            End If
            If glngModul = 1115 Then
                If gbytRegistMode = 0 Then
                    frmDistRoomRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, strOutNO, True
                Else
                    datNow = gobjDatabase.CurrentDate
                    If Format(datNow, "yyyy-mm-dd") < Format(gdatRegistTime, "yyyy-mm-dd") Then
                        frmDistRoomRegist.zlShowMe frmMain, glngModul, gstrDeptIDs, strOutNO, True
                    Else
                        frmDistRoomRegistNew.zlShowMe frmMain, glngModul, gstrDeptIDs, strOutNO, True
                    End If
                End If
            End If
        Case conMenu_View_Option
            frmRegistPara.zlShowMe frmMain, glngModul
        Case 3564
            If Not frmAppRequestManage Is Nothing Then Unload frmAppRequestManage
            If gbytRegistMode = 0 Then
                MsgBox "�ƻ��Ű�ģʽ����ʹ��ԤԼ�Ǽǹ���!", vbInformation, gstrSysName
                Exit Sub
            Else
                frmAppRequestManage.Show 0, frmMain
            End If
        Case conMenu_Edit_AppRequest
            If gbytRegistMode = 0 Then
                MsgBox "�ƻ��Ű�ģʽ����ʹ��ԤԼ�Ǽǹ���!", vbInformation, gstrSysName
                Exit Sub
            Else
                frmAppRequestEdit.ShowMe frmMain, lngPatiID
            End If
    End Select
End Sub
