Attribute VB_Name = "mdlManager"
Option Explicit

Sub Main()
    frmUserLogin.Show 1
    If gcnOracle.State = 0 Then Exit Sub
    
    Call InitCommon(gcnOracle)
    frm医保接口管理.Show
End Sub

