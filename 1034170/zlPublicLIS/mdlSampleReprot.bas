Attribute VB_Name = "mdlSampleReprot"
Option Explicit

Public gstrSysName As String                        '系统名称
Public gstrProductName As String                    'OEM产品名称
Public gstrUnitName As String                       '用户单位名称
Public gcnOracle As New ADODB.Connection                 '公共数据库连接

Public UserInfo As TYPE_USER_INFO

'用户信息
Public Type TYPE_USER_INFO
    ID As Long
    编号 As String
    姓名 As String '人员姓名
    简码 As String
    DeptID As Long '部门ID
    DeptNo As String '部门编号
    DeptName As String '部门名称
    DBUser As String '数据库用户
End Type

Public glngSys As Long                              '系统号
Public glngModule As Long                           '模块号
Public gobjLISInsideComm As Object
Public gobjComLib As Object

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = Nvl(rsTmp!简码)
            UserInfo.姓名 = Nvl(rsTmp!姓名)
            UserInfo.DeptID = Nvl(rsTmp!部门ID, 0)
            UserInfo.DeptNo = rsTmp!部门码 & ""
            UserInfo.DeptName = rsTmp!部门名 & ""
            UserInfo.DBUser = rsTmp!用户名 & ""
            GetUserInfo = True
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub InitObjLis()
'判断如果新版LIS部件为空就初始化
    Dim strErr As String
    If gobjLISInsideComm Is Nothing Then
        On Error Resume Next
        Set gobjLISInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLISInsideComm Is Nothing Then
            If gobjLISInsideComm.InitComponentsHIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS部件初始化错误：" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLISInsideComm = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

