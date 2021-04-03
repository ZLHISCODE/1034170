VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsI_Configure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InitOracle(ByVal cnMain As ADODB.Connection)
    Set gcnOracle = cnMain
End Sub

'该部件用来完成医保接口的安装与卸载
Public Function I_Install(ByVal str主机串 As String) As Boolean
    Dim intRegist As Integer
    Dim strFile As String
    Dim strExchange As String, strTemp As String
    Dim blnExist As Boolean                         '是否存在以本接口部件注册的险种
    Dim blnTrans As Boolean, blnCreate As Boolean, blnRegisted As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim cnInterface As New ADODB.Connection
    
    mstrServer = str主机串
    intRegist = frm医保接口注册.ShowRegist(mintInsure, mstrInsureUser, mstrInsureTablespace, mstrInsureName, mstrDemo, mstrComponent, mstrPath)
    'intRegist:0-取消;1-新安装;2-升级;3-重复注册（指共用相同的医保部件以及中间库用户等，但险类不同）
    If intRegist = 0 Then Exit Function
    
    On Error GoTo ErrHand
    '判断当前安装的接口，是否已正常安装过，如果没有，又有中间库用户的，则先删除中间库，再创建
    mstrSQL = "Select 1 From 保险类别 Where 序号=" & mintInsure
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "判断是否已正常安装")
    blnRegisted = (rsTemp.RecordCount <> 0)
    
    '处理中间库
    If intRegist = 1 Then
        If mstrInsureUser <> "" Then
            '说明本次是第一次安装，可以删除中间库用户等
            On Error Resume Next
            gcnOracle.Execute "Drop User " & mstrInsureUser & " Cascade"
            If mstrInsureTablespace <> "" Then
                gcnOracle.Execute "Drop Tablespace " & mstrInsureTablespace
            End If
            
            On Error GoTo ErrHand
            '如果存在表空间则自动创建
            If mstrInsureTablespace <> "" Then
                gcnOracle.Execute "Create Tablespace " & mstrInsureTablespace & " Datafile '" & mstrInsureTablespace & ".DBF' Size 10M reuse"
            End If
            
            '自动创建中间库，并设置缺省表空间
            gcnOracle.Execute "Create User " & mstrInsureUser & " identified by his " & IIf(mstrInsureTablespace = "", "", " Default Tablespace " & mstrInsureTablespace)
            blnCreate = True
            
            '授权
            gcnOracle.Execute "Grant Resource,connect to " & mstrInsureUser
            gcnOracle.Execute "Grant Create Public Synonym to " & mstrInsureUser
        End If
    End If
    
    blnTrans = True
    gcnOracle.BeginTrans
    
    '插入保险类别基础数据
    If intRegist <> 2 Or blnRegisted = False Then
        mstrSQL = "zl_保险类别_Insert(" & mintInsure & ",'" & mstrInsureName & "'," & _
            "'" & mstrDemo & "',NULL,1,0)"
        gcnOracle.Execute mstrSQL, , adCmdStoredProc
        mstrSQL = "ZL_ZLINSURECOMPONENTS_INSERT(" & mintInsure & ",'" & mstrInsureUser & "','" & mstrInsureTablespace & "','" & mstrComponent & "')"
        gcnOracle.Execute mstrSQL, , adCmdStoredProc
    End If
    
    '1、运行zlManData.sql（包含医保管理数据及初始化数据）
    '2、运行zlInterface.sql（中间库用户的表及初始化数据）
    '3、运行zlStruc.sql（ZLHIS的结构变动）
    If Not RunSetup(mstrPath & "\" & IIf(intRegist <> 2, "Setup", "Spnew") & "\zlManData.sql", gcnOracle) Then
        gcnOracle.RollbackTrans
        blnTrans = False
        Exit Function
    End If
    
    '打开中间库连接
    If intRegist = 1 Then
        If mstrInsureUser <> "" Then
            With cnInterface
                .Open "Driver={Microsoft ODBC for Oracle};Server=" & mstrServer, mstrInsureUser, "his"
            End With
            If Not RunSetup(mstrPath & "\" & IIf(intRegist <> 2, "Setup", "Spnew") & "\zlInterface.sql", cnInterface) Then
                blnTrans = False
                gcnOracle.RollbackTrans
                cnInterface.Close
                Set cnInterface = New ADODB.Connection
                Set cnInterface = Nothing
                Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    blnTrans = False
    
    '运行数据结构修正脚本
    Call RunSetup(mstrPath & "\" & IIf(intRegist <> 2, "Setup", "Spnew") & "\zlStruc.sql", gcnOracle)
    
    '插入权限清单
    mstrSQL = " Select Distinct 100 As 系统,B.序号,B.功能,A.对象,A.权限" & _
              " From zlInsurePrivs A,zlInsureFuncs B" & _
              " Where A.险类 = B.险类 And A.方法 = B.方法 And A.险类=" & mintInsure
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.Open mstrSQL, gcnOracle
    
    '可能存在共用的权限，遇到主键重复时，不报错
    With rsTemp
        Do While Not .EOF
            mstrSQL = "Insert Into zlProgPrivs " & _
                      "(系统,序号,功能,对象,所有者,权限) " & _
                      "Values " & _
                      "(" & !系统 & "," & !序号 & ",'" & !功能 & "'," & _
                      "'" & !对象 & "',USER,'" & !权限 & "')"
            gcnOracle.Execute mstrSQL
            .MoveNext
        Loop
    End With
    
    I_Install = True
    Exit Function
ErrHand:
    If InStr(1, Err.Description, "出现重复") <> 0 Then Resume Next
    If InStr(1, Err.Description, "违反唯一约束条件") <> 0 Then Resume Next
    If InStr(1, Err.Description, "已经存在") <> 0 Then Resume Next
    If InStr(1, Err.Description, "或作用名发生冲突") <> 0 Then Resume Next
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function RunSetup(ByVal strFile As String, ByVal cnMain As ADODB.Connection) As Boolean
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
    Dim strLine As String, StrSQL As String
    Dim blnProcedure As Boolean, blnExecute As Boolean
    Dim strServer As String, strUser As String, strPass As String
    
    'intRegist:1-正常注册;2-重复注册
    On Error GoTo ErrHand
    If Not objFileSys.FileExists(strFile) Then
        RunSetup = True
        Exit Function
    End If
    
    Set objStream = objFileSys.OpenTextFile(strFile)
    Do While Not objStream.AtEndOfStream
        strLine = Trim(UCase(objStream.ReadLine))
        If InStr(1, strLine, "--") Then strLine = Mid(strLine, 1, InStr(1, strLine, "--") - 1)
        If strLine <> "" Then
            If StrSQL = "" Then
                blnProcedure = (InStr(1, strLine, "PROCEDURE") <> 0) Or (InStr(1, strLine, "FUNCTION") <> 0)
            End If
            StrSQL = StrSQL & IIf(Trim(StrSQL) <> "", vbCrLf, "") & strLine
            
            Call CheckExec(StrSQL, blnProcedure, blnExecute)
            If blnExecute Then
                If Not blnProcedure Then
                    StrSQL = Mid(StrSQL, 1, InStr(1, StrSQL, ";") - 1)
                Else
                    StrSQL = Mid(StrSQL, 1, InStr(1, StrSQL, "/") - 1)
                End If
                
                StrSQL = Replace(StrSQL, "GINTINSURE", mintInsure)
                cnMain.Execute StrSQL
                StrSQL = ""
            End If
        End If
    Loop
    
    RunSetup = True
    Exit Function
ErrHand:
    If InStr(1, Err.Description, "违反唯一约束条件") <> 0 Then Resume Next
    If InStr(1, Err.Description, "已经存在") <> 0 Then Resume Next
    If InStr(1, Err.Description, "无法删除约束条件") <> 0 Then Resume Next
    If InStr(1, Err.Description, "名称已被现有对象占用") <> 0 Then Resume Next
    
    MsgBox "出错的SQL：" & StrSQL, vbInformation, gstrSysName
End Function

Private Sub CheckExec(ByVal StrSQL As String, ByVal blnProcedure As Boolean, blnExecute As Boolean)
    Dim lngPos As Long
    Dim strText As String
    
    blnExecute = False
    If Trim(StrSQL) = "" Then Exit Sub
    
    For lngPos = Len(StrSQL) To 1 Step -1
        strText = Trim(Mid(StrSQL, lngPos, 1))
        If strText <> "" Then
            If strText = ";" And blnProcedure = False Then
                blnExecute = True
                Exit Sub
            ElseIf strText = "/" And blnProcedure Then
                blnExecute = True
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    Next
End Sub

Public Function I_Uninstall(ByVal intInsure As Integer) As Boolean
    Dim blnTrans As Boolean, blnDelete As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    mstrSQL = "Select 名称 From 保险类别 Where 序号=" & intInsure
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "取医保接口的名称")
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有找到该接口的注册信息！", vbInformation, gstrSysName
        Exit Function
    End If
    If MsgBox("你确定要卸载 " & Nvl(rsTemp!名称) & " 吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    '取中间库用户名
    mstrSQL = " Select Upper(用户名) 用户名,Upper(表空间) As 表空间,Upper(部件) AS 部件" & _
             " From zlInsureComponents " & _
             " Where 险类=" & intInsure
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "取中间库用户名")
    mstrInsureUser = Nvl(rsTemp!用户名)
    mstrInsureTablespace = Nvl(rsTemp!表空间)
    mstrComponent = Nvl(rsTemp!部件)
    
    mstrSQL = "Select 1 From zlInsureComponents Where Upper(部件)='" & mstrComponent & "' And 险类<>" & intInsure
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "判断是否还存在使用本医保部件的接口")
    blnDelete = (rsTemp.RecordCount = 0)
    
    blnTrans = True
    gcnOracle.BeginTrans
    '删除保险部件表
    gcnOracle.Execute "ZL_ZLINSURECOMPONENTS_DELETE(" & intInsure & ")", , adCmdStoredProc
    '删除支持库说明
    gcnOracle.Execute "ZL_ZLINSUREBASE_DELETE(" & intInsure & ")", , adCmdStoredProc
    '删除支持业务说明
    gcnOracle.Execute "ZL_ZLINSUREOPERATION_DELETE(" & intInsure & ")", , adCmdStoredProc
    '删除权限清单（交集部分权限不能删除）
    '原理：A    :当前险类涉及到的权限集合
    '      Total:所有险类涉及到的权限集合
    '      Total-A:除A外所有险类涉及到的权限集合
    '      A Minus ((Total-A) Intersect A)
    mstrSQL = " Delete ZlProgprivs Where (系统,序号,功能,对象,权限) IN " & _
              " (Select Distinct 100 As 系统,B.序号,B.功能,A.对象,A.权限" & _
              " From zlInsurePrivs A,zlInsureFuncs B" & _
              " Where A.险类 = B.险类 And A.方法 = B.方法 And A.险类=" & intInsure & _
              " Minus " & _
              " (Select Distinct 100 As 系统,B.序号,B.功能,A.对象,A.权限" & _
              " From zlInsurePrivs A,zlInsureFuncs B" & _
              " Where A.险类 = B.险类 And A.方法 = B.方法 And A.险类<>" & intInsure & _
              " Intersect " & _
              " Select Distinct 100 As 系统,B.序号,B.功能,A.对象,A.权限" & _
              " From zlInsurePrivs A,zlInsureFuncs B" & _
              " Where A.险类 = B.险类 And A.方法 = B.方法 And A.险类=" & intInsure & "))"
    gcnOracle.Execute mstrSQL
    
    '删除本接口支持的菜单
    gcnOracle.Execute "ZL_ZLINSUREMODULS_DELETE(" & intInsure & ")", , adCmdStoredProc
    '删除本接口支持的方法
    gcnOracle.Execute "ZL_ZLINSUREFUNCS_DELETE(" & intInsure & ")", , adCmdStoredProc
    '删除本接口所需要访问的对象及权限
    gcnOracle.Execute "ZL_ZLINSUREPRIVS_DELETE(" & intInsure & ")", , adCmdStoredProc
    
    '删除保险类别
    gcnOracle.Execute "zl_保险类别_Delete(" & intInsure & ")", , adCmdStoredProc
    gcnOracle.CommitTrans
    blnTrans = False
    
    '删除中间库用户
    On Error Resume Next
    '如果还存在其他险类使用到本医保接口部件的，则认为都使用相同的中间库，则不删除
    If blnDelete Then
        If mstrInsureUser <> "" Then
            gcnOracle.Execute "Drop User " & mstrInsureUser & " Cascade"
        End If
        If mstrInsureTablespace <> "" Then
            gcnOracle.Execute "Drop Tablespace " & mstrInsureTablespace
        End If
    End If
    
    I_Uninstall = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
End Function
