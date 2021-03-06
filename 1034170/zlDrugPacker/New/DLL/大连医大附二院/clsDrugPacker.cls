VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugPacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
    Set gcolDevice = New Collection
End Sub

Private Sub Class_Terminate()
    Dim frmItem As Form
    
    On Error Resume Next
    For Each frmItem In Forms
        Unload frmItem
    Next
    
    Set gcolDevice = Nothing
    Set frmItem = Nothing
    Set gobjComLib = Nothing
End Sub

Public Function InitCommon(ByVal cnOracle As ADODB.Connection, ByVal frmOwner As Object, _
        ByVal lngSys As Long, ByVal lngModule As Long, Optional ByVal lngStockID As Long) As Boolean
'功能：初始化方法
'参数：
'  cnOracle：HIS的DB连接对象
'  frmOwner：主调窗体对象
'  lngSys：主调系统号
'  lngmodule：主调模块号
'  lngStockID：药房ID，可选，在药房业务时传递
'初始化根据当前业务，药房，通过对应的设备连接信息来进行初始化
'HIS端涉及的接口业务都需要调用初始化：收费（自助收费），药品目录管理，药品处方发药，药品部门发药
'收费：根据设备的门诊属性来找对应的设备，再根据连接信息进行初始化；可能有多个设备
'药品目录管理：对每个设备进行初始化
'药品处方发药、药品部门发药：根据传递的药房ID来找对应的设备，再根据连接信息进行初始化；可能有多个设备


    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim strKey As String
    Dim blnCalled As Boolean
    Dim blnMenuItem As Boolean
    
    On Error GoTo errHandle
    
    Set gobjConn = cnOracle
    Set gfrmOwner = frmOwner
    glngSys = lngSys
    glngModule = lngModule
    
    '创建zl9Comlib部件
    On Error Resume Next
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    If Err <> 0 Or gobjComLib Is Nothing Then
        gstrMessage = "引用zl9ComLib部件失败！" & vbNewLine & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo errHandle
    
    If Not gobjConn Is Nothing Then
        gobjComLib.InitCommon gobjConn
    End If
    
    '注册医院名称
    gstrRegHospital = gobjComLib.zlRegInfo("单位名称", , -1)
    
    '获取用户信息
    Call GetUserInfo
    
'    '获取设备信息
'    Call GetDeviceInfo
'
'    '获取设备参数
'    Call GetDeviceParam

    '嵌入菜单
    Call SetMenuItem
        
    InitCommon = True
        
    Exit Function
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Function

Public Property Get Message() As String
   Message = gstrMessage
End Property

'Public Sub SetInterface(ByVal cnOracle As ADODB.Connection, ByVal frmOwner As Object, ByVal lngSys As Long, ByVal lngModule As Long)
Public Sub SetInterface(ByVal bytType As Byte, Optional ByVal strKey As String)
'功能：各类功能设置
'参数：
'  bytType：
'    0：药品信息上传
'    1：药品库存上传
'    2：设备开关
'    3：设备参数设置
'    4：药房设置药品剂型
'  strKey：
'    bytType=1时，格式：药房ID

    Dim frmTmp As Form
    
    Select Case bytType
        Case 0
            Set frmTmp = New frmDrugInfo
            frmTmp.ShowMe gfrmOwner
        Case 1
            Set frmTmp = New frmDrugStock
            frmTmp.ShowMe Val(strKey), gfrmOwner
        Case 2
            Set frmTmp = New frmDeviceState
            frmTmp.ShowMe 0, gfrmOwner
        Case 3
            Set frmTmp = New frmDeviceSet
            frmTmp.Show vbModal, gfrmOwner
        Case 4
            Set frmTmp = New frmDeviceParam
            frmTmp.ShowMeByStock gfrmOwner, Val(strKey)
    End Select
    Set frmTmp = Nothing
    
End Sub

Public Function HisUpload(ByVal lngModule As Long, ByVal bytType As Byte, Optional ByVal varKey As Variant, Optional ByVal lngStockID As Long) As Boolean
'功能：通知接口按业务需要上传数据
'参数：
'   lngModule：模块号，用于判断业务
'   bytType：
'       1: 门诊处方上传 (配药)
'       2: 门诊发药通知 (发药)
'       3: 住院药品医嘱上传 (配、发药)
'   varKey：
'       当bytType=1时，varKey表示“单据;库房ID;NO”；
'       格式：“单据;库房ID;NO[|单据;库房ID;NO][|...]”
'       当bytType=2时，同bytType=1
'       当bytType=3时，varKey表示药品收发ID；
'       格式：“药品收发ID[,药品收发ID][,...]”
'   lngStockID：可选，药房事务时传递药房ID
'返回：True成功；False失败
    
    Dim rsData As ADODB.Recordset
    Dim objDevice As clsDevice
    Dim lngDeviceID As Long
    
            
    '1.判断当前业务是否需要上传数据
'    If GetJudge_IsNeedUpload(lngModule, bytType) = False Then Exit Function
        
    '得到已格式化的记录集
    Select Case bytType
        Case 1  '处方明细上传
            '获取HIS数据，并格式化成符合上传要求的格式
            Set rsData = mdlProcessData.SetUpload(bytType, varKey, lngModule)
            If rsData.RecordCount > 0 Then rsData.MoveFirst
            Do While rsData.EOF = False
                '设备对象
                If objDevice Is Nothing Then
                    Set objDevice = New clsDevice
                    objDevice.ID = rsData!DeviceID
                    lngDeviceID = rsData!DeviceID
                Else
                    If lngDeviceID <> rsData!DeviceID Then
                        lngDeviceID = rsData!DeviceID
                        objDevice.ID = lngDeviceID
                    End If
                End If
                
                '上传
                mdlDrugPacker.Dispense objDevice, rsData!NO, rsData!单据, rsData!Content
                
                rsData.MoveNext
            Loop
            rsData.Close
            Set objDevice = Nothing
            
        Case 2  '发药信息上传(处方概要信息)
            '获取HIS数据，并格式化成符合上传要求的格式
            Set rsData = mdlProcessData.SetUpload(bytType, varKey, lngModule)
            If rsData.RecordCount > 0 Then rsData.MoveFirst
            Do While rsData.EOF = False
                '设备对象
                If objDevice Is Nothing Then
                    Set objDevice = New clsDevice
                    objDevice.ID = rsData!DeviceID
                    lngDeviceID = rsData!DeviceID
                Else
                    If lngDeviceID <> rsData!DeviceID Then
                        lngDeviceID = rsData!DeviceID
                        objDevice.ID = lngDeviceID
                    End If
                End If
                
                '上传
                mdlDrugPacker.Dispensing objDevice, rsData!Content
                
                rsData.MoveNext
            Loop
            rsData.Close
            
        Case 3  '主要药品医嘱信息上传
            '获取HIS数据，并格式化成符合上传要求的格式
            Dim lng领药部门ID As Long
            Dim strContent As String
            Dim strPageNO As String
            Dim rsTmp As ADODB.Recordset
            
            Set rsData = mdlProcessData.SetUpload(bytType, varKey, lngModule)
            
            '排序
            rsData.Sort = "领药部门ID,DeviceID"
            If rsData.RecordCount > 0 Then rsData.MoveFirst
            
            Do While rsData.EOF = False
                lng领药部门ID = rsData!领药部门ID
                
                If objDevice Is Nothing Then
                    Set objDevice = New clsDevice
                    objDevice.ID = rsData!DeviceID
                    
                    If objDevice.Status = False Then GoTo makLoop
                    
                    '取单据号
                    On Error GoTo errHandle
                    strPageNO = "select convert(char(6),getdate(),12) + right('000000'+cast(isnull(max(substring(page_no,7,len(page_no))),0)+1 as varchar(4)),4) max_no " & _
                                "from atf_ypxx where convert(char(6),getdate(),12)=left(page_no,6) "
                    
                    Set rsTmp = New ADODB.Recordset
                    rsTmp.Open strPageNO, objDevice.DBConnect, adOpenStatic, adLockReadOnly
                    If rsTmp.EOF = False Then
                        strPageNO = "'" & rsTmp.Fields(0).Value & "'"
                    End If
                    rsTmp.Close
                    Set rsTmp = Nothing
                    
                    On Error GoTo errDB
                    objDevice.DBConnect.BeginTrans
                End If
                
                strContent = "insert into atf_ypxx " & _
                             "(DETAIL_SN,inpatient_no,p_id,name,ward_sn,ward_name,doctor,bed_no,comment,comm2,drug_code,drugname" & _
                             ",specification,dosage,dos_unit,atf_no,pri_flag,Mz_flag,dept_name,amount,occ_time,page_no) " & vbCrLf
                strContent = strContent & _
                             "select a.*, " & strPageNO & " from (" & rsData!Detail & ") A " & vbCrLf
                
                '上传明细
                If mdlDrugPacker.Dispensing(objDevice, strContent) = False Then GoTo errDB
                
                rsData.MoveNext
                If rsData.EOF Then
                    GoTo makCommon
                ElseIf lng领药部门ID <> rsData!领药部门ID Then
makCommon:
                    rsData.MovePrevious
                    
                    strContent = "insert into atf_yp_page_no (ward_sn,group_no,atf_no,submit_time,page_no) " & vbCrLf
                    strContent = strContent & rsData!Title & "," & strPageNO
                    
                    '上传单据
                    If mdlDrugPacker.Dispensing(objDevice, strContent) = False Then GoTo errDB
                    
                    objDevice.DBConnect.CommitTrans
                    
                    Set objDevice = Nothing
                End If
                
makLoop:
                rsData.MoveNext
            Loop
            rsData.Close
            Set objDevice = Nothing
            
    End Select
    
    HisUpload = True
    Exit Function
    
errHandle:
    gstrMessage = Err.Description
    If gobjComLib.ErrCenter = 1 Then Resume
    Exit Function

errDB:
    If Not objDevice Is Nothing Then
        gstrMessage = Err.Description
        On Error Resume Next
        objDevice.DBConnect.RollbackTrans
    End If
End Function
