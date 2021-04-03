VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------------
'功能：扩展接口插件模板
'如果要做限制使用 请实现 GetUserName 方法
'---------------------------------------------------------------------------

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口初始化
'参数：cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    
    Set gcnOracle = cnOracle
    
End Sub

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1, Optional ByVal strReserve As String) As String
'功能：获取外挂功能列表；该接口一般在初始化之后调用一次，调用时机和频度可能因不同的主程序而不同
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=保留参数,用于扩展使用
'        医技站中的调用时,传入为当前执行的前提医嘱ID
'返回：外挂功能列表，格式为"功能名1,功能名2,功能名3,..."
'      如果某个功能需要在病人切换(或刷新)时自动调用，则在功能名前面加上"Auto:"前缀，这种功能只允许有一个。
    
    GetFuncNames = "功能名1,功能名2,功能名3,功能名4,功能名5"
End Function

Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, _
    ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String, Optional ByVal int场合 As Integer = -1)
'功能：执行某个外挂功能
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号
'      strFuncName=要执行的功能名，该功能名之前由GetFuncNames函数返回
'      lngPatiID=当前病人ID
'      varRecId=数字或者字符串；对门诊病人，为当前挂号单号或者挂号ID；对住院病人，为当前住院主页ID
'      varKeyId=数字或者字符串；当前的关键业务数据唯一标识ID，如医嘱ID
'      strReserve=保留参数,用于扩展使用
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    
    MsgBox strFuncName & "执行成功！"
    
End Sub
 
Public Function GetUserName() As String
'功能：获取单位名称
'说明：用户名空时表示不限制
'例  GetUserName = "中联技术中心医院信息系统"

    GetUserName = ""
End Function
 
Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口终止
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
     
End Sub

