VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Detail"
Attribute VB_Ext_KEY = "Member0" ,"Detail"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit '要求变量声明

'局部变量，保存集合
Private mCol As Collection

Public Function Add(ID As Long, 药名ID As Long, 类别 As String, 类别名称 As String, 名称 As String, 编码 As String, _
    简码 As String, 别名 As String, 规格 As String, 计算单位 As String, 说明 As String, 屏蔽费别 As Boolean, _
    药房包装 As Double, 药房单位 As String, 分批 As Boolean, 变价 As Boolean, 加班加价 As Boolean, 执行科室 As Byte, _
    服务对象 As Integer, 类型 As String, Optional 补充摘要 As Boolean, Optional 固有从属 As Byte, _
    Optional 从项数次 As Double = 1, Optional 跟踪在用 As Boolean, Optional 库存 As Double, Optional sKey As String, _
    Optional 价格 As Double, Optional 处方限量 As Double = -1, Optional 处方职务 As String = "00", Optional 要求审批 As Boolean, Optional 录入限量 As Double, _
    Optional 中药形态 As Integer = 0, Optional 商品名 As String = "", _
    Optional 批次 As Long = -1, Optional ByVal 商品条码 As String = "", Optional 内部条码 As String = "") As Detail
    '创建新对象
    Dim objNewMember As Detail
    Set objNewMember = New Detail

    '设置传入方法的属性
    objNewMember.Key = sKey
    objNewMember.ID = ID
    objNewMember.药名ID = 药名ID
    objNewMember.类别 = 类别
    objNewMember.类别名称 = 类别名称
    objNewMember.名称 = 名称
    objNewMember.编码 = 编码
    objNewMember.简码 = 简码
    objNewMember.别名 = 别名
    objNewMember.规格 = 规格
    objNewMember.药房包装 = 药房包装
    objNewMember.药房单位 = 药房单位
    objNewMember.计算单位 = 计算单位
    objNewMember.说明 = 说明
    objNewMember.屏蔽费别 = 屏蔽费别
    objNewMember.分批 = 分批
    objNewMember.变价 = 变价
    objNewMember.加班加价 = 加班加价
    objNewMember.补充摘要 = 补充摘要
    objNewMember.执行科室 = 执行科室
    objNewMember.服务对象 = 服务对象
    objNewMember.类型 = 类型
    objNewMember.固有从属 = 固有从属
    objNewMember.从项数次 = 从项数次
    objNewMember.跟踪在用 = 跟踪在用
    objNewMember.库存 = 库存
    objNewMember.价格 = 价格
    objNewMember.处方限量 = 处方限量
    objNewMember.录入限量 = 录入限量
    objNewMember.处方职务 = IIf(处方职务 = "", "00", 处方职务)
    objNewMember.要求审批 = 要求审批
    objNewMember.中药形态 = 中药形态
    objNewMember.商品名 = 商品名
    objNewMember.批次 = 批次
    objNewMember.商品条码 = 商品条码
    objNewMember.内部条码 = 内部条码
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Detail
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：'debug.print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub
