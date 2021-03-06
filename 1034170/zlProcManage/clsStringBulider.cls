VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringBulider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'说明：此类来自"孤帆"(2005-11-11 02:44:32),这个字符串操作类的自增缓冲区的特点是：当缓冲区不够大时将以128（128*2）的大小增加缓冲区大小

Private m_Buffer() As Byte   '缓冲区
Private m_lpBuffer As Long   '分配的内存块指针

Private m_AllocSize As Long  '实际分配的内存大小
Private m_Length As Long     '当前已经使用的内存大小

'*************************************************************************
'   设置缓冲区容量大小
'*************************************************************************
Property Let Capacity(ByVal Size As Long)
    If Size < 1 Then Exit Property

    Size = Size + Size '由于vb的字符串是BSTR字符占2个字节所以分配2倍的内存
    If Size <= m_AllocSize Then Exit Property '内存块还够用

    '   分配一个缓冲区
    If m_AllocSize = 0 Or m_Length = 0 Then '没有数据或还没有分配过内存
        ReDim m_Buffer(0 To Size)
    Else
        ReDim Preserve m_Buffer(0 To Size)
    End If
    m_lpBuffer = VarPtr(m_Buffer(0))        '缓冲区指针

    m_AllocSize = Size
End Property

'*************************************************************************
'   获取当前字符串长度
'*************************************************************************
Property Get Length() As Long
    Length = m_Length / 2
End Property

'*************************************************************************
'   清空数据
'*************************************************************************
Sub Clear()
    m_Length = 0
End Sub

'*************************************************************************
'� 在字符串缓冲区的后面增加字符串
'   参  数：要增加的字符串
'   返回值：无
'*************************************************************************
Sub Append(ByVal str As String)
    Const IncreaseStep = 128
    Dim nLen As Long '新加入的字符串长度

    nLen = Len(str)
    nLen = nLen + nLen

    '   如果m_AllocSize <= m_Length + nLen，则以 128+ nLen 个字符的增量增加内存
    If m_AllocSize <= m_Length + nLen Then
        Me.Capacity = (m_Length + nLen) / 2 + IncreaseStep
    End If

    '   拷贝数据到缓冲区
    Call CopyMemory(ByVal (m_lpBuffer + m_Length), ByVal StrPtr(str), nLen)
    m_Length = m_Length + nLen
End Sub

'*************************************************************************
'   转换为vb的字符串
'*************************************************************************
Property Get ToString() As String
    If m_Length > 1 Then
        ToString = Left(m_Buffer, m_Length / 2) '注意不要用Left$
    End If
End Property

'*************************************************************************
'   寻找 str 在缓冲区中第一次出现的位置(字符位置从 1 数起)
'   参  数：要查找的字符串,[是否忽视大小写(默认忽视大小写)]
'   返回值：成功则不为-1
'*************************************************************************
Function IndexOf(ByVal str As String, Optional ByVal IsIgnoreCase As Boolean = True) As Long
    If m_Length < 1 Then
        IndexOf = -1
        Exit Function
    End If

    If IsIgnoreCase Then
        IndexOf = StrCSpnI(m_lpBuffer, StrPtr(str))
    Else
        IndexOf = StrCSpn(m_lpBuffer, StrPtr(str))
    End If

    If IndexOf >= 0 Then IndexOf = IndexOf + 1
End Function

'*************************************************************************
'   寻找 str 在缓冲区中最后一次出现的位置(字符位置从 1 数起)
'   参  数：要查找的字符串,[是否忽视大小写(默认忽视大小写)]
'   返回值：成功则不为-1
'*************************************************************************
Function LastIndexOf(ByVal str As String, Optional ByVal IsIgnoreCase As Boolean = True) As Long
    If m_Length < 1 Then
        LastIndexOf = -1
        Exit Function
    End If

    If IsIgnoreCase Then
        LastIndexOf = StrRStrI(m_lpBuffer, m_lpBuffer + m_Length, StrPtr(str))
    Else
        LastIndexOf = StrRStr(m_lpBuffer, m_lpBuffer + m_Length, StrPtr(str))
    End If

    If LastIndexOf >= m_lpBuffer Then
        LastIndexOf = (LastIndexOf - m_lpBuffer) / 2 + 1
    Else
        LastIndexOf = -1
    End If
End Function

Private Sub Class_Initialize()
    m_AllocSize = 0
    m_Length = 0
End Sub

Private Sub Class_Terminate()
    Erase m_Buffer
End Sub
