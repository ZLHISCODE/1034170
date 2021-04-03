VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'#########################################################################
'##ģ �� ����cTile.cls
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��4��8��
'##��    ������һ��ͼƬƽ�̵�ָ��DC��������
'##Ҫ    ��
'#########################################################################
Option Explicit

Private m_lHdc As Long              'ͼƬDCֵ
Private m_lHBmp As Long             'ͼƬ������豸�����⣩
Private m_lHBmpOld As Long          '��ͼƬ������ڲ���
Private m_pic As StdPicture         '�洢ͼƬ����
Private m_sFileName As String       'ͼƬ�ļ���
Private m_lXOriginOffset As Long    'Xԭ��ƫ����
Private m_lYOriginOffset As Long    'Yԭ��ƫ����
Private m_lBitmapW As Long          'λͼ����
Private m_lBitmapH As Long          'λͼ�߶�

Private Const cTileErrorBase = 5600

 'ͼƬDCֵ
 Public Property Get PicturehDC() As Long
    PicturehDC = m_lHdc
End Property

 'Xԭ��ƫ����
Public Property Get XOriginOffset() As Long
    XOriginOffset = m_lXOriginOffset
End Property

 'Xԭ��ƫ����
Public Property Let XOriginOffset(ByVal lPixels As Long)
    m_lXOriginOffset = lPixels
End Property

'Yԭ��ƫ����
Public Property Get YOriginOffset() As Long
    YOriginOffset = m_lYOriginOffset
End Property

'Yԭ��ƫ����
Public Property Let YOriginOffset(ByVal lPixels As Long)
    m_lYOriginOffset = lPixels
End Property

'λͼ����
Public Property Get BitmapWidth() As Long
    BitmapWidth = m_lBitmapW
End Property

'λͼ�߶�
Public Property Get BitmapHeight() As Long
    BitmapHeight = m_lBitmapH
End Property

'����������Ϣ����������������ţ���Ϣ��
Private Sub pErr(lNumber As Long, smsg As String)
    MsgBox "����: " & smsg & "��" & lNumber, vbExclamation
End Sub

'ͼƬ�ļ���
Public Property Let FileName(ByVal sFIleName As String)
    '��һ���ļ�����ͼƬ
    If (m_sFileName <> sFIleName) Then
        '���ͼƬ�������
        pClearUp
        If (pbLoadPicture(sFIleName)) Then
            m_sFileName = sFIleName
        End If
    End If
End Property

'ͼƬ�ļ���
Public Property Get FileName() As String
    FileName = m_sFileName
End Property

'��ȡͼƬ����
Public Property Get Picture() As StdPicture
    Set Picture = m_pic
End Property

'ͼƬ���ݸ�ֵ
Public Property Let Picture(oPic As StdPicture)
   ' ��һ�� StdPicture �������ͼƬ
   pClearUp '���λͼ
   If Not oPic Is Nothing Then
      If (pbEnsurePicture()) Then   '�����ڴ�ռ�
          Set m_pic = oPic          '��ֵ
          If (Err.Number = 0) Then
              pbGetBitmapIntoDC     'ͬʱ������ؾ��
          End If
      End If
   End If
End Property

Private Function pbEnsurePicture() As Boolean
'ΪͼƬ�����ڴ�ռ䣬ȷ���ܹ�����ͼƬ��ȷ������ True�����򷵻� False
On Error Resume Next
    pbEnsurePicture = True
    If (m_pic Is Nothing) Then
        Set m_pic = New StdPicture  '��׼λͼ
        If (Err.Number <> 0) Then
            pErr 3, "�޷�ΪͼƬ��������ڴ�ռ䡣"
            pbEnsurePicture = False
        Else
        End If
    End If
On Error GoTo 0
    Exit Function
End Function

Private Function pbLoadPictureFromFile(sFile As String) As Boolean
'����ָ���ļ�����ͼƬ���ɹ����� True
On Error Resume Next
   Set m_pic = LoadPicture(sFile) '����ͼƬ
   If (Err.Number <> 0) Then
       pErr 0, "����ͼƬʧ��: " & Err.Description
   Else
       pbLoadPictureFromFile = True
   End If
On Error GoTo 0
    Exit Function
End Function

Private Function pbLoadPicture(sFile As String) As Boolean
'�������ã�����ָ�����Ƶ�ͼƬ���ɹ����� True
    '�ȷ���ͼƬ�ڴ�ռ� pbEnsurePicture()
    If (pbEnsurePicture()) Then
        '��ʽ����ͼƬ
        If (pbLoadPictureFromFile(sFile)) Then
            '����ɹ�����
            pbLoadPicture = pbGetBitmapIntoDC()
        End If
    End If
    
End Function

Private Function pbGetBitmapIntoDC() As Boolean
'�������ã� ��ȡλͼ��Ϣ��DC�У��ɹ����� True
Dim tB As BITMAP
Dim lHDC As Long, lHwnd As Long
Dim lHDCTemp As Long, lHBmpTempOld As Long

    ' ʹ��һ��DC������ͼƬλͼ���������ǲ���ͨ��������Blt������
    lHwnd = GetDesktopWindow()  '��ȡ���棨����Ļ�����д��塢ͼ��Ļ������򣩵� Window �����
    lHDC = GetDC(lHwnd)         '��ȡ�þ���� DC ������Ϊһ�� Window �����
    m_lHdc = CreateCompatibleDC(lHDC)   '�������ݵ��ڴ��豸���������ڴ�ͼ�񣩡�
    lHDCTemp = CreateCompatibleDC(lHDC) '
    If (m_lHdc <> 0) Then
        ' ��ȡλͼ�ߴ�
        ' GetObjectAPI ȡ�ö�ָ���������˵����һ���ṹ
        GetObjectAPI m_pic.Handle, LenB(tB), tB
        m_lBitmapW = tB.bmWidth     '��ȡͼƬ����
        m_lBitmapH = tB.bmHeight    '��ȡͼƬ�߶�
        
        'SelectObject ѡ��ָ�������豸�����У��Զ����Ǿɵ�ͬ�����
        '����1��DC���������2������ľ����
        lHBmpTempOld = SelectObject(lHDCTemp, m_pic.Handle)
        
        '����һ���豸���λͼ
        '����1���豸�������������2��ͼƬ���ȣ�����3��ͼƬ�߶�
        m_lHBmp = CreateCompatibleBitmap(lHDC, m_lBitmapW, m_lBitmapH)
        m_lHBmpOld = SelectObject(m_lHdc, m_lHBmp)  'ѡȡ��λͼ
        
        '��ֵ����ͼƬ�鵽Ŀ����������
        BitBlt m_lHdc, 0, 0, m_lBitmapW, m_lBitmapH, lHDCTemp, 0, 0, vbSrcCopy
        
        'ѡ��ͼƬ
        SelectObject lHDCTemp, lHBmpTempOld
        'ɾ���豸����
        DeleteDC lHDCTemp
        
        If (m_lHBmpOld <> 0) Then
            pbGetBitmapIntoDC = True
            If (m_sFileName = "") Then
               m_sFileName = "PICTURE"
            End If
        Else
            pClearUp    '���ͼƬ
            pErr 2, "�޷�ѡȡͼƬ���豸�����С�"
        End If
    Else
        pErr 1, "�޷��������ݵ��豸������"
    End If
    '�ͷ��豸����
    ReleaseDC lHwnd, lHDC
    
End Function

'��ɫ�壨���ص�ǰͼƬ���ݣ�
Public Property Get Palette() As StdPicture
    Set Palette = m_pic
End Property

Private Sub pClearUp()
'�������ã����ͼƬ�����Դ���������:m_lHdc/m_lHBmpOld/m_lHBmp��

    ' ����ļ���������
    m_sFileName = ""
    ' ��������Ѿ�����һ��DC����ô�����
    If (m_lHdc <> 0) Then
        ' ѡȡһ��DC��λͼ
        If (m_lHBmpOld <> 0) Then
            SelectObject m_lHdc, m_lHBmpOld
            ' ԭʼλͼ����ɾ������Ϊ���� m_pic ���С�
        End If
        If (m_lHBmp <> 0) Then
            DeleteObject m_lHBmp    'ɾ��ָ�������ͼ�ζ���
        End If
        ' �Ƴ�DC
        DeleteDC m_lHdc
    End If
End Sub

' ƽ��ͼƬ��ָ������
Public Sub TileArea( _
        ByRef hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long _
    )
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = ((X + m_lXOriginOffset) Mod m_lBitmapW)
    lSrcStartY = ((Y + m_lYOriginOffset) Mod m_lBitmapH)
    lSrcStartWidth = (m_lBitmapW - lSrcStartX)
    lSrcStartHeight = (m_lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = Y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (Y + Height)
        If (lDstY + lDstHeight) > (Y + Height) Then
            lDstHeight = Y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = X
        lSrcX = lSrcStartX
        Do While lDstX < (X + Width)
            If (lDstX + lDstWidth) > (X + Width) Then
                lDstWidth = X + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            '���ƾ���ͼƬ����
            BitBlt hdc, lDstX, lDstY, lDstWidth, lDstHeight, m_lHdc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = m_lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = m_lBitmapH
    Loop
End Sub


Private Sub Class_Terminate()
    'ȷ������GDI���󶼱��ͷ�
    pClearUp
    '���ͼƬ
    Set m_pic = Nothing
End Sub


