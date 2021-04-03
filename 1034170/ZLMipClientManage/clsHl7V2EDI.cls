VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHL7V2EDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
'����
Private mstrPath As String
Private mdocXML As DOMDocument
Private mdocSegmentXML As DOMDocument

'######################################################################################################################
'����


'######################################################################################################################
'����

Public Function GetMessageFormat(ByVal strConfigFile As String, ByRef rsFormat As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim xn As IXMLDOMNode
    Dim xlst As IXMLDOMNodeList
    Dim strFileName As String

    On Error GoTo errHand

    '------------------------------------------------------------------------------------------------------------------
    '�жϲ����ļ�

    Dim fso As New FileSystemObject

    strFileName = fso.GetFileName(strConfigFile)

    Set fso = Nothing

    Set mdocXML = New DOMDocument
    mdocXML.Load strConfigFile


    strConfigFile = Mid(strConfigFile, 1, Len(strConfigFile) - Len(strFileName)) & "HL7.Segment"

    Set mdocSegmentXML = New DOMDocument
    mdocSegmentXML.Load strConfigFile

    '------------------------------------------------------------------------------------------------------------------
    '
    Set rsFormat = New ADODB.Recordset
    With rsFormat
        .Fields.Append "id", adVarChar, 32, adFldKeyColumn
        .Fields.Append "parent_id", adVarChar, 32
        .Fields.Append "�ڵ�����", adVarChar, 50
        .Fields.Append "�ڵ����", adVarChar, 30
        .Fields.Append "��������", adVarChar, 20
        .Fields.Append "�ظ�Ƶ��", adVarChar, 30
        .Fields.Append "�ڵ�˵��", adVarChar, 200
        .Fields.Append "�����ظ�", adVarChar, 200
        .Fields.Append "���ݸ�ֵ", adVarChar, 200
        .Open
    End With

    '------------------------------------------------------------------------------------------------------------------
    'ȡ�汾��

    Set xn = mdocXML.selectSingleNode(".//Meta//Release")
'    strMessageVer = xn.Text

    '------------------------------------------------------------------------------------------------------------------
    Set xn = mdocXML.selectSingleNode(".//Message//MessageType")
'    strMessageType = xn.Text

    '------------------------------------------------------------------------------------------------------------------
    Set xn = mdocXML.selectSingleNode(".//Message//Group")

    If Not (xn Is Nothing) Then

        Set xlst = xn.childNodes

        If GetMessageNode(xlst, 0, "", rsFormat) = False Then Exit Function

    End If

    If rsFormat.RecordCount > 0 Then rsFormat.MoveFirst
    '------------------------------------------------------------------------------------------------------------------
    GetMessageFormat = True

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Private Function GetMessageNode(ByVal xlst As IXMLDOMNodeList, ByVal intLayer As Integer, ByVal strParentDataKey As String, ByRef rsFormat As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim xn As IXMLDOMNode
    Dim xlst1 As IXMLDOMNodeList
    Dim xlst2 As IXMLDOMNodeList
    Dim xnf As IXMLDOMNode
'    Dim xe As IXMLDOMElement
    Dim strTemp As String
    Dim i As Long
    Dim objPropertyCollect As Collection
    Dim strItem As String
    Dim strDataKey As String
    Dim strSegmentName As String
    Dim strNodeName As String
    Dim strDataType As String
    Dim strNodeType As String
    Dim strDataMin As String
    Dim strDataMax As String
    Dim strNodeInfo As String

    On Error GoTo errHand

    intLayer = intLayer + 1

    For Each xnf In xlst

        '��ȡ��ǰ�ڵ������
        Set objPropertyCollect = Nothing
        Set objPropertyCollect = New Collection

        If xnf.Attributes.length > 0 Then
            For i = 0 To xnf.Attributes.length - 1
                objPropertyCollect.Add xnf.Attributes(i).Text, "K_" & UCase(xnf.Attributes(i).nodeName)
            Next
        End If

        strNodeType = xnf.nodeName
        strNodeName = ""
        strDataType = ""
        strNodeInfo = ""
        strDataMin = ""
        strDataMax = ""

        On Error Resume Next

        'strNodeName
        '--------------------------------------------------------------------------------------------------------------
        Select Case UCase(strNodeType)
        Case UCase("Group")
            strNodeName = objPropertyCollect(UCase("K_name"))
        Case UCase("Segment")
            strNodeName = objPropertyCollect(UCase("K_nodeName"))
            strSegmentName = objPropertyCollect(UCase("K_ref"))
            If strNodeName = "" Then strNodeName = strSegmentName

        Case UCase("Data")
            strNodeName = objPropertyCollect(UCase("K_nodeName"))
            If strNodeName = "" Then strNodeName = objPropertyCollect(UCase("K_ref"))
        Case UCase("Composite")
            strNodeName = objPropertyCollect(UCase("K_nodeName"))
        End Select

        'strDataType
        '--------------------------------------------------------------------------------------------------------------
        Select Case UCase(strNodeType)
        Case UCase("Composite")
            strDataType = objPropertyCollect(UCase("K_ref"))
        End Select

        'strDataMin
        '--------------------------------------------------------------------------------------------------------------
        strItem = ""
        strItem = objPropertyCollect(UCase("K_minOccurs"))
        If strItem <> "" Then strDataMin = strItem

        'strDataMax
        '--------------------------------------------------------------------------------------------------------------
        strItem = ""
        strItem = objPropertyCollect(UCase("K_maxOccurs"))
        If strItem <> "" Then strDataMax = strItem


        'strNodeInfo
        '--------------------------------------------------------------------------------------------------------------
        Select Case UCase(strNodeType)
        Case UCase("Composite")
            strNodeInfo = objPropertyCollect(UCase("K_info"))
        End Select

        On Error GoTo errHand

        '��д��¼
        '--------------------------------------------------------------------------------------------------------------
        rsFormat.AddNew
        strDataKey = zlCommFun.GetGUID
        rsFormat("id").Value = strDataKey
        rsFormat("parent_id").Value = strParentDataKey
        rsFormat("�ڵ�����").Value = Space(2 * (intLayer - 1)) & strNodeType
        rsFormat("�ڵ����").Value = strNodeName
        rsFormat("��������").Value = strDataType

        Select Case UCase(strDataMax)
        Case UCase("unbounded")
            rsFormat("�ظ�Ƶ��").Value = Val(strDataMin) & " �� u"
        Case Else
            rsFormat("�ظ�Ƶ��").Value = Val(strDataMin) & " �� " & Val(strDataMax)
        End Select

        rsFormat("�ڵ�˵��").Value = strNodeInfo

        '��
        '--------------------------------------------------------------------------------------------------------------
        If UCase(strNodeType) = UCase("Segment") Then
            Set xn = mdocSegmentXML.selectSingleNode(".//Elements//Segment[@name='" & strSegmentName & "']")
            If Not (xn Is Nothing) Then
                Set xlst1 = xn.childNodes
                If GetMessageNode(xlst1, intLayer, strDataKey, rsFormat) = False Then Exit Function
            End If
        End If

        '�����¼�
        '--------------------------------------------------------------------------------------------------------------
        If Not (xnf.childNodes Is Nothing) Then

            Set xlst2 = xnf.childNodes
            If GetMessageNode(xlst2, intLayer, strDataKey, rsFormat) = False Then Exit Function
        End If



    Next

    GetMessageNode = True

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
    Resume
End Function
