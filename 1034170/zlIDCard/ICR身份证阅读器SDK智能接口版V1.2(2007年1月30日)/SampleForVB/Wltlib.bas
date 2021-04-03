Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function InitComm Lib "termb.dll" (ByVal port As Integer) As Integer
Public Declare Function InitCommExt Lib "termb.dll" () As Integer
Public Declare Function Authenticate Lib "termb.dll" () As Integer
Public Declare Function AuthenticateExt Lib "termb.dll" () As Integer
Public Declare Function Read_Content_Path Lib "termb.dll" (ByVal fileName As String, ByVal active As Integer) As Integer
Public Declare Function Read_Content Lib "termb.dll" (ByVal active As Integer) As Integer
Public Declare Function CloseComm Lib "termb.dll" () As Integer
Public Declare Function GetSAMID Lib "termb.dll" () As String


