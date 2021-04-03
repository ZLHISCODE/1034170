VERSION 5.00
Begin VB.Form frmXJCA 
   Caption         =   "新疆CA "
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   12135
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtValue 
      Height          =   495
      Left            =   3600
      TabIndex        =   13
      Top             =   3120
      Width           =   4815
   End
   Begin VB.TextBox txtShow 
      Height          =   2175
      Left            =   3240
      TabIndex        =   12
      Top             =   360
      Width           =   8175
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "证书驱动是否安装"
      Height          =   375
      Index           =   11
      Left            =   360
      TabIndex        =   11
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "证书驱动是否安装"
      Height          =   375
      Index           =   10
      Left            =   360
      TabIndex        =   10
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "集成测试"
      Height          =   375
      Index           =   9
      Left            =   360
      TabIndex        =   9
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "签章"
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "验证签名"
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "签名"
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "登录认证"
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "校验密码"
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "取证书DN值"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "取证书序列号"
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "key是否已就绪"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdCA 
      Caption         =   "证书驱动是否安装"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmXJCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjXJCA As New xjcaTechATLLib.xjcaTechATLLib
'Private mobjXJCA As Object
Private mobjSoapClient As Object
Private mobjXJCAHOS As Object
'Private Declare Function XJCA_SignSeal Lib "XJCA_HOS.dll" (ByRef strSrc As Byte, ByRef strxml As Byte, ByRef intT As Long) As Boolean
'Private Declare Function XJCA_SignSeal Lib "XJCA_HOS.dll" (ByRef strSrc As String, ByRef strxml As String, ByRef intT As Long) As Boolean
'Private Declare Function XJCA_SignSeal Lib "XJCA_HOS.dll" (ByVal strSrc As String, ByVal strxml As String, ByVal intT As Long) As Long
Private Declare Function XJCA_SignSeal Lib "XJCA_HOS.dll" (ByVal strSrc As String, ByVal strXml As Long, ByVal intT As Long) As Boolean
Private Declare Function XJCA_GetSealBMPB Lib "XJCA_HOS.dll" (ByVal lngFilePath As String, ByVal lngTimes As Long) As Boolean
Private Declare Function XJCA_VerifySeal Lib "XJCA_HOS.dll" (ByVal strSrc As String, ByVal strXml As String, ByVal strPic As String, ByVal strCert As String) As Boolean
'功能:验证签章数据.接口原型  bool   XJCA_VerifySeal((char* src,char* xml,char* pct,char* cert)
'参数：
'返回：True\false

'功能:获取签章图片
'参数:lngFilePath 路径字符串首地址：varPtr("C:\XXX\...")
Private Const BASE64CHR          As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private psBase64Chr(0 To 63)     As String
'Private strxml2 = strxml2 & "<?xml version=" & """" & "1.0" & """" & " encoding=" & """" & "Unicode" & """" & " standalone=" & """" & "yes" & """" & "?><XJCASIGN>" & _
'            "<SIGNATUREVALUE>MIIFAgYJKoZIhvcNAQcCoIIE8zCCBO8CAQExCzAJBgUrDgMCGgUAME8GCSqGSIb3DQEHAaBCBEAzNzM2NUVEMDI5N0Q5NDRDQTlCN0YxM" & _
'            "kIzNjg5QUMzQTM5QTJEN0Q2QkY0MURGMjQ4MEZBQTg4MzZFQzUzNzQ0oIIDWzCCA1cwggLAoAMCAQICDCtR8sEgQa9pV1tvTjANBgkqhkiG9w0BAQUFADB0MQ" & _
'            "swCQYDVQQGEwJDTjEhMB8GA1UECAwY5paw55aG57u05ZC+5bCU6Ieq5rK75Yy6MRUwEwYDVQQHDAzkuYzpsoHmnKjpvZAxDTALBgNVBAoMBHhqY2ExDTALBgNVBA" & _
'            "sMBHhqY2ExDTALBgNVBAMMBHhqY2EwHhcNMTQwODEyMDMxMDU5WhcNMTcwODExMDMxMDU5WjCBvTELMAkGA1UEBhMCQ04xLTArBgNVBAgeJAA2ADUAMAAxADAANQAxADk" & _
'            "ANwAwADAAMQAwADEAMAAwADIANjERMA8GA1UEBx4IADAAMAAyADYxETAPBgNVBAoeCGWwdYYAQwBBMREwDwYDVQQLHggAQwBBTi1fwzEjMCEGCSqGSIb3DQEJARYUeGpjYXhtc3"
'  strxml2 = strxml2 & "NAeGpjYS5jb20uY24xITAfBgNVBAMeGGWwdYYAQwBBY6VT421Li9UAMAAwADIANjCBnzANBgkqhkiG9w0BAQEFAAOBjQAwgYkCgYEApZ2Z2NuBIhshRpYXPK5qKkw0F4iO02As6h3IBJSUdDv0iRvthEyrbDl4C" & _
'  "mrMJiUWlj3FTiSKkNN0TQnTFOB9NGhSGy/F4mIABAaAnxOQaGADv7uumss+tLeX4IsDUOd/i3tFnuMqcm+ToIjrXWuxl9g6vV6fSxYYinZN+bD67/ECAwEAAaOBozCBoDAMBgNVHRMEBTADAQEAMAsGA1UdDwQE" & _
'  "AwIAwDAdBgNVHSUEFjAUBggrBgEFBQcDAgYIKwYBBQUHAwgwJAYFKlYVAQEEGzFANjAxN1NGM" & _
'"DY1MDEwNTE5NzAwMTAxMDAyNjAfBgNVHSMEGDAWgBTDhV7v0cusdg8guwKqM75c+FuqjTAdBgNVHQ4EFgQUjXfbCNAfAbiQyJYMsYmyJ01qJakwDQYJKoZIhvcNAQEFBQADgYEAEZyOQzXJS6x66GIF9KA0ife1YLEc+JI" & _
'"sAaslV3r7c6xS8TDlow+ILCmLBc9E+Jy+SjOqcgRiHulGFxaflhVR75EXc2R50fyoK89nTsZ0+" & _
'"73LRio5Z57dY0IzDRfQijKY/uPhtSeN09/JLn84BSYqkaZNAeleis4lc2M+nS9PsA0xggErMIIBJwIBATCBhDB0MQswCQYDVQQGEwJDTjEhMB8GA1UECAwY5paw55aG57u05ZC+5bCU6Ieq5rK75Yy6MRUwEwYDVQQH" & _
'"DAzkuYzpsoHmnKjpvZAxDTALBgNVBAoMBHhqY2ExDTALBgNVBAsMBHhqY2ExDTALBgNVBAMMBHhqY2ECDCtR8sEgQa9pV1tvTjAJBgUrDgMCGgUAMA0GCSqGSIb3DQEBAQUABIGAchGzvgVbEX9OTa1kddSWKKj0CZCPv" & _
'"TfcCwv1aWfThpQEwMUt7T6NiKpRg5UPHEKmokqkF2aM8UFsRc1VHLuFBNjgNa49J4ewBArZi3iM60upz1HrGyi" & _
'"T0s2qBpX9vTqRAFGpSzRVza64SnwZ8KtlCeo4mG4N5HEO5nWZNBeSBxU=|6F95D9C544C5D25F3F904135FA3B95A1DAC1550E3DA0F523467187A4126612F4</SIGNATUREVALUE>"
'
'strxml2 = strxml2 & "<SIGNEDLENGTH>MTcxNg==</SIGNEDLENGTH></XJCASIGN>" & _
'",6Iiw5ZyD6Iiw7ICCzqDEgsiC4qyM74mR4oOB6r2B5Z2p5r2b44GO2I3iqInkoobvnobEjdSBBeeQsOCsseCksMyG0ZXhjIbkjILjhY7jgKHYn+WUg+CghOGgjOmbpu6esOialuuvp+" & _
'"6WtOu6kOuDpe6ilOqqh+uLpu6Wu+uqjOGUseGMsMyG0ZXgsIfukIzosrnri6numoHqopzrt6njhpDjgI3Yi+WUg+CohNCM5qm45oWj4LSx4KywzIbRleCwi+eghOaNquOFoeOAjdiL5ZSDzITQjOapuOaFo+G4sOC0l+OQseOgsOOI" & _
'"seOMsOOAseOkteGdmuOEjeOAt+OEuOOAseOEs+OUsOWoueiEsOOGveOAi9iJ5ZSD2ITIk+S5g+K0seKssMyG0ZXhuIgkNjUwMTA1MTk3MDAxMDEwMDLjhLbjgJHYj+WUg9yE4KCe44CA44CA44iA45iA4YSx4LywzIbRleG4iuaUiOeWsMKGQ+OFgeOAkdi" & _
'"P5ZSD4KyE4KCe5IyA5ISA4rWO7I2f4oyx4oSw4KSG6Jiq6JmI4Le34KSB4ZiB56CU5o2q56Gh542t5IGz5qm45oWj5oyu5rWv5oyu44Wu44Ch2J/llIPMhOGgnuuBpeiZteSMgOSEgOqVo+6Nk+Stre2Wi+OAgOOAgOOIgOOYgOiEsOOCn9iN4qiJ5KKG756GxI3EgQXohIPCjeiEsMqJ6IaB6pSA6aad7a+Y4oqB4oSb6ZmG" & _
'"47CX5qqu5LCq4Zy06LqI5oOT7qis7KCd6ZCE55KU75C74a6J6JOt6q2M46Ws4Km47LGq4pSm6ZiW7JS94pGO6YKK55OT4KWN4ZOT57eg5qC04a2S7JSv5oui0IDogIbhjp/mopDNoOuuv+mqruO7i+uetO6Cl86L7p2Q6K2/5JW77o6e54iq6Y2v6KKg5ber64Wr7aKX67S66b2e4ZmL6KiY5LW264O57r+6y7HEg8SA6Iaj44Kj6oKB4LCwzIbhtZXQk+OAhcSDAeCssMyG4b" & _
'"WV0I/MhALjg4DYneWUg+KUneGYhOGQsOCghtir1IHchciD4KCG2KvUgdyF4KCD4pCw1IblmKrEldCB44Sb45mA44Sw5Yy344GG45S244Sw45Sw46Sx44C344Sw44Sw44Cw45iy4bywzIbhtZXQo+OAmOiAluyMlOW6he2Hr+qzi+C9tuusoOqoguu4s++hnOqpm+OCjdid5ZSD4Lid4ZiE4ZCE556N4KOb4b+Q66CB7KKQ4LKW6Kax4p6y5qmN6qSl4LSw4KSG6Jiq6JmI4Le" & _
'"3xIHUhcyA6IaB4YSA6Lqc45WD5K+J56qs5ouo75CF45Kg756J5oK14bKx6Yu4xKzilqvnqZfnj7vliqzjg7Hqj6XooI/ipKzWi+STj+mzuOSqvuqos9Gy4bmi5Jup4ZiX6Zqf5YSV6Yev54yX56Wk77OR4q6o5p+P7JmO76207K694qmG5py57bae5Imj4LSz7YCX44qK77qY7oej4p617Y6N7Kef57yu1LjiqKbqmpHFjeW7qey6iueMpeO5o+K+neuBjw1QTW1UUE16RFBNL" & _
'"3pQL0FEUC9NelAvWmpQL21UUC96RFAvLzJZQUFHWUFNMllBWm1ZQW1XWUF6R1lBLzJZekFHWXpNMll6Wm1Zem1XWXp6R1l6LzJabUFHWm1NMlptWm1abW1XWm16R1ptLzJhWkFHYVpNMmFaWm1hWm1XYVp6R2FaLz" & _
'"JiTUFHYk1NMmJNWm1iTW1XYk16R2JNLzJiL0FHYi9NMmIvWm1iL21XYi96R2IvLzVrQUFKa0FNNWtBWnBrQW1aa0F6SmtBLzVrekFKa3pNNWt6WnBrem1aa3p6Smt6LzVsbUFKbG1NNWxtWnBsbW1abG16SmxtLzVtWkFKbVpNNW1aWnBtWm1abVp6Sm1aLzVuTUFKbk1NNW5NWnBuTW1abk16Sm5NLzVuL0FKbi9NNW4vWnBuL21abi96Sm4vLzh3QUFNd0FNOHdBWnN3QW1jd0F6TXdBLzh3ekFNd3pNOHd6WnN3e" & _
'"m1jd3p6TXd6Lzh4bUFNeG1NOHhtWnN4bW1jeG16TXhtLzh5WkFNeVpNOHlaWnN5Wm1jeVp6TXlaLzh6TUFNek1NOHo=,6DA1676D5A2B2D926DE15F7CFABB27AB46B1E9DB1D81B755557ED355E5D555C2"


Private Sub cmdCA_Click(Index As Integer)
    Dim lngRet As Long
    Dim strSN As String
    Dim strInfo As String
    Dim strDN As String
'    Dim strXml(40000) As Byte
    Dim i As Long

  
    Dim blnRet As Boolean
    
    lngRet = 0: strSN = "": strInfo = "": strDN = ""
    txtShow.Text = ""
    Select Case Index
        
    Case 0
        lngRet = mobjXJCA.XJCA_CspInstalled("HaiTai Cryptographic Service Provider for xjca") '10000 表示已经安装
'        lngRet = XJCA_CspInstalled("HaiTai Cryptographic Service Provider for xjca")
        strInfo = "是否安装驱动? " & vbCrLf & "返回值：" & lngRet
    Case 1
        lngRet = mobjXJCA.XJCA_KeyInsert("HaiTai Cryptographic Service Provider for xjca")
        strInfo = "证书是否插入? " & vbCrLf & "返回值：" & lngRet
    Case 2
        strSN = Space(100)
        strSN = mobjXJCA.XJCA_GetCertSN

        strInfo = "证书序列号? " & vbCrLf & "返回值：" & strSN
    Case 3 '取证书dn值
        txtShow.Text = mobjXJCA.XJCA_GetCertDN
        strInfo = strDN
    Case 4 '校验PIN
        lngRet = mobjXJCA.XJCA_VerifyPin("123456", Len("123456"))   '10000 标识成功
        strInfo = lngRet
    Case 5 ' 认证
'        strDN = mobjXJCA.XJCA_GetCertDN
'        strDN = Mid(strDN, InStr(strDN, "CN=") + 3)
        strDN = "新疆CA接口测试0026"
        Dim str1 As String
        Dim str2 As String
        str1 = "http://124.117.245.71:18080/webServices/authService"
        str2 = "4028e48a39dd529a0139dd5c383d0010"
        strInfo = mobjXJCA.XJCA_CertAuth(str1, str2, strDN)
        Debug.Print strInfo
'        strInfo = StrConv(strInfo, vbUnicode)
    Case 6 '签名
        strInfo = mobjXJCA.XJCA_SignStr("HaiTai Cryptographic Service Provider for xjca", "数据源")
        Debug.Print strInfo
    Case 7 '验证签名
        strSN = mobjXJCA.XJCA_SignStr("HaiTai Cryptographic Service Provider for xjca", "数据源")
        
'        strInfo = mobjXJCA.XJCA_VerifySignStr(strSrc(0), strxml(0), 0)
    Case 8
        On Error Resume Next
'        Dim bytSrc(1000) As Byte
'        Dim bytXml(1000) As Byte
'        Dim lngLen As Long
'        bytSrc(0) = 178
'        bytSrc(1) = 226
'        blnRet = XJCA_SignSeal(bytSrc(0), bytXml(0), lngLen)
        Dim strSrc1 As String

        Dim lngLen As Long
        
        strSrc1 = "ss"
        Dim strXml As String
        Dim strxml2 As String
        strXml = """"
        Dim bytXml(40000) As Byte
        blnRet = XJCA_SignSeal(strSrc1, VarPtr(bytXml(0)), VarPtr(lngLen))
'         For i = LBound(bytXml) To 40000
'            strXml = strXml & Chr(bytXml(i))
'         Next
         
         Dim blnTm As Boolean
        blnTm = XJCA_VerifySeal(strSrc1, strxml2, Null, Null)
        Debug.Print blnTm
    Case 9
        Dim strName As String, strUniqueID As String, strCert As String, strCertDN As String
        
        Call XJCA_GetCertList(strName, strUniqueID, strCert, strCertDN)
    Case 10
        Dim strFile As String
        
'        strFile = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & mobjXJCA.XJCA_GetCertSN & ".bmp"
        strFile = "d:\www.bmp"
        blnRet = XJCA_GetSealBMPB(strFile, 2)
        If blnRet = False Then Exit Sub
        
    End Select
    txtShow.Text = strInfo
    Debug.Print txtShow.Text
End Sub


Public Function XJCA_GetCertList(ByRef strName As String, Optional ByRef strUniqueID As String, Optional ByRef strCert As String, Optional ByRef strCertDN As String) As Boolean
    '-入参:无
    '-出参
    'strName :      保存接口返回的证书所有者姓名
    'strUniqueID:   保存接口返回的证书所有者唯一标识
    'strCert:       保存接口返回的签名证书
    On Error GoTo errH
    Dim strSrc As String
    Dim strSN As String
    Dim strTmp As String
    Dim lngLen As Long
    Dim bytXml() As Byte
    Dim blnRet As Boolean
    Dim lngCount As Long
    Dim arrTmp As Variant
    Dim i As Long
    
    On Error GoTo errH
    strTmp = CStr(mobjXJCA.XJCA_GetCertDN())      'C=CN, S=650105197001010026, L=0026, O=新疆CA, OU=CA中心, E=xjcaxmss@xjca.com.cn, CN=新疆CA接口测试0026
    arrTmp = Split(strTmp, ",")

     strCertDN = arrTmp(0) & "," & arrTmp(1) & "," & arrTmp(2) & "," & arrTmp(3) & "," & arrTmp(4) & "," & arrTmp(5) & "," & arrTmp(6)
    txtValue.Text = CStr(mobjXJCA.XJCA_GetCertSN)    '证书序号
    strUniqueID = Trim(txtValue.Text)

    If Len(strUniqueID) <= 5 Or strUniqueID = "10002" Then
        MsgBox "获取证书序列号失败！", vbInformation + vbOKOnly, ""
        Exit Function
    End If
    strSrc = "1234567890"
    ReDim bytXml(40000) As Byte
    blnRet = XJCA_SignSeal(VarPtr(strSrc), VarPtr(bytXml(0)), VarPtr(lngLen))
    If blnRet = False Then
        MsgBox "读取证书信息失败！", vbInformation + vbOKOnly, ""
        Exit Function
    Else
        strTmp = ""
        For i = LBound(bytXml) To UBound(bytXml)
            strTmp = strTmp & Chr(bytXml(i))
        Next
        'strSrc返回的值分三部分："签名值,证书信息,证书ID(与证书序号有别)"
        strCert = Split(strTmp, ",")(1)  '证书信息
    End If
    
    strName = Mid(strCertDN, InStr(strCertDN, "CN=") + 3)   '获取证书持有者姓名

    XJCA_GetCertList = True
    Exit Function
errH:
    MsgBox "读取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, ""

End Function

