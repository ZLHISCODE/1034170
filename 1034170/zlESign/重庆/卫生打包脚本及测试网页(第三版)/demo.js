var certid = null;
ezca.checkCrl=true; //false true
function SOF_GetVersion1(){
	alert('版本:'+ezca.SOF_GetVersion());
}
function SOF_SetSignMethod1()
{
	ezca.SOF_SetSignMethod(131328);
}
function SOF_GetSignMethod1()
{
	alert(ezca.SOF_GetSignMethod());
}
function SOF_SetEncryptMethod1()
{
	ezca.SOF_SetEncryptMethod(131328);
}
function SOF_GetEncryptMethod1()
{
	alert(ezca.SOF_GetEncryptMethod());
}
function SOF_GetUserList1()
{
	var certlist = ezca.SOF_GetUserList();
	var infos = certlist.split('||');
	certid = infos[1];
	document.getElementById('SOF_GetUserList').value = certlist;
	//document.getElementById('certid1').value = certid;
	//document.getElementById('certid2').value = certid;
	//document.getElementById('SOF_ChangePassWd1').value = certid;
	//document.getElementById('SOF_ExportExChangeUserCert1').value = certid;
	//document.getElementById('SOF_GetUserInfo1').value = certid;
	//document.getElementById('SOF_SignData1').value = certid;
	//document.getElementById('SOF_SignFile1').value = certid;
	//document.getElementById('SOF_PriKeyDecrypt1').value = certid;
	//document.getElementById('SOF_SignDataByP71').value = certid;
	//document.getElementById('SOF_SignDataXML1').value = certid;
}
function SOF_GetUserInfo()
{
	var SOF_GetUserInfo1 = document.getElementById('SOF_GetUserInfo1').value;
	var SOF_GetUserInfo2 = document.getElementById('SOF_GetUserInfo2').value;
	alert(ezca.SOF_GetUserInfo(SOF_GetUserInfo1,SOF_GetUserInfo2));
}
function SOF_ExportUserCert1()
{
	var certid1 = document.getElementById('certid1').value;
	document.getElementById('SOF_ExportUserCert2').value = ezca.SOF_ExportUserCert(certid1);
	document.getElementById('SOF_GetCertInfo1').value = ezca.SOF_ExportUserCert(certid1);
	document.getElementById('SOF_GetCertInfoByOid1').value = ezca.SOF_ExportUserCert(certid1);
	document.getElementById('SOF_ValidateCert1').value = ezca.SOF_ExportUserCert(certid1);
	document.getElementById('SOF_VerifySignedData1').value = ezca.SOF_ExportUserCert(certid1);
	document.getElementById('SOF_VerifySignedFile1').value = ezca.SOF_ExportUserCert(certid1);
}
function SOF_Login()
{
	var certid2 = document.getElementById('certid2').value;
	var passwd1 = document.getElementById('passwd1').value;
//	if(passwd1=="")
//	{
//		alert('请输入待验证的密码！');
//		return;
//	}
	alert(ezca.SOF_Login(certid2,passwd1));
}
function SOF_ChangePassWd()
{
	var certid = document.getElementById('SOF_ChangePassWd1').value;
	var oldpasswd = document.getElementById('SOF_ChangePassWd2').value;
	var newpasswd = document.getElementById('SOF_ChangePassWd3').value;
//	if(oldpasswd==""||newpasswd=="")
//	{
//		alert('请正确输入密码！');
//		return;
//	}
	alert(ezca.SOF_ChangePassWd(certid,oldpasswd,newpasswd)?'修改成功':'修改失败');
}
function SOF_ExportExChangeUserCert()
{
	var certid1 = document.getElementById('SOF_ExportExChangeUserCert1').value;
	var temp = ezca.SOF_ExportExChangeUserCert(certid1);
	document.getElementById('SOF_ExportExChangeUserCert2').value = temp;
	document.getElementById('SOF_PubKeyEncrypt1').value = temp;
}
function SOF_GetCertInfo()
{
	var temp = document.getElementById('SOF_GetCertInfo1').value;
//	if(temp=="")
//	{
//		alert('证书实体不能为空！');
//		return;
//	}
	alert(ezca.SOF_GetCertInfo(temp,document.getElementById('SOF_GetCertInfo2').value));
}
function SOF_GetCertInfoByOid()
{
	var SOF_GetCertInfoByOid1 = document.getElementById('SOF_GetCertInfoByOid1').value;
	var SOF_GetCertInfoByOid2 = document.getElementById('SOF_GetCertInfoByOid2').value;
//	if(SOF_GetCertInfoByOid1=="")
//	{
//		alert('证书实体不能为空！');
//		return;
//	}
	document.getElementById('SOF_GetCertInfoByOid3').value=ezca.SOF_GetCertInfoByOid(SOF_GetCertInfoByOid1,SOF_GetCertInfoByOid2);
	//alert(ezca.SOF_GetCertInfoByOid(SOF_GetCertInfoByOid1,SOF_GetCertInfoByOid2));
}
function SOF_ValidateCert()
{
	var certid1 = document.getElementById('SOF_ValidateCert1').value;
	alert(ezca.SOF_ValidateCert(certid1));
}
function SOF_SignData()
{
	var certid1 = document.getElementById('SOF_SignData1').value;
	var SOF_SignData2 = document.getElementById('SOF_SignData2').value;
	document.getElementById('SOF_SignData3').value = ezca.SOF_SignData(certid1,SOF_SignData2);
	document.getElementById('SOF_VerifySignedData3').value = document.getElementById('SOF_SignData3').value;
}
function SOF_VerifySignedData()
{
	var SOF_VerifySignedData1 = document.getElementById('SOF_VerifySignedData1').value;
	var SOF_VerifySignedData2 = document.getElementById('SOF_VerifySignedData2').value;
	var SOF_VerifySignedData3 = document.getElementById('SOF_VerifySignedData3').value;
//	if(SOF_VerifySignedData1==""||SOF_VerifySignedData2=="")
//	{
//		alert("输入不能为空");
//		return;
//	}
	alert(ezca.SOF_VerifySignedData(SOF_VerifySignedData1,SOF_VerifySignedData2,SOF_VerifySignedData3));
}
function SOF_SignFile()
{
	var SOF_SignFile1 = document.getElementById('SOF_SignFile1').value;
	var SOF_SignFile2 = document.getElementById('SOF_SignFile2').value;
	var temp = ezca.SOF_SignFile(SOF_SignFile1,SOF_SignFile2);
	document.getElementById('SOF_SignFile3').value = temp;
	document.getElementById('SOF_VerifySignedFile3').value = temp;
}
function SOF_VerifySignedFile()
{
	var SOF_VerifySignedFile1 = document.getElementById('SOF_VerifySignedFile1').value;
	var SOF_VerifySignedFile2 = document.getElementById('SOF_VerifySignedFile2').value;
	var SOF_VerifySignedFile3 = document.getElementById('SOF_VerifySignedFile3').value;
	alert(ezca.SOF_VerifySignedFile(SOF_VerifySignedFile1,SOF_VerifySignedFile2,SOF_VerifySignedFile3));
}
function SOF_EncryptData()
{
	var SOF_EncryptData1 = document.getElementById('SOF_EncryptData1').value;
	var SOF_EncryptData2 = document.getElementById('SOF_EncryptData2').value;
	document.getElementById('SOF_EncryptData3').value = ezca.SOF_EncryptData(SOF_EncryptData1,SOF_EncryptData2);
	document.getElementById('SOF_DecryptData2').value = ezca.SOF_EncryptData(SOF_EncryptData1,SOF_EncryptData2);
}
function SOF_DecryptData()
{
	var SOF_DecryptData1 = document.getElementById('SOF_DecryptData1').value;
	var SOF_DecryptData2 = document.getElementById('SOF_DecryptData2').value;
	document.getElementById('SOF_DecryptData3').value = ezca.SOF_DecryptData(SOF_DecryptData1,SOF_DecryptData2);
}
function SOF_EncryptFile()
{
	var SOF_EncryptFile1 = document.getElementById('SOF_EncryptFile1').value;
	var SOF_EncryptFile2 = document.getElementById('SOF_EncryptFile2').value;
	var SOF_EncryptFile3 = document.getElementById('SOF_EncryptFile3').value;
//	if(SOF_EncryptFile1==""||SOF_EncryptFile2==""||SOF_EncryptFile3=="")
//	{
//		alert('请选择对应文件');
//		return;
//	}
	document.getElementById('SOF_DecryptFile2').value = SOF_EncryptFile3;
	alert(ezca.SOF_EncryptFile(SOF_EncryptFile1,SOF_EncryptFile2,SOF_EncryptFile3));
}
function SOF_DecryptFile()
{
	var SOF_DecryptFile1 = document.getElementById('SOF_DecryptFile1').value;
	var SOF_DecryptFile2 = document.getElementById('SOF_DecryptFile2').value;
	var SOF_DecryptFile3 = document.getElementById('SOF_DecryptFile3').value;
	alert(ezca.SOF_DecryptFile(SOF_DecryptFile1,SOF_DecryptFile2,SOF_DecryptFile3));
}
function SOF_PubKeyEncrypt()
{
	var SOF_PubKeyEncrypt1 = document.getElementById('SOF_PubKeyEncrypt1').value;
	var SOF_PubKeyEncrypt2 = document.getElementById('SOF_PubKeyEncrypt2').value;
	document.getElementById('SOF_PubKeyEncrypt3').value = ezca.SOF_PubKeyEncrypt(SOF_PubKeyEncrypt1,SOF_PubKeyEncrypt2);
	document.getElementById('SOF_PriKeyDecrypt2').value = document.getElementById('SOF_PubKeyEncrypt3').value;
}
function SOF_PriKeyDecrypt()
{
	var SOF_PriKeyDecrypt1 = document.getElementById('SOF_PriKeyDecrypt1').value;
	var SOF_PriKeyDecrypt2 = document.getElementById('SOF_PriKeyDecrypt2').value;
	document.getElementById('SOF_PriKeyDecrypt3').value = ezca.SOF_PriKeyDecrypt(SOF_PriKeyDecrypt1,SOF_PriKeyDecrypt2);
}
function SOF_SignDataByP7()
{
	var SOF_SignDataByP71 = document.getElementById('SOF_SignDataByP71').value;
	var SOF_SignDataByP72 = document.getElementById('SOF_SignDataByP72').value;
	document.getElementById('SOF_SignDataByP73').value = ezca.SOF_SignDataByP7(SOF_SignDataByP71,SOF_SignDataByP72);
	document.getElementById('SOF_VerifySignedDataByP71').value = document.getElementById('SOF_SignDataByP73').value;
	document.getElementById('SOF_GetP7SignDataInfo1').value = document.getElementById('SOF_SignDataByP73').value;
}
function SOF_VerifySignedDataByP7()
{
	var SOF_VerifySignedDataByP71 = document.getElementById('SOF_VerifySignedDataByP71').value;
//	if(SOF_VerifySignedDataByP71=="")
//	{
//		alert('输入不能为空！');
//		return;
//	}
	alert(ezca.SOF_VerifySignedDataByP7(SOF_VerifySignedDataByP71));
}
function SOF_GetP7SignDataInfo()
{
	var SOF_GetP7SignDataInfo1 = document.getElementById('SOF_GetP7SignDataInfo1').value;
	var SOF_GetP7SignDataInfo2 = document.getElementById('SOF_GetP7SignDataInfo2').value;
//	if(SOF_GetP7SignDataInfo1==""||SOF_GetP7SignDataInfo2=="")
//	{
//		alert("输入项不能为空！");
//		return;
//	}
	alert(ezca.SOF_GetP7SignDataInfo(SOF_GetP7SignDataInfo1,SOF_GetP7SignDataInfo2));
}
function SOF_SignDataXML()
{
	var SOF_SignDataXML1 = document.getElementById('SOF_SignDataXML1').value;
	var SOF_SignDataXML2 = document.getElementById('SOF_SignDataXML2').value;
	document.getElementById('SOF_SignDataXML3').value = ezca.SOF_SignDataXML(SOF_SignDataXML1,SOF_SignDataXML2);
	document.getElementById('SOF_VerifySignedDataXML1').value = document.getElementById('SOF_SignDataXML3').value;
	document.getElementById('SOF_GetXMLSignatureInfo1').value = document.getElementById('SOF_SignDataXML3').value;
}
function SOF_VerifySignedDataXML()
{
	var SOF_VerifySignedDataXML1 = document.getElementById('SOF_VerifySignedDataXML1').value;
	alert(ezca.SOF_VerifySignedDataXML(SOF_VerifySignedDataXML1));
}
function SOF_GetXMLSignatureInfo()
{
	var SOF_GetXMLSignatureInfo1 = document.getElementById('SOF_GetXMLSignatureInfo1').value;
	var SOF_GetXMLSignatureInfo2 = document.getElementById('SOF_GetXMLSignatureInfo2').value;
	alert(ezca.SOF_GetXMLSignatureInfo(SOF_GetXMLSignatureInfo1,SOF_GetXMLSignatureInfo2));
}
function SOF_CheckSupport()
{
	alert(ezca.SOF_CheckSupport()==0?'支持':'不支持');
}
function SOF_GenRandom()
{
	alert(ezca.SOF_GenRandom(document.getElementById('SOF_GenRandom1').value));
}