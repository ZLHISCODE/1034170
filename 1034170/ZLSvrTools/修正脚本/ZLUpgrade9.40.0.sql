--ZLHIS��ZLBH�ں��漰����
Alter Table zlTools.zlPrograms Add ��ԴID Varchar2(36)
/

Create Or Replace Function zlTools.f_NewID Return Varchar2
--���ܣ����36λGUID��
 Is
  Guid Varchar(36);
Begin
  Guid := Sys_Guid();
  Guid := Substr(Guid, 1, 8) || '-' || Substr(Guid, 9, 4) || '-' || Substr(Guid, 13, 4) || '-' || Substr(Guid, 17, 4) || '-' ||
          Substr(Guid, 21, 12);
  Return Guid;
End f_NewID;
/

Grant Execute On zlTools.f_NewID To Public
/
Create Public Synonym f_Newid For zlTools.f_NewID
/