--ZLHIS与ZLBH融合涉及调整
Alter Table zlTools.zlPrograms Add 资源ID Varchar2(36)
/

Create Or Replace Function zlTools.f_NewID Return Varchar2
--功能：获得36位GUID串
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