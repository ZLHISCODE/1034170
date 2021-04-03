
ALTER TABLE zlNoticeRec ADD 已读标志 NUMBER(1) DEFAULT 0
/
ALTER TABLE zlNoticeRec Modify 提醒内容	VARCHAR2(1000)
/

drop Public Synonym zlNotices
/
drop Public Synonym zlNoticeUsr
/
drop Public Synonym zlNoticeRec
/
drop Public Synonym ZL_ZLNOTICES_INSERT
/
drop Public Synonym ZL_ZLNOTICES_UPDATE
/
drop Public Synonym ZL_ZLNOTICES_DELETE
/
drop Public Synonym ZL_ZLNOTICEUSR_INSERT
/
drop Public Synonym ZL_ZLNOTICEUSR_DELETE
/
--公共同义词
Create Public Synonym zlNotices for zlNotices
/
Create Public Synonym zlNoticeUsr for zlNoticeUsr
/
Create Public Synonym zlNoticeRec for zlNoticeRec
/
Create Public Synonym ZL_ZLNOTICES_INSERT for ZL_ZLNOTICES_INSERT
/
Create Public Synonym ZL_ZLNOTICES_UPDATE for ZL_ZLNOTICES_UPDATE
/
Create Public Synonym ZL_ZLNOTICES_DELETE for ZL_ZLNOTICES_DELETE
/
Create Public Synonym ZL_ZLNOTICEUSR_INSERT for ZL_ZLNOTICEUSR_INSERT
/
Create Public Synonym ZL_ZLNOTICEUSR_DELETE for ZL_ZLNOTICEUSR_DELETE
/

--权限
Grant select,update on zlNotices to PUBLIC
/
Grant select on zlNoticeUsr to PUBLIC
/
Grant select,insert,delete on zlNoticeRec to PUBLIC
/
Grant execute on ZL_ZLNOTICEUSR_INSERT to PUBLIC
/
Grant execute on ZL_ZLNOTICEUSR_DELETE to PUBLIC
/
Grant execute on ZL_ZLNOTICES_INSERT to PUBLIC
/
Grant execute on ZL_ZLNOTICES_UPDATE to PUBLIC
/
Grant execute on ZL_ZLNOTICES_DELETE to PUBLIC
/
