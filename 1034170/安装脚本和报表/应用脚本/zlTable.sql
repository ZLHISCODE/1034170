--����Ŀ¼
--1.��������,2.ҽ������,3.���˲�������,4.���û���,5.ҩƷ���Ļ���
--6.�ٴ�����,7.�ٴ�·������,8.��������,9.�������,10.�������
--11.������,12.ҽ��ҵ��,13.���˲���ҵ��,14.����ҵ��,15.ҩƷ����ҵ��
--16.�ٴ�ҽ��,17.�ٴ�·��,18.����ҵ��,19.����ҵ��,20.����ҵ��,21.���ҵ��

----------------------------------------------------------------------------
--[[1.��������]]
----------------------------------------------------------------------------
Create Table �ڼ��(
    �ڼ� VARCHAR2(6),
    ��ʼ���� Date,
    ��ֹ���� Date)
    TABLESPACE zl9BaseItem;

Create Table ʱ���(
    ʱ��� VARCHAR2(4),
    ��ʼʱ�� date,
    ��ֹʱ�� date,
    ȱʡʱ�� DATE )
    TABLESPACE zl9BaseItem;

Create Table ������Ʊ�(
    ��Ŀ��� NUMBER(3),
    ��Ŀ���� VARCHAR2(20) Not Null,
    ������ VARCHAR2(20),
    �Զ���ȱ NUMBER(1) Default 0,
		��Ź��� NUMBER(1))
    TABLESPACE zl9BaseItem
    PCTFREE 20 initrans 20
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE ���Һ����(
	��Ŀ��� NUMBER(3),
	����ID NUMBER(18),
	��� VARCHAR2(1),
	������ VARCHAR2(20))
    TABLESPACE zl9BaseItem
    PCTFREE 20 initrans 20
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE �����ű�(
	��Ŀ��� NUMBER(3),
	����ǰ׺ VARCHAR2(20),
	���� DATE,
	������ VARCHAR2(20))
    TABLESPACE zl9BaseItem
    PCTFREE 20 initrans 20
    Cache Storage(Buffer_Pool Keep);

Create Table ��λ���Ʒ���(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(6),
	  ���� VARCHAR2(5),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ���ݻ��ڿ���(
    ���� number(2),
    ����  number(2),
    ���� VARCHAR2(100))
    TABLESPACE zl9BaseItem;

Create Table ��Լ��λ(
    ID NUMBER(18),
    �ϼ�id NUMBER(18),
    ���� VARCHAR2(10) Not Null,
    ���� VARCHAR2(100) Not Null,
    ���� VARCHAR2(10),
    ĩ�� NUMBER(1) default 0,
    ��ַ VARCHAR2(100),
    �绰 VARCHAR2(16),
    �������� VARCHAR2(50),
    �ʺ� VARCHAR2(20),
    ��ϵ�� VARCHAR2(20),
    �����ʼ� varchar2(50),
    ˵�� varchar2(2000),
	  վ�� Varchar2(1),
    ����ʱ�� Date,
    ����ʱ�� Date)
    TABLESPACE zl9BaseItem;

Create Table ����Ŀ¼(
		��� NUMBER(5),
		���� VARCHAR2(100),
		˵�� VARCHAR2(200),
		���� NUMBER(1),
		������ VARCHAR2(50))
		TABLESPACE zl9BaseItem;

Create Table ��������(
		���� NUMBER(5),
		������ NUMBER(5),
		������ VARCHAR2(50),
		����ֵ VARCHAR2(100),
		ȱʡֵ VARCHAR2(100),
		����˵�� VARCHAR2(200))
		TABLESPACE zl9BaseItem;

Create Table ��������(
    ��� NUMBER(18),
    ��� NUMBER(3),
    ���� VARCHAR2(255))
    TABLESPACE zl9BaseItem;

Create Table �ƶ��鷿�ӿ�(
    ��� Number(2),
    ���� Varchar2(20),
		���� Varchar2(1000),
    ���� Number(1))
    TABLESPACE zl9BaseItem;

Create Table �������(
	��� VARCHAR2(10),
	���� VARCHAR2(4),
	�ַ� VARCHAR2(2))
	TABLESPACE zl9BaseItem;

CREATE TABLE �ŶӺ����(
	�������� DATE ,
	�Ŷ����� VARCHAR2 (100),
	������  VARCHAR2(20))
TABLESPACE zl9BaseItem;

Create Table �ŶӽкŶ��� (
    ID NUMBER(18),
    ����ID NUMBER(18),
    �������� VARCHAR2(60),
    ҵ��id Number(18),
    ����ID NUMBER(18),
    �ŶӺ��� varchar2(20),
    �Ŷӱ�� VARCHAR2(10),
    �������� VARCHAR2(100),
    ���� VARCHAR2(20),
    ҽ������ VARCHAR2(64),
    ���� NUMBER(1),
    ������� Number(18),
    �Ŷ�ʱ�� DATE,
    �Ŷ�״̬ NUMBER(1) default 0,
    �Ƿ��ʱ�� number(1) DEFAULT 0,
    ����ҽ�� VARCHAR2(20),
    ҵ������ number(5),
    ����ʱ�� date,
    ��ע Varchar2(64),
    �Ŷ���� Varchar2(30))
    TABLESPACE zl9BaseItem
;

Create Table �Ŷ�����ԭ��
(
  ���� VARCHAR2(5),
  ���� VARCHAR2(64),
  ���� VARCHAR2(20),
  ʹ��Ƶ�� number(5) default 0
)tablespace ZL9BASEITEM;

create table �Ŷ���������
(
  ID       NUMBER(18),
  �������� VARCHAR2(1000),
  ����ID   NUMBER(18),
  �������� VARCHAR2(20),
  ҵ������ number(5),
  ����ʱ�� Date,
  վ�� VARCHAR2(50))
  TABLESPACE zl9BaseItem
  PCTFREE 20 initrans 100;

create table �Ŷ�LED��ʾ����
(
  �������� NUMBER(3),
  ������   VARCHAR2(20),
  ����     NUMBER(1),
  ˵��     VARCHAR2(50))
  TABLESPACE zl9BaseItem;

Create Table �������ʷ���(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(6),
    ������ NUMBER(1),
    ˵�� VARCHAR2(200))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table ��Ա���ʷ���(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(6),
    ˵�� VARCHAR2(200))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table ���Ż������(
	���� varchar2(3),
	���� varchar2(10) not Null,
	���� varchar2(10),
	��Χ varchar2(100))
	TABLESPACE zl9BaseItem;

Create Table ���ű�(
    ID NUMBER(18),
    �ϼ�id NUMBER(18),
    ���� VARCHAR2(10) Not Null,
    ���� VARCHAR2(20) Not Null,
    ���� VARCHAR2(5),
    λ�� VARCHAR2(50),
    ĩ�� NUMBER(1) default 0,
    ����ʱ�� Date,
    ����ʱ�� Date,
    ������� VARCHAR2(10),
    ���Ÿ����� number(18),
    վ�� Varchar2(1))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table ��Ա��(
    ID NUMBER(18),
    ��� VARCHAR2(6) Not Null,
    ���� VARCHAR2(20),
    ���� VARCHAR2(8),
    ���֤�� VARCHAR2(18),
    �������� DATE,
    �Ա� VARCHAR2(4),
    ���� VARCHAR2(20),
    �������� DATE,
    �칫�ҵ绰 VARCHAR2(20),
    �����ʼ� VARCHAR2(20),
    ִҵ��� VARCHAR2(3),
    ִҵ��Χ VARCHAR2(20),
    ִҵ֤�� Varchar2(50),
    ����ְ�� VARCHAR2(30),
    רҵ����ְ�� VARCHAR2(50),
    Ƹ�μ���ְ�� NUMBER(1),
    ѧ�� VARCHAR2(10),
    ��ѧרҵ VARCHAR2(2),
    ��ѧʱ�� NUMBER(2),
    ��ѧ���� VARCHAR2(10),
    ������ѵ VARCHAR2(10),
    ���п��� VARCHAR2(10),
    ���˼�� VARCHAR2(1000),
    ����ʱ�� Date,
    ����ʱ�� Date,
    ����ԭ�� Varchar2(100),
    ���� Varchar2(100),
    ǩ�� varchar2(20),
    ǩ��ͼƬ Long Raw,
    �ʸ�֤��� varchar2(50),
    ִҵ��ʼ���� date,
    ����Ȩ��־ number(1),
    �����ȼ� Varchar2(5),
    վ�� Varchar2(1),
    �ƶ��绰 number(11))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep)
;

Create Table ��Ա��Ƭ(
    ��ԱID NUMBER(18),
    ��Ƭ LONG RAW)
    TABLESPACE zl9BaseItem
    PCTFREE 20;

Create Table ��Ա֤���¼(
    ID NUMBER(18),
    ��ԱID NUMBER(18),
    CertDN VARCHAR2(300),
    CertSN VARCHAR2(100),
    SIGNCERT VARCHAR2(3000),
    EncCert VARCHAR2(2000),
    ע��ʱ�� DATE,
    �Ƿ�ͣ�� Number(1),
    ͣ�ü�¼ XMLType)
    TABLESPACE zl9BaseItem
;

Create Table ������Ա(
    ����id NUMBER(18),
    ��Աid NUMBER(18),
    ȱʡ NUMBER(1))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �ϻ���Ա��(
    �û��� VARCHAR2(20),
    ��Աid NUMBER(18))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table ��������˵��(
    �������� VARCHAR2(10),
    ����id NUMBER(18),
    ������� NUMBER(3))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table ��Ա����˵��(
    ��ԱID NUMBER(18),
    ��Ա���� VARCHAR2(10))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table ���Ű���(
    ����id NUMBER(18),
    ���� NUMBER(1),
    ��ʼʱ�� date,
    ��ֹʱ�� date)
    TABLESPACE zl9BaseItem;

Create Table �������Ҷ�Ӧ(
    ����id NUMBER(18),
    ����id NUMBER(18))
    TABLESPACE zl9Patient    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE ����ǩ�����ò���(
    ����ID NUMBER(18),
    ���� NUMBER(5))
    TABLESPACE ZL9BASEITEM
    Cache Storage(Buffer_Pool Keep);

Create Table ����ְ��(
    ���� VARCHAR2(2),
    ���� VARCHAR2(30),
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem;

Create Table ִҵ���(
    ���� VARCHAR2(3),
    ���� VARCHAR2(20),
    ���� VARCHAR2(8),
    ���� VARCHAR2(16))
    TABLESPACE zl9BaseItem;

Create Table רҵ����ְ��(
    ���� VARCHAR2(3),
    ���� VARCHAR2(50),
    ���� VARCHAR2(10),
    �Ƿ�ѡ�� NUMBER(1))
    TABLESPACE zl9BaseItem;

Create Table ҵ����Ϣ����(
    ���� VARCHAR2(100),
    ���� VARCHAR2(100),
    ˵�� VARCHAR2(4000)
 ) TABLESPACE zl9BaseItem;

----------------------------------------------------------------------------
--[[2.ҽ������]]
----------------------------------------------------------------------------
Create Table �������(
    ��� NUMBER(3),
    ���� VARCHAR2(20),
    ˵�� VARCHAR2(100),
    ҽԺ���� VARCHAR2(20),
    �Ƿ�̶� NUMBER(1),
    �Ƿ��ֹ NUMBER(1),
    �������� NUMBER(1),
    ҽ������ VARCHAR2(30),
    ��� NUMBER (1),
    ��Ŀ��ʾ NUMBER (1) DEFAULT 0,
    ҽ���� varchar2(20))
    TABLESPACE zl9BaseItem;

Create Table ��������Ŀ¼(
    ���� NUMBER(3),
    ��� NUMBER(5),
    ���� VARCHAR2(6),
    ���� VARCHAR2(20))
    TABLESPACE zl9BaseItem;

Create Table ���ղ���(
    ���� NUMBER(3),
    ����  NUMBER(5),
    ������ VARCHAR2(20),
    ����ֵ VARCHAR2(40),
    ��� NUMBER(2),
    �Ƿ�̶� NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE ������Ⱥ(
    ���� NUMBER(3),
    ��� NUMBER(1),
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem;

Create Table ���շ��õ�(
    ���� NUMBER(3),
    ���� NUMBER(5),
    ���� NUMBER(3),
    ���� VARCHAR2(25),
    ���� NUMBER(16,5),
    ���� NUMBER(16,5))
    TABLESPACE zl9BaseItem;

Create Table ���������(
	���� NUMBER(3),
	���� NUMBER(5),
	��ְ NUMBER(1),
	����� NUMBER(3),
	���� VARCHAR2(20),
	���� NUMBER(3),
	���� NUMBER(3),
	ȫ��ͳ�� number(1),
	������ number(1),
	�޷ⶥ�� number(1))
    TABLESPACE zl9BaseItem;

Create Table ����֧������(
	���� NUMBER(3),
	���� NUMBER(5),
	��ְ NUMBER(1),
	����� NUMBER(3),
	���� NUMBER(3),
	��� NUMBER(4),
	���� NUMBER(16,5))
    TABLESPACE zl9BaseItem;

Create Table ����֧���޶�(
	���� NUMBER(3),
	���� NUMBER(5),
	��� NUMBER(4),
	���� VARCHAR2(1),
	��� NUMBER(16,5))
    TABLESPACE zl9BaseItem;

Create Table ������Ŀ(
    ���� NUMBER(3),
    ���� VARCHAR2(20),
    ���� VARCHAR2(100),
    ���� VARCHAR2(30),
    ������� VARCHAR2(6),
    ��ע VARCHAR2(50))
    TABLESPACE zl9BaseItem;

Create Table ����֧������(
	���� NUMBER(3),
	ID NUMBER(18),
	���� VARCHAR2(6),
	���� VARCHAR2(40),
	���� VARCHAR2(10),
	���� NUMBER(3),
	�㷨 NUMBER(3),
	ͳ��ȶ� NUMBER(16,5),
	��׼���� NUMBER(16,5),
	��׼���� NUMBER(5),
	������� NUMBER(1),
	�Ƿ�ҽ�� NUMBER(1) DEFAULT 1)
    TABLESPACE zl9BaseItem;

Create TABLE ���൵�α���(
	����ID		number(18),
	����		number(3),
	����		number(16,5),
	����            number(16,5),
	����		number(16,5))
	TABLESPACE	ZL9BASEITEM;

CREATE TABLE ҽ���������(
    ���� NUMBER(3),
	���� NUMBER(1),		--default=0����ʾȱʡ
	���� VARCHAR2(20),
    ˵�� VARCHAR2(200))
    TABLESPACE ZL9BASEITEM;

CREATE TABLE ҽ��������ϸ(
    ���� NUMBER(3),
	  ��� NUMBER(1) DEFAULT 0,		--default=0����ʾȱʡ
	  �շ�ϸĿID NUMBER(18),
    ��Ŀ���� VARCHAR2(20),
	  ˵�� VARCHAR2(256))
    TABLESPACE ZL9BASEITEM;

Create Table ���ղ���(
	���� NUMBER(3),
	ID NUMBER(18),
	���� VARCHAR2(6),
	���� VARCHAR2(100),
	���� VARCHAR2(10),
	��� VARCHAR2(1),
	����ⶥ�� NUMBER(1),
	�ⶥ�߽�� NUMBER(16,5))
    TABLESPACE zl9BaseItem;

Create Table ����֧����Ŀ(
    ���� NUMBER(3),
    �շ�ϸĿID NUMBER(18),
    ����ID NUMBER(18),
    ��Ŀ���� VARCHAR2(20),
    ��Ŀ���� VARCHAR2(100),
    ��ע VARCHAR2(50),
	�Ƿ�ҽ�� NUMBER(1) DEFAULT 1,
	Ҫ������ NUMBER (1))
    TABLESPACE zl9BaseItem;

Create Table ������׼��Ŀ(
	����ID NUMBER(18),
	�շ�ϸĿID NUMBER(18),
	���� NUMBER(1) DEFAULT 0,
	���� NUMBER(1) DEFAULT 0)
    TABLESPACE zl9BaseItem;

----------------------------------------------------------------------------
--[[3.���˲�������]]
----------------------------------------------------------------------------
Create Table סԺ����ԭ��(
    ���� VARCHAR2(2),
    ���� VARCHAR2(50),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem ;

Create Table ����(
    ����	VARCHAR2(6),
    �ϼ����� VARCHAR2(6),
    ����	VARCHAR2(30),
    ����	VARCHAR2(10),
    ����	number(2),
    ȱʡ��־    number(2) DEFAULT 0)
    TABLESPACE zl9BaseItem;

Create Table ����(
    ���� VARCHAR2(8),
    ���� VARCHAR2(50),
    ���� VARCHAR2(20),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ����(
	���� VARCHAR2(3),
	���� VARCHAR2(30),
	Ӣ�ļ�� varchar2(30),
	���� VARCHAR2(10),
	ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ����״��(
    ���� VARCHAR2(1),
    ���� VARCHAR2(4),
    ���� VARCHAR2(4),
    ȱʡ��־ NUMBER(1) default 0,
	������� varchar2(2))
    TABLESPACE zl9BaseItem;

Create Table ����(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1) default 0,
	������� varchar2(2),
	����ƴд�� varchar2(20),
	��ĸ���� varchar2(10))
    TABLESPACE zl9BaseItem;

Create Table ����ϵ(
    ���� VARCHAR2(2),
    ���� VARCHAR2(30),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ���(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(10),
    ���ȼ� NUMBER(1))
    TABLESPACE zl9BaseItem;

Create Table �Ա�(
    ���� VARCHAR2(1),
    ���� VARCHAR2(4),
    ���� VARCHAR2(4),
    ȱʡ��־ NUMBER(1) default 0,
    ������� varchar2(2))
    TABLESPACE zl9BaseItem;

Create Table ѧ��(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1) default 0,
    ������� varchar2(2))
    TABLESPACE zl9BaseItem;

Create Table Ѫ��(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(2),
    ȱʡ��־ NUMBER(1) default 0,
	  ������� varchar2(2))
    TABLESPACE zl9BaseItem;

Create Table ְҵ(
    ���� VARCHAR2(3),
    ���� VARCHAR2(80),
    ���� VARCHAR2(10),
    �������� VARCHAR2(20),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ��������(
    ���� VARCHAR2(2),
    ���� VARCHAR2(50),
    ���� VARCHAR2(10),
    ��ɫ Number(18),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ���䷽ʽ(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1) Default 0)
    TABLESPACE zl9BaseItem;

Create Table ̥��״��(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1) Default 0)
    TABLESPACE zl9BaseItem;

--������ѯ
----------------------------------------------------------------------------
Create Table ��ѯ����ѡ��(
	��� number(3),
	���� varchar2(30),
	���� varchar2(250))
	TABLESPACE zl9BaseItem;

Create Table ��ѯ���Ԫ��(
	��� number(5),
	���� varchar2(30),
	���� number(2),
	�п� varchar2(250),
	���� number(2),
	�и� varchar2(250),
	�ϲ��� varchar2(250),
	�ϲ��� varchar2(250),
	��ɫ number(18))
	TABLESPACE zl9BaseItem;

Create Table ��ѯ�������(
	��� number(5),
	�к� number(2),
	�к� number(2),
	���� varchar2(200),
	���� number(1),
	��ɫ number(18),
	���� varchar2(200))
	TABLESPACE zl9BaseItem;

Create Table ��ѯͼƬԪ��(
	��� number(5),
	���� number(2),
	���� varchar2(30),
	���� number(3),
	ͼ�� Long Raw,
	��� number(18),
	�߶� number(18),
	�̶� numeric(1) default 0,
	�޸����� Date)
	TABLESPACE zl9BaseItem;

Create Table ��ѯҳ��Ŀ¼(
	ҳ����� number(18),
	�ϼ���� number(18),
	���� varchar2(10),
	ҳ������ varchar2(30),
	���� varchar2(15),
	�̶�ҳ�� number(1),
	ҳ���� number(3),
	�������� number(5),
	ҳ�汳�� number(5),
	�������� number(18),
	������� Varchar2(100),
	ĩ�� number(1))
	TABLESPACE zl9BaseItem;

Create Table ��ѯҳ������(
	��� number(5),
	����� number(5),
	���� varchar2(30),
	ҳ�� number(18),
	ҳ��ͼ�� number(5),
	���� varchar2(20),
	��С Number(2),
	���� Number(1),
	��ɫ Number(18))
	TABLESPACE zl9BaseItem;

Create Table ��ѯ����Ŀ¼(
	ҳ����� number(18),
	������� number(5),
	�����ı� varchar2(30),
	����ͼ�� number(5),
	�������� number(1),
	����λ�� number(1),
	�������� varchar2(50),
	����ҳ�� number(1),
	�����ı� long,
	�������� varchar2(50),
	������ number(5),
	���λ�� number(1),
	��ͼ��� number(5),
	��ͼλ�� number(1))
	TABLESPACE zl9BaseItem;

Create Table ��ѯ��������(
	ҳ����� number(18),
	������� number(5),
	����ҳ�� number(18),
	ҳ�ڶκ� number(18))
	TABLESPACE zl9BaseItem;

Create Table ��ѯר���嵥(
	��� number(5),
	��Աid number(18),
	����id number(18))
	TABLESPACE zl9BaseItem;

Create Table ��ѯ�������(
	��� number(5),
	ͼƬ��� number(5))
	TABLESPACE zl9BaseItem;

--������ҳ
----------------------------------------------------------------------------
Create Table ICU����
(
  ���� VARCHAR2(20),
  ���� VARCHAR2(30),
  ���� VARCHAR2(20)
)
TABLESPACE zl9BaseItem;

CREATE TABLE ��Ⱦ��λ(
  ���� VARCHAR2(6),
  ���� VARCHAR2(20),
  ���� VARCHAR2(10),
  ȱʡ��־ NUMBER(1) default 0
)TABLESPACE zl9BaseItem;

Create Table ��е����Ŀ¼
(
  ���� VARCHAR2(20),
  ���� VARCHAR2(30),
  ���� VARCHAR2(20)
)
TABLESPACE zl9BaseItem;

Create Table ҽԺ��ȾĿ¼
(
  ���� VARCHAR2(20),
  ���� VARCHAR2(30),
  ���� VARCHAR2(20)
)
TABLESPACE zl9BaseItem;

Create Table ��ԭѧĿ¼
(
  ���� VARCHAR2(20),
  ���� VARCHAR2(50),
  ���� VARCHAR2(20)
)
TABLESPACE zl9BaseItem;

CREATE TABLE  ҽѧ��ʾ (
	���� VARCHAR2(4),
	���� varchar2(20),
	���� varchar2(10) ,
	ȱʡ��־ NUMBER (1) DEFAULT 0) 
	TABLESPACE zl9BaseItem;

CREATE TABLE  ֤������ (
	���� VARCHAR2(4),
	���� varchar2(20),
	���� varchar2(10), 
	ȱʡ��־ NUMBER (1) DEFAULT 0) 
	TABLESPACE zl9BaseItem;

CREATE TABLE ������Ŀ(
	���� varchar2(3),
	���� varchar2(20),
	���� varchar2(1000))
	TABLESPACE zl9CisRec 
	Cache Storage(Buffer_Pool Keep);

Create Table ��������(
    ����	Varchar2(10),
    ����	Varchar2(30),
    ����	Varchar2(30),
    ȱʡ��־	Number(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ������Ŀ(
    ���� VARCHAR2(4),
    �ϼ� VARCHAR2(4),
    ���� VARCHAR2(30),
    ���� VARCHAR2(10),
    ĩ�� NUMBER(1) DEFAULT 0)
    TABLESPACE zl9BaseItem;

Create Table �������(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(6),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ����(
    ���� VARCHAR2(1),
    ���� VARCHAR2(20),
    ���� VARCHAR2(4),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table סԺ�����ڼ�(
    ���� VARCHAR2(2),
    ���� VARCHAR2(50),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem ;

Create Table ��Ժ��ʽ(
    ���� VARCHAR2(1),
    ���� VARCHAR2(8),
    ���� VARCHAR2(4),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ��Ժ����(
    ���� VARCHAR2(1),
    ���� VARCHAR2(8),
    ���� VARCHAR2(4),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ����ȥ��(
    ���� VARCHAR2(1),
    ���� VARCHAR2(20),
    ���� VARCHAR2(4),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ��Ժ��ʽ(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ��Ժת��(
    ���� VARCHAR2(5),
    ���� VARCHAR2(100),
    ���� VARCHAR2(50),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ���ƽ��(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table סԺĿ��(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(6),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ��Ⱦ����(
  ���� VARCHAR2(3),
  ���� VARCHAR2(100))
  TABLESPACE zl9BaseItem;

Create Table �����¼�(
       ���� VARCHAR2(3),
       ���� VARCHAR2(100))
  TABLESPACE zl9BaseItem;

Create Table �����п�����(
       ���� VARCHAR2(2),
       ���� VARCHAR2(100))
  TABLESPACE zl9BaseItem;

Create Table ҽ�����(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(6),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ��������(
	���� VARCHAR2(2),
	���� VARCHAR2(20))
	TABLESPACE zl9BaseItem;

Create Table ���Ȳ������(
���� Varchar2(2),
���� Varchar2(50),
���� Varchar2(10))
TABLESPACE ZL9BASEITEM 
Cache Storage(Buffer_Pool Keep);

Create Table �ֻ��̶�(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE ZL9BASEITEM;


--�������
----------------------------------------------------------------------------
Create Table �����������(
	���� VARCHAR2(1),
	��� VARCHAR2(20),
	˵�� VARCHAR2(50),
	���ȼ� NUMBER(3),
	�Ƿ���� NUMBER(1) default 1)
	TABLESPACE zl9BaseItem;

Create Table �����������(
	ID NUMBER(18),
	�ϼ�ID NUMBER(18),
	��� NUMBER(6),
	���� VARCHAR2(150),
	���� VARCHAR2(20),
	��� VARCHAR2(1),
	���뷶Χ  varchar2(60),
	�Ƿ��� NUMBER(1) default 1)
	TABLESPACE zl9BaseItem;

Create Table ��������Ŀ¼(
	ID number(18),
	���� VARCHAR2(20),
	��� NUMBER(3),
	���� VARCHAR2(15),
	ͳ���� VARCHAR2(10),
	���� VARCHAR2(150) not Null,
	���� VARCHAR2(20),
	����� Varchar2(20),
	˵�� VARCHAR2(200),
	�Ա����� VARCHAR2(4),
	��Ч���� VARCHAR2(4),
	�������� VARCHAR2(20),
	���� VARCHAR2(1),
	��� VARCHAR2(1),
	����ID NUMBER(18),
	����ʱ�� DATE ,
	����ʱ�� DATE)
	TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �����������(
	����ID Number(18),
	����ID Number(18),
	��ԱID Number(18))
	TABLESPACE zl9BaseItem;

Create Table �������ֶ�Ӧ(
    ����ID NUMBER(18),
    ���� NUMBER(3),
    ����ID NUMBER(18))
    TABLESPACE zl9BaseItem;

Create Table ����������(
    ���� VARCHAR2(2),
    ���� VARCHAR2(30),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE ZL9BASEITEM;

----------------------------------------------------------------------------
--[[4.���û���]]
----------------------------------------------------------------------------
Create Table ��������(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(6),
    ��� NUMBER(16,2))
    TABLESPACE zl9BaseItem;

CREATE TABLE ����Ԥ��ժҪ(
  ���� varchar2(4),
  ���� varchar2(50),
  ���� VARCHAR2(20),
  ȱʡ��־ number(1))
  TABLESPACE zl9BaseItem;
  
Create Table ���ùҺ�ժҪ
(
  ����   VARCHAR2(4),
  ����   VARCHAR2(50),
  ����   VARCHAR2(25),
  ȱʡ��־ NUMBER(1))
Tablespace ZL9BASEITEM;

CREATE TABLE �����˷�ԭ��(
	���� varchar2(4),
	���� varchar2(50),
	���� VARCHAR2(20),
	ȱʡ��־ number(2)
	) TABLESPACE zl9BaseItem;

Create Table ҽ�۽ӿ�(
    ��� number(2),
    ���� Varchar2(20),
    ҽ�� number(1), --�Ƿ�֧��ҽ�Ƽ�Ŀ����
    ҩƷ number(1), --�Ƿ�֧��ҩƷ��Ŀ����
    ���� number(1), --�Ƿ�֧�����ļ�Ŀ����
    ѡ�� number(1)) --��ǰѡ��ʹ�ñ�־
    TABLESPACE zl9BaseItem;

Create Table ��׼ҽ�۹淶(
    ��Ŀ����	varchar2(20),
    ��Ŀ����	varchar2(200),
    ƴ����   varchar2(10),
    ��Ŀ���� varchar2(100),
    �Ƽ۵�λ varchar2(200),
    ��Ŀ�ں� varchar2(1000),
    �������� varchar2(1000),
    ��Ŀ˵�� varchar2(1000),
    ��Ŀ�۸� number(20,2),
    �ظ���־ char(1),
    ҽԺ�ȼ� char(1),
    ע����־ char(1),
    ������� char(1),
    ����޼� number(20,2),
    ����޼� number(20,2),
    �������� Date)
    TABLESPACE zl9BaseItem;

Create Table Ʊ��ʹ�����(
    ���� varchar2(3),
    ���� VARCHAR2(50),
    ����  VARCHAR2(25),
    ȱʡ��־ NUMBER (1) DEFAULT 0)
    TABLESPACE zl9BaseItem;

Create Table ���㳡��(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table ���㷽ʽ(
    
 ���� VARCHAR2(2),
    
 ���� VARCHAR2(20),
    
 ���� VARCHAR2(4),
    
 ���� NUMBER(2),
    
 Ӧ�տ� NUMBER(1),
    
 Ӧ���� NUMBER(1),
    
 ȱʡ��־ NUMBER(1) default 0,

 �Ƿ�̶� number(1) default 0)
    
 TABLESPACE zl9BaseItem;

Create Table ���㷽ʽӦ��(
    Ӧ�ó��� VARCHAR2(10),
    ���㷽ʽ VARCHAR2(20),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ҽ�Ƹ��ʽ(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1) default 0,
    �Ƿ�ҽ�� number(1) default 0,
    �Ƿ񹫷� number(1)  default 0)
    TABLESPACE zl9BaseItem;

Create Table ��������(
	���� VARCHAR2(2),
	���� VARCHAR2(20),
	���� VARCHAR2(10),
	���� VARCHAR2(1),
	ȱʡ��־ NUMBER(2) default 0)
	TABLESPACE zl9BaseItem;

Create Table �ѱ�(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10) Not Null,
    ���� VARCHAR2(4),
	��Ч��ʼ DATE,
	��Ч���� DATE,
	���ÿ��� NUMBER(1),
	����     NUMBER(1),
	���޳��� NUMBER(1),
    ȱʡ��־ NUMBER(1) default 0,
	������� NUMBER(3) default 3,
    ˵�� VARCHAR2(50))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �ѱ���ϸ(
    �ѱ� VARCHAR2(10),
    ������Ŀid NUMBER(18),
    �շ�ϸĿID Number(18),
    �κ� NUMBER(3) default 1,
    Ӧ�ն���ֵ NUMBER(16,5),
    Ӧ�ն�βֵ NUMBER(16,5) default 10000000000,
    ʵ�ձ��� NUMBER(16,5) default 100,
    ���㷽�� NUMBER(1) default 0)
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �ѱ����ÿ���(
    �ѱ� VARCHAR2(10),
    ����id NUMBER(18))
    TABLESPACE zl9BaseItem;

Create Table ������Ŀ(
    ID NUMBER(18),
    �ϼ�id NUMBER(18),
    ���� VARCHAR2(8) Not Null,
    ���� VARCHAR2(20) Not Null,
    ���� VARCHAR2(10),
    ĩ�� NUMBER(1) default 0,
    ���� NUMBER(1) default 0,
    �վݷ�Ŀ VARCHAR2(20),
    ������Ŀ VARCHAR2(30),
    ����ʱ�� Date,
    ����ʱ�� Date)
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �վݷ�Ŀ(
    ���� VARCHAR2(8),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �վݷ�Ŀ��Ӧ(
    ������ĿID NUMBER(18),
    ���� NUMBER(1),
    �վݷ�Ŀ VARCHAR2(20))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE �շ���Ŀ���(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(10),
    �̶� NUMBER(1),
	��� NUMBER(2))
    TableSpace zl9BaseItem;

CREATE TABLE �շѷ���Ŀ¼(
    ID NUMBER(18),
    �ϼ�id NUMBER(18),
    ���� VARCHAR2(15),
    ���� VARCHAR2(40),
    ���� VARCHAR2(10))
    TableSpace zl9BaseItem;

CREATE TABLE �շ���ĿĿ¼(
    ��� VARCHAR2(1),
    ����ID NUMBER(18),
    ID NUMBER(18),
    ���� VARCHAR2(20),
    ���� VARCHAR2(80),
    ��� VARCHAR2(100),
    ���� VARCHAR2(60),
    ���㵥λ VARCHAR2(20),
    ˵�� VARCHAR2(500),
    ��Ŀ���� NUMBER(3),
    �������� VARCHAR2(20),
    ������� NUMBER(3),
    ���ηѱ� NUMBER(1),
    �Ƿ��� NUMBER(1),
    �Ӱ�Ӽ� NUMBER(1),
    ����ժҪ NUMBER(1),
    ����ȷ�� Number(1),
    ִ�п��� NUMBER(3),
    ��ʶ���� VARCHAR2(20),
    ��ʶ���� VARCHAR2(1),
    ��ѡ�� VARCHAR2(20),
    ����޼� NUMBER(20,2),
    ����޼� NUMBER(20,2),
    ����ʱ�� DATE,
    ����ʱ�� DATE,
    ¼������ Number(16,5),
    ���㷽ʽ Number(1),
    վ�� Varchar2(1),
    ����ԭ�� Varchar2(100),
    ͣ��ԭ�� Varchar2(100),
    ������Ŀ varchar2(30))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE �շ���Ŀ����(
    �շ�ϸĿID NUMBER(18),
    ���� VARCHAR2(80),
    ���� NUMBER(3),
    ���� VARCHAR2(40),
    ���� NUMBER(3))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE �շ����ÿ���(
    ��ĿID NUMBER(18),
    ����id NUMBER(18))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE �շ�ִ�п���(
    �շ�ϸĿID NUMBER(18),
    ������Դ NUMBER(3) DEFAULT 1,
    ��������id NUMBER(18),
    ִ�п���id NUMBER(18))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �շѼ�Ŀ(
    ID NUMBER(18),
    ԭ��id NUMBER(18),
    �շ�ϸĿid NUMBER(18),
    ԭ�� NUMBER(16,7),
    �ּ� NUMBER(16,7),
    ȱʡ�۸� NUMBER(16,7),
    ������Ŀid NUMBER(18),
    �Ӱ�Ӽ��� NUMBER(16,5),
    �����շ��� NUMBER(16,5),
    �䶯ԭ�� NUMBER(3) default 1,
    ����˵�� VARCHAR2(100),
    ����ID NUMBER(18),
    ������ VARCHAR2(20),
    ִ������ Date,
    ��ֹ���� Date,
	No VARCHAR2(8),
	��� NUMBER(5),
    ���ۻ��ܺ� Varchar2(10))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �շѴ�����Ŀ(
    ����ID NUMBER(18),
    ����ID NUMBER(18),
    ���д��� NUMBER(2) default 0,
    �������� NUMBER(16,5))
    TABLESPACE zl9BaseItem;

Create Table �շ��ض���Ŀ(
    �ض���Ŀ VARCHAR2(20),
    �շ�ϸĿid NUMBER(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE  ������Ŀ����(
	ID		NUMBER(18),
	�ϼ�ID	Number(18),
	����		Varchar2(10),
	����		Varchar2(50),
	����		Varchar2(20))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE �����շ���Ŀ(
	ID		NUMBER(18),
	����ID	Number(18),
	����		Varchar2(10),
	����		Varchar2(100),
	ƴ��		Varchar2(20),
	���          VARCHAR2(20),
	��Χ		Number(2),
	��ԱID	Number(18),
	��ע         VARCHAR2(100))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE	������Ŀʹ�ÿ���(
	����ID	NUMBER(18),
	����ID	Number(18))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE	�����շ���Ŀ���(
	����ID		NUMBER(18),
	�շ�ϸĿID	Number(18),
	���			Number(18),
	��������		Number(18),
	����                  Number(3),
	����			Number(16,7),
	����			Number(16,7),
	ִ�п���ID	Number(18))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

create table �շ���Ŀ���
(
  ID         NUMBER(18),
  �շ���ĿID NUMBER(18),
  ����       VARCHAR2(80),
  ����       NUMBER(15,6),
  ���       VARCHAR2(100),
  ���㵥λ   VARCHAR2(20),
  ����       NUMBER(18),
  ˵��       VARCHAR2(500)
)
    TABLESPACE zl9CisRec;

Create Table ���ʱ�����(
    ����id NUMBER(18),
    ���ò��� VARCHAR2(20),
    �������� NUMBER(1),
    ����ֵ NUMBER(16,5),
    ������־1 Varchar2(30),
    ������־2 Varchar2(30),
    ������־3 Varchar2(30),
    �߿����� NUMBER(16,5),
    �߿��׼ NUMBER(16,5))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �Զ��Ƽ���Ŀ(
    ����id NUMBER(18),
    �����־ NUMBER(1),
    �շ�ϸĿid NUMBER(18),
    �������� Date)
    TABLESPACE zl9BaseItem;

Create Table ���ݲ�������(
	��ԱID Number(18),
	���� Number(1),
	ʱ������ Number(3),
	���˵��� Number(1),
	������� Number(20,5))
	TABLESPACE zl9BaseItem;


--ҽ�ƿ�
Create Table һ��ͨĿ¼(
	���	Number(2),
	����	Varchar2(50),
	���㷽ʽ	Varchar2(20),
	ҽԺ����	Varchar2(6),
	����	Number(1))
TABLESPACE zl9BaseItem;

CREATE TABLE ҽ�ƿ����(
	ID Number(18),
	����  VarChar2(4),
	���� Varchar2(50),
	���� Varchar2(4),
	ǰ׺�ı� Varchar2(6),
	���ų��� Number(5),
	ȱʡ��־ Number(1),
	�Ƿ�̶� Number(1) DEFAULT 0,
	�Ƿ��ϸ���� Number(1) DEFAULT 0,
	�Ƿ�ˢ�� number(1) DEFAULT 1,
	�Ƿ����� Number(1) DEFAULT 0,	
	�Ƿ�����ʻ� Number(1),
	�Ƿ�����  number(1) DEFAULT 1,
	�Ƿ�ȫ�� Number(1) DEFAULT 0,
	���� Varchar2(100),
	��ע Varchar2(100),
	�ض���Ŀ Varchar2(6),
	���㷽ʽ Varchar2(20),
	�������� VARCHAR2(10),
	�Ƿ��ظ�ʹ�� number(1) DEFAULT 0,
	�Ƿ�����  NUMBER(1) DEFAULT 0,
	���볤��  number(2) DEFAULT 10,
	���볤������  NUMBER(2) DEFAULT 0,
	�������  NUMBER(2),
	�Ƿ�ģ������ NUMBER(1) DEFAULT 0,
	������������ Number(1) DEFAULT 0,
	�Ƿ�ȱʡ���� Number(1) DEFAULT 0,
	�Ƿ��ƿ� NUMBER(1) DEFAULT 0 ,
	�Ƿ񷢿� NUMBER(1) DEFAULT 0,
	�Ƿ�д�� NUMBER(1),
	����     Number(3),
	�������� Number(2) DEFAULT 0
	) 
	TABLESPACE zl9BaseItem;

CREATE TABLE �����ѽӿ�Ŀ¼(
	��� number(6),
	���� varchar2(50),
	ϵͳ number(2),
	���㷽ʽ VARCHAR2(20),
	���� varchar2(100),
	���� number(2),
	���ƿ� number(2),
	ǰ׺�ı� Varchar2(4),
	���ų��� Number(6),
	�Ƿ�����  NUMBER(1) DEFAULT 0,
	�Ƿ�����  number(1) DEFAULT 1,
	�Ƿ�ȫ�� Number(1) DEFAULT 0,
	�Ƿ�ˢ�� number(1) DEFAULT 1,
	���볤��  number(2),
	���볤������ Number(2) DEFAULT 0,
	������� NUMBER(2)
	) TABLESPACE zl9BaseItem;

CREATE TABLE ���ѿ�����(
	���� varchar2(2),
	���� varchar2(20),
	ȱʡ��� number(16,5),
	ȱʡ�ۿ� number(16,5),
	ȱʡ��־ number(2)
	) TABLESPACE zl9BaseItem;

CREATE TABLE ���÷���ԭ��(
	���� varchar2(4),
	���� varchar2(50),
	���� VARCHAR2(20),
	ȱʡ��־ number(2)
	) TABLESPACE zl9BaseItem;

CREATE TABLE ҽ�ƿ���ʧ��ʽ(
	���� Varchar2 (4),
	���� Varchar2 (20),
	���� varchar2(10),
	��Ч���� number (5),
	ȱʡ��־ Number (1))
	TABLESPACE zl9BaseItem;


--�Һ�
Create Table ����(
	���� VARCHAR2(2),
	���� VARCHAR2(10),
	���� VARCHAR2(4),
	ȱʡ��־ NUMBER(1) default 0,
	˵�� VARCHAR2(50))
	TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table ԤԼ��ʽ(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(6),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table �ҺŰ���(
	ID NUMBER(18),
	���� Varchar2(10),
	���� VARCHAR2(5) Not Null,
	����id NUMBER(18),
	��ĿID NUMBER(18),
	ҽ������ VARCHAR2(20),
	ҽ��ID NUMBER(18),
	���	   NUMBER(18), 
	���� VARCHAR2(4),
	��һ VARCHAR2(4),
	�ܶ� VARCHAR2(4),
	���� VARCHAR2(4),
	���� VARCHAR2(4),
	���� VARCHAR2(4),
	���� VARCHAR2(4),
	�������� Number(1),
	���﷽ʽ Number(1),
	��ſ��� Number(1) Default 0,
	��ʼʱ�� Date,
	��ֹʱ�� Date,
	ͣ������ Date ,
	ִ��ʱ�� Date,
	ִ�мƻ�ID Number(18),
	Ĭ��ʱ�μ�� Number(5),
	�Ƿ�ɾ�� Number(1) Default 0
	)
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
	
Create Table �ҺŰ���ʱ��( 
	 ����ID Number(18),
	 ��� Number(18),
    ���� VARCHAR2(10),
	 ��ʼʱ�� Date,
	 ����ʱ�� Date,
	 �������� Number(18),
	 �Ƿ�ԤԼ Number(1) DEFAULT 0)
	 Tablespace zl9BaseItem;
   
Create Table �ҺŰ������� (
	����ID Number(18),
	������Ŀ Varchar2(10),
	�޺��� Number(5),
	��Լ�� Number(5))
	TableSpace zl9BaseItem;

Create Table �Һ����״̬(
    ���� Varchar2(5),
    ���� Date,
    ��� Number(5),
    ״̬ Number(1),
    ����Ա���� Varchar2(20),
    ԤԼ Number(1) Default(0),
    ��ע VARCHAR2(100),
    �Ǽ�ʱ�� DATE,
    ������ Varchar2(50))
    TABLESPACE zl9Patient
;

Create Table �ҺŰ�������(
    �ű�ID NUMBER(18),
    �������� VARCHAR2(20),
	��ǰ���� Number(1))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE �ҺŰ��żƻ�(
	ID	NUMBER(18),
	����ID	NUMBER (18),
    ��ĿID NUMBER(18),
	����	VARCHAR2(5),
	��Чʱ��  DATE,
	ʧЧʱ�� Date DEFAULT To_date('3000-01-01','yyyy-mm-dd'), 
	����	VARCHAR2(4),
	��һ	VARCHAR2(4),
	�ܶ�	VARCHAR2(4),
	����	VARCHAR2(4),
	����	VARCHAR2(4),
	����	VARCHAR2(4),
	����	VARCHAR2(4),
	���﷽ʽ	NUMBER(1),
	��ſ���	NUMBER(1),
	������    VARCHAR2(20),
	����ʱ��  DATE,
	�����    VARCHAR2(20),
	���ʱ��  DATE,
	ҽ������ Varchar2(20),
	ҽ��ID Number(18),
	ʵ����Ч  DATE DEFAULT To_date('3000-01-01','yyyy-mm-dd'),
	Ĭ��ʱ�μ�� Number(5))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);
	
Create  Table 	�Һżƻ�����(
	�ƻ�ID Number(18),
	������Ŀ Varchar2(10),
	�޺��� Number(5),
	��Լ�� Number(5))
	Tablespace zl9BaseItem;
 
Create Table �Һżƻ�ʱ��(
	�ƻ�ID Number(18),
	��� Number(18),
  ���� VARCHAR2(10),
	��ʼʱ�� Date,
	����ʱ�� Date,
	�������� Number(18),
	�Ƿ�ԤԼ Number(1))
	Tablespace zl9BaseItem;
 
CREATE TABLE �ҺŰ���ͣ��״̬(
	����ID	number(18),
	���         number(18),
	��ʼֹͣʱ��  Date,
	����ֹͣʱ��   Date,
	�ƶ���	varchar2(20),  
	�ƶ�����     date,
	��ע           varchar2(100))
    TABLESPACE zl9BaseItem;

CREATE TABLE ����ͣ��ԭ��(
	���� VARCHAR2(5),
	���� VARCHAR2(50),
	���� VARCHAR2(10),
	ȱʡ��־ number (1))
	TABLESPACE zl9BaseItem;

Create Table �Һżƻ�����(
	�ƻ�ID	NUMBER(18),
	�������� VARCHAR2(20))
	TABLESPACE zl9BaseItem	
	Cache Storage(Buffer_Pool Keep);

 
CREATE TABLE �Һź�����λ(
	���� Varchar2 (4),
	���� Varchar2 (50),
	���� varchar2(10),
	ȱʡ��־ number (1))
	TABLESPACE zl9BaseItem;

CREATE TABLE ������λ���ſ���(
	������λ Varchar2 (50),
	����ID number(18),
	������Ŀ varchar2(10),
	��� Number(18),
	���� number(10))
	TABLESPACE zl9BaseItem;

Create Table ������λ�ƻ�����(
	������λ VARCHAR2(50) ,
	�ƻ�ID   NUMBER(18),
	������Ŀ VARCHAR2(10),
	���     NUMBER(18),
	����     NUMBER(10))
	Tablespace zl9BaseItem; 

Create Table �շѼ��ʵ�(
    ID NUMBER(18),
    ��� VARCHAR2(6),
    ���� VARCHAR2(50),
    �շ���Ŀ�� NUMBER(18),
    ���÷�Χ VARCHAR2(4),
    ��� NUMBER(16),
    �߶� NUMBER(16),
	����ɫ	NUMBER(18),
	����	VARCHAR2(50))
    TABLESPACE zl9Expense;

Create Table �շѼ��ʵ�����(
    ����ID NUMBER(18),
    ��Ӧ�ֶ� VARCHAR2(50),
    ��� NUMBER(18),
    ���� VARCHAR2(20),
    ����ֵ VARCHAR2(30),
    ˳��� NUMBER(5),
	���	NUMBER(16),
	����	NUMBER(16),
	���	NUMBER(16),
	�߶�	NUMBER(16),
	����	VARCHAR2(50),
	ǰ��ɫ	NUMBER(18),
	����ɫ	NUMBER(18),
	�Ƿ���ʾ	NUMBER(1),
	����	NUMBER(1),
	�߿���	NUMBER(1),
	͸��	NUMBER(1))
    TABLESPACE zl9Expense;

----------------------------------------------------------------------------
--[[5.ҩƷ���Ļ���]]
----------------------------------------------------------------------------
Create Table ��Һ������ҩƷ(
    ҩƷid number(18),
    ���� VARCHAR2(50))
    TABLESPACE zl9BaseItem;
    
Create Table ��Һ���ȴ�ӡҩƷ(
    ҩƷid number(18),
    ���� VARCHAR2(50))
    TABLESPACE zl9BaseItem;  

Create Table ��ҩ����(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    �ϰ�� NUMBER(2),
    ר�� NUMBER(1),
    ҩ��id NUMBER(18),
    ��ǰ���� NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE ҩƷ�����λ(
	���� VARCHAR2(2),
	���� VARCHAR2(50),
	���� VARCHAR2(10))
	TABLESPACE zl9BaseItem;

CREATE TABLE ҩƷ������λ(
	���� VARCHAR2(2),
	���� VARCHAR2(50),
	���� VARCHAR2(10))
	TABLESPACE zl9BaseItem;

CREATE TABLE ҩƷ�ⷿ��λ(
	���� VARCHAR2(5),
	�ϼ� VARCHAR2(5),
	���� VARCHAR2(50),
	���� VARCHAR2(10),
	ĩ�� NUMBER(1) DEFAULT 0)
	TABLESPACE zl9BaseItem;

Create Table ҩ�۹�����(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ�������(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ��Դ���(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ����(
    ���� VARCHAR2(3),
    ���� VARCHAR2(20) Not Null,
    ���� VARCHAR2(10),
    ����� VARCHAR2(1))
    TABLESPACE zl9BaseItem;

Create Table ��ҩ����(
    ���� VARCHAR2(3),
    ���� VARCHAR2(20) Not Null,
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ���(
    ���� VARCHAR2(3),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10),
    ȱʡ��־ number(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ҩƷ��ֵ����(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ��Դ����(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

create table ����ҩ��˵��
(
  ���� VARCHAR2(2) not null,
  ���� VARCHAR2(30),
  ���� VARCHAR2(10)
)
tablespace ZL9BASEITEM;

Create Table ҩƷ��ҩ�ݴ�(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ������(
    ���� VARCHAR2(10),
    ���� VARCHAR2(60),
    ���� VARCHAR2(10),
    վ�� Varchar2(1))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ������(
    ID NUMBER(18),
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ϵ�� NUMBER(2))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ���ݷ���(
    ���� NUMBER(2),
    ���� VARCHAR2(16),
    ���� NUMBER(1),
    ˵�� VARCHAR2(200))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ��������(
    ���� NUMBER(2),
    ���id NUMBER(18))
    TABLESPACE zl9BaseItem;

Create Table ������ԭ��(
    ���� VARCHAR2(1),
    ���� VARCHAR2(20),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table �������취(
    ���� VARCHAR2(1),
    ���� VARCHAR2(20),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ�������(
    ���ڿⷿID NUMBER(18),
    �Է��ⷿID NUMBER(18),
    ���� NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE ҩƷ���ÿ���(
	���ò���ID NUMBER(18),
	�Է��ⷿID NUMBER(18))
	TABLESPACE zl9BaseItem;

Create Table ҩƷ�ⷿ��λ(
    �ⷿid NUMBER(18),
	���÷�Χ NUMBER(1),	--1-ҩ��;2-����ҩ��;3-סԺҩ��;4-����(�Ƽ���)
    ���� NUMBER(1))  --1-�ۼ۵�λ,2-���ﵥλ,3-סԺ��λ,4-ҩ�ⵥλ
    TABLESPACE zl9BaseItem;

CREATE TABLE ҩ����ҩ����(
    ҩ��id NUMBER(18),
	���� NUMBER(1),		--1-����;2-סԺ
    ��ҩ NUMBER(1),		--1-��ҩ;����������ҩ
	�Զ���ҩ���� Number(3),
	��ҩȷ�� Number(1))
    TABLESPACE zl9BaseItem;

Create Table ��Ӧ��(
	ID		NUMBER(18),
	�ϼ�id		NUMBER(18),
	����		VARCHAR2(8),
	����		VARCHAR2(80),
	����		VARCHAR2(10),
	ĩ��		NUMBER(1) default 0,
	���֤��	VARCHAR2(30),
	���֤Ч��	DATE,
	ִ�պ�		VARCHAR2(30),
	ִ��Ч��	DATE,
	��Ȩ��		VARCHAR2(30),
	��Ȩ��	DATE,
	˰��ǼǺ�	VARCHAR2(30),
	��ַ		VARCHAR2(50),
	�绰		VARCHAR2(16),
	��������	VARCHAR2(50),
	�ʺ�		VARCHAR2(30),
	��ϵ��		VARCHAR2(20),
	����ʱ��	Date,
	����ʱ��	Date,
	����		varchar2(10),
	������		number(6),
	���ö�		number(18,5),
	����ί����	varchar2(20),
	����ί������	date,
	������֤��	varchar2(20),
	������֤����	date,
	ҩ��ֱ�����	varchar2(20),
	ҩ��ֱ�������	date,
	վ��		Varchar2(1),
	��ӪƷ�� Varchar2(200),
	��ע Varchar2(200))
    TABLESPACE zl9BaseItem;

create table ҩƷ�����̶���
(
       ҩƷid number(18),
       �������� VARCHAR2(60),
       ��׼�ĺ� VARCHAR2(40)
) tablespace ZL9BASEITEM;

CREATE TABLE ҩƷ����(
    ҩ��ID NUMBER(18),
    ҩƷ���� VARCHAR2(20),
    ������� VARCHAR2(10),
    ��Դ��� VARCHAR2(10),
    ��ֵ���� VARCHAR2(10),
    ��ҩ�ݴ� VARCHAR2(10),
    ����ҩ�� NUMBER(1),
    �Ƿ���ҩ NUMBER(1),
    �Ƿ�Ƥ�� NUMBER(1),
    �Ƿ�ԭ�� NUMBER(1),
    �������� NUMBER(16,5),
    ����ְ�� VARCHAR2(2),
    ҩƷ���� NUMBER(1),
    Ʒ��ҽ�� NUMBER(1),
    ������ number(1) DEFAULT 0,
    �ٴ��Թ�ҩ number(1),
    ATCCODE varchar2(50),
    �Ƿ�����ҩ number(1),
    ��ý number(1))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE ҩƷ���(
    ҩ��ID NUMBER(18),
    ҩƷid NUMBER(18),
    ����ϵ�� NUMBER(16,5),
    ���ﵥλ VARCHAR2(8),
    �����װ NUMBER(16,5),
    סԺ��λ VARCHAR2(8),
    סԺ��װ NUMBER(16,5),
    ҩ�ⵥλ VARCHAR2(8),
    ҩ���װ NUMBER(16,5),
    ���Ч�� NUMBER(5),
    ҩƷ��Դ VARCHAR2(10),
    Э��ҩƷ NUMBER(1),
    ����ҩƷ NUMBER(1),
    ��׼�ĺ� VARCHAR2(40),
    ע���̱� VARCHAR2(50),
    ��ʶ�� VARCHAR2(29),
    ҩ�ۼ��� VARCHAR2(10),
    ָ�������� NUMBER(16,7),
    ָ�����ۼ� NUMBER(16,7),
    ָ������� NUMBER(16,5),
    ���� NUMBER(16,5),
    סԺ�ɷ���� NUMBER(3),
    ��̬���� NUMBER(1),
    ҩ����� NUMBER(1),
    ҩ������ NUMBER(1),
    �б�ҩƷ NUMBER(1),
    ��������� NUMBER(16,5),
    GMP��֤ NUMBER(1),
    �ɱ��� number(16,7),
    ����ѱ��� NUMBER(16,5),
    ���쵥λ NUMBER(1),
    ���췧ֵ NUMBER(16,5),
    ��ͬ��λID NUMBER(18),
    �ϴι�Ӧ��ID NUMBER(18),
    �ϴβ���     VARCHAR2(60),
    �ϴ�����     VARCHAR2(20),
    �ϴ��������� DATE,
    �ϴ���׼�ĺ� VARCHAR2(40),
    ��ҩ����  VARCHAR2(20),
    ���� NUMBER(16,5),
    ��ֵ˰�� NUMBER(16,5),
    ����ҩ�� varchar2(30),
    ��ҩ��̬ Number(1) Default 0,  -- 0:ɢװ;  1:��ҩ��Ƭ;  2:����
    �Ƿ񳣱� Number(1),
    ����ɷ����   NUMBER(3),
    DDDֵ number(16,5),
    �ϴ��ۼ� number(16,7),
    ��ΣҩƷ number(1),
    �ͻ���λ varchar2(8),
    �ͻ���װ number(16,5)
    ) 
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create table ҩƷ�ӳɷ���( 
  ���  NUMBER(18) not null,
  ��ͼ�  NUMBER(16,5),
  ��߼�  NUMBER(16,5),
  �ӳ���  NUMBER(16,5),
  ��۶� NUMBER(16,5),
  ˵��   VARCHAR2(50))
  tablespace ZL9BASEITEM;

Create Table Э��ҩƷ����(
    ҩƷID NUMBER(18),
    Э��ҩƷID NUMBER(18),
    ���� NUMBER(16,5),
    ��ĸ NUMBER(16,5))
    TABLESPACE zl9BaseItem;

Create Table ����ҩƷ����(
    ����ҩƷID NUMBER(18),
    ԭ��ҩƷID NUMBER(18),
    ���� NUMBER(16,5),
    ��ĸ NUMBER(16,5))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ�����޶�(
    �ⷿid NUMBER(18),
    ҩƷid NUMBER(18),
    ���� NUMBER(18,5),
    ���� NUMBER(18,5),
    �̵����� VARCHAR2(4),
    �ⷿ��λ VARCHAR2(50),
    ���ñ�־ Number(1) default 1)
    TABLESPACE zl9BaseItem;

Create Table ҩƷ�б굥λ(
		ҩƷid NUMBER(18),
		��λID NUMBER(18),
		����ʱ�� Date,
		����ʱ�� Date,
		�б���� Varchar2(50))
    TABLESPACE zl9BaseItem;

CREATE TABLE ��Һ��ҩ����(
  ���� varchar2(4),
  ���� varchar2(50),
  ���� VARCHAR2(20)
  ) TABLESPACE zl9BaseItem;
  
CREATE TABLE ������������(
  ����id varchar2(20),
  �������� varchar2(20),
  ��ҩ���� varchar2(20),
  ���� number(18),
  ��������ID number(18)
  ) TABLESPACE zl9MedLst;

CREATE TABLE ��ҺҩƷ���ȼ�(
  ����id varchar2(1000),
  �������� varchar2(2000),
  ��ҩ���� varchar2(200),
  Ƶ�� VARCHAR2(200),
  ��Ч number(1),
  ���ȼ� number(3)
  ) TABLESPACE zl9MedLst;

Create Table ��ҩ��������(
    ���� NUMBER(2),
    ��ҩʱ�� Varchar2(20),
    ��ҩʱ�� Varchar2(20),
    ��� Number(1) Default 0,
    ���� Number(1) Default 1,
    ��ɫ NUMBER(18),
    ��������ID number(18))
    TABLESPACE zl9MedLst;

Create Table ��ҺҩƷ����(
    ҩƷID NUMBER(18),
    �洢�¶� NUMBER(1),
    �洢���� NUMBER(1),
    ��ҩ���� VARCHAR2(30),
    �Ƿ������� NUMBER(1),
    ��Һע������ varchar2(200))
    TABLESPACE zl9MedLst
;

Create Table ���ϻ�Դ���(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table ���ϼ�ֵ����(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table ������Դ����(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

Create Table ���ϲ��ʷ���(
    ���� VARCHAR2(4),
    ���� VARCHAR2(30),
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem;
Create Table ���ϴ洢����(
    ���� VARCHAR2(4),
    ���� VARCHAR2(30),
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem;

Create Table ����������(
    ���� VARCHAR2(10),
    ���� VARCHAR2(60),
    ���� VARCHAR2(10),
    ������ҵ���֤ varchar2(40),
    ������ҵ���֤Ч�� Date,
    վ�� Varchar2(1),
    ��Ӫ���֤ varchar2(40),
    ��Ӫ���֤Ч�� date,
    ��ҵ����ִ�� varchar2(40),
    ��ҵ����ִ��Ч�� date)
    TABLESPACE zl9BaseItem;

Create Table �����������(
    ���ڿⷿID NUMBER(18),
    �Է��ⷿID NUMBER(18),
    ���� NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE ���Ͽⷿ��λ(
	���� VARCHAR2(5),
	���� VARCHAR2(50),
	���� VARCHAR2(10))
	TABLESPACE zl9BaseItem;

Create Table ���ϴ����޶�(
    �ⷿid NUMBER(18),
    ����id NUMBER(18),
    ���� NUMBER(18,5),
    ���� NUMBER(18,5),
    �̵����� VARCHAR2(4),
    �ⷿ��λ VARCHAR2(50))
    TABLESPACE zl9BaseItem;

Create Table ���Ʋ��Ϲ���(
    ���Ʋ���ID NUMBER(18),
    ԭ�ϲ���ID NUMBER(18),
    ���� NUMBER(16,5),
    ��ĸ NUMBER(16,5))
    TABLESPACE zl9BaseItem;

Create Table ��������(
    ����ID number(18),
    ����ID number(18),
    ���Ч�� number(5),
    ���Ч�� number(5),
    �޾��Բ��� number(1) default 0,
    һ���Բ��� number(1) default 0,
    ԭ���� number(1) DEFAULT 0,
    ���Ʋ��� number(1) default 0,
    ��Դ��� varchar2(10),
    ���ʷ��� varchar2(30),
    �洢���� varchar2(30),
    ���֤�� varchar2(50),
    ���֤��Ч�� DATE,
    ��׼�ĺ� VARCHAR2(40),
    ע���̱� VARCHAR2(50),
    ע��֤�� Varchar2(50),
    ��װ��λ varchar2(8),
    ����ϵ�� number(16,5),
    ָ�������� number(16,7),
    ָ�����ۼ� number(16,7),
    ָ������� number(16,5),
    �ɱ��� number(16,7),
    ��������� NUMBER(16,5),
    ���� number(16,5),
    �ⷿ���� number(1) default 0,
    ���÷��� number(1) default 0,
    ������Դ varchar2(10),
    ������λ Varchar2(20),
    ����ϵ�� Number(16,5),
    �б���� NUMBER(1),
    �������� number(1),
    ���ٲ��� NUMBER(1)  DEFAULT 0,
    ������� NUMBER(1) DEFAULT 0,
    ��ֵ˰�� NUMBER(16,5),
    ��ֵ���� number(1),
    �Ƿ�������� Number(1),
    �ϴ��ۼ� number(16,7),
    ��е�����ĵ��� number(1),
    �ϴι�Ӧ��id number(18),
    �ϴβ��� varchar2(60))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �����б굥λ(
    ����id NUMBER(18),
    ��λID NUMBER(18),
    �ɱ��� NUMBER(16,5),
    �б���� Varchar2(50))
    TABLESPACE zl9BaseItem;

CREATE TABLE ����������λ(
	���� VARCHAR2(2),
	���� VARCHAR2(50),
	���� VARCHAR2(10))
	TABLESPACE zl9BaseItem;

Create Table ����������;(
    ���� VARCHAR2(6),
    ���� varchar2(50),
    ���� varchar2(10),
    ȱʡ��־ number(2) DEFAULT 0)
    TABLESPACE zl9BaseItem;

CREATE TABLE ���ϼӳɷ���(
	���	number(18),
	��ͼ�	number(16,5),
	��߼�	number(16,5),
	�ӳ���	number(16,5),
	���㷽�� number(1),
	�޼� number(16,5),
	˵�� varchar2(50))
    TableSpace zl9BaseItem;

Create Table ҩƷ���ľ���(
    ����        Number(1),
    ���	Number(1),
    ����	Number(1),
    ��λ	Number(1),
    ����	Number(1))
    TABLESPACE zl9BaseItem;

Create Table ҩƷ������(
	�ⷿID Number(18),
	��鷽ʽ Number(1))
	TABLESPACE zl9BaseItem;

Create Table ���ϳ�����(
	�ⷿID Number(18),
	��鷽ʽ Number(1))
	TABLESPACE zl9BaseItem;

Create Table ����ⷿ����(
  ����ID Number(18),
  �ⷿID Number(18),
  ����ⷿID Number(18))
  TABLESPACE zl9BaseItem;

----------------------------------------------------------------------------
--[[6.�ٴ�����]]
----------------------------------------------------------------------------
Create Table ��Һͨ��(
    ���� VARCHAR2(4),
    ���� VARCHAR2(20),
    ���� VARCHAR2(20))
    TABLESPACE zl9BaseItem;

Create table ��Ѫ�������
(
��Ŀid number(18),
������Ŀid number(18)
)TableSpace zl9BaseItem;

create table ����ҩ�������ҩ(
   ����ID number(18), 
   ��� number(18),
   ҩƷ��� number(5),
   ҩ��ID number(18), 
   ͼ�� Varchar2(50),
   ҩ�� Varchar2(50),
   ���� Varchar2(50),
   Ƶ�� Varchar2(50),
   ;�� Varchar2(50),
   ���� Varchar2(50),
   ��ֹ���� Varchar2(50)
) TABLESPACE zl9CisRec;

create table ����ҩ���������(
   ����ID number(18), 
   ��� number(18),
   ����ID number(18),
   �������� VARCHAR2(100),
   �п� Varchar2(20),
   ��ʼʱ�� Date,
   ����ʱ�� Date,
   Ԥ����ҩ�ڼ� number(2),
   ��ҩ��� number(1)
) TABLESPACE zl9CisRec;

Create Table ������ҩ������Ŀ(
    ���� Varchar2(5),
    ��� Number(5),
    ���� Varchar2(200),
    �Ƿ����� Number(1),  
    �Ƿ���� Number(1),  
    �ϼ� Varchar2(5),
    ĩ�� NUMBER(1)
)TABLESPACE zl9BaseItem;

Create Table ����Ԥ����ҩ�ڼ�(
     ���� Number(5), 
     ���� Varchar2(200)
)TABLESPACE zl9BaseItem;

create table ����ҩ���������(
   ����ID number(18), 
   ��� number(18),
   �������� number(1),
   �Ƿ���� number(1),
   ��Ŀ���� Varchar2(5),
   ����� number(5),
   ��Ŀֵ Varchar2(200)
) TABLESPACE zl9CisRec;

create table ����ҩ�������¼
(
   ID number(18),
   ������ Varchar2(20),
   ����ʱ�� Date,
   ��Χ��ʼʱ�� Date,
   ��Χ����ʱ�� Date
) TABLESPACE zl9CisRec;

create table ����ҩ�������ϸ
(
   ����ID number(18),
   ����ID number(18),
   ��ҳID Number(5),
   ���  Number(18),
   �Ƿ����� Number(1), 
   ��ԭѧ��� 	Number(1), 
   ��ԭѧ�������	Date,	
   ��ԭѧ���걾	varchar2(50), 
   ��ԭѧ�����ϸ����	varchar2(100), 
   ҩ������ 	Number(1), 
   ҩ����������	Date,	
   ҩ�������Ƿ����	Number(1), 
   ��ҩǰ���� varchar2(30),
   ��ҩǰ��ϸ������ varchar2(30), 
   ��ҩǰ������ϸ�� varchar2(30),
   ��ҩǰC��Ӧ���� varchar2(30),
   ��ҩǰ����ת��ø varchar2(30),
   ��ҩǰ���� varchar2(30),
   ��ҩ������ varchar2(30),
   ��ҩ���ϸ������ varchar2(30),
   ��ҩ��������ϸ�� varchar2(30),
   ��ҩ��C��Ӧ���� varchar2(30),
   ��ҩ�����ת��ø varchar2(30),
   ��ҩ���� varchar2(30),
   ��ҩǰ�������� Date, 
   ��ҩǰ��ϸ���������� Date, 
   ��ҩǰ������ϸ������ Date,
   ��ҩǰC��Ӧ�������� Date,
   ��ҩǰ����ת��ø���� Date,
   ��ҩǰ�������� Date,
   ��ҩ���������� Date,
   ��ҩ���ϸ���������� Date,
   ��ҩ��������ϸ������ Date,
   ��ҩ��C��Ӧ�������� Date,
   ��ҩ�����ת��ø���� Date,
   ��ҩ�������� Date, 
   Ӱ��ѧ��� varchar2(200), 
   Ӱ��ѧ��ϲ�λ	varchar2(50),
   Ӱ��ѧ��Ͻ���	varchar2(100),
   �ٴ�֢״ number(18),  
   ��ҩĿ�� number(2),     
   ��Ⱦ��� number(18),   
   ���ƽ�� number(2),    
   ��Ӧ֢ varchar2(500),   
   ҩ��ѡ�� varchar2(500),  
   ���μ��� varchar2(500),  
   ÿ�ո�ҩƵ�� varchar2(500), 
   �ܼ� varchar2(500),  
   ��ҩ;�� varchar2(500),  
   ��ҩ�Ƴ� varchar2(500),  
   ��ǰ��ҩʱ�� varchar2(500),  
   ������ҩ varchar2(500),   
   ������ҩ varchar2(500),   
   ������ҩ varchar2(500), 
   ����ҩ�� varchar2(500),  
   ��ע varchar2(500), 
   �Ƿ��ӡ Number(1),
   �Ƿ�༭ Number(1),
   ��ҩ���� NUMBER(5),
   ����ҩ���� NUMBER(5),
   �Ƿ��ÿ����ҩ Number(1)
) TABLESPACE zl9CisRec;

CREATE TABLE ��ѪĿ��(
    ���� VARCHAR2(4),
    ���� VARCHAR2(100),
    ���� VARCHAR2(20)
)TABLESPACE zl9BaseItem; 

Create Table �ٴ�����(
    �������� VARCHAR2(10),
    ����id NUMBER(18))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �ٴ�����(
    ���� VARCHAR2(10),
    ���� VARCHAR2(30),
    ���� VARCHAR2(15),
		��� NUMBER(4))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table ��������(
    ���� VARCHAR2(3),
    ���� VARCHAR2(20),
    ���� VARCHAR2(6),
    λ�� VARCHAR2(40),
    վ�� Varchar2(1),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table �ٴ�ҽ��С��(
	ID NUMBER(18),
	����ID NUMBER(18),
	���� VARCHAR2(50),
	˵�� VARCHAR2(200),
        ����ʱ�� Date,
	����ʱ�� Date)
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep);

Create Table ҽ��С����Ա(
	С��ID NUMBER(18),
	��ԱID NUMBER(18),
	�Ƿ��鳤 Number(1))
    TABLESPACE zl9BaseItem
    Cache Storage(Buffer_Pool Keep);

 Create Table ��Ա����ҩ��Ȩ��(
	��Աid Number(18), 
	���� Number(1), 
	��¼״̬ Number(3) Default (1), 
	������Ա Varchar2(20), 
	����ʱ�� Date,
	���� Number(2) default(1)) 
	Tablespace Zl9baseitem;

Create Table ��Ա����Ȩ��(
  ��Աid Number(18), 
  ������ĿID Number(18),
  ��¼���� Number(3)) 
  Tablespace Zl9baseitem;  

CREATE TABLE ���������(
	���� VARCHAR2(3),
	���� VARCHAR2(100))
TABLESPACE zl9BaseItem;

Create Table ������Ŀ¼(
    ���� VARCHAR2(2),
    ���� VARCHAR2(50),
    ���� VARCHAR2(25),
    ICD���� VARCHAR2(1000))
    TABLESPACE zl9BaseItem;

Create Table ��������˵��(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(6)
    )
    TABLESPACE zl9BaseItem;

Create Table ҽ�����ݶ���(
    ������� VARCHAR2(1),
    ҽ������ VARCHAR2(500))
    TABLESPACE zl9BaseItem;

Create Table ҽ��δִ��ԭ��(
    ���� VARCHAR2(5),
    ���� VARCHAR2(100),
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem;

Create Table ҽ������ԭ��(
    ���� VARCHAR2(4),
    ���� VARCHAR2(200),
    ���� VARCHAR2(200),
    ���� Number(1),
    ��Ա Varchar2(100))
    TABLESPACE zl9BaseItem;

Create Table ���Ʋο�����(
    ID NUMBER(18),
    �ϼ�id NUMBER(18),
    ���� VARCHAR2(8),
    ���� VARCHAR2(40),
    ���� VARCHAR2(10),
    ���� number(1))
    TableSpace zl9BaseItem;

Create Table ���Ʋο�Ŀ¼(
    ID NUMBER(18),
    ����ID NUMBER(18),
    ���� VARCHAR2(13),
    ���� VARCHAR2(60),
    ˵�� VARCHAR2(4000),
    ���� VARCHAR2(20),
    ���� NUMBER(1))
    TableSPACE zl9BaseItem;

Create Table ���Ʋο�����(
    �ο�Ŀ¼ID NUMBER(18),
    ���� VARCHAR2(60),
    ���� NUMBER(1),
    ���� VARCHAR2(12),
    ���� NUMBER(1))
    TableSpace zl9BaseItem;

Create Table ���Ʋο�����(
    �ο�Ŀ¼ID NUMBER(18),
    ��Ŀ��� NUMBER(5),
    �ο���Ŀ VARCHAR2(20),
    ��Ŀ��� NUMBER(1),
    �����к� NUMBER(5),
    �����ı� VARCHAR2(4000),
    �������� NUMBER(3))
    TableSPACE zl9BaseItem;

Create Table ���Ʋο�����(
    �ο�Ŀ¼ID NUMBER(18),
    �ο���Ŀ VARCHAR2(20),
    �����к� NUMBER(5),
    ����֢ID NUMBER(18),
    �������� NUMBER(1))
    TableSPACE zl9BaseItem;

CREATE TABLE ������Ŀ���(
    ���� VARCHAR2(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(10))
    TableSpace zl9BaseItem;

CREATE TABLE ���Ʒ���Ŀ¼(
    ID NUMBER(18),
    ���� VARCHAR2(8),
    ���� VARCHAR2(40),
    ���� VARCHAR2(10),
    �ϼ�id NUMBER(18),
    ���� NUMBER(1),
		����ʱ�� DATE,
		����ʱ�� DATE)
    TableSpace zl9BaseItem;

CREATE TABLE ������ĿĿ¼(
		��� VARCHAR2(1),
		����ID NUMBER(18),
		ID NUMBER(18),
		���� VARCHAR2(20),
		���� VARCHAR2(60),
		�걾��λ VARCHAR2(60),
		���㵥λ VARCHAR2(20),
		���㷽ʽ NUMBER(1),
		������� Number(1),
		ִ��Ƶ�� NUMBER(1),
		�����Ա� NUMBER(1),
		����Ӧ�� NUMBER(1),
		�����Ŀ NUMBER(1),
		�������� VARCHAR2(20),
		ִ�а��� NUMBER(1),
		ִ�п��� NUMBER(1),
		������� NUMBER(1),
		�Ƽ����� NUMBER(1),
		�ο�Ŀ¼ID NUMBER(18),
		��ԱID NUMBER(18),
		����ʱ�� DATE,
		����ʱ�� DATE,
		¼������ NUMBER(16,5),
		�Թܱ��� Varchar2(4),
		ִ�з��� NUMBER(2),
		ִ�б�� NUMBER(1),
		վ�� Varchar2(1),
		������ Varchar2(20),
		����ϵ��  number(16,5))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

Create Table ���Ƹ�����Ŀ(
    ��ԱID NUMBER(18),
	������ĿID NUMBER(18),
	�շ�ϸĿID NUMBER(18),
	Ƶ�� Number(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE ������Ŀ����(
    ������Ŀid NUMBER(18),
    ���� VARCHAR2(60),
    ���� NUMBER(1),
    ���� VARCHAR2(30),
    ���� NUMBER(1))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE ���ƻ�����Ŀ(
    ���� NUMBER(18),
    ������ VARCHAR2(30),
    ��ĿID NUMBER(18),
    ���� NUMBER(18))
    TableSpace zl9BaseItem;


CREATE TABLE �������ÿ���(
    ��ĿID NUMBER(18),
    ����ID NUMBER(18))
    TableSpace zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE ����ִ�п���(
    ������Ŀid NUMBER(18),
    ������Դ NUMBER(1) DEFAULT 1,
    ��������id NUMBER(18),
    ִ�п���id NUMBER(18))
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

create table �����÷�����
(
  ��Ŀid NUMBER(18),
  ����   NUMBER(3),
  �÷�id NUMBER(18),
  Ƶ��   VARCHAR2(3),
  ���˼��� NUMBER(16,5),
  С������ NUMBER(16,5),
  ҽ������ VARCHAR2(100),
  �Ƴ�   NUMBER(5),
  dddֵ NUMBER(16,5)
)
tablespace ZL9BASEITEM;

CREATE TABLE ������Ŀ���(
		�������ID NUMBER(18),
		��� NUMBER(18),
		������ NUMBER(18),
		��Ч NUMBER(1),
		������ĿID NUMBER(18),
		ҽ������ Varchar2(1000),
		���� NUMBER(16,5),
		�������� NUMBER(16,5),
		�ܸ����� NUMBER(16,5),
		�շ�ϸĿID NUMBER(18),
		�걾��λ VARCHAR2(60),
		��鷽�� Varchar2(30),
		ҽ������ VARCHAR2(100),
		ִ��Ƶ�� VARCHAR2(20),
		Ƶ�ʴ��� NUMBER(3),
		Ƶ�ʼ�� NUMBER(3),
		�����λ VARCHAR2(4),
		ִ������ NUMBER(1),
		ִ�б�� NUMBER(1),
		ִ�п���ID NUMBER(18),
		ʱ�䷽�� VARCHAR2(50),
		�䷽ID Number(18),
		�����ĿID Number(18),
		�䷽���� number(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE �����շѹ�ϵ(
    ������Ŀid NUMBER(18),
    �շ���Ŀid NUMBER(18),
    �շ����� NUMBER(16,5) DEFAULT 1,
    ���ж��� NUMBER(1),
    ������Ŀ Number(1),
		�������� Number(1) default 0 not Null,
    ��鲿λ Varchar2(30),
    ��鷽�� Varchar2(30),
		�շѷ�ʽ Number(1),
		���ÿ���ID NUMBER(18),
		������Դ NUMBER(1) Default 0 Not Null)
    TABLESPACE zl9BaseItem    
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE ����Ƶ����Ŀ(
    ���� VARCHAR2(3),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10),
    Ӣ������ VARCHAR2(50),
    Ƶ�ʴ��� NUMBER(3),
    Ƶ�ʼ�� NUMBER(3),
    �����λ VARCHAR2(4),
    ���÷�Χ NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE ����Ƶ��ʱ��(
    ִ��Ƶ�� VARCHAR2(3),
    ������� NUMBER(3),
    ʱ�䷽�� VARCHAR2(50),
    ��ҩ;��ID NUMBER(18))
    TABLESPACE zl9BaseItem;

Create Table ҩ�ȸ�������(
	��� Varchar2(10),
	���� Varchar2(4000)) 
	TABLESPACE zl9BaseItem;

CREATE TABLE ������������(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE ����������ģ(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(10),
    ȱʡ��־ NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE ��ҩ�����ע(
    ���� VARCHAR2(5),
    ���� VARCHAR2(20),
    ���� VARCHAR2(8))
    TABLESPACE zl9BaseItem;

Create Table ��ҩ������(
	���� VARCHAR2(2),
    ���� VARCHAR2(1) Not Null,
    ��ֵ NUMBER(16,5) Not Null)
    TABLESPACE zl9BaseItem;

CREATE TABLE ������������(
    ���� VARCHAR2(1),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10),
    �̶� NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE ��������(
    ���� VARCHAR2(5),
    ���� VARCHAR2(100),
    ���� VARCHAR2(20),
		��Ա VARCHAR2(20))
    TABLESPACE zl9BaseItem;

CREATE TABLE ���ü�������(
    ���� VARCHAR2(3),
    ���� VARCHAR2(30),
    ���� Number(5,2))
    TABLESPACE zl9BaseItem;

CREATE TABLE �����ο���Ŀ(
    ��� NUMBER(1),
    ����� NUMBER(3),
    ����� NUMBER(3),
    ���� VARCHAR2(20),
    ��ʽ NUMBER(1),
    ��� NUMBER(1),
    ���� NUMBER(1))
    TABLESPACE zl9BaseItem;

Create Table ���Ʋο���Ŀ(
    ���� Number(1),
    ��� NUMBER(3),
    ��� NUMBER(1),
    ���� VARCHAR2(20),
    ���� NUMBER(1))
    TableSPACE zl9BaseItem;

CREATE TABLE ������Ϸ���(
    ID NUMBER(18),
    �ϼ�ID NUMBER(18),
    ���� VARCHAR2(6),
    ���� VARCHAR2(40),
    ���� VARCHAR2(10),
    ��� NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE �����������(
    ����ID NUMBER(18),
    ���ID NUMBER(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE �������Ŀ¼(
    ID NUMBER(18),
    ��� NUMBER(1),
    ���� VARCHAR2(10),
    ���� VARCHAR2(40),
    ˵�� VARCHAR2(4000),
    ���� VARCHAR2(20),
    ���� NUMBER(5),
    �ٴ� NUMBER(5))
    TABLESPACE zl9BaseItem   
    Cache Storage(Buffer_Pool Keep);

Create TABLE ������Ͽ���(
    ���ID NUMBER(18),
    ����ID NUMBER(18),
    ��ԱID Number(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE ������ϱ���(
    ���id NUMBER(18),
    ���� VARCHAR2(40),
    ���� NUMBER(1),
    ���� VARCHAR2(12),
    ���� NUMBER(1))
    TABLESPACE zl9BaseItem   
    Cache Storage(Buffer_Pool Keep);

CREATE TABLE ������ϲο�(
    ���ID NUMBER(18),
    ��Ŀ��� NUMBER(5),
    �ο���Ŀ VARCHAR2(20),
    ��Ŀ��� NUMBER(1),
    ��Ŀ��ʽ NUMBER(1),
    ֤��ID NUMBER(18),
    ֤����� NUMBER(5),
    ֤������ VARCHAR2(20),
    �����к� NUMBER(5),
    �����ı� VARCHAR2(4000),
    �������� NUMBER(1))
    TABLESPACE zl9BaseItem;

Create Table ��ϲ��ֶ�Ӧ(
    ���ID NUMBER(18),
    ���� NUMBER(3),
    ����ID NUMBER(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE �������ƴ�ʩ(
    ���ID NUMBER(18),
    �ο���Ŀ VARCHAR2(20),
    ֤������ VARCHAR2(20),
    �����к� NUMBER(5),
    ������ĿID NUMBER(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE ������Ϲ���(
    ���ID NUMBER(18),
    ����� NUMBER(3),
    ������ VARCHAR2(20),
    ������ NUMBER(3),
    ��ĿID NUMBER(18),
    ��ϵʽ VARCHAR2(10),
    ����ֵ VARCHAR2(250),
    ���ɶ� NUMBER(3))
    TABLESPACE zl9BaseItem;

CREATE TABLE ������϶���(
    ����ID NUMBER(18),
    ���ID NUMBER(18),
    ����ID NUMBER(18))
    TABLESPACE zl9BaseItem;


----------------------------------------------------------------------------
--[[7.�ٴ�·������]]
----------------------------------------------------------------------------
CREATE TABLE ·������Ŀ¼(
	ID		NUMBER(18),
	����    VARCHAR2(64),
	����	VARCHAR2(100),
	�Ƿ�̶� NUMBER(1)
	)
	TABLESPACE zl9CISRec;

CREATE TABLE ·������ṹ(		
	����ID	NUMBER(18),
	�к�	NUMBER(5),
	��Ŀ���	NUMBER(5),
	��Ŀ�ı�1 VARCHAR2(100),
	��Ŀ�ı�2 VARCHAR2(100),
	SQL�ı� VARCHAR2(4000),
	ҳ�� number(3),
	·��ID number(18),
	��ѡ��� number(5)
	)
    TABLESPACE zl9CISRec;

Create Table ·��������� (
   ����ID number(18),
   �к�  NUMBER(5),
   ·��ID number(18),
   ��� Number(8)
) TABLESPACE zl9CISRec;

Create Table ��׼·��Ŀ¼(
 ID NUMBER(8),
 �������� Varchar2(100),
 ���� Varchar2(8),   
 ·������ Varchar2(80),
 ���  NUMBER(2),
 �汾˵�� Varchar2(20) 
)
    
 tablespace ZL9BASEITEM
;

create table ��׼·������(
	  ��׼·��id NUMBER(8) ,
	  ���     NUMBER(3) ,
	  ����     VARCHAR2(100),
	  ����     VARCHAR2(4000)
	)
	tablespace ZL9BASEITEM;

create table ��׼·������(
	  ��׼·��id NUMBER(8),
	  ��������   VARCHAR2(100),
	  ��������   VARCHAR2(100)
	)
	tablespace ZL9BASEITEM;

create table ��׼·����(
	  ��׼·��id NUMBER(8),
	  �����   NUMBER(3),
	  ������   VARCHAR2(100),
	  ����ͷ   Varchar2(500),
	  �������   NUMBER(3),
	  ��������   VARCHAR2(50),
	  �׶����   NUMBER(3),
	  �׶�����   VARCHAR2(100),
	  ·������   VARCHAR2(2000)
	)
	tablespace ZL9BASEITEM;

CREATE TABLE ·����Ŀ˳��(
	˳�� number(2),
	ҽ����Ч NUMBER(1),
	������� VARCHAR2(1),
	�������� VARCHAR2(20),
	ִ�з��� NUMBER(2))
TableSpace zl9BaseItem       
    Cache Storage(Buffer_Pool Keep);

Create Table �ٴ���������(
    ���� VARCHAR2(1),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem;

Create Table ·���������(
    ���� NUMBER(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem;

Create Table ·���������(
    ���� VARCHAR2(5),
    ���� VARCHAR2(100),
    ���� VARCHAR2(10),
		�ϼ� VARCHAR2(5),
		ĩ�� NUMBER(1),
		���� NUMBER(1),
		���� NUMBER(2))
    TABLESPACE zl9BaseItem;

Create Table ���쳣��ԭ��(
    ���� VARCHAR2(6),
    ���� VARCHAR2(200),
    ���� VARCHAR2(20),
	�ϼ� VARCHAR2(6),
	ĩ�� NUMBER(1),
	���� NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE �ٴ�·��ͼ��(
	ID NUMBER(18),
	ͼ�� BLOB,
	���� NUMBER(1))
	LOB(ͼ��) Store as (Cache)
    TABLESPACE zl9BaseItem;

----------------------------------------------------------------------------
--[[8.��������]]
----------------------------------------------------------------------------
CREATE TABLE ��ϵ��ν��ϵ(
    ��� NUMBER(3),
    ���� NUMBER(3),
    ĸ�� NUMBER(3),
    ��ν VARCHAR2(10),
    ��ϵ VARCHAR2(12),
    �Ա� VARCHAR2(4),
    Ψһ��ϵ NUMBER(3),
    ���ֵȼ� NUMBER(5),
    �����ȼ� NUMBER(3),
    ѪԵ��ϵ NUMBER(3),
	������� varchar2(2))
    TABLESPACE zl9BaseItem;

CREATE TABLE ������������(
    ���� VARCHAR2(1),
    ID NUMBER(18),
    �ϼ�ID NUMBER(18),
    ���� VARCHAR2(6),
    ���� VARCHAR2(40),
    ���� VARCHAR2(8))
    TABLESPACE zl9BaseItem;

CREATE TABLE ����������Ŀ(
    ID NUMBER(18),
    ����ID NUMBER(18),
    ���� VARCHAR2(13),
    ������ VARCHAR2(60),
    Ӣ���� VARCHAR2(40),
    �滻�� NUMBER(1),
    ���� NUMBER(3),
    ���� NUMBER(3),
    С�� NUMBER(3),
    ��λ VARCHAR2(20),
    �ٴ����� VARCHAR2(250),
    ��ʾ�� NUMBER(1),
    �Ա��� NUMBER(1),
    ��ֵ�� VARCHAR2(1000),
    ������ VARCHAR2(1000),
    ��ʼֵ VARCHAR2(1000),
    ���ֱ��� NUMBER(1),
    ��ֵ���� VARCHAR2(100),
    ���� Number(1) Default 0,
	��̬�� Number(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE �������ͼ��(
    ���� VARCHAR2(4),
    ���� VARCHAR2(30),
    ���� VARCHAR2(10),
    ͼ�� BLOB)	
    TABLESPACE zl9EprLob;

CREATE TABLE ����������ʽ(
    ��� NUMBER(3),
    ���� VARCHAR2(30),
    ������ʽ VARCHAR2(4000),
    ������ʽ VARCHAR2(4000),
    ϵͳ NUMBER(1))
    TABLESPACE zl9EprDat;

Create Table ������д�¼� (
  ���� NUMBER(3),
  ��� NUMBER(3),
  ���� Varchar2(20),
  ���� Varchar2(10),
  ˵�� Varchar2(100),
  ��ǰ���� Number(1),
  ѭ������ Number(1))
  TABLESPACE zl9EprDat;

CREATE TABLE ����ҳ���ʽ(
    ���� NUMBER(3),
    ��� VARCHAR2(3),
    ���� VARCHAR2(30),
    ���� NUMBER(1),
    ��ʽ VARCHAR2(4000),
    ҳü VARCHAR2(1000),
    ҳ�� VARCHAR2(1000),
    ͼ�� BLOB,
	ҳü�ļ� BLOB,
	ҳ���ļ� BLOB)
    TABLESPACE zl9EprDat;

CREATE TABLE �����ļ��б�(
    ID NUMBER(18),
    ���� NUMBER(3),
    ���� Varchar2(10),
    ��� VARCHAR2(3),
    ���� VARCHAR2(30),
    ˵�� VARCHAR2(2000),
    ҳ�� VARCHAR2(3),
    ���� NUMBER(5),
    ͨ�� NUMBER(3))
    TABLESPACE zl9EprDat;

CREATE TABLE ����ʱ��Ҫ��(
    �ļ�ID NUMBER(18),
    �¼� VARCHAR2(20),
    ���� NUMBER(1),
    Ψһ NUMBER(1),
    ��дʱ�� NUMBER(5),
    ����ʱ�� NUMBER(5),
    ���ʱ�� NUMBER(5),
    һ������ NUMBER(5),
    �������� NUMBER(5),
    ��Σ���� NUMBER(5))
    TABLESPACE zl9EprDat;

CREATE TABLE ���������ϵ(
    �ļ�ID NUMBER(18),
    ���ID NUMBER(18))
    TABLESPACE zl9EprDat;

CREATE TABLE ����Ӧ�ÿ���(
    �ļ�ID NUMBER(18),
    ����ID NUMBER(18))
    TABLESPACE zl9EprDat;

CREATE TABLE ��������ǰ��(
    �ļ�ID NUMBER(18),
    ����ID NUMBER(18),
    ���ID NUMBER(18))
    TABLESPACE zl9EprDat;

Create Table ��������Ӧ��(
    ������ĿID Number(18),
    Ӧ�ó��� Number(3),
    �����ļ�ID Number(18))
    Tablespace zl9EprDat;

Create Table �������ݸ���(
    �ļ�ID Number(18),
    ��Ŀ Varchar2(30),
    ���� Number(1),
    ���� Number(5),
    Ҫ��ID Number(18),
    ֻ�� number(1),
    ���� Varchar2(200))
    Tablespace zl9EprDat;
    
create table ��������ģ��(
    ID NUMBER(18),
    �����ļ�Id NUMBER(18),
    ���ݸ��� VARCHAR2(30),
    ģ����� VARCHAR2(30),
    ģ������ VARCHAR2(512),
    ʹ�ô��� number(8)       
)TABLESPACE zl9EprDat;
    
CREATE TABLE �����ļ���ʽ(
    �ļ�ID NUMBER(18),
    ���� BLOB)
	LOB(����) Store as (Cache)
    TABLESPACE zl9EprLob;

CREATE TABLE �����ļ��ṹ(
    ID NUMBER(18),
    �ļ�ID NUMBER(18),
    ��ID NUMBER(18),
    ������� NUMBER(18),
    �������� NUMBER(1),
    ������ NUMBER(18),
    �������� NUMBER(1),
    �������� VARCHAR2(1000),
    �����д� NUMBER(18),
    �����ı� VARCHAR2(4000),
    �Ƿ��� NUMBER(1),
    Ԥ�����ID NUMBER(18),
    ������� NUMBER(1),
    ʹ��ʱ�� VARCHAR2(8),
    ����Ҫ��ID NUMBER(18),
		�滻�� NUMBER(1),
    Ҫ������ VARCHAR2(40),
    Ҫ������ NUMBER(3),
    Ҫ�س��� NUMBER(3),
    Ҫ��С�� NUMBER(3),
    Ҫ�ص�λ VARCHAR2(50),
    Ҫ�ر�ʾ NUMBER(3),
    ������̬ NUMBER(3),
    Ҫ��ֵ�� VARCHAR2(4000))
    TABLESPACE zl9EprDat;

CREATE TABLE �����ļ�ͼ��(
    ����ID NUMBER(18),
    ͼ�� BLOB)
	LOB(ͼ��) Store as (Cache)
    TABLESPACE zl9EprLob;

Create Table �����ʾ����(
    ID Number(18),
    �ϼ�ID Number(18),
    ���� Varchar2(8),
    ���� Varchar2(30),
    ˵�� Varchar2(200),
    ��Χ Varchar2(8))
    Tablespace zl9EprDat;

Create Table �����ʾ�ʾ��(
    ID Number(18),
    ����ID Number(18),
    ��� Varchar2(13),
    ���� Varchar2(60),
    ͨ�ü� Number(1),
    ����id Number(18),
    ��Աid Number(18))
    Tablespace zl9EprDat;

CREATE TABLE �����ʾ����(
    �ʾ�ID NUMBER(18),
    ���д��� NUMBER(5),
    �������� NUMBER(3),
    �����ı� VARCHAR2(4000),
    ����Ҫ��ID NUMBER(18),
	�滻�� NUMBER(1),
    Ҫ������ VARCHAR2(40),
    Ҫ������ NUMBER(3),
    Ҫ�س��� NUMBER(3),
    Ҫ��С�� NUMBER(3),
    Ҫ�ص�λ VARCHAR2(10),
    Ҫ�ر�ʾ NUMBER(3),
    Ҫ��ֵ�� VARCHAR2(4000),
    ������̬ NUMBER(3),
    �������� Varchar2(1000))
    TABLESPACE zl9EprDat;

Create Table �����ʾ�����(
    �ʾ�ID Number(18),
    ������ Varchar2(20),
    ����ֵ Varchar2(2000))
    Tablespace zl9EprDat;

Create Table ������ٴʾ�(
    ���ID Number(18),
    �ʾ����ID Number(18))
    Tablespace zl9EprDat;

CREATE TABLE ��������Ŀ¼(
    ID NUMBER(18),
    �ļ�ID NUMBER(18),
    ��� VARCHAR2(5),
    ���� VARCHAR2(30),
    ���� VARCHAR2(10),
    ���� Varchar2(50),
    ���� NUMBER(1),
    ˵�� VARCHAR2(100),
    ͨ�ü� NUMBER(1),
    ����id NUMBER(18),
    ��Աid NUMBER(18))
    TABLESPACE zl9EprDat;

CREATE TABLE �������ĸ�ʽ(
    �ļ�ID NUMBER(18),
    ���� BLOB)
	LOB(����) Store as (Cache)
    TABLESPACE zl9EprLob;

CREATE TABLE ������������(
    ID NUMBER(18),
    �ļ�ID NUMBER(18),
    ��ID NUMBER(18),
    ������� NUMBER(18),
    �������� NUMBER(1),
    ������ NUMBER(18),
    �������� NUMBER(1),
    �������� VARCHAR2(1000),
    �����д� NUMBER(18),
    �����ı� VARCHAR2(4000),
    �Ƿ��� NUMBER(1),
    Ԥ�����ID NUMBER(18),
		�������ID Number(18),
    ������� NUMBER(1),
    ʹ��ʱ�� VARCHAR2(2),
    ����Ҫ��ID NUMBER(18),
		�滻�� NUMBER(1),
    Ҫ������ VARCHAR2(40),
    Ҫ������ NUMBER(3),
    Ҫ�س��� NUMBER(3),
    Ҫ��С�� NUMBER(3),
    Ҫ�ص�λ VARCHAR2(50),
    Ҫ�ر�ʾ NUMBER(3),
    ������̬ NUMBER(3),
    Ҫ��ֵ�� VARCHAR2(4000))
    TABLESPACE zl9EprDat;

CREATE TABLE ��������ͼ��(
    ����ID NUMBER(18),
    ͼ�� BLOB)
	LOB(ͼ��) Store as (Cache)
    TABLESPACE zl9EprLob;

Create Table ������������(
    ����ID Number(18),
    ������ Varchar2(20),
    ����ֵ Varchar2(2000))
    Tablespace zl9EprDat;

Create Table �������İ�(
       ID Number(18),
       ��� Varchar2(5),
       ���� Varchar2(30),
       ���� Varchar2(10),
       ˵�� Varchar2(100),
       ͨ�ü� Number(1),
       ����ID Number(18),
       ��ԱID Number(18))
    TABLESPACE zl9EprDat;

Create Table �������İ����(
       ���İ�ID Number(18),
       ����ID  Number(18))
    TABLESPACE zl9EprDat;

--�������
Create Table ������鷽��(
    ID		Number(18),
    ����	Varchar2(50),
    �ܷ�	Number(5,2),
    �ֶ���	Number(5,2),
    ����ʱ��	Date,
    ͣ��ʱ��	Date,
    ˵�� VARCHAR2(200))
    TableSpace zl9BaseItem;

Create Table ����������(
    ID		Number(18),
    �ϼ�id	Number(18),
    ����	Varchar2(10),
    ����	Varchar2(30),
    ����ID	Number(18))
    TableSpace zl9BaseItem;

Create Table �������Ŀ¼(
    ID		Number(18),
    ����ID	Number(18),
    ����	Varchar2(10),
    ����	Varchar2(255),
    ����	Varchar2(255),
    ˵��	Varchar2(2000),
    �������	Varchar2(4000),
    ���ö���	Number(3),
    �ļ�ID	Varchar2(2000),
    ���û���	Varchar2(1),
    ��ֵ	Number(5,2),
    ����	Number(1),
	����Դ Number(1) Default 0)
    TableSpace zl9BaseItem;

----------------------------------------------------------------------------
--[[9.�������]]
----------------------------------------------------------------------------
Create Table �����¼��Ŀ(
    ��Ŀ��� NUMBER(5),
    ��Ŀ���� VARCHAR2(20),
    ������Ŀ NUMBER(1),
    ��Ŀ���� NUMBER(3),
    ��Ŀ���� NUMBER(3),
    ��ĿС�� NUMBER(3),
    ��Ŀ��λ VARCHAR2(10),
    ��Ŀ��ʾ NUMBER(3),
    ��Ŀֵ�� VARCHAR2(4000),
    ��Ŀ���� Number(3),
    ��ĿID NUMBER(18),
    ����ȼ� NUMBER(3),
    ���ÿ��� Number(3),
    ���ò��� Number(3),
    Ӧ�÷�ʽ Number(3),
    �������� Varchar2(2),
    ������ VARCHAR2(20),
    ˵�� VARCHAR2(1000),
    Ӧ�ó��� number(1)  DEFAULT 0)
    TABLESPACE zl9EprDat;

    --����������,,ͬһ���ϵĲ���������ܽ϶�

CREATE TABLE ������Ŀģ��(
	����ID NUMBER (18),
	ģ������ VARCHAR2 (50),
	����ȼ� NUMBER (3),	--0-�ؼ�;1-һ��;2-����;3-����;-1���޻���ȼ�
	��Ŀ��� NUMBER (5),
	������� NUMBER (3))
	TABLESPACE zl9EprDat;

CREATE TABLE ���²���(
	���� VARCHAR2 (50),
	���õ��� VARCHAR2 (50),
	���� VARCHAR2 (50),
	�²��� VARCHAR2 (50),
	���� NUMBER (1) DEFAULT 0)
	TABLESPACE zl9EprDat;

CREATE TABLE ���²�λ(
    ��Ŀ��� NUMBER (5),
    ��λ VARCHAR2 (50),
    ��Ƿ��� VARCHAR2 (10),
    �����ɫ NUMBER (18),
    ���ͼ�� BLOB,
	ȱʡ�� NUMBER (1) DEFAULT 0,
	�̶��� NUMBER(1) DEFAULT 0)
    TABLESPACE ZL9EPRDAT;

CREATE TABLE �����ص����(
    ���	NUMBER(5),
    �ϼ����	NUMBER(5),
    ��Ŀ���	NUMBER(5),
    ���²�λ	VARCHAR2(10),
    �ص���Ŀ	Number(5),
    �ص���Ŀ	Varchar2(2000),
    ��Ƿ���	Varchar2(10),
    �����ɫ	Number(18),
    ���ͼ��	Blob)
    TABLESPACE zl9EprDat;

CREATE TABLE �������ÿ���(
    ��Ŀ��� NUMBER(5),
    ����id NUMBER(18))
    TABLESPACE zl9EprDat;

CREATE TABLE ���¼�¼��Ŀ(
    ��Ŀ��� NUMBER(5),
    ������� NUMBER(3),
    ��¼�� VARCHAR2(20),
    ��¼�� NUMBER(3),
    ��¼�� VARCHAR2(10),
    ��¼ɫ NUMBER(18),
    ���ֵ NUMBER(16,5),
    ��Сֵ NUMBER(16,5),
    ��λֵ NUMBER(16,5),
    ��¼Ƶ�� Number(1),
    ��λ Varchar2(10),
    ����� NUMBER(5),
    �̶ȼ�� NUMBER (16,5),
    ��ʾ�� NUMBER (16,5),
    �ٽ�ֵ Varchar2(30),
    ��Ժ�ײ� NUMBER (1) DEFAULT 0)
    TABLESPACE zl9EprDat;

CREATE TABLE �������ʱ��(
	���� VARCHAR2 (20),
	��ʼ VARCHAR2 (5),
	���� VARCHAR2 (5),
	��� NUMBER (1) DEFAULT 1,
	���� NUMBER (1) DEFAULT 1,
	����ID NUMBER(18))
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ���������Ŀ(
	��� NUMBER (5),
	����� NUMBER (5))
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ��������Ŀ(
	��Ŀ��� NUMBER (5),
	��Ŀ���� varchar2(20))
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ����ͬ����Ŀ(
	����ȼ� NUMBER(3),
	���䷶Χ Varchar2(50),
	������Ŀ Varchar2(100),
	���ÿ��� varchar2(200))
	TABLESPACE ZL9EPRDAT;

CREATE TABLE �����������(
	����ID NUMBER (18),
	������� NUMBER(18),
	������ NUMBER (5),
	˵�� VARCHAR2(20),
	ͼ������ NUMBER (5),
	��Ч���� NUMBER (5))
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ������������ʽ(
	ID NUMBER(18),
	����ID NUMBER (18),
	���� VARCHAR2 (20) NOT NULL,
	���� VARCHAR2 (20),
	������Ŀ XMLTYPE,
	�к� NUMBER (18),
	λ�� NUMBER (1) DEFAULT 1,
	�Ƿ�̶� NUMBER (1),
	�Ƿ����� NUMBER (1),
	���� VARCHAR2 (500),
	ʱ�� DATE )
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ������ĿƵ��(
	Ƶ�� NUMBER (1),
	��� NUMBER (1),
	��ʼ VARCHAR2 (5),
	���� VARCHAR2 (5),
	��� NUMBER (1) DEFAULT 1)
	PCTFREE 20 INITRANS 10  
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ���̲���(
	���� VARCHAR2 (50),
	���õ��� VARCHAR2 (50),
	���� VARCHAR2 (50),
	���� NUMBER (1) DEFAULT 0)
	TABLESPACE zl9EprDat;

CREATE TABLE ����Ҫ������(
	�ļ�ID NUMBER (18),
	Ӥ�� NUMBER (1),
	���� VARCHAR2 (60),
	���� VARCHAR2 (100),
	��ת�� Number(3))
	TABLESPACE zl9EprDat;

CREATE TABLE �����¼��Ŀ(
    ��� NUMBER(3),
    ��¼�� VARCHAR2(20),
    ��¼�� VARCHAR2(2),
    ��¼ɫ NUMBER(18),
    ���ֵ NUMBER(16,5),
    ��Сֵ NUMBER(16,5),
    ��λ Varchar2(10),
    ��¼�� number(3),
    ���� number(1),
    ��ĿID NUMBER(18))
    TABLESPACE zl9EprDat;

----------------------------------------------------------------------------
--[[10.�������]]
----------------------------------------------------------------------------
CREATE TABLE ���鱨����Ŀ(
    ID         Number(18),
    ������ĿID NUMBER(18),
    ����걾 VARCHAR2(20),
    ������� NUMBER(5),
    ������ĿID NUMBER(18),
    ϸ��ID NUMBER(18))
    TABLESPACE zl9BaseItem;

Create Table ���鱸ע����(
	���� varchar2(10),
	���� varchar2(100) not Null,
	���� varchar2(10),
	˵�� varchar2(80),
	���� varchar2(20))
    TABLESPACE zl9BaseItem;

Create Table ���鿹������(
	ID number(18),
	���� varchar2(10),
	���� varchar2(50),
	Ӣ�� Varchar2(50),
	���� Varchar2(20))
    TABLESPACE zl9BaseItem;

Create Table �����ÿ�����(
	ID number(18),
	���� varchar2(10),
	������ varchar2(50),
	Ӣ���� varchar2(50),
	���� varchar2(20),
	˵�� varchar2(100),
	ҩ������ Number(1),
	WHONET�� Varchar2(10),
	�÷�����1 Varchar2(30),
	ѪҩŨ��1 Varchar2(30),
	��ҩŨ��1 Varchar2(30),
	�÷�����2 Varchar2(30),
	ѪҩŨ��2 Varchar2(30),
	��ҩŨ��2 Varchar2(30))
    TABLESPACE zl9BaseItem;

Create Table ���鿹������ҩ(
	������ID number(18),
	�����ط���ID number(18))
    TABLESPACE zl9BaseItem;

Create Table ������������(
	���� varchar2(10),
	���� varchar2(100) not Null,
	���� varchar2(10),
	˵�� varchar2(80))
    TABLESPACE zl9BaseItem;

Create Table ������������(
    ���� VARCHAR2(3),
    ���� VARCHAR2(50),
    ���� VARCHAR2(10),
    ˵�� VARCHAR2(80),
    ���� VARCHAR2(20))
    TABLESPACE zl9BaseItem;

Create Table ����ϸ������(
	ID number(18),
	���� varchar2(13),
	�������� varchar2(40),
	Ӣ������ varchar2(40),
	���� varchar2(10))
    TABLESPACE zl9BaseItem;

Create Table ����ϸ�����(
	���� varchar2(8),
	���� varchar2(30),
	���� varchar2(20),
	ȱʡ��־ NUMBER(1) DEFAULT 0)
	TABLESPACE zl9BaseItem;

Create Table ����ϸ������(
	���� varchar2(8),
	���� varchar2(30),
	���� varchar2(20))
	TABLESPACE zl9BaseItem;

Create Table ����Ⱦɫ����(
	���� varchar2(8),
	���� varchar2(30),
	���� varchar2(20),
	ȱʡ��־ NUMBER(1) DEFAULT 0)
	TABLESPACE zl9BaseItem;

Create Table ����ϸ��(
	ID number(18),
	���� varchar2(10),
	������ varchar2(100),
	Ӣ���� varchar2(100),
	����ID number(18),
	���� varchar2(10),
	Ĭ��ҩ�� Varchar2(1),
	Ĭ�Ϸ��� Varchar2(20),
	WHONET�� Varchar2(10),
	Ĭ�Ͻ�� varchar2(50),
	ϸ����� varchar2(30),
	ϸ������ varchar2(30),
	�����Ϸ���  varchar2(30))
    TABLESPACE zl9BaseItem;

Create Table ����ϸ��������(
	ϸ��ID number(18),
	�����ط���ID number(18),
	ȱʡ��־ number(18))
    TABLESPACE zl9BaseItem;

Create Table ������Ŀ(
	������ĿID number(18),
	��д varchar2(40),
	������� varchar2(10),
	��Ŀ��� number(1),
	������� number(1),
	��λ varchar2(20),
	��ӡ���� number(18),
	��ӡ��� number(18),
	���㹫ʽ varchar2(500),
	���鷽�� varchar2(40),
	�ϲ������ varchar2(10),
	����쳣���� varchar2(10),
	�����Χ Varchar2(20),
	Ĭ��ֵ Varchar2(200),
	�������� Number(16,5),
	�������� Number(16,5),
	���챨���� Number(16,5),
	�ȶԾ�ʾ�� Number(16,5),
	�ȶ�ʧ���� Number(16,5),
	ȡֵ���� Varchar2(200),
	��˽��Ŀ Number(1),
	���Թ�ʽ varchar2(50),
	�����Թ�ʽ varchar2(50),
	CutOff��ʽ varchar2(50),
	������� Number(18),
	���쾯ʾ�� Number(16,5),
	�ٴ����� varchar2(4000),
	��ο� number(1))
    TABLESPACE zl9BaseItem;

Create Table ������Ŀ�ο�(
	ID     Number(18),
	��ĿID number(18),
	�걾���� varchar2(20),
	�Ա��� number(1),
	�������� number(18),
	�������� number(18),
	���䵥λ varchar2(10),
	�ο���ֵ number(21,4),
	�ο���ֵ number(21,4),
	��ע varchar2(50),
	����ID number(18),
	�������ID Number(18),
	�ٴ����� Varchar2(30),
	��ƫ���� Number(10,2),
	Ĭ�� number(1),
	��ʾ���� NUMBER(16,5),
	��ʾ���� NUMBER(16,5),
	�������� NUMBER(16,5),
	�������� NUMBER(16,5))
    TABLESPACE zl9BaseItem;

Create Table ������Ŀȡֵ(
	��ĿID number(18),
	���� varchar2(10),
	ȡֵ varchar2(10),
	�����־ number(1))
    TABLESPACE zl9BaseItem;

Create Table ����걾��̬(
	���� varchar2(10),
	���� varchar2(50) not Null,
	˵�� varchar2(100))
    TABLESPACE zl9BaseItem;

Create Table ��������(
	ID number(18),
	���� varchar2(10),
	���� varchar2(20),
	���� varchar2(10),
	���Ӽ���� varchar2(40),
	ͨѶ������ varchar2(40),
	ͨѶ�˿� varchar2(10),
	������ number(6),
	���� number(2),
	ֹͣλ number(2,1),
	У��λ varchar2(4),
	�������� varchar2(50),
	������־ɫ varchar2(10),
	ʹ��С��ID number(18),
	�ʿر걾�� varchar2(40),
	��ע VARCHAR2(100),
	�ʿ����� Number(5),
	���ڵ�λ Varchar2(2),
	�ʿ�ˮƽ�� Number(1),
	�ϴ��ʿ��� Date,
	QC�� Varchar2(8),
	�Լ���Դ Varchar2(30),
	У׼����Դ Varchar2(30),
	΢���� Number(1),
	ת������ Date,
	ת������ID Number(18),
	���� varchar2(60),
	���Ƶ�� varchar(30),
	���ʱ�� varchar(5),
	���巽ʽ varchar(30),
	�հ���ʽ varchar(30),
	�����ʿ�ͼ Number(1),
	����ʱָ������ Number(1))
    TABLESPACE zl9BaseItem;

Create Table ����������Ŀ(
	��ĿID number(18),
	����ID number(18),
	ͨ������ varchar2(20),
	С��λ�� number(18),
	���λ�� number(18),
	ȱʡ���� number(1),
	����ֵ Number(16,5),
	����� Number(16,5),
	������ID number(18),
	��������Ŀ number(1))
    TABLESPACE zl9BaseItem;

Create Table �����ʿع���(
    ID Number(18),
    ���� Number(1),
    ���� Varchar2(3),
    ���� Varchar2(20),
    ˵�� Varchar2(100),
    ��ʽ Number(1),
    ��ˮƽ Number(1),
    N Number(2),
    X Number(5, 1),
    M Number(2),
    P Number(5, 3),
    K Number(5, 1),
    H Number(5, 1))
    Tablespace zl9BaseItem;

Create Table �����ʿ�Ʒ(
    ID Number(18),
    ����ID Number(18),
    ���� Varchar2(50),
    ���� Varchar2(10),
    Ũ�� Varchar2(30),
    ˮƽ Number(1),
    ���� Varchar2(30),
    ��ʼ���� Date,
    �������� Date,
    �Ƕ�ֵ Number(1),
    �걾�� Varchar2(40),
    �Լ� varchar2(30),
    У׼�� varchar2(30))
    Tablespace zl9BaseItem;

Create Table �����ʿ�Ʒ��Ŀ(
	�ʿ�ƷID NUMBER(18),
	��ĿID   NUMBER(18),
	��ֵ     NUMBER(18,4),
	SD       NUMBER(18,4),
	CV       NUMBER(18,4),
	��ĿQC�� VARCHAR2(8),
	����QC�� VARCHAR2(8),
	����     VARCHAR2(30),
	ȡֵ���� VARCHAR2(500),
	����ֵ VarChar2(500),
	�ʿ�ȡֵ VarChar2(100))
    TABLESPACE zl9BaseItem;

Create Table ����С��(
  ID NUMBER(18),
  ���� VARCHAR2(10),
  ���� VARCHAR2(50))
    TABLESPACE zl9BaseItem;

Create Table ����С���Ա(
  С��ID NUMBER(18),
  ��ԱID NUMBER(18),
  Ĭ��С�� NUMBER(1),
  ��ע   VARCHAR2(100))
    TABLESPACE zl9BaseItem;

Create Table ����С������(
  С��ID NUMBER(18),
  ����ID NUMBER(18),
  �鿴   Number(1),
  ����   Number(1),
  �������� Number(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE ���Ƽ���걾(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(8),
		�����Ա� VARCHAR2(4))
    TABLESPACE zl9BaseItem;

CREATE TABLE ���Ƽ�������(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(8),
    ȱʡ��־ NUMBER(1),
    ���� Varchar2(2))
    TABLESPACE zl9BaseItem;

Create Table �����Լ���ϵ(
    ID     Number(18),
    ��ĿID Number(18),
    ����ID Number(18),
    ����ID Number(18),
	���� number(16,5),
	�̶� number(1))
    TABLESPACE zl9BaseItem;

Create Table ��Ѫ������(
    ���� Varchar2(4),
    ���� Varchar(30),
    ���� Varchar2(10),
    ��Ӽ� Varchar2(30),
    ��Ѫ�� Varchar2(30),
    ��� Varchar2(30),
    ��ɫ number(10),
    ����ID number(18))
    TABLESPACE zl9CisRec;

Create Table ����������(
    ���� VARCHAR2(3),
    ���� VARCHAR2(200),
    ���� VARCHAR2(10),
    ���� VARCHAR2(20))
    TABLESPACE zl9CisRec;

Create Table �ٴ�����(
    ���� VARCHAR2(2),
    ���� VARCHAR2(30),
    ���� VARCHAR2(10))
    TABLESPACE zl9CisRec;


Create Table ������Ŀѡ��(
    ������ĿID Number(18),
    ��������ID Number(18),
    ���������ֽ� Number(1),
    סԺ����ID Number(18),
    סԺ�����ֽ� Number(1),
    �������ID Number(18),
    ��������ֽ� Number(1),    
    �������� Number(5),
    ��ʱ��׼ Number(5),
    ��ʱ��λ Varchar2(4),
    ȡ����ص� Varchar2(50),
    ����˵�� Varchar2(200),
    �ͼ�ʱ�� number(4),
    �����ʱ Number(5))
    TABLESPACE zl9CisRec;

Create Table ����ϸ������(
    ����ID Number(18),
    ͨ������ Varchar2(50),
    ϸ��ID Number(18),
    ������ID Number(18))
    TABLESPACE zl9CisRec;

Create Table ������������(
    ID Number(18),
    �ϼ�ID Number(18),
    ����ID Number(18),
    ��ĿID Number(18),
    �ж� Varchar2(80),
    ����ID Number(18),
    ���� Varchar2(1),
    ����Χ Number(3),
    ��ˮƽ Number(1),
    Y��Ǽ� Number(1),
    Y���� Varchar2(20),
    Y���� Number(1),
    Y��ʾ Varchar2(500),
    N��Ǽ� Number(1),
    N���� Varchar2(20),
    N���� Number(1),
    N��ʾ Varchar2(500),
    �Ƿ�ʹ�� Number(1))
    Tablespace zl9BaseItem;

Create Table ϸ����ⷽ��(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(10))
    TABLESPACE zl9CisRec;

Create Table ϸ����ҩ����(
    ���� VARCHAR2(4),
    ���� VARCHAR2(100),
    ���� VARCHAR2(20))
    TABLESPACE zl9CisRec;

Create Table ����������Ŀ(
    �걾ID Number(18),
    ������ĿID Number(18),
    ��� Number(3),
	��ת�� Number(3))
    TABLESPACE zl9CisRec;

Create Table ����ģ��Ŀ¼(
    ID Number(18),
    ���� Varchar2(6),
	���� Varchar2(30),
	���� Varchar2(10),
    ������ĿID Number(18),
    ˵�� Varchar2(50),
    ������ Varchar2(20),
    ����ʱ�� Date,
    �������� Varchar2(100),
    ���鱸ע Varchar2(400))
    TABLESPACE zl9CisRec;

Create Table ����ģ������(
    ID Number(18),
    ģ��ID Number(18),
    ��ĿID Number(18),
    ������ Varchar2(60),
    ϸ��ID Number(18),
    �������� Varchar2(50))
    TABLESPACE zl9CisRec;

Create Table ����ģ��ҩ��(
    ϸ�����ID Number(18),
    ������ID Number(18),
    ��� Varchar2(20),
    ������� Varchar2(20),
    ҩ������ Number(3))
    TABLESPACE zl9CisRec;

Create Table ����ϲ�����(
	����ĿID number(18),
	�ϲ���ĿID number(18) not null)
    TABLESPACE zl9CisRec;

Create Table �ʿؼ��鷽��(
    ���� Varchar2(6),
    ���� Varchar2(30),
    ���� Varchar2(10))
    Tablespace zl9CisRec;

Create Table �ʿ��Լ���Դ(
    ���� Varchar2(6),
    ���� Varchar2(30),
    ���� Varchar2(10),
    QC���� Varchar2(8))
    Tablespace zl9CisRec;

Create Table �ʿر���ʾ�(
    ���� Varchar2(3),
    ���� Varchar2(80),
    ���� Varchar2(10),
    ���� Varchar2(4))
    Tablespace zl9CisRec;

Create Table �ʿؼ��̷�(
    N Number(3),
    N3S Number(6,2),
    N2S Number(6,2))
    Tablespace zl9CisRec;

Create Table �ʿؿؽ�ϵ��(
    ���� Varchar2(20),
    ���� Number(1),
    N2 Number(6,2),
    N3 Number(6,2),
    N4 Number(6,2),
    N6 Number(6,2),
    N7 Number(6,2),
    N10 Number(6,2),
    N12 Number(6,2),
    N16 Number(6,2),
    N20 Number(6,2),
    �д� Number(2))
    Tablespace zl9CisRec;

Create Table ��������״̬(
    ����ID Number(18),
    ��ĿID Number(18),
    ʧ�ر�� Varchar2(100),
    ʧ������ Date)
    Tablespace zl9CisRec;

Create Table �����ʿؾ�ֵ(
    �ʿ�ƷID Number(18),
    ��ĿID Number(18),
    �ڼ� Varchar2(20),
    ��ʼ���� Date,
    �������� Date,    
    ��ֵ Number(18,4),
    SD Number(18,4),
    CV Number(18,4),
    �������� Date,
    ������ Varchar2(20))
    Tablespace zl9CisRec;

Create Table �����ʿط���(
    ������ Varchar2(30),
    ˮƽ�� Number(1),
    ��� Number(18),
    �ϼ� Number(18),
    �ж� Varchar2(80),
    ������ Varchar2(20),
    ��ʽ Number(1),
    N Number(2),
    X Number(5, 1),
    M Number(2),
    ���� Varchar2(1),
    ����Χ Number(3),
    ��ˮƽ Number(1),
    Y��Ǽ� Number(1),
    Y���� Varchar2(20),
    Y���� Number(1),
    Y��ʾ Varchar2(500),
    N��Ǽ� Number(1),
    N���� Varchar2(20),
    N���� Number(1),
    N��ʾ Varchar2(500))
    Tablespace zl9CisRec;

CREATE TABLE ����������(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(8),
    ȱʡ��־ NUMBER(1))
    TABLESPACE zl9BaseItem;

create table ������˹���(
  ID NUMBER(18),
  ����  Varchar2(3),
  ����  VARCHAR2(30),
  ����   VARCHAR2(30),
  ��ĿID NUMBER(18),
  ����ID NUMBER(18),
  ����ID NUMBER(18),
  �������� VARCHAR2(1),
  �Ա�   VARCHAR2(4),
  �������� NUMBER(9),
  �������� NUMBER(9),
  ���䵥λ VARCHAR2(4),
  ��� VARCHAR2(500),
  ����   VARCHAR2(4000),
  ������� VARCHAR2(4000),
  �����ϵ VARCHAR2(3),
  ��ʾ��Ϣ VARCHAR2(200),
  ����   VARCHAR2(1),
  ��Ч   VARCHAR2(1),
  ���   VARCHAR2(1),
  ��ע   VARCHAR2(200))
  Tablespace zl9BaseItem;

create table ����ϸ�������زο�
(
  ϸ��ID       NUMBER(18),
  �����ط���ID NUMBER(18),
  ������ID     NUMBER(18),
  ҩ������     NUMBER(1),
  �ο���ֵ     NUMBER(21,4),
  �ο���ֵ     NUMBER(21,4),
  �жϷ�ʽ     NUMBER(1),   ---- 1-�����ο�ֵ,0-�ο�ֵ����
  ��ע         VARCHAR2(500),
  ��ֵ���     VARCHAR2(30),
  �м���     VARCHAR2(30),
  ��ֵ���     VARCHAR2(30)
)tablespace ZL9BASEITEM;

Create Table �����������(
	���� varchar2(10),
	���� varchar2(200) not Null)
    TABLESPACE zl9BaseItem;    

 Create Table ���������;(
	���� varchar2(10),
	���� varchar2(200) not Null)
    TABLESPACE zl9BaseItem;

Create Table ����ø��ģ��(
	ID	 NUMBER(18),
	���	 Number(3),
	����     VARCHAR2(20),
	��Ŀ     VARCHAR2(1000),	--��Ŀ��ʽ����ĿA;��ĿB;...��ĿH ��8����Ŀ
	����     VARCHAR2(2000))	--���ݸ�ʽ: ���1;���2...���12|���1;���2...���12 ��8��ÿ��12�����
    TABLESPACE zl9BaseItem;

----------------------------------------------------------------------------
--[[11.������]]
----------------------------------------------------------------------------
Create Table ҽ��ִ�з���(
       ����id Number(18),
       ִ�м� Varchar(20),
       ����   Varchar(20),
       ��ǰ���� Number(1),
       ����豸 Varchar2(3),
       ����ǰ׺ varchar2(10),
       ����ID Number(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE ���Ƽ������(
    ���� VARCHAR2(2),
    ���� VARCHAR2(20),
    ���� VARCHAR2(8),
    ������ NUMBER(1))
    TABLESPACE zl9BaseItem;

Create Table ������Ŀ��λ(
    ID Number(18),
    ��ĿID Number(18),
    ���� Varchar2(20),
    ��λ Varchar2(30),
		���� Varchar2(30),
		Ĭ�� Number(1))
    Tablespace zl9BaseItem;

Create Table ���Ƽ�鲿λ(
    ���� Varchar2(20),
    ���� Varchar2(4),
    ���� Varchar2(30),
    ���� Varchar2(30),
    ��ע Varchar2(200),
    ���� Varchar2(1000),
    �����Ա� Number(1))
    Tablespace zl9BaseItem;

Create Table ��Ӱ��(
    ���� Varchar2(2),
    ���� Varchar2(30),
    ���� Varchar2(10))
    Tablespace zl9BaseItem;

Create Table ��ݹ�����Ϣ(
    ID Number(18), 
    ��Ŀ varchar2(128),
    ģ��� number(18),
    �˵����� varchar2(128),
    ������� number(18) default 0,
    �˵�ID Number(18),    
    �˵�˵�� varchar2(128),
    ���Ƽ� Number(18),
    �ַ��� Number(18),
    Ĭ�ϼ� varchar2(64),
    ����� varchar2(64)
    )
    TABLESPACE zl9BaseItem; 

Create Table ��ݹ��ܹ���
(
   ID number(18),
   ��ݹ���ID number(18),
   �û�ID number(18),
   ���Ƽ� number(18),
   �ַ��� number(18),
   ����� varchar2(64))
   TABLESPACE zl9BaseItem;
    
Create Table Ӱ��ִ�з���(
       ID Number(18),
       ����id Number(18),
       ���� Varchar2(30),
       ����ǰ׺  Varchar2(10)
       )
    TABLESPACE zl9BaseItem;
    
Create Table Ӱ��������(
       ID Number(18),
       ����ID Number(18),
       ����ID Number(18),
       ������ĿID Number(18)
       )
    TABLESPACE zl9BaseItem;
           
create table Ӱ���˾�ģ��
(
  ID         NUMBER(18) not null,
  Ӱ������   VARCHAR2(30),
  �˾�����   VARCHAR2(50),
  ��ǿǿ������ NUMBER(3),
  ��ǿǿ�ȼ��� NUMBER(3),
  ��ǿ�������� NUMBER(3),
  ��ǿ���ȼ��� NUMBER(3),
  ƽ������     NUMBER(3),
  ƽ������     NUMBER(3))
    TABLESPACE zl9BaseItem;

CREATE TABLE Ӱ��ͼ��ע(
    ���� VARCHAR2(5),
    ���� VARCHAR2(100),
    ���� VARCHAR2(20),
    ��Ա VARCHAR2(20))
    TABLESPACE zl9BaseItem;

create table Ӱ��MWL��λ����
(
  ID           number(18),
  ����ID       number(18),
  PACS��λ���� varchar2(30),
  �豸��λ���� varchar2(64),
  �豸��λ���� varchar2(64)
)
    TABLESPACE zl9BaseItem;

Create Table Ӱ�������¼(
  IP��ַ VARCHAR2 (15),
  ���� VARCHAR2 (20),
  ����ʱ�� DATE)
    TABLESPACE zl9BaseItem;
    
Create Table Ӱ���ѯ����(
       Id Number(18),
       �������� varchar2(30),
       ����˵�� varchar2(512),
       ��ѯ��� varchar2(1024),
       �Ƿ�Ĭ�� Number(1) default 0,
       ʹ��״̬ Number(1) default 1,
       ������� Number(18),
       �������� Number(18) default 0,
	   �Ƿ����ù��� Number(1) default 0,
       �Ƿ�ϵͳ��ѯ Number(1) default 0
)TABLESPACE zl9BaseItem;     

Create Table Ӱ���ѯ����(
       Id Number(18),
       ����ID Number(18),
       ¼����Ŀ varchar2(30),
       ¼������ Number(1),
       Ĭ��ֵ   varchar2(512),
       ������Դ varchar2(1024),
       ¼��˳�� number(18)
)TABLESPACE zl9BaseItem;


Create Table Ӱ����Ϸ���(
  ���� VARCHAR2(2),
  ���� VARCHAR2(20),
  ���� VARCHAR2(8),
  �������� VARCHAR2(20))
    TABLESPACE zl9BaseItem;

create table Ӱ�����̲���
(
    ID     NUMBER(18),
    ����ID NUMBER(18),
    ������ VARCHAR2(100),
    ����ֵ VARCHAR2(1000))
    TABLESPACE zl9BaseItem;

Create Table Ӱ�������(
    ���� varchar2(10),
    ���� varchar2(20),
    ���� varchar2(10),
    ���� number(3),
    ������ number(18))
    TABLESPACE zl9BaseItem;

Create Table Ӱ������Ŀ(
    ������Ŀid number(18),
    Ӱ����� varchar2(10),
    ���в��� number(1),
    �ɷ���Ƭ number(1),
    ���׼�� varchar2(50),
    ����ͼ�� number(1))
    TABLESPACE zl9BaseItem;

Create Table Ӱ���豸Ŀ¼(
    �豸�� varchar2(3),
    �豸�� varchar2(100),
    ���� number(1),
    IP��ַ varchar2(15),
    �˿ں� varchar2(5),
    ����Ŀ¼ varchar2(100),
    FTPĿ¼ varchar2(100),
    Ŀ¼�� number(1),
    FTP�û��� varchar2(20),
    FTP���� varchar2(20),
    ����Ŀ¼�û��� VARCHAR2(20),
    ����Ŀ¼���� VARCHAR2(20),
    ����Ŀ¼ VARCHAR2(100),
    ����AE varchar2(20),
    �豸AE varchar2(20),
    ״̬ NUMBER(1),
    Ӱ����� VARCHAR2(20))
    TABLESPACE zl9BaseItem;

--Ӱ�����
Create Table Ӱ����ɫ�嵥(
    ��� number(5),
    ��ɫ varchar2(20),
    ��ɫ���� varchar2(4000),
    ϵͳ���� number(1))
    TABLESPACE zl9BaseItem;

Create Table Ӱ���ӡ��ʽ(
    ���� number(5),
    ���� varchar2(50),
    ��ʽ varchar2(20),
    ����1 number(3),
    ����2 number(3),
    ����3 number(3),
    ����4 number(3),
    ����5 number(3),
    ����6 number(3),
    ����7 number(3))
    TABLESPACE zl9BaseItem;

Create Table Ӱ��Ƭ���(
    ���� number(5),
    ���� varchar2(20),
    ��Ƭ��� varchar2(20),
    ��Ƭ���� varchar2(20),
    ��λ varchar2(20))
    TABLESPACE zl9BaseItem;

Create Table Ӱ���ע�洢��(
    ��� number(5),
    VGroup varchar2(20),
    Element varchar2(20),
    VR varchar2(20),
    ��ע���� varchar2(20))
    TABLESPACE zl9BaseItem;

Create Table Ӱ��ͼ����Ϣ��(
    ID number(5),
    ��ʼ��ַ varchar2(20),
    ������ַ varchar2(20),
    Ӣ������ varchar2(50),
    �������� varchar2(50),
    ���ļ�� varchar2(50),
    Ӣ�ļ�� varchar2(50),
    ���� number(1),
    ��ѡ�� number(1),
    λ�� number(2),
    ������� number(2),
    �ɵ��� number(1),
    ʹ�ü��� number(1))
    TABLESPACE zl9BaseItem;

Create Table Ӱ����������(
    ��ԱID number(18),
    ����ͼ��߿���ɫ varchar(20),
    ����ͼ��߿����� varchar(5),
    ����ͼ��߿��߿� varchar(5),
    ѡ��ͼ��߿���ɫ varchar(20),
    ѡ�����б߿���ɫ varchar(20),
    ѡ��ͼ��߿����� varchar(5),
    ѡ��ͼ��߿��߿� varchar(5),
    ͼ������ɫ varchar(20),
    ͼ���Ǵ�С varchar(5),
    ��עѡ������ɫ varchar(20),
    ��עѡ������С varchar(5),
    ��λ����ɫ varchar(20),
    ��λ������ varchar(5),
    ��λ�߼�� varchar(5),
    ���м��� varchar(5),
    ����������� varchar(5),
    ����������� varchar(5),
    ͼ���� varchar(5),
    ��ʾ����߿� varchar(5),
    ������ɫ varchar(20),
    ���򱳾���ɫ varchar(20),
    ��ע������ɫ varchar(20),
    ��ע�������� varchar(5),
    ��ע�����߿� varchar(5),
    ��עѡ����ɫ varchar(20),
    ��עѡ������ varchar(5),
    ��עѡ���߿� varchar(5),
    ��ע���ִ�С varchar(5),
    ������ʾ��� varchar(5),
    ������ʾƽ��ֵ varchar(5),
    ������ʾ������ varchar(5),
    ������ʾ���� varchar(20),
    ����X����ƫ�� varchar(5),
    ����Y����ƫ�� varchar(5),
    ������ͼ������ varchar(5),
    ��ʾ��λ��� varchar(20),
    ������λ��� varchar(20),
    ��ʾ��� varchar(5),
    ������ұ߾� varchar(5),
    ������±߾� varchar(5),
    ��߿�� varchar(5),
    ��߸߶� varchar(5),
    ����߿� varchar(5),
    �����ɫ varchar(20),
    ����λλ�� varchar(5),
    ��괩�󲽳� varchar(5),
    ������β��� varchar(5),
    ���������� varchar(5),
    ������Ų��� varchar(5),
    ������Ϣ���±߾� varchar(5),
    ������Ϣ���ұ߾� varchar(5),
    ������Ϣ��ɫ varchar(20),
    ������Ϣ��ʾ��Сֵ varchar(5),
    ������Ϣ��ͼ������ varchar(5),
    ������Ϣ���� varchar(50),
    ������Ϣ��ͷ varchar(20),
    ֱ������ varchar(5),
    ������ͼ���С varchar(5),
    ������λ�� varchar(5),
    ��������ʾ varchar(5),
    ״̬�������С varchar(5),
    ����Ѫ����ֵ varchar(10),
    ��խѪ����ֵ varchar(10),
    Ѫ�ܱڿ�� varchar(10),
    ������ʾ�ܳ� varchar(10),
    ������ʾ���ֵ varchar2(10),
    ������ʾ��Сֵ varchar2(10),
    �����ֲ��� varchar2(2),
    ��ʾ��ӡ��� VARCHAR2(2))
    TABLESPACE zl9BaseItem;

Create Table Ӱ����갴ť����(
    ��ԱID number(18),
    ֱ�� varchar2(200),
    ���� varchar2(200),
    ��Բ varchar2(200),
    ��ͷ varchar2(200),
    ����� varchar2(200),
    ����� varchar2(200),
    �Ƕ� varchar2(200),
    ���� varchar2(200),
    ����λ varchar2(200),
    ����λ varchar2(200),
    ���� varchar2(200),
    ���� varchar2(200),
    �ü�_��ע���� varchar2(200),
    ����Ӧ���� varchar2(200),
    ��ά��� varchar2(200),
    ����ע varchar2(200))
    TABLESPACE zl9BaseItem;

Create Table Ӱ��ͼ��������(
    ID  number(18),
    ��ԱID number(18),
    Ӱ������ varchar2(30),
    �������� varchar2(30),
    Բ��X number(10),
    Բ��Y number(10),
    Բ�ΰ뾶 number(10),
    ������߽� number(10),
    �����ұ߽� number(10),
    �����ϱ߽� number(10),
    �����±߽� number(10),
    ����ζ��� varchar2(50),
    ������ɫ number(20))
    TABLESPACE zl9BaseItem;

Create Table Ӱ��Ԥ�贰��λ(
    ID  number(18),
    ��ԱID  NUMBER(18),
    Ӱ������ varchar2(30),
    ��ݼ� number(5),
    �������� varchar2(50),
    ����Ӣ���� varchar2(60),
    ���� number(10),
    ��λ number(10),
    �Ƿ�Ĭ�� number(5))
    TABLESPACE zl9BaseItem;

Create Table Ӱ����Ļ����(
    ID  number(18),
    ��ԱID number(18),
    Ӱ������ varchar2(30),
    �Զ����в��� number(5),
    �Զ�ͼ�񲼾� number(5),
    �������� number(5),
    �������� number(5),
    ͼ������ number(5),
    ͼ������ number(5),
    �Զ����� number(1),
    ��ʾ������Ϣ number(1),
    ѡ��λ�� number(1),
    ѡ������ͬ�� number(1),
    ��ֵģʽ number(1))
    TABLESPACE zl9BaseItem;

Create Table Ӱ���ӡ������(
    ID  number(18),
    ��ӡ���� varchar2(50),
    IP��ַ varchar2(18),
    �˿ں� number(5),
    AE���� varchar(50),
    ��ӡ��ʽ varchar(50),
    ���ȼ� varchar(30),
    ��ӡ���� number(5),
    ���� varchar(30),
    ���� varchar(30),
    ��Ƭ��� varchar(30),
    ѡ��Ƭ�� varchar(30),
    �ֱ��� varchar(30),
    �Ŵ�ģʽ varchar(30),
    ƽ��ģʽ varchar(30),
    ���� varchar(30),
    ��С�ܶ� varchar(30),
    ����ܶ� varchar(30),
    �հ��ܶ� varchar(30),
    �߿��ܶ� varchar(30),
    ���� varchar(30),
    ͼ��λ�� number(5),
    �û�AE���� VARCHAR2(50),
    ͼ��߿��� NUMBER(2),
    ͼƬ�ֱ��� number(3))
    TABLESPACE zl9BaseItem;

Create Table Ӱ��Ƭ��ӡ����(
    Ӱ����� varchar2(50),
    �����С number(5),
    �Ƿ���ͼ������ number(5),
    ��λ��ע�����С NUMBER(5),
    ��λ��ע��ͼ������ NUMBER(5),
    ���巴ɫ NUMBER(1) default 0,
    ������Ӱ NUMBER(1) default 0,
    ���屳��͸�� NUMBER(1) default 1)
    TABLESPACE zl9BaseItem;

Create Table ������Ӱ��(
       ҽ��ID         NUMBER(18),
       ��Ӱ��         Varchar2(30),
       ����           Varchar2(30),
       Ũ��           Varchar2(30))
    TABLESPACE zl9CisRec
    PCTFREE 5;

create table Ӱ��DICOM�����
(
  ����ID   number(18),
  ������    varchar2(20),
  �豸��   varchar2(3),
  ������ varchar2(20),
  PACS��ɫ varchar2(3),
  PACSIP��ַ    varchar2(15),
  PACSAE����   varchar2(20),
  PACS�˿� varchar2(5),
  �豸IP��ַ    varchar2(15),
  �豸AE����   varchar2(20),
  �豸�˿� varchar2(5)
  )
  TABLESPACE ZL9CISREC;

create table Ӱ��DICOM�������
(
  �������ID  number(18),
  ����ID   number(18),
  ��������  varchar2(100),
  ����ֵ    varchar2(100))
  TABLESPACE ZL9CISREC;

create table Ӱ��MWL�����
(
  ID         number(18),
  ����ID    number(18),
  ���       varchar2(4),
  Ԫ�غ�     varchar2(4),
  �ϼ�ID	 number(18),
  ���ı���   varchar2(50),
  Ӣ�ı���   varchar2(50),
  ����ֵ     varchar2(100),
  �Ƿ�Ƕ������ number(1),
  �Ƿ����     number(1),
  ֵ����	   varchar2(10),
  ѡ��         number(1),
  Ԫ������     varchar2(5),
  ǿ�ƽ��ֵ   varchar2(100),
  Ĭ��ֵ       varchar2(100),
  Ĭ��ѡ��     number(1),
  Ĭ��ǿ�ƽ��ֵ varchar2(100))
  TABLESPACE ZL9CISREC;

create table Ӱ������豸
(
  ����ID   number(18),
  IP��ַ   varchar2(20),
  �豸���� varchar2(100),
  Ӱ����� varchar2(20))
  TABLESPACE ZL9CISREC;
    
Create Table Ӱ���ղ����(
    ID   NUMBER(18),      
    �ϼ�ID   NUMBER(18),    
    �ղ����  Varchar2(64),   
    �Ƿ��� NUMBER(1),        
    ������   Varchar2(20),    
    ����ʱ�� Date             
)TABLESPACE zl9CISREC;


--�������
Create Table Ӱ�������
(
    ���� varchar2(10),
    ���� varchar2(20),
    ���� varchar2(10),
    ǰ����� varchar2(1),
    ������ number(18))
    TABLESPACE zl9BaseItem;

Create Table ��������(
    ���� VARCHAR2(2),
    ���� VARCHAR2(10),
    ���� VARCHAR2(6),
    ȱʡ��־ NUMBER(1) default 0)
    TABLESPACE zl9BaseItem;

Create Table ����������(
       ID Number(5),
       ���� Number(1) default -1,
       ǰ׺ Varchar2(5),
       ��   Number(1) default 0,
       ��   Number(1) default 0,
       ��   Number(1) default 0,
       ���λ�� Number(2) default 4,
       ���λ�� Number(2) default 5,
       ��ʼ��   Number(18) default 1,
       ��ͬ���� Number(1) default 0
)TABLESPACE zl9CisRec;

Create Table ��������¼(
       ID Number(5),
       ���� Number(1) default -1,
       ��   Number(4) default 0,
       ��   Number(2) default 0,
       ��   Number(2) default 0,
       ��ǰ��� Number(18) default 1
)TABLESPACE zl9CisRec;


Create Table ������걾(
       ID Number(18),
       �걾���� Varchar2(64),
       �걾��λ Varchar2(20),
       �걾���� Number(1) default 0,
       Ĭ�ϱ걾��   Varchar2(20),
       Ĭ����Ƭ�� Number(2) default 1,
       ����    varchar2(10),
       ��ע     Varchar2(255)       
) TABLESPACE zl9CisRec;

Create Table �����ײ���Ϣ(
    �ײ�ID Number(18), 
    �ײ����� VARCHAR2(64),
    �ײ���� VARCHAR2(64),
    �ײ�˵�� VARCHAR2(1024),
    ������ VARCHAR2(64),
    ����ʱ�� Date)
    TABLESPACE zl9CisRec;  
    
Create Table �����ײ͹���(
    ID Number(18),    
    �ײ�ID Number(18), 
    ����ID Number(18),
    ����˳�� Number(5))
    TABLESPACE zl9CisRec;  
    
    
Create Table ����������(
       ID Number(18),
       �������� Varchar2(64),
       �������� Number(1),
       �������� Varchar2(30),
       ������ Varchar2(64),
       ����ʱ�� date,
       ��ע Varchar2(1024)
  )TABLESPACE zl9CisRec;    



----------------------------------------------------------------------------
--[[12.ҽ��ҵ��]]
----------------------------------------------------------------------------
CREATE TABLE ����ǼǼ�¼(
	���� NUMBER(18),
	����ID NUMBER(18),
	��ҳID NUMBER(18),
	����ʱ�� DATE ,
	״̬ NUMBER(2),		--1-������;0-δ����
	ҽ����� VARCHAR2(3),
	�ʻ���� NUMBER(16,5),
	����ID NUMBER(18),
	�������� VARCHAR2(100),
	����֢ VARCHAR2(200),
	IC����Ϣ VARCHAR2(200),
	HIS��ˮ�� VARCHAR2(30),
	YB��ˮ�� VARCHAR2(30),
	��¼ID NUMBER(18),	--����ID��������ô��ֶ���������סԺ����
	��ע VARCHAR2(200),
	ȷ�� NUMBER(1))
    TABLESPACE ZL9BASEITEM;

CREATE TABLE ҽ�����˹�����(
	���� NUMBER (3),
	���� NUMBER (5),
	ҽ���� VARCHAR2 (30),
	����ID NUMBER (18),
	����ʱ�� DATE,
	��־ NUMBER (1) DEFAULT 0)
    TABLESPACE ZL9BASEITEM;

CREATE TABLE ������־(
	���� NUMBER(1) DEFAULT 0,	--1-����
	NO VARCHAR2(20),
	ҽ���� VARCHAR2 (50),
	���� VARCHAR2(100),
	�����ܶ� NUMBER(16,5),
	����ʱ�� DATE )
    TABLESPACE ZL9BASEITEM;

Create Table ҽ�����˵���(
    ���� NUMBER(3),
    ���� NUMBER(5),
    ���� VARCHAR2(25),
	ҽ���� VARCHAR2(30),
    ���� VARCHAR2(8),
    ��Ա��� VARCHAR2(8),
    ��λ���� VARCHAR2(15),
    ˳��� VARCHAR2(20),
	����֤�� VARCHAR2(26),
    �ʻ���� NUMBER(16,5),
    ��ǰ״̬ NUMBER(2),
    ����ID NUMBER(18),
    ��ְ NUMBER(1),
    ����� NUMBER(3),
    �Ҷȼ� VARCHAR2(1),
	����ʱ�� DATE)
    TABLESPACE zl9Patient
    PCTFREE 5;

Create Table �ʻ������Ϣ(
	����ID NUMBER(18),
	���� NUMBER(3),
	��� NUMBER(4),
	�ʻ������ۼ� NUMBER(16,5),
	�ʻ�֧���ۼ� NUMBER(16,5),
	����ͳ���ۼ� NUMBER(16,5),
	ͳ�ﱨ���ۼ� NUMBER(16,5),
	סԺ�����ۼ� NUMBER(3),
	��������   NUMBER(16,5),
	����ͳ���޶� NUMBER(16,5),
	���ͳ���޶� NUMBER(16,5),
	�����ۼ�   NUMBER(16,5),
	���ͳ���ۼ� NUMBER(16,5),
	������Ϣ  VARCHAR2(100))
    TABLESPACE zl9Patient
    PCTFREE 5;

Create Table ���ս����¼(
	���� NUMBER(2),
	��¼ID NUMBER(18),
	����ID NUMBER (18),
	����ʱ�� DATE ,
	���� NUMBER(3),
	����ID NUMBER(18),
	��� NUMBER(4),
	�ʻ��ۼ����� NUMBER(16,5),
	�ʻ��ۼ�֧�� NUMBER(16,5),
	�ۼƽ���ͳ�� NUMBER(16,5),
	�ۼ�ͳ�ﱨ�� NUMBER(16,5),
    סԺ���� NUMBER(5),
	���� NUMBER(16,5),
	�ⶥ�� NUMBER(16,5),
	ʵ������ NUMBER(16,5),
	�������ý�� NUMBER(16,5),
	ȫ�Ը���� NUMBER(16,5),
	�����Ը���� NUMBER(16,5),
	����ͳ���� NUMBER(16,5),
	ͳ�ﱨ����� NUMBER(16,5),
	���Ը���� NUMBER(16,5),
	�����Ը���� NUMBER(16,5),
	�����ʻ�֧�� NUMBER(16,5),
	֧��˳��� VARCHAR2(20),
	��;����   NUMBER(1),
	��ҳID   NUMBER(5),
	�Ƿ��ϴ� NUMBER(1),
	��ע     VARCHAR2(500),
	У�� NUMBER(1),
	������ˮ�� VARCHAR2(30),
	����ʱ�� DATE ,
	����վ VARCHAR2(50),
	�汾�� VARCHAR2(15),
	ҽ����� VARCHAR2(3),
	����ID NUMBER(18),
	�������� VARCHAR2(100),
	����֢ VARCHAR2(200),
	ȷ�� NUMBER(1),
  ��� Number(18))
    TABLESPACE zl9Expense
    PCTFREE 5;

Create Table ���ս������(
	����ID NUMBER(18),
	���� NUMBER(3),
	����ͳ���� NUMBER(16,5),
	ͳ�ﱨ����� NUMBER(16,5),
	���� NUMBER(3))
    TABLESPACE zl9Expense
    PCTFREE 5;

Create Table ���ս�����ϸ(
	����ID number(18),
	���㷽ʽ varchar2(20),
	��� number(16,5),
	��־ NUMBER(1) DEFAULT 0)
	TABLESPACE zl9Expense
	PCTFREE 5;

Create Table ����ģ�����(
    ����ID Number(18),
    ��ҳID Number(5),
    ���㷽ʽ Varchar2(20),
    ��� Number(16,5),
    ����ʱ�� Date)
    TABLESPACE zl9Expense
    PCTFREE 5;

----------------------------------------------------------------------------
--[[13.���˲���ҵ��]]
----------------------------------------------------------------------------
Create Table ������Ϣ(
    ����ID NUMBER(18),
    ��ҳID NUMBER(5),
    ����� NUMBER(18),
    סԺ�� NUMBER(18),
    ���￨�� VARCHAR2(50),
    ����֤�� VARCHAR2(50),
    �ѱ� VARCHAR2(10),
    ҽ�Ƹ��ʽ Varchar2(20),
    ���� VARCHAR2(100),
    �Ա� VARCHAR2(4),
    ���� varchar2(20),
    �������� Date,
    �����ص� VARCHAR2(100),
    ���֤�� VARCHAR2(18),
    ����֤�� VARCHAR2(20),
    ��� VARCHAR2(10),
    ְҵ VARCHAR2(80),
    ���� VARCHAR2(20),
    ���� VARCHAR2(30),
    ���� VARCHAR2(100),
    ���� VARCHAR2(30),
    ѧ�� VARCHAR2(10),
    ����״�� VARCHAR2(4),
    ��ͥ��ַ VARCHAR2(100),
    ��ͥ�绰 VARCHAR2(20),
    ��ͥ��ַ�ʱ� VARCHAR2(6),
    �໤�� VARCHAR2(64),
    ��ϵ������ VARCHAR2(64),
    ��ϵ�˹�ϵ VARCHAR2(30),
    ��ϵ�˵�ַ VARCHAR2(100),
    ��ϵ�˵绰 VARCHAR2(20),
    ���ڵ�ַ VARCHAR2(100),
    ���ڵ�ַ�ʱ� VARCHAR2(6),
    Email Varchar2(30),
    QQ Varchar2(12),
    ��ͬ��λid NUMBER(18),
    ������λ VARCHAR2(100),
    ��λ�绰 VARCHAR2(20),
    ��λ�ʱ� VARCHAR2(6),
    ��λ������ VARCHAR2(50),
    ��λ�ʺ� VARCHAR2(20),
    ������ VARCHAR2(100),
    ������ NUMBER(16,5),
    �������� NUMBER(1),
    ����ʱ�� Date,
    ����״̬ Number(1) Default 0,
    �������� Varchar2(20),
    סԺ���� number(3),
    ��ǰ����id number(18),
    ��ǰ����id number(18),
    ��ǰ���� VARCHAR2(10),
    ��Ժʱ�� DATE,
    ��Ժʱ�� Date,
    ��Ժ number(1),
    IC���� varchar2(50),
    ������ varchar2(50),
    ҽ���� VARCHAR2(30),
    ���� NUMBER(3),
    ��ѯ���� Varchar2(50),
    �Ǽ�ʱ�� Date,
    ͣ��ʱ�� Date,
    ���� Number(1),
    ��ϵ�����֤�� varchar(18),
    �������� Varchar2(50),
    ����ģʽ number(2))
    TABLESPACE zl9Patient initrans 20
;

Create Table ��Ժ����
(
	����ID NUMBER(18),
	����ID NUMBER(18),
	����ID NUMBER(18)
)
TABLESPACE zl9Patient
Initrans 20;

CREATE TABLE ������Ϣ�䶯(
	����ID Number(18),
	�䶯��Ŀ VARCHAR2(10) not NULL,
	ԭ��Ϣ VARCHAR2(100),
	����Ϣ VARCHAR2(100),
	�䶯ʱ�� DATE,
	�䶯�� Varchar2(20),
	�䶯ģ�� Varchar2(100),
	˵�� varchar2(4000)
	)TABLESPACE zl9Patient;

CREATE TABLE ����ҽ�ƿ���Ϣ(
	����ID number(18),
	�����ID Number(18),
	���� Varchar2(50),
	���� Varchar2(50),
	״̬ Number(2) DEFAULT 0,
	��ʧʱ�� Date,
	��ʧ��ʽ Varchar2(20),
	��ʧ��   varchar2(20),
	�������� Date,
	������ Varchar2(20))
	TABLESPACE zl9Patient;


CREATE TABLE ���˷�����¼ (
	����ID NUMBER (18),
	�ɿ��� VARCHAR2 (50),
	�ɿ����� NUMBER (3) DEFAULT 2,
	�ɿ�����ҽԺ VARCHAR2 (50),
	�¿��� VARCHAR2 (50),
	����ʱ�� DATE ,
	�ϴ���־ NUMBER (1) DEFAULT 0)
TABLESPACE zl9BaseItem
PCTFREE 5;

Create Table ����ҽ�ƿ�����(
	����ID Number(18),
	�����ID Number(18),
	���� varchar2(50),
	��Ϣ�� Varchar2(20),
	��Ϣֵ Varchar2(100))
	Tablespace zl9Patient ;

Create Table ������Ϣ�ӱ�(
	����ID Number(18),
	����ID Number(18),
	��Ϣ�� Varchar2(20),
	��Ϣֵ Varchar2(100))
	Tablespace zl9Patient ;

CREATE TABLE ����ҽ�ƿ��䶯(
	ID 	Number(18),
	����ID 	Number(18),
	�����ID Number(18),
	���� 	 VarChar2(50),
	�䶯ID 	 Number(18),
	�䶯��� Number(3),
	ԭ����   VARCHAR2(50),
	������   VARCHAR2(50),
	�䶯ʱ�� Date,
	�䶯ԭ�� Varchar2(100),
	��ʧ��ʽ Varchar2(30),
	����Ա���� Varchar2(20),
	�Ǽ�ʱ�� Date)
	TABLESPACE zl9Patient;

Create Table ������Ƭ(
    ����ID NUMBER(18),
    ��Ƭ Long Raw)
    TABLESPACE zl9Patient
    PCTFREE 20;
    --��Ƭ����ʱ��Ҫʹ�ý϶��Ԥ���ռ�

Create Table ���ⲡ��(
	��� NUMBER(18),
  ����ID NUMBER(18),
	����ԭ�� VARCHAR2(200),
	����ʱ�� DATE,
	�Ǽ��� VARCHAR2(20),
	����ԭ�� VARCHAR2(200),
	����ʱ�� DATE,
	������ VARCHAR2(20))
    TABLESPACE zl9Patient;

Create Table ���˵�����¼(
    ����ID      NUMBER(18),
	��ҳID		NUMBER(5),
    ������      VARCHAR2(64),
    ������      NUMBER(16,5),
	��������    NUMBER(1),
	����ԭ��	   VARCHAR2(50),
	�ۼƺ�      NUMBER(5),
    ����Ա���  VARCHAR2(6),
    ����Ա����  VARCHAR2(20),
    �Ǽ�ʱ��    Date,
	����ʱ��	Date,
	ɾ����־  NUMBER(1) default 1,
	ɾ������Ա��� VARCHAR2(6),
	ɾ������Ա���� VARCHAR2(20),
	ɾ��ʱ��  Date)
    TABLESPACE zl9Patient;

Create Table ���˺ϲ���¼(
    ����ID				NUMBER(18),
    ԭ��Ϣ			VARCHAR2(1000),
	�ϲ�ԭ��		VARCHAR2(250),
    ����Ա����	VARCHAR2(20),
    �ϲ�ʱ��		Date)
    TABLESPACE zl9Patient;

CREATE TABLE ����������Ϣ(
		����ID NUMBER(18),
		���� NUMBER(5),
		������ VARCHAR2(20),
		��־ NUMBER(1),
		�������� NUMBER(1),
		����ʱ�� DATE)
		TABLESPACE zl9Patient;

Create Table ���ﲡ����¼(
    ����ID NUMBER(18),
    ������ NUMBER(18),
    �������� Date,
    ������� VARCHAR2(10),
    �洢״̬ VARCHAR2(4),
    ���λ�� VARCHAR2(20))
    TABLESPACE zl9Patient
    initrans 20;

Create Table סԺ������¼(
    ����ID NUMBER(18),
    ��ҳID NUMBER(5),
    ������ VARCHAR2(20),
		������ VARCHAR2(20),
    �������� Date,
    ������� VARCHAR2(10),
    �洢״̬ VARCHAR2(8),
    ���λ�� VARCHAR2(20))
    TABLESPACE zl9Patient
    PCTFREE 5;

Create Table ������ҳ(
    ����ID NUMBER(18),
    ��ҳID NUMBER(5),
    סԺ�� NUMBER(18),
    �������� NUMBER(1),
    ҽ�Ƹ��ʽ VARCHAR2(20),
    �ѱ� VARCHAR2(10),
    ����Ժ NUMBER(1),
    ��Ժ����ID NUMBER(18),
    ��Ժ����id NUMBER(18),
    ҽ��С��id NUMBER(18),
    ��Ժ���� Date,
    ��Ժ���� VARCHAR2(20),
    ��Ժ��ʽ VARCHAR2(8),
    ��Ժ���� VARCHAR2(8),
    ����Ժת�� VARCHAR2(1),
    סԺĿ�� VARCHAR2(10),
    ��Ժ���� VARCHAR2(10),
    �Ƿ���� NUMBER(1),
    ��ǰ���� VARCHAR2(20),
    ��ǰ����id NUMBER(18),
    ����ȼ�id NUMBER(18),
    ��Ժ����id NUMBER(18),
    ��Ժ���� VARCHAR2(10),
    ��Ժ���� Date,
    סԺ���� NUMBER(4),
    ��Ժ��ʽ VARCHAR2(10),
    �Ƿ�ȷ�� NUMBER(1),
    ȷ������ Date,
    �·����� number(1),
    Ѫ�� VARCHAR2(10),
    ���ȴ��� NUMBER(5),
    �ɹ����� NUMBER(5),
    �����־ NUMBER(1),
    �������� NUMBER(5),
    ʬ���־ NUMBER(1),
    ����ҽʦ VARCHAR2(20),
    ���λ�ʿ VARCHAR2(20),
    סԺҽʦ VARCHAR2(20),
    ������ VARCHAR2(20),
    ��ĿԱ��� VARCHAR2(6),
    ��ĿԱ���� VARCHAR2(20),
    ��Ŀ���� Date,
    ״̬ NUMBER(3),
    ���ú� NUMBER(16,5),
    ���� VARCHAR2(100),
    �Ա� VARCHAR2(4),
    ���� varchar2(20),
    ��� NUMBER(16,5),
    ���� NUMBER(16,5),
    ����״�� VARCHAR2(4),
    ְҵ VARCHAR2(80),
    ���� VARCHAR2(30),
    ѧ�� VARCHAR2(10),
    ��λ�绰 VARCHAR2(20),
    ��λ�ʱ� VARCHAR2(6),
    ��λ��ַ VARCHAR2(100),
    ���� VARCHAR2(30),
    ��ͥ��ַ VARCHAR2(100),
    ��ͥ�绰 VARCHAR2(20),
    ��ͥ��ַ�ʱ� VARCHAR2(6),
    ��ϵ������ VARCHAR2(64),
    ��ϵ�˹�ϵ VARCHAR2(30),
    ��ϵ�˵�ַ VARCHAR2(100),
    ��ϵ�˵绰 VARCHAR2(20),
    ��ϵ�����֤�� VARCHAR2(18),
    ���ڵ�ַ VARCHAR2(100),
    ���ڵ�ַ�ʱ� VARCHAR2(6),
    ��ҽ������� VARCHAR2(4),
    ���� NUMBER(3),
    ���� Number(5),
    ��˱�־ NUMBER(1),
    ����� VARCHAR2(20),
    ������� DATE,
    �Ƿ��ϴ� NUMBER(1),
    ����ת�� Number(1),
    �Ǽ��� Varchar2(20),
    �Ǽ�ʱ�� Date,
    ��ע Varchar2(100),
    ����״̬ Number(3),
    �������� Varchar2(50),
    ���ʱ�� Date,
    ·��״̬ number(1),
    ������ varchar2(2),
    ��ת�� Number(3),
    Ӥ������ID NUMber(18),
    Ӥ������ID NUMber(18),
    ĸӤת�Ʊ�־ varchar2(100),
    ҽ������ʱ�� Date)
    TABLESPACE zl9Patient initrans 20
;

Create Table ������ҳ�ӱ�(
	����ID NUMBER(18),
	��ҳID NUMBER(5),
	��Ϣ�� VARCHAR2(20),
	��Ϣֵ VARCHAR2(100))
    TABLESPACE zl9Patient
    initrans 20;

Create Table ���˱䶯��¼(
    Id Number(18),
    ����ID number(18) Not Null,
    ��ҳID number(5) Not Null,
    ��ʼʱ�� Date,
    ��ʼԭ�� number(3),
    ���Ӵ�λ number(1),
    ����id number(18),
    ����id number(18),
    ҽ��С��id number(18),
    ����ȼ�id number(18),
    ��λ�ȼ�id number(18),
    ���� VARCHAR2(10),
    ���λ�ʿ varchar2(20),
    ����ҽʦ varchar2(20),
    ����ҽʦ varchar2(20),
    ����ҽʦ varchar2(20),
    ����         varchar2(20),
    ��ֹ��Ա varchar2(20),
    ��ֹʱ�� Date,
    ��ֹԭ�� number(3),
    ����Ա��� varchar2(6),
    ����Ա���� varchar2(20),
    �ϴμ���ʱ�� Date)
    TABLESPACE zl9Patient
    PCTFREE 5 initrans 20;

Create Table ���˹���ҩ��(
    ����ID NUMBER(18),
    ����ҩ��id NUMBER(18),
    ����ҩ�� VARCHAR2(60),
	������Ӧ varchar2(100))
    TABLESPACE zl9Patient
    PCTFREE 5;

Create Table ��λ״����¼(
    ����id NUMBER(18),
    ���� VARCHAR2(10),
    ����id NUMBER(18),
    ����� VARCHAR2(10),
    �Ա���� VARCHAR2(10),
    ��λ���� VARCHAR2(10),
    �ȼ�id NUMBER(18),
    ״̬ VARCHAR2(4),
    ����id NUMBER(18),
	���� NUMBER(1) Default 0)
    TABLESPACE zl9Patient
    PCTFREE 20 initrans 20
    Cache Storage(Buffer_Pool Keep);

Create Table ��λ������¼(
    ���� Date,
    �䶯 NUMBER(5),
    ����id NUMBER(18),
    ���� VARCHAR2(10),
    ����id NUMBER(18),
	��λ���� VARCHAR2(10))
    TABLESPACE zl9Patient;

Create Table ����������Ŀ(
    ����ID      NUMBER(18),
    ��ҳID	NUMBER(5),
    ��ĿID      NUMBER(18),
    ������      VARCHAR2(20),
    ����ʱ��	Date,
    ʹ������	NUMBER(16,5),
    ��������	NUMBER(16,5))
    TABLESPACE zl9Patient;

CREATE TABLE ������Ŀģ��(
	����  number(5),
	����  varchar2(20),
	��ĿID NUMBER(18))
	TABLESPACE zl9Patient;

Create Table ���˱�ע��Ϣ(
    Id Number(18),
    ����ID number(18) Not Null,
    ��ҳID number(5) Not Null,
    ���� varchar2(200),
    �Ǽ�ʱ�� Date,
    �Ǽ��� varchar2(20),
    �Ƿ���� Number(1),
    ���ʱ�� Date,
    ����� varchar2(20))
    TABLESPACE zl9Patient;

Create Table ���˵�ַ��Ϣ(
    ����ID NUMBER(18),
    ��ҳID NUMBER(5),
    ��ַ��� Number(5),
    ʡ varchar2(100),
    �� varchar2(100),
    �� varchar2(100),
    ���� varchar2(100)) 
    tablespace zl9Patient;

--���˲���
----------------------------------------------------------------------------
CREATE TABLE ���˹�����¼(
    ID NUMBER(18),
    ����ID NUMBER(18),
    ��ҳID NUMBER(18),
    ��¼��Դ NUMBER(1),
    ҩ��ID NUMBER(18),
    ҩ���� VARCHAR2(60),
    ��� NUMBER(1),
    ����ʱ�� DATE,
    ��¼ʱ�� DATE,
    ��¼�� VARCHAR2(20),
    ������Ӧ varchar2(100),
    ����Դ���� Varchar2(10),
    ��ת�� Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;

CREATE TABLE ����֢״��¼(
    ����ID NUMBER(18),  
    ��ҳID NUMBER(18),	--���ﲡ����Һ�ID
    ���   NUMBER(4),
    ����   VARCHAR2(10),
    ����   VARCHAR2(100),
    ��ʼ���� DATE,
    �������� DATE,
    ��¼�� VARCHAR2(20),
    ��¼ʱ�� DATE)
    TABLESPACE zl9CisRec;

CREATE TABLE  �������߼�¼ (
	����ID NUMBER(18),
	����ʱ�� Date,
	�������� varchar2(200)) 
	TABLESPACE zl9Patient;	

CREATE TABLE ������ϼ�¼(
    ID NUMBER(18),
    ����ID NUMBER(18),
    ��ҳID NUMBER(18),
    ҽ��ID NUMBER(18),
    ��¼��Դ NUMBER(1),
    ��ϴ��� NUMBER(2) DEFAULT 1,
    ������� NUMBER(2) DEFAULT 1,
    ����ID NUMBER(18),
    ����ID NUMBER(18),
    ������� NUMBER(2),
    ����ID NUMBER(18),
    ���ID NUMBER(18),
    ֤��ID NUMBER(18),
    ������� VARCHAR2(200),
    ��Ժ���� varchar2(30),
    ��Ժ��� VARCHAR2(10),
    �Ƿ�δ�� NUMBER(1),
    �Ƿ����� NUMBER(1),
    ��ע VARCHAR2(50),
    ��¼���� DATE,
    ��¼�� VARCHAR2(20),
    ȡ��ʱ�� DATE,
    ȡ���� VARCHAR2(20),
	����ʱ�� date,
	��ת�� Number(3))
    TABLESPACE zl9CisRec;

CREATE TABLE �������ҽ��(
    ���ID NUMBER(18),
    ҽ��ID NUMBER(18),
	��ת�� Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;

Create Table ��Ϸ������(
	����ID number(18),
	��ҳID number(5),
	�������� number(2),
	������� number(2))
	TABLESPACE zl9CisRec
	PCTFREE 5;

CREATE TABLE ���������¼(
    ID NUMBER(18),
    ����ID NUMBER(18),
    ��ҳID NUMBER(5),
    ��¼��Դ NUMBER(1),
    �������� DATE,
    ׼������ Number(3),
    ������� VARCHAR2(8),
    �ٴ����� NUMBER(1),
    ������ʼʱ�� DATE,
    ��������ʱ�� DATE,
    ������ҩʱ�� Date,
    �������� VARCHAR2(100),
    ��������ID NUMBER(18),
    ������ĿID NUMBER(18),
    �������� VARCHAR2(100),
    ����ҽʦ VARCHAR2(20),
    ������ʿ VARCHAR2(20),
    ��һ���� VARCHAR2(20),
    �ڶ����� VARCHAR2(20),
    ������ʿ VARCHAR2(20),
    ����ʼʱ�� DATE,
    �������ʱ�� DATE,
    ����ʽ NUMBER(18),
    ASA�ּ� VARCHAR2(20),
    NNIS�ּ� VARCHAR2(20),
    �������� number(2),
    �������� VARCHAR2(20),
    �������� VARCHAR2(6),
    ��Һ���� NUMBER(5),
    ����ҽʦ VARCHAR2(20),
    ������ʼʱ�� DATE,
    ��������ʱ�� DATE,
    �п� VARCHAR2(2),
    ���� VARCHAR2(6),
    �пڲ�λ VARCHAR2(100),
    �ط��ƻ� NUMBER(1),
    �ط�Ŀ�� VARCHAR2(100),
    �пڸ�Ⱦ NUMBER(1),
    ����֢ NUMBER(1),
	��ǰ������ҩ   NUMBER(1),
    ������ҩ����   NUMBER(5),
    ��Ԥ�ڵĶ������� NUMBER(1),
    ������֢    NUMBER(1),
    ������������   NUMBER(1),
    ��������֢    NUMBER(1),
    �����Ѫ��Ѫ��  NUMBER(1),
    �����˿��ѿ�   NUMBER(1),
    �������Ѫ˨  NUMBER(1),
    ���������л���� NUMBER(1),
    �������˥��   NUMBER(1),
    �����˨��    NUMBER(1),
    �����Ѫ֢    NUMBER(1),
    �����Źؽڹ���  NUMBER(1),
    ��¼���� DATE,
    ��¼�� VARCHAR2(20),
    ȡ��ʱ�� DATE,
    ȡ���� VARCHAR2(20),
	��ת�� Number(3),
	������Դ number(1))
    TABLESPACE zl9CisRec;

CREATE TABLE ������������¼(
	����ID NUMBER(18),
	��ҳID NUMBER(18),
	��� NUMBER(3),
	Ӥ������ VARCHAR2(100),
	Ӥ���Ա� VARCHAR2(4),
	������� NUMBER(3),
	���䷽ʽ VARCHAR2(20),
	̥��״�� VARCHAR2(20),
	����ʱ�� DATE)
    TABLESPACE zl9CisRec
    PCTFREE 5;

Create Table ���˿����ؼ�¼(
       ����Id NUMBER(18),
       ��ҳId NUMBER(5),
       ҩ��id NUMBER(18),
       ҩƷ���� VARCHAR2(80),
       ��ҩĿ�� VARCHAR2(200),
       ʹ�ý׶� VARCHAR2(30),
       ʹ������ NUMBER(18,2),
       ��¼�� VarCHAR2(30),
       ��¼ʱ�� Date,
       һ���п�Ԥ���� Number(1),
       DDD�� Number(16,4),
       ������ҩ varchar2(30))
  TABLESPACE zl9CisRec;


Create Table ������֢�໤���(
	����ID number(18),
	��ҳID number(5),
	���   number(18), 
	�໤������ varchar2(100),
	����ʱ�� Date,
	�˳�ʱ�� Date,
	����ס�ƻ� number(1),
	����סԭ�� varchar2(100),
	�˹������ѳ�   NUMBER(1),
    �ط���֢ҽѧ��  NUMBER(1),
    �ط����ʱ��   VARCHAR2(30)
)TABLESPACE zl9CisRec;

Create Table �������Ƽ�¼(
  ����ID NUMBER(18),
  ��ҳID NUMBER(5),
  ���   number(18),
  ����ID NUMBER(18),
  ��ʼ���� DATE,
  �������� DATE,
  �Ƴ���   number(16,5),
  ����     number(16,5),
  ���Ʒ��� VARCHAR2(50),
  ����Ч�� VARCHAR2(10))
    TABLESPACE zl9CisRec ;
    
Create Table �������Ƽ�¼(
  ����ID NUMBER(18),
  ��ҳID NUMBER(5),
  ���   number(18),
  ����ID NUMBER(18),
  ��ʼ���� DATE,
  �������� DATE,
  ��Ұ��λ VARCHAR2(50),
  �������   NUMBER(16,5),
  �ۼ���     NUMBER(16,5),
  ����Ч�� VARCHAR2(10))
    TABLESPACE zl9CisRec ;

Create Table ������������(
	����ID NUMBER(18),
	��ҳID NUMBER(5),
	���   number(18),
	ҩƷID number(18),
	ҩ������ varchar2(200),
	�Ƴ�	varchar2(50),
	������� varchar2(50),
	���ⷴӦ VARCHAR2(100),
	��Ч VARCHAR2(50))
    TABLESPACE zl9CisRec ;

Create Table ��е����ʹ�����(
	����ID number(18),
	��ҳID number(5),
	��� number(18),
	�໤������ VARCHAR2(50),
	��е������ Varchar2(20),
	��ʼʹ��ʱ�� Date,
	����ʹ��ʱ�� Date,
	��Ⱦ�ۼ�ʱ�� varchar(20))
TABLESPACE zl9CisRec;

Create Table ���˸�Ⱦ��¼(
	��� number(5),
	����ID NUMBER(18),
	��ҳID NUMBER(5),
	�Ǽ�ʱ�� Date,
	�Ǽ��� VARCHAR2(20),
	ȷ������ Date,
	��Ⱦ��λ VARCHAR2(20),
	��Ⱦ���� VARCHAR2(30)
)TABLESPACE zl9CisRec PCTFREE 5;

Create Table ���˲�ԭѧ���(
	��� number(5),
	����ID NUMBER(18),
	��ҳID NUMBER(5),
	�Ǽ�ʱ�� Date,
	�Ǽ��� VARCHAR2(20),
	�걾 VARCHAR2(20),
	��ԭѧ���� VARCHAR2(20),
	�ͼ����� Date
)TABLESPACE zl9CisRec PCTFREE 5;


----------------------------------------------------------------------------
--[[14.����ҵ��]]
----------------------------------------------------------------------------
Create Table ƾ����ӡ��¼(
  NO varchar2(8),
  ��¼���� NUMBER(3),
  ��ӡʱ�� Date,
  ��ӡ���� NUMBER(3),
  ��ӡ�� varchar2(100),
  ������ varchar2(100),
  IP��ַ varchar2(100),
  ��ע varchar2(500),
  ��ת�� NUMBER(3)) 
TABLESPACE zl9Expense PCTFREE 5 initrans 20;
Create Table ���˹Һż�¼(
    ID NUMBER(18),
    NO VARCHAR2(8),
    ��¼���� number(3) default(1),
    ��¼״̬ NUMBER(3)default(1),
    ����ID NUMBER(18),
    ����� NUMBER(18),
    ���� VARCHAR2(100),
    �Ա� VARCHAR2(4),
    ���� varchar2(20),
    ���� NUMBER(1),
    �ű� VARCHAR2(5),
    ���� NUMBER(5),
    ���� NUMBER(1),
    ���� VARCHAR2(20),
    ���ӱ�־ NUMBER(1),
    ִ�в���ID NUMBER(18),
    ִ���� VARCHAR2(20),
    ִ��״̬ NUMBER(1),
    ִ��ʱ�� DATE,
    ���ʱ�� DATE,
    �Ǽ�ʱ�� DATE,
    ����ʱ�� DATE,
    ����Ա��� VARCHAR2(6),
    ����Ա���� VARCHAR2(20),
    ��Ⱦ���ϴ� Number(1),
    ����ʱ�� Date,
    ������ַ varchar2(200),
    ����ʱ�� Date,
    ���� Number(5),
    ժҪ Varchar2(1000),
    ת��ű� VARCHAR2(5),
    ת�����ID Number(18),
    ת������ VARCHAR2(20),
    ת��ҽ�� VARCHAR2(20),
    ת��״̬ Number(1),
    �������ID Number(18),
    ��������� VARCHAR2(250),
    ԤԼ number(2),
    ԤԼ��ʽ varchar2(10),
    ��¼��־ number(2),
    �˺������ VARCHAR2(20),
    �˺����ʱ�� DATE,
    ԤԼʱ�� DATE,
    ������ VARCHAR2(20),
    ����ʱ�� Date,
    ������ˮ�� VARCHAR2(50),
    ����˵�� VARCHAR2(50),
    ������λ VARCHAR2(50),
    ԤԼ����Ա VARCHAR2(20),
    ԤԼ����Ա��� VARCHAR2(6),
    ���� number(3),
    ��ת�� Number(3))
    TABLESPACE zl9Patient PCTFREE 5 initrans 20
;

Create Table ����ת���¼(
    �Һ�ID NUMBER(18),
    NO VARCHAR2(8),
    �������ID NUMBER(18),
    ����ҽ�� VARCHAR2(20),
    ���տ���ID NUMBER(18),
    ����ҽ�� VARCHAR2(20),
	����ʱ�� Date,
	��ת�� Number(3)
	)
    TABLESPACE zl9Patient;

Create Table ���˹ҺŻ���(
    ���� date,
    ����id NUMBER(18),
    ��ĿID NUMBER(18),
    ҽ������ VARCHAR2(20),
    ҽ��ID NUMBER(18),
    ���� VARCHAR2(5),
    �ѹ��� NUMBER(5),
    ��Լ�� NUMBER(5),
    �����ѽ��� Number(5),
	��ת�� Number(3))
    TABLESPACE zl9Expense
    PCTFREE 5 initrans 20;

Create  Table ������λ�ҺŻ���(
	���� Date,
	���� Varchar2(5),
	������λ Varchar2(50),
	��� Number(5),
	��Լ�� Number(10),
	�ѽ��� Number(10)
	) Tablespace zl9Expense
	PCTFREE 5 initrans 20;

Create Table �������(
    ����id NUMBER(18),
    ���� NUMBER(1),
    ���� NUMBER(2) DEFAULT 2,
    Ԥ����� NUMBER(16,5),
    ������� NUMBER(16,5))
    TABLESPACE zl9Expense
    initrans 20;

Create Table ���˽ɿ��¼(
    ID NUMBER(18),
    ����ID NUMBER(18),
    No VARCHAR2(8),
    ��¼״̬ Number(3),
    ���㷽ʽ VARCHAR2(20),
    ����� VARCHAR2(10),
    ��� NUMBER(16,5),
    ժҪ VARCHAR2(50),
    �Ǽ�ʱ�� Date,
    �Ǽ��� VARCHAR2(20))
    TABLESPACE zl9Expense
    PCTFREE 5;

Create Table ���˽ɿ����(
    �ɿ VARCHAR2(8),
    ����ID NUMBER(18),
    ��� NUMBER(16,5))
    TABLESPACE zl9Expense
    PCTFREE 5;

Create Table ���˴߿��¼(
    ID		NUMBER(18),
    ����ID  NUMBER(18),
    ��ҳID  NUMBER(18),
    Ԥ�����  NUMBER(16,5),
    δ�����  NUMBER(16,5),
    �Էѽ��  NUMBER(16,5),
    ҽ��Ԥ��  NUMBER(16,5),
    ��ǰ���  NUMBER(16,5),
    �߿�����  NUMBER(16,5),
    �߿��׼  NUMBER(16,5),
    �߿���  NUMBER(16,5), 
    ��ӡ���� DATE ,
    ��ӡ��     VARCHAR2(20))
    TABLESPACE zl9Expense;

Create Table ���˽��ʼ�¼(
    ID NUMBER(18),
    NO VARCHAR2(8),
    ʵ��Ʊ�� VARCHAR2(20),
    ��¼״̬ NUMBER(3),
    ��;���� NUMBER(1),
    ����id NUMBER(18),
    ����Ա��� VARCHAR2(6),
    ����Ա���� VARCHAR2(20),
    ��ע   VARCHAR2(50),
    ԭ��  VARCHAR2(100),
    �շ�ʱ�� Date,
    ��ʼ���� Date,
    �������� Date,
    �ɿ���ID number(18),
    �������� NUMBER(1),
	��ת�� Number(3))
    TABLESPACE zl9Expense
    PCTFREE 5;

Create Table סԺ���ü�¼(
    ID NUMBER(18),
    ��¼���� NUMBER(3),
    NO VARCHAR2(8),
    ʵ��Ʊ�� VARCHAR2(50),
    ��¼״̬ NUMBER(3),
    ��� NUMBER(18),
    �������� NUMBER(5),
    �۸񸸺� NUMBER(5),
    �ಡ�˵� NUMBER(1) default 0,
    ���ʵ�ID NUMBER(18) default 0,
    ����id NUMBER(18),
    ��ҳid NUMBER(5),
    ҽ����� NUMBER(18),
    �����־ NUMBER(3) default 1,
    ���ʷ��� NUMBER(1) default 0,
    ���� VARCHAR2(100),
    �Ա� VARCHAR2(4),
    ���� varchar2(20),
    ��ʶ�� NUMBER(18),
    ���� VARCHAR2(10),
    ���˲���id NUMBER(18),
    ���˿���id NUMBER(18),
    �ѱ� VARCHAR2(10),
    �շ���� VARCHAR2(1),
    �շ�ϸĿid NUMBER(18),
    ���㵥λ VARCHAR2(20),
    ���� NUMBER(3) default 1,
    ��ҩ���� VARCHAR2(50),
    ���� NUMBER(16,5),
    �Ӱ��־ NUMBER(1),
    ���ӱ�־ NUMBER(1),
    Ӥ���� NUMBER(1),
    ������Ŀid NUMBER(18),
    �վݷ�Ŀ VARCHAR2(20),
    ��׼���� NUMBER(16,5),
    Ӧ�ս�� NUMBER(16,5),
    ʵ�ս�� NUMBER(16,5),
    ������ VARCHAR2(20),
    ��������id NUMBER(18),
    ������ VARCHAR2(20),
    ����ʱ�� Date,
    �Ǽ�ʱ�� Date,
    ִ�в���id NUMBER(18),
    ִ���� VARCHAR2(20),
    ִ��״̬ NUMBER(2),
    ִ��ʱ�� date,
    ���� Varchar2(500),
    ����Ա��� VARCHAR2(6),
    ����Ա���� VARCHAR2(20),
    ����id NUMBER(18),
    ���ʽ�� NUMBER(16,5),
    ���մ���id number(18),
    ������Ŀ�� number(1),
    ���ձ��� varchar2(20),
    �������� varchar2(20),
    ͳ���� number(16,5),
    �Ƿ��ϴ� number(1),
    ժҪ Varchar2(1000),
    �Ƿ��� Number(1) Default 0,
    �ɿ���ID number(18),
    ҽ��С��ID NUMBER(18),
    ��ת�� Number(3))
    TABLESPACE zl9Expense initrans 20
;

Create Table ������ü�¼(
    ID NUMBER(18),
    ��¼���� NUMBER(3),
    NO VARCHAR2(8),
    ʵ��Ʊ�� VARCHAR2(50),
    ��¼״̬ NUMBER(3),
    ��� NUMBER(18),
    �������� NUMBER(5),
    �۸񸸺� NUMBER(5),
    ���ʵ�ID NUMBER(18) default 0,
    ����id NUMBER(18),
    ҽ����� NUMBER(18),
    �����־ NUMBER(3) default 1,
    ���ʷ��� NUMBER(1) default 0,
    ���� VARCHAR2(100),
    �Ա� VARCHAR2(4),
    ���� varchar2(20),
    ��ʶ�� NUMBER(18),
    ���ʽ VARCHAR2(10),
    ���˿���id NUMBER(18),
    �ѱ� VARCHAR2(10),
    �շ���� VARCHAR2(1),
    �շ�ϸĿid NUMBER(18),
    ���㵥λ VARCHAR2(20),
    ���� NUMBER(3) default 1,
    ��ҩ���� VARCHAR2(50),
    ���� NUMBER(16,5),
    �Ӱ��־ NUMBER(1),
    ���ӱ�־ NUMBER(1),
    Ӥ���� NUMBER(1),
    ������Ŀid NUMBER(18),
    �վݷ�Ŀ VARCHAR2(20),
    ��׼���� NUMBER(16,5),
    Ӧ�ս�� NUMBER(16,5),
    ʵ�ս�� NUMBER(16,5),
    ������ VARCHAR2(20),
    ��������id NUMBER(18),
    ������ VARCHAR2(20),
    ����ʱ�� Date,
    �Ǽ�ʱ�� Date,
    ִ�в���id NUMBER(18),
    ִ���� VARCHAR2(20),
    ִ��״̬ NUMBER(2),
    ִ��ʱ�� date,
    ���� Varchar2(500),
    ����Ա��� VARCHAR2(6),
    ����Ա���� VARCHAR2(20),
    ����id NUMBER(18),
    ���ʽ�� NUMBER(16,5),
    ���մ���id number(18),
    ������Ŀ�� number(1),
    ���ձ��� varchar2(20),
    �������� varchar2(20),
    ͳ���� number(16,5),
    �Ƿ��ϴ� number(1),
    ժҪ Varchar2(1000),
    �Ƿ��� Number(1) Default 0,
    �ɿ���ID number(18),
    ����״̬ number(4),
    ��ת�� Number(3))
    TABLESPACE zl9Expense initrans 20
;

Create Table ���˷�������(
    ����ID NUMBER(18),
    ������� number(2) DEFAULT 0,
    �շ�ϸĿid NUMBER(18),
    ���벿��id NUMBER(18),
    ��˲���id NUMBER(18),
    ���� NUMBER(16,5),
    ������ VARCHAR2(20),
    ����ʱ�� Date,
    ����� VARCHAR2(20),
    ���ʱ�� Date,
    ״̬ NUMBER(1))
    TABLESPACE zl9Expense
    PCTFREE 5;

Create Table �����˷�����(
    NO VARCHAR2(8),
    ��¼���� NUMBER(3),
    ������ VARCHAR2(20),
    ����ʱ�� Date,
    ����� VARCHAR2(20),
    ���ʱ�� Date,
    ����ԭ�� Varchar2(100))
    TABLESPACE zl9Expense PCTFREE 5
;

Create Table ���˽��ʻ���(    
	����ʱ�� date,
	����ID  NUMBER(18),
	��ҳID NUMBER(5),
	����id NUMBER(18),
	���˲���id NUMBER(18),
	���˿���id NUMBER(18),
	��������id NUMBER(18),
	ִ�в���id NUMBER(18),
	������Ŀid NUMBER(18),    
	Ӧ�ս�� NUMBER(16,5),
	ʵ�ս�� NUMBER(16,5),
	���ʽ�� NUMBER(16,5))
	TABLESPACE zl9Expense
	PCTFREE 5;

CREATE TABLE  ������˼�¼(
	����   number(2),
	����ID NUMBER(18),
	����ID NUMBER(18),
	��ҳID NUMBER(18),
	����� VARCHAR2(20),
	������� Date ,
	ת��ID number(18),
	ת���� VARCHAR2(20),
	ת��ʱ�� DATE,
	��¼״̬ NUMBER(2))
     TABLESPACE zl9Expense;

Create Table ҽ���������(
    ���� date,
    ������ VARCHAR2(20),
    ִ���� VARCHAR2(20),
    ���˲���id NUMBER(18),
    ���˿���id NUMBER(18),
    ��������id NUMBER(18),
    ִ�в���id NUMBER(18),
    ������Ŀid NUMBER(18),
    ��Դ;�� NUMBER(3),
    ���ʷ��� NUMBER(1),
    Ӧ�ս�� NUMBER(16,5),
    ʵ�ս�� NUMBER(16,5),
    ���ʽ�� NUMBER(16,5))
	TABLESPACE zl9Expense
	PCTFREE 5;

Create Table ���˷��û���(
    ���� date,
    ���˲���id NUMBER(18),
    ���˿���id NUMBER(18),
    ��������id NUMBER(18),
    ִ�в���id NUMBER(18),
    ������Ŀid NUMBER(18),
    ��Դ;�� NUMBER(3),
    ���ʷ��� NUMBER(1),
    Ӧ�ս�� NUMBER(16,5),
    ʵ�ս�� NUMBER(16,5),
    ���ʽ�� NUMBER(16,5))
    TABLESPACE zl9Expense
    PCTFREE 5;

Create Table ����δ�����(
    ����id NUMBER(18),
    ��ҳid NUMBER(5),
    ���˲���id NUMBER(18),
    ���˿���id NUMBER(18),
    ��������id NUMBER(18),
    ִ�в���id NUMBER(18),
    ������Ŀid NUMBER(18),
    ��Դ;�� NUMBER(3),
    ��� NUMBER(16,5))
    TABLESPACE zl9Expense;

Create Table ����Ԥ����¼(
    ID NUMBER(18),
    ��¼���� NUMBER(3),
    NO VARCHAR2(8),
    ʵ��Ʊ�� VARCHAR2(20),
    ��¼״̬ NUMBER(3),
    ����id NUMBER(18),
    ��ҳid NUMBER(18),
    ����id NUMBER(18),
    �ɿλ VARCHAR2(50),
    ��λ������ VARCHAR2(50),
    ��λ�ʺ� VARCHAR2(20),
    ժҪ VARCHAR2(50),
    ��� NUMBER(16,5),
    ���㷽ʽ VARCHAR2(20),
    ������� VARCHAR2(30),
    �տ�ʱ�� Date,
    ����Ա��� VARCHAR2(6),
    ����Ա���� VARCHAR2(20),
    ��Ԥ�� NUMBER(16,5),
    ����id NUMBER(18),
	�ɿ� NUMBER(16,5),
	�Ҳ� NUMBER(16,5),
	�ɿ���ID number(18),
	Ԥ����� number(1),
	�����ID number(18),
	���㿨��� number(18),
	���� varchar2(50),
	������ˮ�� varchar2(50),
	����˵�� varchar2(500),
	������λ  VARCHAR2(50),
	������� NUMBER(18),
	У�Ա�־ number(2),
	��ת�� Number(3))
    TABLESPACE zl9Expense
    initrans 20;
 

CREATE TABLE �������㽻��(
	����ID Number(18),
	������Ŀ varchar2(50),
	�������� VARCHAR2(100),
	��ת�� Number(3))
	TABLESPACE zl9Expense;

Create Table Ʊ������¼(
    ID NUMBER(18),
    Ʊ�� NUMBER(1),
    ʹ����� VARCHAR2(50),
    ����Ʊ�� number(1),
    ǰ׺�ı� VARCHAR2(2),
    ��ʼ���� VARCHAR2(50),
    ��ֹ���� VARCHAR2(50),
    ������� NUMBER(10),
    ʣ������ NUMBER(10),
    ��ע VARCHAR2(200),
    �Ǽ��� VARCHAR2(20),
    �Ǽ�ʱ�� DATE)
    TABLESPACE zl9Expense;

Create Table Ʊ�ݱ����¼(
    ID NUMBER(18),
    ���ID NUMBER(18),
    ��ʼ���� VARCHAR2(50),
    ��ֹ���� VARCHAR2(50),
    ���� NUMBER(10),
    ����ԭ�� VARCHAR2(200),
    ������ VARCHAR2(20),
    ����ʱ�� DATE)
    TABLESPACE zl9Expense;

Create Table Ʊ�����ü�¼(
    ID NUMBER(18),
    Ʊ�� NUMBER(1),
    ʹ����� VARCHAR2(50),
    ������ VARCHAR2(20),
    ǰ׺�ı� VARCHAR2(2),
    ��ʼ���� VARCHAR2(50),
    ��ֹ���� VARCHAR2(50),
    ʹ�÷�ʽ NUMBER(1),
    �Ǽ�ʱ�� DATE,
    ʹ��ʱ�� DATE,
    �Ǽ��� VARCHAR2(20),
    ��ǰ���� VARCHAR2(50),
    ʣ������ NUMBER(10),
	���� VARCHAR2(20),
	�˶��� VARCHAR2(20),
	�˶�ʱ�� DATE,
	�˶Խ�� NUMBER(1),
	�˶�ģʽ NUMBER(1),
	ǩ���� varchar2(20),
	ǩ��ʱ�� DATE,
	��ע VARCHAR2(200),
	��ת�� Number(3)
	)
    TABLESPACE zl9Expense
    PCTFREE 5 initrans 20;

Create Table Ʊ��ʹ����ϸ(
	ID	Number(18),
	Ʊ�� NUMBER(1),
	���� VARCHAR2(50),
	���� NUMBER(1),
	ԭ�� NUMBER(1),
	����ID NUMBER(18),
	���մ��� NUMBER(3),
	��ӡID NUMBER(18),
	ʹ��ʱ�� DATE,
	ʹ���� VARCHAR2(20),
	�˶��� VARCHAR2(20),
	�˶�ʱ�� DATE,
	�˶Խ�� NUMBER(1),
	��ע VARCHAR2(200),
	��ת�� Number(3))
    TABLESPACE zl9Expense
    PCTFREE 5 initrans 20;

Create Table Ʊ�ݴ�ӡ����(
	ID NUMBER(18),
	�������� NUMBER(3),
	NO VARCHAR2(8),
	��ת�� Number(3))
	TABLESPACE zl9Expense
    PCTFREE 5 initrans 20;

Create Table Ʊ�ݴ�ӡ��ϸ(
	ʹ��ID	NUMBER(18),
	Ʊ��	NUMBER(1),
	NO	VARCHAR2(8),
	Ʊ��	VARCHAR2(50),
	�Ƿ���� NUMBER(1),
	����Ʊ����� NUMBER(18), 
	���	VARCHAR2(4000),
	��ת�� Number(3))
TABLESPACE zl9Expense
PCTFREE 5 initrans 20;

Create Table ��Ա�ɿ����(
    �տ�Ա VARCHAR2(20),
    ���㷽ʽ VARCHAR2(20),
    ���� NUMBER,
    ��� NUMBER(16,5),
    �ϴ�����ʱ�� DATE)
    TABLESPACE zl9Expense
    PCTFREE 20 initrans 20;

Create Table ��Ա�սɼ�¼(
	ID		Number(18),
	��¼����	Number(2),
	NO		varchar2(20),
	�տ�Ա		varchar2(20),
	�տ��ID	Number(18),
	��Ԥ����	Number(16,5),
	����ϼ�	Number(16,5),
	����ϼ�	Number(16,5),
	ժҪ		varchar2(50),
	��ʼʱ��	Date,	
	��ֹʱ��	Date,	
	�ɿ���ID	Number(18),
	�Ǽ���		varchar2(20),
	�Ǽ�ʱ��	Date,
	С���տ���	varchar2(20),
	С���տ�ʱ��	Date,	
	С���տ�ID	Number(18),
	С������ID	Number(18),
	�����տ���	varchar2(20),
	�����տ�ʱ��	Date,
	�����տ�ID	Number(18),
	������		varchar2(20),
	����ʱ��	Date,
	�սɱ�־        number(2),
	��ת��		Number(3))
TABLESPACE zl9Expense
PCTFREE 20;

Create Table ��Ա�ս���ϸ(
	�ս�ID	 number(18),
	���㷽ʽ Varchar2(20),
	�����	 Varchar2(10),
	���	 number(16,5),
	���	 number(16,5),
	��ת��	 number(3))
TABLESPACE zl9Expense
PCTFREE 5;


Create Table ��Ա�ս�Ʊ��(
	�ս�ID		Number (18),
	Ʊ��		Number(2),
	����		number(2),
	���		Number(18),
	Ʊ������	Number(18),
	��ʼƱ��	Varchar2(50),
	��ֹƱ��	Varchar2(50),
	���		number(16,5),
	����ʱ��	date,
	��ת��		Number(3))
TABLESPACE zl9Expense
PCTFREE 5;

Create Table ��Ա�սɶ���(
	�ս�ID Number(18),
	����   Number(2),
	��¼ID Number(18),
	��ת�� Number(3))
TABLESPACE zl9Expense
PCTFREE 5;
  
Create Table ��Ա�ݴ��¼(
	ID	 number(18),
	�ս�ID	 number(18),
	��¼���� number(2),
	NO	 varchar2(20),
	���㷽ʽ  varchar2(20),
	���	 number(16,5),
	�տ�Ա	 varchar2(20),
	����ʱ�� Date,
	�ջ���   varchar2(20),
	�ջ�ʱ�� Date,	
	��ע     varchar2(50),
	�Ǽ���   varchar2(20),
	�Ǽ�ʱ�� Date,	
	��ת�� number(3))
TABLESPACE zl9Expense
PCTFREE 5;

CREATE TABLE ��Ա����¼(
	ID number(18),
	����� number(16,5),
	��ע varchar2(100),
	����� varchar2(20),
	����ʱ�� Date,
	���㷽ʽ VARCHAR2(20)  NOT NULL,
	����� varchar2(20),
	���ʱ�� date,
	ȡ��ʱ�� DATE,
	ȡ��ԭ�� varchar2(100),
	��ת�� Number(3))
	TABLESPACE zl9Expense
	PCTFREE 5;

Create Table ����ɿ����(
    ID NUMBER(18),
    ������ VARCHAR2(50),
    ����     VARCHAR2(20),
    ˵��	VARCHAR2(50),
    ������ID Number(18),
    ɾ������ Date ,
    �ϴ�����ʱ�� Date )
    TABLESPACE zl9Expense
    Cache Storage(Buffer_Pool Keep);

Create Table �ɿ��Ա���(
    ��ID NUMBER(18),
    ��ԱID number(18))
    TABLESPACE zl9Expense
    Cache Storage(Buffer_Pool Keep);

Create Table �շ�����¼(
    ���� Date,
    �տ�Ա VARCHAR2(20),
    ���� NUMBER(1),  --1-Ԥ����,2-����,3-�շ�,4-�Һ�
    ��ʼʱ�� Date,
    ��ֹʱ�� Date)
    TABLESPACE zl9Expense
    PCTFREE 5;

--���ѿ�����
CREATE TABLE ���ѿ�Ŀ¼(
	ID		Number(18),
	�ӿڱ��        number(6),
	������		Varchar2(20),
	����		Varchar2(20),
	���		Number(18),	
	����		Varchar2(50),
	�������	Varchar2(500),
	�ɷ��ֵ	Number(2) DEFAULT 0,
	��Ч��		Date,
	����ԭ��	Varchar2(50),
	������		Varchar2(20),
	�쿨����ID      number(18),
	�쿨��		Varchar2(20),	
	����ʱ��	Date,	
	������		Varchar2(20),	
	����ʱ��	Date,	
	ͣ����          VARCHAR2(20),
	ͣ������        DATE,
	��ǰ״̬	Number(2) DEFAULT 1,
	��ע		varchar2(100),
	���㷽ʽ  varchar2(20),
	������	Number(16,5),
	���۽��	Number(16,5),
	��ֵ�ۿ���	Number(16,5),
	���	Number(16,5),
	������� number(18),
	�ɿ���ID number(18),
	������ID number(18),
	��λ������ VARCHAR2(50),
    ��λ�ʺ�   VARCHAR2(20),
    �������   VARCHAR2(30)
	) TABLESPACE zl9Expense;

CREATE TABLE ���ѿ���ֵ��¼ (
	ID		Number(18),
	���ѿ�ID	number(18),
	���            number(18),
	��¼״̬	number(18),--1-����,2-����
	���㷽ʽ  varchar2(20),
	��ֵ���	Number(16,5),
	��ֵ�ۿ�	Number(16,5),
	�ɿ���	Number(16,5),
	��ֵʱ��	Date,
	����Ա����	Varchar2(20),
	�ɿ���	Varchar2(20),
	��ע    varchar2(100),
	�ɿ���ID number(18),
	��λ������ VARCHAR2(50),
    ��λ�ʺ�   VARCHAR2(20),
    �������   VARCHAR2(30)
	) TABLESPACE zl9Expense
	PCTFREE 5;

CREATE TABLE ���˿������¼ (
	ID	Number(18),
	�ӿڱ�� NUMBER(18),
	���ѿ�ID Number(18),
	���     number(18),
	��¼״̬ number(18),
	���㷽ʽ Varchar2(20),
	������ Number(16,5),
	����    Varchar2(50),	
	������ˮ�� Varchar2(50),
	����ʱ�� DATE,
	��ע Varchar2(100),
	�����־ number(2) DEFAULT 1,
	��ת�� Number(3)
	) TABLESPACE zl9Expense
	PCTFREE 5;

CREATE TABLE ���˿�������� (
	Ԥ��ID	Number(18),
	������ID NUMBER(18),
	��ת�� Number(3)
	 ) TABLESPACE zl9Expense
	PCTFREE 5;


----------------------------------------------------------------------------
--[[15.ҩƷ����ҵ��]]
----------------------------------------------------------------------------
Create table ҩƷ���Ŷ���
(
ҩƷid number(18),
�������� varchar2(60),
���� varchar2(20),
���� number(18),
�ɱ��� number(16,7),
�ۼ� number(16,7)
) TABLESPACE zl9MedLst;
create table ҩƷ�������
(
�ⷿid number(18),
���� number(2),
ԭʼNO varchar2(8),
�ϴ�NO varchar2(8),
����NO varchar2(8),
����� varchar2(100),
������� date,
ժҪ varchar2(1000))
TABLESPACE zl9MedLst;

create table ���ۻ��ܼ�¼ 
(
       ���ۺ� varchar2(10),
       ���� number(1),
       ִ������ date,
       �������� date,
       ������ varchar2(20),
       ˵�� varchar2(100),
       ���� number(1)
) tablespace zl9MedLst;


Create Table ҩƷ�ɹ��ƻ�(
    ID NUMBER(18),
    No varchar(8),
    �ƻ����� NUMBER(3),
    �ڼ� VARCHAR2(8),
    �ⷿid NUMBER(18),
    ҩ��id NUMBER(18),
    ���Ʒ��� NUMBER(3),
    ����˵�� VARCHAR2(250),
    ������ VARCHAR2(20),
    �������� date,
    ����� VARCHAR2(20),
    ������� date,
    ������ VARCHAR2(20),
    �������� date)
    TABLESPACE zl9MedLst
    PCTFREE 5;

Create Table ҩƷ�ƻ�����(
    �ƻ�ID NUMBER(18),
    ҩƷid NUMBER(18),
    ��� NUMBER(5),
    ǰ������ NUMBER(16,5),
    �������� NUMBER(16,5),
    �������� NUMBER(16,5),
    �������� NUMBER(16,5),
    ������� NUMBER(16,5),
    �ƻ����� NUMBER(16,5),
    ִ������ NUMBER(16,5),
    ���� NUMBER (19,7),
    ��� NUMBER(18,5),
    �ϴι�Ӧ�� VARCHAR2(50),
    �ϴ������� VARCHAR2(60),
    ˵�� Varchar2(100),
    �ۼ� NUMBER (19,7),
    �ۼ۽�� NUMBER(18,5),
    �Ƿ��ϴ� Number(1) Default 0,
    �ͻ����� number(16,5))
    TABLESPACE zl9MedLst
    PCTFREE 5;

Create Table ҩƷ��ҩ�ƻ�(
    ID NUMBER(18),
    No VARCHAR2(8),
    ��� NUMBER(5),
    ҩƷid NUMBER(18),
    ��ҩ��λid NUMBER(18),
    ʵ������ NUMBER(16,5),
    �ɱ��� NUMBER(16,7),
    �ɱ���� NUMBER(16,5),
    ���� VARCHAR2(40),
    ���� VARCHAR2(20),
    Ч�� Date,
    ������ VARCHAR2(20),
    �������� Date,
    ժҪ VARCHAR2(100),
    ����� VARCHAR2(20),
    ������� Date)
    TABLESPACE zl9MedLst
    PCTFREE 5;

Create Table ���ϲɹ��ƻ�(
    ID NUMBER(18),
    ���� NUMBER(2) DEFAULT 0,
    No varchar(8),
    �ƻ����� NUMBER(3),
    �ڼ� VARCHAR2(6),
    �ⷿid NUMBER(18),
    ����ID NUMBER(18),
    ���Ʒ��� NUMBER(3),
    ����˵�� VARCHAR2(250),
    ������ VARCHAR2(20),
    �������� date,
    ����� VARCHAR2(20),
    ������� date)
    TABLESPACE zl9MedLst
    PCTFREE 5;

Create Table ���ϼƻ�����(
    �ƻ�ID NUMBER(18),
    ����id NUMBER(18),
    ��� NUMBER(5),
    ǰ������ NUMBER(16,5),
    �������� NUMBER(16,5),
    ������� NUMBER(16,5),
    �빺���� number(16,5),
    �ƻ����� NUMBER(16,5),
    ���� NUMBER (19,7),
    ��� NUMBER(18,5),
    �ϴι�Ӧ�� VARCHAR2(50),
    �ϴ������� VARCHAR2(60),
    �������� number(16,5),
    �������� number(16,5))
    TABLESPACE zl9MedLst
    PCTFREE 5;

Create Table ҩƷ���(
    �ⷿid NUMBER(18),
    ҩƷid NUMBER(18),
    ���� NUMBER(18),
    Ч�� DATE,
    ���� NUMBER(1),
    �������� NUMBER(18,5),
    ʵ������ NUMBER(18,5),
    ʵ�ʽ�� NUMBER(18,5),
    ʵ�ʲ�� NUMBER(18,5),
    �ϴι�Ӧ��id NUMBER(18),
    �ϴβɹ��� NUMBER(16,7),
    �ϴ����� VARCHAR2(20),
    �ϴ���������  date ,
    �ϴβ��� VARCHAR2(60),
    ���Ч�� Date,
    ��׼�ĺ� VARCHAR2(40),
    ���ۼ� NUMBER(16,7),
    �ϴο��� NUMBER(16,7),
    ��Ʒ���� Varchar2(50),
    �ڲ����� Varchar2(50),
    ƽ���ɱ��� number(16,7))
    TABLESPACE zl9MedLst
    initrans 20;

Create Table ҩƷ���(
    �ⷿid NUMBER(18),
    ҩƷid NUMBER(18),
    ���� NUMBER(18),
    �������� Date,
    ������� Date,
    ʵ������ NUMBER(18,5),
    ʵ�ʽ�� NUMBER(18,5),
    ʵ�ʲ�� NUMBER(18,5),
    ������� Date,
    ����־ Number(1),
	�Ƿ��ʼ Number(1))
    TABLESPACE zl9MedLst;

Create Table ҩƷ����(
    �ڼ� Varchar(8),
    ����id NUMBER(18),
    �ⷿid NUMBER(18),
    ҩƷid NUMBER(18),
    �������� NUMBER(18,5),
    ʵ������ NUMBER(18,5),
    ʵ�ʽ�� NUMBER(18,5))
    TABLESPACE zl9MedLst
    PCTFREE 5;

Create Table ҩƷ����ƻ�(
    ����id NUMBER(18),
    �ⷿid NUMBER(18),
    ҩƷid NUMBER(18),
    �������� NUMBER(18,5),
	����ID Number(18),
	״̬ Number(1),
	�Ǽ��� Varchar2(20),
	�Ǽ�ʱ�� Date,
	��ת�� Number(3))
    TABLESPACE zl9MedLst
    PCTFREE 5;

CREATE TABLE ҩƷǩ����¼(
	ID NUMBER(18),
	ǩ������ NUMBER(2),
	ǩ����Ϣ VARCHAR2(4000),
	ʱ��� DATE,
	ʱ�����Ϣ Varchar2(4000),
	֤��ID	NUMBER(18),
	ǩ��ʱ�� DATE,
	ǩ���� VARCHAR2(20),
	���� NUMBER(2),
	��ת�� Number(3))
	TABLESPACE zl9MedLst;

Create TABLE ҩƷǩ����ϸ(
	ǩ��ID NUMBER(18),
	�շ�ID NUMBER(18),
	��ת�� Number(3))
	TABLESPACE zl9MedLst;

Create Table ҩƷ�շ�����(
    ���� Date,
    �ⷿid NUMBER(18),
    ҩƷid NUMBER(18),
    ���id NUMBER(18),
    ���� NUMBER(2),
    ���� NUMBER(18,5),
    ��� NUMBER(18,5),
    ��� NUMBER(18,5))
    TABLESPACE zl9MedLst
    PCTFREE 5;

Create Table δ��ҩƷ��¼(
    ���� NUMBER(2),
    No VARCHAR2(8),
    ����id NUMBER(18),
    ��ҳid NUMBER(18),
    ���� VARCHAR2(100),
    ���ȼ� NUMBER(1),
    �Է�����id NUMBER(18),
    �ⷿid NUMBER(18),
    ��ҩ���� VARCHAR2(10),
    �������� Date,
    ���շ� NUMBER(1),
    ��ҩ�� VARCHAR2(20),
    ��ӡ״̬ NUMBER(1) Default 0,
    δ���� NUMBER(5),
    �������� Number(2),
    ��ҩ�� Varchar2(20),
    �Ŷ�״̬ Number(1),
    ����ʱ�� date,
    �������� varchar2(50))
    TABLESPACE zl9MedLst
    initrans 20;

Create Table �ɱ��۵�����Ϣ(
    Id NUMBER(18),
    �շ�id NUMBER(18),
    ��ҩ��λid NUMBER(18),
    �ⷿid NUMBER(18),
    ҩƷid NUMBER(18),
    ���� NUMBER(18),
    ���� VARCHAR2(20),
    Ч�� DATE,
    ���� VARCHAR2(60),
    ���Ч�� Date,
    ԭ�ɱ��� NUMBER(16,7),
    �³ɱ��� NUMBER(16,7),
    ��Ʊ�� VARCHAR2(200),
    ��Ʊ���� Date,
    ��Ʊ��� NUMBER(18,5),
    Ӧ����䶯 Number(1),
    ִ������ Date,
    ���ۻ��ܺ� Varchar2(10))
    TABLESPACE zl9MedLst
    PCTFREE 5;

Create Table ҩƷ�շ���¼(
    ID NUMBER(18),
    ��¼״̬ NUMBER(3),
    ���� NUMBER(2),
    No VARCHAR2(8),
    ��� NUMBER(5),
    �ⷿid NUMBER(18),
    ��ҩ��λid NUMBER(18),
    ������id NUMBER(18),
    �Է�����id NUMBER(18),
    ���ϵ�� NUMBER(2),
    ҩƷid NUMBER(18),
    ���� NUMBER(18),
    ���� VARCHAR2(60),
    ���� VARCHAR2(20),
    �������� date,
    Ч�� Date,
    ���� NUMBER(3) default 1,
    ��д���� NUMBER(16,5),
    ʵ������ NUMBER(16,5),
    �ɱ��� NUMBER(16,7),
    �ɱ���� NUMBER(16,5),
    ���� NUMBER(16,7),
    ���ۼ� NUMBER(16,7),
    ���۽�� NUMBER(16,5),
    ��� NUMBER(16,5),
    ժҪ VARCHAR2(1000),
    ������ VARCHAR2(20),
    �������� Date,
    ��ҩ�� VARCHAR2(20),
    ��ҩ���� DATE,
    ����� VARCHAR2(20),
    ������� Date,
    �۸�id NUMBER(18),
    ����id NUMBER(18),
    ���� NUMBER(18,7),
    Ƶ�� VARCHAR2(20),
    �÷� VARCHAR2(30),
    ��� VARCHAR2(100),
    ������� Date,
    ���Ч�� Date,
    ��Ʒ�ϸ�֤ VARCHAR2(100),
    ��ҩ��ʽ NUMBER(1),
    ��ҩ���� VARCHAR2(10),
    ������ VARCHAR2(20),
    ��׼�ĺ� VARCHAR2(40),
    ���ܷ�ҩ�� NUMBER(18),
    ע��֤�� varchar2(50),
    �ⷿ��λ Varchar2(50),
    ��Ʒ���� Varchar2(50),
    �ڲ����� Varchar2(50),
    �˲��� Varchar2(200),
    �˲����� date,
    ǩ��ȷ���� varchar2(20),
    ǩ��ʱ�� date,
    ��ת�� Number(3),
    �ƻ�id number(18))
    TABLESPACE zl9MedLst initrans 20
;

Create Table �շ���¼������Ϣ(
	�շ�ID number(18),
	���� varchar2(20),
	�������� varchar2(100),
	סԺ�� number(18),
	���� varchar2(10),
	��ת�� Number(3))
    TABLESPACE zl9MedLst
    PCTFREE 5;

Create Table �ݴ�ҩƷ��¼ (
       NO             VARCHAR2(8),
       ���           NUMBER(5),
       ����ID         Number(18),
       ����ID         Number(18),
       ҽ��ID         Number(18),
       ���ͺ�         Number(18),
       ҩƷID         Number(18),
       ҩƷ����       Varchar2(80),
       ���           Varchar2(100),
       ִ�з���       Number(2),    -- 0-���������� 1-��Һ�� 2-ע���� 3-Ƥ����
       ʹ��״̬       Number(1),    -- 0-δ��,1-����
       ժҪ           Varchar2(200),
       ���ϵ��       Number(2),    -- 1-���ݴ�ҩƷ -1-ʹ���ݴ�ҩƷ
       ��λ           varchar2(20), -- Ŀ¼�ڵ�ҩƷ��ҽ��ҩƷΪ���㵥λ ,Ŀ¼��ҩƷΪ���ﵥλ
       ����           Number(16,5),
       ����           Number(16,5), -- ��������,Ŀ¼�ڼ�¼���Ǽ��㵥λ����,Ŀ¼��Ϊ���ﵥλ����
       ����           Number(16,5),
       ���           Number(16,5),
       ����Ա         Varchar2(10),
       �Ǽ�ʱ��       Date,
       ����ʱ��       Date) --	ʹ��״̬Ϊ1�ļ�¼��������
       TABLESPACE zl9CisRec;

Create Table ����������Ϣ(
    �շ�ID NUMBER(18),
    ����ID NUMBER(18),
    ����ID NUMBER(18),
    ��ҳID NUMBER(18),
    ���� VARCHAR2(100),
    �Ա� VARCHAR2(4),
    ���� varchar2(20),
    ���� VARCHAR2(10),
    ҽ�Ƹ��ʽ VARCHAR2(20),
    ��ǰ����id NUMBER(18),
    ��ǰ����id NUMBER(18),
    ʹ��ʱ�� DATE,
    ���� VARCHAR2(20))
    TABLESPACE zl9MedLst
;

Create Table ҩƷ������¼(
    ID NUMBER(18),
    �ⷿid number(18),
    ҩƷid NUMBER(18),
    ���� NUMBER(18),
    ���� VARCHAR2(60),
    ���� VARCHAR2(20),
    ����ԭ�� VARCHAR2(20),
    �������� NUMBER(16,5),
    �ɱ����� number(16,7),
    �ɱ���� number(16,7),
    ���۵��� number(16,7),
    ���۽�� number(16,7),
    ˵��	VARCHAR2(50),
    ��ҩ��λid number(18),
    �Ǽ��� VARCHAR2(20),
    �Ǽ�ʱ�� Date,
    ����취 VARCHAR2(20),
    ������ VARCHAR2(20),
    ����ʱ�� Date,
    ���ⵥNO VARCHAR2(8))
    TABLESPACE zl9MedLst
    PCTFREE 5;

Create Table ҩƷ����¼(
    Id Number(18),
    �ⷿid Number(18),
    �ڳ����� Date,
    ��ĩ���� Date,
    ������ Varchar2(20),
    �������� Date,
    ����� Varchar2(20),
    ������� Date,
    �ϴν��ID Number(18),
    �ڼ� varchar2(6),
    ���� Number(1))
    TABLESPACE zl9MedLst;

Create Table ҩƷ�����ϸ(
    ���id Number(18),
    �ⷿid Number(18),
    ҩƷid Number(18),
    ���� Number(18),
    �ڳ����� Number(16,5),
    �ڳ���� Number(16,5),
    �ڳ���� Number(16,5),
    ��ĩ���� Number(16,5),
    ��ĩ��� Number(16,5),
    ��ĩ��� Number(16,5))
    TABLESPACE zl9MedLst;

Create Table ҩƷ������(
    Id Number(18),
    ���id Number(18),
    �ⷿid Number(18),
    ҩƷid Number(18),
    ���� Number(18),
    ������ Number(16,5),
    ���� Number(16,5),
    ��۲� Number(16,5))
    TABLESPACE zl9MedLst;

Create Table ��Һ��ҩ��¼(
    ID NUMBER(18),
    ����ID NUMBER(18),
    ��� NUMBER(18),
    ��ҩ���� Number(2),
    ���� VARCHAR2(100),
    �Ա� VARCHAR2(4),
    ���� varchar2(20),
    סԺ�� NUMBER(18),
    ���� VARCHAR2(10),
    ���˲���id NUMBER(18),
    ���˿���id NUMBER(18),
    ִ��ʱ�� Date,
    ƿǩ�� Varchar2(20),
    ��ӡ��־ Number(1),
    ҽ��ID Number(18),
    ���ͺ� Number(18),
    �Ƿ��� Number(1),
    ��ҩ���� NUMBER(18),
    ���ȼ� VARCHAR2(30),
    ���ʱ�� date,
    �Ƿ����� number(1),
    �Ƿ�������� number(1),
    �ֹ��������� number(1),
    ����״̬ number(2),
    ������Ա varchar2(20),
    ����ʱ�� date,
    ��ӡ��� number(5),
    ��ӡʱ�� date,
    ��ת�� Number(3),
    �Ƿ�ȷ�ϵ��� number(1))
    TABLESPACE zl9MedLst initrans 20
;

Create Table ��Һ��ҩ״̬(
  ��ҩid Number(18),
  �������� Number(2), 
  ������Ա Varchar2(20),
  ����ʱ�� Date,
  ����˵�� Varchar2(200),
  ��ת�� Number(1))
Tablespace Zl9medlst
    Initrans 20;

CREATE TABLE ��Һ��ҩ����(
    ��ҩID NUMBER(18),
    NO VARCHAR2(8),
	  ��ת�� Number(3))
    TABLESPACE zl9MedLst
    initrans 20;

Create Table ��Һ��ҩ����(
    ��¼ID NUMBER(18),
    �շ�ID NUMBER(18),
    ���� NUMBER(16,5),
	��ת�� Number(3))
    TABLESPACE zl9MedLst
    initrans 20;

Create Table �ⷿȷ�ϼ�¼(
    �ⷿid NUMBER(18),
    �·� VARCHAR2(6),
    ���� NUMBER(1),  --1-ҩƷ,2-����
    ��ʼʱ�� Date,
    ��ֹʱ�� Date)
    TABLESPACE zl9MedLst;

Create Table Ӧ����¼(
	ID number(18),
	��¼���� number(3),
	��¼״̬ NUMBER(3),
	NO varchar2(8),
	��Ŀid number(18),
	��� number(18),
	�շ�id NUMBER(18),
	��λID NUMBER(18),
        �ⷿID Number(18),
	Ʒ�� varchar2(80),
	��� varchar2(100),
	���� varchar2(50),
	���� varchar2(20),
	������λ varchar2(8),
	��ⵥ�ݺ� varchar2(8),
	���ݽ�� number(16,5),
	���� number(16,5),
	�ɹ��� number(19,5),
	�ɹ���� number(16,5),
	������� VARCHAR2(200),
	��Ʊ�� VARCHAR2(200),
	��Ʊ���� Date,
	��Ʊ��� NUMBER(18,5),
	��Ʊ�޸�ʱ�� Date,
	�ƶ����� Date,
	�ƻ���� number(16,5),
	�ƻ��� varchar2(20),
	�ƻ����� Date,
	������ varchar2(20),
	�������� Date,
	����� VARCHAR2(20),
	������� Date,
	ժҪ varchar2(1000),
	������� number(18),
	�ƻ���� number(18) Default 0,
	�����־ number(1) default 0,
	Ԥ�� number(1) default 0,
	ϵͳ��ʶ number(1),
	��Ʊ���� varchar2(20))
    TABLESPACE zl9DueRec
    PCTFREE 5;

Create Table Ӧ�����(
    ��λid NUMBER(18),
    ���� NUMBER(1),
    ��� NUMBER(18,5))
    TABLESPACE zl9DueRec;

Create Table �����¼(
    ID NUMBER(18),
    ��¼״̬ NUMBER(3),
    No VARCHAR2(8),
    ��� NUMBER(5),
    Ԥ���� NUMBER(1),
    ��λid NUMBER(18),
    ��� NUMBER(16,5),
    ���㷽ʽ VARCHAR2(20),
    ������� VARCHAR2(10),
    ժҪ VARCHAR2(50),
    ������ VARCHAR2(20),
    �������� Date,
    Ԥ���� VARCHAR2(20),
    Ԥ������ Date,
    ����� VARCHAR2(20),
    ������� Date,
    ������� NUMBER(18),
	�ܸ���־ number(1) Default 0)
    TABLESPACE zl9DueRec
    PCTFREE 5;


----------------------------------------------------------------------------
--[[16.�ٴ�ҽ��]]
----------------------------------------------------------------------------
Create Table ��Ѫ������(
  ҽ��ID number(18),
  ���   number(18),
  ������ĿID number(18),
  ָ����� varchar2(20),
  ָ�������� varchar2(60),
  ָ��Ӣ���� varchar2(40),
  ָ���� varchar2(500),
  �����λ varchar2(50),
  �����־ varchar2(10),
  ����ο� varchar2(500),
  ȡֵ���� varchar2(4000),
  �Ƿ��˹���д number(1),
  ��ת�� Number(3)
) TABLESPACE zl9CisRec;
Create Table ����ҽ����¼(
    ID NUMBER(18),
    ���ID NUMBER(18),
    ǰ��ID Number(18),
    ������Դ NUMBER(1),
    ����id NUMBER(18),
    ��ҳid NUMBER(5),
    �Һŵ� VARCHAR2(8),
    Ӥ�� NUMBER(3),
    ���� VARCHAR2(100),
    �Ա� VARCHAR2(4),
    ���� varchar2(20),
    ���˿���id NUMBER(18),
    ��� NUMBER(18),
    ҽ��״̬ NUMBER(3),
    ҽ����Ч NUMBER(1),
    ������� VARCHAR2(1),
    ������Ŀid NUMBER(18),
    �걾��λ VARCHAR2(60),
    ��鷽�� Varchar2(30),
    �շ�ϸĿid NUMBER(18),
    ���� Number(16,5),
    �������� NUMBER(16,5),
    �״����� NUMBER(16,5),
    �ܸ����� NUMBER(16,5),
    ҽ������ VARCHAR2(1000),
    ҽ������ VARCHAR2(100),
    ִ�п���id NUMBER(18),
    Ƥ�Խ�� VARCHAR2(10),
    ִ��Ƶ�� VARCHAR2(20),
    Ƶ�ʴ��� NUMBER(3),
    Ƶ�ʼ�� NUMBER(3),
    �����λ VARCHAR2(4),
    ִ��ʱ�䷽�� VARCHAR2(100),
    �Ƽ����� NUMBER(1),
    ִ������ NUMBER(1),
    ִ�б�� Number(1),
    ��˱�� Number(1),
    �ɷ���� NUMBER(3),
    ������־ NUMBER(1),
    ��ʼִ��ʱ�� DATE,
    ִ����ֹʱ�� DATE,
    �ϴ�ִ��ʱ�� DATE,
    �ϴδ�ӡʱ�� Date,
    ��������id NUMBER(18),
    ����ҽ�� VARCHAR2(41),
    ����ʱ�� DATE,
    У�Ի�ʿ VARCHAR2(20),
    У��ʱ�� DATE,
    ͣ��ҽ�� VARCHAR2(20),
    ͣ��ʱ�� DATE,
    ȷ��ͣ��ʱ�� Date,
    ȷ��ͣ����ʿ Varchar2(20),
    ����ʱ�� Date,
    �Ƿ��ϴ� number(1),
    ����� Number(1),
    ���δ�ӡ Number(1),
    ժҪ Varchar2(1000),
    ��Ѽ��� Number(1),
    ��ҩĿ�� Number(1),
    ��ҩ���� Varchar2(1000),
    ���״̬ Number(1),
    ������� NUMBER(18),
    ����˵�� varchar2(1000),
    �Ƿ������� number(1),
    �䷽ID Number(18),
    ������� number(2),
    �����ĿID Number(18),
    ������־ Number(1),
    �¿�ǩ��ID NUMBER(18),
    ��ת�� Number(3),
    ҩʦ��˱�־ number(1),
    ҩʦ���ʱ�� date)
    TABLESPACE zl9CisRec initrans 20
;

CREATE TABLE ����ҽ���Ƽ�(
		ҽ��ID NUMBER(18),
		�շ�ϸĿID NUMBER(18),
		���� NUMBER(16,5),
		���� NUMBER(16,5),
		���� Number(1),
		ִ�п���ID Number(18),
		�������� Number(1),
		�շѷ�ʽ Number(1),
		��ת�� Number(3))
    TABLESPACE zl9CisRec
    initrans 20;

CREATE TABLE ����ҽ��״̬(
    ҽ��ID NUMBER(18),
    �������� NUMBER(3),
    ������Ա VARCHAR2(20),
    ����ʱ�� DATE,
	����˵�� VARCHAR2(200),
	ǩ��ID Number(18),
	��ת�� Number(3))
    TABLESPACE zl9CisRec
    initrans 20;

Create Table ����ҽ������(
    ҽ��ID Number(18),
    ��Ŀ Varchar2(30),
    ���� Number(1),
    ���� Number(5),
    Ҫ��ID Number(18),
    ���� Varchar2(4000),
	��ת�� Number(3))
    Tablespace zl9CISRec
    PCTFREE 5;

Create Table ����ҽ������(
    ҽ��ID Number(18),
    ����ID Number(18),
	����״̬ Number(1),
	��ת�� Number(3))
    Tablespace zl9CISRec
    PCTFREE 5;

Create Table ������ļ�¼(
    ҽ��ID Number(18),
    ����ID Number(18),
	������ Varchar2(20),
	����ʱ�� Date,
	���Ĵ��� Number(5),
	ȡ��ʱ�� Date,
	��ת�� Number(3))
    Tablespace zl9CISRec
    PCTFREE 5;

Create Table ҽ��ǩ����¼(
    ID NUMBER(18),
    ǩ������ NUMBER(2),
    ǩ����Ϣ VARCHAR2(4000),
    ʱ��� DATE,
    ֤��ID NUMBER(18),
    ǩ��ʱ�� DATE,
    ǩ���� VARCHAR2(20),
    ��ת�� Number(3),
    ʱ�����Ϣ Varchar2(4000))
    TABLESPACE zl9CisRec
;

CREATE TABLE ����ҽ������(
		ҽ��ID NUMBER(18),
		���ͺ� NUMBER(18),
		��¼���� NUMBER(3),
		������� NUMBER(1),
		NO VARCHAR2(8),
		��¼��� NUMBER(18),
		�������� NUMBER(16,5),
		������ VARCHAR2(20),
		����ʱ�� DATE,
		�״�ʱ�� DATE,
		ĩ��ʱ�� DATE,
		����ʱ�� Date,
		ִ��״̬ NUMBER(3),
		ִ�в���id NUMBER(18),
		����� Varchar2(20),
		���ʱ�� Date,
		�Ʒ�״̬ NUMBER(3),
		ִ�м� Varchar(20),
		����ʱ�� Date,
		ִ�й��� Number(1),
		������ varchar2(20),
		����ʱ�� DATE,
		�������� VARCHAR2(18),
		������� Number(1),
		ִ��˵�� Varchar2(1000),
		������ varchar2(20),
		����ʱ�� date,
		�������� number(18),
		�ͼ��� varchar2(20),
		�����ӡ number(3),
		�걾�ͳ�ʱ�� Date,
		�زɱ걾 number(1),
		��ת�� Number(3))
    TABLESPACE zl9CisRec
    initrans 20;

CREATE TABLE ����ҽ��ִ��(
    ҽ��ID NUMBER(18),
    ���ͺ� NUMBER(18),
    Ҫ��ʱ�� DATE,
    �������� NUMBER(16,5),
    ִ��ժҪ VARCHAR2(200),
    ִ���� VARCHAR2(20),
    ִ��ʱ�� DATE,
    ִ�н�� number(1),
    �Ǽ��� VARCHAR2(20),
    �Ǽ�ʱ�� DATE,
    �˶��� VARCHAR2(20),
    �˶�ʱ�� date,
    ��ˮ��	Number(18), -- ��¼�ļ���ҽ��һ��ִ�е�
    �ӵ���	Varchar2(20),
    ��ҩ��	Varchar2(20),
    ����	Number(18), -- ���汾��ִ��һ���м���
    ���	Number(18), -- �������
    ����	Number(10,5), -- ����ĵ���
    ��ϵ��	Number(10,5), -- ����ĵ�ϵ��
    Һ����	Number(16,5), -- ҩƷ��Һ����
    ��ʱ	Number(10), --ִ������Ҫ�õ�ʱ�䣬��λ��
    ����	Number(10), --��ǰ����ʱ��������ѣ���λ��,-1��ʾ�����ѣ�0��ʾ�������ѣ�>0��ʾ��ǰ��ʱ��
    ˵�� 	Varchar2(200), --�ӵ���ʿ��дҩƷִ��ʱ�����˵���������䣬�ܹ�
    ��Һʱ�� Date,
    ��ת�� Number(3),
    ��Һͨ�� Varchar2(20))
    TABLESPACE zl9CisRec
    initrans 20;

Create Table ҽ��ִ��ʱ��
(
Ҫ��ʱ�� DATE,
ҽ��ID NUMBER(18),
���ͺ� NUMBER(18),
��ת�� Number(3)
)
TABLESPACE zl9CisRec
Initrans 20;

CREATE TABLE ҽ��ִ�мƼ�(
    ҽ��ID NUMBER(18),
    ���ͺ� NUMBER(18),
    Ҫ��ʱ�� DATE,
    �շ�ϸĿID NUMBER(18),
	�������� NUMBER(1) default(0),
    ���� NUMBER(16,5),
	��ת�� Number(3)
)
TABLESPACE zl9CisRec
initrans 20;

Create Table ҽ��ִ�д�ӡ
(
ҽ��ID   NUMBER(18),    
����ID   NUMBER(18),
�ϴδ�ӡʱ�� Date,
��ת�� Number(3)
)
TABLESPACE zl9CisRec
Initrans 20;

CREATE TABLE ����ִ�е���ӡ(
		����ID Number(18),
		��ҳID Number(18),
		Ӥ�� Number(3),
		����ID Number(18),
		ĩҳĩ�к� Number(3))
    TABLESPACE zl9CisRec;

CREATE TABLE ����ҽ������(
    ҽ��ID NUMBER(18),
    ���ͺ� NUMBER(18),
    ��¼���� NUMBER(3),
    NO VARCHAR2(8),
	��ת�� Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5
    initrans 20;

CREATE TABLE ����ҽ����ӡ(
    ҽ��ID NUMBER(18),
		ҳ�� NUMBER(5),
		�к� NUMBER(5),
		���� NUMBER(5),
		����ID Number(18),
		��ҳID Number(18),
		Ӥ�� Number(3),
		��Ч Number(1),
		��ӡ��� Number(1),
		��ӡʱ�� DATE,
		��ӡ�� VARCHAR2(20),
		����ҽ�� number(1),
		��ת�� Number(3))
    TABLESPACE zl9CisRec;

Create Global Temporary Table ҽ����ӡ��¼(
		ҽ��ID NUMBER(18),
		˳�� NUMBER(18),
		��ӡ��� NUMBER(1),
		����ҽ�� number(1))
    On Commit Delete Rows;

CREATE TABLE ���Ƶ��ݴ�ӡ(
    ��¼���� NUMBER(3),
    NO VARCHAR2(8),
	��ӡ���� Number(1),
	��ӡ�� Varchar2(20),
	��ӡʱ�� Date,
	��ת�� Number(3))
    TABLESPACE zl9CisRec;

Create Table ��Ѫ�����¼(
  ҽ��ID number(18),
  �Ƿ����  number(2),
  ��Ѫ����  number(2),
  ������Ѫʷ  number(2),
  �в����  number(2),
  ��Ѫ������  number(2),
  ��ѪѪ��  number(2),
  RHD number(2),
  ��Ѫ��Ѫ��  number(2),
  HCT  number(10,2),
  ALT  number(10,2),
  HBSAG  number(2),
  ÷��  number(2),
  Ѫ�쵰��  number(10,2),
  ѪС��  number(10,2),
  ANTIHCV  number(2),
  ANTIHIV12  number(2),
	��ת�� Number(3)
) TABLESPACE zl9CisRec;

Create Table ִ�д�ӡ��¼ (
       ҽ��ID     Number(18),
       ���ͺ�         Number(18),
       ��ˮ��     Number(18),
       ��ӡ˵��       Varchar2(1000),
       ��ӡʱ��       Date,
       ��ӡ��         Varchar2(20),
	   ��ת�� Number(3))
       TABLESPACE zl9CisRec
       Pctfree 5;


Create Table ��λ״����¼(
       ����ID         Number(18),
       ����ID         Number(18),
       ����           Varchar2(30), -- ���࣬�û�����������
       ���           Varchar2(30), -- ��λ���
       ���           Number(1),    -- 0-��ͨ��λ 1-���� 2-����ҩƷ��λ 3-VIP��λ
       �շ�ϸĿID     Number(18), -- ��Ҫ�շѣ����Ŷ�Ӧ���շ�ϸĿID
       ״̬           Number(1), -- 0-��,1-����,2-������,������ά��
       ����           Number(1), -- 0-��λ,1-��λ
       ��ע           Varchar2(100),
       NO             Varchar2(8),
       ���������  varchar2(50))
       TABLESPACE zl9CisRec;

Create Table �ŶӼ�¼(
       ����ID         Number(18),
       ����ID         Number(18),
       ����           Date Default Sysdate,
       ˳���         Number(5), -- �����Ŷӵ�˳���
       ��Ȩ��         Number(10), -- ���ⲡ�������¸ı�˳����
       ״̬           Number(2), -- 0-���� 1-��� 2-���� 3-�˺� 4-���� 5-������ 6-��ִ�� 7-ִ����
       ��ʼ����Ա  Varchar2(20),
       ��ʼʱ��    Date,
       ��������Ա  Varchar2(20),
       ����ʱ��    Date,
       �Һŵ�      Varchar2(8),
       ���б�־ NUMBER(1) default 0 not null,
       ��ע           Varchar2(100),
       ����̨	number(2))
       TABLESPACE zl9CisRec;

Create Table ���ﴩ��̨(
	ID	Number(18),
	����ID	Number(18),
	���	Number(2),
	��Ч	Number(1),
	���������	Varchar2(50),
	�Һŵ�1 Varchar2(8),
	�Һŵ�2 Varchar2(8))
	TABLESPACE zl9CisRec;

Create Table ��������־(
	ID	Number(18),
	����ID	Number(18),
	���	Number(1),
	����Դ	Varchar2(20),
	���������	Varchar2(50),
	���д���	Varchar2(20),
	����ʱ��	date,
	�������	number(2),
	ҽ��ID	number(18),
	���ͺ�	number(18),
	Ҫ��ʱ��	date,	
	ʣ��Һ����	number(18),
	��Ӧ��	Varchar2(20),
	��Ӧʱ��	date)
	TABLESPACE zl9CisRec;

Create Table ������Һ������־(
	ID	Number(18),
	����ID  Number(18),
	�Һŵ�	Varchar2(8),
	���	Number(2),    --1-�ŶӲ�����־ 2��ҽ��������־ 3-���в�����־ 4-��λ������־
	ʱ��	Date,
	���� Varchar2(4000),
	����Ա	Varchar2(20))
	TABLESPACE zl9CisRec PCTFREE 5;

create table ҵ����Ϣ�嵥
(
   ID Number(18),
   ����ID number(18),
   ����ID Number(18),
   �������ID Number(18),
   ���ﲡ��ID Number(18),
   ������Դ Number(1),
   ��Ϣ���� Varchar2(4000),
   ���ѳ��� varchar2(50),
   ���ͱ��� varchar2(100),
   ҵ���ʶ  varchar2(50),
   ���ȳ̶�  Number(3),
   �Ƿ�����  Number(1),
   �Ǽ�ʱ�� Date
) TABLESPACE zl9CisRec;

create table ҵ����Ϣ���Ѳ���
(
   ��ϢID Number(18),
   ����ID number(18) 
) TABLESPACE zl9CisRec;

create table ҵ����Ϣ������Ա
(
   ��ϢID Number(18),
   ������Ա varchar2(20)
) TABLESPACE zl9CisRec;

create table ҵ����Ϣ״̬
(
   ��ϢID Number(18),
   �Ķ����� number(3),
   �Ķ��� varchar2(20),
   �Ķ�ʱ�� date,
   �Ķ�����ID number(18)
) TABLESPACE zl9CisRec;

----------------------------------------------------------------------------
--[[17.�ٴ�·��]]
----------------------------------------------------------------------------
create table ����·������
(
·��ִ��ID  number(18),
����ID     varchar2(32)
)
TABLESPACE zl9CISRec;

CREATE TABLE �ٴ�·��Ŀ¼(
    ID NUMBER(18),
	���� VARCHAR2(50),
    ���� VARCHAR2(5),
    ���� VARCHAR2(100),
    ͨ�� NUMBER(1),
    ���°汾 NUMBER(3),
    �������� VARCHAR2(20),
    ���ò��� VARCHAR2(20),
	�����Ա� NUMBER(1),
	�������� VARCHAR2(10),
    ˵�� VARCHAR2(200),
	ȷ������ NUMBER(3),
	����·������ number(1),
    ���� NUMBER(1) default(0))
    TABLESPACE zl9BaseItem;

CREATE TABLE �ٴ�·����֧(
    ID  NUMBER(18),
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
    ����  VARCHAR2(50),
    ˵��  VARCHAR2(200),
    ǰһ�׶�ID NUMBER(18),
    ��׼סԺ�� VARCHAR2(10),
    ��׼���� VARCHAR2(20),
    ������ VARCHAR2(20),
    ����ʱ�� DATE)
    TABLESPACE zl9BaseItem;

CREATE TABLE �ٴ�·������(
    ·��ID NUMBER(18),
    ����ID NUMBER(18),
	���ID NUMBER(18),
	���� number(2))
    TABLESPACE zl9BaseItem;

CREATE TABLE �ٴ�·������(
    ·��ID NUMBER(18),
    ����ID NUMBER(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE �ٴ�·���ļ�(
    ·��ID NUMBER(18),
	�ļ��� VARCHAR2(200),
    ���� BLOB,
	������ VARCHAR2(20),
	����ʱ�� DATE,
	��� number(2)
	)
    TABLESPACE zl9BaseItem;

CREATE TABLE �ٴ�·���汾(
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
    ��׼סԺ�� VARCHAR2(10),
    ��׼���� VARCHAR2(20),
    �汾˵�� VARCHAR2(200),
    ������ VARCHAR2(20),
    ����ʱ�� DATE,
    ����� VARCHAR2(20),
    ���ʱ�� DATE,
		ͣ���� VARCHAR2(20),
    ͣ��ʱ�� DATE)
    TABLESPACE zl9BaseItem;

CREATE TABLE �ٴ�·���׶�(
		ID NUMBER(18),
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
		��ID NUMBER(18),
	��֧ID NUMBER(18),
    ��� NUMBER(5),
    ���� VARCHAR2(50),
    ��ʼ���� NUMBER(3),
    �������� NUMBER(3),
    ��־ VARCHAR2(10),
		���� VARCHAR2(50),
    ˵�� VARCHAR2(200))
    TABLESPACE zl9BaseItem;

CREATE TABLE �ٴ�·������(
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
    ��� NUMBER(5),
		���� VARCHAR2(50),
	��֧ID NUMBER(18))
    TABLESPACE zl9BaseItem;

Create Table �ٴ�·����Ŀ(
    ID NUMBER(18),
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
    �׶�ID NUMBER(18),
    ��֧ID NUMBER(18),
    ���� VARCHAR2(50),
    ��Ŀ��� NUMBER(5),
    ��Ŀ���� VARCHAR2(1000),
    ִ�з�ʽ NUMBER(1),
    ִ���� NUMBER(1),
    ��Ŀ��� VARCHAR2(500),
    ͼ��ID NUMBER(18),
    ����ο� varchar2(1500),
    ������ number(1),
    ����Ҫ�� NUMBER(1),
    ������ number(1))
    TABLESPACE zl9BaseItem
;

CREATE TABLE ·��ҽ���䶯(
    ��ĿID  NUMBER(18),
    ����ʱ��  Date,
    ����Ա  VARCHAR2(100),
    ҽ������ID  NUMBER(18),
    ���ID NUMBER(18),
    ��� NUMBER(5),
    ��Ч NUMBER(1),
    ������ĿID NUMBER(18),
    �շ�ϸĿID NUMBER(18),
    ҽ������ VARCHAR2(1000),
    �������� NUMBER(16,5),
    �ܸ����� NUMBER(16,5),
    �걾��λ VARCHAR2(60),
    ��鷽�� VARCHAR2(30),
    ҽ������ VARCHAR2(1000),
    ִ��Ƶ�� VARCHAR2(20),
    Ƶ�ʴ��� NUMBER(3),
    Ƶ�ʼ�� NUMBER(3),
    �����λ VARCHAR2(4),
    ִ������ NUMBER(1),
    ִ�б�� NUMBER(1),
    ִ�п���ID NUMBER(18),
    ʱ�䷽�� VARCHAR2(50),
    �Ƿ�ȱʡ Number(1) Default 0,
    �Ƿ�ѡ number(1) default 0,
    �䷽ID Number(18),
    �����ĿID Number(18))
   TABLESPACE zl9CISRec;
CREATE TABLE ·��ҽ������(
		ID NUMBER(18),
    ���ID NUMBER(18),
    ��� NUMBER(5),
    ��Ч NUMBER(1),
    ������ĿID NUMBER(18),
		�շ�ϸĿID NUMBER(18),
		ҽ������ VARCHAR2(1000),
		�������� NUMBER(16,5),
		�ܸ����� NUMBER(16,5),
		�걾��λ VARCHAR2(60),
		��鷽�� VARCHAR2(30),
		ҽ������ VARCHAR2(1000),
		ִ��Ƶ�� VARCHAR2(20),
		Ƶ�ʴ��� NUMBER(3),
		Ƶ�ʼ�� NUMBER(3),
		�����λ VARCHAR2(4),
		ִ������ NUMBER(1),
		ִ�б�� NUMBER(1),
		ִ�п���ID NUMBER(18),
		ʱ�䷽�� VARCHAR2(50),
		�Ƿ�ȱʡ Number(1) Default 0,
		�Ƿ�ѡ number(1) default(0),
		�䷽ID Number(18),
		�����ĿID Number(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE �ٴ�·��ҽ��(
		·����ĿID NUMBER(18),
    ҽ������ID NUMBER(18))
    TABLESPACE zl9BaseItem;

Create Table �ٴ�·������(
    ��ĿID NUMBER(18),
    �ļ�ID NUMBER(18),
    ԭ��ID VARCHAR2(32),
    ���� varchar2(100),
    ��� Number(5))
    TABLESPACE zl9BaseItem;

CREATE TABLE �ٴ�·������(
		ID NUMBER(18),
    ·��ID NUMBER(18),
    �汾�� NUMBER(3),
		�׶�ID NUMBER(18),
		�������� NUMBER(1),
	��֧ID NUMBER(18))
    TABLESPACE zl9BaseItem;

CREATE TABLE ·������ָ��(
		ID NUMBER(18),
    ����ID NUMBER(18),
    ��� NUMBER(5),
		����ָ�� VARCHAR2(200),
		ָ������ NUMBER(1),
		ָ���� VARCHAR2(500))
    TABLESPACE zl9BaseItem;

CREATE TABLE ·����������(
		����ID NUMBER(18),
    ָ��ID NUMBER(18),
    ��ĿID NUMBER(18),
		��ϵʽ VARCHAR2(5),
		����ֵ VARCHAR2(50),
		������� NUMBER(1))
    TABLESPACE zl9BaseItem;

CREATE TABLE �����ٴ�·��(
		ID NUMBER(18),
		����ID NUMBER(18),
		��ҳID NUMBER(5),
		����ID NUMBER(18),
		·��ID NUMBER(18),
		�汾�� NUMBER(3),
		������ VARCHAR2(20),
		����ʱ�� DATE,
		����˵�� VARCHAR2(1000),
		δ����ԭ�� Varchar2(6),
		��ʼʱ�� DATE,
		����ʱ�� DATE,
		״̬ NUMBER(1),
		��ǰ����   NUMBER(18),
		��ǰ�׶�ID NUMBER(18),
		ǰһ�׶�ID NUMBER(18),
		������� NUMBER(2),
		�����Դ NUMBER(1),
		����ID NUMBER(18),
		���ID NUMBER(18),
        �ϲ�·������ NUMBER(2),
		��ת�� Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;

Create Table ���˺ϲ�·��(
    ID         NUMBER(18),      
    ����ID     NUMBER(18),
    ��ҳID     NUMBER(5),
    ����ID     NUMBER(18),
    ·��ID     NUMBER(18),
    �汾��     NUMBER(3),
    ������     VARCHAR2(20),
    ����ʱ��     DATE,
    ����˵��     VARCHAR2(1000),
    ��ǰ����     NUMBER(18),
    ��ǰ�׶�ID   NUMBER(18),
    ǰһ�׶�ID   NUMBER(18),
    �������     NUMBER(2),
    �����Դ     NUMBER(1),
    ����ID     NUMBER(18),
    ��Ҫ·����¼ID  NUMBER(18),
    ��Ҫ·���׶�ID NUMBER(18),
    ��Ҫ·������   NUMBER(18),
    ����ʱ��     DATE,
	��ת�� Number(3)) 
TABLESPACE zl9CISRec;

Create Table ���˺ϲ�·������(
    ·����¼ID  NUMBER(18),      
    �׶�ID NUMBER(18),
	���� DATE,
    �ϲ�·����¼ID  NUMBER(18),
    �ϲ�·���׶�ID NUMBER(18),
    �ϲ�·������   NUMBER(18),
    �Ǽ�ʱ�� date,
	��ת�� Number(3)) 
TABLESPACE zl9CISRec;

CREATE TABLE ����·������(
	·����¼ID NUMBER(18),
	�׶�ID NUMBER(18),
	���� DATE,
    	����ԭ�� VARCHAR2(6),
	��ת�� Number(3))
    TABLESPACE zl9CISRec;
Create Table ����·��ִ��(
    ID NUMBER(18),
    ·����¼ID NUMBER(18),
    �׶�ID NUMBER(18),
    ���� DATE,
    ���� NUMBER(5),
    ���� VARCHAR2(50),
    ��ĿID NUMBER(18),
    ��Ŀ��� NUMBER(5),
    ��Ŀ���� VARCHAR2(1000),
    ִ���� NUMBER(1),
    ��Ŀ��� VARCHAR2(500),
    ����ԭ�� Varchar2(6),
    ���ԭ�� VARCHAR2(1000),
    ͼ��ID NUMBER(18),
    ִ���� VARCHAR2(20),
    ִ��ʱ�� DATE,
    ִ�н�� VARCHAR2(50),
    ִ��˵�� VARCHAR2(200),
    �Ǽ��� VARCHAR2(20),
    �Ǽ�ʱ�� DATE,
    �ϲ�·����¼ID NUMBER(18),
    �ϲ�·���׶�ID NUMBER(18),
    ��ת�� Number(3),
    ������ number(1))
    TABLESPACE zl9CISRec PCTFREE 5
;

CREATE TABLE ����·������(
		·����¼ID NUMBER(18),
		�׶�ID NUMBER(18),
		���� DATE,
		���� NUMBER(5),
		������ VARCHAR2(50),
		����ʱ�� DATE,
		������� NUMBER(2),
		����˵�� VARCHAR2(1000),
		����ԭ�� Varchar2(6),
		ʱ����� Number(1) Default 0,
		�Ǽ��� VARCHAR2(20),
		�Ǽ�ʱ�� DATE,
		��������� Varchar2(20),
		�������ʱ�� Date,
		��ת����� varchar2(20),
		��ת���ʱ�� date,
		ԭ·��ID NUMBER(18),
		ԭ·���汾 NUMBER(3),
		��ת�� Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;

CREATE TABLE ����·��ָ��(
		·����¼ID NUMBER(18),
		�׶�ID NUMBER(18),
		���� DATE,
		���� NUMBER(5),
		�������� NUMBER(1),
		����ָ�� VARCHAR2(200),
		ָ������ NUMBER(1),
		ָ���� VARCHAR2(50),
		�ϲ�·����¼ID Number(18),
		��ת�� Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;

CREATE TABLE ����·��ҽ��(
	·��ִ��ID NUMBER(18),
    ����ҽ��ID NUMBER(18),
	��ת�� Number(3))
    TABLESPACE zl9CISRec
    PCTFREE 5;

CREATE TABLE ���˳�����¼(
	����ID		NUMBER(18),
	��ҳID		NUMBER(18),
	�к�		NUMBER(5),	
    ·����¼ID  number(18),
	����ֵ		NUMBER(18),
	�ַ�ֵ		VARCHAR2(100),
	����ֵ		Date,
	��ע		VARCHAR2(1000),
	�Ǽ���		VARCHAR2(20),
	�Ǽ�ʱ��	DATE,
	��ת�� Number(3)
	)
TABLESPACE zl9CISRec;

CREATE TABLE ����·��ȡ��(
  ����ʱ�� Date,
  ������  VARCHAR2(20),
  �����  VARCHAR2(20),
  ����ID    NUMBER(18),
  ��ҳID    NUMBER(18)
  )
TABLESPACE zl9CISRec;

CREATE TABLE ·�������ļ�(
	ID		 NUMBER(18),	
	����ID	 NUMBER(18),
	�ڼ�	 VARCHAR2(20),
	��ʼʱ�� DATE,
	����ʱ�� DATE,
	·��ID	 NUMBER(18),	
	��д��	 VARCHAR2(20),	
	��дʱ�� DATE
	)
    TABLESPACE zl9CISRec;

CREATE TABLE ·�������¼(
	�ļ�ID	NUMBER(18),	
	�к�	NUMBER(3),
	��Ŀֵ	VARCHAR2(100),
	��ע	VARCHAR2(1000)
	)
    TABLESPACE zl9CISRec;

----------------------------------------------------------------------------
--[[18.����ҵ��]]
----------------------------------------------------------------------------
CREATE TABLE ���Ӳ�����¼(
    ID NUMBER(18),
    ��� NUMBER(4),
    ������Դ NUMBER(3),
    ����ID NUMBER(18),
    ��ҳID NUMBER(18),
    Ӥ�� NUMBER(5),
    ����ID NUMBER(18),
    �������� NUMBER(3),
    �ļ�ID NUMBER(18),
    �������� VARCHAR2(30),
    ������ VARCHAR2(20),
    ����ʱ�� DATE,
    ���ʱ�� DATE,
    ������ VARCHAR2(20),
    ����ʱ�� Date,
    ���汾 NUMBER(5),
    ǩ������ NUMBER(1),
    �鵵�� VARCHAR2(20),
    �鵵���� DATE,
    ����״̬ NUMBER(3),
	��ӡ��	Varchar2(20),
	��ӡʱ�� Date,
    �༭��ʽ Number(1) Default 0,
	·��ִ��ID Number(18),
	��ת�� Number(3))
    TABLESPACE zl9EprDat;

CREATE TABLE ���Ӳ�����ʽ(
    �ļ�ID NUMBER(18),
    ���� BLOB,
	��ת�� Number(3))
	LOB(����) Store as (Cache)
    TABLESPACE zl9EprLob
    PCTFREE 20;

CREATE TABLE ���Ӳ�������(
    ����ID NUMBER(18),
    ��� NUMBER(5),
    �ļ��� VARCHAR2(50),
    ���� BLOB,
    ��С NUMBER(12,2),
    ������ VARCHAR2(20),
    ���� Date,
	��ת�� Number(3))
	LOB(����) Store as (Cache)
    TABLESPACE zl9EprLob;

CREATE TABLE ���Ӳ�������(
    ID NUMBER(18),
    �ļ�ID NUMBER(18),
    ��ʼ�� NUMBER(5),
    ��ֹ�� NUMBER(5),
    ��ID NUMBER(18),
    ������� NUMBER(18),
    �������� NUMBER(1),
    ������ NUMBER(18),
    �������� NUMBER(1),
    �������� VARCHAR2(1000),
    �����д� NUMBER(18),
    �����ı� VARCHAR2(4000),
    �Ƿ��� NUMBER(1),
    Ԥ�����ID NUMBER(18),
		�������ID Number(18),
    ������� NUMBER(1),
    ʹ��ʱ�� VARCHAR2(2),
    ����Ҫ��ID NUMBER(18),
		�滻�� NUMBER(1),
    Ҫ������ VARCHAR2(40),
    Ҫ������ NUMBER(3),
    Ҫ�س��� NUMBER(3),
    Ҫ��С�� NUMBER(3),
    Ҫ�ص�λ VARCHAR2(50),
    Ҫ�ر�ʾ NUMBER(3),
    ������̬ NUMBER(3),
    Ҫ��ֵ�� VARCHAR2(4000),
	��ת�� Number(3))
    TABLESPACE zl9EprDat;

CREATE TABLE ���Ӳ���ͼ��(
    ����ID NUMBER(18),
    ͼ�� BLOB,
	��ת�� Number(3))
	LOB(ͼ��) Store as (Cache)
    TABLESPACE zl9EprLob;

CREATE TABLE �����䶯ԭ��(
	ID      Number(18),
	�����ļ�id  Number(18),
	�䶯ԭ��  Number(1),
	ԭ��Ҫ��id  Number(18),
	ԭ��Ҫ��  Varchar2(40),
	ԭ������  Varchar2(50))
	TABLESPACE zl9EprDat;

CREATE TABLE �����䶯���(
	ID          Number(18),
	�䶯ԭ��id  Number(18),
	�䶯���    Number(1),
	�������id  Number(18),
	���Ҫ��id  Number(18),
	���Ҫ��    Varchar2(40),
	���ֵ��  Varchar2(500),
	ԭʼֵ��  Varchar2(500))
	TABLESPACE zl9EprDat;

Create Table ���Ӳ���ʱ��(
    ID        Number(18),
    ����ID    Number(18),
    ��ҳID    Number(18),
    ������Դ  Number(1),
    ����ID    Number(18),
    ������    Varchar2(64),
    �ļ�ID    Number(18),
    ��������  Number(3),
    �������  Varchar2(3),
    ��������  Varchar2(30),
    �¼�      Varchar2(1000),
    ����      Number(1),
    Ψһ      Number(1),
    �¼�ʱ��   Date,
    ��ʼʱ��   Date,
    ����ʱ��   Date,
    һ������   Number(5),
    ��������   Number(5),
    ��Σ����   Number(5),
    ���ں�     Number(5),
    ��ɼ�¼ID Number(18),
    ���ʱ��   Date)
    TABLESPACE zl9EprDat
	PCTFREE 20 initrans 20;

CREATE TABLE ���Ӳ�����ӡ(
    ID NUMBER(18),
    �ļ�ID NUMBER(18),
    ���� NUMBER(3),
    ����ID NUMBER(18),
    ��ҳID NUMBER(18),
    ��ӡ�� Varchar2(64),
    ��ӡʱ��	Date)
    TABLESPACE zl9EprDat;

Create Table �����걨��¼(
    �ļ�ID NUMBER(18),
    ����״̬ NUMBER(3),
    �վ��� VARCHAR2(20),
    �վ�ʱ�� DATE,
    �վ�˵�� VARCHAR2(100),
    ������ VARCHAR2(20),
    ����ʱ�� DATE,
    ���͵�λ VARCHAR2(30),
    ���ͱ�ע VARCHAR2(100),
    �Ǽ��� VARCHAR2(20),
    �Ǽ�ʱ�� DATE,
    ���� VARCHAR2(100),
    �Ա� VARCHAR2(4),
    ���� varchar2(20),
    ְҵ VARCHAR2(80),
    ��ͥ��ַ VARCHAR2(100),
    ��ͥ�绰 VARCHAR2(20),
    �������� DATE,
    ȷ������ DATE,
    �������1 VARCHAR2(150),
    �������2 VARCHAR2(150),
    ���ע VARCHAR2(100),
    �ĵ�ID Varchar2(32),
    ��ת�� Number(3))
    TABLESPACE zl9EprDat
;

Create Table �����걨��Ӧ(
    �걨��Ŀ VARCHAR2(30),
    ��ӦҪ�� VARCHAR2(40))
    TABLESPACE zl9EprDat;

Create Global Temporary Table ��ʱ��������(
		ID NUMBER(18),
		�ļ�ID NUMBER(18),
		��ID NUMBER(18),
		������� NUMBER(18),
		�������� NUMBER(1),
		������ NUMBER(18),
		�������� NUMBER(1),
		�������� VARCHAR2(1000),
		��ʼ�� NUMBER(5),
		��ֹ�� NUMBER(5),
		�����д� NUMBER(18),
		�����ı� VARCHAR2(4000),
		�Ƿ��� NUMBER(1),
		Ԥ�����ID NUMBER(18),
		�������ID Number(18),
		������� NUMBER(1),
		ʹ��ʱ�� VARCHAR2(2),
		����Ҫ��ID NUMBER(18),
		�滻�� NUMBER(1),
		Ҫ������ VARCHAR2(40),
		Ҫ������ NUMBER(3),
		Ҫ�س��� NUMBER(3),
		Ҫ��С�� NUMBER(3),
		Ҫ�ص�λ VARCHAR2(50),
		Ҫ�ر�ʾ NUMBER(3),
		������̬ NUMBER(3),
		Ҫ��ֵ�� VARCHAR2(4000))
    On Commit Delete Rows;

Create Global Temporary Table ����ʱ�޼��(
    ����ID NUMBER(18),
    ��ҳID NUMBER(18),
    ������Դ NUMBER(3),
    �䶯�¼� VARCHAR2(40),
    �¼�ʱ�� DATE,
    �ļ�ID NUMBER(18),
    �������� NUMBER(3),
    ������� VARCHAR2(3),
    �������� VARCHAR2(30),
    Ψһ NUMBER(1),
    ����ID NUMBER(18),
    ������ VARCHAR2(20),
    ���ں� NUMBER(3),
    ����ʱ�� DATE,
    Ҫ��ʱ�� DATE,
    ��ɼ�¼ID NUMBER(18),
    ���ʱ�� DATE)
	On Commit Preserve Rows;

Create Global Temporary Table �������ݼ��(
    ����id NUMBER(18),
    ��ҳid NUMBER(18),
    ������Դ NUMBER(3),
    ������¼id NUMBER(18),
    �������� NUMBER(3),
    �������� VARCHAR2(30),
    ������� DATE,
    ���id NUMBER(18),
    ��ٸ�id NUMBER(18),
    ��ٲ�� NUMBER(3),
    ������ NUMBER(5),
    ����ı� VARCHAR2(200),
    ��ʾ���� NUMBER(1),	--˵����0-��;1-��ʾ;2-����
    ��ʾ���� VARCHAR2(4000))
	On Commit Preserve Rows;

--���鵵
Create Table �����ύ��¼(
    ID			Number(18),
    ����id		Number(18),
    ��ҳid		Number(5),
    ��¼״̬	Number(3),
    �ύ��		Varchar2(20),
    �ύʱ��	Date,
    ������		Varchar2(20),
    ����ʱ��	Date,
    �鵵��		Varchar2(20),
    �鵵ʱ��	Date,
    ������		Varchar2(20),
    ����ʱ��	Date,
    ��������	Varchar2(255))
    TableSpace zl9CISAudit;

Create Table �������ռ�¼(
	ID NUMBER(18),
	����id NUMBER(18),
	��ҳID NUMBER(18),
	������ varchar2(20),
	������ varchar2(20),
	����ʱ�� Date,
	��¼ʱ�� date)
	TABLESPACE zl9CISAudit;

Create Table ����������ǩ(
    �ύid		Number(18),
    ���Ķ���	Number(3),
    �ļ�id		Number(18),
    ����ʱ��	Date)
    TableSpace zl9CISAudit;

Create Table ������ӡ��¼(
	����id		Number(18),
	��ҳid		Number(5),
	��ӡ����	Number(5),
	��ӡ���	Number(5),
	��ӡ����	Varchar2(100), 
	��ӡ��		Varchar2(20),	
	��ӡʱ��	Date)
	TABLESPACE zl9CISAudit;

Create Table ����������¼(
    ID			Number(18),
    ���id		Number(18),
    �ύid		Number(18),
    ����id		Number(18),
    ��ҳid		Number(5),
    ��������	Number(3),
    �ļ�id		Varchar2(32),
    ҽ��id		Number(18),
    ����id		Number(18),
    ��¼����	Number(3),
    ��¼״̬	Number(3),
    �������	Varchar2(255),
    ������Ŀid	Number(18),
    ������		Varchar2(20),
    ����ʱ��	Date,
    ��������	Date,
    ����˵��	Varchar2(255),
    ������		Varchar2(20),
    ����ʱ��	Date,
    ��ֵ NUMBER(8,2),
    ����˵�� VARCHAR2(255),
    �������� NUMBER(5),
    ���ּ��� Varchar2(1),
    ���� Number(1),
    ������¼ Varchar2(200),
	���ĵ�ID Varchar2(32))
    TableSpace zl9CISAudit
    PCTFREE 5;

Create Table ����������ʷ(
    ID			Number(18),
    ���id		Number(18),
    �ύid		Number(18),
    ����id		Number(18),
    ��ҳid		Number(5),
    ��������		Number(3),
    �ļ�id		Varchar2(32),
    ҽ��id		Number(18),
    ����id		Number(18),
    ��¼����	Number(3),
    ��¼״̬	Number(3),
    �������	Varchar2(255),
    ������Ŀid	Number(18),
    ������		Varchar2(20),
    ����ʱ��	Date,
    ��������	Date,
    ����˵��	Varchar2(255),
    ������		Varchar2(20),
    ����ʱ��	Date,
    ��ֵ NUMBER(8,2),
    ����˵�� VARCHAR2(255),
    �������� NUMBER(5),
    ���ּ��� Varchar2(1),
    ���� Number(1),
    ������¼ Varchar2(200),
	���ĵ�ID Varchar2(32))
    TableSpace zl9CISAudit
    PCTFREE 5;

Create Table �������ļ�¼(
	ID		Number(18),
	No		Varchar2(10),
	��¼״̬	Number(3),   
	������	Varchar2(20),	
	��������	Varchar2(255),
	����ʱ��	Date,
	��������	Date,
	����ʱ��	Date,
	��������	Date,
	��׼��	Varchar2(20),
	��׼ʱ��	Date,
	�ܽ�����	Varchar2(255),
	�ܽ���	Varchar2(20),
	�ܽ�ʱ��	Date,
	�Ǽ�ʱ��	Date,
	�ջ���		Varchar2(20),
	�黹ʱ��	Date)
	TABLESPACE zl9CISAudit;

Create Table ������������(
    ����id		Number(18),
    ����id		Number(18),
    ��ҳid		Number(5))
    TableSpace zl9CISAudit;

Create Table ����������Ա(
    ����id		Number(18),
    ��Աid		Number(18))
    TableSpace zl9CISAudit;

Create Table ��������¼(
    ID			Number(18),
    ����id		Number(18),
    ��ҳid		Number(5),
    ��¼״̬	Number(3),
    �����		Varchar2(20),
    ���ʱ��	Date,
    �������	Varchar2(255),
    �����		Varchar2(20),
    ���ʱ��	Date)
    TableSpace zl9CISAudit;

--��������
Create Table �������ַ���(
	ID number(18),
	���� varchar2(50),
	�ܷ� number(8,2) default 100,
	��ֵ number(8,2),
	��ֵ number(8,2),
	���� varchar2(10),
	���� varchar2(10),
	ѡ�� number(1) default 0,
	����ʱ�� Date,
	ͣ��ʱ�� Date)
    	TABLESPACE zl9CISAudit;

Create Table �������ֱ�׼(
	ID number(18),
	�ϼ�ID number(18),
	����ID number(18),
	���� varchar2(50),
	���� varchar2(4000),
	��׼��ֵ number(8,2),
	ȱ�ݵȼ� varchar2(2),
	���ֵ�λ varchar2(8),
	�ϼ���� NUMBER(18),
	��� NUMBER(18),
	�ж����� Varchar2(4000),
	����ȼ� varchar2(2),
	����Դ Number(1) Default 0)
    	TABLESPACE zl9CISAudit;

Create Table �������ֽ��(
	ID number(18),
	����ID number(18),
	��ҳID number(5),
	����ID number(18),
	�ܷ� number(8,2),
	�ȼ� varchar2(2),
	�����޸� number(1),
	�������� varchar2(20),
	��ע	varchar(50),
	������ varchar2(20),
	����ʱ�� Date,
	����� varchar2(20),
	���ʱ�� Date)
	TABLESPACE zl9CISAudit;

Create Table ����������ϸ(
	ID number(18),
	����ID number(18),
	���ֱ�׼ID number(18),
	������� number(8,2),
	ȱ�ݵȼ� varchar2(2),
	�ɷ��޸� Number(1) Default 0,
	��ע	varchar(50))
	TABLESPACE zl9CISAudit;

CREATE TABLE ��˽������Ŀ(
    ��ĿID NUMBER(18))
    TABLESPACE zl9EprDat;

CREATE TABLE �������͵�λ(
    ���� VARCHAR2(2),
    ���� VARCHAR2(30),
    ���� VARCHAR2(10))
    TABLESPACE zl9BaseItem;

----------------------------------------------------------------------------
--[[19.����ҵ��]]
----------------------------------------------------------------------------

CREATE TABLE ���˻����ļ�(
	ID NUMBER (18),
	����ID NUMBER (18),
	����ID NUMBER (18),
	��ҳID NUMBER (18),
	Ӥ�� NUMBER (3),
	��ʽID NUMBER (18),
	�ļ����� VARCHAR2 (50),
	��ʼʱ�� DATE ,
	����ʱ�� DATE,
	����ID NUMBER (18),			--ͬһ�����˵��ļ���,ֻ���ļ�ID��ͬ���ļ��������������µ����⣨�ϲ�ʱ��ģ����㣩
	�鵵�� VARCHAR2 (20),
	�鵵ʱ�� DATE ,
	������ VARCHAR2 (20),
	����ʱ�� DATE,
	��ת�� Number(3))
	PCTFREE 20 initrans 10  
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ���˻����ӡ(
	�ļ�ID NUMBER (18),
	��¼ID NUMBER (18),
	����ʱ�� DATE ,
	���� NUMBER (3),
	��ʼҳ�� NUMBER (5),
	��ʼ�к� NUMBER (5),
	����ҳ�� NUMBER (5),
	�����к� NUMBER (5),
	�в� NUMBER (5) DEFAULT 0,	--��¼���ϴ��޸ĺ����������У�0��ʾ����δ�����仯
	��ӡ�� VARCHAR2 (20),
	��ӡʱ�� DATE,
	��ӡҳ�� NUMBER (5),
	��ӡ�к� NUMBER (5),
	��ӡ��ʶ NUMBER(1),
	��ӡ����ҳ�� NUMBER(5),
	��ת�� Number(3))
	PCTFREE 20 initrans 10   
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ���˻�����Ŀ(
	�ļ�ID NUMBER (18),
	ҳ�� NUMBER (5),
	�к� NUMBER (5),
	��ͷ���� VARCHAR2 (20),		--���еı�ͷ���ݣ�ȱʡ���ж���
	��� NUMBER(1),				--��Ŀ�ڵ�ǰ�е�������ţ���1��ʼ�����Ϊ2
	��Ŀ��� NUMBER (5),		--ÿ��ֻ�ܰ�һ����Ŀ������ѡ�����͵���Ŀ
	��λ VARCHAR2 (50),
	����Ա VARCHAR2 (20),
	����ʱ�� DATE,
	��ת�� Number(3))
	PCTFREE 20 INITRANS 10  
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ���˻�������(
	ID NUMBER (18),
	�ļ�ID NUMBER (18),
	����ʱ�� DATE ,
	��ʾ NUMBER(2) DEFAULT 0,
	���汾 NUMBER (5),
	������ VARCHAR2 (20),
	����ʱ�� DATE,
	ǩ���� VARCHAR2 (50),
	����ǩ���� VARCHAR2(20),
	ǩ��ʱ�� VARCHAR2 (50),
	ǩ������ NUMBER(3),
	������� NUMBER(3) DEFAULT 0,
	�����ı� VARCHAR2(50),
	���ܱ�� NUMBER(2),
	��ʼʱ�� VARCHAR2(5),
	����ʱ�� VARCHAR2(5),
	��ת�� Number(3))
	PCTFREE 20 initrans 10  
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ���˻�����ϸ(
	ID NUMBER (18),
	��¼ID NUMBER (18),
	��¼���� NUMBER (3),	--������Ŀ=1���ϱ�˵��=2�������ձ��=4��ǩ����¼=5���±�˵��=6����ǩ��¼=15
	��Ŀ���� VARCHAR2 (20),
	��ĿID NUMBER (18),
	������ NUMBER (5),
	��Ŀ��� NUMBER (5),
	��Ŀ���� VARCHAR2 (20),
	��Ŀ���� NUMBER (3),
	��¼���� VARCHAR2 (4000),
	��Ŀ��λ VARCHAR2 (10),
	��¼��� NUMBER (3),	--ͨ����дΪ0����������дΪ1��������穵�������дΪ1
	���²�λ VARCHAR2 (10),
	��¼��� NUMBER (3),
	���Ժϸ� NUMBER (1),
	������Դ NUMBER (1) DEFAULT 0 ,	--0-�ֹ�¼��;1-��Դ�ڼ�¼��;2-��Դ�����µ�;3-��Դ��PDA;9-��ʷ���ݣ�Ϊ�˱�֤�������ݲ����µ����¼���ʽ���ܣ��������������µ��鿴����ȷ
	��ԴID NUMBER(18),				--��ϸID
	���� NUMBER (1) DEFAULT 0,		--1��ʾ��������¼ʹ��,���ڿ���ͬ������
	δ��˵�� VARCHAR2 (4000),
	��ʼ�汾 NUMBER (5),
	��ֹ�汾 NUMBER (5),
	��¼�� VARCHAR2 (20),
	��¼ʱ�� DATE ,
	��ʾ NUMBER(1) DEFAULT 0,
	��ת�� Number(3))
	PCTFREE 20 initrans 10  
	TABLESPACE ZL9EPRDAT;

CREATE TABLE ���˻���Ҫ������
(
  �ļ�id NUMBER(18),
  ҳ��   NUMBER(5),
  ���� VARCHAR2(60),
  ���� varchar2(1000),
  ����Ա  VARCHAR2(20),
  ����ʱ�� DATE,
  ��ת��  NUMBER(3)
)tablespace ZL9EPRDAT;

CREATE TABLE ������Ǽ�¼(
	����ID NUMBER(18),
	����ID NUMBER(18),
	��ҳID NUMBER(18),
	������� NUMBER(18),
	������ NUMBER(5),
	���� DATE)
	TABLESPACE ZL9EPRDAT;


Create Table ���˻����¼(
    ID NUMBER(18),
    ������Դ NUMBER(3),
    ����ID NUMBER(18),
    ��ҳID NUMBER(18),
    Ӥ�� NUMBER(5),
    ����ID NUMBER(18),
    ������ NUMBER(1),
    ����ʱ�� DATE,
    ���汾 Number(5),
    �鵵�� Varchar2(20),
    �鵵ʱ�� Date,
    ������ VARCHAR2(20),
    ����ʱ�� DATE,
    ��ת�� number(3))
    TABLESPACE zl9EprDat
;

Create Table ���˻�������(
    ID NUMBER(18),
    ��¼ID NUMBER(18),
    ��¼���� NUMBER(3),
    ��Ŀ���� VARCHAR2(20),
    ��ĿID NUMBER(18),
    ��Ŀ��� NUMBER(5),
    ��Ŀ���� VARCHAR2(20),
    ��Ŀ���� NUMBER(3),
    ��¼���� VARCHAR2(4000),
    ��Ŀ��λ VARCHAR2(10),
    ��¼��� NUMBER(3),
    ���²�λ Varchar2(10),
    ��¼��� Number(3),
    ���Ժϸ� Number(1),
    δ��˵�� varchar2(4000),
    ��ʼ�汾 Number(5),
    ��ֹ�汾 Number(5),
    ��¼�� VARCHAR2(20),
    �޸�ʱ�� DATE,
    ��ת�� number(3))
    TABLESPACE zl9EprDat 
;

----------------------------------------------------------------------------
--[[20.����ҵ��]]
----------------------------------------------------------------------------
Create TABLE ������ˮ�߱걾(
  ID        NUMBER(18),
  �걾ID    NUMBER(18),
  �����Ƿ����  number(1),
  ��ת�� Number(3)
) TABLESPACE zl9CisRec;
Create TABLE ������ˮ��ָ��(
  ID        NUMBER(18),
  �걾ID    NUMBER(18),
  ��Ŀid    NUMBER(18),
  �����Ƿ����  number(1),
  �������  varchar2(4000),
  ��ת�� Number(3)
) TABLESPACE zl9CisRec;
Create Table ����걾��¼(
	ID number(18),
	ҽ��ID number(18),
	�걾��� varchar2(20) not Null,
	����ʱ�� Date,
	������ varchar2(20),
	�걾���� varchar2(200),
	������ varchar2(20),
	����ʱ�� Date,
	����״̬ number(1),
	������ varchar2(20),
	����ʱ�� Date,
	����� varchar2(20),
	���ʱ�� Date,
	�ϲ������ number(18),
	��ӡ���� number(18),
	�������� number(1),
	����ID number(18),
	�������� varchar2(18),
	������ number(2),
	��ע varchar2(4000),
	δͨ�����ԭ�� varchar2(40),
	����ʱ�� Date,
	�걾��̬ varchar2(50),
	�Ƿ��ʿ�Ʒ number(1),
	ִ�п���id number(18),
	΢����걾 Number(1),
	No Varchar2(20),
	�Ƿ��� NUMBER(1),
	�걾��� NUMBER(1),
	���鱸ע VARCHAR2(400),
	������Դ NUMBER(1),
	����id NUMBER(18),  --���Ϊ��ͨ����ҽ��,��Ӧ�Ĳ�����Ϣ��¼;���ΪӤ��ҽ��,��ʾ��ĸ�׶�Ӧ�Ĳ�����Ϣ
	Ӥ�� NUMBER(3),	    --�ǵڼ���Ӥ��ҽ������������,��ͨΪ0
	���� VARCHAR2(100),
	�Ա� VARCHAR2(4),
	���� VARCHAR2(10),
	�������� NUMBER(4),
	���䵥λ VARCHAR2(10),
	������ Varchar2(20),
	�������id Number(20),
	�ϲ�ID number(18),
	���� VARCHAR2(10),
	��ʶ�� number(18),
	���˿��� varchar(24),
	���� number(1),
	�Һŵ� varchar2(8),
	����� number(18),
	סԺ�� number(18),
	�������� date,
	��ҳID number(5),
	������Ŀ varchar(1000),
	�������� varchar(20),
	������ varchar(20),
	����ʱ�� date,
	���� varchar2(20),
	������ varchar2(20),
	����ʱ�� date,
	һ������ varchar2(1000),
	�������� varchar2(1000),
	�������� varchar2(1000),
	���δͨ�� varchar2(2000),
	���Ϊ�� number(3),
	������ varchar2(30),
	����ʱ�� date,
	����λ�� varchar2(100),
	���滷�� varchar2(500),
	������ varchar2(30),
	����ʱ�� date,
	���ٷ�ʽ varchar2(100),
	��ת�� Number(3))
    TABLESPACE zl9CisRec;

--������Ŀ�ֲ�
Create Table ������Ŀ�ֲ�(
	ID		Number(18),
	�걾id		Number(18),
	��Ŀid		Number(18),
	ҽ��id		Number(18),
	ϸ��ID		number(18),
	��Χ		Number(1),
	��ת�� Number(3))		--���ֶ���ʱδ��
	TABLESPACE zl9CisRec
  PCTFREE 5;

Create Table �����Լ���¼(
	ҽ��id		Number(18),
	No		Varchar2(20),
	���		Number(18),
	����id		Number(18),
	����		Number(16,5),
	�̶� number(1),
	��ת�� Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;

Create Table ������ͨ���(
	ID number(18),
	����걾ID number(18) not Null,
	������ĿID number(18),
	������ varchar2(500),
	�����־ number(1),
	����ο� varchar2(500),
	�޸��� varchar2(20),
	�޸�ʱ�� Date,
	��¼���� number(2),
	ԭʼ��� varchar2(500),
	ԭʼ��¼ʱ�� Date,
	��¼�� varchar2(20),
	�Ƿ���� number(1),
	�޸�ԭ�� number(1),
	ϸ��ID number(18),
	����ID number(18),
	�������� varchar2(50),
	������ĿID number(18),
	������� NUMBER(5),
	OD varchar2(20),
	CUTOFF varchar2(20),
	SCO varchar2(20),
	ø���ID number(18),
	���ý�� number(3),
	��ҩ���� varchar2(100),
	ҩ����ID number(18),
	ϡ�ͱ��� NUMBER(16,5),
	���鱸ע varchar2(4000),
	��ת�� Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;

Create Table ����ҩ�����(
	ϸ�����ID number(18),
	������ID number(18),
	�޸��� varchar2(20),
	�޸�ʱ�� Date,
	��� varchar2(20),
	������� varchar2(10),
	��¼���� number(2),
	����ID number(18),
	ҩ������ Number(1),
	��ת�� Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;

Create Table �����ʿؼ�¼(
    �걾ID Number(18),
    �걾��� Varchar2(20),
    ������ Varchar2(20),
    ����ID Number(18),
    ����ʱ�� Date,
    ʱ�� Varchar2(8),
    �ʿ�ƷID Number(18),
    ���Դ��� Number(3),
    ���ü�¼ Number(3),
    �������Լ� Number(1),
    �°�װ�Լ� Number(1),
    ������У׼�� Number(1),
    �°�װУ׼�� Number(1),
    �°�װ������ Number(1),
    ����ά������ Number(1),
	��ת�� Number(3))
    Tablespace zl9CisRec;

Create Table �����ʿر���(
    ���ID Number(18),
    ��� Number(1),
    ���� Varchar2(100),
    ��ʾ Varchar2(500),
    ԭ�� Varchar2(500),
    ��ʩ Varchar2(500),
    ���� Varchar2(500),
    ������ Varchar2(20),
    ����ʱ�� Date,
    �鵵�� Varchar2(20),
    �鵵ʱ�� Date,
    ��ĿID number(18),
	��ת�� Number(3))
    Tablespace zl9CisRec;

Create Table �������ñ���(
    ���ID Number(18),
    ԭ�� Varchar2(500),
    ������ Varchar2(20),
    ����ʱ�� Date)
    Tablespace zl9CisRec;

Create Table ����ͼ����(
	ID			Number(18),
	�걾id		Number(18),
	ͼ������		varchar2(20),
	��ת��		Number(3),
	ͼ���		CLOB,
	ͼ��λ��		varchar2(4000))		
	TABLESPACE zl9CisRec PCTFREE 5;

Create Table ����ø���¼(
	ID		Number(18),
	���		VARCHAR(20) Not Null,
	����ʱ��	Date,
	����		VARCHAR(10),
	�ο�����	VARCHAR(10),
	���Ƶ��	VARCHAR(10),
	���ʱ��        varchar(10),
	���巽ʽ        VarChar(10),
	�հ���ʽ	Varchar(10),
	�Լ�����	Varchar(20),
	�Լ�Ч��	Date,
	�Լ�����	VarChar(50),
	���Է���	VarChar(30),
	����ID		number(18),
	�Ƿ���	Number(1),		--�Ƿ��͵���ʦ����վ
	OD���հ�	Number(1),		--�Ƿ��ȥ�հ׿��� 1=Ҫ��
	���λ��	varchar(50),		--���ڱ�����λ��
	���嵥��	Number(1),		--�Ƿ��ǽ��еĵ��浥����� 1=���浥��
	������Ŀ	VarChar(300),		--���ǵ��嵥���ֻ��һ����Ŀ �絥������ʽ��A_ID;B_ID:C_ID... ��8����Ŀ
	���Թ�ʽ	VarChar(1000),		--���ǵ��嵥���ֻ��һ����ʽ �絥������ʽ: A_��ʽ;B_��ʽ;C_��ʽ...��8����ʽ
	�����Թ�ʽ	VarChar(1000),		--���ǵ��嵥���ֻ��һ����ʽ �絥������ʽ: A_��ʽ;B_��ʽ;C_��ʽ...��8����ʽ
	CutOff��ʽ	VarChar(1000),		--���ǵ��嵥���ֻ��һ����ʽ �絥������ʽ: A_��ʽ;B_��ʽ;C_��ʽ...��8����ʽ
	���Խ��	varchar(3000),		--���1^���;���2^���...���12^���|���1^���;���2���...���12^��� ��8��ÿ��12�����Ϊ����Ϊ"^"
	�Լ���¼	varchar(1000)		--��¼�Լ�:�Լ�����;�Լ�Ч��;�Լ�����;���Է���|�Լ�����;�Լ�Ч��;�Լ�����;���Է���|.....
	)
	TABLESPACE zl9CisRec;

Create Table ���������¼(
	ID Number(18),
	�걾ID number(18),
	�������� number(2),  -- 0=��� 1=ȡ�����
	����Ա varchar(20),
	����ʱ�� date,
	��ת�� Number(3))
  TABLESPACE zl9CisRec;

Create Table ������ռ�¼(
	ID     number(18),
	ҽ��ID number(18),
	������ varchar(20),
	����ʱ�� date,
	�������� varchar(200),
	�ز���	varchar(20),
	�ز�ʱ�� date,
	��ת�� Number(3))
  TABLESPACE zl9CisRec;

Create Table ����ø���Լ�
(
  �Լ����� VARCHAR2(30),
  �Լ�Ч�� DATE,
  �Լ����� VARCHAR2(100),
  ���Է��� VARCHAR2(100),
  ������ĿID number(18)
) TABLESPACE zl9CisRec;

Create Table ���������¼(
	ID     NUMBER(18),
	�걾ID number(18),
	��; varchar(10),
	��ת�� Number(3))
  TABLESPACE zl9CisRec;

Create Table ����ǩ����¼(
    ����걾ID NUMBER(18),
    ǩ������ NUMBER(2),
    ǩ����Ϣ VARCHAR2(4000),
    ʱ��� DATE,
    ֤��ID NUMBER(18),
    ǩ��ʱ�� DATE,
    ǩ���� VARCHAR2(20),
    ��ת�� Number(3),
    ʱ�����Ϣ varchar2(4000))
    TABLESPACE zl9CisRec
;

Create global temporary Table ����ø����ӡ(
  ����     VARCHAR2(20),
  ����	   VARCHAR2(20),
  Col1	   VARCHAR2(20),
  Col2	   VARCHAR2(20),
  Col3	   VARCHAR2(20),
  Col4	   VARCHAR2(20),
  Col5	   VARCHAR2(20),
  Col6	   VARCHAR2(20),
  Col7	   VARCHAR2(20),
  Col8	   VARCHAR2(20),
  Col9	   VARCHAR2(20),
  Col10	   VARCHAR2(20),
  Col11	   VARCHAR2(20),
  Col12	   VARCHAR2(20))
  on commit delete rows;

Create global temporary Table �ʿؼ��̷���ӡ(
  �������� VARCHAR2(10),
  ����     Varchar2(2),
  �ⶨֵ   Varchar2(18),
  ��ֵ     Varchar2(18),
  SD       Varchar2(18),
  SI����   Varchar2(18),
  SI����   Varchar2(18),
  ���     VARCHAR2(10),
  ������   VARCHAR2(30))
  on commit preserve rows;    

create global temporary table �ʿؼ���ͼ��ӡ
(
  ��Ŀ     VARCHAR2(10),
  A01      varchar2(10),
  A02      varchar2(10),
  A03      varchar2(10),
  A04      varchar2(10),
  A05      varchar2(10),
  A06      varchar2(10),
  A07      varchar2(10),
  A08      varchar2(10),
  A09      varchar2(10),
  A10      varchar2(10),
  A11      varchar2(10),
  A12      varchar2(10),
  A13      varchar2(10),
  A14      varchar2(10),
  A15      varchar2(10),
  A16      varchar2(10),
  A17      varchar2(10),
  A18      varchar2(10),
  A19      varchar2(10),
  A20      varchar2(10)
)
on commit preserve rows;

----------------------------------------------------------------------------
--[[21.���ҵ��]]
----------------------------------------------------------------------------
Create Table Ӱ�����¼(
    ҽ��ID number(18),
    ���ͺ� number(18),
    Ӱ����� varchar2(10),
    ִ�п���ID number(18),
    ���� number(18),
    ���� varchar2(100),
    Ӣ���� varchar2(100),
    �Ա� varchar(4),
    ���� varchar2(20),
    �������� date,
    ��� number(16,5),
    ���� number(16,5),
    ������ number(1),
    ���UID varchar2(64),
    λ��һ varchar2(3),
    λ�ö� varchar2(3),
    λ���� varchar2(3),
    ����豸 varchar2(30),
    �Ƿ��ӡ number(1),
    ��鼼ʦ Varchar2(20),
    ��鼼ʦ�� Varchar2(20),
    Ӱ������ Varchar2(10),
    �������� Varchar2(10),
    Σ��״̬ number(1),
    ������� Varchar2(10),
    �������� Varchar2(200),
    ����ͼ�� varchar2(4000),
    �������� DATE,
    ������ varchar2(20),
    ����� varchar2(20),
    ������� Varchar2(20),
    ��ɫͨ�� Number(1),
    �����ӡ Number(1),
    ������ Varchar2(64),
    ������ Varchar2(64),
    ������� Varchar2(200),
    ��Ϸ��� VARCHAR2(100),
    ���Ž�Ƭ number(1),
    ���淢�� number(1),
    ����ID NUMBER(18),
    ���淢���� varchar2(10),
    ��Ƭ������ varchar2(10),
    ͼ��λ�� Number(1),
    ͼ������ Number(5),
    �Ƿ�ʦȷ�� Number(1),
    �Ƿ���ӽ�Ƭ Number(1),
    ��ת�� Number(3))
    TABLESPACE zl9CisRec
;

create table Ӱ��Σ��ֵ��¼(
    id number(18),
    ҽ��id number(18),
    �Ǽ��� varchar2(30),
    �Ǽ�ʱ�� date,
    ֪ͨʱ�� date,
    ֪ͨ��ʽ varchar2(20),
    ���ܿ��� varchar2(30),
    ������Ա varchar2(30),
    ������ varchar2(512),
	��ת�� Number(3))
    tablespace zl9CisRec;

Create Table Ӱ��������(
    ����UID varchar2(64),
    ���UID varchar2(64),
    ���к� number(10),
    �������� varchar2(64),
    �ɼ�ʱ�� Date,
	��ת�� Number(3))
    TABLESPACE zl9CisRec;

Create Table Ӱ����ͼ��(
    ͼ��UID varchar2(64),
    ����UID varchar2(64),
    ͼ��� number(10),
    ͼ������ varchar2(64),
    �ɼ�ʱ�� date,
    ͼ��ʱ�� date,
    ��� VARCHAR2(20),
    ͼ��λ�ò��� VARCHAR2(64),
    ͼ������ VARCHAR2(120),
    �ο�֡UID VARCHAR2(64),
    ��Ƭλ�� VARCHAR2(20),
    ���� VARCHAR2(10),
    ���� VARCHAR2(10),
    ���ؾ��� VARCHAR2(64),
    ��̬ͼ NUMBER(1),
    ��Ƭ��ӡ NUMBER(1),
    �������� varchar(64),
    ¼�Ƴ��� number(18),
	��ת�� Number(3))
    TABLESPACE zl9CisRec
    PCTFREE 5;

Create Table Ӱ����ʱ��¼(
    Ӱ����� varchar2(10),
    ���� number(18),
    ���� varchar2(100),
    Ӣ���� varchar2(100),
    �Ա� varchar(4),
    ���� varchar2(20),
    �������� date,
    ��� number(5),
    ���� number(5),
    ������ number(1),
    ���Ž�Ƭ number(1),
    ���UID varchar2(64),
    λ��һ varchar2(3),
    λ�ö� varchar2(3),
    λ���� varchar2(3),
    ����豸 varchar2(64),
    ����ͼ�� varchar2(2000),
    �������� DATE)
    TABLESPACE zl9CisRec
;

Create Table Ӱ����ʱ����(
    ����UID varchar2(64),
    ���UID varchar2(64),
    ���к� number(10),
    �������� varchar2(64),
    �ɼ�ʱ�� Date)
    TABLESPACE zl9CisRec;

Create Table Ӱ����ʱͼ��(
    ͼ��UID varchar2(64),
    ����UID varchar2(64),
    ͼ��� number(10),
    ͼ������ varchar2(64),
    �ɼ�ʱ�� date,
    ͼ��ʱ�� date,
    ��� VARCHAR2(20),
    ͼ��λ�ò��� VARCHAR2(64),
    ͼ������ VARCHAR2(120),
    �ο�֡UID VARCHAR2(64),
    ��Ƭλ�� VARCHAR2(20),
    ���� VARCHAR2(10),
    ���� VARCHAR2(10),
    ���ؾ��� VARCHAR2(64),
    ��̬ͼ NUMBER(1),
    �������� varchar(64),
    ¼�Ƴ��� number(18))    
    TABLESPACE zl9CisRec;
    
Create Table Ӱ�񱨸沵��(
  ID       Number(18),     
  ҽ��ID   Number(18) ,
  ����ID   Number(18),
  �������� Varchar2(512),
  ����ʱ�� Date,
  ������   Varchar2(64),
  ��ת�� Number(3))
  Tablespace zl9CisRec;    

Create Table Ӱ��鵵��ҵ(
    ���� number(10),
    ���� varchar2(20),
    ִ��ʱ�� Date,
    Դ�豸 varchar2(1),
    Ŀ���豸 varchar2(1),
    ָ���豸 varchar2(3),
    �Ƿ�Ǩ�� number(1),
    �Ƿ�ɾ�� number(1),
    ��ʼʱ�� Date,
    ����ʱ�� Date,
    ������� number(10),
    �Զ����� number(1),
    ִ�й��� number(1),
    �������� varchar2(250))
    TABLESPACE zl9CisRec;

Create Table ��Ƭ��ӡ��¼(
    ID	Number(18),
    ���ID	Number(18),
    ҽ��ID	Number(18),
    ��Ƭ��С	Varchar(20),
    ��ӡ��	Varchar(64),
    ��ӡʱ��	Date)
    TABLESPACE zl9CisRec;

Create Table Ӱ���ղ�����(
    ID   NUMBER(18),       
    �ղ�ID  NUMBER(18), 
    ҽ��ID  NUMBER(18), 
    �ղ�ʱ�� Date,
	��ת�� Number(3)
)TABLESPACE zl9CISREC;

create table Ӱ�����뵥ͼ��
(
    ID          NUMBER(18),      
    ҽ��ID      NUMBER(18),    
    ���뵥ͼ��  varchar2(64),           
    FTP·��     varchar2(100),
    �豸��      varchar2(3),
    ɨ����      varchar2(20),
    ɨ��ʱ��    date,
	��ת�� Number(3)
)
TABLESPACE zl9CISREC;

---����ҵ��
----------------------------------------------------------------------------
Create Table ��������Ϣ(
    ����ҽ��ID Number(18), 
    ҽ��ID Number(18),   
    ����� VARCHAR2(20),           
    ������� Number(1),
    ȡ�Ĺ��� Number(1) default 0,
    ��Ƭ���� Number(1) default 0,
    ���߹��� Number(1) default 0,
    ��Ⱦ���� Number(1) default 0,
    ���ӹ��� Number(1) default 0,
    �޼����� Varchar2(2048),
    ʣ��λ�� Varchar2(64),
    �������� Varchar2(64),
    �ۺ����� Varchar2(10),
    �ۺ���� varchar2(255),
    ����ʱ�� Date)
    TABLESPACE zl9CisRec; 

Create Table ����������Ϣ(
    ID Number(18),   
    ����ҽ��ID Number(18), 
    ������Ŀ VARCHAR2(20),   
    ���۽�� VARCHAR2(10),   
    ������� Varchar2(255),
    �Ľ����� Varchar2(255),
    ��ע Varchar2(1024),
    ������ Varchar2(64),
    ����ʱ�� date)
    TABLESPACE zl9CisRec; 

Create Table ����걾��Ϣ(
    �걾ID NUMBER(18),
    ҽ��ID Number(18),
    �ͼ�ID Number(18),
    �걾���� VARCHAR2(64),
    ������� NUMBER(1) default 0,
    �걾���� NUMBER(1) default 0,
    �ɼ���λ VARCHAR2(20),
    ԭ�б�� VARCHAR2(20),
    ���� Number(2) default 0,
    ���λ�� VARCHAR2(64),
    �������� Date,
    ��ע VARCHAR2(1024))
    TABLESPACE zl9CisRec;    

Create Table �����ͼ���Ϣ(
    ID NUMBER(18),   
    ҽ��ID NUMBER(18),
    �ͼ쵥λ VARCHAR2(64),
    �ͼ���� VARCHAR2(64),
    �ͼ��� VARCHAR2(64),
    �ͼ����� DATE Not Null,
    ��ϵ��ʽ VARCHAR2(64),
    �Ǽ��� VARCHAR2(64),
    ����״̬ NUMBER(1) default 1,
    ����ԭ�� VARCHAR2(1024),
    ֪ͨ�� VARCHAR2(64),
    ��ע VARCHAR2(1024))
    TABLESPACE zl9CisRec;
    
Create Table ����������Ϣ(
    ����ID Number(18),  
    ����ҽ��ID Number(18), 
    ������ Varchar2(64),
    ����ʱ�� Date,        
    �������� Number(1) default 0,
    ����ϸĿ Number(1) default 0,
    ����״̬ Number(1) default 0,
    �������� Varchar2(1024),
    �Ƿ��ӡ Number(1) default 0,
    ����״̬ Number(1) default 0,
    ���ʱ�� Date)
    TABLESPACE zl9CisRec;    
   
Create Table ����ȡ����Ϣ(
    �Ŀ�ID Number(18),
    ��� Number(18),
    ����ҽ��ID Number(18), 
    ����ID Number(18),
    �걾ID Number(18),
    �걾���� Varchar2(64),
    ȡ��λ�� Varchar2(64),
    ��״ Varchar2(64),
    ��ɫ Varchar2(20),
    ���� Varchar2(20),
    �걾�� Varchar2(20),
    ������ Number(2) default 1,   
    �Ƿ���� Number(1) default 0,
    �Ƿ��Ѹ� Number(1) default 0,
    �Ƿ����� Number(1) default 0,
    ��ȡҽʦ Varchar2(64),
    ��ȡҽʦ Varchar2(64),
    ��¼ҽʦ Varchar2(64),
    ȷ��״̬ Number(1) default 0,
    �鵵״̬ number(1) default 0,
    ȡ��ʱ�� Date)
    TABLESPACE zl9CisRec;   
   
Create Table �����Ѹ���Ϣ(
    ID Number(18),   
    �걾ID Number(18),
    ��ʼʱ�� Date,
    ����ʱ�� Number(5),
    ��ǰ�״� Number(2),
    ���״̬ Number(1) default 0,
    ����Ա Varchar2(64))
    TABLESPACE zl9CisRec;     

Create Table ������Ƭ��Ϣ(
    ID Number(18),  
    ����ҽ��ID Number(18), 
    �Ŀ�ID Number(18),
    ����ID Number(18),
    ��Ƭ���� Number(1) default 0,
    ��Ƭ��ʽ Number(1) default 0,
    ��Ƭʱ�� Date,
    ��Ƭ�� Number(2),
    ��Ƭ�� Varchar2(64),       
    ��ǰ״̬ Number(1) default 0,
    �鵵״̬ number(1) default 0,
    �嵥״̬ Number(1) default 0)
    TABLESPACE zl9CisRec;     

Create Table ������̱���(
    ID Number(18),  
    ����ҽ��ID Number(18), 
    �걾���� Varchar2(64),
    �������� Number(1),
    �������� Number(1),
    ����� Varchar2(2048),
    ������ Varchar2(2048),
    ����ͼ�� Varchar2(2048),
    ����ҽʦ Varchar2(64),        
    �������� Date,       
    ��ǰ״̬ Number(1) default 0,
    ��ע Varchar2(1024))
    TABLESPACE zl9CisRec;  

Create Table ��������Ϣ(
    ����ID Number(18), 
    �������� VARCHAR2(64),
    ʹ���˷� Number(5),
    �����˷� Number(5),
    �������� Date,
    ��Ч�� Number(4),
    �������� Date,
    ��¡�� Number(1),
    ���ö��� Varchar2(20),
    ������ Varchar2(10),
    Ӧ����� Varchar2(1024),
    �Ǽ��� Varchar2(64),
    �Ǽ�ʱ�� Date,
    ʹ��״̬ Number(1) default 1,
    ��ע Varchar2(1024))
    TABLESPACE zl9CisRec;  
   
Create Table �����ؼ���Ϣ(
    ID Number(18),    
    ����ҽ��ID Number(18), 
    �Ŀ�ID Number(18),
    ����ID Number(18),        
    ����ID Number(18),
    ��Ŀ˳�� Varchar2(20),
    �ؼ����� Number(1) default 0,
    �ؼ�ϸĿ Number(1) default 0,
    �������� Number(1) default 0,
    ��ǰ״̬ NUMBER(1) default 0,
    ���ʱ�� Date,    
    �ؼ�ҽʦ Varchar2(64),
    �嵥״̬ Number(1) default 0,
    �鵵״̬ number(1) default 0,
    ��Ŀ��� Varchar2(20) null)
    TABLESPACE zl9CisRec; 

Create Table �������ӳ�(
    ID Number(18),    
    ����ҽ��ID Number(18), 
    �ӳ�ԭ�� Varchar2(1024),        
    �ӳ����� Number(2) default 0,
    ��ʱ��� Varchar2(1024),
    ת���� Varchar2(64),
    �Ǽ��� Varchar2(64),
    �Ǽ�ʱ�� Date,    
    ��ǰ״̬ Number(1) default 0)
    TABLESPACE zl9CisRec; 

Create Table ���������Ϣ(
    ID Number(18),    
    ����ҽ��ID Number(18), 
    ����ҽʦ Varchar2(64),
    ����ҽʦ Varchar2(64),
    ���ﵥλ Varchar2(64),         
    ����ʱ�� Date not null,
    ��ֹʱ�� Date not null,
    �������� Number(1) default 0,
    ������� Varchar2(2048),
    ��Ͻ�� Varchar2(2048),
    ������ Varchar2(2048),    
    ���ʱ�� Date,
    ��ע Varchar2(1024),
    ��ǰ״̬ Number(1) default 0)
    TABLESPACE zl9CisRec; 

Create Table �����巴��(
    ID Number(18),   
    ����ID Number(18), 
    �ο������ VARCHAR2(200),
    ʵ������ Number(1) default 0,
    �������� VARCHAR2(10),
    ������� VARCHAR2(1024),
    ����ҽ�� VARCHAR2(64),
    ����ʱ�� Date)
    TABLESPACE zl9CisRec;   

Create Table ��������Ϣ(
       ID Number(18),
       �������� Varchar2(64),
       ������� Varchar2(20),
       ����ID Number(18),
       ��鷶Χ Varchar2(64),
       ��ʼ���� date,
       �������� date,
       ������ Varchar2(64),
       �������� date,
       ����˵�� Varchar2(1024),
       �������� Varchar2(32),
       ������� Varchar2(32),       
       �������� Varchar2(32),
       ��ϸ��ַ Varchar2(128),
       ����״̬ number(1) default 0,
       �鵵ʱ�� Date
  )TABLESPACE zl9CisRec;

Create Table ����鵵��Ϣ(
    ID Number(18), 
    ������Դ Number(1) default 0,
    ����ҽ��ID Number(18),
    �Ŀ�ID   Number(18),
    ��ƬID   Number(18),
    �ؼ�ID   Number(18),
    ����ID   Number(18),
    ���״̬ Number(1)  default 0,
    ����״̬ Number(1)  default 0
  )TABLESPACE zl9CisRec; 
  
Create Table ���������Ϣ(
    ID Number(18), 
    ������ Varchar2(64),
    ����ʱ�� Date,
    ֤������ Number(1) default 0,
    ֤������ Varchar2(20),
    ��ϵ�绰 Varchar2(20),
    ��ϵ��ַ Varchar2(128),
    Ѻ�� Number(16, 5),
    �������� Number(1),
    �������� Number(5),
    ����ԭ�� Varchar2(1024),
    �Ǽ��� Varchar2(64) ,
    �黹״̬ Number(1)  default 0,
    ȷ��״̬ Number(1)  default 0,
    ��ע Varchar2(1024)
  )TABLESPACE zl9CisRec;       

Create Table ������ʧ��Ϣ(
       ID Number(18),
       ����ID number(18),
       �鵵ID Number(18),
       ��ʧ���� Number(18),
       ��ʧԭ�� Varchar2(1024),
       ��ʧ���� date,
       �Ǽ���  Varchar2(64),
       ��ע Varchar2(1024)
  )TABLESPACE zl9CisRec;

Create Table ����黹��Ϣ(
       ID Number(18),
       ����ID Number(18),
       �黹�� Varchar2(64),
       �黹���� date,
       �˻�Ѻ�� Number(16,5),
       ����ҽԺ Varchar2(64),
       ����ҽʦ Varchar2(64),
       ������� Varchar2(2048),
       �Ǽ���  Varchar2(64),
       ��ע Varchar2(1024)
  )TABLESPACE zl9CisRec;

Create Table ������Ĺ���(
       ����ID Number(18),
       �鵵ID Number(18),
       �������� Number(2),
       �黹���� Number(2),
       �黹״̬ Number(1) default 0
  )TABLESPACE zl9CisRec;  

Create Table ����Ƭ��Ϣ(
       Id Number(18),
       ��ԴID Number(18),
       ��Դ���� Number(1),
       �Ŀ�ID Number(18),
       ����ҽ��ID Number(18),
       ����� Varchar2(30),
       �鵵״̬ Number(1),
       ��Ƭ���� Varchar2(10),
       ������ Varchar2(30),
       �������� date,
       ��ע   varchar2(512)
)TABLESPACE zl9CisRec;   
  
