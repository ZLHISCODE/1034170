------------------------------------------------------------
--数据组成分析：
------------------------------------------------------------
--  1.  咨询图片元素
------------------------------------------------------------
--  1.  咨询图片元素
Insert into 咨询图片元素( 序号,性质,名称,类型,图形,宽度,高度,修改日期,固定) 
Select 124,5,'当前选中的状态',0,NULL,120,50,trunc(sysdate)-28,0 From Dual Union All
Select 125,5,'返回',0,NULL,91,49,trunc(sysdate)-28,0 From Dual Union All
Select 126,5,'非当前选择的状态',0,NULL,120,50,trunc(sysdate)-28,0 From Dual Union All
Select 127,5,'挂号',0,NULL,91,49,trunc(sysdate)-28,0 From Dual Union All
Select 128,5,'关闭',0,NULL,91,49,trunc(sysdate)-28,0 From Dual Union All
Select 129,5,'上翻',0,NULL,80,43,trunc(sysdate)-28,0 From Dual Union All
Select 130,5,'我要挂号',0,NULL,128,50,trunc(sysdate)-28,0 From Dual Union All
Select 131,5,'下翻',0,NULL,80,43,trunc(sysdate)-28,0 From Dual Union All
Select 132,5,'专家介绍',0,NULL,130,49,trunc(sysdate)-28,0 From Dual Union All
Select 82,3,'项目2',1,NULL,16,16,trunc(sysdate)-30,0 From Dual Union All
Select 81,3,'项目1',1,NULL,16,16,trunc(sysdate)-30,0 From Dual Union All
Select 83,3,'项目3',1,NULL,16,16,trunc(sysdate)-30,0 From Dual Union All
Select 85,3,'项目5',1,NULL,16,16,trunc(sysdate)-30,0 From Dual Union All
Select 84,3,'项目4',1,NULL,16,16,trunc(sysdate)-30,0 From Dual Union All
Select 70,2,'健康美丽',2,NULL,120,63,trunc(sysdate)-30,0 From Dual Union All
Select 69,2,'保护环境',2,NULL,120,63,trunc(sysdate)-30,0 From Dual Union All
Select 71,2,'请勿吸烟',2,NULL,120,63,trunc(sysdate)-30,0 From Dual Union All
Select 72,2,'仁爱之心施天下2',2,NULL,120,63,trunc(sysdate)-30,0 From Dual Union All
Select 73,2,'珍爱生命',2,NULL,120,63,trunc(sysdate)-30,0 From Dual Union All
Select 74,2,'中联软件',2,NULL,120,64,trunc(sysdate)-30,0 From Dual Union All
Select 75,1,'诚实敬业',2,NULL,667,63,trunc(sysdate)-30,0 From Dual Union All
Select 77,1,'让你放心',2,NULL,667,63,trunc(sysdate)-30,0 From Dual Union All
Select 76,1,'精心的医护',2,NULL,677,63,trunc(sysdate)-30,0 From Dual Union All
Select 90,4,'背景4',0,NULL,60,60,trunc(sysdate)-30,0 From Dual Union All
Select 89,4,'背景3',0,NULL,420,350,trunc(sysdate)-30,0 From Dual Union All
Select 91,4,'背景5',0,NULL,332,148,trunc(sysdate)-30,0 From Dual Union All
Select 92,4,'花儿1',0,NULL,242,157,trunc(sysdate)-30,0 From Dual Union All
Select 87,4,'背景1',0,NULL,64,64,trunc(sysdate)-30,0 From Dual Union All
Select 86,3,'项目6',1,NULL,32,32,trunc(sysdate)-30,0 From Dual Union All
Select 64,9,'integris',0,NULL,254,200,trunc(sysdate)-30,0 From Dual Union All
Select 79,1,'一流技术2',2,NULL,677,63,trunc(sysdate)-30,0 From Dual Union All
Select 59,9,'medic3',0,NULL,116,82,trunc(sysdate)-30,0 From Dual Union All
Select 57,9,'medic1',0,NULL,90,90,trunc(sysdate)-30,0 From Dual Union All
Select 55,9,'baby',0,NULL,592,343,trunc(sysdate)-30,0 From Dual Union All
Select 56,9,'medic',0,NULL,72,108,trunc(sysdate)-30,0 From Dual Union All
Select 54,9,'baby1',0,NULL,81,92,trunc(sysdate)-30,0 From Dual Union All
Select 51,0,'whole',0,NULL,400,263,trunc(sysdate)-30,0 From Dual Union All
Select 50,9,'second',0,NULL,482,585,trunc(sysdate)-30,0 From Dual Union All
Select 49,9,'scan',0,NULL,200,200,trunc(sysdate)-30,0 From Dual Union All
Select 48,9,'first',0,NULL,484,319,trunc(sysdate)-30,0 From Dual Union All
Select 47,9,'dual',0,NULL,299,200,trunc(sysdate)-30,0 From Dual Union All
Select 46,9,'asu',0,NULL,200,200,trunc(sysdate)-30,0 From Dual Union All
Select 45,9,'alcyon',0,NULL,200,214,trunc(sysdate)-30,0 From Dual Union All
Select 42,9,'special-1',0,NULL,208,145,trunc(sysdate)-30,0 From Dual Union All
Select 43,9,'time',0,NULL,103,103,trunc(sysdate)-30,0 From Dual Union All
Select 44,9,'flower',0,NULL,72,101,trunc(sysdate)-30,0 From Dual Union All
Select 80,1,'早日健康—诚实友爱',2,NULL,667,63,trunc(sysdate)-30,0 From Dual Union All
Select 19,9,'food1',0,NULL,145,180,trunc(sysdate)-30,0 From Dual Union All
Select 40,9,'food',0,NULL,132,95,trunc(sysdate)-30,0 From Dual Union All
Select 78,1,'先进的设备',2,NULL,667,64,trunc(sysdate)-30,0 From Dual Union All
Select 38,9,'doctor',0,NULL,95,99,trunc(sysdate)-30,0 From Dual Union All
Select 35,9,'activi1',0,NULL,136,128,trunc(sysdate)-30,0 From Dual Union All
Select 37,9,'doctor1',0,NULL,84,92,trunc(sysdate)-30,0 From Dual Union All
Select 36,9,'check',0,NULL,100,120,trunc(sysdate)-30,0 From Dual Union All
Select 32,4,'flower',0,NULL,72,101,trunc(sysdate)-30,0 From Dual Union All
Select 93,4,'花儿2',0,NULL,255,157,trunc(sysdate)-30,0 From Dual Union All
Select 95,9,'优美宁静的环境',0,NULL,667,63,trunc(sysdate)-30,0 From Dual Union All
Select 97,9,'宣传1',0,NULL,667,64,trunc(sysdate)-30,0 From Dual Union All
Select 96,9,'春季保养',0,NULL,466,40,trunc(sysdate)-30,0 From Dual Union All
Select 98,9,'宣传2',0,NULL,463,60,trunc(sysdate)-30,0 From Dual Union All
Select 109,4,'背景图片1',0,NULL,640,480,trunc(sysdate)-30,0 From Dual Union All
Select 101,9,'优质服务',0,NULL,667,63,trunc(sysdate)-30,0 From Dual Union All
Select 102,4,'早日康复',0,NULL,667,507,trunc(sysdate)-30,0 From Dual Union All
Select 104,9,'cup',0,NULL,70,101,trunc(sysdate)-30,0 From Dual Union All
Select 120,4,'就医指南',0,NULL,660,510,trunc(sysdate)-30,0 From Dual Union All
Select 112,4,'背景图片3',0,NULL,640,500,trunc(sysdate)-30,0 From Dual Union All
Select 113,4,'背景图片4',0,NULL,660,500,trunc(sysdate)-30,0 From Dual Union All
Select 114,4,'背景图片5',0,NULL,660,500,trunc(sysdate)-30,0 From Dual Union All
Select 119,4,'主页背景',0,NULL,660,510,trunc(sysdate)-30,0 From Dual;