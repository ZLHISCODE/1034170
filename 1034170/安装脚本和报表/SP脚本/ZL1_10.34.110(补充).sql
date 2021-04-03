--[连续升级]1
--[管理工具版本号]10.34.110
--本脚本支持从ZLHIS+ v10.34.100 升级到 v10.34.110
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--109686:刘鹏飞,2017-05-31,主键或唯一键缺失补充
Alter Table 护理波动项目 Add Constraint 护理波动项目_PK Primary Key (项目序号) Using Index Tablespace zl9Indexcis;

Alter Table 护理适用科室 Add Constraint 护理适用科室_PK Primary Key (项目序号,科室ID) Using Index Tablespace zl9Indexcis;

--109412:张德婷,2017-05-23,添加静配相关表的主键约束
alter table 输液不配置药品 add constraint 输液不配置药品_PK primary key(药品id) Using Index Tablespace zl9Indexhis;

alter table 输液配药类型 add constraint 输液配药类型_PK primary key(编码) Using Index Tablespace zl9Indexhis;

alter table 输液药品优先级 add constraint 输液药品优先级_PK primary key(科室id,配药类型,频次) Using Index Tablespace zl9Indexhis;

alter table 输液优先打印药品 add constraint 输液优先打印药品_PK primary key(药品id) Using Index Tablespace zl9Indexhis;

alter table 配置收费方案 drop constraint 配置收费方案_UQ_配药类型 cascade drop index; 

alter table 配置收费方案 add constraint 配置收费方案_PK primary key(配药类型,项目id) Using Index Tablespace zl9Indexhis;

--109164:李南春,2017-05-23,增加病人免疫记录主键和外键
declare 
  n_count Number;   
begin   
  --增加病人免疫记录的主键前要处理可能存在的重复记录
  n_count := 0;
  For C_免疫 in (Select 病人ID,接种时间 From 病人免疫记录 group by 病人ID,接种时间 having count(1) > 1) Loop
      Update 病人免疫记录 set 接种时间 = 接种时间 + RowNum * 1/24/60/60  Where 病人ID = C_免疫.病人ID And 接种时间 = C_免疫.接种时间;
      n_count := 1;
  end Loop;

  If n_count = 1 Then 
     Commit;
  end if;
end;
/

Alter Table 病人免疫记录 add Constraint 病人免疫记录_PK Primary Key (病人ID,接种时间) Using Index Tablespace zl9indexhis;

--103974:李南春,2017-05-19,自动记账从病案主页获取病人基本信息
Create Or Replace View 在院病人自动记帐 As
Select p.病人id, p.主页id, Nvl(A.姓名,I.姓名) as 姓名, Nvl(A.性别,I.性别) as 性别, Nvl(A.年龄,I.年龄) as 年龄, i.住院号, a.费别, p.科室id, p.病区id, p.床号, p.附加床位, p.收费细目id, p.收入项目id, 1 As 标志,
       p.现价 As 标准单价, p.开始日期, p.终止日期, p.终止日期 - p.开始日期 As 天数, p.数量, p.经治医师, p.责任护士, p.操作员编号, p.操作员姓名
From 病人信息 I, 病案主页 A,
     (Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量
       From 自动计价项目 A,
            (Select a.病人id, a.主页id, a.开始时间, a.附加床位, a.病区id, a.科室id, a.床号, a.床位等级id, 1 As 数量, a.责任护士, a.经治医师, a.终止时间,
                     a.操作员编号, a.操作员姓名, a.上次计算时间
              From 病人变动记录 A, 病人信息 B
              Where a.开始原因 <> 10 And a.病人id = b.病人id And a.主页id = b.主页id And b.在院 = 1
              Union All
              Select b.病人id, b.主页id, 开始时间, 附加床位, b.病区id, b.科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号,
                     操作员姓名, 上次计算时间
              From 病人变动记录 B, 收费从属项目 I, 病人信息 C
              Where b.病人id = c.病人id And b.主页id = c.主页id And c.在院 = 1 And b.床位等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B,
            收费价目 P
       Where a.病区id = b.病区id And Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.计算标志 = 1 And b.床位等级id = p.收费细目id And Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))
       Union All
       Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量
       From 自动计价项目 A,
            (Select a.病人id, a.主页id, 开始时间, 附加床位, a.病区id, a.科室id, 床号, 护理等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间
              From 病人变动记录 A, 病人信息 B
              Where 开始原因 <> 10 And a.病人id = b.病人id And a.主页id = b.主页id And b.在院 = 1
              Union All
              Select b.病人id, b.主页id, 开始时间, 附加床位, b.病区id, b.科室id, 床号, i.从项id As 护理等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号,
                     操作员姓名, 上次计算时间
              From 病人变动记录 B, 收费从属项目 I, 病人信息 C
              Where b.护理等级id = i.主项id And b.病人id = c.病人id And b.主页id = c.主页id And c.在院 = 1 And b.开始原因 <> 10 And i.固有从属 > 0) B,
            收费价目 P, 收费项目目录 C
       Where a.病区id = b.病区id And b.附加床位 <> 1 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And a.计算标志 = 2 And
             b.护理等级id = p.收费细目id And b.护理等级id = c.Id And Nvl(c.计算方式, 0) <> 1 And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))
       Union All
       Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, a.数量
       From (Select 病区id, 计算标志, 收费细目id, 1 As 数量, 启用日期
              From 自动计价项目
              Union All
              Select 病区id, 计算标志, 从项id, i.从项数次 As 数量, 启用日期
              From 自动计价项目 A, 收费从属项目 I
              Where a.收费细目id = i.主项id And i.固有从属 > 0) A, 病人变动记录 B, 收费价目 P, 病人信息 C
       Where a.病区id = b.病区id And b.病人id = c.病人id And b.主页id = c.主页id And c.在院 = 1 And b.附加床位 <> 1 And b.开始原因 <> 10 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.收费细目id = p.收费细目id And (a.计算标志 = 6 And b.床位等级id Is Not Null Or a.计算标志 = 7) And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))) P
Where i.病人id = p.病人id And a.病人id = p.病人id And a.主页id = p.主页id;


Create Or Replace View 出院病人自动记帐 As
Select p.病人id, p.主页id, Nvl(A.姓名,I.姓名) as 姓名, Nvl(A.性别,I.性别) as 性别, Nvl(A.年龄,I.年龄) as 年龄, i.住院号, a.费别, p.科室id, p.病区id, p.床号, p.附加床位, p.收费细目id, p.收入项目id, 1 As 标志,
       p.现价 As 标准单价, p.开始日期, p.终止日期, p.终止日期 - p.开始日期 As 天数, p.数量, p.经治医师, p.责任护士, p.操作员编号, p.操作员姓名
From 病人信息 I, 病案主页 A,
     (Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量
       From 自动计价项目 A,
            (Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, 床位等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间
              From 病人变动记录 A
              Where 开始原因 <> 10
              Union All
              Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, i.从项id As 床位等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名,
                     上次计算时间
              From 病人变动记录 B, 收费从属项目 I
              Where b.床位等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B, 收费价目 P
       Where a.病区id = b.病区id And Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.计算标志 = 1 And b.床位等级id = p.收费细目id And Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))
       Union All
       Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, b.数量
       From 自动计价项目 A,
            (Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, 护理等级id, 1 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名, 上次计算时间
              From 病人变动记录
              Where 开始原因 <> 10
              Union All
              Select 病人id, 主页id, 开始时间, 附加床位, 病区id, 科室id, 床号, i.从项id As 护理等级id, i.从项数次 As 数量, 责任护士, 经治医师, 终止时间, 操作员编号, 操作员姓名,
                     上次计算时间
              From 病人变动记录 B, 收费从属项目 I
              Where b.护理等级id = i.主项id And b.开始原因 <> 10 And i.固有从属 > 0) B, 收费价目 P, 收费项目目录 C
       Where a.病区id = b.病区id And b.附加床位 <> 1 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And a.计算标志 = 2 And
             b.护理等级id = p.收费细目id And b.护理等级id = c.Id And Nvl(c.计算方式, 0) <> 1 And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))
       Union All
       Select b.病人id, b.主页id, b.科室id, b.病区id, b.床号, b.附加床位, p.收费细目id, p.收入项目id, p.现价, b.经治医师, b.责任护士, b.操作员编号, b.操作员姓名,
              Zl_Date_Half(Greatest(Least(Nvl(b.上次计算时间, b.开始时间), Nvl(b.终止时间, Greatest(Nvl(b.上次计算时间, b.开始时间))),
                                           Greatest(Nvl(b.上次计算时间, b.开始时间))), p.执行日期, Nvl(a.启用日期, Add_Months(Sysdate, -2)))) As 开始日期,
              Zl_Date_Half(Least(Nvl(b.终止时间, Greatest(b.开始时间, Sysdate)), Nvl(p.终止日期, Sysdate + 30) + 1)) As 终止日期, a.数量
       From (Select 病区id, 计算标志, 收费细目id, 1 As 数量, 启用日期
              From 自动计价项目
              Union All
              Select 病区id, 计算标志, 从项id, i.从项数次 As 数量, 启用日期
              From 自动计价项目 A, 收费从属项目 I
              Where a.收费细目id = i.主项id And i.固有从属 > 0) A, 病人变动记录 B, 收费价目 P
       Where a.病区id = b.病区id And b.附加床位 <> 1 And b.开始原因 <> 10 And
             Zl_Date_Half(Nvl(b.上次计算时间, b.开始时间)) <> Zl_Date_Half(Nvl(b.终止时间, Sysdate)) And p.现价 <> 0 And
             a.收费细目id = p.收费细目id And (a.计算标志 = 6 And b.床位等级id Is Not Null Or a.计算标志 = 7) And
             Zl_Date_Half(Nvl(b.终止时间, Sysdate)) >= Zl_Date_Half(p.执行日期) And
             Zl_Date_Half(b.开始时间) <= Zl_Date_Half(Nvl(p.终止日期, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.终止时间, Sysdate), Nvl(p.终止日期, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.启用日期, Add_Months(Sysdate, -2)))) P
Where i.病人id = p.病人id And a.病人id = p.病人id And a.主页id = p.主页id;

--109168:谢荣,2017-05-16,医保病人档案和医保病人关联表的主键顺序不一致。
Alter Table 医保病人档案 Drop Constraint 医保病人档案_PK Cascade Drop Index;

Alter Table 医保病人档案 Add Constraint 医保病人档案_PK Primary Key (医保号,险类,中心) Using Index Tablespace zl9Indexhis;

--108762:冉俊明,2017-05-15,临床出诊安排医生姓名前显示职称标识符
Alter Table 专业技术职务 Add 标识符 Varchar2(5);

--109137:胡俊勇,2017-05-15,基础表主键缺失
Alter Table 医嘱常用原因 Add Constraint 医嘱常用原因_PK Primary Key (编码) Using Index Tablespace zl9Indexhis;

Alter Table 医嘱常用原因 Add Constraint 医嘱常用原因_UQ_人员 Unique (人员,名称,性质) Using Index Tablespace zl9Indexhis;

--105165:张德婷,2017-04-20,处方发药可以添加该叫号窗口的所有处方
alter table 发药窗口 add 叫号窗口 varchar2(10);

--107559:冉俊明,2017-04-17,增加终止停诊安排功能
Alter Table 临床出诊停诊记录 Add 失效时间 Date;

--105791:冉俊明,2017-04-05,法定假日表字段名调整
Alter Table 法定假日表 Rename Column 允许预约 To 允许预约日期;
Alter Table 法定假日表 Rename Column 允许挂号 To 允许挂号日期;

--109164:李南春,2017-05-23,增加病人免疫记录主键和外键
Alter Table 病人免疫记录 add Constraint 病人免疫记录_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID);

--109421:李业庆,2017-05-23,药品加成方案表增加主键
Alter Table 药品加成方案 Add Constraint 药品加成方案_PK Primary Key (序号) Using Index Tablespace zl9Indexhis;



-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--109289:冉俊明,2017-05-23,使用“全部启用序号控制”功能时，对于启用序号但没有启用分时段的安排，没有生成对应的时段序号数据
--基础数据，临床出诊安排，数据量不会太大
Begin
  --1.规则数据
  --不分时段的序号控制号先生成序号,开始时间、终止时间填写时间段的开始时间和结束时间
  For c_安排 In (Select a.Id, b.号类, Nvl(c.站点, '-') As 站点
               From 临床出诊安排 A, 临床出诊号源 B, 部门表 C, 临床出诊表 D
               Where a.号源id = b.Id And b.科室id = c.Id And a.出诊id = d.Id And Nvl(d.排班方式, 0) In (0, 3)) Loop
  
    For c_记录 In (With c_时间段 As
                    (Select 时间段, 开始时间, 终止时间
                    From (Select 时间段,
                                  To_Date('3000-01-01' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                  To_Date('3000-01-01' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 终止时间,
                                  Row_Number() Over(Partition By 时间段 Order By 时间段, 站点 Asc, 号类 Asc) As 组号
                           From 时间段
                           Where Nvl(站点, c_安排.站点) = c_安排.站点 And Nvl(号类, c_安排.号类) = c_安排.号类)
                    Where 组号 = 1)
                   Select a.Id, a.限号数,
                          To_Date(To_Char(Sysdate, 'yyyy-mm-dd ') || To_Char(b.开始时间, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                          To_Date(To_Char(Sysdate, 'yyyy-mm-dd ') || To_Char(b.终止时间, 'hh24:mi:ss'),
                                  'yyyy-mm-dd hh24:mi:ss') + Case
                            When b.终止时间 <= b.开始时间 Then
                             1
                            Else
                             0
                          End As 终止时间
                   From 临床出诊限制 A, c_时间段 B
                   Where a.上班时段 = b.时间段 And a.安排id = c_安排.Id And Nvl(a.限号数, 0) <> 0 And Nvl(a.是否序号控制, 0) = 1 And
                         Nvl(a.是否分时段, 0) = 0 And Not Exists (Select 1 From 临床出诊时段 Where 限制id = a.Id)) Loop
    
      For I In 1 .. c_记录.限号数 Loop
        Insert Into 临床出诊时段
          (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
        Values
          (c_记录.Id, I, c_记录.开始时间, c_记录.终止时间, 1, 1);
      End Loop;
    End Loop;
  End Loop;

  --2.记录数据
  --不分时段的序号控制号先生成序号,开始时间、终止时间填写时间段的开始时间和结束时间
  For c_记录 In (Select a.Id, a.限号数, a.开始时间, a.终止时间
               From 临床出诊记录 A
               Where a.出诊日期 > Trunc(Sysdate) And Nvl(a.已挂数, 0) = 0 And Nvl(a.已约数, 0) = 0 And Nvl(a.限号数, 0) <> 0 And
                     Nvl(a.是否序号控制, 0) = 1 And Nvl(a.是否分时段, 0) = 0 And Not Exists
                (Select 1 From 临床出诊序号控制 Where 记录id = a.Id)) Loop
  
    For I In 1 .. c_记录.限号数 Loop
      Insert Into 临床出诊序号控制
        (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约)
      Values
        (c_记录.Id, I, c_记录.开始时间, c_记录.终止时间, 1, 1);
    End Loop;
  End Loop;
End;
/

--106248:蒋廷中,2017-5-15,病理诊断检查ICD附码检查
Update zlParameters
Set 参数说明='主要出院诊断编码开头为C00到D48时,病理诊断：1-必须填写，2-提示是否填写，0-不检查。'
Where 参数名 = '病理诊断检查' And 模块 = 1261 And 系统 = &n_System;
Update zlParameters
Set 参数说明='主要出院诊断编码开头为C00到D48时，主要出院诊断的ICD附码：1-必须填写，2-提示是否填写，0-不检查。'
Where 参数名 = 'ICD附码检查' And 模块 = 1261 And 系统 = &n_System;

--108423:李小东,2017-05-11,检验标本登记，紧急标本提示
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1212, 0, 1, 0, 0, 6, '紧急标本提示', '0', '0', '检验标本登记，登记标本为紧急标本时是否提示'
  From Dual;

--106148:刘尔旋,2017-04-17,新版提前挂号颜色处理
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1111, 0, 0, 0, 0, 72, '提前挂号颜色', Null, '0',
         '新版挂号时，提前挂号安排的字体颜色显示。'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 系统 = &n_System And 模块 = 1111 And 参数名 = '提前挂号颜色');

--89759:李南春,2017-04-14,消费卡刷卡是否定位到密码框
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, -Null, -Null, -Null, -Null, -Null, 276, '消费卡刷卡消费须定位到密码框', '1', '1',
         '如果启用参数，消费卡刷卡消费的时候，没有卡密码的情况下，光标也定位到密码文本框。'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 参数号 = 276 And Nvl(模块, 0) = 0 And Nvl(系统, 0) = &n_System);

--105443:刘尔旋,2017-04-06,预约有效时间规则变更
Update zlParameters
Set 参数说明 = '表示预约与实际预约接收时间的有效范围,以分钟为单位,0表示不限制,>0表示提前接收的限制分钟数,<0表示延后接收的限制分钟数'
Where 系统 = &n_System And 模块 = 1111 And 参数名 = '预约有效时间';

--104983:冉俊明,2017-04-05,体检病人按单据分别打印
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1121, 0, 0, 0, 0, 113, '体检病人分单据打印', Null, '0',
         '在病人收费管理中，如果票据是按“根据实际打印分配票号”，且设置了“门诊收费每张单据分别打印”时，体检病人是否按每张收费单进行分别打印发票。1-体检病人按每张单据分别打印，0或NULL-体检病人不按每张单据分别打印'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where 系统 = &n_System And 模块 = 1121 And 参数名 = '体检病人分单据打印');

--107799:胡俊勇,2017-04-05,医嘱清单过滤条件个性化参数记录
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1252, 1, 1, 0, 1, 24, '医嘱状态过滤', Null, '0',
         '门诊医嘱清单页签选择记录:0-医嘱,3-报告'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where 参数名 = '医嘱过滤方式' And Nvl(模块, 0) = 1252 And Nvl(系统, 0) = &n_System);

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1252, 1, 1, 0, 1, 59, '报告查看类型', Null, '0',
         '门诊医嘱清单选择报告页签时的过滤条件记录:0-全部,1-检查,2-检验,3-其他'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where 参数名 = '报告查看类型' And Nvl(模块, 0) = 1252 And Nvl(系统, 0) = &n_System);

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1252, 1, 1, 0, 1, 60, '过滤条件自动隐藏', Null, '0',
         '门诊医嘱清单中的过滤条件工具栏是否自动隐藏:0-不隐藏,1-隐藏'
  From Dual
  Where Not Exists (Select 1
         From zlParameters
         Where 参数名 = '过滤条件自动隐藏' And Nvl(模块, 0) = 1252 And Nvl(系统, 0) = &n_System);

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1253, 1, 1, 0, 1, 59, '报告查看类型', Null, '0',
         '住院医嘱清单选择报告页签时的过滤条件记录:0-全部,1-检查,2-检验,3-其他'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where 参数名 = '报告查看类型' And Nvl(模块, 0) = 1253 And Nvl(系统, 0) = &n_System);

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1253, 1, 1, 0, 1, 60, '过滤条件自动隐藏', Null, '0',
         '住院医嘱嘱清单中的过滤条件工具栏是否自动隐藏:0-不隐藏,1-隐藏'
  From Dual
  Where Not Exists (Select 1
         From zlParameters
         Where 参数名 = '过滤条件自动隐藏' And Nvl(模块, 0) = 1253 And Nvl(系统, 0) = &n_System);

Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1253, 1, 1, 0, 1, 61, '医嘱显示在用', Null, '0', '住院医嘱清单中的过滤条件:0-在用医嘱,1-所有医嘱'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where 参数名 = '医嘱显示在用' And Nvl(模块, 0) = 1253 And Nvl(系统, 0) = &n_System);


--107566:黄捷,2017-04-05,图像四角信息增加医嘱ID
Insert Into 影像图像信息表
  (Id, 开始地址, 结束地址, 英文名称, 中文名称, 中文简称, 英文简称, 常用, 被选用, 位置, 角内序号, 使用计算)
Values
  (影像图像信息表_Id.Nextval, 3, 3, 'cal', 'DB医嘱ID', '[医嘱ID]', '[OrderID]', -1, 0, 0, 0, 0);


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--109070:张险华,2017-06-02,增加范文审核模块
Insert Into zlPrograms
  (序号, 标题, 说明, 系统, 部件)
  Select 2228 序号, '范文审核' 标题, '用于对范文进行审核操作' As 说明, &n_System 系统, 'zl9EmrInterface' 部件
  From Dual
  Where Not Exists (Select 1 From zlPrograms Where 序号 = 2228 And 标题='范文审核');

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,2228,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '基本',0,'基本信息',1 From Dual Union All
Select '科内审核',1,'对操作人员所属科室和操作本人私有范文的审核权限',1 From Dual Union All
Select '全院审核',2,'对所有范文的审核权限',0 From Dual Union All
Select '范文修改',3,'对范文标题,名称,说明,所属原型,适用范文的修改权限',0 From Dual Union All
Select '内容编辑',4,'对范文内容进行修改的权限',0 From Dual) A
Where Not Exists (Select 1 From zlProgFuncs Where 序号 = 2228 And 功能='基本');

Insert Into zlMenus(组别, ID, 上级id, 标题, 说明, 系统, 模块, 短标题, 图标)
Select A.组别,ZlMenus_ID.Nextval,A.ID,B.* From (
Select 组别,ID From zlMenus Where 标题 = '病历文档基础' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null) A,
(Select 标题, 说明, 系统, 模块, 短标题, 图标 From zlMenus Where 1=0 Union ALL
Select '范文审核','用于对范文进行审核操作',&n_System,2228,'范文审核',114 From Dual) B
Where Not Exists (Select 1 From zlMenus Where 模块 = 2228 And 标题='范文审核');

--100722:张德婷,2017-06-01,修正窗口改变不能读取数据的问题
Insert Into zlProgPrivs(系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1340, '增删改', User, 'zl_发药窗口_业务调整', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1340 And 功能 = '增删改' And Upper(对象) = Upper('zl_发药窗口_业务调整'));

--106745:刘尔旋,2017-05-31,挂号检查封装
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1115, '预约挂号登记', User, 'Zl_Fun_病人挂号记录_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1115 And 功能 = '预约挂号登记' And Upper(对象) = Upper('Zl_Fun_病人挂号记录_Check'));  

--106745:刘尔旋,2017-05-31,挂号封装检查
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1260, '病人挂号', User, 'Zl_Fun_病人挂号记录_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1260 And 功能 = '病人挂号' And Upper(对象) = Upper('Zl_Fun_病人挂号记录_Check'));  
         
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1260, '预约挂号', User, 'Zl_Fun_病人挂号记录_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1260 And 功能 = '预约挂号' And Upper(对象) = Upper('Zl_Fun_病人挂号记录_Check'));  

--106745:刘尔旋,2017-05-31,挂号封装检查
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1111, '挂收费号', User, 'Zl_Fun_病人挂号记录_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1111 And 功能 = '挂收费号' And Upper(对象) = Upper('Zl_Fun_病人挂号记录_Check'));  

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1111, '挂免费号', User, 'Zl_Fun_病人挂号记录_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1111 And 功能 = '挂免费号' And Upper(对象) = Upper('Zl_Fun_病人挂号记录_Check'));  

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1111, '预约挂号', User, 'Zl_Fun_病人挂号记录_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1111 And 功能 = '预约挂号' And Upper(对象) = Upper('Zl_Fun_病人挂号记录_Check'));  

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1111, '接收预约', User, 'Zl_Fun_病人挂号记录_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1111 And 功能 = '接收预约' And Upper(对象) = Upper('Zl_Fun_病人挂号记录_Check'));  

--108762:冉俊明,2017-05-15,临床出诊安排医生姓名前显示职称标识符
Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明, 缺省值)
  Select &n_System, 1114, '职称标识设置', 24, '具有该权限时，可以对显示在临床安排中医生姓名前的医生职称标识符进行设置。', 0
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where 系统 = &n_System And 序号 = 1114 And 功能 = '职称标识设置');

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1114, '基本', User, '专业技术职务', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1114 And 功能 = '基本' And 对象 = '专业技术职务');

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1114, '职称标识设置', User, 'Zl_专业技术职务_更新标识符', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1114 And 功能 = '职称标识设置' And Upper(对象) = Upper('Zl_专业技术职务_更新标识符'));

--108534:张险华,2017-5-11,增加取消完成审批模块
Insert Into zlPrograms
  (序号, 标题, 说明, 系统, 部件)
  Select 2227 序号, '取消完成审批' 标题, '用于在病历完成后需要再次修改时进行审批操作' As 说明, &n_System 系统, 'zl9EmrInterface' 部件
  From Dual
  Where Not Exists (Select 1 From zlPrograms Where 序号 = 2227 And 标题='取消完成审批');

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,2227,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '基本',0,'基本信息',1 From Dual Union All
Select '系统报表',1,'对系统自带报表的访问',1 From Dual) A
Where Not Exists (Select 1 From zlProgFuncs Where 序号 = 2227 And 功能='基本');

Insert Into zlMenus(组别, ID, 上级id, 标题, 说明, 系统, 模块, 短标题, 图标)
Select A.组别,ZlMenus_ID.Nextval,A.ID,B.* From (
Select 组别,ID From zlMenus Where 标题 = '质控系统管理' And 组别 = '缺省' And 系统 = &n_System And 模块 Is Null) A,
(Select 标题, 说明, 系统, 模块, 短标题, 图标 From zlMenus Where 1=0 Union ALL
Select '取消完成审批','用于在病历完成后需要再次修改时进行审批操作',&n_System,2227,'取消完成审批',109 From Dual) B
Where Not Exists (Select 1 From zlMenus Where 模块 = 2227 And 标题='取消完成审批');

--100642:董露露,2017-05-11,处理病人信息管理中增、删、改权限独立，可以分开进行授权的需求
Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明, 缺省值)
  Select &n_System, 1101, '增加', 22, '增加病人信息的操作权限。有该权限时，允许对病人信息进行增加操作。', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where 系统 = &n_System And 序号 = 1101 And 功能 = '增加');

Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明, 缺省值)
  Select &n_System, 1101, '修改', 23, '修改病人信息的操作权限。有该权限时，允许对病人信息进行修改操作。', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where 系统 = &n_System And 序号 = 1101 And 功能 = '修改');

Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明, 缺省值)
  Select &n_System, 1101, '删除', 24, '删除病人信息的操作权限。有该权限时，允许对病人信息进行修改操作。', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where 系统 = &n_System And 序号 = 1101 And 功能 = '删除');

Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明, 缺省值)
  Select &n_System, 1101, '启停', 25, '启用和停用病人信息的操作权限。有该权限时，允许对病人信息进行停用、取消停用操作。', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where 系统 = &n_System And 序号 = 1101 And 功能 = '启停');

Insert Into Zlprogrelas
  (系统, 序号, 组号, 功能, 关系, 主项, 主项关系)
  Select &n_System, 1101, 1, '增加', 2, 0, 0 From Dual
  Where Not Exists (Select 1 From Zlprogrelas Where 系统 = &n_System And 序号 = 1101 and 组号=1 And 功能 = '增加');
  
Insert Into Zlprogrelas
  (系统, 序号, 组号, 功能, 关系, 主项, 主项关系)
  Select &n_System, 1101, 1, '修改', 2, 0, 0 From Dual
   Where Not Exists (Select 1 From Zlprogrelas Where 系统 = &n_System And 序号 = 1101 and 组号=1 And 功能 = '修改');
  
Insert Into Zlprogrelas
  (系统, 序号, 组号, 功能, 关系, 主项, 主项关系)
  Select &n_System, 1101, 1, '删除', 2, 0, 0 From Dual
   Where Not Exists (Select 1 From Zlprogrelas Where 系统 = &n_System And 序号 = 1101 and 组号=1 And 功能 = '删除');
  
Insert Into Zlprogrelas
  (系统, 序号, 组号, 功能, 关系, 主项, 主项关系)
  Select &n_System, 1101, 1, '启停', 2, 0, 0 From Dual
   Where Not Exists (Select 1 From Zlprogrelas Where 系统 = &n_System And 序号 = 1101 and 组号=1 And 功能 = '启停');

--107898:冉俊明,2017-05-08,修正临床出诊号源“适用年龄”设置问题，同步增加“出诊号源设置”对“性别”表的“SELECT”权限
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1114, '出诊号源设置', User, '性别', 'SELECT'
  From Dual
  Where Not Exists (Select 1 From zlProgPrivs Where 系统 = &n_System And 序号 = 1114 And 功能 = '出诊号源设置' And 对象 = '性别');

--108825:李南春,2017-05-05,自助票据打印授权
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1802, '基本', User, 'Zl_病人挂号票据_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1802 And 功能 = '基本' And Upper(对象) = Upper('Zl_病人挂号票据_Insert'));

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1803, '基本', User, 'Zl_病人挂号票据_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1803 And 功能 = '基本' And Upper(对象) = Upper('Zl_病人挂号票据_Insert'));

--98580:张德婷,2017-05-04,成套方案根据权限判断能否修改
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1054,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All 
    Select '修改个人成套方案',10,'有该权限时，操作员可以修改个人的成套方案。',1 From Dual Union All 
    Select '修改科室成套方案',11,'有该权限时，操作员可以修改科室的成套方案。',1 From Dual Union All
    Select '修改全院成套方案',12,'有该权限时，操作员可以修改全院的成套方案。',1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1009,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All 
    Select '修改个人成套方案',16,'有该权限时，操作员可以修改个人的成套方案。',1 From Dual Union All 
    Select '修改科室成套方案',17,'有该权限时，操作员可以修改科室的成套方案。',1 From Dual Union All
    Select '修改全院成套方案',18,'有该权限时，操作员可以修改全院的成套方案。',1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

--107559:冉俊明,2017-04-17,增加终止停诊安排功能
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1114, '停诊审批', User, 'Zl_临床出诊停诊_Stop', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1114 And 功能 = '停诊审批' And Upper(对象) = Upper('Zl_临床出诊停诊_Stop'));

--106708:冉俊明,2017-04-07,违反规范，调整
Delete From zlProgPrivs
Where 系统 = &n_System And 序号 = 1114 And 功能 = '出诊安排' And Upper(对象) = Upper('Zl_Buildregisterfixedrule');

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1114, '出诊安排', User, 'Zl_临床出诊表_Addbyfixedrule', 'EXECUTE'
  From Dual
  Where Not Exists
   (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1114 And 功能 = '出诊安排' And Upper(对象) = Upper('Zl_临床出诊表_Addbyfixedrule'));

Delete From zlProgPrivs
Where 系统 = &n_System And 序号 = 1114 And 功能 = '出诊安排' And Upper(对象) = Upper('Zl_Buildregisterplanbyrecord');

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1114, '出诊安排', User, 'Zl_临床出诊表_Addbyrecord', 'EXECUTE'
  From Dual
  Where Not Exists
   (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1114 And 功能 = '出诊安排' And Upper(对象) = Upper('Zl_临床出诊表_Addbyrecord'));

Delete From zlProgPrivs
Where 系统 = &n_System And 序号 = 1114 And 功能 = '出诊安排' And Upper(对象) = Upper('Zl_Buildregisterplanbytemplet');

Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1114, '出诊安排', User, 'Zl_临床出诊表_Addbytemplet', 'EXECUTE'
  From Dual
  Where Not Exists
   (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1114 And 功能 = '出诊安排' And Upper(对象) = Upper('Zl_临床出诊表_Addbytemplet'));

--105824:廖思奇,2017-04-05,在1290 1291 中增加 报告发放 功能包括Zl_影像报告发放 执行权限 
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1290,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All 
Select '报告发放',39,'诊断报告的发放',1 From Dual Union All 
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1291,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All 
Select '报告发放',36,'诊断报告的发放',1 From Dual Union All 
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select  &n_System,1290,'报告发放',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select 'Zl_影像报告发放','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select  &n_System,1291,'报告发放',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select 'Zl_影像报告发放','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--107370:李小东,2017-03-27,新版护士站-浏览检验结果，过滤该病区所有病人
Insert Into zlProgFuncs(系统,序号,功能,排列,说明,缺省值)
Select &n_System,1265,A.* From (
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0 Union All
Select '本科病人',9,'拥有该权限时浏览检验结果可过滤出该病区所有病人，无该权限时只能过滤出单个病人',1 From Dual Union All
Select 功能,排列,说明,缺省值 From zlProgFuncs Where 1 = 0) A;

--107370:李小东,2017-03-27,新版护士站-浏览检验结果，过滤该病区所有病人
Insert Into zlRoleGrant
  (系统, 序号, 角色, 功能)
  Select Distinct 系统, 序号, 角色, '本科病人' 功能
  From zlRoleGrant A
  Where 系统 = &n_System And 序号 = 1265 And 功能 = '基本';

--99878:冉俊明,2017-05-27,同一收费项目多个价格管理
Insert Into Zlmodulerelas
  (系统, 模块, 功能, 相关系统, 相关模块, 相关类型, 相关功能, 缺省值)
  Select &n_System, 1107, Null, &n_System, 9000, 1, '基本', 1 From Dual union all
  Select &n_System, 1111, Null, &n_System, 9000, 1, '基本', 1 From Dual union all
  Select &n_System, 1115, Null, &n_System, 9000, 1, '基本', 1 From Dual union all
  Select &n_System, 1120, Null, &n_System, 9000, 1, '基本', 1 From Dual union all
  Select &n_System, 1121, Null, &n_System, 9000, 1, '基本', 1 From Dual union all
  Select &n_System, 1122, Null, &n_System, 9000, 1, '基本', 1 From Dual union all
  Select &n_System, 1133, Null, &n_System, 9000, 1, '基本', 1 From Dual union all
  Select &n_System, 1134, Null, &n_System, 9000, 1, '基本', 1 From Dual union all
  Select &n_System, 1135, Null, &n_System, 9000, 1, '基本', 1 From Dual union all
  Select &n_System, 1139, Null, &n_System, 9000, 1, '基本', 1 From Dual;




-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------






-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--100722:张德婷,2017-06-01,修改发药窗口同步调整业务数据
Create Or Replace Procedure Zl_发药窗口_业务调整
(
  药房id_In In Number,
  旧窗口_In In Varchar2,
  新窗口_In In Varchar2
) Is

  Cursor c_未发数据 Is
    Select 单据, No, 库房id
    From 未发药品记录
    Where 填制日期 Between Sysdate - 3 And Sysdate And 发药窗口 = 旧窗口_In And 库房id = 药房id_In;

  --药房参数
  Cursor c_药房参数 Is
    Select a.参数值
    From (Select 机器名, 参数值 From Zluserparas Where 参数id = 1687) a,
         (Select 机器名, 参数值 From Zluserparas Where 参数id = 1688) b
    Where a.机器名 = b.机器名 And b.参数值 = 药房id_In;

  v_未发数据 c_未发数据%Rowtype;
  v_药房参数 c_药房参数%Rowtype;
Begin
  --费用参数
  Update Zluserparas
  Set 参数值 = 药房id_In || ':' || 新窗口_In
  Where 参数值 = 药房id_In || ':' || 旧窗口_In And
        参数id In (Select Id From Zlparameters Where 参数名 In ('西药房窗口', '中药房窗口', '成药房窗口'));

  --业务数据
  For v_未发数据 In c_未发数据 Loop
    Update 药品收发记录
    Set 发药窗口 = 新窗口_In
    Where 单据 = v_未发数据.单据 And No = v_未发数据.No And 库房id = v_未发数据.库房id And 发药窗口 = 旧窗口_In;
    Update 门诊费用记录
    Set 发药窗口 = 新窗口_In
    Where No = v_未发数据.No And 执行部门id = v_未发数据.库房id And 发药窗口 = 旧窗口_In;
    Update 住院费用记录
    Set 发药窗口 = 新窗口_In
    Where No = v_未发数据.No And 执行部门id = v_未发数据.库房id And 发药窗口 = 旧窗口_In;
  End Loop;

  Update 未发药品记录
  Set 发药窗口 = 新窗口_In
  Where 填制日期 Between Sysdate - 3 And Sysdate And 发药窗口 = 旧窗口_In And 库房id = 药房id_In;

  --药品参数
  Update Zluserparas
  Set 参数值 = Replace(参数值, 旧窗口_In, 新窗口_In)
  Where 参数id = 1687 And 机器名 In (Select 机器名 From Zluserparas Where 参数值 = 药房id_In And 参数id = 1688);

  --叫号窗口
  Update 发药窗口 Set 叫号窗口 = 新窗口_In Where 药房id = 药房id_In And 叫号窗口 = 旧窗口_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_发药窗口_业务调整;
/

--106745:刘尔旋,2017-06-01,挂号检查封装
Create Or Replace Function Zl_Fun_病人挂号记录_Check
(
  操作方式_In   Integer,
  病人id_In     门诊费用记录.病人id%Type,
  号码_In       挂号安排.号码%Type,
  出诊记录id_In 临床出诊记录.Id%Type := Null,
  发生时间_In   门诊费用记录.发生时间%Type,
  专家号_In     Number := 0
) Return Varchar2 As
  --功能：挂号有效性检查(包含预约;预约挂号不扣款;预约挂号扣款)
  --入参:操作方式_IN:0-挂号(包含收款预约),1-预约,2-预约接收
  --     是否加号_In:是否加号调用，0-非加号调用，1-加号调用
  --返回:0-检查通过
  --     1-特定检查项检查失败，同时返回错误提示文本
  --     2-其他错误导致的检查失败，同时返回错误提示文本
  Err_Item Exception;
  n_病人预约科室数 Number(18);
  n_已约科室       Number(18);
  v_Temp           Varchar2(500);
  v_加入原因       特殊病人.加入原因%Type;
  n_同科限号数     Number;
  n_同科限约数     Number;
  n_科室id         挂号安排.科室id%Type;
  n_Count          Number(18);
  n_病人挂号科室数 Number;
  n_专家号挂号限制 Number;
  n_专家号预约限制 Number;
  n_专家号         Number;
  d_生效时间       Date;
  n_计划id         挂号安排计划.Id%Type;

  Cursor c_Pati(n_病人id 病人信息.病人id%Type) Is
    Select a.病人id, a.姓名, a.性别, a.年龄, a.住院号, a.门诊号, a.费别, a.险类, c.编码 As 付款方式
    From 病人信息 A, 医疗付款方式 C
    Where a.病人id = n_病人id And a.医疗付款方式 = c.名称(+);

  r_Pati c_Pati%RowType;

  Function Zl_操作员
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
    -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
  Begin
    If Type_In = 0 Then
      --缺省部门
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --操作员编码
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --操作员姓名
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

Begin
  --检测病人相关
  Open c_Pati(病人id_In);
  n_Count := 0;
  Begin
    Fetch c_Pati
      Into r_Pati;
    n_Count := 1;
  Exception
    When Others Then
      n_Count := -1;
  End;
  If n_Count <= 0 Then
    Return '1|病人未找到，不能继续。';
  End If;
  --预约检测黑名单
  If 操作方式_In = 1 Then
    Begin
      Select 加入原因 Into v_加入原因 From 特殊病人 Where 撤消时间 Is Null And 病人id = 病人id_In And Rownum = 1;
      Return '1|此病人在特殊病人名单中，原因：【' || v_加入原因 || '】不能继续！';
    Exception
      When Others Then
        Null;
    End;
  End If;

  --检测挂号时间
  If Trunc(Sysdate) > Trunc(发生时间_In) Then
    Return '1|不能挂以前的号(' || To_Char(发生时间_In, 'yyyy-mm-dd') || ')。';
  End If;

  --部门ID,部门名称;人员ID,人员编号,人员姓名
  v_Temp := Zl_Identity(0);
  If Nvl(v_Temp, ' ') = ' ' Then
    Return '1|当前操作人员未设置对应的人员关系,不能继续。';
  End If;

  n_专家号 := 专家号_In;
  If 出诊记录id_In Is Null Then
    Select 科室id Into n_科室id From 挂号安排 Where 号码 = 号码_In;
  Else
    Select 科室id Into n_科室id From 临床出诊记录 Where ID = 出诊记录id_In;
  End If;

  --检测系统参数
  v_Temp           := Nvl(zl_GetSysParameter('病人同科限挂N个号', 1111), '0|0') || '|';
  n_同科限号数     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
  n_同科限约数     := To_Number(Nvl(zl_GetSysParameter('病人同科限约N个号', 1111), '0'));
  n_病人预约科室数 := To_Number(Nvl(zl_GetSysParameter('病人预约科室数', 1111), '0'));
  n_病人挂号科室数 := To_Number(Nvl(zl_GetSysParameter('病人挂号科室限制', 1111), '0'));
  n_专家号挂号限制 := To_Number(Nvl(zl_GetSysParameter('专家号挂号限制'), '0'));
  n_专家号预约限制 := To_Number(Nvl(zl_GetSysParameter('专家号预约限制'), '0'));
  --对参数控制进行检查
  If 操作方式_In = 1 Then
    If Nvl(n_同科限约数, 0) <> 0 Or Nvl(n_病人预约科室数, 0) <> 0 Then
      n_已约科室 := 0;
      For c_Chkitem In (Select Distinct 执行部门id
                        From 病人挂号记录
                        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
                              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
        n_已约科室 := n_已约科室 + 1;
      End Loop;
      If n_已约科室 >= Nvl(n_病人预约科室数, 0) And Nvl(n_病人预约科室数, 0) > 0 Then
        Return '1|同一病人最多同时能预约[' || Nvl(n_病人预约科室数, 0) || ']个科室,不能再预约！';
      End If;
    
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
            Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
      If n_Count >= Nvl(n_同科限约数, 0) And Nvl(n_同科限约数, 0) > 0 Then
        Return '1|该病人已经在该科室预约了' || n_Count || '次,不能再预约！';
      End If;
    End If;
    If Nvl(n_专家号预约限制, 0) <> 0 And n_专家号 = 1 Then
      If 出诊记录id_In Is Null Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 预约时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = 号码_In;
      Else
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 2 And 出诊记录id = 出诊记录id_In;
      End If;
      If n_Count >= Nvl(n_专家号预约限制, 0) And Nvl(n_专家号预约限制, 0) > 0 Then
        Return '1|该病人已经超过本号预约限制,不能再预约！';
      End If;
    End If;
  Else
    If (Nvl(n_同科限号数, 0) <> 0 Or Nvl(n_病人挂号科室数, 0) <> 0) And 操作方式_In = 0 Then
      n_已约科室 := 0;
      For c_Chkitem In (Select Distinct 执行部门id
                        From 病人挂号记录
                        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
                              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id <> n_科室id) Loop
        n_已约科室 := n_已约科室 + 1;
      End Loop;
      If n_已约科室 >= Nvl(n_病人挂号科室数, 0) And Nvl(n_病人挂号科室数, 0) > 0 Then
        Return '1|同一病人最多同时能挂号[' || Nvl(n_病人挂号科室数, 0) || ']个科室,不能再挂号！';
      End If;
    
      Select Count(1)
      Into n_Count
      From 病人挂号记录
      Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
            Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 执行部门id = n_科室id;
      If n_Count >= Nvl(n_同科限号数, 0) And Nvl(n_同科限号数, 0) > 0 Then
        Return '1|该病人已经在该科室挂号了' || n_Count || '次,不能再挂号！';
      End If;
    End If;
  
    If Nvl(n_专家号挂号限制, 0) <> 0 And n_专家号 = 1 Then
      If 出诊记录id_In Is Null Then
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 发生时间 Between Trunc(发生时间_In) And
              Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 号别 = 号码_In;
      Else
        Select Count(1)
        Into n_Count
        From 病人挂号记录
        Where 病人id = 病人id_In And 记录状态 = 1 And 记录性质 = 1 And 出诊记录id = 出诊记录id_In;
      End If;
      If n_Count >= Nvl(n_专家号挂号限制, 0) And Nvl(n_专家号挂号限制, 0) > 0 Then
        Return '1|该病人已经超过本号挂号限制,不能再挂号！';
      End If;
    End If;
  End If;

  Return '0-号源正常';

Exception
  When Others Then
    Return '2-' || SQLErrM;
End Zl_Fun_病人挂号记录_Check;
/

--109002:陈刘,2017-05-31,单一病人重算数据行

Create Or Replace Procedure Zl_病人护理打印_Update
(
  文件id_In   In 病人护理打印.文件id%Type,
  发生时间_In In 病人护理打印.发生时间%Type,
  行数_In     In 病人护理打印.行数%Type,
  删除_In     Number := 0,
  继续重算_In Number := 0
) Is
  n_Actives   Number;
  n_Rows      Number; --0-新增,>0表示修改 
  n_Startpage Number; --开始页 
  n_Startrow  Number; --开始行 
  n_Endpage   Number; --结束页 
  n_Endrow    Number; --结束行 
  n_Count     Number; --发生时间之后的数据条数 
  n_Pagerows  Number; --每页有效数据行 
  n_Del       Number;
  n_行数      病人护理打印.行数%Type;
  n_Firstdata Number; --是否是录入的第一条数据 
  n_记录id    病人护理数据.Id%Type;
  n_记录oldid 病人护理打印.记录id%Type;
  n_格式id    病人护理文件.格式id%Type;
  d_发生时间  病人护理打印.发生时间%Type;
  v_Username  人员表.姓名%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(1000);
  v_Print   Varchar2(800);
Begin
  n_Del      := 删除_In;
  n_行数     := 行数_In;
  v_Username := Zl_Username;
  Select 格式id Into n_格式id From 病人护理文件 Where ID = 文件id_In;

  If n_行数 = 0 Then
    v_Err_Msg := '有效数据行不能等于零，请记录本次错误的操作过程！';
    Raise Err_Item;
  End If;

  Begin
    Select 记录id, 行数, 开始页号, 开始行号, 结束页号, 结束行号
    Into n_记录oldid, n_Rows, n_Startpage, n_Startrow, n_Endpage, n_Endrow
    From 病人护理打印
    Where 文件id = 文件id_In And 发生时间 = 发生时间_In;
  Exception
    When Others Then
      n_Rows := 0;
  End;

  --提取该护理文件格式每页有效数据行（不加错误处理） 
  Select To_Number(内容文本)
  Into n_Pagerows
  From 病历文件结构
  Where 对象属性 = '有效数据行' And 父id = (Select ID From 病历文件结构 Where 文件id = n_格式id And 对象序号 = 1 And 父id Is Null);

  Select Count(*) Into n_Count From 病人护理数据 Where 文件id = 文件id_In And 发生时间 = 发生时间_In;

  --修改数据时,也可能删除 
  If n_Del = 0 Then
    Begin
      If n_Count = 0 Then
        n_Del := 1;
      End If;
      If n_Count > 1 Then
        v_Err_Msg := '在发生时间【' || To_Char(发生时间_In, 'YYYY-MM-DD hh24:mi:ss') || '】已经存在相应的数据，您不能再次录入或修改数据的时间为此发生时间！';
        Raise Err_Item;
      End If;
    End;
  Elsif n_Del = 1 And n_Count > 0 Then
    n_Del  := 0;
    n_行数 := 1;
  End If;

  n_Firstdata := 0;
  If n_Del = 1 Then
    Delete 病人护理打印 Where 文件id = 文件id_In And 发生时间 = 发生时间_In;
    n_Rows := n_Rows * -1;
  Else
    Select ID Into n_记录id From 病人护理数据 Where 文件id = 文件id_In And 发生时间 = 发生时间_In;
  
    If n_Rows = 0 Then
      --根据现有打印数据及将要插入的数据，计算出开始页号，行号，结束页号，行号 
      Select Min(发生时间) Into d_发生时间 From 病人护理打印 Where 文件id = 文件id_In And 发生时间 > 发生时间_In;
      If d_发生时间 Is Null Then
        Select Max(发生时间) Into d_发生时间 From 病人护理打印 Where 文件id = 文件id_In And 发生时间 < 发生时间_In;
        If d_发生时间 Is Null Then
          n_Startpage := 1;
          n_Startrow  := 1;
          n_Firstdata := 1;
        Else
          Select 结束页号, 结束行号
          Into n_Startpage, n_Startrow
          From 病人护理打印
          Where 文件id = 文件id_In And 发生时间 = d_发生时间;
          n_Startrow := n_Startrow + 1;
        End If;
      Else
        Select 开始页号, 开始行号
        Into n_Startpage, n_Startrow
        From 病人护理打印
        Where 文件id = 文件id_In And 发生时间 = d_发生时间;
      End If;
    
      --校正页号,行号 
      If n_Startrow > n_Pagerows Then
        n_Startpage := n_Startpage + 1;
        n_Startrow  := n_Startrow - n_Pagerows;
      
        --翻页时，自动依据当前页的设置产生新页的活动项目设置 
        Begin
          Select 1 Into n_Actives From 病人护理活动项目 Where 文件id = 文件id_In And 页号 = n_Startpage And Rownum < 2;
        Exception
          When Others Then
            n_Actives := 0;
        End;
      
        If n_Actives = 0 Then
          Insert Into 病人护理活动项目
            (文件id, 页号, 列号, 列头名称, 序号, 项目序号, 部位, 操作员, 操作时间)
            Select 文件id, n_Startpage, 列号, 列头名称, 序号, 项目序号, 部位, v_Username, Sysdate
            From 病人护理活动项目
            Where 文件id = 文件id_In And 页号 = n_Startpage - 1;
        End If;
      End If;
      n_Endpage := n_Startpage;
      n_Endrow  := n_Startrow + n_行数 - 1;
      If n_Endrow > n_Pagerows Then
        --不考虑输入的数据超过一页的情况 
        n_Endpage := n_Endpage + 1;
        n_Endrow  := n_Endrow - n_Pagerows;
      
        --翻页时，自动依据当前页的设置产生新页的活动项目设置 
        Begin
          Select 1 Into n_Actives From 病人护理活动项目 Where 文件id = 文件id_In And 页号 = n_Endpage And Rownum < 2;
        Exception
          When Others Then
            n_Actives := 0;
        End;
      
        If n_Actives = 0 Then
          Insert Into 病人护理活动项目
            (文件id, 页号, 列号, 列头名称, 序号, 项目序号, 部位, 操作员, 操作时间)
            Select 文件id, n_Endpage, 列号, 列头名称, 序号, 项目序号, 部位, v_Username, Sysdate
            From 病人护理活动项目
            Where 文件id = 文件id_In And 页号 = n_Endpage - 1;
        End If;
      End If;
      --不允许录入跨两页的数据 
      If n_Endrow > n_Pagerows Or n_Endpage - n_Startpage > 1 Then
        If 继续重算_In = 1 Then
          n_行数   := n_行数 - n_Endrow + n_Pagerows;
          n_Endrow := n_Pagerows;
        Else
          v_Err_Msg := '您在发生时间【' || To_Char(发生时间_In, 'YYYY-MM-DD hh24:mi:ss') || '】录入的数据存在错误，录入的数据不能连续跨一页以上！';
          Raise Err_Item;
        End If;
      End If;
    
      Insert Into 病人护理打印
        (记录id, 文件id, 发生时间, 行数, 开始页号, 开始行号, 结束页号, 结束行号)
      Values
        (n_记录id, 文件id_In, 发生时间_In, n_行数, n_Startpage, n_Startrow, n_Endpage, n_Endrow);
      --新插入的数据的行数就是差值 
      n_Rows := n_行数;
    Else
      --计算与原行数的差值 
      n_Rows := n_行数 - n_Rows;
      --校正页号,行号 
      n_Endrow := n_Endrow + n_Rows;
      If n_Endrow <= 0 Then
        n_Endrow  := n_Pagerows + n_Endrow;
        n_Endpage := n_Endpage - 1;
      End If;
      If n_Endrow > n_Pagerows Then
        --不考虑输入的数据超过一页的情况 
        n_Endpage := n_Endpage + 1;
        n_Endrow  := n_Endrow - n_Pagerows;
      End If;
    
      --不允许录入跨两页的数据 
      If n_Endrow > n_Pagerows Or n_Endpage - n_Startpage > 1 Then
        v_Err_Msg := '您在发生时间【' || To_Char(发生时间_In, 'YYYY-MM-DD hh24:mi:ss') || '】录入的数据存在错误，录入的数据不能连续跨一页以上！';
        Raise Err_Item;
      End If;
    
      --更新打印数据（当前数据的打印人与打印时间更新为NULL，其后数据不动） 
      Update 病人护理打印
      Set 文件id = 文件id_In, 记录id = n_记录id, 发生时间 = 发生时间_In, 行数 = n_行数, 开始页号 = n_Startpage, 开始行号 = n_Startrow,
          结束页号 = n_Endpage, 结束行号 = n_Endrow, 行差 = Decode(打印人, Null, 0, n_Rows),
          --只有打印过的数据才记录行差 
          打印人 = Null, 打印时间 = Null
      Where 记录id = n_记录oldid;
    End If;
  End If;
  --无行差，退出 
  If n_Rows = 0 Then
    Return;
  End If;

  --之后是否存在数据？ 
  Begin
    Select 1 Into n_Count From 病人护理打印 Where 文件id = 文件id_In And 发生时间 > 发生时间_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;

  If n_Count = 1 Then
    --更新之后数据的打印相关数据（除打印人与打印时间） 
    If n_Rows > 0 Then
      Update 病人护理打印
      Set 开始页号 = 开始页号 + Decode(Sign(开始行号 + n_Rows - n_Pagerows), 1, 1, 0),
          结束页号 = 结束页号 + Decode(Sign(结束行号 + n_Rows - n_Pagerows), 1, 1, 0),
          开始行号 = Decode(Mod(开始行号 + n_Rows, n_Pagerows), 0, n_Pagerows, Mod(开始行号 + n_Rows, n_Pagerows)),
          结束行号 = Decode(Mod(结束行号 + n_Rows, n_Pagerows), 0, n_Pagerows, Mod(结束行号 + n_Rows, n_Pagerows)), 打印人 = Null,
          打印时间 = Null
      Where 文件id = 文件id_In And 发生时间 > 发生时间_In;
    Else
      --新的行号小于1则页号-1 
      --新的行号+每页的有效行后再进行判断 
      Update 病人护理打印
      Set 开始页号 = 开始页号 - Decode(Sign(开始行号 + n_Rows - 1), -1, 1, 0),
          结束页号 = 结束页号 - Decode(Sign(结束行号 + n_Rows - 1), -1, 1, 0),
          开始行号 = Decode(Mod(开始行号 + n_Pagerows + n_Rows, n_Pagerows), 0, n_Pagerows,
                         Mod(开始行号 + n_Pagerows + n_Rows, n_Pagerows)),
          结束行号 = Decode(Mod(结束行号 + n_Pagerows + n_Rows, n_Pagerows), 0, n_Pagerows,
                         Mod(结束行号 + n_Pagerows + n_Rows, n_Pagerows)), 打印人 = Null, 打印时间 = Null
      Where 文件id = 文件id_In And 发生时间 > 发生时间_In;
      --程序应该是先删除了数据才更新的，所以不会存在页号为零的，页号为零的肯定已经删除了。 
      --DELETE 病人护理打印 WHERE 开始页号=0; 
    End If;
    --检查更新之后的打印数据是否存在连续跨一页以上，如果存在则禁止。 
    v_Print := '';
    For r_Print In (Select 发生时间, 开始页号
                    From 病人护理打印
                    Where 文件id = 文件id_In And 发生时间 > 发生时间_In And 结束页号 - 开始页号 > 1
                    Order By 发生时间) Loop
      If Lengthb(v_Print || Chr(13) || Chr(10) || '页号【' || r_Print.开始页号 || '】    发生时间【' ||
                 To_Char(r_Print.发生时间, 'YYYY-MM-DD hh24:mi:ss') || '】') < 800 Then
        v_Print := v_Print || Chr(13) || Chr(10) || '页号【' || r_Print.开始页号 || '】    发生时间【' ||
                   To_Char(r_Print.发生时间, 'YYYY-MM-DD hh24:mi:ss') || '】';
      End If;
    End Loop;
    If v_Print Is Not Null Then
      v_Err_Msg := '您在发生时间【' || To_Char(发生时间_In, 'YYYY-MM-DD hh24:mi:ss') || '】录入的数据影响了后续数据位置，导致以下数据连续跨了一页以上：';
      v_Err_Msg := v_Err_Msg || v_Print || Chr(13) || Chr(10) || '目前产品暂不支持对跨一页以上的数据进行展示和打印，操作终止！';
      Raise Err_Item;
    End If;
  End If;
  --进行关联文件的页号修正 
  Zl_病人护理打印_Batchretrypage(文件id_In, n_Firstdata || ';' || n_Firstdata);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人护理打印_Update;
/

--109518:张永康,2017-05-24,每批转出后不重建待转出索引
Create Or Replace Procedure Zl1_Datamoveout1
(
  Demoded_In        In Number,
  Optmode_In        In Number := 0,
  Curtime_In        In Number := 1,
  Totaltime_In      In Number := 1,
  Speedmode_In      In Number := 0,
  Disabletrigger_In In Number := 0,
  Disablejob_In     In Number := 0,
  Parallel_In       In Number := 0,
  Sysowner_In       In Varchar2 := Null,
  Peissysowner_In   In Varchar2 := Null,
  Opersysowner_In   In Varchar2 := Null
) As
  --功能：标记并转出n天前的数据到历史表空间 
  --参数:Demoded_in:          标记转出多少天以前的数据,当参数Optmode_In为0或1时才有效 
  --     Optmode_in:           0-标记并执行转出,1-只进行标记，2-只执行转出(将已标记的) 
  --     Curtime_in,Totaltime_in，连续多次转出时的当前次数和总次数，如果都为1表示一次性转出 
  --                首次时会检查在线表与历史表的结构一致性、在线表的子表是否转出，并且禁用他表外键，禁用转出表引用非转出表的外键索引 
  --                最后一次执行后，需在界面程序中手工恢复禁用的外键和索引 
  --     Speedmode_in:        0-在线模式，1-离线模式（在客户端停用时，转出期间禁用转出表的主键、唯一键、外键约束和索引，以加快删除操作） 
  --                          历史库的约束和索引禁用在应用程序时进行（因为需要用到历史库的连接） 
  --     Disabletrigger_in:   1=转出期间禁用当前所有者的触发器，0-不禁用 
  --     Disablejob_in:       1=转出期间禁用当前所有者的自动作业，0-不禁用 
  --     parallel_in:         重建标记查询所需索引时的并行度，缺省为不并行执行
  --     SysOwner_In:         标准系统指定转出历史表空间所有者
  --     PeisSysOwner_In:     体检系统指定转出历史表空间所有者
  --     OperSysOwner_In:     手麻系统指定转出历史表空间所有者
  --说明：1.标记要转出的数据，可以多次标记，然后分批执行转出 
  --      2.转出时，根据zlBakTables中定义的分组和顺序转出数据，分组提交事务; 
  --      3.为了避免查询范围太大导致性能问题，及Undo表空间增长太大，建议每次不要转出太多的数据(界面程序调用时自动拆分为每次调用转一个月); 
  d_End        Date;
  n_System     Number(5);
  v_Systems    Varchar2(100);
  n_Peissystem Number(5);
  n_Opersystem Number(5);
  n_Reset      Number(1) := 0;
  v_Sql        Varchar2(4000);
  v_Owner      Varchar2(20);

  v_Pre组号      Number(2);
  v_当前批次     Number(8);
  v_序列         Number(8);
  n_重建索引间隔 Zldatamove.重建索引间隔%Type;
  n_重建索引范围 Zldatamove.重建索引范围%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(4000);

  --功能：转移数据（插入后删除，按组提交事务） 
  Procedure Movedata
  (
    v_Table    In Varchar2,
    v_当前批次 In Varchar2,
    v_Owner    In Varchar2
  ) As
    v_Colstr Varchar2(4000);
  Begin
    Select f_List2str(Cast(Collect(Column_Name) As t_Strlist)) As Colsstr
    Into v_Colstr
    From (Select Column_Name From User_Tab_Columns Where Table_Name = v_Table Order By Column_Id);
  
    v_Sql := 'Insert Into /*+ append */ ' || v_Owner || '.' || v_Table || '(' || v_Colstr || ') Select ' || v_Colstr ||
             ' From ' || v_Table || ' Where 待转出 = ' || v_当前批次;
    Execute Immediate v_Sql;
  
    v_Sql := 'Delete ' || v_Table || ' Where 待转出 = ' || v_当前批次;
    Execute Immediate v_Sql;
    Commit;
    --每张表提交一次，避免Undo占用过多，耗时的业务查询可能报ora-01555快照太旧的错误
  End Movedata;

  --检查历史表等 
  Function Checkvalid(v_Systems In Varchar2) Return Varchar2 Is
    n_只读 Number(3);
    n_状态 Number(1);
    v_Err  Varchar2(4000);
    v_Tmp1 Varchar2(4000);
    v_Tmp2 Varchar2(4000);
    v_Tmp3 Varchar2(4000);
  Begin
    Select Count(1)
    Into n_只读
    From zlBakSpaces
    Where 系统 In (Select Column_Value From Table(f_Num2list(v_Systems))) And
          (所有者 = Sysowner_In Or 所有者 = Peissysowner_In Or 所有者 = Opersysowner_In) And 只读 = 1;
  
    If n_只读 > 0 Then
      v_Err := '[ZLSOFT]存在只读状态的当前历史数据空间,操作不能继续![ZLSOFT]';
      Return(v_Err);
    End If;
  
    --并发检查，避免人工转出期间，自动作业又调用本过程 
    Select Nvl(状态, 0) Into n_状态 From zlDataMove Where 系统 = n_System And 组号 = 1;
    If n_状态 = 1 Then
      v_Err := '[ZLSOFT]其他用户正在进行转出操作，如果不是这种情况，请手工更新"zlDataMove.状态"的值为空![ZLSOFT]';
      Return(v_Err);
    End If;
    --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- 
    If Optmode_In <> 2 Then
      --检查在线表与后备表的字段是否一致,以避免数据转移了一部分时才报错。 
      For R In (Select 表名 From zlBakTables Where 系统 In (Select Column_Value From Table(f_Num2list(v_Systems)))) Loop
        v_Tmp1 := '';
        v_Tmp2 := '';
        v_Tmp3 := '';
        For C In (Select *
                  From (Select a.Column_Name, a.Data_Type, a.Data_Precision, b.Column_Name As Bcolumn_Name,
                                b.Data_Type As Bdata_Type, b.Data_Precision As Bdata_Precision
                         From (Select Column_Name, Data_Type,
                                       Decode(Data_Type, 'VARCHAR2', Data_Length, Data_Precision) As Data_Precision
                                From User_Tab_Columns A
                                Where Table_Name = r.表名) A,
                              (Select Column_Name, Data_Type,
                                       Decode(Data_Type, 'VARCHAR2', Data_Length, Data_Precision) As Data_Precision
                                From All_Tab_Columns
                                Where Table_Name = r.表名 And
                                      Owner In
                                      (Select 所有者
                                       From zlBakSpaces
                                       Where 系统 In (Select Column_Value From Table(f_Num2list(v_Systems))) And
                                             (所有者 = Sysowner_In Or 所有者 = Peissysowner_In Or 所有者 = Opersysowner_In))) B
                         Where a.Column_Name = b.Column_Name(+))
                  Where Bcolumn_Name Is Null Or Data_Type <> Bdata_Type Or Data_Precision > Bdata_Precision) Loop
        
          If c.Bcolumn_Name Is Null Then
            v_Tmp1 := v_Tmp1 || ',' || c.Column_Name || ' ' || c.Data_Type || '(' || c.Data_Precision || ')';
          Elsif c.Data_Type <> c.Bdata_Type Then
            If c.Data_Type = 'DATE' Then
              v_Tmp2 := v_Tmp2 || ',' || c.Column_Name || ' ' || c.Data_Type || ',历史表的为' || c.Bdata_Type;
            Else
              v_Tmp2 := v_Tmp2 || ',' || c.Column_Name || ' ' || c.Data_Type || '(' || c.Data_Precision || '),历史表的为' ||
                        c.Bdata_Type;
            End If;
          Else
            v_Tmp3 := v_Tmp3 || ',' || c.Column_Name || ' ' || c.Data_Type || '(' || c.Data_Precision || '),历史表的为' ||
                      c.Bdata_Precision;
          End If;
        End Loop;
      
        If v_Tmp1 Is Not Null Then
          v_Err := v_Err || Chr(10) || ',缺字段：' || r.表名 || ' ' || v_Tmp1;
        End If;
        If v_Tmp2 Is Not Null Then
          v_Err := v_Err || Chr(10) || ',类型不同：' || r.表名 || ' ' || v_Tmp2;
        End If;
        If v_Tmp3 Is Not Null Then
          v_Err := v_Err || Chr(10) || ',长度较小：' || r.表名 || ' ' || v_Tmp3;
        End If;
      
        If Lengthb(v_Err) > 3000 Then
          v_Err := '[ZLSOFT]请到【管理工具】中执行【历史库修正】' || Substr(v_Err, 1, 3000) || '......[ZLSOFT]';
          Return(v_Err);
        End If;
      End Loop;
    
      If v_Err Is Not Null Then
        v_Err := '[ZLSOFT]请到【管理工具】中执行【历史库修正】:' || Substr(v_Err, 1, 3000) || '[ZLSOFT]';
        --重建H表视图的脚本生成语句示例： 
        --Select 'Create or replace view  ZLHIS.H' || 表名 || ' as Select * From ZLBAK1.' || 表名 || ';' From Zlbaktables Where 系统 In(Select Column_Value From Table(f_num2list(v_Systems))) 
        Return(v_Err);
      End If;
    
      --可能由于历史升级脚本的遗漏，有些不再使用的外键或子表没有删除，为了避免转移到中途时才报错，先检查一遍 
      For P In (Select Constraint_Name
                From (Select Constraint_Name,
                              Row_Number() Over(Partition By Constraint_Name Order By Decode(Constraint_Type, 'P', 0, 1)) Rn
                       From User_Constraints A, zlBakTables B
                       Where b.表名 = a.Table_Name And b.系统 In (Select Column_Value From Table(f_Num2list(v_Systems))) And
                             a.Constraint_Type In ('P', 'U'))
                Where Rn = 1) Loop
        For R In (Select a.Table_Name, a.Constraint_Name, a.Delete_Rule
                  From User_Constraints A
                  Where a.r_Constraint_Name = p.Constraint_Name And Not Exists
                   (Select 1
                         From zlBakTables B
                         Where b.表名 = a.Table_Name And b.系统 In (Select Column_Value From Table(f_Num2list(v_Systems))))
                  Order By a.r_Constraint_Name) Loop
          v_Err := v_Err || Chr(10) || r.Table_Name || '(' || r.Constraint_Name || ',' || r.Delete_Rule || '->' ||
                   p.Constraint_Name || ')';
          If Lengthb(v_Err) > 2000 Then
            v_Err := '[ZLSOFT]子表未转出:' || Substr(v_Err, 1, 2000) || '......[ZLSOFT]';
            Return(v_Err);
          End If;
        End Loop;
      End Loop;
    
      If v_Err Is Not Null Then
        v_Err := '[ZLSOFT]子表未转出:' || Substr(v_Err, 1, 2000) || '[ZLSOFT]';
        Return(v_Err);
      End If;
    End If;
    Return('');
  End Checkvalid;
Begin
  If Optmode_In <> 2 Then
    Select Trunc(Sysdate) - Demoded_In Into d_End From Dual;
  End If;
  v_Owner := Zl_Owner;
  Select 编号 Into n_System From zlSystems Where Upper(所有者) = v_Owner And 编号 Like '1%';

  Select Nvl(Min(编号), 0) Into n_Peissystem From zlSystems Where Upper(所有者) = v_Owner And 编号 Like '21%';
  Select Nvl(Min(编号), 0) Into n_Opersystem From zlSystems Where Upper(所有者) = v_Owner And 编号 Like '24%';

  --1.安全性检查 
  ----------------------------------------------------------------------------------- 
  If Curtime_In = 1 Then
    v_Systems := n_System;
    If n_Peissystem > 0 Then
      v_Systems := v_Systems || ',' || n_Peissystem;
    End If;
    If n_Opersystem > 0 Then
      v_Systems := v_Systems || ',' || n_Opersystem;
    End If;
  
    v_Err_Msg := Checkvalid(v_Systems);
    If v_Err_Msg Is Not Null Then
      Raise Err_Item;
    End If;
  
    --一批中的首次调用时禁用触发器和作业 
    If Disabletrigger_In = 1 Then
      Zl1_Datamove_Reb(n_System, Speedmode_In, 1, 0);
      If n_Peissystem > 0 Then
        Zl1_Datamove_Reb(n_Peissystem, Speedmode_In, 1, 0);
      End If;
      If n_Opersystem > 0 Then
        Zl1_Datamove_Reb(n_Opersystem, Speedmode_In, 1, 0);
      End If;
    End If;
  
    If Disablejob_In = 1 Then
      Zl1_Datamove_Reb(n_System, Speedmode_In, 2, 0);
      If n_Peissystem > 0 Then
        Zl1_Datamove_Reb(n_Peissystem, Speedmode_In, 2, 0);
      End If;
      If n_Opersystem > 0 Then
        Zl1_Datamove_Reb(n_Opersystem, Speedmode_In, 2, 0);
      End If;
    End If;
  
    Update zlDataMove Set 状态 = 1 Where 系统 = n_System And 组号 = 1;
    Commit;
  End If;

  --2.标记要转出的数据 
  ----------------------------------------------------------------------------------- 
  If Optmode_In <> 2 Then
    --上次标记转出出错后继续进行标记转出 
    Select Nvl(Max(批次), 0) Into v_当前批次 From Zldatamovelog Where 系统 = n_System And 待转出 = 2;
  
    If v_当前批次 = 0 Then
      Select Nvl(Max(批次), 0) + 1, Decode(Curtime_In, 1, Nvl(Max(序列), 0) + 1, Max(序列))
      Into v_当前批次, v_序列
      From Zldatamovelog
      Where 系统 = n_System;
    
      Insert Into Zldatamovelog
        (系统, 批次, 序列, 截止时间, 标记开始时间, 待转出, 当前进度)
      Values
        (n_System, v_当前批次, v_序列, d_End, Sysdate, 2, '正在标记待转出数据');
      Commit;
    Else
      Update Zldatamovelog
      Set 标记开始时间 = Sysdate, 当前进度 = '正在标记待转出数据'
      Where 系统 = n_System And 批次 = v_当前批次;
      Commit;
    End If;
  
    Zl1_Datamove_Tag(d_End, v_当前批次, n_System);
    If n_Peissystem > 0 Then
      Execute Immediate 'Begin Zl21_Datamove_Tag(:1, :2, :3); End;'
        Using d_End, v_当前批次, n_Peissystem;
    End If;
    If n_Opersystem > 0 Then
      Execute Immediate 'Begin Zl24_Datamove_Tag(:1, :2, :3); End;'
        Using d_End, v_当前批次, n_Opersystem;
    End If;
  
    Update Zldatamovelog
    Set 标记结束时间 = Sysdate, 当前进度 = '标记待转出数据完成', 待转出 = 1
    Where 系统 = n_System And 批次 = v_当前批次;
    Commit;
  End If;

  --3.转移数据处理 
  ----------------------------------------------------------------------------------- 
  If Optmode_In = 1 Then
    If Curtime_In = Totaltime_In Then
      Update zlDataMove Set 状态 = Null Where 系统 = n_System And 组号 = 1;
    End If;
    Commit;
  Else
    --从最小的批次开始执行转出 
    If Optmode_In = 2 Then
      Select Nvl(Min(批次), 0), Max(截止时间)
      Into v_当前批次, d_End
      From Zldatamovelog
      Where 系统 = n_System And 待转出 = 1;
    
      If v_当前批次 = 0 Then
        Update zlDataMove Set 状态 = Null Where 系统 = n_System And 组号 = 1;
        Return;
      End If;
    End If;
  
    --禁用约束和索引 
    If Curtime_In = 1 Then
      Update Zldatamovelog Set 当前进度 = '正在禁用约束和索引' Where 系统 = n_System And 批次 = v_当前批次;
      --要先禁用约束，否则主键或唯一键的索引被禁用后，会导致查询或插入操作报错，而禁用主键或唯一键则会删除对应的索引 
      n_Reset := 1;
      Zl1_Datamove_Reb(n_System, Speedmode_In, 3, 0);
      If n_Peissystem > 0 Then
        Zl1_Datamove_Reb(n_Peissystem, Speedmode_In, 3, 0);
      End If;
      If n_Opersystem > 0 Then
        Zl1_Datamove_Reb(n_Opersystem, Speedmode_In, 3, 0);
      End If;
    
      Zl1_Datamove_Reb(n_System, Speedmode_In, 4, 0);
      If n_Peissystem > 0 Then
        Zl1_Datamove_Reb(n_Peissystem, Speedmode_In, 4, 0);
      End If;
      If n_Opersystem > 0 Then
        Zl1_Datamove_Reb(n_Opersystem, Speedmode_In, 4, 0);
      End If;
    End If;
  
    --数据转出处理 
    ----------------------------------------------------------------------------------- 
    --不更新汇总表：病人费用汇总，药品收发汇总，药品库存，人员缴款余额等，虽然只是更新的期初数， 
    --但是，由于并非这个时间以前的数据都转出了（部分未符合转出条件的数据未转出），更新后，如果按时间查询，会发现在线中这些日期的数据非常小，容易引起误解 
    --即使出于某些特殊原因需要更新汇总表，也可以通过汇总表处理的过程进行重新汇总，所以，不必在转出过程中逐条更新。 
  
    --"标记结束时间=转出开始时间"时不记录 
    If Optmode_In = 2 Then
      Update Zldatamovelog Set 转出开始时间 = Sysdate Where 系统 = n_System And 批次 = v_当前批次;
    End If;
  
    --a.转出标准版数据
    For R In (Select 表名, 组号 From zlBakTables Where 系统 = n_System And 直接转出 = 1 Order By 组号, 序号) Loop
      If Nvl(v_Pre组号, -1) <> r.组号 Then
        Update Zldatamovelog
        Set 当前进度 = '正在转出第' || r.组号 || '组(' || r.表名 || '...)数据'
        Where 系统 = n_System And 批次 = v_当前批次;
        Commit;
      End If;
    
      Movedata(r.表名, v_当前批次, Sysowner_In);
      v_Pre组号 := r.组号;
    End Loop;
  
    --b.转出体检数据
    v_Pre组号 := -1;
    For R In (Select 表名, 组号 From zlBakTables Where 系统 = n_Peissystem And 直接转出 = 1 Order By 组号, 序号) Loop
      If Nvl(v_Pre组号, -1) <> r.组号 Then
        Update Zldatamovelog
        Set 当前进度 = '正在转出体检第' || r.组号 || '组(' || r.表名 || '...)数据'
        Where 系统 = n_System And 批次 = v_当前批次;
        Commit;
      End If;
    
      Movedata(r.表名, v_当前批次, Peissysowner_In);
      v_Pre组号 := r.组号;
    End Loop;
  
    --c.转出手麻数据
    v_Pre组号 := -1;
    For R In (Select 表名, 组号 From zlBakTables Where 系统 = n_Opersystem And 直接转出 = 1 Order By 组号, 序号) Loop
      If Nvl(v_Pre组号, -1) <> r.组号 Then
        Update Zldatamovelog
        Set 当前进度 = '正在转出手麻第' || r.组号 || '组(' || r.表名 || '...)数据'
        Where 系统 = n_System And 批次 = v_当前批次;
        Commit;
      End If;
    
      Movedata(r.表名, v_当前批次, Opersysowner_In);
      v_Pre组号 := r.组号;
    End Loop;
    Commit;
  
    Update 病案主页 Set 待转出 = Null, 数据转出 = 1 Where 待转出 = v_当前批次;
  
    Update zlDataMove Set 上次日期 = d_End Where 系统 = n_System And 组号 = 1;
  
    v_Sql := 'Update ' || Sysowner_In || '.zlBakInfo Set 最后转储日期 = Sysdate Where 系统 = ' || n_System;
    Execute Immediate v_Sql;
  
    If n_Peissystem > 0 Then
      v_Sql := 'Update ' || Peissysowner_In || '.zlBakInfo Set 最后转储日期 = Sysdate Where 系统 = ' || n_Peissystem;
      Execute Immediate v_Sql;
    End If;
  
    If n_Opersystem > 0 Then
      v_Sql := 'Update ' || Opersysowner_In || '.zlBakInfo Set 最后转储日期 = Sysdate Where 系统 = ' || n_Opersystem;
      Execute Immediate v_Sql;
    End If;
  
    Update Zldatamovelog
    Set 转出结束时间 = Sysdate, 待转出 = Null, 当前进度 = '转出数据完成,正在重建待转出索引'
    Where 系统 = n_System And 批次 = v_当前批次;
    Commit;
  
    If Curtime_In = Totaltime_In Then
      Update zlDataMove
      Set 状态 = Null, 本次最终日期 = Decode(Sign(d_End - 本次最终日期), -1, 本次最终日期, Null)
      Where 系统 = n_System And 组号 = 1;
      Commit;
    End If;
  
    --4.索引重建（以提高下次标记转出查询的速度） 
    ----------------------------------------------------------------------------------- 
    --每次转完后不要重建待转出索引，在线重建容易出现卡死，并且索引无法重建和删除（ORA-08104）
   
    --收缩标记转出查询所需的索引被删除后的空闲空间，下次标记转出时减少范围扫描的数据块 
    --如果每次转完后进行，则耗时较多，所以，可根据查询的耗时来动态决定间隔次数(界面缺省为24次转出后重建一次) 
    Select Nvl(重建索引间隔, 0), Nvl(重建索引范围, 0)
    Into n_重建索引间隔, n_重建索引范围
    From zlDataMove
    Where 系统 = n_System And 组号 = 1;
  
    If Mod(Curtime_In, n_重建索引间隔) = 0 And n_重建索引间隔 <> 0 Then
      Zl1_Datamove_Reb(n_System, Speedmode_In, 6, 1, Parallel_In, n_重建索引范围);
    
      If n_Peissystem > 0 Then
        Zl1_Datamove_Reb(n_Peissystem, Speedmode_In, 6, 1, Parallel_In, n_重建索引范围);
      End If;
      If n_Opersystem > 0 Then
        Zl1_Datamove_Reb(n_Opersystem, Speedmode_In, 6, 1, Parallel_In, n_重建索引范围);
      End If;
    End If;
  
    Update Zldatamovelog Set 重建结束时间 = Sysdate, 当前进度 = '完成' Where 系统 = n_System And 批次 = v_当前批次;
    Commit;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    --可能部分数据插入成功，删除操作由于主键或唯一的索引被禁用而失败 
    Rollback;
    Update zlDataMove Set 状态 = Null Where 系统 = n_System And 组号 = 1;
  
    v_Err_Msg := Substr(SQLErrM, 1, 60);
    If Curtime_In = 1 And n_Reset = 0 Then
      Update Zldatamovelog Set 当前进度 = '转出标记出错：' || v_Err_Msg Where 系统 = n_System And 批次 = v_当前批次;
    Else
      Update Zldatamovelog
      Set 当前进度 = '转出出错：' || v_Err_Msg || Substr(v_Sql, 1, 30)
      Where 系统 = n_System And 批次 = v_当前批次;
    End If;
    Commit;
  
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamoveout1;
/

--109518:张永康,2017-05-24,电子病历内容触发器的修改
Create Or Replace Procedure Zl1_Datamove_Reb
(
  System_In    In Number,
  Speedmode_In In Number,
  Func_In      In Number,
  Enable_In    In Number := 0,
  Parallel_In  In Number := 0,
  Rebscope_In  In Number := 0
) As
  --功能：在历史数据转出之前，禁用触发器、自动作业、约束、索引，转出之后启用这些对象，以及重建待转出索引，收回标记转出所用索引的空间 
  --参数： 
  --System_In:    应用系统编号,100=标准版 
  --speedmode_in：数据转出模式，0-在线模式，1-离线模式（在客户端停用时，转出期间禁用转出表的主键、唯一键、外键约束和索引，以加快已转数据的删除操作） 
  --func_in:      1=触发器，2=自动作业，3=约束，4=索引，5=重建待转出索引，6-收回标记转出所用索引的空间，7-重组表的存储空间（move），并恢复被禁用的约束和索引 ,8-重建标记转出查询所需索引以外的其他索引 
  --Enable_in:    0-禁用，1=启用，对func_in值为1-4有效 
  --rebScope_in:   Func_In=6时，指重建索引的范围(0-经济核算类,1-经济核算类及医嘱类,2-全部)，Func_In=7时指Move表的范围(0-经济核算类，1-全部) 

  v_Sql      Varchar2(4000);
  n_Do       Number(1);
  n_Parallel Number(1);
  v_Tbs      Varchar2(100);
  v_Prompt   Varchar2(100);
  d_Curdate  Date;

  --功能：1.禁用或启用引用转出表主键的他表外键,避免删除主表记录时对子表每行记录执行一次SQL查询或删除 
  --      2.禁用或启用主键或唯一键约束（禁用时会自动删除对应的索引，启用时自动创建），以提高数据删除性能 
  --例如：病人医嘱发送_FK_医嘱ID，如果这些外键所在的表，数据未转出（未在zlbaktables表中定义），执行前会检查并限制转出。 
  Procedure Setconstraintstatus As
    v_Pcol Varchar2(50);
    v_Fcol Varchar2(50);
    v_Del  Varchar2(4000);
  Begin
    --禁用时，先禁用引用转出表主键的他表外键，再禁用转出表的主键 
    If Enable_In = 0 Then
      --1.在线模式转出时，由于有业务产生删除操作，所以，对于级联删除的外键，用触发器来替代对子表数据的删除操作
      If Speedmode_In = 0 Then
        For Rp In (Select Distinct a.Table_Name As Ptable_Name, a.Constraint_Name
                   From User_Constraints A, User_Constraints C, zlBakTables B
                   Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
                         c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And
                         c.Delete_Rule = 'CASCADE'
                   Order By a.Table_Name) Loop
        
          Select f_List2str(Cast(Collect(Column_Name Order By Position) As t_Strlist))
          Into v_Pcol
          From User_Cons_Columns
          Where Constraint_Name = Rp.Constraint_Name;
        
          v_Del := '';
          For Rf In (Select b.Table_Name, b.Constraint_Name,
                            f_List2str(Cast(Collect(b.Column_Name Order By b.Position) As t_Strlist)) As r_Col
                     From User_Constraints A, User_Cons_Columns B
                     Where a.r_Constraint_Name = Rp.Constraint_Name And a.Constraint_Name = b.Constraint_Name
                     Group By b.Table_Name, b.Constraint_Name) Loop
            If Instr(v_Pcol, ',') > 0 Then
              v_Del := v_Del || Chr(10) || '        Delete ' || Rf.Table_Name || ' Where (' || Rf.r_Col ||
                       ') in ((:Old.' || Replace(v_Pcol, ',', ',:Old.') || '));';
            Else
              v_Del := v_Del || Chr(10) || '        Delete ' || Rf.Table_Name || ' Where ' || Rf.r_Col || ' = :Old.' ||
                       v_Pcol || ';';
            End If;
          End Loop;
        
          --对级联删除的外键，引用自身所在表的字段的情况，加上条件，只在删除父记录时才级联删除子记录
          --并且加上自治事务，否则会产生ora-2099,ora-04091错误,表XX的数据发生了变化，触发器函数不能读它
          Select Max(Column_Name)
          Into v_Fcol
          From User_Cons_Columns A, User_Constraints B
          Where a.Constraint_Name = b.Constraint_Name And b.r_Constraint_Name = b.Table_Name || '_PK' And
                b.r_Constraint_Name = Rp.Constraint_Name;
        
          If v_Fcol Is Not Null Then
            v_Sql := 'Create Or Replace Trigger ' || Rp.Ptable_Name || '_Cascade_Del' || Chr(10) || '    After Delete On ' ||
                     Rp.Ptable_Name || Chr(10) || '    For Each Row' || Chr(10) || 'Declare' || Chr(10) ||
                     ' Pragma Autonomous_Transaction;' || Chr(10) || 'Begin' || Chr(10) ||
                     '    If :Old.待转出 Is Null And :Old.' || v_Fcol || ' Is Null Then ' || v_Del || Chr(10) || '    Commit;' ||
                     Chr(10) || '    End If; ' || Chr(10) || 'End ' || Rp.Ptable_Name || '_Cascade_Del;';
          Else
            v_Sql := 'Create Or Replace Trigger ' || Rp.Ptable_Name || '_Cascade_Del' || Chr(10) || '    After Delete On ' ||
                     Rp.Ptable_Name || Chr(10) || '    For Each Row' || Chr(10) || 'Begin' || Chr(10) ||
                     '    If :Old.待转出 Is Null Then ' || v_Del || Chr(10) || '    End If; ' || Chr(10) || 'End ' ||
                     Rp.Ptable_Name || '_Cascade_Del;';
          End If;
        
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.禁用引用转出表主键的他表外键
      For R In (Select c.Table_Name, c.Constraint_Name, a.Table_Name As Ptable_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'ENABLED'
                Order By a.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.禁用主键或唯一键索引(离线转出时)
      If Speedmode_In = 1 Then
        --必须删除索引，否则即使skip_unusable_indexes为true，也无法删除存在Unusable状态的唯一性索引的表中的记录
        --保留转出标记中的SQL查询所需的索引(主键和唯一键对应的索引) 
        For R In (Select a.Table_Name, a.Constraint_Name
                  From User_Constraints A, zlBakTables T, User_Tables B
                  Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And a.Status = 'ENABLED' And
                        a.Constraint_Type In ('P', 'U') And a.Table_Name = b.Table_Name And b.Iot_Type Is Null And
                        a.Constraint_Name Not In (Select Upper(索引名) From Zlbaktableindex Where 系统 = System_In)
                  Order By Constraint_Name) Loop
          v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name ||
                   ' Cascade Drop Index';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    Else
      --启用时
      --1.先启用主键和唯一键，再启用引用转出表主键的他表外键 
      If Speedmode_In = 1 Then
        --先重建索引，再启用约束，以便重建索引时利用并行执行缩短时间，并且启用约束时也可以采用novalidate方式 
        For R In (Select d.Table_Name, d.Constraint_Name,
                         f_List2str(Cast(Collect(d.Column_Name Order By d.Position) As t_Strlist)) Colstr
                  From User_Cons_Columns D,
                       (Select a.Table_Name, a.Constraint_Name
                         From User_Constraints A, zlBakTables T
                         Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And a.Status = 'DISABLED' And
                               a.Constraint_Type In ('P', 'U')) A
                  Where a.Constraint_Name = d.Constraint_Name And a.Table_Name = d.Table_Name
                  Group By d.Table_Name, d.Constraint_Name
                  Order By Constraint_Name) Loop
        
          Update zlDataMove Set 说明 = '正在恢复约束:' || r.Constraint_Name Where 系统 = 100 And 组号 = 1;
        
          Select Tablespace_Name Into v_Tbs From User_Indexes Where Table_Name = r.Table_Name And Rownum < 2;
        
          --禁用主键或唯一键时，索引是被删除了的，所以这里要用Create 
          v_Sql := 'Create Unique Index ' || r.Constraint_Name || ' On ' || r.Table_Name || '(' || r.Colstr ||
                   ') Tablespace ' || v_Tbs || ' Nologging';
          Begin
            Execute Immediate v_Sql;
          Exception
            When Others Then
              Null; --可能有些主键或唯一键不是本次转出期间被禁用的，之前就存在不唯一数据，创建唯一索引会出错 
          End;
        
          --会自动建立约束与索引的关联 
          v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.启用引用转出表主键的他表外键 
      For R In (Select c.Table_Name, c.Constraint_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'DISABLED'
                Order By a.Table_Name) Loop
        --为了加快速度，采用novalidate，不验证已有数据 
        --可能引用转出表主键的他表，在zlbaktables中定义了，但没有编写对应的数据转出脚本，未验证的数据可能有违反约束的情况。 
        v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.在线模式转出时，删除之前创建的用来替代级联删除外键的触发器
      If Speedmode_In = 0 Then
        For R In (Select a.Trigger_Name
                  From User_Triggers A, zlBakTables B
                  Where a.Table_Name = b.表名 And b.直接转出 = 1 And b.系统 = System_In And
                        Trigger_Name = Table_Name || '_CASCADE_DEL' And Triggering_Event = 'DELETE') Loop
          v_Sql := 'Drop Trigger ' || r.Trigger_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    End If;
  End Setconstraintstatus;

  --功能：高速模式时禁用LOB以外的所有索引，在线模式时仅禁用转出表引用非转出表的外键索引(例如：病人医嘱计价_IX_收费细目ID) 
  --说明：禁用索引是为了提高删除数据的性能 
  Procedure Setindexstatus As
  Begin
    If Speedmode_In = 1 Then
      --保留转出标记中的SQL查询所需的索引 
      For R In (Select /*+ rule*/
                 a.Index_Name
                From User_Indexes A, zlBakTables T
                Where a.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And t.直接转出 = 1 And
                      a.Index_Name <> a.Table_Name || '_IX_待转出' And
                      a.Index_Name Not In (Select Upper(索引名) From Zlbaktableindex Where 系统 = System_In) And
                      a.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And a.Index_Type = 'NORMAL' And Not Exists
                 (Select 1
                       From User_Constraints C
                       Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
          Execute Immediate v_Sql;
        Else
          Update zlDataMove Set 说明 = '正在重建索引:' || r.Index_Name Where 系统 = 100 And 组号 = 1;
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
          Begin
            Execute Immediate v_Sql;
            --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
          
          Exception
            When Others Then
              If SQLErrM Like 'ORA-00054%' Then
                v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
                Execute Immediate v_Sql;
              End If;
          End;
        End If;
      End Loop;
    Else
      For R In (Select a.Index_Name
                From (Select d.Table_Name, d.Index_Name,
                              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr
                       From User_Ind_Columns D, zlBakTables T, User_Indexes C
                       Where c.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And c.Uniqueness = 'NONUNIQUE' And
                             c.Index_Type = 'NORMAL' And c.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And
                             c.Index_Name = d.Index_Name And c.Table_Name = d.Table_Name
                       Group By d.Table_Name, d.Index_Name) A,
                     (Select e.Table_Name,
                              f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr
                       From User_Cons_Columns E, User_Constraints F, zlBakTables T, User_Constraints C
                       Where e.Table_Name = t.表名 And t.直接转出 = 1 And t.系统 = System_In And
                             e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And
                             c.Constraint_Name = f.r_Constraint_Name And c.Table_Name Not In ('病案主页', '病人信息') And
                             Not Exists
                        (Select 1 From zlBakTables G Where g.表名 = c.Table_Name And g.系统 = System_In)
                       Group By e.Table_Name, e.Constraint_Name) B
                Where a.Table_Name = b.Table_Name And a.Colstr = b.Colstr
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          --特殊处理：以下两个索引不禁用，是由于药品目录修改规格，财务缴款需要使用 
          If r.Index_Name Not In ('病人医嘱记录_IX_收费细目ID', '药品收发记录_IX_药品ID', '药品收发记录_IX_价格ID') Then
            v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
            Execute Immediate v_Sql;
          End If;
        Else
          Update zlDataMove Set 说明 = '正在重建索引:' || r.Index_Name Where 系统 = 100 And 组号 = 1;
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
          Execute Immediate v_Sql;
          --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
        End If;
      End Loop;
    End If;
  End Setindexstatus;

  --功能：转出数据期间，停用转出表上的所有触发器，转出后再恢复 
  Procedure Settriggerstatus As
  Begin
    For R In (Select Distinct a.Table_Name, t.停用触发器
              From User_Triggers A, zlBakTables T
              Where a.Status = Decode(Enable_In, 0, 'ENABLED', 'DISABLED') And a.Table_Name = t.表名 And t.直接转出 = 1 And
                    t.系统 = System_In) Loop
      If Enable_In = 0 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' DISABLE ALL TRIGGERS';
        Update zlBakTables Set 停用触发器 = 1 Where 系统 = System_In And 表名 = r.Table_Name;
      Elsif Nvl(r.停用触发器, 0) = 1 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' ENABLE ALL TRIGGERS';
        Update zlBakTables Set 停用触发器 = Null Where 系统 = System_In And 表名 = r.Table_Name;
      End If;
      Execute Immediate v_Sql;
    End Loop;
    Commit;
  End Settriggerstatus;

  --功能：转出数据期间，停用当前所有者的所有自动作业，转出后再启用 
  Procedure Setjobstatus As
    v_Jobs Varchar2(4000);
  Begin
    --停用 
    If Enable_In = 0 Then
      For R In (Select Job From User_Jobs Where Broken = 'N') Loop
        Dbms_Job.Broken(r.Job, True);
        v_Jobs := v_Jobs || ',' || r.Job;
      End Loop;
    
      If v_Jobs Is Not Null Then
        v_Jobs := Substr(v_Jobs, 2);
        Update zlDataMove Set 停用作业号 = v_Jobs Where 系统 = System_In And 组号 = 1;
      End If;
    Else
      --启用 
      Select 停用作业号 Into v_Jobs From zlDataMove Where 系统 = System_In And 组号 = 1;
      If v_Jobs Is Not Null Then
        For R In (Select Job
                  From User_Jobs
                  Where Broken = 'Y' And Job In (Select Column_Value From Table(f_Num2list(v_Jobs)))) Loop
          Dbms_Job.Broken(r.Job, False);
        End Loop;
        Update zlDataMove Set 停用作业号 = Null Where 系统 = System_In And 组号 = 1;
      End If;
    End If;
    --作业设置后必须提交事务才生效 
    Commit;
  End Setjobstatus;
Begin
  If Parallel_In < 2 Then
    Execute Immediate 'Alter Session DISABLE PARALLEL DDL';
  Else
    If Func_In In (6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
      --为重建索引设置并行执行（由于通常受限于IO设备的性能，设置太高的并行度反而会降低性能，如有高性能存储设备，可加大并行度） 
      --执行重建索引后会自动为索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢),在后面取消索引的并行度 
      --恢复在线库的约束和索引时，不管是不是在线模式，都加上并行，否则太慢
      Execute Immediate 'Alter Session FORCE PARALLEL DDL PARALLEL ' || Parallel_In;
      n_Parallel := 1;
    End If;
  End If;

  --提高索引创建速度（缩短40%以上的时间）
  If Func_In In (6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
    Execute Immediate 'alter session set workarea_size_policy=MANUAL';
  
    --设置直接路径IO的大小
    Execute Immediate 'alter session set events ''10351 trace name context forever, level 128''';
    Execute Immediate 'alter session SET db_file_multiblock_read_count=128';
    Execute Immediate 'alter session set "_sort_multiblock_read_count"=128';
    Begin
      --由于10G的BUG，sort_area_size需执行两次才会生效
      Execute Immediate 'alter session SET sort_area_size=512000000';
      Execute Immediate 'alter session SET sort_area_size=512000000';
    Exception
      When Others Then
        Null; --如果可用内存不足500M，失败则忽略
    End;
    Execute Immediate 'alter session SET db_block_checking=false';
  End If;

  If Func_In In (5, 6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
    d_Curdate := Sysdate;
  End If;

  If Func_In = 1 Then
    --1.设置触发器 
    Settriggerstatus;
  Elsif Func_In = 2 Then
    --2.设置自动作业 
    Setjobstatus;
  Elsif Func_In = 3 Then
    --3.设置约束状态    
    Setconstraintstatus;
    v_Prompt := '恢复禁用的约束';
  Elsif Func_In = 4 Then
    --4.设置索引状态 
    Setindexstatus;
    v_Prompt := '恢复禁用的索引';
  Elsif Func_In = 5 Then
    --5.重建"待转出"索引    
    For R In (Select Index_Name
              From (Select b.Index_Name
                     From zlBakTables A, User_Indexes B
                     Where a.表名 = b.Table_Name And a.直接转出 = 1 And a.系统 = System_In And
                           b.Index_Name = b.Table_Name || '_IX_待转出'
                     Union All
                     Select '病案主页_IX_待转出'
                     From Dual
                     Where System_In = 100)
              Order By 1) Loop
      Update zlDataMove Set 说明 = '正在重建待转出索引:' || r.Index_Name Where 系统 = 100 And 组号 = 1;
    
      --耗时太短，无须并行DDL 
      --在线转出时如果重建索引会锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
      --在线重建索引太慢，所以，即使在线转出模式也不用在线重建
      v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Begin
        Execute Immediate v_Sql;
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
    v_Prompt := '重建待转出索引';
  
  Elsif Func_In = 6 Then
    --6.重建标记转出查询所用到的索引（测试表明重建后最多可缩短一半的查询时间） 
    --根据业务的启用阶段来决定重建哪些索引，以避免一些不必要的重建耗时    
    For R In (Select b.Index_Name, a.组号
              From User_Indexes B, zlBakTables A
              Where a.表名 = b.Table_Name And a.系统 = System_In And
                    (b.Table_Name, b.Index_Name) In
                    (Select 表名, Upper(索引名) From Zlbaktableindex Where 系统 = System_In)
              Order By Index_Name) Loop
      n_Do := 0;
      If Rebscope_In = 0 Then
        If r.组号 < 5 Then
          n_Do := 1; --仅经济核算类 
        End If;
      Elsif Rebscope_In = 1 Then
        If r.组号 < 5 Or r.组号 = 8 Then
          n_Do := 1; --仅经济核算类、医嘱类 
        End If;
      Else
        n_Do := 1;
      End If;
    
      If n_Do = 1 Then
        Update zlDataMove Set 说明 = '正在重建索引:' || r.Index_Name Where 系统 = 100 And 组号 = 1;
      
        --v_Sql := 'Alter Index ' || r.Index_Name || ' shrink Space'; 
        --使用shrink方式不能并行执行,试验表明速度比rebuild PARALLEL 8 慢6倍 
        If Speedmode_In = 1 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
        Else
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
        End If;
        Begin
          Execute Immediate v_Sql;
          --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源 
        
        Exception
          When Others Then
            If SQLErrM Like 'ORA-00054%' Then
              v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
              Execute Immediate v_Sql;
            End If;
        End;
      End If;
    End Loop;
    v_Prompt := '重建标记转出所需索引';
  
    --重组表的数据
  Elsif Func_In = 7 Then
    --rebScope_in=0,只重组组号小于5的经济核算类表（费用、药品、票据），否则全部重组     
    For R In (Select a.表名 As Table_Name
              From zlBakTables A
              Where a.直接转出 = 1 And a.系统 = System_In And (组号 < Decode(Rebscope_In, 0, 5, 100))
              Order By 组号, 序号) Loop
    
      Update zlDataMove Set 说明 = '正在重组表:' || r.Table_Name Where 系统 = 100 And 组号 = 1;
    
      --如果有空闲的空间，最好移到其他表空间，只有这样才能绝对移动文件尾部的数据块，以便进行表空间文件的收缩 
      --在前面设置了会话级的强制并行 
      v_Sql := 'Alter Table ' || r.Table_Name || ' Move Nologging';
      Execute Immediate v_Sql;
    
      --单独移动Lob对象 
      For L In (Select Column_Name, Tablespace_Name From User_Lobs Where Table_Name = r.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Move Lob(' || l.Column_Name || ') Store as (Tablespace ' ||
                 l.Tablespace_Name || ') Nologging';
        Execute Immediate v_Sql;
      End Loop;
    
      v_Sql := 'Alter Table ' || r.Table_Name || ' Noparallel';
      Execute Immediate v_Sql;
    
      --move后，表相关的索引会全部失效，需要全部重建 
      For S In (Select Index_Name
                From User_Indexes
                Where Table_Name = r.Table_Name And Status = 'UNUSABLE'
                Order By Index_Name) Loop
      
        Update zlDataMove Set 说明 = '正在恢复失效索引:' || s.Index_Name Where 系统 = 100 And 组号 = 1;
      
        --在前面设置了会话级的强制并行 
        v_Sql := 'Alter Index ' || s.Index_Name || ' Rebuild Nologging';
        Execute Immediate v_Sql;
      End Loop;
    End Loop;
    v_Prompt := '重组表数据';
  
    --重建转出表上标记转出以外的其他索引（用于转出完成后收回空闲空间）
    --失效的索引不重建，因为转出完后有单独的重建功能
  Elsif Func_In = 8 Then
    For R In (Select b.Index_Name, a.组号
              From User_Indexes B, zlBakTables A
              Where a.表名 = b.Table_Name And a.系统 = System_In And b.Status = 'VALID' And b.Index_Type = 'NORMAL' And
                    b.Index_Name Not Like 'BIN$%' And
                    b.Index_Name Not In (Select Upper(索引名) From Zlbaktableindex Where 系统 = System_In)
              Order By Index_Name) Loop
    
      Update zlDataMove Set 说明 = '正在重建索引:' || r.Index_Name Where 系统 = 100 And 组号 = 1;
    
      If Speedmode_In = 1 Then
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Else
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
      End If;
      Begin
        Execute Immediate v_Sql;
        --在线重建比较慢，不在线重建则需要锁表，如果有其他并发事务，则会出错：ORA-00054: 资源正忙, 但指定以 NOWAIT 方式获取资源    
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
    v_Prompt := '重建标记转出以外的其他索引';
  End If;

  If Func_In In (5, 6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
    Update zlDataMove
    Set 说明 = To_Char(Sysdate, 'mm-dd hh24:mi') || v_Prompt || ':' || Trunc((Sysdate - d_Curdate) * 24 * 60) || '分钟'
    Where 系统 = 100 And 组号 = 1;
  End If;

  --执行重建索引后会自动为索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢) 
  --------------------------------------------------------------------------------------------------- 
  If n_Parallel = 1 Then
    Execute Immediate 'ALTER Session DISABLE PARALLEL DDL';
  
    For R In (Select Index_Name From User_Indexes Where Degree Not In ('1', '0')) Loop
      v_Sql := 'Alter Index ' || r.Index_Name || ' Noparallel';
      Execute Immediate v_Sql;
    End Loop;
  End If;

  Commit;
  --本过程不进行错误处理，错误由调用过程处理 
End Zl1_Datamove_Reb;
/


--108964:李小东,2017-05-24,不同年龄段参考值提取
CREATE OR REPLACE Function Zl_Get_Reference
(
  Type_In       In Number, --0=参考 1=参考ID 2=危急参考 3=危急参考下限 4=危急参考上限
  项目id_In     In Number,
  标本类型_In   In Varchar2,
  性别_In       In Number,
  出生日期_In   In Date,
  仪器id_In     In Number := Null,
  年龄_In       In Varchar2 := Null,
  申请科室id_In In Number := Null
) Return Varchar2 As

  Cursor v_Reference_Type Is
    Select a.Id,
           Trim(To_Char(a.参考低值, c.格式)) || '～' || Trim(To_Char(a.参考高值, c.格式)) ||
            Decode(a.临床特征, Null, '', '成人', '', '婴儿', '', ' ' || a.临床特征) As 结果参考, b.结果类型, b.取值序列,
           Trim(To_Char(a.警示下限, c.格式)) || '～' || Trim(To_Char(a.警示上限, c.格式)) ||
            Decode(a.临床特征, Null, '', '成人', '', '婴儿', '', ' ' || a.临床特征) As 危急参考, a.警示下限, a.警示上限, Nvl(b.多参考, 0) 多参考
    From 检验项目参考 A, 检验项目 B,
         (Select '9999990' ||
                   Decode(Max(Nvl(c.小数位数, -1)), 0, '', -1, '.00', Substr('.000000', 1, 1 + Max(Nvl(c.小数位数, -1)))) As 格式



           From 检验仪器项目 C, 检验项目 D
           Where d.诊治项目id = 项目id_In And d.诊治项目id = c.项目id(+)) C
    Where a.项目id = 项目id_In And a.项目id = b.诊治项目id;

  v_Return Varchar2(4000);
  v_Sql    Varchar2(4000);

  Type c_Type Is Ref Cursor; --声明REF游标类型
  r_Emp v_Reference_Type%RowType; --声明一个行类型变量
  Cur   c_Type; --声明REF游标类型的变量

  v_结果类型 Number(1);

  v_年数     Number(18, 1);
  v_月数     Number(18, 1);
  v_日数     Number(18, 1);
  v_小时     Number(18, 1);
  v_出生日期 Date;
  v_Pos      Number(4);
  v_多参考   Number(4);
  v_Value    Number(18);
  v_Valuerec Varchar2(255);
  v_年龄     Varchar2(50);
  v_结果参考 Varchar2(1000);
  v_参考id   Number(18);
  v_危紧参考 Varchar2(1000);
  v_警示下限 Varchar2(1000);
  v_警示上限 Varchar2(1000);
  d_Sysdate  Date;

  v_项目id_Bound   检验项目参考.项目id%Type;
  v_标本类型_Bound 检验项目参考.标本类型%Type;
  v_性别域1_Bound  检验项目参考.性别域%Type;
  v_性别域2_Bound  检验项目参考.性别域%Type;
  v_性别域3_Bound  检验项目参考.性别域%Type;
  v_仪器id_Bound   检验项目参考.仪器id%Type;

  v_年龄单位日_Bound   检验项目参考.年龄单位%Type;
  v_年龄单位月_Bound   检验项目参考.年龄单位%Type;
  v_年龄单位小时_Bound 检验项目参考.年龄单位%Type;
  v_年龄单位年_Bound   检验项目参考.年龄单位%Type;

  v_年龄单位日1_Bound   检验项目参考.年龄单位%Type;
  v_年龄单位月1_Bound   检验项目参考.年龄单位%Type;
  v_年龄单位小时1_Bound 检验项目参考.年龄单位%Type;
  v_年龄单位年1_Bound   检验项目参考.年龄单位%Type;

  v_临床特征_Bound   检验项目参考.临床特征%Type;
  v_申请科室id_Bound 检验项目参考.申请科室id%Type;
  v_年龄_1           Varchar2(50);
  v_年龄_2           Varchar2(50);

  Function Sub_Is_Number(v_In In Varchar2) Return Boolean Is
    n_Tmp Number;
  Begin
    n_Tmp := To_Number(v_In);
    If n_Tmp Is Not Null Then
      Return True;
    Else
      Return False;
    End If;
  Exception
    When Others Then
      Return False;
  End Sub_Is_Number;

  Function Zlsplit
  (
    v_Str       In Varchar2,
    v_Delimiter In Varchar2,
    v_Number    In Number
  ) Return Varchar2 Is
    v_Record     Varchar2(1000);
    v_Currrecord Varchar2(1000);
    v_Currnum    Number;
  Begin
    v_Record  := v_Str || v_Delimiter;
    v_Currnum := 0;
    While v_Record Is Not Null Loop
      v_Currrecord := Substr(v_Record, 1, Instr(v_Record, v_Delimiter) - 1);
      If v_Currnum = v_Number Then
        Return(v_Currrecord);
      End If;

      v_Currnum := v_Currnum + 1;
      v_Record  := Replace(v_Delimiter || v_Record, v_Delimiter || v_Currrecord || v_Delimiter);
    End Loop;

    Return('');
  End Zlsplit;
  Function Zlval(Vstr In Varchar2) Return Number Is
    Result Number(16, 6);
    Intbit Number(8);
    Strnum Varchar(10);
  Begin
    Strnum := '';
    For Intbit In 1 .. 10 Loop
      If Instr('0123456789.', Substr(Vstr, Intbit, 1)) = 0 Then
        Exit;
      End If;
      Strnum := Strnum || Substr(Vstr, Intbit, 1);
      Null;
    End Loop;
    Result := To_Number(Strnum);
    Return(Result);
  End Zlval;

Begin
  d_Sysdate := Sysdate;

  v_Sql := ' Select a.id,Trim(To_Char(A.参考低值, C.格式)) || ''～'' || Trim(To_Char(A.参考高值, C.格式)) || ' ||
           ' Decode(A.临床特征, Null, '''', ''成人'', '''', ''婴儿'','''', '' '' || A.临床特征) As 结果参考, B.结果类型, B.取值序列, ' ||
           ' Trim(To_Char(A.警示下限, C.格式)) || ''～'' || Trim(To_Char(A.警示上限, C.格式)) || ' || ' Decode(A.临床特征, Null, '''', ''成人'', '''', ''婴儿'','''', '' '' || A.临床特征) As 危急参考,a.警示下限,a.警示上限,
             nvl(b.多参考,0) 多参考 ' || ' From 检验项目参考 A, 检验项目 B, ' || ' (Select ''9999990'' || ' ||
           ' Decode(Max(Nvl(C.小数位数, -1)), 0, '''', -1, ''.00'', Substr(''.000000'', 1, 1 + Max(Nvl(C.小数位数, -1)))) As 格式 ' ||
           ' From 检验仪器项目 C, 检验项目 D ' || ' Where D.诊治项目ID = :项目ID And D.诊治项目ID = C.项目ID(+)) C ' ||
           ' Where A.项目ID = :项目ID And A.项目ID = B.诊治项目ID ';

  v_项目id_Bound := 项目id_In;

  v_年龄 := 年龄_In;
  If v_年龄 = '岁' Then
    v_年龄 := Null;
  End If;

  If v_年龄 = '月' Then
    v_年龄 := Null;
  End If;

  If v_年龄 = '小时' Then
    v_年龄 := Null;
  End If;

  If v_年龄 = '天' Then
    v_年龄 := Null;
  End If;

  If Nvl(标本类型_In, '') <> '' Or 标本类型_In Is Not Null Then
    v_Sql := v_Sql || ' And A.标本类型 = :标本类型 ';
  Else
    v_Sql := v_Sql || ' And (A.标本类型 = :标本类型 or 1=1) ';
  End If;
  v_标本类型_Bound := 标本类型_In;

  If Nvl(性别_In, '') <> '' Or 性别_In Is Not Null Then
    --V_Sql := V_Sql || ' And A.性别域 = Nvl(' || 性别_In || ', 1) ';
    v_Sql := v_Sql || ' And decode(A.性别域,null,:性别,0,:性别,A.性别域) = Nvl(:性别, 1) ';

  Else
    v_Sql := v_Sql || ' And (decode(A.性别域,null,:性别1,0,:性别2,A.性别域) = Nvl(:性别3, 1) or 1 = 1) ';
  End If;
  v_性别域1_Bound := 性别_In;
  v_性别域2_Bound := 性别_In;
  v_性别域3_Bound := 性别_In;

  If Nvl(仪器id_In, '') <> '' Or 仪器id_In Is Not Null Then
    v_Sql := v_Sql || ' And (A.仪器id = :仪器ID Or A.仪器id Is Null) ';
  Else
    v_Sql := v_Sql || ' And (A.仪器id = :仪器ID Or A.仪器id Is Null or 1=1) ';
  End If;
  v_仪器id_Bound := 仪器id_In;

  If Nvl(v_年龄, '') <> '' Or v_年龄 Is Not Null Then
    If Instr(v_年龄, '岁') > 0 Or Instr(v_年龄, '月') > 0 Or Instr(v_年龄, '天') > 0 Or Instr(v_年龄, '小时') > 0 Or
       Sub_Is_Number(v_年龄) Then
      --处理日期
      v_出生日期 := 出生日期_In;
      v_年龄_1   := v_年龄;
      If Instr(v_年龄_1, '岁') > 0 Then
        v_年龄   := Substr(v_年龄_1, 1, Instr(v_年龄_1, '岁'));
        v_年龄_2 := Substr(v_年龄_1, Instr(v_年龄_1, '岁') + 1);
      Elsif Instr(v_年龄, '月') > 0 Then
        v_年龄   := Substr(v_年龄_1, 1, Instr(v_年龄_1, '月'));
        v_年龄_2 := Substr(v_年龄_1, Instr(v_年龄_1, '月') + 1);
      Elsif Instr(v_年龄, '天') > 0 Then
        v_年龄   := Substr(v_年龄_1, 1, Instr(v_年龄_1, '天'));
        v_年龄_2 := Substr(v_年龄_1, Instr(v_年龄_1, '天') + 1);
      Elsif Instr(v_年龄, '小时') > 0 Then
        v_年龄   := Substr(v_年龄_1, 1, Instr(v_年龄_1, '小时') + 1);
        v_年龄_2 := Substr(v_年龄_1, Instr(v_年龄_1, '小时') + 2);
        If v_年龄 = '0小时' Or v_年龄 = '0时' Then
          v_年龄 := ' ';
        End If;
      End If;
      If v_年龄 Is Not Null And (v_年龄 = '成人' Or v_年龄 = '婴儿' Or v_年龄 = '岁') = False Then
        If Substr(v_年龄, 1, 1) = '*' Then
          v_出生日期 := Add_Months(d_Sysdate, -216);
        Else
          If Substr(v_年龄, Length(v_年龄)) = '月' Then
            v_出生日期 := Add_Months(d_Sysdate, -1 * Nvl(Zlval(v_年龄), 0));
          Else
            If Substr(v_年龄, Length(v_年龄)) = '天' Then
              v_出生日期 := d_Sysdate - Nvl(Zlval(v_年龄), 0);
            Else
              If Substr(v_年龄, Length(v_年龄) - 1) = '小时' Then
                If Nvl(Zlval(v_年龄), 0) <> 0 Then
                  v_出生日期 := d_Sysdate - Nvl(Zlval(v_年龄), 0) / 24;
                End If;
              Else
                v_出生日期 := Add_Months(d_Sysdate, -12 * Nvl(Zlval(v_年龄), 0)) - 1;
              End If;
            End If;
          End If;
          If v_年龄_2 Is Not Null Then
            If Substr(v_年龄_2, Length(v_年龄_2)) = '月' Then
              v_出生日期 := Add_Months(v_出生日期, -1 * Nvl(Zlval(v_年龄_2), 0));
            Else
              If Substr(v_年龄_2, Length(v_年龄_2)) = '天' Then
                v_出生日期 := v_出生日期 - Nvl(Zlval(v_年龄_2), 0);
              Else
                If Substr(v_年龄_2, Length(v_年龄_2) - 1) = '小时' Then
                  If Nvl(Zlval(v_年龄_2), 0) <> 0 Then
                    v_出生日期 := v_出生日期 - Nvl(Zlval(v_年龄_2), 0) / 24;
                  End If;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
      If Not (v_出生日期 Is Null) Then
        v_年数 := Round(Months_Between(d_Sysdate, v_出生日期) / 12 ,1);
        v_月数 := Round(Months_Between(d_Sysdate, v_出生日期) ,1);
        v_日数 := Round(d_Sysdate - v_出生日期 ,1);
        v_小时 := Round((d_Sysdate - (v_出生日期 - 1 / 24)) * 24 - 1);
      End If;
      v_Sql := v_Sql || 'And (Decode(A.年龄单位, ''日'',:日, ''月'',:月,''小时'',:小时,:年) ' ||
               ' Between Nvl(A.年龄下限, -9999) And Nvl(A.年龄上限, 9999) )';
    Else
      v_Sql := v_Sql || 'And (Decode(A.年龄单位, ''日'',:日, ''月'',:月,''小时'',:小时,:年) ' ||
               ' Between Nvl(A.年龄下限, -9999) And Nvl(A.年龄上限, 9999) or 1=1 )';
    End If;

  Else
    v_Sql := v_Sql || 'And (Decode(A.年龄单位, ''日'',:日, ''月'',:月,''小时'',:小时,:年) ' ||
             ' Between Nvl(A.年龄下限, -9999) And Nvl(A.年龄上限, 9999) or 1=1 )';
  End If;
  v_年龄单位日_Bound   := v_日数;
  v_年龄单位月_Bound   := v_月数;
  v_年龄单位小时_Bound := v_小时;
  v_年龄单位年_Bound   := v_年数;
  If Instr(v_年龄, '成人') > 0 Or Instr(v_年龄, '婴儿') > 0 Or Instr(v_年龄, '分钟') > 0 Then
    --处理成人和婴儿
    v_Sql := v_Sql || ' And A.临床特征 =:年龄';
  Else
    v_Sql := v_Sql || ' And (A.临床特征 =:年龄 or 1=1)';
    v_Sql := v_Sql || ' And instr(''婴儿,成人'',nvl(临床特征,'' '')) <= 0  ';
  End If;

  v_临床特征_Bound := Replace(v_年龄, '分钟', '婴儿');

  If Nvl(申请科室id_In, '') <> '' Or 申请科室id_In Is Not Null Then
    v_Sql := v_Sql || ' And (A.申请科室ID = :申请科室ID Or nvl(A.申请科室ID,0) = 0) ';
  Else
    v_Sql := v_Sql || ' And ((A.申请科室ID = :申请科室ID Or nvl(A.申请科室ID,0) = 0) or 1=1) ';
  End If;
  v_申请科室id_Bound := 申请科室id_In;

  If (Nvl(v_年龄, '') = '' Or v_年龄 Is Null) And (出生日期_In <> '' Or 出生日期_In Is Not Null) Then
    --按出生日期查询
    If Not (出生日期_In Is Null) Then
      v_年数 := Round(Months_Between(d_Sysdate, 出生日期_In) / 12 - 0.5);
      v_月数 := Round(Months_Between(d_Sysdate, 出生日期_In) - 0.5);
      v_日数 := Round(d_Sysdate - 出生日期_In - 0.5);
      v_小时 := Round((d_Sysdate - (出生日期_In - 1 / 24)) * 24 - 1);

      v_Sql := v_Sql || 'And (Decode(A.年龄单位, ''日'',:日, ''月'',:月,''小时'',:小时,:年) ' ||
               ' Between Nvl(A.年龄下限, -9999) And Nvl(A.年龄上限, 9999) )';
    Else
      v_Sql := v_Sql || 'And (Decode(A.年龄单位, ''日'',:日, ''月'',:月,''小时'',:小时,:年) ' ||
               ' Between Nvl(A.年龄下限, -9999) And Nvl(A.年龄上限, 9999) or 1=1 )';
    End If;

  Else
    v_Sql := v_Sql || 'And (Decode(A.年龄单位, ''日'',:日, ''月'',:月,''小时'',:小时,:年) ' ||
             ' Between Nvl(A.年龄下限, -9999) And Nvl(A.年龄上限, 9999) or 1=1 )';
  End If;
  v_年龄单位日1_Bound   := v_日数;
  v_年龄单位月1_Bound   := v_月数;
  v_年龄单位小时1_Bound := v_小时;
  v_年龄单位年1_Bound   := v_年数;

  --加上排序
  v_Sql := v_Sql || ' Order By a.默认 desc,A.临床特征 ';

  If Nvl(申请科室id_In, '') <> '' Or 申请科室id_In Is Not Null Then
    v_Sql := v_Sql || ' ,a.申请科室ID  ';
  End If;

  If Nvl(性别_In, '') <> '' Or 性别_In Is Not Null Then
    v_Sql := v_Sql || ' ,a.性别域 desc  ';
  Else
    v_Sql := v_Sql || ' ,a.性别域 ';
  End If;

  v_Sql := v_Sql || ' ,a.id ';

  v_Return := '';
  Open Cur For v_Sql
    Using v_项目id_Bound, v_项目id_Bound, v_标本类型_Bound, v_性别域1_Bound, v_性别域2_Bound, v_性别域3_Bound, v_仪器id_Bound, v_年龄单位日_Bound, v_年龄单位月_Bound, v_年龄单位小时_Bound, v_年龄单位年_Bound, v_临床特征_Bound, v_申请科室id_Bound, v_年龄单位日1_Bound, v_年龄单位月1_Bound, v_年龄单位小时1_Bound, v_年龄单位年1_Bound;

  Loop
    Fetch Cur
      Into r_Emp;
    Exit When Cur%NotFound;
    If Cur%RowCount > 0 Then

      v_结果类型 := r_Emp.结果类型;
      v_Valuerec := r_Emp.取值序列;
      v_参考id   := r_Emp.Id;
      v_多参考   := r_Emp.多参考;

      If Nvl(v_Return, '') = '' Or v_Return Is Null Then
        If Type_In = 2 Then
          v_Return := r_Emp.危急参考;
        Else
          v_Return := r_Emp.结果参考;
        End If;
      Else
        If Type_In = 2 Then
          v_Return := v_Return || Chr(13) || Chr(10) || r_Emp.危急参考;
        Else
          If v_多参考 = 1 Then
            v_Return := v_Return || Chr(13) || Chr(10) || r_Emp.结果参考;
          End If;
        End If;
      End If;

      --只增加第一个选出的警示参考
      If v_警示下限 = '' Or v_警示下限 Is Null Then
        v_警示下限 := r_Emp.警示下限;
      End If;
      If v_警示上限 = '' Or v_警示上限 Is Null Then
        v_警示上限 := r_Emp.警示上限;
      End If;
    End If;
  End Loop;

  If v_Return = '' Or v_Return Is Null Then
    Begin
      Select 结果参考, 结果类型, 取值序列, ID, 危急参考, 警示下限, 警示上限
      Into v_结果参考, v_结果类型, v_Valuerec, v_参考id, v_危紧参考, v_警示下限, v_警示上限
      From (Select a.Id,
                    Trim(To_Char(a.参考低值, c.格式)) || '～' || Trim(To_Char(a.参考高值, c.格式)) ||
                     Decode(a.临床特征, Null, '', '成人', '', '婴儿', '', ' ' || a.临床特征) As 结果参考, b.结果类型, b.取值序列,
                    Trim(To_Char(a.警示下限, c.格式)) || '～' || Trim(To_Char(a.警示上限, c.格式)) ||
                     Decode(a.临床特征, Null, '', '成人', '', '婴儿', '', ' ' || a.临床特征) As 危急参考, a.警示下限, a.警示上限
             From 检验项目参考 A, 检验项目 B,
                  (Select '9999990' ||
                            Decode(Max(Nvl(c.小数位数, -1)), 0, '', -1, '.00', Substr('.000000', 1, 1 + Max(Nvl(c.小数位数, -1)))) As 格式
                    From 检验仪器项目 C, 检验项目 D
                    Where d.诊治项目id = 项目id_In And d.诊治项目id = c.项目id(+)) C
             Where a.项目id = 项目id_In And a.项目id = b.诊治项目id
             Order By a.默认 Desc, a.临床特征, a.性别域)
      Where Rownum = 1;
      If Type_In = 2 Then
        v_Return := v_危紧参考;
      Else
        v_Return := v_结果参考;
      End If;
      --只增加第一个选出的警示参考
      If v_警示下限 = '' Or v_警示下限 Is Null Then
        v_警示下限 := r_Emp.警示下限;
      End If;
      If v_警示上限 = '' Or v_警示上限 Is Null Then
        v_警示上限 := r_Emp.警示上限;
      End If;
    Exception
      When Others Then
        v_Return := Null;
    End;
  End If;
  If v_Return <> '' Or v_Return Is Not Null Then

    If v_Return = '～' Then
      v_Return := '';
    Else
      If v_结果类型 = 2 Then
        v_Pos := Instr(v_Return, '～');

        Begin
          Select To_Number(Substr(v_Return, 1, v_Pos - 1)) Into v_Value From Dual;
        Exception
          When Others Then
            v_Value := 0;
        End;
        v_Return := Zlsplit(v_Valuerec, ';', v_Value);
      End If;
    End If;
    If Type_In = 0 Then
      Return v_Return;
    Elsif Type_In = 1 Then
      Return v_参考id;
    Elsif Type_In = 2 Then
      Return v_Return;
    Elsif Type_In = 3 Then
      Return v_警示下限;
    Elsif Type_In = 4 Then
      Return v_警示上限;
    End If;
  End If;
  Close Cur; --关闭游标
  Return v_Return;
End Zl_Get_Reference;
/

--108762:焦博,2017-05-23,在临床出诊安排医生姓名前增加职称标识符
Create Or Replace Procedure Zl_专业技术职务_更新标识符(编码标识符_In Varchar2) As
  --功能：修改医生职务标识符
  --格式：编码1,标识符1;编码2,标识符2;...
  v_编码   专业技术职务.编码%Type;
  v_标识符 专业技术职务.标识符%Type;
Begin
  For c_编码标识符 In (Select C1 As 编码, C2 As 标识符 From Table(f_Str2list2(编码标识符_In, ';', ',')) Order By 编码) Loop
    v_编码   := c_编码标识符.编码;
    v_标识符 := c_编码标识符.标识符;
    Update 专业技术职务 Set 标识符 = v_标识符 Where 编码 = v_编码;
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_专业技术职务_更新标识符;
/

--109314:陈刘,2017-05-18,取消失效函数的使用

CREATE OR REPLACE Procedure Zl_病区公告栏样式_Updatedata
(
  病区id_In In 病区公告栏样式.病区id%Type,
  Id_In     In 病区公告栏样式.Id%Type := Null
) Is

  v_Content Varchar2(2000);
  v_Xh      Varchar2(4000);

  --只提取系统项
  Cursor c_Callboard Is
    Select Id, 名称, 别名, 行号, 位置, 是否固定, 是否隐藏, 内容, 时间
    From 病区公告栏样式
    Where 病区id = 病区id_In And (Id_In Is Null Or Id = Id_In)
    Order By 行号, 位置;

  Cursor c_Xry Is
    Select f_List2str(Cast(Collect(出院病床) As t_Strlist)) As 新入院
    From (Select b.出院病床
           From 病人信息 a, 病案主页 b,
                (Select 病人id, 主页id
                  From 病人变动记录
                  Where 病区id = 病区id_In And (开始原因 = 2 Or 开始原因 = 1) And
                        开始时间 Between To_Date(To_Char(Sysdate, 'yyyy-MM-dd') || ' 00:00:00', 'yyyy-MM-dd hh24:mi:ss') And
                        Sysdate
                  Group By 病人id, 主页id) c, 在院病人 d
           Where a.病人id = b.病人id And a.主页ID = b.主页id And b.病人id = c.病人id And b.主页id = c.主页id And a.病人id = d.病人id And
                 a.当前病区id = d.病区id And d.病区id = 病区id_In And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null
           Order By b.出院病床);
  r_Xry c_Xry%Rowtype;

  Cursor c_Xzy Is
    Select f_List2str(Cast(Collect(出院病床) As t_Strlist)) As 新转入
    From (Select b.出院病床
           From 病人信息 a, 病案主页 b,
                (Select 病人id, 主页id
                  From 病人变动记录
                  Where 病区id = 病区id_In And (开始原因 = 3 Or 开始原因 = 15) And
                        开始时间 Between To_Date(To_Char(Sysdate, 'yyyy-MM-dd') || ' 00:00:00', 'yyyy-MM-dd hh24:mi:ss') And
                        Sysdate
                  Group By 病人id, 主页id) c, 在院病人 d
           Where a.病人id = b.病人id And a.主页ID = b.主页id And b.病人id = c.病人id And b.主页id = c.主页id And a.病人id = d.病人id And
                 a.当前病区id = d.病区id And d.病区id = 病区id_In And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null
           Order By b.出院病床);
  r_Xzy c_Xzy%Rowtype;

  Cursor c_Yjhl Is
    Select f_List2str(Cast(Collect(出院病床) As t_Strlist)) As 一级护理
    From (Select b.出院病床
           From 病人信息 a, 病案主页 b, 收费项目目录 c, 在院病人 d
           Where a.病人id = b.病人id And a.主页ID = b.主页id And b.护理等级id = c.Id And a.病人id = d.病人id And a.当前病区id = d.病区id And
                 d.病区id = 病区id_In And
                 (Instr(c.名称, '一') > 0 Or Instr(c.名称, 'I') > 0 Or Instr(c.名称, 'Ⅰ') > 0 Or Instr(c.名称, '1') > 0) And
                 Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null
           Order By b.出院病床);
  r_Yjhl c_Yjhl%Rowtype;

  Cursor c_Tjhl Is
    Select f_List2str(Cast(Collect(出院病床) As t_Strlist)) As 特级护理
    From (Select b.出院病床
           From 病人信息 a, 病案主页 b, 收费项目目录 c, 在院病人 d
           Where a.病人id = b.病人id And b.护理等级id = c.Id And a.病人id = d.病人id And a.当前病区id = d.病区id And d.病区id = 病区id_In And
                 (Instr(c.名称, '特') > 0 Or Instr(c.名称, '重') > 0) And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null
           Order By b.出院病床);
  r_Tjhl c_Tjhl%Rowtype;

  Cursor c_Bw Is
    Select f_List2str(Cast(Collect(出院病床) As t_Strlist)) As 病危
    From (Select b.出院病床
           From 病人信息 a, 病案主页 b, 在院病人 d
           Where a.病人id = b.病人id And a.主页ID = b.主页id And a.病人id = d.病人id And a.当前病区id = d.病区id And d.病区id = 病区id_In And
                 Instr(b.当前病况, '危') > 0 And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null
           Order By b.出院病床);
  r_Bw c_Bw%Rowtype;

  Cursor c_Ycy Is
    Select f_List2str(Cast(Collect(出院病床) As t_Strlist)) As 预出院
    From (Select b.出院病床
           From 病人信息 a, 病案主页 b, 在院病人 c
           Where a.病人id = b.病人id And a.主页ID = b.主页id And a.病人id = c.病人id And a.当前病区id = c.病区id And c.病区id = 病区id_In And
                 b.状态 = 3 And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null
           Order By b.出院病床);
  r_Ycy c_Ycy%Rowtype;

  Cursor c_Ss Is
    Select f_List2str(Cast(Collect(出院病床) As t_Strlist)) As 手术
    From (Select Distinct d.出院病床
           From 病人信息 b, 病案主页 d, 病人医嘱记录 a, 诊疗项目目录 c, 在院病人 e
           Where b.病人id = d.病人id And b.主页ID = d.主页id And d.病人id = a.病人id And d.主页id = a.主页id And b.病人id = e.病人id And
                 b.当前病区id = e.病区id And e.病区id = 病区id_In And
                 ((a.医嘱期效 = 0 And a.医嘱状态 In (3, 5, 6, 7, 8, 9) And (a.执行终止时间 Is Null Or a.执行终止时间 >= Sysdate)) Or
                 (a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8))) And
                 a.开嘱时间 Between To_Date(To_Char(Sysdate - 7, 'yyyy-MM-dd') || ' 00:00:00', 'yyyy-MM-dd hh24:mi:ss') And
                 Sysdate And
                 Substr(Nvl(a.标本部位, To_Char(开始执行时间, 'YYYY-MM-DD HH24:MI')), 1, 10) = To_Char(Sysdate, 'YYYY-MM-DD') And
                 Nvl(a.婴儿, 0) = 0 And a.诊疗项目id = c.Id And c.类别 = 'F' And Nvl(d.病案状态, 0) <> 5 And d.封存时间 Is Null
           Order By d.出院病床);
  r_Ss c_Ss%Rowtype;

  Cursor c_Fs Is
    Select f_List2str(Cast(Collect(出院病床) As t_Strlist)) As 发烧
    From (Select Distinct b.出院病床
           From 病人信息 a, 病案主页 b, 在院病人 f
           Where a.病人id = b.病人id And a.主页ID = b.主页id And a.病人id = f.病人id And a.当前病区id = f.病区id And f.病区id = 病区id_In And
                 Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And Exists
            (Select c.Id
                  From 病人护理文件 c, 病人护理数据 d, 病人护理明细 e
                  Where c.Id = d.文件id And d.Id = e.记录id And e.记录类型 = 1 And e.项目序号 = 1 And
                        Length(Translate(e.记录内容, '-.0123456789' || e.记录内容, '-.0123456789')) = Length(e.记录内容) And
                        Zl_To_Number(e.记录内容) >= 37.2 And e.终止版本 Is Null And b.病人id = c.病人id And b.主页id = c.主页id And
                        Nvl(c.婴儿, 0) = 0 And d.发生时间 Between Sysdate - 3 And Sysdate)
           Order By b.出院病床);
  r_Fs c_Fs%Rowtype;

  Cursor c_Gms Is
    Select f_List2str(Cast(Collect(出院病床) As t_Strlist)) As 过敏史
    From (Select Distinct b.出院病床
           From 病人信息 a, 病案主页 b, 病人过敏记录 c, 在院病人 d
           Where a.病人id = b.病人id And a.主页ID = b.主页id And b.病人id = c.病人id And a.病人id = d.病人id And a.当前病区id = d.病区id And
                 d.病区id = 病区id_In And Nvl(b.病案状态, 0) <> 5 And b.封存时间 Is Null And c.结果 = 1 And Not Exists
            (Select 药物id
                  From 病人过敏记录
                  Where (Nvl(药物id, 0) = Nvl(c.药物id, 0) Or Nvl(药物名, 'Null') = Nvl(c.药物名, 'Null')) And Nvl(结果, 0) = 0 And
                        记录时间 > c.记录时间 And 病人id = c.病人id)
           Order By b.出院病床);
  r_Gms c_Gms%Rowtype;

  Cursor c_Diy Is
    Select /*+ Rule */
     f_List2str(Cast(Collect(当前床号) As t_Strlist)) As 床号列表
    From (Select Distinct b.当前床号
           From 病人信息 b, 病案主页 c, 病人医嘱记录 a, 在院病人 d, ((Select Column_Value From Table(f_Num2list(v_Xh)))) e
           Where b.病人id = c.病人id And b.主页ID = c.主页id And c.病人id = a.病人id And c.主页id = a.主页id And b.病人ID=d.病人ID And b.当前病区id = d.病区id And
                 d.病区id = 病区id_In And
                 ((a.医嘱期效 = 0 And a.医嘱状态 In (3, 5, 6, 7, 8, 9) And a.开始执行时间 >= b.入院时间 And
                 (a.执行终止时间 Is Null Or a.执行终止时间 >= Sysdate)) Or
                 (a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8) And a.开始执行时间 Between Sysdate - 1 And Sysdate)) And
                 Nvl(a.婴儿, 0) = 0 And a.诊疗项目id + 0 = e.Column_Value And Nvl(c.病案状态, 0) <> 5 And c.封存时间 Is Null
           Order By b.当前床号);
  r_Diy c_Diy%Rowtype;

Begin
  For r_Board In c_Callboard Loop
    v_Content := '';
    If Instr(',新入院列表,新转入列表,一级护理列表,特级护理列表,病危列表,预出院列表,手术列表,发烧列表,过敏史列表,',
             ',' || r_Board.名称 || ',') > 0 Then
      --系统固定项
      If r_Board.名称 = '新入院列表' Then
        Open c_Xry;
        Fetch c_Xry
          Into r_Xry;
        If c_Xry%Rowcount > 0 Then
          v_Content := Nvl(r_Xry.新入院, '');
        End If;

        Update 病区公告栏样式 Set 内容 = v_Content, 时间 = Sysdate Where Id = r_Board.Id;
        Close c_Xry;
      Elsif r_Board.名称 = '新转入列表' Then
        Open c_Xzy;
        Fetch c_Xzy
          Into r_Xzy;
        If c_Xzy%Rowcount > 0 Then
          v_Content := Nvl(r_Xzy.新转入, '');
        End If;

        Update 病区公告栏样式 Set 内容 = v_Content, 时间 = Sysdate Where Id = r_Board.Id;
        Close c_Xzy;
      Elsif r_Board.名称 = '一级护理列表' Then
        Open c_Yjhl;
        Fetch c_Yjhl
          Into r_Yjhl;
        If c_Yjhl%Rowcount > 0 Then
          v_Content := Nvl(r_Yjhl.一级护理, '');
        End If;

        Update 病区公告栏样式 Set 内容 = v_Content, 时间 = Sysdate Where Id = r_Board.Id;
        Close c_Yjhl;
      Elsif r_Board.名称 = '特级护理列表' Then
        Open c_Tjhl;
        Fetch c_Tjhl
          Into r_Tjhl;
        If c_Tjhl%Rowcount > 0 Then
          v_Content := Nvl(r_Tjhl.特级护理, '');
        End If;

        Update 病区公告栏样式 Set 内容 = v_Content, 时间 = Sysdate Where Id = r_Board.Id;
        Close c_Tjhl;
      Elsif r_Board.名称 = '病危列表' Then
        Open c_Bw;
        Fetch c_Bw
          Into r_Bw;
        If c_Bw%Rowcount > 0 Then
          v_Content := Nvl(r_Bw.病危, '');
        End If;

        Update 病区公告栏样式 Set 内容 = v_Content, 时间 = Sysdate Where Id = r_Board.Id;
        Close c_Bw;
      Elsif r_Board.名称 = '预出院列表' Then
        Open c_Ycy;
        Fetch c_Ycy
          Into r_Ycy;
        If c_Ycy%Rowcount > 0 Then
          v_Content := Nvl(r_Ycy.预出院, '');
        End If;

        Update 病区公告栏样式 Set 内容 = v_Content, 时间 = Sysdate Where Id = r_Board.Id;
        Close c_Ycy;
      Elsif r_Board.名称 = '手术列表' Then
        Open c_Ss;
        Fetch c_Ss
          Into r_Ss;
        If c_Ss%Rowcount > 0 Then
          v_Content := Nvl(r_Ss.手术, '');
        End If;

        Update 病区公告栏样式 Set 内容 = v_Content, 时间 = Sysdate Where Id = r_Board.Id;
        Close c_Ss;
      Elsif r_Board.名称 = '发烧列表' Then
        Open c_Fs;
        Fetch c_Fs
          Into r_Fs;
        If c_Fs%Rowcount > 0 Then
          v_Content := Nvl(r_Fs.发烧, '');
        End If;

        Update 病区公告栏样式 Set 内容 = v_Content, 时间 = Sysdate Where Id = r_Board.Id;
        Close c_Fs;
      Else
        Open c_Gms;
        Fetch c_Gms
          Into r_Gms;
        If c_Gms%Rowcount > 0 Then
          v_Content := Nvl(r_Gms.过敏史, '');
        End If;

        Update 病区公告栏样式 Set 内容 = v_Content, 时间 = Sysdate Where Id = r_Board.Id;
        Close c_Gms;
      End If;
    Else
      --自添加已绑定的项目
      v_Content := '';
      Begin
        Select f_List2str(Cast(Collect(a.Xh) As t_Strlist))
        Into v_Xh
        From 病区公告栏样式 p, Xmltable('/ITEMLIST/ITEM/XH' Passing p.诊疗项目 Columns Xh Varchar2(256) Path '/XH') a
        Where p.Id = r_Board.Id;
      Exception
        When Others Then
          v_Xh := '';
      End;

      If v_Xh Is Not Null Then
        Open c_Diy;
        Fetch c_Diy
          Into r_Diy;
        If c_Diy%Rowcount > 0 Then
          v_Content := Nvl(r_Diy.床号列表, '');
        End If;

        Update 病区公告栏样式 Set 内容 = v_Content, 时间 = Sysdate Where Id = r_Board.Id;
        Close c_Diy;
      Else
        Update 病区公告栏样式 Set 内容 = Null, 时间 = Sysdate Where Id = r_Board.Id;
      End If;
    End If;
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_病区公告栏样式_Updatedata;
/

--109214:刘尔旋,2017-05-18,取消失效函数的使用
Create Or Replace Procedure Zl_Third_Getvisitdetails
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:根据挂号单号获取该次就诊详情
  --入参:Xml_In: 
  --<IN>
  --    <GHDH>挂号单号</GHDH>
  --    <JSKLB>结算卡类别</JSKLB>
  --</IN>
  --出参:Xml_Out 
  --<OUTPUT>
  --    <DJLIST>  //如果为空表示为找到数据
  --        <DJ>
  --            <NO>单据号</NO>
  --            <DJLX>单据类型</DJLX> //1-收费单;4-挂号单
  --            <KDSJ>开单时间</KDSJ>
  --            <ZFZT>支付状态</ZFZT>    //0未支付1已支付
  --            <SFJSK>是否结算卡支付</SFJSK> //即该单据是否存在入参<JSKLB>来进行支付的,是返回1,否则返回0
  --            <LX>类型</LX> //挂号单固定为挂号,其他按收费类别汇总
  --            <ZXKS>执行科室</ZXKS>
  --            <ZXKSID>执行科室ID</ZXKSID>
  --            <MXLIST> 
  --                     <MX>
  --                                <JZSJ>就诊时间</JZSJ>    //挂号有效:yyyy-mm-dd hh24:mi:ss
  --                                <BW>部位</BW>               //检查,检验时有效
  --                                <XM>项目名称</XM>     //挂号无效:其他项目有效
  --                                <ZXZT>执行状态</ZXZT> //挂号:未接诊;已接诊;完成就诊;收费:未执行;已执行;部分执行
  --                                <BG>报告状态</BG>// 1-已出报告;0未出报告,检查,检验时有效 
  --                                <BLID>病历ID</BLID>  //如果<BG>字段为1，该值不为空,检查,检验时有效
  --                                <GG>规格</GG>                       //药品有效
  --                                <SL>数量</SL> //非挂号有效
  --                                <DW>单位</DW> //非挂号有效
  --                                <DJ>单价</DJ> //非挂号有效
  --                                <JE>金额</JE>  
  --                     </MX>
  --             </MXLIST>
  --             <DL> //队列
  --                        <XH>序号</XH>
  --                        <QMRS>前面人数</QMRS>  //(由Oracle函数zl_GetSequenceBeforPerons获取)
  --             </DL>
  --        </DJ>
  --    </DJLIST>
  --    <ERROR><MSG></MSG></ERROR>                      //如果错误返回
  --</OUTPUT>

  -------------------------------------------------------------------------------------------------- 
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  x_Templet Xmltype; --模板XML 

  v_卡类别   Varchar2(100);
  n_卡类别id Number(18);
  v_挂号单   Varchar2(10);
  v_排队号码 Varchar2(10);
  n_Temp     Number(18);

  n_Count Number(18);

  v_Temp       Varchar2(32767); --临时XML 
  v_队列       Varchar2(32767);
  v_No         Varchar2(50);
  v_Tmp        Varchar2(4000);
  n_Add_Djlist Number(1); --是否增加了DJLIST的;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB')
  Into v_挂号单, v_卡类别
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_挂号单 Is Null Then
    v_Err_Msg := '不能找到指定的挂号单号(当前挂号单号为空)';
    Raise Err_Item;
  End If;
  n_Add_Djlist := 0;

  v_Err_Msg := Null;
  If v_卡类别 Is Not Null Then
    Begin
      n_卡类别id := To_Number(v_卡类别);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
  
    If n_卡类别id = 0 Then
      Begin
        Select ID, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_Err_Msg
        From 医疗卡类别
        Where 名称 = v_卡类别;
      Exception
        When Others Then
          v_Err_Msg := '卡类别:' || v_卡类别 || '不存在!';
      End;
    
    Else
    
      Begin
        Select ID, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_Err_Msg
        From 医疗卡类别
        Where ID = n_卡类别id;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  End If;
  --1.获取挂号数据
  n_Count := 0;
  For c_挂号 In (Select a.Id, a.No, a.记录性质, a.执行部门id, c.名称 As 执行部门, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间,
                      a.预约时间, a.接收时间, To_Char(a.发生时间, 'yyyy-mm-dd HH24:mi') As 就诊时间, a.号别, a.号序, b.金额, a.记录状态,
                      Decode(Nvl(a.执行状态, 0), 0, '等待接诊', 1, '完成就诊', 2, '正在就诊', -1, '取消就诊') As 执行状态,
                      Decode(Nvl(b.结帐id, 0), 0, 0, 1) As 支付标志
               From 病人挂号记录 A,
                    (Select NO, Max(Nvl(结帐id, 0)) As 结帐id, Sum(实收金额) As 金额
                      From 门诊费用记录 B
                      Where 记录性质 = 4 And NO = v_挂号单
                      Group By NO) B, 部门表 C
               Where a.No = v_挂号单 And a.No = b.No And a.执行部门id = c.Id(+)) Loop
    If Nvl(c_挂号.记录状态, 0) <> 1 Then
      v_Err_Msg := '单据号:' || v_挂号单 || '已经被退号!';
      Raise Err_Item;
    End If;
    Begin
      Select 排队号码 Into v_排队号码 From 排队叫号队列 Where 业务id = c_挂号.Id And Nvl(业务类型, 0) = 0;
    Exception
      When Others Then
        v_排队号码 := Null;
    End;
    If v_排队号码 Is Not Null Then
      --业务id_In ,业务类型_In 排队号码_In Number := Null
      n_Temp := Zl_Getsequencebeforperons(c_挂号.Id, 0, v_排队号码);
      v_队列 := v_队列 || '<DL><XH>' || v_排队号码 || '</XH><QMRS>' || n_Temp || '</QMRS></DL>';
    End If;
    n_Temp := 0;
    If Nvl(n_卡类别id, 0) <> 0 Then
      Begin
        Select 1
        Into n_Temp
        From 病人预交记录
        Where NO = v_挂号单 And 记录性质 = 4 And 记录状态 In (1, 3) And 卡类别id = n_卡类别id And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    End If;
    v_Temp := '<NO>' || c_挂号.No || '</NO>';
    v_Temp := v_Temp || '<DJLX>' || 4 || '</DJLX>';
    v_Temp := v_Temp || '<KDSJ>' || c_挂号.登记时间 || '</KDSJ>';
    v_Temp := v_Temp || '<ZFZT>' || c_挂号.支付标志 || '</ZFZT>';
    v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
    v_Temp := v_Temp || '<LX>挂号</LX>';
    v_Temp := v_Temp || '<ZXKS>' || c_挂号.执行部门 || '</ZXKS>';
    v_Temp := v_Temp || '<ZXKSID>' || c_挂号.执行部门id || '</ZXKSID>';
    v_Temp := v_Temp || '<MXLIST><MX><JZSJ>' || c_挂号.就诊时间 || '</JZSJ><JE>' || c_挂号.金额 || '</JE></MX></MXLIST>';
    If v_队列 Is Not Null Then
      v_Temp := v_Temp || v_队列;
    End If;
  
    If Nvl(n_Add_Djlist, 0) = 0 Then
      --增加DJList节点
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<DJLIST></DJLIST>')) Into x_Templet From Dual;
      n_Add_Djlist := 1;
    End If;
    v_Temp := '<DJ>' || v_Temp || '</DJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT/DJLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    n_Count := n_Count + 1;
  End Loop;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '未找到指定的挂号单据:' || v_挂号单 || '!';
    Raise Err_Item;
  End If;

  --2.分类汇总收费单
  v_No := '-_';

  For c_费用 In (Select j.医嘱id, j.相关id As 组号, j.No, j.序号, j.收费类别, i.名称 As 收费类别名, j.执行部门id, q.名称 As 执行部门, j.收费细目id, m.名称,
                      m.规格, Max(j.计算单位) As 计算单位, Decode(Max(j.执行状态), 0, '未执行', 1, '完全执行', 2, '部分执行', '') As 执行状态,
                      Max(j.付款状态) As 付款状态, To_Char(Max(j.登记时间), 'yyyy-mm-dd hh24:mi:ss') As 登记时间, Max(j.单价) As 单价,
                      Sum(j.数量) As 数量, Sum(j.实收金额) As 实收金额
               From (Select a.相关id, a.Id As 医嘱id, b.No, b.收费类别, Max(Decode(b.记录状态, 0, 0, 1)) As 付款状态, b.结帐id, b.执行部门id,
                             Max(Decode(b.记录状态, 2, 0, b.执行状态)) As 执行状态,
                             Max(Decode(b.记录状态, 2, Null + Sysdate, b.登记时间)) As 登记时间, Nvl(b.价格父号, b.序号) As 序号, b.收费细目id,
                             b.计算单位, Sum(b.标准单价) As 单价, Avg(Nvl(b.付数, 1) * b.数次) As 数量, Sum(b.实收金额) As 实收金额

                      
                      From 门诊费用记录 B, 病人医嘱记录 A
                      Where Mod(b.记录性质, 10) = 1 And a.Id = b.医嘱序号 And Nvl(b.费用状态, 0) = 0 And a.挂号单 = v_挂号单
                      Group By a.相关id, a.Id, b.No, b.收费类别, b.结帐id, b.执行部门id, Nvl(b.价格父号, b.序号), b.收费细目id, b.计算单位) J,
                    收费项目目录 M, 部门表 Q, 收费项目类别 I
               Where j.收费细目id = m.Id And j.执行部门id = q.Id(+) And j.收费类别 = i.编码(+)
               Group By j.医嘱id, j.相关id, j.No, j.序号, j.收费类别, i.名称, j.执行部门id, q.名称, j.收费细目id, m.名称, m.规格
               Order By 登记时间 Desc, NO Desc, 收费类别, 序号) Loop
    If c_费用.No <> v_No Then
      n_Temp := 0;
      --单据不同,则产生的结构不同
      If Nvl(c_费用.付款状态, 0) = 1 Then
        --是否结算卡支付的
        Begin
          Select 1
          Into n_Temp
          From 病人预交记录 A, 门诊费用记录 B
          Where a.结帐id = b.结帐id And b.No = c_费用.No And Mod(b.记录性质, 10) = 1 And b.记录状态 In (1, 3) And a.卡类别id = n_卡类别id And
                Rownum < 2;
        Exception
          When Others Then
            n_Temp := 0;
        End;
      End If;
      v_Tmp := Null;
      Begin
        Select f_List2str(Cast(Collect(名称) As t_Strlist))
        Into v_Tmp
        From (Select Distinct b.名称
               From 门诊费用记录 A, 收费项目类别 B
               Where a.收费类别 = b.编码 And a.No = c_费用.No And a.记录性质 = 1 And a.记录状态 In (1, 3));
      Exception
        When Others Then
          Null;
      End;
    
      If Nvl(n_Add_Djlist, 0) = 0 Then
        --增加DJList节点
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<DJLIST></DJLIST>')) Into x_Templet From Dual;
        n_Add_Djlist := 1;
      End If;
    
      v_No   := c_费用.No;
      v_Temp := '<NO>' || c_费用.No || '</NO>';
      v_Temp := v_Temp || '<DJLX>' || 1 || '</DJLX>';
      v_Temp := v_Temp || '<KDSJ>' || c_费用.登记时间 || '</KDSJ>';
      v_Temp := v_Temp || '<ZFZT>' || c_费用.付款状态 || '</ZFZT>';
      v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
    
      v_Temp := v_Temp || '<LX>' || Nvl(Replace(v_Tmp, ',', '/'), '') || '</LX>';
      v_Temp := v_Temp || '<ZXKS>' || c_费用.执行部门 || '</ZXKS>';
      v_Temp := v_Temp || '<ZXKSID>' || c_费用.执行部门id || '</ZXKSID>';
      v_Temp := v_Temp || '<MXLIST></MXLIST>' || Nvl(v_队列, '') || '';
      v_Temp := '<DJ NO="' || c_费用.No || '">' || v_Temp || '</DJ>';
      Select Appendchildxml(x_Templet, '/OUTPUT/DJLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    End If;
  
    v_Temp := '<XM>' || Nvl(c_费用.名称, '') || '</XM>';
    If c_费用.收费类别 = 'D' Then
      --检查获取部位
      Begin
        Select f_List2str(Cast(Collect(标本部位) As t_Strlist))
        Into v_Tmp
        From 病人医嘱记录
        Where 相关id = c_费用.医嘱id;
      Exception
        When Others Then
          v_Tmp := Null;
      End;
      v_Temp := v_Temp || '<BW>' || Nvl(v_Tmp, '') || '</BW>';
    Elsif c_费用.收费类别 = 'C' Then
      --检验
      Begin
        Select Max(Decode(b.审核时间, Null, 0, 1))
        Into n_Temp
        From 病人医嘱记录 A, 检验标本记录 B
        Where a.Id = c_费用.医嘱id And a.Id = b.医嘱id(+) And Exists
         (Select 1 From 病人医嘱记录 Where 相关id = c_费用.医嘱id And 诊疗类别 = 'C');
      Exception
        When Others Then
          n_Temp := 0;
      End;
      v_Temp := v_Temp || '<BG>' || n_Temp || '</BG>';
      If n_Temp = 1 Then
        --取病历ID
        Begin
          Select 病历id
          Into n_Temp
          From 病人医嘱报告
          Where 医嘱id = c_费用.医嘱id And Nvl(病历id, 0) <> 0 And Rownum < 2;
        Exception
          When Others Then
            n_Temp := Null;
        End;
        v_Temp := v_Temp || '<BLID>' || Nvl(n_Temp, '') || '</BLID>';
      End If;
    End If;
  
    v_Temp := v_Temp || '<GG>' || Nvl(c_费用.规格, '') || '</GG>';
    v_Temp := v_Temp || '<SL>' || Nvl(c_费用.数量, 0) || '</SL>';
    v_Temp := v_Temp || '<DW>' || Nvl(c_费用.计算单位, '') || '</DW>';
    v_Temp := v_Temp || '<DJ>' || Nvl(c_费用.单价, 0) || '</DJ>';
    v_Temp := v_Temp || '<JE>' || Nvl(c_费用.实收金额, 0) || '</JE>';
    v_Temp := '<MX>' || v_Temp || '</MX>';
    Select Appendchildxml(x_Templet, '/OUTPUT/DJLIST/DJ[@NO="' || v_No || '"]/MXLIST', Xmltype(v_Temp))
    Into x_Templet
    From Dual;
  
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getvisitdetails;
/

--109174:李南春,2017-05-16,删除医疗卡同时删除对应特定项目
Create Or Replace Procedure Zl_医疗卡类别_Delete(Id_In In 医疗卡类别.ID%Type) Is
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_是否启用 Number;
  n_是否固定 Number;
  v_特定项目 Varchar2(20);
Begin
  Begin
    Select 是否启用, 是否固定, 特定项目 Into n_是否启用, n_是否固定, v_特定项目 From 医疗卡类别 Where ID = Id_In;
  Exception
    When Others Then
      n_是否启用 := -1;
  End;
  If Nvl(n_是否启用, 0) = -1 Then
    v_Err_Msg := '[ZLSOFT]医疗卡类别可能被人他人删除，不能再次删除![ZLSOFT]';
    Raise Err_Item;
  End If;
  If Nvl(n_是否启用, 0) = 0 Then
    v_Err_Msg := '[ZLSOFT]医疗卡类别已经被停用，不能删除![ZLSOFT]';
    Raise Err_Item;
  End If;
  If Nvl(n_是否固定, 0) = 1 Then
    v_Err_Msg := '[ZLSOFT]医疗卡类别是系统固定的，不能删除![ZLSOFT]';
    Raise Err_Item;
  End If;

  Delete From 医疗卡类别 Where ID = Id_In And Nvl(是否启用, 0) = 1 And Nvl(是否固定, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]医疗卡类别可能被人他人删除，不能再次删除![ZLSOFT]';
    Raise Err_Item;
  End If;
  
  IF Not v_特定项目 is Null then
    Delete From 收费特定项目 Where 特定项目 = v_特定项目;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_医疗卡类别_Delete;
/

--108046:刘尔旋,2017-05-15,费用审核管理取消审核出错问题
Create Or Replace Procedure Zl_费用审核记录_Delete
(
  费用id_In In 费用审核记录.费用id%Type,
  性质_In   In 费用审核记录.性质%Type
) Is
  v_Err_Msg Varchar2(100);
  Err_Item Exception;
  n_Count Number(4);
  v_No    门诊费用记录.No%Type;
Begin
  Select Count(a.Id), Max(a.No)
  Into n_Count, v_No
  From 门诊费用记录 A, (Select Mod(记录性质, 10) As 记录性质, NO, 序号 From 门诊费用记录 Where ID = 费用id_In) B
  Where a.No = b.No And Mod(a.记录性质, 10) = b.记录性质 And a.序号 = b.序号
  Group By a.No, Mod(a.记录性质, 10), a.序号
  Having Sum(a.数次) <> 0 Or Sum(a.实收金额) <> 0;
  If n_Count = 0 Then
    v_Err_Msg := '[ZLSOFT]单据『' || v_No || '』可能因并发原因,已经被他人转出或退费,不能取消审核![ZLSOFT]';
    Raise Err_Item;
  End If;
  Delete From 费用审核记录 Where 费用id = 费用id_In And 性质 = 性质_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_费用审核记录_Delete;
/

--109116:刘尔旋,2017-05-15,挂号医保校正处理误差费
Create Or Replace Procedure Zl_病人结算记录_Update
(
  结帐id_In       病人预交记录.结帐id%Type,
  保险结算_In     Varchar2, --"结算方式|结算金额||....."
  结帐_In         Number := 0,
  缺省结算方式_In Varchar2 := Null,
  缺省冲预交_In   Number := 0, --0-用现金缴款,1:剩于款项用冲预交支付(门诊预交),2-剩于款项用冲预交支付(住院预交)
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  结算卡序号_In   病人预交记录.结算卡序号%Type := Null,
  卡号_In         病人预交记录.卡号%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  合作单位_In     病人预交记录.合作单位%Type := Null,
  仅修正医保_In   Number := 0
) As
  --该游标为要删除的由费用记录产生的结算记录

  Cursor c_Del Is
    Select a.Id, a.记录性质, a.冲预交, a.结算方式, b.性质, a.预交类别
    From 病人预交记录 A, 结算方式 B
    Where a.结算方式 = b.名称 And a.结帐id = 结帐id_In;

  Cursor c_Del_医保 Is
    Select a.Id, a.记录性质, a.冲预交, a.结算方式, b.性质, a.预交类别
    From 病人预交记录 A, 结算方式 B
    Where a.结算方式 = b.名称 And b.性质 In (3, 4) And a.结帐id = 结帐id_In;

  --相关信息
  v_No         病人预交记录.No%Type;
  v_病人id     住院费用记录.病人id%Type;
  v_主页id     住院费用记录.主页id%Type;
  v_发生时间   住院费用记录.发生时间%Type;
  v_登记时间   住院费用记录.登记时间%Type;
  v_操作员编号 住院费用记录.操作员编号%Type;
  v_操作员姓名 住院费用记录.操作员姓名%Type;

  --本次结算变量
  v_金额合计 病人预交记录.冲预交%Type;

  --保险结算
  v_保险结算 Varchar2(255);
  v_当前结算 Varchar2(50);
  v_现金结算 病人预交记录.结算方式%Type;
  v_结算方式 病人预交记录.结算方式%Type;
  v_结算金额 病人预交记录.冲预交%Type;

  v_记录性质 病人预交记录.记录性质%Type;
  v_缺省     病人预交记录.结算方式%Type;

  --分币处理及误差变量
  v_现金金额   病人预交记录.冲预交%Type;
  v_Cashcented 病人预交记录.冲预交%Type;
  v_误差金额   病人预交记录.冲预交%Type;
  v_费用id     住院费用记录.Id%Type;
  v_序号       住院费用记录.序号%Type;
  v_收费类别   住院费用记录.收费类别%Type;
  v_收费细目id 住院费用记录.收费细目id%Type;
  v_收入项目id 住院费用记录.收入项目id%Type;
  v_收据费目   住院费用记录.收据费目%Type;
  n_Noexists   Number(3);
  n_医疗小组id 住院费用记录.医疗小组id%Type;
  n_结算序号   病人预交记录.结算序号%Type;
  n_费用状态   门诊费用记录.费用状态%Type;
  n_预交金额   病人预交记录.金额%Type;
  n_当前金额   病人预交记录.金额%Type;
  v_误差项     结算方式.名称%Type;

  --临时变量
  Err_Item Exception;
  v_Err_Msg  Varchar2(255);
  n_组id     财务缴款分组.Id%Type;
  n_执行状态 门诊费用记录.执行状态%Type;
Begin
  --如果缺省结算方式为空，则取现金结算方式
  If 缺省结算方式_In Is Null Then
    Begin
      Select 名称 Into v_缺省 From 结算方式 Where 性质 = 1 And Rownum < 2;
    Exception
      When Others Then
        v_缺省 := '现金';
    End;
  Else
    v_缺省 := 缺省结算方式_In;
  End If;

  --取得本次结算的相关信息
  If Nvl(结帐_In, 0) = 1 Then
    Select NO, 病人id, 收费时间, 操作员编号, 操作员姓名, 缴款组id, 0
    Into v_No, v_病人id, v_登记时间, v_操作员编号, v_操作员姓名, n_组id, n_执行状态
    From 病人结帐记录
    Where ID = 结帐id_In;
  Else
    Begin
      n_Noexists := 0;
      Select NO, 病人id, 登记时间, 操作员编号, 操作员姓名, 缴款组id, 执行状态, 费用状态
      Into v_No, v_病人id, v_登记时间, v_操作员编号, v_操作员姓名, n_组id, n_执行状态, n_费用状态
      From 门诊费用记录
      Where 结帐id = 结帐id_In And Rownum < 2;
    Exception
      When Others Then
        n_Noexists := 1;
    End;
    If n_Noexists = 1 Then
      --费用记录不存在，从补充记录中找
      Select NO, 病人id, 登记时间, 操作员编号, 操作员姓名, 缴款组id, 费用状态
      Into v_No, v_病人id, v_登记时间, v_操作员编号, v_操作员姓名, n_组id, n_费用状态
      From 费用补充记录
      Where 结算id = 结帐id_In And Rownum < 2;
    End If;
    If Nvl(n_费用状态, 0) = 1 Then
      --异常单据为空:
      v_缺省 := Null;
    End If;
  
    Begin
      --20051027 陈东
      Select 记录性质
      Into v_记录性质
      From 病人预交记录
      Where 结帐id = 结帐id_In And Rownum = 1 And Mod(记录性质, 10) <> 1;
    Exception
      When Others Then
        v_记录性质 := -1;
    End;
    If v_记录性质 = -1 Then
      Begin
        Select Decode(记录性质, 1, 3, 11, 3, 4, 4, 记录性质)
        Into v_记录性质
        From 门诊费用记录
        Where 结帐id = 结帐id_In And Rownum = 1;
      Exception
        When Others Then
          --可能是卡费
          Select 记录性质 Into v_记录性质 From 住院费用记录 Where 结帐id = 结帐id_In And Rownum = 1;
      End;
    End If;
  End If;

  If Nvl(v_病人id, 0) <> 0 And Nvl(结帐_In, 0) = 1 Then
    Select 主页id Into v_主页id From 病人信息 Where 病人id = v_病人id;
  End If;
  Select 结算序号 Into n_结算序号 From 病人预交记录 Where 结帐id = 结帐id_In And Rownum = 1;

  ----回退缴款,预交不动,因为没有改冲预交的
  --收费未最未最终完成的,代表按异常单据修正,不处理人员缴款余额
  v_金额合计 := 0;
  If Nvl(仅修正医保_In, 0) = 0 Then
    For r_Del In c_Del Loop
      If r_Del.记录性质 Not In (1, 11) Then
        If Nvl(n_费用状态, 0) <> 1 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) - r_Del.冲预交
          Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = r_Del.结算方式;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (v_操作员姓名, r_Del.结算方式, 1, -1 * r_Del.冲预交);
          End If;
        End If;
        v_金额合计 := v_金额合计 + r_Del.冲预交;
        Delete From 病人预交记录 Where ID = r_Del.Id;
      Else
        --检查是否冲预交
        If Nvl(缺省冲预交_In, 0) <> 0 Then
          v_金额合计 := v_金额合计 + r_Del.冲预交;
          If Nvl(n_费用状态, 0) <> 1 Then
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Del.冲预交, 0)
            Where 病人id = v_病人id And 类型 = Nvl(r_Del.预交类别, 2);
            If Sql%NotFound Then
              Insert Into 病人余额
                (病人id, 性质, 预交余额, 费用余额, 类型)
              Values
                (v_病人id, 1, Nvl(r_Del.冲预交, 0), 0, Nvl(r_Del.预交类别, 2));
            End If;
          End If;
          If r_Del.记录性质 = 1 Then
            Update 病人预交记录 Set 冲预交 = 0 Where ID = r_Del.Id;
          Else
            Delete 病人预交记录 Where ID = r_Del.Id;
          End If;
        End If;
      End If;
    End Loop;
  Else
    For r_Del In c_Del_医保 Loop
      If r_Del.记录性质 Not In (1, 11) Then
        If Nvl(n_费用状态, 0) <> 1 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) - r_Del.冲预交
          Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = r_Del.结算方式;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (v_操作员姓名, r_Del.结算方式, 1, -1 * r_Del.冲预交);
          End If;
        End If;
        v_金额合计 := v_金额合计 + r_Del.冲预交;
        Delete From 病人预交记录 Where ID = r_Del.Id;
      Else
        --检查是否冲预交
        If Nvl(缺省冲预交_In, 0) <> 0 Then
          v_金额合计 := v_金额合计 + r_Del.冲预交;
          If Nvl(n_费用状态, 0) <> 1 Then
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Del.冲预交, 0)
            Where 病人id = v_病人id And 类型 = Nvl(r_Del.预交类别, 2);
            If Sql%NotFound Then
              Insert Into 病人余额
                (病人id, 性质, 预交余额, 费用余额, 类型)
              Values
                (v_病人id, 1, Nvl(r_Del.冲预交, 0), 0, Nvl(r_Del.预交类别, 2));
            End If;
          End If;
          If r_Del.记录性质 = 1 Then
            Update 病人预交记录 Set 冲预交 = 0 Where ID = r_Del.Id;
          Else
            Delete 病人预交记录 Where ID = r_Del.Id;
          End If;
        End If;
      End If;
    End Loop;
  End If;

  --------------------------------------------------------------------------------------------------------------
  --------------------------------------------------------------------------------------------------------------
  --产生医保支付结算
  If 保险结算_In Is Not Null Then
    --各个保险结算
    v_保险结算 := 保险结算_In || '||';
    While v_保险结算 Is Not Null Loop
      v_当前结算 := Substr(v_保险结算, 1, Instr(v_保险结算, '||') - 1);
    
      v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, '|') - 1);
      v_结算金额 := To_Number(Substr(v_当前结算, Instr(v_当前结算, '|') + 1));
    
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
      Values
        (病人预交记录_Id.Nextval, Decode(结帐_In, 1, 2, v_记录性质), v_No, 1, v_病人id, v_主页id, '保险部份', v_结算方式, v_登记时间, v_操作员编号,
         v_操作员姓名, v_结算金额, 结帐id_In, n_组id, n_结算序号, Mod(Decode(结帐_In, 1, 2, v_记录性质), 10));
    
      v_金额合计 := v_金额合计 - v_结算金额;
    
      v_保险结算 := Substr(v_保险结算, Instr(v_保险结算, '||') + 2);
    End Loop;
  End If;
  --剩余部分用预交
  If Nvl(缺省冲预交_In, 0) <> 0 And v_金额合计 <> 0 Then
    n_预交金额 := v_金额合计;
    --先缴先用
    --不包含结算方式为代收款项的预交款。
    For c_预交 In (Select a.No, Sum(Nvl(a.金额, 0) - Nvl(a.冲预交, 0)) As 金额, Nvl(Max(a.结帐id), 0) As 结帐id, a.预交类别,
                        Max(Decode(a.记录性质, 1, a.记录状态, 1)) As 记录状态,
                        Max(Decode(a.记录性质, 1, Decode(a.记录状态, 1, a.Id, 3, a.Id, 0), 0)) As ID,
                        Max(Decode(a.记录性质, 1, Decode(a.记录状态, 1, a.收款时间, 3, a.收款时间, Null, Null))) As 收款时间
                 From 病人预交记录 A
                 Where a.记录性质 In (1, 11) And a.病人id = v_病人id And Nvl(a.预交类别, 2) = 缺省冲预交_In And
                       a.结算方式 Not In (Select 名称 From 结算方式 Where 性质 = 5)
                 Group By a.No, a.预交类别
                 Having Sum(Nvl(a.金额, 0) - Nvl(a.冲预交, 0)) <> 0
                 Order By 收款时间) Loop
    
      n_当前金额 := Case
                  When c_预交.金额 - n_预交金额 < 0 Then
                   c_预交.金额
                  Else
                   n_预交金额
                End;
    
      If c_预交.结帐id = 0 Then
        --第一次冲预交(将第一次标上结帐ID,冲预交标记为0)
        Update 病人预交记录
        Set 冲预交 = 0, 结帐id = 结帐id_In, 结算序号 = n_结算序号, 结算性质 = Mod(Decode(结帐_In, 1, 2, v_记录性质), 10)
        Where ID = c_预交.Id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
         冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
               v_登记时间, v_操作员姓名, v_操作员编号, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结算序号,
               Mod(Decode(结帐_In, 1, 2, v_记录性质), 10)
        From 病人预交记录
        Where NO = c_预交.No And 记录状态 = c_预交.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
      Where 病人id = v_病人id And 性质 = 1 And 类型 = Nvl(c_预交.预交类别, 2);
      --检查是否已经处理完
      If c_预交.金额 < n_预交金额 Then
        n_预交金额 := n_预交金额 - c_预交.金额;
      Else
        n_预交金额 := 0;
      End If;
    
      If n_预交金额 = 0 Then
        Exit;
      End If;
    End Loop;
    If n_预交金额 <> 0 Then
      v_Err_Msg := '[ZLSOFT]预交余不够支付本次支付金额,不能继续操作！[ZLSOFT]';
      Raise Err_Item;
    End If;
    Delete From 病人余额 Where 病人id = v_病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    v_金额合计 := n_预交金额;
  End If;

  --剩余部份全部用缺省结算方式结算，(小于零也不进行额外处理)
  If v_金额合计 <> 0 Then
    Update 病人预交记录
    Set 冲预交 = 冲预交 + v_金额合计, 卡类别id = 卡类别id_In, 结算卡序号 = 结算卡序号_In, 卡号 = 卡号_In, 交易流水号 = 交易流水号_In, 交易说明 = 交易说明_In,
        合作单位 = 合作单位_In, 结算序号 = n_结算序号
    
    Where 结帐id = 结帐id_In And Nvl(结算方式, 'LXH_Test') = Nvl(v_缺省, 'LXH_Test') And 记录性质 = Decode(结帐_In, 1, 2, v_记录性质);
    If Sql%RowCount = 0 Then
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 卡类别id, 结算卡序号, 卡号, 交易流水号,
         交易说明, 合作单位, 结算性质)
      Values
        (病人预交记录_Id.Nextval, Decode(结帐_In, 1, 2, v_记录性质), v_No, 1, v_病人id, v_主页id, '保险结算修正', v_缺省, v_登记时间, v_操作员编号,
         v_操作员姓名, v_金额合计, 结帐id_In, n_组id, n_结算序号, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In,
         Mod(Decode(结帐_In, 1, 2, v_记录性质), 10));
    End If;
  
    --挂号结算,分币处理(由于挂号界面没有预结算,所以在此过程中根据分币处理规则来修正)
    If v_记录性质 = 4 Then
    
      Begin
        Select a.冲预交, a.结算方式
        Into v_现金金额, v_现金结算
        From 病人预交记录 A, 结算方式 B
        Where a.结算方式 = b.名称 And b.性质 = 1 And a.结帐id = 结帐id_In And a.记录性质 = 4;
      Exception
        When Others Then
          v_现金金额 := 0;
      End;
      If Floor(Abs(v_现金金额) * 10) <> Abs(v_现金金额) * 10 Then
        --误差处理
        v_Cashcented := Zl_Cent_Money(v_现金金额, 1);
        v_误差金额   := v_现金金额 - v_Cashcented;
        If v_误差金额 <> 0 Then
          Begin
            Select 名称 Into v_误差项 From 结算方式 Where 性质 = 9;
          Exception
            When Others Then
              v_误差项 := Null;
          End;
          If v_误差项 Is Not Null Then
            --10.34之后误差数据
            Insert Into 病人预交记录
              (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 结算序号, 结算性质)
            Values
              (病人预交记录_Id.Nextval, Decode(结帐_In, 1, 2, v_记录性质), v_No, 1, v_病人id, v_主页id, '误差费', v_误差项, v_登记时间, v_操作员编号,
               v_操作员姓名, v_误差金额, 结帐id_In, n_组id, n_结算序号, Mod(Decode(结帐_In, 1, 2, v_记录性质), 10));
            Update 病人预交记录
            Set 冲预交 = v_Cashcented
            Where 结帐id = 结帐id_In And 记录性质 = 4 And 结算方式 = v_现金结算;
          Else
            --1.更新预交记录(一定存在记录)
            Update 病人预交记录
            Set 冲预交 = v_Cashcented
            Where 结算方式 = (Select 名称 From 结算方式 Where 性质 = 1 And Rownum = 1) And 结帐id = 结帐id_In;
          
            --2.生成误差费用记录(注:计算单位记录的是号别,所以不取误差项的)
            Begin
              Select a.类别, a.Id, c.Id, c.收据费目
              Into v_收费类别, v_收费细目id, v_收入项目id, v_收据费目
              From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费特定项目 D
              Where d.特定项目 = '误差项' And d.收费细目id = a.Id And a.Id = b.收费细目id And b.收入项目id = c.Id And
                    Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'));
            Exception
              When Others Then
                v_Err_Msg := '不能正确读取收费误差项的信息，请先检查该项目是否设置正确。';
                Raise Err_Item;
            End;
            If Nvl(结帐_In, 0) = 1 Then
              Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
              Select Max(序号) + 1, Max(发生时间) Into v_序号, v_发生时间 From 住院费用记录 Where 结帐id = 结帐id_In;
              n_医疗小组id := Zl_医疗小组_Get(0, v_操作员姓名, v_病人id, v_主页id, v_发生时间);
            
              Insert Into 住院费用记录
                (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 床号, 姓名, 性别, 年龄, 病人病区id, 病人科室id, 费别, 收费类别,
                 收费细目id, 计算单位, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间,
                 登记时间, 执行部门id, 执行人, 执行状态, 结帐id, 结帐金额, 操作员编号, 操作员姓名, 是否上传, 缴款组id, 医疗小组id)
                Select v_费用id, 记录性质, NO, 实际票号, 记录状态, v_序号, Null, Null, 门诊标志, 病人id, 标识号, 床号, 姓名, 性别, 年龄, 病人病区id, 病人科室id,
                       费别, v_收费类别, v_收费细目id, 计算单位, 发药窗口, 1, 1, 加班标志, 9, v_收入项目id, v_收据费目, v_误差金额, v_误差金额, v_误差金额, 记帐费用,
                       划价人, 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 结帐id_In, v_误差金额, 操作员编号, 操作员姓名, 1, 缴款组id,
                       Decode(n_医疗小组id, Null, 医疗小组id, n_医疗小组id)
                From 住院费用记录
                Where 结帐id = 结帐id_In And Rownum = 1;
            Else
              Select 病人费用记录_Id.Nextval Into v_费用id From Dual;
              Select Max(序号) + 1 Into v_序号 From 门诊费用记录 Where 结帐id = 结帐id_In;
              Insert Into 门诊费用记录
                (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id,
                 计算单位, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 记帐费用, 划价人, 开单部门id, 开单人, 发生时间, 登记时间,
                 执行部门id, 执行人, 执行状态, 费用状态, 结帐id, 结帐金额, 操作员编号, 操作员姓名, 是否上传, 缴款组id)
                Select v_费用id, 记录性质, NO, 实际票号, 记录状态, v_序号, Null, Null, 门诊标志, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 病人科室id, 费别,
                       v_收费类别, v_收费细目id, 计算单位, 发药窗口, 1, 1, 加班标志, 9, v_收入项目id, v_收据费目, v_误差金额, v_误差金额, v_误差金额, 记帐费用, 划价人,
                       开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 费用状态, 结帐id_In, v_误差金额, 操作员编号, 操作员姓名, 1, 缴款组id
                From 门诊费用记录
                Where 结帐id = 结帐id_In And Rownum = 1;
            End If;
          End If;
          --3.更新汇总表
          --只可能产生误差金额的变化.仅为了变量处理方便而用游标
        End If;
      End If;
    End If;
  End If;

  --最后再处理"人员缴款余额"(没有动冲预交那部分,所以"病人余额"的预交余额不用更新)
  For r_Del In c_Del Loop
    If r_Del.记录性质 Not In (1, 11) Then
      If Nvl(n_费用状态, 0) <> 1 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + r_Del.冲预交
        Where 收款员 = v_操作员姓名 And 性质 = 1 And 结算方式 = r_Del.结算方式;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (v_操作员姓名, r_Del.结算方式, 1, r_Del.冲预交);
        End If;
      End If;
    End If;
  End Loop;
  Delete From 人员缴款余额 Where 性质 = 1 And 收款员 = v_操作员姓名 And Nvl(余额, 0) = 0;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人结算记录_Update;
/

--89348:胡俊勇,2017-05-23,医嘱单打印
--109076:胡俊勇,2017-05-12,医嘱单打印
Create Or Replace Procedure Zl_病人医嘱打印_Insert
(
  病人id_In 病人医嘱记录.病人id%Type,
  主页id_In 病人医嘱记录.主页id%Type,
  婴儿_In   病人医嘱记录.婴儿%Type,
  期效_In   病人医嘱记录.医嘱期效%Type,
  行数_In   Number
  --功能：将病人没有打印过的医嘱插入 病人医嘱打印
  --参数：行数_In：报表医嘱单一页可以打多少行
  --      行数_In医嘱单报表的行数，通常是28行。
) Is
  n_序号       病人医嘱记录.序号%Type;
  n_医嘱id     病人医嘱记录.Id%Type;
  n_重整标记   Number;
  v_Max_Date   Date;
  d_重整       Date;
  d_Pdate      Date;
  n_换页打     Number;
  n_打重开     Number;
  n_转科       Number;
  n_页号       Number;
  n_行号       Number;
  n_位置       Number;
  n_打印模式   Number;
  n_打给药方式 Number;
  n_Lzzkhy     Number;
  n_Cnt        Number;
  v_Tmp        Varchar2(200);

  --c_Advice 取出待打印的医嘱，在打印临嘱时转科医嘱都会读取出来，后面要判断是不是要生成打印记录
  Cursor c_Advice Is
    Select 医嘱id, 顺序, 打印标记, 特殊医嘱, 换页
    From (With Printtable As (Select a.Id As 医嘱id, a.序号 As 顺序, 0 As 打印标记, Null As 特殊医嘱,
                                     Decode(a.诊疗类别, 'Z', Decode(b.操作类型, '3', 3, '4', 4, 0), 0) As 换页, a.诊疗项目id
                              From 病人医嘱记录 A, 诊疗项目目录 B
                              Where a.病人id = 病人id_In And a.主页id = 主页id_In And Nvl(a.婴儿, 0) = 婴儿_In And a.诊疗项目id = b.Id(+) And
                                    (期效_In = 0 And (a.医嘱期效 = 0 Or n_位置 In (-1, 0, 2) And a.医嘱期效 = 1 And a.诊疗类别 = 'Z' And
                                    b.操作类型 In ('5', '3', '11')) Or
                                    期效_In = 1 And a.医嘱期效 = 1 And
                                    Not (n_位置 = 0 And Nvl(a.诊疗类别, 'X') = 'Z' And Nvl(b.操作类型, 'X') In ('5', '3', '11')) Or
                                    期效_In = 1 And a.医嘱期效 = 1 And n_位置 = 0 And a.诊疗类别 = 'Z' And b.操作类型 = '3') And
                                    a.医嘱状态 Not In (-1, 2) And (n_打印模式 = 1 And a.医嘱状态 = 1 Or a.医嘱状态 <> 1) And
                                    Nvl(a.屏蔽打印, 0) = 0 And a.序号 > n_序号 And a.病人来源 = 2)
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From 病人医嘱记录 L, 诊疗项目目录 I, Printtable P
           Where l.Id = p.医嘱id And l.诊疗项目id = i.Id And
                 (l.诊疗类别 Not In ('5', '6', '7', 'E') Or l.诊疗类别 = 'E' And Nvl(i.操作类型, '0') Not In ('2', '3') Or
                 i.Id Is Null) And l.相关id Is Null
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From 病人医嘱记录 L, Printtable P
           Where l.Id = p.医嘱id And l.诊疗类别 In ('5', '6')
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From 病人医嘱记录 L, 诊疗项目目录 I, Printtable P
           Where l.Id = p.医嘱id And l.诊疗项目id = i.Id And l.诊疗类别 = 'E' And i.操作类型 = '2' And l.相关id Is Null And n_打给药方式 = 1
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From Printtable P
           Where p.诊疗项目id Is Null
           Order By 顺序);


  Cursor c_Advice_Redo Is
    Select 医嘱id, 顺序, 打印标记, 特殊医嘱, 换页
    From (With Printtable As (Select a.Id As 医嘱id, a.序号 As 顺序, 0 As 打印标记, Null As 特殊医嘱,
                                     Decode(a.诊疗类别, 'Z', Decode(b.操作类型, '3', 3, '4', 4, 0), 0) As 换页, a.诊疗项目id
                              From 病人医嘱记录 A, 诊疗项目目录 B
                              Where a.病人id = 病人id_In And a.主页id = 主页id_In And Nvl(a.婴儿, 0) = 婴儿_In And a.诊疗项目id = b.Id(+) And
                                    (期效_In = 0 And (a.医嘱期效 = 0 Or n_位置 In (-1, 0, 2) And a.医嘱期效 = 1 And a.诊疗类别 = 'Z' And
                                    b.操作类型 In ('5', '3', '11')) Or
                                    期效_In = 1 And a.医嘱期效 = 1 And
                                    Not (n_位置 = 0 And Nvl(a.诊疗类别, 'X') = 'Z' And Nvl(b.操作类型, 'X') In ('5', '3', '11'))) And
                                    a.医嘱状态 Not In (-1, 2) And (n_打印模式 = 1 And a.医嘱状态 = 1 Or a.医嘱状态 <> 1) And
                                    Nvl(a.屏蔽打印, 0) = 0 And a.序号 > n_序号 And Exists
                               (Select 1 From 病人医嘱状态 C Where a.Id = c.医嘱id And c.操作时间 >= v_Max_Date) And a.病人来源 = 2)
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From 病人医嘱记录 L, 诊疗项目目录 I, Printtable P
           Where l.Id = p.医嘱id And l.诊疗项目id = i.Id And
                 (l.诊疗类别 Not In ('5', '6', '7', 'E') Or l.诊疗类别 = 'E' And Nvl(i.操作类型, '0') Not In ('2', '3') Or
                 i.Id Is Null) And l.相关id Is Null
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From 病人医嘱记录 L, Printtable P
           Where l.Id = p.医嘱id And l.诊疗类别 In ('5', '6')
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From 病人医嘱记录 L, 诊疗项目目录 I, Printtable P
           Where l.Id = p.医嘱id And l.诊疗项目id = i.Id And l.诊疗类别 = 'E' And i.操作类型 = '2' And l.相关id Is Null And n_打给药方式 = 1
           Union All
           Select p.医嘱id, p.顺序, p.打印标记, p.特殊医嘱, p.换页
           From Printtable P
           Where p.诊疗项目id Is Null
           Order By 顺序);


  --获取下一个或用的行号和页号
  Function Getnextpos
  (
    v_页号 病人医嘱打印.页号%Type,
    v_行号 病人医嘱打印.行号%Type,
    v_行数 Number
  ) Return Varchar2 Is
    n_p Number;
    n_r Number;
  Begin
    If v_行号 = 0 Then
      n_p := 1;
      n_r := 1;
    Elsif v_行号 = v_行数 Then
      n_p := v_页号 + 1;
      n_r := 1;
    Else
      n_p := v_页号;
      n_r := v_行号 + 1;
    End If;
    Return(n_p || ',' || n_r);
  End;

Begin
  n_位置       := Zl_To_Number(Nvl(zl_GetSysParameter('转科和出院打印', 1254), 0));
  n_打印模式   := Zl_To_Number(Nvl(zl_GetSysParameter('医嘱单打印模式', 1253), 0));
  n_打给药方式 := Zl_To_Number(Nvl(zl_GetSysParameter('药品用法单独打印一行', 1254), 0));
  n_Lzzkhy     := Zl_To_Number(Nvl(zl_GetSysParameter('临嘱单转科换页', 1254), 0));
  n_换页打     := Zl_To_Number(Nvl(zl_GetSysParameter('重整和术后医嘱换页打印', 1254), 0));
  n_打重开     := Zl_To_Number(Nvl(zl_GetSysParameter('转科换页后在首行打印重开医嘱', 1254), 0));

  --判断是不是重整后打印医嘱
  If 期效_In = 1 Then
    d_重整 := To_Date('1900-01-01', 'YYYY-MM-DD');
  Else
    Select 医嘱重整时间 Into d_重整 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
    If d_重整 Is Null Then
      d_重整 := To_Date('1900-01-01', 'YYYY-MM-DD');
    End If;
  End If;
  v_Max_Date := d_重整;
  Begin
    Select 医嘱id, 打印时间, 页号, 行号
    Into n_医嘱id, d_Pdate, n_页号, n_行号
    From (Select 医嘱id, 打印时间, 页号, 行号
           From 病人医嘱打印
           Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(婴儿, 0) = 婴儿_In And 期效 = 期效_In And 医嘱id Is Not Null
           Order By 页号 Desc, 行号 Desc)
    Where Rownum < 2;
  
    Select Nvl(Max(序号), 0)
    Into n_序号
    From 病人医嘱记录
    Where ID = (Select Nvl(a.相关id, a.Id) From 病人医嘱记录 A Where a.Id = n_医嘱id);
  
    If 期效_In = 0 Then
      If d_Pdate Is Not Null Then
        If d_Pdate < d_重整 And d_重整 <> To_Date('1900-01-01', 'YYYY-MM-DD') Then
          n_重整标记 := 1;
          n_序号     := 0;
        End If;
      End If;
    End If;
  Exception
    When Others Then
      n_页号 := 0;
      n_行号 := 0;
      n_序号 := 0;
  End;

  If n_医嘱id Is Not Null Then
    Select Max(b.操作类型)
    Into v_Tmp
    From 病人医嘱记录 A, 诊疗项目目录 B
    Where a.诊疗项目id = b.Id(+) And a.Id = n_医嘱id And a.诊疗类别 = 'Z';
  End If;
  If v_Tmp = '3' Then
    n_Cnt := 3;
  Elsif v_Tmp = '4' Then
    n_Cnt := 4;
  End If;

  v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
  n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
  n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);

  If n_Cnt = 3 And n_Lzzkhy = 1 And 期效_In = 1 Then
    --临时医嘱转科换页
    If n_行号 <> 1 Then
      n_行号 := 1;
      n_页号 := n_页号 + 1;
    End If;
  Elsif 期效_In = 0 Then
    --重整，术后，转科重开，这些只针对于长期医嘱单
    --重整标记
    If n_重整标记 = 1 Then
      If n_换页打 = 1 Then
        If n_行号 <> 1 Then
          n_行号 := 1;
          n_页号 := n_页号 + 1;
        End If;
      End If;
      Insert Into 病人医嘱打印
        (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
      Values
        (-1 * Null, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, 0, Null);
      v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
      n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
      n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
    End If;
  
    --转科换页打重开字样
    If n_打重开 = 1 And n_Cnt = 3 Then
      If n_重整标记 = 1 Then
        --前面打了重整就不换页了
        If n_打重开 = 1 Then
          Insert Into 病人医嘱打印
            (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
          Values
            (-1 * Null, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, 0, 1);
          v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
          n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
          n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
        End If;
      Else
        --打重开字样
        If n_行号 <> 1 Then
          n_行号 := 1;
          n_页号 := n_页号 + 1;
        End If;
        Insert Into 病人医嘱打印
          (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
        Values
          (-1 * Null, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, 0, 1);
        v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
        n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
        n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End If;
    End If;
  
    --转科医嘱换页
    If Nvl(n_重整标记, 0) <> 1 And Nvl(n_打重开, 0) <> 1 And n_Cnt = 3 And n_换页打 = 1 Then
      If n_行号 <> 1 Then
        n_行号 := 1;
        n_页号 := n_页号 + 1;
      End If;
    End If;
  
    --术后医嘱换页
    If Nvl(n_重整标记, 0) <> 1 And n_Cnt = 4 And n_换页打 = 1 Then
      If n_行号 <> 1 Then
        n_行号 := 1;
        n_页号 := n_页号 + 1;
      End If;
    End If;
  End If;
  n_转科 := 0;

  --最近次重整后,需要打印的医嘱，考虑换页打印情况转科术后
  ---r_Print.换页 对特殊医嘱标记，4－术后，3－转科
  If v_Max_Date = To_Date('1900-01-01', 'YYYY-MM-DD') Then
    For r_Print In c_Advice Loop
      ----换页或者打医嘱重开字样
      If n_换页打 = 1 And n_转科 = 1 And 期效_In = 0 Then
        If n_打重开 = 1 Then
          --打重开字样
          If n_行号 <> 1 Then
            n_行号 := 1;
            n_页号 := n_页号 + 1;
          End If;
          Insert Into 病人医嘱打印
            (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
          Values
            (-1 * Null, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, 0, 1);
          v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
          n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
          n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
        Else
          --只是单纯换一页
          If n_行号 <> 1 Then
            n_行号 := 1;
            n_页号 := n_页号 + 1;
          End If;
        End If;
        n_转科 := 0;
      End If;
    
      If 期效_In = 1 And n_转科 = 1 And n_Lzzkhy = 1 Then
        If n_行号 <> 1 Then
          n_行号 := 1;
          n_页号 := n_页号 + 1;
        End If;
        n_转科 := 0;
      End If;
    
      If r_Print.换页 = 4 And n_换页打 = 1 Then
        --术后医嘱换页
        --如果行号为1说明已经是新的一页的第一行,否则换页
        If n_行号 <> 1 Then
          n_行号 := 1;
          n_页号 := n_页号 + 1;
        End If;
      End If;
    
      If 期效_In = 0 Or 期效_In = 1 And (n_位置 = 2 Or n_位置 = 1 Or r_Print.换页 <> 3) Then
        Insert Into 病人医嘱打印
          (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
        Values
          (r_Print.医嘱id, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, r_Print.打印标记, r_Print.特殊医嘱);
        v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
        n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
        n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End If;
      --启用了转科换页打重开字样，则插入一重开医嘱标记，此处一定换页，因为转科换页前要先打出转科医嘱，
      --这里不插入数据，只进行标记，再下一次循时才插入。如果转科医嘱是最后一条是不用打印新开字样的。
      If r_Print.换页 = 3 Then
        n_转科 := 1;
      End If;
    End Loop;
  Else
    For r_Print In c_Advice_Redo Loop
      ----换页或者打医嘱重开字样
      If n_换页打 = 1 And n_转科 = 1 And 期效_In = 0 Then
        If n_打重开 = 1 Then
          --打重开字样
          If n_行号 <> 1 Then
            n_行号 := 1;
            n_页号 := n_页号 + 1;
          End If;
          Insert Into 病人医嘱打印
            (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
          Values
            (-1 * Null, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, 0, 1);
          v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
          n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
          n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
        Else
          --只是单纯换一页
          If n_行号 <> 1 Then
            n_行号 := 1;
            n_页号 := n_页号 + 1;
          End If;
        End If;
        n_转科 := 0;
      End If;
    
      If r_Print.换页 = 4 And n_换页打 = 1 Then
        --术后医嘱换页
        --如果行号为1说明已经是新的一页的第一行,否则换页
        If n_行号 <> 1 Then
          n_行号 := 1;
          n_页号 := n_页号 + 1;
        End If;
      End If;
      Insert Into 病人医嘱打印
        (医嘱id, 页号, 行号, 行数, 病人id, 主页id, 婴儿, 期效, 打印标记, 特殊医嘱)
      Values
        (r_Print.医嘱id, n_页号, n_行号, 1, 病人id_In, 主页id_In, 婴儿_In, 期效_In, r_Print.打印标记, r_Print.特殊医嘱);
      v_Tmp  := Getnextpos(n_页号, n_行号, 行数_In);
      n_页号 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
      n_行号 := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      --启用了转科换页打重开字样，则插入一重开医嘱标记，此处一定换页，因为转科换页前要先打出转科医嘱
      --这里不插入数据，只进行标记，再下一次循时才插入。如果转科医嘱是最后一条是不用打印新开字样的。
    
      If r_Print.换页 = 3 And 期效_In = 0 Then
        n_转科 := 1;
      End If;
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱打印_Insert;
/


--95807:陈刘,2017-05-12,记录单中曲线项目增加未记说明信息录入

CREATE OR REPLACE Procedure Zl_病人护理数据_Update
(
  文件id_In   In 病人护理数据.文件id%Type,
  发生时间_In In 病人护理数据.发生时间%Type,
  记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，签名记录=5，审签记录=15
  项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
  记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容；37或38/37
  体温部位_In In 病人护理明细.体温部位%Type := Null,
  他人记录_In In Number := 1,
  数据来源_In In 病人护理明细.数据来源%Type := 0,
  审签_In     In Number := 0,
  操作员_In   In 病人护理数据.保存人%Type := Null,
  记录组号_In In 病人护理明细.记录组号%Type := Null, --适用分类汇总(一条数据对应多条相同项目的明细)
  相关序号_In In 病人护理明细.相关序号%Type := Null, --适用分类汇总(记录汇总项目关联的名称项目序号)
  未记说明_In In 病人护理明细.未记说明%Type := Null --入量导入存储医嘱ID:发送号
) Is
  Intins      Number(18);
  Int共用     Number(1);
  n_Newid     病人护理数据.Id%Type;
  n_Oldid     病人护理数据.Id%Type;
  n_行数      病人护理打印.行数%Type;
  n_Mutilbill Number(1);
  n_Syntend   Number(1);
  n_Synchro   Number(1);
  n_未记说明  Number(1);
  n_曲线      Number(1);
  n_Num       Number(18);

  n_汇总类别     病人护理数据.汇总类别%Type;
  v_科室id       部门表.Id%Type;
  v_保存人       人员表.姓名%Type;
  v_记录人       人员表.姓名%Type;
  n_文件id       病人护理数据.文件id%Type;
  n_记录id       病人护理数据.Id%Type;
  n_明细id       病人护理明细.Id%Type;
  n_来源id       病人护理明细.来源id%Type;
  v_数据来源     病人护理明细.数据来源%Type;
  n_最高版本     病人护理明细.开始版本%Type;
  n_项目性质     护理记录项目.项目性质%Type;
  n_病人id       病人护理文件.病人id%Type;
  n_主页id       病人护理文件.主页id%Type;
  n_婴儿         病人护理文件.婴儿%Type;
  d_婴儿出院时间 病人医嘱记录.开始执行时间%Type;
  d_文件开始时间 病人护理文件.开始时间%Type;
  --提取该病人当前科室所有未结束的护理文件，且文件开始时间小于等于记录发生时间的文件列表供同步数据使用
  Cursor Cur_Fileformats Is
    Select a.Id As 格式id, b.Id As 文件id, a.保留, a.子类, b.婴儿
    From 病历文件列表 A, 病人护理文件 B, 病人护理文件 C, 病人护理数据 D
    Where a.种类 = 3 And a.保留 <> 1 And a.Id = b.格式id And b.Id <> c.Id And b.结束时间 Is Null And b.开始时间 <= d.发生时间 And
          (a.通用 = 1 Or (a.通用 = 2 And b.科室id = c.科室id)) And c.病人id = b.病人id And c.主页id = b.主页id And c.婴儿 = b.婴儿 And
          c.Id = d.文件id And d.Id = n_记录id And c.Id = 文件id_In
    Order By a.编号;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --取记录ID
  Int共用     := 0;
  n_记录id    := 0;
  n_Mutilbill := 0;
  n_Syntend   := 0;
  n_未记说明  := 0;
  n_曲线      := 0;

  If 操作员_In Is Null Then
    v_保存人 := Zl_Username;
  Else
    v_保存人 := 操作员_In;
  End If;

  --如果是对应多份护理文件值为1，表示需同步其它护理文件；否则不处理文件同步
  n_Mutilbill := Zl_To_Number(zl_GetSysParameter('对应多份护理文件', 1255));
  --如果允许多份护理文件之间数据同步,则自动同步,否则不同步
  n_Syntend := Zl_To_Number(zl_GetSysParameter('允许数据同步', 1255));

  Begin
    Select ID, 汇总类别
    Into n_记录id, n_汇总类别
    From 病人护理数据
    Where 文件id = 文件id_In And 发生时间 = 发生时间_In;
  Exception
    When Others Then
      n_记录id := 0;
  End;

  --检查是不是本人的记录
  ---------------------------------------------------------------------------------------------------------------------
  If 他人记录_In = 0 And n_记录id > 0 And 审签_In = 0 Then
    v_记录人 := '';
    Begin
      Select 记录人
      Into v_记录人
      From 病人护理明细
      Where 记录id = n_记录id And 项目序号 = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
            Nvl(记录组号, 0) = Nvl(记录组号_In, 0) And 终止版本 Is Null;
    Exception
      When Others Then
        v_记录人 := '';
    End;
    If v_记录人 Is Not Null And v_记录人 <> v_保存人 Then
      v_Error := '你无权修改他人登记的护理数据！';
      Raise Err_Custom;
    End If;
  End If;

  --检查是否入科
  Select 病人id, 主页id, Nvl(婴儿, 0), 开始时间
  Into n_病人id, n_主页id, n_婴儿, d_文件开始时间
  From 病人护理文件
  Where ID = 文件id_In;
  d_婴儿出院时间 := Null;
  If n_婴儿 <> 0 Then
    Begin
      Select 开始执行时间
      Into d_婴儿出院时间
      From 病人医嘱记录 B, 诊疗项目目录 C
      Where b.诊疗项目id + 0 = c.Id And b.医嘱状态 = 8 And Nvl(b.婴儿, 0) <> 0 And c.类别 = 'Z' And
            Instr(',3,5,11,', ',' || c.操作类型 || ',', 1) > 0 And b.病人id = n_病人id And b.主页id = n_主页id And b.婴儿 = n_婴儿;
    Exception
      When Others Then
        d_婴儿出院时间 := Null;
    End;
  End If;
  If d_婴儿出院时间 Is Null Then
    v_科室id := 0;
    Begin
      Select a.科室id
      Into v_科室id
      From 病人变动记录 A, 病人护理文件 B
      Where a.科室id Is Not Null And a.病人id = b.病人id And a.主页id = b.主页id And b.Id = 文件id_In And
            (To_Date(To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI') || '59', 'YYYY-MM-DD HH24:MI:SS') >= a.开始时间 And
            (To_Date(To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI') || '00', 'YYYY-MM-DD HH24:MI:SS') < = Nvl(a.终止时间, Sysdate) Or
            a.终止时间 Is Null)) And Rownum < 2;
    Exception
      When Others Then
        v_科室id := 0;
    End;
    If v_科室id = 0 Then
      v_Error := '数据发生时间 ' || To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI:SS') || ' 不在病人有效变动时间范围内，不能操作！';
      Raise Err_Custom;
    End If;
  Else
    If 发生时间_In < d_文件开始时间 Or 发生时间_In > d_婴儿出院时间 Then
      v_Error := '数据发生时间 ' || To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI:SS') || ' 不在病人有效变动时间范围内，不能操作！';
      Raise Err_Custom;
    End If;
  End If;

  --如果数据来源<>0则退出
  n_来源id := 0;
  If n_记录id > 0 Then
    Begin
      Select 数据来源, Nvl(来源id, 0)
      Into v_数据来源, n_来源id
      From 病人护理明细
      Where 记录id = n_记录id And Nvl(项目序号, 0) = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
            Nvl(记录组号, 0) = Nvl(记录组号_In, 0);
    Exception
      When Others Then
        v_数据来源 := 0;
    End;
    If v_数据来源 > 0 And n_来源id > 0 Then
      Return;
    End If;
  End If;

  --取最高版本
  Select Nvl(Max(Nvl(a.开始版本, 1)), 0) + 1, Count(b.Id)
  Into n_最高版本, Intins
  From 病人护理明细 A, 病人护理数据 B
  Where b.Id = n_记录id And a.记录id = b.Id And Mod(a.记录类型, 10) = 5;

  --目前已经签名的数据不能修改，只有在审签模式下进行修改，即审签_In=1
  If 审签_In <> 1 And Intins > 0 Then
    v_Error := '发生时间 ' || To_Char(发生时间_In, 'YYYY-MM-DD HH24:MI:SS') || ' 所对应的数据已经签名或审签，不能继续操作！' || Chr(13) || Chr(10) ||
               '这可能是由于网络并发操作引起的，请刷新后再试！';
    Raise Err_Custom;
  End If;
  Intins := 0;

  --无内容时,要清除数据（审签回退时会自动清除审签过程中修改的数据，所以此处只需考虑普签即可）
  If 记录内容_In Is Null Then
    Begin
      Select ID
      Into n_明细id
      From 病人护理明细
      Where 记录id = n_记录id And Nvl(项目序号, 0) = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
            Nvl(记录组号, 0) = Nvl(记录组号_In, 0) And 终止版本 Is Null;
    Exception
      --无数据退出
      When Others Then
        Return;
    End;

    --查找除了本条要删除的数据，是否还存其他有效的数据，如果存在只删除本条数据，否则删除此发生时间对应的所有数据。
    Select Count(ID)
    Into Intins
    From 病人护理明细
    Where 记录id = n_记录id And Mod(记录类型, 10) <> 5 And 终止版本 Is Null And ID <> n_明细id;
    If Intins = 0 Then
      Delete From 病人护理明细 Where 记录id = n_记录id;
    Else
      Delete From 病人护理明细 Where ID = n_明细id;
    End If;

    Delete From 病人护理数据 A
    Where a.Id = n_记录id And Not Exists (Select 1 From 病人护理明细 B Where b.记录id = a.Id);

    --如果是删除签名后修改产生的最后一条数据,则应将签名记录的终止版本清为空
    Begin
      Select 1
      Into Intins
      From 病人护理明细
      Where 开始版本 = n_最高版本 And 终止版本 Is Null And 记录类型 = 1 And 记录id = n_记录id;
    Exception
      When Others Then
        Intins := 0;
    End;
    If Intins = 0 Then
      Update 病人护理明细 Set 终止版本 = Null Where 记录类型 = 5 And 开始版本 = n_最高版本 - 1 And 记录id = n_记录id;
    End If;
    If Nvl(n_汇总类别, 0) <> 0 Then
      Return;
    End If;

    --############
    --清除共用数据
    --############
    For Rsdel In (Select Distinct 记录id From 病人护理明细 Where 来源id = n_明细id) Loop

      Delete 病人护理明细 Where 来源id = n_明细id And 记录id = Rsdel.记录id;
      --删除对应的打印数据
      Begin
        Select Count(*) Into Intins From 病人护理明细 Where 记录id = Rsdel.记录id;
      Exception
        When Others Then
          Intins := 0;
      End;
      If Intins = 0 Then
        --提取清除数据对应的文件ID
        Begin
          Select b.Id, a.保留
          Into n_文件id, Intins
          From 病历文件列表 A, 病人护理文件 B, 病人护理数据 C
          Where a.Id = b.格式id And b.Id = c.文件id And c.Id = Rsdel.记录id;
        Exception
          When Others Then
            n_文件id := 0;
        End;
        Delete 病人护理数据 Where ID = Rsdel.记录id;
        If Intins <> -1 Then
          Zl_病人护理打印_Update(n_文件id, 发生时间_In, 1, 1);
        End If;
      End If;
    End Loop;
  Else
    --检查录入的项目是否属于该记录单
    Begin
      Select 1
      Into Intins
      From (Select b.项目序号
             From 病历文件结构 A, 护理记录项目 B
             Where a.要素名称 = b.项目名称 And b.项目序号 = 项目序号_In And
                   父id = (Select b.Id
                          From 病人护理文件 A, 病历文件结构 B
                          Where a.Id = 文件id_In And a.格式id = b.文件id And b.父id Is Null And b.对象序号 = 4)
             Union
             Select 项目序号
             From 护理记录项目
             Where 项目性质 = 2 And 项目序号 = 项目序号_In);
    Exception
      When Others Then
        Intins := 0;
    End;
    If Intins = 0 Then
      Return;
    End If;
    If n_记录id = 0 Then
      Select 病人护理数据_Id.Nextval Into n_记录id From Dual;

      Insert Into 病人护理数据
        (ID, 文件id, 发生时间, 最后版本, 保存人, 保存时间)
      Values
        (n_记录id, 文件id_In, 发生时间_In, n_最高版本, v_保存人, Sysdate);
    End If;

    --插入本次登记的病人护理明细
    Update 病人护理明细
    Set 记录内容 = 记录内容_In, 数据来源 = 数据来源_In, 未记说明 = 未记说明_In, 记录人 = v_保存人, 记录时间 = Sysdate
    Where 记录id = n_记录id And 项目序号 = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
          Nvl(记录组号, 0) = Nvl(记录组号_In, 0) And 开始版本 = n_最高版本 And 终止版本 Is Null;
    If Sql%RowCount = 0 Then
      Select 病人护理明细_Id.Nextval Into n_明细id From Dual;
      Insert Into 病人护理明细
        (ID, 记录id, 记录类型, 项目分组, 项目id, 相关序号, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 记录组号, 体温部位, 数据来源, 共用, 未记说明, 开始版本, 终止版本,
         记录人, 记录时间)
        Select n_明细id, n_记录id, 记录类型_In, a.分组名, a.项目id, 相关序号_In, a.项目序号, Upper(a.项目名称), a.项目类型, 记录内容_In, a.项目单位, 0,
               记录组号_In, 体温部位_In, 数据来源_In, Nvl(b.共用, 0), 未记说明_In, n_最高版本, Null, v_保存人, Sysdate
        From 护理记录项目 A, 病人护理明细 B
        Where a.项目序号 = b.项目序号(+) And b.终止版本(+) Is Null And b.记录id(+) = n_记录id And a.项目序号 = 项目序号_In And Rownum < 2;
    End If;
    Select ID
    Into n_明细id
    From 病人护理明细
    Where 记录id = n_记录id And 项目序号 = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
          Nvl(记录组号, 0) = Nvl(记录组号_In, 0) And 开始版本 = n_最高版本 And 终止版本 Is Null;
    --填写历史数据及签名记录的终止版本
    Update 病人护理明细
    Set 终止版本 = n_最高版本
    Where 记录id = n_记录id And ((Mod(记录类型, 10) <> 5 And 项目序号 = 项目序号_In And Nvl(体温部位, 'TWBW') = Nvl(体温部位_In, 'TWBW') And
          Nvl(记录组号, 0) = Nvl(记录组号_In, 0)) Or 记录类型 = Decode(审签_In, 1, 15, 5)) And 开始版本 <= n_最高版本 - 1 And 终止版本 Is Null;

    --如果是未签名数据，最后修改操作员做为该记录的保存人更新
    If n_最高版本 = 1 Then
      Update 病人护理数据 Set 保存人 = v_保存人, 保存时间 = Sysdate Where ID = n_记录id;
    End If;

    If Nvl(n_汇总类别, 0) <> 0 Then
      Return;
    End If;

    --############
    --同步共用数据
    --############
    --1\先处理体温单（一个病人始终只存在一份有效的体温单文件）
    --如果体温表存在相同发生时间的数据，使用它的ID
    --CL,2015-12-30,记录单同步文字项目到体温单
    For Row_Format In Cur_Fileformats Loop
      If Row_Format.保留 = -1 Then
        If Row_Format.子类 = '1' Then
          Begin
            Select 1, h.项目性质
            Into Intins, n_项目性质
            From (Select To_Char(f.项目序号) As 序号, g.项目性质
                   From 体温记录项目 F, 护理记录项目 G
                   Where f.项目序号 = g.项目序号 And g.项目性质 = 2 And
                         (g.适用科室 = 1 Or
                         (g.适用科室 = 2 And Exists
                          (Select 1 From 护理适用科室 D Where g.项目序号 = d.项目序号 And d.科室id = v_科室id))) And Nvl(g.应用方式, 0) <> 0 And
                         (Nvl(g.适用病人, 0) = 0 Or Nvl(g.适用病人, 0) = Decode(Nvl(Row_Format.婴儿, 0), 0, 1, 2))
                   Union All
                   Select b.内容文本 As 序号, 1 As 项目性质
                   From 病历文件结构 A, 病历文件结构 B
                   Where a.文件id = Row_Format.格式id And a.父id Is Null And a.对象序号 In (2, 3) And b.父id = a.Id) H
            Where Instr(',' || h.序号 || ',', ',' || 项目序号_In || ',', 1) > 0;
          Exception
            When Others Then
              Intins := 0;
          End;
        Else
          Begin
            Select 1, g.项目性质
            Into Intins, n_项目性质
            From 体温记录项目 F, 护理记录项目 G
            Where f.项目序号 = g.项目序号 And Nvl(g.应用方式, 0) <> 0 And g.护理等级 >= 0 And
                  (Nvl(g.适用病人, 0) = 0 Or Nvl(g.适用病人, 0) = Decode(Nvl(Row_Format.婴儿, 0), 0, 1, 2)) And f.项目序号 = 项目序号_In And
                  (g.适用科室 = 1 Or (g.适用科室 = 2 And Exists
                   (Select 1 From 护理适用科室 D Where g.项目序号 = d.项目序号 And d.科室id = v_科室id)));
          Exception
            When Others Then
              Intins := 0;
          End;
        End If;

        If Intins > 0 Then
          --LPF,2013-01-23,检查此项目是否需要进行同步(对于以前已经同步过的数据，为了保证记录单和体温单数据一直将不根据此函数判断。)
          n_Synchro := Zl_Temperatureprogram(文件id_In, v_科室id, 项目序号_In, 发生时间_In);
          Begin
            Select b.Id
            Into n_Newid
            From 病人护理文件 A, 病人护理数据 B
            Where a.Id = Row_Format.文件id And b.文件id = a.Id And b.发生时间 = 发生时间_In;
          Exception
            When Others Then
              n_Newid := 0;
          End;
          n_Oldid := n_Newid;
          If n_Newid = 0 And n_Synchro = 1 Then
            Select 病人护理数据_Id.Nextval Into n_Newid From Dual;
            --产生体温单主记录
            Insert Into 病人护理数据
              (ID, 文件id, 保存人, 保存时间, 发生时间, 最后版本)
            Values
              (n_Newid, Row_Format.文件id, v_保存人, Sysdate, 发生时间_In, 1);
          End If;

          Begin
            Select To_Number(记录内容_In) Into n_Num From Dual;
          Exception
            When Invalid_Number Then
              Begin
                Select 1 Into n_曲线 From 体温记录项目 Where 项目序号 = 项目序号_In And 记录法 = 1;
              Exception
                When Others Then
                  n_曲线 := 0;
              End;
              Begin
                Select 1 Into n_未记说明 From 常用体温说明 Where 名称 = 记录内容_In;
              Exception
                When Others Then
                  n_未记说明 := 0;
              End;
          End;

          If n_Newid > 0 Then
            --插入未同步的体温单数据(仍然要联接多表查询)
            Select Count(*)
            Into v_数据来源
            From 病人护理明细
            Where 记录id = n_Newid And 项目序号 = 项目序号_In And
                  Decode(n_项目性质, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无');
            If v_数据来源 = 0 Then
              --说明在同步开始已经进行过检查
              If n_Synchro = 1 Then
                --没有检查此项目是否需要同步
                If n_曲线 = 1 And n_未记说明 = 1 Then
                  Insert Into 病人护理明细
                    (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 体温部位, 数据来源, 来源id, 未记说明, 开始版本, 终止版本,
                     记录人, 记录时间, 记录组号)
                    Select 病人护理明细_Id.Nextval, n_Newid, b.记录类型, b.项目分组, b.项目id, b.项目序号, b.项目名称, b.项目类型, Null, b.项目单位,
                           b.记录标记, b.体温部位, 1, b.Id, b.记录内容, 1, Null, b.记录人, Sysdate, 1
                    From (Select 项目序号_In As 项目序号, Nvl(体温部位_In, '无') As 体温部位
                           From Dual
                           Minus
                           Select f.项目序号, Decode(Nvl(f.项目性质, 1), 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无'))
                           From 病人护理明细 E, 护理记录项目 F
                           Where e.记录id = n_Newid And e.项目序号 = f.项目序号) A, 病人护理明细 B
                    Where a.项目序号 = b.项目序号 And b.记录id = n_记录id And b.Id = n_明细id;
                  If Sql%RowCount > 0 Then
                    Int共用 := 1;
                  End If;
                Else
                  Insert Into 病人护理明细
                    (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 体温部位, 数据来源, 来源id, 开始版本, 终止版本, 记录人,
                     记录时间, 记录组号)
                    Select 病人护理明细_Id.Nextval, n_Newid, b.记录类型, b.项目分组, b.项目id, b.项目序号, b.项目名称, b.项目类型, b.记录内容, b.项目单位,
                           b.记录标记, b.体温部位, 1, b.Id, 1, Null, b.记录人, Sysdate, 1
                    From (Select 项目序号_In As 项目序号, Nvl(体温部位_In, '无') As 体温部位
                           From Dual
                           Minus
                           Select f.项目序号, Decode(Nvl(f.项目性质, 1), 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无'))
                           From 病人护理明细 E, 护理记录项目 F
                           Where e.记录id = n_Newid And e.项目序号 = f.项目序号) A, 病人护理明细 B
                    Where a.项目序号 = b.项目序号 And b.记录id = n_记录id And b.Id = n_明细id;
                  If Sql%RowCount > 0 Then
                    Int共用 := 1;
                  End If;
                end if;
              End If;
            Else
              If n_曲线 = 1 And n_未记说明 = 1 Then
                Update 病人护理明细
                Set 未记说明 = 记录内容_In, 来源id = n_明细id, 记录内容 = Null
                Where 记录id = n_Newid And 项目序号 = 项目序号_In And
                      Decode(n_项目性质, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无') And 数据来源 > 0;
                If Sql%RowCount > 0 Then
                  Int共用 := 1;
                End If;
              Else
                Update 病人护理明细
                Set 记录内容 = 记录内容_In, 来源id = n_明细id
                Where 记录id = n_Newid And 项目序号 = 项目序号_In And
                      Decode(n_项目性质, 2, Nvl(体温部位, '无'), Nvl(体温部位_In, '无')) = Nvl(体温部位_In, '无') And 数据来源 > 0;
                If Sql%RowCount > 0 Then
                  Int共用 := 1;
                End If;
              End If;
            End If;
          End If;
        End If;
        --2\再循环处理记录单
      Else
        If n_Mutilbill = 1 And n_Syntend = 1 Then
          --提取记录单与当前记录单存在重叠的且有数据的固定项目
          Select Count(*)
          Into Intins
          From (Select b.项目序号
                 From 病历文件结构 A, 护理记录项目 B
                 Where a.要素名称 = b.项目名称 And b.项目表示 In (0, 4, 5) And
                       父id =
                       (Select ID From 病历文件结构 Where 文件id = Row_Format.格式id And 父id Is Null And 对象序号 = 4)
                 Intersect
                 Select b.项目序号
                 From 病历文件结构 A, 护理记录项目 B, 病人护理文件 C, 病人护理数据 D, 病人护理明细 G
                 Where c.Id = d.文件id And a.文件id = c.格式id And d.Id = g.记录id And d.Id = n_记录id And g.Id = n_明细id And
                       b.项目序号 = g.项目序号 And b.项目表示 In (0, 4, 5) And g.记录类型 = 1 And a.要素名称 = b.项目名称 And
                       a.父id = (Select ID From 病历文件结构 E Where e.文件id = c.格式id And 父id Is Null And 对象序号 = 4));

          If Intins > 0 Then
            n_Newid := 0;
            --可能指定文件已经存在相同发生时间的数据，直接用它的ID即可
            Begin
              Select c.Id
              Into n_Newid
              From 病人护理数据 C
              Where c.文件id = Row_Format.文件id And c.发生时间 = 发生时间_In;
            Exception
              When Others Then
                n_Newid := 0;
            End;

            If n_Newid = 0 Then
              --产生记录单主记录
              Select 病人护理数据_Id.Nextval Into n_Newid From Dual;

              Insert Into 病人护理数据
                (ID, 文件id, 保存人, 保存时间, 发生时间, 最后版本)
                Select n_Newid, Row_Format.文件id, c.保存人, c.保存时间, c.发生时间, 1
                From 病人护理数据 C
                Where c.Id = n_记录id;
            End If;

            If n_Newid > 0 Then
              --插入未同步的记录单数据
              Select Count(*) Into v_数据来源 From 病人护理明细 Where 记录id = n_Newid And 项目序号 = 项目序号_In;
              If v_数据来源 = 0 Then
                Insert Into 病人护理明细
                  (ID, 记录id, 记录类型, 项目分组, 项目id, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记, 体温部位, 数据来源, 来源id, 未记说明, 开始版本, 终止版本,
                   记录人, 记录时间)
                  Select 病人护理明细_Id.Nextval, n_Newid, b.记录类型, b.项目分组, b.项目id, b.项目序号, b.项目名称, b.项目类型, b.记录内容, b.项目单位,
                         b.记录标记, b.体温部位, 1, b.Id, b.未记说明, 1, Null, b.记录人, Sysdate
                  From (Select b.项目序号
                         From 病历文件结构 A, 护理记录项目 B
                         Where a.要素名称 = b.项目名称 And b.项目表示 In (0, 4, 5) And
                               父id = (Select ID
                                      From 病历文件结构
                                      Where 文件id = Row_Format.格式id And 父id Is Null And 对象序号 = 4)
                         Intersect
                         Select b.项目序号
                         From 病历文件结构 A, 护理记录项目 B, 病人护理文件 C, 病人护理数据 D, 病人护理明细 G
                         Where c.Id = d.文件id And a.文件id = c.格式id And d.Id = g.记录id And d.Id = n_记录id And g.Id = n_明细id And
                               b.项目序号 = g.项目序号 And b.项目表示 In (0, 4, 5) And g.记录类型 = 1 And a.要素名称 = b.项目名称 And
                               a.父id =
                               (Select ID From 病历文件结构 E Where e.文件id = c.格式id And 父id Is Null And 对象序号 = 4)) A, 病人护理明细 B
                  Where a.项目序号 = b.项目序号 And b.记录id = n_记录id And b.Id = n_明细id;
                If Sql%RowCount > 0 Then
                  Int共用 := 1;
                  --原行数不要动
                  Begin
                    Select 行数 Into n_行数 From 病人护理打印 Where 文件id = Row_Format.文件id And 记录id = n_Newid;
                  Exception
                    When Others Then
                      n_行数 := 1;
                  End;
                  Zl_病人护理打印_Update(Row_Format.文件id, 发生时间_In, n_行数, 0);
                End If;
              Else
                Update 病人护理明细
                Set 记录内容 = 记录内容_In, 未记说明 = 未记说明_In, 来源id = n_明细id
                Where 记录id = n_Newid And 项目序号 = 项目序号_In And 数据来源 > 0;
                If Sql%RowCount > 0 Then
                  Int共用 := 1;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    End Loop;

    If Int共用 = 1 Then
      Update 病人护理明细 Set 共用 = 1 Where ID = n_明细id;
      --将历史数据的共用标志设置为NULL
      Update 病人护理明细 Set 共用 = Null Where 记录id = n_记录id And 项目序号 = 项目序号_In And ID <> n_明细id;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人护理数据_Update;
/


--108753:冉俊明,2017-05-09,修正在根据票据分配规则自动分配票据明细数据时，未按照性能规范加cardinality关键词的问题
Create Or Replace Procedure Zl_Custom_Invoice_Autoallot
(
  操作类型_In       Number,
  模拟计算_In       Number,
  票种_In           票据使用明细.票种%Type,
  领用id_In         票据使用明细.领用id%Type,
  病人id_In         门诊费用记录.病人id%Type,
  Nos_In            Varchar2,
  起始发票号_In     门诊费用记录.实际票号%Type,
  使用人_In         票据使用明细.使用人%Type,
  使用时间_In       票据使用明细.使用时间%Type,
  发票号_In         In Out Varchar2,
  发票张数_In       Out Number,
  按病人补打票据_In Number := 0,
  打印id_In         票据使用明细.打印id%Type := Null,
  Print_Nos_In      t_Strlist := Null
) As
  -------------------------------------------------------------------------------------------------------------
  --功能：根据票据分配规则,自动分配票据明细数据
  --入参：
  --     操作类型_In :1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
  --     模拟计算_IN :0-不进行模拟计算;1-进行模拟计算,模拟计算时不保存数据
  --     票种_IN     :1-门诊收费;暂无其他类型票据
  --     病人ID_IN   :病人ID,如果Nos和发票号_In为空时,表示针对该病人的所有未打印的票据进行打印
  --     NOs_IN      :单据号,多个用逗号分离,最多有400张单据,格式为:A00001,A00002.....
  --     退费NOs:退费所涉及的单据
  --     启始发票号_IN:重打票据或发出票据的启始票号;
  --     发票号_In   :可以为多个,用逗号分隔,当操作类型为3-重打票据时和4-退费回收票据有效,表示本次需要回收的票据
  --     打印id_In:按病人补打示据时，传入了相关的打印ID,以外面传入的打印ID为准
  --     按病人补打票据_In：1-表示按病人补打票据,不分结算次数
  --     Print_Nos_in:当前的所涉及的收费单据号，主要是控制超过varchar2的大小限制，主要是按病人补打发票时会出现超长的情况，因此通过集合传入,主要是歉容用，本次打印单据>3000时，Nos_in传入值为空。
  --出参:
  --     发票号_In   :可以为多个,用逗号分隔,当操作类型为3-重打票据时和4-退费回收票据有效,表示重打或退费重新发出的票据
  --     发票张数_IN :返回本次收费所需要的发票张数
  -------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_分单据打印 Number(3);
  n_执行科室   Number(3);
  n_收据费目   Number(3);
  n_汇总条件   Number(3);
  n_收费细目   Number(3);

  --------------------------------------------------------
  --定义内部票据处理的数据集
  Type Ty_Rec_Bill Is Record(
    票号     票据打印明细.票号%Type,
    NO       票据打印明细.No%Type,
    序号     票据打印明细.序号%Type,
    关联序号 票据打印明细.关联票号序号%Type,
    修改标志 Number(1));
  Type Ty_Tb_Bill Is Table Of Ty_Rec_Bill;
  c_Invoce Ty_Tb_Bill := Ty_Tb_Bill();
  --------------------------------------------------------
  --按元素1,元素2,元素3,元素4,分别统计各单据的序号
  Type Ty_Rec_No Is Record(
    NO   门诊费用记录.No%Type,
    序号 Varchar2(1000));
  Type Ty_Tb_No Is Table Of Ty_Rec_No;
  c_No Ty_Tb_No := Ty_Tb_No();
  --------------------------------------------------------
  Cursor c_Fact Is
    Select 前缀文本, 剩余数量, 开始号码, 终止号码, 当前号码 From 票据领用记录 Where ID = Nvl(领用id_In, 0);
  r_Factrow c_Fact%RowType;

  v_Nos        Varchar2(4000);
  v_发票号     票据打印明细.票号%Type;
  v_开始发票号 票据打印明细.票号%Type;
  v_当前发票号 票据打印明细.票号%Type;
  v_回收票据号 Varchar2(4000);
  n_Find       Number(3);

  n_元素1_Count Number(3);
  n_元素2_Count Number(3);
  n_元素3_Count Number(3);
  n_元素4_Count Number(3);

  v_元素1    门诊费用记录.No%Type;
  n_元素2    门诊费用记录.执行部门id%Type;
  v_元素3    门诊费用记录.收据费目%Type;
  n_元素4    门诊费用记录.收费细目id%Type;
  v_发票信息 Varchar2(4000);
  n_误差项   Number(1);
  n_打印id   票据使用明细.打印id%Type;
  n_使用id   票据使用明细.Id%Type;
  n_返回数   Number(18);
  n_关联序号 Number(18);
  r_单据号   t_Strlist := t_Strlist();
  r_单据序号 t_Strlist := t_Strlist();
  l_使用id   t_Numlist := t_Numlist();
  l_关联序号 t_Numlist := t_Numlist();

  v_打印内容 Varchar2(4000);
  v_Temp     Varchar2(4000);
  Procedure Invoice_Split_Notgroup
  (
    Print_Nos        t_Strlist,
    回收发票_In      Varchar2,
    本次打印发票_Out In Out Varchar2,
    本次发票张数_Out In Out Number,
    Invoce_Out       In Out Ty_Tb_Bill
  ) As
    ----------------------------------------------------------------------------
    --入参:
    --   收费收费NOs_IN:本次需要处理的发票所涉及的单据,多个用逗号分离
    --   回收发票_IN-退费时有效,多个用逗号分离，表示本次需要回收的发票号 
    --出参:
    -- 本次打印发票_Out-本次需要的发票号,多个用逗号分离
    -- 本次发票张数_Out-本次需要的发票数
    -- Invoce_Out:本次返回的发票号与单据的对应关系
    n_Count Number(18);
    n_分页  Number(18);
  
    Cursor Cr_Bill Is
      Select NO As 元素1, 执行部门id As 元素2, 收据费目 As 元素3, NO As 元素4, NO As 单据, 序号, 0 As 个数
      From 门诊费用记录
      Where Rownum <= 1;
    c_Bill Cr_Bill%RowType;
    --------------------------------------------------------------------------------------------
    --根据相关传入的数据,取对应的数据集
    Type Ty_费用明细 Is Ref Cursor;
    c_费用明细 Ty_费用明细; --游标变量 
  
  Begin
    --按单据分配票据
    If 操作类型_In = 3 Or 操作类型_In = 4 Then
      --1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
      Open c_费用明细 For
        With c_费用 As
         (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, '-', a.No) As 元素4, a.No As 单据,
                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
          From 门诊费用记录 A,
               (Select /*+cardinality(j,10)*/
                  NO, 序号
                 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                 Where m.票号 = j.Column_Value) B
          Where Mod(a.记录性质, 10) = 1 And a.No = b.No And Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
          Having Sum(Nvl(a.实收金额, 0)) <> 0)
        Select 元素1, 元素2, 元素3, 元素4, 单据, 序号, Count(*) As 个数
        From c_费用
        Group By 元素1, 元素2, 元素3, 元素4, 单据, 序号
        Order By 元素1, 元素2, 元素3, 元素4, 单据, 序号;
    Else
      Open c_费用明细 For
        With c_费用 As
         (Select /*+cardinality(b,10)*/
           Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
           Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, '-', a.No) As 元素4, a.No As 单据,
           Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
          From 门诊费用记录 A, Table(Print_Nos) B
          Where Mod(a.记录性质, 10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
          Having Sum(Nvl(a.实收金额, 0)) <> 0)
        Select 元素1, 元素2, 元素3, 元素4, 单据, 序号, Count(*) As 个数
        From c_费用
        Group By 元素1, 元素2, 元素3, 元素4, 单据, 序号
        Order By 元素1, 元素2, 元素3, 元素4, 单据, 序号;
    End If;
  
    v_元素1          := '+';
    n_元素2          := 0;
    v_元素3          := '+';
    n_元素4          := 0;
    n_元素1_Count    := 0;
    n_元素2_Count    := 0;
    n_元素3_Count    := 0;
    n_元素4_Count    := 0;
    本次发票张数_Out := 0;
    If n_汇总条件 <> 0 Then
      n_关联序号 := 1;
    Else
      n_关联序号 := 0;
    End If;
    n_Count := 0;
    c_No.Delete;
    Loop
      Fetch c_费用明细
        Into c_Bill;
      Exit When c_费用明细%NotFound;
      n_Count := 1;
    
      n_分页 := 0;
      If (v_元素1 <> c_Bill.元素1) Or (n_元素2 <> c_Bill.元素2 And n_元素2_Count >= n_执行科室 And n_执行科室 <> 0) Or
         (v_元素3 <> c_Bill.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Or (n_元素4_Count >= n_收费细目 And n_收费细目 <> 0) Then
      
        If (v_元素1 <> '+' Or n_元素2 <> 0 Or v_元素3 <> '+' Or n_元素4 <> 0) Then
          n_分页 := 1;
        End If;
        n_元素2_Count := 0;
        n_元素3_Count := 0;
        n_元素4_Count := 0;
        n_元素1_Count := 0;
        v_元素1       := '+';
        n_元素2       := 0;
        v_元素3       := '+';
      End If;
    
      If n_分页 = 1 Then
        --分页:计算发票号及相关的
        For I In 1 .. c_No.Count Loop
          Invoce_Out.Extend;
          Invoce_Out(Invoce_Out.Count).票号 := v_发票号;
          Invoce_Out(Invoce_Out.Count).No := c_No(I).No;
          Invoce_Out(Invoce_Out.Count).序号 := Case
                                               When Instr(c_No(I).序号, ',') > 0 Then
                                                Substr(c_No(I).序号, 2)
                                               Else
                                                c_No(I).序号
                                             End;
          Invoce_Out(Invoce_Out.Count).关联序号 := n_关联序号;
        End Loop;
      
        本次发票张数_Out := 本次发票张数_Out + 1;
        本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
        v_发票号         := Zl_Incstr(v_发票号);
        c_No.Delete;
      End If;
      If (v_元素1 <> c_Bill.元素1) Then
        n_元素1_Count := n_元素1_Count + 1;
        v_元素1       := c_Bill.元素1;
      End If;
      If (n_元素2 <> c_Bill.元素2) Then
        n_元素2_Count := n_元素2_Count + 1;
        n_元素2       := c_Bill.元素2;
      End If;
      If (v_元素3 <> c_Bill.元素3) Then
        n_元素3_Count := n_元素3_Count + 1;
        v_元素3       := c_Bill.元素3;
      End If;
      If n_收费细目 <> 0 Then
        n_元素4_Count := n_元素4_Count + 1;
      End If;
    
      -------------------------------------------
      --分配单据号及序号
      n_Find := 0;
      For J In 1 .. c_No.Count Loop
        If c_No(J).No = c_Bill.单据 Then
          --单据号相同,将序号合并
          c_No(J).序号 := c_No(J).序号 || ',' || c_Bill.序号;
          n_Find := 1;
          Exit;
        End If;
      End Loop;
      If n_Find = 0 Then
        c_No.Extend;
        c_No(c_No.Count).No := c_Bill.单据;
        c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_Bill.序号;
      End If;
    End Loop;
  
    --是否有发票数据
    If n_Count >= 1 Then
      --最后一个发票分配
      本次发票张数_Out := 本次发票张数_Out + 1;
      本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
    Else
      本次发票张数_Out := 0;
      本次打印发票_Out := '';
    End If;
    If c_No.Count <> 0 Then
      For I In 1 .. c_No.Count Loop
        Invoce_Out.Extend;
        Invoce_Out(Invoce_Out.Count).票号 := v_发票号;
        Invoce_Out(Invoce_Out.Count).No := c_No(I).No;
        If Instr(c_No(I).序号, ',') > 0 Then
          c_No(I).序号 := Substr(c_No(I).序号, 2);
        End If;
        Invoce_Out(Invoce_Out.Count).序号 := c_No(I).序号;
        Invoce_Out(Invoce_Out.Count).关联序号 := n_关联序号;
      End Loop;
      c_No.Delete;
    End If;
    If Instr(Nvl(本次打印发票_Out, '-'), ',') > 0 Then
      本次打印发票_Out := Substr(本次打印发票_Out, 2);
    End If;
  End Invoice_Split_Notgroup;

Begin
  --处理票据数据
  If 票种_In <> 1 Then
    --暂不支持其他,只支持门诊收费
    Return;
  End If;
  v_发票号 := 起始发票号_In;
  v_Nos    := Nos_In;
  -----------------------------------------------------------------------------------------------------------------------------
  --一、获取发票分配的相关规则
  --**开始
  --1.确定是否分单据分配票号,缺省不按单据分号
  n_分单据打印 := 0;
  --2.确定是否按执行科室分单据号,缺省为按1个执行科室分号
  n_执行科室 := 1;

  --3.确定是否按收据费目分单据号,缺省为按3个收据费目分号
  n_收据费目 := 3;

  --4.确定是否按收费细目分单据号,缺省为不按收费细目分号
  n_收费细目 := 0;

  --5.决定是否首页汇总,缺省为不汇总
  n_汇总条件 := 0;

  v_回收票据号 := 发票号_In;
  发票张数_In  := 0;
  --**结束
  If Nvl(按病人补打票据_In, 0) <> 0 Then
    --按病人补打票据时，只按收费费目打印
    n_执行科室 := 0;
  End If;

  -----------------------------------------------------------------------------------------------------------------------------
  --二、进行发票分配
  Invoice_Split_Notgroup(Print_Nos_In, 发票号_In, v_发票信息, 发票张数_In, c_Invoce);

  -----------------------------------------------------------------------------------------------------------------------------
  --*****************************************************************************************************************************
  --注意:
  --以下代码，不轻意更改,在上面的代码中需要确定两个变量的值:一是v_发票信息;二是c_Invoce集合中的值
  --  v_发票信息:本次所涉及的发票信息,多个用逗号分离,最好按升序排序
  --  c_Invoce:为集合数据，为发票号和单据的对应关系

  发票号_In := v_发票信息;
  If 模拟计算_In = 1 Then
    --模拟计算,只返回票据张数和使用的票据号,直接退出
    Return;
  End If;
  -----------------------------------------------------------------------------------------------------------------------------
  --四、退费时，需要先处理回收发票
  v_开始发票号 := Null;
  v_当前发票号 := Null;
  --1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
  If 操作类型_In = 3 Or 操作类型_In = 4 Then
    --收回票据
    Select 使用id Bulk Collect
    Into l_使用id
    From (Select /*+cardinality(j,10)*/
           Distinct b.使用id
           From 票据使用明细 A, 票据打印明细 B, Table(f_Str2list(v_回收票据号)) J
           Where a.Id = b.使用id And b.票号 = j.Column_Value And Nvl(b.票种, 0) = 1);
  
    --插入回收记录
    Forall I In 1 .. l_使用id.Count
      Insert Into 票据使用明细
        (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
        Select 票据使用明细_Id.Nextval, 票种, 号码, 2, Decode(操作类型_In, 3, 4, 2), 领用id, 打印id, 使用人_In, 使用时间_In
        From 票据使用明细
        Where ID = l_使用id(I);
    Forall I In 1 .. l_使用id.Count
      Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I);
  End If;

  If c_Invoce.Count = 0 Then
    --无发票数据,则直接返回,退费时，表示只收回票据
    Return;
  End If;

  -----------------------------------------------------------------------------------------------------------------------------
  --五、重新处理发出的票据(含退费重新发出的票据处理)
  If 起始发票号_In Is Null Then
    v_Err_Msg := '未传入起始发票号,不能进行票据分配处理';
    Raise Err_Item;
  End If;

  If Nvl(领用id_In, 0) <> 0 Then
    Open c_Fact;
    Fetch c_Fact
      Into r_Factrow;
    If c_Fact%RowCount = 0 Then
      v_Err_Msg := '无效的票据领用批次，无法完成收费票据分配操作。';
      Close c_Fact;
      Raise Err_Item;
    Elsif Nvl(r_Factrow.剩余数量, 0) < 发票张数_In Then
      v_Err_Msg := '当前批次的剩余数量不足' || 发票张数_In || '张，无法完成收费票据分配操作。';
      Close c_Fact;
      Raise Err_Item;
    End If;
  End If;

  --1.实际处理票据信息
  If Nvl(n_分单据打印, 0) <> 1 Or Nvl(按病人补打票据_In, 0) = 1 Then
    --不分单据打印时,表示一次打印,打印ID填成一致
    n_打印id := 打印id_In;
    If Nvl(n_打印id, 0) = 0 Then
      Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
    End If;
  End If;

  发票张数_In := 0;
  v_打印内容  := '';
  For c_Invoce_No In (Select Column_Value As 发票号 From Table(f_Str2list(v_发票信息)) Order By 发票号) Loop
    --2.检查票据范围是否正确
    If Nvl(领用id_In, 0) <> 0 Then
      If Not (Upper(c_Invoce_No.发票号) >= Upper(r_Factrow.开始号码) And Upper(c_Invoce_No.发票号) <= Upper(r_Factrow.终止号码) And
          Length(c_Invoce_No.发票号) = Length(r_Factrow.终止号码)) Then
        v_Err_Msg := '该单据需要打印多张票据,但票据号"' || c_Invoce_No.发票号 || '"超出票据领用的号码范围！';
        Close c_Fact;
        Raise Err_Item;
      End If;
    End If;
  
    --3.处理票据打印明细
    r_单据号.Delete;
    r_单据序号.Delete;
    l_关联序号.Delete;
  
    Select 票据使用明细_Id.Nextval Into n_使用id From Dual;
  
    n_关联序号 := 0;
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).票号 = c_Invoce_No.发票号 Then
        n_关联序号 := c_Invoce(I).关联序号;
        Exit;
      End If;
    End Loop;
  
    --处理关联票据,以便回收票据
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).关联序号 = n_关联序号 And Nvl(c_Invoce(I).修改标志, 0) = 0 Then
        If n_关联序号 <> 0 Then
          c_Invoce(I).关联序号 := n_使用id;
        End If;
        c_Invoce(I).修改标志 := 1;
      End If;
    End Loop;
  
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).票号 = c_Invoce_No.发票号 Then
        r_单据号.Extend;
        r_单据号(r_单据号.Count) := c_Invoce(I).No;
        r_单据序号.Extend;
        r_单据序号(r_单据序号.Count) := c_Invoce(I).序号;
        l_关联序号.Extend;
        If Nvl(c_Invoce(I).关联序号, 0) <> 0 Then
          --检查是否存在其他的票据
          n_Find := 0;
          For J In 1 .. c_Invoce.Count Loop
            If c_Invoce(I).关联序号 = c_Invoce(J).关联序号 And c_Invoce(I).票号 <> c_Invoce(J).票号 Then
              n_Find := 1;
              Exit;
            End If;
          End Loop;
        
          If n_Find = 0 Then
            l_关联序号(l_关联序号.Count) := Null;
            c_Invoce(I).关联序号 := 0;
          Else
            l_关联序号(l_关联序号.Count) := c_Invoce(I).关联序号;
          End If;
        Else
          l_关联序号(l_关联序号.Count) := Null;
        End If;
      End If;
    End Loop;
  
    --1.处理门打印内容
    If n_分单据打印 = 1 Then
      --分单据打印,需按单据进行处理
      --票据打印内容
      n_Find := 0;
      v_Temp := '';
      For I In 1 .. r_单据号.Count Loop
        v_Temp := v_Temp || ',' || r_单据号(I);
        If Instr(Nvl(v_打印内容, '-') || ',', ',' || r_单据号(I) || ',') > 0 Then
          --已经找到
          n_Find := 1;
        End If;
      End Loop;
      v_打印内容 := v_打印内容 || Nvl(v_Temp, '+');
    
      If Nvl(n_Find, 0) = 0 Then
        Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
        Forall I In 1 .. r_单据号.Count
          Insert Into 票据打印内容
            (ID, 数据性质, NO, 打印类型)
          Values
            (n_打印id, 1, r_单据号(I), Decode(Nvl(按病人补打票据_In, 0), 1, 1, 0));
        --以便更新门诊费用记录中的实际票号
        v_开始发票号 := c_Invoce_No.发票号;
        Forall I In 1 .. r_单据号.Count
          Update 门诊费用记录 Set 实际票号 = v_开始发票号 Where Mod(记录性质, 10) = 1 And NO = r_单据号(I);
      End If;
    Else
    
      If v_开始发票号 Is Null Then
        --以便更新门诊费用记录中的实际票号
        v_开始发票号 := c_Invoce_No.发票号;
      
        --票据打印内容
        Insert Into 票据打印内容
          (ID, 数据性质, NO, 打印类型)
          Select n_打印id, 1, Column_Value, Decode(Nvl(按病人补打票据_In, 0), 1, 1, 0) From Table(Print_Nos_In);
      
        Update 门诊费用记录
        Set 实际票号 = v_开始发票号
        Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(Print_Nos_In));
      End If;
    End If;
    --2.处理票据打印明细
  
    发票张数_In := 发票张数_In + 1;
    --处理票据使用明细
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
    Values
      (n_使用id, 1, c_Invoce_No.发票号, 1, Decode(操作类型_In, 3, 3, 1), Decode(Nvl(领用id_In, 0), 0, Null, 领用id_In), n_打印id,
       使用人_In, 使用时间_In);
  
    Forall I In 1 .. r_单据号.Count
      Insert Into 票据打印明细
        (使用id, 票种, 是否回收, NO, 票号, 序号, 关联票号序号)
      Values
        (n_使用id, 1, 0, r_单据号(I), c_Invoce_No.发票号, r_单据序号(I), l_关联序号(I));
  
    v_当前发票号 := c_Invoce_No.发票号;
  End Loop;

  If Nvl(领用id_In, 0) <> 0 Then
    Close c_Fact;
    Update 票据领用记录
    Set 使用时间 = 使用时间_In, 当前号码 = v_当前发票号, 剩余数量 = Nvl(剩余数量, 0) - 发票张数_In
    Where ID = 领用id_In
    Returning 剩余数量 Into n_返回数;
    If n_返回数 < 0 Then
      v_Err_Msg := '当前批次的剩余数量不足' || 发票张数_In || '张，无法完成收费票据分配操作。';
      Raise Err_Item;
    End If;
  End If;
  --*****************************************************************************************************************************
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Custom_Invoice_Autoallot;
/

--108753:冉俊明,2017-05-09,修正在根据票据分配规则自动分配票据明细数据时，未按照性能规范加cardinality关键词的问题
Create Or Replace Procedure Zl_Invoice_Autoallot
(
  操作类型_In   Number,
  模拟计算_In   Number,
  票种_In       票据使用明细.票种%Type,
  领用id_In     票据使用明细.领用id%Type,
  病人id_In     门诊费用记录.病人id%Type,
  Nos_In        Varchar2,
  起始发票号_In 门诊费用记录.实际票号%Type,
  使用人_In     票据使用明细.使用人%Type,
  使用时间_In   票据使用明细.使用时间%Type,
  发票号_In     In Out Varchar2,
  发票张数_In   Out Number,
  打印id_In     票据使用明细.打印id%Type := 0
) As
  ---------------------------------------------------------------------------------------------
  --功能：根据票据分配规则,自动分配票据明细数据
  --入参：
  --     操作类型_In :1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
  --     票种_IN     :1-门诊收费;暂无其他类型票据
  --     病人ID_IN   :病人ID,如果Nos和发票号_In为空时,表示针对该病人的所有未打印的票据进行打印
  --     NOs_IN      :单据号,多个用逗号分离,最多;有400张单据,格式为:A00001,A00002.....
  --     启始发票号_IN:重打票据或发出票据的启始票号;
  --     发票号_In   :可以为多个,当操作类型为3-重打票据时,有效
  --     模拟计算_IN :0-不进行模拟计算;1-进行模拟计算,模拟计算时不保存数据
  --     打印ID_In :打印ID_In<>0时，表示根据临时表"临时票据打印内容"所对应的NO来产生打印数据(主要解决按病人补打发票不分结算次数的情况)
  --出参:
  --     发票张数_IN :返回本次收费所需要的发票张数
  ---------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  v_Para       Varchar2(1000);
  v_Temp       Varchar2(32767);
  n_启用模式   Number(3);
  n_分单据打印 Number(3);
  n_执行科室   Number(3);
  n_收据费目   Number(3);
  n_汇总条件   Number(3);
  n_收费细目   Number(3);

  ---------------------------------------------------------
  Type Ty_Rec_Splitno Is Record(
    元素1    票据打印明细.No%Type,
    元素2集  Varchar2(4000),
    元素3集  Varchar2(4000),
    关联序号 Number(18));

  Type Ty_Tb_Splitno Is Table Of Ty_Rec_Splitno;
  c_Split_No   Ty_Tb_Splitno := Ty_Tb_Splitno();
  c_Split_费目 Ty_Tb_Splitno := Ty_Tb_Splitno();

  --------------------------------------------------------
  --定义内部票据处理的数据集
  Type Ty_Rec_Bill Is Record(
    票号     票据打印明细.票号%Type,
    NO       票据打印明细.No%Type,
    序号     票据打印明细.序号%Type,
    关联序号 票据打印明细.关联票号序号%Type,
    修改标志 Number(1));
  Type Ty_Tb_Bill Is Table Of Ty_Rec_Bill;
  c_Invoce Ty_Tb_Bill := Ty_Tb_Bill();
  --------------------------------------------------------
  --按元素1,元素2,元素3,元素4,分别统计各单据的序号
  Type Ty_Rec_No Is Record(
    NO   门诊费用记录.No%Type,
    序号 Varchar2(1000));
  Type Ty_Tb_No Is Table Of Ty_Rec_No;
  c_No Ty_Tb_No := Ty_Tb_No();
  --------------------------------------------------------
  Cursor c_Fact Is
    Select 前缀文本, 剩余数量, 开始号码, 终止号码, 当前号码 From 票据领用记录 Where ID = Nvl(领用id_In, 0);
  r_Factrow c_Fact%RowType;

  --------------------------------------------------------------------------------------------
  --根据相关传入的数据,取对应的数据集

  v_Nos        Varchar2(32767);
  v_发票号     票据打印明细.票号%Type;
  v_开始发票号 票据打印明细.票号%Type;
  v_当前发票号 票据打印明细.票号%Type;
  v_回收票据号 Varchar2(4000);
  n_Find       Number(3);

  n_元素1_Count Number(3);
  n_元素2_Count Number(3);
  n_元素3_Count Number(3);
  n_元素4_Count Number(3);

  v_元素1     门诊费用记录.No%Type;
  n_元素2     门诊费用记录.执行部门id%Type;
  v_元素3     门诊费用记录.收据费目%Type;
  n_元素4     门诊费用记录.收费细目id%Type;
  v_发票信息  Varchar2(4000);
  n_误差项    Number(1);
  n_打印id    票据使用明细.打印id%Type;
  n_使用id    票据使用明细.Id%Type;
  n_返回数    Number(18);
  n_关联序号  Number(18);
  r_单据号    t_Strlist := t_Strlist();
  l_Print_Nos t_Strlist := t_Strlist();
  r_单据序号  t_Strlist := t_Strlist();
  l_使用id    t_Numlist := t_Numlist();
  l_关联序号  t_Numlist := t_Numlist();

  v_打印内容       Varchar2(4000);
  l_元素2          t_Numlist := t_Numlist();
  l_元素3          t_Strlist := t_Strlist();
  v_起始发票号     票据领用记录.开始号码%Type;
  n_按病人补打票据 Number(2);
  n_打印类型       票据打印内容.打印类型%Type;

  -------------------------------------------------------------------------------------------------------------------
  --Invoice_Split_Notgroup:不进行分组汇总或首页汇总时调用此过程
  Procedure Invoice_Split_Notgroup
  (
    Print_Nos        t_Strlist,
    回收发票_In      Varchar2,
    本次打印发票_Out Out Varchar2,
    本次发票张数_Out Out Number
  ) As
    ----------------------------------------------------------------------------
    --入参:
    --   收费收费NOs_IN:本次需要处理的发票所涉及的单据,多个用逗号分离
    --   回收发票_IN-退费时有效,多个用逗号分离，表示本次需要回收的发票号
    --出参:
    -- 本次打印发票_Out-本次需要的发票号,多个用逗号分离
    -- 本次发票张数_Out-本次需要的发票数
    -- 本次退费单据_Out-退费回收所涉及的NO号,多个用逗号分离
  
    n_Count Number(18);
    n_分页  Number(18);
  
    Cursor Cr_Bill Is
      Select NO As 元素1, 执行部门id As 元素2, 收据费目 As 元素3, NO As 元素4, NO As 单据, 序号, 0 As 个数
      From 门诊费用记录
      Where Rownum <= 1;
    c_Bill Cr_Bill%RowType;
    --------------------------------------------------------------------------------------------
    --根据相关传入的数据,取对应的数据集
    Type Ty_费用明细 Is Ref Cursor;
    c_费用明细 Ty_费用明细; --游标变量
  
  Begin
    --按单据分配票据
    If 操作类型_In = 3 Or 操作类型_In = 4 Then
      Open c_费用明细 For
        With c_费用 As
         (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, '-', a.No) As 元素4, a.No As 单据,
                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
          From 门诊费用记录 A,
               (Select /*+cardinality(j,10)*/
                  NO, 序号
                 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                 Where m.票号 = j.Column_Value) B
          Where Mod(记录性质, 10) = 1 And a.No = b.No And Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
          Having Sum(Nvl(a.实收金额, 0)) <> 0)
        Select 元素1, 元素2, 元素3, 元素4, 单据, 序号, Count(*) As 个数
        From c_费用
        Group By 元素1, 元素2, 元素3, 元素4, 单据, 序号
        Order By 元素1, 元素2, 元素3, 元素4, 单据, 序号;
    Else
      Open c_费用明细 For
        With c_费用 As
         (Select /*+cardinality(b,10)*/
           Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
           Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, '-', a.No) As 元素4, a.No As 单据,
           Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
          From 门诊费用记录 A, Table(Print_Nos) B
          Where Mod(记录性质, 10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
          Having Sum(Nvl(a.实收金额, 0)) <> 0)
        Select 元素1, 元素2, 元素3, 元素4, 单据, 序号, Count(*) As 个数
        From c_费用
        Group By 元素1, 元素2, 元素3, 元素4, 单据, 序号
        Order By 元素1, 元素2, 元素3, 元素4, 单据, 序号;
    End If;
  
    v_元素1          := '+';
    n_元素2          := 0;
    v_元素3          := '+';
    n_元素4          := 0;
    n_元素1_Count    := 0;
    n_元素2_Count    := 0;
    n_元素3_Count    := 0;
    n_元素4_Count    := 0;
    本次发票张数_Out := 0;
    If n_汇总条件 <> 0 Then
      n_关联序号 := 1;
    Else
      n_关联序号 := 0;
    End If;
    n_Count := 0;
    c_No.Delete;
    Loop
      Fetch c_费用明细
        Into c_Bill;
      Exit When c_费用明细%NotFound;
      n_Count := 1;
    
      n_分页 := 0;
      If (v_元素1 <> c_Bill.元素1) Or (n_元素2 <> c_Bill.元素2 And n_元素2_Count >= n_执行科室 And n_执行科室 <> 0) Or
         (v_元素3 <> c_Bill.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Or (n_元素4_Count >= n_收费细目 And n_收费细目 <> 0) Then
      
        If (v_元素1 <> '+' Or n_元素2 <> 0 Or v_元素3 <> '+' Or n_元素4 <> 0) Then
          n_分页 := 1;
        End If;
        n_元素2_Count := 0;
        n_元素3_Count := 0;
        n_元素4_Count := 0;
        n_元素1_Count := 0;
        v_元素1       := '+';
        n_元素2       := 0;
        v_元素3       := '+';
      End If;
    
      If n_分页 = 1 Then
        --分页:计算发票号及相关的
        For I In 1 .. c_No.Count Loop
          c_Invoce.Extend;
          c_Invoce(c_Invoce.Count).票号 := v_发票号;
          c_Invoce(c_Invoce.Count).No := c_No(I).No;
          c_Invoce(c_Invoce.Count).序号 := Case
                                           When Instr(c_No(I).序号, ',') > 0 Then
                                            Substr(c_No(I).序号, 2)
                                           Else
                                            c_No(I).序号
                                         End;
          c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
        End Loop;
      
        本次发票张数_Out := 本次发票张数_Out + 1;
        本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
        v_发票号         := Zl_Incstr(v_发票号);
        c_No.Delete;
      End If;
      If (v_元素1 <> c_Bill.元素1) Then
        n_元素1_Count := n_元素1_Count + 1;
        v_元素1       := c_Bill.元素1;
      End If;
      If (n_元素2 <> c_Bill.元素2) Then
        n_元素2_Count := n_元素2_Count + 1;
        n_元素2       := c_Bill.元素2;
      End If;
      If (v_元素3 <> c_Bill.元素3) Then
        n_元素3_Count := n_元素3_Count + 1;
        v_元素3       := c_Bill.元素3;
      End If;
      If n_收费细目 <> 0 Then
        n_元素4_Count := n_元素4_Count + 1;
      End If;
    
      -------------------------------------------
      --分配单据号及序号
      n_Find := 0;
      For J In 1 .. c_No.Count Loop
        If c_No(J).No = c_Bill.单据 Then
          --单据号相同,将序号合并
          c_No(J).序号 := c_No(J).序号 || ',' || c_Bill.序号;
          n_Find := 1;
          Exit;
        End If;
      End Loop;
      If n_Find = 0 Then
        c_No.Extend;
        c_No(c_No.Count).No := c_Bill.单据;
        c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_Bill.序号;
      End If;
    End Loop;
  
    --是否有发票数据
    If n_Count >= 1 Then
      --最后一个发票分配
      本次发票张数_Out := 本次发票张数_Out + 1;
      本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
    Else
      本次发票张数_Out := 0;
      本次打印发票_Out := '';
    End If;
    If c_No.Count <> 0 Then
      For I In 1 .. c_No.Count Loop
        c_Invoce.Extend;
        c_Invoce(c_Invoce.Count).票号 := v_发票号;
        c_Invoce(c_Invoce.Count).No := c_No(I).No;
        If Instr(c_No(I).序号, ',') > 0 Then
          c_No(I).序号 := Substr(c_No(I).序号, 2);
        End If;
        c_Invoce(c_Invoce.Count).序号 := c_No(I).序号;
        c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
      End Loop;
      c_No.Delete;
    End If;
    If Instr(Nvl(本次打印发票_Out, '-'), ',') > 0 Then
      本次打印发票_Out := Substr(本次打印发票_Out, 2);
    End If;
  End Invoice_Split_Notgroup;
  --结束:不进行分组汇总或首页汇总时调用此过程
  -------------------------------------------------------------------------------------------------------------------
  --按组汇总
  Procedure Invoice_Split_Group
  (
    Print_Nos        t_Strlist,
    回收发票_In      Varchar2,
    本次打印发票_Out Out Varchar2,
    本次发票张数_Out Out Number
  ) As
  Begin
    v_元素1          := '+';
    n_元素2          := 0;
    v_元素3          := '+';
    n_元素4          := 0;
    n_元素1_Count    := 0;
    n_元素2_Count    := 0;
    n_元素3_Count    := 0;
    n_元素4_Count    := 0;
    本次发票张数_Out := 0;
  
    c_No.Delete;
    l_元素2.Delete;
  
    --按单据分配票据
    If 操作类型_In = 3 Or 操作类型_In = 4 Then
      --******************************************************************************************************************************
      --退费和重打按发票号处理(开始)
      --4.收据费目+收费细目
      If n_分单据打印 = 0 And n_执行科室 = 0 And n_收据费目 <> 0 And n_收费细目 <> 0 Then
        v_元素3 := '+';
        c_Split_费目.Delete;
        For c_分页 In (With c_费用 As
                        (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                               Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                               Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A,
                             (Select /*+cardinality(j,10)*/
                                NO, 序号
                               From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                               Where m.票号 = j.Column_Value) B
                        Where Mod(记录性质, 10) = 1 And a.No = b.No And
                              Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                              Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select a.元素3, Count(*) As 个数 From c_费用 A Group By 元素3 Order By 元素3) Loop
          If (v_元素3 <> c_分页.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Then
            If v_元素3 <> '+' Then
              c_Split_费目.Extend;
              For J In 1 .. l_元素3.Count Loop
                --单据号相同,将序号合并
                c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
              End Loop;
              v_元素3       := '+';
              n_元素3_Count := 0;
              l_元素3.Delete;
            End If;
          End If;
          If (v_元素3 <> c_分页.元素3) Then
            n_元素3_Count := n_元素3_Count + 1;
            v_元素3       := c_分页.元素3;
            l_元素3.Extend;
            l_元素3(l_元素3.Count) := v_元素3;
          End If;
        End Loop;
        If l_元素3.Count <> 0 Then
          c_Split_费目.Extend;
          For J In 1 .. l_元素3.Count Loop
            --单据号相同,将序号合并
            c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
          End Loop;
        End If;
        n_关联序号 := 0;
        For I In 1 .. c_Split_费目.Count Loop
          c_No.Delete;
          n_关联序号    := n_关联序号 + 1;
          n_元素4_Count := 0;
          For c_分页 In (With c_费用 As
                          (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                          From 门诊费用记录 A,
                               (Select /*+cardinality(j,10)*/
                                  NO, 序号
                                 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                                 Where m.票号 = j.Column_Value) B
                          Where Mod(记录性质, 10) = 1 And a.No = b.No And
                                Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                          Having Sum(Nvl(a.实收金额, 0)) <> 0)
                         Select m.元素1, 元素2, 元素3, m.元素4, m.单据, m.序号, Count(*) As 个数
                         From c_费用 M
                         Where Instr(',' || c_Split_费目(I).元素3集 || ',', ',' || m.元素3 || ',') > 0
                         Group By m.元素1, 元素2, m.元素4, 元素3, m.单据, m.序号
                         Order By m.元素1, 元素2, m.元素4, 元素3, m.单据, m.序号) Loop
            If n_元素4_Count >= n_收费细目 And n_收费细目 <> 0 Then
              --分页:计算发票号及相关的
              For J In 1 .. c_No.Count Loop
                c_Invoce.Extend;
                c_Invoce(c_Invoce.Count).票号 := v_发票号;
                c_Invoce(c_Invoce.Count).No := c_No(J).No;
                c_Invoce(c_Invoce.Count).序号 := Case
                                                 When Instr(c_No(J).序号, ',') > 0 Then
                                                  Substr(c_No(J).序号, 2)
                                                 Else
                                                  c_No(J).序号
                                               End;
                c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
              End Loop;
              本次发票张数_Out := 本次发票张数_Out + 1;
              本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
              v_发票号         := Zl_Incstr(v_发票号);
              c_No.Delete;
              n_元素4_Count := 0;
              --分页
            End If;
            n_元素4_Count := n_元素4_Count + 1;
            -------------------------------------------
            --分配单据号及序号
            n_Find := 0;
            For J In 1 .. c_No.Count Loop
              If c_No(J).No = c_分页.单据 Then
                --单据号相同,将序号合并
                c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
                n_Find := 1;
                Exit;
              End If;
            End Loop;
            If n_Find = 0 Then
              c_No.Extend;
              c_No(c_No.Count).No := c_分页.单据;
              c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
            End If;
          End Loop;
          If c_No.Count <> 0 Then
            --分页:计算发票号及相关的
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).票号 := v_发票号;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).序号 := Case
                                               When Instr(c_No(J).序号, ',') > 0 Then
                                                Substr(c_No(J).序号, 2)
                                               Else
                                                c_No(J).序号
                                             End;
              c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
            End Loop;
            本次发票张数_Out := 本次发票张数_Out + 1;
            本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
            v_发票号         := Zl_Incstr(v_发票号);
            c_No.Delete;
          End If;
        End Loop;
      End If;
    
      If (n_分单据打印 = 1 Or n_执行科室 > 0) And (n_收据费目 <> 0 Or n_收费细目 <> 0) Then
        n_元素2_Count := 0;
        v_元素1       := '+';
        n_元素2       := 0;
        c_Split_No.Delete;
        For c_分页 In (With c_费用 As
                        (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                               Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                               Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A,
                             (Select /*+cardinality(j,10)*/
                                NO, 序号
                               From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                               Where m.票号 = j.Column_Value) B
                        Where Mod(记录性质, 10) = 1 And a.No = b.No And
                              Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                             
                              Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select a.元素1, a.元素2, b.编码, Count(*) As 个数
                       From c_费用 A, 部门表 B
                       Where a.元素2 = b.Id(+)
                       Group By 元素1, b.编码, 元素2
                       Order By 元素1, b.编码, 元素2) Loop
          If (v_元素1 <> c_分页.元素1) Or (n_元素2 <> c_分页.元素2 And n_元素2_Count >= n_执行科室 And n_执行科室 <> 0) Then
            c_Split_No.Extend;
            n_元素2_Count := 0;
            v_元素1       := '+';
            n_元素2       := 0;
          End If;
          If (v_元素1 <> c_分页.元素1) Then
            v_元素1 := c_分页.元素1;
            c_Split_No(c_Split_No.Count).元素1 := v_元素1;
          End If;
          If (n_元素2 <> c_分页.元素2) Then
            n_元素2_Count := n_元素2_Count + 1;
            n_元素2 := c_分页.元素2;
            c_Split_No(c_Split_No.Count).元素2集 := c_Split_No(c_Split_No.Count).元素2集 || ',' || n_元素2;
          End If;
        End Loop;
      End If;
    
      --6.(no Or 执行科室)+收费细目
      If (n_分单据打印 = 0 Or n_执行科室 > 0) And n_收据费目 = 0 And n_收费细目 <> 0 Then
      
        For I In 1 .. c_Split_No.Count Loop
          v_元素3 := '+';
          --只有首页汇总的,才有关联序号
          n_关联序号    := n_关联序号 + 1;
          n_元素4_Count := 0;
          For c_分页 In (With c_费用 As
                          (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                          From 门诊费用记录 A,
                               (Select /*+cardinality(j,10)*/
                                  NO, 序号
                                 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                                 Where m.票号 = j.Column_Value) B
                          Where Mod(记录性质, 10) = 1 And a.No = b.No And
                                Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                               
                                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                          Having Sum(Nvl(a.实收金额, 0)) <> 0)
                         Select 元素1, 元素2, 元素3, a.元素4, a.单据, a.序号, Count(*) As 个数
                         From c_费用 A
                         Where a.元素1 = c_Split_No(I).元素1 And
                               Instr(',' || c_Split_No(I).元素2集 || ',', ',' || a.元素2 || ',') > 0
                         Group By 元素1, 元素2, 元素4, 元素3, 单据, a.序号
                         Order By 元素1, 元素2, 元素4, 单据, 序号) Loop
            If n_元素4_Count >= n_收费细目 And n_收费细目 <> 0 Then
              --分配单据
              If c_No.Count <> 0 Then
                --分页:计算发票号及相关的
                For J In 1 .. c_No.Count Loop
                  c_Invoce.Extend;
                  c_Invoce(c_Invoce.Count).票号 := v_发票号;
                  c_Invoce(c_Invoce.Count).No := c_No(J).No;
                  c_Invoce(c_Invoce.Count).序号 := Case
                                                   When Instr(c_No(J).序号, ',') > 0 Then
                                                    Substr(c_No(J).序号, 2)
                                                   Else
                                                    c_No(J).序号
                                                 End;
                  c_Invoce(c_Invoce.Count).关联序号 := n_元素4_Count;
                End Loop;
                本次发票张数_Out := 本次发票张数_Out + 1;
                本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
                v_发票号         := Zl_Incstr(v_发票号);
                c_No.Delete;
              End If;
              n_元素4_Count := 0;
            End If;
            n_元素4_Count := n_元素4_Count + 1;
          
            -------------------------------------------
            --分配单据号及序号
            n_Find := 0;
            For J In 1 .. c_No.Count Loop
              If c_No(J).No = c_分页.单据 Then
                --单据号相同,将序号合并
                c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
                n_Find := 1;
                Exit;
              End If;
            End Loop;
            If n_Find = 0 Then
              c_No.Extend;
              c_No(c_No.Count).No := c_分页.单据;
              c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
            End If;
          End Loop;
          --分配单据
          If c_No.Count <> 0 Then
            --分页:计算发票号及相关的
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).票号 := v_发票号;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).序号 := Case
                                               When Instr(c_No(J).序号, ',') > 0 Then
                                                Substr(c_No(J).序号, 2)
                                               Else
                                                c_No(J).序号
                                             End;
              c_Invoce(c_Invoce.Count).关联序号 := n_元素4_Count;
            End Loop;
            本次发票张数_Out := 本次发票张数_Out + 1;
            本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
            v_发票号         := Zl_Incstr(v_发票号);
            c_No.Delete;
          End If;
        End Loop;
      End If;
    
      --7.(no Or 执行科室)+收据费目+收费细目
      n_关联序号 := 0;
      If (n_分单据打印 = 0 Or n_执行科室 > 0) And n_收据费目 <> 0 And n_收费细目 <> 0 Then
        c_Split_费目.Delete;
        For I In 1 .. c_Split_No.Count Loop
        
          n_关联序号    := n_关联序号 + 1;
          v_元素3       := '+';
          n_元素3_Count := 0;
          l_元素3.Delete;
          For c_分页 In (With c_费用 As
                          (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                          From 门诊费用记录 A,
                               (Select /*+cardinality(j,10)*/
                                  NO, 序号
                                 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                                 Where m.票号 = j.Column_Value) B
                          Where Mod(记录性质, 10) = 1 And a.No = b.No And
                                Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                               
                                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                          Having Sum(Nvl(a.实收金额, 0)) <> 0)
                         Select a.元素3, Count(*) As 个数
                         From c_费用 A
                         Where a.元素1 = c_Split_No(I).元素1 And
                               Instr(',' || c_Split_No(I).元素2集 || ',', ',' || a.元素2 || ',') > 0
                         Group By 元素3
                         Order By 元素3) Loop
          
            If (v_元素3 <> c_分页.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Then
              If v_元素3 <> '+' Then
                c_Split_费目.Extend;
                c_Split_费目(c_Split_费目.Count).元素1 := c_Split_No(I).元素1;
                c_Split_费目(c_Split_费目.Count).元素2集 := c_Split_No(I).元素2集;
                c_Split_费目(c_Split_费目.Count).关联序号 := n_关联序号;
                For J In 1 .. l_元素3.Count Loop
                  --单据号相同,将序号合并
                  c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
                End Loop;
              End If;
              v_元素3       := '+';
              n_元素3_Count := 0;
              l_元素3.Delete;
            End If;
            If (v_元素3 <> c_分页.元素3) Then
              n_元素3_Count := n_元素3_Count + 1;
              v_元素3       := c_分页.元素3;
              l_元素3.Extend;
              l_元素3(l_元素3.Count) := v_元素3;
            End If;
          End Loop;
        
          If l_元素3.Count <> 0 Then
            c_Split_费目.Extend;
            c_Split_费目(c_Split_费目.Count).元素1 := c_Split_No(I).元素1;
            c_Split_费目(c_Split_费目.Count).元素2集 := c_Split_No(I).元素2集;
            c_Split_费目(c_Split_费目.Count).关联序号 := n_关联序号;
            For J In 1 .. l_元素3.Count Loop
              --单据号相同,将序号合并
              c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
            End Loop;
          End If;
        End Loop;
      
        For I In 1 .. c_Split_费目.Count Loop
          c_No.Delete;
          n_元素4_Count := 0;
          For c_分页 In (With c_费用 As
                          (Select Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                                 Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                                 Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                          From 门诊费用记录 A,
                               (Select /*+cardinality(j,10)*/
                                  NO, 序号
                                 From 票据打印明细 M, Table(f_Str2list(回收发票_In)) J
                                 Where m.票号 = j.Column_Value) B
                          Where Mod(记录性质, 10) = 1 And a.No = b.No And
                                Instr(',' || b.序号 || ',', ',' || Nvl(a.价格父号, a.序号) || ',') > 0 And
                               
                                Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                          Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                          Having Sum(Nvl(a.实收金额, 0)) <> 0)
                         Select 元素1, 元素2, 元素3, a.元素4, a.单据, a.序号, Count(*) As 个数
                         From c_费用 A
                         Where a.元素1 = c_Split_费目(I).元素1 And
                               Instr(',' || c_Split_费目(I).元素2集 || ',', ',' || a.元素2 || ',') > 0 And
                               Instr(',' || c_Split_费目(I).元素3集 || ',', ',' || a.元素3 || ',') > 0
                         Group By 元素1, 元素2, 元素4, 元素3, a.单据, a.序号
                         Order By 元素1, 元素2, 元素4, 元素3, 单据, 序号) Loop
            If (n_元素4_Count >= n_收费细目 And n_收费细目 <> 0) Then
              --分配单据
              If c_No.Count <> 0 Then
                --分页:计算发票号及相关的
                For J In 1 .. c_No.Count Loop
                  c_Invoce.Extend;
                  c_Invoce(c_Invoce.Count).票号 := v_发票号;
                  c_Invoce(c_Invoce.Count).No := c_No(J).No;
                  c_Invoce(c_Invoce.Count).序号 := Case
                                                   When Instr(c_No(J).序号, ',') > 0 Then
                                                    Substr(c_No(J).序号, 2)
                                                   Else
                                                    c_No(J).序号
                                                 End;
                  c_Invoce(c_Invoce.Count).关联序号 := c_Split_费目(I).关联序号;
                End Loop;
                本次发票张数_Out := 本次发票张数_Out + 1;
                本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
                v_发票号         := Zl_Incstr(v_发票号);
                c_No.Delete;
              End If;
              n_元素4_Count := 0;
            End If;
            n_元素4_Count := n_元素4_Count + 1;
            -------------------------------------------
            --分配单据号及序号
            n_Find := 0;
            For J In 1 .. c_No.Count Loop
              If c_No(J).No = c_分页.单据 Then
                --单据号相同,将序号合并
                c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
                n_Find := 1;
                Exit;
              End If;
            End Loop;
            If n_Find = 0 Then
              c_No.Extend;
              c_No(c_No.Count).No := c_分页.单据;
              c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
            End If;
          End Loop;
        
          --分配单据
          If c_No.Count <> 0 Then
            --分页:计算发票号及相关的
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).票号 := v_发票号;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).序号 := Case
                                               When Instr(c_No(J).序号, ',') > 0 Then
                                                Substr(c_No(J).序号, 2)
                                               Else
                                                c_No(J).序号
                                             End;
              c_Invoce(c_Invoce.Count).关联序号 := c_Split_费目(I).关联序号;
            End Loop;
            本次发票张数_Out := 本次发票张数_Out + 1;
            本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
            v_发票号         := Zl_Incstr(v_发票号);
            c_No.Delete;
          End If;
        End Loop;
      End If;
    
      --退费和重打按发票号处理(结束)
      --******************************************************************************************************************************
      If Instr(Nvl(本次打印发票_Out, '-'), ',') > 0 Then
        本次打印发票_Out := Substr(本次打印发票_Out, 2);
      End If;
      Return;
    
    End If;
  
    --******************************************************************************************************************************
    --以下是按正常分配单据(开始)
    --4.收据费目+收费细目
    If n_分单据打印 = 0 And n_执行科室 = 0 And n_收据费目 <> 0 And n_收费细目 <> 0 Then
      v_元素3 := '+';
      c_Split_费目.Delete;
    
      For c_分页 In (With c_费用 As
                      (Select /*+cardinality(b,10)*/
                       Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                       Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                       Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                      From 门诊费用记录 A, Table(Print_Nos) B
                      Where Mod(记录性质, 10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                      Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                      Having Sum(Nvl(a.实收金额, 0)) <> 0)
                     Select a.元素3, Count(*) As 个数 From c_费用 A Group By 元素3 Order By 元素3) Loop
        If (v_元素3 <> c_分页.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Then
          If v_元素3 <> '+' Then
            c_Split_费目.Extend;
            For J In 1 .. l_元素3.Count Loop
              --单据号相同,将序号合并
              c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
            End Loop;
            v_元素3       := '+';
            n_元素3_Count := 0;
            l_元素3.Delete;
          End If;
        End If;
        If (v_元素3 <> c_分页.元素3) Then
          n_元素3_Count := n_元素3_Count + 1;
          v_元素3       := c_分页.元素3;
          l_元素3.Extend;
          l_元素3(l_元素3.Count) := v_元素3;
        End If;
      End Loop;
      If l_元素3.Count <> 0 Then
        c_Split_费目.Extend;
        For J In 1 .. l_元素3.Count Loop
          --单据号相同,将序号合并
          c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
        End Loop;
      End If;
      n_关联序号 := 0;
      For I In 1 .. c_Split_费目.Count Loop
        c_No.Delete;
        n_关联序号    := n_关联序号 + 1;
        n_元素4_Count := 0;
        For c_分页 In (With c_费用 As
                        (Select /*+cardinality(b,10)*/
                         Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                         Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                         Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A, Table(Print_Nos) B
                        Where Mod(记录性质, 10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select m.元素1, 元素2, 元素3, m.元素4, m.单据, m.序号, Count(*) As 个数
                       From c_费用 M
                       Where Instr(',' || c_Split_费目(I).元素3集 || ',', ',' || m.元素3 || ',') > 0
                       Group By m.元素1, 元素2, m.元素4, 元素3, m.单据, m.序号
                       Order By m.元素1, 元素2, m.元素4, 元素3, m.单据, m.序号) Loop
          If n_元素4_Count >= n_收费细目 And n_收费细目 <> 0 Then
            --分页:计算发票号及相关的
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).票号 := v_发票号;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).序号 := Case
                                               When Instr(c_No(J).序号, ',') > 0 Then
                                                Substr(c_No(J).序号, 2)
                                               Else
                                                c_No(J).序号
                                             End;
              c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
            End Loop;
            本次发票张数_Out := 本次发票张数_Out + 1;
            本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
            v_发票号         := Zl_Incstr(v_发票号);
            c_No.Delete;
            n_元素4_Count := 0;
            --分页
          End If;
          n_元素4_Count := n_元素4_Count + 1;
          -------------------------------------------
          --分配单据号及序号
          n_Find := 0;
          For J In 1 .. c_No.Count Loop
            If c_No(J).No = c_分页.单据 Then
              --单据号相同,将序号合并
              c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
              n_Find := 1;
              Exit;
            End If;
          End Loop;
          If n_Find = 0 Then
            c_No.Extend;
            c_No(c_No.Count).No := c_分页.单据;
            c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
          End If;
        End Loop;
        If c_No.Count <> 0 Then
          --分页:计算发票号及相关的
          For J In 1 .. c_No.Count Loop
            c_Invoce.Extend;
            c_Invoce(c_Invoce.Count).票号 := v_发票号;
            c_Invoce(c_Invoce.Count).No := c_No(J).No;
            c_Invoce(c_Invoce.Count).序号 := Case
                                             When Instr(c_No(J).序号, ',') > 0 Then
                                              Substr(c_No(J).序号, 2)
                                             Else
                                              c_No(J).序号
                                           End;
            c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
          End Loop;
          本次发票张数_Out := 本次发票张数_Out + 1;
          本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
          v_发票号         := Zl_Incstr(v_发票号);
          c_No.Delete;
        End If;
      End Loop;
    End If;
  
    If (n_分单据打印 = 1 Or n_执行科室 > 0) And (n_收据费目 <> 0 Or n_收费细目 <> 0) Then
      n_元素2_Count := 0;
      v_元素1       := '+';
      n_元素2       := 0;
      c_Split_No.Delete;
      For c_分页 In (With c_费用 As
                      (Select /*+cardinality(b,10)*/
                       Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                       Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                       Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                      From 门诊费用记录 A, Table(Print_Nos) B
                      Where Mod(记录性质, 10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                      Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                      Having Sum(Nvl(a.实收金额, 0)) <> 0)
                     Select a.元素1, a.元素2, b.编码, Count(*) As 个数
                     From c_费用 A, 部门表 B
                     Where a.元素2 = b.Id(+)
                     Group By 元素1, b.编码, 元素2
                     Order By 元素1, b.编码, 元素2) Loop
        If (v_元素1 <> c_分页.元素1) Or (n_元素2 <> c_分页.元素2 And n_元素2_Count >= n_执行科室 And n_执行科室 <> 0) Then
          c_Split_No.Extend;
          n_元素2_Count := 0;
          v_元素1       := '+';
          n_元素2       := 0;
        End If;
        If (v_元素1 <> c_分页.元素1) Then
          v_元素1 := c_分页.元素1;
          c_Split_No(c_Split_No.Count).元素1 := v_元素1;
        End If;
        If (n_元素2 <> c_分页.元素2) Then
          n_元素2_Count := n_元素2_Count + 1;
          n_元素2 := c_分页.元素2;
          c_Split_No(c_Split_No.Count).元素2集 := c_Split_No(c_Split_No.Count).元素2集 || ',' || n_元素2;
        End If;
      End Loop;
    End If;
  
    --3.(no Or 执行科室)+收费细目
    If (n_分单据打印 = 0 Or n_执行科室 > 0) And n_收据费目 = 0 And n_收费细目 <> 0 Then
    
      For I In 1 .. c_Split_No.Count Loop
        v_元素3 := '+';
        --只有首页汇总的,才有关联序号
        n_关联序号    := Nvl(n_关联序号, 0) + 1;
        n_元素4_Count := 0;
        For c_分页 In (With c_费用 As
                        (Select /*+cardinality(b,10)*/
                         Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                         Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                         Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A, Table(Print_Nos) B
                        Where Mod(记录性质, 10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select 元素1, 元素2, 元素3, a.元素4, a.单据, a.序号, Count(*) As 个数
                       From c_费用 A
                       Where a.元素1 = c_Split_No(I).元素1 And
                             Instr(',' || c_Split_No(I).元素2集 || ',', ',' || a.元素2 || ',') > 0
                       Group By 元素1, 元素2, 元素4, 元素3, 单据, a.序号
                       Order By 元素1, 元素2, 元素4, 单据, 序号) Loop
          If n_元素4_Count >= n_收费细目 And n_收费细目 <> 0 Then
            --分配单据
            If c_No.Count <> 0 Then
              --分页:计算发票号及相关的
              For J In 1 .. c_No.Count Loop
                c_Invoce.Extend;
                c_Invoce(c_Invoce.Count).票号 := v_发票号;
                c_Invoce(c_Invoce.Count).No := c_No(J).No;
                c_Invoce(c_Invoce.Count).序号 := Case
                                                 When Instr(c_No(J).序号, ',') > 0 Then
                                                  Substr(c_No(J).序号, 2)
                                                 Else
                                                  c_No(J).序号
                                               End;
                c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
              End Loop;
              本次发票张数_Out := 本次发票张数_Out + 1;
              本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
              v_发票号         := Zl_Incstr(v_发票号);
              c_No.Delete;
            End If;
            n_元素4_Count := 0;
          End If;
          n_元素4_Count := n_元素4_Count + 1;
        
          -------------------------------------------
          --分配单据号及序号
          n_Find := 0;
          For J In 1 .. c_No.Count Loop
            If c_No(J).No = c_分页.单据 Then
              --单据号相同,将序号合并
              c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
              n_Find := 1;
              Exit;
            End If;
          End Loop;
          If n_Find = 0 Then
            c_No.Extend;
            c_No(c_No.Count).No := c_分页.单据;
            c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
          End If;
        End Loop;
        --分配单据
        If c_No.Count <> 0 Then
          --分页:计算发票号及相关的
          For J In 1 .. c_No.Count Loop
            c_Invoce.Extend;
            c_Invoce(c_Invoce.Count).票号 := v_发票号;
            c_Invoce(c_Invoce.Count).No := c_No(J).No;
            c_Invoce(c_Invoce.Count).序号 := Case
                                             When Instr(c_No(J).序号, ',') > 0 Then
                                              Substr(c_No(J).序号, 2)
                                             Else
                                              c_No(J).序号
                                           End;
            c_Invoce(c_Invoce.Count).关联序号 := n_关联序号;
          End Loop;
          本次发票张数_Out := 本次发票张数_Out + 1;
          本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
          v_发票号         := Zl_Incstr(v_发票号);
          c_No.Delete;
        End If;
      End Loop;
    End If;
  
    --7.(no Or 执行科室)+收据费目+收费细目
    n_关联序号 := 0;
    If (n_分单据打印 = 0 Or n_执行科室 > 0) And n_收据费目 <> 0 And n_收费细目 <> 0 Then
      c_Split_费目.Delete;
    
      For I In 1 .. c_Split_No.Count Loop
      
        n_关联序号    := n_关联序号 + 1;
        v_元素3       := '+';
        n_元素3_Count := 0;
        l_元素3.Delete;
        For c_分页 In (With c_费用 As
                        (Select /*+cardinality(b,10)*/
                         Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                         Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                         Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A, Table(Print_Nos) B
                        Where Mod(记录性质, 10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select a.元素3, Count(*) As 个数
                       From c_费用 A
                       Where a.元素1 = c_Split_No(I).元素1 And
                             Instr(',' || c_Split_No(I).元素2集 || ',', ',' || a.元素2 || ',') > 0
                       Group By 元素3
                       Order By 元素3) Loop
          If (v_元素3 <> c_分页.元素3 And n_元素3_Count >= n_收据费目 And n_收据费目 <> 0) Then
            If v_元素3 <> '+' Then
              c_Split_费目.Extend;
              c_Split_费目(c_Split_费目.Count).元素1 := c_Split_No(I).元素1;
              c_Split_费目(c_Split_费目.Count).元素2集 := c_Split_No(I).元素2集;
              c_Split_费目(c_Split_费目.Count).关联序号 := n_关联序号;
              For J In 1 .. l_元素3.Count Loop
                --单据号相同,将序号合并
                c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
              End Loop;
            End If;
            v_元素3       := '+';
            n_元素3_Count := 0;
            l_元素3.Delete;
          End If;
          If (v_元素3 <> c_分页.元素3) Then
            n_元素3_Count := n_元素3_Count + 1;
            v_元素3       := c_分页.元素3;
            l_元素3.Extend;
            l_元素3(l_元素3.Count) := v_元素3;
          End If;
        End Loop;
      
        If l_元素3.Count <> 0 Then
          c_Split_费目.Extend;
          c_Split_费目(c_Split_费目.Count).元素1 := c_Split_No(I).元素1;
          c_Split_费目(c_Split_费目.Count).元素2集 := c_Split_No(I).元素2集;
          c_Split_费目(c_Split_费目.Count).关联序号 := n_关联序号;
          For J In 1 .. l_元素3.Count Loop
            --单据号相同,将序号合并
            c_Split_费目(c_Split_费目.Count).元素3集 := c_Split_费目(c_Split_费目.Count).元素3集 || ',' || l_元素3(J);
          End Loop;
        End If;
      End Loop;
    
      For I In 1 .. c_Split_费目.Count Loop
        c_No.Delete;
        n_元素4_Count := 0;
        --收费细目,按条数计数,还是要按执行科室+收据费目
        For c_分页 In (With c_费用 As
                        (Select /*+cardinality(b,10)*/
                         Decode(n_分单据打印, 0, '-', a.No) As 元素1, Decode(n_执行科室, 0, 0, a.执行部门id) As 元素2,
                         Decode(n_收据费目, 0, '-', a.收据费目) As 元素3, Decode(n_收费细目, 0, 0, a.收费细目id) As 元素4, a.No As 单据,
                         Nvl(a.价格父号, a.序号) As 序号, Sum(Nvl(a.实收金额, 0)) As 实收金额
                        From 门诊费用记录 A, Table(Print_Nos) B
                        Where Mod(记录性质, 10) = 1 And a.No = b.Column_Value And Decode(n_误差项, 1, Nvl(a.附加标志, 0), 0) <> 9
                        Group By a.No, a.执行部门id, a.收据费目, a.收费细目id, Nvl(a.价格父号, a.序号)
                        Having Sum(Nvl(a.实收金额, 0)) <> 0)
                       Select 元素1, 元素2, 元素3, a.元素4, a.单据, a.序号, Count(*) As 个数
                       From c_费用 A
                       Where a.元素1 = c_Split_费目(I).元素1 And
                             Instr(',' || c_Split_费目(I).元素2集 || ',', ',' || a.元素2 || ',') > 0 And
                             Instr(',' || c_Split_费目(I).元素3集 || ',', ',' || a.元素3 || ',') > 0
                       Group By 元素1, 元素2, 元素4, 元素3, a.单据, a.序号
                       Order By 元素1, 元素2, 元素4, 元素3, 单据, 序号) Loop
          If (n_元素4_Count >= n_收费细目 And n_收费细目 <> 0) Then
            --分配单据
            If c_No.Count <> 0 Then
              --分页:计算发票号及相关的
              For J In 1 .. c_No.Count Loop
                c_Invoce.Extend;
                c_Invoce(c_Invoce.Count).票号 := v_发票号;
                c_Invoce(c_Invoce.Count).No := c_No(J).No;
                c_Invoce(c_Invoce.Count).序号 := Case
                                                 When Instr(c_No(J).序号, ',') > 0 Then
                                                  Substr(c_No(J).序号, 2)
                                                 Else
                                                  c_No(J).序号
                                               End;
                c_Invoce(c_Invoce.Count).关联序号 := c_Split_费目(I).关联序号;
              End Loop;
              本次发票张数_Out := 本次发票张数_Out + 1;
              本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
              v_发票号         := Zl_Incstr(v_发票号);
              c_No.Delete;
            End If;
            n_元素4_Count := 0;
          End If;
          n_元素4_Count := n_元素4_Count + 1;
          -------------------------------------------
          --分配单据号及序号
          n_Find := 0;
          For J In 1 .. c_No.Count Loop
            If c_No(J).No = c_分页.单据 Then
              --单据号相同,将序号合并
              c_No(J).序号 := c_No(J).序号 || ',' || c_分页.序号;
              n_Find := 1;
              Exit;
            End If;
          End Loop;
          If n_Find = 0 Then
            c_No.Extend;
            c_No(c_No.Count).No := c_分页.单据;
            c_No(c_No.Count).序号 := c_No(c_No.Count).序号 || ',' || c_分页.序号;
          End If;
        End Loop;
        --分配单据
        If c_No.Count <> 0 Then
          --分页:计算发票号及相关的
          For J In 1 .. c_No.Count Loop
            c_Invoce.Extend;
            c_Invoce(c_Invoce.Count).票号 := v_发票号;
            c_Invoce(c_Invoce.Count).No := c_No(J).No;
            c_Invoce(c_Invoce.Count).序号 := Case
                                             When Instr(c_No(J).序号, ',') > 0 Then
                                              Substr(c_No(J).序号, 2)
                                             Else
                                              c_No(J).序号
                                           End;
            c_Invoce(c_Invoce.Count).关联序号 := c_Split_费目(I).关联序号;
          End Loop;
          本次发票张数_Out := 本次发票张数_Out + 1;
          本次打印发票_Out := Nvl(本次打印发票_Out, '') || ',' || v_发票号;
          v_发票号         := Zl_Incstr(v_发票号);
          c_No.Delete;
        End If;
      End Loop;
    End If;
    --正常分配单据结束
    --******************************************************************************************************************************
    If Instr(Nvl(本次打印发票_Out, '-'), ',') > 0 Then
      本次打印发票_Out := Substr(本次打印发票_Out, 2);
    End If;
  End Invoice_Split_Group;
  -------------------------------------------------------------------------------------------------------------------
Begin

  --启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
  v_Para := Nvl(zl_GetSysParameter('票据分配规则', 1121), '0||0;0;0,0;0');
  If Instr(v_Para, '||') = 0 Then
    v_Para := v_Para || '||';
  End If;
  v_Temp := Substr(v_Para, 1, Instr(v_Para, '||') - 1);
  If v_Temp Is Null Then
    --无设置值,代表无启用,直接返回
    Return;
  End If;

  --0-根据实际打印分配票号;1-根据预定规则分配票号;2-.根据自定义规则分配票号
  n_启用模式 := Zl_To_Number(v_Temp);
  If Nvl(n_启用模式, 0) = 0 Then
    --0-根据实际打印分配票号:按原来的处理方式分配票据,直接退出
    Return;
  End If;
  v_Temp       := Nvl(zl_GetSysParameter('误差项不使用票据', 1121), '0');
  n_误差项     := Zl_To_Number(v_Temp);
  v_起始发票号 := 起始发票号_In;

  If v_起始发票号 Is Null Then
    --模拟计算时,可以不传入起始发票号
    If Nvl(领用id_In, 0) <> 0 Then
      Open c_Fact;
      Fetch c_Fact
        Into r_Factrow;
    
      If c_Fact%RowCount <> 0 Then
        If Nvl(r_Factrow.当前号码, '-') <> '-' Then
          v_起始发票号 := Zl_Incstr(r_Factrow.当前号码);
        Else
          v_起始发票号 := r_Factrow.开始号码;
        End If;
      End If;
    End If;
    If v_起始发票号 Is Null Then
      v_起始发票号 := 'J0000001';
    End If;
  End If;

  v_发票号   := v_起始发票号;
  v_发票信息 := Null;

  n_按病人补打票据 := 0;
  n_打印类型       := Null;
  --按单据分配票据
  If 操作类型_In = 3 Or 操作类型_In = 4 Then
    --1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
    If 发票号_In Is Null Then
      v_Err_Msg := '未传入指定的回收票据,不允许' || Case
                     When 操作类型_In = 1 Then
                      '重打票据。'
                     Else
                      '补打票据。'
                   End;
      Raise Err_Item;
    End If;
  
    Select 单据号 Bulk Collect
    Into l_Print_Nos
    From (Select /*+cardinality(j,10)*/
           Distinct c.No As 单据号
           From 票据打印明细 A, 票据使用明细 B, 票据打印内容 C, Table(f_Str2list(发票号_In)) J
           Where a.使用id = b.Id And b.打印id = c.Id And a.票号 = j.Column_Value
           Order By 单据号);
  
    If l_Print_Nos.Count = 0 Then
      v_Err_Msg := '未找到指定发票(' || 发票号_In || '所对应的收费单据!';
      Raise Err_Item;
    End If;
  
    Select /*+cardinality(b,10)*/
     Max(打印类型)
    Into n_打印类型
    From 票据打印内容 A, Table(l_Print_Nos) B
    Where a.No = b.Column_Value And a.数据性质 = 1;
  
    If Nvl(n_打印类型, 0) = 1 Then
      --一次打印有多次结算的，则表示以前为按病人打印的
      n_按病人补打票据 := 1;
      n_打印类型       := 1;
    End If;
  
  Elsif 打印id_In <> 0 Then
    n_按病人补打票据 := 1;
    n_打印类型       := 1;
    Select 单据号 Bulk Collect
    Into l_Print_Nos
    From (Select Distinct NO As 单据号
           From 临时票据打印内容 A
           Where a.Id = 打印id_In And Nvl(a.性质, 0) = 1
           Order By 单据号);
    If l_Print_Nos.Count = 0 Then
      v_Err_Msg := '未找到本次需要分配票据的单据信息(打印ID=' || 打印id_In || ')!';
      Raise Err_Item;
    End If;
  
  Else
    Select Column_Value Bulk Collect Into l_Print_Nos From Table(f_Str2list(Nos_In)) J;
    If l_Print_Nos.Count = 0 Then
      v_Err_Msg := '未找到本次需要分配票据的单据信息(单据信息：' || Nvl(Nos_In, '') || ')!';
      Raise Err_Item;
    End If;
  End If;

  v_Nos := Null;
  If n_启用模式 = 2 Then
    If l_Print_Nos.Count < 3000 Then
      --1.只有自定义模式时，才会涉及可能存在用户调整的情况，加入此判断，主要是为了歉容
      --2.以前不可能超过3000张单据，如果超过3000张单据，需要改造对应的票据,主要适用按病人补打票据的情况
      For I In 1 .. l_Print_Nos.Count Loop
        v_Nos := Nvl(v_Nos, '') || ',' || l_Print_Nos(I);
      
      End Loop;
      v_Nos := Substr(v_Nos, 2);
    End If;
  
    --根据自定义规则分配票号,调用:Zl_Custom_Invoice_Autoallot过程
    Zl_Custom_Invoice_Autoallot(操作类型_In, 模拟计算_In, 票种_In, 领用id_In, 病人id_In, v_Nos, 起始发票号_In, 使用人_In, 使用时间_In, 发票号_In,
                                发票张数_In, n_按病人补打票据, 打印id_In, l_Print_Nos);
    Return;
  End If;

  --参数获取:
  --1.根据预定规则分配票号
  --   NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
  v_Para := Substr(v_Para, Instr(v_Para, '||') + 2);
  If Instr(v_Para, ';') > 0 Then
    --NO:票据是否按单据进行分别打印,1表示按单据打印;0-不按单据打印
    v_Temp       := Substr(v_Para, 1, Instr(v_Para, ';') - 1);
    n_分单据打印 := Zl_To_Number(v_Temp);
    v_Para       := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --执行科室
    v_Temp     := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
    n_执行科室 := Zl_To_Number(v_Temp);
    v_Para     := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --收据费目
    v_Temp     := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
    n_收据费目 := Zl_To_Number(v_Temp);
    v_Para     := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --收据费目
    v_Temp     := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
    n_收费细目 := Zl_To_Number(v_Temp);
    v_Para     := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --执行科室
    v_Temp := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
  Else
    v_Temp := Nvl(v_Para, '0');
  End If;
  n_汇总条件 := Zl_To_Number(v_Temp);

  If n_按病人补打票据 = 1 Then
    --如果打印ID<>0的情况,如果不等零，表示按病人补打发票，则票据将自动不分单据打印，按执行科室打印及收据细目打印
    n_分单据打印 := 0;
    n_执行科室   := 0;
    n_收费细目   := 0;
  End If;

  v_回收票据号 := 发票号_In;
  发票张数_In  := 0;
  --一、首页汇总或不汇总
  If n_汇总条件 <> 2 Then
    Invoice_Split_Notgroup(l_Print_Nos, 发票号_In, v_发票信息, 发票张数_In);
  Else
    --二、分组汇总
    Invoice_Split_Group(l_Print_Nos, 发票号_In, v_发票信息, 发票张数_In);
  End If;
  发票号_In := v_发票信息;
  If 模拟计算_In = 1 Then
    --模拟计算,只返回票据张数和使用的票据号,直接退出
    Return;
  End If;

  v_开始发票号 := Null;
  v_当前发票号 := Null;
  --1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
  If 操作类型_In = 3 Or 操作类型_In = 4 Then
    --收回票据
    Select 使用id Bulk Collect
    Into l_使用id
    From (Select /*+cardinality(j,10)*/
           Distinct b.使用id
           From 票据使用明细 A, 票据打印明细 B, Table(f_Str2list(v_回收票据号)) J
           Where a.Id = b.使用id And b.票号 = j.Column_Value And Nvl(b.票种, 0) = 1);
  
    --插入回收记录
    Forall I In 1 .. l_使用id.Count
      Insert Into 票据使用明细
        (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
        Select 票据使用明细_Id.Nextval, 票种, 号码, 2, Decode(操作类型_In, 3, 4, 2), 领用id, 打印id, 使用人_In, 使用时间_In
        From 票据使用明细
        Where ID = l_使用id(I);
    Forall I In 1 .. l_使用id.Count
      Update 票据打印明细 Set 是否回收 = 1 Where 使用id = l_使用id(I);
  End If;

  If c_Invoce.Count = 0 Then
    --无数据,直接返回
    Return;
  End If;

  If 起始发票号_In Is Null Then
    v_Err_Msg := '未传入起始发票号,不能进行票据分配处理';
    Raise Err_Item;
  End If;

  If Nvl(领用id_In, 0) <> 0 Then
    Open c_Fact;
    Fetch c_Fact
      Into r_Factrow;
    If c_Fact%RowCount = 0 Then
      v_Err_Msg := '无效的票据领用批次，无法完成收费票据分配操作。';
      Close c_Fact;
      Raise Err_Item;
    Elsif Nvl(r_Factrow.剩余数量, 0) < 发票张数_In Then
      v_Err_Msg := '当前批次的剩余数量不足' || 发票张数_In || '张，无法完成收费票据分配操作。';
      Close c_Fact;
      Raise Err_Item;
    End If;
  End If;

  --实际处理票据信息
  If Nvl(n_分单据打印, 0) <> 1 Or Nvl(n_按病人补打票据, 0) = 1 Then
    --不分单据打印时,表示一次打印,打印ID填成一致
    n_打印id := 打印id_In;
    If Nvl(n_打印id, 0) = 0 Then
      Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
    End If;
  End If;

  发票张数_In := 0;
  v_打印内容  := '';
  For c_Invoce_No In (Select Column_Value As 发票号 From Table(f_Str2list(v_发票信息)) Order By 发票号) Loop
    --检查票据范围是否正确
    If Nvl(领用id_In, 0) <> 0 Then
      If Not (Upper(c_Invoce_No.发票号) >= Upper(r_Factrow.开始号码) And Upper(c_Invoce_No.发票号) <= Upper(r_Factrow.终止号码) And
          Length(c_Invoce_No.发票号) = Length(r_Factrow.终止号码)) Then
        v_Err_Msg := '该单据需要打印多张票据,但票据号"' || c_Invoce_No.发票号 || '"超出票据领用的号码范围！';
        Close c_Fact;
        Raise Err_Item;
      End If;
    End If;
  
    --处理票据打印明细
    r_单据号.Delete;
    r_单据序号.Delete;
    l_关联序号.Delete;
  
    Select 票据使用明细_Id.Nextval Into n_使用id From Dual;
  
    n_关联序号 := 0;
  
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).票号 = c_Invoce_No.发票号 Then
        n_关联序号 := c_Invoce(I).关联序号;
        Exit;
      End If;
    End Loop;
    --处理关联票据,以便回收票据
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).关联序号 = n_关联序号 And Nvl(c_Invoce(I).修改标志, 0) = 0 Then
        If n_关联序号 <> 0 Then
          c_Invoce(I).关联序号 := n_使用id;
        End If;
        c_Invoce(I).修改标志 := 1;
      End If;
    End Loop;
  
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).票号 = c_Invoce_No.发票号 Then
        r_单据号.Extend;
        r_单据号(r_单据号.Count) := c_Invoce(I).No;
        r_单据序号.Extend;
        r_单据序号(r_单据序号.Count) := c_Invoce(I).序号;
        l_关联序号.Extend;
        If Nvl(c_Invoce(I).关联序号, 0) <> 0 Then
          --检查是否存在其他的票据
          n_Find := 0;
          For J In 1 .. c_Invoce.Count Loop
            If c_Invoce(I).关联序号 = c_Invoce(J).关联序号 And c_Invoce(I).票号 <> c_Invoce(J).票号 Then
              n_Find := 1;
              Exit;
            End If;
          End Loop;
        
          If n_Find = 0 Then
            l_关联序号(l_关联序号.Count) := Null;
            c_Invoce(I).关联序号 := 0;
          Else
            l_关联序号(l_关联序号.Count) := c_Invoce(I).关联序号;
          End If;
        Else
          l_关联序号(l_关联序号.Count) := Null;
        End If;
      End If;
    End Loop;
  
    --1.处理门打印内容
    If n_分单据打印 = 1 Then
      --分单据打印,需按单据进行处理
      --票据打印内容
      n_Find := 0;
      v_Temp := '';
      For I In 1 .. r_单据号.Count Loop
        v_Temp := v_Temp || ',' || r_单据号(I);
        If Instr(Nvl(v_打印内容, '-') || ',', ',' || r_单据号(I) || ',') > 0 Then
          --已经找到
          n_Find := 1;
        End If;
      End Loop;
      v_打印内容 := v_打印内容 || Nvl(v_Temp, '+');
    
      If Nvl(n_Find, 0) = 0 Then
        Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
        Forall I In 1 .. r_单据号.Count
          Insert Into 票据打印内容 (ID, 数据性质, NO, 打印类型) Values (n_打印id, 1, r_单据号(I), n_打印类型);
        --以便更新门诊费用记录中的实际票号
        v_开始发票号 := c_Invoce_No.发票号;
        Forall I In 1 .. r_单据号.Count
          Update 门诊费用记录 Set 实际票号 = v_开始发票号 Where Mod(记录性质, 10) = 1 And NO = r_单据号(I);
      End If;
    Else
    
      If v_开始发票号 Is Null Then
        --以便更新门诊费用记录中的实际票号
        v_开始发票号 := c_Invoce_No.发票号;
      
        --票据打印内容
        Insert Into 票据打印内容
          (ID, 数据性质, NO, 打印类型)
          Select n_打印id, 1, Column_Value, n_打印类型 From Table(l_Print_Nos);
      
        Update 门诊费用记录
        Set 实际票号 = v_开始发票号
        Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(l_Print_Nos));
      End If;
    End If;
  
    --2.处理票据打印明细
  
    发票张数_In := 发票张数_In + 1;
    --处理票据使用明细
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用人, 使用时间)
    Values
      (n_使用id, 1, c_Invoce_No.发票号, 1, Decode(操作类型_In, 3, 3, 1), Decode(Nvl(领用id_In, 0), 0, Null, 领用id_In), n_打印id,
       使用人_In, 使用时间_In);
  
    Forall I In 1 .. r_单据号.Count
      Insert Into 票据打印明细
        (使用id, 票种, 是否回收, NO, 票号, 序号, 关联票号序号)
      Values
        (n_使用id, 1, 0, r_单据号(I), c_Invoce_No.发票号, r_单据序号(I), l_关联序号(I));
  
    v_当前发票号 := c_Invoce_No.发票号;
  End Loop;

  If Nvl(领用id_In, 0) <> 0 Then
    Close c_Fact;
  
    Update 票据领用记录
    Set 使用时间 = 使用时间_In, 当前号码 = v_当前发票号, 剩余数量 = Nvl(剩余数量, 0) - 发票张数_In
    Where ID = 领用id_In
    Returning 剩余数量 Into n_返回数;
    If n_返回数 < 0 Then
      v_Err_Msg := '当前批次的剩余数量不足' || 发票张数_In || '张，无法完成收费票据分配操作。';
      Raise Err_Item;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Invoice_Autoallot;
/

--108706:刘尔旋,2017-05-08,门诊转住院产生三方卡原样住院预交单
Create Or Replace Procedure Zl_门诊转住院_三方卡结算
(
  No_In         住院费用记录.No%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  退费时间_In   住院费用记录.发生时间%Type,
  门诊退费_In   Number := 0,
  入院科室id_In 住院费用记录.开单部门id%Type := Null,
  主页id_In     住院费用记录.主页id%Type := Null,
  三方退费_In   Number := 0,
  结帐id_In     病人预交记录.结帐id%Type := Null
) As
  v_结帐ids    Varchar2(3000);
  n_组id       财务缴款分组.Id%Type;
  n_退现       Number;
  v_预交no     病人预交记录.No%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  v_Nos        Varchar2(3000);
  v_Info       Varchar2(5000);
  v_当前结算   Varchar2(3000);
  v_原结帐ids  Varchar2(5000);
  n_Tempid     病人预交记录.Id%Type;
  v_流水号     病人预交记录.交易流水号%Type;
  v_说明       病人预交记录.交易说明%Type;
  n_预交id     病人预交记录.Id%Type;
  n_原预交id   病人预交记录.Id%Type;
  n_病人id     病人信息.病人id%Type;
  n_原结帐id   病人预交记录.结帐id%Type;
  n_冲销金额   病人预交记录.冲预交%Type;
  n_卡序号     病人预交记录.卡类别id%Type;
  n_三方卡     Number;
  n_返回值     人员缴款余额.余额%Type;
  v_结算方式   病人预交记录.结算方式%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_卡号       病人预交记录.卡号%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  v_交易说明   病人预交记录.交易说明%Type;
  n_原样退     Number;
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  Procedure Zl_Square_Update
  (
    结帐ids_In    Varchar2,
    现结帐id_In   病人预交记录.结帐id%Type,
    缴款组id_In   病人预交记录.缴款组id%Type,
    退款时间_In   病人预交记录.收款时间%Type,
    结算序号_In   病人预交记录.结算序号%Type,
    结算内容_In   Varchar2 := Null,
    退费金额_In   病人预交记录.冲预交%Type := Null,
    结算卡序号_In 病人预交记录.结算卡序号%Type := Null
  ) As
    n_记录状态 病人卡结算记录.记录状态%Type;
    n_预交id   病人预交记录.Id%Type;
    v_卡号     病人卡结算记录.卡号%Type;
    n_存在卡片 Number;
    d_停用日期 消费卡目录.停用日期%Type;
    n_最大序号 病人卡结算记录.序号%Type;
    n_序号     病人卡结算记录.序号%Type;
    n_余额     消费卡目录.余额%Type;
    n_接口编号 病人卡结算记录.接口编号%Type;
    d_回收时间 消费卡目录.回收时间%Type;
    n_Id       病人预交记录.Id%Type;
  Begin
    n_预交id := 0;
  
    --处理消费卡,结算卡在上面就已经处理了
    For v_校对 In (Select Min(a.Id) As 预交id, c.消费卡id, Sum(c.结算金额) As 结算金额, c.接口编号, c.卡号, Min(c.序号) As 序号, Min(c.Id) As ID
                 From 病人预交记录 A, 病人卡结算对照 B, 病人卡结算记录 C
                 Where a.Id = b.预交id And a.结算卡序号 = 结算卡序号_In And b.卡结算id = c.Id And a.记录性质 = 3 And
                       Instr(Nvl(结算内容_In, '_LXH'), ',' || a.结算方式 || ',') = 0 And
                       a.结帐id In (Select Column_Value From Table(f_Str2list(结帐ids_In)))
                 Group By c.消费卡id, c.接口编号, c.卡号) Loop
    
      If Nvl(v_校对.消费卡id, 0) <> 0 Then
        Select Max(记录状态)
        Into n_记录状态
        From 病人卡结算记录
        Where 接口编号 = v_校对.接口编号 And 消费卡id = Nvl(v_校对.消费卡id, 0) And 卡号 = v_校对.卡号 And Nvl(序号, 0) = Nvl(v_校对.序号, 0);
      Else
        Select Max(记录状态)
        Into n_记录状态
        From 病人卡结算记录
        Where 接口编号 = v_校对.接口编号 And 消费卡id Is Null And 卡号 = v_校对.卡号 And Nvl(序号, 0) = Nvl(v_校对.序号, 0);
      End If;
    
      If n_记录状态 = 1 Then
        n_记录状态 := 2;
      Else
        n_记录状态 := n_记录状态 + 2;
      End If;
      --多条时,只更新一条
      If n_预交id = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id,
           预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
          Select n_预交id, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退款时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
                 -1 * 退费金额_In, 现结帐id_In, 缴款组id_In, 预交类别, 卡类别id, Nvl(结算卡序号, v_校对.接口编号), 卡号, 交易流水号, 交易说明, 合作单位, 2, 结算序号_In,
                 结算性质
          From 病人预交记录 A
          Where ID = v_校对.预交id;
      End If;
    
      If Nvl(v_校对.消费卡id, 0) <> 0 Then
        --消费卡,直接退回卡数据中
        Begin
          Select 卡号, 1, 停用日期, (Select Max(序号) From 消费卡目录 B Where a.卡号 = b.卡号 And a.接口编号 = b.接口编号), 序号, 余额, 接口编号, 回收时间
          Into v_卡号, n_存在卡片, d_停用日期, n_最大序号, n_序号, n_余额, n_接口编号, d_回收时间
          From 消费卡目录 A
          Where ID = v_校对.消费卡id;
        Exception
          When Others Then
            n_存在卡片 := 0;
        End;
      
        --取消停用
        If n_存在卡片 = 0 Then
          v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡被他人删除，不能再启用该卡片,请检查！';
          Raise Err_Item;
        End If;
        If Nvl(n_序号, 0) < Nvl(n_最大序号, 0) Then
          v_Err_Msg := '不能启用历史发卡记录(卡号为"' || v_卡号 || '"),请检查！';
          Raise Err_Item;
        End If;
        If Nvl(d_停用日期, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡已经被他人停用，不能再进行退费,请检查！';
          Raise Err_Item;
        End If;
      
        If d_回收时间 < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡已经回收，不能退费,请检查！';
          Raise Err_Item;
        End If;
        Update 消费卡目录 Set 余额 = Nvl(余额, 0) + 退费金额_In Where ID = Nvl(v_校对.消费卡id, 0);
      End If;
    
      Select 病人卡结算记录_Id.Nextval Into n_Id From Dual;
      Insert Into 病人卡结算记录
        (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
        Select n_Id, 接口编号, 消费卡id, 序号, n_记录状态, 结算方式, -1 * 退费金额_In, 卡号, 交易流水号, 交易时间, 备注,
               Decode(消费卡id, Null, 0, 0, 0, 1) As 标志
        From 病人卡结算记录
        Where ID = v_校对.Id;
      Insert Into 病人卡结算对照 (预交id, 卡结算id) Values (n_预交id, n_Id);
    
      If n_记录状态 <> 2 And n_记录状态 <> 1 Then
        Update 病人卡结算记录 Set 记录状态 = 3 Where ID = v_校对.Id;
      End If;
    End Loop;
  End;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  If 结帐id_In Is Null Then
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  Else
    n_结帐id := 结帐id_In;
  End If;

  Select 结帐id, 病人id
  Into n_原结帐id, n_病人id
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3) And Rownum < 2;

  For r_结账id In (Select Distinct 结帐id
                 From 门诊费用记录
                 Where NO In (Select Distinct NO
                              From 门诊费用记录
                              Where 结帐id In (Select 结帐id
                                             From 病人预交记录
                                             Where 结算序号 In (Select b.结算序号
                                                            From 门诊费用记录 A, 病人预交记录 B
                                                            Where a.No = No_In And b.结算序号 < 0 And Mod(a.记录性质, 10) = 1 And
                                                                  a.记录状态 <> 0 And a.结帐id = b.结帐id))) And
                       Mod(记录性质, 10) = 1 And 记录状态 <> 0
                 Union
                 Select Distinct 结帐id
                 From 门诊费用记录
                 Where NO In (Select Distinct NO
                              From 门诊费用记录
                              Where 结帐id In (Select a.结帐id
                                             From 门诊费用记录 A, 病人预交记录 B
                                             Where a.No = No_In And b.结算序号 > 0 And Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And
                                                   a.结帐id = b.结帐id))) Loop
    v_原结帐ids := v_原结帐ids || ',' || r_结账id.结帐id;
  End Loop;
  v_原结帐ids := Substr(v_原结帐ids, 2);

  Begin
    Select 摘要
    Into v_Info
    From 病人预交记录
    Where 结算方式 Is Null And 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id;
  Exception
    When Others Then
      v_Info := '';
  End;
  --处理卡结算信息
  If v_Info Is Not Null Then
    While v_Info Is Not Null Loop
      v_当前结算 := Substr(v_Info, 1, Instr(v_Info, '|') - 1);
      n_三方卡   := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
      n_卡序号   := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
    
      v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
      n_冲销金额 := -1 * To_Number(v_当前结算);
    
      If n_三方卡 = 0 Then
        --消费卡
        Select 结算方式
        Into v_结算方式
        From 病人预交记录
        Where 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And 结算卡序号 = n_卡序号 And Rownum < 2;
        Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, n_冲销金额, n_卡序号);
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) - n_冲销金额
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
        Returning 余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, v_结算方式, 1, -1 * n_冲销金额);
          n_返回值 := n_冲销金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
        End If;
      Else
        --结算卡
        Select 结算方式, 卡类别id, 卡号, 交易流水号, 交易说明
        Into v_结算方式, n_卡类别id, v_卡号, v_交易流水号, v_交易说明
        From 病人预交记录
        Where 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And 卡类别id = n_卡序号 And Rownum < 2;
      
        If Nvl(门诊退费_In, 0) = 1 Then
          If 三方退费_In = 0 Then
            v_Err_Msg := '存在无法退现的三方账户,无法进行退费!';
            Raise Err_Item;
          End If;
          Update 病人预交记录
          Set 冲预交 = 冲预交 - n_冲销金额
          Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
               缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
            Values
              (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_冲销金额, v_结算方式, Null, 退费时间_In,
               Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, n_卡类别id, Null, v_卡号, v_交易流水号, v_交易说明, Null, n_结帐id,
               -1 * n_结帐id, 0, 3);
          End If;
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) - n_冲销金额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
          Returning 余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, v_结算方式, 1, -1 * n_冲销金额);
            n_返回值 := -1 * n_冲销金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
          End If;
        Else
          Begin
            Select 1 Into n_退现 From 医疗卡类别 Where ID = n_卡类别id And 是否退现 = 1;
          Exception
            When Others Then
              n_退现 := 0;
          End;
        
          If 三方退费_In = 1 Or n_退现 = 0 Then
            v_结算方式 := v_结算方式;
            n_原样退   := 1;
          Else
            n_原样退 := 0;
            Begin
              Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
            Exception
              When Others Then
                Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
            End;
          End If;
        
          If 三方退费_In = 0 Then
            If n_原样退 = 1 Then
              Select 交易流水号, 交易说明, ID
              Into v_流水号, v_说明, n_原预交id
              From 病人预交记录
              Where 结帐id = n_原结帐id And 结算方式 = v_结算方式 And Rownum < 2;
            
              Update 病人预交记录
              Set 冲预交 = 冲预交 - n_冲销金额
              Where 记录性质 = 3 And 记录状态 = 2 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式 And 结帐id = n_结帐id;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_冲销金额, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, n_卡类别id, Null, v_卡号, v_交易流水号, v_交易说明, Null, n_结帐id,
                   -1 * n_结帐id, 0, 3);
              End If;
            
              v_预交no := Nextno(11);
              Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位)
              Values
                (n_预交id, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, n_冲销金额, v_结算方式, Null, 退费时间_In, Null, Null,
                 Null, 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, 2, n_卡类别id, Null, v_卡号, v_流水号, v_说明, Null);
              Update 三方结算交易 Set 交易id = n_预交id Where 交易id = n_原预交id;
            Else
              Update 病人预交记录
              Set 冲预交 = 冲预交 - n_冲销金额
              Where 记录性质 = 3 And 记录状态 = 2 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式 And 结帐id = n_结帐id;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_冲销金额, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, Null, n_结帐id,
                   -1 * n_结帐id, 0, 3);
              End If;
            
              Update 病人预交记录
              Set 金额 = 金额 + n_冲销金额
              Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                v_预交no := Nextno(11);
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 预交类别)
                Values
                  (病人预交记录_Id.Nextval, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, n_冲销金额, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, 2);
              End If;
            End If;
          
            --病人余额
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + n_冲销金额
            Where 性质 = 1 And 病人id = n_病人id And 类型 = 2
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, 2, n_冲销金额, 0);
              n_返回值 := n_冲销金额;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End If;
          --4.2缴款数据处理
          --   因为没有实际收病人的钱,所以不处理
          --部分退费情况，退原预交记录
          If 三方退费_In = 1 Then
            Update 人员缴款余额
            Set 余额 = Nvl(余额, 0) - n_冲销金额
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
            Returning 余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 人员缴款余额
                (收款员, 结算方式, 性质, 余额)
              Values
                (操作员姓名_In, v_结算方式, 1, -1 * n_冲销金额);
              n_返回值 := -1 * n_冲销金额;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 人员缴款余额
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
            End If;
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * n_冲销金额)
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_冲销金额, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, n_卡类别id, Null, v_卡号, v_交易流水号, v_交易说明, Null, n_结帐id,
                 -1 * n_结帐id, 0, 3);
            End If;
          End If;
        End If;
      End If;
      v_Info := Substr(v_Info, Instr(v_Info, '|') + 1);
    End Loop;
  End If;

  Delete From 病人预交记录 Where 结帐id = n_结帐id And 记录状态 = 2 And 结算方式 Is Null;
  Update 门诊费用记录 Set 费用状态 = 0 Where 结帐id = n_结帐id;
  Update 门诊费用记录 Set 费用状态 = 0 Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 2;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_三方卡结算;
/

--108432:冉俊明,2017-05-08,修正固定出诊表临时安排取消审核后没有删除由该临时安排生成的出诊记录，导致清除该临时安排时报错的问题
Create Or Replace Procedure Zl_临床出诊安排_Publish
(
  Id_In       临床出诊表.Id%Type,
  发布人_In   临床出诊表.发布人%Type := Null,
  发布时间_In 临床出诊表.发布时间%Type := Null,
  取消发布_In Number := 0
) As
  --发布和取消发布安排
  --参数：
  --        取消发布_In 是否取消发布
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count    Number(2);
  n_排班方式 临床出诊表.排班方式%Type;
  l_记录id   t_Numlist := t_Numlist();

  d_开始时间 临床出诊安排.开始时间%Type;
  d_终止时间 临床出诊安排.终止时间%Type;

  n_跨月周出诊id 临床出诊表.Id%Type;

  Function Get跨月周出诊id(出诊id_In 临床出诊表.Id%Type) Return 临床出诊表.Id%Type Is
    ----------------------------------------
    --如果原周出诊表不是整周(不足7天)，则需要查找到另一个出诊表构成整周
    ----------------------------------------
    n_出诊id 临床出诊表.Id%Type;
    n_年份   临床出诊表.年份%Type;
    n_月份   临床出诊表.月份%Type;
    n_周数   临床出诊表.周数%Type;
  
    d_开始时间 临床出诊安排.开始时间%Type;
    d_结束时间 临床出诊安排.终止时间%Type;
  
    --根据日期计算当月的周数，以及每一周的时间范围
    Cursor c_Weekrange(Date_In Date) Is
      Select Rownum As 周数, 开始日期, 结束日期
      From (With Month_Range As (Select Trunc(Date_In) As First_Day, Last_Day(Trunc(Date_In)) As Last_Day From Dual)
             Select Decode(To_Char(First_Day, 'day'), '星期日', First_Day, Null) As 开始日期,
                    Decode(To_Char(First_Day, 'day'), '星期日', First_Day, Null) As 结束日期
             From Month_Range
             Union All
             Select Decode(Sign(Trunc(First_Day + 7 * Week, 'day') + 1 - First_Day), 1,
                            Trunc(First_Day + 7 * Week, 'day') + 1, First_Day) As 开始日期,
                    Decode(Sign(Trunc(First_Day + 7 * Week, 'day') + 7 - Last_Day), 1, Last_Day,
                            Trunc(First_Day + 7 * Week, 'day') + 7) As 结束日期
             From Month_Range A, (Select Level - 1 As Week From Dual Connect By Level <= 6) B)
             Where 开始日期 <= 结束日期;
  
  
  Begin
    Begin
      Select 年份, 月份, 周数 Into n_年份, n_月份, n_周数 From 临床出诊表 Where ID = 出诊id_In;
    Exception
      When Others Then
        Return 0;
    End;
  
    If n_年份 Is Null Or n_月份 Is Null Or n_周数 Is Null Then
      Return 0;
    End If;
  
    For r_Weekrange In c_Weekrange(To_Date(n_年份 || '-' || n_月份 || '-01', 'yyyy-mm-dd')) Loop
      If r_Weekrange.周数 = n_周数 Then
        d_开始时间 := r_Weekrange.开始日期;
        d_结束时间 := r_Weekrange.结束日期;
        Exit;
      End If;
    End Loop;
  
    If d_开始时间 Is Null Or d_结束时间 Is Null Then
      Return 0;
    End If;
    If Trunc(d_结束时间) - Trunc(d_开始时间) >= 6 Then
      Return 0;
    End If;
  
    --存在跨月的，查找另一个出诊表的年月周
    n_年份 := Null;
    n_月份 := Null;
    n_周数 := Null;
    If Trunc(d_开始时间 - 1, 'month') <> Trunc(d_开始时间, 'month') Then
      --当前是第一周,获取另一个出诊表的年月
      n_年份 := To_Number(To_Char(d_开始时间 - 1, 'yyyy'));
      n_月份 := To_Number(To_Char(d_开始时间 - 1, 'mm'));
    Elsif Trunc(d_结束时间 + 1, 'month') <> Trunc(d_结束时间, 'month') Then
      --当前是最后一周,获取另一个出诊表的年月
      n_年份 := To_Number(To_Char(d_结束时间 + 1, 'yyyy'));
      n_月份 := To_Number(To_Char(d_结束时间 + 1, 'mm'));
      n_周数 := 1;
    Else
      Return 0;
    End If;
  
    --获取跨月的另一个出诊表的ID
    Begin
      Select ID
      Into n_出诊id
      From (Select Rownum As 行号, ID
             From 临床出诊表
             Where Nvl(排班方式, 0) = 2 And 年份 = n_年份 And 月份 = n_月份 And (n_周数 Is Null Or 周数 = n_周数)
             Order By 周数 Desc)
      Where 行号 < 2;
    Exception
      When Others Then
        Return 0;
    End;
  
    Return n_出诊id;
  End;
Begin
  Begin
    Select Nvl(排班方式, 0) Into n_排班方式 From 临床出诊表 Where ID = Id_In;
  Exception
    When Others Then
      v_Err_Msg := '出诊表信息未找到！';
      Raise Err_Item;
  End;

  If Nvl(取消发布_In, 0) = 0 Then
    --发布安排
    If Nvl(n_排班方式, 0) = 0 Then
      Select Max(1)
      Into n_Count
      From 临床出诊安排 A, 临床出诊限制 B, 临床出诊表 C
      Where a.Id = b.安排id And a.出诊id = c.Id And c.排班方式 = 0 And c.Id = Id_In And Rownum < 2;
      If Nvl(n_Count, 0) = 0 Then
        v_Err_Msg := '当前出诊表无有效的安排，不能发布！';
        Raise Err_Item;
      End If;
    Else
      If Nvl(n_排班方式, 0) = 2 Then
        n_跨月周出诊id := Get跨月周出诊id(Id_In);
      End If;
      Select Max(1)
      Into n_Count
      From 临床出诊安排 A, 临床出诊记录 B, 临床出诊表 C
      Where a.Id = b.安排id And a.出诊id = c.Id And c.排班方式 In (1, 2) And (c.Id = Id_In Or c.Id = n_跨月周出诊id) And Rownum < 2;
      If Nvl(n_Count, 0) = 0 Then
        v_Err_Msg := '当前出诊表无有效的安排，不能发布！';
        Raise Err_Item;
      End If;
    
      Select Max(1)
      Into n_Count
      From 临床出诊记录 A, 临床出诊安排 B
      Where a.号源id = b.号源id And a.出诊日期 Between b.开始时间 And b.终止时间 And a.安排id <> b.Id And b.出诊id = Id_In And Rownum < 2;
      If Nvl(n_Count, 0) <> 0 Then
        v_Err_Msg := '当前出诊表中的部分号源在当前出诊表的生效时间范围内已经存在有效的安排，不能发布！';
        Raise Err_Item;
      End If;
    End If;
  
    --如果存在多个未发布的安排表，则不允许发布后面日期的安排，必须按最小有效时间进行发布
    Select Max(1)
    Into n_Count
    From (Select ID, 年份 || LPad(月份, 2, '0') || 周数 As 日期
           From 临床出诊表
           Where Nvl(排班方式, 0) = Nvl(n_排班方式, 0) And 发布人 Is Null) A,
         (Select ID, 年份 || LPad(月份, 2, '0') || 周数 As 日期 From 临床出诊表 Where ID = Id_In) B
    Where a.日期 < b.日期 And Rownum < 2;
    If Nvl(n_Count, 0) <> 0 Then
      If Nvl(n_排班方式, 0) = 0 Then
        v_Err_Msg := '当前出诊表前面还有未发布的固定出诊表，必须先将其发布或删除后才能发布该出诊表！';
      Elsif Nvl(n_排班方式, 0) = 1 Then
        v_Err_Msg := '当前出诊表前面还有未发布的月出诊表，必须先将其发布或删除后才能发布该出诊表！';
      Elsif Nvl(n_排班方式, 0) = 2 Then
        v_Err_Msg := '当前出诊表前面还有未发布的周出诊表，必须先将其发布或删除后才能发布该出诊表！';
      End If;
      Raise Err_Item;
    End If;
  
    Update 临床出诊表 Set 发布人 = 发布人_In, 发布时间 = 发布时间_In Where ID = Id_In;
    Update 临床出诊安排 Set 审核人 = 发布人_In, 审核时间 = 发布时间_In Where 出诊id = Id_In;
  
    --删除发布时有安排，但是号源已被停用的记录
    For c_安排 In (Select a.Id
                 From 临床出诊安排 A, 临床出诊号源 B, 部门表 C, 人员表 D, 收费项目目录 E
                 Where a.号源id = b.Id And b.科室id = c.Id And a.医生id = d.Id(+) And b.项目id = e.Id And a.出诊id = Id_In And
                       Not (Nvl(b.是否删除, 0) = 0 And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                        Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                        Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                        Nvl(e.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd'))) Loop
      Zl_临床出诊安排_Delete(c_安排.Id, Nvl(n_排班方式, 0));
    End Loop;
  
    If Nvl(n_排班方式, 0) <> 0 Then
      --月安排/周安排根据停诊安排和法定节假日调整出诊记录的出诊/预约情况
      Select 开始时间, 终止时间 Into d_开始时间, d_终止时间 From 临床出诊安排 Where 出诊id = Id_In And Rownum < 2;
      For c_安排 In (Select a.Id, a.号源id, b.日期
                   From 临床出诊安排 A,
                        (Select Trunc(d_开始时间) + Level - 1 As 日期
                          From Dual
                          Connect By Level <= Trunc(d_终止时间) - Trunc(d_开始时间) + 1) B
                   Where a.出诊id = Id_In
                   Order By 号源id, 日期) Loop
      
        Zl_Clinicvisitmodify(c_安排.号源id, c_安排.Id, c_安排.日期, 发布人_In, 发布时间_In);
      End Loop;
    
      --修改临床出诊记录中的"是否发布"
      Select a.Id Bulk Collect
      Into l_记录id
      From 临床出诊记录 A, 临床出诊安排 B
      Where a.安排id = b.Id And b.出诊id = Id_In;
    
      Forall I In 1 .. l_记录id.Count
        Update 临床出诊记录 Set 是否发布 = 1 Where ID = l_记录id(I);
    End If;
    Return;
  End If;

  --==================================================================================================================
  --取消发布
  Select Max(1)
  Into n_Count
  From (Select ID, 年份 || LPad(月份, 2, '0') || 周数 As 日期
         From 临床出诊表
         Where Nvl(排班方式, 0) = Nvl(n_排班方式, 0) And 发布人 Is Not Null) A,
       (Select ID, 年份 || LPad(月份, 2, '0') || 周数 As 日期 From 临床出诊表 Where ID = Id_In) B
  Where a.日期 > b.日期 And Rownum < 2;
  If Nvl(n_Count, 0) <> 0 Then
    If Nvl(n_排班方式, 0) = 0 Then
      v_Err_Msg := '当前出诊后面还有已发布的固定出诊表，必须先将其取消发布后才能取消发布该出诊表！';
    Elsif Nvl(n_排班方式, 0) = 1 Then
      v_Err_Msg := '当前出诊后面还有已发布的月出诊表，必须先将其取消发布后才能取消发布该出诊表！';
    Elsif Nvl(n_排班方式, 0) = 2 Then
      v_Err_Msg := '当前出诊后面还有已发布的周出诊表，必须先将其取消发布后才能取消发布该出诊表！';
    End If;
    Raise Err_Item;
  End If;

  Select Max(1)
  Into n_Count
  From 病人挂号记录 C, 临床出诊记录 A, 临床出诊安排 B
  Where c.出诊记录id = a.Id And a.安排id = b.Id And b.出诊id = Id_In And Rownum < 2;
  If Nvl(n_Count, 0) <> 0 Then
    v_Err_Msg := '当前出诊表的安排已被使用，不允许取消发布！';
    Raise Err_Item;
  End If;

  Update 临床出诊表 Set 发布人 = Null, 发布时间 = Null Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '出诊表信息未找到！';
    Raise Err_Item;
  End If;
  Update 临床出诊安排 Set 审核人 = Null, 审核时间 = Null Where 出诊id = Id_In;

  --固定安排取消发布时删除出诊记录
  If Nvl(n_排班方式, 0) = 0 Then
    --删除出诊记录
    Select a.Id Bulk Collect
    Into l_记录id
    From 临床出诊记录 A, 临床出诊安排 B
    Where a.安排id = b.Id And b.出诊id = Id_In;
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
  Else
    --删除备份的出诊记录
    Select a.Id Bulk Collect
    Into l_记录id
    From 临床出诊记录 A, 临床出诊安排 B
    Where a.安排id = b.Id And b.出诊id = Id_In And a.相关id Is Not Null;
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
  
    --月安排/周安排清除停诊信息，并修改是否发布
    Select a.Id Bulk Collect
    Into l_记录id
    From 临床出诊记录 A, 临床出诊安排 B
    Where a.安排id = b.Id And b.出诊id = Id_In;
  
    Forall I In 1 .. l_记录id.Count
      Delete From 临床出诊停诊记录 Where 记录id = l_记录id(I);
  
    --修改临床出诊记录中的"是否发布"
    Select a.Id Bulk Collect
    Into l_记录id
    From 临床出诊记录 A, 临床出诊安排 B
    Where a.安排id = b.Id And b.出诊id = Id_In;
  
    Forall I In 1 .. l_记录id.Count
      Update 临床出诊记录
      Set 停诊开始时间 = Null, 停诊终止时间 = Null, 停诊原因 = Null, 是否发布 = 0
      Where ID = l_记录id(I);
  
    --恢复临床出诊序号控制的"是否预约"及"是否停诊"
    For c_记录 In (Select a.Id, a.是否分时段, a.是否序号控制
                 From 临床出诊记录 A, 临床出诊安排 B
                 Where a.安排id = b.Id And b.出诊id = Id_In) Loop
      If Nvl(c_记录.是否分时段, 0) = 1 Then
        If Nvl(c_记录.是否序号控制, 0) = 0 Then
          Update 临床出诊序号控制 Set 是否预约 = 1 Where 记录id = c_记录.Id;
        Else
          Update 临床出诊序号控制 Set 是否预约 = Nvl(预约顺序号, 0), 是否停诊 = 0 Where 记录id = c_记录.Id;
        End If;
      End If;
    End Loop;
  
    --换休的不再恢复
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊安排_Publish;
/

--108432:冉俊明,2017-05-08,修正固定出诊表临时安排取消审核后没有删除由该临时安排生成的出诊记录，导致清除该临时安排时报错的问题
Create Or Replace Procedure Zl_临床出诊临时安排_Cancel(安排id_In In 临床出诊安排.Id%Type) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  n_Count  Number(2);
  l_记录id t_Numlist := t_Numlist();
Begin
  Select Count(1)
  Into n_Count
  From 临床出诊记录 A, 病人挂号记录 B
  Where a.Id = b.出诊记录id And a.安排id = 安排id_In And Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '当前安排已存在预约挂号数据，不能取消审核！';
    Raise Err_Item;
  End If;

  Select Count(1)
  Into n_Count
  From 临床出诊安排 A, 临床出诊安排 B, 临床出诊表 C
  Where a.号源id = b.号源id And a.出诊id = c.Id And c.排班方式 = 0 And a.Id <> b.Id And b.Id = 安排id_In And a.登记时间 > b.登记时间 And
        a.审核时间 Is Not Null And Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '该号源在当前安排之后还存在已审核的安排，你不能取消审核当前安排！';
    Raise Err_Item;
  End If;

  Update 临床出诊安排 Set 审核人 = Null, 审核时间 = Null Where ID = 安排id_In And 审核时间 Is Not Null;
  If Sql%NotFound Then
    v_Err_Msg := '当前安排已被他人取消审核或删除，不能再取消审核！';
    Raise Err_Item;
  End If;

  --删除该安排已生成的出诊记录
  Select a.Id Bulk Collect Into l_记录id From 临床出诊记录 A Where a.安排id = 安排id_In;
  Zl_临床出诊记录_Batchdelete(l_记录id);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊临时安排_Cancel;
/

--108667:冉俊明,2017-05-08,修正自动清除划价单时，当同一个医嘱既有药品也有其它项目，而其它项目正在执行时报错
Create Or Replace Procedure Zl_门诊划价记录_Clear(Day_In Number) As
  --功能：自动清除划价单 
  --参数：Day_IN=删除划价后超过Day_IN天未收费的单据 
  Cursor c_Price Is
    Select Distinct a.No, f_List2str(Cast(Collect(To_Char(a.序号)) As t_Strlist)) As 序号
    From 门诊费用记录 A, 未发药品记录 B
    Where a.记录性质 = 1 And a.记录状态 = 0 And a.执行状态 Not In (1, 2) And a.划价人 Is Not Null And a.操作员姓名 Is Null And
          b.单据 In (8, 24) And Nvl(b.已收费, 0) = 0 And a.No = b.No And Nvl(a.执行部门id, 0) = Nvl(b.库房id, 0) And
          Sysdate - b.填制日期 >= Day_In
    Group By a.No;
Begin
  For r_Price In c_Price Loop
    If Not r_Price.序号 Is Null Then
      Zl_门诊划价记录_Delete(r_Price.No, r_Price.序号, 1);
      Commit;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊划价记录_Clear;
/

--97423:余伟节,2017-05-05,处理再入院病人取消登记时找不到数据的问题
Create Or Replace Procedure Zl_入院病案主页_Delete
(
  病人id_In     病案主页.病人id%Type,
  主页id_In     病案主页.主页id%Type,
  转留观_In     Number := 0,
  清除住院号_In Number := 0
  --功能：取消病人入院/预约登记
  --     主页ID_IN:为0时表示取消预约登记
  --     转留观_IN:将正常入院登记病人转为住院留观病人
  --     清除住院号_In:第一次住院的病人转留观时是否清除住院号
) As
  v_入院时间   病案主页.入院日期%Type;
  v_入院科室   病案主页.入院科室id%Type;
  v_出院时间   病案主页.出院日期%Type;
  v_住院号     病案主页.住院号%Type;
  v_再入院     病案主页.再入院%Type;
  v_出院科室id 病案主页.出院科室id%Type;
  n_病人性质   病案主页.病人性质%Type;
  n_主页id     病案主页.主页id%Type;

  v_Count Number;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Select Nvl(状态, 0), Nvl(病人性质, 0)
  Into v_Count, n_病人性质
  From 病案主页
  Where 病人id = 病人id_In And 主页id = 主页id_In;
  If v_Count <> 1 Then
    v_Error := '该病人已经入科,请先将病人撤消至入院状态。';
    Raise Err_Custom;
  End If;

  --删除电子病历时机
  Select 出院科室id, 再入院 Into v_出院科室id, v_再入院 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
  If v_再入院 = 0 Then
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '入院', v_出院科室id);
  Else
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '再次入院', v_出院科室id);
  End If;

  --提取最近一次不为空的住院号
  Begin
    If 主页id_In = 0 Then
      Select 住院号
      Into v_住院号
      From 病案主页
      Where 病人id = 病人id_In And
            主页id =
            (Select Max(主页id) From 病案主页 Where 病人id = 病人id_In And Nvl(主页id, 0) <> 0 And Nvl(住院号, 0) <> 0);
    Else
      Select 住院号
      Into v_住院号
      From 病案主页
      Where 病人id = 病人id_In And
            主页id =
            (Select Max(主页id) From 病案主页 Where 病人id = 病人id_In And 主页id < 主页id_In And Nvl(住院号, 0) <> 0);
    End If;
  Exception
    When Others Then
      Null;
  End;

  If 转留观_In = 1 And Nvl(主页id_In, 0) <> 0 Then
    Update 病案主页
    Set 病人性质 = 2, 住院号 = Decode(清除住院号_In, 1, Null, 住院号)
    Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(病人性质, 0) = 0;
  
    --调整住院次数
    Update 病人信息 Set 住院次数 = Decode(Sign(住院次数 - 1), 1, 住院次数 - 1, Null) Where 病人id = 病人id_In;
    If 清除住院号_In = 1 Then
      Update 病人信息 Set 住院号 = v_住院号 Where 病人id = 病人id_In;
    End If;
  Else
    Begin
      Select b.入院日期, b.出院日期, b.入院科室id
      Into v_入院时间, v_出院时间, v_入院科室
      From 病人信息 A, 病案主页 B
      Where a.病人id = 病人id_In And a.病人id = b.病人id And a.主页id = b.主页id And Nvl(b.主页id, 0) <> 0;
    Exception
      When Others Then
        Null;
    End;
    --撤消预约登记病人不检查住院日报
    If Nvl(主页id_In, 0) <> 0 Then
      Select Zl_住院日报_Count(v_入院科室, v_入院时间) Into v_Count From Dual;
      If v_Count > 0 Then
        v_Error := '已产生业务时间内的住院日报,不能办理该业务!';
        Raise Err_Custom;
      End If;
    End If;
    --门诊留观病人下达入院通知后存在两条有效的病案主页记录（36549）
    Select Count(*) Into v_Count From 病案主页 Where 病人id = 病人id_In And 入院日期 Is Not Null And 出院日期 Is Null;
    If Not v_Count > 1 Then
      v_Count := 0;
      If Nvl(主页id_In, 0) <> 0 And Nvl(n_病人性质, 0) = 0 Then
        v_Count := 1;
      End If;
      --再入院病人,取消入院登记时,病人信息的入院时间和出院时间应该回退到上一次入院日期和出院日期
      If v_再入院 = 1 Then
        Begin
          Select 入院日期, 出院日期
          Into v_入院时间, v_出院时间
          From 病案主页
          Where 病人id = 病人id_In And
                主页id = (Select Max(主页id)
                        From 病案主页
                        Where 病人id = 病人id_In And 主页id < 主页id_In And Nvl(住院号, 0) <> 0);
        Exception
          When Others Then
            --异常处理是为了屏蔽取不到数据的异常情况
            Null;
        End;
      End If;    
      Update 病人信息
      Set 住院号 = v_住院号, 住院次数 = Decode(v_Count, 0, 住院次数, Decode(Sign(住院次数 - 1), 1, 住院次数 - 1, Null)), 当前科室id = Null,
          当前病区id = Null, 当前床号 = Null, 入院时间 = v_入院时间, 出院时间 = v_出院时间, 担保人 = Null, 担保额 = Null, 担保性质 = Null, 在院 = Null
      Where 病人id = 病人id_In;
      Delete From 在院病人 Where 病人id = 病人id_In;
    End If;
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
    Delete From 病人诊断记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 记录来源 = 2;
  
    --本次住院如果交了预交款,改为当作门诊交的
    Update 病人预交记录 Set 主页id = Null Where 病人id = 病人id_In And 主页id = 主页id_In;
  
    --本次发卡的,改变门诊发卡
    Update 住院费用记录 Set 主页id = Null Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 5;
  
    --本次住院的所有费用记录无结算且已全部冲销，则将对应费用记录中的"主页ID"清除。
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From 住院费用记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记帐费用 = 1 And 结帐id Is Not Null;
  
    If v_Count = 0 Then
      Begin
        Select Nvl(Count(*), 0)
        Into v_Count
        From 住院费用记录
        Where 病人id = 病人id_In And 主页id = 主页id_In And 记帐费用 = 1
        Group By NO, 记录性质, 序号
        Having Nvl(Sum(实收金额), 0) <> 0;
      Exception
        When Others Then
          v_Count := 0;
      End;
    
      If v_Count = 0 Then
        Delete 病人未结费用 Where 病人id = 病人id_In And 主页id = 主页id_In And 金额 = 0;
        Update 住院费用记录 Set 主页id = Null Where 病人id = 病人id_In And 主页id = 主页id_In And 记帐费用 = 1;
      End If;
    End If;
  
    --本次住院所有医嘱记录都已作废
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From 病人医嘱记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(医嘱状态, 0) <> 4;
    If v_Count = 0 Then
      Delete From 病人医嘱记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
    End If;
  
    --以下表,没有建病案主页(病人ID,主页ID)的外键,因为其主页ID可能是挂号ID
    Delete From 病人过敏记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 病人诊断记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 病人新生儿记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 电子病历记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 电子病历打印 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    --如果入院发放了就诊卡,则删除会失败(病人费用记录主页ID有外键约束)
    Delete From 病案主页 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    --修改病人信息的主页ID和住院次数
    Select Max(主页id) Into n_主页id From 病案主页 Where 病人id = 病人id_In And Nvl(主页id, 0) <> 0;
    Update 病人信息 Set 主页id = n_主页id Where 病人id = 病人id_In;
    If n_主页id Is Null Then
      Update 病人信息 Set 住院次数 = Null Where 病人id = 病人id_In;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_入院病案主页_Delete;
/

--108667:冉俊明,2017-04-28,修正自动清除划价单时，当同一个医嘱既有药品也有其它项目，而其它项目正在执行时报错
Create Or Replace Procedure Zl_门诊划价记录_Delete
(
  No_In       门诊费用记录.No%Type,
  序号_In     Varchar2 := Null,
  自动清除_In Number := 0
) As
  --功能：删除一张门诊划价单据
  --入参：
  --       序号_In：主要用于门诊医生站作废单个药品
  --      自动清除_in：是否自动清除划价单 zl_门诊划价记录_clear 在调用
  --该光标用于处理药品库存可用数量
  Cursor c_Stock Is
    Select 发药方式, 库房id, 批次, 药品id, 实际数量, 付数, 灭菌效期, 产地, 批号, 效期, ID, 商品条码, 内部条码, 费用id
    From 药品收发记录
    Where 单据 In (8, 24) And NO = No_In And Mod(记录状态, 3) = 1 And 审核人 Is Null And
          费用id In (Select ID
                   From 门诊费用记录
                   Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 And 收费类别 In ('4', '5', '6', '7') And
                         (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
    Order By 药品id;
  --该游标用于处理费用记录序号
  Cursor c_Serial Is
    Select ID, 价格父号 From 门诊费用记录 Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 Order By 序号;

  v_医嘱ids  Varchar2(4000);
  l_医嘱id   t_Numlist := t_Numlist();
  l_药品收发 t_Numlist := t_Numlist();
  v_医嘱id   病人医嘱记录.Id%Type;
  l_费用id   t_Numlist := t_Numlist();
  n_备货卫材 Number;

  n_父号         门诊费用记录.序号%Type;
  n_Count        Number;
  n_医嘱数       Number(5);
  n_已执行_Count Number;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin

  --是否已经删除或收费
  Select Nvl(Count(ID), 0), Sum(Decode(医嘱序号, Null, 0, 1)), Max(医嘱序号), Sum(Decode(Nvl(执行状态, 0), 1, 1, 2, 1, 0))
  Into n_Count, n_医嘱数, v_医嘱id, n_已执行_Count
  From 门诊费用记录
  Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And
        (Instr(',' || 序号_In || ',', ',' || Nvl(价格父号, 序号) || ',') > 0 Or 序号_In Is Null);

  If n_Count = 0 Then
    If Nvl(自动清除_In, 0) = 1 Then
      --自动清除划价单调用时不报错，直接退出
      Return;
    Else
      v_Err_Msg := '要删除的费用记录不存在，可能已经删除或已经收费。';
      Raise Err_Item;
    End If;
  End If;
  --是否已经执行
  If Nvl(n_已执行_Count, 0) > 0 Then
    v_Err_Msg := '要删除的费用记录中包含已执行的内容！';
    Raise Err_Item;
  End If;

  --医嘱费用：检查正在执行的医嘱(注意已执行的情况在下面检查,因为不传 序号_IN 这种情况费用界面已限制)
  --自动清除划价单调用时，由于只会传入药品卫材的对应序号，所以不用检查医嘱；
  --如果检查医嘱，可能同一个医嘱中既有药品，也有其它项目，而其它项目正在执行或已执行时该药品划价记录将删除不掉
  If Nvl(自动清除_In, 0) = 0 Then
    Select Nvl(Count(*), 0)
    Into n_Count
    From 病人医嘱发送
    Where 执行状态 = 3 And (NO, 记录性质, 医嘱id) In
          (Select NO, 记录性质, 医嘱序号
                        From 门诊费用记录
                        Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 And 医嘱序号 Is Not Null And
                              (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null));
    If n_Count > 0 Then
      v_Err_Msg := '要删除的费用中存在对应的医嘱正在执行的情况，不能删除！';
      Raise Err_Item;
    End If;
  End If;

  --药品相关内容
  --先处理备货材料
  For v_出库 In (Select 发药方式, 库房id, 批次, 药品id, 实际数量, 付数, 灭菌效期, 产地, 批号, 效期, ID, 商品条码, 内部条码, 费用id
               From 药品收发记录
               Where 单据 = 21 And Mod(记录状态, 3) = 1 And 审核人 Is Null And
                     费用id In (Select ID
                              From 门诊费用记录
                              Where NO = No_In And 记录性质 = 1 And 记录状态 = 0 And 收费类别 = '4' And
                                    (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
               Order By 药品id) Loop
  
    If v_出库.库房id Is Not Null Then
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) + Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0)
      Where 库房id = v_出库.库房id And 药品id = v_出库.药品id And Nvl(批次, 0) = Nvl(v_出库.批次, 0) And 性质 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 药品库存
          (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码)
        Values
          (v_出库.库房id, v_出库.药品id, 1, v_出库.批次, v_出库.效期,
           Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0), v_出库.批号, v_出库.产地, v_出库.灭菌效期,
           v_出库.商品条码, v_出库.内部条码);
      End If;
    End If;
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := v_出库.Id;
  
    l_费用id.Extend;
    l_费用id(l_费用id.Count) := v_出库.费用id;
  End Loop;

  For r_Stock In c_Stock Loop
  
    If r_Stock.库房id Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_备货卫材
      From Table(l_费用id)
      Where Column_Value = r_Stock.费用id;
      If Nvl(n_备货卫材, 0) = 0 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0)
        Where 库房id = r_Stock.库房id And 药品id = r_Stock.药品id And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码)
          Values
            (r_Stock.库房id, r_Stock.药品id, 1, r_Stock.批次, r_Stock.效期,
             Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0), r_Stock.批号, r_Stock.产地,
             r_Stock.灭菌效期, r_Stock.商品条码, r_Stock.内部条码);
        End If;
      End If;
    End If;
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := r_Stock.Id;
  End Loop;

  --删除药品收发记录
  Forall I In 1 .. l_药品收发.Count
    Delete From 药品收发记录 Where ID = l_药品收发(I);

  ------------------------------------------------------------------------------------------------------------------------
  --批量删未发药品记录
  Delete From 未发药品记录 A
  Where NO = No_In And 单据 In (8, 24) And Not Exists
   (Select 1
         From 药品收发记录
         Where 单据 = a.单据 And Nvl(库房id, 0) = Nvl(a.库房id, 0) And NO = No_In And Mod(记录状态, 3) = 1 And 审核人 Is Null);
  ------------------------------------------------------------------------------------------------------------------------

  --删除病人医嘱附费(最后一次删除时)
  If 序号_In Is Null Then
    --Begin
    --  Select 医嘱序号
    --  Into v_医嘱id
    --  From 门诊费用记录
    --  Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And Rownum = 1;
    -- Exception
    --  When Others Then
    --    Null;
    -- End;
  
    If v_医嘱id Is Not Null Then
      Delete From 病人医嘱附费 Where 医嘱id = v_医嘱id And NO = No_In And 记录性质 = 1;
    End If;
  End If;

  If n_医嘱数 > 0 Then
    If n_医嘱数 = 1 Then
      l_医嘱id.Extend;
      l_医嘱id(l_医嘱id.Count) := v_医嘱id;
    Else
      Select Distinct 医嘱序号 Bulk Collect
      Into l_医嘱id
      From 门诊费用记录
      Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And 医嘱序号 Is Not Null And
            (Instr(',' || 序号_In || ',', ',' || Nvl(价格父号, 序号) || ',') > 0 Or 序号_In Is Null);
    End If;
  End If;

  --门诊费用记录
  Delete From 门诊费用记录
  Where 记录性质 = 1 And 记录状态 = 0 And NO = No_In And
        (Instr(',' || 序号_In || ',', ',' || Nvl(价格父号, 序号) || ',') > 0 Or 序号_In Is Null);
  If Sql%RowCount = 0 Then
    If Nvl(自动清除_In, 0) = 1 Then
      --自动清除划价单调用时不报错，直接退出
      Return;
    Else
      v_Err_Msg := '要删除的费用记录不存在，可能已经删除或已经收费。';
      Raise Err_Item;
    End If;
  End If;

  If 序号_In Is Not Null Then
    --重新调整剩余费用费用记录的序号
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.价格父号 Is Null Then
        n_父号 := n_Count;
      End If;
      Update 门诊费用记录 Set 序号 = n_Count, 价格父号 = Decode(价格父号, Null, Null, n_父号) Where ID = r_Serial.Id;
    
      n_Count := n_Count + 1;
    End Loop;
  End If;
  v_医嘱ids := Null;
  For I In 1 .. l_医嘱id.Count Loop
    v_医嘱ids := Nvl(v_医嘱ids, '') || ',' || l_医嘱id(I);
  End Loop;
  If v_医嘱ids Is Not Null Then
    v_医嘱ids := Substr(v_医嘱ids, 2);
    --场合_In    Integer, --0:门诊;1-住院
    --性质_In    Integer, --1-收费单;2-记帐单
    --操作_In    Integer, --0:删除划价单;1-收费或记帐;2-退费或销帐
    --No_In      门诊费用记录.No%Type,
    --医嘱ids_In Varchar2
    Zl_医嘱发送_计费状态_Update(0, 1, 0, No_In, v_医嘱ids);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊划价记录_Delete;
/

--108251:刘尔旋,2017-04-27,产生划价单的挂号单退号处理
Create Or Replace Procedure Zl_Third_Registdelcheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:HIS退号检查
  --入参:Xml_In:
  --<IN>
  --  <GHDH>A000001</GHDH>    //挂号单号
  --  <JSKLB>支付宝</JSKLB>      //结算卡类别
  --  <JCFP>1</JCFP>            //检查发票
  --  <GHJE>20</GHJE>            //挂号金额
  --  <LSH>34563</LSH>           //交易流水号
  --  <JKFS>0</JKFS>             //缴款方式,0-挂号或预约缴款;1-预约不缴款
  --  <YYFS></YYFS>              //缴款方式=1时传入，预约的预约方式
  --  <XL></XL>                  //险类
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <ERROR><MSG></MSG></ERROR> //为空表示检查成功
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_卡类别     Varchar2(100);
  v_No         病人挂号记录.No%Type;
  n_挂号金额   门诊费用记录.实收金额%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_结算方式   医疗卡类别.结算方式%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  n_存在       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --临时XML
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  n_已开医嘱   Number(2);
  n_检查发票   Number(3);
  n_是否打印   Number(3);
  n_缴款方式   Number(3);
  n_险类       病人信息.险类%Type;
  v_预约方式   病人挂号记录.预约方式%Type;
  v_收费单     门诊费用记录.No%Type;
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS'),
         To_Number(Extractvalue(Value(A), 'IN/XL'))
  Into v_No, v_卡类别, n_挂号金额, v_交易流水号, n_检查发票, n_缴款方式, v_预约方式, n_险类
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Max(收费单) Into v_收费单 From 病人挂号记录 Where NO = v_No;

  n_缴款方式 := Nvl(n_缴款方式, 0);

  If v_卡类别 Is Not Null And n_缴款方式 = 0 Then
    Select Nvl2(Translate(v_卡类别, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
    If v_Type = 'Num' Then
      --传入的是卡类别ID
      Select 结算方式 Into v_结算方式 From 医疗卡类别 Where ID = To_Number(v_卡类别);
    Else
      --传入的是卡类别名称
      Select 结算方式 Into v_结算方式 From 医疗卡类别 Where 名称 = v_卡类别;
    End If;
    If Nvl(n_缴款方式, 0) = 0 Then
      If Nvl(n_险类, 0) = 0 Then
        Select Nvl(Max(1), 0)
        Into n_存在
        From 病人预交记录 A,
             (Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_No And 记录性质 = 4
               Union
               Select Distinct 结帐id
               From 住院费用记录
               Where NO = v_No And 记录性质 = 5
               Union
               Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_收费单 And 记录性质 = 1) B
        Where a.结帐id = b.结帐id And 结算方式 <> v_结算方式 And Mod(记录性质, 10) <> 1 And Rownum < 2;
      Else
        Select Nvl(Max(1), 0)
        Into n_存在
        From 病人预交记录 A,
             (Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_No And 记录性质 = 4
               Union
               Select Distinct 结帐id
               From 住院费用记录
               Where NO = v_No And 记录性质 = 5
               Union
               Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_收费单 And 记录性质 = 1) B, 结算方式 C
        Where a.结帐id = b.结帐id And 结算方式 <> v_结算方式 And Mod(记录性质, 10) <> 1 And a.结算方式 = c.名称 And c.性质 Not In (3, 4) And
              Rownum < 2;
        If n_存在 = 0 Then
          Select Nvl(Max(1), 0)
          Into n_存在
          From 保险结算记录 A,
               (Select Distinct 结帐id
                 From 门诊费用记录
                 Where NO = v_No And 记录性质 = 4
                 Union
                 Select Distinct 结帐id
                 From 住院费用记录
                 Where NO = v_No And 记录性质 = 5
                 Union
                 Select Distinct 结帐id
                 From 门诊费用记录
                 Where NO = v_收费单 And 记录性质 = 1) B
          Where a.记录id = b.结帐id And 险类 <> n_险类 And Rownum < 2;
        End If;
      End If;
      If n_存在 = 1 Then
        v_Err_Msg := '传入的挂号单据包含' || v_结算方式 || '以外的结算方式,无法退号!';
        Raise Err_Item;
      End If;
    Else
      Begin
        Select 1 Into n_存在 From 病人挂号记录 A Where a.No = v_No And a.预约方式 = v_预约方式 And Rownum < 2;
      Exception
        When Others Then
          n_存在 := 0;
      End;
      If n_存在 = 0 Then
        v_Err_Msg := '传入的挂号单据不是' || v_预约方式 || '预约的,无法退号!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If n_缴款方式 = 0 Then
    If v_收费单 Is Null Then
      Select Sum(实收金额) Into n_实收金额 From 门诊费用记录 Where NO = v_No And 记录性质 = 4;
    Else
      Select Sum(实收金额) Into n_实收金额 From 门诊费用记录 Where NO = v_收费单 And 记录性质 = 1;
    End If;
    If n_实收金额 <> n_挂号金额 Then
      v_Err_Msg := '传入的退款金额与实际挂号金额不符，请检查!';
      Raise Err_Item;
    End If;
  End If;

  --补充结算检查，已存在补结算数据的，不能退号
  Begin
    Select 1
    Into n_存在
    From 费用补充记录 A,
         (Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_No And 记录性质 = 4
           Union
           Select Distinct 结帐id
           From 住院费用记录
           Where NO = v_No And 记录性质 = 5
           Union
           Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_收费单 And 记录性质 = 1) B
    Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 1 And Nvl(a.费用状态, 0) <> 2 And Rownum < 2;
  Exception
    When Others Then
      n_存在 := 0;
  End;
  If n_存在 = 1 Then
    v_Err_Msg := '传入的挂号单据已经进行了二次结算,无法退号!';
    Raise Err_Item;
  End If;
  --医嘱检查，已经开过医嘱的，不能退号
  Begin
    Select Distinct 1 Into n_已开医嘱 From 病人医嘱记录 Where 挂号单 = v_No;
  Exception
    When Others Then
      n_已开医嘱 := 0;
  End;
  If n_已开医嘱 = 1 Then
    v_Err_Msg := '传入的挂号单据已经开过医嘱,无法退号!';
    Raise Err_Item;
  End If;
  If Nvl(n_检查发票, 0) = 1 Then
    Select Max(Decode(a.实际票号, Null, 0, 1)) Into n_是否打印 From 门诊费用记录 A Where NO = v_No And 记录性质 = 4;
    If Nvl(n_是否打印, 0) = 1 Then
      v_Err_Msg := '本次退号的单据已开发票,不能退费!';
      Raise Err_Item;
    End If;
    Select Max(Decode(a.实际票号, Null, 0, 1))
    Into n_是否打印
    From 门诊费用记录 A
    Where NO = v_收费单 And 记录性质 = 1;
    If Nvl(n_是否打印, 0) = 1 Then
      v_Err_Msg := '本次退号的单据已开发票,不能退费!';
      Raise Err_Item;
    End If;
  End If;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registdelcheck;
/

--108251:刘尔旋,2017-04-27,产生划价单的挂号单处理
Create Or Replace Procedure Zl_Third_Registdel
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:HIS退号
  --入参:Xml_In:
  --<IN>
  --  <GHDH>A000001</GHDH>    //挂号单号
  --  <JSKLB>支付宝</JSKLB>      //结算卡类别
  --  <JCFP>1</JCFP>            //检查发票
  --  <GHJE>20</GHJE>            //挂号金额
  --  <LSH>34563</LSH>           //交易流水号
  --  <JKFS>0</JKFS>             //缴款方式,0-挂号或预约缴款;1-预约不缴款
  --  <YYFS></YYFS>              //缴款方式=1时传入，预约的预约方式
  --</IN>

  --出参:Xml_Out
  --<OUTPUT>
  -- <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  -- <YJZID>原结帐ID</YJZID>
  -- <CXID>冲销ID</CXID>
  -- <ERROR><MSG></MSG></ERROR> //为空表示取消挂号成功
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_卡类别     Varchar2(100);
  v_No         病人挂号记录.No%Type;
  n_挂号金额   门诊费用记录.实收金额%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_结算方式   医疗卡类别.结算方式%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  n_存在       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --临时XML
  x_Templet    Xmltype; --模板XML
  v_Err_Msg    Varchar2(200);
  n_已开医嘱   Number(2);
  n_检查发票   Number(3);
  n_是否打印   Number(3);
  n_缴款方式   Number(3);
  n_结帐id     门诊费用记录.结帐id%Type;
  n_冲销id     门诊费用记录.结帐id%Type;
  d_登记时间   Date;
  v_预约方式   病人挂号记录.预约方式%Type;
  v_收费单     门诊费用记录.No%Type;
  n_记录状态   门诊费用记录.记录状态%Type;
  n_病人id     门诊费用记录.病人id%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_退费结算   Varchar2(1000);
  Err_Item Exception;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS')
  Into v_No, v_卡类别, n_挂号金额, v_交易流水号, n_检查发票, n_缴款方式, v_预约方式
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Max(收费单) Into v_收费单 From 病人挂号记录 Where NO = v_No;

  n_缴款方式 := Nvl(n_缴款方式, 0);

  If n_缴款方式 = 1 Then
    Begin
      Select 1 Into n_存在 From 门诊费用记录 Where NO = v_No And 记录性质 = 4 And 结帐id Is Not Null And Rownum < 2;
      Select 1
      Into n_存在
      From 门诊费用记录
      Where NO = v_收费单 And 记录性质 = 1 And 结帐id Is Not Null And Rownum < 2;
    Exception
      When Others Then
        n_存在 := 0;
    End;
    If n_存在 = 1 Then
      v_Err_Msg := '传入的挂号单据不是预约挂号单,无法取消预约!';
      Raise Err_Item;
    End If;
    Begin
      Select 1 Into n_存在 From 病人挂号记录 A Where a.No = v_No And a.预约方式 = v_预约方式 And Rownum < 2;
    Exception
      When Others Then
        n_存在 := 0;
    End;
    If n_存在 = 0 Then
      v_Err_Msg := '传入的挂号单据不是' || v_预约方式 || '预约的,无法取消预约!';
      Raise Err_Item;
    End If;
  End If;

  If v_卡类别 Is Not Null And n_缴款方式 = 0 Then
    Select Nvl2(Translate(v_卡类别, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
    If v_Type = 'Num' Then
      --传入的是卡类别ID
      Select 结算方式, ID Into v_结算方式, n_卡类别id From 医疗卡类别 Where ID = To_Number(v_卡类别);
    Else
      --传入的是卡类别名称
      Select 结算方式, ID Into v_结算方式, n_卡类别id From 医疗卡类别 Where 名称 = v_卡类别;
    End If;
  
    Select Sum(实收金额) Into n_实收金额 From 门诊费用记录 Where NO = v_No And 记录性质 = 4;
  
    If Nvl(n_缴款方式, 0) = 0 Then
      --要退的单据不是以该结算卡结算的，则禁止退号
      Begin
        Select 1
        Into n_存在
        From 病人预交记录 A,
             (Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_No And 记录性质 = 4
               Union
               Select Distinct 结帐id
               From 住院费用记录
               Where NO = v_No And 记录性质 = 5
               Union
               Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_收费单 And 记录性质 = 1) B
        Where a.结帐id = b.结帐id And 结算方式 = v_结算方式 And Rownum < 2;
      Exception
        When Others Then
          n_存在 := 0;
      End;
      If n_存在 = 0 Then
        v_Err_Msg := '传入的挂号单据不是' || v_结算方式 || '结算的,无法退号!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --补充结算检查，已存在补结算数据的，不能退号
  Begin
    Select 1
    Into n_存在
    From 费用补充记录 A,
         (Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_No And 记录性质 = 4
           Union
           Select Distinct 结帐id
           From 住院费用记录
           Where NO = v_No And 记录性质 = 5
           Union
           Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_收费单 And 记录性质 = 1) B
    Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 1 And Nvl(a.费用状态, 0) <> 2 And Rownum < 2;
  Exception
    When Others Then
      n_存在 := 0;
  End;
  If n_存在 = 1 Then
    v_Err_Msg := '传入的挂号单据已经进行了二次结算,无法退号!';
    Raise Err_Item;
  End If;
  --医嘱检查，已经开过医嘱的，不能退号
  Begin
    Select Distinct 1 Into n_已开医嘱 From 病人医嘱记录 Where 挂号单 = v_No;
  Exception
    When Others Then
      n_已开医嘱 := 0;
  End;
  If n_已开医嘱 = 1 Then
    v_Err_Msg := '传入的挂号单据已经开过医嘱,无法退号!';
    Raise Err_Item;
  End If;
  If Nvl(n_检查发票, 0) = 1 Then
    Select Max(Decode(a.实际票号, Null, 0, 1)) Into n_是否打印 From 门诊费用记录 A Where NO = v_No And 记录性质 = 4;
    If Nvl(n_是否打印, 0) = 1 Then
      v_Err_Msg := '本次退号的单据已开发票,不能退费!';
      Raise Err_Item;
    End If;
    Select Max(Decode(a.实际票号, Null, 0, 1))
    Into n_是否打印
    From 门诊费用记录 A
    Where NO = v_收费单 And 记录性质 = 1;
    If Nvl(n_是否打印, 0) = 1 Then
      v_Err_Msg := '本次退号的单据已开发票,不能退费!';
      Raise Err_Item;
    End If;
  End If;
  --获取操作员信息
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
  d_登记时间 := Sysdate;

  Zl_三方机构挂号_Delete(v_No, v_交易流水号, '移动平台退号', d_登记时间);

  --同步处理划价单
  If v_收费单 Is Not Null Then
    Select Max(记录状态), Max(病人id) Into n_记录状态, n_病人id From 门诊费用记录 Where NO = v_收费单 And 记录性质 = 1;
    If n_记录状态 = 0 Then
      Zl_门诊划价记录_Delete(v_收费单);
    End If;
    If n_记录状态 = 1 Then
      If v_结算方式 Is Null Then
        v_Err_Msg := '本次挂号单据退款失败,请检查!';
        Raise Err_Item;
      End If;
      Select 病人结帐记录_Id.Nextval Into n_冲销id From Dual;
      Zl_门诊收费记录_销帐(v_收费单, v_操作员编号, v_操作员姓名, Null, d_登记时间, Null, n_冲销id);
    
      v_退费结算 := v_结算方式 || '|' || -1 * n_挂号金额 || '|' || ' |' || ' ';
      Zl_门诊退费结算_Modify(2, n_病人id, n_冲销id, v_退费结算, 0, n_卡类别id, Null, v_交易流水号, Null, 0, 0, 0, 2);
    End If;
  End If;

  If v_收费单 Is Null Then
    Select Max(结帐id) Into n_结帐id From 门诊费用记录 Where NO = v_No And 记录性质 = 4 And 记录状态 = 3;
    Select Max(结帐id) Into n_冲销id From 门诊费用记录 Where NO = v_No And 记录性质 = 4 And 记录状态 = 2;
  Else
    Select Max(结帐id) Into n_结帐id From 门诊费用记录 Where NO = v_收费单 And 记录性质 = 1 And 记录状态 = 3;
    Select Max(结帐id) Into n_冲销id From 门诊费用记录 Where NO = v_收费单 And 记录性质 = 1 And 记录状态 = 2;
  End If;

  v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<YJZID>' || n_结帐id || '</YJZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CXID>' || n_冲销id || '</CXID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registdel;
/

--108251:刘尔旋,2017-04-27,产生划价单的挂号单处理
Create Or Replace Procedure Zl_Third_Getvisitinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:根据挂号单号获取该次就诊详情(医嘱为主要显示)
  --入参:Xml_In:
  --<IN>
  --    <GHDH>挂号单号</GHDH>
  --    <JSKLB>结算卡类别</JSKLB>
  --    <MXGL>明细过滤</MXGL> 0-不过滤,明细包含治疗 1-过滤,明细不包含治疗,默认为1
  --</IN>
  --出参:Xml_Out
  --<OUTPUT>
  --  <GH>
  --     <GHDH>挂号单号</GHDH> //本次查询的挂号单号
  --     <YYSJ>预约时间</YYSJ> //yyyy-mm-dd hh24:mi:ss
  --     <JZSJ></JZSJ>      //实际就诊时间
  --     <DJH></DJH>        //单据号
  --     <JE></JE>          //金额
  --     <DJLX></DJLX>      //单据类型,1-收费单，4-挂号单
  --     <KDSJ></KDSJ>      //开单时间
  --     <JKFS></JKFS>      //缴款方式,0-挂号或预约缴款;1-预约不缴款
  --     <ZFZT></ZFZT>  //支付状态,0-待支付，1-已支付，2-已退费
  --     <SFJSK></SFJSK>    //是否结算卡支付，0-否，1-是
  --  </GH>
  --  <YZLIST>
  --     <YZ>                   //医嘱返回与HIS中显示的内容相同
  --        <YZID><YZID>        //医嘱ID，返回组医嘱ID
  --        <YZLX><YZLX>        //医嘱类型,如处方、检查、检验
  --         <YZMC></YZMC>        //医嘱名称
  --        <ZXKS></ZXKS>       //执行科室
  --        <ZXKSID></ZXKSID>   //执行科室ID
  --        <FYCK></FYCK>       //发药窗口
  --        <YZMX>
  --           <MX>
  --              <YZNR></YZNR>        //医嘱内容
  --              <ZXZT></ZXZT>        //医嘱执行状态
  --              <SFFY>是否发药</SFFY> // 0-否 ，1-是
  --              <GG>规格</GG>
  --              <SL>数量</SL>
  --              <DW>计算单位</DW>
  --              <BZDJ>标准单价</BZDJ>
  --              <YSJE>应收金额</YSJE>
  --              <SSJE>实收金额</SSJE>
  --           </MX>
  --           <MX/>
  --        </YZMX>
  --        <BG></BG>                   //是否已出报告，是否签名
  --        <BGLY></BGLY>               //是否外检项目,1-院内项目，2-外检项目
  --        <BGLYSM></BGLYSM>           //外检项目说明
  --        <JZBG></JZBG>                //禁止显示报告。0-允许，1-禁止
  --        <JZTS></JZTS>                 //提示文字。对于禁止查看的报告，可返回用于提示病人的信息
  --        <BLID></BLID>              //病历ID，如果<BG>字段为1，该值不为空
  --        <DJLIST>
  --           <DJ>                //费用单据信息
  --              <DJH></DJH>      //费用单据号
  --              <DJLX></DJLX>    //单据类型
  --              <JE></JE>        //单据总金额
  --              <KDSJ></KDSJ>    //开单时间
  --              <ZFZT></ZFZT>    //支付状态,0-待支付，1-已支付，2-已退费,3-退费申请中,4-审核通过,5-审核未通过
  --              <SHSM></SHSM>    //审核说明,审核未通过原因
  --              <SFJSK></SFJSK>  //是否结算卡支付，0-否，1-是
  --           </DJ>
  --           <DJ/>
  --        </DJLIST>
  --     </YZ>
  --  </YZLIST>
  --    <ERROR><MSG></MSG></ERROR>                      //如果错误返回
  --</OUTPUT>

  --------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  x_Templet Xmltype; --模板XML

  v_卡类别   Varchar2(100);
  n_卡类别id Number(18);
  v_挂号单   Varchar2(10);
  v_排队号码 Varchar2(10);
  n_Temp     Number(18);
  v_队列名称 排队叫号队列.队列名称%Type;

  n_Count Number(18);

  v_Temp       Varchar2(32767); --临时XML
  v_队列       Varchar2(32767);
  v_No         Varchar2(50);
  n_Add_Djlist Number(1); --是否增加了DJLIST的
  n_性质       Number(2);
  n_组医嘱id   Number(18);
  n_独立医嘱   Number(8);
  n_执行科室id Number(18);
  v_执行科室   Varchar2(50);
  n_退款金额   病人预交记录.冲预交%Type;
  n_明细过滤   Number(3);
  n_退费状态   病人退费申请.状态%Type;
  v_申请原因   病人退费申请.申请原因%Type;
  v_审核原因   病人退费申请.审核原因%Type;
  v_发药窗口   门诊费用记录.发药窗口%Type;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/MXGL')
  Into v_挂号单, v_卡类别, n_明细过滤
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_挂号单 Is Null Then
    v_Err_Msg := '不能找到指定的挂号单号(当前挂号单号为空)';
    Raise Err_Item;
  End If;
  If n_明细过滤 Is Null Then
    n_明细过滤 := 1;
  End If;
  n_Add_Djlist := 0;

  v_Err_Msg := Null;
  If v_卡类别 Is Not Null Then
    Begin
      n_卡类别id := To_Number(v_卡类别);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
  
    If n_卡类别id = 0 Then
      Begin
        Select ID, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_Err_Msg
        From 医疗卡类别
        Where 名称 = v_卡类别;
      Exception
        When Others Then
          v_Err_Msg := '卡类别:' || v_卡类别 || '不存在!';
      End;
    
    Else
    
      Begin
        Select ID, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_Err_Msg
        From 医疗卡类别
        Where ID = n_卡类别id;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  End If;
  n_性质 := 4;
  --1.获取挂号数据
  Begin
    Select 收费单 Into v_No From 病人挂号记录 Where NO = v_挂号单;
  Exception
    When Others Then
      v_No := Null;
  End;

  If v_No Is Not Null Then
    Select Count(*) Into n_Count From 门诊费用记录 Where NO = v_No And 记录性质 = 1;
    If n_Count <> 0 Then
      n_性质 := 1;
    End If;
  End If;
  If n_性质 = 4 Then
    v_No := v_挂号单;
  End If;

  n_Count := 0;
  For c_挂号 In (Select a.Id, v_No As NO, n_性质 As 记录性质, a.执行部门id, c.名称 As 执行部门,
                      To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, To_Char(a.预约时间, 'yyyy-mm-dd hh24:mi:ss') As 预约时间,
                      a.接收时间, To_Char(a.发生时间, 'yyyy-mm-dd HH24:mi:ss') As 就诊时间, a.号别, a.号序, b.金额, a.记录状态,
                      Decode(Nvl(a.执行状态, 0), 0, '等待接诊', 1, '完成就诊', 2, '正在就诊', -1, '取消就诊') As 执行状态,
                      Decode(Nvl(b.结帐id, 0), 0, 0, 1) As 支付标志, Decode(Nvl(a.记录性质, 0), 2, 1, 0) As 缴款方式, b.结帐id As 结帐id
               From 病人挂号记录 A,
                    (Select Max(Decode(记录状态, 0, 0, 2, 0, Nvl(结帐id, 0))) As 结帐id, Sum(实收金额) As 金额
                      From 门诊费用记录 B
                      Where 记录性质 = n_性质 And NO = v_No) B, 部门表 C
               Where a.No = v_挂号单 And a.执行部门id = c.Id(+)) Loop
  
    If Nvl(c_挂号.记录状态, 0) <> 1 Then
      v_Err_Msg := '单据号:' || v_挂号单 || '已经被退号!';
      Raise Err_Item;
    End If;
  
    Begin
      Select 排队号码, 队列名称
      Into v_排队号码, v_队列名称
      From 排队叫号队列
      Where 业务id = c_挂号.Id And Nvl(业务类型, 0) = 0;
    Exception
      When Others Then
        v_排队号码 := Null;
    End;
    If v_排队号码 Is Not Null Then
      --业务id_In ,业务类型_In 排队号码_In Number := Null
      n_Temp := Zl_Getsequencebeforperons(c_挂号.Id, 0, v_排队号码, v_队列名称);
      v_队列 := v_队列 || '<DL><XH>' || v_排队号码 || '</XH><QMRS>' || n_Temp || '</QMRS></DL>';
    End If;
    n_Temp := 0;
    If Nvl(n_卡类别id, 0) <> 0 Then
      Begin
        Select 1
        Into n_Temp
        From 病人预交记录
        Where 结帐id = c_挂号.结帐id And 记录性质 = 4 And 记录状态 In (1, 3) And 卡类别id = n_卡类别id And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    End If;
  
    v_Temp := '<GHDH>' || v_挂号单 || '</GHDH>';
    v_Temp := v_Temp || '<DJH>' || c_挂号.No || '</DJH>';
    v_Temp := v_Temp || '<YYSJ>' || c_挂号.预约时间 || '</YYSJ>';
    v_Temp := v_Temp || '<JZSJ>' || c_挂号.就诊时间 || '</JZSJ>';
    v_Temp := v_Temp || '<KDSJ>' || c_挂号.登记时间 || '</KDSJ>';
    v_Temp := v_Temp || '<JKFS>' || c_挂号.缴款方式 || '</JKFS>';
    v_Temp := v_Temp || '<JE>' || c_挂号.金额 || '</JE>';
    v_Temp := v_Temp || '<DJLX>' || n_性质 || '</DJLX>';
    v_Temp := v_Temp || '<ZFZT>' || c_挂号.支付标志 || '</ZFZT>';
    v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
    If v_队列 Is Not Null Then
      v_Temp := v_Temp || v_队列;
    End If;
    v_Temp := '<GH>' || v_Temp || '</GH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    n_Count := n_Count + 1;
  End Loop;

  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '未找到指定的挂号单据:' || v_挂号单 || '!';
    Raise Err_Item;
  End If;

  --2.组建医嘱及费用相关数据
  n_组医嘱id := 0;

  For c_医嘱 In (With 医嘱费用 As
                  (Select 医嘱id, 发送号, 记录性质, NO, Max(Nvl(执行状态, 0)) As 执行状态
                  From (Select b.医嘱id, b.发送号, b.记录性质, b.No, Nvl(b.执行状态, 0) As 执行状态
                         From 病人医嘱记录 A, 病人医嘱发送 B
                         Where a.挂号单 = v_挂号单 And a.Id = b.医嘱id(+)
                         Union All
                         Select b.医嘱id, b.发送号, b.记录性质, b.No, Nvl(c.执行状态, 0) As 执行状态
                         From 病人医嘱记录 A, 病人医嘱附费 B, 病人医嘱发送 C
                         Where a.挂号单 = v_挂号单 And a.Id = b.医嘱id(+) And b.医嘱id = c.医嘱id(+) And b.发送号 = c.发送号(+))
                  Group By 医嘱id, 发送号, 记录性质, NO)
                 
                 Select Nvl(a.相关id, a.Id) As 组id, Decode(a.相关id, Null, 0, 1) As 附医嘱, a.Id, a.相关id, e.发药窗口,
                        Max(Decode(a.诊疗类别, 'E', Decode(q.操作类型, '2', '处方', '4', '处方', '6', '检验', m.名称), m.名称)) As 医嘱类型,
                        a.执行科室id, d.名称 As 执行科室, Decode(a.相关id, Null, a.医嘱内容, Null) As 组医嘱内容,
                        Max(Decode(a.诊疗类别, '5', 1, '6', 1, '7', 1, 0) * Decode(Nvl(e.执行状态, 0), 1, 1, 3, 1, 0)) As 发药状态,
                        Decode(a.相关id, Null, Null, q.名称) As 明细医嘱内容, s.规格, (e.数次 * e.付数) As 数量, e.计算单位 As 单位,
                        Decode(Nvl(b.执行状态, 0), 0, '未执行', 1, '完全执行', 2, '拒绝执行', 3, '正在执行', '正在执行') As 执行状态,
                        Max(Decode(p.审核时间, Null, Decode(C1.完成时间, Null, 0, 1), 1)) As 是否已出报告, c.病历id, e.No, e.记录性质 As 单据类型,
                        Max(e.标准单价) As 标准单价, Sum(e.应收金额) As 应收金额, Sum(e.实收金额) As 实收金额,
                        To_Char(e.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 开单时间, Decode(Nvl(e.记录状态, 0), 0, 0, 3, 2, 1) As 支付状态,
                        a.病人id
                 
                 From 病人医嘱记录 A, 医嘱费用 B, 病人医嘱报告 C, 电子病历记录 C1, 部门表 D, 门诊费用记录 E, 诊疗项目类别 M, 诊疗项目目录 Q, 收费项目目录 S, 检验标本记录 P
                 Where a.Id = b.医嘱id(+) And a.执行科室id = d.Id(+) And c.病历id = C1.Id(+) And a.Id = c.医嘱id(+) And
                       a.Id = p.医嘱id(+) And b.医嘱id = e.医嘱序号(+) And e.收费细目id = s.Id(+) And b.No = e.No(+) And
                       b.记录性质 = e.记录性质(+) And e.记录状态(+) <> 2 And a.挂号单 = v_挂号单 And a.诊疗类别 = m.编码(+) And
                       a.诊疗项目id = q.Id(+) And a.医嘱状态 In (3, 8)
                 Group By a.Id, a.婴儿, a.序号, a.相关id, e.发药窗口, a.诊疗类别, a.执行科室id, d.名称, a.医嘱内容, q.名称, s.规格, e.数次 * e.付数,
                          e.计算单位, Decode(Nvl(b.执行状态, 0), 0, '未执行', 1, '完全执行', 2, '拒绝执行', 3, '正在执行', '正在执行'), C1.完成时间,
                          Decode(c.病历id, Null, 0, 1), c.病历id, e.No, e.记录性质, e.登记时间, Decode(Nvl(e.记录状态, 0), 0, 0, 3, 2, 1),
                          p.审核时间, a.病人id
                 Order By 组id, 附医嘱, Nvl(a.婴儿, 0), a.序号) Loop
    If Nvl(n_Add_Djlist, 0) = 0 Then
      --增加DJList节点
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<YZLIST></YZLIST>')) Into x_Templet From Dual;
      n_Add_Djlist := 1;
    End If;
  
    If n_组医嘱id <> Nvl(c_医嘱.组id, 0) Then
      n_组医嘱id := Nvl(c_医嘱.组id, 0);
    
      Zl_Third_Custom_Getdeptinfo(n_组医嘱id, n_执行科室id, v_执行科室);
    
      If Nvl(n_执行科室id, 0) = 0 Then
        If c_医嘱.医嘱类型 = '检验' Then
          --检验医嘱以显示采集科室
          n_执行科室id := c_医嘱.执行科室id;
          v_执行科室   := c_医嘱.执行科室;
        Else
          Begin
            Select b.Id, b.名称, c.发药窗口
            Into n_执行科室id, v_执行科室, v_发药窗口
            From 病人医嘱记录 A, 部门表 B, 门诊费用记录 C
            Where a.Id = c.医嘱序号 And a.相关id = n_组医嘱id And a.执行科室id = b.Id And Rownum <= 1;
          Exception
            When Others Then
              n_执行科室id := c_医嘱.执行科室id;
              v_执行科室   := c_医嘱.执行科室;
              v_发药窗口   := c_医嘱.发药窗口;
          End;
        End If;
      End If;
    
      v_Temp := '<YZID>' || n_组医嘱id || '</YZID>';
      v_Temp := v_Temp || '<YZLX>' || c_医嘱.医嘱类型 || '</YZLX>';
      v_Temp := v_Temp || '<YZMC>' || c_医嘱.组医嘱内容 || '</YZMC>';
      v_Temp := v_Temp || '<ZXKS>' || v_执行科室 || '</ZXKS>';
      v_Temp := v_Temp || '<ZXKSID>' || n_执行科室id || '</ZXKSID>';
      v_Temp := v_Temp || '<FYCK>' || v_发药窗口 || '</FYCK>';
      v_Temp := v_Temp || '<BG>' || c_医嘱.是否已出报告 || '</BG>';
      v_Temp := v_Temp || Zl_Third_Custom_Getrptfrom(n_组医嘱id);
      v_Temp := v_Temp || Zl_Third_Custom_Rptlimit(c_医嘱.病人id, n_组医嘱id);
      If Nvl(c_医嘱.是否已出报告, 0) = 1 And c_医嘱.病历id Is Not Null Then
        v_Temp := v_Temp || '<BLID>' || c_医嘱.病历id || '</BLID>';
      End If;
      v_Temp := '<YZ 医嘱ID="' || n_组医嘱id || '">' || v_Temp || '<YZMX></YZMX><DJLIST></DJLIST></YZ>';
      Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      For v_费用 In (
                   
                   Select a.No, Mod(a.记录性质, 10) As 单据类型, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 开单时间,
                           Max(Decode(Nvl(a.记录状态, 0), 0, 0, 3, 2, 1)) As 支付状态, Sum(a.实收金额) As 单据金额, Max(a.结帐id) As 结算卡支付
                   From 门诊费用记录 A
                   Where (a.No, a.记录性质) In
                         (Select Distinct q.No, q.记录性质
                          From 病人医嘱记录 M, 病人医嘱发送 Q
                          Where m.Id = q.医嘱id(+) And (m.Id = n_组医嘱id Or m.相关id = n_组医嘱id)
                          Union All
                          Select Distinct q.No, q.记录性质
                          From 病人医嘱记录 M, 病人医嘱附费 Q
                          Where m.Id = q.医嘱id(+) And (m.Id = n_组医嘱id Or m.相关id = n_组医嘱id)) And
                         Nvl(a.记录状态, 0) In (0, 1, 3)
                   Group By a.No, Mod(a.记录性质, 10), To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss')) Loop
        Begin
          Select 1
          Into n_Temp
          From 病人预交记录 A, 门诊费用记录 B
          Where a.结帐id = b.结帐id And b.No = v_费用.No And Mod(b.记录性质, 10) = 1 And b.记录状态 In (1, 3) And a.卡类别id = n_卡类别id And
                Rownum < 2;
        Exception
          When Others Then
            n_Temp := 0;
        End;
        Begin
          Select -1 * Sum(结帐金额)
          Into n_退款金额
          From 门诊费用记录 B
          Where b.No = v_费用.No And Mod(b.记录性质, 10) = 1 And b.记录状态 = 2;
        Exception
          When Others Then
            n_退款金额 := 0;
        End;
        Begin
          Select 状态, 申请原因, 审核原因
          Into n_退费状态, v_申请原因, v_审核原因
          From 病人退费申请
          Where NO = v_费用.No And Mod(记录性质, 10) = Mod(v_费用.单据类型, 10);
        Exception
          When Others Then
            n_退费状态 := -1;
            v_申请原因 := '';
            v_审核原因 := '';
        End;
      
        v_Temp := '<DJH>' || v_费用.No || '</DJH>';
        v_Temp := v_Temp || '<DJLX>' || v_费用.单据类型 || '</DJLX>';
        v_Temp := v_Temp || '<JE>' || v_费用.单据金额 || '</JE>';
        v_Temp := v_Temp || '<KDSJ>' || v_费用.开单时间 || '</KDSJ>';
        If n_退费状态 = -1 Then
          v_Temp := v_Temp || '<ZFZT>' || v_费用.支付状态 || '</ZFZT>';
        Else
          If n_退费状态 = 0 Then
            v_Temp := v_Temp || '<ZFZT>3</ZFZT>';
          End If;
          If n_退费状态 = 1 Then
            If v_费用.支付状态 = 2 Then
              v_Temp := v_Temp || '<ZFZT>2</ZFZT>';
            Else
              v_Temp := v_Temp || '<ZFZT>4</ZFZT>';
            End If;
          End If;
          If n_退费状态 = 2 Then
            v_Temp := v_Temp || '<ZFZT>5</ZFZT>';
          End If;
        End If;
      
        If n_退费状态 = -1 Then
          v_Temp := v_Temp || '<SHSM>' || '' || '</SHSM>';
        Else
          If n_退费状态 = 0 Then
            v_Temp := v_Temp || '<SHSM>' || v_申请原因 || '</SHSM>';
          End If;
          If n_退费状态 = 1 Then
            v_Temp := v_Temp || '<SHSM>' || v_审核原因 || '</SHSM>';
          End If;
          If n_退费状态 = 2 Then
            v_Temp := v_Temp || '<SHSM>' || v_审核原因 || '</SHSM>';
          End If;
        End If;
      
        v_Temp := v_Temp || '<YTJE>' || Nvl(n_退款金额, 0) || '</YTJE>';
        v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
        v_Temp := '<DJ>' || v_Temp || '</DJ>';
        Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@医嘱ID="' || n_组医嘱id || '"]/DJLIST', Xmltype(v_Temp))
        Into x_Templet
        From Dual;
      End Loop;
    End If;
  
    --只有一条记录的医嘱，在明细中增加该条医嘱，以获取执行状态
    Select Decode(Count(*), 0, 1, 0) Into n_独立医嘱 From 病人医嘱记录 Where 相关id = n_组医嘱id;
    If n_独立医嘱 = 1 Then
      v_Temp := '<YZNR>' || c_医嘱.组医嘱内容 || '</YZNR>';
      v_Temp := v_Temp || '<GG>' || c_医嘱.规格 || '</GG>';
      v_Temp := v_Temp || '<SFFY>' || c_医嘱.发药状态 || '</SFFY>';
      v_Temp := v_Temp || '<SL>' || c_医嘱.数量 || '</SL>';
      v_Temp := v_Temp || '<DW>' || c_医嘱.单位 || '</DW>';
      v_Temp := v_Temp || '<BZDJ>' || Nvl(c_医嘱.标准单价, 0) || '</BZDJ>';
      v_Temp := v_Temp || '<YSJE>' || Nvl(c_医嘱.应收金额, 0) || '</YSJE>';
      v_Temp := v_Temp || '<SSJE>' || Nvl(c_医嘱.实收金额, 0) || '</SSJE>';
      v_Temp := v_Temp || '<ZXZT>' || c_医嘱.执行状态 || '</ZXZT>';
      v_Temp := '<MX>' || v_Temp || '</MX>';
      Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@医嘱ID="' || n_组医嘱id || '"]/YZMX', Xmltype(v_Temp))
      Into x_Templet
      From Dual;
    End If;
  
    If Nvl(c_医嘱.附医嘱, 0) = 1 Then
      If n_明细过滤 = 0 Or (n_明细过滤 = 1 And c_医嘱.医嘱类型 <> '治疗') Then
        v_Temp := '<YZNR>' || c_医嘱.明细医嘱内容 || '</YZNR>';
        v_Temp := v_Temp || '<GG>' || c_医嘱.规格 || '</GG>';
        v_Temp := v_Temp || '<SL>' || c_医嘱.数量 || '</SL>';
        v_Temp := v_Temp || '<DW>' || c_医嘱.单位 || '</DW>';
        v_Temp := v_Temp || '<SFFY>' || c_医嘱.发药状态 || '</SFFY>';
        v_Temp := v_Temp || '<ZXZT>' || c_医嘱.执行状态 || '</ZXZT>';
        v_Temp := v_Temp || '<BZDJ>' || Nvl(c_医嘱.标准单价, 0) || '</BZDJ>';
        v_Temp := v_Temp || '<YSJE>' || Nvl(c_医嘱.应收金额, 0) || '</YSJE>';
        v_Temp := v_Temp || '<SSJE>' || Nvl(c_医嘱.实收金额, 0) || '</SSJE>';
        v_Temp := '<MX>' || v_Temp || '</MX>';
        Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@医嘱ID="' || n_组医嘱id || '"]/YZMX', Xmltype(v_Temp))
        Into x_Templet
        From Dual;
      End If;
    End If;
  
  End Loop;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getvisitinfo;
/

--108594:刘尔旋,2017-04-25,门诊转住院单张单据退费问题
Create Or Replace Procedure Zl_门诊转住院_收费转出
(
  No_In         住院费用记录.No%Type,
  操作员编号_In 住院费用记录.操作员编号%Type,
  操作员姓名_In 住院费用记录.操作员姓名%Type,
  退费时间_In   住院费用记录.发生时间%Type,
  门诊退费_In   Number := 0,
  入院科室id_In 住院费用记录.开单部门id%Type := Null,
  主页id_In     住院费用记录.主页id%Type := Null,
  结算方式_In   病人预交记录.结算方式%Type := Null,
  结帐id_In     病人预交记录.结帐id%Type := Null,
  原结帐id_In   病人预交记录.结帐id%Type := Null,
  误差费_In     病人预交记录.冲预交%Type := Null
) As
  --门诊退费_In:0-门诊转住院立即销帐;1-门诊退费模式
  -- 门诊退费_In为1时:入院科室id_In和主页ID_IN可以不传入
  n_Count      Number(5);
  n_原结帐id   住院费用记录.结帐id%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  n_预交使用额 病人预交记录.冲预交%Type;
  n_实际冲销   病人预交记录.冲预交%Type;
  n_组id       财务缴款分组.Id%Type;
  n_病人id     病人信息.病人id%Type;
  v_预交no     病人预交记录.No%Type;
  n_预交金额   病人预交记录.冲预交%Type;
  n_打印id     票据使用明细.打印id%Type;
  n_开单部门id 住院费用记录.开单部门id%Type;
  v_开单人     门诊费用记录.开单人%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_误差费     病人预交记录.冲预交%Type;
  v_误差费     结算方式.名称%Type;
  n_返回值     病人余额.费用余额%Type;
  v_结算方式   结算方式.名称%Type;
  v_Nos        Varchar2(3000);
  v_结帐ids    Varchar2(3000);
  v_原结帐ids  Varchar2(3000);
  n_Tempid     病人预交记录.Id%Type;
  n_预交id     病人预交记录.Id%Type;
  n_医保       Number;
  n_存在       Number;
  n_退现       Number;
  n_部分退费   Number;
  n_退费条数   Number;
  n_异常标志   Number;
  n_计算误差   Number;
  n_费用状态   门诊费用记录.费用状态%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  Procedure Zl_Square_Update
  (
    结帐ids_In    Varchar2,
    现结帐id_In   病人预交记录.结帐id%Type,
    缴款组id_In   病人预交记录.缴款组id%Type,
    退款时间_In   病人预交记录.收款时间%Type,
    结算序号_In   病人预交记录.结算序号%Type,
    结算内容_In   Varchar2 := Null,
    退费金额_In   病人预交记录.冲预交%Type := Null,
    结算卡序号_In 病人预交记录.结算卡序号%Type := Null
  ) As
    n_记录状态 病人卡结算记录.记录状态%Type;
    n_预交id   病人预交记录.Id%Type;
    v_卡号     病人卡结算记录.卡号%Type;
    n_存在卡片 Number;
    d_停用日期 消费卡目录.停用日期%Type;
    n_最大序号 病人卡结算记录.序号%Type;
    n_序号     病人卡结算记录.序号%Type;
    n_余额     消费卡目录.余额%Type;
    n_接口编号 病人卡结算记录.接口编号%Type;
    d_回收时间 消费卡目录.回收时间%Type;
    n_Id       病人预交记录.Id%Type;
  Begin
    n_预交id := 0;
  
    --处理消费卡,结算卡在上面就已经处理了
    For v_校对 In (Select Min(a.Id) As 预交id, c.消费卡id, Sum(c.结算金额) As 结算金额, c.接口编号, c.卡号, Max(c.序号) As 序号, Max(c.Id) As ID
                 From 病人预交记录 A, 病人卡结算对照 B, 病人卡结算记录 C
                 Where a.Id = b.预交id And a.结算卡序号 = 结算卡序号_In And b.卡结算id = c.Id And a.记录性质 = 3 And
                       Instr(Nvl(结算内容_In, '_LXH'), ',' || a.结算方式 || ',') = 0 And
                       a.结帐id In (Select Column_Value From Table(f_Str2list(结帐ids_In)))
                 Group By c.消费卡id, c.接口编号, c.卡号) Loop
    
      If Nvl(v_校对.消费卡id, 0) <> 0 Then
        Select Max(记录状态)
        Into n_记录状态
        From 病人卡结算记录
        Where 接口编号 = v_校对.接口编号 And 消费卡id = Nvl(v_校对.消费卡id, 0) And 卡号 = v_校对.卡号 And Nvl(序号, 0) = Nvl(v_校对.序号, 0);
      Else
        Select Max(记录状态)
        Into n_记录状态
        From 病人卡结算记录
        Where 接口编号 = v_校对.接口编号 And 消费卡id Is Null And 卡号 = v_校对.卡号 And Nvl(序号, 0) = Nvl(v_校对.序号, 0);
      End If;
    
      If n_记录状态 = 1 Then
        n_记录状态 := 2;
      Else
        n_记录状态 := n_记录状态 + 2;
      End If;
      --多条时,只更新一条
      If n_预交id = 0 Then
        Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      
        Insert Into 病人预交记录
          (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id,
           预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 校对标志, 结算序号, 结算性质)
          Select n_预交id, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退款时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
                 -1 * 退费金额_In, 现结帐id_In, 缴款组id_In, 预交类别, 卡类别id, Nvl(结算卡序号, v_校对.接口编号), 卡号, 交易流水号, 交易说明, 合作单位, 2, 结算序号_In,
                 结算性质
          From 病人预交记录 A
          Where ID = v_校对.预交id;
      End If;
    
      If Nvl(v_校对.消费卡id, 0) <> 0 Then
        --消费卡,直接退回卡数据中
        Begin
          Select 卡号, 1, 停用日期, (Select Max(序号) From 消费卡目录 B Where a.卡号 = b.卡号 And a.接口编号 = b.接口编号), 序号, 余额, 接口编号, 回收时间
          Into v_卡号, n_存在卡片, d_停用日期, n_最大序号, n_序号, n_余额, n_接口编号, d_回收时间
          From 消费卡目录 A
          Where ID = v_校对.消费卡id;
        Exception
          When Others Then
            n_存在卡片 := 0;
        End;
      
        --取消停用
        If n_存在卡片 = 0 Then
          v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡被他人删除，不能再启用该卡片,请检查！';
          Raise Err_Item;
        End If;
        If Nvl(n_序号, 0) < Nvl(n_最大序号, 0) Then
          v_Err_Msg := '不能启用历史发卡记录(卡号为"' || v_卡号 || '"),请检查！';
          Raise Err_Item;
        End If;
        If Nvl(d_停用日期, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡已经被他人停用，不能再进行退费,请检查！';
          Raise Err_Item;
        End If;
      
        If d_回收时间 < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '卡号为"' || v_卡号 || '"的消费卡已经回收，不能退费,请检查！';
          Raise Err_Item;
        End If;
        Update 消费卡目录 Set 余额 = Nvl(余额, 0) + 退费金额_In Where ID = Nvl(v_校对.消费卡id, 0);
      End If;
    
      Select 病人卡结算记录_Id.Nextval Into n_Id From Dual;
      Insert Into 病人卡结算记录
        (ID, 接口编号, 消费卡id, 序号, 记录状态, 结算方式, 结算金额, 卡号, 交易流水号, 交易时间, 备注, 结算标志)
        Select n_Id, 接口编号, 消费卡id, 序号, n_记录状态, 结算方式, -1 * 退费金额_In, 卡号, 交易流水号, 交易时间, 备注,
               Decode(消费卡id, Null, 0, 0, 0, 1) As 标志
        From 病人卡结算记录
        Where ID = v_校对.Id;
      Insert Into 病人卡结算对照 (预交id, 卡结算id) Values (n_预交id, n_Id);
    
      If n_记录状态 <> 2 And n_记录状态 <> 1 Then
        Update 病人卡结算记录 Set 记录状态 = 3 Where ID = v_校对.Id;
      End If;
    End Loop;
  End;
Begin
  n_组id := Zl_Get组id(操作员姓名_In);
  --误差费
  Begin
    Select 名称 Into v_误差费 From 结算方式 Where 性质 = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := '没有发现误差结算方式，请检查是否正确设置！';
      Raise Err_Item;
  End;

  If 原结帐id_In Is Null Then
  
    Select Count(NO), Sum(实收金额) Into n_Count, n_实收金额 From 门诊费用记录 Where NO = No_In And Mod(记录性质,10) = 1;
    If n_Count = 0 Or n_实收金额 = 0 Then
      v_Err_Msg := '单据' || No_In || '不是收费单据或因并发原因他人操作了该单据,不能转为住院费用.';
      Raise Err_Item;
    End If;
  
    Select 结帐id, 病人id, 开单部门id, 开单人
    Into n_原结帐id, n_病人id, n_开单部门id, v_开单人
    From 门诊费用记录
    Where NO = No_In And Mod(记录性质,10) = 1 And 记录状态 In (1, 3) And Rownum < 2;
  
    --1.1作废费用记录
    If 结帐id_In Is Null Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    Else
      n_结帐id := 结帐id_In;
    End If;
  
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
       计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 执行时间,
       操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id, 费用状态)
      Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id,
             收费类别, 收费细目id, 计算单位, 付数, 发药窗口, -1 * 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, -1 * 应收金额, -1 * 实收金额, 开单部门id,
             开单人, 执行部门id, 划价人, 执行人, -1, 执行时间, 操作员编号_In, 操作员姓名_In, 发生时间, 退费时间_In, n_结帐id, -1 * 结帐金额, 保险项目否, 保险大类id, 统筹金额,
             摘要, Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id, 0
      From 门诊费用记录
      Where NO = No_In And Mod(记录性质,10) = 1 And 记录状态 = 1;
  
    --Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
  
    --1.2作废预交记录
    --作废冲预交部分
    For r_结账id In (Select Distinct 结帐id
                   From 门诊费用记录
                   Where NO In (Select Distinct NO
                                From 门诊费用记录
                                Where 结帐id In (Select 结帐id
                                               From 病人预交记录
                                               Where 结算序号 In (Select b.结算序号
                                                              From 门诊费用记录 A, 病人预交记录 B
                                                              Where a.No = No_In And b.结算序号 < 0 And Mod(a.记录性质, 10) = 1 And
                                                                    a.记录状态 <> 0 And a.结帐id = b.结帐id))) And
                         Mod(记录性质, 10) = 1 And 记录状态 <> 0
                   Union
                   Select Distinct 结帐id
                   From 门诊费用记录
                   Where NO In (Select Distinct NO
                                From 门诊费用记录
                                Where 结帐id In (Select a.结帐id
                                               From 门诊费用记录 A, 病人预交记录 B
                                               Where a.No = No_In And b.结算序号 > 0 And Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 And
                                                     a.结帐id = b.结帐id)) And Mod(记录性质, 10) = 1 And 记录状态 <> 0) Loop
      v_原结帐ids := v_原结帐ids || ',' || r_结账id.结帐id;
    End Loop;
    v_原结帐ids := Substr(v_原结帐ids, 2);
  
    Begin
      Select 1
      Into n_医保
      From 保险结算记录
      Where 记录id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And Rownum < 2;
    Exception
      When Others Then
        n_医保 := 0;
    End;
  
    If n_医保 = 1 Then
      Begin
        Select 1
        Into n_存在
        From 医保结算明细
        Where NO = No_In And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And Rownum < 2;
      Exception
        When Others Then
          v_Err_Msg := '当前单据' || No_In || '不存在医保结算明细,无法进行门诊转住院!';
          Raise Err_Item;
      End;
    End If;
  
    --医保退款
    For r_医保 In (Select 结帐id, NO, 结算方式, 金额, 备注
                 From 医保结算明细
                 Where NO = No_In And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids)))) Loop
      Update 人员缴款余额
      Set 余额 = Nvl(余额, 0) - r_医保.金额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_医保.结算方式
      Returning 余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 人员缴款余额
          (收款员, 结算方式, 性质, 余额)
        Values
          (操作员姓名_In, r_医保.结算方式, 1, -1 * r_医保.金额);
        n_返回值 := r_医保.金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 人员缴款余额
        Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_医保.结算方式 And Nvl(余额, 0) = 0;
      End If;
    
      Update 病人预交记录
      Set 冲预交 = 冲预交 + (-1 * r_医保.金额)
      Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = r_医保.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
           缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
        Values
          (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_医保.金额, r_医保.结算方式, Null, 退费时间_In,
           Null, Null, Null, 操作员编号_In, 操作员姓名_In, r_医保.备注, n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id,
           0, 3);
      End If;
    
      Update 病人预交记录
      Set 记录状态 = 3
      Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
            结算方式 = r_医保.结算方式;
    
      Update 医保结算明细
      Set 金额 = 金额 + (-1 * r_医保.金额)
      Where NO = No_In And 结帐id = n_结帐id And 结算方式 = r_医保.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 医保结算明细
          (结帐id, NO, 结算方式, 金额)
        Values
          (n_结帐id, No_In, r_医保.结算方式, -1 * r_医保.金额);
      End If;
      n_实收金额 := n_实收金额 - r_医保.金额;
    End Loop;
  
    Begin
      Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
    Exception
      When Others Then
        Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
    End;
  
    If n_实收金额 <> 0 Then
      For r_Prepay In (Select NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, Sum(冲预交) As 冲预交, 卡类别id, 结算卡序号,
                              卡号, 交易流水号, 交易说明, 合作单位
                       From 病人预交记录 A
                       Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids)))
                       Group By n_Tempid, NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 卡类别id, 结算卡序号, 卡号,
                                交易流水号, 交易说明, 合作单位) Loop
        If n_实收金额 <> 0 Then
          If r_Prepay.冲预交 >= n_实收金额 Then
            Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 缴款组id)
              Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                     r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                     操作员编号_In, -1 * n_实收金额, n_结帐id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                     r_Prepay.交易说明, r_Prepay.合作单位, 1, -1 * n_结帐id, n_组id
              From Dual;
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(n_实收金额, 0)
            Where 病人id = n_病人id And 类型 = 1 And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, n_实收金额, 1);
              n_返回值 := n_实收金额;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
            n_实收金额 := 0;
          Else
            Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 缴款组id)
              Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                     r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                     操作员编号_In, -1 * r_Prepay.冲预交, n_结帐id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                     r_Prepay.交易说明, r_Prepay.合作单位, 1, -1 * n_结帐id, n_组id
              From Dual;
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(r_Prepay.冲预交, 0)
            Where 病人id = n_病人id And 类型 = 1 And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, r_Prepay.冲预交, 1);
              n_返回值 := r_Prepay.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
            n_实收金额 := n_实收金额 - r_Prepay.冲预交;
          End If;
        End If;
      End Loop;
    End If;
    --2.票据收回
    --可能以前没有打印,无收回
    Select Nvl(Max(ID), 0)
    Into n_打印id
    From (Select b.Id
           From 票据使用明细 A, 票据打印内容 B
           Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = No_In
           Order By a.使用时间 Desc)
    Where Rownum < 2;
    If n_打印id > 0 Then
      --多张单据循环调用时只能收回一次
      Select Count(打印id) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
      If n_Count = 0 Then
        Insert Into 票据使用明细
          (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
          Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 退费时间_In, 操作员姓名_In
          From 票据使用明细
          Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
      End If;
    End If;
  
    --3.缴款数据处理(
    --   现有两种情况:
    --    1. 转出过程直接销帐的,则缴款数据不增加;
    --    2. 先转出,再到门诊退款退票,则需要进行缴款数据处理
    If Nvl(门诊退费_In, 0) = 1 Then
      For c_预交 In (Select a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, Min(a.交易流水号) As 交易流水号,
                          Min(a.交易说明) As 交易说明, Min(a.合作单位) As 合作单位, b.性质
                   From 病人预交记录 A, 结算方式 B
                   Where a.记录性质 = 3 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                         a.结算方式 = b.名称 And b.性质 In (1, 2, 7, 8) And a.结算方式 Is Not Null
                   Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质
                   Having Sum(a.冲预交) <> 0
                   Order By a.卡类别id, 性质 Desc) Loop
        If n_实收金额 <> 0 Then
          Begin
            Select 是否退现 Into n_退现 From 医疗卡类别 Where ID = c_预交.卡类别id;
          Exception
            When Others Then
              n_退现 := 0;
          End;
          If (c_预交.性质 = 7 Or (c_预交.性质 = 8 And c_预交.卡类别id Is Not Null)) And n_退现 = 0 Then
            If c_预交.冲预交 > n_实收金额 Then
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '1' || ',' || c_预交.卡类别id || ',' || -1 * n_实收金额 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || c_预交.卡类别id || ',' || -1 * n_实收金额 || '|', n_组id,
                   Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = c_预交.结算方式;
              n_费用状态 := 1;
              n_实收金额 := 0;
            Else
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|', n_组id,
                   Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
            
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = c_预交.结算方式;
              n_费用状态 := 1;
              n_实收金额 := n_实收金额 - c_预交.冲预交;
            End If;
          Else
            n_实际冲销 := 0;
            If c_预交.性质 In (3, 4) Or (c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null) Then
              v_结算方式 := c_预交.结算方式;
            Else
              If 结算方式_In Is Null Then
                Begin
                  Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
                Exception
                  When Others Then
                    Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
                End;
              Else
                v_结算方式 := 结算方式_In;
              End If;
            End If;
          
            If c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null Then
              If n_实收金额 >= c_预交.冲预交 Then
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, c_预交.冲预交, c_预交.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null,
                     退费时间_In, Null, Null, Null, 操作员编号_In, 操作员姓名_In,
                     '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|', n_组id, Null, Null, Null, Null, Null, Null,
                     n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := c_预交.冲预交;
              Else
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, n_实收金额, c_预交.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '0' || ',' || c_预交.结算卡序号 || ',' || -1 * n_实收金额 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                     Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || c_预交.结算卡序号 || ',' || -1 * n_实收金额 || '|', n_组id,
                     Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := n_实收金额;
              End If;
            Else
              If c_预交.冲预交 > n_实收金额 Then
                n_实际冲销 := n_实收金额;
              Else
                n_实际冲销 := c_预交.冲预交;
              End If;
            End If;
          
            If c_预交.结算卡序号 Is Null Then
              Update 人员缴款余额
              Set 余额 = Nvl(余额, 0) - n_实际冲销
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
              Returning 余额 Into n_返回值;
              If Sql%RowCount = 0 Then
                Insert Into 人员缴款余额
                  (收款员, 结算方式, 性质, 余额)
                Values
                  (操作员姓名_In, v_结算方式, 1, -1 * n_实际冲销);
                n_返回值 := n_实际冲销;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 人员缴款余额
                Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
              End If;
            
              --退原预交记录
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实际冲销)
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实际冲销, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, c_预交.合作单位, n_结帐id,
                   -1 * n_结帐id, 0, 3);
              End If;
            End If;
            Update 病人预交记录
            Set 记录状态 = 3
            Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                  结算方式 = c_预交.结算方式;
            n_实收金额 := n_实收金额 - n_实际冲销;
          End If;
        End If;
      End Loop;
    
      --更新费用审核记录
      Update 费用审核记录
      Set 记录状态 = 2
      Where 费用id In (Select ID From 门诊费用记录 Where NO = No_In And 记录性质 = 1 And 记录状态 In (1, 3)) And 性质 = 1;
      --作废门诊记录
      Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And Mod(记录性质,10) = 1 And 记录状态 = 1;
      For r_Clinic In (Select 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码, 费用类型,
                              发药窗口, 付数, Sum(数次) As 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, Sum(应收金额) As 应收金额,
                              Sum(实收金额) As 实收金额, Sum(统筹金额) As 统筹金额, 开单部门id, 开单人, 执行部门id, 划价人, Max(记帐单id) As 记帐单id, 发生时间,
                              实际票号
                       From 门诊费用记录
                       Where NO = No_In And Mod(记录性质,10) = 1 And 记录状态 In (2, 3) And Nvl(附加标志, 0) Not In (8, 9)
                       Group By 序号, 从属父号, 价格父号, 病人id, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否, 保险大类id, 保险编码,
                                费用类型, 发药窗口, 付数, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 开单部门id, 开单人, 执行部门id, 划价人, 发生时间, 实际票号
                       Having Sum(数次) <> 0) Loop
        Insert Into 门诊费用记录
          (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
           保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间,
           登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 缴款组id, 结帐id, 结帐金额, 费用状态)
        Values
          (病人费用记录_Id.Nextval, 1, No_In, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号, 1, r_Clinic.病人id,
           '', r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别, r_Clinic.收费细目id,
           r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口, r_Clinic.付数,
           -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价,
           -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 0, r_Clinic.开单部门id, r_Clinic.开单人, r_Clinic.发生时间,
           退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', n_组id, n_结帐id,
           -1 * r_Clinic.实收金额, 0);
      End Loop;
    Else
      --4.退款转预交(不产生票据,由操作员通过重打进行)
      For r_Pay In (Select Min(a.Id) As 预交id, a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号,
                           a.交易说明, a.合作单位, b.性质
                    From 病人预交记录 A, 结算方式 B
                    Where a.记录性质 = 3 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                          a.结算方式 = b.名称 And (b.性质 In (1, 2, 7, 8)) And a.结算方式 Is Not Null
                    Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质, a.交易流水号, a.交易说明, a.合作单位


                    
                    Having Sum(a.冲预交) <> 0
                    Order By a.卡类别id, 性质 Desc) Loop
        --4.1产生预交款单据 (不存在部分退费的情况)
        --所有单据,按规则生成预交款单据
        --因为收款后立即缴款,所以人员缴款余额无变化
        If n_实收金额 <> 0 Then
          If r_Pay.性质 = 7 Or (r_Pay.性质 = 8 And r_Pay.卡类别id Is Not Null) Then
            If r_Pay.冲预交 > n_实收金额 Then
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '1' || ',' || r_Pay.卡类别id || ',' || -1 * n_实收金额 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || r_Pay.卡类别id || ',' || -1 * n_实收金额 || '|', n_组id,
                   Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
            
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = r_Pay.结算方式;
              n_费用状态 := 1;
              n_实收金额 := 0;
            Else
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|'
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|',
                   n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
              End If;
            
              Update 病人预交记录
              Set 记录状态 = 3
              Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                    结算方式 = r_Pay.结算方式;
              n_费用状态 := 1;
              n_实收金额 := n_实收金额 - r_Pay.冲预交;
            End If;
          Else
            n_实际冲销 := 0;
            If r_Pay.性质 In (3, 4) Or (r_Pay.性质 = 8 And r_Pay.结算卡序号 Is Not Null) Then
              v_结算方式 := r_Pay.结算方式;
            Else
              If 结算方式_In Is Null Then
                Begin
                  Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
                Exception
                  When Others Then
                    Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
                End;
              Else
                v_结算方式 := 结算方式_In;
              End If;
            End If;
          
            If r_Pay.性质 = 8 And r_Pay.结算卡序号 Is Not Null Then
              If n_实收金额 >= r_Pay.冲预交 Then
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, r_Pay.冲预交, r_Pay.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null,
                     退费时间_In, Null, Null, Null, 操作员编号_In, 操作员姓名_In,
                     '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|', n_组id, Null, Null, Null, Null, Null,
                     Null, n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := r_Pay.冲预交;
              Else
                --Zl_Square_Update(v_原结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, n_实收金额, r_Pay.结算卡序号);
                Update 病人预交记录
                Set 冲预交 = 冲预交 + (-1 * n_实收金额), 摘要 = 摘要 || '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * n_实收金额 || '|'
                Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into 病人预交记录
                    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                     摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
                  Values
                    (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实收金额, Null, Null, 退费时间_In,
                     Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * n_实收金额 || '|', n_组id,
                     Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
                End If;
                n_费用状态 := 1;
                n_实际冲销 := n_实收金额;
              End If;
            Else
              If r_Pay.冲预交 > n_实收金额 Then
                n_实际冲销 := n_实收金额;
              Else
                n_实际冲销 := r_Pay.冲预交;
              End If;
            End If;
          
            If r_Pay.性质 Not In (3, 4, 7, 8) Then
              Update 病人预交记录
              Set 金额 = 金额 + n_实际冲销
              Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                v_预交no := Nextno(11);
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 预交类别)
                Values
                  (病人预交记录_Id.Nextval, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, n_实际冲销, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, r_Pay.预交类别);
              End If;
            
              --病人余额
              Update 病人余额
              Set 预交余额 = Nvl(预交余额, 0) + n_实际冲销
              Where 性质 = 1 And 病人id = n_病人id And 类型 = 2
              Returning 预交余额 Into n_返回值;
              If Sql%RowCount = 0 Then
                Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, 2, n_实际冲销, 0);
                n_返回值 := n_实际冲销;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 病人余额
                Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
              End If;
            End If;
            --4.2缴款数据处理
            --   因为没有实际收病人的钱,所以不处理
            --部分退费情况，退原预交记录
            If r_Pay.性质 In (3, 4) Then
              Update 人员缴款余额
              Set 余额 = Nvl(余额, 0) - n_实际冲销
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式
              Returning 余额 Into n_返回值;
              If Sql%RowCount = 0 Then
                Insert Into 人员缴款余额
                  (收款员, 结算方式, 性质, 余额)
                Values
                  (操作员姓名_In, r_Pay.结算方式, 1, -1 * n_实际冲销);
                n_返回值 := n_实际冲销;
              End If;
              If Nvl(n_返回值, 0) = 0 Then
                Delete From 人员缴款余额
                Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式 And Nvl(余额, 0) = 0;
              End If;
            End If;
          
            If r_Pay.性质 <> 8 Then
              Update 病人预交记录
              Set 冲预交 = 冲预交 + (-1 * n_实际冲销)
              Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
              If Sql%RowCount = 0 Then
                Insert Into 病人预交记录
                  (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名,
                   摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
                Values
                  (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * n_实际冲销, v_结算方式, Null, 退费时间_In,
                   Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, r_Pay.卡类别id, r_Pay.结算卡序号, r_Pay.卡号, r_Pay.交易流水号,
                   r_Pay.交易说明, r_Pay.合作单位, n_结帐id, -1 * n_结帐id, 0, 3);
              End If;
            End If;
          
            Update 病人预交记录
            Set 记录状态 = 3
            Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_原结帐ids))) And
                  结算方式 = r_Pay.结算方式;
            n_实收金额 := n_实收金额 - n_实际冲销;
          
          End If;
        End If;
      End Loop;
    End If;
  
    If 误差费_In Is Not Null Then
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id,
         卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
      Values
        (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, 误差费_In, v_误差费, Null, 退费时间_In, Null, Null,
         Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 0, 3);
    End If;
    Delete From 病人预交记录
    Where 结帐id = n_结帐id And 记录性质 = 3 And 记录状态 = 2 And 冲预交 = 0 And 结算方式 Is Not Null;
    Delete From 病人预交记录 Where 结帐id = n_原结帐id And 摘要 = '预交临时记录' And 记录性质 = 3;
    Update 门诊费用记录 Set 费用状态 = Nvl(n_费用状态, 0) Where NO = No_In And Mod(记录性质, 10) = 1 And 记录状态 = 2;
  Else
    --医保按结算转出
    For r_Nos In (Select Distinct a.No
                  From 门诊费用记录 A
                  Where Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And a.结帐id = 原结帐id_In) Loop
      v_Nos := v_Nos || ',' || r_Nos.No;
    End Loop;
    v_Nos := Substr(v_Nos, 2);
  
    For r_结帐ids In (Select Distinct a.结帐id
                    From 门诊费用记录 A
                    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And
                          a.记录状态 <> 0) Loop
      v_结帐ids := v_结帐ids || ',' || r_结帐ids.结帐id;
    End Loop;
    v_结帐ids := Substr(v_结帐ids, 2);
    Select Count(a.No), Sum(a.实收金额)
    Into n_Count, n_实收金额
    From 门诊费用记录 A
    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1;
    If n_Count = 0 Or n_实收金额 = 0 Then
      v_Err_Msg := '本次结算不是收费或因并发原因他人操作了该结算,不能转为住院费用.';
      Raise Err_Item;
    End If;
  
    Select 结帐id, 病人id, 开单部门id, 开单人
    Into n_原结帐id, n_病人id, n_开单部门id, v_开单人
    From 门诊费用记录
    Where 结帐id = 原结帐id_In And Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And Rownum < 2;
  
    Begin
      Select 1
      Into n_部分退费
      From 门诊费用记录 A
      Where Mod(a.记录性质, 10) = 1 And a.记录状态 = 2 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
            Rownum < 2;
    Exception
      When Others Then
        n_部分退费 := 0;
    End;
  
    Begin
      Select 0
      Into n_部分退费
      From 门诊费用记录 A
      Where 记录性质 = 11 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select Count(Avg(1))
      Into n_退费条数
      From 病人预交记录 A
      Where a.记录性质 = 3 And a.记录状态 <> 0 And 结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))
      Group By a.结算方式;
    Exception
      When Others Then
        n_退费条数 := 0;
    End;
    --1.1作废费用记录
    If 结帐id_In Is Null Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    Else
      n_结帐id := 结帐id_In;
    End If;
    Insert Into 门诊费用记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别, 收费细目id,
       计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, 执行状态, 执行时间,
       操作员编号, 操作员姓名, 发生时间, 登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id, 费用状态)
      Select 病人费用记录_Id.Nextval, a.No, a.实际票号, a.记录性质, 2, a.序号, a.从属父号, a.价格父号, a.病人id, a.医嘱序号, a.门诊标志, a.姓名, a.性别, a.年龄,
             a.标识号, a.付款方式, a.费别, a.病人科室id, a.收费类别, a.收费细目id, a.计算单位, a.付数, a.发药窗口, -1 * a.数次, a.加班标志, a.附加标志, a.收入项目id,
             a.收据费目, a.记帐费用, a.标准单价, -1 * a.应收金额, -1 * a.实收金额, a.开单部门id, a.开单人, a.执行部门id, a.划价人, a.执行人, -1, a.执行时间,
             操作员编号_In, 操作员姓名_In, a.发生时间, 退费时间_In, n_结帐id, -1 * a.结帐金额, a.保险项目否, a.保险大类id, a.统筹金额, a.摘要,
             Decode(Nvl(a.附加标志, 0), 9, 1, 0), a.保险编码, a.费用类型, n_组id, 0
      From 门诊费用记录 A
      Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And a.记录状态 = 1;
  
    --作废医保
    For r_医保 In (Select 结帐id, NO, 结算方式, 金额, 备注
                 From 医保结算明细
                 Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And
                       结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))) Loop
      Update 医保结算明细
      Set 金额 = 金额 + (-1 * r_医保.金额)
      Where NO = r_医保.No And 结帐id = r_医保.结帐id And 结算方式 = r_医保.结算方式;
      If Sql%RowCount = 0 Then
        Insert Into 医保结算明细
          (结帐id, NO, 结算方式, 金额)
        Values
          (r_医保.结帐id, r_医保.No, r_医保.结算方式, -1 * r_医保.金额);
      End If;
    End Loop;
  
    --Update 门诊费用记录 Set 记录状态 = 3 Where NO = No_In And 记录性质 = 1 And 记录状态 = 1;
    --1.2作废预交记录
    --作废冲预交部分
    If n_部分退费 = 0 And Nvl(门诊退费_In, 0) = 0 Then
      For r_Prepay In (Select NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, -1 * Sum(冲预交) As 冲预交,
                              卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质
                       From 病人预交记录 A
                       Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                             Nvl(冲预交, 0) <> 0
                       Group By n_Tempid, NO, 实际票号, 病人id, 主页id, 科室id, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, 卡类别id, 结算卡序号,
                                卡号, 交易流水号, 交易说明, 合作单位, 结算性质) Loop
        Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 预交类别, 结算性质)
          Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                 r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                 操作员编号_In, r_Prepay.冲预交, n_结帐id, n_组id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                 r_Prepay.交易说明, r_Prepay.合作单位, -1 * n_结帐id, 1, r_Prepay.结算性质
          From Dual;
      End Loop;
    
      For v_预交 In (Select 病人id, Nvl(预交类别, 2) As 预交类别, Nvl(Sum(Nvl(冲预交, 0)), 0) As 预交金额
                   From 病人预交记录 A
                   Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                         a.结帐id <> n_结帐id
                   Group By 病人id, Nvl(预交类别, 2)
                   Having Sum(Nvl(冲预交, 0)) <> 0) Loop
      
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) + Nvl(v_预交.预交金额, 0)
        Where 病人id = v_预交.病人id And 类型 = Nvl(v_预交.预交类别, 2) And 性质 = 1
        Returning 预交余额 Into n_返回值;
        If Sql%RowCount = 0 Then
          Insert Into 病人余额
            (病人id, 类型, 预交余额, 性质)
          Values
            (v_预交.病人id, Nvl(v_预交.预交类别, 2), v_预交.预交金额, 1);
          n_返回值 := v_预交.预交金额;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 病人余额
          Where 病人id = v_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
        End If;
      End Loop;
    Else
      If n_退费条数 = 0 And Nvl(门诊退费_In, 0) = 0 Then
        --只使用了预交，原样退回预交
        For r_Prepay In (Select NO, 实际票号, 病人id, 主页id, 科室id, Max(结算方式) As 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间,
                                -1 * Sum(冲预交) As 冲预交, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质
                         From 病人预交记录 A
                         Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                               Nvl(冲预交, 0) <> 0
                         Group By n_Tempid, NO, 实际票号, 病人id, 主页id, 科室id, 结算号码, 缴款单位, 单位开户行, 单位帐号, 收款时间, 卡类别id, 结算卡序号, 卡号,
                                  交易流水号, 交易说明, 合作单位, 结算性质) Loop
          Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
          Insert Into 病人预交记录
            (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
             结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 预交类别, 结算性质)
            Select n_Tempid, r_Prepay.No, r_Prepay.实际票号, 11, 1, r_Prepay.病人id, r_Prepay.主页id, r_Prepay.科室id, Null,
                   r_Prepay.结算方式, r_Prepay.结算号码, Null, r_Prepay.缴款单位, r_Prepay.单位开户行, r_Prepay.单位帐号, 退费时间_In, 操作员姓名_In,
                   操作员编号_In, r_Prepay.冲预交, n_结帐id, n_组id, r_Prepay.卡类别id, r_Prepay.结算卡序号, r_Prepay.卡号, r_Prepay.交易流水号,
                   r_Prepay.交易说明, r_Prepay.合作单位, -1 * n_结帐id, 1, r_Prepay.结算性质
            From Dual;
          Select -1 * 冲预交 Into n_预交金额 From 病人预交记录 Where ID = n_Tempid;
          Update 病人余额
          Set 预交余额 = Nvl(预交余额, 0) + Nvl(n_预交金额, 0)
          Where 病人id = r_Prepay.病人id And 类型 = 1 And 性质 = 1
          Returning 预交余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (n_病人id, 1, n_预交金额, 1);
            n_返回值 := n_预交金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 病人余额
            Where 病人id = r_Prepay.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
          End If;
        End Loop;
      Else
        Begin
          Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
        Exception
          When Others Then
            Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
        End;
        Select 病人预交记录_Id.Nextval Into n_Tempid From Dual;
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
           结帐id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
          Select n_Tempid, Max(NO), Max(实际票号), 3, 3, 病人id, 主页id, 科室id, Null, v_结算方式, Max(结算号码), '预交临时记录', Null, Null,
                 Null, Max(收款时间), 操作员姓名_In, 操作员编号_In, Sum(冲预交), n_原结帐id, Null, Null, Null, Null, Null, Null,
                 -1 * n_原结帐id, 3
          From 病人预交记录 A
          Where 记录性质 In (1, 11) And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
                Nvl(冲预交, 0) <> 0
          Group By n_Tempid, 3, 3, 病人id, 主页id, 科室id, Null, v_结算方式, '预交临时记录', 操作员姓名_In, 操作员编号_In, n_原结帐id;
      End If;
    End If;
  
    --作废门诊缴费及医保部分
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退费时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
             0, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, -1 * n_结帐id, 结算性质
      From 病人预交记录 A, 结算方式 B
      Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
            a.结算方式 = b.名称 And b.性质 Not In (7, 8);
  
    Insert Into 病人预交记录
      (ID, 记录性质, NO, 记录状态, 病人id, 主页id, 摘要, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 预交类别,
       卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 结算性质, 校对标志)
      Select 病人预交记录_Id.Nextval, 记录性质, NO, 2, 病人id, 主页id, 摘要, 结算方式, 结算号码, 退费时间_In, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In, 操作员姓名_In,
             0, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, -1 * n_结帐id, 结算性质, 1
      From 病人预交记录 A, 结算方式 B
      Where a.记录性质 = 3 And a.记录状态 = 1 And a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And
            a.结算方式 = b.名称 And b.性质 = 7;
    If Sql%RowCount <> 0 Then
      n_费用状态 := 1;
    End If;
  
    Update 病人预交记录
    Set 记录状态 = 3
    Where 记录性质 = 3 And 记录状态 = 1 And 结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)));
  
    --2.票据收回
    --可能以前没有打印,无收回
    For r_Nos In (Select Distinct a.No
                  From 门诊费用记录 A
                  Where Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And
                        a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))) Loop
    
      Select Nvl(Max(ID), 0)
      Into n_打印id
      From (Select b.Id
             From 票据使用明细 A, 票据打印内容 B
             Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = r_Nos.No
             Order By a.使用时间 Desc)
      Where Rownum < 2;
      If n_打印id > 0 Then
        --多张单据循环调用时只能收回一次
        Select Count(打印id) Into n_Count From 票据使用明细 Where 票种 = 1 And 性质 = 2 And 打印id = n_打印id;
        If n_Count = 0 Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, 退费时间_In, 操作员姓名_In
            From 票据使用明细
            Where 打印id = n_打印id And 票种 = 1 And 性质 = 1;
        End If;
      End If;
    End Loop;
  
    --3.缴款数据处理(
    --   现有两种情况:
    --    1. 转出过程直接销帐的,则缴款数据不增加;
    --    2. 先转出,再到门诊退款退票,则需要进行缴款数据处理
    If Nvl(门诊退费_In, 0) = 1 Then
      For c_预交 In (Select a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, Min(a.交易流水号) As 交易流水号,
                          Min(a.交易说明) As 交易说明, Min(a.合作单位) As 合作单位, b.性质
                   From 病人预交记录 A, 结算方式 B
                   Where a.记录性质 = 3 And a.记录状态 In (2, 3) And
                         a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And a.结算方式 = b.名称 And
                         b.性质 In (1, 2, 3, 4, 7, 8) And a.结算方式 Is Not Null
                   Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质
                   Having Sum(a.冲预交) <> 0) Loop
        Begin
          Select 是否退现 Into n_退现 From 医疗卡类别 Where ID = c_预交.卡类别id;
        Exception
          When Others Then
            n_退现 := 0;
        End;
        If (c_预交.性质 = 7 Or (c_预交.性质 = 8 And c_预交.卡类别id Is Not Null)) And n_退现 = 0 Then
          Update 病人预交记录
          Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|'
          Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
               缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
            Values
              (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null, 退费时间_In,
               Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || c_预交.卡类别id || ',' || -1 * c_预交.冲预交 || '|', n_组id,
               Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
          End If;
          n_费用状态 := 1;
        Else
          If c_预交.性质 In (3, 4) Or (c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null) Then
            v_结算方式 := c_预交.结算方式;
          Else
            If 结算方式_In Is Null Then
              Begin
                Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
              Exception
                When Others Then
                  Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
              End;
            Else
              v_结算方式 := 结算方式_In;
            End If;
          End If;
        
          If c_预交.性质 = 8 And c_预交.结算卡序号 Is Not Null Then
            --Zl_Square_Update(v_结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, c_预交.冲预交, c_预交.结算卡序号);
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交), 摘要 = 摘要 || '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|'
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, Null, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || c_预交.结算卡序号 || ',' || -1 * c_预交.冲预交 || '|', n_组id,
                 Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
            End If;
            n_费用状态 := 1;
          End If;
          If c_预交.结算卡序号 Is Null Then
            Update 人员缴款余额
            Set 余额 = Nvl(余额, 0) - c_预交.冲预交
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式
            Returning 余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 人员缴款余额
                (收款员, 结算方式, 性质, 余额)
              Values
                (操作员姓名_In, v_结算方式, 1, -1 * c_预交.冲预交);
              n_返回值 := c_预交.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 人员缴款余额
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_结算方式 And Nvl(余额, 0) = 0;
            End If;
            --部分退费情况，退原预交记录
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * c_预交.冲预交)
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * c_预交.冲预交, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, c_预交.合作单位, n_结帐id,
                 -1 * n_结帐id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    
      --更新费用审核记录
      Update 费用审核记录
      Set 记录状态 = 2
      Where 费用id In (Select a.Id
                     From 门诊费用记录 A
                     Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And
                           a.记录状态 In (1, 3)) And 性质 = 1;
      --作废门诊记录
      For r_Nos In (Select Distinct NO
                    From 门诊费用记录
                    Where Mod(记录性质, 10) = 1 And 记录状态 In (1, 3) And
                          结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids)))) Loop
        Update 门诊费用记录 Set 记录状态 = 3 Where NO = r_Nos.No And Mod(记录性质, 10) = 1 And 记录状态 = 1;
      End Loop;
      For r_Clinic In (Select Min(a.记录性质) As 记录性质, a.No, a.序号, a.从属父号, a.价格父号, a.病人id, a.姓名, a.性别, a.年龄, a.病人科室id, a.费别,
                              a.收费类别, a.收费细目id, a.计算单位, a.保险项目否, a.保险大类id, a.保险编码, a.费用类型, a.发药窗口, a.付数, Sum(a.数次) As 数次,
                              a.加班标志, a.附加标志, a.收入项目id, a.收据费目, a.标准单价, Sum(a.应收金额) As 应收金额, Sum(a.实收金额) As 实收金额,
                              Sum(a.统筹金额) As 统筹金额, a.开单部门id, a.开单人, a.执行部门id, a.划价人, Max(a.记帐单id) As 记帐单id,
                              Max(a.是否急诊) As 是否急诊, a.发生时间, Min(a.实际票号) As 实际票号
                       From 门诊费用记录 A
                       Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.记录性质, 10) = 1 And
                             a.记录状态 In (2, 3) And Nvl(a.附加标志, 0) Not In (8, 9)
                       Group By a.No, a.序号, a.从属父号, a.价格父号, a.病人id, a.姓名, a.性别, a.年龄, a.病人科室id, a.费别, a.收费类别, a.收费细目id,
                                a.计算单位, a.保险项目否, a.保险大类id, a.保险编码, a.费用类型, a.发药窗口, a.付数, a.加班标志, a.附加标志, a.收入项目id, a.收据费目,
                                a.标准单价, a.开单部门id, a.开单人, a.执行部门id, a.划价人, a.发生时间
                       Having Sum(a.数次) <> 0) Loop
        Insert Into 门诊费用记录
          (ID, 记录性质, NO, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 门诊标志, 病人id, 标识号, 姓名, 性别, 年龄, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 保险项目否,
           保险大类id, 保险编码, 费用类型, 发药窗口, 付数, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 统筹金额, 记帐费用, 开单部门id, 开单人, 发生时间,
           登记时间, 执行部门id, 划价人, 操作员编号, 操作员姓名, 记帐单id, 摘要, 是否急诊, 缴款组id, 结帐id, 结帐金额, 执行状态, 费用状态)
        Values
          (病人费用记录_Id.Nextval, r_Clinic.记录性质, r_Clinic.No, r_Clinic.实际票号, 2, r_Clinic.序号, r_Clinic.从属父号, r_Clinic.价格父号,
           1, r_Clinic.病人id, '', r_Clinic.姓名, r_Clinic.性别, r_Clinic.年龄, r_Clinic.病人科室id, r_Clinic.费别, r_Clinic.收费类别,
           r_Clinic.收费细目id, r_Clinic.计算单位, r_Clinic.保险项目否, r_Clinic.保险大类id, r_Clinic.保险编码, r_Clinic.费用类型, r_Clinic.发药窗口,
           r_Clinic.付数, -1 * r_Clinic.数次, r_Clinic.加班标志, r_Clinic.附加标志, r_Clinic.收入项目id, r_Clinic.收据费目, r_Clinic.标准单价,
           -1 * r_Clinic.应收金额, -1 * r_Clinic.实收金额, -1 * r_Clinic.统筹金额, 0, r_Clinic.开单部门id, r_Clinic.开单人, r_Clinic.发生时间,
           退费时间_In, r_Clinic.执行部门id, r_Clinic.划价人, 操作员编号_In, 操作员姓名_In, r_Clinic.记帐单id, '', r_Clinic.是否急诊, n_组id, n_结帐id,
           -1 * r_Clinic.实收金额, -1, 0);
      End Loop;
    Else
      --4.退款转预交(不产生票据,由操作员通过重打进行)
    
      For r_Pay In (Select Min(a.Id) As 预交id, a.结算方式, Sum(a.冲预交) As 冲预交, 2 As 预交类别, a.卡类别id, a.结算卡序号, a.卡号, a.交易流水号,
                           a.交易说明, a.合作单位, b.性质
                    From 病人预交记录 A, 结算方式 B
                    Where a.记录性质 = 3 And a.记录状态 In (2, 3) And
                          a.结帐id In (Select Column_Value From Table(f_Str2list(v_结帐ids))) And a.结算方式 = b.名称 And
                          b.性质 In (1, 2, 3, 4, 7, 8) And a.结算方式 Is Not Null
                    Group By a.结算方式, 预交类别, a.卡类别id, a.结算卡序号, a.卡号, b.性质, a.交易流水号, a.交易说明, a.合作单位


                    
                    Having Sum(a.冲预交) <> 0) Loop
        --4.1产生预交款单据 (不存在部分退费的情况)
        --所有单据,按规则生成预交款单据
        --因为收款后立即缴款,所以人员缴款余额无变化
        If r_Pay.性质 = 7 Or (r_Pay.性质 = 8 And r_Pay.卡类别id Is Not Null) Then
          Update 病人预交记录
          Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|'
          Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
          If Sql%RowCount = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
               缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
            Values
              (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null, 退费时间_In,
               Null, Null, Null, 操作员编号_In, 操作员姓名_In, '1' || ',' || r_Pay.卡类别id || ',' || -1 * r_Pay.冲预交 || '|', n_组id,
               Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
          End If;
          n_费用状态 := 1;
        Else
          If r_Pay.性质 In (3, 4) Or (r_Pay.性质 = 8 And r_Pay.结算卡序号 Is Not Null) Then
            v_结算方式 := r_Pay.结算方式;
          Else
            Begin
              Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
            Exception
              When Others Then
                Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
            End;
          End If;
        
          If r_Pay.性质 = 8 Then
            --Zl_Square_Update(v_结帐ids, n_结帐id, n_组id, 退费时间_In, -1 * n_结帐id, Null, r_Pay.冲预交, r_Pay.结算卡序号);
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交), 摘要 = 摘要 || '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|'
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 Is Null;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 结算性质, 校对标志)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, Null, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '0' || ',' || r_Pay.结算卡序号 || ',' || -1 * r_Pay.冲预交 || '|', n_组id,
                 Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 3, 1);
            End If;
            n_费用状态 := 1;
          End If;
          If r_Pay.性质 Not In (3, 4, 7, 8) Then
            Update 病人预交记录
            Set 金额 = 金额 + r_Pay.冲预交
            Where 记录性质 = 1 And 记录状态 = 1 And 收款时间 = 退费时间_In And 病人id + 0 = n_病人id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              v_预交no := Nextno(11);
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 预交类别)
              Values
                (病人预交记录_Id.Nextval, v_预交no, Null, 1, 1, n_病人id, 主页id_In, 入院科室id_In, r_Pay.冲预交, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '门诊转住院预交', n_组id, r_Pay.预交类别);
            End If;
          
            --病人余额
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + r_Pay.冲预交
            Where 性质 = 1 And 病人id = n_病人id And 类型 = 2
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额 (病人id, 性质, 类型, 预交余额, 费用余额) Values (n_病人id, 1, 2, r_Pay.冲预交, 0);
              n_返回值 := r_Pay.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = n_病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End If;
          --4.2缴款数据处理
          --   因为没有实际收病人的钱,所以不处理
          --部分退费情况，退原预交记录
          If r_Pay.性质 In (3, 4) Then
            Update 人员缴款余额
            Set 余额 = Nvl(余额, 0) - r_Pay.冲预交
            Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式
            Returning 余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 人员缴款余额
                (收款员, 结算方式, 性质, 余额)
              Values
                (操作员姓名_In, r_Pay.结算方式, 1, -1 * r_Pay.冲预交);
              n_返回值 := r_Pay.冲预交;
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 人员缴款余额
              Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = r_Pay.结算方式 And Nvl(余额, 0) = 0;
            End If;
          End If;
        
          If r_Pay.结算卡序号 Is Null Then
            Update 病人预交记录
            Set 冲预交 = 冲预交 + (-1 * r_Pay.冲预交)
            Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
            If Sql%RowCount = 0 Then
              Insert Into 病人预交记录
                (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
                 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
              Values
                (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, -1 * r_Pay.冲预交, v_结算方式, Null, 退费时间_In,
                 Null, Null, Null, 操作员编号_In, 操作员姓名_In, '', n_组id, r_Pay.卡类别id, r_Pay.结算卡序号, r_Pay.卡号, r_Pay.交易流水号,
                 r_Pay.交易说明, r_Pay.合作单位, n_结帐id, -1 * n_结帐id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    End If;
    If 误差费_In Is Not Null Then
      Begin
        Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And 名称 Like '%现金%' And Rownum < 2;
      Exception
        When Others Then
          Select 名称 Into v_结算方式 From 结算方式 Where 性质 = 1 And Rownum < 2;
      End;
      Update 病人预交记录
      Set 冲预交 = 冲预交 - 误差费_In
      Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_结算方式;
      Update 病人预交记录
      Set 冲预交 = 冲预交 + 误差费_In
      Where 记录性质 = 3 And 记录状态 = 2 And 结帐id = n_结帐id And 结算方式 = v_误差费;
      If Sql%RowCount = 0 Then
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 冲预交, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要,
           缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 校对标志, 结算性质)
        Values
          (病人预交记录_Id.Nextval, Null, Null, 3, 2, n_病人id, 主页id_In, 入院科室id_In, 误差费_In, v_误差费, Null, 退费时间_In, Null, Null,
           Null, 操作员编号_In, 操作员姓名_In, '', n_组id, Null, Null, Null, Null, Null, Null, n_结帐id, -1 * n_结帐id, 0, 3);
      End If;
    End If;
    Delete From 病人预交记录 Where 结帐id = n_原结帐id And 摘要 = '预交临时记录' And 记录性质 = 3;
    Delete From 病人预交记录
    Where 结帐id = n_结帐id And 记录性质 = 3 And 记录状态 = 2 And 冲预交 = 0 And 结算方式 Is Not Null;
    Update 门诊费用记录
    Set 费用状态 = Nvl(n_费用状态, 0)
    Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(记录性质, 10) = 1 And 记录状态 = 2;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊转住院_收费转出;
/

--105165:张德婷,2017-04-20,处方发药可以添加该叫号窗口的所有处方
CREATE OR REPLACE Procedure zl_发药窗口_Insert (
    编码_IN IN 发药窗口.编码%Type, 
    名称_IN IN 发药窗口.名称%Type, 
    上班_IN IN 发药窗口.上班否%Type, 
    药房ID_IN IN 发药窗口.药房ID%Type, 
    专家_IN IN 发药窗口.专家%Type,
    叫号窗口_IN In 发药窗口.叫号窗口%type

) 
IS 
    Msg VARCHAR2 (30); 
Begin 
    Insert INTO 发药窗口 
                    (编码, 名称, 上班否, 药房ID, 专家,叫号窗口) 
          VALUES (编码_IN, 名称_IN, 上班_IN, 药房ID_IN, 专家_IN,叫号窗口_IN); 
Exception 
    When Others Then 
        Zl_ErrorCenter (SQLCODE, SQLERRM); 
End zl_发药窗口_Insert;
/

--105165:张德婷,2017-04-20,处方发药可以添加该叫号窗口的所有处方
CREATE OR REPLACE Procedure zl_发药窗口_UPDATE (
    编码_IN IN 发药窗口.编码%Type,
    名称_IN IN 发药窗口.名称%Type,
    药房ID_IN IN 发药窗口.药房ID%Type,
    专家_IN IN 发药窗口.专家%Type,
    Old编码_IN IN 发药窗口.编码%Type,
    Old药房ID_IN IN 发药窗口.药房ID%Type,
    叫号窗口_IN In 发药窗口.叫号窗口%type
)
IS
    Msg VARCHAR2 (30);
Begin
    UPDATE 发药窗口
        SET 编码 = 编码_IN,
             名称 = 名称_IN,
             药房ID = 药房ID_IN,
             专家 = 专家_IN,
             叫号窗口=叫号窗口_IN
     WHERE 编码 = Old编码_IN
        AND 药房ID = Old药房ID_IN;
Exception
    When Others Then
        Zl_ErrorCenter (SQLCODE, SQLERRM);
End zl_发药窗口_UPDATE;
/

--105165:张德婷,2017-04-20,处方发药可以添加该叫号窗口的所有处方
CREATE OR REPLACE Procedure Zl_未发药品记录_呼叫
(
  No_In       药品收发记录.NO%Type,
  单据_In     药品收发记录.单据%Type,
  药房id_In   药品收发记录.库房id%Type,
  发药窗口_In 药品收发记录.发药窗口%Type,
  呼叫内容_In 未发药品记录.呼叫内容%Type := Null
) Is
Begin
  If 呼叫内容_In Is Null Then
    --呼叫内容为空时，将当前的呼叫状态的单据的呼叫内容清空
    Update 未发药品记录
    Set 呼叫内容 = Null
    Where 库房id = 药房id_In And 单据 = 单据_In and (发药窗口 = 发药窗口_In or 发药窗口 in(select 名称 from 发药窗口 where 叫号窗口=发药窗口_In)) And NO = No_In  And 排队状态 = 3 and 填制日期 between sysdate-3 and sysdate;
  Else
    --呼叫内容不为空时，先将以前的呼叫状态中的单据设置为已呼叫，再将当前单据设置为呼叫状态，并填写呼叫内容和呼叫时间
    --可以满足同一单据反复呼叫的情况
    Update 未发药品记录
    Set 排队状态 = 4, 呼叫内容 = Null
    Where 库房id = 药房id_In And (发药窗口 = 发药窗口_In or 发药窗口 in(select 名称 from 发药窗口 where 叫号窗口=发药窗口_In)) And 排队状态 = 3 and 填制日期 between sysdate-3 and sysdate;

    Update 未发药品记录
    Set 排队状态 = 3, 呼叫内容 = 呼叫内容_In, 呼叫时间 = Sysdate
    Where 库房id = 药房id_In And 单据 = 单据_In And NO = No_In and 填制日期 between sysdate-3 and sysdate;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_未发药品记录_呼叫;
/

--108439:余伟节,2017-04-19,住院病人下达医嘱时取病人年龄时从病案主页中提取
CREATE OR REPLACE Procedure Zl_病人医嘱记录_Insert
(
  Id_In           病人医嘱记录.Id%Type,
  相关id_In       病人医嘱记录.相关id%Type,
  序号_In         病人医嘱记录.序号%Type,
  病人来源_In     病人医嘱记录.病人来源%Type,
  病人id_In       病人医嘱记录.病人id%Type,
  主页id_In       病人医嘱记录.主页id%Type,
  婴儿_In         病人医嘱记录.婴儿%Type,
  医嘱状态_In     病人医嘱记录.医嘱状态%Type,
  医嘱期效_In     病人医嘱记录.医嘱期效%Type,
  诊疗类别_In     病人医嘱记录.诊疗类别%Type,
  诊疗项目id_In   病人医嘱记录.诊疗项目id%Type,
  收费细目id_In   病人医嘱记录.收费细目id%Type,
  天数_In         病人医嘱记录.天数%Type,
  单次用量_In     病人医嘱记录.单次用量%Type,
  总给予量_In     病人医嘱记录.总给予量%Type,
  医嘱内容_In     病人医嘱记录.医嘱内容%Type,
  医生嘱托_In     病人医嘱记录.医生嘱托%Type,
  标本部位_In     病人医嘱记录.标本部位%Type,
  执行频次_In     病人医嘱记录.执行频次%Type,
  频率次数_In     病人医嘱记录.频率次数%Type,
  频率间隔_In     病人医嘱记录.频率间隔%Type,
  间隔单位_In     病人医嘱记录.间隔单位%Type,
  执行时间方案_In 病人医嘱记录.执行时间方案%Type,
  计价特性_In     病人医嘱记录.计价特性%Type,
  执行科室id_In   病人医嘱记录.执行科室id%Type,
  执行性质_In     病人医嘱记录.执行性质%Type,
  紧急标志_In     病人医嘱记录.紧急标志%Type,
  开始执行时间_In 病人医嘱记录.开始执行时间%Type,
  执行终止时间_In 病人医嘱记录.执行终止时间%Type,
  病人科室id_In   病人医嘱记录.病人科室id%Type,
  开嘱科室id_In   病人医嘱记录.开嘱科室id%Type,
  开嘱医生_In     病人医嘱记录.开嘱医生%Type,
  开嘱时间_In     病人医嘱记录.开嘱时间%Type,
  挂号单_In       病人医嘱记录.挂号单%Type := Null,
  前提id_In       病人医嘱记录.前提id%Type := Null,
  检查方法_In     病人医嘱记录.检查方法%Type := Null,
  执行标记_In     病人医嘱记录.执行标记%Type := Null,
  可否分零_In     病人医嘱记录.可否分零%Type := Null,
  摘要_In         病人医嘱记录.摘要%Type := Null,
  操作员姓名_In   病人医嘱状态.操作人员%Type := Null,
  零费记帐_In     病人医嘱记录.零费记帐%Type := Null,
  用药目的_In     病人医嘱记录.用药目的%Type := Null,
  用药理由_In     病人医嘱记录.用药理由%Type := Null,
  审核状态_In     病人医嘱记录.审核状态%Type := Null,
  申请序号_In     病人医嘱记录.申请序号%Type := Null,
  超量说明_In     病人医嘱记录.超量说明%Type := Null,
  首次用量_In     病人医嘱记录.首次用量%Type := Null,
  配方id_In       病人医嘱记录.配方id%Type := Null,
  手术情况_In     病人医嘱记录.手术情况%Type := Null,
  组合项目id_In   病人医嘱记录.组合项目id%Type := Null,
  皮试结果_In     病人医嘱记录.皮试结果%Type := Null
  --功能：医生或护士新开,补录医嘱时新产生的医嘱记录。可用于门诊或住院。
) Is
  v_Temp     Varchar2(255);
  v_人员姓名 病人医嘱状态.操作人员%Type;

  v_姓名 病人信息.姓名%Type;
  v_性别 病人信息.性别%Type;
  v_年龄 病人信息.年龄%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --当前操作人员
  If 操作员姓名_In Is Not Null Then
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  If Nvl(主页id_In, 0) <> 0 Then
    Select 姓名, 性别, 年龄 Into v_姓名, v_性别, v_年龄 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Select 姓名, 性别, 年龄 Into v_姓名, v_性别, v_年龄 From 病人信息 Where 病人id = 病人id_In;
  End If;

  --病人医嘱记录
  Insert Into 病人医嘱记录
    (ID, 相关id, 序号, 病人来源, 病人id, 主页id, 姓名, 性别, 年龄, 婴儿, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id, 收费细目id, 天数, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 标本部位,
     检查方法, 执行标记, 执行频次, 频率次数, 频率间隔, 间隔单位, 执行时间方案, 计价特性, 执行科室id, 执行性质, 紧急标志, 可否分零, 开始执行时间, 执行终止时间, 病人科室id, 开嘱科室id, 开嘱医生,
     开嘱时间, 挂号单, 前提id, 摘要, 零费记帐, 手术时间, 用药目的, 用药理由, 审核状态, 申请序号, 超量说明, 首次用量, 配方id, 手术情况, 组合项目id, 皮试结果)
  Values
    (Id_In, 相关id_In, 序号_In, 病人来源_In, 病人id_In, 主页id_In, v_姓名, v_性别, v_年龄, 婴儿_In, 医嘱状态_In, 医嘱期效_In, 诊疗类别_In, 诊疗项目id_In,
     收费细目id_In, 天数_In, 单次用量_In, 总给予量_In, 医嘱内容_In, 医生嘱托_In, 标本部位_In, 检查方法_In, 执行标记_In, 执行频次_In, 频率次数_In, 频率间隔_In, 间隔单位_In,
     执行时间方案_In, 计价特性_In, 执行科室id_In, 执行性质_In, 紧急标志_In, 可否分零_In, 开始执行时间_In, 执行终止时间_In, 病人科室id_In, 开嘱科室id_In, 开嘱医生_In,
     开嘱时间_In, 挂号单_In, 前提id_In, 摘要_In, 零费记帐_In,
     Decode(诊疗类别_In, 'F', To_Date(标本部位_In, 'yyyy-mm-dd hh24:mi:ss'), 'K', To_Date(标本部位_In, 'yyyy-mm-dd hh24:mi:ss'),
             Null), 用药目的_In, 用药理由_In, 审核状态_In, 申请序号_In, 超量说明_In, 首次用量_In, 配方id_In, 手术情况_In, 组合项目id_In, 皮试结果_In);

  --病人医嘱状态
  If 医嘱状态_In <> -1 Then
    Delete From 病人医嘱状态 Where 医嘱id = Id_In And 操作类型 = 1;
    If Sql%RowCount <> 0 Then
      v_Error := '相同ID的新开医嘱已经存在。';
      Raise Err_Custom;
    End If;
    --因为可能同时：新开->自动校对(住院医生发送)->互斥自动停止(住院医生发送临嘱停止),因此分别-2,-1秒
    Insert Into 病人医嘱状态
      (医嘱id, 操作类型, 操作人员, 操作时间)
    Values
      (Id_In, 1, v_人员姓名, Sysdate - 2 / 60 / 60 / 24);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_Insert;
/

--108340:蔡青松,2017-04-18,执行相关操作之前先判断是否收费
Create Or Replace Procedure Zl_检验报告单_Insert
(
  Id_In   In 病人医嘱记录.Id%Type,
  Type_In In Number -- 0=新增 1=删除
) Is
  --HIS和其他LIS接口使用
  v_主页id     病人医嘱记录.主页id%Type;
  v_医嘱id     病人医嘱记录.Id%Type;
  v_开嘱科室id 病人医嘱记录.开嘱科室id%Type;
  v_病人来源   检验标本记录.病人来源%Type;
  v_病人id     检验标本记录.病人id%Type;
  v_婴儿       检验标本记录.婴儿%Type;
  v_病历文件id 病历单据应用.病历文件id%Type;
  v_病历文件名 病历文件列表.名称%Type;
  v_文件id     电子病历内容.文件id%Type;
  v_Temp       Varchar2(255);
  v_人员部门id 部门人员.部门id%Type;
  v_人员编号   人员表.编号%Type;
  v_人员姓名   人员表.姓名%Type;
  v_No         病人医嘱发送.No%Type;
  v_性质       病人医嘱发送.记录性质%Type;
  v_序号       Varchar2(1000);
  v_查阅       Number;
  v_Error      Varchar2(255);
  Err_Custom Exception;
  n_Par Number;
  --查找当前标本的相关申请
  Cursor c_Samplequest Is
    Select Distinct ID As 医嘱id From 病人医嘱记录 Where Id_In In (ID, 相关id);

  --未审核的费用行(不包含药品)
  Cursor c_Verify(v_医嘱id In Number) Is
    Select Distinct 记录性质, NO, 序号, 记录状态, 门诊标志
    From 住院费用记录
    Where 收费类别 Not In ('5', '6', '7') And 医嘱序号 + 0 In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id)) And 记帐费用 = 1 And
          记录状态 = 0 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id = v_医嘱id
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id)))
    Order By 记录性质, NO, 序号;

Begin
  --操作员信息:部门ID,部门名称;人员ID,人员编号,人员姓名
  v_Temp       := Zl_Identity;
  v_人员部门id := To_Number(Substr(v_Temp, 1, Instr(v_Temp, ',') - 1));

  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Distinct Nvl(b.主页id, 0), Nvl(b.相关id, 0), Decode(b.病人来源, 2, 2, 4, 4, 1), Nvl(b.病人id, 0), Nvl(b.开嘱科室id, 0),
                  Nvl(b.婴儿, 0)
  Into v_主页id, v_医嘱id, v_病人来源, v_病人id, v_开嘱科室id, v_婴儿
  From 病人医嘱记录 B
  Where b.相关id = Id_In;
  If v_病人来源 = 1 Then
    --主页ID： 门诊病人填挂号ID
    Select Nvl(Max(b.Id), 0)
    Into v_主页id
    From 病人挂号记录 B, 病人医嘱记录 A
    Where a.挂号单 = b.No(+) And a.Id = Id_In;
  End If;
  Begin
    Select 病历文件id, c.名称
    Into v_病历文件id, v_病历文件名
    From 病人医嘱记录 A, 病历单据应用 B, 病历文件列表 C
    Where a.诊疗项目id = b.诊疗项目id And b.病历文件id = c.Id And a.相关id = v_医嘱id And b.应用场合 = v_病人来源 And Rownum <= 1;
  Exception
    When Others Then
      Return;
  End;

  If Type_In = 0 Then
    --检查是否收费
    n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
    If n_Par = 1 Then
      For r_Samplequest In c_Samplequest Loop
        For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
          If r_Verify.记录状态 = 0 Then
            If r_Verify.门诊标志 = 1 Then
              v_Error := '标本未收费，不允许执行，请联系管理员！';
              Raise Err_Custom;
            Elsif r_Verify.门诊标志 = 2 Then
              v_Error := '标本未记账，不允许执行，请联系管理员！';
              Raise Err_Custom;
            End If;
          End If;
        End Loop;
      End Loop;
    End If;
  
    --新增
    --删除以前的报告记录
    Begin
      Select Nvl(病历id, 0) Into v_文件id From 病人医嘱报告 Where 医嘱id = v_医嘱id And Rownum <= 1;
      If v_文件id > 0 Then
        Delete 电子病历记录 Where ID = v_文件id;
        Delete 电子病历内容 Where 文件id = v_文件id;
      End If;
    Exception
      When Others Then
        Select 电子病历记录_Id.Nextval Into v_文件id From Dual;
        --Insert Into 病人医嘱报告 (医嘱id, 病历id) Values (v_医嘱id, v_文件id);
    End;
  
    Insert Into 电子病历记录
      (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 保存人, 保存时间, 最后版本, 签名级别)
    Values
      (v_文件id, v_病人来源, v_病人id, v_主页id, v_婴儿, v_开嘱科室id, 7, v_病历文件id, v_病历文件名, Null, Sysdate, Null, Sysdate, 1, 0);
  
    Insert Into 病人医嘱报告 (医嘱id, 病历id) Values (v_医嘱id, v_文件id);
  
    Insert Into 电子病历内容
      (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行)
    Values
      (电子病历内容_Id.Nextval, v_文件id, 1, 1, Null, 1, 2, Null, Null, 0, 0, 0, 0);
  
    Update 病人医嘱发送 Set 执行状态 = 1 Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id));
  
    --2.检查当前标本相关的申请的相关标本是否完成审核
    For r_Samplequest In c_Samplequest Loop
    
      --r_SampleQuest.医嘱id申请已经完成,处理后续环节
    
      --2.费用执行处理
      If v_性质 = 1 Then
        Update 门诊费用记录
        Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = v_人员姓名
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
      Else
        Update 住院费用记录
        Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = v_人员姓名
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
      End If;
      --3.自动审核记帐
      For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
        If r_Verify.No || ',' || r_Verify.记录性质 <> v_No || ',' || v_性质 Then
          If v_序号 Is Not Null Then
            If v_性质 = 1 Then
              Zl_门诊记帐记录_Verify(v_No, v_人员编号, v_人员姓名, Substr(v_序号, 2));
            Elsif v_性质 = 2 Then
              Zl_住院记帐记录_Verify(v_No, v_人员编号, v_人员姓名, Substr(v_序号, 2));
            End If;
          End If;
          v_序号 := Null;
        End If;
        v_No   := r_Verify.No;
        v_性质 := r_Verify.记录性质;
        v_序号 := v_序号 || ',' || r_Verify.序号;
      End Loop;
      If v_序号 Is Not Null Then
        If v_性质 = 1 Then
          Zl_门诊记帐记录_Verify(v_No, v_人员编号, v_人员姓名, Substr(v_序号, 2));
        Elsif v_性质 = 2 Then
          Zl_住院记帐记录_Verify(v_No, v_人员编号, v_人员姓名, Substr(v_序号, 2));
        End If;
      End If;
    
    End Loop;
  Else
    --删除
  
    v_查阅 := 0;
    Select Nvl(查阅状态, 0) Into v_查阅 From 病人医嘱报告 Where 医嘱id = v_医嘱id;
    If v_查阅 = 0 Then
      Select 病历id Into v_文件id From 病人医嘱报告 Where 医嘱id = v_医嘱id And Rownum <= 1;
      Delete 病人医嘱报告 Where 医嘱id = v_医嘱id;
      Delete 电子病历记录 Where ID = v_文件id;
      Delete 电子病历内容 Where 文件id = v_文件id;
      Update 病人医嘱发送
      Set 执行状态 = 0
      Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id));
      For r_Samplequest In c_Samplequest Loop
        --2.费用执行处理
        If v_性质 = 1 Then
          Update 门诊费用记录
          Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
          Where 收费类别 Not In ('5', '6', '7') And
                (医嘱序号, 记录性质, NO) In
                (Select 医嘱id, 记录性质, NO
                 From 病人医嘱附费
                 Where 医嘱id = r_Samplequest.医嘱id
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
        Else
          Update 住院费用记录
          Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
          Where 收费类别 Not In ('5', '6', '7') And
                (医嘱序号, 记录性质, NO) In
                (Select 医嘱id, 记录性质, NO
                 From 病人医嘱附费
                 Where 医嘱id = r_Samplequest.医嘱id
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
        End If;
      End Loop;
    Else
      v_Error := '该报告已经被医生查阅，不能取消，请联系医生。';
      Raise Err_Custom;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验报告单_Insert;
/

--108340:蔡青松,2017-04-18,执行相关操作之前先判断是否收费
Create Or Replace Procedure Zl_病人医嘱发送_Sampleinput
(
  医嘱id      In Varchar2,
  接收人_In   In 病人医嘱发送.接收人%Type := Null,
  接收批次_In In 病人医嘱发送.接收批次%Type := 0,
  人员编号_In In 人员表.编号%Type := Null,
  人员姓名_In In 人员表.姓名%Type := Null,
  送检人_In   In 病人医嘱发送.送检人%Type := Null
) Is
  --未审核的费用行(不包含药品)
  Cursor c_Verify(v_医嘱id In Number) Is
    Select Distinct 记录性质, NO, 序号, 记录状态,门诊标志
    From 住院费用记录
    Where 收费类别 Not In ('5', '6', '7') And 医嘱序号 + 0 = v_医嘱id And 记帐费用 = 1 And 记录状态 = 0 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id = v_医嘱id
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id)))
    Union All
    Select Distinct 记录性质, NO, 序号, 记录状态, 门诊标志
    From 门诊费用记录
    Where 收费类别 Not In ('5', '6', '7') And 医嘱序号 + 0 = v_医嘱id And 记帐费用 = 1 And 记录状态 = 0 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id = v_医嘱id
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id)))
    Order By 记录性质, NO, 序号;

  --查找当前标本的相关申请
  Cursor c_Samplequest(v_医嘱id In Number) Is
    Select Distinct ID As 医嘱id, 病人来源 From 病人医嘱记录 Where v_医嘱id In (ID, 相关id);

  v_执行 Number(1);
  v_No   病人医嘱发送.No%Type;
  v_性质 病人医嘱发送.记录性质%Type;
  v_序号 Varchar2(1000);

  v_医嘱id   病人医嘱发送.医嘱id%Type;
  v_相关id   病人医嘱记录.相关id%Type;
  v_费用性质 病人医嘱发送.记录性质%Type;
  v_样本条码 病人医嘱发送.样本条码%Type;
  v_Records  Varchar2(2000);
  v_Currrec  Varchar2(50);
  v_Fields   Varchar2(50);
  v_Count    Number(18);
  v_病人id   病人医嘱记录.病人id%Type;
  v_主页id   病人医嘱记录.主页id%Type;
  v_是否出院 Number; --0=出院,1=在院
  v_记录状态 Number;
  v_病人来源 病人医嘱记录.病人来源%Type;
  v_Date     Date;
  Err_Custom Exception;
  v_Error Varchar2(100);
  n_Par   Number;
Begin
  Select Sysdate Into v_Date From Dual;
  --执行后自动审核对应的记帐划价单(不包含药品)
  Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_执行 From Dual;

  v_Records := 医嘱id || '|';

  While v_Records Is Not Null Loop
  
    v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields  := v_Currrec;
    v_医嘱id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    v_相关id  := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    If 接收人_In Is Null Then
      Update 病人医嘱发送 Set 接收人 = Null, 接收时间 = Null, 接收批次 = Null Where 医嘱id In (v_医嘱id, v_相关id);
      Update 病人医嘱发送
      Set 执行状态 = Decode(样本条码, Null, 0, 1)
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID In (v_医嘱id, v_相关id) And 相关id Is Null);
      For r_Samplequest In c_Samplequest(v_相关id) Loop
        If r_Samplequest.病人来源 = 2 Then
          Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
          Into v_费用性质
          From 病人医嘱发送
          Where 医嘱id = r_Samplequest.医嘱id;
        Else
          v_费用性质 := 1;
        End If;
        If v_费用性质 = 2 Then
          --2.费用执行处理
          Update 住院费用记录
          Set 执行状态 = 0, 执行时间 = Null, 执行人 = 接收人_In
          Where 收费类别 Not In ('5', '6', '7') And
                (医嘱序号, 记录性质, NO) In
                (Select 医嘱id, 记录性质, NO
                 From 病人医嘱附费
                 Where 医嘱id = r_Samplequest.医嘱id
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null)
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Null) And 采样人 Is Null);
        Else
          Update 门诊费用记录
          Set 执行状态 = 0, 执行时间 = Null, 执行人 = 接收人_In
          Where 收费类别 Not In ('5', '6', '7') And
                (医嘱序号, 记录性质, NO) In
                (Select 医嘱id, 记录性质, NO
                 From 病人医嘱附费
                 Where 医嘱id = r_Samplequest.医嘱id
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null)
                 Union All
                 Select 医嘱id, 记录性质, NO
                 From 病人医嘱发送
                 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Null) And 采样人 Is Null);
        End If;
      End Loop;
    Else
      --判断是否已出院，如果已出院负不完成登记
      Begin
        If v_主页id Is Null Then
          Select a.病人id, a.主页id, a.病人来源
          Into v_病人id, v_主页id, v_病人来源
          From 病人医嘱记录 A, 病案主页 B
          Where a.病人id = b.病人id And a.主页id = b.主页id(+) And a.Id = v_医嘱id;
        End If;
      Exception
        When Others Then
          v_病人来源 := 1;
      End;
      If v_病人来源 = 2 Then
        If Nvl(v_主页id, 0) > 0 Then
          Select Decode(出院日期, Null, 1, 0)
          Into v_是否出院
          From 病案主页
          Where 病人id = v_病人id And 主页id = v_主页id;
        Else
          v_是否出院 := 0;
        End If;
      
        If v_是否出院 = 0 Then
          --出院的才处理
          Begin
            Select Nvl(记录状态, 0)
            Into v_记录状态
            From 住院费用记录
            Where 医嘱序号 = v_医嘱id And Nvl(记录状态, 0) = 0 And Rownum = 1;
          Exception
            When Others Then
              v_记录状态 := 1;
          End;
        
          Select Nvl(样本条码, 0) Into v_样本条码 From 病人医嘱发送 Where 医嘱id = v_医嘱id;
          If v_样本条码 = 0 Then
            v_Error := '病人已出院不能完成登记!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      --检查医嘱是否收费
      n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
      If n_Par = 1 Then
        For r_Samplequest In c_Samplequest(v_相关id) Loop
          For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
            If r_Verify.记录状态 = 0 Then
              If r_Verify.门诊标志 = 1 Then
                v_Error := '标本未收费，不允许执行，请联系管理员！';
                Raise Err_Custom;
              Elsif r_Verify.门诊标志 = 2 Then
                v_Error := '标本未记账，不允许执行，请联系管理员！';
                Raise Err_Custom;
              End If;
            End If;
          End Loop;
        End Loop;
      End If;
    
      Update 病人医嘱发送
      Set 接收人 = 接收人_In, 接收时间 = v_Date, 接收批次 = 接收批次_In, 重采标本 = Null, 送检人 = 送检人_In
      Where 医嘱id In (v_医嘱id, v_相关id);
      Update 病人医嘱发送
      Set 执行状态 = 1
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID In (v_医嘱id, v_相关id) And 相关id Is Null);
      --记帐划价单是否转为记帐单
      --2.检查当前标本相关的申请的相关标本是否完成审核
      For r_Samplequest In c_Samplequest(v_相关id) Loop
        v_Count := 0;
        --r_SampleQuest.医嘱id申请已经完成,处理后续环节
        If v_Count = 0 Then
          If r_Samplequest.病人来源 = 2 Then
            Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
            Into v_费用性质
            From 病人医嘱发送
            Where 医嘱id = r_Samplequest.医嘱id;
          Else
            v_费用性质 := 1;
          End If;
          If v_费用性质 = 2 Then
            --2.费用执行处理
            Update 住院费用记录
            Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
            Where 收费类别 Not In ('5', '6', '7') And
                  (医嘱序号, 记录性质, NO) In
                  (Select 医嘱id, 记录性质, NO
                   From 病人医嘱附费
                   Where 医嘱id = r_Samplequest.医嘱id
                   Union All
                   Select 医嘱id, 记录性质, NO
                   From 病人医嘱发送
                   Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null)
                   Union All
                   Select 医嘱id, 记录性质, NO
                   From 病人医嘱发送
                   Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Null) And 采样人 Is Null);
          Else
            Update 门诊费用记录
            Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
            Where 收费类别 Not In ('5', '6', '7') And
                  (医嘱序号, 记录性质, NO) In
                  (Select 医嘱id, 记录性质, NO
                   From 病人医嘱附费
                   Where 医嘱id = r_Samplequest.医嘱id
                   Union All
                   Select 医嘱id, 记录性质, NO
                   From 病人医嘱发送
                   Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Not Null)
                   Union All
                   Select 医嘱id, 记录性质, NO
                   From 病人医嘱发送
                   Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = r_Samplequest.医嘱id And 相关id Is Null) And 采样人 Is Null);
          End If;
          --3.自动审核记帐
          If v_执行 = 1 Then
            For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
              If r_Verify.No || ',' || r_Verify.记录性质 <> v_No || ',' || v_性质 Then
                If v_序号 Is Not Null Then
                  If v_费用性质 = 1 Then
                    Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
                  Elsif v_费用性质 = 2 Then
                    Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
                  End If;
                End If;
                v_序号 := Null;
              End If;
              v_No   := r_Verify.No;
              v_性质 := r_Verify.记录性质;
              v_序号 := v_序号 || ',' || r_Verify.序号;
            End Loop;
            If v_序号 Is Not Null Then
              If v_费用性质 = 1 Then
                Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
              Elsif v_费用性质 = 2 Then
                Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
              End If;
            End If;
          End If;
        End If;
      End Loop;
    End If;
    v_Records := Substr('|' || v_Records, Length('|' || v_Currrec || '|') + 1);
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱发送_Sampleinput;
/

--108340:蔡青松,2017-04-18,执行相关操作之前先判断是否收费
Create Or Replace Procedure Zl_检验预置条码_采集完成
(
  医嘱内容_In Varchar2, --内容包括多个医嘱ID使用","分隔 
  人员编号_In 人员表.编号%Type := Null,
  人员姓名_In 人员表.姓名%Type := Null --Null=取消，不为空时完成采集 
) Is
  --查找当前标本的相关申请 
  Cursor c_Samplequest(v_医嘱id In Varchar2) Is
    Select /*+ rule */
    Distinct ID As 医嘱id, 病人来源
    From 病人医嘱记录 A, 病人医嘱发送 B
    Where a.Id = b.医嘱id And b.接收人 Is Null And 相关id Is Null And
          a.Id In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist)));

  --未审核的费用行(不包含药品) 
  Cursor c_Verify(v_医嘱id In Varchar2) Is
    Select /*+ rule */
    Distinct 记录性质, NO, 序号, 记录状态, 门诊标志
    From 住院费用记录
    Where 收费类别 Not In ('5', '6', '7') And
          医嘱序号 + 0 In
          (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist))) And 相关id Is Null) And
          记帐费用 = 1 And 记录状态 = 0 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist)))
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID
                                        From 病人医嘱记录
                                        Where ID In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist))) And
                                              相关id Is Null) And 接收人 Is Null)
    Union All
    Select /*+ rule */
    Distinct 记录性质, NO, 序号, 记录状态, 门诊标志
    From 门诊费用记录
    Where 收费类别 Not In ('5', '6', '7') And
          医嘱序号 + 0 In
          (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist))) And 相关id Is Null) And
          记帐费用 = 1 And 记录状态 = 0 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist)))
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID
                                        From 病人医嘱记录
                                        Where ID In (Select * From Table(Cast(f_Num2list(v_医嘱id) As Zltools.t_Numlist))) And
                                              相关id Is Null) And 接收人 Is Null)
    Order By 记录性质, NO, 序号;

  v_检验标本记录 Number(18);
  v_执行状态     Number(1);
  v_接收人       Varchar2(50);
  v_Error        Varchar2(100);
  V_执行         Number;
  v_No           病人医嘱发送.No%Type;
  v_性质         病人医嘱发送.记录性质%Type;
  v_序号         Varchar2(1000);
  Err_Custom Exception;
  n_Par Number;
Begin

  If 人员姓名_In Is Not Null Then
    --检查标本是否被核收或接收 
    Begin
      Select /*+ rule */
       Nvl(c.Id, 0), b.执行状态, b.接收人
      Into v_检验标本记录, v_执行状态, v_接收人
      From 病人医嘱记录 A, 病人医嘱发送 B, 检验标本记录 C
      Where a.Id = b.医嘱id And a.相关id = c.医嘱id(+) And
            a.Id In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist)));
    Exception
      When Others Then
        v_检验标本记录 := 0;
    End;
  
    If v_检验标本记录 <> 0 Then
      v_Error := '标本已被检验科核收不能完成采集!';
      Raise Err_Custom;
    End If;
  
    If v_执行状态 <> 2 And v_接收人 Is Not Null Then
      v_Error := '标本已被检验科签收不能完成采集!';
      Raise Err_Custom;
    End If;
  
    --检查医嘱是否收费
    n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
    If n_Par = 1 Then
      For r_Verify In c_Verify(医嘱内容_In) Loop
        If r_Verify.记录状态 = 0 Then
          If r_Verify.门诊标志 = 1 Then
            v_Error := '标本未收费，不允许执行，请联系管理员！';
            Raise Err_Custom;
          Elsif r_Verify.门诊标志 = 2 Then
            v_Error := '标本未记账，不允许执行，请联系管理员！';
            Raise Err_Custom;
          End If;
        End If;
      End Loop;
    End If;
  
    Update /*+ rule */ 检验拒收记录
    Set 重采人 = 人员姓名_In, 重采时间 = Sysdate
    Where 医嘱id In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist)));
  
    --更新采集信息(检验和采集） 
    Update /*+ rule */ 病人医嘱发送
    Set 采样人 = 人员姓名_In, 采样时间 = Sysdate, 执行状态 = Decode(执行状态, 2, 0, 执行状态),
        重采标本 = Decode(Nvl(重采标本, 0), 0, Decode(执行状态, 2, 1, 0), 重采标本), 执行说明 = Null
    Where 医嘱id In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist)));
  
    --更新医嘱和费用记录 
    For r_Samplequest In c_Samplequest(医嘱内容_In) Loop
      If r_Samplequest.病人来源 = 2 Then
        --2.费用执行处理 
        Update 住院费用记录
        Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In (Select 医嘱id, 记录性质, NO
                                   From 病人医嘱附费
                                   Where 医嘱id = r_Samplequest.医嘱id
                                   Union All
                                   Select 医嘱id, 记录性质, NO
                                   From 病人医嘱发送
                                   Where 医嘱id In (Select ID
                                                  From 病人医嘱记录 A, 病人医嘱发送 B
                                                  Where a.Id = b.医嘱id And r_Samplequest.医嘱id In (a.Id) And a.相关id Is Null And
                                                        b.执行状态 In (0, 2) And b.接收人 Is Null));
      Else
        --2.费用执行处理 
        Update 门诊费用记录
        Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In (Select 医嘱id, 记录性质, NO
                                   From 病人医嘱附费
                                   Where 医嘱id = r_Samplequest.医嘱id
                                   Union All
                                   Select 医嘱id, 记录性质, NO
                                   From 病人医嘱发送
                                   Where 医嘱id In (Select ID
                                                  From 病人医嘱记录 A, 病人医嘱发送 B
                                                  Where a.Id = b.医嘱id And r_Samplequest.医嘱id In (a.Id) And a.相关id Is Null And
                                                        b.执行状态 In (0, 2) And b.接收人 Is Null));
      End If;
    End Loop;
  
    --更新执行状态(只更新采集） 
    Update /*+ rule */ 病人医嘱发送
    Set 执行状态 = 1, 完成人 = 人员姓名_In, 完成时间 = Sysdate
    Where 医嘱id In
          (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist))) And 相关id Is Null);
    --执行后自动审核对应的记帐划价单(不包含药品)
    Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_执行 From Dual;
    --3.自动审核记帐 
    For r_Verify In c_Verify(医嘱内容_In) Loop
      If r_Verify.No || ',' || r_Verify.记录性质 <> v_No || ',' || v_性质 Then
        If v_序号 Is Not Null Then
          If v_性质 = 1 Then
            Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          Elsif v_性质 = 2 Then
            Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          End If;
        End If;
        v_序号 := Null;
      End If;
      v_No   := r_Verify.No;
      v_性质 := r_Verify.记录性质;
      v_序号 := v_序号 || ',' || r_Verify.序号;
    End Loop;
    If v_序号 Is Not Null Then
      If v_性质 = 1 Then
        Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
      Elsif v_性质 = 2 Then
        Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
      End If;
    End If;
  
  Else
    --检查标本是否被核收或接收 
    Begin
      Select /*+ rule */
       Nvl(c.Id, 0), b.执行状态, b.接收人
      Into v_检验标本记录, v_执行状态, v_接收人
      From 病人医嘱记录 A, 病人医嘱发送 B, 检验标本记录 C
      Where a.Id = b.医嘱id And a.相关id = c.医嘱id(+) And
            a.Id In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist)));
    Exception
      When Others Then
        v_检验标本记录 := 0;
    End;
  
    If v_检验标本记录 <> 0 Then
      v_Error := '标本已被检验科核收不能取消完成采集!';
      Raise Err_Custom;
    End If;
  
    If v_执行状态 <> 2 And v_接收人 Is Not Null Then
      v_Error := '标本已被检验科签收不能取消完成采集!';
      Raise Err_Custom;
    End If;
  
    Update /*+ rule */ 病人医嘱发送
    Set 采样人 = Null, 采样时间 = Null, 执行状态 = 0, 执行说明 = Null, 完成人 = Null, 完成时间 = Null
    Where 医嘱id In (Select ID
                   From 病人医嘱记录
                   Where ID In (Select * From Table(Cast(f_Num2list(医嘱内容_In) As Zltools.t_Numlist))));
  
    For r_Samplequest In c_Samplequest(医嘱内容_In) Loop
    
      If r_Samplequest.病人来源 = 2 Then
        --2.费用执行处理 
        Update 住院费用记录
        Set 执行状态 = 0, 执行时间 = Null, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID) And 相关id Is Null) And
                     执行状态 In (0, 2) And 接收人 Is Null);
      Else
        Update 门诊费用记录
        Set 执行状态 = 0, 执行时间 = Null, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID) And 相关id Is Null) And
                     执行状态 In (0, 2) And 接收人 Is Null);
      End If;
    End Loop;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验预置条码_采集完成;
/

--108340:蔡青松,2017-04-18,执行相关操作之前先判断是否收费
Create Or Replace Procedure Zl_检验标本记录_报告审核
(
  Id_In       检验标本记录.Id%Type,
  审核人_In   检验标本记录.审核人%Type := Null,
  人员编号_In 人员表.编号%Type := Null,
  人员姓名_In 人员表.姓名%Type := Null
) Is

  --未审核的费用行(不包含药品) 
  Cursor c_Verify(v_医嘱id In Number) Is
    Select Distinct 2 As 记录性质, NO, 序号, 记录状态, 门诊标志
    From 住院费用记录
    Where 收费类别 Not In ('5', '6', '7') And 记帐费用 = 1 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id = v_医嘱id
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id))) And 医嘱序号 = v_医嘱id
    Union All
    Select Distinct 1 As 记录性质, NO, 序号, 记录状态,门诊标志
    From 门诊费用记录
    Where 收费类别 Not In ('5', '6', '7') And 记帐费用 = 1 And 价格父号 Is Null And
          (记录性质, NO) In (Select 记录性质, NO
                         From 病人医嘱附费
                         Where 医嘱id = v_医嘱id
                         Union All
                         Select 记录性质, NO
                         From 病人医嘱发送
                         Where 医嘱id In (Select ID From 病人医嘱记录 Where v_医嘱id In (ID, 相关id))) And 医嘱序号 = v_医嘱id
    Order By 记录性质, NO, 序号;

  --查找当前标本的相关申请 
  Cursor c_Samplequest(v_微生物 In Number) Is
    Select Distinct 医嘱id, 病人来源
    From (Select a.医嘱id, b.病人来源
           From 检验项目分布 A, 检验标本记录 B
           Where 0 = v_微生物 And a.标本id = Id_In And a.医嘱id Is Not Null And a.标本id = b.Id
           Union
           Select a.医嘱id, b.病人来源
           From 检验项目分布 A, 检验标本记录 B
           Where 1 = v_微生物 And a.标本id = Id_In And a.医嘱id Is Not Null And a.标本id = b.Id
           Union
           Select b.Id As 医嘱id, a.病人来源
           From 检验标本记录 A, 病人医嘱记录 B
           Where a.Id = Id_In And a.医嘱id = b.相关id);

  Cursor c_Stuff
  (
    v_No     Varchar2,
    v_主页id Number
  ) Is
    Select NO, 单据, 库房id
    From 未发药品记录
    Where NO = v_No And 单据 In (24, 25, 26) And 库房id Is Not Null And Not Exists
     (Select 1 From Dual Where zl_GetSysParameter(Decode(v_主页id, Null, 92, 63)) = '1') And Exists
     (Select a.序号
           From 住院费用记录 A, 材料特性 B
           Where a.记录性质 = 2 And a.记录状态 = 1 And a.No = v_No And a.收费细目id = b.材料id And b.跟踪在用 = 1
           Union All
           Select a.序号
           From 门诊费用记录 A, 材料特性 B
           Where a.记录性质 = 2 And a.记录状态 = 1 And a.No = v_No And a.收费细目id = b.材料id And b.跟踪在用 = 1)
    Order By 库房id;

  v_执行  Number(1);
  v_No    病人医嘱发送.No%Type;
  v_Nonew 病人医嘱发送.No%Type;
  v_性质  病人医嘱发送.记录性质%Type;
  v_序号  Varchar2(1000);

  v_Count      Number(18);
  v_Counts     Number(18);
  v_微生物标本 Number(1) := 0;
  v_主页id     Number(18);
  v_婴儿       Number(1);
  v_年龄       Varchar2(100);
  v_仪器       Number(18);
  v_Intloop    Number;
  Err_Custom Exception;
  v_Error Varchar2(100);

  n_Par Number;
Begin
  Select Nvl(婴儿, 0), 年龄 Into v_婴儿, v_年龄 From 检验标本记录 Where ID = Id_In;

  --执行后自动审核对应的记帐划价单(不包含药品) 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_执行 From Dual;

  v_微生物标本 := 0;
  Begin
    Select 1 Into v_微生物标本 From 检验标本记录 Where 微生物标本 = 1 And ID = Id_In;
  Exception
    When Others Then
      v_微生物标本 := 0;
  End;

  --先判断医嘱是否收费
  n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
  If n_Par = 1 Then
    For r_Samplequest In c_Samplequest(v_微生物标本) Loop
      For r_相关医嘱 In (Select ID As 医嘱id From 病人医嘱记录 Where 相关id = r_Samplequest.医嘱id) Loop
        For r_Verify In c_Verify(r_相关医嘱.医嘱id) Loop
          If r_Verify.记录状态 = 0 Then
            If r_Verify.门诊标志 = 1 Then
              v_Error := '标本未收费，不允许执行，请联系管理员！';
              Raise Err_Custom;
            Elsif r_Verify.门诊标志 = 2 Then
              v_Error := '标本未记账，不允许执行，请联系管理员！';
              Raise Err_Custom;
            End If;
          End If;
        End Loop;
      End Loop;
    End Loop;
  End If;

  --1.置本标本的状态及审核人和时间 
  Update 检验标本记录
  Set 审核人 = Decode(审核人_In, Null, 人员姓名_In, 审核人_In), 审核时间 = Sysdate, 样本状态 = 2
  Where ID = Id_In;

  --记录审核过程 
  Insert Into 检验操作记录
    (ID, 标本id, 操作类型, 操作员, 操作时间)
  Values
    (检验操作记录_Id.Nextval, Id_In, 0, Decode(审核人_In, Null, 人员姓名_In, 审核人_In), Sysdate);

  --2.检查当前标本相关的申请的相关标本是否完成审核 
  For r_Samplequest In c_Samplequest(v_微生物标本) Loop
  
    v_Count := 0;
  
    If v_微生物标本 = 0 Then
      Begin
        Select Nvl(Count(1), 0)
        Into v_Count
        From 检验标本记录
        Where 样本状态 < 2 And ID In (Select 标本id From 检验项目分布 Where 医嘱id = r_Samplequest.医嘱id);
      Exception
        When Others Then
          v_Count := 0;
      End;
    End If;
  
    --r_SampleQuest.医嘱id申请已经完成,处理后续环节 
    If v_Count = 0 Then
    
      --1.置申请单的执行状态 
      Update 病人医嘱发送
      Set 执行状态 = 1, 完成人 = 人员姓名_In, 完成时间 = Sysdate
      Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id));
    
      Update 病人医嘱发送
      Set 执行状态 = 1, 完成人 = 人员姓名_In, 完成时间 = Sysdate
      Where 医嘱id In (Select 相关id
                     From 病人医嘱记录
                     Where ID In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
    
      If r_Samplequest.病人来源 = 2 Then
        --2.费用执行处理 
        Update 住院费用记录
        Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
      Else
        Update 门诊费用记录
        Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
      End If;
    
      --3.自动审核记帐 
      If v_执行 = 1 Then
        Select Count(*) Into v_Counts From 病人医嘱记录 Where 相关id = r_Samplequest.医嘱id;
        If v_Counts > 0 Then
          For r_相关医嘱 In (Select ID As 医嘱id From 病人医嘱记录 Where 相关id = r_Samplequest.医嘱id) Loop
            For r_Verify In c_Verify(r_相关医嘱.医嘱id) Loop
              If r_Verify.No || ',' || r_Verify.记录性质 <> v_No || ',' || v_性质 Then
                If v_序号 Is Not Null Then
                  If v_性质 = 1 Then
                    Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
                  Elsif v_性质 = 2 Then
                    Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
                  End If;
                End If;
                v_序号 := Null;
              End If;
              v_No   := r_Verify.No;
              v_性质 := r_Verify.记录性质;
              v_序号 := v_序号 || ',' || r_Verify.序号;
            End Loop;
          End Loop;
        Else
          For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
            If r_Verify.No || ',' || r_Verify.记录性质 <> v_No || ',' || v_性质 Then
              If v_序号 Is Not Null Then
                If v_性质 = 1 Then
                  Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
                Elsif v_性质 = 2 Then
                  Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
                End If;
              End If;
              v_序号 := Null;
            End If;
            v_No   := r_Verify.No;
            v_性质 := r_Verify.记录性质;
            v_序号 := v_序号 || ',' || r_Verify.序号;
          End Loop;
        End If;
        If v_序号 Is Not Null Then
          If v_性质 = 1 Then
            Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          Elsif v_性质 = 2 Then
            Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          End If;
          v_序号 := Null;
          --  v_性质 := null; 
        End If;
      End If;
    
      --审核试剂消耗单 
      v_Intloop := 1;
    
      Select 仪器id Into v_仪器 From 检验标本记录 Where ID = Id_In;
      For r_检验试剂 In (Select c.材料id, c.数量
                     From 病人医嘱记录 A, 检验报告项目 B, 检验试剂关系 C
                     Where a.相关id = r_Samplequest.医嘱id And a.诊疗项目id = b.诊疗项目id And b.报告项目id = c.项目id And c.仪器id = v_仪器) Loop
        Zl_检验试剂记录_Insert(r_Samplequest.医嘱id, v_Intloop, r_检验试剂.材料id, r_检验试剂.数量);
        v_Intloop := v_Intloop + 1;
      End Loop;
      Select Count(*) Into v_Intloop From 检验试剂记录 Where 医嘱id = r_Samplequest.医嘱id And NO Is Null;
      If v_Intloop > 1 Then
        v_Nonew := Nextno(14);
        Update 检验试剂记录 Set NO = v_Nonew Where 医嘱id = r_Samplequest.医嘱id;
      End If;
      If v_Nonew Is Not Null Then
      
        Zl_检验试剂记录_Bill(r_Samplequest.医嘱id, v_Nonew);
      
        v_主页id := Null;
        Select 主页id Into v_主页id From 病人医嘱记录 A Where ID = r_Samplequest.医嘱id;
      
        If v_主页id Is Null Then
          Zl_门诊记帐记录_Verify(v_Nonew, 人员编号_In, 人员姓名_In);
        Else
          Zl_住院记帐记录_Verify(v_Nonew, 人员编号_In, 人员姓名_In);
        End If;
      
        --如果记帐没有自动发料,则自动发料,否则不处理 
        For r_Stuff In c_Stuff(v_Nonew, v_主页id) Loop
          Zl_材料收发记录_处方发料(r_Stuff.库房id, 25, v_Nonew, 人员姓名_In, 人员姓名_In, 人员姓名_In, 1, Sysdate);
        End Loop;
      End If;
    End If;
  End Loop;
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 9, 0 || ',' || Id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验标本记录_报告审核;
/

--108340:蔡青松,2017-04-18,执行相关操作之前先判断是否收费
Create Or Replace Procedure Zl_检验标本记录_标本核收
(
  Id_In         In 检验标本记录.Id%Type,
  医嘱id_In     In 检验标本记录.医嘱id%Type,
  多个医嘱_In   In Varchar2, --用于更新多个医嘱的执行状态 
  复盖标本id_In In 检验标本记录.Id%Type := 0, --补填时指向另一个标本时复盖指向的标本 
  标本序号_In   In 检验标本记录.标本序号%Type,
  采样时间_In   In 检验标本记录.采样时间%Type,
  采样人_In     In 检验标本记录.采样人%Type,
  仪器id_In     In 检验标本记录.仪器id%Type,
  核收时间_In   In 检验标本记录.核收时间%Type,
  标本形态_In   In 检验标本记录.标本形态%Type,
  检验人_In     In 检验标本记录.检验人%Type := Null,
  检验时间_In   In 检验标本记录.检验时间%Type := Null,
  微生物标本_In In 检验标本记录.微生物标本%Type := Null,
  标本类别_In   In 检验标本记录.标本类别%Type := 0,
  检验备注_In   In 检验标本记录.检验备注%Type := Null,
  姓名_In       In 检验标本记录.姓名%Type := Null,
  性别_In       In 检验标本记录.性别%Type := Null,
  年龄_In       In 检验标本记录.年龄%Type := Null,
  No_In         In 检验标本记录.No%Type := Null,
  标本类型_In   In 检验标本记录.标本类型%Type := Null,
  申请科室id_In In 检验标本记录.申请科室id%Type := Null,
  申请人_In     In 检验标本记录.申请人%Type := Null,
  标识号_In     In 检验标本记录.标识号%Type := Null,
  床号_In       In 检验标本记录.床号%Type := Null,
  病人科室_In   In 检验标本记录.病人科室%Type := Null,
  检验项目_In   In 检验标本记录.检验项目%Type := Null,
  申请类型_In   In 检验标本记录.申请类型%Type := Null,
  病人id_In     In 检验标本记录.病人id%Type := Null,
  执行科室_In   In 检验标本记录.执行科室id%Type := Null,
  人员编号_In   In 人员表.编号%Type := Null,
  人员姓名_In   In 人员表.姓名%Type := Null
) Is

  Cursor v_Advice Is
    Select /*+ Rule */
    Distinct a.Id, a.开嘱时间, a.标本部位, f.样本条码, a.执行科室id, a.诊疗项目id, a.开嘱科室id, a.开嘱医生, a.病人id, a.病人来源, a.婴儿, a.紧急标志 As 紧急,
             b.门诊号, b.住院号, b.出生日期, a.挂号单, Decode(c.主页id, 0, Null, c.主页id) As 主页id, d.操作类型, f.接收人, f.接收时间
    From 病人医嘱记录 A, 病人医嘱发送 F, 病人信息 B, 病案主页 C, 诊疗项目目录 D
    Where a.相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist))) And a.Id = f.医嘱id And
          a.病人id = b.病人id And a.病人id = c.病人id(+) And a.主页id = c.主页id(+) And a.诊疗项目id = d.Id(+);

  Cursor v_Advice_1 Is
    Select /*+ Rule */
    Distinct b.No As 单据号, a.相关id
    From 病人医嘱记录 A, 病人医嘱发送 B
    Where a.Id = b.医嘱id And a.相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
    Union All
    Select /*+ Rule */
    Distinct b.No As 单据号, a.相关id
    From 病人医嘱记录 A, 病人医嘱发送 B
    Where a.Id = b.医嘱id And a.Id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)));

  Cursor v_Patient Is
    Select 病人id, 住院号, 门诊号, 出生日期 From 病人信息 Where 病人id = 病人id_In;

  --未审核的费用行(不包含药品) 
  Cursor c_Verify(v_医嘱id In Number) Is
    Select /*+ Rule */
    Distinct a.记录性质, a.No, a.序号, a.医嘱序号, a.门诊标志, a.记录状态
    From 住院费用记录 A, 病人医嘱发送 B,
         (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From 病人医嘱记录
           Where 相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))) C
    Where a.收费类别 Not In ('5', '6', '7') And a.医嘱序号 = c.Id And a.记录状态 = 0 And 价格父号 Is Null And a.医嘱序号 = b.医嘱id And
          a.记录性质 = b.记录性质 And a.No = b.No And a.记帐费用 = 1
    Union All
    Select /*+ Rule */
    Distinct a.记录性质, a.No, a.序号, a.医嘱序号, a.门诊标志, a.记录状态
    From 住院费用记录 A, 病人医嘱附费 B,
         (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From 病人医嘱记录
           Where 相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))) C
    Where a.收费类别 Not In ('5', '6', '7') And a.医嘱序号 = c.Id And a.记录状态 = 0 And 价格父号 Is Null And a.医嘱序号 = b.医嘱id And
          a.记录性质 = b.记录性质 And a.No = b.No And a.记帐费用 = 1
    Union All
    Select /*+ Rule */
    Distinct a.记录性质, a.No, a.序号, a.医嘱序号, a.门诊标志, a.记录状态
    From 门诊费用记录 A, 病人医嘱发送 B,
         (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From 病人医嘱记录
           Where 相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))) C
    Where a.收费类别 Not In ('5', '6', '7') And a.医嘱序号 = c.Id And a.记录状态 = 0 And 价格父号 Is Null And a.医嘱序号 = b.医嘱id And
          a.记录性质 = b.记录性质 And a.No = b.No And a.记帐费用 = 1
    Union All
    Select /*+ Rule */
    Distinct a.记录性质, a.No, a.序号, a.医嘱序号, a.门诊标志, a.记录状态
    From 门诊费用记录 A, 病人医嘱附费 B,
         (Select ID
           From 病人医嘱记录
           Where ID In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From 病人医嘱记录
           Where 相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))) C
    Where a.收费类别 Not In ('5', '6', '7') And a.医嘱序号 = c.Id And a.记录状态 = 0 And 价格父号 Is Null And a.医嘱序号 = b.医嘱id And
          a.记录性质 = b.记录性质 And a.No = b.No And a.记帐费用 = 1
    Order By 记录性质, NO, 序号;

  --查找当前标本的相关申请 
  Cursor c_Samplequest(v_微生物 In Number) Is
    Select Distinct 医嘱id, 病人来源
    From (Select Decode(a.医嘱id, Null, b.医嘱id, a.医嘱id) As 医嘱id, b.病人来源
           From 检验项目分布 A, 检验标本记录 B
           Where Nvl(v_微生物, 0) = 0 And a.标本id = b.Id And b.医嘱id In (Select 医嘱id From 检验标本记录 Where ID = Id_In) And
                 a.医嘱id Is Not Null
           Union
           Select Decode(a.医嘱id, Null, b.医嘱id, a.医嘱id) As 医嘱id, b.病人来源
           From 检验项目分布 A, 检验标本记录 B
           Where 1 = v_微生物 And b.Id = a.标本id And b.Id = Id_In
           Union
           Select b.Id As 医嘱id, b.病人来源
           From 检验标本记录 A, 病人医嘱记录 B
           Where a.Id = Id_In And a.医嘱id In (b.Id, b.相关id));

  Cursor c_Stuff
  (
    v_No     Varchar2,
    v_主页id Number
  ) Is
    Select NO As 单据号, 单据, 库房id
    From 未发药品记录
    Where NO = v_No And 单据 In (24, 25, 26) And 库房id Is Not Null And Not Exists
     (Select 1 From Dual Where zl_GetSysParameter(Decode(v_主页id, Null, 92, 63)) = '1') And Exists
     (Select a.序号
           From 住院费用记录 A, 材料特性 B
           Where a.记录性质 = 2 And a.No = v_No And a.收费细目id = b.材料id And b.跟踪在用 = 1
           Union All
           Select a.序号
           From 门诊费用记录 A, 材料特性 B
           Where a.记录性质 = 2 And a.No = v_No And a.收费细目id = b.材料id And b.跟踪在用 = 1)
    Order By 库房id;

  r_Advice   v_Advice%RowType;
  r_Advice_1 v_Advice_1%RowType;
  r_Patient  v_Patient%RowType;

  Err_Custom Exception;
  v_Error Varchar2(1000);
  v_Flag  Number(18);

  v_Temp      Varchar2(255);
  v_Seq       Number;
  v_Union     Number;
  v_Patientid Number;
  v_Itemid    Number;
  v_Count     Number;
  v_执行      Number;
  v_No        病人医嘱发送.No%Type;
  v_性质      病人医嘱发送.记录性质%Type;
  v_序号      Varchar2(1000);
  v_主页id    Number(18);
  v_门诊标志  住院费用记录.门诊标志%Type;
  n_Count     Number;
  v_姓名      病人医嘱记录.姓名%Type;
  v_性别      病人医嘱记录.性别%Type;
  v_年龄      病人医嘱记录.年龄%Type;
  v_病人来源  病人医嘱记录.病人来源%Type;
  v_婴儿      病人医嘱记录.婴儿%Type;
  v_婴儿姓名  病人医嘱记录.姓名%Type;
  v_婴儿性别  病人医嘱记录.性别%Type;

  n_Par Number;
Begin

  If Nvl(复盖标本id_In, 0) > 0 Then
    Begin
      Select 姓名 Into v_Temp From 检验标本记录 Where ID = 复盖标本id_In And 姓名 Is Null;
    Exception
      When Others Then
        v_Error := '指定复盖的标本已被核收或已删除，请重新指定！';
        Raise Err_Custom;
    End;
  End If;

  If Nvl(医嘱id_In, 0) > 0 Then
    Select 姓名, 性别, 年龄, 病人来源, 婴儿
    Into v_姓名, v_性别, v_年龄, v_病人来源, v_婴儿
    From 病人医嘱记录
    Where ID = 医嘱id_In;
  
    If v_病人来源 <> 3 Then
      If Nvl(v_婴儿, 0) = 0 Then
        If v_姓名 <> 姓名_In Or v_性别 <> 性别_In Then
          v_Error := '病人姓名、性别、年龄和医嘱不符不能保存，请检查或修改病人信息后再进行保存！';
          Raise Err_Custom;
        End If;
      Else
        Select b.婴儿姓名, b.婴儿性别
        Into v_婴儿姓名, v_婴儿性别
        From 病人医嘱记录 A, 病人新生儿记录 B
        Where a.病人id = b.病人id And a.主页id = b.主页id And a.婴儿 = b.序号 And
              a.相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist))) And Rownum = 1;
      
        If v_婴儿姓名 <> 姓名_In Or v_婴儿性别 <> 性别_In Then
          v_Error := '病人姓名、性别、年龄和医嘱不符不能保存，请检查或修改病人信息后再进行保存！';
          Raise Err_Custom;
        End If;
      End If;
    End If;
  
    Select Count(ID) Into v_Flag From 检验标本记录 Where 医嘱id = 医嘱id_In And ID <> Id_In;
    If v_Flag > 0 Then
      Select Count(Distinct b.报告项目id)
      Into v_Flag
      From 病人医嘱记录 A, 检验报告项目 B
      Where a.诊疗项目id = b.诊疗项目id And a.相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)));
    
      Select Count(a.项目id)
      Into n_Count
      From 检验项目分布 A
      Where a.医嘱id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist))) And a.标本id <> Id_In;
      If (v_Flag - n_Count) <= 0 Then
        v_Error := '当前医嘱已被核收，不能重复核收！';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  --判断医嘱是否收费
  n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
  If n_Par = 1 Then
    For r_Advice_1 In v_Advice_1 Loop
      For r_Verify In c_Verify(r_Advice_1.相关id) Loop
        If r_Verify.记录状态 = 0 Then
          If r_Verify.门诊标志 = 1 Then
            v_Error := '标本未收费，不允许执行，请联系管理员！';
            Raise Err_Custom;
          Elsif r_Verify.门诊标志 = 2 Then
            v_Error := '标本未记账，不允许执行，请联系管理员！';
            Raise Err_Custom;
          End If;
        End If;
      End Loop;
    End Loop;
  End If;

  If 医嘱id_In = 0 Then
    Open v_Patient;
    Fetch v_Patient
      Into r_Patient;
  
    If v_Patient%Found Then
      Zl_病人信息_锁定检查(r_Patient.病人id);
    End If;
  
    Update 检验标本记录
    Set 采样时间 = Decode(采样时间_In, Null, 采样时间, 采样时间_In), 采样人 = Decode(采样人_In, Null, 采样人, 采样人_In), 标本类型 = Nvl(标本类型_In, 标本类型),
        检验时间 = 检验时间_In, 姓名 = Decode(姓名_In, Null, 姓名, 姓名_In), 性别 = Decode(性别_In, Null, 性别, 性别_In),
        年龄 = Decode(年龄_In, Null, 年龄, 年龄_In), 年龄数字 = Decode(年龄_In, Null, Null, Zl_Val(年龄_In)),
        年龄单位 = Decode(年龄_In, Null, 年龄单位,
                       Decode(年龄_In, Null, Null, '成人', '成人', '婴儿', '婴儿',
                               Decode(Sign(Instr(年龄_In, '岁')), 1, '岁',
                                       Decode(Sign(Instr(年龄_In, '月')), 1, '月',
                                               Decode(Sign(Instr(年龄_In, '天')), 1, '天',
                                                       Decode(Sign(Instr(年龄_In, '小时')), 1, '小时', Null)))))),
        申请科室id = Decode(申请科室id_In, Null, 申请科室id, 申请科室id_In), 申请人 = Decode(申请人_In, Null, 申请人, 申请人_In),
        标本形态 = Decode(标本形态_In, Null, 标本形态, 标本形态_In), 标识号 = Decode(标识号_In, Null, 标识号, 标识号_In),
        床号 = Decode(床号_In, Null, 床号, 床号_In), 病人科室 = Decode(病人科室_In, Null, 病人科室, 病人科室_In),
        检验项目 = Decode(检验项目_In, Null, 检验项目, 检验项目_In), 病人id = Decode(病人id_In, Null, 病人id, 病人id_In),
        医嘱id = Decode(医嘱id_In, Null, 医嘱id, 0, 医嘱id, 医嘱id_In)
    Where ID = Id_In;
    If Sql%NotFound Then
      Insert Into 检验标本记录
        (ID, 医嘱id, 标本序号, 采样时间, 采样人, 标本类型, 核收人, 核收时间, 样本状态, 申请类型, 仪器id, 样本条码, 申请时间, 标本形态, 报告结果, 执行科室id, 检验人, 检验时间, 微生物标本,
         标本类别, 检验备注, 申请科室id, 申请人, 姓名, 性别, 年龄, 年龄数字, 年龄单位, 病人id, 病人来源, 婴儿, NO, 合并id, 标识号, 床号, 病人科室, 紧急, 门诊号, 住院号, 出生日期,
         挂号单, 主页id, 检验项目, 操作类型, 接收人, 接收时间)
      Values
        (Id_In, Decode(医嘱id_In, 0, Null, 医嘱id_In), 标本序号_In, 采样时间_In, 采样人_In, 标本类型_In, 人员姓名_In, 核收时间_In, 1, 申请类型_In,
         Decode(仪器id_In, 0, Null, 仪器id_In), Null, Null, 标本形态_In, 0, 执行科室_In, 检验人_In, 检验时间_In, 微生物标本_In, 标本类别_In, 检验备注_In,
         申请科室id_In, 申请人_In, 姓名_In, 性别_In, 年龄_In, Zl_Val(年龄_In),
         Decode(年龄_In, Null, Null, '成人', '成人', '婴儿', '婴儿',
                 Decode(Sign(Instr(年龄_In, '岁')), 1, '岁',
                         Decode(Sign(Instr(年龄_In, '月')), 1, '月',
                                 Decode(Sign(Instr(年龄_In, '天')), 1, '天', Decode(Sign(Instr(年龄_In, '小时')), 1, '小时', Null))))),
         病人id_In, Decode(r_Patient.住院号, Null, Decode(r_Patient.门诊号, Null, 3, 1), 2), 0, Null, Null, 标识号_In, 床号_In,
         病人科室_In, 标本类别_In, r_Patient.门诊号, r_Patient.住院号, r_Patient.出生日期, Null, Null, 检验项目_In, Null, Null, Null);
    End If;
    If Nvl(复盖标本id_In, 0) > 0 Then
      Zl_检验标本记录_Union(Id_In, 复盖标本id_In);
    End If;
    --记录核收和补填操作 
    Insert Into 检验操作记录
      (ID, 标本id, 操作类型, 操作员, 操作时间)
    Values
      (检验操作记录_Id.Nextval, Id_In, 2, 人员姓名_In, Sysdate);
    Close v_Patient;
  Else
    Open v_Advice;
    Fetch v_Advice
      Into r_Advice;
  
    If v_Advice%Found Then
      Zl_病人信息_锁定检查(r_Advice.病人id);
    End If;
  
    Update 检验标本记录
    Set 医嘱id = Decode(医嘱id_In, Null, 医嘱id, 0, 医嘱id, 医嘱id_In), 采样时间 = Decode(采样时间_In, Null, 采样时间, 采样时间_In),
        采样人 = Decode(采样人_In, Null, 采样人, 采样人_In), 标本序号 = Decode(标本序号_In, Null, 标本序号, 标本序号_In),
        标本类型 = Decode(标本类型_In, Null, Decode(标本类型, Null, r_Advice.标本部位, 标本类型), 标本类型_In),
        申请时间 = Decode(r_Advice.开嘱时间, Null, 申请时间, r_Advice.开嘱时间), 核收人 = Decode(核收人, Null, 人员姓名_In, 核收人),
        样本条码 = Decode(r_Advice.样本条码, Null, 样本条码, r_Advice.样本条码), 申请类型 = Decode(申请类型_In, Null, 申请类型, 申请类型_In),
        执行科室id = Decode(执行科室_In, Null, 执行科室id, 执行科室_In), 检验人 = Decode(检验人_In, Null, 检验人, 检验人_In),
        检验时间 = Decode(检验时间_In, Null, 检验时间, 检验时间_In), 检验备注 = Decode(检验备注_In, Null, 检验备注, 检验备注_In),
        申请科室id = Decode(申请科室id_In, Null, 申请科室id, 申请科室id_In), 申请人 = Decode(申请人_In, Null, 申请人, 申请人_In),
        姓名 = Decode(姓名_In, Null, 姓名, 姓名_In), 性别 = Decode(性别_In, Null, 性别, 性别_In), 年龄 = Decode(年龄_In, Null, 年龄, 年龄_In),
        年龄数字 = Decode(年龄_In, Null, 年龄数字, Zl_Val(年龄_In)),
        年龄单位 = Decode(年龄_In, Null, 年龄单位,
                       Decode(年龄_In, Null, Null, '成人', '成人', '婴儿', '婴儿',
                               Decode(Sign(Instr(年龄_In, '岁')), 1, '岁',
                                       Decode(Sign(Instr(年龄_In, '月')), 1, '月',
                                               Decode(Sign(Instr(年龄_In, '天')), 1, '天',
                                                       Decode(Sign(Instr(年龄_In, '小时')), 1, '小时', Null)))))),
        病人id = Decode(r_Advice.病人id, Null, 病人id, r_Advice.病人id), 病人来源 = Decode(r_Advice.病人来源, Null, 病人来源, r_Advice.病人来源),
        婴儿 = Decode(r_Advice.婴儿, 婴儿, r_Advice.婴儿), NO = Decode(No_In, Null, NO, No_In), 合并id = v_Union,
        标本形态 = Decode(标本形态_In, Null, 标本形态, 标本形态_In), 标识号 = Decode(标识号_In, Null, 标识号, 标识号_In),
        床号 = Decode(床号_In, Null, 床号, 床号_In), 病人科室 = Decode(病人科室_In, Null, 病人科室, 病人科室_In), 标本类别 = 标本类别_In,
        门诊号 = r_Advice.门诊号, 住院号 = r_Advice.住院号, 出生日期 = r_Advice.出生日期, 挂号单 = r_Advice.挂号单, 主页id = r_Advice.主页id,
        检验项目 = Decode(检验项目_In, Null, 检验项目, 检验项目_In), 操作类型 = r_Advice.操作类型, 接收人 = r_Advice.接收人, 接收时间 = r_Advice.接收时间
    Where ID = Id_In;
  
    If Sql%NotFound Then
      Insert Into 检验标本记录
        (ID, 医嘱id, 标本序号, 采样时间, 采样人, 标本类型, 核收人, 核收时间, 样本状态, 申请类型, 仪器id, 样本条码, 申请时间, 标本形态, 报告结果, 执行科室id, 检验人, 检验时间, 微生物标本,
         标本类别, 检验备注, 申请科室id, 申请人, 姓名, 性别, 年龄, 年龄数字, 年龄单位, 病人id, 病人来源, 婴儿, NO, 合并id, 标识号, 床号, 病人科室, 紧急, 门诊号, 住院号, 出生日期,
         挂号单, 主页id, 检验项目, 操作类型, 接收人, 接收时间)
      Values
        (Id_In, Decode(医嘱id_In, 0, Null, 医嘱id_In), 标本序号_In, 采样时间_In, 采样人_In, Nvl(标本类型_In, r_Advice.标本部位), 人员姓名_In,
         核收时间_In, 1, 申请类型_In, Decode(仪器id_In, 0, Null, 仪器id_In), r_Advice.样本条码, r_Advice.开嘱时间, 标本形态_In, 0, 执行科室_In,
         检验人_In, 检验时间_In, 微生物标本_In, 标本类别_In, 检验备注_In, 申请科室id_In, 申请人_In, 姓名_In, 性别_In, 年龄_In, Zl_Val(年龄_In),
         Decode(年龄_In, Null, Null, '成人', '成人', '婴儿', '婴儿',
                 Decode(Sign(Instr(年龄_In, '岁')), 1, '岁',
                         Decode(Sign(Instr(年龄_In, '月')), 1, '月',
                                 Decode(Sign(Instr(年龄_In, '天')), 1, '天', Decode(Sign(Instr(年龄_In, '小时')), 1, '小时', Null))))),
         r_Advice.病人id, r_Advice.病人来源, r_Advice.婴儿, No_In, v_Union, 标识号_In, 床号_In, 病人科室_In, r_Advice.紧急, r_Advice.门诊号,
         r_Advice.住院号, r_Advice.出生日期, r_Advice.挂号单, r_Advice.主页id, 检验项目_In, r_Advice.操作类型, r_Advice.接收人, r_Advice.接收时间);
    End If;
    If Nvl(复盖标本id_In, 0) > 0 Then
      Zl_检验标本记录_Union(Id_In, 复盖标本id_In);
    End If;
    Insert Into 检验操作记录
      (ID, 标本id, 操作类型, 操作员, 操作时间)
    Values
      (检验操作记录_Id.Nextval, Id_In, 2, 人员姓名_In, Sysdate);
  
    --查找主项目有时填写合并ID 
    Begin
      Select a.Id
      Into v_Union
      From 检验标本记录 A, 检验标本记录 B, 病人医嘱记录 C, 检验合并规则 D, 病人医嘱记录 E
      Where a.病人id = b.病人id And b.Id = Id_In And a.样本状态 = 1 And Nvl(a.病人id, 0) <> 0 And a.医嘱id = c.相关id And
            d.主项目id = c.诊疗项目id And d.合并项目id = e.诊疗项目id And e.Id = r_Advice.Id And Rownum = 1
      Order By a.核收时间 Desc;
    Exception
      When Others Then
        v_Union := Null;
    End;
    If Nvl(v_Union, 0) <> 0 Then
      Update 检验标本记录 Set 合并id = v_Union Where (ID = Id_In Or 医嘱id = r_Advice.Id);
    End If;
    --查找有了主项目时填写合并项目 
    Begin
      Select a.Id, a.病人id, c.主项目id
      Into v_Union, v_Patientid, v_Itemid
      From 检验标本记录 A, 病人医嘱记录 B, 检验合并规则 C
      Where a.医嘱id = b.相关id And b.诊疗项目id = c.主项目id And a.Id = Id_In And Rownum = 1;
    Exception
      When Others Then
        v_Union := Null;
    End;
    If Nvl(v_Union, 0) <> 0 Then
      Update 检验标本记录
      Set 合并id = v_Union
      Where ID In (Select a.Id
                   From 检验标本记录 A, 病人医嘱记录 B, 检验合并规则 C
                   Where a.医嘱id = b.相关id And b.诊疗项目id = c.合并项目id And c.主项目id = v_Itemid And a.病人id = v_Patientid And
                         a.样本状态 = 1);
    End If;
  
    v_Seq := 1;
    Close v_Advice;
    v_Flag := 0;
    Begin
      Select Nvl(Max(1), 0) Into v_Flag From 检验申请项目 Where 标本id = Id_In;
    Exception
      When Others Then
        v_Flag := 0;
    End;
    If v_Flag = 0 Then
      For r_Advice In v_Advice Loop
        Update 检验申请项目
        Set 标本id = Id_In, 诊疗项目id = r_Advice.诊疗项目id
        Where 标本id = Id_In And 诊疗项目id = r_Advice.诊疗项目id;
        If Sql%RowCount = 0 Then
          Insert Into 检验申请项目 (标本id, 诊疗项目id, 序号) Values (Id_In, r_Advice.诊疗项目id, v_Seq);
        End If;
        v_Seq := v_Seq + 1;
      End Loop;
    End If;
  
  End If;

  --根据参数来判断是否发料 
  For r_Advice_1 In v_Advice_1 Loop
    --如果记帐没有自动发料,则自动发料,否则不处理 
    For r_Stuff In c_Stuff(r_Advice_1.单据号, v_主页id) Loop
    
      Zl_材料收发记录_处方发料(r_Stuff.库房id, r_Stuff.单据, r_Stuff.单据号, 人员姓名_In, 人员姓名_In, 人员姓名_In, 1, Sysdate);
    End Loop;
  End Loop;

  Update /*+ Rule */ 病人医嘱发送
  Set 执行状态 = 3
  Where 执行状态 = 0 And
        医嘱id In (Select ID
                 From 病人医嘱记录
                 Where ID In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist)))
                 Union All
                 Select ID
                 From 病人医嘱记录
                 Where 相关id In (Select * From Table(Cast(f_Num2list(多个医嘱_In) As Zltools.t_Numlist))));
  --执行后自动审核对应的记帐划价单(不包含药品)
  Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_执行 From Dual;
  --2.检查当前标本相关的申请的相关标本是否完成审核 
  For r_Samplequest In c_Samplequest(微生物标本_In) Loop
  
    v_Count := 0;
  
    --r_SampleQuest.医嘱id申请已经完成,处理后续环节 
    If v_Count = 0 Then
    
      If r_Samplequest.病人来源 = 2 Then
        Update 住院费用记录
        Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
      Else
        Update 门诊费用记录
        Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 人员姓名_In
        Where 收费类别 Not In ('5', '6', '7') And
              (医嘱序号, 记录性质, NO) In
              (Select 医嘱id, 记录性质, NO
               From 病人医嘱附费
               Where 医嘱id = r_Samplequest.医嘱id
               Union All
               Select 医嘱id, 记录性质, NO
               From 病人医嘱发送
               Where 医嘱id In (Select ID From 病人医嘱记录 Where r_Samplequest.医嘱id In (ID, 相关id)));
      End If;
      --3.自动审核记帐 
      If Nvl(v_执行, 0) = 1 Then
        For r_Verify In c_Verify(r_Samplequest.医嘱id) Loop
          If r_Verify.No || ',' || r_Verify.记录性质 <> v_No || ',' || v_性质 Then
            If v_序号 Is Not Null Then
              If r_Verify.门诊标志 = 1 Then
                Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
              Elsif r_Verify.门诊标志 = 2 Then
                Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
              End If;
            End If;
            v_序号 := Null;
          End If;
          v_门诊标志 := r_Verify.门诊标志;
          v_No       := r_Verify.No;
          v_性质     := r_Verify.记录性质;
          v_序号     := v_序号 || ',' || r_Verify.序号;
        End Loop;
        If v_序号 Is Not Null Then
          If v_门诊标志 = 1 Then
            Zl_门诊记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          Elsif v_门诊标志 = 2 Then
            Zl_住院记帐记录_Verify(v_No, 人员编号_In, 人员姓名_In, Substr(v_序号, 2));
          End If;
        End If;
      End If;
    End If;
  End Loop;

  If Nvl(申请类型_In, 0) = 1 Then
    Zl_病人医嘱记录_屏蔽打印(医嘱id_In, 1);
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_检验标本记录_标本核收;
/

--107559:冉俊明,2017-04-17,增加终止停诊安排功能
Create Or Replace Procedure Zl_临床出诊停诊_Apply
(
  操作类型_In Number,
  Id_In       临床出诊停诊记录.Id%Type,
  开始时间_In 临床出诊停诊记录.开始时间%Type := Null,
  终止时间_In 临床出诊停诊记录.终止时间%Type := Null,
  停诊原因_In 临床出诊停诊记录.停诊原因%Type := Null,
  申请人_In   临床出诊停诊记录.申请人%Type := Null,
  申请时间_In 临床出诊停诊记录.申请时间%Type := Null,
  登记人_In   临床出诊停诊记录.登记人%Type := Null
) As
  --功能：退费申请以及取消申请
  --参数：
  --        操作类型_In：0-申请，else-取消申请
  --说明：
  n_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If 操作类型_In = 0 Then
    --申请
    If 开始时间_In <= Sysdate Then
      v_Error := '停诊时间的开始时间必须大于当前时间！';
      Raise Err_Custom;
    End If;
  
    If 开始时间_In >= 终止时间_In Then
      v_Error := '停诊时间的结束时间必须大于开始时间！';
      Raise Err_Custom;
    End If;
  
    Select Count(1)
    Into n_Count
    From 临床出诊停诊记录
    Where 记录id Is Null And Not (开始时间 > 终止时间_In Or Nvl(失效时间, 终止时间) < 开始时间_In) And 申请人 = 申请人_In And Rownum < 2;
    If n_Count <> 0 Then
      v_Error := '医生 ' || 申请人_In || ' 在当前停诊时间范围内已存在停诊安排，不能重复申请！';
      Raise Err_Custom;
    End If;
  
    Insert Into 临床出诊停诊记录
      (ID, 开始时间, 终止时间, 停诊原因, 申请人, 申请时间, 登记人)
    Values
      (临床出诊停诊记录_Id.Nextval, 开始时间_In, 终止时间_In, 停诊原因_In, 申请人_In, 申请时间_In, 登记人_In);
  
    Return;
  End If;

  --取消申请
  Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In And 审批人 Is Not Null;
  If n_Count <> 0 Then
    v_Error := '该申请已被审批，不能取消申请。';
    Raise Err_Custom;
  End If;

  Delete 临床出诊停诊记录 Where ID = Id_In;
  If Sql%NotFound Then
    v_Error := '该申请可能已被他人取消申请，请刷新后查看...';
    Raise Err_Custom;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊停诊_Apply;
/

--107559:冉俊明,2017-04-17,增加终止停诊安排功能
Create Or Replace Procedure Zl_临床出诊停诊_Stop
(
  Id_In       临床出诊停诊记录.Id%Type,
  终止人_In   临床出诊停诊记录.取消人%Type,
  终止时间_In 临床出诊停诊记录.失效时间%Type := Null
) As
  --功能：终止停诊安排
  --参数：
  --       终止时间_In：Null-立即终止，其它-具体的终止时间
  v_Error Varchar2(255);
  Err_Custom Exception;

  n_Count Number;
Begin
  If 终止时间_In Is Not Null Then
    If 终止时间_In < Sysdate Then
      v_Error := '终止时间必须大于当前时间！';
      Raise Err_Custom;
    End If;
  End If;

  Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In And 终止时间 < Sysdate;
  If n_Count <> 0 Then
    v_Error := '该停诊安排已失效，不能终止！';
    Raise Err_Custom;
  End If;

  Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In And 失效时间 Is Not Null;
  If n_Count <> 0 Then
    v_Error := '该停诊安排已被终止，不能再终止！';
    Raise Err_Custom;
  End If;

  Update 临床出诊停诊记录
  Set 失效时间 = Nvl(终止时间_In, Sysdate), 取消人 = 终止人_In, 取消时间 = Sysdate
  Where ID = Id_In And 审批人 Is Not Null;
  If Sql%NotFound Then
    v_Error := '该停诊安排还未审批，不能终止！';
    Raise Err_Custom;
  End If;

  For c_记录 In (Select a.Id, c.号码, a.停诊终止时间, a.是否序号控制, a.是否分时段
               From 临床出诊记录 A, 临床出诊停诊记录 B, 临床出诊号源 C
               Where ((a.替诊医生姓名 Is Null And a.医生id Is Not Null And a.医生姓名 = b.申请人) Or
                     (a.替诊医生姓名 Is Not Null And a.替诊医生id Is Not Null And a.替诊医生姓名 = b.申请人)) And a.号源id = c.Id And
                     b.Id = Id_In And (a.开始时间 Between b.开始时间 And b.终止时间 Or a.终止时间 Between b.开始时间 And b.终止时间) And
                     Nvl(a.是否发布, 0) = 1 And a.停诊终止时间 > Nvl(终止时间_In, Sysdate)) Loop
  
    Update 临床出诊记录
    Set 停诊开始时间 = Case
                   When 停诊开始时间 >= Nvl(终止时间_In, Sysdate) Then
                    Null
                   Else
                    停诊开始时间
                 End,
        停诊终止时间 = Case
                   When 停诊开始时间 >= Nvl(终止时间_In, Sysdate) Then
                    Null
                   Else
                    Nvl(终止时间_In, Sysdate)
                 End
    Where ID = c_记录.Id;
  
    --调整"临床出诊序号控制.是否停诊"为0
    Update 临床出诊序号控制
    Set 是否停诊 = 0
    Where 记录id = c_记录.Id And Nvl(是否停诊, 0) = 1 And 开始时间 Between Nvl(终止时间_In, Sysdate) And c_记录.停诊终止时间 And
          Nvl(c_记录.是否序号控制, 0) = 1 And Nvl(c_记录.是否分时段, 0) = 1;
  
    --消息推送
    -- 停诊类型(1-停诊,2-取消停诊),出诊记录ID,停诊号码
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 17, 2 || ',' || c_记录.Id || ',' || c_记录.号码;
    Exception
      When Others Then
        Null;
    End;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊停诊_Stop;
/

--108264:刘尔旋,2017-04-17,销账附加费用处理问题
Create Or Replace Procedure Zl_门诊记帐记录_Delete
(
  No_In         门诊费用记录.No%Type,
  序号_In       Varchar2,
  操作员编号_In 门诊费用记录.操作员编号%Type,
  操作员姓名_In 门诊费用记录.操作员姓名%Type
) As
  --功能：冲销一张门诊记帐单据中指定序号行
  --序号：格式如"1,3,5,7,8",为空表示冲销所有可冲销行
  --该光标用于销帐指定费用行

  --该游标为要退费单据的所有原始记录
  Cursor c_Bill(n_标志 Number) Is
    Select a.Id, a.价格父号, a.序号, a.执行状态, a.收费类别, a.医嘱序号, a.病人id, a.收入项目id, a.开单部门id, a.执行部门id, a.病人科室id, a.实收金额,
           Decode(a.记录状态, 0, 1, 0) As 划价, j.诊疗类别, m.跟踪在用
    From 门诊费用记录 A, 病人医嘱记录 J, 材料特性 M
    Where a.医嘱序号 = j.Id(+) And a.收费细目id + 0 = m.材料id(+) And a.No = No_In And a.记录性质 = 2 And a.记录状态 In (0, 1, 3) And
          a.门诊标志 = n_标志
    Order By a.收费细目id, a.序号;

  --该游标用于处理药品库存可用数量
  --不要管费用的执行状态,因为先于此步处理
  Cursor c_Stock(n_标志 Number) Is
    Select ID, 库房id, 药品id, 批次, 批号, 产地, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码, 费用id
    From 药品收发记录
    Where NO = No_In And 单据 In (9, 25) And Mod(记录状态, 3) = 1 And 审核人 Is Null And
          费用id In (Select ID
                   From 门诊费用记录
                   Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = n_标志 And
                         (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
    Order By 药品id;

  --该游标用于处理费用记录序号
  Cursor c_Serial Is
    Select 序号, 价格父号 From 门诊费用记录 Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) Order By 序号;
  l_药品收发 t_Numlist := t_Numlist();
  l_划价     t_Numlist := t_Numlist();
  l_费用id   t_Numlist := t_Numlist();
  n_备货卫材 Number;

  v_医嘱ids Varchar2(4000);

  n_医嘱id   病人医嘱记录.Id%Type;
  n_父号     门诊费用记录.价格父号%Type;
  n_门诊标志 门诊费用记录.门诊标志%Type;

  --部分退费计算变量
  n_剩余数量 Number;
  n_剩余应收 Number;
  n_剩余实收 Number;
  n_剩余统筹 Number;

  n_准退数量 Number;
  n_退费次数 Number;

  n_应收金额 Number;
  n_实收金额 Number;
  n_统筹金额 Number;

  n_Dec Number;

  n_Count   Number;
  d_Curdate Date;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --是否已经全部完全执行(只是整张单据的检查)
  Select Nvl(Count(*), 0), Max(Nvl(门诊标志, 1))
  Into n_Count, n_门诊标志
  From 门诊费用记录
  Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  If Nvl(n_门诊标志, 0) = 0 Then
    n_门诊标志 := 1;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查)
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 门诊费用记录
                Where NO = No_In And 记录性质 = 2 And 门诊标志 = n_门诊标志 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 门诊费用记录
                       Where NO = No_In And 记录性质 = 2 And 门诊标志 = n_门诊标志 And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以销帐的费用！';
    Raise Err_Item;
  End If;

  ---------------------------------------------------------------------------------
  --公用变量
  Select Sysdate Into d_Curdate From Dual;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --循环处理每行费用(收入项目行)
  For r_Bill In c_Bill(n_门诊标志) Loop
    If Instr(',' || 序号_In || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or 序号_In Is Null Then
    
      If r_Bill.划价 = 0 Then
        If Nvl(r_Bill.执行状态, 0) <> 1 Then
          --求剩余数量,剩余应收,剩余实收
          Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
          Into n_剩余数量, n_剩余应收, n_剩余实收, n_剩余统筹
          From 门诊费用记录
          Where NO = No_In And 记录性质 = 2 And 序号 = r_Bill.序号;
        
          If n_剩余数量 = 0 Then
            If 序号_In Is Not Null Then
              v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部销帐！';
              Raise Err_Item;
            End If;
            --情况：未限定行号,原始单据中的该笔已经全部销帐(执行状态=0的一种可能)
          Else
            --准销数量(非药品项目为剩余数量,原始数量)
            If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Or (r_Bill.收费类别 = '4' And Nvl(r_Bill.跟踪在用, 0) = 0) Then
            
              --@@@
              --非药品部分(以具体医嘱执行为准进行检查)
              --: 1.存在医嘱发送的,则以医嘱执行为准(但不能包含:检查;检验;手术;麻醉及输血)
              --: 2.对于病人医吃计价中的收费方式为:0-正常收取 的,才支持部分退;如果是其他的,则只能全退
              --: 3.不存在医嘱的,则以剩余数量为准
              n_Count := 0;
              If Instr(',C,D,F,G,K,', ',' || r_Bill.诊疗类别 || ',') = 0 And r_Bill.诊疗类别 Is Not Null Then
              
                Select Nvl(Sum(数量), 0), Count(*)
                Into n_准退数量, n_Count
                From (Select j.医嘱序号 As 医嘱id, j.收费细目id, Nvl(j.付数, 1) * Nvl(j.数次, 1) As 数量
                       From 门诊费用记录 J, 病人医嘱记录 M
                       Where j.医嘱序号 = m.Id And j.No = No_In And j.记录性质 = 2 And j.序号 = r_Bill.序号 And j.记录状态 In (1, 3) And
                             Exists
                        (Select 1
                              From 病人医嘱发送 A
                              Where a.医嘱id = j.医嘱序号 And Nvl(a.执行状态, 0) <> 1 And a.No || '' = No_In) And Exists
                        (Select 1
                              From 病人医嘱计价 A
                              Where a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) = 0) And j.价格父号 Is Null And
                             Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             (j.记录状态 In (1, 3) And Not Exists
                              (Select 1
                               From 药品收发记录
                               Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) Or
                              j.记录状态 = 2 And Not Exists
                              (Select 1 From 药品收发记录 Where NO = No_In And 单据 In (8, 24) And 药品id = j.收费细目id))
                       Union All
                       Select a.医嘱id, a.收费细目id, -1 * Nvl(a.数量, 1) * Nvl(c.本次数次, 1) As 数量
                       From 病人医嘱计价 A, 病人医嘱发送 B, 病人医嘱执行 C, 门诊费用记录 J, 病人医嘱记录 M
                       Where a.医嘱id = b.医嘱id And b.医嘱id = c.医嘱id And Nvl(a.收费方式, 0) = 0 And b.发送号 = c.发送号 And
                             a.医嘱id = m.Id And Nvl(c.执行结果, 1) = 1 And Nvl(b.执行状态, 0) <> 1 And a.医嘱id = j.医嘱序号 And
                             a.收费细目id = j.收费细目id And j.No = No_In And j.记录性质 = 2 And j.序号 = r_Bill.序号 And
                             j.记录状态 In (1, 3) And j.价格父号 Is Null And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And
                             Not Exists
                        (Select 1
                              From 药品收发记录
                              Where 费用id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0) And Not Exists
                        (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)
                       Union All
                       Select a.医嘱id, a.收费细目id, 0 As 数量
                       From 病人医嘱计价 A, 门诊费用记录 J, 病人医嘱记录 M
                       Where a.医嘱id = m.Id And a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id And Nvl(a.收费方式, 0) <> 0 And
                             j.No = No_In And j.记录性质 = 2 And Nvl(j.执行状态, 0) = 2 And Not Exists
                        (Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1) And
                             Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0);
              
              End If;
            
              If Nvl(n_Count, 0) = 0 Then
                n_准退数量 := n_剩余数量;
              End If;
            
            Else
              Select Sum(Nvl(付数, 1) * 实际数量)
              Into n_准退数量
              From 药品收发记录
              Where NO = No_In And 单据 In (9, 25) And Mod(记录状态, 3) = 1 And 审核人 Is Null And 费用id = r_Bill.Id;
            
              --不跟踪在用的卫生材料
              If r_Bill.收费类别 = '4' And Nvl(n_准退数量, 0) = 0 Then
                n_准退数量 := n_剩余数量;
              End If;
            End If;
          
            --处理门诊费用记录
          
            --该笔项目第几次销帐
            Select Nvl(Max(Abs(执行状态)), 0) + 1
            Into n_退费次数
            From 门诊费用记录
            Where NO = No_In And 记录性质 = 2 And 记录状态 = 2 And 序号 = r_Bill.序号;
          
            --金额=剩余金额*(准退数/剩余数)
            n_应收金额 := Round(n_剩余应收 * (n_准退数量 / n_剩余数量), n_Dec);
            n_实收金额 := Round(n_剩余实收 * (n_准退数量 / n_剩余数量), n_Dec);
            n_统筹金额 := Round(n_剩余统筹 * (n_准退数量 / n_剩余数量), n_Dec);
          
            --插入退费记录
            Insert Into 门诊费用记录
              (ID, NO, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别, 病人科室id, 收费类别,
               收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 划价人,
               执行人, 执行状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 记帐单id, 摘要, 保险编码, 是否急诊, 结论)
              Select 病人费用记录_Id.Nextval, NO, 记录性质, 2, 序号, 从属父号, 价格父号, 病人id, 医嘱序号, 门诊标志, 婴儿费, 姓名, 性别, 年龄, 标识号, 付款方式, 费别,
                     病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                     Decode(Sign(n_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * n_准退数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用,
                     标准单价, -1 * n_应收金额, -1 * n_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * n_退费次数, 执行时间, 操作员编号_In,
                     操作员姓名_In, 发生时间, d_Curdate, 保险项目否, 保险大类id, -1 * n_统筹金额, 记帐单id, 摘要, 保险编码, 是否急诊, 结论
              From 门诊费用记录
              Where ID = r_Bill.Id;
          
            --记录病人医嘱附费对应的医嘱ID(不是主费用)
            If n_医嘱id Is Null And r_Bill.医嘱序号 Is Not Null Then
              n_医嘱id := r_Bill.医嘱序号;
            End If;
          
            --病人余额
            Update 病人余额
            Set 费用余额 = Nvl(费用余额, 0) - n_实收金额
            Where 病人id = r_Bill.病人id And 性质 = 1 And 类型 = 1;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 性质, 类型, 费用余额, 预交余额)
              Values
                (r_Bill.病人id, 1, 1, -1 * n_实收金额, 0);
            End If;
          
            --病人未结费用
            Update 病人未结费用
            Set 金额 = Nvl(金额, 0) - n_实收金额
            Where 病人id = r_Bill.病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And
                  Nvl(病人科室id, 0) = Nvl(r_Bill.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Bill.开单部门id, 0) And
                  Nvl(执行部门id, 0) = Nvl(r_Bill.执行部门id, 0) And 收入项目id + 0 = r_Bill.收入项目id And 来源途径 + 0 = n_门诊标志;
            If Sql%RowCount = 0 Then
              Insert Into 病人未结费用
                (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
              Values
                (r_Bill.病人id, Null, Null, r_Bill.病人科室id, r_Bill.开单部门id, r_Bill.执行部门id, r_Bill.收入项目id, n_门诊标志,
                 -1 * n_实收金额);
            End If;
          
            --标记原费用记录
            --执行状态:全部退完(准退数=剩余数)标记为0,否则标记为1
            Update 门诊费用记录
            Set 记录状态 = 3, 执行状态 = Decode(Sign(n_准退数量 - n_剩余数量), 0, 0, 1)
            Where ID = r_Bill.Id;
          End If;
        Else
          If 序号_In Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能销帐！';
            Raise Err_Item;
          End If;
          --情况:没限定行号,原始单据中包括已经完全执行的
        End If;
      End If;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --药品相关内容
  ------------------------------------------------------------------------------------------------------------------------
  --先处理备货材料
  For v_出库 In (Select ID, 库房id, 药品id, 批次, 批号, 产地, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码, 费用id
               From 药品收发记录
               Where 单据 = 21 And Mod(记录状态, 3) = 1 And 审核人 Is Null And
                     费用id In (Select ID
                              From 门诊费用记录
                              Where NO = No_In And 记录性质 = 2 And 记录状态 In (0, 1, 3) And 收费类别 = '4' And 门诊标志 = n_门诊标志 And
                                    (Instr(',' || 序号_In || ',', ',' || 序号 || ',') > 0 Or 序号_In Is Null))
               Order By 药品id) Loop
    --处理药品库存
    If v_出库.库房id Is Not Null Then
      Update 药品库存
      Set 可用数量 = Nvl(可用数量, 0) + Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0)
      Where 库房id = v_出库.库房id And 药品id = v_出库.药品id And Nvl(批次, 0) = Nvl(v_出库.批次, 0) And 性质 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 药品库存
          (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码)
        Values
          (v_出库.库房id, v_出库.药品id, 1, v_出库.批次, v_出库.效期,
           Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0), v_出库.批号, v_出库.产地, v_出库.灭菌效期,
           v_出库.商品条码, v_出库.内部条码);
      End If;
    End If;
    l_费用id.Extend;
    l_费用id(l_费用id.Count) := v_出库.费用id;
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := v_出库.Id;
  End Loop;

  For r_Stock In c_Stock(n_门诊标志) Loop
  
    --处理药品库存
    If r_Stock.库房id Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_备货卫材
      From Table(l_费用id)
      Where Column_Value = r_Stock.费用id;
      If Nvl(n_备货卫材, 0) = 0 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0)
        Where 库房id = r_Stock.库房id And 药品id = r_Stock.药品id And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码)
          Values
            (r_Stock.库房id, r_Stock.药品id, 1, r_Stock.批次, r_Stock.效期,
             Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0), r_Stock.批号, r_Stock.产地,
             r_Stock.灭菌效期, r_Stock.商品条码, r_Stock.内部条码);
        End If;
      End If;
    End If;
  
    l_药品收发.Extend;
    l_药品收发(l_药品收发.Count) := r_Stock.Id;
  End Loop;

  --删除药品收发记录
  Forall I In 1 .. l_药品收发.Count
    Delete From 药品收发记录 Where ID = l_药品收发(I);

  ------------------------------------------------------------------------------------------------------------------------
  --批量删未发药品记录

  Delete From 未发药品记录 A
  Where NO = No_In And 单据 In (9, 25) And Not Exists
   (Select 1
         From 药品收发记录
         Where 单据 = a.单据 And Nvl(库房id, 0) = Nvl(a.库房id, 0) And NO = No_In And Mod(记录状态, 3) = 1 And 审核人 Is Null);
  ------------------------------------------------------------------------------------------------------------------------

  ---------------------------------------------------------------------------------
  --如果是划价,直接删除费用记录(药品处理后)
  n_Count   := 0;
  v_医嘱ids := Null;
  For r_Bill In c_Bill(n_门诊标志) Loop
    If Instr(',' || 序号_In || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or 序号_In Is Null Then
      If r_Bill.划价 = 1 Then
        If Nvl(r_Bill.执行状态, 0) <> 1 Then
          l_划价.Extend;
          l_划价(l_划价.Count) := r_Bill.Id;
        
          --Delete From 门诊费用记录 Where ID = r_Bill.ID;
          n_Count := n_Count + 1; --记录是否有删除行
        
          If r_Bill.医嘱序号 Is Not Null Then
            If Instr(',' || Nvl(v_医嘱ids, '') || ',', ',' || r_Bill.医嘱序号 || ',') = 0 Then
              v_医嘱ids := Nvl(v_医嘱ids, '') || ',' || r_Bill.医嘱序号;
            End If;
          
            --记录病人医嘱附费对应的医嘱ID(不是主费用)
            If n_医嘱id Is Null Then
              n_医嘱id := r_Bill.医嘱序号;
            End If;
          End If;
        Else
          If 序号_In Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能销帐！';
            Raise Err_Item;
          End If;
          --情况:没限定行号,原始单据中包括已经完全执行的
        End If;
      End If;
    End If;
  End Loop;

  --删除划价记录
  Forall I In 1 .. l_划价.Count
    Delete From 门诊费用记录 Where ID = l_划价(I);

  --删除之后再统一调整序号
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.价格父号 Is Null Then
        n_父号 := n_Count;
      End If;
    
      Update 门诊费用记录
      Set 序号 = n_Count, 价格父号 = Decode(价格父号, Null, Null, n_父号)
      Where NO = No_In And 记录性质 = 2 And 序号 = r_Serial.序号;
    
      Update 门诊费用记录 Set 从属父号 = n_Count Where NO = No_In And 记录性质 = 2 And 从属父号 = r_Serial.序号;
    
      n_Count := n_Count + 1;
    End Loop;
  
  End If;

  --整张单据全部冲完时，删除病人医嘱附费
  For c_医嘱 In (Select Distinct 医嘱序号
               From 门诊费用记录
               Where NO = No_In And 记录性质 = 2 And 记录状态 = 3 And 医嘱序号 Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 门诊费用记录
                  Where 记录性质 = 2 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 2 And NO = No_In;
    End If;
  End Loop;

  If v_医嘱ids Is Not Null Then
    --医嘱处理
    --场合_In    Integer:=0, --0:门诊;1-住院
    --性质_In    Integer:=1, --1-收费单;2-记帐单
    --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
    --No_In      门诊费用记录.No%Type,
    --医嘱ids_In Varchar2 := Null
    v_医嘱ids := Substr(v_医嘱ids, 2);
    Zl_医嘱发送_计费状态_Update(0, 2, 0, No_In, v_医嘱ids);
  Else
    Zl_医嘱发送_计费状态_Update(0, 2, 2, No_In, v_医嘱ids);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊记帐记录_Delete;
/

--108264:刘尔旋,2017-04-17,销账附加费用处理问题
Create Or Replace Procedure Zl_住院记帐记录_Delete
(
  No_In           住院费用记录.No%Type,
  序号_In         Varchar2,
  操作员编号_In   住院费用记录.操作员编号%Type,
  操作员姓名_In   住院费用记录.操作员姓名%Type,
  记录性质_In     住院费用记录.记录性质%Type := 2,
  操作状态_In     Number := 0,
  输液配药检查_In Number := 1,
  登记时间_In     住院费用记录.登记时间%Type := Sysdate
) As
  --功能：冲销一张住院记帐单据中指定序号行
  --序号：格式如"1,3,5,7,8",或"1:2:33456,3:2,5:2,7:2,8:2",冒号前面的数字表示行号,中间的数字表示退的数量,后面的数字表示配药记录的ID,目前仅在销帐审核时才传入
  --      为空表示冲销所有可冲销行
  --记录性质:    2-人工记帐单,3-自动记帐单
  --输液配药检查:    0-医嘱调用，不检查药品是否进入输液配药中心；1-非医嘱调用，检查药品是否进入配药中心
  --该光标用于销帐指定费用行
  --操作状态_In:0-表示直接销帐;1-表示审核销帐(通过销帐申请-->销帐审核流程)
  --该游标为要退费单据的所有原始记录
  Cursor c_Bill Is
    Select ID, 价格父号, 序号, 执行状态, 记录性质, 收费类别, 医嘱序号, 收费细目id, 病人id, 主页id, 收入项目id, 开单部门id, 病人科室id, 执行部门id, 病人病区id, 付数, 数次
    From 住院费用记录
    Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 门诊标志 = 2
    Order By 收费细目id, 序号;

  --该游标用于处理药品库存可用数量
  --不要管费用的执行状态,因为先于此步处理
  Cursor c_Stock(v_序号_In Varchar2) Is
    Select ID, 单据, NO, 库房id, 药品id, 批次, 发药方式, 付数, 实际数量, 灭菌效期, 效期, 产地, 批号, 填制日期, 费用id, 商品条码, 内部条码
    From 药品收发记录
    Where NO = No_In And 单据 In (9, 10, 25, 26) And Mod(记录状态, 3) = 1 And 审核人 Is Null And
          费用id In (Select ID
                   From 住院费用记录
                   Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And
                         门诊标志 = 2 And (Instr(',' || v_序号_In || ',', ',' || 序号 || ',') > 0 Or v_序号_In Is Null))
    Order By 药品id, 填制日期 Desc;

  r_Stock c_Stock%RowType;
  --该游标用于处理费用记录序号
  Cursor c_Serial Is
    Select 序号, 价格父号
    From 住院费用记录
    Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3)
    Order By 序号;

  Cursor Cr_药品 Is
    Select ID, 单据, NO, 库房id, 药品id, 批次, 发药方式, 0 As 数量, 灭菌效期, 效期, 产地, 批号, 填制日期, 费用id
    From 药品收发记录
    Where Rownum <= 1;
  v_药品 Cr_药品%RowType;

  v_医嘱id     病人医嘱记录.Id%Type;
  n_划价       Number;
  v_父号       住院费用记录.价格父号%Type;
  v_序号       Varchar2(2000);
  v_Tmp        Varchar2(4000);
  v_医嘱ids    Varchar2(4000);
  l_药品收发   t_Numlist := t_Numlist();
  l_划价       t_Numlist := t_Numlist();
  l_费用id     t_Numlist := t_Numlist();
  n_付数       Number;
  n_虚拟库房id 药品收发记录.库房id%Type;
  n_其他出库id 药品收发记录.Id%Type;
  n_库房id     药品收发记录.库房id%Type;
  n_返回值     Number;
  --部分退费计算变量
  v_剩余数量 Number;
  v_剩余应收 Number;
  v_剩余实收 Number;
  v_剩余统筹 Number;

  v_准退数量 Number;
  v_退费次数 Number;
  v_应收金额 Number;
  v_实收金额 Number;
  v_统筹金额 Number;
  n_Temp     Number;
  n_部分销帐 Number;
  v_Dec      Number;
  n_Count    Number;
  v_Curdate  Date;
  Err_Item Exception;
  v_Err_Msg        Varchar2(255);
  n_备货卫材       Number;
  n_病人id         病案主页.病人id%Type;
  n_主页id         病案主页.主页id%Type;
  n_审核标志       病案主页.审核标志%Type;
  n_住院状态       病案主页.状态%Type;
  n_病人审核方式   Number(2);
  n_未入科禁止记账 Number(2);
  v_配药id         Varchar2(4000);
  Type Ty_药品 Is Ref Cursor;
  c_药品 Ty_药品; --游标变量

Begin
  --销帐审核时,非药品会传入行号的销帐数量
  If Not 序号_In Is Null Then
    If Instr(序号_In, ':') > 0 Then
      v_Tmp := 序号_In || ',';
      While Not v_Tmp Is Null Loop
        v_序号 := v_序号 || ',' || Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
        If Instr(Substr(v_Tmp, Instr(v_Tmp, ':') + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':') - 1), ':') > 0 Then
          v_配药id := v_配药id || ',' ||
                    Substr(v_Tmp, Instr(v_Tmp, ':', 1, 2) + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':', 1, 2) - 1);
        End If;
        v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End Loop;
      v_序号 := Substr(v_序号, 2);
      If v_配药id Is Not Null Then
        v_配药id := Substr(v_配药id, 2);
      End If;
    Else
      v_序号 := 序号_In;
    End If;
  End If;

  --是否已经全部完全执行(只是整张单据的检查)
  Select Nvl(Count(*), 0), Nvl(Max(病人id), 0), Nvl(Max(主页id), 0)
  Into n_Count, n_病人id, n_主页id
  From 住院费用记录
  Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1 And 门诊标志 = 2;
  If n_Count = 0 Then
    v_Err_Msg := '该单据中的项目已经全部完全执行！';
    Raise Err_Item;
  End If;

  n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0);
  n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0);
  If n_病人审核方式 = 1 Or n_未入科禁止记账 = 1 Then
  
    Begin
      Select 审核标志, 状态 Into n_审核标志, n_住院状态 From 病案主页 Where 病人id = n_病人id And 主页id = n_主页id;
    Exception
      When Others Then
        n_审核标志 := 0;
        n_住院状态 := 0;
    End;
    If n_未入科禁止记账 = 1 And n_住院状态 = 1 Then
      v_Err_Msg := '病人未入科,禁止对病人相关费用的操作!';
      Raise Err_Item;
    End If;
  
    If n_病人审核方式 = 1 Then
    
      If Nvl(n_审核标志, 0) = 1 Then
        v_Err_Msg := '该病人目前正在审核费用,不能进行费用相关调整!';
        Raise Err_Item;
      End If;
      If Nvl(n_审核标志, 0) = 2 Then
        v_Err_Msg := '该病人目前已经完成了费用审核,不能进行费用相关调整!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --未完全执行的项目是否有剩余数量(只是整张单据的检查)
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select 序号, Sum(数量) As 剩余数量
         From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                From 住院费用记录
                Where NO = No_In And 记录性质 = 记录性质_In And 门诊标志 = 2 And
                      Nvl(价格父号, 序号) In
                      (Select Nvl(价格父号, 序号)
                       From 住院费用记录
                       Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And Nvl(执行状态, 0) <> 1)
                Group By 记录状态, Nvl(价格父号, 序号))
         Group By 序号
         Having Sum(数量) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '该单据中未完全执行部分项目剩余数量为零,没有可以销帐的费用！';
    Raise Err_Item;
  End If;

  --医嘱费用：检查正在执行的医嘱(注意已执行的情况在下面检查,因为不传 序号_IN 这种情况费用界面已限制)
  If Nvl(操作状态_In, 0) <> 1 Then
    --走销帐申请流程的，不检查医保执行状态
    Select Nvl(Count(*), 0)
    Into n_Count
    From 病人医嘱发送
    Where 执行状态 = 3 And (NO, 记录性质, 医嘱id) In
          (Select NO, 记录性质, 医嘱序号
                        From 住院费用记录
                        Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 医嘱序号 Is Not Null And
                              (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null));
    If n_Count > 0 Then
      v_Err_Msg := '要销帐的费用中存在对应的医嘱正在执行的情况，不能销帐！';
      Raise Err_Item;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --先打开药品对应数据集,以确保当前条件下有数据,为了处理并发判断
  --不能在游标条件中取消"审核人 is Null"条件，因为多次退药可能部份又已发
  Open c_Stock(v_序号);

  --公用变量
  Select 登记时间_In Into v_Curdate From Dual;

  --金额小数位数
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;

  For c_编目病案 In (Select a.姓名
                 From 病人信息 A, 病案主页 B
                 Where a.病人id = b.病人id And b.编目日期 Is Not Null And
                       (b.病人id, b.主页id) In
                       (Select Distinct 病人id, 主页id
                        From 住院费用记录
                        Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 门诊标志 = 2)) Loop
    v_Err_Msg := '病人『' || c_编目病案.姓名 || '』 已经被病案编目,不能被销帐！';
    Raise Err_Item;
  End Loop;
  v_医嘱ids := Null;
  --循环处理每行费用(收入项目行)
  For r_Bill In c_Bill Loop
    --检查已经存在病案编目的,则不能进行销帐处理
    If Instr(',' || v_序号 || ',', ',' || Nvl(r_Bill.价格父号, r_Bill.序号) || ',') > 0 Or v_序号 Is Null Then
      Select Decode(记录状态, 0, 1, 0) Into n_划价 From 住院费用记录 Where ID = r_Bill.Id;
      If Nvl(r_Bill.执行状态, 0) <> 1 Then
        --求剩余数量,剩余应收,剩余实收
        Select Sum(Nvl(付数, 1) * 数次), Sum(应收金额), Sum(实收金额), Sum(统筹金额)
        Into v_剩余数量, v_剩余应收, v_剩余实收, v_剩余统筹
        From 住院费用记录
        Where NO = No_In And 记录性质 = 记录性质_In And 序号 = r_Bill.序号;
        n_部分销帐 := 0;
        If v_剩余数量 = 0 Then
          If v_序号 Is Not Null Then
            v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经全部销帐！';
            Raise Err_Item;
          End If;
          --情况：未限定行号,原始单据中的该笔已经全部销帐(执行状态=0的一种可能)
        Else
        
          If Instr(序号_In, ':') > 0 Then
            v_Tmp := ',' || 序号_In;
            v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || r_Bill.序号 || ':') + Length(',' || r_Bill.序号 || ':'));
            v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
            If Instr(v_Tmp, ':') > 0 Then
              v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
            End If;
            v_准退数量 := v_Tmp;
            n_部分销帐 := 1;
          End If;
        
          --准销数量(非药品项目为剩余数量,原始数量)
          If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Then
            If Instr(序号_In, ':') = 0 Or 序号_In Is Null Then
              v_准退数量 := v_剩余数量;
            End If;
          Else
            --医嘱超期收回时,卫材可能没有发放,但申请销帐的是部分数量,所以要以申请的为准
            If Instr(序号_In, ':') = 0 Or 序号_In Is Null Then
              Select Nvl(Sum(Nvl(付数, 1) * 实际数量), 0), Count(*)
              Into v_准退数量, n_Count
              From 药品收发记录
              Where NO = No_In And 单据 In (9, 10, 25, 26) And Mod(记录状态, 3) = 1 And 审核人 Is Null And 费用id = r_Bill.Id;
            End If;
          
            --有剩余数量无准退数量的有两种情况：
            --1.不跟踪在用的卫材无对应的收发记录,这时使用剩余数量
            --2.并发操作,此时已发药或发料
            If v_准退数量 = 0 Then
              If r_Bill.收费类别 = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发料,须退料后再退费！';
                  Raise Err_Item;
                Else
                  v_准退数量 := v_剩余数量;
                End If;
              Else
                v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已发药,须退药后再退费！';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          --处理住院费用记录
          If Nvl(n_划价, 0) = 0 Then
            --划价时,直接更改数量,所以不须查划冲销次数
            --该笔项目第几次销帐
            Select Nvl(Max(Abs(执行状态)), 0) + 1
            Into v_退费次数
            From 住院费用记录
            Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 = 2 And 序号 = r_Bill.序号 And 门诊标志 = 2;
          End If;
        
          --金额=剩余金额*(准退数/剩余数)
          v_应收金额 := Round(v_剩余应收 * (v_准退数量 / v_剩余数量), v_Dec);
          v_实收金额 := Round(v_剩余实收 * (v_准退数量 / v_剩余数量), v_Dec);
          v_统筹金额 := Round(v_剩余统筹 * (v_准退数量 / v_剩余数量), v_Dec);
          If Nvl(n_划价, 0) = 1 Then
            If Nvl(n_部分销帐, 0) = 0 Then
              l_划价.Extend;
              l_划价(l_划价.Count) := r_Bill.Id;
              n_返回值 := 0;
            Else
              --更新数量
              --划价的,先将相关的数据处理在内部表集中
              n_付数 := 0;
              If r_Bill.付数 > 1 Then
                --如果是中药,超期回收肯定是回收的付数,而不是次数.因此,需要检查准退数量是否可以整 除
                If Trunc(v_准退数量 / r_Bill.数次) <> (v_准退数量 / r_Bill.数次) Then
                  v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用为中药,请按付数进行退费！';
                  Raise Err_Item;
                End If;
                n_付数 := Trunc(v_准退数量 / r_Bill.数次);
                If Nvl(r_Bill.付数, 0) - n_付数 < 0 Then
                  v_准退数量 := r_Bill.数次;
                Else
                  v_准退数量 := 0;
                End If;
              End If;
              Update 住院费用记录
              Set 付数 = 付数 - n_付数, 数次 = 数次 - v_准退数量, 应收金额 = Nvl(应收金额, 0) - v_应收金额, 实收金额 = Nvl(实收金额, 0) - v_实收金额,
                  登记时间 = v_Curdate, 统筹金额 = Nvl(统筹金额, 0) - v_统筹金额
              Where ID = r_Bill.Id
              Returning Nvl(数次, 0) * Nvl(付数, 0) Into n_返回值;
            End If;
            If Nvl(n_返回值, 0) <= 0 Then
              l_划价.Extend;
              l_划价(l_划价.Count) := r_Bill.Id;
            End If;
            If r_Bill.医嘱序号 Is Not Null Then
              If Instr(',' || Nvl(v_医嘱ids, '') || ',', ',' || r_Bill.医嘱序号 || ',') = 0 Then
                v_医嘱ids := Nvl(v_医嘱ids, '') || ',' || r_Bill.医嘱序号;
              End If;
              --记录病人医嘱附费对应的医嘱ID(不是主费用)
              If v_医嘱id Is Null Then
                v_医嘱id := r_Bill.医嘱序号;
              End If;
            End If;
          
          End If;
        
          If Nvl(n_划价, 0) = 0 Then
            --划价时,直接更改数量,所以不须查划冲销次数
            --插入退费记录
            Insert Into 住院费用记录
              (ID, NO, 记录性质, 记录状态, 序号, 从属父号, 价格父号, 主页id, 病人id, 医嘱序号, 门诊标志, 多病人单, 婴儿费, 姓名, 性别, 年龄, 标识号, 床号, 费别, 病人病区id,
               病人科室id, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人,
               执行部门id, 划价人, 执行人, 执行状态, 执行时间, 操作员编号, 操作员姓名, 发生时间, 登记时间, 保险项目否, 保险大类id, 统筹金额, 保险编码, 记帐单id, 摘要, 费用类型, 是否急诊,
               结论, 医疗小组id)
              Select 病人费用记录_Id.Nextval, NO, 记录性质, 2, 序号, 从属父号, 价格父号, 主页id, 病人id, 医嘱序号, 门诊标志, 多病人单, 婴儿费, 姓名, 性别, 年龄, 标识号,
                     床号, 费别, 病人病区id, 病人科室id, 收费类别, 收费细目id, 计算单位, Decode(Sign(v_准退数量 - Nvl(付数, 1) * 数次), 0, 付数, 1), 发药窗口,
                     Decode(Sign(v_准退数量 - Nvl(付数, 1) * 数次), 0, -1 * 数次, -1 * v_准退数量), 加班标志, 附加标志, 收入项目id, 收据费目, 记帐费用,
                     标准单价, -1 * v_应收金额, -1 * v_实收金额, 开单部门id, 开单人, 执行部门id, 划价人, 执行人, -1 * v_退费次数, 执行时间, 操作员编号_In,
                     操作员姓名_In, 发生时间, v_Curdate, 保险项目否, 保险大类id, -1 * v_统筹金额, 保险编码, 记帐单id, 摘要, 费用类型, 是否急诊, 结论, 医疗小组id
              From 住院费用记录
              Where ID = r_Bill.Id;
          
            --记录病人医嘱附费对应的医嘱ID(不是主费用)
            If v_医嘱id Is Null And r_Bill.医嘱序号 Is Not Null Then
              v_医嘱id := r_Bill.医嘱序号;
            End If;
          
            Update 病人审批项目
            Set 已用数量 = Nvl(已用数量, 0) - v_准退数量
            Where 病人id = r_Bill.病人id And 主页id = r_Bill.主页id And 项目id = r_Bill.收费细目id And Nvl(使用限量, 0) <> 0;
          
            --病人余额
            Update 病人余额
            Set 费用余额 = Nvl(费用余额, 0) - v_实收金额
            Where 病人id = r_Bill.病人id And 类型 = 2 And 性质 = 1;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 类型, 性质, 费用余额, 预交余额)
              Values
                (r_Bill.病人id, 2, 1, -1 * v_实收金额, 0);
            End If;
          
            --病人未结费用
            Update 病人未结费用
            Set 金额 = Nvl(金额, 0) - v_实收金额
            Where 病人id = r_Bill.病人id And Nvl(主页id, 0) = Nvl(r_Bill.主页id, 0) And Nvl(病人病区id, 0) = Nvl(r_Bill.病人病区id, 0) And
                  Nvl(病人科室id, 0) = Nvl(r_Bill.病人科室id, 0) And Nvl(开单部门id, 0) = Nvl(r_Bill.开单部门id, 0) And
                  Nvl(执行部门id, 0) = Nvl(r_Bill.执行部门id, 0) And 收入项目id + 0 = r_Bill.收入项目id And 来源途径 + 0 = 2;
            If Sql%RowCount = 0 Then
              Insert Into 病人未结费用
                (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
              Values
                (r_Bill.病人id, r_Bill.主页id, r_Bill.病人病区id, r_Bill.病人科室id, r_Bill.开单部门id, r_Bill.执行部门id, r_Bill.收入项目id, 2,
                 -1 * v_实收金额);
            End If;
          
            --标记原费用记录
            --执行状态:全部退完(准退数=剩余数)标记为0,否则保持原状态
            If Instr(',4,5,6,7,', r_Bill.收费类别) = 0 Then
              --一般情况非药品和卫材的项目,不存在部分销帐的情况,只有销帐申请和销帐审核时,才会出现部分销帐,所以
              --执行状态只有两种:0.未执行;1已执行;
              --由于在销帐审核过程中将已执行强制改为了2部分执行,因此需要在此处改为1已执行.未执行的不变.
              Update 住院费用记录
              Set 记录状态 = 3, 执行状态 = Decode(Sign(v_准退数量 - v_剩余数量), 0, 0, Decode(执行状态, 2, 1, 执行状态))
              Where ID = r_Bill.Id;
            Else
              Update 住院费用记录
              Set 记录状态 = 3, 执行状态 = Decode(Sign(v_准退数量 - v_剩余数量), 0, 0, 执行状态)
              Where ID = r_Bill.Id;
            End If;
          End If;
        End If;
      Else
        If v_序号 Is Not Null Then
          v_Err_Msg := '单据中第' || Nvl(r_Bill.价格父号, r_Bill.序号) || '行费用已经完全执行,不能销帐！';
          Raise Err_Item;
        End If;
        --情况:没限定行号,原始单据中包括已经完全执行的
      End If;
    End If;
  End Loop;

  --不存在配药ID,检查该药品是否在输液配药中心
  If v_配药id Is Null And 输液配药检查_In = 1 Then
    For v_费用 In (Select ID
                 From 住院费用记录
                 Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = 2 And
                       (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null)) Loop
      Begin
        Select Count(1)
        Into n_Count
        From 输液配药内容 A, 药品收发记录 B
        Where a.收发id = b.Id And b.费用id = v_费用.Id And Instr(',8,9,10,21,24,25,26,', ',' || b.单据 || ',') > 0;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '存在已经进入输液配药中心的待销帐药品，无法完成销帐！';
        Raise Err_Item;
      End If;
    End Loop;
  End If;

  n_部分销帐 := 0;
  ---------------------------------------------------------------------------------
  --药品相关处理:主要是对销帐审核有效.(可以是部分)
  For v_费用 In (Select ID, 序号, 收费类别
               From 住院费用记录
               Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 In ('4', '5', '6', '7') And 门诊标志 = 2 And
                     (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null)
               Order By 收费细目id) Loop
    --根据费用ID来进行相关的处理
    v_准退数量 := 0;
    If Instr(序号_In, ':') > 0 Then
      v_Tmp := ',' || 序号_In;
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || v_费用.序号 || ':') + Length(',' || v_费用.序号 || ':'));
      v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
      If Instr(v_Tmp, ':') > 0 Then
        v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
      End If;
      v_准退数量 := v_Tmp;
    End If;
    If v_准退数量 <> 0 Then
      n_部分销帐 := 1;
      n_Temp     := 0;
      --------------------------------------------------------------------------------------
      --检查是否备货记帐卫材,规则如下
      -- a.如果存在存在未审核的其他出库且部分销帐时,直接在原来的基础上更改其他出库数量
      -- b.如果存在存在未审核的其他出库且完全销帐时,直接删除
      -- c.库存处理:还原为虚拟库房的可用数量;发料部门不处理
      -- d.如果已经发了料,这个时间由于其他出库单已经审核,因此就按正常情况流转,库存恢复到发料部门中
      n_虚拟库房id := Null;
      n_其他出库id := Null;
      If v_费用.收费类别 = '4' Then
        Begin
          Select 1, 库房id, ID
          Into n_备货卫材, n_虚拟库房id, n_其他出库id
          From 药品收发记录
          Where 费用id = v_费用.Id And 审核日期 Is Null And 单据 = 21 And Rownum = 1;
        Exception
          When Others Then
            n_备货卫材 := 0;
        End;
      Else
        n_备货卫材 := 0;
      End If;
      --------------------------------------------------------------------------------------
      If v_配药id Is Not Null Then
        Open c_药品 For
          Select /*+ rule*/
           a.Id, a.单据, a.No, a.库房id, a.药品id, a.批次, a.发药方式,
           Decode(a.发药方式, Null, 1, -1, 0, 1) * Nvl(a.付数, 1) * Nvl(a.实际数量, 0) As 数量, a.灭菌效期, a.效期, a.产地, a.批号, a.填制日期,
           a.费用id
          From 药品收发记录 A, Table(f_Str2list(v_配药id)) B, 输液配药内容 C
          Where a.No = No_In And a.单据 In (9, 10, 25, 26) And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null And a.费用id = v_费用.Id And
                a.Id = c.收发id And c.记录id = b.Column_Value
          Order By 填制日期;
      Else
        Open c_药品 For
          Select ID, 单据, NO, 库房id, 药品id, 批次, 发药方式, Decode(发药方式, Null, 1, -1, 0, 1) * Nvl(付数, 1) * Nvl(实际数量, 0) As 数量,
                 灭菌效期, 效期, 产地, 批号, 填制日期, 费用id
          From 药品收发记录
          Where NO = No_In And 单据 In (9, 10, 25, 26) And Mod(记录状态, 3) = 1 And 审核人 Is Null And 费用id = v_费用.Id
          Order By 填制日期;
      End If;
      Loop
        Fetch c_药品
          Into v_药品;
        Exit When c_药品%NotFound;
        n_Temp := v_药品.数量;
        If v_准退数量 >= n_Temp Then
          l_药品收发.Extend;
          l_药品收发(l_药品收发.Count) := v_药品.Id;
          If Nvl(n_其他出库id, 0) > 0 Then
            l_药品收发.Extend;
            l_药品收发(l_药品收发.Count) := n_其他出库id;
          End If;
          v_准退数量 := v_准退数量 - n_Temp;
        Else
          If v_费用.收费类别 = '7' Then
            --当前行的数量要大
            Update 药品收发记录
            Set 付数 = 1, 实际数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量,
                填写数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(填写数量, 0) - v_准退数量,
                成本金额 =
                 (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 成本价,
                零售金额 =
                 (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 零售价,
                差价 = Round((Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 零售价 -
                            (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 成本价, 5)
            Where ID = v_药品.Id;
          Else
            Update 药品收发记录
            Set 实际数量 = Nvl(实际数量, 0) - v_准退数量, 填写数量 = Nvl(填写数量, 0) - v_准退数量,
                成本金额 =
                 (Nvl(实际数量, 0) - v_准退数量) * 成本价,
                零售金额 =
                 (Nvl(实际数量, 0) - v_准退数量) * 零售价,
                差价 = Round((Nvl(实际数量, 0) - v_准退数量) * 零售价 - (Nvl(实际数量, 0) - v_准退数量) * 成本价, 5)
            Where ID = v_药品.Id;
          End If;
          --更新其他出库单
          If Nvl(n_其他出库id, 0) <> 0 Then
            If v_费用.收费类别 = '7' Then
              Update 药品收发记录
              Set 付数 = 1, 实际数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量,
                  填写数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量,
                  成本金额 =
                   (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 成本价,
                  零售金额 =
                   (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 零售价,
                  差价 = Round((Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 零售价 -
                              (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - v_准退数量) * 成本价, 5)
              Where ID = Nvl(n_其他出库id, 0);
            Else
              Update 药品收发记录
              Set 实际数量 = Nvl(实际数量, 0) - v_准退数量, 填写数量 = Nvl(实际数量, 0) - v_准退数量,
                  成本金额 =
                   (Nvl(实际数量, 0) - v_准退数量) * 成本价,
                  零售金额 =
                   (Nvl(实际数量, 0) - v_准退数量) * 零售价,
                  差价 = Round((Nvl(实际数量, 0) - v_准退数量) * 零售价 - (Nvl(实际数量, 0) - v_准退数量) * 成本价, 5)
              Where ID = Nvl(n_其他出库id, 0);
            End If;
          End If;
          n_Temp     := v_准退数量;
          v_准退数量 := 0;
        End If;
        If Nvl(n_备货卫材, 0) = 1 Then
          n_库房id := n_虚拟库房id;
        Else
          n_库房id := v_药品.库房id;
        End If;
      
        If n_库房id Is Not Null Then
          Update 药品库存
          Set 可用数量 = Nvl(可用数量, 0) + n_Temp
          Where 库房id = n_库房id And 药品id = v_药品.药品id And Nvl(批次, 0) = Nvl(v_药品.批次, 0) And 性质 = 1;
          If Sql%RowCount = 0 Then
            Insert Into 药品库存
              (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期)
            Values
              (n_库房id, v_药品.药品id, 1, v_药品.批次, v_药品.效期, n_Temp, v_药品.批号, v_药品.产地, v_药品.灭菌效期);
          End If;
          Delete 药品库存
          Where 库房id = n_库房id And 药品id = v_药品.药品id And Nvl(批次, 0) = Nvl(v_药品.批次, 0) And 性质 = 1 And Nvl(可用数量, 0) = 0 And
                Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
        End If;
      
        If Nvl(n_备货卫材, 0) = 1 Then
          Update 药品库存
          Set 可用数量 = Nvl(可用数量, 0) + n_Temp
          Where 库房id = v_药品.库房id And 药品id = v_药品.药品id And Nvl(批次, 0) = Nvl(v_药品.批次, 0) And 性质 = 1;
          If Sql%RowCount = 0 Then
            Insert Into 药品库存
              (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期)
            Values
              (v_药品.库房id, v_药品.药品id, 1, v_药品.批次, v_药品.效期, n_Temp, v_药品.批号, v_药品.产地, v_药品.灭菌效期);
          End If;
        End If;
      
        If v_准退数量 = 0 Then
          Exit;
        End If;
      End Loop;
      --不跟踪卫材的,不检查:因为不跟噻的话,不会在药品收发记录中存在
      If Nvl(v_准退数量, 0) <> 0 And Not (v_费用.收费类别 = '4' And n_Temp = 0) Then
        --未分配完成,表示此药品可能已经执行.
        v_Err_Msg := '要销帐的费用中存在已发的药品或卫材，或已被其他人销帐；这可能是并发操作引起的。';
        Raise Err_Item;
      End If;
    End If;
  End Loop;

  If n_部分销帐 = 0 Then
    ------------------------------------------------------------------------------------------------------------------------
    --先处理备货材料
    For v_出库 In (Select ID, 单据, NO, 库房id, 药品id, 批次, 发药方式, 付数, 实际数量, 灭菌效期, 效期, 产地, 批号, 填制日期, 费用id, 商品条码, 内部条码
                 From 药品收发记录
                 Where 单据 = 21 And Mod(记录状态, 3) = 1 And 审核人 Is Null And
                       费用id In (Select ID
                                From 住院费用记录
                                Where NO = No_In And 记录性质 = 记录性质_In And 记录状态 In (0, 1, 3) And 收费类别 = '4' And 门诊标志 = 2 And
                                      (Instr(',' || v_序号 || ',', ',' || 序号 || ',') > 0 Or v_序号 Is Null))
                 Order By 药品id, 填制日期 Desc) Loop
      --处理药品库存
      If v_出库.库房id Is Not Null Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0)
        Where 库房id = v_出库.库房id And 药品id = v_出库.药品id And Nvl(批次, 0) = Nvl(v_出库.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码)
          Values
            (v_出库.库房id, v_出库.药品id, 1, v_出库.批次, v_出库.效期,
             Decode(v_出库.发药方式, Null, 1, -1, 0, 1) * Nvl(v_出库.付数, 1) * Nvl(v_出库.实际数量, 0), v_出库.批号, v_出库.产地, v_出库.灭菌效期,
             v_出库.商品条码, v_出库.内部条码);
        End If;
        Delete 药品库存
        Where 库房id = v_出库.库房id And 药品id = v_出库.药品id And Nvl(批次, 0) = Nvl(v_出库.批次, 0) And 性质 = 1 And Nvl(可用数量, 0) = 0 And
              Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
      End If;
      l_费用id.Extend;
      l_费用id(l_费用id.Count) := v_出库.费用id;
      l_药品收发.Extend;
      l_药品收发(l_药品收发.Count) := v_出库.Id;
    End Loop;
  
    --药品相关内容
    Fetch c_Stock
      Into r_Stock;
    While c_Stock%Found Loop
    
      --处理药品库存
      If r_Stock.库房id Is Not Null Then
      
        Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
        Into n_备货卫材
        From Table(l_费用id)
        Where Column_Value = r_Stock.费用id;
      
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0)
        Where 库房id = r_Stock.库房id And 药品id = r_Stock.药品id And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期)
          Values
            (r_Stock.库房id, r_Stock.药品id, 1, r_Stock.批次, r_Stock.效期,
             Decode(r_Stock.发药方式, Null, 1, -1, 0, 1) * Nvl(r_Stock.付数, 1) * Nvl(r_Stock.实际数量, 0), r_Stock.批号, r_Stock.产地,
             r_Stock.灭菌效期);
        End If;
        Delete 药品库存
        Where 库房id = r_Stock.库房id And 药品id = r_Stock.药品id And Nvl(批次, 0) = Nvl(r_Stock.批次, 0) And 性质 = 1 And
              Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
      End If;
    
      --删除药品收发记录(加上并发操作检查:审核人 Is Null)
      --Delete From 药品收发记录 Where ID = r_Stock.ID And 审核人 Is Null;
    
      l_药品收发.Extend;
      l_药品收发(l_药品收发.Count) := r_Stock.Id;
      Fetch c_Stock
        Into r_Stock;
    End Loop;
    Close c_Stock;
  
    --删除药品收发记录
    Forall I In 1 .. l_药品收发.Count
      Delete From 药品收发记录 Where ID = l_药品收发(I) And 审核人 Is Null;
    If Sql%RowCount <> l_药品收发.Count And l_药品收发.Count <> 0 Then
      v_Err_Msg := '要销帐的费用中存在已发的药品或卫材，或已被其他人销帐；这可能是并发操作引起的。';
      Raise Err_Item;
    End If;
  Else
    --删除药品收发记录
    Forall I In 1 .. l_药品收发.Count
      Delete From 药品收发记录 Where ID = l_药品收发(I) And 审核人 Is Null;
  End If;
  --未发药品记录
  Delete From 未发药品记录 A
  Where NO = No_In And 单据 In (9, 10, 25, 26) And Not Exists
   (Select 1
         From 药品收发记录
         Where 单据 = a.单据 And Nvl(库房id, 0) = Nvl(a.库房id, 0) And NO = No_In And Mod(记录状态, 3) = 1 And 审核人 Is Null);

  ---------------------------------------------------------------------------------
  --如果是划价,直接删除费用记录(药品处理后)
  n_Count := l_划价.Count;
  --删除划价记录
  Forall I In 1 .. l_划价.Count
    Delete From 住院费用记录 Where ID = l_划价(I);

  --删除之后再统一调整序号
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.价格父号 Is Null Then
        v_父号 := n_Count;
      End If;
    
      Update 住院费用记录
      Set 序号 = n_Count, 价格父号 = Decode(价格父号, Null, Null, v_父号)
      Where NO = No_In And 记录性质 = 记录性质_In And 序号 = r_Serial.序号;
    
      Update 住院费用记录
      Set 从属父号 = n_Count
      Where NO = No_In And 记录性质 = 记录性质_In And 从属父号 = r_Serial.序号;
    
      n_Count := n_Count + 1;
    End Loop;
  
  End If;

  --整张单据全部冲完时，删除病人医嘱附费
  For c_医嘱 In (Select Distinct 医嘱序号
               From 住院费用记录
               Where NO = No_In And 记录性质 = 2 And 记录状态 = 3 And 医嘱序号 Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select 序号, Sum(数量) As 剩余数量
           From (Select 记录状态, Nvl(价格父号, 序号) As 序号, Avg(Nvl(付数, 1) * 数次) As 数量
                  From 住院费用记录
                  Where 记录性质 = 2 And 医嘱序号 + 0 = c_医嘱.医嘱序号 And NO = No_In
                  Group By 记录状态, Nvl(价格父号, 序号))
           Group By 序号
           Having Sum(数量) <> 0);
  
    If n_Count = 0 Then
      Delete From 病人医嘱附费 Where 医嘱id = c_医嘱.医嘱序号 And 记录性质 = 2 And NO = No_In;
    End If;
  End Loop;

  If v_医嘱ids Is Not Null Then
    --医嘱处理
    --场合_In    Integer:=0, --0:门诊;1-住院
    --性质_In    Integer:=1, --1-收费单;2-记帐单
    --操作_In    Integer:=0, --0:删除划价单;1-收费或记帐;2-退费或销帐
    --No_In      门诊费用记录.No%Type,
    --医嘱ids_In Varchar2 := Null
    v_医嘱ids := Substr(v_医嘱ids, 2);
    Zl_医嘱发送_计费状态_Update(1, 2, 0, No_In, v_医嘱ids);
  Else
    Zl_医嘱发送_计费状态_Update(1, 2, 2, No_In);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_住院记帐记录_Delete;
/

--107321:冉俊明,2017-04-14,修改无“所有科室”权限时的功能控制问题
Create Or Replace Procedure Zl_临床出诊表_Delete
(
  Id_In     临床出诊表.Id%Type,
  人员id_In 人员表.Id%Type := Null,
  站点_In   部门表.站点%Type
) As
  --功能：删除临床出诊表 
  --参数： 
  --        人员id_In 除固定安排外有效，不为0或null表示临床科室人员在删除 
  n_Count    Number;
  n_排班方式 临床出诊表.排班方式%Type;
  n_出诊id   临床出诊表.Id%Type;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  l_记录id t_Numlist := t_Numlist();
  l_限制id t_Numlist := t_Numlist();
Begin
  Begin
    Select 1 Into n_Count From 临床出诊表 Where 排班方式 <> 3 And 发布人 Is Not Null And ID = Id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '当前出诊表已发布，不能删除！';
    Raise Err_Item;
  End If;

  Begin
    Select 排班方式 Into n_排班方式 From 临床出诊表 Where ID = Id_In;
  Exception
    When Others Then
      v_Err_Msg := '出诊表信息未找到！';
      Raise Err_Item;
  End;

  --按天排班的月模板数据保存在出诊记录中的 
  If Nvl(n_排班方式, 0) In (0, 3) Then
    --固定安排/模板 
    --删除临床出诊限制 
    Select b.Id Bulk Collect
    Into l_限制id
    From 临床出诊安排 A, 临床出诊限制 B
    Where a.Id = b.安排id And a.出诊id = Id_In;
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊诊室 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊挂号控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊限制 Where ID = l_限制id(I);
  
    --删除临床出诊安排 
    Delete From 临床出诊安排 Where 出诊id = Id_In;
  
    --删除临床出诊表 
    Delete 临床出诊表 Where ID = Id_In;
  
    Return;
  End If;

  --======================================================================================================== 
  --月出诊表/周出诊表 
  --月出诊表/周出诊表只能从最后一个开始删除 
  Begin
    Select ID
    Into n_出诊id
    From (Select a.Id
           From 临床出诊表 A, 临床出诊安排 B, 临床出诊号源 C, 部门表 D
           Where a.排班方式 = n_排班方式 And a.Id = b.出诊id(+) And b.号源id = c.Id(+) And c.科室id = d.Id(+)
                --当前人员可操作的号源 
                 And (Nvl(人员id_In, 0) = 0 Or (Nvl(c.是否临床排班, 0) = 1 And Exists
                  (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
                --站点 
                 And (d.站点 Is Null Or d.站点 = 站点_In)
           Order By a.年份 Desc, a.月份 Desc, a.周数 Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      n_出诊id := 0;
  End;
  If Nvl(n_出诊id, 0) <> 0 And Nvl(n_出诊id, 0) <> Id_In Then
    v_Err_Msg := '必须从最后一个出诊表开始删除！';
    Raise Err_Item;
  End If;

  If Nvl(人员id_In, 0) <> 0 Then
    --没有"所有科室"权限
    Select Count(1)
    Into n_Count
    From 临床出诊安排 A, 临床出诊号源 B
    Where a.号源id = b.Id And a.出诊id = Id_In And
          Not (Nvl(b.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = b.科室id And 人员id = 人员id_In)) And
          Rownum < 2;
    If n_Count <> 0 Then
      v_Err_Msg := '当前出诊表中含有其它人员已经制定的安排，不能删除！';
      Raise Err_Item;
    End If;
  End If;

  --删除临床出诊记录 
  Select a.Id Bulk Collect
  Into l_记录id
  From 临床出诊记录 A, 临床出诊安排 B, 临床出诊号源 C, 部门表 D
  Where a.安排id = b.Id And a.号源id = c.Id And c.科室id = d.Id And b.出诊id = Id_In
       --当前人员可操作的号源 
        And (Nvl(人员id_In, 0) = 0 Or
        (Nvl(c.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where c.科室id = 部门id And 人员id = 人员id_In)))
       --站点 
        And (d.站点 Is Null Or d.站点 = 站点_In);

  Zl_临床出诊记录_Batchdelete(l_记录id);

  --删除临床出诊安排 
  Delete From 临床出诊安排 A
  Where a.出诊id = Id_In And Exists
   (Select 1
         From 临床出诊号源 B, 部门表 D
         Where a.号源id = b.Id And b.科室id = d.Id
              --当前人员可操作的号源 
               And (Nvl(人员id_In, 0) = 0 Or (Nvl(b.是否临床排班, 0) = 1 And Exists
                (Select 1 From 部门人员 Where b.科室id = 部门id And 人员id = 人员id_In)))
              --站点 
               And (d.站点 Is Null Or d.站点 = 站点_In));

  --删除临床出诊表 
  Delete 临床出诊表 A
  Where a.Id = Id_In And Not Exists (Select 1 From 临床出诊安排 Where 出诊id = a.Id And 号源id Is Not Null);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊表_Delete;
/

--109289:冉俊明,2017-05-23,使用“全部启用序号控制”功能时，对于启用序号但没有启用分时段的安排，没有生成对应的时段序号数据
--107321:冉俊明,2017-04-14,修改无“所有科室”权限时的功能控制问题
Create Or Replace Procedure Zl_临床出诊安排_序号控制
(
  出诊id_In   临床出诊表.Id%Type,
  序号控制_In 临床出诊限制.是否序号控制%Type,
  站点_In     部门表.站点%Type := Null,
  人员id_In   人员表.Id%Type := 0
) As
  --全部启用序号控制或者全部取消序号控制
  --参数：
  --      人员id_In 不等于0则修改人员所在科室的所有号源安排，否则修改所有号源的安排
  n_Count    Number(2);
  n_出诊记录 Number(2);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_安排id t_Numlist := t_Numlist();
  l_记录id t_Numlist := t_Numlist();

  --该游标用于读取所有临床出诊安排的ID
  Cursor c_安排
  (
    出诊id_In 临床出诊表.Id%Type,
    人员id_In 人员表.Id%Type := 0
  ) Is
    Select b.Id
    From 临床出诊安排 B, 临床出诊号源 C
    Where b.号源id = c.Id And b.出诊id = 出诊id_In And
          (Nvl(人员id_In, 0) = 0 Or (Nvl(人员id_In, 0) <> 0 And Nvl(c.是否临床排班, 0) = 1 And Exists
           (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In))) And Exists
     (Select 1 From 部门表 Where ID = c.科室id And (站点_In Is Null Or (站点 Is Null Or 站点 = 站点_In)));
Begin
  Select Count(1)
  Into n_Count
  From 临床出诊表 A
  Where a.Id = 出诊id_In And a.发布人 Is Not Null And a.排班方式 <> 3 And Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '当前出诊表已发布，不允许修改！';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From 临床出诊表 A Where a.Id = 出诊id_In And a.排班方式 In (1, 2) And Rownum < 2;
  If n_Count <> 0 Then
    n_出诊记录 := 1;
  End If;

  Open c_安排(出诊id_In, 人员id_In);
  Fetch c_安排 Bulk Collect
    Into l_安排id;
  Close c_安排;

  If Nvl(n_出诊记录, 0) = 0 Then
    --临床出诊限制或模板
    Forall I In 1 .. l_安排id.Count
      Update 临床出诊限制
      Set 是否序号控制 = 序号控制_In
      Where (限号数 Is Not Null Or 限约数 Is Not Null) And 安排id = l_安排id(I);
  
    If Nvl(序号控制_In, 0) = 0 Then
      --取消序号控制，删除序号数据
      Select /*+cardinality(b,10)*/
       ID Bulk Collect
      Into l_记录id
      From 临床出诊限制 A, Table(l_安排id) B
      Where a.安排id = b.Column_Value And (a.限号数 Is Not Null Or a.限约数 Is Not Null) And Nvl(a.是否序号控制, 0) = 0 And
            Nvl(a.是否分时段, 0) = 0;
    
      Forall I In 1 .. l_记录id.Count
        Delete From 临床出诊时段 Where 限制id = l_记录id(I);
    Else
      --不分时段的序号控制号先生成序号,开始时间、终止时间填写时间段的开始时间和结束时间
      For c_安排 In (Select /*+cardinality(d,10)*/
                    a.Id, b.号类, c.站点
                   From 临床出诊安排 A, 临床出诊号源 B, 部门表 C, Table(l_安排id) D
                   Where a.号源id = b.Id And b.科室id = c.Id And a.Id = d.Column_Value) Loop
      
        For c_记录 In (With c_时间段 As
                        (Select 时间段, 开始时间, 终止时间
                        From (Select 时间段,
                                      To_Date('3000-01-01' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                      To_Date('3000-01-01' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 终止时间,
                                      Row_Number() Over(Partition By 时间段 Order By 时间段, 站点 Asc, 号类 Asc) As 组号
                               From 时间段
                               Where Nvl(站点, c_安排.站点) = c_安排.站点 And Nvl(号类, c_安排.号类) = c_安排.号类)
                        Where 组号 = 1)
                       Select a.Id, a.限号数,
                              To_Date(To_Char(Sysdate, 'yyyy-mm-dd ') || To_Char(b.开始时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                              To_Date(To_Char(Sysdate, 'yyyy-mm-dd ') || To_Char(b.终止时间, 'hh24:mi:ss'),
                                      'yyyy-mm-dd hh24:mi:ss') + Case
                                When b.终止时间 <= b.开始时间 Then
                                 1
                                Else
                                 0
                              End As 终止时间
                       From 临床出诊限制 A, c_时间段 B
                       Where a.上班时段 = b.时间段 And 安排id = c_安排.Id And Nvl(限号数, 0) <> 0 And Nvl(是否序号控制, 0) = 1 And
                             Nvl(是否分时段, 0) = 0 And Not Exists (Select 1 From 临床出诊时段 Where 限制id = a.Id)) Loop
        
          For I In 1 .. c_记录.限号数 Loop
            Insert Into 临床出诊时段
              (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
            Values
              (c_记录.Id, I, c_记录.开始时间, c_记录.终止时间, 1, 1);
          End Loop;
        End Loop;
      End Loop;
    End If;
  Else
    --临床出诊记录
    Forall I In 1 .. l_安排id.Count
      Update 临床出诊记录
      Set 是否序号控制 = 序号控制_In
      Where (限号数 Is Not Null Or 限约数 Is Not Null) And 安排id = l_安排id(I);
  
    If Nvl(序号控制_In, 0) = 0 Then
      --取消序号控制，删除序号数据
      Select /*+cardinality(b,10)*/
       a.Id Bulk Collect
      Into l_记录id
      From 临床出诊记录 A, Table(l_安排id) B
      Where a.安排id = b.Column_Value And Nvl(a.限号数, 0) <> 0 And Nvl(a.是否序号控制, 0) = 0 And Nvl(a.是否分时段, 0) = 0;
    
      Forall I In 1 .. l_记录id.Count
        Delete From 临床出诊序号控制 Where 记录id = l_记录id(I);
    Else
      --不分时段的序号控制号先生成序号,开始时间、终止时间填写时间段的开始时间和结束时间
      For c_记录 In (Select /*+cardinality(b,10)*/
                    a.Id, a.限号数, a.开始时间, a.终止时间
                   From 临床出诊记录 A, Table(l_安排id) B
                   Where a.安排id = b.Column_Value And Nvl(a.限号数, 0) <> 0 And Nvl(a.是否序号控制, 0) = 1 And Nvl(a.是否分时段, 0) = 0 And
                         Not Exists (Select 1 From 临床出诊序号控制 Where 记录id = a.Id)) Loop
      
        For I In 1 .. c_记录.限号数 Loop
          Insert Into 临床出诊序号控制
            (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约)
          Values
            (c_记录.Id, I, c_记录.开始时间, c_记录.终止时间, 1, 1);
        End Loop;
      End Loop;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊安排_序号控制;
/

--107321:冉俊明,2017-04-14,修改无“所有科室”权限时的功能控制问题
Create Or Replace Procedure Zl_临床出诊安排_Batchdelete
(
  出诊id_In   临床出诊表.Id%Type,
  人员id_In   人员表.Id%Type := 0,
  站点_In     部门表.站点%Type := Null,
  号源id_In   临床出诊安排.号源id%Type := 0,
  安排id_In   临床出诊安排.Id%Type := 0,
  临时安排_In 临床出诊安排.是否临时安排%Type := 0
) As
  --功能：批量删除临床出诊安排 
  --参数： 
  --      人员id_In 不等于0则删除人员所在科室的所有号源安排 
  --      号源id_In 不等于0则删除该号源的所有安排 
  --      安排ID_in 不等于0则删除该号源的当前安排(一般是临时安排) 
  --说明：如果人员id_In=0且号源id_In=0 则删除该出诊表的所有号源的所有安排 
  n_Count    Number(8);
  n_出诊记录 Number(1);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_限制id t_Numlist := t_Numlist();
  l_记录id t_Numlist := t_Numlist();
Begin
  If Nvl(临时安排_In, 0) = 0 Then
    Begin
      Select 1
      Into n_Count
      From 临床出诊表 A
      Where a.Id = 出诊id_In And a.发布人 Is Not Null And a.排班方式 <> 3 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count <> 0 Then
      v_Err_Msg := '当前出诊表已发布，不允许修改安排！';
      Raise Err_Item;
    End If;
  End If;

  Begin
    Select 1 Into n_出诊记录 From 临床出诊表 A Where a.Id = 出诊id_In And a.排班方式 In (1, 2) And Rownum < 2;
  Exception
    When Others Then
      n_出诊记录 := 0;
  End;

  If Nvl(n_出诊记录, 0) = 0 Then
    --删除临床出诊规则/模板 
    Select a.Id Bulk Collect
    Into l_限制id
    From 临床出诊限制 A, 临床出诊安排 B, 临床出诊号源 C, 部门表 D
    Where a.安排id = b.Id And b.号源id = c.Id And c.科室id = d.Id And b.出诊id = 出诊id_In And
          (
          --删除该出诊表的所有号源的所有安排 
           (Nvl(号源id_In, 0) = 0 And Nvl(人员id_In, 0) = 0)
          --删除该号源的所有安排 
           Or (Nvl(号源id_In, 0) <> 0 And b.号源id = 号源id_In And Nvl(安排id_In, 0) = 0)
          --删除该号源的选择安排 
           Or (Nvl(安排id_In, 0) <> 0 And b.Id = 安排id_In)
          --删除人员所在科室的所有号源安排 
           Or (Nvl(人员id_In, 0) <> 0 And Nvl(c.是否临床排班, 0) = 1 And Exists
            (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
         --站点 
          And (d.站点 Is Null Or d.站点 = 站点_In);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊挂号控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊诊室 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete From 临床出诊限制 Where ID = l_限制id(I);
  
    --删除临床出诊安排 
    For c_安排 In (Select b.Id
                 From 临床出诊安排 B, 临床出诊号源 C, 部门表 D
                 Where b.号源id = c.Id And c.科室id = d.Id And b.出诊id = 出诊id_In And
                       (
                       --删除该出诊表的所有号源的所有安排 
                        (Nvl(号源id_In, 0) = 0 And Nvl(人员id_In, 0) = 0)
                       --删除该号源的所有安排 
                        Or (Nvl(号源id_In, 0) <> 0 And b.号源id = 号源id_In And Nvl(安排id_In, 0) = 0)
                       --删除该号源的选择安排 
                        Or (Nvl(安排id_In, 0) <> 0 And b.Id = 安排id_In)
                       --删除人员所在科室的所有号源安排 
                        Or (Nvl(人员id_In, 0) <> 0 And Nvl(c.是否临床排班, 0) = 1 And Exists
                         (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
                      --站点 
                       And (d.站点 Is Null Or d.站点 = 站点_In) And Not Exists
                  (Select 1 From 临床出诊限制 Where 安排id = b.Id)) Loop
      Zl_临床出诊安排_Delete(c_安排.Id);
    End Loop;
  Else
    --删除临床出诊记录 
    Select a.Id Bulk Collect
    Into l_记录id
    From 临床出诊记录 A, 临床出诊安排 B, 临床出诊号源 C, 部门表 D
    Where a.安排id = b.Id And b.号源id = c.Id And c.科室id = d.Id And b.出诊id = 出诊id_In And
          (
          --删除该出诊表的所有号源的所有安排 
           (Nvl(号源id_In, 0) = 0 And Nvl(人员id_In, 0) = 0)
          --删除该号源的所有安排 
           Or (Nvl(号源id_In, 0) <> 0 And b.号源id = 号源id_In And Nvl(安排id_In, 0) = 0)
          --删除该号源的选择安排 
           Or (Nvl(安排id_In, 0) <> 0 And b.Id = 安排id_In)
          --删除人员所在科室的所有号源安排 
           Or (Nvl(人员id_In, 0) <> 0 And Nvl(c.是否临床排班, 0) = 1 And Exists
            (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
         --站点 
          And (d.站点 Is Null Or d.站点 = 站点_In);
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
  
    --删除临床出诊安排 
    For c_安排 In (Select b.Id
                 From 临床出诊安排 B, 临床出诊号源 C, 部门表 D
                 Where b.号源id = c.Id And c.科室id = d.Id And b.出诊id = 出诊id_In And
                       (
                       --删除该出诊表的所有号源的所有安排 
                        (Nvl(号源id_In, 0) = 0 And Nvl(人员id_In, 0) = 0)
                       --删除该号源的所有安排 
                        Or (Nvl(号源id_In, 0) <> 0 And b.号源id = 号源id_In And Nvl(安排id_In, 0) = 0)
                       --删除该号源的选择安排 
                        Or (Nvl(安排id_In, 0) <> 0 And b.Id = 安排id_In)
                       --删除人员所在科室的所有号源安排 
                        Or (Nvl(人员id_In, 0) <> 0 And Nvl(c.是否临床排班, 0) = 1 And Exists
                         (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
                      --站点 
                       And (d.站点 Is Null Or d.站点 = 站点_In) And Not Exists
                  (Select 1 From 临床出诊记录 Where 安排id = b.Id)) Loop
      Zl_临床出诊安排_Delete(c_安排.Id);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊安排_Batchdelete;
/

--108192:蒋廷中,2017-04-14,解决icd附码删除错误的问题
Create Or Replace Procedure Zl_病人诊断记录_Delete
(
  --功能：删除病人诊断记录 
  --参数：诊断类型_IN=为空时表示所有类型,否则为字符串,如'1,2,3...' 
  --      诊断s_In=需要删除的诊断ID串 ,格式为 'ID1,ID2,ID3...'  
  病人id_In   病人诊断记录.病人id%Type,
  主页id_In   病人诊断记录.主页id%Type,
  记录来源_In 病人诊断记录.记录来源%Type := Null,
  病历id_In   病人诊断记录.病历id%Type := Null,
  诊断类型_In Varchar2 := Null,
  诊断ids_In  Varchar2 := Null
) Is
  V_类型串 Varchar2(255);
  V_类型   病人诊断记录.诊断类型%Type;
Begin
  If 诊断类型_In Is Null Then
    If Not 诊断ids_In Is Null Then
      For Rdiag In (Select /*+ Rule*/
                     ID, 记录来源, 诊断类型, 诊断次序
                    From 病人诊断记录
                    Where ID In (Select Column_Value From Table(F_Str2list(诊断ids_In)))
                    Order By 记录来源, 诊断类型, 诊断次序) Loop
        If Rdiag.记录来源 = 3 And Rdiag.诊断类型 = 2 And Rdiag.诊断次序 = 1 Then
          Update 病案主页 Set 单病种 = Null Where 病人id = 病人id_In And 主页id = Nvl(主页id_In, 0);
        End If;
        --如果诊断类型是当前传入的诊断类型，则删除诊断
        If Rdiag.记录来源 = 记录来源_In Or 记录来源_In Is Null Then
          Delete From 病人诊断医嘱 Where 诊断id = Rdiag.Id;
          Delete From 病人诊断记录
          Where 病人id = 病人id_In And 主页id = 主页id_In And 记录来源 = Rdiag.记录来源 And 诊断类型 = Rdiag.诊断类型 And 诊断次序 = Rdiag.诊断次序 And
                Nvl(编码序号, 1) = 2;
          Delete From 病人诊断记录 Where ID = Rdiag.Id;
        End If;
      End Loop;
    Else
      Delete From 病人诊断医嘱
      Where 诊断id In (Select ID
                     From 病人诊断记录
                     Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And (记录来源 = 记录来源_In Or 记录来源_In Is Null) And
                           (病历id = 病历id_In Or 病历id_In Is Null));
    
      Delete From 病人诊断记录
      Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And (记录来源 = 记录来源_In Or 记录来源_In Is Null) And
            (病历id = 病历id_In Or 病历id_In Is Null);
      --删除单病种标识 
      If 记录来源_In = 3 Then
        Update 病案主页 Set 单病种 = Null Where 病人id = 病人id_In And 主页id = Nvl(主页id_In, 0);
      End If;
    End If;
  Else
    V_类型串 := 诊断类型_In || ',';
    While V_类型串 Is Not Null Loop
      V_类型 := To_Number(Substr(V_类型串, 1, Instr(V_类型串, ',') - 1));
    
      Delete From 病人诊断医嘱
      Where 诊断id In (Select ID
                     From 病人诊断记录
                     Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And (记录来源 = 记录来源_In Or 记录来源_In Is Null) And
                           (病历id = 病历id_In Or 病历id_In Is Null) And 诊断类型 = V_类型);
    
      Delete From 病人诊断记录
      Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And (记录来源 = 记录来源_In Or 记录来源_In Is Null) And
            (病历id = 病历id_In Or 病历id_In Is Null) And 诊断类型 = V_类型;
    
      V_类型串 := Substr(V_类型串, Instr(V_类型串, ',') + 1);
    
      --如果是入院诊断则删除单病种标识 
      If V_类型 = 2 And 记录来源_In = 3 Then
        Update 病案主页 Set 单病种 = Null Where 病人id = 病人id_In And 主页id = Nvl(主页id_In, 0);
      End If;
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人诊断记录_Delete;
/

--106872:刘尔旋,2017-04-12,预约接收保存摘要
Create Or Replace Procedure Zl_预约挂号接收_Insert
(
  No_In            门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  领用id_In        票据使用明细.领用id%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  结算方式_In      病人预交记录.结算方式%Type, --现金的结算名称
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额
  预交支付_In      病人预交记录.冲预交%Type, --挂号时使用的预交金额
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额
  发生时间_In      门诊费用记录.发生时间%Type,
  号序_In          挂号序号状态.序号%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  生成队列_In      Number := 0,
  登记时间_In      门诊费用记录.登记时间%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  冲预交病人ids_In Varchar2 := Null,
  三方调用_In      Number := 0,
  更新交款余额_In  Number := 0, --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
  摘要_In          病人挂号记录.摘要%Type := Null
) As
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, NO;

  v_Err_Msg Varchar2(255);
  Err_Item    Exception;
  Err_Special Exception;

  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;
  v_号别     门诊费用记录.计算单位%Type;
  v_号序     门诊费用记录.发药窗口%Type;
  v_排队号码 排队叫号队列.排队号码 %Type;
  v_预约方式 病人挂号记录.预约方式 %Type;

  n_打印id        票据打印内容.Id%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  v_冲预交病人ids Varchar2(4000);

  n_挂号id         病人挂号记录.Id%Type;
  n_分诊台签到排队 Number;
  n_组id           财务缴款分组.Id%Type;
  n_Count          Number(18);
  n_排队           Number;
  n_当天排队       Number;
  n_当前金额       病人预交记录.金额%Type;
  n_预交id         病人预交记录.Id%Type;
  n_消费卡id       消费卡目录.Id%Type;
  n_自制卡         Number;

  d_Date       Date;
  d_预约时间   门诊费用记录.发生时间%Type;
  d_发生时间   Date;
  d_排队时间   Date;
  n_时段       Number := 0;
  n_存在       Number := 0;
  v_排队序号   排队叫号队列.排队序号%Type;
  n_结算模式   病人信息.结算模式%Type;
  n_票种       票据使用明细.票种%Type;
  v_付款方式   病人挂号记录.医疗付款方式%Type;
  v_操作员姓名 病人挂号记录.接收人%Type;
  n_接收模式   Number := 0;
Begin
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  n_接收模式      := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);

  --获取结算方式名称
  Begin
    Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_现金 := '现金';
  End;
  Begin
    Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
  Exception
    When Others Then
      v_个人帐户 := '个人帐户';
  End;
  If 登记时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 登记时间_In;
  End If;

  --更新挂号序号状态
  Begin
    Select 号别, 号序, Trunc(发生时间), 发生时间, 预约方式
    Into v_号别, v_号序, d_预约时间, d_发生时间, v_预约方式
    From 病人挂号记录
    Where 记录性质 = 2 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(接收人) Into v_操作员姓名 From 病人挂号记录 Where 记录性质 = 2 And 记录状态 In (1, 3) And NO = No_In;
      If v_操作员姓名 Is Null Then
        v_Err_Msg := '当前预约挂号单已被取消';
        Raise Err_Item;
      Else
        If v_操作员姓名 = 操作员姓名_In Then
          v_Err_Msg := '当前预约挂号单已被接收';
          Raise Err_Special;
        Else
          v_Err_Msg := '当前预约挂号单已被其它人接收';
          Raise Err_Item;
        End If;
      End If;
  End;

  --判断是否分时段
  Begin
    Select 1
    Into n_时段
    From Dual
    Where Exists (Select 1
           From 挂号安排时段 A, 挂号安排 B
           Where a.安排id = b.Id And b.号码 = v_号别 And Rownum < 2
           Union All
           Select 1
           From 挂号计划时段 C, 挂号安排计划 D 　
           Where c.计划id = d.Id And d.号码 = v_号别 And d.生效时间 > Sysdate And Rownum < 2);
  Exception
    When Others Then
      n_时段 := 0;
  End;

  If n_时段 = 0 And 三方调用_In = 0 Then
    If n_接收模式 = 0 Then
      If Trunc(发生时间_In) = Trunc(Sysdate) Then
        d_发生时间 := 发生时间_In;
      Else
        d_发生时间 := Sysdate;
      End If;
    Else
      d_发生时间 := 发生时间_In;
    End If;
  Else
    If Not 发生时间_In Is Null Then
      d_发生时间 := 发生时间_In;
    End If;
  End If;
  If Not v_号序 Is Null Then
    If 号序_In Is Null Then
      Delete 挂号序号状态 Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
    Else
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
      
        If n_时段 = 0 And 三方调用_In = 0 Then
          --提前接收或延迟接收
          Delete 挂号序号状态 Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
          Begin
            Select 1 Into n_存在 From 挂号序号状态 Where 号码 = v_号别 And 日期 = Trunc(Sysdate) And 序号 = v_号序;
          Exception
            When Others Then
              n_存在 := 0;
          End;
          If n_存在 = 0 Then
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
            Values
              (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
          Else
            --号码已被使用的情况
            Begin
              v_号序 := 1;
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
              Values
                (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            Exception
              When Others Then
                Select Min(序号 + 1)
                Into v_号序
                From 挂号序号状态 A
                Where 号码 = v_号别 And 日期 = Trunc(Sysdate) And Not Exists
                 (Select 1 From 挂号序号状态 Where 号码 = a.号码 And 日期 = a.日期 And 序号 = a.序号 + 1);
                Insert Into 挂号序号状态
                  (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
                Values
                  (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            End;
          End If;
        Else
          Update 挂号序号状态
          Set 状态 = 1, 登记时间 = Sysdate
          Where Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序 And 号码 = v_号别 And 状态 = 2;
          If Sql% NotFound Then
            Begin
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
              Values
                (v_号别, Trunc(Sysdate), v_号序, 1, 操作员姓名_In, Sysdate);
            Exception
              When Others Then
                v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
                Raise Err_Item;
            End;
          End If;
        
        End If;
      
      Else
        Update 挂号序号状态
        Set 序号 = 号序_In, 状态 = 1, 登记时间 = Sysdate
        Where 号码 = v_号别 And Trunc(日期) = Trunc(d_预约时间) And 序号 = v_号序;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
            Values
              (v_号别, Trunc(d_发生时间), v_号序, 1, 操作员姓名_In, Sysdate);
          Exception
            When Others Then
              v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  Else
    If Not 号序_In Is Null Then
      Begin
        Insert Into 挂号序号状态
          (号码, 日期, 序号, 状态, 操作员姓名, 登记时间)
        Values
          (v_号别, Trunc(Sysdate), 号序_In, 1, 操作员姓名_In, Sysdate);
      Exception
        When Others Then
          v_Err_Msg := '序号' || 号序_In || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
      End;
      v_号序 := 号序_In;
    Else
      v_号序 := Null;
    End If;
  End If;

  --更新门诊费用记录
  Update 门诊费用记录
  Set 记录状态 = 1, 实际票号 = Decode(Nvl(记帐费用_In, 0), 1, Null, 票据号_In), 结帐id = Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In),
      结帐金额 = Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额), 发药窗口 = 诊室_In, 病人id = 病人id_In, 标识号 = 门诊号_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
      性别 = 性别_In, 付款方式 = 付款方式_In, 费别 = 费别_In, 发生时间 = d_发生时间, 登记时间 = d_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In,
      缴款组id = n_组id, 记帐费用 = Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 摘要 = Nvl(摘要_In, 摘要)
  Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;

  --病人挂号记录
  Update 病人挂号记录
  Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = 病人id_In, 门诊号 = 门诊号_In, 发生时间 = d_发生时间, 姓名 = 姓名_In, 性别 = 性别_In,
      年龄 = 年龄_In, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In), 号序 = v_号序, 诊室 = 诊室_In,
      摘要 = Nvl(摘要_In, 摘要)
  Where 记录状态 = 1 And NO = No_In And 记录性质 = 2
  Returning ID Into n_挂号id;
  If Sql%NotFound Then
    Begin
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         摘要, 号序, 预约, 预约方式, 接收人, 接收时间, 预约时间, 险类, 医疗付款方式)
        Select n_挂号id, No_In, 1, 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 计算单位, 加班标志, 诊室_In, Null, 执行部门id, 执行人, 0, Null,
               登记时间, 发生时间, 操作员编号, 操作员姓名, Nvl(摘要_In, 摘要), v_号序, 1, Substr(结论, 1, 10) As 预约方式, 操作员姓名_In,
               Nvl(登记时间_In, Sysdate), 发生时间, Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '由于并发原因,单据号为【' || No_In || '】的病人' || 姓名_In || '已经被接收';
        Raise Err_Item;
    End;
  End If;

  --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
  If Nvl(生成队列_In, 0) <> 0 Then
    n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
    If Nvl(n_分诊台签到排队, 0) = 0 Then
      For v_挂号 In (Select ID, 姓名, 诊室, 执行人, 执行部门id, 发生时间, 号别, 号序 From 病人挂号记录 Where NO = No_In) Loop
      
        Begin
          Select 1,
                 Case
                   When 排队时间 < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_排队, n_当天排队
          From 排队叫号队列
          Where 业务类型 = 0 And 业务id = v_挂号.Id And Rownum <= 1;
        Exception
          When Others Then
            n_排队 := 0;
        End;
        If n_排队 = 0 Then
          --产生队列
          --按”执行部门”产生队列
          n_挂号id   := v_挂号.Id;
          v_队列名称 := v_挂号.执行部门id;
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, n_挂号id, v_挂号.号别 || '|' || v_挂号.号序);
          v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        
          --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
          d_排队时间 := Zl_Get_Queuedate(n_挂号id, v_挂号.号别, v_挂号.号序, v_挂号.发生时间);
          --   队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In,
          Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, v_挂号.执行部门id, v_排队号码, Null, 姓名_In, 病人id_In, v_挂号.诊室, v_挂号.执行人, d_排队时间,
                           v_预约方式, Null, v_排队序号);
        Elsif Nvl(n_当天排队, 0) = 1 Then
          --更新队列号
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, v_挂号.Id, v_挂号.号别 || '|' || Nvl(v_挂号.号序, 0));
          v_排队序号 := Zlgetsequencenum(0, v_挂号.Id, 1);
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人, v_排队号码, v_排队序号);
        
        Else
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人);
        End If;
        --预约接收时，改变记录标志
        Update 病人挂号记录 Set 记录标志 = 1 Where ID = n_挂号id;
      End Loop;
    End If;
  End If;

  --汇总结算到病人预交记录
  If (Nvl(现金支付_In, 0) <> 0 Or (Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0)) And
     Nvl(记帐费用_In, 0) = 0 Then
    Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 结算序号,
       结算性质)
    Values
      (n_预交id, 4, 1, No_In, 病人id_In, Nvl(结算方式_In, v_现金), Nvl(现金支付_In, 0), d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
       n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 结帐id_In, 4);
  
    If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
    
      n_消费卡id := Null;
      Begin
        Select Nvl(自制卡, 0), 1 Into n_自制卡, n_Count From 卡消费接口目录 Where 编号 = 结算卡序号_In;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count = 0 Then
        v_Err_Msg := '[ZLSOFT]没有发现原结算卡的相应类别,不能继续操作！[ZLSOFT]';
        Raise Err_Item;
      End If;
      If n_自制卡 = 1 Then
        Select ID
        Into n_消费卡id
        From 消费卡目录
        Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In And
              序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In);
      End If;
      Zl_病人卡结算记录_Insert(结算卡序号_In, n_消费卡id, 结算方式_In, 现金支付_In, 卡号_In, Null, 登记时间_In, Null, 结帐id_In, n_预交id);
    End If;
  
  End If;

  --对于就诊卡通过预交金挂号
  If Nvl(预交支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    n_预交金额 := 预交支付_In;
    For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
      n_当前金额 := Case
                  When r_Deposit.金额 - n_预交金额 < 0 Then
                   r_Deposit.金额
                  Else
                   n_预交金额
                End;
      If r_Deposit.结帐id = 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 登记时间_In,
               操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结帐id_In, 4
        From 病人预交记录
        Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
      Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (r_Deposit.病人id, Nvl(1, 2), -1 * n_当前金额, 1);
        n_返回值 := -1 * n_当前金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --检查是否已经处理完
      If r_Deposit.金额 < n_预交金额 Then
        n_预交金额 := n_预交金额 - r_Deposit.金额;
      Else
        n_预交金额 := 0;
      End If;
    
      If n_预交金额 = 0 Then
        Exit;
      End If;
    End Loop;
  End If;

  --对于医保挂号
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       预交类别, 结算序号, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 4, 1, No_In, 病人id_In, v_个人帐户, 个帐支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '医保挂号', n_组id,
       Null, Null, Null, Null, Null, Null, Null, 结帐id_In, 4);
  End If;

  --相关汇总表的处理
  --人员缴款余额
  If Nvl(现金支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 0) = 0 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 现金支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金)
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额
        (收款员, 结算方式, 性质, 余额)
      Values
        (操作员姓名_In, Nvl(结算方式_In, v_现金), 1, 现金支付_In);
      n_返回值 := 现金支付_In;
    
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = Nvl(结算方式_In, v_现金) And Nvl(余额, 0) = 0;
    End If;
  End If;

  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 0) = 0 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 个帐支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
      n_返回值 := 个帐支付_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And Nvl(余额, 0) = 0;
    End If;
  End If;

  --处理票据使用情况
  If 票据号_In Is Not Null And Nvl(记帐费用_In, 0) = 0 Then
    Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
  
    --当前票据的票种
    Select 票种 Into n_票种 From 票据领用记录 Where ID = Nvl(领用id_In, 0);
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, No_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
    Values
      (票据使用明细_Id.Nextval, n_票种, 票据号_In, 1, 1, 领用id_In, n_打印id, d_Date, 操作员姓名_In);
  
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = d_Date
    Where ID = Nvl(领用id_In, 0);
  End If;

  If Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 1 And NO = No_In And Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (病人id_In, 1, 1, Nvl(c_费用.实收金额, 0), 0);
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    n_结算模式 := 0;
    Update 病人信息
    Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In
    Where 病人id = 病人id_In
    Returning Nvl(结算模式, 0) Into n_结算模式;
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  End If;

  --病人担保信息
  If 病人id_In Is Not Null Then
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, d_Date) >= d_Date;
    End If;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 1, n_挂号id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约挂号接收_Insert;
/

--106872:刘尔旋,2017-04-12,预约接收保存摘要
Create Or Replace Procedure Zl_预约挂号接收_出诊_Insert
(
  No_In            门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  领用id_In        票据使用明细.领用id%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  结算方式_In      Varchar2, --现金的结算名称
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额
  预交支付_In      病人预交记录.冲预交%Type, --挂号时使用的预交金额
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额
  发生时间_In      门诊费用记录.发生时间%Type,
  号序_In          挂号序号状态.序号%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  生成队列_In      Number := 0,
  登记时间_In      门诊费用记录.登记时间%Type := Null,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  冲预交病人ids_In Varchar2 := Null,
  三方调用_In      Number := 0,
  更新交款余额_In  Number := 0, --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
  摘要_In          病人挂号记录.摘要%Type := Null
) As
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, NO;

  v_Err_Msg Varchar2(255);
  Err_Item    Exception;
  Err_Special Exception;
  v_操作员姓名 病人挂号记录.接收人%Type;
  v_现金       结算方式.名称%Type;
  v_个人帐户   结算方式.名称%Type;
  v_队列名称   排队叫号队列.队列名称%Type;
  v_号别       门诊费用记录.计算单位%Type;
  v_号序       门诊费用记录.发药窗口%Type;
  v_排队号码   排队叫号队列.排队号码 %Type;
  v_预约方式   病人挂号记录.预约方式 %Type;

  n_打印id        票据打印内容.Id%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  v_冲预交病人ids Varchar2(4000);

  n_挂号id         病人挂号记录.Id%Type;
  n_分诊台签到排队 Number;
  n_组id           财务缴款分组.Id%Type;
  n_Count          Number(18);
  n_排队           Number;
  n_当天排队       Number;
  n_当前金额       病人预交记录.金额%Type;
  n_预交id         病人预交记录.Id%Type;
  n_消费卡id       消费卡目录.Id%Type;
  n_自制卡         Number;

  d_Date         Date;
  d_预约时间     门诊费用记录.发生时间%Type;
  d_发生时间     Date;
  d_排队时间     Date;
  n_时段         Number := 0;
  n_存在         Number := 0;
  v_结算内容     Varchar2(2000);
  v_当前结算     Varchar2(500);
  n_结算金额     病人预交记录.冲预交%Type;
  v_结算号码     病人预交记录.结算号码%Type;
  v_结算方式     病人预交记录.结算方式%Type;
  n_三方卡标志   Number(3);
  v_排队序号     排队叫号队列.排队序号%Type;
  n_结算模式     病人信息.结算模式%Type;
  n_票种         票据使用明细.票种%Type;
  v_付款方式     病人挂号记录.医疗付款方式%Type;
  n_接收模式     Number := 0;
  n_出诊记录id   病人挂号记录.出诊记录id%Type;
  n_新出诊记录id 病人挂号记录.出诊记录id%Type;
  n_号源id       临床出诊记录.号源id%Type;
  n_预约顺序号   临床出诊序号控制.预约顺序号%Type;
  n_旧分时段     临床出诊记录.是否分时段%Type;
  n_旧序号控制   临床出诊记录.是否序号控制%Type;
  n_旧科室id     临床出诊记录.科室id%Type;
  n_旧项目id     临床出诊记录.项目id%Type;
  n_旧医生id     临床出诊记录.医生id%Type;
  n_挂号模式     Number(3);
  d_启用时间     Date;
  v_Paratemp     Varchar2(500);
  v_Registtemp   Varchar2(500);
  n_检查         Number(3);
  n_序号控制     临床出诊记录.是否序号控制%Type;
  v_旧上班时段   临床出诊记录.上班时段%Type;
Begin
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);
  v_Paratemp      := Nvl(zl_GetSysParameter('挂号排班模式'), 0);
  n_接收模式      := Nvl(zl_GetSysParameter('预约接收模式', 1111), 0);
  n_挂号模式      := To_Number(Substr(v_Paratemp, 1, 1));
  If n_挂号模式 = 1 Then
    Begin
      d_启用时间 := To_Date(Substr(v_Paratemp, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
  End If;

  --获取结算方式名称
  Begin
    Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
  Exception
    When Others Then
      v_现金 := '现金';
  End;
  Begin
    Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
  Exception
    When Others Then
      v_个人帐户 := '个人帐户';
  End;
  If 登记时间_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := 登记时间_In;
  End If;

  --更新挂号序号状态
  Begin
    Select 号别, 号序, Trunc(发生时间), 发生时间, 预约方式, 出诊记录id
    Into v_号别, v_号序, d_预约时间, d_发生时间, v_预约方式, n_出诊记录id
    From 病人挂号记录
    Where 记录性质 = 2 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(接收人) Into v_操作员姓名 From 病人挂号记录 Where 记录性质 = 2 And 记录状态 In (1, 3) And NO = No_In;
      If v_操作员姓名 Is Null Then
        v_Err_Msg := '当前预约挂号单已被取消';
        Raise Err_Item;
      Else
        If v_操作员姓名 = 操作员姓名_In Then
          v_Err_Msg := '当前预约挂号单已被接收';
          Raise Err_Special;
        Else
          v_Err_Msg := '当前预约挂号单已被其它人接收';
          Raise Err_Item;
        End If;
      End If;
  End;

  --判断是否分时段
  Select Nvl(是否分时段, 0), 号源id, Nvl(是否序号控制, 0)
  Into n_时段, n_号源id, n_序号控制
  From 临床出诊记录
  Where ID = n_出诊记录id;

  If n_时段 = 1 And 三方调用_In = 0 And n_接收模式 = 0 Then
    If Trunc(发生时间_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '分时段的预约挂号单只能当天接收！';
      Raise Err_Item;
    End If;
  End If;

  If n_时段 = 0 And 三方调用_In = 0 Then
    If n_接收模式 = 0 Then
      If Trunc(发生时间_In) = Trunc(Sysdate) Then
        d_发生时间 := 发生时间_In;
      Else
        d_发生时间 := Sysdate;
      End If;
    Else
      d_发生时间 := 发生时间_In;
    End If;
  Else
    If Not 发生时间_In Is Null Then
      d_发生时间 := 发生时间_In;
    End If;
  End If;

  If d_启用时间 Is Not Null Then
    If d_发生时间 < d_启用时间 Then
      v_Err_Msg := '当前预约挂号单属于出诊表排班模式安排，不能在' || To_Char(d_启用时间, 'yyyy-mm-dd hh24:mi:ss') || '之前接收!';
      Raise Err_Item;
    End If;
  End If;

  If Not v_号序 Is Null Then
    If 号序_In Is Null Then
      Update 临床出诊序号控制 Set 挂号状态 = 0 Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id;
    Else
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
        If n_时段 = 0 And 三方调用_In = 0 Then
          --提前接收或延迟接收
          Update 临床出诊序号控制 Set 挂号状态 = 0 Where 序号 = v_号序 And 记录id = n_出诊记录id;
        
          Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
          Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
          From 临床出诊记录
          Where ID = n_出诊记录id;
          Begin
            Select ID
            Into n_新出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                  Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
              Raise Err_Item;
          End;
        
          Begin
            Select 1
            Into n_存在
            From 临床出诊序号控制
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          Exception
            When Others Then
              n_存在 := 0;
          End;
        
          If n_存在 = 1 Then
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          Else
            --号码已被使用的情况
            Select Min(序号) Into v_号序 From 临床出诊序号控制 Where 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
            If v_号序 Is Null Then
              v_Err_Msg := '接收当天没有可用序号,无法接收!';
              Raise Err_Item;
            End If;
            Update 临床出诊序号控制
            Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
            Where 记录id = n_新出诊记录id And 序号 = v_号序 And Nvl(挂号状态, 0) = 0;
          End If;
        Else
          Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
          Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
          From 临床出诊记录
          Where ID = n_出诊记录id;
          Begin
            Select ID
            Into n_新出诊记录id
            From 临床出诊记录
            Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                  Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
              Raise Err_Item;
          End;
          Update 临床出诊序号控制
          Set 挂号状态 = 0, 操作员姓名 = 操作员姓名_In
          Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id And Nvl(挂号状态, 0) = 2
          Returning 预约顺序号 Into n_预约顺序号;
        
          Update 临床出诊序号控制
          Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In, 预约顺序号 = n_预约顺序号
          Where 序号 = v_号序 And 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
          If Sql% RowCount = 0 Then
            v_Err_Msg := '接收当天序号' || v_号序 || '已被其它人使用,无法接收.';
            Raise Err_Item;
          End If;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = v_号序 Or 备注 = v_号序) And 记录id = n_出诊记录id;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '序号' || v_号序 || '已被其它人使用,请重新选择一个序号.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  Else
    If Not 号序_In Is Null Then
      If Trunc(d_预约时间) <> Trunc(Sysdate) And n_接收模式 = 0 Then
        Select 是否分时段, 是否序号控制, 科室id, 医生id, 项目id, 上班时段
        Into n_旧分时段, n_旧序号控制, n_旧科室id, n_旧医生id, n_旧项目id, v_旧上班时段
        From 临床出诊记录
        Where ID = n_出诊记录id;
        Begin
          Select ID
          Into n_新出诊记录id
          From 临床出诊记录
          Where 号源id = n_号源id And 是否分时段 = n_旧分时段 And 是否序号控制 = n_旧序号控制 And 科室id = n_旧科室id And
                Nvl(医生id, 0) = Nvl(n_旧医生id, 0) And 上班时段 = v_旧上班时段 And Nvl(是否发布, 0) = 1 And 出诊日期 = Trunc(Sysdate) And
                Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '接收当天没有对应的出诊安排,无法接收!';
            Raise Err_Item;
        End;
        Update 临床出诊序号控制
        Set 挂号状态 = 0, 操作员姓名 = 操作员姓名_In
        Where (序号 = 号序_In Or 备注 = 号序_In) And 记录id = n_出诊记录id And Nvl(挂号状态, 0) = 2
        Returning 预约顺序号 Into n_预约顺序号;
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In, 预约顺序号 = n_预约顺序号
        Where 序号 = 号序_In And 记录id = n_新出诊记录id And Nvl(挂号状态, 0) = 0;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '接收当天序号' || 号序_In || '已被其它人使用,无法接收.';
          Raise Err_Item;
        End If;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = 1, 操作员姓名 = 操作员姓名_In
        Where (序号 = 号序_In Or 备注 = 号序_In) And 记录id = n_出诊记录id;
      
      End If;
      v_号序 := 号序_In;
    Else
      v_号序 := Null;
    End If;
  End If;

  --更新门诊费用记录
  Update 门诊费用记录
  Set 记录状态 = 1, 实际票号 = Decode(Nvl(记帐费用_In, 0), 1, Null, 票据号_In), 结帐id = Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In),
      结帐金额 = Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额), 发药窗口 = 诊室_In, 病人id = 病人id_In, 标识号 = 门诊号_In, 姓名 = 姓名_In, 年龄 = 年龄_In,
      性别 = 性别_In, 付款方式 = 付款方式_In, 费别 = 费别_In, 发生时间 = d_发生时间, 登记时间 = d_Date, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In,
      缴款组id = n_组id, 记帐费用 = Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 摘要 = Nvl(摘要_In, 摘要)
  Where 记录性质 = 4 And 记录状态 = 0 And NO = No_In;

  v_Registtemp := zl_GetSysParameter('挂号排班模式');
  If Substr(v_Registtemp, 1, 1) = 1 Then
    Begin
      If To_Date(Substr(v_Registtemp, 3), 'yyyy-mm-dd hh24:mi:ss') > d_发生时间 Then
        v_Err_Msg := '接收时间' || To_Char(d_发生时间, 'yyyy-mm-dd hh24:mi:ss') || '未启用出诊表排班模式,目前无法接收!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select 1
      Into n_检查
      From 临床出诊记录
      Where ID = Nvl(n_新出诊记录id, n_出诊记录id) And d_发生时间 Between 停诊开始时间 And 停诊终止时间;
    Exception
      When Others Then
        n_检查 := 0;
    End;
    If n_检查 = 1 And Not (n_时段 = 1 And n_序号控制 = 1) Then
      v_Err_Msg := '接收时间' || To_Char(d_发生时间, 'yyyy-mm-dd hh24:mi:ss') || '的安排已经被停诊,无法接收!';
      Raise Err_Item;
    End If;
  End If;

  --病人挂号记录
  Update 病人挂号记录
  Set 接收人 = 操作员姓名_In, 接收时间 = d_Date, 记录性质 = 1, 病人id = 病人id_In, 门诊号 = 门诊号_In, 发生时间 = d_发生时间, 姓名 = 姓名_In, 性别 = 性别_In,
      年龄 = 年龄_In, 操作员编号 = 操作员编号_In, 操作员姓名 = 操作员姓名_In, 险类 = Decode(Nvl(险类_In, 0), 0, Null, 险类_In), 号序 = v_号序, 诊室 = 诊室_In,
      出诊记录id = Nvl(n_新出诊记录id, n_出诊记录id), 摘要 = Nvl(摘要_In, 摘要)
  Where 记录状态 = 1 And NO = No_In And 记录性质 = 2
  Returning ID Into n_挂号id;
  If Sql%NotFound Then
    Begin
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Begin
        Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
      Exception
        When Others Then
          v_付款方式 := Null;
      End;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         摘要, 号序, 预约, 预约方式, 接收人, 接收时间, 预约时间, 险类, 医疗付款方式, 出诊记录id)
        Select n_挂号id, No_In, 1, 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 计算单位, 加班标志, 诊室_In, Null, 执行部门id, 执行人, 0, Null,
               登记时间, 发生时间, 操作员编号, 操作员姓名, Nvl(摘要_In, 摘要), v_号序, 1, Substr(结论, 1, 10) As 预约方式, 操作员姓名_In,
               Nvl(登记时间_In, Sysdate), 发生时间, Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, Nvl(n_新出诊记录id, n_出诊记录id)
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '由于并发原因,单据号为【' || No_In || '】的病人' || 姓名_In || '已经被接收';
        Raise Err_Item;
    End;
  End If;

  --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
  If Nvl(生成队列_In, 0) <> 0 Then
    n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
    If Nvl(n_分诊台签到排队, 0) = 0 Then
      For v_挂号 In (Select ID, 姓名, 诊室, 执行人, 执行部门id, 发生时间, 号别, 号序 From 病人挂号记录 Where NO = No_In) Loop
      
        Begin
          Select 1,
                 Case
                   When 排队时间 < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_排队, n_当天排队
          From 排队叫号队列
          Where 业务类型 = 0 And 业务id = v_挂号.Id And Rownum <= 1;
        Exception
          When Others Then
            n_排队 := 0;
        End;
        If n_排队 = 0 Then
          --产生队列
          --按”执行部门”产生队列
          n_挂号id   := v_挂号.Id;
          v_队列名称 := v_挂号.执行部门id;
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, n_挂号id, v_挂号.号别 || '|' || v_挂号.号序);
          v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        
          --挂号id_In,号码_In,号序_In,缺省日期_In,扩展_In(暂无用)
          d_排队时间 := Zl_Get_Queuedate(n_挂号id, v_挂号.号别, v_挂号.号序, v_挂号.发生时间);
          --   队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In,
          Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, v_挂号.执行部门id, v_排队号码, Null, 姓名_In, 病人id_In, v_挂号.诊室, v_挂号.执行人, d_排队时间,
                           v_预约方式, Null, v_排队序号);
        Elsif Nvl(n_当天排队, 0) = 1 Then
          --更新队列号
          v_排队号码 := Zlgetnextqueue(v_挂号.执行部门id, v_挂号.Id, v_挂号.号别 || '|' || Nvl(v_挂号.号序, 0));
          v_排队序号 := Zlgetsequencenum(0, v_挂号.Id, 1);
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人, v_排队号码, v_排队序号);
        
        Else
          --新队列名称_IN, 业务类型_In, 业务id_In , 科室id_In , 患者姓名_In , 诊室_In, 医生姓名_In ,排队号码_In
          Zl_排队叫号队列_Update(v_挂号.执行部门id, 0, v_挂号.Id, v_挂号.执行部门id, v_挂号.姓名, v_挂号.诊室, v_挂号.执行人);
        End If;
      End Loop;
    End If;
  End If;

  --汇总结算到病人预交记录
  If Nvl(记帐费用_In, 0) = 0 Then
    If Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, 登记时间_In, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
         n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, 4);
    End If;
    If Nvl(现金支付_In, 0) <> 0 Then
      v_结算内容 := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);
      
        If n_三方卡标志 = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, Null, Null, Null, Null, Null, Null, 4, v_结算号码);
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, No_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, Null, 4, v_结算号码);
          If Nvl(结算卡序号_In, 0) <> 0 Then
            n_消费卡id := Null;
            Begin
              Select Nvl(自制卡, 0), 1 Into n_自制卡, n_Count From 卡消费接口目录 Where 编号 = 结算卡序号_In;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 0 Then
              v_Err_Msg := '没有发现原结算卡的相应类别,不能继续操作！';
              Raise Err_Item;
            End If;
            If n_自制卡 = 1 Then
              Select ID
              Into n_消费卡id
              From 消费卡目录
              Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In And
                    序号 = (Select Max(序号) From 消费卡目录 Where 接口编号 = 结算卡序号_In And 卡号 = 卡号_In);
            End If;
            Zl_病人卡结算记录_Insert(结算卡序号_In, n_消费卡id, v_结算方式, n_结算金额, 卡号_In, Null, 登记时间_In, Null, 结帐id_In, n_预交id);
          End If;
        End If;
      
        If Nvl(更新交款余额_In, 0) = 0 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + n_结算金额
          Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金)
          Returning 余额 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, Nvl(v_结算方式, v_现金), 1, n_结算金额);
            n_返回值 := n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End If;
      
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
    End If;
  End If;

  --对于就诊卡通过预交金挂号
  If Nvl(预交支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    n_预交金额 := 预交支付_In;
    For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
      n_当前金额 := Case
                  When r_Deposit.金额 - n_预交金额 < 0 Then
                   r_Deposit.金额
                  Else
                   n_预交金额
                End;
      If r_Deposit.结帐id = 0 Then
        --第一次冲预交(填上结帐ID,金额为0)
        Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
      End If;
      --冲上次剩余额
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号, 冲预交,
         结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结算序号, 结算性质)
        Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 登记时间_In,
               操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 预交类别, 结帐id_In, 4
        From 病人预交记录
        Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
    
      --更新病人预交余额
      Update 病人余额
      Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
      Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2)
      Returning 预交余额 Into n_返回值;
      If Sql%RowCount = 0 Then
        Insert Into 病人余额 (病人id, 类型, 预交余额, 性质) Values (r_Deposit.病人id, Nvl(1, 2), -1 * n_当前金额, 1);
        n_返回值 := -1 * n_当前金额;
      End If;
      If Nvl(n_返回值, 0) = 0 Then
        Delete From 病人余额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
      End If;
    
      --检查是否已经处理完
      If r_Deposit.金额 < n_预交金额 Then
        n_预交金额 := n_预交金额 - r_Deposit.金额;
      Else
        n_预交金额 := 0;
      End If;
    
      If n_预交金额 = 0 Then
        Exit;
      End If;
    End Loop;
  End If;

  --对于医保挂号
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 Then
    Insert Into 病人预交记录
      (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
       预交类别, 结算序号, 结算性质)
    Values
      (病人预交记录_Id.Nextval, 4, 1, No_In, 病人id_In, v_个人帐户, 个帐支付_In, d_Date, 操作员编号_In, 操作员姓名_In, 结帐id_In, '医保挂号', n_组id,
       Null, Null, Null, Null, Null, Null, Null, 结帐id_In, 4);
  End If;

  --相关汇总表的处理
  --人员缴款余额
  If Nvl(个帐支付_In, 0) <> 0 And Nvl(记帐费用_In, 0) = 0 And Nvl(更新交款余额_In, 0) = 0 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 个帐支付_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
    Returning 余额 Into n_返回值;
  
    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
      n_返回值 := 个帐支付_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And Nvl(余额, 0) = 0;
    End If;
  End If;

  --处理票据使用情况
  If 票据号_In Is Not Null And Nvl(记帐费用_In, 0) = 0 Then
    Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
  
    --当前票据的票种
    Select 票种 Into n_票种 From 票据领用记录 Where ID = Nvl(领用id_In, 0);
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 4, No_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人)
    Values
      (票据使用明细_Id.Nextval, n_票种, 票据号_In, 1, 1, 领用id_In, n_打印id, d_Date, 操作员姓名_In);
  
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = d_Date
    Where ID = Nvl(领用id_In, 0);
  End If;

  If Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
    For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                 From 门诊费用记录
                 Where 记录性质 = 4 And 记录状态 = 1 And NO = No_In And Nvl(记帐费用, 0) = 1) Loop
      --病人余额
      Update 病人余额
      Set 费用余额 = Nvl(费用余额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人余额
          (病人id, 性质, 类型, 费用余额, 预交余额)
        Values
          (病人id_In, 1, 1, Nvl(c_费用.实收金额, 0), 0);
      End If;
    
      --病人未结费用
      Update 病人未结费用
      Set 金额 = Nvl(金额, 0) + Nvl(c_费用.实收金额, 0)
      Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
            Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And 收入项目id + 0 = c_费用.收入项目id And
            来源途径 + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into 病人未结费用
          (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
        Values
          (病人id_In, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, Nvl(c_费用.实收金额, 0));
      End If;
    End Loop;
  End If;
  If Nvl(病人id_In, 0) <> 0 Then
    n_结算模式 := 0;
    Update 病人信息
    Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In
    Where 病人id = 病人id_In
    Returning Nvl(结算模式, 0) Into n_结算模式;
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  End If;

  --病人担保信息
  If 病人id_In Is Not Null Then
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = d_Date
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, d_Date) >= d_Date;
    End If;
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
      Using 1, n_挂号id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_预约挂号接收_出诊_Insert;
/

--108159:梁唐彬,2017-04-12,医嘱执行修改登记报错
CREATE OR REPLACE Procedure Zl_病人医嘱执行_Update
( 
  原执行时间_In 病人医嘱执行.执行时间%Type, 
  医嘱id_In     病人医嘱执行.医嘱id%Type, 
  发送号_In     病人医嘱执行.发送号%Type, 
  要求时间_In   病人医嘱执行.要求时间%Type, 
  本次数次_In   病人医嘱执行.本次数次%Type, 
  执行摘要_In   病人医嘱执行.执行摘要%Type, 
  执行人_In     病人医嘱执行.执行人%Type, 
  执行时间_In   病人医嘱执行.执行时间%Type, 
  执行结果_In   病人医嘱执行.执行结果%Type := 1, 
  未执行原因_In 病人医嘱执行.说明%Type := Null, 
  单独执行_In   Number := 0, 
  操作员编号_In 人员表.编号%Type := Null, 
  操作员姓名_In 人员表.姓名%Type := Null, 
  执行部门id_In 门诊费用记录.执行部门id%Type := 0 
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。 
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门 
) Is 
  --除了要执行的主记录,还包含了附加手术,检查部位的记录 
  --手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同 
  V_Temp     Varchar2(255); 
  V_人员编号 人员表.编号%Type; 
  V_人员姓名 人员表.姓名%Type; 
 
  V_组id        病人医嘱记录.Id%Type; 
  V_诊疗类别    病人医嘱记录.诊疗类别%Type; 
  V_执行结果old 病人医嘱执行.执行结果%Type; 
  N_本次数次old 病人医嘱执行.本次数次%Type; 
 
  V_病人来源 病人医嘱记录.病人来源%Type; 
  V_费用性质 病人医嘱发送.记录性质%Type; 
 
  N_执行次数 Number; 
  N_剩余次数 Number; 
  N_执行状态 Number; 
  n_发送数次 Number;
  n_单次数次 Number;
  v_Count    Number;
  n_登记数次 Number;
  d_要求时间 date;
  
  D_登记时间 病人医嘱执行.登记时间%Type; 
  N_取消执行 Number; 
  N_Diffday  Number(18, 3); 
 
  V_Date  Date; 
  V_Error Varchar2(255); 
  Err_Custom Exception; 
Begin 
  --当前操作人员 
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then 
    V_人员编号 := 操作员编号_In; 
    V_人员姓名 := 操作员姓名_In; 
  Else 
    V_Temp     := Zl_Identity; 
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ';') + 1); 
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ',') + 1); 
    V_人员编号 := Substr(V_Temp, 1, Instr(V_Temp, ',') - 1); 
    V_人员姓名 := Substr(V_Temp, Instr(V_Temp, ',') + 1); 
  End If; 
 
  Select Sysdate Into V_Date From Dual; 
  Select Nvl(执行结果, 1), Nvl(本次数次, 0), 登记时间 
  Into V_执行结果old, N_本次数次old, D_登记时间 
  From 病人医嘱执行 
  Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 执行时间 = 原执行时间_In; 
  -----取消执行有效天数限制 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(220), '999')) Into N_取消执行 From Dual; 
  Select V_Date - D_登记时间 Into N_Diffday From Dual; 
  --登记时间超过取消执行天数的记录，不允许修改医嘱执行情况 
  If N_Diffday > N_取消执行 Then 
    V_Error := '医嘱执行登记时间超过了取消执行有效天数，不能修改医嘱执行情况！'; 
    Raise Err_Custom; 
  End If; 
  --病人医嘱执行 
  Update 病人医嘱执行 
  Set 要求时间 = 要求时间_In, 本次数次 = 本次数次_In, 执行摘要 = 执行摘要_In, 执行人 = 执行人_In, 执行时间 = 执行时间_In, 登记时间 = V_Date, 登记人 = V_人员姓名, 
      执行结果 = 执行结果_In, 说明 = 未执行原因_In 
  Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 执行时间 = 原执行时间_In; 
  --本次执行次数或这执行结果修改后需要更新单据的执行状态 
  If V_执行结果old <> 执行结果_In Or N_本次数次old <> 本次数次_In Then 
    Select 病人来源, Nvl(相关id, ID), 诊疗类别 
    Into V_病人来源, V_组id, V_诊疗类别 
    From 病人医嘱记录 
    Where ID = 医嘱id_In; 
 
    If v_病人来源 = 2 Then 
      Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2)) 
      Into v_费用性质 
      From 病人医嘱发送 
      Where 发送号 = 发送号_In And 医嘱id = 医嘱id_In; 
    Else 
      v_费用性质 := 1; 
    End If; 
   
    Select Decode(a.执行状态, 1, a.发送数次, c.登记次数), Decode(a.执行状态, 1, 0, a.发送数次 - c.登记次数) ,A.发送数次,C.登记次数
    Into n_执行次数, n_剩余次数 ,n_发送数次,n_登记数次
    From 病人医嘱发送 A, 
         (Select 医嘱id_In 医嘱id, 发送号_In 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数 
           From 病人医嘱执行 B 
           Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In And Nvl(b.执行结果, 1) <> 0) C 
    Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In; 
   
    --如果全部执行则状态为1，未执行状态为0，部分执行状态为2 
    Select Decode(N_剩余次数, 0, 1, Decode(N_执行次数, 0, 0, 2)) Into N_执行状态 From Dual; 
    
    --更新医嘱执行计价.执行状态
    If n_发送数次 > 0 Then
      Select Count(distinct 要求时间) Into v_Count From 医嘱执行计价 Where 医嘱ID = 医嘱ID_IN And 发送号 = 发送号_IN;
      If v_Count > 0 Then
        n_单次数次 := n_发送数次 / v_Count;
        --已执行数量+本次数次 总共能够执行多少个时间点,取最大整数
        v_Count := ceil((n_登记数次 ) / n_单次数次);
		If n_登记数次 = 0 Then
			Update 医嘱执行计价 Set 执行状态 = 0 Where 医嘱id = 医嘱ID_IN And 发送号 = 发送号_IN And NVL(执行状态,0) <> 2;
		Else
	        --获取执行截至要求时间 
	        Select 要求时间 Into d_要求时间
	        From (Select 要求时间, Rownum As 次数
	               From (Select Distinct 要求时间 From 医嘱执行计价 Where 医嘱id = 医嘱ID_IN And 发送号 = 发送号_IN Order By 要求时间))
	        Where 次数 = v_Count;
	        
	        If Not d_要求时间 Is Null Then
	          --先检查是否已经退费
	          Select Max(NVL(执行状态,0)) Into v_Count From 医嘱执行计价 Where 医嘱id = 医嘱ID_IN And 发送号 = 发送号_IN And 要求时间 <= d_要求时间;
	          If v_Count = 2 Then
	            v_Error := '您指定的执行时间段的医嘱费用已经被退费，不允许再执行。'; 
	            Raise Err_Custom; 
	          End If;
	          --更新截至要求时间之前(含)的记录执行状态；
	          Update 医嘱执行计价 Set 执行状态 = 1 Where 医嘱id = 医嘱ID_IN And 发送号 = 发送号_IN And 要求时间 <= d_要求时间 And NVL(执行状态,0) <> 2;
	          Update 医嘱执行计价 Set 执行状态 = 0 Where 医嘱id = 医嘱ID_IN And 发送号 = 发送号_IN And 要求时间 > d_要求时间 And NVL(执行状态,0) <> 2;
	        End If;
		End If;
      End If;
    End If;
 
    --执行次数不为0就标记为正在执行 
    If Nvl(单独执行_In, 0) = 1 Then 
      Update 病人医嘱发送 
      Set 执行状态 = Decode(N_执行次数, 0, 0, 3), 完成人 = Null, 完成时间 = Null 
      Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In; 
    Else 
      Update 病人医嘱发送 
      Set 执行状态 = Decode(N_执行次数, 0, 0, 3), 完成人 = Null, 完成时间 = Null 
      Where 执行状态 In (0, 3) And 发送号 + 0 = 发送号_In And 
            医嘱id In (Select ID From 病人医嘱记录 Where (ID = V_组id Or 相关id = V_组id) And 诊疗类别 = V_诊疗类别); 
    End If; 
 
    If V_费用性质 = 2 Then 
      If Nvl(单独执行_In, 0) = 1 Then 
        Update 住院费用记录 A 
        Set 执行状态 = N_执行状态, 执行人 = Decode(N_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(N_执行状态, 0, Null, 执行时间_In) 
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or A.执行部门id = 执行部门id_In) And Not Exists 
         (Select 1 From 材料特性 Where 材料id = A.收费细目id And 跟踪在用 = 1) And A.记录状态 In (0, 1, 3) And 
              (医嘱序号, NO, 记录性质) In 
              (Select 医嘱id, NO, 记录性质 
               From 病人医嘱发送 
               Where 执行状态 = Decode(N_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In); 
      Else 
        Update 住院费用记录 A 
        Set 执行状态 = N_执行状态, 执行人 = Decode(N_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(N_执行状态, 0, Null, 执行时间_In) 
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or A.执行部门id = 执行部门id_In) And Not Exists 
         (Select 1 From 材料特性 Where 材料id = A.收费细目id And 跟踪在用 = 1) And A.记录状态 In (0, 1, 3) And 
              (医嘱序号, NO, 记录性质) In 
              (Select 医嘱id, NO, 记录性质 
               From 病人医嘱发送 
               Where 执行状态 = Decode(N_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And 
                     医嘱id In 
                     (Select ID From 病人医嘱记录 Where (ID = V_组id Or 相关id = V_组id) And 诊疗类别 = V_诊疗类别)); 
      End If; 
    Else 
      If Nvl(单独执行_In, 0) = 1 Then 
        Update 门诊费用记录 A 
        Set 执行状态 = N_执行状态, 执行人 = Decode(N_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(N_执行状态, 0, Null, 执行时间_In) 
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or A.执行部门id = 执行部门id_In) And Not Exists 
         (Select 1 From 材料特性 Where 材料id = A.收费细目id And 跟踪在用 = 1) And A.记录状态 In (0, 1, 3) And 
              (医嘱序号, NO, 记录性质) In 
              (Select 医嘱id, NO, 记录性质 
               From 病人医嘱发送 
               Where 执行状态 = Decode(N_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In); 
      Else 
        Update 门诊费用记录 A 
        Set 执行状态 = N_执行状态, 执行人 = Decode(N_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(N_执行状态, 0, Null, 执行时间_In) 
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or A.执行部门id = 执行部门id_In) And Not Exists 
         (Select 1 From 材料特性 Where 材料id = A.收费细目id And 跟踪在用 = 1) And A.记录状态 In (0, 1, 3) And 
              (医嘱序号, NO, 记录性质) In 
              (Select 医嘱id, NO, 记录性质 
               From 病人医嘱发送 
               Where 执行状态 = Decode(N_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And 
                     医嘱id In 
                     (Select ID From 病人医嘱记录 Where (ID = V_组id Or 相关id = V_组id) And 诊疗类别 = V_诊疗类别)); 
      End If; 
    End If; 
  End If; 
Exception 
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Error || '[ZLSOFT]'); 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_病人医嘱执行_Update;
/

--107950:涂建华,2017-04-10,电子病历格式报告中的患者基本信息调整提示
CREATE OR REPLACE Procedure Zl_病人信息_基本信息调整_PACS
( 
  病人id_In 病人信息变动.病人id%Type, 
  就诊id_In Varchar2, --门诊病人为挂号ID;住院病人为主页ID;体检病人为体检单号 
  姓名_In   病人信息.姓名%Type, 
  性别_In   病人信息.性别%Type, 
  年龄_In   病人信息.年龄%Type, 
  场合_In   Number,--1-门诊;2-住院;3-体检 
  说明_Out  Out 病人信息变动.说明%type --出参 
) As 
  Cursor c_AdviceID1 is select a.id as 医嘱id from 病人医嘱记录 a, 影像检查记录 b 
                      where a.id=b.医嘱id and a.挂号单=(select NO from 病人挂号记录 where id=to_number(就诊id_In)) and a.相关id is null; --门诊 
  Cursor c_AdviceID2 is Select a.id as 医嘱id from 病人医嘱记录 a, 影像检查记录 b 
                      where a.id=b.医嘱id and a.病人id=病人id_In and a.主页id=to_number(就诊id_In) and a.相关id is null;     --住院 
  Cursor c_AdviceID3 is Select a.id as 医嘱id from 病人医嘱记录 a where a.挂号单=就诊id_In and a.病人id=病人id_In and a.病人来源=4;    --体检 
  Err_Custom Exception; 
  V_Error Varchar2(2000); 
  v_执行科室 部门表.名称%type;  --当前医嘱的执行科室 
  v_所有执行科室组合 Varchar2(2000);--多条医嘱时，所有的执行科室 
  v_项目名称 诊疗项目目录.名称%type;  --当前医嘱对应的项目名称 
  v_执行项目名称组合 Varchar2(2000);--多条医嘱是，所有的项目名称 
  n_Type Number(1);        --签名类型：1，数字签名、2，电子签名 
Begin 
  --当有电子签名时，所有记录不能修改 
  if 场合_In= 1 then  --门诊 
    For Row_Cols1 In c_AdviceID1 Loop 
      begin 
        select 签名类型 into n_Type 
        from(Select substr(对象属性,1,1) as 签名类型 From 电子病历内容 Where 对象类型= 8 And 文件ID=(Select 病历ID From 病人医嘱报告 Where 医嘱ID= Row_Cols1.医嘱id) order by 签名类型 desc) 
        where rownum=1; 
 
        select 名称 into v_项目名称 from 诊疗项目目录 where id=(select 诊疗项目id from 病人医嘱记录 where id=Row_Cols1.医嘱id); 
      Exception 
      When Others Then 
        null; 
      end; 
 
      if n_Type is not null then 
        if n_Type=1 then 
          v_执行项目名称组合:=v_执行项目名称组合||'、'||v_项目名称; 
        elsif n_Type=2 then 
          V_Error:='病人【'||姓名_In||'】的 '||v_项目名称||' 项目已进行过电子签名，不能进行病人信息修改操作！'; 
          Raise Err_Custom; 
        End if; 
      end if; 
    end loop; 
  elsif 场合_In=2 then  --住院 
    For Row_Cols2 In c_AdviceID2 Loop 
      begin 
        select 签名类型 into n_Type 
        from(Select substr(对象属性,1,1) as 签名类型 From 电子病历内容 Where 对象类型= 8 And 文件ID=(Select 病历ID From 病人医嘱报告 Where 医嘱ID= Row_Cols2.医嘱id) order by 签名类型 desc) 
        where rownum=1; 
 
        select 名称 into v_项目名称 from 诊疗项目目录 where id=(select 诊疗项目id from 病人医嘱记录 where id=Row_Cols2.医嘱id); 
      Exception 
      When Others Then 
        null; 
      end; 
 
      if n_Type is not null then 
        if n_Type=1 then 
          v_执行项目名称组合:=v_执行项目名称组合||'、'||v_项目名称; 
        elsif n_Type=2 then 
          V_Error:='病人【'||姓名_In||'】的 '||v_项目名称||' 项目已进行过电子签名，不能进行病人信息修改操作！'; 
          Raise Err_Custom; 
        End if; 
      end if; 
    end loop; 
  elsif 场合_In=3 then  --体检 
    For Row_Cols3 In c_AdviceID3 Loop 
      begin 
        select 签名类型 into n_Type 
        from(Select substr(对象属性,1,1) as 签名类型 From 电子病历内容 Where 对象类型= 8 And 文件ID=(Select 病历ID From 病人医嘱报告 Where 医嘱ID= Row_Cols3.医嘱id) order by 签名类型 desc) 
        where rownum=1; 
 
        select 名称 into v_项目名称 from 诊疗项目目录 where id=(select 诊疗项目id from 病人医嘱记录 where id=Row_Cols3.医嘱id); 
      Exception 
      When Others Then 
        null; 
      end; 
 
      if n_Type is not null then 
        if n_Type=1 then 
          v_执行项目名称组合:=v_执行项目名称组合||'、'||v_项目名称; 
        elsif n_Type=2 then 
          V_Error:='病人【'||姓名_In||'】的 '||v_项目名称||' 项目已进行过电子签名，不能进行病人信息修改操作！'; 
          Raise Err_Custom; 
        End if; 
      end if; 
    end loop; 
  end if; 
 
  --修改信息 
  if 场合_In= 1 then  --门诊 
    For Row_Cols1 In c_AdviceID1 Loop 
       Begin
         Update 影像检查记录 Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = Decode(年龄_In, Null, 年龄, 年龄_In) Where 医嘱id=Row_Cols1.医嘱id; 
 
         Select 名称 Into v_执行科室 from 部门表 where id=(select 执行科室id from 影像检查记录 where 医嘱id=Row_Cols1.医嘱id); 
       Exception 
         When Others Then 
           null; 
       End;

       if nvl(instr(v_所有执行科室组合,v_执行科室),0)<=0 then 
         v_所有执行科室组合:=v_所有执行科室组合||','||v_执行科室; 
       end if; 
    end loop; 
  elsif 场合_In=2 then  --住院 
    For Row_Cols2 In c_AdviceID2 Loop 
      Begin
        Update 影像检查记录 Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = Decode(年龄_In, Null, 年龄, 年龄_In) Where 医嘱id=Row_Cols2.医嘱id; 
 
        Select 名称 Into v_执行科室 from 部门表 where id=(select 执行科室id from 影像检查记录 where 医嘱id=Row_Cols2.医嘱id); 
      Exception 
        When Others Then 
          null; 
      End;

      if nvl(instr(v_所有执行科室组合,v_执行科室),0)<=0 then 
        v_所有执行科室组合:=v_所有执行科室组合||','||v_执行科室; 
      end if; 
    end loop; 
  elsif 场合_In=3 then --体检 
    For Row_Cols3 In c_AdviceID3 Loop 
      Begin
         Update 影像检查记录 Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = Decode(年龄_In, Null, 年龄, 年龄_In) Where 医嘱id=Row_Cols3.医嘱id; 
 
         Select 名称 Into v_执行科室 from 部门表 where id=(select 执行科室id from 影像检查记录 where 医嘱id=Row_Cols3.医嘱id); 
      Exception 
         When Others Then 
            null; 
      End;
      
      if nvl(instr(v_所有执行科室组合,v_执行科室),0)<=0 then 
        v_所有执行科室组合:=v_所有执行科室组合||','||v_执行科室; 
      end if; 
    end loop; 
  end if; 
 
  if nvl(v_执行项目名称组合,' ')<>' ' then 
     说明_Out:=substr(v_所有执行科室组合,2)||':【'||姓名_In||'】的【'|| substr(v_执行项目名称组合,2) ||'】对应检查报告已签名，需要手工调整！'; 
  else
     说明_Out:=substr(v_所有执行科室组合,2)||':【'||姓名_In||'】的基本信息已修改！'; 
  end if;
Exception 
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Error || '[ZLSOFT]'); 
  When Others Then 
    Zl_Errorcenter(Sqlcode, Sqlerrm); 
End Zl_病人信息_基本信息调整_PACS;
/

--106708:冉俊明,2017-04-07,违反规范，调整
Drop Procedure Zl_Buildregisterfixedrule;

--106708:冉俊明,2017-04-07,违反规范，调整
Create Or Replace Procedure Zl_临床出诊表_Addbyfixedrule
(
  Id_In         临床出诊表.Id%Type,
  Newid_In      临床出诊表.Id%Type,
  出诊表名_In   临床出诊表.出诊表名%Type,
  开始时间_In   临床出诊安排.开始时间%Type,
  终止时间_In   临床出诊安排.终止时间%Type,
  操作员姓名_In 临床出诊安排.操作员姓名%Type := Null,
  登记时间_In   临床出诊安排.登记时间%Type := Null,
  站点_In       部门表.站点%Type
) As
  -------------------------------------------------------------------------
  --功能：根据现有固定出诊表规则生成成新的固定出诊表
  -------------------------------------------------------------------------
  n_Count Number;

  n_出诊id 临床出诊表.Id%Type;

  v_操作员   临床出诊安排.操作员姓名%Type;
  d_登记时间 Date;
  v_Err_Msg  Varchar2(255);
  Err_Item Exception;
Begin
  Begin
    Select 1 Into n_Count From 临床出诊表 Where ID = Id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    v_Err_Msg := '未发现原出诊表信息！';
    Raise Err_Item;
  End If;

  --检查是否有有效号源
  Begin
    Select 1
    Into n_Count
    From 临床出诊号源 A, 部门表 B, 人员表 C, 收费项目目录 D
    Where a.科室id = b.Id And a.医生id = c.Id(+) And a.项目id = d.Id And a.排班方式 = 0 And Nvl(a.是否删除, 0) = 0 And
          (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd')
         --站点
          And (b.站点 Is Null Or b.站点 = 站点_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '当前出诊表中已无可按固定排班的号源，不能生成新的固定安排！';
    Raise Err_Item;
  End If;

  n_出诊id := Newid_In;
  If Nvl(n_出诊id, 0) = 0 Then
    Select 临床出诊表_Id.Nextval Into n_出诊id From Dual;
  End If;

  Insert Into 临床出诊表
    (ID, 排班方式, 出诊表名, 年份)
  Values
    (n_出诊id, 0, 出诊表名_In, To_Number(To_Char(开始时间_In, 'yyyy')));

  d_登记时间 := Nvl(登记时间_In, Sysdate);
  v_操作员   := Nvl(操作员姓名_In, Zl_Username);

  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, n_出诊id As 出诊id, 原安排id, 号源id, 项目id, 医生id, 医生姓名
               From (Select b.Id As 原安排id, b.号源id, c.项目id, c.医生id, c.医生姓名,
                             Row_Number() Over(Partition By c.Id Order By b.开始时间 Desc) As 组号
                      From 临床出诊安排 B, 临床出诊号源 C, 部门表 D, 人员表 E, 收费项目目录 F
                      Where b.号源id = c.Id And c.科室id = d.Id And c.医生id = e.Id(+) And c.项目id = f.Id And b.出诊id = Id_In
                           --号源限制
                            And c.排班方式 = 0 And Nvl(c.是否删除, 0) = 0 And
                            (c.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.撤档时间 Is Null) And
                            Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                            Nvl(e.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                            Nvl(f.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd')
                           --站点
                            And (d.站点 Is Null Or d.站点 = 站点_In)) M
               Where 组号 = 1) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, v_操作员, d_登记时间);
  
    --出诊限制
    For c_限制 In (Select ID, 安排id, 限制项目, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 分诊方式, 诊室id, 是否独占
                 From 临床出诊限制
                 Where 安排id = c_号源.原安排id) Loop
    
      Zl_临床出诊限制_Copy(c_限制.Id, c_号源.安排id);
    End Loop;
  End Loop;

  --加入没有的出诊安排的号源
  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, n_出诊id As 出诊id, a.Id As 号源id, a.项目id, a.医生id, a.医生姓名
               From 临床出诊号源 A, 部门表 D, 人员表 B, 收费项目目录 C
               Where a.科室id = d.Id And a.医生id = b.Id(+) And a.项目id = c.Id And a.排班方式 = 0 And Nvl(a.是否删除, 0) = 0 And
                     (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                     Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd')
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)
                    
                     And Not Exists (Select 1 From 临床出诊安排 Where 出诊id = n_出诊id And 号源id = a.Id)) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, v_操作员, d_登记时间);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊表_Addbyfixedrule;
/

--106708:冉俊明,2017-04-07,违反规范，调整
Drop Procedure Zl_Buildregisterplanbyrecord;

--106708:冉俊明,2017-04-07,违反规范，调整
Create Or Replace Procedure Zl_临床出诊表_Addbyrecord
(
  原出诊id_In   临床出诊表.Id%Type,
  新出诊id_In   临床出诊表.Id%Type,
  排班方式_In   临床出诊表.排班方式%Type,
  出诊表名_In   临床出诊表.出诊表名%Type,
  年份_In       临床出诊表.年份%Type,
  月份_In       临床出诊表.月份%Type,
  周数_In       临床出诊表.周数%Type,
  开始时间_In   临床出诊安排.开始时间%Type,
  终止时间_In   临床出诊安排.终止时间%Type,
  操作员姓名_In 临床出诊安排.操作员姓名%Type,
  登记时间_In   临床出诊安排.登记时间%Type,
  站点_In       部门表.站点%Type,
  人员id_In     人员表.Id%Type := Null,
  删除安排_In   Number := 0
) As
  -------------------------------------------------------------------------
  --功能：根据出诊记录生成新的出诊记录（月安排/周安排）
  --参数：
  --        人员id_In 除固定安排外有效，不为0或null表示临床科室人员在添加
  --        删除安排_In 固定排班转为月排班/周排班时，在制定月排班/周排班时是否删除新出诊表时间内未使用的出诊记录
  --说明：
  -------------------------------------------------------------------------
  n_Count Number;

  l_记录id t_Numlist := t_Numlist();
  n_安排id 临床出诊安排.Id%Type;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_跨月周出诊id 临床出诊表.Id%Type;

  Function Get跨月周出诊id(出诊id_In 临床出诊表.Id%Type) Return 临床出诊表.Id%Type Is
    ----------------------------------------
    --如果原周出诊表不是整周(不足7天)，则需要查找到另一个出诊表构成整周
    ----------------------------------------
    n_出诊id 临床出诊表.Id%Type;
    n_年份   临床出诊表.年份%Type;
    n_月份   临床出诊表.月份%Type;
    n_周数   临床出诊表.周数%Type;
  
    d_开始时间 临床出诊安排.开始时间%Type;
    d_结束时间 临床出诊安排.终止时间%Type;
  
    --根据日期计算当月的周数，以及每一周的时间范围
    Cursor c_Weekrange(Date_In Date) Is
      Select Rownum As 周数, 开始日期, 结束日期
      From (With Month_Range As (Select Trunc(Date_In) As First_Day, Last_Day(Trunc(Date_In)) As Last_Day From Dual)
             Select Decode(To_Char(First_Day, 'day'), '星期日', First_Day, Null) As 开始日期,
                    Decode(To_Char(First_Day, 'day'), '星期日', First_Day, Null) As 结束日期
             From Month_Range
             Union All
             Select Decode(Sign(Trunc(First_Day + 7 * Week, 'day') + 1 - First_Day), 1,
                            Trunc(First_Day + 7 * Week, 'day') + 1, First_Day) As 开始日期,
                    Decode(Sign(Trunc(First_Day + 7 * Week, 'day') + 7 - Last_Day), 1, Last_Day,
                            Trunc(First_Day + 7 * Week, 'day') + 7) As 结束日期
             From Month_Range A, (Select Level - 1 As Week From Dual Connect By Level <= 6) B)
             Where 开始日期 <= 结束日期;
  
  
  Begin
    Begin
      Select 年份, 月份, 周数 Into n_年份, n_月份, n_周数 From 临床出诊表 Where ID = 出诊id_In;
    Exception
      When Others Then
        Return 0;
    End;
  
    If n_年份 Is Null Or n_月份 Is Null Or n_周数 Is Null Then
      Return 0;
    End If;
  
    For r_Weekrange In c_Weekrange(To_Date(n_年份 || '-' || n_月份 || '-01', 'yyyy-mm-dd')) Loop
      If r_Weekrange.周数 = n_周数 Then
        d_开始时间 := r_Weekrange.开始日期;
        d_结束时间 := r_Weekrange.结束日期;
        Exit;
      End If;
    End Loop;
  
    If d_开始时间 Is Null Or d_结束时间 Is Null Then
      Return 0;
    End If;
    If Trunc(d_结束时间) - Trunc(d_开始时间) >= 6 Then
      Return 0;
    End If;
  
    --存在跨月的，查找另一个出诊表的年月周
    n_年份 := Null;
    n_月份 := Null;
    n_周数 := Null;
    If Trunc(d_开始时间 - 1, 'month') <> Trunc(d_开始时间, 'month') Then
      --当前是第一周,获取另一个出诊表的年月
      n_年份 := To_Number(To_Char(d_开始时间 - 1, 'yyyy'));
      n_月份 := To_Number(To_Char(d_开始时间 - 1, 'mm'));
    Elsif Trunc(d_结束时间 + 1, 'month') <> Trunc(d_结束时间, 'month') Then
      --当前是最后一周,获取另一个出诊表的年月
      n_年份 := To_Number(To_Char(d_结束时间 + 1, 'yyyy'));
      n_月份 := To_Number(To_Char(d_结束时间 + 1, 'mm'));
      n_周数 := 1;
    Else
      Return 0;
    End If;
  
    --获取跨月的另一个出诊表的ID
    Begin
      Select ID
      Into n_出诊id
      From (Select Rownum As 行号, ID
             From 临床出诊表
             Where Nvl(排班方式, 0) = 2 And 年份 = n_年份 And 月份 = n_月份 And (n_周数 Is Null Or 周数 = n_周数)
             Order By 周数 Desc)
      Where 行号 < 2;
    Exception
      When Others Then
        Return 0;
    End;
  
    Return n_出诊id;
  End;
Begin
  Begin
    Select 1
    Into n_Count
    From 临床出诊号源 A, 部门表 B, 人员表 C, 收费项目目录 D
    Where a.科室id = b.Id And a.医生id = c.Id(+) And a.项目id = d.Id
         --有效号源
          And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          (
          --月排班
           Nvl(排班方式_In, 0) = 1 And a.排班方式 = 1
          --周排班
           Or Nvl(排班方式_In, 0) = 2 And
           (
           --当前出诊表所在时间范围内不能有月排班
            a.排班方式 = 2 And Not Exists
            (Select 1
                From 临床出诊安排 P, 临床出诊表 Q
                Where p.出诊id = q.Id And p.号源id = a.Id And
                      Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
           --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
            Or a.排班方式 = 1 And Exists
            (Select 1
                From 临床出诊安排 P, 临床出诊表 Q
                Where p.出诊id = q.Id And p.号源id = a.Id And
                      Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
         --号源在该出诊表时间范围内无出诊记录
          And Not Exists
     (Select 1
           From 临床出诊记录 O, 临床出诊安排 P, 临床出诊表 Q
           Where o.安排id = p.Id And p.出诊id = q.Id And p.号源id = a.Id And o.出诊日期 Between 开始时间_In And 终止时间_In And
                 (q.排班方式 In (1, 2)
                 --原来为固定出诊安排
                 Or q.排班方式 = 0 And (Nvl(删除安排_In, 0) = 0 Or Nvl(删除安排_In, 0) = 1 And Exists
                  (Select 1 From 病人挂号记录 Where 出诊记录id = a.Id))))
         --当前人员可操作的号源
          And (Nvl(人员id_In, 0) = 0 Or
          (Nvl(a.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
         --站点
          And (b.站点 Is Null Or b.站点 = 站点_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    If Nvl(排班方式_In, 0) = 1 Then
      v_Err_Msg := '当前出诊表中已无可按月排班的号源，不能生成新的出诊表！';
    Else
      v_Err_Msg := '当前出诊表中已无可按周排班的号源，不能生成新的出诊表！';
    End If;
    Raise Err_Item;
  End If;

  --检查出诊表是否存在
  Begin
    Select 1 Into n_Count From 临床出诊表 Where ID = 新出诊id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    Insert Into 临床出诊表
      (ID, 排班方式, 出诊表名, 年份, 月份, 周数)
    Values
      (新出诊id_In, 排班方式_In, 出诊表名_In, 年份_In, 月份_In, 周数_In);
  End If;

  --如果当前出诊表时间范围内无挂号且无预约的出诊记录(固定安排)，则删除这部分出诊记录(在删除出诊表时可恢复)，
  --并修改固定安排的终止时间，程序中已询问
  If Nvl(删除安排_In, 0) = 1 Then
    For c_安排 In (Select b.Id As 安排id
                 From 临床出诊安排 B, 临床出诊表 C, 临床出诊号源 D
                 Where b.出诊id = c.Id And b.号源id = d.Id
                      --号源
                       And Nvl(d.是否删除, 0) = 0 And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                       Nvl(d.排班方式, 0) = 排班方式_In
                      --安排有被使用了的出诊记录
                       And c.排班方式 = 0 And b.终止时间 >= 开始时间_In And Not Exists
                  (Select 1
                        From 临床出诊记录 M, 病人挂号记录 N
                        Where m.安排id = b.Id And m.Id = n.出诊记录id And m.出诊日期 >= 开始时间_In)
                      --当前人员可操作的号源
                       And (Nvl(人员id_In, 0) = 0 Or (Nvl(d.是否临床排班, 0) = 1 And Exists
                        (Select 1 From 部门人员 Where 部门id = d.科室id And 人员id = 人员id_In)))) Loop
    
      For c_记录 In (Select ID As 记录id From 临床出诊记录 Where 安排id = c_安排.安排id And 出诊日期 >= 开始时间_In) Loop
        l_记录id.Extend();
        l_记录id(l_记录id.Count) := c_记录.记录id;
      End Loop;
    End Loop;
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
  
  End If;

  --如果原周出诊表不是整周(不足7天)，则需要查找到另一个出诊表构成整周
  If Nvl(排班方式_In, 0) = 2 Then
    n_跨月周出诊id := Get跨月周出诊id(原出诊id_In);
  End If;

  For c_号源 In (Select 新出诊id_In As 出诊id, b.Id As 原安排id, b.号源id, c.项目id, c.医生id, c.医生姓名
               From 临床出诊安排 B, 临床出诊号源 C, 部门表 D, 人员表 E, 收费项目目录 F
               Where b.号源id = c.Id And c.科室id = d.Id And b.医生id = e.Id(+) And c.项目id = f.Id And
                     (b.出诊id = 原出诊id_In Or b.出诊id = n_跨月周出诊id)
                    --有效号源
                     And Nvl(c.是否删除, 0) = 0 And (c.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.撤档时间 Is Null) And
                     Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(e.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(f.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (
                     --月排班
                      Nvl(排班方式_In, 0) = 1 And c.排班方式 = 1
                     -- 周排班
                      Or Nvl(排班方式_In, 0) = 2 And
                      (
                      --当前出诊表所在时间范围内不能有月排班
                       c.排班方式 = 2 And Not Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = c.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
                      --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
                       Or c.排班方式 = 1 And Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = c.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
                    --号源在该出诊表时间范围内无出诊记录
                     And Not Exists
                (Select 1
                      From 临床出诊记录 P
                      Where p.号源id = c.Id And p.出诊日期 Between 开始时间_In And 终止时间_In)
                    --当前人员可操作的号源
                     And (Nvl(人员id_In, 0) = 0 Or (Nvl(c.是否临床排班, 0) = 1 And Exists
                      (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)) Loop
  
    Begin
      Select ID Into n_安排id From 临床出诊安排 Where 出诊id = c_号源.出诊id And 号源id = c_号源.号源id;
    Exception
      When Others Then
        n_安排id := Null;
    End;
  
    If Nvl(n_安排id, 0) = 0 Then
      Select 临床出诊安排_Id.Nextval Into n_安排id From Dual;
    
      Insert Into 临床出诊安排
        (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间)
      Values
        (n_安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员姓名_In, 登记时间_In);
    End If;
  
    --出诊记录
    For c_记录 In (Select Decode(b.Id, Null, a.Id, b.Id) As ID, c.日期
                 From 临床出诊记录 A, 临床出诊记录 B,
                      (Select Trunc(开始时间_In) + Level - 1 As 日期
                        From Dual
                        Connect By Level <= Trunc(终止时间_In) - Trunc(开始时间_In) + 1) C
                 Where a.Id = b.相关id(+) And a.安排id = c_号源.原安排id And a.相关id Is Null And Nvl(a.是否临时出诊, 0) = 0
                      --月排班
                       And (Nvl(排班方式_In, 0) = 1 And To_Char(a.出诊日期, 'dd') = To_Char(c.日期, 'dd')
                       --周排班
                       Or Nvl(排班方式_In, 0) = 2 And To_Char(a.出诊日期, 'D') = To_Char(c.日期, 'D'))) Loop
      Zl_临床出诊记录_Copy(c_记录.Id, n_安排id, c_记录.日期, 操作员姓名_In, 登记时间_In);
    End Loop;
  End Loop;

  --加入没有的出诊安排的号源
  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, 新出诊id_In As 出诊id, a.Id As 号源id, a.项目id, a.医生id, a.医生姓名
               From 临床出诊号源 A, 部门表 D, 人员表 E, 收费项目目录 F
               Where a.科室id = d.Id And a.医生id = e.Id(+) And a.项目id = f.Id
                    --有效号源
                     And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.撤档时间 Is Null) And
                     Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(e.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(f.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (
                     --月排班
                      Nvl(排班方式_In, 0) = 1 And a.排班方式 = 1
                     -- 周排班
                      Or Nvl(排班方式_In, 0) = 2 And
                      (
                      --当前出诊表所在时间范围内不能有月排班
                       a.排班方式 = 2 And Not Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = a.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
                      --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
                       Or a.排班方式 = 1 And Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = a.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
                    --号源在该出诊表时间范围内无出诊记录
                     And Not Exists
                (Select 1
                      From 临床出诊记录 P
                      Where p.号源id = a.Id And p.出诊日期 Between 开始时间_In And 终止时间_In)
                    --当前人员可操作的号源
                     And (Nvl(人员id_In, 0) = 0 Or (Nvl(a.是否临床排班, 0) = 1 And Exists
                      (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)
                    
                     And Not Exists (Select 1 From 临床出诊安排 Where 出诊id = 新出诊id_In And 号源id = a.Id)) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员姓名_In, 登记时间_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊表_Addbyrecord;
/

--106708:冉俊明,2017-04-07,违反规范，调整
Drop Procedure Zl_Buildregisterplanbytemplet;

--106708:冉俊明,2017-04-07,违反规范，调整
Create Or Replace Procedure Zl_临床出诊表_Addbytemplet
(
  模板id_In   临床出诊表.Id%Type,
  人员id_In   人员表.Id%Type,
  出诊id_In   临床出诊表.Id%Type,
  排班方式_In 临床出诊表.排班方式%Type,
  出诊表名_In 临床出诊表.出诊表名%Type,
  年份_In     临床出诊表.年份%Type,
  月份_In     临床出诊表.月份%Type,
  周数_In     临床出诊表.周数%Type,
  开始时间_In 临床出诊安排.开始时间%Type,
  终止时间_In 临床出诊安排.终止时间%Type,
  操作员_In   临床出诊安排.操作员姓名%Type,
  登记时间_In 临床出诊安排.登记时间%Type,
  站点_In     部门表.站点%Type,
  删除安排_In Number := 0
) As
  -------------------------------------------------------------------------
  --功能说明：根据模板自动生成临床出诊记录
  --参数：
  --        人员id_In 除固定安排外有效，不为0或null表示临床科室人员在添加
  --        删除安排_In 固定排班转为月排班/周排班时，在制定月排班/周排班时是否删除新出诊表时间内未使用的出诊记录
  --说明：
  -------------------------------------------------------------------------
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  n_Count   Number(18);

  d_轮询日期 Date;
  n_轮询天数 Number;
  v_限制项目 临床出诊限制.限制项目%Type;

  n_是否出诊 Number(2);
  d_开始时间 临床出诊记录.开始时间%Type;

  l_记录id t_Numlist := t_Numlist();

  Procedure Isvisit
  (
    安排id_In       临床出诊安排.Id%Type,
    排班规则_In     临床出诊安排.排班规则%Type,
    出诊日期_In     临床出诊记录.出诊日期%Type,
    轮询开始时间_In 临床出诊安排.开始时间%Type,
    限制项目_In     Out 临床出诊限制.限制项目%Type,
    是否出诊_In     Out Number
  ) As
    --判断是否出诊，并获取出诊项目
    d_轮询日期 Date;
    n_轮询天数 Number;
  Begin
    是否出诊_In := 1;
    --检查这天是否出诊
    If 排班规则_In = 1 Then
      --星期排班
      Select Decode(To_Char(出诊日期_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',
                     Null)
      Into 限制项目_In
      From Dual;
      Select Count(1) Into n_Count From 临床出诊限制 Where 安排id = 安排id_In And 限制项目 = 限制项目_In;
      If Nvl(n_Count, 0) = 0 Then
        是否出诊_In := 0;
      End If;
    Elsif 排班规则_In = 2 Then
      --单日排班
      限制项目_In := '单日';
      If Mod(To_Number(To_Char(出诊日期_In, 'dd')), 2) <> 1 Then
        是否出诊_In := 0;
      End If;
    Elsif 排班规则_In = 3 Then
      --双日排班
      限制项目_In := '双日';
      If Mod(To_Number(To_Char(出诊日期_In, 'dd')), 2) <> 0 Then
        是否出诊_In := 0;
      End If;
    Elsif 排班规则_In = 4 Or 排班规则_In = 5 Then
      --4-月内轮循,5-轮循不限制
      If 排班规则_In = 4 Then
        d_轮询日期 := To_Date(To_Char(出诊日期_In, 'yyyy-mm') || To_Char(轮询开始时间_In, '-dd'), 'yyyy-mm-dd');
      Else
        d_轮询日期 := 轮询开始时间_In;
      End If;
      Begin
        Select To_Number(Substr(限制项目, 1, Instr(限制项目, '天') - 1))
        Into n_轮询天数
        From 临床出诊限制
        Where 安排id = 安排id_In And Rownum < 2;
      Exception
        When Others Then
          n_轮询天数 := 0;
      End;
      If Nvl(n_轮询天数, 0) > 0 Then
        限制项目_In := n_轮询天数 || '天';
        If Mod(Trunc(出诊日期_In) - Trunc(d_轮询日期), n_轮询天数 + 1) <> 0 Then
          是否出诊_In := 0;
        End If;
      End If;
    Elsif 排班规则_In = 6 Then
      --特定日期
      限制项目_In := To_Number(To_Char(出诊日期_In, 'dd')) || '日';
      Select Count(1) Into n_Count From 临床出诊限制 Where 安排id = 安排id_In And 限制项目 = 限制项目_In;
      If Nvl(n_Count, 0) = 0 Then
        是否出诊_In := 0;
      End If;
    End If;
  End;
Begin
  Begin
    Select 1
    Into n_Count
    From 临床出诊号源 A, 部门表 B, 人员表 C, 收费项目目录 D
    Where a.科室id = b.Id And a.医生id = c.Id(+) And a.项目id = d.Id
         --有效号源
          And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          (
          --月排班
           Nvl(排班方式_In, 0) = 1 And a.排班方式 = 1
          --周排班
           Or Nvl(排班方式_In, 0) = 2 And
           (
           --当前出诊表所在时间范围内不能有月排班
            a.排班方式 = 2 And Not Exists
            (Select 1
                From 临床出诊安排 P, 临床出诊表 Q
                Where p.出诊id = q.Id And p.号源id = a.Id And
                      Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
           --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
            Or a.排班方式 = 1 And Exists
            (Select 1
                From 临床出诊安排 P, 临床出诊表 Q
                Where p.出诊id = q.Id And p.号源id = a.Id And
                      Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
         --号源在该出诊表时间范围内无出诊记录
          And Not Exists
     (Select 1
           From 临床出诊记录 O, 临床出诊安排 P, 临床出诊表 Q
           Where o.安排id = p.Id And p.出诊id = q.Id And p.号源id = a.Id And o.出诊日期 Between 开始时间_In And 终止时间_In And
                 (q.排班方式 In (1, 2)
                 --原来为固定出诊安排
                 Or q.排班方式 = 0 And (Nvl(删除安排_In, 0) = 0 Or Nvl(删除安排_In, 0) = 1 And Exists
                  (Select 1 From 病人挂号记录 Where 出诊记录id = a.Id))))
         --当前人员可操作的号源
          And (Nvl(人员id_In, 0) = 0 Or
          (Nvl(a.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
         --站点
          And (b.站点 Is Null Or b.站点 = 站点_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    If Nvl(排班方式_In, 0) = 1 Then
      v_Err_Msg := '当前出诊表中已无可按月排班的号源，不能生成新的出诊表！';
    Else
      v_Err_Msg := '当前出诊表中已无可按周排班的号源，不能生成新的出诊表！';
    End If;
    Raise Err_Item;
  End If;

  --检查出诊表是否存在
  Begin
    Select 1 Into n_Count From 临床出诊表 Where ID = 出诊id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    Insert Into 临床出诊表
      (ID, 排班方式, 出诊表名, 年份, 月份, 周数)
    Values
      (出诊id_In, 排班方式_In, 出诊表名_In, 年份_In, 月份_In, 周数_In);
  End If;

  --如果当前出诊表时间范围内无挂号且无预约的出诊记录(固定安排)，则删除这部分出诊记录(在删除出诊表时可恢复)，
  --并修改固定安排的终止时间，程序中已询问
  If Nvl(删除安排_In, 0) = 1 Then
    For c_安排 In (Select b.Id As 安排id
                 From 临床出诊安排 B, 临床出诊表 C, 临床出诊号源 D
                 Where b.出诊id = c.Id And b.号源id = d.Id
                      --号源
                       And Nvl(d.是否删除, 0) = 0 And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                       Nvl(d.排班方式, 0) = 排班方式_In
                      --安排有被使用了的出诊记录
                       And c.排班方式 = 0 And b.终止时间 >= 开始时间_In And Not Exists
                  (Select 1
                        From 临床出诊记录 M, 病人挂号记录 N
                        Where m.安排id = b.Id And m.Id = n.出诊记录id And m.出诊日期 >= 开始时间_In)
                      --当前人员可操作的号源
                       And (Nvl(人员id_In, 0) = 0 Or (Nvl(d.是否临床排班, 0) = 1 And Exists
                        (Select 1 From 部门人员 Where 部门id = d.科室id And 人员id = 人员id_In)))) Loop
    
      For c_记录 In (Select ID As 记录id From 临床出诊记录 Where 安排id = c_安排.安排id And 出诊日期 >= 开始时间_In) Loop
        l_记录id.Extend();
        l_记录id(l_记录id.Count) := c_记录.记录id;
      End Loop;
    End Loop;
  
    Zl_临床出诊记录_Batchdelete(l_记录id);
  End If;

  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, 出诊id_In As 出诊id, b.Id As 原安排id, b.号源id, c.科室id, c.项目id, c.医生id, c.医生姓名,
                      b.排班规则, b.是否周六出诊, b.是否周日出诊, b.开始时间, c.号类, Nvl(d.站点, '-') As 站点
               From 临床出诊安排 B, 临床出诊号源 C, 部门表 D, 人员表 E, 收费项目目录 F
               Where b.号源id = c.Id And c.科室id = d.Id And c.医生id = e.Id(+) And c.项目id = f.Id And b.出诊id = 模板id_In
                    --有效号源
                     And Nvl(c.是否删除, 0) = 0 And (c.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.撤档时间 Is Null) And
                     Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(e.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(f.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (
                     --月排班
                      Nvl(排班方式_In, 0) = 1 And c.排班方式 = 1
                     -- 周排班
                      Or Nvl(排班方式_In, 0) = 2 And
                      (
                      --当前出诊表所在时间范围内不能有月排班
                       c.排班方式 = 2 And Not Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = c.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
                      --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
                       Or c.排班方式 = 1 And Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = c.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
                    --号源在该出诊表时间范围内无出诊记录
                     And Not Exists
                (Select 1
                      From 临床出诊记录 P
                      Where p.号源id = c.Id And p.出诊日期 Between 开始时间_In And 终止时间_In)
                    --当前人员可操作的号源
                     And (Nvl(人员id_In, 0) = 0 Or (Nvl(c.是否临床排班, 0) = 1 And Exists
                      (Select 1 From 部门人员 Where 部门id = c.科室id And 人员id = 人员id_In)))
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员_In, 登记时间_In);
  
    --临床出诊记录
    For c_日期 In (Select Trunc(开始时间_In) + Level - 1 As 日期,
                        Decode(To_Char(Trunc(开始时间_In) + Level - 1, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5',
                                '周四', '6', '周五', '7', '周六', Null) As 星期
                 From Dual
                 Connect By Level <= Trunc(终止时间_In) - Trunc(开始时间_In) + 1) Loop
    
      Isvisit(c_号源.原安排id, c_号源.排班规则, c_日期.日期, c_号源.开始时间, v_限制项目, n_是否出诊);
    
      --是否周六、周日不出诊
      --排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
      If Instr(',2,3,4,5,', c_号源.排班规则) > 0 And
         (Nvl(c_号源.是否周六出诊, 0) = 0 And c_日期.星期 = '周六' Or Nvl(c_号源.是否周日出诊, 0) = 0 And c_日期.星期 = '周日') Then
        n_是否出诊 := 0;
      End If;
    
      If Nvl(n_是否出诊, 0) = 1 Then
        For c_记录 In (With c_时间段 As
                        (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间
                        From (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间,
                                      Row_Number() Over(Partition By 时间段 Order By 时间段, 站点 Asc, 号类 Asc) As 组号
                               From 时间段
                               Where Nvl(站点, c_号源.站点) = c_号源.站点 And Nvl(号类, c_号源.号类) = c_号源.号类)
                        Where 组号 = 1)
                       Select 临床出诊记录_Id.Nextval As 记录id, m.Id As 限制id, m.上班时段,
                              To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.开始时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                              To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.终止时间, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.终止时间 <= j.开始时间 Then
                                  1
                                 Else
                                  0
                               End As 终止时间,
                              To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.缺省时间, j.开始时间), 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.缺省时间 < j.开始时间 Then
                                  1
                                 Else
                                  0
                               End As 缺省预约时间,
                              To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.提前时间, j.开始时间), 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.开始时间 < j.提前时间 Then
                                  -1
                                 Else
                                  0
                               End As 提前挂号时间, m.限号数, m.限约数, m.是否序号控制, m.是否分时段, m.预约控制, a.项目id, a.医生id, a.医生姓名, m.分诊方式,
                              m.诊室id, m.是否独占
                       From 临床出诊安排 A, 临床出诊限制 M, c_时间段 J
                       Where a.Id = m.安排id And m.上班时段 = j.时间段 And a.Id = c_号源.原安排id And m.限制项目 = v_限制项目) Loop
        
          Insert Into 临床出诊记录
            (ID, 安排id, 号源id, 出诊日期, 上班时段, 开始时间, 终止时间, 缺省预约时间, 提前挂号时间, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 项目id, 科室id, 医生id,
             医生姓名, 分诊方式, 诊室id, 登记人, 登记时间, 是否独占)
          Values
            (c_记录.记录id, c_号源.安排id, c_号源.号源id, c_日期.日期, c_记录.上班时段, c_记录.开始时间, c_记录.终止时间, c_记录.缺省预约时间, c_记录.提前挂号时间,
             c_记录.限号数, c_记录.限约数, c_记录.是否序号控制, c_记录.是否分时段, c_记录.预约控制, c_记录.项目id, c_号源.科室id, c_记录.医生id, c_记录.医生姓名,
             c_记录.分诊方式, c_记录.诊室id, 操作员_In, 登记时间_In, c_记录.是否独占);
        
          Begin
            Select 开始时间 Into d_开始时间 From 临床出诊时段 Where 限制id = c_记录.限制id And 序号 = 1;
          Exception
            When Others Then
              d_开始时间 := Null;
          End;
          --插入临床出诊序号控制
          If Nvl(c_记录.是否分时段, 0) = 1 And Nvl(c_记录.是否序号控制, 0) = 1 Then
            --分时段且启用序号控制，使用"预约顺序号"记录"是否预约"
            Insert Into 临床出诊序号控制
              (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 预约顺序号)
              Select c_记录.记录id, 序号,
                     To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                        When Trunc(开始时间) > Trunc(d_开始时间) Then
                         1
                        Else
                         0
                      End,
                     To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                        When Trunc(终止时间) > Trunc(d_开始时间) Then
                         1
                        Else
                         0
                      End, 限制数量, 是否预约, 是否预约
              From 临床出诊时段
              Where 限制id = c_记录.限制id;
          Else
            Insert Into 临床出诊序号控制
              (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约)
              Select c_记录.记录id, 序号,
                     To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                       When Trunc(开始时间) > Trunc(d_开始时间) Then
                        1
                       Else
                        0
                     End,
                     To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                       When Trunc(终止时间) > Trunc(d_开始时间) Then
                        1
                       Else
                        0
                     End, 限制数量, 是否预约
              From 临床出诊时段
              Where 限制id = c_记录.限制id;
          End If;
        
          --插入合作单位挂号控制记录
          Insert Into 临床出诊挂号控制记录
            (类型, 性质, 名称, 记录id, 序号, 控制方式, 数量)
            Select 类型, 性质, 名称, c_记录.记录id, 序号, 控制方式, 数量
            From 临床出诊挂号控制
            Where 限制id = c_记录.限制id;
        
          --插入临床出诊诊室记录
          Insert Into 临床出诊诊室记录
            (记录id, 诊室id)
            Select c_记录.记录id, 诊室id From 临床出诊诊室 Where 限制id = c_记录.限制id;
        End Loop;
      End If;
    End Loop;
  End Loop;

  --加入没有的出诊安排的号源
  For c_号源 In (Select 临床出诊安排_Id.Nextval As 安排id, 出诊id_In As 出诊id, a.Id As 号源id, a.项目id, a.医生id, a.医生姓名
               From 临床出诊号源 A, 部门表 D, 人员表 E, 收费项目目录 F
               Where a.科室id = d.Id And a.医生id = e.Id(+) And a.项目id = f.Id
                    --有效号源
                     And Nvl(a.是否删除, 0) = 0 And (a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.撤档时间 Is Null) And
                     Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(e.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(f.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (
                     --月排班
                      Nvl(排班方式_In, 0) = 1 And a.排班方式 = 1
                     -- 周排班
                      Or Nvl(排班方式_In, 0) = 2 And
                      (
                      --当前出诊表所在时间范围内不能有月排班
                       a.排班方式 = 2 And Not Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = a.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 1)
                      --当前已调整为了月排班,但是本月又用了周排班，则本月剩下的部分将继续按周进行排班
                       Or a.排班方式 = 1 And Exists
                       (Select 1
                           From 临床出诊安排 P, 临床出诊表 Q
                           Where p.出诊id = q.Id And p.号源id = a.Id And
                                 Not (p.终止时间 < Trunc(开始时间_In, 'MONTH') Or p.开始时间 > Last_Day(开始时间_In)) And q.排班方式 = 2)))
                    --号源在该出诊表时间范围内无出诊记录
                     And Not Exists
                (Select 1
                      From 临床出诊记录 P
                      Where p.号源id = a.Id And p.出诊日期 Between 开始时间_In And 终止时间_In)
                    --当前人员可操作的号源
                     And (Nvl(人员id_In, 0) = 0 Or (Nvl(a.是否临床排班, 0) = 1 And Exists
                      (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = 人员id_In)))
                    --站点
                     And (d.站点 Is Null Or d.站点 = 站点_In)
                    
                     And Not Exists (Select 1 From 临床出诊安排 Where 出诊id = 出诊id_In And 号源id = a.Id)) Loop
  
    Insert Into 临床出诊安排
      (ID, 出诊id, 号源id, 项目id, 医生id, 医生姓名, 开始时间, 终止时间, 操作员姓名, 登记时间)
    Values
      (c_号源.安排id, c_号源.出诊id, c_号源.号源id, c_号源.项目id, c_号源.医生id, c_号源.医生姓名, 开始时间_In, 终止时间_In, 操作员_In, 登记时间_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊表_Addbytemplet;
/

--106708:冉俊明,2017-04-07,违反规范，调整
Create Or Replace Procedure Zl_临床出诊记录_Stopvisit
(
  记录id_In     临床出诊停诊记录.记录id%Type,
  开始时间_In   临床出诊停诊记录.开始时间%Type := Null,
  终止时间_In   临床出诊停诊记录.终止时间%Type := Null,
  停诊原因_In   临床出诊停诊记录.停诊原因%Type := Null,
  操作员_In     临床出诊停诊记录.申请人%Type := Null,
  操作时间_In   临床出诊停诊记录.申请时间%Type := Null,
  取消停诊_In   Number := 0,
  是否不检查_In Number := 0
) As
  --功能：停诊或者取消停诊
  --入参：
  --       是否不检查_in 主要用于停用/启用号源时使用
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
  d_Cur   Date;

  v_号码 临床出诊号源.号码%Type;
Begin
  If Nvl(取消停诊_In, 0) = 0 Then
    --停诊
    If Nvl(是否不检查_In, 0) = 0 Then
      Select Count(1) Into n_Count From 临床出诊记录 A Where ID = 记录id_In And 停诊开始时间 Is Not Null;
      If Nvl(n_Count, 0) <> 0 Then
        v_Err_Msg := '当前安排已被他人进行了停诊，请刷新数据后查看！';
        Raise Err_Item;
      End If;
    
      If 开始时间_In <= Sysdate Then
        v_Err_Msg := '停诊时间的开始时间小于了当前时间，不能进行停诊操作！';
        Raise Err_Item;
      End If;
    End If;
  
    Insert Into 临床出诊停诊记录
      (ID, 记录id, 开始时间, 终止时间, 停诊原因, 申请人, 申请时间, 审批人, 审批时间, 登记人)
      Select 临床出诊停诊记录_Id.Nextval, 记录id_In, 开始时间_In, 终止时间_In, 停诊原因_In, Nvl(a.医生姓名, 操作员_In), 操作时间_In, 操作员_In, 操作时间_In,
             操作员_In
      From 临床出诊记录 A
      Where ID = 记录id_In;
  
    --保存原始临床出诊记录
    Select Count(1) Into n_Count From 临床出诊记录 Where 相关id = 记录id_In;
    If Nvl(n_Count, 0) = 0 Then
      For c_记录 In (Select ID, 安排id, To_Date('1900-01-01', 'yyyy-mm-dd') As 出诊日期, 登记人, 登记时间, 是否发布
                   From 临床出诊记录
                   Where ID = 记录id_In) Loop
        Zl_临床出诊记录_Copy(c_记录.Id, c_记录.安排id, c_记录.出诊日期, c_记录.登记人, c_记录.登记时间, c_记录.是否发布, c_记录.Id);
      End Loop;
    End If;
  
    Update 临床出诊记录
    Set 停诊开始时间 = 开始时间_In, 停诊终止时间 = 终止时间_In, 停诊原因 = 停诊原因_In
    Where ID = 记录id_In;
  
    --调整"临床出诊序号控制.是否停诊"为1
    Update 临床出诊序号控制 A
    Set 是否停诊 = 1
    Where 记录id = 记录id_In And 开始时间 Between 开始时间_In And 终止时间_In And Exists
     (Select 1 From 临床出诊记录 Where ID = a.记录id And Nvl(是否序号控制, 0) = 1 And Nvl(是否分时段, 0) = 1);
  
    Insert Into 病人服务信息记录
      (ID, 通知类型, 记录id, 挂号id, 号源id, 号码, 科室id, 项目id, 医生id, 医生姓名, 病人id, 登记人, 登记时间, 通知原因)
      Select 病人服务信息记录_Id.Nextval, 1, 记录id_In, 挂号id, 号源id, 号码, 科室id, 项目id, 医生id, 医生姓名, 病人id, 操作员_In, 操作时间_In,
             '医生' || 停诊原因_In || '，已停诊'
      From (Select b.Id As 挂号id, c.Id As 号源id, c.号码, c.科室id, a.项目id, a.医生id, a.医生姓名, b.病人id
             From 临床出诊记录 A, 病人挂号记录 B, 临床出诊号源 C
             Where a.Id = b.出诊记录id And a.号源id = c.Id And b.记录状态 = 1 And a.Id = 记录id_In And
                   (b.记录性质 = 1 And b.发生时间 Between a.停诊开始时间 And a.停诊终止时间 Or
                   b.记录性质 = 2 And b.预约时间 Between a.停诊开始时间 And a.停诊终止时间));
  
    --消息推送
    -- 停诊类型(1-停诊,2-取消停诊),出诊记录ID,停诊号码
    Begin
      Select b.号码 Into v_号码 From 临床出诊记录 A, 临床出诊号源 B Where a.号源id = b.Id And a.Id = 记录id_In;
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 17, 1 || ',' || 记录id_In || ',' || v_号码;
    Exception
      When Others Then
        Null;
    End;
  Else
    --取消停诊
    --数据检查
    Select Count(1) Into n_Count From 临床出诊记录 A Where ID = 记录id_In And 停诊开始时间 Is Null;
    If Nvl(n_Count, 0) <> 0 Then
      If Nvl(是否不检查_In, 0) = 1 Then
        Return;
      End If;
      v_Err_Msg := '当前安排已被他人取消停诊，请刷新数据后查看！';
      Raise Err_Item;
    End If;
  
    If Nvl(是否不检查_In, 0) = 0 Then
      Select 停诊终止时间 Into d_Cur From 临床出诊记录 Where ID = 记录id_In And 停诊开始时间 Is Not Null;
      If d_Cur <= Sysdate Then
        v_Err_Msg := '停诊时间的终止时间小于了当前时间，不能进行取消停诊操作！';
        Raise Err_Item;
      End If;
      Select Count(1)
      Into n_Count
      From 病人服务信息记录
      Where 记录id = 记录id_In And 通知类型 = 1 And 处理人 Is Not Null;
      If n_Count <> 0 Then
        v_Err_Msg := '该出诊记录存在病人服务信息记录，且已被处理，不允许取消停诊操作！';
        Raise Err_Item;
      End If;
    End If;
  
    Select Count(1)
    Into n_Count
    From (Select 开始时间, 停诊开始时间 As 终止时间
           From (Select a.开始时间, a.终止时间, a.停诊开始时间, a.停诊终止时间
                  From 临床出诊记录 A, 临床出诊记录 B
                  Where a.号源id = b.号源id And a.出诊日期 = b.出诊日期 And b.Id = 记录id_In And a.Id <> b.Id)
           Where 开始时间 < 停诊开始时间 And 终止时间 = 停诊终止时间
           Union All
           Select 停诊终止时间 As 开始时间, 终止时间
           From (Select a.开始时间, a.终止时间, a.停诊开始时间, a.停诊终止时间
                  From 临床出诊记录 A, 临床出诊记录 B
                  Where a.号源id = b.号源id And a.出诊日期 = b.出诊日期 And b.Id = 记录id_In And a.Id <> b.Id)
           Where 开始时间 = 停诊开始时间 And 终止时间 > 停诊终止时间
           Union All
           Select 开始时间, 停诊开始时间 As 终止时间
           From (Select a.开始时间, a.终止时间, a.停诊开始时间, a.停诊终止时间
                  From 临床出诊记录 A, 临床出诊记录 B
                  Where a.号源id = b.号源id And a.出诊日期 = b.出诊日期 And b.Id = 记录id_In And a.Id <> b.Id)
           Where 开始时间 < 停诊开始时间 And 终止时间 > 停诊终止时间
           Union All
           Select 停诊终止时间 As 开始时间, 终止时间
           From (Select a.开始时间, a.终止时间, a.停诊开始时间, a.停诊终止时间
                  From 临床出诊记录 A, 临床出诊记录 B
                  Where a.号源id = b.号源id And a.出诊日期 = b.出诊日期 And b.Id = 记录id_In And a.Id <> b.Id)
           Where 开始时间 < 停诊开始时间 And 终止时间 > 停诊终止时间) M, 临床出诊记录 N
    Where m.开始时间 < n.终止时间 And m.终止时间 > n.开始时间 And n.Id = 记录id_In And Rownum < 2;
    If n_Count <> 0 Then
      If Nvl(是否不检查_In, 0) = 1 Then
        Return;
      End If;
      v_Err_Msg := '当前上班时段的时间范围与该号源今日目前有效的上班时段的时间范围有交叉，你不能取消停诊！';
      Raise Err_Item;
    End If;
  
    Update 临床出诊停诊记录
    Set 取消人 = 操作员_In, 取消时间 = 操作时间_In
    Where 记录id = 记录id_In And 替诊医生姓名 Is Null And 取消人 Is Null;
  
    Update 临床出诊记录
    Set 停诊开始时间 = Null, 停诊终止时间 = Null, 停诊原因 = Null
    Where ID = 记录id_In And 停诊开始时间 Is Not Null;
  
    --调整"临床出诊序号控制.是否停诊"为0
    Update 临床出诊序号控制 Set 是否停诊 = 0 Where 记录id = 记录id_In And Nvl(是否停诊, 0) = 1;
  
    Delete 病人服务信息记录 Where 记录id = 记录id_In And 通知类型 = 1 And 处理人 Is Null;
  
    --消息推送
    -- 停诊类型(1-停诊,2-取消停诊),出诊记录ID,停诊号码
    Begin
      Select b.号码 Into v_号码 From 临床出诊记录 A, 临床出诊号源 B Where a.号源id = b.Id And a.Id = 记录id_In;
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 17, 2 || ',' || 记录id_In || ',' || v_号码;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊记录_Stopvisit;
/

--107559:冉俊明,2017-04-17,增加终止停诊安排功能
--106712:冉俊明,2017-04-07,SQL语句优化
Create Or Replace Procedure Zl_临床出诊停诊_Audit
(
  操作类型_In Number,
  Id_In       临床出诊停诊记录.Id%Type,
  审批人_In   临床出诊停诊记录.审批人%Type := Null,
  审批时间_In 临床出诊停诊记录.审批时间%Type := Null
) As
  --功能：审批停诊安排
  --参数：
  --       状态_In：1-审批，2-取消审批
  n_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If Nvl(操作类型_In, 0) = 1 Then
    --审批
    Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In And 审批人 Is Not Null;
    If n_Count <> 0 Then
      v_Error := '该申请已被审批，不能再次审批！';
      Raise Err_Custom;
    End If;
  
    Update 临床出诊停诊记录 Set 审批人 = 审批人_In, 审批时间 = 审批时间_In Where ID = Id_In;
    If Sql%NotFound Then
      v_Error := '该申请可能已被取消申请，请刷新后查看...';
      Raise Err_Custom;
    End If;
  
    --对出诊记录进行停诊标记
    For c_记录 In (Select a.Id, Greatest(a.开始时间, b.开始时间) As 停诊开始时间, Least(a.终止时间, b.终止时间) As 停诊终止时间, b.停诊原因, c.号码, a.是否序号控制,
                        a.是否分时段
                 From 临床出诊记录 A, 临床出诊停诊记录 B, 临床出诊号源 C
                 Where ((a.替诊医生姓名 Is Null And a.医生id Is Not Null And a.医生姓名 = b.申请人) Or
                       (a.替诊医生姓名 Is Not Null And a.替诊医生id Is Not Null And a.替诊医生姓名 = b.申请人)) And a.号源id = c.Id And
                       b.Id = Id_In And Not (a.开始时间 > b.终止时间 Or a.终止时间 < b.开始时间)
                      --只处理已发布了的
                       And Nvl(a.是否发布, 0) = 1) Loop
    
      Update 临床出诊记录
      Set 停诊开始时间 = c_记录.停诊开始时间, 停诊终止时间 = c_记录.停诊终止时间, 停诊原因 = c_记录.停诊原因
      Where ID = c_记录.Id;
    
      --调整"临床出诊序号控制.是否停诊"为1
      Update 临床出诊序号控制 A
      Set 是否停诊 = 1
      Where 记录id = c_记录.Id And 开始时间 Between c_记录.停诊开始时间 And c_记录.停诊终止时间 And Nvl(c_记录.是否序号控制, 0) = 1 And
            Nvl(c_记录.是否分时段, 0) = 1;
    
      Insert Into 病人服务信息记录
        (ID, 通知类型, 记录id, 挂号id, 号源id, 号码, 科室id, 项目id, 医生id, 医生姓名, 病人id, 登记人, 登记时间)
        Select 病人服务信息记录_Id.Nextval, 1, a.Id, b.Id, c.Id, c.号码, c.科室id, a.项目id, a.医生id, a.医生姓名, b.病人id, 审批人_In, 审批时间_In
        From 临床出诊记录 A, 病人挂号记录 B, 临床出诊号源 C
        Where a.Id = b.出诊记录id And a.号源id = c.Id And b.记录状态 = 1 And a.Id = c_记录.Id And
              (b.记录性质 = 1 And b.发生时间 Between a.停诊开始时间 And a.停诊终止时间 Or
              b.记录性质 = 2 And b.预约时间 Between a.停诊开始时间 And a.停诊终止时间);
    
      --消息推送
      -- 停诊类型(1-停诊,2-取消停诊),出诊记录ID,停诊号码
      Begin
        Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
          Using 17, 1 || ',' || c_记录.Id || ',' || c_记录.号码;
      Exception
        When Others Then
          Null;
      End;
    End Loop;
    Return;
  End If;

  --取消审批
  Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In And 终止时间 < Sysdate;
  If n_Count <> 0 Then
    v_Error := '该停诊安排已失效，不能取消审批！';
    Raise Err_Custom;
  End If;

  Select Count(1) Into n_Count From 临床出诊停诊记录 Where ID = Id_In And 失效时间 Is Not Null;
  If n_Count <> 0 Then
    v_Error := '该停诊安排已被终止，不能取消审批！';
    Raise Err_Custom;
  End If;

  Select Count(1)
  Into n_Count
  From 临床出诊记录 A, 临床出诊停诊记录 B, 病人服务信息记录 C
  Where Nvl(a.替诊医生姓名, a.医生姓名) = b.申请人 And Nvl(a.替诊医生id, a.医生id) Is Not Null And a.Id = c.记录id And
        (a.开始时间 Between b.开始时间 And b.终止时间 Or a.终止时间 Between b.开始时间 And b.终止时间) And c.处理人 Is Not Null And b.Id = Id_In;
  If Nvl(n_Count, 0) <> 0 Then
    v_Error := '该停诊安排的部分停诊信息已被处理，不能取消审批！';
    Raise Err_Custom;
  End If;

  Update 临床出诊停诊记录 Set 审批人 = Null, 审批时间 = Null Where ID = Id_In And 审批时间 Is Not Null;
  If Sql%NotFound Then
    v_Error := '该安排可能已被他人取消审批，请刷新后查看...';
    Raise Err_Custom;
  End If;

  For c_记录 In (Select a.Id, c.号码
               From 临床出诊记录 A, 临床出诊停诊记录 B, 临床出诊号源 C
               Where ((a.替诊医生姓名 Is Null And a.医生id Is Not Null And a.医生姓名 = b.申请人) Or
                     (a.替诊医生姓名 Is Not Null And a.替诊医生id Is Not Null And a.替诊医生姓名 = b.申请人)) And a.号源id = c.Id And
                     b.Id = Id_In And (a.开始时间 Between b.开始时间 And b.终止时间 Or a.终止时间 Between b.开始时间 And b.终止时间) And
                     Nvl(a.是否发布, 0) = 1) Loop
  
    Update 临床出诊记录 Set 停诊开始时间 = Null, 停诊终止时间 = Null, 停诊原因 = Null Where ID = c_记录.Id;
  
    --调整"临床出诊序号控制.是否停诊"为0
    Update 临床出诊序号控制 Set 是否停诊 = 0 Where 记录id = c_记录.Id And Nvl(是否停诊, 0) = 1;
  
    Delete 病人服务信息记录 Where 记录id = c_记录.Id And 通知类型 = 1 And 处理人 Is Null;
  
    --消息推送
    -- 停诊类型(1-停诊,2-取消停诊),出诊记录ID,停诊号码
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 17, 2 || ',' || c_记录.Id || ',' || c_记录.号码;
    Exception
      When Others Then
        Null;
    End;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊停诊_Audit;
/

--106556:涂建华,2017-04-06,将组句内容参数类型调整为xml的方式进行处理
--影像报告组句管理(---定义部分---)***************************************************
CREATE OR REPLACE Package b_PACS_RptCombo Is
  --Create By Hwei;
  --2014/11/25
  Type t_Refcur Is Ref Cursor;

  --1.功  能：获得影像报告组句列表
  Procedure p_GetComboList(
    Val Out t_Refcur
	);
  --2.功  能：添加影像报告组句信息
  Procedure p_AddComboInfo(
    ID_In     In 影像报告组句清单.ID%Type,
    编码_In   In 影像报告组句清单.编码%Type,
    名称_In   In 影像报告组句清单.名称%Type,
    说明_In   In 影像报告组句清单.说明%Type,
    分组_In   In 影像报告组句清单.分组%Type,
    多组_In   In 影像报告组句清单.多组%Type,
    组成_In   In 影像报告组句清单.组成%Type,
    编辑人_In In 影像报告组句清单.编辑人%Type
	);
  --3.功  能;修改影像报告组句信息
  Procedure p_EditComboInfo(
    ID_In     In 影像报告组句清单.ID%Type,
    编码_In   In 影像报告组句清单.编码%Type,
    名称_In   In 影像报告组句清单.名称%Type,
    说明_In   In 影像报告组句清单.说明%Type,
    分组_In   In 影像报告组句清单.分组%Type,
    多组_In   In 影像报告组句清单.多组%Type,
    组成_In   In 影像报告组句清单.组成%Type,
    编辑人_In In 影像报告组句清单.编辑人%Type
	);
  --4.功  能：通过ID删除影像报告组句信息
  Procedure p_DelComboInfo(
    ID_In In 影像报告组句清单.ID%Type
	);
  --5.功  能：根据ID获得影像报告组句信息
  Procedure p_GetComboInfoByID(
	Val           Out t_Refcur,
	ID_In In 影像报告组句清单.ID%Type
	);
  --6.功  能：获得影像报告组句的所有分组信息
  Procedure p_GetComboAllGroup(
    Val Out t_Refcur
	);
  --7.功  能：获得ID对应的影像报告组句的短语信息
  Procedure p_GetComboContent(
	Val           Out t_Refcur,
	ID_In In 影像报告组句清单.ID%Type
	);
  --8.功  能：更新ID对应的影像报告组句的短语信息
  Procedure p_EditComboContent(
	ID_In   In 影像报告组句清单.ID%Type,
	组成_In  In 影像报告组句清单.组成%Type
	);
  --9.功 能：获取编辑人对应的最后修改影像报告组句信息
  Procedure p_GetComboInfoByEditor(
	Val           Out t_Refcur,
	编辑人_In In 影像报告组句清单.编辑人%Type
	);
  --10.功  能：新增片段到组合句
  Procedure p_Append_Fragment_Tocombo(
    Text_In In XmlType,
    Id_In   In 影像报告组句清单.ID%Type
	);

  --11.功  能：修改片段到组合句
  Procedure p_Update_Combo_Fragment(
    Text_In In XmlType,
    Id_In   In 影像报告组句清单.ID%Type,
    Pid_In  In Varchar2
	);
  --12.功  能：根据分类ID查询词句
  Procedure p_Get_Fragment_By_Typeid(
	Val           Out t_Refcur,
	Id_In In 影像报告组句清单.ID%Type
	);
  --13.功  能：获取下一个编码
  Procedure p_Get_ComboNextCode(
    Val Out t_Refcur
	);
end b_PACS_RptCombo;
/

--影像报告组句管理(---实现部分---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptCombo Is
  --Create By Hwei;
  --2014/11/25

  --1.功  能：获得影像报告组句列表
  Procedure p_GetComboList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             编码,
             名称,
             说明,
             分组,
             多组,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             编辑人,
             最后编辑时间
        From 影像报告组句清单 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboList;

  --2.功  能：添加影像报告组句信息
  Procedure p_AddComboInfo(
    ID_In     In 影像报告组句清单.ID%Type,
    编码_In   In 影像报告组句清单.编码%Type,
    名称_In   In 影像报告组句清单.名称%Type,
    说明_In   In 影像报告组句清单.说明%Type,
    分组_In   In 影像报告组句清单.分组%Type,
    多组_In   In 影像报告组句清单.多组%Type,
    组成_In   In 影像报告组句清单.组成%Type,
    编辑人_In In 影像报告组句清单.编辑人%Type
	) As
  Begin
    Insert Into 影像报告组句清单
      (ID, 编码, 名称, 说明, 分组, 多组, 组成, 编辑人, 最后编辑时间)
    Values
      (ID_In, 编码_In, 名称_In, 说明_In, 分组_In, 多组_In, 组成_In, 编辑人_In, Sysdate);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddComboInfo;

  --3.功  能;修改影像报告组句信息
  Procedure p_EditComboInfo(
    ID_In     In 影像报告组句清单.ID%Type,
    编码_In   In 影像报告组句清单.编码%Type,
    名称_In   In 影像报告组句清单.名称%Type,
    说明_In   In 影像报告组句清单.说明%Type,
    分组_In   In 影像报告组句清单.分组%Type,
    多组_In   In 影像报告组句清单.多组%Type,
    组成_In   In 影像报告组句清单.组成%Type,
    编辑人_In In 影像报告组句清单.编辑人%Type
	) As
  Begin
    Update 影像报告组句清单
       set 编码         = 编码_In,
           名称         = 名称_In,
           说明         = 说明_In,
           分组         = 分组_In,
           多组         = 多组_In,
           组成         = 组成_In,
           编辑人       = 编辑人_In,
           最后编辑时间 = SysDate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditComboInfo;

  --4.功  能：通过ID删除影像报告组句信息
  Procedure p_DelComboInfo(
    ID_In In 影像报告组句清单.ID%Type
	) As
  Begin
    Delete From 影像报告组句清单 Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelComboInfo;

  --5.功  能：根据ID获得影像报告组句信息
  Procedure p_GetComboInfoByID(
	Val           Out t_Refcur,
	ID_In In 影像报告组句清单.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             编码,
             名称,
             说明,
             分组,
             多组,
             (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             编辑人,
             最后编辑时间
        From 影像报告组句清单 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboInfoByID;

  --6.功  能：获得影像报告组句的所有分组信息
  Procedure p_GetComboAllGroup(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Distinct 分组 From 影像报告组句清单;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboAllGroup;

  --7.功  能：获得ID对应的影像报告组句的短语信息
  Procedure p_GetComboContent(
	Val           Out t_Refcur,
	ID_In In 影像报告组句清单.ID%Type
	) As
  Begin
    Open Val For
      Select (Nvl(t.组成, XmlType('<NULL/>'))).GetClobVal() As 组成
        From 影像报告组句清单 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboContent;

  --8.功  能：更新ID对应的影像报告组句的短语信息
  Procedure p_EditComboContent(
    ID_In   In 影像报告组句清单.ID%Type,
    组成_In In 影像报告组句清单.组成%Type
	) As
  Begin
    Update 影像报告组句清单 Set 组成 = 组成_In Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditComboContent;

  --9.功 能：获取编辑人对应的最后修改影像报告组句信息
  Procedure p_GetComboInfoByEditor(
	Val           Out t_Refcur,
	编辑人_In In 影像报告组句清单.编辑人%Type
	) AS
  Begin
    Open Val For
      Select RawToHex(ID) ID, 编辑人, 最后编辑时间
        From 影像报告组句清单 t1
       Where Not Exists (Select 1
                From 影像报告组句清单
               Where 最后编辑时间 > t1.最后编辑时间)
         And 编辑人 = 编辑人_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboInfoByEditor;

  --10.功  能：新增片段到组合句
  Procedure p_Append_Fragment_Tocombo(
    Text_In In XmlType,
	Id_In   In 影像报告组句清单.ID%Type
	) As
  Begin
    Update 影像报告组句清单 A
       Set a.组成 = Appendchildxml(a.组成, '/root', Text_In)
     Where a.ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Append_Fragment_Tocombo;

  --11.功  能：修改片段到组合句
  Procedure p_Update_Combo_Fragment(
    Text_In In XmlType,
    Id_In   In 影像报告组句清单.ID%Type,
    Pid_In  In Varchar2
	) As
  Begin
    Update 影像报告组句清单 A
       Set a.组成 = Updatexml(a.组成,
                            '/root/sentence[@sid="' || Pid_In || '"]',
                            Text_In)
     Where a.ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Combo_Fragment;

  --12.功  能：根据分类ID查询词句
  Procedure p_Get_Fragment_By_Typeid(
	Val           Out t_Refcur,
	Id_In In 影像报告组句清单.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             上级id,
             编码,
             名称,
             说明,
             节点类型,
             (Nvl(a.组成, XmlType('<NULL/>'))).GetClobVal() As 组成,
             学科,
             标签,
             是否私有,
             作者,
			 (Nvl(a.适应条件, XmlType('<NULL/>'))).GetClobVal() As 适应条件, 
             最后编辑时间
        From 影像报告片段清单 A
       Where a.上级id = Id_In
         And a.节点类型 <> 0;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Fragment_By_Typeid;

  --13.功  能：获取下一个编码
  Procedure p_Get_ComboNextCode(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select b_pacs_rptpublic.f_Get_Nextcode('影像报告组句清单') As 编码
        From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ComboNextCode;
End b_PACS_RptCombo;
/

--107584:冉俊明,2017-04-06,修正跨天的上班时段序号分配时日期规则错误
Create Or Replace Procedure Zl_临床出诊号源限制_Modify
(
  Id_In           临床出诊号源限制.Id%Type,
  号源id_In       临床出诊号源限制.号源id%Type,
  上班时段_In     临床出诊号源限制.上班时段%Type,
  限号数_In       临床出诊号源限制.限号数%Type,
  限约数_In       临床出诊号源限制.限约数%Type,
  是否序号控制_In 临床出诊号源限制.是否序号控制%Type,
  是否分时段_In   临床出诊号源限制.是否分时段%Type,
  预约控制_In     临床出诊号源限制.预约控制%Type,
  是否独占_In     临床出诊号源限制.是否独占%Type,
  分诊方式_In     临床出诊号源限制.分诊方式%Type,
  诊室id_In       临床出诊号源限制.诊室id%Type,
  号源诊室_In     Varchar2 := Null,
  号源时段_In     Varchar2 := Null,
  号源控制_In     Varchar2 := Null,
  删除号源限制_In Integer := 0
) As
  --号源时段_IN:序号,开始时间(HH:MM:SS),终止时(HH:MM:SS)间,数量,是否预约|...
  --号源诊室_IN:诊室id1,诊室id2,....
  --号源控制_IN:类型,性质,名称,控制方式,序号,数量|
  --删除号源限制_in:1-插入数据前，先删除号源限制,0-不删除数据，直接插入,-1-仅删除号源限制,不插入数据
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  l_限制id t_Numlist := t_Numlist();
  n_Count  Number;

  n_序号     临床出诊号源时段.序号%Type;
  d_开始时间 临床出诊号源时段.开始时间%Type;
  d_终止时间 临床出诊号源时段.终止时间%Type;
  n_数量     临床出诊号源时段.限制数量%Type;
  n_是否预约 临床出诊号源时段.是否预约%Type;

  n_类型     临床出诊号源控制.类型%Type;
  n_性质     临床出诊号源控制.性质%Type;
  v_名称     临床出诊号源控制.名称%Type;
  n_控制方式 临床出诊号源控制.控制方式%Type;
  n_限制数量 临床出诊号源控制.数量%Type;
Begin
  If Nvl(删除号源限制_In, 0) = 1 Or Nvl(删除号源限制_In, 0) = -1 Then
    Select ID Bulk Collect Into l_限制id From 临床出诊号源限制 Where 号源id = 号源id_In;
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源时段 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源控制 Where 限制id = l_限制id(I);
  
    Forall I In 1 .. l_限制id.Count
      Delete 临床出诊号源诊室 Where 限制id = l_限制id(I);
  
    Delete 临床出诊号源限制 Where 号源id = 号源id_In;
    Delete From 临床出诊号源限制 Where 号源id = 号源id_In;
  
    If Nvl(删除号源限制_In, 0) = -1 Then
      Return;
    End If;
  End If;

  Select Count(1) Into n_Count From 临床出诊号源限制 Where ID = Id_In;
  If n_Count = 0 Then
    Insert Into 临床出诊号源限制
      (ID, 号源id, 上班时段, 限号数, 限约数, 是否序号控制, 是否分时段, 预约控制, 是否独占, 分诊方式, 诊室id)
    Values
      (Id_In, 号源id_In, 上班时段_In, 限号数_In, 限约数_In, 是否序号控制_In, 是否分时段_In, 预约控制_In, 是否独占_In, 分诊方式_In, 诊室id_In);
  
  End If;

  If 号源时段_In Is Not Null Then
    --插入号源缺省时间段
    For c_时间段集 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(号源时段_In, '|'))) Loop
      n_序号     := Null;
      n_数量     := Null;
      n_是否预约 := Null;
      For c_时间段 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(c_时间段集.值)) Order By 序号) Loop
        If c_时间段.序号 = 1 Then
          n_序号 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 2 Then
          d_开始时间 := To_Date(c_时间段.值, 'yyyy-mm-dd hh24:mi:ss');
        End If;
      
        If c_时间段.序号 = 3 Then
          d_终止时间 := To_Date(c_时间段.值, 'yyyy-mm-dd hh24:mi:ss');
        End If;
      
        If c_时间段.序号 = 4 Then
          n_数量 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 5 Then
          n_是否预约 := To_Number(c_时间段.值);
        End If;
      
      End Loop;
    
      If Nvl(n_序号, 0) <> 0 Then
        Insert Into 临床出诊号源时段
          (限制id, 序号, 开始时间, 终止时间, 限制数量, 是否预约)
        Values
          (Id_In, n_序号, d_开始时间, d_终止时间, n_数量, n_是否预约);
      End If;
    End Loop;
  
  End If;

  --插入号源的缺省控制
  --号源控制_IN:类型,性质,名称,控制方式,序号,数量|
  If 号源控制_In Is Not Null Then
    For c_时间段集 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(号源控制_In, '|'))) Loop
      n_类型     := Null;
      n_性质     := Null;
      v_名称     := Null;
      n_序号     := Null;
      n_控制方式 := Null;
      n_限制数量 := Null;
    
      --类型,性质,名称,控制方式,序号,数量|
      For c_时间段 In (Select Rownum As 序号, Column_Value As 值 From Table(f_Str2list(c_时间段集.值)) Order By 序号) Loop
        If c_时间段.序号 = 1 Then
          n_类型 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 2 Then
          n_性质 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 3 Then
          v_名称 := c_时间段.值;
        End If;
      
        If c_时间段.序号 = 4 Then
          n_控制方式 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 5 Then
          n_序号 := To_Number(c_时间段.值);
        End If;
      
        If c_时间段.序号 = 6 Then
          n_限制数量 := To_Number(c_时间段.值);
        End If;
      
      End Loop;
    
      If v_名称 Is Not Null Then
        Insert Into 临床出诊号源控制
          (限制id, 类型, 性质, 名称, 序号, 控制方式, 数量)
        Values
          (Id_In, n_类型, n_性质, v_名称, n_序号, n_控制方式, n_限制数量);
      
      End If;
    End Loop;
  End If;
  --插入号源诊室
  If 号源诊室_In Is Not Null Then
    Insert Into 临床出诊号源诊室
      (限制id, 诊室id)
      Select Id_In As 限制id, Column_Value As 科室id From Table(f_Num2list(号源诊室_In));
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床出诊号源限制_Modify;
/

--108432:冉俊明,2017-05-08,修正固定出诊表临时安排取消审核后没有删除由该临时安排生成的出诊记录，导致清除该临时安排时报错的问题
--107584:冉俊明,2017-04-06,修正跨天的上班时段序号分配时日期规则错误
Create Or Replace Procedure Zl1_Auto_Buildingregisterplan
(
  挂号时间_In In Date := Null,
  号源id_In   临床出诊号源.Id%Type := Null
) As
  -------------------------------------------------------------------------
  --功能说明：自动生成临床出诊记录
  --          1、根据号源自动生成预约数内的临床出诊记录;
  --          2、预约天数的确定:号源预约天数-->预约方式的天数（取最大)-->系统预约天数
  --入参:挂号时间_IN:NULL时，自动生成;否则只检查指定日期是否生成了出诊记录没有
  --    号源id_In:NULL时处理所有号源，否则只处理指定号源
  -------------------------------------------------------------------------
  n_缺省预约天数 临床出诊号源.预约天数%Type;
  v_操作员姓名   临床出诊安排.操作员姓名%Type;
  d_登记日期     临床出诊安排.登记时间%Type;
  n_安排id       临床出诊安排.Id%Type;
  n_项目id       临床出诊安排.项目id %Type;

  n_记录id   临床出诊记录.Id%Type;
  d_当前日期 临床出诊记录.出诊日期%Type;

  n_是否出诊 Number(2);
  l_固定时段 t_Strlist := t_Strlist();
  n_Count    Number(18);

  n_加预约天数 Number := 0;
  d_开始时间   临床出诊记录.开始时间%Type;
Begin

  Select Max(预约天数) Into n_缺省预约天数 From 预约方式;
  If Nvl(n_缺省预约天数, 0) = 0 Then
    n_缺省预约天数 := To_Number(Nvl(zl_GetSysParameter('挂号允许预约天数'), '0'));
  End If;
  If Nvl(n_缺省预约天数, 0) = 0 Then
    n_缺省预约天数 := 7;
  End If;

  --以半天为单位,如果参数“号源开放时间”在12:00:00-23:59:59期间的，则开放预约天数+1天
  n_加预约天数 := Zl_Fun_Getappointmentdays;

  d_当前日期   := Trunc(Nvl(挂号时间_In, Sysdate));
  d_登记日期   := Sysdate;
  v_操作员姓名 := Zl_Username;

  --第一层循环，号源信息
  For c_号源 In (Select c.Id, c.号类, c.号码, c.项目id, c.科室id, c.医生姓名,
                      Decode(Nvl(c.预约天数, 0), 0, n_缺省预约天数, c.预约天数) + n_加预约天数 As 预约天数, Nvl(b.站点, '-') As 站点,
                      Nvl(c.是否假日换休, 0) As 是否假日换休, Nvl(c.假日控制状态, 0) As 假日控制状态, Nvl(c.排班方式, 0) As 排班方式
               From 临床出诊号源 C, 部门表 B, 人员表 A, 收费项目目录 D
               Where c.科室id = b.Id And c.医生id = a.Id(+) And c.项目id = d.Id And Nvl(c.是否删除, 0) = 0 And
                     Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(a.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (号源id_In Is Null Or c.Id = 号源id_In)
                    --
                     And Exists (Select 1
                      From 临床出诊安排 M, 临床出诊表 N
                      Where m.出诊id = n.Id And m.号源id = c.Id And Nvl(n.排班方式, 0) = 0 And n.发布时间 Is Not Null And
                            m.审核时间 Is Not Null And d_当前日期 <= m.终止时间)) Loop
  
    --检查当前日期所在的安排的收费项目是否为号源中的收费项目，如果不是，则更新号源中的收费项目
    Begin
      Select 项目id
      Into n_项目id
      From (Select a.项目id
             From 临床出诊安排 A, 临床出诊表 B
             Where a.出诊id = b.Id And a.号源id = c_号源.Id And a.审核时间 Is Not Null And d_当前日期 Between a.开始时间 And a.终止时间 And
                   Nvl(b.排班方式, 0) = 0 And b.发布时间 Is Not Null
             Order By a.登记时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        n_项目id := Null;
    End;
    If Nvl(n_项目id, 0) <> 0 Then
      If Nvl(c_号源.项目id, 0) <> n_项目id Then
        Update 临床出诊号源 Set 项目id = n_项目id Where ID = c_号源.Id;
        Commit;
      End If;
    End If;
  
    --第二层循环，出诊日期
    --从头一天开始生成，避免如全日(8:00-7:59)在0:00-7:59没有出诊记录
    For c_日期 In (Select m.日期,
                        Decode(To_Char(m.日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                                '周六', Null) As 星期
                 From (Select Trunc(d_当前日期) + 天数 As 日期
                        From (Select Level - 1 As 天数 From Dual Connect By Level <= c_号源.预约天数 + 1)
                        Where 号源id_In Is Not Null
                        Union All
                        Select Trunc(d_当前日期 - 1) + 天数 As 日期
                        From (Select Level - 1 As 天数 From Dual Connect By Level <= c_号源.预约天数 + 2)
                        Where 号源id_In Is Null) M) Loop
    
      l_固定时段 := t_Strlist();
      --检查当日是否在月/周出诊表中,若在，则不生成出诊记录
      Select Count(1)
      Into n_Count
      From 临床出诊安排 A, 临床出诊表 B
      Where a.出诊id = b.Id And a.号源id = c_号源.Id And c_日期.日期 Between Trunc(a.开始时间) And Trunc(a.终止时间) And
            Nvl(b.排班方式, 0) In (1, 2) And Rownum < 2;
    
      --当前号源为按月/周排班，且当前日期之前已有按月/周排班的出诊记录就不再按固定安排生成出诊记录了
      If Nvl(n_Count, 0) = 0 And Nvl(c_号源.排班方式, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 临床出诊安排 A, 临床出诊表 B
        Where a.出诊id = b.Id And Nvl(b.排班方式, 0) In (1, 2) And a.号源id = c_号源.Id And a.开始时间 < c_日期.日期 And Rownum < 2;
      End If;
    
      If Nvl(n_Count, 0) = 0 Then
        If 号源id_In Is Null Then
          --出诊安排,取最后登记的一个
          Begin
            Select 安排id
            Into n_安排id
            From (Select a.Id As 安排id
                   From 临床出诊安排 A, 临床出诊表 B
                   Where a.号源id = c_号源.Id And a.出诊id = b.Id And Nvl(b.排班方式, 0) = 0 And b.发布时间 Is Not Null And
                         a.审核时间 Is Not Null And c_日期.日期 Between a.开始时间 And a.终止时间
                   Order By a.登记时间 Desc)
            Where Rownum < 2;
          Exception
            When Others Then
              n_安排id := 0;
          End;
        Else
          --如果指定了号源ID，肯定是发布后新增了临时安排重新生成出诊记录，最后登记的一个肯定是本次新增的，
          --只需要处理这个安排即可，不在这个安排有效时间范围内的就不处理
          Begin
            Select 安排id
            Into n_安排id
            From (Select a.Id As 安排id, a.开始时间, a.终止时间, Row_Number() Over(Order By a.登记时间 Desc) As 行号
                   From 临床出诊安排 A, 临床出诊表 B
                   Where a.号源id = c_号源.Id And a.出诊id = b.Id And Nvl(b.排班方式, 0) = 0 And b.发布时间 Is Not Null And
                         a.审核时间 Is Not Null And c_日期.日期 Between 开始时间 And 终止时间)
            Where 行号 = 1;
          Exception
            When Others Then
              n_安排id := 0;
          End;
        End If;
      
        If Nvl(n_安排id, 0) <> 0 Then
          If 号源id_In Is Null Then
            --确定当日是否有出诊记录
            Select Count(1)
            Into n_Count
            From 临床出诊记录 A
            Where a.号源id = c_号源.Id And a.出诊日期 = c_日期.日期 And Rownum < 2;
          
            --1.未指定号源ID，则是正常生成出诊记录，有出诊记录的日期将不再处理
            If Nvl(n_Count, 0) = 0 Then
              --1.1无出诊记录，正常生成
              n_是否出诊 := 1;
            Else
              --1.2有出诊记录，不再处理
              n_是否出诊 := 0;
            End If;
          Else
            --2.指定了号源ID，肯定是发布后新增了临时安排重新生成出诊记录
            n_是否出诊 := 1;
            --当日有出诊记录，需要做如下处理
            For c_记录 In (Select a.安排id, a.Id As 记录id, a.出诊日期, a.上班时段, a.是否分时段, a.是否序号控制
                         From 临床出诊记录 A
                         Where a.号源id = c_号源.Id And a.出诊日期 = c_日期.日期) Loop
            
              Select Count(1) Into n_Count From 病人挂号记录 Where 出诊记录id = c_记录.记录id;
              If Nvl(n_Count, 0) = 0 Then
                --2.2.1如果时段不存在预约挂号数据，则删除重新生成
                Zl_临床出诊上班时段_Delete(c_记录.安排id, To_Char(c_记录.出诊日期, 'yyyy-mm-dd'), 1, c_记录.上班时段);
              Else
                --2.2.2如果时段存在预约挂号数据，则只需调整出诊记录的安排ID即可
                Update 临床出诊记录 Set 安排id = n_安排id Where ID = c_记录.记录id;
                l_固定时段.Extend();
                l_固定时段(l_固定时段.Count) := c_记录.上班时段;
              End If;
            End Loop;
          End If;
        
          --检查这天是否出诊
          If n_是否出诊 = 1 Then
            Select Count(1) Into n_Count From 临床出诊限制 Where 安排id = n_安排id And 限制项目 = c_日期.星期;
            If Nvl(n_Count, 0) = 0 Then
              n_是否出诊 := 0;
            End If;
          End If;
        
          If Nvl(n_是否出诊, 0) = 0 Then
            --如果不存在临床出诊记录，则增加临床出诊记录(时间段为NULL 的空记录)
            Insert Into 临床出诊记录
              (ID, 安排id, 号源id, 出诊日期, 登记人, 登记时间)
              Select 临床出诊记录_Id.Nextval, n_安排id, a.Id As ID, c_日期.日期, v_操作员姓名, d_登记日期 As 登记时间
              From 临床出诊号源 A, 临床出诊安排 B
              Where a.Id = b.号源id And b.Id = n_安排id
                   --
                    And Not Exists (Select 1 From 临床出诊记录 Where 号源id = a.Id And 出诊日期 = c_日期.日期);
          Else
            For c_记录 In (With c_时间段 As
                            (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间
                            From (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间,
                                          Row_Number() Over(Partition By 时间段 Order By 时间段, 站点 Asc, 号类 Asc) As 组号
                                   From 时间段
                                   Where Nvl(站点, c_号源.站点) = c_号源.站点 And Nvl(号类, c_号源.号类) = c_号源.号类)
                            Where 组号 = 1)
                           Select n_安排id As 安排id, B1.号源id, c_日期.日期 As 出诊日期, m.上班时段, m.Id As 限制id,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.开始时间, 'hh24:mi:ss'),
                                           'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.终止时间, 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.终止时间 <= j.开始时间 Then
                                     1
                                    Else
                                     0
                                  End As 终止时间, Null As 停诊开始时间, Null As 停诊终止时间, Null As 停诊原因,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.缺省时间, j.开始时间), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.缺省时间 < j.开始时间 Then
                                     1
                                    Else
                                     0
                                  End As 缺省预约时间,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.提前时间, j.开始时间), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.开始时间 < j.提前时间 Then
                                     -1
                                    Else
                                     0
                                  End As 提前挂号时间, m.限号数, 0 As 已挂数, m.限约数, 0 As 已约数, 0 As 其中已接收, m.是否序号控制, m.是否分时段, m.预约控制,
                                  m.是否独占, B1.项目id, B1.医生id, B1.医生姓名, Null As 替诊医生id, Null As 替诊医生姓名, m.分诊方式, m.诊室id,
                                  0 As 是否锁定, 0 As 是否临时出诊, v_操作员姓名 As 操作员姓名, d_登记日期 As 登记时间, c_日期.星期 As 限制项目
                           From 临床出诊安排 B1, 临床出诊限制 M, c_时间段 J
                           Where B1.Id = n_安排id And B1.Id = m.安排id And m.限制项目 = c_日期.星期 And m.上班时段 = j.时间段 And
                                 To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.开始时间, 'hh24:mi:ss'),
                                         'yyyy-mm-dd hh24:mi:ss') >= B1.开始时间) Loop
              Begin
                Select 1 Into n_Count From Table(l_固定时段) Where Column_Value = c_记录.上班时段;
              Exception
                When Others Then
                  n_Count := 0;
              End;
            
              If Nvl(n_Count, 0) = 0 Then
                Select 临床出诊记录_Id.Nextval Into n_记录id From Dual;
                Insert Into 临床出诊记录
                  (ID, 安排id, 号源id, 出诊日期, 上班时段, 开始时间, 终止时间, 停诊开始时间, 停诊终止时间, 停诊原因, 缺省预约时间, 提前挂号时间, 限号数, 已挂数, 限约数, 已约数,
                   其中已接收, 是否序号控制, 是否分时段, 预约控制, 是否独占, 项目id, 科室id, 医生id, 医生姓名, 替诊医生id, 替诊医生姓名, 分诊方式, 诊室id, 是否锁定, 是否临时出诊,
                   登记人, 登记时间, 是否发布)
                Values
                  (n_记录id, c_记录.安排id, c_记录.号源id, c_记录.出诊日期, c_记录.上班时段, c_记录.开始时间, c_记录.终止时间, c_记录.停诊开始时间, c_记录.停诊终止时间,
                   c_记录.停诊原因, c_记录.缺省预约时间, c_记录.提前挂号时间, c_记录.限号数, c_记录.已挂数, c_记录.限约数, c_记录.已约数, c_记录.其中已接收, c_记录.是否序号控制,
                   c_记录.是否分时段, c_记录.预约控制, c_记录.是否独占, c_记录.项目id, c_号源.科室id, c_记录.医生id, c_记录.医生姓名, c_记录.替诊医生id,
                   c_记录.替诊医生姓名, c_记录.分诊方式, c_记录.诊室id, c_记录.是否锁定, c_记录.是否临时出诊, c_记录.操作员姓名, d_登记日期, 1);
              
                d_开始时间 := c_记录.开始时间;
                --插入临床出诊序号控制
                If Nvl(c_记录.是否分时段, 0) = 1 And Nvl(c_记录.是否序号控制, 0) = 1 Then
                  --分时段且启用序号控制，使用"预约顺序号"记录"是否预约"
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 预约顺序号)
                    Select n_记录id, 序号,
                           To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'),
                                    'yyyy-mm-dd hh24:mi:ss') + Case
                              When d_开始时间 > To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                               1
                              Else
                               0
                            End,
                           To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'),
                                    'yyyy-mm-dd hh24:mi:ss') + Case
                              When d_开始时间 >= To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                               1
                              Else
                               0
                            End, 限制数量, 是否预约, 是否预约
                    From 临床出诊时段
                    Where 限制id = c_记录.限制id;
                Else
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约)
                    Select n_记录id, 序号,
                           To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') + Case
                             When d_开始时间 > To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'),
                                                   'yyyy-mm-dd hh24:mi:ss') Then
                              1
                             Else
                              0
                           End,
                           To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') + Case
                             When d_开始时间 >= To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') Then
                              1
                             Else
                              0
                           End, 限制数量, 是否预约
                    From 临床出诊时段
                    Where 限制id = c_记录.限制id;
                End If;
              
                --插入合作单位挂号控制记录
                Insert Into 临床出诊挂号控制记录
                  (类型, 性质, 名称, 记录id, 序号, 控制方式, 数量)
                  Select 类型, 性质, 名称, n_记录id, 序号, 控制方式, 数量
                  From 临床出诊挂号控制
                  Where 限制id = c_记录.限制id;
              
                --插入临床出诊诊室记录
                Insert Into 临床出诊诊室记录
                  (记录id, 诊室id)
                  Select n_记录id, 诊室id From 临床出诊诊室 Where 限制id = c_记录.限制id;
              End If;
            End Loop;
          
            --根据停诊安排和法定节假日调整出诊记录的出诊/预约情况
            Zl_Clinicvisitmodify(c_号源.Id, n_安排id, c_日期.日期, v_操作员姓名, d_登记日期);
          End If;
        End If;
      End If;
      --一天一提交
      Commit;
    End Loop;
  End Loop;
End Zl1_Auto_Buildingregisterplan;
/

--105791:冉俊明,2017-04-19,法定假日表字段名调整
Create Or Replace Procedure Zl_法定假日表_Modify
(
  操作类型_In     Number,
  年份_In         法定假日表.年份%Type,
  节日名称_In     法定假日表.节日名称%Type,
  开始日期_In     法定假日表.开始日期%Type,
  终止日期_In     法定假日表.终止日期%Type,
  备注_In         法定假日表.备注%Type,
  换休情况_In     Varchar2 := Null,
  允许预约日期_In 法定假日表.允许预约日期%Type,
  允许挂号日期_In 法定假日表.允许挂号日期%Type
) As
  --新增、修改法定节假日
  --      操作类型_In 0-新增，1-修改
  --      换休情况_In 格式：调休时间1~ 原上班时间1;调休时间2~ 原上班时间2;
  --      允许预约日期_In 允许预约的日期,格式：yyyy-mm-dd;yyyy-mm-dd;...
  --      允许挂号日期_In 允许挂号的日期,格式：yyyy-mm-dd;yyyy-mm-dd;...
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;

  v_换休情况 Varchar2(4000);
  v_当前项目 Varchar2(4000);
  d_开始日期 Date;
  d_终止日期 Date;
Begin
  If 操作类型_In = 0 Then
    --新增
    Begin
      Select 1
      Into n_Count
      From 法定假日表
      Where 性质 = 0 And 年份 = 年份_In And 节日名称 = 节日名称_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := 年份_In || '年已存在“' || 节日名称_In || '”！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From 临床出诊记录 A
      Where a.出诊日期 >= 开始日期_In And a.上班时段 Is Not Null And Nvl(a.是否发布, 0) = 1 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) <> 0 Then
      v_Err_Msg := '当前节假日开始时间之后已有有效的出诊安排，不能再新增！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From 临床出诊记录
      Where 出诊日期 Between 开始日期_In And 终止日期_In And Nvl(是否发布, 0) = 1 And (Nvl(已约数, 0) <> 0 Or Nvl(已挂数, 0) <> 0) And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '当前节假日的时间范围内已有预约挂号病人，不能设置！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From 法定假日表
      Where 性质 = 0 And 终止日期_In > 开始日期 And 开始日期_In < 终止日期 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '当前节假日的时间范围内已存在其它节假日！';
      Raise Err_Item;
    End If;
  
    Insert Into 法定假日表
      (年份, 节日名称, 性质, 开始日期, 终止日期, 备注, 允许预约日期, 允许挂号日期)
    Values
      (年份_In, 节日名称_In, 0, 开始日期_In, 终止日期_In, 备注_In, 允许预约日期_In, 允许挂号日期_In);
  
    If 换休情况_In Is Not Null Then
      v_换休情况 := 换休情况_In || ';';
    End If;
    While v_换休情况 Is Not Null Loop
      v_当前项目 := Substr(v_换休情况, 0, Instr(v_换休情况, ';') - 1);
      d_开始日期 := To_Date(Substr(v_当前项目, 0, Instr(v_当前项目, '~') - 1), 'yyyy-mm-dd');
      d_终止日期 := To_Date(Substr(v_当前项目, Instr(v_当前项目, '~') + 1), 'yyyy-mm-dd');
    
      Insert Into 法定假日表
        (年份, 节日名称, 性质, 开始日期, 终止日期, 备注)
      Values
        (年份_In, 节日名称_In, 1, d_开始日期, d_终止日期, Null);
    
      v_换休情况 := Substr(v_换休情况, Instr(v_换休情况, ';') + 1);
    End Loop;
  
  Elsif 操作类型_In = 1 Then
    --修改
    Begin
      Select 开始日期
      Into d_开始日期
      From 法定假日表
      Where 性质 = 0 And 年份 = 年份_In And 节日名称 = 节日名称_In And Rownum < 2;
    Exception
      When Others Then
        d_开始日期 := Null;
    End;
    If d_开始日期 Is Null Then
      v_Err_Msg := 年份_In || '年不存在“' || 节日名称_In || '”！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From 临床出诊记录 A
      Where a.出诊日期 >= d_开始日期 And a.上班时段 Is Not Null And Nvl(a.是否发布, 0) = 1 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) <> 0 Then
      v_Err_Msg := '当前节假日开始时间之后已有有效的出诊安排，不能修改！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From 临床出诊记录
      Where 出诊日期 Between 开始日期_In And 终止日期_In And Nvl(是否发布, 0) = 1 And (Nvl(已约数, 0) <> 0 Or Nvl(已挂数, 0) <> 0) And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '当前节假日的时间范围内已有预约挂号病人，不能修改！';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From 法定假日表
      Where 性质 = 0 And 终止日期_In > 开始日期 And 开始日期_In < 终止日期 And Not (年份 = 年份_In And 节日名称 = 节日名称_In) And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '当前节假日的时间范围内已存在其它节假日！';
      Raise Err_Item;
    End If;
  
    Update 法定假日表
    Set 开始日期 = 开始日期_In, 终止日期 = 终止日期_In, 备注 = 备注_In, 允许预约日期 = 允许预约日期_In, 允许挂号日期 = 允许挂号日期_In
    Where 年份 = 年份_In And Nvl(性质, 0) = 0 And 节日名称 = 节日名称_In;
  
    --先删除换休数据
    Delete From 法定假日表 Where 年份 = 年份_In And Nvl(性质, 0) = 1 And 节日名称 = 节日名称_In;
    If 换休情况_In Is Not Null Then
      v_换休情况 := 换休情况_In || ';';
    End If;
    While v_换休情况 Is Not Null Loop
      v_当前项目 := Substr(v_换休情况, 0, Instr(v_换休情况, ';') - 1);
      d_开始日期 := To_Date(Substr(v_当前项目, 0, Instr(v_当前项目, '~') - 1), 'yyyy-mm-dd');
      d_终止日期 := To_Date(Substr(v_当前项目, Instr(v_当前项目, '~') + 1), 'yyyy-mm-dd');
    
      Insert Into 法定假日表
        (年份, 节日名称, 性质, 开始日期, 终止日期, 备注)
      Values
        (年份_In, 节日名称_In, 1, d_开始日期, d_终止日期, Null);
    
      v_换休情况 := Substr(v_换休情况, Instr(v_换休情况, ';') + 1);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_法定假日表_Modify;
/

--107559:冉俊明,2017-04-18,增加终止停诊安排功能
--105791:冉俊明,2017-04-19,法定假日表字段名调整
Create Or Replace Procedure Zl_Clinicvisitmodify
(
  号源id_In     In 临床出诊记录.号源id%Type,
  安排id_In     In 临床出诊记录.号源id%Type,
  出诊日期_In   In 临床出诊记录.出诊日期%Type,
  登记人_In     In 临床出诊记录.登记人%Type,
  登记时间_In   In 临床出诊记录.登记时间%Type,
  是否已换休_In In Number := 0
) As
  --功能：根据停诊安排和法定节假日调整出诊记录的出诊/预约情况
  --入参：
  --     是否已换休_In 主要用于换休后进行停诊处理
  --说明：
  --     临床出诊号源.假日控制状态：0-不上班;1-上班且开放预约;2-上班但不开放预约;3-受节假日设置控制
  --     1-停诊，在停诊安排时间范围内
  --     2-停诊，在法定节假日内
  --       2.1临床出诊号源.假日控制状态=0
  --       2.2临床出诊号源.假日控制状态=3，且设置了允许预约/允许挂号，但该上班时段不在设置的允许预约/允许挂号的时间范围内
  --     3-禁止预约，在法定节假日内，
  --       3.1临床出诊号源.假日控制状态=2
  --       3.3临床出诊号源.假日控制状态=3，设置了允许预约/允许挂号，且该上班时段在设置的允许挂号的时间范围内，但不在设置的允许预约时间范围内
  --     else-正常出诊

  n_假日控制状态 临床出诊号源.假日控制状态%Type;
  n_是否假日换休 临床出诊号源.是否假日换休%Type;

  d_原上班日期 临床出诊记录.出诊日期%Type;
  d_调休日期   临床出诊记录.出诊日期%Type;

  d_停诊开始时间 临床出诊记录.停诊开始时间%Type;
  d_停诊终止时间 临床出诊记录.停诊终止时间%Type;
  v_停诊原因     临床出诊记录.停诊原因%Type;

  d_假日开始日期 法定假日表.开始日期%Type;
  d_假日终止日期 法定假日表.终止日期%Type;
  v_允许预约     法定假日表.允许预约日期%Type;
  v_允许挂号     法定假日表.允许挂号日期%Type;

  d_停止预约开始时间 临床出诊记录.停诊开始时间%Type;
  d_停止预约终止时间 临床出诊记录.停诊终止时间%Type;

  n_Count    Number(2);
  n_允许预约 Number(2);
  n_允许挂号 Number(2);

  Procedure Stopbespeak
  (
    记录id_In   In 临床出诊记录.Id%Type,
    开始时间_In In 临床出诊记录.开始时间%Type,
    终止时间_In In 临床出诊记录.终止时间%Type
  ) As
    --功能：禁止预约
    --说明：
    --      分时段且序号控制的，修改"临床出诊序号控制.是否预约"等于1的为0；取消发布时根据"预约顺序号"恢复
    --      分时段且不序号控制的，修改"临床出诊序号控制.是否预约"为0；取消发布时根据恢复为1
    --      不分时段的，提供公共函数在挂号预约时检查预约时间是否在不允许预约的时间范围内
  Begin
    Update 临床出诊序号控制 Set 是否预约 = 0 Where 记录id = 记录id_In And 开始时间 Between 开始时间_In And 终止时间_In;
  End Stopbespeak;

  Procedure Stopvisit
  (
    记录id_In       In 临床出诊记录.Id%Type,
    停诊开始时间_In In 临床出诊记录.停诊开始时间%Type,
    停诊终止时间_In In 临床出诊记录.停诊终止时间%Type,
    停诊原因_In     In 临床出诊记录.停诊原因%Type
  ) As
    --功能：停诊
    --说明：
    --     同一条出诊记录可以存在多条停诊记录，临床出诊记录的停诊开始时间为多条停诊记录的最小开始时间，停诊终止时间为多条停诊记录的最大终止时间

  
    d_停诊开始时间 临床出诊记录.停诊开始时间%Type;
    d_停诊终止时间 临床出诊记录.停诊终止时间%Type;
    v_停诊原因     临床出诊记录.停诊原因%Type;
  Begin
    If 停诊开始时间_In >= 停诊终止时间_In Then
      Return;
    End If;
  
    --产生停诊记录
    Insert Into 临床出诊停诊记录
      (ID, 记录id, 开始时间, 终止时间, 停诊原因, 申请人, 申请时间, 审批人, 审批时间, 登记人)
      Select 临床出诊停诊记录_Id.Nextval, ID, 停诊开始时间_In, 停诊终止时间_In, 停诊原因_In, Nvl(医生姓名, 登记人_In), 登记时间_In, 登记人_In, 登记时间_In, 登记人_In
      From 临床出诊记录
      Where ID = 记录id_In;
  
    Begin
      Select Min(a.开始时间), Max(a.终止时间), Max(a.停诊原因)
      Into d_停诊开始时间, d_停诊终止时间, v_停诊原因
      From 临床出诊停诊记录 A
      Where a.记录id = 记录id_In And a.取消时间 Is Null;
    Exception
      When Others Then
        d_停诊开始时间 := Null;
        d_停诊终止时间 := Null;
        v_停诊原因     := Null;
    End;
  
    Update 临床出诊记录
    Set 停诊开始时间 = d_停诊开始时间, 停诊终止时间 = d_停诊终止时间, 停诊原因 = v_停诊原因
    Where ID = 记录id_In;
  
    --调整"临床出诊序号控制.是否停诊"为1
    Update 临床出诊序号控制 A
    Set 是否停诊 = 1
    Where 记录id = 记录id_In And 开始时间 Between 停诊开始时间_In And 停诊终止时间_In And Exists
     (Select 1 From 临床出诊记录 Where ID = a.记录id And Nvl(是否序号控制, 0) = 1 And Nvl(是否分时段, 0) = 1);
  End Stopvisit;

  Procedure Changedaysoff
  (
    号源id_In     In 临床出诊记录.号源id%Type,
    安排id_In     In 临床出诊记录.安排id%Type,
    出诊日期_In   In 临床出诊记录.出诊日期%Type,
    原上班日期_In In 临床出诊记录.出诊日期%Type,
    调休日期_In   In 临床出诊记录.出诊日期%Type
  ) As
    --功能：换休处理
    n_安排id 临床出诊记录.安排id%Type;
    l_记录id t_Numlist := t_Numlist();
    n_Count  Number(2);
  Begin
    --1.前面的安排换到今日
    If 原上班日期_In Is Not Null Then
      --1.1.前面的日期没有安排则不处理
      Select Count(1)
      Into n_Count
      From 临床出诊记录
      Where 号源id = 号源id_In And 出诊日期 = 原上班日期_In And Rownum < 2;
    
      If Nvl(n_Count, 0) > 0 Then
        --[1]删除今日现有的安排
        Select ID Bulk Collect Into l_记录id From 临床出诊记录 Where 号源id = 号源id_In And 出诊日期 = 出诊日期_In;
        Zl_临床出诊记录_Batchdelete(l_记录id);
      
        --[2]复制安排
        For c_换休记录 In (Select ID, 是否发布 From 临床出诊记录 Where 号源id = 号源id_In And 出诊日期 = 原上班日期_In) Loop
          Zl_临床出诊记录_Copy(c_换休记录.Id, 安排id_In, 出诊日期_In, 登记人_In, 登记时间_In, c_换休记录.是否发布);
        End Loop;
      
        --[3]重新对今日进行停诊安排和法定节假日调整
        For c_记录 In (Select ID From 临床出诊记录 Where 号源id = 号源id_In And 出诊日期 = 出诊日期_In) Loop
          Zl_Clinicvisitmodify(号源id_In, 安排id_In, 出诊日期_In, 登记人_In, 登记时间_In, 1);
        End Loop;
      End If;
    End If;
  
    --2.今日的安排换到前面
    If 调休日期_In Is Not Null Then
      --2.1.今日没有安排则不处理
      Select Count(1)
      Into n_Count
      From 临床出诊记录
      Where 号源id = 号源id_In And 出诊日期 = 出诊日期_In And Rownum < 2;
    
      If Nvl(n_Count, 0) > 0 Then
        --2.2.前面那一天的安排已存在预约挂号记录则不替换(有漏洞)
        Select Count(1)
        Into n_Count
        From 临床出诊记录 A, 病人挂号记录 B
        Where a.Id = b.出诊记录id And a.号源id = 号源id_In And a.出诊日期 = 调休日期_In And Rownum < 2;
      
        If Nvl(n_Count, 0) = 0 Then
          --[1]记录前面那一天的原安排ID,没有就不处理
          Begin
            Select ID
            Into n_安排id
            From (Select Rownum As Rn, ID
                   From 临床出诊安排
                   Where 号源id = 号源id_In And 调休日期_In Between 开始时间 And 终止时间 And 审核时间 Is Not Null
                   Order By 登记时间 Desc)
            Where Rn < 2;
          Exception
            When Others Then
              n_安排id := 0;
          End;
        
          If Nvl(n_安排id, 0) <> 0 Then
            --[2]删除前面那一天现有的安排
            Select ID Bulk Collect Into l_记录id From 临床出诊记录 Where 号源id = 号源id_In And 出诊日期 = 调休日期_In;
            Zl_临床出诊记录_Batchdelete(l_记录id);
          
            --[3]复制安排
            For c_换休记录 In (Select ID From 临床出诊记录 Where 号源id = 号源id_In And 出诊日期 = 出诊日期_In) Loop
              --肯定是发布了的
              Zl_临床出诊记录_Copy(c_换休记录.Id, n_安排id, 调休日期_In, 登记人_In, 登记时间_In, 1);
            
            End Loop;
          
            --[4]重新对前面那一天进行停诊安排和法定节假日调整
            For c_记录 In (Select ID From 临床出诊记录 Where 号源id = 号源id_In And 出诊日期 = 调休日期_In) Loop
              Zl_Clinicvisitmodify(号源id_In, 安排id_In, 调休日期_In, 登记人_In, 登记时间_In, 1);
            End Loop;
          End If;
        End If;
      End If;
    End If;
  End Changedaysoff;
Begin
  Begin
    Select Nvl(b.假日控制状态, 0), Nvl(b.是否假日换休, 0)
    Into n_假日控制状态, n_是否假日换休
    From 临床出诊号源 B
    Where b.Id = 号源id_In;
  Exception
    When Others Then
      --没有找到号源，直接退出
      Return;
  End;

  --================================================================================
  --【1】假日换休处理
  --说明：只能用后面的日期向前面检查，因为后面的日期可能还没有制定安排
  --================================================================================
  If Nvl(是否已换休_In, 0) = 0 Then
    --确定法定节假日是否需要换休
    If Nvl(n_是否假日换休, 0) = 1 Then
      --1.前面的安排换到今日
      Begin
        --开始日期：原本休息日(即调休日) ， 终止日期：原本上班日(即被调休日)
        Select a.终止日期
        Into d_原上班日期
        From 法定假日表 A
        Where a.性质 = 1 And 出诊日期_In = a.开始日期 And a.终止日期 < 出诊日期_In And Rownum < 2;
      Exception
        When Others Then
          d_原上班日期 := Null;
      End;
    
      --2.今日的安排换到前面
      Begin
        --开始日期：原本休息日(即调休日) ， 终止日期：原本上班日(即被调休日)
        Select a.开始日期
        Into d_调休日期
        From 法定假日表 A
        Where a.性质 = 1 And 出诊日期_In = a.终止日期 And a.开始日期 < 出诊日期_In And Rownum < 2;
      Exception
        When Others Then
          d_调休日期 := Null;
      End;
    
      Changedaysoff(号源id_In, 安排id_In, 出诊日期_In, d_原上班日期, d_调休日期);
    End If;
  End If;

  For c_记录 In (Select ID, 出诊日期, 开始时间, 终止时间
               From 临床出诊记录
               Where 号源id = 号源id_In And 出诊日期 = 出诊日期_In And 上班时段 Is Not Null) Loop
    --================================================================================
    --【2】停诊安排停诊处理
    --================================================================================
    For c_停诊 In (Select a.开始时间, Nvl(a.失效时间, a.终止时间) As 终止时间, a.停诊原因
                 From 临床出诊停诊记录 A, 临床出诊记录 B
                 Where a.申请人 = b.医生姓名 And a.记录id Is Null And a.审批时间 Is Not Null And b.医生id Is Not Null And
                       b.Id = c_记录.Id And c_记录.开始时间 < Nvl(a.失效时间, a.终止时间) And c_记录.终止时间 > a.开始时间
                 Order By a.审批时间) Loop
    
      d_停诊开始时间 := c_停诊.开始时间;
      d_停诊终止时间 := c_停诊.终止时间;
      If d_停诊开始时间 < c_记录.开始时间 Then
        d_停诊开始时间 := c_记录.开始时间;
      End If;
      If d_停诊终止时间 > c_记录.终止时间 Then
        d_停诊终止时间 := c_记录.终止时间;
      End If;
      Stopvisit(c_记录.Id, d_停诊开始时间, d_停诊终止时间, c_停诊.停诊原因);
    End Loop;
  
    --================================================================================
    --【3】法定节假日停诊及禁止预约处理
    --================================================================================
    --1.查找含有上班时段时间的节假日，以第一个为准（开始时间升序排序），一般也只有一个
    Begin
      Select 开始日期, 终止日期, 节日名称, 允许预约日期, 允许挂号日期
      Into d_假日开始日期, d_假日终止日期, v_停诊原因, v_允许预约, v_允许挂号
      From (Select a.开始日期, a.终止日期, a.节日名称, a.允许预约日期, a.允许挂号日期
             From 法定假日表 A
             Where a.性质 = 0 And c_记录.开始时间 < a.终止日期 And c_记录.终止时间 > a.开始日期
             Order By a.开始日期)
      Where Rownum < 2;
    Exception
      When Others Then
        d_假日开始日期 := Null;
        d_假日终止日期 := Null;
        v_停诊原因     := Null;
        v_允许预约     := Null;
        v_允许挂号     := Null;
    End;
  
    If v_停诊原因 Is Not Null Then
      --假日控制状态:0-不上班;1-上班且开放预约;2-上班但不开放预约;3-受节假日设置控制
      If Nvl(n_假日控制状态, 0) = 0 Then
        --不上班，停诊
        d_停诊开始时间 := d_假日开始日期;
        d_停诊终止时间 := d_假日终止日期 + 1 - 1 / 24 / 60 / 60;
        If d_停诊开始时间 < c_记录.开始时间 Then
          d_停诊开始时间 := c_记录.开始时间;
        End If;
        If d_停诊终止时间 > c_记录.终止时间 Then
          d_停诊终止时间 := c_记录.终止时间;
        End If;
        Stopvisit(c_记录.Id, d_停诊开始时间, d_停诊终止时间, v_停诊原因);
      Elsif Nvl(n_假日控制状态, 0) = 2 Then
        --允许挂号，但禁止预约
        d_停止预约开始时间 := d_假日开始日期;
        d_停止预约终止时间 := d_假日终止日期 + 1 - 1 / 24 / 60 / 60;
        If d_停止预约开始时间 < c_记录.开始时间 Then
          d_停止预约开始时间 := c_记录.开始时间;
        End If;
        If d_停止预约终止时间 > c_记录.终止时间 Then
          d_停止预约终止时间 := c_记录.终止时间;
        End If;
        Stopbespeak(c_记录.Id, d_停止预约开始时间, d_停止预约终止时间);
      Elsif Nvl(n_假日控制状态, 0) = 3 Then
        --没有"允许挂号"的就一定没有"允许预约"的
        If v_允许挂号 Is Not Null Then
          --2.检查是否有包含上班时段时间的"允许挂号"
          --因为上班时段最多24小时，所以查出的结果最多两天，且这两天一定是连续的
          n_允许挂号 := 0;
          For c_允许挂号 In (With 临时表 As
                            (Select To_Date(Column_Value, 'yyyy-mm-dd') As 开始时间,
                                   To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 As 终止时间
                            From Table(f_Str2list(v_允许挂号, ';'))
                            Where c_记录.开始时间 < To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And
                                  c_记录.终止时间 > To_Date(Column_Value, 'yyyy-mm-dd')
                            Order By To_Date(Column_Value, 'yyyy-mm-dd'))
                           Select a.开始时间, Nvl(b.终止时间, a.终止时间) As 终止时间
                           From 临时表 A, 临时表 B
                           Where a.终止时间 = b.开始时间(+) - 1 / 24 / 60 / 60 And Rownum < 2) Loop
          
            n_允许挂号 := 1;
            n_允许预约 := 0;
            --3.检查是否有包含上班时段时间的"允许预约"
            For c_允许预约 In (With 临时表 As
                              (Select To_Date(Column_Value, 'yyyy-mm-dd') As 开始时间,
                                     To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 As 终止时间
                              From Table(f_Str2list(v_允许预约, ';'))
                              Where c_记录.开始时间 < To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And
                                    c_记录.终止时间 > To_Date(Column_Value, 'yyyy-mm-dd')
                              Order By To_Date(Column_Value, 'yyyy-mm-dd'))
                             Select a.开始时间, Nvl(b.终止时间, a.终止时间) As 终止时间
                             From 临时表 A, 临时表 B
                             Where a.终止时间 = b.开始时间(+) - 1 / 24 / 60 / 60 And Rownum < 2) Loop
            
              n_允许预约 := 1;
              --在"允许挂号","允许预约"时间范围内的不需要处理
            
              --检查前后是否需要禁止预约
              If c_记录.开始时间 < c_允许预约.开始时间 And c_允许挂号.开始时间 < c_允许预约.开始时间 Then
                If c_记录.开始时间 < c_允许挂号.开始时间 Then
                  d_停止预约开始时间 := c_允许挂号.开始时间;
                Else
                  d_停止预约开始时间 := c_记录.开始时间;
                End If;
                d_停止预约终止时间 := c_允许预约.开始时间;
                Stopbespeak(c_记录.Id, d_停止预约开始时间, d_停止预约终止时间);
              End If;
            
              If c_记录.终止时间 > c_允许预约.终止时间 And c_允许挂号.终止时间 > c_允许预约.终止时间 Then
                d_停止预约开始时间 := c_允许预约.终止时间;
                If c_记录.终止时间 > c_允许挂号.终止时间 Then
                  d_停止预约开始时间 := c_允许挂号.终止时间;
                Else
                  d_停止预约开始时间 := c_记录.终止时间;
                End If;
                Stopbespeak(c_记录.Id, d_停止预约开始时间, d_停止预约终止时间);
              End If;
            End Loop;
          
            --允许挂号，但禁止预约
            If Nvl(n_允许预约, 0) = 0 Then
              d_停止预约开始时间 := c_允许挂号.开始时间;
              d_停止预约终止时间 := c_允许挂号.终止时间;
              If d_停止预约开始时间 < c_记录.开始时间 Then
                d_停止预约开始时间 := c_记录.开始时间;
              End If;
              If d_停止预约终止时间 > c_记录.终止时间 Then
                d_停止预约终止时间 := c_记录.终止时间;
              End If;
              Stopbespeak(c_记录.Id, d_停止预约开始时间, d_停止预约终止时间);
            End If;
          
            --检查前后是否需要停诊
            If c_记录.开始时间 < c_允许挂号.开始时间 And d_假日开始日期 < c_允许挂号.开始时间 Then
              If c_记录.开始时间 < d_假日开始日期 Then
                d_停诊开始时间 := d_假日开始日期;
              Else
                d_停诊开始时间 := c_记录.开始时间;
              End If;
              d_停诊终止时间 := c_允许挂号.开始时间;
              Stopvisit(c_记录.Id, d_停诊开始时间, d_停诊终止时间, v_停诊原因);
            End If;
          
            If c_记录.终止时间 > c_允许挂号.终止时间 And d_假日终止日期 > c_允许挂号.终止时间 Then
              d_停诊开始时间 := c_允许挂号.终止时间;
              If c_记录.终止时间 > d_假日终止日期 Then
                d_停诊终止时间 := d_停诊终止时间;
              Else
                d_停诊终止时间 := c_记录.终止时间;
              End If;
              Stopvisit(c_记录.Id, d_停诊开始时间, d_停诊终止时间, v_停诊原因);
            End If;
          End Loop;
        
          --不在设置的允许挂号时间范围内，停诊
          If Nvl(n_允许挂号, 0) = 0 Then
            d_停诊开始时间 := d_假日开始日期;
            d_停诊终止时间 := d_假日终止日期 + 1 - 1 / 24 / 60 / 60;
            If d_停诊开始时间 < c_记录.开始时间 Then
              d_停诊开始时间 := c_记录.开始时间;
            End If;
            If d_停诊终止时间 > c_记录.终止时间 Then
              d_停诊终止时间 := c_记录.终止时间;
            End If;
            Stopvisit(c_记录.Id, d_停诊开始时间, d_停诊终止时间, v_停诊原因);
          End If;
        Else
          --未设置允许挂号/允许预约，则停诊
          d_停诊开始时间 := d_假日开始日期;
          d_停诊终止时间 := d_假日终止日期 + 1 - 1 / 24 / 60 / 60;
          If d_停诊开始时间 < c_记录.开始时间 Then
            d_停诊开始时间 := c_记录.开始时间;
          End If;
          If d_停诊终止时间 > c_记录.终止时间 Then
            d_停诊终止时间 := c_记录.终止时间;
          End If;
          Stopvisit(c_记录.Id, d_停诊开始时间, d_停诊终止时间, v_停诊原因);
        End If;
      End If;
    End If;
  End Loop;
End Zl_Clinicvisitmodify;
/

--105791:冉俊明,2017-04-19,法定假日表字段名调整
Create Or Replace Function Zl_Fun_Get临床出诊预约状态
(
  记录id_In   In 临床出诊记录.Id%Type,
  预约时间_In In 病人挂号记录.预约时间%Type,
  序号_In     临床出诊序号控制.序号%Type := Null,
  预约方式_In 预约方式.名称%Type := Null,
  合作单位_In 挂号合作单位.名称%Type := Null,
  收费预约_In Number := 0
) Return Varchar2 As
  --功能：判断出诊记录在预约时间是否可预约
  --入参：
  --返回：
  --     格式：预约状态|提示信息，如："1|预约时间不在当前上班时段时间范围内。"
  --     预约状态：
  --         0-可预约
  --         ======================================================
  --         1-不可预约，预约时间不在当前上班时段时间范围内
  --         2-不可预约，当前上班时段禁止预约
  --         3-不可预约，当前上班时段在预约时间时已停诊
  --         4-不可预约，当前上班时段剩余可预约数为零
  --         ======================================================
  --         5-不可预约，当前预约时间在法定节假日时间范围内，不上班
  --         6-不可预约，当前预约时间在法定节假日时间范围内，禁止预约
  --         7-不可预约，当前预约时间在法定节假日不允许预约的时间范围内
  --         8-不可预约，当前预约时间在法定节假日不允许挂号的时间范围内
  --         9-不可预约，当前预约时间在法定节假日时间范围内，已停诊
  --         ======================================================
  --         10-不可预约，当前预约方式禁止预约
  --         11-不可预约，当前预约方式可预约数不足
  --         ======================================================
  --         12-不可预约，当前合作单位禁止预约
  --         13-不可预约，当前合作单位可预约数不足
  --         ======================================================
  --         14-不可预约，当前序号禁止预约
  --         15-不可预约，当前序号已经被使用
  --         16-不可预约，当前序号不可用
  --
  n_号源id         临床出诊记录.号源id%Type;
  n_是否分时段     临床出诊记录.是否分时段%Type;
  n_预约控制       临床出诊记录.预约控制%Type;
  d_停诊开始时间   临床出诊记录.停诊开始时间%Type;
  d_停诊终止时间   临床出诊记录.停诊终止时间%Type;
  v_停诊原因       临床出诊记录.停诊原因%Type;
  n_限约数         临床出诊记录.限约数%Type;
  n_已约数         临床出诊记录.已约数%Type;
  n_独占           临床出诊记录.是否独占%Type;
  n_控制方式       临床出诊挂号控制记录.控制方式%Type;
  n_数量           临床出诊挂号控制记录.数量%Type;
  n_数量限制       临床出诊挂号控制记录.数量%Type;
  n_序号控制       临床出诊记录.是否序号控制%Type;
  v_预约方式       临床出诊挂号控制记录.名称%Type;
  n_类型           临床出诊挂号控制记录.类型%Type;
  n_预约方式限约数 临床出诊记录.限约数%Type;
  n_预约方式已约数 临床出诊记录.已约数%Type;
  n_挂号状态       临床出诊序号控制.挂号状态%Type;
  n_是否预约       临床出诊序号控制.是否预约%Type;

  n_假日控制状态 临床出诊号源.假日控制状态%Type;

  v_允许预约 法定假日表.允许预约日期%Type;
  v_允许挂号 法定假日表.允许挂号日期%Type;
  n_Count    Number(2);
  n_已使用   Number(5);
Begin
  Begin
    Select a.号源id, a.是否分时段, a.预约控制, a.停诊开始时间, a.停诊终止时间, a.停诊原因, Nvl(限约数, 限号数), 已约数, 是否独占, 是否序号控制
    Into n_号源id, n_是否分时段, n_预约控制, d_停诊开始时间, d_停诊终止时间, v_停诊原因, n_限约数, n_已约数, n_独占, n_序号控制
    From 临床出诊记录 A
    Where a.Id = 记录id_In And 预约时间_In Between 开始时间 And 终止时间;
  Exception
    When Others Then
      Return '1|预约时间不在当前上班时段时间范围内。';
  End;

  --预约方式检查
  If 预约方式_In Is Not Null Then
    Begin
      Select 控制方式
      Into n_控制方式
      From 临床出诊挂号控制记录
      Where 类型 = 2 And 性质 = 1 And 记录id = 记录id_In And 名称 = 预约方式_In And Rownum < 2;
    Exception
      When Others Then
        Begin
          Select 控制方式
          Into n_控制方式
          From 临床出诊挂号控制记录
          Where 类型 = 2 And 性质 = 1 And 记录id = 记录id_In And Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
    End;
    If n_控制方式 = 0 Then
      Return '10|当前预约方式禁止预约。';
    End If;
    If n_控制方式 = 1 Or n_控制方式 = 2 Then
      Select Nvl(限约数, 限号数) Into n_预约方式限约数 From 临床出诊记录 Where ID = 记录id_In;
      If n_独占 = 0 Then
        Begin
          Select 数量
          Into n_数量
          From 临床出诊挂号控制记录
          Where 类型 = 2 And 性质 = 1 And 名称 = 预约方式_In And 记录id = 记录id_In;
        Exception
          When Others Then
            n_数量 := Null;
        End;
        If n_数量 Is Not Null Then
          If n_控制方式 = 1 Then
            n_预约方式限约数 := Round(n_预约方式限约数 * n_数量 / 100);
          Else
            n_预约方式限约数 := n_数量;
          End If;
          Select Count(1)
          Into n_预约方式已约数
          From 病人挂号记录
          Where 出诊记录id = 记录id_In And 记录状态 = 1 And 预约方式 = 预约方式_In;
          If n_预约方式已约数 >= n_预约方式限约数 Then
            Return '11|当前预约方式可预约数不足。';
          End If;
        End If;
      Else
        --限数量独占
        Begin
          Select 数量
          Into n_数量
          From 临床出诊挂号控制记录
          Where 类型 = 2 And 性质 = 1 And 名称 = 预约方式_In And 记录id = 记录id_In;
        Exception
          When Others Then
            n_数量 := Null;
        End;
        If n_数量 Is Not Null Then
          If n_控制方式 = 1 Then
            n_预约方式限约数 := Round(n_预约方式限约数 * n_数量 / 100);
          Else
            n_预约方式限约数 := n_数量;
          End If;
          Select Count(1)
          Into n_预约方式已约数
          From 病人挂号记录
          Where 出诊记录id = 记录id_In And 记录状态 = 1 And 预约方式 = 预约方式_In;
          If n_预约方式已约数 >= n_预约方式限约数 Then
            Return '11|当前预约方式可预约数不足。';
          End If;
        Else
          If 收费预约_In = 0 Then
            For r_限制 In (Select 数量, 名称, 类型 From 临床出诊挂号控制记录 Where 性质 = 1 And 记录id = 记录id_In) Loop
              If r_限制.类型 = 1 Then
                Select Count(1)
                Into n_已使用
                From 病人挂号记录
                Where 出诊记录id = 记录id_In And 合作单位 = r_限制.名称 And 记录状态 = 1;
              Else
                Select Count(1)
                Into n_已使用
                From 病人挂号记录
                Where 出诊记录id = 记录id_In And 预约方式 = r_限制.名称 And 记录状态 = 1;
              End If;
              If n_控制方式 = 1 Then
                n_数量限制 := Nvl(n_数量限制, 0) + Round(r_限制.数量 * n_预约方式限约数 / 100) - Nvl(n_已使用, 0);
              Else
                n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
              End If;
            End Loop;
            Select Count(1) Into n_已使用 From 病人挂号记录 Where 出诊记录id = 记录id_In And 记录状态 = 1;
            If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
              Null;
            Else
              Return '11|当前预约方式可预约数不足。';
            End If;
          Else
            For r_限制 In (Select 数量, 名称, 类型
                         From 临床出诊挂号控制记录
                         Where 性质 = 1 And 类型 = 2 And 记录id = 记录id_In) Loop
              Select Count(1)
              Into n_已使用
              From 病人挂号记录
              Where 出诊记录id = 记录id_In And 预约方式 = r_限制.名称 And 记录状态 = 1;
              If n_控制方式 = 1 Then
                n_数量限制 := Nvl(n_数量限制, 0) + Round(r_限制.数量 * n_预约方式限约数 / 100) - Nvl(n_已使用, 0);
              Else
                n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
              End If;
            End Loop;
            Select Count(1) Into n_已使用 From 病人挂号记录 Where 出诊记录id = 记录id_In And 记录状态 = 1;
            If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
              Null;
            Else
              Return '11|当前预约方式可预约数不足。';
            End If;
          End If;
        End If;
      End If;
    End If;
    If n_控制方式 = 3 Then
      If n_序号控制 = 1 Then
        If 收费预约_In = 0 Then
          Begin
            Select 数量, 名称, 类型
            Into n_预约方式限约数, v_预约方式, n_类型
            From 临床出诊挂号控制记录
            Where 性质 = 1 And 记录id = 记录id_In And 序号 = 序号_In;
          Exception
            When Others Then
              n_预约方式限约数 := Null;
          End;
          If n_预约方式限约数 Is Not Null Then
            If v_预约方式 <> 预约方式_In Or n_类型 = 1 Then
              Return '11|当前预约方式可预约数不足。';
            End If;
            Select Nvl(Max(1), 0)
            Into n_预约方式已约数
            From 病人挂号记录
            Where 出诊记录id = 记录id_In And 号序 = 序号_In;
            If n_预约方式已约数 >= n_预约方式限约数 Then
              Return '11|当前预约方式可预约数不足。';
            End If;
          End If;
        Else
          Begin
            Select 数量, 名称, 类型
            Into n_预约方式限约数, v_预约方式, n_类型
            From 临床出诊挂号控制记录
            Where 性质 = 1 And 类型 = 2 And 记录id = 记录id_In And 序号 = 序号_In;
          Exception
            When Others Then
              n_预约方式限约数 := Null;
          End;
          If n_预约方式限约数 Is Not Null Then
            If v_预约方式 <> 预约方式_In Then
              Return '11|当前预约方式可预约数不足。';
            End If;
            Select Nvl(Max(1), 0)
            Into n_预约方式已约数
            From 病人挂号记录
            Where 出诊记录id = 记录id_In And 号序 = 序号_In;
            If n_预约方式已约数 >= n_预约方式限约数 Then
              Return '11|当前预约方式可预约数不足。';
            End If;
          End If;
        End If;
      Else
        If 收费预约_In = 0 Then
          For r_限制 In (Select 数量, 名称, 类型
                       From 临床出诊挂号控制记录
                       Where 性质 = 1 And 记录id = 记录id_In And 序号 = 序号_In) Loop
            If r_限制.名称 <> 预约方式_In Or r_限制.类型 = 1 Then
              If r_限制.类型 = 1 Then
                Select Count(1)
                Into n_已使用
                From 临床出诊序号控制 A, 病人挂号记录 B
                Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                      b.合作单位 = r_限制.名称 And b.记录状态 = 1;
              Else
                Select Count(1)
                Into n_已使用
                From 临床出诊序号控制 A, 病人挂号记录 B
                Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                      b.预约方式 = r_限制.名称 And b.记录状态 = 1;
              End If;
              n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
            Else
              Select Count(1)
              Into n_预约方式已约数
              From 临床出诊序号控制 A, 病人挂号记录 B
              Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                    b.预约方式 = 预约方式_In And b.记录状态 = 1;
              If n_预约方式已约数 >= n_预约方式限约数 Then
                Return '11|当前预约方式可预约数不足。';
              End If;
            End If;
          End Loop;
          Select Count(1)
          Into n_已使用
          From 临床出诊序号控制 A
          Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And 序号 = 序号_In;
          Select Nvl(限约数, 限号数) Into n_预约方式限约数 From 临床出诊记录 Where ID = 记录id_In;
          If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
            Null;
          Else
            Return '11|当前预约方式可预约数不足。';
          End If;
        Else
          For r_限制 In (Select 数量, 名称, 类型
                       From 临床出诊挂号控制记录
                       Where 性质 = 1 And 类型 = 2 And 记录id = 记录id_In And 序号 = 序号_In) Loop
            If r_限制.名称 <> 预约方式_In Then
              Select Count(1)
              Into n_已使用
              From 临床出诊序号控制 A, 病人挂号记录 B
              Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                    b.预约方式 = r_限制.名称 And b.记录状态 = 1;
              n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
            Else
              Select Count(1)
              Into n_预约方式已约数
              From 临床出诊序号控制 A, 病人挂号记录 B
              Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                    b.预约方式 = 预约方式_In And b.记录状态 = 1;
              If n_预约方式已约数 >= n_预约方式限约数 Then
                Return '11|当前预约方式可预约数不足。';
              End If;
            End If;
          End Loop;
          Select Count(1)
          Into n_已使用
          From 临床出诊序号控制 A
          Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And 序号 = 序号_In;
          Select Nvl(限约数, 限号数) Into n_预约方式限约数 From 临床出诊记录 Where ID = 记录id_In;
          If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
            Null;
          Else
            Return '11|当前预约方式可预约数不足。';
          End If;
        End If;
      End If;
    End If;
  End If;

  --合作单位检查
  If 合作单位_In Is Not Null Then
    Begin
      Select 控制方式
      Into n_控制方式
      From 临床出诊挂号控制记录
      Where 类型 = 1 And 性质 = 1 And 记录id = 记录id_In And 名称 = 合作单位_In And Rownum < 2;
    Exception
      When Others Then
        Begin
          Select 控制方式
          Into n_控制方式
          From 临床出诊挂号控制记录
          Where 类型 = 1 And 性质 = 1 And 记录id = 记录id_In And Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
    End;
    If n_控制方式 = 0 Then
      Return '12|当前合作单位禁止预约。';
    End If;
    If n_控制方式 = 1 Or n_控制方式 = 2 Then
      Select Nvl(限约数, 限号数) Into n_预约方式限约数 From 临床出诊记录 Where ID = 记录id_In;
      If n_独占 = 0 Then
        Begin
          Select 数量
          Into n_数量
          From 临床出诊挂号控制记录
          Where 类型 = 1 And 性质 = 1 And 名称 = 合作单位_In And 记录id = 记录id_In;
        Exception
          When Others Then
            n_数量 := Null;
        End;
        If n_数量 Is Not Null Then
          If n_控制方式 = 1 Then
            n_预约方式限约数 := Round(n_预约方式限约数 * n_数量 / 100);
          Else
            n_预约方式限约数 := n_数量;
          End If;
          Select Count(1)
          Into n_预约方式已约数
          From 病人挂号记录
          Where 出诊记录id = 记录id_In And 记录状态 = 1 And 合作单位 = 合作单位_In;
          If n_预约方式已约数 >= n_预约方式限约数 Then
            Return '13|当前合作单位可预约数不足。';
          End If;
        End If;
      Else
        --限数量独占
        Begin
          Select 数量
          Into n_数量
          From 临床出诊挂号控制记录
          Where 类型 = 1 And 性质 = 1 And 名称 = 合作单位_In And 记录id = 记录id_In;
        Exception
          When Others Then
            n_数量 := Null;
        End;
        If n_数量 Is Not Null Then
          If n_控制方式 = 1 Then
            n_预约方式限约数 := Round(n_预约方式限约数 * n_数量 / 100);
          Else
            n_预约方式限约数 := n_数量;
          End If;
          Select Count(1)
          Into n_预约方式已约数
          From 病人挂号记录
          Where 出诊记录id = 记录id_In And 记录状态 = 1 And 合作单位 = 合作单位_In;
          If n_预约方式已约数 >= n_预约方式限约数 Then
            Return '13|当前合作单位可预约数不足。';
          End If;
        Else
          For r_限制 In (Select 数量, 名称, 类型 From 临床出诊挂号控制记录 Where 性质 = 1 And 记录id = 记录id_In) Loop
            If r_限制.类型 = 1 Then
              Select Count(1)
              Into n_已使用
              From 病人挂号记录
              Where 出诊记录id = 记录id_In And 合作单位 = r_限制.名称 And 记录状态 = 1;
            Else
              Select Count(1)
              Into n_已使用
              From 病人挂号记录
              Where 出诊记录id = 记录id_In And 预约方式 = r_限制.名称 And 记录状态 = 1;
            End If;
            If n_控制方式 = 1 Then
              n_数量限制 := Nvl(n_数量限制, 0) + Round(r_限制.数量 * n_预约方式限约数 / 100) - Nvl(n_已使用, 0);
            Else
              n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
            End If;
          End Loop;
          Select Count(1) Into n_已使用 From 病人挂号记录 Where 出诊记录id = 记录id_In And 记录状态 = 1;
          If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
            Null;
          Else
            Return '13|当前合作单位可预约数不足。';
          End If;
        End If;
      End If;
    End If;
    If n_控制方式 = 3 Then
      If n_序号控制 = 1 Then
        Begin
          Select 数量, 名称, 类型
          Into n_预约方式限约数, v_预约方式, n_类型
          From 临床出诊挂号控制记录
          Where 性质 = 1 And 记录id = 记录id_In And 序号 = 序号_In;
        Exception
          When Others Then
            n_预约方式限约数 := Null;
        End;
        If n_预约方式限约数 Is Not Null Then
          If v_预约方式 <> 合作单位_In Or n_类型 = 1 Then
            Return '13|当前合作单位可预约数不足。';
          End If;
          Select Nvl(Max(1), 0)
          Into n_预约方式已约数
          From 病人挂号记录
          Where 出诊记录id = 记录id_In And 号序 = 序号_In;
          If n_预约方式已约数 >= n_预约方式限约数 Then
            Return '13|当前合作单位可预约数不足。';
          End If;
        End If;
      Else
        For r_限制 In (Select 数量, 名称, 类型
                     From 临床出诊挂号控制记录
                     Where 性质 = 1 And 记录id = 记录id_In And 序号 = 序号_In) Loop
          If r_限制.名称 <> 合作单位_In Or r_限制.类型 = 1 Then
            If r_限制.类型 = 1 Then
              Select Count(1)
              Into n_已使用
              From 临床出诊序号控制 A, 病人挂号记录 B
              Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                    b.合作单位 = r_限制.名称 And b.记录状态 = 1;
            Else
              Select Count(1)
              Into n_已使用
              From 临床出诊序号控制 A, 病人挂号记录 B
              Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And
                    b.预约方式 = r_限制.名称 And b.记录状态 = 1;
            End If;
            n_数量限制 := Nvl(n_数量限制, 0) + r_限制.数量 - Nvl(n_已使用, 0);
          Else
            Select Count(1)
            Into n_预约方式已约数
            From 临床出诊序号控制 A, 病人挂号记录 B
            Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And a.备注 = b.号序 And b.合作单位 = 合作单位_In And
                  b.记录状态 = 1;
            If n_预约方式已约数 >= n_预约方式限约数 Then
              Return '13|当前合作单位可预约数不足。';
            End If;
          End If;
        End Loop;
        Select Count(1)
        Into n_已使用
        From 临床出诊序号控制 A
        Where a.记录id = 记录id_In And a.预约顺序号 Is Not Null And Nvl(a.挂号状态, 0) <> 0 And 序号 = 序号_In;
        Select Nvl(限约数, 限号数) Into n_预约方式限约数 From 临床出诊记录 Where ID = 记录id_In;
        If n_预约方式限约数 - n_数量限制 - n_已使用 > 0 Then
          Null;
        Else
          Return '13|当前合作单位可预约数不足。';
        End If;
      End If;
    End If;
  End If;

  --0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
  If Nvl(n_预约控制, 0) = 1 Then
    Return '2|当前上班时段禁止预约。';
  End If;

  If d_停诊开始时间 Is Not Null And Not (Nvl(n_序号控制, 0) = 1 And Nvl(n_是否分时段, 0) = 1) Then
    If 预约时间_In >= d_停诊开始时间 And 预约时间_In <= d_停诊终止时间 Then
      Return '3|当前上班时段在预约时间时已停诊，不能预约！';
    End If;
  End If;

  If Nvl(n_限约数, 0) > 0 Then
    If Nvl(n_限约数, 0) - Nvl(n_已约数, 0) <= 0 Then
      Return '4|当前上班时段剩余可预约数为零，不能继续预约！';
    End If;
  End If;

  If Nvl(n_是否分时段, 0) = 0 Then
    --不分时段
    Begin
      Select Nvl(b.假日控制状态, 0) Into n_假日控制状态 From 临床出诊号源 B Where b.Id = n_号源id;
    Exception
      When Others Then
        n_假日控制状态 := 0;
    End;
  
    --1.查找包含预约时间的节假日
    Begin
      Select a.允许预约日期, a.允许挂号日期
      Into v_允许预约, v_允许挂号
      From 法定假日表 A
      Where a.性质 = 0 And 预约时间_In Between a.开始日期 And a.终止日期 + 1 - 1 / 24 / 60 / 60 And Rownum < 2;
    Exception
      When Others Then
        Return '0|正常预约。';
    End;
  
    --假日控制状态：0-不上班;1-上班且开放预约;2-上班但不开放预约;3-受节假日设置控制
    If Nvl(n_假日控制状态, 0) = 0 Then
      --不上班的肯定是不能预约的
      Return '5|当前预约时间在法定节假日时间范围内，不上班。';
    Elsif Nvl(n_假日控制状态, 0) = 1 Then
      Return '0|正常预约。';
    Elsif Nvl(n_假日控制状态, 0) = 2 Then
      --在节假日时间范围内，则不能预约
      Return '6|当前预约时间在法定节假日时间范围内，禁止预约。';
    Elsif Nvl(n_假日控制状态, 0) = 3 Then
      --没有"允许挂号"就一定没有"允许预约"
      If v_允许挂号 Is Not Null Then
        --2.检查是否有包含预约时间的"允许挂号"
        Select Max(1)
        Into n_Count
        From Table(f_Str2list(v_允许挂号, ';'))
        Where 预约时间_In Between To_Date(Column_Value, 'yyyy-mm-dd') And
              To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And Rownum < 2;
      
        If Nvl(n_Count, 0) <> 0 Then
          --3.检查是否有包含预约时间的"允许预约"
          Select Max(1)
          Into n_Count
          From Table(f_Str2list(v_允许预约, ';'))
          Where 预约时间_In Between To_Date(Column_Value, 'yyyy-mm-dd') And
                To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And Rownum < 2;
        
          If Nvl(n_Count, 0) = 0 Then
            --不在"允许预约"时间范围内，则不能预约
            Return '7|当前预约时间在法定节假日不允许预约的时间范围内，不能预约。';
          Else
            Return '0|正常预约。';
          End If;
        Else
          Return '8|当前预约时间在法定节假日不允许挂号的时间范围内，不能预约。';
        End If;
      Else
        --没有设置"允许挂号"/"允许预约"表示停诊，肯定不能预约
        Return '9|当前预约时间在法定节假日时间范围内，已停诊，不能预约。';
      End If;
    End If;
  Else
    --分时段
    If Nvl(序号_In, 0) <> 0 Then
      Begin
        Select Nvl(是否预约, 0), Nvl(挂号状态, 0)
        Into n_是否预约, n_挂号状态
        From 临床出诊序号控制
        Where 记录id = 记录id_In And 序号 = 序号_In;
      Exception
        When Others Then
          Return '16|当前选择的序号不可用。';
      End;
      If n_是否预约 = 0 Then
        Return '14|当前选择的序号禁止预约。';
      End If;
      If n_挂号状态 <> 0 Then
        Return '15|当前选择的序号已经被使用。';
      End If;
    End If;
    Return '0|正常预约。';
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Get临床出诊预约状态;
/

--101575:余伟节,2017-04-24,解决路径项目对应的继承医嘱数据超长问题

Create Or Replace Procedure Zl_临床路径版本_Copy
(
  源路径id_In     临床路径版本.路径id%Type,
  源版本号_In     临床路径版本.版本号%Type,
  目标路径id_In   临床路径版本.路径id%Type,
  目标版本号_In   临床路径版本.版本号%Type,
  源分支id_In     临床路径分支.Id%Type := Null,
  是否分支路径_In Number := Null,
  目标分支id_In   临床路径分支.Id%Type := Null
  --功能：复制产生新的临床路径版本
  --参数：
  --  源版本号_In：如果未指定(0或NULL)，则取最新有效的版本号
  --  目标本号_In：如果未指定(0或NULL)，则产生新的版本号
  --  是否分支路径_In：编辑分支路径时从其他分支或主路径复制路径结构,1-是，0否。
  --  目标分支ID_In:编辑分支路径时复制其他分支的结构，其他分支的ID。
) Is
  n_源版本号   临床路径版本.版本号%Type;
  n_目标版本号 临床路径版本.版本号%Type;

  n_Advice_New_Id    Number;
  n_Advice_Parent_Id Number;

  n_Step_New_Id    Number;
  n_Step_Parent_Id Number;

  n_Item_New_Id Number;

  n_Eval_New_Id Number;
  n_Eval_Old_Id Number;

  n_Mark_New_Id Number;

  n_Branch_New_Id Number;

  v_Error Varchar2(255);
  Err_Custom Exception;

  n_前一阶段序号 临床路径阶段.序号%Type;
  n_结束天数     临床路径阶段.结束天数%Type;
  v_标准住院日   临床路径分支.标准住院日%Type;
  t_Advice       t_Numlist2 := t_Numlist2(); --缓存继承的医嘱ID C1:=上一个版本的医嘱ID;C2:=新生成的医嘱ID
  --临床路径分支
  Procedure 临床路径分支_Insert
  (
    源id_In       Number,
    New_Id_In     Number,
    路径id_In     Number,
    版本号_In     Number,
    名称_In       临床路径分支.名称%Type := Null,
    说明_In       临床路径分支.说明%Type := Null,
    前一阶段id_In 临床路径分支.前一阶段id%Type := Null,
    标准住院日_In 临床路径分支.标准住院日%Type := Null,
    标准费用_In   临床路径分支.标准费用%Type := Null
  ) Is
  Begin
    If Nvl(源id_In, 0) <> 0 Then
      Insert Into 临床路径分支
        (ID, 路径id, 版本号, 名称, 说明, 前一阶段id, 标准住院日, 标准费用, 创建人, 创建时间)
        Select New_Id_In, 路径id_In, 版本号_In, Nvl(名称_In, 名称), 说明, 前一阶段id, 标准住院日, 标准费用, Zl_Username, Sysdate
        From 临床路径分支
        Where ID = 源id_In;
    Else
      --如果是复制主路径，则如果标准住院日超出了，自动修改。
      Insert Into 临床路径分支
        (ID, 路径id, 版本号, 名称, 说明, 前一阶段id, 标准住院日, 标准费用, 创建人, 创建时间)
        Select New_Id_In, 路径id_In, 版本号_In, 名称_In, 说明_In, 前一阶段id_In, 标准住院日_In, 标准费用_In, Zl_Username, Sysdate
        From Dual;
    End If;
  
  End;

  --临床路径阶段
  Procedure 临床路径阶段_Insert
  (
    源id_In       Number,
    New_Id_In     Number,
    路径id_In     Number,
    版本号_In     Number,
    New_父id_In   Number,
    分支id_Old_In Number := Null,
    分支id_New_In Number := Null
  ) Is
  Begin
    Insert Into 临床路径阶段
      (ID, 路径id, 版本号, 父id, 序号, 名称, 开始天数, 结束天数, 标志, 说明, 分支id)
      Select New_Id_In, 路径id_In, 版本号_In, New_父id_In, 序号, 名称, 开始天数, 结束天数, 标志, 说明, 分支id_New_In
      From 临床路径阶段
      Where ID = 源id_In And Nvl(分支id, 0) = Nvl(分支id_Old_In, 0);
  End;
  ---临床路径项目
  Procedure 临床路径项目_Insert
  (
    源id_In       Number,
    New_Id_In     Number,
    路径id_In     Number,
    版本号_In     Number,
    New_阶段id_In Number,
    分支id_Old_In Number := Null,
    分支id_New_In Number := Null
  ) Is
  Begin
    Insert Into 临床路径项目
      (ID, 路径id, 版本号, 阶段id, 分类, 项目序号, 项目内容, 内容要求, 执行方式, 执行者, 生成者, 项目结果, 图标id, 分支id)
      Select New_Id_In, 路径id_In, 版本号_In, New_阶段id_In, 分类, 项目序号, 项目内容, 内容要求, 执行方式, 执行者, 生成者, 项目结果, 图标id, 分支id_New_In
      From 临床路径项目
      Where ID = 源id_In And Nvl(分支id, 0) = Nvl(分支id_Old_In, 0);
  End;
  --路径医嘱内容
  Procedure 路径医嘱内容_Insert
  (
    源id_In       Number,
    New_Id_In     Number,
    New_相关id_In Number
  ) Is
  Begin
    Insert Into 路径医嘱内容
      (ID, 相关id, 序号, 期效, 诊疗项目id, 医嘱内容, 单次用量, 总给予量, 收费细目id, 标本部位, 检查方法, 执行频次, 频率次数, 频率间隔, 间隔单位, 医生嘱托, 执行性质, 执行科室id, 时间方案,
       是否缺省, 是否备选, 组合项目id)
      Select New_Id_In, New_相关id_In, 序号, 期效, 诊疗项目id, 医嘱内容, 单次用量, 总给予量, 收费细目id, 标本部位, 检查方法, 执行频次, 频率次数, 频率间隔, 间隔单位, 医生嘱托,
             执行性质, 执行科室id, 时间方案, 是否缺省, 是否备选, 组合项目id
      From 路径医嘱内容
      Where ID = 源id_In;
  End;
  --临床路径医嘱
  Procedure 临床路径医嘱_Inset
  (
    路径项目id_In Number,
    医嘱内容id_In Number
  ) Is
  Begin
    Insert Into 临床路径医嘱 (路径项目id, 医嘱内容id) Values (路径项目id_In, 医嘱内容id_In);
  End;
  --临床路径病历
  Procedure 临床路径病历_Inset
  (
    源项目id_In   Number,
    项目id_New_In Number
  ) Is
  Begin
    Insert Into 临床路径病历
      (项目id, 文件id, 原型id, 名称, 序号)
      Select 项目id_New_In, 文件id, 原型id, 名称, 序号 From 临床路径病历 Where 项目id = 源项目id_In;
  End;
  ---临床路径评估
  Procedure 临床路径评估_Insert
  (
    源id_In       Number,
    New_Id_In     Number,
    路径id_In     Number,
    版本号_In     Number,
    阶段id_In     Number,
    分支id_Old_In Number := Null,
    分支id_New_In Number := Null
  ) Is
  Begin
    Insert Into 临床路径评估
      (ID, 路径id, 版本号, 阶段id, 评估类型, 分支id)
      Select New_Id_In, 路径id_In, 版本号_In, 阶段id_In, 评估类型, 分支id_New_In
      From 临床路径评估
      Where ID = 源id_In And Nvl(分支id, 0) = Nvl(分支id_Old_In, 0);
  End;

  Procedure 路径评估指标_Insert
  (
    源id_In   Number,
    New_Id_In Number,
    评估id_In Number
  ) Is
  Begin
    Insert Into 路径评估指标
      (ID, 评估id, 序号, 评估指标, 指标类型, 指标结果)
      Select New_Id_In, 评估id_In, 序号, 评估指标, 指标类型, 指标结果 From 路径评估指标 Where ID = 源id_In;
  End;
  --路径评估条件
  Procedure 路径评估条件_Insert
  (
    源评估id_In   Number,
    源指标id_In   Number,
    源项目id_In   Number,
    New_评估id_In Number,
    New_指标id_In Number,
    New_项目id_In Number
  ) Is
  Begin
    If 源指标id_In Is Null Then
      Insert Into 路径评估条件
        (评估id, 指标id, 项目id, 关系式, 条件值, 条件组合)
        Select New_评估id_In, New_指标id_In, New_项目id_In, 关系式, 条件值, 条件组合
        From 路径评估条件
        Where 评估id = 源评估id_In And 指标id Is Null And 项目id = 源项目id_In;
    Elsif 源项目id_In Is Null Then
      Insert Into 路径评估条件
        (评估id, 指标id, 项目id, 关系式, 条件值, 条件组合)
        Select New_评估id_In, New_指标id_In, New_项目id_In, 关系式, 条件值, 条件组合
        From 路径评估条件
        Where 评估id = 源评估id_In And 指标id = 源指标id_In And 项目id Is Null;
    End If;
  End;
  --临床路径阶段
  Procedure 临床路径阶段cascade_Insert
  (
    源id_In       Number,
    New_Id_In     Number,
    Old路径id_In  Number,
    New路径id_In  Number,
    Old版本号_In  Number,
    New版本号_In  Number,
    分支id_Old_In Number := Null,
    分支id_New_In Number := Null
  ) Is
    n_After   Number(10);
    n_Count   Number(10);
    n_Inherit Number;
    v_Oldid   Varchar2(4000);
    n_Start   Number(10);
    Arr_Id    t_Numlist;
  
  Begin
    ---临床路径评估(阶段，指标类评估条件）
    Select Max(a.Id)
    Into n_Eval_Old_Id
    From 临床路径评估 A
    Where a.路径id = Old路径id_In And a.版本号 = Old版本号_In And a.阶段id = 源id_In And a.评估类型 = 2 And
          Nvl(a.分支id, 0) = Nvl(分支id_Old_In, 0);
  
    If Nvl(n_Eval_Old_Id, 0) <> 0 Then
      Select 临床路径评估_Id.Nextval Into n_Eval_New_Id From Dual;
      临床路径评估_Insert(n_Eval_Old_Id, n_Eval_New_Id, New路径id_In, New版本号_In, New_Id_In, 分支id_Old_In, 分支id_New_In);
      ---路径评估指标
      For r_路径评估指标 In (Select ID From 路径评估指标 Where 评估id = n_Eval_Old_Id) Loop
        Select 路径评估指标_Id.Nextval Into n_Mark_New_Id From Dual;
        路径评估指标_Insert(r_路径评估指标.Id, n_Mark_New_Id, n_Eval_New_Id);
        ---路径评估条件
        路径评估条件_Insert(n_Eval_Old_Id, r_路径评估指标.Id, Null, n_Eval_New_Id, n_Mark_New_Id, Null);
      End Loop;
    End If;
    --临床路径项目
    For r_临床路径项目 In (Select ID
                     From 临床路径项目
                     Where 阶段id = 源id_In And 路径id = Old路径id_In And 版本号 = Old版本号_In And
                           Nvl(分支id, 0) = Nvl(分支id_Old_In, 0)) Loop
    
      Select 临床路径项目_Id.Nextval Into n_Item_New_Id From Dual;
      临床路径项目_Insert(r_临床路径项目.Id, n_Item_New_Id, New路径id_In, New版本号_In, New_Id_In, 分支id_Old_In, 分支id_New_In);
      ---临床路径评估（阶段评估，项目类评估条件）
      If Nvl(n_Eval_Old_Id, 0) <> 0 Then
        ---路径评估条件
        路径评估条件_Insert(n_Eval_Old_Id, Null, r_临床路径项目.Id, n_Eval_New_Id, Null, n_Item_New_Id);
      End If;
      ---临床路径病历
      临床路径病历_Inset(r_临床路径项目.Id, n_Item_New_Id);
    
      --路径医嘱内容
      For r_临床路径医嘱 In (Select b.Id
                       From 临床路径医嘱 A, 路径医嘱内容 B
                       Where a.路径项目id = r_临床路径项目.Id And a.医嘱内容id = b.Id And b.相关id Is Null) Loop
        --继承的医嘱判断
        Select Count(1) Into n_Inherit From 临床路径医嘱 Where 医嘱内容id = r_临床路径医嘱.Id;
        v_Oldid := Null;
        If n_Inherit > 1 Then
          Begin
            Select a.C2 Into v_Oldid From Table(t_Advice) A Where a.C1 = r_临床路径医嘱.Id;
          Exception
            When No_Data_Found Then
              v_Oldid := Null;
          End;
        End If;
        If v_Oldid Is Null Then
          ---b.序号 > a.序号 and b.ID >a.ID --获取父医嘱ID大于子医嘱ID并且父医嘱序号大于子医嘱序号的记录数
          Select Count(1)
          Into n_After
          From 路径医嘱内容 A
          Where a.相关id = r_临床路径医嘱.Id And Exists
           (Select 1 From 路径医嘱内容 B Where b.Id = r_临床路径医嘱.Id And b.序号 > a.序号 And b.Id > a.Id);
        
          Select Count(1) + 1 Into n_Count From 路径医嘱内容 A Where a.相关id = r_临床路径医嘱.Id;
          Select 路径医嘱内容_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
          If n_After = 0 Then
            n_Advice_Parent_Id := Arr_Id(1);
            n_Start            := 2;
          Else
            n_Advice_Parent_Id := Arr_Id(n_Count);
            n_Start            := 1;
          End If;
        
          路径医嘱内容_Insert(r_临床路径医嘱.Id, n_Advice_Parent_Id, Null);
          If n_Inherit > 1 Then
            t_Advice.Extend;
            t_Advice(t_Advice.Count) := t_Numobj2(r_临床路径医嘱.Id, n_Advice_Parent_Id);
          End If;
        Else
          n_Advice_Parent_Id := To_Number(v_Oldid);
        End If;
        ---临床路径医嘱
        临床路径医嘱_Inset(n_Item_New_Id, n_Advice_Parent_Id);
        --路径医嘱内容相应子节点
        For r_路径医嘱内容 In (Select ID From 路径医嘱内容 Where 相关id = r_临床路径医嘱.Id) Loop
          If v_Oldid Is Null Then
            n_Advice_New_Id := Arr_Id(n_Start);
            n_Start         := n_Start + 1;
          
            路径医嘱内容_Insert(r_路径医嘱内容.Id, n_Advice_New_Id, n_Advice_Parent_Id);
            If n_Inherit > 1 Then
              t_Advice.Extend;
              t_Advice(t_Advice.Count) := t_Numobj2(r_路径医嘱内容.Id, n_Advice_New_Id);
            End If;
          Else
            --继承医嘱，未产生新的ID
            If n_Inherit > 1 Then
              Select a.C2 Into n_Advice_New_Id From Table(t_Advice) A Where a.C1 = r_路径医嘱内容.Id;
            End If;
          End If;
          ---临床路径医嘱
          临床路径医嘱_Inset(n_Item_New_Id, n_Advice_New_Id);
        End Loop;
      End Loop;
    End Loop;
  End;
Begin
  --确定源路径版本号
  n_源版本号 := Nvl(源版本号_In, 0);
  If n_源版本号 = 0 Then
    Select 最新版本 Into n_源版本号 From 临床路径目录 Where ID = 源路径id_In;
    If Nvl(n_源版本号, 0) = 0 Then
      v_Error := '要复制的来源临床路径中没有可用的有效版本。';
      Raise Err_Custom;
    End If;
  End If;

  --确定目标路径版本号
  n_目标版本号 := Nvl(目标版本号_In, 0);
  If n_目标版本号 = 0 Then
    Select Nvl(Max(版本号), 0) + 1 Into n_目标版本号 From 临床路径版本 Where 路径id = 目标路径id_In;
  Else
    If Nvl(是否分支路径_In, 0) = 1 Then
      --从其他分支或主路径复制时
      --记录下前一阶段序号
      Select Max(a.序号)
      Into n_前一阶段序号
      From 临床路径阶段 A, 临床路径分支 B
      Where a.Id = b.前一阶段id And b.Id = Nvl(目标分支id_In, 0);
    
      For r_目标分支 In (Select * From 临床路径分支 Where ID = Nvl(目标分支id_In, 0)) Loop
        Zl_临床路径分支_Delete(目标分支id_In);
        Select 临床路径分支_Id.Nextval Into n_Branch_New_Id From Dual;
        --先确定是否超出标准住院日
        v_标准住院日 := r_目标分支.标准住院日;
        If 源分支id_In = 0 Then
          Select Max(Nvl(结束天数, 开始天数))
          Into n_结束天数
          From 临床路径阶段
          Where 路径id = 源路径id_In And 版本号 = n_目标版本号 And Nvl(分支id, 0) = Nvl(源分支id_In, 0);
          If Instr(v_标准住院日, '-') > 0 Then
            If Substr(v_标准住院日, Instr(v_标准住院日, '-') + 1) < n_结束天数 Then
              v_标准住院日 := Substr(v_标准住院日, 1, Instr(v_标准住院日, '-')) || n_结束天数;
            End If;
          End If;
        End If;
        临床路径分支_Insert(源分支id_In, n_Branch_New_Id, 目标路径id_In, n_目标版本号, r_目标分支.名称, r_目标分支.说明, r_目标分支.前一阶段id, v_标准住院日,
                      r_目标分支.标准费用);
      End Loop;
    Else
      --从其他路径复制或是新增版本是
      Zl_临床路径版本_Delete(目标路径id_In, 目标版本号_In);
    End If;
  End If;
  If Nvl(是否分支路径_In, 0) <> 1 Then
    --从其他路径复制或是新增版本是
    --临床路径版本
    Insert Into 临床路径版本
      (路径id, 版本号, 标准住院日, 标准费用, 版本说明, 创建人, 创建时间)
      Select 目标路径id_In, n_目标版本号, 标准住院日, 标准费用, 版本说明, Zl_Username, Sysdate
      From 临床路径版本
      Where 路径id = 源路径id_In And 版本号 = n_源版本号;
    --路径导入评估
    Select Max(ID)
    Into n_Eval_Old_Id
    From 临床路径评估
    Where 路径id = 源路径id_In And 版本号 = n_源版本号 And 评估类型 = 1;
    If Nvl(n_Eval_Old_Id, 0) <> 0 Then
      Select 临床路径评估_Id.Nextval Into n_Eval_New_Id From Dual;
      临床路径评估_Insert(n_Eval_Old_Id, n_Eval_New_Id, 目标路径id_In, n_目标版本号, Null);
      ---路径评估指标
      For r_路径评估指标 In (Select ID From 路径评估指标 Where 评估id = n_Eval_Old_Id) Loop
        Select 路径评估指标_Id.Nextval Into n_Mark_New_Id From Dual;
        路径评估指标_Insert(r_路径评估指标.Id, n_Mark_New_Id, n_Eval_New_Id);
        ---路径评估条件
        路径评估条件_Insert(n_Eval_Old_Id, r_路径评估指标.Id, Null, n_Eval_New_Id, n_Mark_New_Id, Null);
      End Loop;
    End If;
  Else
    --从其他分支或主路径复制时
    Insert Into 临床路径分类
      (路径id, 版本号, 序号, 名称, 分支id)
      Select 目标路径id_In, n_目标版本号, 序号, 名称, n_Branch_New_Id
      From 临床路径分类
      Where 路径id = 源路径id_In And 版本号 = n_目标版本号 And Nvl(分支id, 0) = Nvl(源分支id_In, 0);
  
    For r_临床路径阶段 In (Select ID, 序号
                     From 临床路径阶段
                     Where 路径id = 源路径id_In And 版本号 = n_目标版本号 And 父id Is Null And Nvl(分支id, 0) = Nvl(源分支id_In, 0)
                     Order By 序号) Loop
      If Nvl(源分支id_In, 0) <> 0 Or r_临床路径阶段.序号 > n_前一阶段序号 Then
        --临床路径阶段的父级行插入
        Select 临床路径阶段_Id.Nextval Into n_Step_Parent_Id From Dual;
        临床路径阶段_Insert(r_临床路径阶段.Id, n_Step_Parent_Id, 目标路径id_In, n_目标版本号, Null, 源分支id_In, n_Branch_New_Id);
      
        临床路径阶段cascade_Insert(r_临床路径阶段.Id, n_Step_Parent_Id, 源路径id_In, 目标路径id_In, 源版本号_In, n_目标版本号, 源分支id_In,
                             n_Branch_New_Id);
        --临床路径阶段的子级行
        For r_临床路径子阶段 In (Select ID
                          From 临床路径阶段
                          Where 路径id = 源路径id_In And 版本号 = n_目标版本号 And 父id = r_临床路径阶段.Id And
                                Nvl(分支id, 0) = Nvl(源分支id_In, 0)) Loop
          --生成新的阶段ID
          Select 临床路径阶段_Id.Nextval Into n_Step_New_Id From Dual;
          临床路径阶段_Insert(r_临床路径子阶段.Id, n_Step_New_Id, 目标路径id_In, n_目标版本号, n_Step_Parent_Id, 源分支id_In, n_Branch_New_Id);
        
          临床路径阶段cascade_Insert(r_临床路径子阶段.Id, n_Step_New_Id, 源路径id_In, 目标路径id_In, 源版本号_In, n_目标版本号, 源分支id_In,
                               n_Branch_New_Id);
        End Loop;
      End If;
    End Loop;
  End If;

  --临床路径分支
  If Nvl(源分支id_In, 0) = 0 And Nvl(目标分支id_In, 0) = 0 Then
    --新增版本时
    For r_临床路径分支 In (Select ID From 临床路径分支 Where 路径id = 源路径id_In And 版本号 = n_源版本号) Loop
      Select 临床路径分支_Id.Nextval Into n_Branch_New_Id From Dual;
      临床路径分支_Insert(r_临床路径分支.Id, n_Branch_New_Id, 目标路径id_In, n_目标版本号);
    
      Insert Into 临床路径分类
        (路径id, 版本号, 序号, 名称, 分支id)
        Select 目标路径id_In, n_目标版本号, 序号, 名称, n_Branch_New_Id
        From 临床路径分类
        Where 路径id = 源路径id_In And 版本号 = n_源版本号 And 分支id = r_临床路径分支.Id;
    
      For r_临床路径阶段 In (Select ID
                       From 临床路径阶段
                       Where 路径id = 源路径id_In And 版本号 = n_源版本号 And 父id Is Null And 分支id = r_临床路径分支.Id
                       Order By 序号) Loop
        --临床路径阶段的父级行插入
        Select 临床路径阶段_Id.Nextval Into n_Step_Parent_Id From Dual;
        临床路径阶段_Insert(r_临床路径阶段.Id, n_Step_Parent_Id, 目标路径id_In, n_目标版本号, Null, r_临床路径分支.Id, n_Branch_New_Id);
      
        临床路径阶段cascade_Insert(r_临床路径阶段.Id, n_Step_Parent_Id, 源路径id_In, 目标路径id_In, 源版本号_In, n_目标版本号, r_临床路径分支.Id,
                             n_Branch_New_Id);
        --临床路径阶段的子级行
        For r_临床路径子阶段 In (Select ID
                          From 临床路径阶段
                          Where 路径id = 源路径id_In And 版本号 = n_源版本号 And 父id = r_临床路径阶段.Id And 分支id = r_临床路径分支.Id) Loop
          --生成新的阶段ID
          Select 临床路径阶段_Id.Nextval Into n_Step_New_Id From Dual;
          临床路径阶段_Insert(r_临床路径子阶段.Id, n_Step_New_Id, 目标路径id_In, n_目标版本号, n_Step_Parent_Id, r_临床路径分支.Id, n_Branch_New_Id);
        
          临床路径阶段cascade_Insert(r_临床路径子阶段.Id, n_Step_New_Id, 源路径id_In, 目标路径id_In, 源版本号_In, n_目标版本号, r_临床路径分支.Id,
                               n_Branch_New_Id);
        End Loop;
      End Loop;
    End Loop;
  
  End If;

  If Nvl(是否分支路径_In, 0) <> 1 Then
    --从其他路径复制或是新增版本是
    --临床路径分类
    Insert Into 临床路径分类
      (路径id, 版本号, 序号, 名称)
      Select 目标路径id_In, n_目标版本号, 序号, 名称
      From 临床路径分类
      Where 路径id = 源路径id_In And 版本号 = n_源版本号 And 分支id Is Null;
  
    --临床路径项目
    --临床路径医嘱
    --路径医嘱内容
    --临床路径病历
    --临床路径评估
    --路径评估指标
    --路径评估条件
  
    For r_临床路径阶段 In (Select ID
                     From 临床路径阶段
                     Where 路径id = 源路径id_In And 版本号 = n_源版本号 And 父id Is Null And 分支id Is Null
                     Order By 序号) Loop
      --临床路径阶段的父级行插入
      Select 临床路径阶段_Id.Nextval Into n_Step_Parent_Id From Dual;
      临床路径阶段_Insert(r_临床路径阶段.Id, n_Step_Parent_Id, 目标路径id_In, n_目标版本号, Null);
      If Nvl(源分支id_In, 0) = 0 And Nvl(目标分支id_In, 0) = 0 Then
        --新增版本时,更新前一阶段ID
        Update 临床路径分支
        Set 前一阶段id = n_Step_Parent_Id
        Where 前一阶段id = r_临床路径阶段.Id And 版本号 = n_目标版本号;
      End If;
    
      临床路径阶段cascade_Insert(r_临床路径阶段.Id, n_Step_Parent_Id, 源路径id_In, 目标路径id_In, 源版本号_In, n_目标版本号);
      --临床路径阶段的子级行
      For r_临床路径子阶段 In (Select ID
                        From 临床路径阶段
                        Where 路径id = 源路径id_In And 版本号 = n_源版本号 And 父id = r_临床路径阶段.Id And 分支id Is Null) Loop
        --生成新的阶段ID
        Select 临床路径阶段_Id.Nextval Into n_Step_New_Id From Dual;
        临床路径阶段_Insert(r_临床路径子阶段.Id, n_Step_New_Id, 目标路径id_In, n_目标版本号, n_Step_Parent_Id);
      
        临床路径阶段cascade_Insert(r_临床路径子阶段.Id, n_Step_New_Id, 源路径id_In, 目标路径id_In, 源版本号_In, n_目标版本号);
      End Loop;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_临床路径版本_Copy;
/

--108821:李业庆,2017-05-15,退料后未向库存回写条码信息
Create Or Replace Procedure Zl_材料收发记录_部门退料
(
  收发id_In   In 药品收发记录.Id%Type,
  审核人_In   In 药品收发记录.审核人%Type,
  审核日期_In In 药品收发记录.审核日期%Type,
  批号_In     In 药品库存.上次批号%Type := Null,
  效期_In     In 药品库存.效期%Type := Null,
  产地_In     In 药品库存.上次产地%Type := Null,
  退料数量_In In 药品收发记录.实际数量%Type := Null,
  自动销帐_In Integer := 0,
  退料人_In   In 药品收发记录.领用人%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);
  v_No      药品收发记录.No%Type;

  n_记录状态   药品收发记录.记录状态%Type;
  n_执行状态   住院费用记录.执行状态%Type;
  n_部分退料   Number;
  n_入出类别id Number(18);
  n_单据       药品收发记录.单据%Type;
  n_库房id     药品收发记录.库房id%Type;
  n_药品id     药品收发记录.药品id%Type;
  n_实际数量   药品收发记录.实际数量%Type;
  n_实际金额   药品收发记录.零售金额%Type;
  n_实际成本   药品收发记录.成本金额%Type;
  n_实际差价   药品收发记录.差价%Type;
  n_费用id     药品收发记录.费用id%Type;
  n_零售价     药品收发记录.零售价%Type;
  n_实价卫材   收费项目目录.是否变价%Type;

  --处理退料时，分批核算性质改变后的处理
  n_新批次       药品收发记录.批次%Type;
  n_批次         药品收发记录.批次%Type;
  n_分批         材料特性.在用分批%Type;
  n_小数         Number(2);
  n_上次供应商id 药品库存.上次供应商id%Type;
  n_成本价       药品收发记录.成本价%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;
  d_灭菌效期     药品库存.灭菌效期%Type;
  v_批准文号     药品库存.批准文号%Type;
  v_产地         药品收发记录.产地%Type;
  v_费用no       住院费用记录.No%Type;
  v_Temp         Varchar2(255);
  v_人员编号     人员表.编号%Type;
  v_人员姓名     人员表.姓名%Type;
  n_主页id       住院费用记录.主页id%Type;
  n_序号         住院费用记录.序号%Type;
  v_病人来源     病人医嘱记录.病人来源%Type;

  v_备货id     药品收发记录.Id%Type;
  v_入库no     药品收发记录.No%Type;
  v_入库序号   Number(5) := 0;
  v_执行时间   药品收发记录.审核日期%Type;
  n_平均成本价 药品库存.平均成本价%Type;
  n_冲销记录id 药品收发记录.Id%Type;
  n_移库       Number(1) := 0;
  v_商品条码   药品库存.商品条码%Type;
  v_内部条码   药品库存.内部条码%Type;
Begin
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_小数 From Dual;

  If 退料数量_In Is Not Null Then
    If 退料数量_In = 0 Then
      Return;
    End If;
  End If;

  --1、判断当前数据是否是备货卫材
  Begin
    Select 汇总发药号
    Into v_备货id
    From 药品收发记录
    Where 单据 = 21 And 审核日期 Is Not Null And
          汇总发药号 =
          (Select Max(a.Id)
           From 药品收发记录 A, 药品收发记录 B
           Where a.单据 = b.单据 And a.No = b.No And a.序号 = b.序号 And b.Id = 收发id_In And (Mod(a.记录状态, 3) = 1 Or a.记录状态 = 1)) And
          Rownum = 1;
  Exception
    When Others Then
      v_备货id := 0;
  End;

  --获取该收发记录的单据、药品ID、库房ID
  Select 单据, NO, 库房id, 药品id, 费用id, 入出类别id, 记录状态, Nvl(批次, 0), 生产日期, 灭菌效期, 批准文号, 供药单位id, 成本价, 产地, 零售价, 商品条码, 内部条码
  Into n_单据, v_No, n_库房id, n_药品id, n_费用id, n_入出类别id, n_记录状态, n_批次, d_上次生产日期, d_灭菌效期, v_批准文号, n_上次供应商id, n_成本价, v_产地,
       n_零售价, v_商品条码, v_内部条码
  From 药品收发记录
  Where ID = 收发id_In;

  --获取该笔记录剩余未退数量、金额及差价
  --尽量避免金额及差价未出完的现象
  Select Sum(Nvl(实际数量, 0) * Nvl(付数, 1)), Sum(Nvl(零售金额, 0)), Sum(Nvl(成本金额, 0)), Sum(Nvl(差价, 0))
  Into n_实际数量, n_实际金额, n_实际成本, n_实际差价
  From 药品收发记录
  Where 审核人 Is Not Null And NO = v_No And 单据 = n_单据 And 序号 = (Select 序号 From 药品收发记录 Where ID = 收发id_In);

  --如果允许退药数为零，表示已退药
  If n_实际数量 = 0 Then
    v_Err_Msg := '该单据已被其他操作员退料，请刷新后再试！';
    Raise Err_Item;
  End If;

  If Nvl(退料数量_In, 0) > n_实际数量 Then
    v_Err_Msg := '该单据已被其他操作员部分退料，请刷新后再试！';
    Raise Err_Item;
  End If;

  --获取该材料当前是否分批的信息
  Select Nvl(在用分批, 0) Into n_分批 From 材料特性 Where 材料id = n_药品id;

  --如果是部分退料，则重新计算零售金额及差价
  n_部分退料 := 0;
  If Not (退料数量_In Is Null Or Nvl(退料数量_In, 0) = n_实际数量) Then
    n_部分退料 := 1;
  End If;

  If n_部分退料 = 1 Then
    n_实际金额 := Round(n_实际金额 * 退料数量_In / n_实际数量, n_小数);
    n_实际成本 := Round(n_实际成本 * 退料数量_In / n_实际数量, n_小数);
    n_实际差价 := Round(n_实际差价 * 退料数量_In / n_实际数量, n_小数);
    n_实际数量 := 退料数量_In;
  End If;

  --n_分批:0-不分批;1-分批;2-原分批，现不分批，按不分批处理;3-原不分批，现分批，产生新批次
  If n_分批 = 0 And n_批次 <> 0 Then
    --原分批，现不分批，按不分批处理
    n_分批 := 2;
  Elsif n_分批 <> 0 And n_批次 = 0 Then
    --原不分批,现分批,产生新的批次，并在新产生的发药记录中使用
    n_分批 := 3;
  Else
    If n_批次 = 0 Then
      n_分批 := 0;
    Else
      n_分批 := 1;
    End If;
  End If;
  If 产地_In Is Not Null Then
    v_产地 := 产地_In;
  End If;
  --记录状态的含义有所变化
  --冲销的记录状态        :iif(n_记录状态=1,0,1)+1
  --被冲销的记录状态        :iif(n_记录状态=1,0,1)+2
  --等待发料的记录状态    :iif(n_记录状态=1,0,1)+3
  Select 药品收发记录_Id.Nextval Into n_冲销记录id From Dual;
  --产生冲销记录
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价,
     零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 领用人, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
    Select n_冲销记录id, n_记录状态 + Decode(n_记录状态, 1, 0, 1) + 1, n_单据, v_No, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号,
           效期, 灭菌效期, 1, -n_实际数量, -n_实际数量, 成本价, -n_实际成本, 扣率, 零售价, -n_实际金额, -n_实际差价, 摘要, 审核人_In, 审核日期_In, 配药人, 审核人_In,
           审核日期_In, 费用id, 单量, 频次, 用法, 发药窗口, 退料人_In, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码
    From 药品收发记录
    Where ID = 收发id_In;

  --如果是部分冲销，则付数填为1，实际数量为付数与实际数量的积
  --产生正常记录以供继续发料
  Select 药品收发记录_Id.Nextval Into n_新批次 From Dual;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 灭菌效期, 付数, 填写数量, 实际数量, 成本价, 成本金额, 扣率, 零售价,
     零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 审核人, 审核日期, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码)
    Select n_新批次, n_记录状态 + Decode(n_记录状态, 1, 0, 1) + 3, n_单据, v_No, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id,
           Decode(n_分批, 1, 批次, 3, n_新批次, Null), Decode(n_分批, 3, 产地_In, 1, 产地, Null), Decode(n_分批, 3, 批号_In, 1, 批号, Null),
           Decode(n_分批, 3, 效期_In, 1, 效期, Null), 灭菌效期, 1, n_实际数量, n_实际数量, 成本价, n_实际成本, 扣率, 零售价, n_实际金额, n_实际差价, 摘要, 填制人,
           填制日期, Null, Null, Null, 费用id, 单量, 频次, 用法, 发药窗口, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码
    From 药品收发记录
    Where ID = 收发id_In;

  --更新病人费用记录的执行状态(0-未执行;1-完全执行;2-部分执行)
  Select Decode(Sum(Nvl(付数, 1) * 实际数量), Null, 0, 0, 0, 2)
  Into n_执行状态
  From 药品收发记录
  Where 单据 = n_单据 And NO = v_No And 费用id = n_费用id And 审核人 Is Not Null;

  If n_执行状态 = 0 Then
    Update 住院费用记录 Set 执行状态 = n_执行状态, 执行人 = Null, 执行时间 = Null Where ID = n_费用id;
    Update 门诊费用记录
    Set 执行状态 = n_执行状态, 执行人 = Null, 执行时间 = Null
    Where NO = v_No And
          序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = 收发id_In)) And
          (Mod(记录性质, 10) = 1 Or Mod(记录性质, 10) = 2) And 记录状态 <> 2 And 执行部门id = n_库房id;
  Else
    Update 住院费用记录 Set 执行状态 = n_执行状态 Where ID = n_费用id;
    Update 门诊费用记录
    Set 执行状态 = n_执行状态
    Where NO = v_No And
          序号 = (Select 序号 From 门诊费用记录 Where ID = (Select 费用id From 药品收发记录 Where ID = 收发id_In)) And
          (Mod(记录性质, 10) = 1 Or Mod(记录性质, 10) = 2) And 记录状态 <> 2 And 执行部门id = n_库房id;
  End If;

  --插入未发药品记录
  Begin
    Insert Into 未发药品记录
      (单据, NO, 病人id, 主页id, 姓名, 优先级, 对方部门id, 库房id, 发药窗口, 填制日期, 已收费, 配药人, 打印状态, 未发数)
      Select a.单据, a.No, a.病人id, a.主页id, a.姓名, Nvl(b.优先级, 0) 优先级, a.对方部门id, a.库房id, a.发药窗口, a.填制日期, a.已收费, Null, 1, 1
      From (Select b.单据, b.No, a.病人id, a.主页id, a.姓名, Decode(a.记录性质, 1, Decode(a.操作员姓名, Null, 0, 1), 1) 已收费, b.对方部门id,
                    b.库房id, b.发药窗口, b.填制日期, c.身份
             From 住院费用记录 A, 药品收发记录 B, 病人信息 C
             Where b.Id = 收发id_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)
             Union All
             Select b.单据, b.No, a.病人id, Null As 主页id, a.姓名, Decode(a.记录性质, 1, Decode(a.操作员姓名, Null, 0, 1), 1) 已收费,
                    b.对方部门id, b.库房id, b.发药窗口, b.填制日期, c.身份
             From 门诊费用记录 A, 药品收发记录 B, 病人信息 C
             Where b.Id = 收发id_In And a.Id = b.费用id + 0 And a.病人id = c.病人id(+)) A, 身份 B
      Where b.名称(+) = a.身份;
  Exception
    When Others Then
      Null;
  End;

  --修改原记录为被冲销记录
  Update 药品收发记录 Set 记录状态 = n_记录状态 + Decode(n_记录状态, 1, 0, 1) + 2 Where ID = 收发id_In;

  --修改药品库存(反冲库存)
  Select 是否变价 Into n_实价卫材 From 收费项目目录 Where ID = n_药品id;

  If n_分批 <> 3 Then
  
    Update 药品库存
    Set 实际数量 = Nvl(实际数量, 0) + n_实际数量, 实际金额 = Nvl(实际金额, 0) + n_实际金额, 实际差价 = Nvl(实际差价, 0) + n_实际差价,
        零售价 = Decode(n_实价卫材, 1, Decode(Nvl(n_批次, 0), 0, Null, Decode(Nvl(零售价, 0), 0, n_零售价, 零售价)), Null)
    Where 库房id + 0 = n_库房id And 药品id = n_药品id And 性质 = 1 And Nvl(批次, 0) = n_批次;
  
    If Sql%RowCount = 0 Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价, 平均成本价, 商品条码,
         内部条码)
      Values
        (n_库房id, n_药品id, Decode(n_分批, 2, Null, n_批次), 1, n_实际数量, n_实际金额, n_实际差价, Decode(n_分批, 1, 效期_In, Null), d_灭菌效期,
         n_上次供应商id, n_成本价, Decode(n_分批, 1, 批号_In, Null), d_上次生产日期, v_产地, v_批准文号,
         Decode(n_实价卫材, 1, Decode(Nvl(n_批次, 0), 0, Null, n_零售价), Null), n_成本价, v_商品条码, v_内部条码);
    End If;
  Else
    Insert Into 药品库存
      (库房id, 药品id, 批次, 性质, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价, 平均成本价, 商品条码, 内部条码)
    Values
      (n_库房id, n_药品id, n_新批次, 1, n_实际数量, n_实际金额, n_实际差价, 效期_In, d_灭菌效期, n_上次供应商id, n_成本价, 批号_In, d_上次生产日期, v_产地, v_批准文号,
       Decode(n_实价卫材, 1, Decode(Nvl(n_新批次, 0), 0, Null, n_零售价), Null), n_成本价, v_商品条码, v_内部条码);
  End If;

  Delete 药品库存
  Where 库房id + 0 = n_库房id And 药品id = n_药品id And 性质 = 1 And Nvl(可用数量, 0) = 0 And Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And
        Nvl(实际差价, 0) = 0;

  If 自动销帐_In = 1 And n_单据 <> 24 Then
    Begin
      Select 主页id, NO, 序号 Into n_主页id, v_费用no, n_序号 From 住院费用记录 Where ID = n_费用id;
    Exception
      When Others Then
        Begin
          Select Null, NO, 序号 Into n_主页id, v_费用no, n_序号 From 门诊费用记录 Where ID = n_费用id;
        Exception
          When Others Then
            n_主页id := Null;
        End;
    End;
    If n_主页id Is Null Then
      Zl_门诊记帐记录_Delete(v_费用no, n_序号, v_人员编号, v_人员姓名);
    Else
      Zl_住院记帐记录_Delete(v_费用no, n_序号, v_人员编号, v_人员姓名);
    End If;
  End If;

  --备货卫材处理
  If v_备货id > 0 Then
    --2、自动冲销已审核的其他出库单据
    Begin
      Select 1
      Into n_移库
      From 药品收发记录
      Where 单据 = 15 And 审核日期 Is Null And
            费用id In (Select Distinct 费用id From 药品收发记录 Where NO = v_No And 药品id = n_药品id And 批次 = n_批次);
    Exception
      When Others Then
        n_移库 := 0;
    End;
    If n_移库 <> 0 Then
      For v_出库冲销 In (Select 1 行次, 记录状态, NO, 序号, 药品id
                     From 药品收发记录
                     Where 单据 = 21 And 审核日期 Is Not Null And 汇总发药号 = v_备货id) Loop
      
        Zl_材料其他出库_Strike(v_出库冲销.行次, v_出库冲销.记录状态, v_出库冲销.No, v_出库冲销.序号, v_出库冲销.药品id, 退料数量_In, 审核人_In, 审核日期_In, 1);
      End Loop;
    
      --3、产生新的其他出库单据
      If v_入库no Is Null Then
        v_入库no := Nextno(74, n_库房id);
      End If;
      v_入库序号 := v_入库序号 + 1;
    
      For v_入库 In (Select 入出类别id, 库房id, 药品id, 批次, 填写数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 产地, 批号, 效期, 灭菌效期, 摘要, 单量, 发药窗口
                   From 药品收发记录
                   Where 单据 = 21 And 审核日期 Is Not Null And 汇总发药号 = v_备货id) Loop
      
        Zl_材料其他出库_Insert(v_入库.入出类别id, v_入库no, v_入库序号, v_入库.库房id, v_入库.药品id, v_入库.批次, v_入库.填写数量, v_入库.成本价, v_入库.成本金额,
                         v_入库.零售价, v_入库.零售金额, v_入库.差价, 审核人_In, 审核日期_In, v_入库.产地, v_入库.批号, v_入库.效期, v_入库.灭菌效期, v_入库.摘要,
                         v_入库.单量, v_入库.发药窗口);
      
        Update 药品收发记录
        Set 费用id = n_费用id, 汇总发药号 = n_新批次
        Where 单据 = 21 And NO = v_入库no And 序号 = v_入库序号;
      End Loop;
    
      --4、删除未审核的外购入库单据（已审核则不管）
      Delete 药品收发记录
      Where 单据 = 15 And 药品id = n_药品id And Nvl(批次, 0) = n_批次 And 费用id = n_费用id And 审核日期 Is Null;
    End If;
  End If;
  --处理调价修正单据
  Zl_材料收发记录_调价修正(n_冲销记录id);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料收发记录_部门退料;
/

--109194:李业庆,2017-05-23,盘点单允许多个新增批次
Create Or Replace Procedure Zl_药品盘点记录单_Insert
(
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  库房id_In     In 药品收发记录.库房id%Type,
  批次_In       In 药品收发记录.批次%Type,
  入出类别id_In In 药品收发记录.入出类别id%Type,
  入出系数_In   In 药品收发记录.入出系数%Type,
  药品id_In     In 药品收发记录.药品id%Type,
  帐面数量_In   In 药品收发记录.填写数量%Type,
  实盘数量_In   In 药品收发记录.扣率%Type,
  数量差_In     In 药品收发记录.实际数量%Type,
  售价_In       In 药品收发记录.零售价%Type,
  金额差_In     In 药品收发记录.零售金额%Type,
  差价差_In     In 药品收发记录.差价%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  摘要_In       In 药品收发记录.摘要%Type := Null,
  产地_In       In 药品收发记录.产地%Type := Null,
  批号_In       In 药品收发记录.批号%Type := Null,
  效期_In       In 药品收发记录.效期%Type := Null,
  盘点时间_In   In 药品收发记录.频次%Type := Null,
  库存金额_In   In 药品收发记录.成本价%Type := Null,
  库存差价_In   In 药品收发记录.成本金额%Type := Null,
  批准文号_In   In 药品收发记录.批准文号%Type := Null,
  成本价_In     In 药品收发记录.单量%Type := Null,
  库房货位_In   In 药品收发记录.库房货位%Type := Null
) Is
  v_批次 药品收发记录.批次%Type;
Begin
  v_批次 := 批次_In;
  If v_批次 < 0 Then
    v_批次 := Zl_Fun_Getbatchnum(药品id_In, 产地_In, 批号_In, 成本价_In, 售价_In, 批次_In);
  End If;

  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 扣率, 实际数量, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 频次,
     成本价, 成本金额, 批准文号, 单量, 库房货位)
  Values
    (药品收发记录_Id.Nextval, 1, 14, No_In, 序号_In, 库房id_In, 入出类别id_In, 入出系数_In, 药品id_In, v_批次, 产地_In, 批号_In, 效期_In, 帐面数量_In,
     实盘数量_In, 数量差_In, 售价_In, 金额差_In, 差价差_In, 摘要_In, 填制人_In, 填制日期_In, 盘点时间_In, 库存金额_In, 库存差价_In, 批准文号_In, 成本价_In, 库房货位_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品盘点记录单_Insert;
/


--109194:李业庆,2017-05-23,盘点单允许多个新增批次
Create Or Replace Function Zl_Fun_Getbatchnum
(
  药品id_In   药品批号对照.药品id%Type,
  生产厂家_In 药品批号对照.生产厂家%Type,
  批号_In     药品批号对照.批号%Type,
  成本价_In   药品批号对照.成本价%Type,
  售价_In     药品批号对照.售价%Type,
  新批次_In   药品批号对照.批次%Type
) Return Number Is
  --功能：药品入库产生入库记录时根据传递过来的参数找对应的批次
  --返回值：查询到的批次，如果批次>0则说明找到了批次,如果批次=0则说明没有找到
  --参数：
  --     生产厂家_in：入库传递过来的生产商
  --     批号_in：入库时录入的批号
  --     成本价_in 入库时的成本价
  --     售价_in  入库时的售价
  --
  n_批次     药品批号对照.批次%Type;
  n_药库包装 药品规格.药库包装%Type;
  n_是否变价 收费项目目录.是否变价%Type;
  n_Count    Number(1);
Begin
  --只处理生产厂家和批号不为空的情况
  If 生产厂家_In Is Not Null And 批号_In Is Not Null Then
    Begin
      Select 批次
      Into n_批次
      From 药品批号对照
      Where 药品id = 药品id_In And Nvl(生产厂家, 'a') = Nvl(生产厂家_In, 'a') And Nvl(批号, 'b') = Nvl(批号_In, 'b') And 成本价 = 成本价_In And
            售价 = 售价_In;
    Exception
      When Others Then
        n_批次 := 新批次_In;
      
        If n_批次 > 0 Then
          --检查有无重复记录
          Begin
            Select 1
            Into n_Count
            From 药品批号对照
            Where 药品id = 药品id_In And Nvl(生产厂家, 'a') = Nvl(生产厂家_In, 'a') And Nvl(批号, 'b') = Nvl(批号_In, 'b') And
                  批次 = n_批次;
          Exception
            When Others Then
              n_Count := 0;
          End;
        
          --没有重复记录才能插入
          If n_Count = 0 Then
            Insert Into 药品批号对照
              (药品id, 生产厂家, 批号, 批次, 成本价, 售价)
            Values
              (药品id_In, 生产厂家_In, 批号_In, 新批次_In, 成本价_In, 售价_In);
          End If;
        End If;
    End;
  End If;

  Return(n_批次);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Getbatchnum;
/

---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--104825:黄捷,2017-04-07,RIS接口中医嘱窗体修改成独立exe
Insert Into Zlfilesupgrade
  (文件类型, 文件名, 版本号, 修改日期, 所属系统, 业务部件, 安装路径, 文件说明, 强制覆盖, 自动注册, 加入日期, 序号)
  Select 1, 'ZL9XWINTERFACE.DLL', '', Null, '1,21', 'ZLSVRSTUDIO.EXE,ZLHIS+.EXE,zl9BaseItem.dll,zl9CISJob.dll,zl9PACSWork.dll,zlCISKernel.dll,zl9peimanage.dll', '[APPSOFT]\APPLY', 'XWRIS接口部件', '1', '1',
         Sysdate, 序号
  From Dual a, (Select Max(To_Number(序号)) + 1 序号 From Zlfilesupgrade) b
  Where Not Exists (Select 1 From Zlfilesupgrade Where Upper(文件名) = 'ZL9XWINTERFACE.DLL');

Insert Into Zlfilesupgrade
  (文件类型, 文件名, 版本号, 修改日期, 所属系统, 业务部件, 安装路径, 文件说明, 强制覆盖, 自动注册, 加入日期, 序号)
  Select 1, 'ZLSOFTSHOWHISFORMS.EXE', '', Null, '1', 'zl9XWInterface.dll', '[APPSOFT]\APPLY',
         'RIS查看HIS中电子病历，门诊医嘱，住院医嘱等功能的独立exe程序', '1', '0', Sysdate, 序号
  From Dual a, (Select Max(To_Number(序号)) + 1 序号 From Zlfilesupgrade) b
  Where Not Exists (Select 1 From Zlfilesupgrade Where Upper(文件名) = 'ZLSOFTSHOWHISFORMS.EXE');

--系统版本号
Update zlSystems Set 版本号='10.34.110' Where 编号=&n_System;
--部件版本号
Commit;