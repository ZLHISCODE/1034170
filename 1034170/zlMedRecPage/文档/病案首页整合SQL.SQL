Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( &n_System,'病原学目录',0,'标本对应的病原学代码已经名称','医疗工作' ); 
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( &n_System,'医院感染目录',0,'医院感染的标准分类名称','医疗工作' ); 
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( &n_System,'器械导管目录',0,'重症监护常用器械目录','医疗工作' ); 
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( &n_System,'Icu类型',0,'重症监护的类别目录','医疗工作' ); 

CREATE TABLE 病人过敏记录(
    ID NUMBER(18),
    病人ID NUMBER(18),
    主页ID NUMBER(18),
    记录来源 NUMBER(1),
    药物ID NUMBER(18),
    药物名 VARCHAR2(60),
    结果 NUMBER(1),
    过敏时间 DATE,
    记录时间 DATE,
    记录人 VARCHAR2(20),
    过敏反应 varchar2(100),
    过敏源编码 Varchar2(10),
    待转出 Number(3))
    TABLESPACE zl9MedRec
    PCTFREE 5;

Alter Table 病人过敏记录 Add Constraint 病人过敏记录_PK Primary Key (ID) Using Index Tablespace zl9IndexMdr;
Alter Table 病人过敏记录 Add Constraint 病人过敏记录_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID) On Delete Cascade;

Create Index 病人过敏记录_IX_病人ID On 病人过敏记录(病人ID) Tablespace zl9IndexMdr;
Create Index 病人过敏记录_IX_待转出 On 病人过敏记录(待转出) Tablespace zl9IndexMdr;

Create Sequence 病人过敏记录_ID Start With 1;
Update zlBaseCode  set 说明='病人的住院死亡原因' where 系统=&n_System and  表名='住院死亡原因'
Create Sequence 病人信息_ID Start With 1;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,200,'增加修改',User,'Zl_病人手麻记录_Update','EXECUTE' From Dual Union All 
Select &n_System,200,'增加修改',User,'Zl_病人过敏记录_Update','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl_病人感染记录_Update','EXECUTE' From Dual Union All 
Select &n_System,200,'增加修改',User,'Zl_病人病原学检查_Update','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl_器械导管使用情况_Update','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl_病案重症监护情况_Update','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl_病案化疗记录_Update','EXECUTE' From Dual Union All 
Select &n_System,200,'增加修改',User,'Zl_病案放疗记录_Update','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl_病案精神治疗_Update','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl_病人过敏记录_Delete','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl_病人过敏记录_Insert','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl3_病人信息_首页整理','EXECUTE' From Dual  Union All
Select &n_System,200,'增加修改',User,'Zl3_病案主页_首页整理','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl3_病人诊断记录_Delete','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl3_病人诊断记录_Update','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl3_病人诊断记录_Insert','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl_病人费用_Delete','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl_分娩信息_Delete','EXECUTE' From Dual  Union All
Select &n_System,200,'增加修改',User,'ZL_病案主页从表_首页整理','EXECUTE' From Dual Union All
Select &n_System,200,'增加修改',User,'Zl_病人手麻记录_Update','EXECUTE' From Dual  Union All
Select &n_System,200,'增加修改',User,'Zl_病人手麻记录_Delete','EXECUTE' From Dual 


 
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,200,'基本',User,'病人过敏药物','SELECT' From Dual Union All
Select &n_System,200,'基本',User,'Zl3_病案主页_首页整理','SELECT' From Dual

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,200,'删除病案',User,'Zl3_病案主页_首页整理','SELECT' From Dual

Drop PROCEDURE ZL_病人过敏药物_病案删除;
Drop PROCEDURE ZL_病人过敏药物_病案增加;
drop Procedure ZL_随诊记录_reset;
drop Procedure ZL_病人信息_病案修改
drop Procedure ZL_病人信息_病案增加
drop Procedure ZL_病案主页_病案修改
drop Procedure Zl_住院病案记录_病案修改
drop Procedure ZL_病案主页_病案增加
drop Procedure ZL_病案主页_删除附加信息
drop Procedure ZL_病案主页从表_INSERT
drop Procedure ZL_病案主页_取消编目
drop Procedure ZL_病案主页_取消编码
drop Procedure Zl_病人诊断情况_Insert
drop Procedure Zl_病案主页从表_Insert
--200:病案管理(增加修改)
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_病人过敏药物_病案删除';
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_病人过敏药物_病案增加';
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_病人信息_病案修改';
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_病人信息_病案增加';
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_病案主页_病案修改';
Delete zlProgPrivs Where 系统=&n_System And 对象='Zl_住院病案记录_病案修改';
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_病案主页_病案增加';
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_病案主页_删除附加信息';
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_病案主页从表_INSERT';
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_病案主页_取消编目';
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_病案主页_取消编码';
Delete zlProgPrivs Where 系统=&n_System And 对象='Zl_病人诊断情况_Insert';
Delete zlProgPrivs Where 系统=&n_System And 对象='Zl_病案主页从表_Insert';
--200:病案管理(增加修改)
--202:随诊管理(增删改)
Delete zlProgPrivs Where 系统=&n_System And 对象='ZL_随诊记录_reset';

Create Or Replace Function Zl_获取住院次数或主页id
(
  病人id_In 病案主页.病人id%Type,
  次数_In   病案主页.主页id%Type,
  标识_In   Integer
)
-------------------------------------------------------------------------------------------------------------------------------------- 
  --功能:获取住院次数或主页ID 
  --入参: 
  --      病人ID_IN:病人ID值 
  --      次数_IN:当标识_IN为1时传入实际的住院次数(排除了留观病人),当标识_IN为0时,表示传入的是主页id. 
  -- 标识_IN:1-返回指定次数的主页id,0-根据主页id返回住院次数(排除了留观病人) 
  --返回:根据病人ID，返回主页id或实际住院次数 
  -------------------------------------------------------------------------------------------------------------------------------------- 

 Return Integer Is
  N_次数 病案主页.主页id%Type;
  N_留观 病案主页.主页id%Type;
  N_住院 病案主页.主页id%Type;
Begin
  N_留观 := 0;
  N_住院 := 0;
  N_次数 := 0;
  For Rs In (Select 主页id, Nvl(病人性质, 0) 病人性质
             From 病案主页
             Where 病人id = 病人id_In And 主页id <> 0
             Order By 主页id) Loop
    If Rs.病人性质 <> 0 Then
      N_留观 := N_留观 + 1;
    Else
      N_住院 := N_住院 + 1;
    End If;
    If 标识_In = 1 Then
      --根据住院次数，获取主页id   
      If N_住院 <= 次数_In Then
        N_次数 := Rs.主页id;
      End If;
      Exit When(N_住院 >= 次数_In);
    Else
      --根据主页id，获取实际的住院次数 
      If Rs.主页id <= 次数_In Then
        N_次数 := N_住院;
      End If;
      Exit When(Rs.主页id >= 次数_In);
    End If;
  End Loop;
  Return N_次数;
Exception
  When Others Then
    Return 0;
End Zl_获取住院次数或主页id;
/
Create Or Replace Procedure Zl_病案主页从表_首页整理
(
  病人id_In 病案主页从表.病人id%Type,
  主页id_In 病案主页从表.主页id%Type,
  信息名_In 病案主页从表.信息名%Type,
  信息值_In 病案主页从表.信息值%Type
) Is
  --功能：用于住院医护站对病人进行首页整理
Begin
  If 信息值_In Is Null Then
    Delete From 病案主页从表 Where 病人id = 病人id_In And 主页id = 主页id_In And Upper(信息名) = Upper(信息名_In);
  Else
    Update 病案主页从表
    Set 信息值 = 信息值_In
    Where 病人id = 病人id_In And 主页id = 主页id_In And Upper(信息名) = Upper(信息名_In);
    If Sql%RowCount = 0 Then
      Insert Into 病案主页从表 (病人id, 主页id, 信息名, 信息值) Values (病人id_In, 主页id_In, 信息名_In, 信息值_In);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案主页从表_首页整理;
/
Create Or Replace Function Zl3_Get护理天数
(
  病人id_In   病案主页.病人id%Type,
  主页id_In   病案主页.主页id%Type,
  项目名称_In Varchar2
) Return Varchar2 Is
  N_护理id 病人变动记录.护理等级id%Type;
  N_天数   Number(18);
  V_Sql    Varchar2(200);
  V_名称   Varchar2(200);
  V_名称1  Varchar2(200);
  V_Return Varchar2(500);
Begin
  For Rs In (Select Column_Value 项目 From Table(Cast(F_Str2list(项目名称_In) As Zltools.T_Strlist))) Loop
    N_天数 := 0;
    If Not Rs.项目 Is Null Then
      --需要将I或一或规换成统一的查找 
      If Rs.项目 = '一级护理' Then
        V_名称  := Replace(项目名称_In, '一', 'I');
        V_名称1 := Replace(项目名称_In, '一', 'Ⅰ');
      Elsif Rs.项目 = '二级护理' Then
        V_名称  := Replace(项目名称_In, '二', 'II');
        V_名称1 := Replace(项目名称_In, '二', 'Ⅱ');
      Elsif Rs.项目 = '三级护理' Then
        V_名称  := Replace(项目名称_In, '三', 'III');
        V_名称1 := Replace(项目名称_In, '三', 'Ⅲ');
      Elsif Rs.项目 = 'CCU' Then
        V_名称  := 'ccu';
        V_名称1 := 'Ccu';
      Elsif Rs.项目 = 'Icu' Then
        V_名称  := 'Icu ';
        V_名称1 := 'Icu';
      End If;
      Begin
        V_Sql := ' Select ID From 收费项目目录 Where(名称 = :1 Or 名称 = :2 Or 名称 = :3) And Rownum = 1 '; --动态SQL 
        Execute Immediate V_Sql
          Into N_护理id
          Using 项目名称_In, V_名称, V_名称1;
      Exception
        When Others Then
          N_护理id := 0;
      End;
    
      If N_护理id <= 0 Then
        N_天数 := 0;
      Else
        Select Sum(终止时间 - 开始时间)
        Into N_天数
        From (Select Distinct 护理等级id, Trunc(开始时间) As 开始时间, Trunc(Decode(终止原因, 1, 1, 0) + 终止时间) As 终止时间
               From 病人变动记录
               Where 病人id = 病人id_In And 主页id = 主页id_In And 护理等级id = N_护理id);
      End If;
    End If;
    V_Return := V_Return || '|' || N_天数;
  End Loop;
  If V_Return Is Not Null Then
    V_Return := Substr(V_Return, 2);
  End If;
  Return V_Return;
Exception
  When Others Then
    Return Null;
End Zl3_Get护理天数;
/

Create Or Replace Function Zl3_Nextno
(
  序号_In In 号码控制表.项目序号%Type,
  判断_In In Number := Null,
  编码_In In 部门表.编码%Type := Null
) Return Varchar2
--    功能：根据特定规则产生新的号码,规则如下： 
  --    一、项目序号： 
  --       1   病人ID         数字 
  --       2   住院号         数字 
  --       4   病案号         数字 
  --    二、年度位确定原则： 
  --       以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码 
  -- 
  --    说明：最大号码-10存入号码控制表,用于并发情况下补缺号(取了号,但未使用) 
  --          For Update在并发情况下锁定行,不用Wait选项以避免向调用者返回空 
  --    返回：最大号码 
 Is
  Pragma Autonomous_Transaction;
  V_No        号码控制表.最大号码%Type;
  V_Maxno     号码控制表.最大号码%Type;
  N_Maxno     Number;
  N_Mod       号码控制表.编号规则%Type;
  V_Tmp       Varchar2(10);
  V_Site_No   Varchar2(1);
  V_Code      Varchar2(20);
  N_Codelenth Number;

  V_Error Varchar2(255);
  Err_Custom Exception;
Begin
  V_Site_No := F_Get_Node_No;

  --1.病人ID 
  If 序号_In = 1 Then
    Select Nvl(编号规则, 0) Into N_Mod From 号码控制表 Where 项目序号 = 序号_In;
  
    --从序列取值，用于不要求病人ID必须连续的用户减少并发争用 
    If N_Mod = 1 Then
      Select 病人信息_Id.Nextval Into V_No From Dual;
    Else
      Select Nvl(最大号码, '0') Into V_Maxno From 号码控制表 Where 项目序号 = 序号_In For Update;
      If V_Site_No Is Null Then
        Select Nvl(Max(病人id), 0) + 1 Into N_Maxno From 病人信息 Where 病人id >= To_Number(V_Maxno);
      Else
        Select Trunc(Nvl(Max(病人id), 0) / 10) + 1
        Into N_Maxno
        From 病人信息
        Where 病人id >= To_Number(V_Maxno) * 10 And Mod(病人id, 10) = To_Number(V_Site_No);
      End If;
      Update 号码控制表 Set 最大号码 = Decode(Sign(N_Maxno - 10), 1, N_Maxno - 10, 1) Where 项目序号 = 序号_In;
      V_No := To_Char(N_Maxno) || V_Site_No;
    End If;
    --2.住院号 
  Elsif 序号_In = 2 Then
    Select Nvl(最大号码, '0'), Nvl(编号规则, 0)
    Into V_Maxno, N_Mod
    From 号码控制表
    Where 项目序号 = 序号_In
    For Update;
  
    If N_Mod = 0 Then
      --0.顺序编号 
      If V_Site_No Is Null Then
        Select Nvl(Max(住院号), 0) + 1 Into N_Maxno From 病案主页 Where 住院号 >= To_Number(V_Maxno);
      Else
        Select Trunc(Nvl(Max(住院号), 0) / 10) + 1
        Into N_Maxno
        From 病案主页
        Where 住院号 >= To_Number(V_Maxno) * 10 And Mod(住院号, 10) = To_Number(V_Site_No);
      End If;
    
      Update 号码控制表 Set 最大号码 = Decode(Sign(N_Maxno - 10), 1, N_Maxno - 10, 1) Where 项目序号 = 序号_In;
    Elsif N_Mod = 1 Then
      --1.年月(YYMM)+顺序号(0000) 
      V_Tmp := To_Char(Sysdate, 'YYMM');
      If V_Site_No Is Null Then
        Select Nvl(Max(住院号), To_Number(V_Tmp || '0000')) + 1
        Into N_Maxno
        From 病案主页
        Where 住院号 Like To_Number(V_Tmp) || '%' And 住院号 >= To_Number(V_Maxno);
      Else
        Select Trunc(Nvl(Max(住院号), To_Number(V_Tmp || '0000')) / 10) + 1
        Into N_Maxno
        From 病案主页
        Where 住院号 Like To_Number(V_Tmp) || '%' And 住院号 >= To_Number(V_Maxno) * 10 And
              Mod(住院号, 10) = To_Number(V_Site_No);
      End If;
    
      Update 号码控制表
      Set 最大号码 = Decode(Sign(N_Maxno - 10 - To_Number(V_Tmp || '0000')), 1, N_Maxno - 10, To_Number(V_Tmp || '0001'))
      Where 项目序号 = 序号_In;
    Elsif N_Mod = 2 Then
      --2.年(YYYY)+顺序号(00000) 
      V_Tmp := To_Char(Sysdate, 'YYYY');
      If V_Site_No Is Null Then
        Select Nvl(Max(住院号), To_Number(V_Tmp || '00000')) + 1
        Into N_Maxno
        From 病案主页
        Where 住院号 Like To_Number(V_Tmp) || '%' And 住院号 >= To_Number(V_Maxno);
      Else
        Select Trunc(Nvl(Max(住院号), To_Number(V_Tmp || '00000')) / 10) + 1
        Into N_Maxno
        From 病案主页
        Where 住院号 Like To_Number(V_Tmp) || '%' And 住院号 >= To_Number(V_Maxno) * 10 And
              Mod(住院号, 10) = To_Number(V_Site_No);
      End If;
    
      Update 号码控制表
      Set 最大号码 = Decode(Sign(N_Maxno - 10 - To_Number(V_Tmp || '00000')), 1, N_Maxno - 10, To_Number(V_Tmp || '00001'))
      Where 项目序号 = 序号_In;
    End If;
    V_No := To_Char(N_Maxno) || V_Site_No;
  
    --4.病案号 
  Elsif 序号_In = 4 Then
    Select Nvl(最大号码, '0'), Nvl(编号规则, 0)
    Into V_Maxno, N_Mod
    From 号码控制表
    Where 项目序号 = 序号_In
    For Update;
  
    If 判断_In = 1 And N_Mod <> 0 Then
      V_No := ' ';
      Return V_No;
    End If;
  
    If N_Mod = 0 Then
      --0.顺序编号 
      If V_Site_No Is Null Then
        --Select Nvl(Max(病案号), '') Into v_Maxno From 住院病案记录 Where 病案号 >= v_Maxno; 
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := To_Char(To_Number(Nvl(V_Maxno, '0') + 1));
          V_Maxno := LTrim(RTrim(V_Maxno));
        End If;
        If 判断_In = 1 Then
          Update 号码控制表 Set 最大号码 = V_Maxno Where 项目序号 = 序号_In;
        End If;
      Else
        --第1位为:站点号,后面为顺序号+1 
        --Select Nvl(Max(病案号), '') 
        --Into v_Maxno 
        --From 住院病案记录 
        --Where 病案号 >= v_Maxno And 病案号 Like v_Site_No || '%'; 
      
        V_Maxno := Substr(Nvl(V_Maxno, '0'), 2);
      
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := To_Char(To_Number(Nvl(V_Maxno, '0') + 1));
          V_Maxno := LTrim(RTrim(V_Maxno));
        End If;
        V_Maxno := V_Site_No || V_Maxno;
        If 判断_In = 1 Then
          Update 号码控制表 Set 最大号码 = V_Maxno Where 项目序号 = 序号_In;
        End If;
      End If;
      V_No := V_Maxno;
    Elsif N_Mod = 1 Then
      --1.年月(YYMM)+顺序号(00000) 
      V_Tmp := To_Char(Sysdate, 'YYMM');
      If V_Site_No Is Null Then
        Select Nvl(Max(病案号), To_Number(V_Tmp || '00000'))
        Into V_Maxno
        From 住院病案记录
        Where 病案号 Like V_Tmp || '%' And 病案号 >= V_Maxno;
        V_Maxno := Substr(V_Maxno, 5);
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 5, '0');
        End If;
      Else
        --年月(YYMM)+站点号+顺序号(0000) 
        Select Nvl(Max(病案号), To_Number(V_Tmp || '00000'))
        Into V_Maxno
        From 住院病案记录
        Where 病案号 Like V_Tmp || V_Site_No || '%' And 病案号 >= V_Maxno;
      
        V_Maxno := Substr(V_Maxno, 6);
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || V_Site_No || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || V_Site_No || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 4, '0');
        End If;
      End If;
      Update 号码控制表 Set 最大号码 = V_Maxno Where 项目序号 = 序号_In;
      V_No := V_Maxno;
    Elsif N_Mod = 2 Then
      --2.年(YYYY)+顺序号(00000) 
      V_Tmp := To_Char(Sysdate, 'YYYY');
    
      If V_Site_No Is Null Then
        Select Nvl(Max(病案号), To_Number(V_Tmp || '00000'))
        Into V_Maxno
        From 住院病案记录
        Where 病案号 Like V_Tmp || '%' And 病案号 >= V_Maxno;
        V_Maxno := Substr(V_Maxno, 5);
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 5, '0');
        End If;
      
      Else
        --年月(YYYY)+站点号+顺序号(0000) 
        Select Nvl(Max(病案号), To_Number(V_Tmp || '00000'))
        Into V_Maxno
        From 住院病案记录
        Where 病案号 Like V_Tmp || V_Site_No || '%' And 病案号 >= V_Maxno;
        V_Maxno := Substr(V_Maxno, 6);
      
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || V_Site_No || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || V_Site_No || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 4, '0');
        End If;
      End If;
      Update 号码控制表 Set 最大号码 = V_Maxno Where 项目序号 = 序号_In;
      V_No := V_Maxno;
    Elsif N_Mod = 3 Then
      --3.年(YYYY)+科室编码+顺序号(00000) 
      V_Tmp := To_Char(Sysdate, 'YYYY');
      If 编码_In Is Null Then
        V_Code := '';
      Else
        V_Code := 编码_In;
      End If;
    
      --获取年+科室编码的固定长度 
      N_Codelenth := Length(V_Tmp || V_Code) + 1;
    
      If V_Site_No Is Null Then
        Select Nvl(Max(病案号), To_Number(V_Tmp || V_Code || '00000'))
        Into V_Maxno
        From 住院病案记录
        Where 病案号 Like V_Tmp || V_Code || '%' And 病案号 >= V_Maxno;
      
        V_Maxno := Substr(V_Maxno, N_Codelenth);
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || V_Code || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || V_Code || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 5, '0');
        End If;
      
      Else
        --年月(YYYY)+站点号+顺序号(0000) 
        Select Nvl(Max(病案号), To_Number(V_Tmp || V_Code || '00000'))
        Into V_Maxno
        From 住院病案记录
        Where 病案号 Like V_Tmp || V_Code || V_Site_No || '%' And 病案号 >= V_Maxno;
        V_Maxno := Substr(V_Maxno, N_Codelenth + 1);
      
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || V_Code || V_Site_No || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || V_Code || V_Site_No || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 4, '0');
        End If;
      End If;
      Update 号码控制表 Set 最大号码 = V_Maxno Where 项目序号 = 序号_In;
      V_No := V_Maxno;
    End If;
    --档案号 
  Elsif 序号_In = 5 Then
    Select Nvl(最大号码, '0'), Nvl(编号规则, 0)
    Into V_Maxno, N_Mod
    From 号码控制表
    Where 项目序号 = 序号_In
    For Update;
  
    If N_Mod = 0 Then
      --0.顺序编号 
      If V_Site_No Is Null Then
        Select Substr(Nvl(Max(LPad(档案号, 20, 0)), ''), 20 - Nvl(Max(Length(档案号)), 20) + 1, Nvl(Max(Length(档案号)), 20))
        Into V_Maxno
        From 住院病案记录
        Where LPad(档案号, 20, 0) >= LPad(V_Maxno, 20, 0) And 档案号 > '0';
      
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := To_Char(To_Number(Nvl(V_Maxno, '0') + 1));
          V_Maxno := LTrim(RTrim(V_Maxno));
        End If;
      Else
        --第1位为:站点号,后面为顺序号+1 
        Select Substr(Nvl(Max(LPad(档案号, 20, 0)), ''), 20 - Nvl(Max(Length(档案号)), 20) + 1, Nvl(Max(Length(档案号)), 20))
        Into V_Maxno
        From 住院病案记录
        Where LPad(档案号, 20, 0) >= LPad(V_Maxno, 20, 0) And 档案号 > '0' And 档案号 Like V_Site_No || '%';
      
        V_Maxno := Substr(Nvl(V_Maxno, '0'), 2);
      
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := To_Char(To_Number(Nvl(V_Maxno, '0') + 1));
          V_Maxno := LTrim(RTrim(V_Maxno));
        End If;
        V_Maxno := V_Site_No || V_Maxno;
      End If;
      If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 1 Then
        Update 号码控制表 Set 最大号码 = Decode(Sign(V_Maxno - 10), 1, V_Maxno - 10, 1) Where 项目序号 = 序号_In;
      End If;
      V_No := V_Maxno;
    Elsif N_Mod = 1 Then
      --1.年月(YYMM)+顺序号(00000) 
      V_Tmp := To_Char(Sysdate, 'YYMM');
      If V_Site_No Is Null Then
        Select Substr(Nvl(Max(LPad(档案号, 20, 0)), To_Number(V_Tmp || '00000')), 20 - Nvl(Max(Length(档案号)), 20) + 1,
                       Nvl(Max(Length(档案号)), 20))
        Into V_Maxno
        From 住院病案记录
        Where 档案号 Like V_Tmp || '%' And LPad(档案号, 20, 0) >= LPad(V_Maxno, 20, 0) And 档案号 > '0';
        V_Maxno := Substr(V_Maxno, 5);
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 5, '0');
        End If;
      Else
        --年月(YYMM)+站点号+顺序号(0000) 
        Select Substr(Nvl(Max(LPad(档案号, 20, 0)), To_Number(V_Tmp || '00000')), 20 - Nvl(Max(Length(档案号)), 20) + 1,
                       Nvl(Max(Length(档案号)), 20))
        Into V_Maxno
        From 住院病案记录
        Where 档案号 Like V_Tmp || V_Site_No || '%' And LPad(档案号, 20, 0) >= LPad(V_Maxno, 20, 0) And 档案号 > '0';
      
        V_Maxno := Substr(V_Maxno, 6);
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || V_Site_No || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || V_Site_No || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 4, '0');
        End If;
      End If;
      If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 1 Then
        Update 号码控制表
        Set 最大号码 = Decode(Sign(V_Maxno - 10 - To_Number(V_Tmp || '00000')), 1, V_Maxno - 10, To_Number(V_Tmp || '00001'))
        Where 项目序号 = 序号_In;
      End If;
      V_No := V_Maxno;
    Elsif N_Mod = 2 Then
      --2.年(YYYY)+顺序号(00000) 
      V_Tmp := To_Char(Sysdate, 'YYYY');
    
      If V_Site_No Is Null Then
        Select Substr(Nvl(Max(LPad(档案号, 20, 0)), To_Number(V_Tmp || '00000')), 20 - Nvl(Max(Length(档案号)), 20) + 1,
                       Nvl(Max(Length(档案号)), 20))
        Into V_Maxno
        From 住院病案记录
        Where 档案号 Like V_Tmp || '%' And LPad(档案号, 20, 0) >= LPad(V_Maxno, 20, 0) And 档案号 > '0';
        V_Maxno := Substr(V_Maxno, 5);
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 5, '0');
        End If;
      Else
        --年月(YYYY)+站点号+顺序号(0000) 
        Select Substr(Nvl(Max(LPad(档案号, 20, 0)), To_Number(V_Tmp || '00000')), 20 - Nvl(Max(Length(档案号)), 20) + 1,
                       Nvl(Max(Length(档案号)), 20))
        Into V_Maxno
        From 住院病案记录
        Where 档案号 Like V_Tmp || V_Site_No || '%' And LPad(档案号, 20, 0) >= LPad(V_Maxno, 20, 0) And 档案号 > '0';
        V_Maxno := Substr(V_Maxno, 6);
      
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || V_Site_No || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || V_Site_No || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 4, '0');
        End If;
      End If;
      If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 1 Then
        Update 号码控制表
        Set 最大号码 = Decode(Sign(V_Maxno - 10 - To_Number(V_Tmp || '00000')), 1, V_Maxno - 10, To_Number(V_Tmp || '00001'))
        Where 项目序号 = 序号_In;
      End If;
      V_No := V_Maxno;
    Elsif N_Mod = 3 Then
      --3.年(YYYY)+科室编码+顺序号(00000) 
      V_Tmp := To_Char(Sysdate, 'YYYY');
      If 编码_In Is Null Then
        V_Code := '';
      Else
        V_Code := 编码_In;
      End If;
    
      --获取年+科室编码的固定长度 
      N_Codelenth := Length(V_Tmp || V_Code) + 1;
    
      If V_Site_No Is Null Then
        Select Substr(Nvl(Max(LPad(档案号, 20, 0)), To_Number(V_Tmp || V_Code || '00000')),
                       20 - Nvl(Max(Length(档案号)), 20) + 1, Nvl(Max(Length(档案号)), 20))
        Into V_Maxno
        From 住院病案记录
        Where 档案号 Like V_Tmp || V_Code || '%' And LPad(档案号, 20, 0) >= LPad(V_Maxno, 20, 0) And 档案号 > '0';
      
        V_Maxno := Substr(V_Maxno, N_Codelenth);
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || V_Code || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || V_Code || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 5, '0');
        End If;
      
      Else
        --年月(YYYY)+站点号+顺序号(0000) 
        Select Substr(Nvl(Max(LPad(档案号, 20, 0)), To_Number(V_Tmp || V_Code || '00000')),
                       20 - Nvl(Max(Length(档案号)), 20) + 1, Nvl(Max(Length(档案号)), 20))
        Into V_Maxno
        From 住院病案记录
        Where 档案号 Like V_Tmp || V_Code || V_Site_No || '%' And LPad(档案号, 20, 0) >= LPad(V_Maxno, 20, 0) And 档案号 > '0';
        V_Maxno := Substr(V_Maxno, N_Codelenth + 1);
      
        If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 0 Then
          --是字符,需处理字符加1 
          V_Maxno := V_Tmp || V_Code || V_Site_No || Zl3_Incstr(V_Maxno);
        Else
          V_Maxno := V_Tmp || V_Code || V_Site_No || LPad(To_Number(Nvl(V_Maxno, '0') + 1), 4, '0');
        End If;
      End If;
      If Zl3_Isnumber(Nvl(V_Maxno, '0')) = 1 Then
        Update 号码控制表
        Set 最大号码 = Decode(Sign(V_Maxno - 10 - To_Number(V_Tmp || V_Code || '00000')), 1, V_Maxno - 10,
                           To_Number(V_Tmp || V_Code || '00001'))
        Where 项目序号 = 序号_In;
      End If;
      V_No := V_Maxno;
    End If;
  Else
    V_Error := '序号为' || 序号_In || '的号码,其规则值:' || N_Mod || ',当前系统不支持！';
    Raise Err_Custom;
  End If;

  Commit;
  Return V_No;
Exception
  When Err_Custom Then
    Rollback;
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Error || '[ZLSOFT]');
  When Others Then
    Rollback;
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl3_Nextno;
/

Create Or Replace Procedure Zl_病人过敏记录_Delete
(
  病人id_In  病人过敏记录.病人id%Type,
  主页id_In  病人过敏记录.主页id%Type,
  来源_In    病人过敏记录.记录来源%Type := Null,
  过敏ids_In Varchar2 := Null
) Is
Begin
  If 过敏ids_In Is Null Then
    Delete From 病人过敏记录
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And (记录来源 = 来源_In Or 来源_In Is Null);
  Else
    Delete From 病人过敏记录
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And (记录来源 = 来源_In Or 来源_In Is Null) And
          ID In (Select Column_Value From Table(F_Str2list(过敏ids_In)));
  End If;
  Delete From 病人过敏药物
  Where (过敏药物id, 过敏药物) Not In (Select Distinct 药物id, 药物名 From 病人过敏记录 Where 病人id = 病人id_In) And 病人id = 病人id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人过敏记录_Delete;
/
Create Or Replace Procedure Zl_病人过敏记录_Update
(
  Id_In         病人过敏记录.Id%Type,
  病人id_In     病人过敏记录.病人id%Type,
  主页id_In     病人过敏记录.主页id%Type,
  来源_In       病人过敏记录.记录来源%Type,
  药物id_In     病人过敏记录.药物id%Type,
  药物名_In     病人过敏记录.药物名%Type,
  结果_In       病人过敏记录.结果%Type := 1,
  过敏时间_In   病人过敏记录.过敏时间%Type := Null,
  过敏反应_In   病人过敏记录.过敏反应%Type := Null,
  过敏源编码_In 病人过敏记录.过敏源编码%Type := Null
) Is
  V_Temp     Varchar2(255);
  V_人员姓名 病人过敏记录.记录人%Type;
  N_Count    Number;
Begin
  --先判断当前指定的过敏信息是指定的来源的诊断，是则更新，否则新增
  Select Count(1) Into N_Count From 病人过敏记录 Where ID = Id_In And (记录来源 = 来源_In Or 来源_In Is Null);
  If N_Count = 1 Then
    Update 病人过敏记录
    Set 记录来源 = 来源_In, 药物id = 药物id_In, 药物名 = 药物名_In, 结果 = 结果_In, 过敏时间 = Nvl(过敏时间_In, Sysdate), 过敏反应 = 过敏反应_In,
        过敏源编码 = 过敏源编码_In
    Where ID = Id_In;
  Else
    V_Temp     := Zl_Identity;
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ';') + 1);
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ',') + 1);
    V_人员姓名 := Substr(V_Temp, Instr(V_Temp, ',') + 1);
  
    Insert Into 病人过敏记录
      (ID, 病人id, 主页id, 记录来源, 药物id, 药物名, 结果, 过敏时间, 记录时间, 记录人, 过敏反应, 过敏源编码)
      Select 病人过敏记录_Id.Nextval, 病人id_In, 主页id_In, 来源_In, 药物id_In, 药物名_In, 结果_In, Nvl(过敏时间_In, Sysdate), Sysdate, V_人员姓名,
             过敏反应_In, 过敏源编码_In
      From Dual;
  End If;

  Select Count(1) Into N_Count From 病人过敏药物 Where 病人id = 病人id_In And 过敏药物 = 药物名_In;
  If N_Count = 0 Then
    Insert Into 病人过敏药物
      (病人id, 过敏药物id, 过敏药物, 过敏反应)
    Values
      (病人id_In, 药物id_In, 药物名_In, 过敏反应_In);
  Else
    Update 病人过敏药物
    Set 过敏药物id = 药物id_In, 过敏反应 = 过敏反应_In
    Where 病人id = 病人id_In And 过敏药物 = 药物名_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人过敏记录_Update;
/
Create Or Replace Procedure Zl_病人过敏记录_Insert
(
  病人id_In     病人过敏记录.病人id%Type,
  主页id_In     病人过敏记录.主页id%Type,
  来源_In       病人过敏记录.记录来源%Type,
  药物id_In     病人过敏记录.药物id%Type,
  药物名_In     病人过敏记录.药物名%Type,
  结果_In       病人过敏记录.结果%Type := 1,
  过敏时间_In   病人过敏记录.过敏时间%Type := Null,
  记录时间_In   病人过敏记录.记录时间%Type := Null,
  过敏反应_In   病人过敏记录.过敏反应%Type := Null,
  过敏源编码_In 病人过敏记录.过敏源编码%Type := Null
) Is
  V_Date     Date;
  V_Gdate    Date;
  V_Temp     Varchar2(255);
  V_人员姓名 病人过敏记录.记录人%Type;
  N_Count    Number;
Begin
  V_Temp     := Zl_Identity;
  V_Temp     := Substr(V_Temp, Instr(V_Temp, ';') + 1);
  V_Temp     := Substr(V_Temp, Instr(V_Temp, ',') + 1);
  V_人员姓名 := Substr(V_Temp, Instr(V_Temp, ',') + 1);

  If 记录时间_In Is Not Null Then
    V_Date := 记录时间_In;
  Else
    Select Sysdate Into V_Date From Dual;
  End If;

  If 过敏时间_In Is Not Null Then
    V_Gdate := 过敏时间_In;
  Else
    Select Sysdate Into V_Gdate From Dual;
  End If;

  Insert Into 病人过敏记录
    (ID, 病人id, 主页id, 记录来源, 药物id, 药物名, 结果, 过敏时间, 记录时间, 记录人, 过敏反应,过敏源编码)
  Values
    (病人过敏记录_Id.Nextval, 病人id_In, 主页id_In, 来源_In, 药物id_In, 药物名_In, 结果_In, V_Gdate, V_Date, V_人员姓名, 过敏反应_In,过敏源编码_In);

  Select Count(1) Into N_Count From 病人过敏药物 Where 病人id = 病人id_In And 过敏药物 = 药物名_In;
  If N_Count = 0 Then
    Insert Into 病人过敏药物
      (病人id, 过敏药物id, 过敏药物, 过敏反应)
    Values
      (病人id_In, 药物id_In, 药物名_In, 过敏反应_In);
  Else
    Update 病人过敏药物
    Set 过敏药物id = 药物id_In, 过敏反应 = 过敏反应_In
    Where 病人id = 病人id_In And 过敏药物 = 药物名_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人过敏记录_Insert;
/
Create Or Replace Procedure Zl3_病人信息_首页整理
(
  --功能：用于对病案首页进行首页整理
  --N_Mode=0,修改病案首页，1-新增病案首页，2-删除病案首页
  N_Mode          Number,
  病人id_In       病人信息.病人id%Type,
  住院号_In       病人信息.住院号%Type := Null,
  姓名_In         病人信息.姓名%Type := Null,
  性别_In         病人信息.性别%Type := Null,
  年龄_In         病人信息.年龄%Type := Null,
  民族_In         病人信息.民族%Type := Null,
  国籍_In         病人信息.国籍%Type := Null,
  区域_In         病人信息.区域%Type := Null,
  籍贯_In         病人信息.籍贯%Type := Null,
  职业_In         病人信息.职业%Type := Null,
  出生日期_In     病人信息.出生日期%Type := Null,
  出生地点_In     病人信息.出生地点%Type := Null,
  身份证号_In     病人信息.身份证号%Type := Null,
  其他证件_In     病人信息.其他证件%Type := Null,
  婚姻状况_In     病人信息.婚姻状况%Type := Null,
  医疗付款方式_In 病人信息.医疗付款方式%Type := Null,
  家庭地址_In     病人信息.家庭地址%Type := Null,
  家庭电话_In     病人信息.家庭电话%Type := Null,
  家庭地址邮编_In 病人信息.家庭地址邮编%Type := Null,
  户口地址_In     病案主页.户口地址%Type := Null := Null,
  户口邮编_In     病案主页.户口地址邮编%Type := Null,
  工作单位id_In   病人信息.合同单位id%Type := Null,
  工作单位_In     病人信息.工作单位%Type := Null,
  单位电话_In     病人信息.单位电话%Type := Null,
  单位邮编_In     病人信息.单位邮编%Type := Null,
  联系人姓名_In   病人信息.联系人姓名%Type := Null,
  联系人关系_In   病人信息.联系人关系%Type := Null,
  联系人电话_In   病人信息.联系人电话%Type := Null,
  联系人地址_In   病人信息.联系人地址%Type := Null,
  Email_In        病人信息.Email%Type := Null,
  Qq_In           病人信息.Qq%Type := Null,
  入院时间_In     病人信息.入院时间%Type := Null,
  出院时间_In     病人信息.出院时间%Type := Null,
  住院次数_In     病人信息.住院次数%Type := 1
) Is
  N_住院次数 病案主页.主页id%Type;
Begin
  If N_Mode = 0 Then
    Select Count(病人id) Into N_住院次数 From 病案主页 Where 病人id = 病人id_In;
    ----修改最后一次病案主页 
    --如果在修改以前的病案主页，那病人信息只需要改住院号，这样才不会影响到它保存最新的信息 
    Update 病人信息
    Set 住院号 = Decode(住院次数_In, N_住院次数, 住院号_In, 住院号), 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 出生日期 = 出生日期_In, 出生地点 = 出生地点_In,
        身份证号 = 身份证号_In, 民族 = 民族_In, 国籍 = 国籍_In, 区域 = 区域_In, 婚姻状况 = 婚姻状况_In, 职业 = 职业_In, 医疗付款方式 = 医疗付款方式_In,
        家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In, 家庭地址邮编 = 家庭地址邮编_In, 工作单位 = 工作单位_In, 单位电话 = 单位电话_In, 单位邮编 = 单位邮编_In,
        联系人姓名 = Decode(联系人姓名_In, Null, 联系人姓名, 联系人姓名_In), 联系人关系 = Decode(联系人姓名_In, Null, 联系人关系, 联系人关系_In),
        联系人电话 = Decode(联系人姓名_In, Null, 联系人电话, 联系人电话_In), 联系人地址 = Decode(联系人姓名_In, Null, 联系人地址, 联系人地址_In), 其他证件 = 其他证件_In,
        合同单位id = 工作单位id_In, 户口地址 = 户口地址_In, 户口地址邮编 = 户口邮编_In, 籍贯 = 籍贯_In, Email = Email_In, Qq = Qq_In, 住院次数 = 住院次数_In,
        入院时间 = 入院时间_In, 出院时间 = 出院时间_In
    Where 病人id = 病人id_In;
  Elsif N_Mode = 1 Then
    ----只有第一次入院才新增病人信息，所以入院次数是固定值 
    Insert Into 病人信息
      (病人id, 住院号, 姓名, 性别, 年龄, 出生日期, 出生地点, 身份证号, 民族, 国籍, 区域, 婚姻状况, 职业, 医疗付款方式, 家庭地址, 家庭电话, 家庭地址邮编, 工作单位, 单位电话, 单位邮编,
       联系人姓名, 联系人关系, 联系人电话, 联系人地址, 其他证件, 合同单位id, 户口地址, 户口地址邮编, 籍贯, Email, Qq, 住院次数, 登记时间, 入院时间, 出院时间)
    Values
      (病人id_In, 住院号_In, 姓名_In, 性别_In, 年龄_In, 出生日期_In, 出生地点_In, 身份证号_In, 民族_In, 国籍_In, 区域_In, 婚姻状况_In, 职业_In, 医疗付款方式_In,
       家庭地址_In, 家庭电话_In, 家庭地址邮编_In, 工作单位_In, 单位电话_In, 单位邮编_In, 联系人姓名_In, 联系人关系_In, 联系人电话_In, 联系人地址_In, 其他证件_In,
       工作单位id_In, 户口地址_In, 户口邮编_In, 籍贯_In, Email_In, Qq_In, 1, Sysdate, 入院时间_In, 出院时间_In);
  Elsif N_Mode = 2 Then
    Delete From 病人信息 Where 病人id = 病人id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl3_病人信息_首页整理;
/
Create Or Replace Procedure Zl3_病案主页_首页整理
(
  --功能：用于对病案首页进行首页整理
  --N_Mode=0,修改病案主页，1-新增病案主页，2-删除病案主页，3-取消编码，4-取消编目
  N_Mode          Number,
  病人id_In       病案主页.病人id%Type,
  主页id_In       病案主页.主页id%Type,
  住院号_In       病案主页.住院号%Type := Null,
  病案号_In       病案主页.病案号%Type := Null,
  档案号_In       住院病案记录.档案号%Type := Null,
  年龄_In         病案主页.年龄%Type := Null,
  国籍_In         病案主页.国籍%Type := Null,
  区域_In         病案主页.区域%Type := Null,
  职业_In         病案主页.职业%Type := Null,
  身高_In         病案主页.身高%Type := Null,
  体重_In         病案主页.体重%Type := Null,
  血型_In         病案主页.血型%Type := Null,
  婚姻状况_In     病案主页.婚姻状况%Type := Null,
  医疗付款方式_In 病案主页.医疗付款方式%Type := Null,
  家庭地址_In     病案主页.家庭地址%Type := Null,
  家庭电话_In     病案主页.家庭电话%Type := Null,
  家庭邮编_In     病案主页.家庭地址邮编%Type := Null,
  户口地址_In     病案主页.户口地址%Type := Null,
  户口邮编_In     病案主页.户口地址邮编%Type := Null,
  单位地址_In     病案主页.单位地址%Type := Null,
  单位电话_In     病案主页.单位电话%Type := Null,
  单位邮编_In     病案主页.单位邮编%Type := Null,
  联系人姓名_In   病案主页.联系人姓名%Type := Null,
  联系人关系_In   病案主页.联系人关系%Type := Null,
  联系人电话_In   病案主页.联系人电话%Type := Null,
  联系人地址_In   病案主页.联系人地址%Type := Null,
  入院病况_In     病案主页.入院病况%Type := Null,
  入院方式_In     病案主页.入院方式%Type := Null,
  入院科室id_In   病案主页.入院科室id%Type := Null,
  入院日期_In     病案主页.入院日期%Type := Null,
  出院方式_In     病案主页.出院方式%Type := Null,
  出院科室id_In   病案主页.出院科室id%Type := Null,
  出院日期_In     病案主页.出院日期%Type := Null,
  再入院_In       病案主页.再入院%Type := Null,
  是否确诊_In     病案主页.是否确诊%Type := Null,
  确诊日期_In     病案主页.确诊日期%Type := Null,
  尸检标志_In     病案主页.尸检标志%Type := Null,
  随诊标志_In     病案主页.随诊标志%Type := Null,
  随诊期限_In     病案主页.随诊期限%Type := Null,
  新发肿瘤_In     病案主页.新发肿瘤%Type := Null,
  中医治疗类别_In 病案主页.中医治疗类别%Type := Null,
  抢救次数_In     病案主页.抢救次数%Type := Null,
  成功次数_In     病案主页.成功次数%Type := Null,
  住院天数_In     病案主页.住院天数%Type := Null,
  费用和_In       病案主页.费用和%Type := Null,
  门诊医师_In     病案主页.门诊医师%Type := Null,
  住院医师_In     病案主页.住院医师%Type := Null,
  责任护士_In     病案主页.责任护士%Type := Null,
  编目员姓名_In   病案主页.编目员姓名%Type := Null,
  编目日期_In     病案主页.编目日期%Type := Null
) As
  D_随诊期限 随诊记录.随诊期限%Type;
  D_应发时间 随诊记录.应发时间%Type;
  N_Count    Number(8);
  V_Err_Msg  Varchar2(100);
  Err_Item Exception;
Begin
  If N_Mode = 0 Then
    Update 病案主页
    Set 住院号 = 住院号_In, 病案号 = 病案号_In, 年龄 = 年龄_In, 职业 = 职业_In, 国籍 = 国籍_In, 区域 = 区域_In, 身高 = 身高_In, 体重 = 体重_In, 血型 = 血型_In,
        婚姻状况 = 婚姻状况_In, 医疗付款方式 = 医疗付款方式_In, 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In, 家庭地址邮编 = 家庭邮编_In, 单位电话 = 单位电话_In,
        单位邮编 = 单位邮编_In, 单位地址 = 单位地址_In, 联系人姓名 = 联系人姓名_In, 联系人关系 = 联系人关系_In, 联系人地址 = 联系人地址_In, 联系人电话 = 联系人电话_In,
        户口地址 = 户口地址_In, 户口地址邮编 = 户口邮编_In, 入院科室id = 入院科室id_In, 入院日期 = 入院日期_In, 入院病况 = 入院病况_In, 入院方式 = 入院方式_In,
        出院方式 = 出院方式_In, 出院科室id = 出院科室id_In, 出院日期 = 出院日期_In, 再入院 = 再入院_In, 是否确诊 = 是否确诊_In, 确诊日期 = 确诊日期_In, 尸检标志 = 尸检标志_In,
        随诊标志 = 随诊标志_In, 随诊期限 = 随诊期限_In, 新发肿瘤 = 新发肿瘤_In, 中医治疗类别 = 中医治疗类别_In, 抢救次数 = 抢救次数_In, 成功次数 = 成功次数_In,
        住院天数 = 住院天数_In, 费用和 = 费用和_In, 门诊医师 = 门诊医师_In, 住院医师 = 住院医师_In, 责任护士 = 责任护士_In, 编目员姓名 = 编目员姓名_In, 编目日期 = 编目日期_In
    Where 病人id = 病人id_In And 主页id = 主页id_In;
  Elsif N_Mode = 1 Then
    Insert Into 病案主页
      (病人id, 主页id, 住院号, 病案号, 年龄, 职业, 国籍, 区域, 身高, 体重, 血型, 婚姻状况, 医疗付款方式, 家庭地址, 家庭电话, 家庭地址邮编, 单位电话, 单位邮编, 单位地址, 联系人姓名,
       联系人关系, 联系人地址, 联系人电话, 户口地址, 户口地址邮编, 入院科室id, 入院日期, 入院病况, 入院方式, 出院方式, 出院科室id, 出院日期, 再入院, 是否确诊, 确诊日期, 尸检标志, 随诊标志,
       随诊期限, 新发肿瘤, 中医治疗类别, 抢救次数, 成功次数, 住院天数, 费用和, 门诊医师, 住院医师, 责任护士, 编目员姓名, 编目日期, 登记人, 登记时间)
    Values
      (病人id_In, 主页id_In, 住院号_In, 病案号_In, 年龄_In, 职业_In, 国籍_In, 区域_In, 身高_In, 体重_In, 血型_In, 婚姻状况_In, 医疗付款方式_In, 家庭地址_In,
       家庭电话_In, 家庭邮编_In, 单位电话_In, 单位邮编_In, 单位地址_In, 联系人姓名_In, 联系人关系_In, 联系人地址_In, 联系人电话_In, 户口地址_In, 户口邮编_In, 入院科室id_In,
       入院日期_In, 入院病况_In, 入院方式_In, 出院方式_In, 出院科室id_In, 出院日期_In, 再入院_In, 是否确诊_In, 确诊日期_In, 尸检标志_In, 随诊标志_In, 随诊期限_In,
       新发肿瘤_In, 中医治疗类别_In, 抢救次数_In, 成功次数_In, 住院天数_In, 费用和_In, 门诊医师_In, 住院医师_In, 责任护士_In, 编目员姓名_In, 编目日期_In, 编目员姓名_In,
       编目日期_In);
    --刘兴宏:增加主页时，需要增加住院次数
    Update 病人信息 Set 住院次数 = 主页id_In Where 病人id = 病人id_In;
  Elsif N_Mode = 2 Then
    Delete From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Elsif N_Mode = 3 Or N_Mode = 4 Then
    -- 取消编目 与 取消编码的区别 
    -- 1.取消编目 不删除病案主页从表
    -- 2.取消编目 不删除随诊记录 
    -- 3.取消编目 不删除诊断符合情况 
  
    --取消编码\编目标志 
    Update 病案主页
    Set 编目员编号 = Decode(N_Mode, 4, Null, 编目员编号), 编目员姓名 = Decode(N_Mode, 4, Null, 编目员姓名), 编目日期 = Null,
        病案号 = Decode(N_Mode, 3, Null, 编目员姓名)
    Where 病人id = 病人id_In And 主页id = 主页id_In;
    --删除一些附表 
    If N_Mode = 3 Then
      Delete From 病案主页从表 Where 病人id = 病人id_In And 主页id = 主页id_In;
      Delete From 诊断符合情况 Where 病人id = 病人id_In And 主页id = 主页id_In;
      Delete From 随诊记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
    End If;
    Delete From 病人费用 Where 病人id = 病人id_In And 主页id = 主页id_In;
    Delete From 病人手麻记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 记录来源 = 4;
    Delete From 病人诊断记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 记录来源 = 4;
    Delete From 病人分娩信息 Where 病人id = 病人id_In And 主页id = 主页id_In;
    Delete From 新生儿诊断记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
  
    --检查是否有借阅记录,如果有,则不能取消 
    Begin
      Select 'Err' Into V_Err_Msg From 借阅记录 Where 病人id = 病人id_In And 主页id = 主页id_In And Rownum = 1;
    Exception
      When Others Then
        V_Err_Msg := 'NO';
    End;
    If V_Err_Msg = 'Err' Then
      V_Err_Msg := '[ZLSOFT]已经存在借阅数据,不能进行取消' || Decode(N_Mode, 3, '编码', '编目') || '![ZLSOFT]';
    End If;
    Delete From 住院病案记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
  End If;

  If N_Mode = 1 Or N_Mode = 0 Then
    --随诊记录
    If Nvl(随诊标志_In, 0) = 0 Then
      Zl_随诊记录_Delete(病人id_In, 主页id_In);
    Else
      --终身随诊
      If Nvl(随诊标志_In, 0) = 9 Then
        D_随诊期限 := To_Date('3000-01-01', 'YYYY-MM-DD HH24:MI');
        --月
      Elsif Nvl(随诊标志_In, 0) = 1 Then
        D_随诊期限 := Add_Months(Nvl(出院日期_In, Sysdate), Nvl(随诊期限_In, 0));
        --年
      Elsif Nvl(随诊标志_In, 0) = 2 Then
        D_随诊期限 := Add_Months(Nvl(出院日期_In, Sysdate), Nvl(随诊期限_In, 0) * 12);
        --周
      Elsif Nvl(随诊标志_In, 0) = 3 Then
        D_随诊期限 := Nvl(出院日期_In, Sysdate) + Nvl(随诊期限_In, 0) * 7;
        --天
      Elsif Nvl(随诊标志_In, 0) = 4 Then
        D_随诊期限 := Nvl(出院日期_In, Sysdate) + Nvl(随诊期限_In, 0);
      End If;
      --缺省为三个月后
      If Months_Between(D_随诊期限, Nvl(出院日期_In, Sysdate)) > 3 Then
        D_应发时间 := Add_Months(Nvl(出院日期_In, Sysdate), 3);
        --当随诊期限少于三个月时以实际期限为准
      Else
        D_应发时间 := D_随诊期限;
      End If;
    
      Select Count(1) Into N_Count From 随诊记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
      If N_Count = 0 Then
        Insert Into 随诊记录
          (ID, 主页id, 病人id, 随诊期限, 应发时间)
        Values
          (随诊记录_Id.Nextval, 主页id_In, 病人id_In, D_随诊期限, D_应发时间);
      Else
        Update 随诊记录 Set 随诊期限 = D_随诊期限 Where 病人id = 病人id_In And 主页id = 主页id_In And 实发时间 Is Null;
      End If;
    End If;
    Zl_住院病案记录_病案修改(病人id_In, 主页id_In, 病案号_In, 档案号_In, '一般', '在院', '院内');
  Elsif N_Mode <> 4 Then
    Zl_病案重症监护情况_Delete(病人id_In, 主页id_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, V_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl3_病案主页_首页整理;
/

Create Or Replace Procedure Zl3_病人诊断记录_Delete
(
  病人id_In    病人诊断记录.病人id%Type,
  主页id_In    病人诊断记录.主页id%Type,
  记录来源_In  病人诊断记录.记录来源%Type := Null,
  病历id_In    病人诊断记录.病历id%Type := Null,
  诊断类型s_In Varchar2 := Null,
  诊断ids_In   Varchar2 := Null
) Is
Begin
  If Not 诊断ids_In Is Null Then
    Delete From 病人诊断记录 Where ID In (Select Column_Value From Table(F_Str2list(诊断ids_In)));
  Else
    Delete From 病人诊断记录
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And
          诊断类型 In (Select Column_Value From Table(F_Str2list(诊断类型s_In))) And 记录来源 = 记录来源_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl3_病人诊断记录_Delete;
/
Create Or Replace Procedure Zl3_病人诊断记录_Update
(
  Id_In       病人诊断记录.Id%Type,
  诊断类型_In 病人诊断记录.诊断类型%Type,
  疾病id_In   病人诊断记录.疾病id%Type,
  诊断id_In   病人诊断记录.诊断id%Type,
  证候id_In   病人诊断记录.证候id%Type,
  诊断描述_In 病人诊断记录.诊断描述%Type,
  出院情况_In 病人诊断记录.出院情况%Type,
  是否未治_In 病人诊断记录.是否未治%Type,
  是否疑诊_In 病人诊断记录.是否疑诊%Type,
  诊断次序_In 病人诊断记录.诊断次序%Type := 1,
  备注_In     病人诊断记录.备注%Type := Null,
  入院病情_In 病人诊断记录.入院病情%Type := Null,
  编码序号_In 病人诊断记录.编码序号%Type
) Is
  --功能：更新病人诊断记录 
  V_诊断id 病人诊断记录.Id%Type;
  N_Count  Number;

  V_Temp     Varchar2(255);
  V_人员姓名 人员表.姓名%Type;
Begin
  --先判断当前诊断是指定的诊断来源的诊断，是则更新，否则新增
  Select Count(1) Into N_Count From 病人诊断记录 Where ID = Id_In And 记录来源 = 4;
  If N_Count = 1 Then
    Update 病人诊断记录
    Set 诊断次序 = 诊断次序_In, 疾病id = 疾病id_In, 诊断id = 诊断id_In, 证候id = 证候id_In, 诊断描述 = 诊断描述_In, 入院病情 = 入院病情_In, 出院情况 = 出院情况_In,
        是否未治 = 是否未治_In, 是否疑诊 = 是否疑诊_In, 备注 = 备注_In
    Where ID = Id_In;
  Else
    Select 病人诊断记录_Id.Nextval Into V_诊断id From Dual;
  
    V_Temp     := Zl_Identity;
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ';') + 1);
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ',') + 1);
    V_人员姓名 := Substr(V_Temp, Instr(V_Temp, ',') + 1);
    Insert Into 病人诊断记录
      (ID, 病人id, 主页id, 记录来源, 诊断类型, 诊断次序, 编码序号, 疾病id, 诊断id, 证候id, 诊断描述, 入院病情, 出院情况, 是否未治, 是否疑诊, 记录日期, 记录人, 备注)
      Select V_诊断id, 病人id, 主页id, 4, 诊断类型_In, 诊断次序_In, 编码序号_In, 疾病id_In, 诊断id_In, 证候id_In, 诊断描述_In, 入院病情_In, 出院情况_In,
             是否未治_In, 是否疑诊_In, Sysdate, V_人员姓名, 备注_In
      From 病人诊断记录
      Where ID = Id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl3_病人诊断记录_Update;
/
Create Or Replace Procedure Zl3_病人诊断记录_Insert
(
  病人id_In   病人诊断记录.病人id%Type,
  主页id_In   病人诊断记录.主页id%Type,
  诊断类型_In 病人诊断记录.诊断类型%Type,
  疾病id_In   病人诊断记录.疾病id%Type,
  诊断id_In   病人诊断记录.诊断id%Type,
  证候id_In   病人诊断记录.证候id%Type,
  诊断描述_In 病人诊断记录.诊断描述%Type,
  出院情况_In 病人诊断记录.出院情况%Type,
  是否未治_In 病人诊断记录.是否未治%Type,
  是否疑诊_In 病人诊断记录.是否疑诊%Type,
  诊断次序_In 病人诊断记录.诊断次序%Type := 1,
  备注_In     病人诊断记录.备注%Type := Null,
  入院病情_In 病人诊断记录.入院病情%Type := Null,
  编码序号_In 病人诊断记录.编码序号%Type
) Is
  V_记录人  病人诊断记录.记录人%Type;
  N_Id      病人诊断记录.Id%Type;
  V_Temp    Varchar2(255);
  V_Err_Msg Varchar2(100);
  Err_Item Exception;
Begin
  V_Temp   := Zl_Identity;
  V_Temp   := Substr(V_Temp, Instr(V_Temp, ';') + 1);
  V_Temp   := Substr(V_Temp, Instr(V_Temp, ',') + 1);
  V_记录人 := Substr(V_Temp, Instr(V_Temp, ',') + 1);
  Select 病人诊断记录_Id.Nextval Into N_Id From Dual;
  If V_记录人 Is Null Then
    V_Err_Msg := '[ZLSOFT]当前用户未设置对应的人员信息,请与系统管理员' || Chr(10) || Chr(13) || '联系,先到用户授权管理中设置！[ZLSOFT]';
    Raise Err_Item;
  End If;

  Insert Into 病人诊断记录
    (ID, 病人id, 主页id, 医嘱id, 记录来源, 病历id, 诊断类型, 疾病id, 诊断id, 证候id, 诊断描述, 入院病情, 出院情况, 是否未治, 是否疑诊, 记录日期, 记录人, 取消时间, 取消人, 编码序号,
     诊断次序, 备注)
  Values
    (N_Id, 病人id_In, 主页id_In, Null, 4, Null, 诊断类型_In, 疾病id_In, 诊断id_In, 证候id_In, 诊断描述_In, 入院病情_In, 出院情况_In, 是否未治_In,
     是否疑诊_In, Sysdate, V_记录人, Null, Null, 编码序号_In, 诊断次序_In, 备注_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, V_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl3_病人诊断记录_Insert;
/
Create Or Replace Procedure Zl_病人费用_Delete
(
  病人id_In 病人费用.病人id%Type,
  主页id_In 病人费用.主页id%Type
) Is
Begin
  Delete From 病人费用 Where 病人id = 病人id_In And 主页id = 主页id_In;
End Zl_病人费用_Delete;
/
Create Or Replace Procedure Zl_分娩信息_Delete
(
  病人id_In 病案主页.病人id%Type,
  主页id_In 病案主页.主页id%Type,
  N_类型    Number := 0
  --N_类型:0 Null,删除新生儿信息，由于新生儿与新生儿诊断存在父子关系，所以删除新生儿信息就可以删除新生儿诊断信息
  --    1,只删除新生儿诊断信息
) As

Begin
  If Nvl(N_类型, 0) = 0 Then
    Delete 病人分娩信息 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Delete 新生儿诊断记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_分娩信息_Delete;
/
Create Or Replace Procedure Zl_器械导管使用情况_Update
(
  病人id_In       In 器械导管使用情况.病人id%Type,
  主页id_In       In 器械导管使用情况.主页id%Type,
  序号_In         In 器械导管使用情况.序号%Type,
  监护室名称_In   In 器械导管使用情况.监护室名称%Type,
  器械及导管_In   In 器械导管使用情况.器械及导管%Type,
  开始使用时间_In In 器械导管使用情况.开始使用时间%Type,
  结束使用时间_In In 器械导管使用情况.结束使用时间%Type,
  感染累计时间_In In 器械导管使用情况.感染累计时间%Type := Null
) Is
Begin
  Update 器械导管使用情况
  Set 监护室名称 = 监护室名称_In, 器械及导管 = 器械及导管_In, 开始使用时间 = 开始使用时间_In, 结束使用时间 = 结束使用时间_In, 感染累计时间 = 感染累计时间_In
  Where 病人id = 病人id_In And 序号 = 序号_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_器械导管使用情况_Update;
/

Create Or Replace Procedure Zl_器械导管使用情况_Delete
(
  病人id_In In 器械导管使用情况.病人id%Type,
  主页id_In In 器械导管使用情况.主页id%Type,
  序号s_In  In Varchar2 := Null
) Is
Begin
  If 序号s_In Is Null Then
    Delete From 器械导管使用情况 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Delete From 器械导管使用情况
    Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 In (Select Column_Value From Table(F_Str2list(序号s_In)));
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_器械导管使用情况_Delete;
/

Create Or Replace Procedure Zl_病人感染记录_Delete
(
  病人id_In In 病人感染记录.病人id%Type,
  主页id_In In 病人感染记录.主页id%Type,
  序号s_In  In Varchar2 := Null
) Is
Begin
  If 序号s_In Is Null Then
    Delete From 病人感染记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Delete From 病人感染记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 In (Select Column_Value From Table(F_Str2list(序号s_In)));
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人感染记录_Delete;
/

Create Or Replace Procedure Zl_病人感染记录_Update
(
  序号_In     In 病人感染记录.序号%Type,
  病人id_In   In 病人感染记录.病人id%Type,
  主页id_In   In 病人感染记录.主页id%Type,
  确诊日期_In In 病人感染记录.确诊日期%Type,
  感染部位_In In 病人感染记录.感染部位%Type,
  感染名称_In In 病人感染记录.感染名称%Type
) Is
Begin
  Update 病人感染记录
  Set 登记时间 = Sysdate, 登记人 = Zl_Username, 确诊日期 = 确诊日期_In, 感染部位 = 感染部位_In, 感染名称 = 感染名称_In
  Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 = 序号_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人感染记录_Update;
/

Create Or Replace Procedure Zl_病人病原学检查_Delete
(
  病人id_In In 病人病原学检查.病人id%Type,
  主页id_In In 病人病原学检查.主页id%Type,
  序号s_In  In Varchar2 := Null
) Is
Begin
  If 序号s_In Is Null Then
    Delete From 病人病原学检查 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Delete From 病人病原学检查
    Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 In (Select Column_Value From Table(F_Str2list(序号s_In)));
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人病原学检查_Delete;
/

Create Or Replace Procedure Zl_病人病原学检查_Update
(
  病人id_In     In 病人病原学检查.病人id%Type,
  主页id_In     In 病人病原学检查.主页id%Type,
  序号_In       In 病人病原学检查.序号%Type,
  标本_In       In 病人病原学检查.标本%Type,
  病原学代码_In In 病人病原学检查.病原学代码%Type,
  送检日期_In   In 病人病原学检查.送检日期%Type
) Is
Begin
  Update 病人病原学检查
  Set 登记时间 = Sysdate, 登记人 = Zl_Username, 标本 = 标本_In, 病原学代码 = 病原学代码_In, 送检日期 = 送检日期_In
  Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 = 序号_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人病原学检查_Update;
/

Create Or Replace Procedure Zl_病人感染记录_Insert
(
  病人id_In   In 病人感染记录.病人id%Type,
  主页id_In   In 病人感染记录.主页id%Type,
  序号_In     In 病人感染记录.序号%Type,
  确诊日期_In In 病人感染记录.确诊日期%Type,
  感染部位_In In 病人感染记录.感染部位%Type,
  感染名称_In In 病人感染记录.感染名称%Type
) Is
Begin
  Insert Into 病人感染记录
    (序号, 病人id, 主页id, 登记时间, 登记人, 确诊日期, 感染部位, 感染名称)
  Values
    (序号_In, 病人id_In, 主页id_In, Sysdate, Zl_Username, 确诊日期_In, 感染部位_In, 感染名称_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人感染记录_Insert;
/

Create Or Replace Procedure Zl_病人病原学检查_Insert
(
  病人id_In     In 病人病原学检查.病人id%Type,
  主页id_In     In 病人病原学检查.主页id%Type,
  序号_In       In 病人病原学检查.序号%Type,
  标本_In       In 病人病原学检查.标本%Type,
  病原学代码_In In 病人病原学检查.病原学代码%Type,
  送检日期_In   In 病人病原学检查.送检日期%Type
) Is
Begin
  Insert Into 病人病原学检查
    (序号, 病人id, 主页id, 登记时间, 登记人, 标本, 病原学代码, 送检日期)
  Values
    (序号_In, 病人id_In, 主页id_In, Sysdate, Zl_Username, 标本_In, 病原学代码_In, 送检日期_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人病原学检查_Insert;
/
Create Or Replace Procedure Zl_病案重症监护情况_Update
(
  病人id_In         In 病案重症监护情况.病人id%Type,
  主页id_In         In 病案重症监护情况.主页id%Type,
  序号_In           In 病案重症监护情况.序号%Type,
  监护室名称_In     In 病案重症监护情况.监护室名称%Type,
  进入时间_In       In 病案重症监护情况.进入时间%Type := Null,
  退出时间_In       In 病案重症监护情况.退出时间%Type := Null,
  再入住计划_In     In 病案重症监护情况.再入住计划%Type := Null,
  再入住原因_In     In 病案重症监护情况.再入住原因%Type := Null,
  人工气道脱出_In   In 病案重症监护情况.人工气道脱出%Type := Null,
  重返重症医学科_In In 病案重症监护情况.重返重症医学科%Type := Null,
  重返间隔时间_In   In 病案重症监护情况.重返间隔时间%Type := Null
) Is
Begin
  Update 病案重症监护情况
  Set 监护室名称 = 监护室名称_In, 进入时间 = 进入时间_In, 退出时间 = 退出时间_In, 再入住计划 = 再入住计划_In, 再入住原因 = 再入住原因_In, 人工气道脱出 = 人工气道脱出_In,
      重返重症医学科 = 重返重症医学科_In, 重返间隔时间 = 重返间隔时间_In
  Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 = 序号_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案重症监护情况_Update;
/
Create Or Replace Procedure Zl_病人抗生素记录_Update
(
  功能_In           Number,
  病人id_In         病人抗生素记录.病人id%Type,
  主页id_In         病人抗生素记录.主页id%Type,
  药名id_In         病人抗生素记录.药名id%Type,
  药品名称_In       病人抗生素记录.药品名称%Type,
  用药目的_In       病人抗生素记录.用药目的%Type,
  使用阶段_In       病人抗生素记录.使用阶段%Type,
  使用天数_In       病人抗生素记录.使用天数%Type,
  记录人_In         病人抗生素记录.记录人%Type,
  记录时间_In       病人抗生素记录.记录时间%Type,
  一类切口预防用_In 病人抗生素记录.一类切口预防用%Type := Null,
  Ddd数_In          病人抗生素记录.Ddd数%Type := Null,
  联合用药_In       病人抗生素记录.联合用药%Type := Null
) Is
  --功能：首页整理中抗生素的使用记录 
  --参数：功能_In 0-新增,1-修改   2-删除 
Begin
  If 功能_In = 0 Then
    If 药名id_In Is Null Or 药名id_In = 0 Then
      Update 病人抗生素记录
      Set 用药目的 = 用药目的_In, 使用阶段 = 使用阶段_In, 使用天数 = 使用天数_In, 记录人 = 记录人_In, 记录时间 = 记录时间_In, 一类切口预防用 = 一类切口预防用_In,
          Ddd数 = Ddd数_In, 联合用药 = 联合用药_In
      Where 病人id = 病人id_In And 主页id = 主页id_In And 药品名称 = 药品名称_In;
    Else
      Update 病人抗生素记录
      Set 用药目的 = 用药目的_In, 使用阶段 = 使用阶段_In, 使用天数 = 使用天数_In, 记录人 = 记录人_In, 记录时间 = 记录时间_In, 药品名称 = 药品名称_In,
          一类切口预防用 = 一类切口预防用_In, Ddd数 = Ddd数_In, 联合用药 = 联合用药_In
      Where 病人id = 病人id_In And 主页id = 主页id_In And 药名id = 药名id_In;
    End If;
  Elsif 功能_In = 1 Then
    Insert Into 病人抗生素记录
      (病人id, 主页id, 药名id, 药品名称, 用药目的, 使用阶段, 使用天数, 记录人, 记录时间, 一类切口预防用, Ddd数, 联合用药)
    Values
      (病人id_In, 主页id_In, 药名id_In, 药品名称_In, 用药目的_In, 使用阶段_In, 使用天数_In, 记录人_In, 记录时间_In, 一类切口预防用_In, Ddd数_In, 联合用药_In);
  Elsif 功能_In = 2 Then
    If 药名id_In Is Null Or 药名id_In = 0 Then
      Delete From 病人抗生素记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 药品名称 = 药品名称_In;
    Else
      Delete From 病人抗生素记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 药名id = 药名id_In;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人抗生素记录_Update;
/
Create Or Replace Procedure Zl_病人手麻记录_Delete
(
  病人id_In   病人手麻记录.病人id%Type,
  主页id_In   病人手麻记录.主页id%Type,
  记录来源_In 病人手麻记录.记录来源%Type,
  数据来源_In 病人手麻记录.数据来源%Type := Null,
  手麻ids_In  Varchar2 := Null
) Is
  --功能：删除病人手麻记录
Begin
  ---差异，为了提交VSS
  If 数据来源_In Is Null Then
    If Not 手麻ids_In Is Null Then
      Delete From 病人手麻记录
      Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And 记录来源 = 记录来源_In And
            ID In (Select Column_Value From Table(F_Str2list(手麻ids_In)));
    Else
      Delete From 病人手麻记录
      Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And 记录来源 = 记录来源_In;
    End If;
  Else
    Delete From 病人手麻记录
    Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0) And 记录来源 = 记录来源_In And 数据来源 = 数据来源_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人手麻记录_Delete;
/
Create Or Replace Procedure Zl_病人手麻记录_Update
(
  Id_In               病人手麻记录.Id%Type,
  病人id_In           病人手麻记录.病人id%Type,
  主页id_In           病人手麻记录.主页id%Type,
  记录来源_In         病人手麻记录.记录来源%Type,
  手术日期_In         病人手麻记录.手术日期%Type,
  手术开始时间_In     病人手麻记录.手术开始时间%Type,
  手术结束时间_In     病人手麻记录.手术结束时间%Type,
  拟行手术_In         病人手麻记录.拟行手术%Type,
  手术操作id_In       病人手麻记录.手术操作id%Type,
  诊疗项目id_In       病人手麻记录.诊疗项目id%Type,
  已行手术_In         病人手麻记录.已行手术%Type,
  主刀医师_In         病人手麻记录.主刀医师%Type,
  助产护士_In         病人手麻记录.助产护士%Type,
  第一助手_In         病人手麻记录.第一助手%Type,
  第二助手_In         病人手麻记录.第二助手%Type,
  手术护士_In         病人手麻记录.手术护士%Type,
  麻醉开始时间_In     病人手麻记录.麻醉开始时间%Type,
  麻醉结束时间_In     病人手麻记录.麻醉结束时间%Type,
  麻醉方式_In         病人手麻记录.麻醉方式%Type,
  麻醉类型_In         病人手麻记录.麻醉类型%Type,
  麻醉质量_In         病人手麻记录.麻醉质量%Type,
  输液总量_In         病人手麻记录.输液总量%Type,
  麻醉医师_In         病人手麻记录.麻醉医师%Type,
  输氧开始时间_In     病人手麻记录.输氧开始时间%Type,
  输氧结束时间_In     病人手麻记录.输氧结束时间%Type,
  切口_In             病人手麻记录.切口%Type,
  愈合_In             病人手麻记录.愈合%Type,
  手术情况_In         病人手麻记录.手术情况%Type := Null,
  Asa分级_In          病人手麻记录.Asa分级%Type := Null,
  再次手术_In         病人手麻记录.再次手术%Type := Null,
  Nnis分级_In         病人手麻记录.Nnis分级%Type := Null,
  手术级别_In         病人手麻记录.手术级别%Type := Null,
  准备天数_In         病人手麻记录.准备天数%Type := Null,
  抗菌用药时间_In     病人手麻记录.抗菌用药时间%Type := Null,
  切口部位_In         病人手麻记录.切口部位%Type := Null,
  切口感染_In         病人手麻记录.切口感染%Type := Null,
  并发症_In           病人手麻记录.并发症%Type := Null,
  重返计划_In         病人手麻记录.重返计划%Type := Null,
  重返目的_In         病人手麻记录.重返目的%Type := Null,
  术前抗菌用药_In     病人手麻记录.术前抗菌用药%Type := Null,
  抗菌用药天数_In     病人手麻记录.抗菌用药天数%Type := Null,
  非预期的二次手术_In 病人手麻记录.非预期的二次手术%Type := Null,
  麻醉并发症_In       病人手麻记录.麻醉并发症%Type := Null,
  术中异物遗留_In     病人手麻记录.术中异物遗留%Type := Null,
  手术并发症_In       病人手麻记录.手术并发症%Type := Null,
  术后出血或血肿_In   病人手麻记录.术后出血或血肿%Type := Null,
  手术伤口裂开_In     病人手麻记录.手术伤口裂开%Type := Null,
  术后深静脉血栓_In   病人手麻记录.术后深静脉血栓%Type := Null,
  术后生理代谢紊乱_In 病人手麻记录.术后生理代谢紊乱%Type := Null,
  术后呼吸衰竭_In     病人手麻记录.术后呼吸衰竭%Type := Null,
  术后肺栓塞_In       病人手麻记录.术后肺栓塞%Type := Null,
  术后败血症_In       病人手麻记录.术后败血症%Type := Null,
  术后髋关节骨折_In   病人手麻记录.术后髋关节骨折%Type := Null,
  数据来源_In         病人手麻记录.数据来源%Type := Null
) Is
  --功能：插入病人手麻记录 
  N_Count    Number(8);
  V_Temp     Varchar2(255);
  V_人员编号 人员表.编号%Type;
  V_人员姓名 人员表.姓名%Type;
Begin
  --先判断当前手术是指定的来源的手术，是则更新，否则新增
  Select Count(1) Into N_Count From 病人手麻记录 Where ID = Id_In And 记录来源 = 记录来源_In;
  If N_Count = 1 Then
    Update 病人手麻记录
    Set 手术日期 = 手术日期_In, 手术开始时间 = 手术开始时间_In, 手术结束时间 = 手术结束时间_In, 拟行手术 = 拟行手术_In, 手术操作id = 手术操作id_In, 诊疗项目id = 诊疗项目id_In,
        已行手术 = 已行手术_In, 主刀医师 = 主刀医师_In, 助产护士 = 助产护士_In, 第一助手 = 第一助手_In, 第二助手 = 第二助手_In, 手术护士 = 手术护士_In,
        麻醉开始时间 = 麻醉开始时间_In, 麻醉结束时间 = 麻醉开始时间_In, 麻醉方式 = 麻醉方式_In, 麻醉类型 = 麻醉类型_In, 麻醉质量 = 麻醉质量_In, 输液总量 = 输液总量_In,
        麻醉医师 = 麻醉医师_In, 输氧开始时间 = 输氧开始时间_In, 输氧结束时间 = 输氧结束时间_In, 切口 = 切口_In, 愈合 = 愈合_In, 手术情况 = 手术情况_In, Asa分级 = Asa分级_In,
        再次手术 = 再次手术_In, Nnis分级 = Nnis分级_In, 手术级别 = 手术级别_In, 准备天数 = 准备天数_In, 抗菌用药时间 = 抗菌用药时间_In, 切口部位 = 切口部位_In,
        切口感染 = 切口感染_In, 并发症 = 并发症_In, 重返计划 = 重返计划_In, 重返目的 = 重返目的_In, 术前抗菌用药 = 术前抗菌用药_In, 抗菌用药天数 = 抗菌用药天数_In,
        非预期的二次手术 = 非预期的二次手术_In, 麻醉并发症 = 麻醉并发症_In, 术中异物遗留 = 术中异物遗留_In, 手术并发症 = 手术并发症_In, 术后出血或血肿 = 术后出血或血肿_In,
        手术伤口裂开 = 手术伤口裂开_In, 术后深静脉血栓 = 术后深静脉血栓_In, 术后生理代谢紊乱 = 术后生理代谢紊乱_In, 术后呼吸衰竭 = 术后呼吸衰竭_In, 术后肺栓塞 = 术后肺栓塞_In,
        术后败血症 = 术后败血症_In, 术后髋关节骨折 = 术后髋关节骨折_In, 数据来源 = 数据来源_In
    Where ID = Id_In;
  Else
    --当前操作人员 
    V_Temp     := Zl_Identity;
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ';') + 1);
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ',') + 1);
    V_人员编号 := Substr(V_Temp, 1, Instr(V_Temp, ',') - 1);
    V_人员姓名 := Substr(V_Temp, Instr(V_Temp, ',') + 1);
  
    Insert Into 病人手麻记录
      (ID, 病人id, 主页id, 记录来源, 手术日期, 手术开始时间, 手术结束时间, 拟行手术, 手术操作id, 诊疗项目id, 已行手术, 主刀医师, 助产护士, 第一助手, 第二助手, 手术护士, 麻醉开始时间,
       麻醉结束时间, 麻醉方式, 麻醉类型, 麻醉质量, 输液总量, 麻醉医师, 输氧开始时间, 输氧结束时间, 切口, 愈合, 记录日期, 记录人, 手术情况, Asa分级, 再次手术, Nnis分级, 手术级别, 准备天数,
       抗菌用药时间, 切口部位, 切口感染, 并发症, 重返计划, 重返目的, 术前抗菌用药, 抗菌用药天数, 非预期的二次手术, 麻醉并发症, 术中异物遗留, 手术并发症, 术后出血或血肿, 手术伤口裂开, 术后深静脉血栓,
       术后生理代谢紊乱, 术后呼吸衰竭, 术后肺栓塞, 术后败血症, 术后髋关节骨折, 数据来源)
      Select 病人手麻记录_Id.Nextval, 病人id_In, 主页id_In, 记录来源_In, 手术日期_In, 手术开始时间_In, 手术结束时间_In, 拟行手术_In, 手术操作id_In, 诊疗项目id_In,
             已行手术_In, 主刀医师_In, 助产护士_In, 第一助手_In, 第二助手_In, 手术护士_In, 麻醉开始时间_In, 麻醉结束时间_In, 麻醉方式_In, 麻醉类型_In, 麻醉质量_In,
             输液总量_In, 麻醉医师_In, 输氧开始时间_In, 输氧结束时间_In, 切口_In, 愈合_In, Sysdate, V_人员姓名, 手术情况_In, Asa分级_In, 再次手术_In,
             Nnis分级_In, 手术级别_In, 准备天数_In, 抗菌用药时间_In, 切口部位_In, 切口感染_In, 并发症_In, 重返计划_In, 重返目的_In, 术前抗菌用药_In, 抗菌用药天数_In,
             非预期的二次手术_In, 麻醉并发症_In, 术中异物遗留_In, 手术并发症_In, 术后出血或血肿_In, 手术伤口裂开_In, 术后深静脉血栓_In, 术后生理代谢紊乱_In, 术后呼吸衰竭_In,
             术后肺栓塞_In, 术后败血症_In, 术后髋关节骨折_In, 数据来源_In
      From Dual;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人手麻记录_Update;
/
Create Or Replace Procedure Zl_病案精神治疗_Delete
(
  病人id_In In 病案精神治疗.病人id%Type,
  主页id_In In 病案精神治疗.主页id%Type,
  序号s_In  In Varchar2 := Null
) Is
Begin
  If 序号s_In Is Null Then
    Delete From 病案精神治疗 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Delete From 病案精神治疗
    Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 In (Select Column_Value From Table(F_Str2list(序号s_In)));
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案精神治疗_Delete;
/
Create Or Replace Procedure Zl_病案精神治疗_Update
(
  病人id_In   In 病案精神治疗.病人id%Type,
  主页id_In   In 病案精神治疗.主页id%Type,
  序号_In     In 病案精神治疗.序号%Type,
  药品id_In   In 病案精神治疗.药品id%Type,
  药物名称_In In 病案精神治疗.药物名称%Type,
  疗程_In     In 病案精神治疗.疗程%Type,
  最高日量_In In 病案精神治疗.最高日量%Type,
  特殊反应_In In 病案精神治疗.特殊反应%Type,
  疗效_In     In 病案精神治疗.疗效%Type
) Is
Begin
  
  Update 病案精神治疗
  Set 药品id = 药品id_In, 药物名称 = 药物名称_In, 疗程 = 疗程_In, 最高日量 = 最高日量_In, 特殊反应 = 特殊反应_In, 疗效 = 疗效_In
  Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 = 序号_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案精神治疗_Update;
/
Create Or Replace Procedure Zl_病案化疗记录_Delete
(
  病人id_In In 病案化疗记录.病人id%Type,
  主页id_In In 病案化疗记录.主页id%Type,
  序号s_In  In Varchar2 := Null
) Is
Begin
  If 序号s_In Is Null Then
    Delete From 病案化疗记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Delete From 病案化疗记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 In (Select Column_Value From Table(F_Str2list(序号s_In)));
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案化疗记录_Delete;
/
Create Or Replace Procedure Zl_病案放疗记录_Delete
(
  病人id_In In 病案放疗记录.病人id%Type,
  主页id_In In 病案放疗记录.主页id%Type,
  序号s_In  In Varchar2 := Null
) Is
Begin
  If 序号s_In Is Null Then
    Delete From 病案放疗记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Delete From 病案放疗记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 In (Select Column_Value From Table(F_Str2list(序号s_In)));
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案放疗记录_Delete;
/
Create Or Replace Procedure Zl_病案化疗记录_Update
(
  病人id_In   In 病案化疗记录.病人id%Type,
  主页id_In   In 病案化疗记录.主页id%Type,
  序号_In     In 病案化疗记录.序号%Type,
  疾病id_In   In 病案化疗记录.疾病id%Type,
  开始日期_In In 病案化疗记录.开始日期%Type,
  结束日期_In In 病案化疗记录.结束日期%Type,
  化疗方案_In In 病案化疗记录.化疗方案%Type,
  疗程数_In   In 病案化疗记录.疗程数%Type,
  总量_In     In 病案化疗记录.总量%Type,
  化疗效果_In In 病案化疗记录.化疗效果%Type
) Is
Begin
  Update 病案化疗记录
  Set 疾病id = 疾病id_In, 开始日期 = 开始日期_In, 结束日期 = 结束日期_In, 疗程数 = 疗程数_In, 总量 = 总量_In, 化疗方案 = 化疗方案_In, 化疗效果 = 化疗效果_In
  Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 = 序号_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案化疗记录_Update;
/
Create Or Replace Procedure Zl_病案放疗记录_Update
(
  病人id_In   In 病案放疗记录.病人id%Type,
  主页id_In   In 病案放疗记录.主页id%Type,
  序号_In     In 病案放疗记录.序号%Type,
  疾病id_In   In 病案放疗记录.疾病id%Type,
  开始日期_In In 病案放疗记录.开始日期%Type,
  结束日期_In In 病案放疗记录.结束日期%Type,
  设野部位_In In 病案放疗记录.设野部位%Type,
  放射剂量_In In 病案放疗记录.放射剂量%Type,
  累计量_In   In 病案放疗记录.累计量%Type,
  放疗效果_In In 病案放疗记录.放疗效果%Type
) Is
Begin

  Update 病案放疗记录
  Set 疾病id = 疾病id_In, 开始日期 = 开始日期_In, 结束日期 = 结束日期_In, 设野部位 = 设野部位_In, 放射剂量 = 放射剂量_In, 累计量 = 累计量_In, 放疗效果 = 放疗效果_In
  Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 = 序号_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案放疗记录_Update;
/
Create Or Replace Procedure Zl_病案重症监护情况_Delete
(
  病人id_In In 病案重症监护情况.病人id%Type,
  主页id_In In 病案重症监护情况.主页id%Type,
  序号s_In  In Varchar2 := Null
) Is
Begin
  If 序号s_In Is Null Then
    Delete From 病案重症监护情况 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Else
    Delete From 病案重症监护情况
    Where 病人id = 病人id_In And 主页id = 主页id_In And 序号 In (Select Column_Value From Table(F_Str2list(序号s_In)));
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案重症监护情况_Delete;
/