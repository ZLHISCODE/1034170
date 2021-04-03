Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( &n_System,'病原学目录',0,'标本对应的病原学代码已经名称','医疗工作' ); 
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( &n_System,'医院感染目录',0,'医院感染的标准分类名称','医疗工作' ); 
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( &n_System,'器械导管目录',0,'重症监护常用器械目录','医疗工作' ); 
Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values( &n_System,'Icu类型',0,'重症监护的类别目录','医疗工作' ); 
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1261,'首页整理',User,'Zl_病人手麻记录_Update','EXECUTE' From Dual Union All 
Select &n_System,1261,'首页整理',User,'Zl_病人诊断记录_Update','EXECUTE' From Dual Union All
Select &n_System,1261,'首页整理',User,'Zl_病人过敏记录_Update','EXECUTE' From Dual Union All
Select &n_System,1261,'首页整理',User,'Zl_病人感染记录_Update','EXECUTE' From Dual Union All 
Select &n_System,1261,'首页整理',User,'Zl_病人病原学检查_Update','EXECUTE' From Dual Union All
Select &n_System,1261,'首页整理',User,'Zl_器械导管使用情况_Update','EXECUTE' From Dual Union All
Select &n_System,1261,'首页整理',User,'Zl_病案重症监护情况_Update','EXECUTE' From Dual Union All
Select &n_System,1261,'首页整理',User,'Zl_病案化疗记录_Update','EXECUTE' From Dual Union All 
Select &n_System,1261,'首页整理',User,'Zl_病案放疗记录_Update','EXECUTE' From Dual Union All
Select &n_System,1261,'首页整理',User,'Zl_病案精神治疗_Update','EXECUTE' From Dual

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1260,'门诊首页',User,'Zl_病人诊断记录_Update','EXECUTE' From Dual Union All 
Select &n_System,1260,'门诊首页',User,'Zl_病人过敏记录_Update','EXECUTE' From Dual 

Insert Into zlBaseCode(系统,表名,固定,说明,分类) Values(&n_System,'住院死亡原因',0,'病人的住院死亡原因','医疗工作');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1261,'首页整理',User,'住院死亡原因','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1262,'首页整理',User,'住院死亡原因','SELECT');
--住院死亡原因:病情危重,入院过迟,医疗事故,其它
insert into 住院死亡原因(编码,名称,简码,缺省标志) values ('1','病情危重','BQWZ',0);
insert into 住院死亡原因(编码,名称,简码,缺省标志) values ('2','入院过迟','RYGC',0);
insert into 住院死亡原因(编码,名称,简码,缺省标志) values ('3','医疗事故','YLSG',1);
insert into 住院死亡原因(编码,名称,简码,缺省标志) values ('4','其它','QT',0);
Create Table 住院死亡原因(
    编码 VARCHAR2(2),
    名称 VARCHAR2(50),
    简码 VARCHAR2(10),
    缺省标志 NUMBER(1) default 0)
    TABLESPACE zl9BaseItem ;

Alter Table 住院死亡原因 Add Constraint 住院死亡原因_PK Primary Key (编码) Using Index  Tablespace  zl9Indexhis;
Alter Table 住院死亡原因 Add Constraint 住院死亡原因_UQ_名称 Unique (名称) Using Index  Tablespace  zl9Indexhis;
Alter Table 住院死亡期间 Add Constraint 住院死亡期间_PK Primary Key (编码) Using Index  Tablespace  zl9Indexhis;
Alter Table 住院死亡期间 Add Constraint 住院死亡期间_UQ_名称 Unique (名称) Using Index  Tablespace  zl9Indexhis;

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

Create Or Replace Procedure Zl_病人诊断记录_Update
(
  Id_In       病人诊断记录.Id%Type,
  病人id_In   病人诊断记录.病人id%Type,
  主页id_In   病人诊断记录.主页id%Type,
  记录来源_In 病人诊断记录.记录来源%Type,
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
  发病时间_In 病人诊断记录.发病时间%Type := Null
) Is
  --功能：更新病人诊断记录 
  V_诊断id     病人诊断记录.Id%Type;
  D_记录日期   病人诊断记录.记录日期%Type;
  V_病人科室id 病人信息.当前科室id%Type;
  V_经治医师   人员表.姓名%Type;
  V_编码       疾病编码目录.编码%Type;
  N_Count      Number;
  V_发病时间   病人诊断记录.发病时间%Type;

  V_Temp     Varchar2(255);
  V_人员姓名 人员表.姓名%Type;
Begin
  --先判断当前诊断是指定的诊断来源的诊断，是则更新，否则新增
  Select Count(1) Into N_Count From 病人诊断记录 Where ID = Id_In And (记录来源 = 记录来源_In Or 记录来源_In Is Null);
  If N_Count = 1 Then
    Update 病人诊断记录
    Set 记录来源 = 记录来源_In, 诊断次序 = 诊断次序_In, 疾病id = 疾病id_In, 诊断id = 诊断id_In, 证候id = 证候id_In, 诊断描述 = 诊断描述_In, 入院病情 = 入院病情_In,
        出院情况 = 出院情况_In, 是否未治 = 是否未治_In, 是否疑诊 = 是否疑诊_In, 备注 = 备注_In, 发病时间 = 发病时间_In
    Where ID = Id_In
    Returning 记录日期 Into D_记录日期;
  Else
    Select 病人诊断记录_Id.Nextval Into V_诊断id From Dual;
  
    V_Temp     := Zl_Identity;
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ';') + 1);
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ',') + 1);
    V_人员姓名 := Substr(V_Temp, Instr(V_Temp, ',') + 1);
    D_记录日期 := Sysdate;
    Insert Into 病人诊断记录
      (ID, 病人id, 主页id, 记录来源, 诊断类型, 诊断次序, 疾病id, 诊断id, 证候id, 诊断描述, 入院病情, 出院情况, 是否未治, 是否疑诊, 记录日期, 记录人, 备注, 发病时间)
      Select V_诊断id, 病人id, 主页id, 记录来源_In, 诊断类型_In, 诊断次序_In, 疾病id_In, 诊断id_In, 证候id_In, 诊断描述_In, 入院病情_In, 出院情况_In,
             是否未治_In, 是否疑诊_In, Sysdate, V_人员姓名, 备注_In, 发病时间_In
      From 病人诊断记录
      Where ID = Id_In;
  End If;

  --如果是门诊第一诊断则更新病人挂号记录.发病时间 
  V_发病时间 := 发病时间_In;
  If 诊断类型_In = 1 And 诊断次序_In = 1 Then
    If 发病时间_In Is Null Then
      --检查中医的发病时间，有则取中医的，否则清空 
      Select Max(发病时间)
      Into V_发病时间
      From 病人诊断记录
      Where 病人id = 病人id_In And 主页id = 主页id_In And 诊断类型 = 11 And 诊断次序 = 1;
    End If;
    If V_发病时间 Is Null Then
      --如果都为NULL，则取挂号记录中的 
      Select Max(发病时间) Into V_发病时间 From 病人挂号记录 Where 病人id = 病人id_In And ID = 主页id_In;
    End If;
    Update 病人挂号记录 Set 发病时间 = V_发病时间 Where 病人id = 病人id_In And ID = 主页id_In;
  End If;
  If 诊断类型_In = 11 And 诊断次序_In = 1 Then
    --如果是中医，则判断是否填写了西医的发病时间，没有填写，则修改，否则以西医发病时间为准 
    Select Count(*)
    Into N_Count
    From 病人诊断记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 诊断类型 = 1 And 诊断次序 = 1 And 发病时间 Is Not Null;
    If N_Count = 0 Then
      If V_发病时间 Is Null Then
        --如果都为NULL，则取挂号记录中的 
        Select Max(发病时间) Into V_发病时间 From 病人挂号记录 Where 病人id = 病人id_In And ID = 主页id_In;
      End If;
      Update 病人挂号记录 Set 发病时间 = V_发病时间 Where 病人id = 病人id_In And ID = 主页id_In;
    End If;
  End If;

  --如果是入院第一诊断，则判断是否是单病种 
  If 诊断类型_In = 2 And 诊断次序_In = 1 And 记录来源_In = 3 Then
    If 疾病id_In Is Not Null Then
      Select 编码 Into V_编码 From 疾病编码目录 Where ID = 疾病id_In;
      Select Max(Upper(编码))
      Into V_编码
      From 单病种目录
      Where Instr('/' || Replace(Upper(Icd编码), ' ', '') || '/', '/' || Upper(V_编码) || '/') > 0 And Rownum < 2;
    Else
      V_编码 := '';
    End If;
    Update 病案主页 Set 单病种 = V_编码 Where 病人id = 病人id_In And 主页id = 主页id_In;
  End If;

  --根据传入的主页id_In查询挂号记录来区分是门诊首页还是住院首页调用 
  Begin
    Select 执行人, 执行部门id Into V_病人科室id, V_经治医师 From 病人挂号记录 Where ID = 主页id_In;
    Zl_电子病历时机_Insert(病人id_In, 主页id_In, 1, '诊断', V_病人科室id, V_经治医师, D_记录日期, D_记录日期);
  Exception
    When Others Then
      Null;
  End;
  If V_病人科室id Is Null And (诊断类型_In <> 1 Or 诊断类型_In <> 11) Then
    Begin
      Select 出院科室id, 住院医师
      Into V_病人科室id, V_经治医师
      From 病案主页
      Where 病人id = 病人id_In And 主页id = 主页id_In;
      Zl_电子病历时机_Insert(病人id_In, 主页id_In, 2, '诊断', V_病人科室id, V_经治医师, D_记录日期, D_记录日期);
    Exception
      When Others Then
        Null;
    End;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人诊断记录_Update;
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
Create Or Replace Procedure Zl_病案化疗记录_Insert
(
  病人id_In   In 病案化疗记录.病人id%Type,
  主页id_In   In 病案化疗记录.主页id%Type,
  序号_In     In 病案化疗记录.序号%Type,
  疾病id_In   In 病案化疗记录.疾病id%Type,
  开始日期_In In 病案化疗记录.开始日期%Type,
  结束日期_In In 病案化疗记录.结束日期%Type,
  疗程数_In   In 病案化疗记录.疗程数%Type,
  化疗方案_In In 病案化疗记录.化疗方案%Type,
  总量_In     In 病案化疗记录.总量%Type,
  化疗效果_In In 病案化疗记录.化疗效果%Type
) Is
Begin
  Insert Into 病案化疗记录
    (病人id, 主页id, 序号, 疾病id, 开始日期, 结束日期, 疗程数, 总量, 化疗方案, 化疗效果)
  Values
    (病人id_In, 主页id_In, 序号_In, 疾病id_In, 开始日期_In, 结束日期_In, 疗程数_In, 总量_In, 化疗方案_In, 化疗效果_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病案化疗记录_Insert;
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

--重整参数顺序
Create Or Replace Procedure Zl_病案主页_首页整理
(
  病人id_In       病案主页.病人id%Type,
  主页id_In       病案主页.主页id%Type,
  年龄_In         病案主页.年龄%Type,
  国籍_In         病案主页.国籍%Type,
  区域_In         病案主页.区域%Type,
  职业_In         病案主页.职业%Type,
  身高_In         病案主页.身高%Type,
  体重_In         病案主页.体重%Type,
  血型_In         病案主页.血型%Type,
  婚姻状况_In     病案主页.婚姻状况%Type,
  医疗付款方式_In 病案主页.医疗付款方式%Type,
  家庭地址_In     病案主页.家庭地址%Type,
  家庭电话_In     病案主页.家庭电话%Type,
  家庭邮编_In     病案主页.家庭地址邮编%Type,
  户口地址_In     病案主页.户口地址%Type,
  户口邮编_In     病案主页.户口地址邮编%Type,
  单位地址_In     病案主页.单位地址%Type,
  单位电话_In     病案主页.单位电话%Type,
  单位邮编_In     病案主页.单位邮编%Type,
  联系人姓名_In   病案主页.联系人姓名%Type,
  联系人关系_In   病案主页.联系人关系%Type,
  联系人电话_In   病案主页.联系人电话%Type,
  联系人地址_In   病案主页.联系人地址%Type,
  入院病况_In     病案主页.入院病况%Type,
  入院方式_In     病案主页.入院方式%Type,
  出院方式_In     病案主页.出院方式%Type,
  再入院_In       病案主页.再入院%Type,
  是否确诊_In     病案主页.是否确诊%Type,
  确诊日期_In     病案主页.确诊日期%Type,
  尸检标志_In     病案主页.尸检标志%Type,
  随诊标志_In     病案主页.随诊标志%Type,
  随诊期限_In     病案主页.随诊期限%Type,
  新发肿瘤_In     病案主页.新发肿瘤%Type,
  中医治疗类别_In 病案主页.中医治疗类别%Type,
  抢救次数_In     病案主页.抢救次数%Type,
  成功次数_In     病案主页.成功次数%Type,
  门诊医师_In     病案主页.门诊医师%Type,
  住院医师_In     病案主页.住院医师%Type,
  主治医师_In     病案主页.住院医师%Type,
  主任医师_In     病案主页.住院医师%Type,
  责任护士_In     病案主页.责任护士%Type,
  操作员编号_In   病案主页.编目员编号%Type := Null,
  操作员姓名_In   病案主页.编目员姓名%Type := Null
) As
  --功能：用于住院医护工作站对病人进行首页整理 
  V_住院医师 病案主页.住院医师%Type;
  V_主治医师 病案主页.住院医师%Type;
  V_主任医师 病案主页.住院医师%Type;
  V_责任护士 病案主页.责任护士%Type;
  V_病人性质 病案主页.病人性质%Type;
  V_原因     病人变动记录.开始原因%Type;
  V_出院科室 病案主页.出院科室id%Type;
  V_Curdate  Date;
  V_Count    Number;

  Cursor C_Oldinfo Is
    Select B.*
    From (Select C.*
           From 病人变动记录 C
           Where C.病人id = 病人id_In And C.主页id = 主页id_In And
                 C.开始时间 = (Select Min(开始时间)
                           From 病人变动记录
                           Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > V_Curdate)) A, 病人变动记录 B
    Where B.病人id = 病人id_In And B.主页id = 主页id_In And A.开始时间 = B.终止时间 And A.开始原因 = B.终止原因 And A.附加床位 = B.附加床位
    Union
    Select *
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And 开始时间 <= V_Curdate;

  Cursor C_Endinfo Is
    Select * From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
  R_Oldinfo  C_Oldinfo%RowType;
  R_Endinfo  C_Endinfo%RowType;
  V_终止原因 病人变动记录.终止原因%Type;
  V_终止时间 病人变动记录.终止时间%Type;
  V_终止人员 病人变动记录.终止人员%Type;

  V_Error Varchar2(255);
  Err_Custom Exception;

Begin
  --取更改前的内容(用NoneData和新的比较) 
  Select 病人性质, Nvl(住院医师, 'NoneData'), Nvl(出院科室id, 入院科室id), Nvl(责任护士, 'NoneData')
  Into V_病人性质, V_住院医师, V_出院科室, V_责任护士
  From 病案主页
  Where 病人id = 病人id_In And 主页id = 主页id_In;

  Begin
    Select Nvl(信息值, 'NoneData')
    Into V_主治医师
    From 病案主页从表
    Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主治医师';
  
  Exception
    When Others Then
      V_主治医师 := 'NoneData';
    
  End;

  Begin
    Select Nvl(信息值, 'NoneData')
    Into V_主任医师
    From 病案主页从表
    Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主任医师';
  
  Exception
    When Others Then
      V_主任医师 := 'NoneData';
    
  End;

  Update 病案主页
  Set 婚姻状况 = 婚姻状况_In, 年龄 = 年龄_In, 职业 = 职业_In, 国籍 = 国籍_In, 区域 = 区域_In, 医疗付款方式 = 医疗付款方式_In, 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In,
      家庭地址邮编 = 家庭邮编_In, 单位地址 = 单位地址_In, 单位电话 = 单位电话_In, 单位邮编 = 单位邮编_In, 联系人姓名 = 联系人姓名_In, 联系人关系 = 联系人关系_In,
      联系人电话 = 联系人电话_In, 联系人地址 = 联系人地址_In, 再入院 = 再入院_In, 入院病况 = 入院病况_In, 是否确诊 = 是否确诊_In, 确诊日期 = 确诊日期_In, 抢救次数 = 抢救次数_In,
      成功次数 = 成功次数_In, 尸检标志 = 尸检标志_In, 随诊标志 = 随诊标志_In, 随诊期限 = 随诊期限_In, 血型 = 血型_In, 门诊医师 = 门诊医师_In, 住院医师 = 住院医师_In,
      新发肿瘤 = 新发肿瘤_In, 中医治疗类别 = 中医治疗类别_In, 身高 = 身高_In, 体重 = 体重_In, 出院方式 = 出院方式_In, 入院方式 = 入院方式_In, 责任护士 = 责任护士_In,
      户口地址 = 户口地址_In, 户口地址邮编 = 户口邮编_In
  Where 病人id = 病人id_In And 主页id = 主页id_In;

  If V_住院医师 <> Nvl(住院医师_In, 'NoneData') Or V_主治医师 <> Nvl(主治医师_In, 'NoneData') Or V_主任医师 <> Nvl(主任医师_In, 'NoneData') Or
     V_责任护士 <> Nvl(责任护士_In, 'NoneData') Then
  
    Select Sysdate Into V_Curdate From Dual;
    Open C_Oldinfo;
    Fetch C_Oldinfo
      Into R_Oldinfo;
    Open C_Endinfo;
    Fetch C_Endinfo
      Into R_Endinfo;
    If C_Endinfo%RowCount = 0 Then
      --出院病人不进行变动处理 
      Close C_Endinfo;
    Else
      --如果终止时间<>NULL ，就记录下终止时间和终止原因。 
      If R_Oldinfo.终止时间 Is Not Null Then
        V_终止时间 := R_Oldinfo.终止时间;
        V_终止原因 := R_Oldinfo.终止原因;
        V_终止人员 := R_Oldinfo.终止人员;
      End If;
      --如果是待入住的病人，则不产生变动，直接修改已有变动 
      If R_Oldinfo.开始原因 = 1 And R_Oldinfo.终止时间 Is Null Then
        Update 病人变动记录
        Set 经治医师 = 住院医师_In, 主治医师 = 主治医师_In, 责任护士 = 责任护士_In, 主任医师 = 主任医师_In
        Where ID = R_Oldinfo.Id;
      Else
        If V_住院医师 <> Nvl(住院医师_In, 'NoneData') Then
          V_原因 := 7;
          If V_终止时间 Is Null Then
            Update 病人变动记录
            Set 终止时间 = V_Curdate, 终止原因 = V_原因, 终止人员 = 操作员姓名_In
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
          Else
            Update 病人变动记录
            Set 终止时间 = V_Curdate, 终止原因 = V_原因, 终止人员 = 操作员姓名_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = V_终止时间 And 终止原因 = V_终止原因;
            --更新将来的记录如果有停止到将来的则删除上次计算时间 
            Update 病人变动记录
            Set 经治医师 = 住院医师_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > V_Curdate;
          End If;
        
          --产生病历书写时机 
          Zl_电子病历时机_Insert(病人id_In, 主页id_In, 2, '交班', R_Oldinfo.科室id, 住院医师_In, V_Curdate, V_Curdate);
        
          While C_Oldinfo%Found Loop
            --注意:有附加床位时有多条记录 
            Insert Into 病人变动记录
              (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 医疗小组id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情,
               操作员编号, 操作员姓名, 终止时间, 终止原因, 终止人员)
            Values
              (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, V_Curdate, V_原因, R_Oldinfo.附加床位, R_Oldinfo.病区id, R_Oldinfo.科室id,
               R_Oldinfo.医疗小组id, R_Oldinfo.护理等级id, R_Oldinfo.床位等级id, R_Oldinfo.床号, R_Oldinfo.责任护士, 住院医师_In,
               R_Oldinfo.主治医师, R_Oldinfo.主任医师, R_Oldinfo.病情, 操作员编号_In, 操作员姓名_In, V_终止时间, V_终止原因, V_终止人员);
          
            Fetch C_Oldinfo
              Into R_Oldinfo;
          End Loop;
        
          --如果存在停止到将来的变动就更新终止原因 
          If V_终止时间 Is Not Null Then
            V_终止原因 := V_原因;
            V_终止时间 := V_Curdate;
            V_终止人员 := 操作员姓名_In;
          End If;
        
          Close C_Oldinfo;
          Open C_Oldinfo; --重新打开,以便取最新信息 
          Fetch C_Oldinfo
            Into R_Oldinfo;
        End If;
      
        If V_主治医师 <> Nvl(主治医师_In, 'NoneData') Then
          Update 病案主页从表
          Set 信息值 = 主治医师_In
          Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主治医师';
        
          If Sql%RowCount = 0 Then
            Insert Into 病案主页从表
              (病人id, 主页id, 信息名, 信息值)
            Values
              (病人id_In, 主页id_In, '主治医师', 主治医师_In);
          
          End If;
        
          V_原因 := 11;
          If V_终止时间 Is Null Then
            Update 病人变动记录
            Set 终止时间 = V_Curdate, 终止原因 = V_原因, 终止人员 = 操作员姓名_In
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
          
          Else
            Update 病人变动记录
            Set 终止时间 = V_Curdate, 终止原因 = V_原因, 终止人员 = 操作员姓名_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = V_终止时间 And 终止原因 = V_终止原因;
            --更新将来的记录如果有停止到将来的则删除上次计算时间 
            Update 病人变动记录
            Set 主治医师 = 主治医师_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > V_Curdate;
          End If;
        
          While C_Oldinfo%Found Loop
            Insert Into 病人变动记录
              (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 医疗小组id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情,
               操作员编号, 操作员姓名, 终止时间, 终止原因, 终止人员)
            Values
              (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, V_Curdate, V_原因, R_Oldinfo.附加床位, R_Oldinfo.病区id, R_Oldinfo.科室id,
               R_Oldinfo.医疗小组id, R_Oldinfo.护理等级id, R_Oldinfo.床位等级id, R_Oldinfo.床号, R_Oldinfo.责任护士, R_Oldinfo.经治医师,
               主治医师_In, R_Oldinfo.主任医师, R_Oldinfo.病情, 操作员编号_In, 操作员姓名_In, V_终止时间, V_终止原因, V_终止人员);
          
            Fetch C_Oldinfo
              Into R_Oldinfo;
          End Loop;
        
          --如果存在停止到将来的变动就更新终止原因 
          If V_终止时间 Is Not Null Then
            V_终止原因 := V_原因;
            V_终止时间 := V_Curdate;
            V_终止人员 := 操作员姓名_In;
          End If;
        
          Close C_Oldinfo;
          Open C_Oldinfo;
          Fetch C_Oldinfo
            Into R_Oldinfo;
        End If;
      
        If V_责任护士 <> Nvl(责任护士_In, 'NoneData') Then
          V_原因 := 8;
          If V_终止时间 Is Null Then
            Update 病人变动记录
            Set 终止时间 = V_Curdate, 终止原因 = V_原因, 终止人员 = 操作员姓名_In
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
          Else
            Update 病人变动记录
            Set 终止时间 = V_Curdate, 终止原因 = V_原因, 终止人员 = 操作员姓名_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = V_终止时间 And 终止原因 = V_终止原因;
            --更新将来的记录，如果有停止到将来的则删除上次计算时间 
            Update 病人变动记录
            Set 责任护士 = 责任护士_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > V_Curdate;
          End If;
          While C_Oldinfo%Found Loop
            Insert Into 病人变动记录
              (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 医疗小组id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情,
               操作员编号, 操作员姓名, 终止时间, 终止原因, 终止人员)
            Values
              (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, V_Curdate, V_原因, R_Oldinfo.附加床位, R_Oldinfo.病区id, R_Oldinfo.科室id,
               R_Oldinfo.医疗小组id, R_Oldinfo.护理等级id, R_Oldinfo.床位等级id, R_Oldinfo.床号, 责任护士_In, R_Oldinfo.经治医师,
               R_Oldinfo.主治医师, R_Oldinfo.主任医师, R_Oldinfo.病情, 操作员编号_In, 操作员姓名_In, V_终止时间, V_终止原因, V_终止人员);
            Fetch C_Oldinfo
              Into R_Oldinfo;
          End Loop;
          --如果存在停止到将来的变动就更新终止原因 
          If V_终止时间 Is Not Null Then
            V_终止原因 := V_原因;
            V_终止时间 := V_Curdate;
            V_终止人员 := 操作员姓名_In;
          End If;
          Close C_Oldinfo;
          Open C_Oldinfo;
          Fetch C_Oldinfo
            Into R_Oldinfo;
        End If;
      
        If V_主任医师 <> Nvl(主任医师_In, 'NoneData') Then
          Update 病案主页从表
          Set 信息值 = 主任医师_In
          Where 病人id = 病人id_In And 主页id = 主页id_In And 信息名 = '主任医师';
        
          If Sql%RowCount = 0 Then
            Insert Into 病案主页从表
              (病人id, 主页id, 信息名, 信息值)
            Values
              (病人id_In, 主页id_In, '主任医师', 主任医师_In);
          
          End If;
        
          V_原因 := 12;
          If V_终止时间 Is Null Then
            Update 病人变动记录
            Set 终止时间 = V_Curdate, 终止原因 = V_原因, 终止人员 = 操作员姓名_In
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null;
          
          Else
            Update 病人变动记录
            Set 终止时间 = V_Curdate, 终止原因 = V_原因, 终止人员 = 操作员姓名_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = V_终止时间 And 终止原因 = V_终止原因;
            --更新将来的记录如果有停止到将来的则删除上次计算时间 
            Update 病人变动记录
            Set 主任医师 = 主任医师_In, 上次计算时间 = Null
            Where 病人id = 病人id_In And 主页id = 主页id_In And 开始时间 > V_Curdate;
          End If;
        
          While C_Oldinfo%Found Loop
            Insert Into 病人变动记录
              (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 医疗小组id, 护理等级id, 床位等级id, 床号, 责任护士, 经治医师, 主治医师, 主任医师, 病情,
               操作员编号, 操作员姓名, 终止时间, 终止原因, 终止人员)
            Values
              (病人变动记录_Id.Nextval, 病人id_In, 主页id_In, V_Curdate, V_原因, R_Oldinfo.附加床位, R_Oldinfo.病区id, R_Oldinfo.科室id,
               R_Oldinfo.医疗小组id, R_Oldinfo.护理等级id, R_Oldinfo.床位等级id, R_Oldinfo.床号, R_Oldinfo.责任护士, R_Oldinfo.经治医师,
               R_Oldinfo.主治医师, 主任医师_In, R_Oldinfo.病情, 操作员编号_In, 操作员姓名_In, V_终止时间, V_终止原因, V_终止人员);
          
            Fetch C_Oldinfo
              Into R_Oldinfo;
          End Loop;
        
          Close C_Oldinfo;
          Open C_Oldinfo;
          Fetch C_Oldinfo
            Into R_Oldinfo;
        End If;
      
        Close C_Oldinfo;
        Close C_Endinfo;
        Select Count(*)
        Into V_Count
        From 病人变动记录
        Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(附加床位, 0) = 0 And 开始时间 Is Not Null And 终止时间 Is Null;
      
        If V_Count > 1 Then
          V_Error := '发现病人存在非法的变动记录,当前操作不能继续！' || Chr(13) || Chr(10) || '这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
          Raise Err_Custom;
        End If;
      
      End If;
    
    End If;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
  
End Zl_病案主页_首页整理;
/
Create Or Replace Procedure Zl_病人信息_首页整理
(
  --功能：用于门诊/住院医生站对病人进行首页整理 
  --参数： 
  --      标识号_IN:门诊为门诊号,住院为住院号 
  --      挂号单_IN:门诊医生站调用时才需要传入(用于区分) 
  --      就诊摘要_IN:门诊医生站调用时才需要传入 
  病人id_In       病人信息.病人id%Type,
  标识号_In       病人信息.住院号%Type,
  姓名_In         病人信息.姓名%Type,
  性别_In         病人信息.性别%Type,
  年龄_In         病人信息.年龄%Type,
  民族_In         病人信息.民族%Type,
  国籍_In         病人信息.国籍%Type,
  区域_In         病人信息.区域%Type,
  籍贯_In         病人信息.籍贯%Type,
  职业_In         病人信息.职业%Type,
  出生日期_In     病人信息.出生日期%Type,
  出生地点_In     病人信息.出生地点%Type,
  身份证号_In     病人信息.身份证号%Type,
  其他证件_In     病人信息.其他证件%Type,
  婚姻状况_In     病人信息.婚姻状况%Type,
  医疗付款方式_In 病人信息.医疗付款方式%Type,
  家庭地址_In     病人信息.家庭地址%Type,
  家庭电话_In     病人信息.家庭电话%Type,
  家庭地址邮编_In 病人信息.家庭地址邮编%Type,
  户口地址_In     病案主页.户口地址%Type,
  户口邮编_In     病案主页.户口地址邮编%Type,
  工作单位id_In   病人信息.合同单位id%Type,
  工作单位_In     病人信息.工作单位%Type,
  单位电话_In     病人信息.单位电话%Type,
  单位邮编_In     病人信息.单位邮编%Type,
  联系人姓名_In   病人信息.联系人姓名%Type := Null,
  联系人关系_In   病人信息.联系人关系%Type := Null,
  联系人电话_In   病人信息.联系人电话%Type := Null,
  联系人地址_In   病人信息.联系人地址%Type := Null,
  Email_In        病人信息.Email%Type := Null,
  Qq_In           病人信息.Qq%Type := Null,
  监护人_In       病人信息.监护人%Type := Null,
  挂号单_In       病人挂号记录.No%Type := Null,
  复诊_In         病人挂号记录.复诊%Type := Null,
  摘要_In         病人挂号记录.摘要%Type := Null,
  传染病上传_In   病人挂号记录.传染病上传%Type := Null,
  发病时间_In     病人挂号记录.发病时间%Type := Null,
  发病地址_In     病人挂号记录.发病地址%Type := Null
) Is
Begin
  Update 病人信息
  Set 门诊号 = Decode(挂号单_In, Null, 门诊号, 标识号_In), 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 出生日期 = 出生日期_In, 出生地点 = 出生地点_In,
      身份证号 = 身份证号_In, 民族 = 民族_In, 国籍 = 国籍_In, 区域 = 区域_In, 婚姻状况 = 婚姻状况_In, 职业 = 职业_In, 医疗付款方式 = 医疗付款方式_In, 家庭地址 = 家庭地址_In,
      家庭电话 = 家庭电话_In, 家庭地址邮编 = 家庭地址邮编_In, 工作单位 = 工作单位_In, 合同单位id = 工作单位id_In, 单位电话 = 单位电话_In, 单位邮编 = 单位邮编_In,
      联系人姓名 = Decode(联系人姓名_In, Null, 联系人姓名, 联系人姓名_In), 联系人关系 = Decode(联系人姓名_In, Null, 联系人关系, 联系人关系_In),
      联系人电话 = Decode(联系人姓名_In, Null, 联系人电话, 联系人电话_In), 联系人地址 = Decode(联系人姓名_In, Null, 联系人地址, 联系人地址_In),
      监护人 = Nvl(监护人_In, 监护人), 其他证件 = 其他证件_In, 户口地址 = 户口地址_In, 户口地址邮编 = 户口邮编_In, 籍贯 = 籍贯_In, Email = Nvl(Email_In, Email),
      Qq = Nvl(Qq_In, Qq)
  Where 病人id = 病人id_In;

  If 挂号单_In Is Not Null Then
    --门诊病案记录 
    If 标识号_In Is Not Null Then
      Update 门诊病案记录 Set 病案号 = 标识号_In Where 病人id = 病人id_In;
      If Sql%RowCount = 0 Then
        Insert Into 门诊病案记录
          (病人id, 病案号, 建立日期, 病案类别, 存储状态, 存放位置)
        Values
          (病人id_In, 标识号_In, Sysdate, '一般', '正常', Null);
      End If;
    Else
      Delete From 门诊病案记录 Where 病人id = 病人id_In;
    End If;
  
    --同步修改挂号记录 
    Update 病人挂号记录
    Set 门诊号 = 标识号_In, 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 复诊 = 复诊_In, 摘要 = 摘要_In, 传染病上传 = 传染病上传_In, 发病时间 = 发病时间_In,
        发病地址 = 发病地址_In
    Where NO = 挂号单_In;
  
    Update 门诊费用记录
    Set 标识号 = 标识号_In, 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 结论 = 摘要_In
    Where NO = 挂号单_In And 记录性质 = 4;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人信息_首页整理;
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
  病人id_In   In 病人感染记录.病人id%Type,
  主页id_In   In 病人感染记录.主页id%Type,
  序号_In     In 病人感染记录.序号%Type,
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