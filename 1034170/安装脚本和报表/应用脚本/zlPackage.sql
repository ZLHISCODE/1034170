--XWPACS接口包
create or replace package b_XINWANGInterface is
  Type t_Refcur Is Ref Cursor;
-- 功    能：PACS状态改变信息
  Procedure PacsStatusChange
  (
    状态ID_In In Number,
    医嘱ID_In In 影像检查记录.医嘱ID%Type, 
    影像类别_In In 影像检查记录.影像类别%Type, 
    检查号_In In 影像检查记录.检查号%Type,
    处理时间  In date,
    执行人 In  Varchar2,
    胶片大小 In Varchar2
  );
-- 功    能：取消图像关联
  Procedure PacsUnmatchImage
  (
    医嘱ID_In In 影像检查记录.医嘱ID%Type 
  );
-- 功    能：填写报告图的存储设备
  Procedure PacsSetFTPDeviceNo
  (
    医嘱ID_In In 影像检查记录.医嘱ID%Type ,
    设备号_In In 影像检查记录.位置一%Type
  );
-- 功    能：更新图像数
  Procedure UpdateImgCount
  (
    医嘱ID_In 影像检查记录.医嘱ID%Type,
    图像数_In NUMBER
  );  

end b_XINWANGInterface ;
/


create or replace package body b_XINWANGInterface  is

-- 功    能：PACS状态改变信息
  Procedure PacsStatusChange
  (
    状态ID_In In Number,
    医嘱ID_In In 影像检查记录.医嘱ID%Type, 
    影像类别_In In 影像检查记录.影像类别%Type, 
    检查号_In In 影像检查记录.检查号%Type,
    处理时间  In date,
    执行人 In  Varchar2,
    胶片大小 In Varchar2
  ) Is
   Strsql           Varchar2(2000);
    Cursor c_Advice Is
    Select id
    From  病人医嘱记录 
    Where Id = 医嘱id_In Or (相关id = 医嘱id_In And 诊疗类别 In ('F', 'G', 'D'));
    
  Begin
       --状态ID_In:1-匹配成功;2-匹配失败;3-新检查（收到第一幅图像）;4-收到每一幅图像;5-删除检查;6-胶片打印成功
       If 状态ID_In = 1 Then 
          --图象匹配成功
      
          --填写影像检查记录表的 检查UID，接收日期等，但是不填写序列级别的表,检查UID填写医嘱ID
          update 影像检查记录 set 检查UID= 医嘱ID_In ,接收日期 = Decode(处理时间, Null, Sysdate, 处理时间),图像位置=1 Where 医嘱ID = 医嘱ID_In ;
          
          --设置医嘱执行状态
          For r_Advice In c_Advice Loop
              Update 病人医嘱发送
                     Set 执行状态 = 3, 执行过程 = Decode(Sign(执行过程 - 2), 1, 执行过程, 3)
                     Where 医嘱id = r_Advice.id;
          End Loop;
       Elsif 状态ID_In = 2 Then 
          Strsql :='dd';
       Elsif 状态ID_In = 3 Then 
          -- 3-新检查（收到第一幅图像），暂时不处理
          Strsql :='dd';
       Elsif 状态ID_In = 4 Then 
          --  4-收到每一幅图像 ，暂时不处理
          Strsql :='dd';
       Elsif 状态ID_In = 5 Then 
          -- 5-删除检查
          -- 删除影像检查记录表中对应的检查UID，接收日期等
          update 影像检查记录 set 检查UID=null,位置一=null,位置二=null,位置三=null,报告图象=null,接收日期=null
                 where 医嘱ID = 医嘱ID_IN;
       Elsif 状态ID_In = 6 Then 
          -- 6-胶片打印成功
          --记录胶片大小，打印人等

          --一个医嘱打印一张或者多张胶片的情况，每张胶片调用一过程，相关ID为空
          Insert Into 胶片打印记录 (ID, 相关id, 医嘱id, 胶片大小, 打印人, 打印时间)
                 Values (胶片打印记录_Id.Nextval, Null, 医嘱ID_In, 胶片大小, 执行人,Decode(处理时间, Null, Sysdate, 处理时间));
          Update 影像检查记录 Set 是否打印 = 1 Where  医嘱ID =医嘱ID_In;
       Elsif 状态ID_In = 7 then
         --更新电子胶片状态
          Update 影像检查记录 Set 是否电子胶片 = 1 Where  医嘱ID =医嘱ID_In;        
       End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End PacsStatusChange;
  
-- 功    能：PACS图像取消关联
  Procedure PacsUnmatchImage
  (
    医嘱ID_In In 影像检查记录.医嘱ID%Type 
  ) Is
   v_执行过程  病人医嘱发送.执行过程%Type; 
   v_发送号  病人医嘱发送.发送号%Type; 
  Begin
       --设置影像检查记录表的状态
       update 影像检查记录 set 检查UID= Null ,接收日期 = Null,图像位置=null ,位置一 =Null Where 医嘱ID = 医嘱ID_In ;
       
       --调用 Zl_影像检查_State 改变检查过程的状态
       Select 执行过程,发送号 Into v_执行过程,v_发送号 From 病人医嘱发送 Where 医嘱ID = 医嘱ID_In;
       
       --如果执行过程是3，则将过程修改成2
       If v_执行过程 = 3 Then 
          Zl_影像检查_State(医嘱ID_In,v_发送号,2);
       End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End PacsUnmatchImage;
  
-- 功    能：填写报告图的存储设备
  Procedure PacsSetFTPDeviceNo
  (
    医嘱ID_In In 影像检查记录.医嘱ID%Type,
    设备号_In In 影像检查记录.位置一%Type
  ) Is
  Begin
       --设置影像检查记录表的状态
       update 影像检查记录 set 位置一= 设备号_In Where 医嘱ID = 医嘱ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End PacsSetFTPDeviceNo;


--RISPACS更新图像数量需要调用的存储过程
  procedure UpdateImgCount
  (
    医嘱ID_IN    影像检查记录.医嘱ID%Type,
    图像数_In    NUMBER
  ) is
  begin
      update 影像检查记录 set 图像数量=图像数_In where 医嘱ID=医嘱ID_In;
  Exception
      When Others Then
          Zl_Errorcenter(Sqlcode, Sqlerrm);  
  end UpdateImgCount;

end b_XINWANGInterface;
/
