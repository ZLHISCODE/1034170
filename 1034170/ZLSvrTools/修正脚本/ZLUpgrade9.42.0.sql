----10.31.0---》9.42.0
--53460:涂建华,2012-9-11
alter table zldiarylog modify 窗体名 varchar2(40)
/

--56119:陈福容,2012-11-21
--52728:陈福容,2012-08-27
CREATE OR REPLACE PROCEDURE zl_zlRoleGrant_BatchInsert(
	角色_In		In		zlRoleGrant.角色%Type,
	权限_In		In		Varchar2
)
Is
	a_Return t_Strlist := t_Strlist(); 
	---------------------------------------------------------------------------------------------------------------
	Function GetSplitString(Str_In In Varchar2) Return t_Strlist As

	  v_Str   Long Default Str_In || ''''; 
	  v_Index Number; 
	  v_List  t_Strlist := t_Strlist(); 

	Begin 
	  Loop 
	    v_Index := Instr(v_Str, ''''); 
	    Exit When(Nvl(v_Index, 0) = 0); 
	    v_List.Extend; 
	    v_List(v_List.Count) := Trim(Substr(v_Str, 1, v_Index - 1)); 
	    v_Str := Substr(v_Str, v_Index + 1); 
	  End Loop; 
	  Return v_List;
	End;
Begin
	If 权限_In Is Not Null Then
		a_Return:=GetSplitString(权限_In);
		For n_Count In 1 .. a_Return.Count Loop
			If Mod(n_Count,3)=1 Then
				If Upper(a_Return(n_Count+0))='NULL' Then
					Insert Into zlRoleGrant(角色,系统,序号,功能)
					Select 角色_In,Null ,zl_To_Number(a_Return(n_Count+1)),a_Return(n_Count+2) 
					From zlProgFuncs Where 系统 Is Null And 序号=zl_To_Number(a_Return(n_Count+1)) And 功能=a_Return(n_Count+2) 
						And Not Exists (Select 1 From zlRoleGrant Where 角色=角色_In And 系统 Is Null And 序号=zl_To_Number(a_Return(n_Count+1)) And 功能=a_Return(n_Count+2));
				Else
					Insert Into zlRoleGrant(角色,系统,序号,功能)
					Select 角色_In,zl_To_Number(a_Return(n_Count+0)) ,zl_To_Number(a_Return(n_Count+1)),a_Return(n_Count+2) 
					From zlProgFuncs Where 系统 = zl_To_Number(a_Return(n_Count+0)) And 序号=zl_To_Number(a_Return(n_Count+1)) And 功能=a_Return(n_Count+2) 
						And Not Exists (Select 1 From zlRoleGrant Where 角色=角色_In And 系统=zl_To_Number(a_Return(n_Count+0)) And 序号=zl_To_Number(a_Return(n_Count+1)) And 功能=a_Return(n_Count+2));
				End If;
			End If;
		End Loop; 
	End If;
End zl_zlRoleGrant_BatchInsert;
/

--52728:陈福容,2012-08-27
CREATE OR REPLACE PROCEDURE zl_zlRoleGrant_BatchDelete(
	角色_In		In		zlRoleGrant.角色%Type,
	权限_In		In		Varchar2
)
Is
	a_Return t_Strlist := t_Strlist(); 
	---------------------------------------------------------------------------------------------------------------
	Function GetSplitString(Str_In In Varchar2) Return t_Strlist As

	  v_Str   Long Default Str_In || ''''; 
	  v_Index Number; 
	  v_List  t_Strlist := t_Strlist(); 

	Begin 
	  Loop 
	    v_Index := Instr(v_Str, ''''); 
	    Exit When(Nvl(v_Index, 0) = 0); 
	    v_List.Extend; 
	    v_List(v_List.Count) := Trim(Substr(v_Str, 1, v_Index - 1)); 
	    v_Str := Substr(v_Str, v_Index + 1); 
	  End Loop; 
	  Return v_List;
	End;
Begin
	If 权限_In Is Not Null Then
		a_Return:=GetSplitString(权限_In);
		For n_Count In 1 .. a_Return.Count Loop
			If Mod(n_Count,3)=1 Then
				If zl_To_Number(a_Return(n_Count+0))=0 Then
					Delete From zlRoleGrant Where 角色=角色_In And 功能=a_Return(n_Count+2)  And 序号=zl_To_Number(a_Return(n_Count+1)) And 系统 Is Null;
				Else
					Delete From zlRoleGrant Where 角色=角色_In And 功能=a_Return(n_Count+2)  And 序号=zl_To_Number(a_Return(n_Count+1)) And 系统=zl_To_Number(a_Return(n_Count+0));
				End If;
			End If;
		End Loop; 
	End If;

End zl_zlRoleGrant_BatchDelete;
/
