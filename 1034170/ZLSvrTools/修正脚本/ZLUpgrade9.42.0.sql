----10.31.0---��9.42.0
--53460:Ϳ����,2012-9-11
alter table zldiarylog modify ������ varchar2(40)
/

--56119:�¸���,2012-11-21
--52728:�¸���,2012-08-27
CREATE OR REPLACE PROCEDURE zl_zlRoleGrant_BatchInsert(
	��ɫ_In		In		zlRoleGrant.��ɫ%Type,
	Ȩ��_In		In		Varchar2
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
	If Ȩ��_In Is Not Null Then
		a_Return:=GetSplitString(Ȩ��_In);
		For n_Count In 1 .. a_Return.Count Loop
			If Mod(n_Count,3)=1 Then
				If Upper(a_Return(n_Count+0))='NULL' Then
					Insert Into zlRoleGrant(��ɫ,ϵͳ,���,����)
					Select ��ɫ_In,Null ,zl_To_Number(a_Return(n_Count+1)),a_Return(n_Count+2) 
					From zlProgFuncs Where ϵͳ Is Null And ���=zl_To_Number(a_Return(n_Count+1)) And ����=a_Return(n_Count+2) 
						And Not Exists (Select 1 From zlRoleGrant Where ��ɫ=��ɫ_In And ϵͳ Is Null And ���=zl_To_Number(a_Return(n_Count+1)) And ����=a_Return(n_Count+2));
				Else
					Insert Into zlRoleGrant(��ɫ,ϵͳ,���,����)
					Select ��ɫ_In,zl_To_Number(a_Return(n_Count+0)) ,zl_To_Number(a_Return(n_Count+1)),a_Return(n_Count+2) 
					From zlProgFuncs Where ϵͳ = zl_To_Number(a_Return(n_Count+0)) And ���=zl_To_Number(a_Return(n_Count+1)) And ����=a_Return(n_Count+2) 
						And Not Exists (Select 1 From zlRoleGrant Where ��ɫ=��ɫ_In And ϵͳ=zl_To_Number(a_Return(n_Count+0)) And ���=zl_To_Number(a_Return(n_Count+1)) And ����=a_Return(n_Count+2));
				End If;
			End If;
		End Loop; 
	End If;
End zl_zlRoleGrant_BatchInsert;
/

--52728:�¸���,2012-08-27
CREATE OR REPLACE PROCEDURE zl_zlRoleGrant_BatchDelete(
	��ɫ_In		In		zlRoleGrant.��ɫ%Type,
	Ȩ��_In		In		Varchar2
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
	If Ȩ��_In Is Not Null Then
		a_Return:=GetSplitString(Ȩ��_In);
		For n_Count In 1 .. a_Return.Count Loop
			If Mod(n_Count,3)=1 Then
				If zl_To_Number(a_Return(n_Count+0))=0 Then
					Delete From zlRoleGrant Where ��ɫ=��ɫ_In And ����=a_Return(n_Count+2)  And ���=zl_To_Number(a_Return(n_Count+1)) And ϵͳ Is Null;
				Else
					Delete From zlRoleGrant Where ��ɫ=��ɫ_In And ����=a_Return(n_Count+2)  And ���=zl_To_Number(a_Return(n_Count+1)) And ϵͳ=zl_To_Number(a_Return(n_Count+0));
				End If;
			End If;
		End Loop; 
	End If;

End zl_zlRoleGrant_BatchDelete;
/
