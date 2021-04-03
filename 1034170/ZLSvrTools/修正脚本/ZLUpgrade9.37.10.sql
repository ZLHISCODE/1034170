--30537
Alter Table zlTools.zlReports Modify หตร๗ Varchar2(2000)
/

--30993
Create Or Replace Function Zltools.f_Num2list
(
  Str_In   In Varchar2,
  Split_In In Varchar2 := ','
) Return t_Numlist
  Pipelined As
  v_Str Long;
  P     Number;
Begin
  If Str_In Is Null Then
    Return;
  End If;
  v_Str := Str_In || Split_In;
  Loop
    P := Instr(v_Str, Split_In);
    Exit When(Nvl(P, 0) = 0);
    Pipe Row(To_Number(Trim(Substr(v_Str, 1, P - 1))));
    v_Str := Substr(v_Str, P + 1);
  End Loop;
  Return;
End;
/

Create Or Replace Function Zltools.f_Num2list2
(
  Str_In      In Varchar2,
  Split_In    In Varchar2 := ',',
  Subsplit_In In Varchar2 := ':'
) Return t_Numlist2
  Pipelined As
  v_Str   Long;
  P       Number;
  v_Tmp   Varchar2(4000);
  Out_Rec t_Numobj2 := t_Numobj2(Null, Null);
Begin
  If Str_In Is Null Then
    Return;
  End If;
  v_Str := Str_In || Split_In;
  Loop
    P := Instr(v_Str, Split_In);
    Exit When(Nvl(P, 0) = 0);
    v_Tmp      := Trim(Substr(v_Str, 1, P - 1));
    Out_Rec.C1 := To_Number(Substr(v_Tmp, 1, Instr(v_Tmp, Subsplit_In) - 1));
    Out_Rec.C2 := To_Number(Substr(v_Tmp, Instr(v_Tmp, Subsplit_In) + 1));
    Pipe Row(Out_Rec);
    v_Str := Substr(v_Str, P + 1);
  End Loop;
  Return;
End;
/
