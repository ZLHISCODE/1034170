copy .\1034170\�������ؼ�\OLEGUIDS.TLB c:\Windows\System32 /Y
copy .\1034170\�������ؼ�\olelib.tlb c:\Windows\System32 /Y
copy .\1034170\�������ؼ�\ISHF_Ex.tlb c:\Windows\System32 /Y
copy .\1034170\�������ؼ�\SHLEXT.tlb c:\Windows\System32 /Y
for %%c in (.\1034170\�������ؼ�\*.ocx) do regsvr32.exe /s %%c 
.\1034170\�������ؼ�\c1regsvr.exe .\1034170\�������ؼ�\olch2x8.ocx -s