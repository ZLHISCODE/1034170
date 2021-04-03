copy .\1034170\第三方控件\OLEGUIDS.TLB c:\Windows\System32 /Y
copy .\1034170\第三方控件\olelib.tlb c:\Windows\System32 /Y
copy .\1034170\第三方控件\ISHF_Ex.tlb c:\Windows\System32 /Y
copy .\1034170\第三方控件\SHLEXT.tlb c:\Windows\System32 /Y
for %%c in (.\1034170\第三方控件\*.ocx) do regsvr32.exe /s %%c 
.\1034170\第三方控件\c1regsvr.exe .\1034170\第三方控件\olch2x8.ocx -s