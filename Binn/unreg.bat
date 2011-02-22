for %%i in (*.dll;*.ocx) do regsvr32 /s /u %1 %%i
rem regtlib *.tlb