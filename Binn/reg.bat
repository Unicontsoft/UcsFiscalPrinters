for %%i in (*.dll;*.ocx) do regsvr32 /s %1 %%i
regtlib *.tlb