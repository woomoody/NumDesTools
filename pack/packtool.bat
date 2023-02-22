set packpath=%1%

%packpath%ExcelDnaPack %packpath%NumDesTools-AddIn64.dna /Y /O %packpath%NumDesToolsPack64.XLL
%packpath%ExcelDnaPack %packpath%NumDesTools-AddIn.dna /Y /O %packpath%NumDesToolsPack.XLL

pause
