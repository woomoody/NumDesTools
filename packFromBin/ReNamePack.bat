@echo off

set "destinationFolder=%~dp0"
cd /d "%destinationFolder%"

REM 删除Bat下旧文件
set "targetFileName64=NumDesToolsPack64.XLL"  
set "targetFileName32=NumDesToolsPack.XLL"  
if exist "%targetFileName64%" (
    del "%targetFileName64%"
    echo File %targetFileName64% deleted successfully.
) else (
    echo File %targetFileName64% not found.
)
if exist "%targetFileName32%" (
    del "%targetFileName32%"
    echo File %targetFileName32% deleted successfully.
) else (
    echo File %targetFileName32% not found.
)

REM 复制bin下新文件
set "sourceFileName64=NumDesTools-AddIn64-packed.xll"   
set "sourceFileName32=NumDesTools-AddIn-packed.xll"   

set "sourceFile64=NumDesTools\bin\Debug\%sourceFileName64%" 
set "sourceFile32=NumDesTools\bin\Debug\%sourceFileName32%"   


set "beforDestination=%destinationFolder:~0,-1%"
for %%I in ("%beforDestination%") do set "beforDestination=%%~dpI"

set "fullSourcePath64=%beforDestination%%sourceFile64%"
set "fullSourcePath32=%beforDestination%%sourceFile32%"

if exist "%fullSourcePath64%" (
    copy "%fullSourcePath64%" "%destinationFolder%"
    echo File %fullSourcePath64% copied successfully.
) else (
    echo Source %fullSourcePath64% file not found.
)
if exist "%fullSourcePath32%" (
    copy "%fullSourcePath32%" "%destinationFolder%"
    echo File %fullSourcePath32% copied successfully.
) else (
    echo Source %fullSourcePath32% file not found.
)


REM 新文件重命名为旧文件
if exist "%sourceFileName64%" (
    ren "%sourceFileName64%" "%targetFileName64%"
    echo File #%sourceFileName64%# renamed #%targetFileName64%# successfully.
) else (
    echo File %sourceFileName64% not found.
)
if exist "%sourceFileName32%" (
    ren "%sourceFileName32%" "%targetFileName32%"
    echo File #%sourceFileName32%# renamed #%targetFileName32%# successfully.
) else (
    echo File %sourceFileName32% not found.
)

exit /b
