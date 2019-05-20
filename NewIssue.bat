@echo off
cls

for /f "delims=" %%a in ('wmic OS Get localdatetime ^| find "."') do set DateTime=%%a

set Yr=%DateTime:~0,4%
set Mon=%DateTime:~4,2%
set Day=%DateTime:~6,2%
set Hr=%DateTime:~8,2%
set Min=%DateTime:~10,2%
set Sec=%DateTime:~12,2%

set timestamp=%Yr%%Mon%%Day%%Hr%%Min%%Sec%

echo ##########################################
echo #                                        #
echo #            Issue Management            #
echo #                  by                    #  
echo #               David Tsang              #
echo #               20 May 2019              #
echo #                                        #
echo ##########################################
echo.
echo.
echo.
SET /P uname=Name: 
echo.
echo.
SET /P utitle=Title: 
echo.
echo.
SET /P udesc=Description: 
echo.
echo.
SET /P uprogramId=Program Id: 
echo.
echo.

SET dname=%timestamp%-%uname%-%utitle%
SET dname=%dname: =-%

mkdir %dname%

SET fpath=%dname%\%dname%-Note.txt

set timestamp2=%Yr%-%Mon%-%Day% %Hr%:%Min%:%Sec%

echo Name: %uname%>>%fpath%
echo Date: %timestamp2%>>%fpath%
echo Program Id: %uprogramId%>>%fpath%
echo Description: %udesc%>>%fpath%

echo %dname% created
echo.
echo.
pause

NewIssue
