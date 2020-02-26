@echo off
setlocal
set "out_file=%~dp0combined-MySQL.sql"
set "sql_dir=%~dp0MySQL"

echo -- This is an amalgamation of all MySQL scripts> %out_file%
echo.>> %out_file%
type %sql_dir%\schema.sql >> %out_file%
echo.>> %out_file%
type %sql_dir%\usp_umq_setup_service.sql >> %out_file%
echo.>> %out_file%
type %sql_dir%\usp_umq_send.sql >> %out_file%
echo.>> %out_file%
type %sql_dir%\usp_umq_wait_request.sql >> %out_file%
echo.>> %out_file%
type %sql_dir%\usp_umq_test.sql >> %out_file%

