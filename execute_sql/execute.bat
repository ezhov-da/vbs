@rem==================
@rem кодировка в CP866
@rem==================

@echo off

SET NAME_VBS=execute_sql.vbs

SET PATH_TO_EXECUTE_DIRECTORY=%~dp0
echo Путь к папке запуска bat файла: %PATH_TO_EXECUTE_DIRECTORY%

SET LOG_FILE="%PATH_TO_EXECUTE_DIRECTORY%log_minpart_mp.txt"
echo Лог файл: %LOG_FILE%

SET CONNECTIONS="teradata->Provider=MSDASQL.1;Persist Security Info=False;Data Source=teradata_auth"
echo Подключения: %CONNECTIONS%

SET SCRIPTS="teradata->E:\_own_repository_data\projects_dev\macros\minpart\SQL_QUERY.sql"
echo Скрипты: %SCRIPTS%

SET EXECUTE_VBS=%PATH_TO_EXECUTE_DIRECTORY%%NAME_VBS%

cscript %EXECUTE_VBS% %LOG_FILE% %CONNECTIONS% %SCRIPTS%
