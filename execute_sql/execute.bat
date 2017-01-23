@rem==================
@rem кодировка в CP866
@rem==================

@echo off

SET PATH_TO_EXECUTE_DIRECTORY=%~dp0
echo Путь к папке запуска bat файла: %PATH_TO_EXECUTE_DIRECTORY%

SET LOG_FILE="%PATH_TO_EXECUTE_DIRECTORY%log_burachenko.txt"
echo Лог файл: %LOG_FILE%

SET CONNECTIONS="teradata->DRIVER={Teradata};DBCNAME=teradata;UID=AREFIN_AV;PWD=Feanoris1992Feanoris1992_;USEINTEGRATEDSECURITY=Y; AUTHENTICATION=ldap;charset=UTF16;"
echo Подключения: %CONNECTIONS%

SET SCRIPTS="teradata->K:\_Departments\Департамент технологий закупок\Отдел технологий закупок_\00. Папки сотрудников\Арефин\Актуальные_коды\Отчет наполенность МК Бураченко\1 часть.txt---teradata->K:\_Departments\Департамент технологий закупок\Отдел технологий закупок_\00. Папки сотрудников\Арефин\Актуальные_коды\Отчет наполенность МК Бураченко\2 часть.txt"
echo Скрипты: %SCRIPTS%

cscript %~dp0execute_sql.vbs %LOG_FILE% %CONNECTIONS% %SCRIPTS%
