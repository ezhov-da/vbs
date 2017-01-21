@rem==================
@rem кодировка в CP866
@rem==================

@echo off
chcp 866 && cscript %~dp0execute_sql.vbs "log_inactive_mp.txt" "teradata->Provider=MSDASQL.1;Persist Security Info=False;Data Source=teradata_auth---mssql->Provider=MSDASQL.1;Persist Security Info=False;Data Source=OTZ-prod1" "teradata->K:\_Departments\Департамент технологий закупок\Отдел технологий закупок_\00. Папки сотрудников\Арефин\Актуальные_коды\Местные поставщики\script\Неактивные МП код.txt---mssql->E:\_own_repository_data\projects_dev\scripts\inactive_mp\execute_load_to_mssql.sql"
pause
