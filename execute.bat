@echo off

SET PATH_TO_BAT=%~dp0

SET VBS=executor.vbs  

SET CONNECTIONS="Provider=MSDASQL.1;Persist Security Info=False;Data Source=OTZ-prod1"

SET QUERY="execute OTZ.[dbo].prc_aa_diagnostik_da_only"

SET LOG_FILE="%PATH_TO_BAT%log_execute_prc_aa_diagnostik_da_only.txt"

SET EXECUTE_VBS=%PATH_TO_BAT%%VBS%

chcp 866 && cscript %EXECUTE_VBS% %CONNECTIONS% %QUERY% %LOG_FILE%
