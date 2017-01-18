@rem ===================================
@rem разработчик Ежов Д.А.
@rem GitHub: https://github.com/ezhov-da
@rem ===================================

@echo off

@rem !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
@rem КОДИРОВКА ФАЙЛА ДОЛЖНА БЫТЬ В OEM866
@rem !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

@rem этот BAT файл запускает скрипт, который производит установку указанной в данном файле надстройки
@rem в этом файле указываются:
@rem - путь к скрипту. Пример: E:\_own_repository_data\vbs\install_excel_addins\install_excel_addins.vbs
@rem - путь к надстройке. Пример: X:\Категорийные\Кожанов\
@rem - название надстройки без расширени. Пример: test_ezhov
@rem - расширение. Пример: xla

@rem путь к исполняемому скрипту
SET PATH_TO_VBS="X:\Категорийные\Автоматизированная система анализа ассортимента\Инструменты АСАА\Минпартия\install_excel_addins.vbs"
@rem путь к надстройке
SET PATH_TO_ADDIN="X:\Категорийные\Автоматизированная система анализа ассортимента\Инструменты АСАА\Минпартия\"
@rem название файла
SET ADDIN_NAME="PROBLEM-MINQUANTITY(20161017)_0.22"
@rem расширеним надстройки
SET ADDIN_EXT="xlam"
@rem 1 - показывать, 0 - не показывать
SET IS_SHOW_MSG="1" 

echo Входные параметры:
echo PATH_TO_ADDIN: %PATH_TO_ADDIN%
echo PATH_TO_VBS: %PATH_TO_VBS%
echo ADDIN_NAME: %ADDIN_NAME%
echo ADDIN_EXT: %ADDIN_EXT%
echo IS_SHOW_MSG: %IS_SHOW_MSG%

CScript %PATH_TO_VBS% %PATH_TO_ADDIN% %ADDIN_NAME% %ADDIN_EXT% %IS_SHOW_MSG%

echo ErrorLevel: %ErrorLevel%

if ErrorLevel 1 (
	C:\Windows\system32\cscript.exe %PATH_TO_VBS% %PATH_TO_ADDIN% %ADDIN_NAME% %ADDIN_EXT% %IS_SHOW_MSG%
)

if ErrorLevel 1 (
	pause
)
