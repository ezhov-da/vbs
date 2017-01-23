'------------------------------------------------------------------------
'--> автор				: Ежов Д.А.
'--> дата создания		: 2015-10-08	
'--> описание			: универсальный скрипт для запуска запросов в бд
'						  данный vbs запускается через cmd или bat:
'						  cscript path\executor.vbs param1, param2
'						  для запуска файла из дериктории cmd или bat,
'						  пишем %~dp0
'--> входные параметры	: 1. строка подключения, 2. запрос на выполнение
'------------------------------------------------------------------------
Dim connectString	'строка подключения
Dim command			'команда на выполнение
Dim nameLogFile		'название файла для логирования

'пробуем прочитать входные параметры
on error resume next
'присваиваем аргументы
	connectString 	= WScript.Arguments(0)
	command 		= WScript.Arguments(1)
	nameLogFile 	= WScript.Arguments(2)
on error goto 0	

WScript.echo ">" & now() & " получили входные параметры"
WScript.echo ">" & now() & " connectString:" & connectString
WScript.echo ">" & now() & " command:" & command
WScript.echo ">" & now() & " nameLogFile:" & nameLogFile

'БЛОК ЛОГИРОВАНИЯ VBS---------------------------------------------------
Dim WshShell
Dim objFS
Dim objFile
Dim file
Dim tfile
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFS = CreateObject("Scripting.FileSystemObject")
file = nameLogFile
if objFS.FileExists(file) then
	Set tfile = objFS.GetFile(file)
	'8 - открываем для добавления
	'1 - для чтения
	'2 - для записи. 
	'--
	'-2 используем системную кодировку
	Set objFile = tfile.OpenAsTextStream(8, -2)	
	else
		Set objFile = objFS.CreateTextFile(file, False)
end if
' ----------------------------------------------------------------------
objFile.WriteLine now() & " получаем параметры" 
'проверяем наличие параметров
if connectString = "" or command = "" or nameLogFile = "" then
	objFile.WriteLine now() & " не указаны входные параметры" 
	msgBox "Не указаны три входных параметра:" & Chr(13) & "1. строка подключения" & Chr(13) & "2. запрос на выполнение" & Chr(13) & "3. название файла для лога"
	WScript.Quit
end if

objFile.WriteLine now() & " проверка на параметры пройдена" 	
objFile.WriteLine now() & " 1.command: " & connectString	
objFile.WriteLine now() & " 2.connectString: " & command	
objFile.WriteLine now() & " 3.nameLogFile: " & nameLogFile	
	
On Error Resume Next	'отключаем ошибки, чтоб самим их обрабатывать
	
'начинаем выполнение запроса
WScript.echo ">" & now() & " начинаем выполнение запроса"

dim ADO	
Set ADO = CreateObject("ADODB.Connection")
ADO.ConnectionTimeout = 0
ADO.CommandTimeout = 0
ADO.Open connectString

'проверяем наличие ошибок при подключении
if Err.Number <> 0 then
		objFile.WriteLine now() & " скрипт выполнен с ошибками: " & Err.Description  
		MsgBox "Cкрипт выполнен с ошибками:" & Chr(13) & Err.Description  
		WScript.Quit
End if

objFile.WriteLine now() & " создали подключение" 
WScript.echo ">" & now() & " создали подключение"
objFile.WriteLine now() & " начали выполнять запрос" 
WScript.echo ">" & now() & " начали выполнять запрос"
ADO.Execute command

'проверяем наличие ошибок при выполнении запроса
Select Case Err.Number
	Case 0 'Все в порядке
		objFile.WriteLine now() & " запрос выполнен" 
		WScript.echo ">" & now() & " запрос выполнен"
	Case Else
		objFile.WriteLine now() & " скрипт выполнен с ошибками: " & Err.Description  
		MsgBox "Cкрипт выполнен с ошибками:" & Chr(13) & Err.Description  
		WScript.echo ">" & now() & " скрипт выполнен с ошибками: " & Err.Description  
End Select

ADO.close
set ADO = nothing
'ЗАВЕРШЕНИЕ БЛОКА ЛОГИРОВАНИЯ-------------------------------------------	
objFile.Close
Set objFile = Nothing
Set objFS = Nothing
' ----------------------------------------------------------------------
