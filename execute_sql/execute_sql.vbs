'===================================
'разработчик Ежов Д.А.
'GitHub: https://github.com/ezhov-da
'===================================
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'КОДИРОВКА ФАЙЛА ДОЛЖНА БЫТЬ В ANSI
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Данный скрипт выполняет запросы из файлов.
'
'Файлы со скриптами в кодировке ANSI
'
'Принимаемые парметры:
'1. Название файла для логирования: log_inactive_mp.txt
'2. список подключений: teradata->Provider=MSDASQL.1;Persist Security Info=False;Data Source=teradata_auth---mssql->...
'все подключения, которые будут использоваться файлами со скриптами
'3. Подключения и файлы со скриптами: teradata->C:\aaa.txt---mssql->C:\bbb.sql///
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'разделитель для свойств
dim CONST_SEPARATOR_PROP: CONST_SEPARATOR_PROP = "->"
'разделитель для списков
dim CONST_SEPARATOR_LIST: CONST_SEPARATOR_LIST = "---"

Dim ADO


'получаем входные параметры
Dim objArgs : set objArgs = WScript.Arguments
Dim fileLog 		: fileLog = 		objArgs(0)		'файл для логирования
Dim connections 	: connections = 	objArgs(1)	'настройки подключения
Dim filesExecute	: filesExecute =	objArgs(2)	'файлы для обработки

Wscript.Echo "[лог файл]: " & fileLog
Wscript.Echo "[подключения]: " & connections
Wscript.Echo "[файлы]: " & filesExecute

Dim dicConnections 	: set dicConnections = CreateObject("Scripting.Dictionary")
Dim dicFiles 		: set dicFiles = CreateObject("Scripting.Dictionary")

call fillDicConnections
call fillDicFiles

'on error resume next

'БЛОК ЛОГИРОВАНИЯ VBS---------------------------------------------------
Dim WshShell	: set WshShell = WScript.CreateObject("WScript.Shell")
Dim objFS 		: set objFS = CreateObject("Scripting.FileSystemObject")
Dim file 		: file = fileLog
Dim objFile
Dim tfile

WScript.Echo "[путь к файлу лога]: " & objFS.GetAbsolutePathName(file)
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

call log(" ")
call log("[начинаем работу...]====================================================>")

dim dicItems
dicItems = dicFiles.Items()

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

dim elementCount
for elementCount = lbound(dicItems) to ubound(dicItems)
	'Получаем файл и подключение для работы
	Dim querysFiles
	querysFiles = dicItems(elementCount)
	'Парсим на подключение и путь к файлу
	Dim arrayFile 
	arrayFile  = split(querysFiles, CONST_SEPARATOR_PROP)
	
	Dim connection
	connection	= arrayFile(0)
	Dim filePath
	filePath = arrayFile(1)
	
	call log("[используемая строка подключения]: " & connection)
	call log("[используемый файл]: " & filePath)
	
	'Получаем данные из файла
	Set txtfile = FSO.OpenTextFile(filePath)
	dim textFromFile
	textFromFile = txtfile.ReadALL
	
	Set ADO = CreateObject("ADODB.Connection")                                                              
	connectString = dicConnections.Item(connection)
	ADO.ConnectionTimeout = 0                                                                               
	ADO.CommandTimeout = 0                                                                                  
	ADO.Open connectString  	
	
	call log("[создали и открыли подключение]")
	
	Dim resultExecute
	resultExecute = executeTextQuerys(textFromFile)
	
	if (resultExecute) then
			call log("[файл выполнен успешно]")
		else
			call log("[~!файл выполнен с ошибками!~]")
			ADO.close  
			set ADO = nothing  
			exit for
	end if
	
	ADO.close     
	set ADO = nothing  
next

call log("[~все скрипты выполнены~]<====================================================")
call log(" ")

set FSO = nothing

'ЗАВЕРШЕНИЕ БЛОКА ЛОГИРОВАНИЯ-------------------------------------------	
objFile.Close
Set objFile = Nothing
Set objFS = Nothing
' ----------------------------------------------------------------------
'||
'||
'||
'||
'||
'||
'VV
'выполнение запроса из файла,
'сюда передаем текст файла полностью, внутри происходит парсинг
function executeTextQuerys(textFromFile)
	'для этого получаем запросы разделенные точкой запятой
	Dim massiveQuerys : massiveQuerys = split(textFromFile, ";") 
	for index = lbound(massiveQuerys) to ubound(massiveQuerys)
		
		dim strForExecute : strForExecute = trim(massiveQuerys(index))
		
		call log("[выполняем]: " & strForExecute)
		On Error Resume Next
		
		if (strForExecute <> "") then
			ADO.execute	strForExecute
		end if
		
		'обработка ошибок, чтоб в случае ошибки, мы могли понять в чем дело
		Select Case Err.Number
			Case 0 'Все в порядке
				call log("[выполнили]")
			Case Else
				call log("[ошибка]: " & Err.Number & " - " & Err.Description)
				Err.Clear
				executeTextQuerys = false
				exit function
		End Select	
		On Error GoTo 0
	next
	executeTextQuerys = true
end function

'заполняем словарь подключениями
sub fillDicConnections()
	Dim arrayConnections : arrayConnections = split(connections, CONST_SEPARATOR_LIST)
	Dim c
	for c = lbound(arrayConnections) to ubound(arrayConnections)
		Dim strConnection : strConnection = split(arrayConnections(c), CONST_SEPARATOR_PROP)
		dicConnections.add strConnection(0), strConnection(1)
	next
end sub

'заполняем словарь файлами со скриптами
sub fillDicFiles()
	Dim arrayConnections : arrayConnections = split(filesExecute, CONST_SEPARATOR_LIST)
	Dim c
	for c = lbound(arrayConnections) to ubound(arrayConnections)
		dicFiles.add dicFiles.count, arrayConnections(c)
	next
end sub

sub log(textLog)
	Dim strLog : strLog = now() & " " & textLog 	
	Wscript.echo strLog
	objFile.WriteLine strLog
end sub
