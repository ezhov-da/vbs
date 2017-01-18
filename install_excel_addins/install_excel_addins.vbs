'===================================
'разработчик Ежов Д.А. 
'GitHub: https://github.com/ezhov-da
'===================================

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'КОДИРОВКА ФАЙЛА ДОЛЖНА БЫТЬ В ANSI
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim objArgs
Set objArgs = WScript.Arguments

dim pathToAddins
pathToAddins = objArgs(0)
WScript.Echo "Путь к надстройке: " & pathToAddins

dim nameAddins
nameAddins = objArgs(1)
WScript.Echo "Название надстройки: " & nameAddins

dim extensionAddins
extensionAddins = objArgs(2)
WScript.Echo "Расширение надстройки: " & extensionAddins

Dim source
source = pathToAddins & nameAddins & "." & extensionAddins
WScript.Echo "Копирование надстройки: " & source

Dim questionCloseExcel
questionCloseExcel = MsgBox ("Перед установкой надстройки, необходимо закрыть Excel." & chr(10) & "Excel закрыт?", vbYesNo, "Закрытие Excel")
Select Case questionCloseExcel
Case vbYes

	dim excel
	set excel  = CreateObject("Excel.Application")    	

	Dim targetFolder
	targetFolder = excel.Application.UserLibraryPath
	WScript.Echo "Место хранения надстроек: " & targetFolder

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.CopyFile source, targetFolder

	on error resume next
	excel.Application.EnableEvents = False
	excel.Application.AddIns(nameAddins).Installed = True
	excel.Application.EnableEvents = False
	excel.Quit
	set excel = nothing

	if Err > 0 then
		MsgBox "Ошибка : " & Err.Description
	else
		dim result
		result = "Надстройка [" & source & "] установлена."
		WScript.Echo result
		MsgBox result
	end if
	
Case vbNo

    MsgBox "Установка надстройки отменена." & chr(10) & "Закройте Excel и запустите установку надстройки повторно." & chr(10) & "Спасибо."
	
End Select


