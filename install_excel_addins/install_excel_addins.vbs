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

dim isShowMsg
isShowMsg = objArgs(3)
WScript.Echo "Показывать сообщения: " & isShowMsg

dim nameAddinsWithoutVersion
nameAddinsWithoutVersion = objArgs(4)
WScript.Echo "Название надстройки без версии для отключения старых надстроек: " & nameAddinsWithoutVersion

Dim source
source = pathToAddins & nameAddins & "." & extensionAddins
WScript.Echo "Копирование надстройки: " & source

Dim questionCloseExcel
if (isShowMsg = "1") then
	questionCloseExcel = MsgBox ("Перед установкой надстройки, необходимо закрыть Excel." & chr(10) & "Excel закрыт?", vbYesNo, "Закрытие Excel")
else
	questionCloseExcel = vbYes
end if

Select Case questionCloseExcel
Case vbYes

	dim excel
	set excel  = CreateObject("Excel.Application")    	

	Dim targetFolder
	targetFolder = excel.Application.UserLibraryPath
	WScript.Echo "Место хранения надстроек: " & targetFolder

	call reconnectAddins(nameAddinsWithoutVersion, excel)
	
	call delFileOnMask(targetFolder, nameAddinsWithoutVersion)
	
	excel.Quit
	set excel = nothing
	
	set excel  = CreateObject("Excel.Application")  
	
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	WScript.Echo "Копирование файла..."
	
	fso.CopyFile source, targetFolder

	on error resume next
	
	excel.Application.EnableEvents = False
	
	excel.Application.AddIns(nameAddins).Installed = True
	
	WScript.Echo "Надстройка подключена..."
	
	excel.Application.EnableEvents = False
	
	excel.Quit
	
	set excel = nothing
	
	WScript.Echo "Обнуление EXCEL..."	

	if Err > 0 then
		if (isShowMsg = "1") then
			MsgBox "Ошибка : " & Err.Description
		end if
	else
		dim result
		result = "Надстройка [" & source & "] установлена."
		WScript.Echo result
		if (isShowMsg = "1") then
			MsgBox result
		end if
	end if
	
Case vbNo
	if (isShowMsg = "1") then
		MsgBox "Установка надстройки отменена." & chr(10) & "Закройте Excel и запустите установку надстройки повторно." & chr(10) & "Спасибо."
	end if
End Select

'==================================================================================================================================
'Переработали функцию, столкнулись с тем, 
'что не удаляли старые надстройки из-за того, 
'что не находили в папке нужные имена,
'а не находили из-за того, что имя постоянно менялось (дата и версия)'
'В связи с этим перешли на поиск по регулярке, 
'то есть ищем конкретный инструмент, и удаляем файлы по нему
Function delFileOnMask(s, sMask)
	dim objRegExp
	dim objMatches
	dim counter
	dim objMatch

	dim  oFSO
	Set oFSO = CreateObject("Scripting.FileSystemObject")

    Dim col'коллекция для удаления файлов
    Set col = CreateObject("System.Collections.ArrayList")
	
	Set objRegExp = CreateObject("VBScript.RegExp")
	objRegExp.Pattern = sMask
	
	Dim oFld, arrMask, v, i
	Set oFld = oFSO.GetFolder(s)
	
	For Each v In oFld.Files
	  Set objMatches = objRegExp.Execute(v.name)
	  If objMatches.Count > 0 Then
		col.add v
	  End If
	Next
	
    For Each i In col
		WScript.Echo "Надстройка удалена: " & i.name	
		i.Delete
	Next
	
	WScript.Echo "Удаление надстроек завершено..."
End Function

'отключаем надстройки, которые совпали
Function reconnectAddins(nameAddinForReconnect, excel)
  dim objRegExp
  dim objMatches
  dim counter
  dim objMatch

  Set objRegExp = CreateObject("VBScript.RegExp")
  objRegExp.Pattern = nameAddinForReconnect

    Dim c
	Dim addin
    For c = 1 To excel.Application.AddIns.count
		set addin = excel.Application.AddIns.Item(c)
		Set objMatches = objRegExp.Execute(addin.name)
		If objMatches.Count > 0 Then
			addin.Installed = False
			WScript.Echo "Надстройка отключена: "& addin.name
		End If		
    Next
End Function
