'===================================
'разработчик Ежов Д.А. 
'GitHub: https://github.com/ezhov-da
'===================================

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'КОДИРОВКА ФАЙЛА ДОЛЖНА БЫТЬ В ANSI
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Итак, данный скрипт отключает и удаляет все надстройки, которые подходят под название
'Копирует новую надстройку из указанного места
'И подключает ее
'
'Входные параметры:
'1. Название надстройки, которую нужно отключить/удалить: EXAMPLE
'2. Название надстройки для подключения: EXAMPLE_0.2
'3. Расширение надстройки: xlam
'4. Путь откуда брать новую надстройку: C:\test\

Option Explicit
'on error resume next
'получаем название надстройки, которую необходимо обновить

dim objArgs
Set objArgs = WScript.Arguments
dim nameXLAMForDelete
nameXLAMForDelete = objArgs(0)
'WScript.Echo "Название надстройки для удаления: " & nameXLAMForDelete

dim nameAddinsForInstall
nameAddinsForInstall = objArgs(1)
'WScript.Echo "Название надстройки для установки: " & nameAddinsForInstall

dim extAddinsForInstall
extAddinsForInstall = objArgs(2)
'WScript.Echo "Расширение надстройки для установки: " & extAddinsForInstall

dim pathFromCopyNewAddins
pathFromCopyNewAddins = objArgs(3)
'WScript.Echo "Путь откуда брать новую надстройку: " & pathFromCopyNewAddins

Dim excel
set excel  = CreateObject("Excel.Application")    	

Dim targetFolder
targetFolder = excel.Application.UserLibraryPath
'WScript.Echo "Место хранения надстроек: " & targetFolder

Dim oFSO: Set oFSO = CreateObject("Scripting.FileSystemObject")

'WScript.Echo "Запустили отключение надстроек"
call reconnectAddins(nameXLAMForDelete)
'WScript.Echo "Отключение надстроек завершено"

'закрываем excel
excel.Quit
set excel = nothing

'Удаляем надстройки
'WScript.Echo "Запустили удаление надстроек"
call delFileOnMask(targetFolder, nameXLAMForDelete)
'WScript.Echo "Удаление надстроек завершено"

Dim nameNewAddinsAndExtension
nameNewAddinsAndExtension = nameAddinsForInstall & "." & extAddinsForInstall
'WScript.Echo "Название новой надстройки с расширением: " & nameNewAddinsAndExtension

Dim fillPathToNewAddins
fillPathToNewAddins = pathFromCopyNewAddins & nameNewAddinsAndExtension
'WScript.Echo "Полный путь к новой надстройке: " & fillPathToNewAddins

oFSO.CopyFile fillPathToNewAddins, targetFolder
'WScript.Echo "Скопировали надстройку в место хранения"

'создаем новый экземпляр excel для подключения
set excel  = CreateObject("Excel.Application")  

'WScript.Echo "Подключаем надстройку..."
excel.Application.EnableEvents = False
excel.Application.AddIns(nameAddinsForInstall).Installed = True
excel.Application.EnableEvents = False
'WScript.Echo "Надстройка подключена"

excel.Quit
set excel = nothing
set oFSO = nothing

if Err.Number = 0 then
	MsgBox "Надстройка обновлена."
else
	MsgBox "Ошибка обновления: " & Err.Description
end if

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

  Set objRegExp = CreateObject("VBScript.RegExp")
  objRegExp.Pattern = sMask
  Dim oFld, arrMask, v, i
  Set oFld = oFSO.GetFolder(s)
  arrMask = Split(LCase(sMask), " ")
  For Each v In oFld.Files
    For i = LBound(arrMask) To UBound(arrMask)
      Set objMatches = objRegExp.Execute(v.name)
      If objMatches.Count > 0 Then
		'WScript.Echo "Надстройка удалена: "& v.name
        v.Delete
        Exit For
      End If
    Next
  Next
End Function

'отключаем надстройки, которые совпали
Function reconnectAddins(nameAddinForReconnect)
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
			'WScript.Echo "Надстройка отключена: "& addin.name
		End If		
    Next
End Function
