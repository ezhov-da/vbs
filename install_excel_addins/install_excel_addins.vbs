'===================================
'����������� ���� �.�. 
'GitHub: https://github.com/ezhov-da
'===================================

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'��������� ����� ������ ���� � ANSI
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim objArgs
Set objArgs = WScript.Arguments

dim pathToAddins
pathToAddins = objArgs(0)
WScript.Echo "���� � ����������: " & pathToAddins

dim nameAddins
nameAddins = objArgs(1)
WScript.Echo "�������� ����������: " & nameAddins

dim extensionAddins
extensionAddins = objArgs(2)
WScript.Echo "���������� ����������: " & extensionAddins

Dim source
source = pathToAddins & nameAddins & "." & extensionAddins
WScript.Echo "����������� ����������: " & source

Dim questionCloseExcel
questionCloseExcel = MsgBox ("����� ���������� ����������, ���������� ������� Excel." & chr(10) & "Excel ������?", vbYesNo, "�������� Excel")
Select Case questionCloseExcel
Case vbYes

	dim excel
	set excel  = CreateObject("Excel.Application")    	

	Dim targetFolder
	targetFolder = excel.Application.UserLibraryPath
	WScript.Echo "����� �������� ���������: " & targetFolder

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
		MsgBox "������: " & Err.Description
	else
		dim result
		result = "���������� [" & source & "] �����������."
		WScript.Echo result
		MsgBox result
	end if
	
Case vbNo

    MsgBox "��������� ���������� ��������." & chr(10) & "�������� Excel � ��������� ��������� ���������� ��������." & chr(10) & "�������."
	
End Select


