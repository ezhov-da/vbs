Option Explicit

Public Sub checkVersion()
    Dim v_version  As Boolean
    v_version = False
    
    Dim name As String
    name = getNameXLAM()
    
    Dim v_addin
    
    For Each v_addin In Application.AddIns
      If LCase(v_addin.Title) = LCase(name) Then
         v_version = True
         'v_addin.Installed = True
      End If
    Next
    
    Dim otvet
    Dim PROCED
    
    'Если установлена не та версия:
    If Not v_version Then
        otvet = MsgBox("Текущая версия инструмента устарела. Произвести обновление? Приложение будет закрыто.", vbYesNo, "Обновление")
            If otvet = 6 Then
                Dim script As String: script = """X:\Категорийные\Кожанов\reinstall_addins.vbs"" "
                Dim nameAddinForDelete As String: nameAddinForDelete = """INACTIVE"" "
                Dim nameNewAddins As String: nameNewAddins = """" & name & """ "
                Dim extNewAddins As String: extNewAddins = """xlam"" "
                Dim fromNewAddins As String: fromNewAddins = """X:\Категорийные\Кожанов\"""
                Dim scriptCommand As String
                scriptCommand = "Wscript.exe  " & script & nameAddinForDelete & nameNewAddins & extNewAddins & fromNewAddins
            
                PROCED = Shell(scriptCommand, vbNormalFocus)
                Application.DisplayAlerts = False
                On Error Resume Next
                Application.Quit
                End
             End If
    End If
End Sub

Private Function getNameXLAM() As String
    Const PATH_TO_FILE_VERSION As String = "X:\Категорийные\Кожанов\e_inactive_version_name.txt"
    
    Dim myFile As String
    myFile = PATH_TO_FILE_VERSION
    Open myFile For Input As #1
    
    Dim text As String
    Dim textline
    
    Do Until EOF(1)
        Line Input #1, textline
        text = text & textline
    Loop
    
    Close #1
    getNameXLAM = text
End Function

