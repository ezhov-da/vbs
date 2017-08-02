Option Explicit


Private Const PATH_TO_REINSTALL_VBS As String = "Категорийные\Кожанов\reinstall_addins.vbs"
Private Const PATH_TO_FOLDER_WITH_ADDINS As String = "Категорийные\Кожанов\"

Private Const WORK_PATH As String = "Категорийные\Кожанов\"

Private Const TXT_FILE_VERSION As String = "e_wassort_version_name_new.txt"

Private Const NAME_ADDIN_FOR_DELETE_WITHOUT_VERSION As String = "PROBLEM-MINQUANTITY"

Public Sub checkVersion()
    Dim v_version  As Boolean
    v_version = False
    
    
    Dim path As String: path = getPath()
    Dim name As String
    name = getNameXLAM(path)
    
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
                Dim pathToReinstallVbs As String: pathToReinstallVbs = path & PATH_TO_REINSTALL_VBS
                Dim pathToFolderVersion As String: pathToFolderVersion = path & PATH_TO_FOLDER_WITH_ADDINS
            
                Dim script As String: script = """" & pathToReinstallVbs & """ "
                Dim nameAddinForDelete As String: nameAddinForDelete = """" & NAME_ADDIN_FOR_DELETE_WITHOUT_VERSION & """ "
                Dim nameNewAddins As String: nameNewAddins = """" & name & """ "
                Dim extNewAddins As String: extNewAddins = """xlam"" "
                Dim fromNewAddins As String: fromNewAddins = """" & pathToFolderVersion & """"
                Dim scriptCommand As String
                scriptCommand = _
                    "Wscript.exe  " & _
                    script & _
                    nameAddinForDelete & _
                    nameNewAddins & _
                    extNewAddins & _
                    fromNewAddins
            
                'Debug.Print scriptCommand
            
                PROCED = Shell(scriptCommand, vbNormalFocus)
                Application.DisplayAlerts = False
                On Error Resume Next
                Application.Quit
                End
             End If
    End If
End Sub

Private Function getNameXLAM(path As String) As String
    Dim pathToFileVersion As String
    pathToFileVersion = path & WORK_PATH & TXT_FILE_VERSION
     
    Dim myFile As String
    myFile = pathToFileVersion
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

Private Function getPath() As String
    'пути для получения файлов
    Dim pathArr: pathArr = Array("W:\_Departments\", "X:\", "R:\")

    Dim find As String

    Dim i
    For i = LBound(pathArr) To UBound(pathArr)
    
        On Error Resume Next
    
        'прибавляем папку, так как именно к ней может не быть доступа
        find = Dir(pathArr(i) & "Категорийные\", vbDirectory)
               
        'а раз нет доступа, значит ошибка
        'ее и обрабатываем
        If (err.number = 0 And find <> "") Then
            getPath = pathArr(i)
            Exit Function
        End If
    Next i
End Function

