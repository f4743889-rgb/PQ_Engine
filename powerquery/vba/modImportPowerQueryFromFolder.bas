Attribute VB_Name = "modImportPowerQueryFromFolder"
Option Explicit

Sub ImportPQQueriesFromFolder()
    Dim fd As FileDialog
    Dim folderPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim queryName As String
    Dim mCode As String
    Dim pqQuery As WorkbookQuery
    Dim stream As Object
    Dim countImported As Long
    Dim countUpdated As Long
    
    ' 1. Діалог вибору папки з файлами .pq
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "Оберіть папку з файлами .pq"
    fd.AllowMultiSelect = False
    
    If fd.Show <> -1 Then Exit Sub
    folderPath = fd.SelectedItems(1)
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    Application.ScreenUpdating = False
    
    countImported = 0
    countUpdated = 0
    
    ' 2. Перебір усіх файлів у папці
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "pq" Then
            ' Отримуємо ім'я файлу без розширення (це буде назва запиту)
            queryName = fso.GetBaseName(file.Name)
            
            ' Зчитуємо вміст .pq файлу в кодуванні UTF-8 (щоб зберегти кирилицю)
            Set stream = CreateObject("ADODB.Stream")
            stream.Type = 2 ' adTypeText
            stream.charset = "utf-8"
            stream.Open
            stream.LoadFromFile file.path
            mCode = stream.ReadText
            stream.Close
            
            ' 3. Перевірка: чи існує вже запит з такою назвою
            Set pqQuery = Nothing
            On Error Resume Next
            Set pqQuery = ActiveWorkbook.Queries(queryName)
            On Error GoTo 0
            
            If pqQuery Is Nothing Then
                ' Створюємо новий запит
                ActiveWorkbook.Queries.Add Name:=queryName, formula:=mCode
                countImported = countImported + 1
            Else
                ' Якщо запит існує — оновлюємо його M-код
                pqQuery.formula = mCode
                countUpdated = countUpdated + 1
            End If
        End If
    Next file
    
    Application.ScreenUpdating = True
    
    ' 4. Звіт про виконання
    MsgBox "Імпорт успішно завершено!" & vbCrLf & vbCrLf & _
           "• Створено нових запитів: " & countImported & vbCrLf & _
           "• Оновлено існуючих: " & countUpdated, _
           vbInformation, "Power Query Import"
End Sub




