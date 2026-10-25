Attribute VB_Name = "modImportPowerQueryFromFolder"
Option Explicit

' ============================================================================
' modImportPowerQueryFromFolder: імпорт або оновлення запитів Power Query з файлів .pq (v4.11)
' Самодостатній (без modEngineCommon): ним збирають книгу з нуля.
'
' ImportPQQueriesFromFolder: ім'я файла = ім'я запиту, вміст - код M у UTF-8.
' Шукає .pq у вибраній теці; якщо там їх немає - у відомих підтеках: powerquery\queries
' (Git-репозиторій), engine\powerquery і for_workbook\powerquery (пакет). Запити пишуться
' в книгу з цим модулем (ThisWorkbook), а не в активну.
' ============================================================================

Public Sub ImportPQQueriesFromFolder()
    Dim fd As FileDialog, fso As Object, file As Object, pqQuery As WorkbookQuery, stream As Object
    Dim folderPath As String, queryName As String, mCode As String, sub_ As Variant
    Dim countImported As Long, countUpdated As Long, oldScreen As Boolean

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "Оберіть теку з файлами .pq (або корінь репозиторію чи пакета)"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Sub
    folderPath = fd.SelectedItems(1)
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If CountPq(fso, folderPath) = 0 Then
        For Each sub_ In Array("powerquery\queries\", "engine\powerquery\", "for_workbook\powerquery\")
            If CountPq(fso, folderPath & sub_) > 0 Then folderPath = folderPath & sub_: Exit For
        Next sub_
    End If
    If CountPq(fso, folderPath) = 0 Then
        MsgBox "У вибраній теці і в підтеках powerquery\queries, engine\powerquery, for_workbook\powerquery " & _
               "немає файлів .pq.", vbExclamation, "Power Query Import"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Fail

    For Each file In fso.GetFolder(folderPath).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "pq" Then
            queryName = fso.GetBaseName(file.Name)

            Set stream = CreateObject("ADODB.Stream")
            stream.Type = 2                                  ' adTypeText
            stream.charset = "utf-8"
            stream.Open
            stream.LoadFromFile file.path
            mCode = stream.ReadText
            stream.Close

            Set pqQuery = Nothing
            On Error Resume Next
            Set pqQuery = ThisWorkbook.Queries(queryName)
            On Error GoTo Fail

            If pqQuery Is Nothing Then
                ThisWorkbook.Queries.Add Name:=queryName, formula:=mCode
                countImported = countImported + 1
            Else
                pqQuery.formula = mCode
                countUpdated = countUpdated + 1
            End If
        End If
    Next file

    Application.ScreenUpdating = oldScreen
    MsgBox "Імпорт завершено." & vbCrLf & vbCrLf & _
           "Тека: " & folderPath & vbCrLf & _
           "• Створено нових запитів: " & countImported & vbCrLf & _
           "• Оновлено існуючих: " & countUpdated, _
           vbInformation, "Power Query Import"
    Exit Sub
Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "Імпорт перервано на «" & queryName & "»: " & Err.Description & vbCrLf & _
           "Встигли: створено " & countImported & ", оновлено " & countUpdated & ".", _
           vbCritical, "Power Query Import"
End Sub

Private Function CountPq(fso As Object, dirPath As String) As Long
    Dim f As Object
    If Not fso.FolderExists(dirPath) Then Exit Function
    For Each f In fso.GetFolder(dirPath).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "pq" Then CountPq = CountPq + 1
    Next f
End Function
