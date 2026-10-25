Attribute VB_Name = "modCleanExportSelectedSheets"
Option Explicit

' Потребує modEngineCommon: імена файлів очищує engSafeFileName (етап D, аудит D13).

Sub CleanExportSelectedSheets()

    Dim wbSrc       As Workbook    ' < вихідна книга
    Dim ws          As Worksheet
    Dim wbNew       As Workbook
    Dim wsName      As String
    Dim arrSheets() As String
    Dim i           As Long
    Dim exportCount As Long
    Dim extraText   As String
    Dim exportPath  As String
    Dim fd          As FileDialog
    Dim lo          As ListObject

    Set wbSrc = ActiveWorkbook     ' < запам'ятовуємо до будь-яких дій

    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' -- 1. Запит суфіксу ---------------------------------------------------
    extraText = InputBox("Введіть текст, який буде додано до назви кожного файлу:" & vbLf & _
                         "(залиште порожнім, якщо не потрібно)", "Суфікс для файлів")

    If StrPtr(extraText) = 0 Then
        MsgBox "Експорт скасовано.", vbInformation
        GoTo CleanUp
    End If

    extraText = engSafeFileName(Trim(extraText))
    If Len(extraText) > 0 Then extraText = " " & extraText

    ' -- 2. Вибір папки експорту --------------------------------------------
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .title = "Оберіть папку для збереження файлів"
        .InitialFileName = wbSrc.path & "\"
        .AllowMultiSelect = False
        If .Show = False Then
            MsgBox "Експорт скасовано.", vbInformation
            GoTo CleanUp
        End If
        exportPath = .SelectedItems(1)
    End With

    ' -- 3. Зберігаємо імена виділених аркушів -----------------------------
    ReDim arrSheets(1 To wbSrc.Windows(1).SelectedSheets.Count)
    i = 1
    For Each ws In wbSrc.Windows(1).SelectedSheets
        arrSheets(i) = ws.Name
        i = i + 1
    Next ws

    ' Знімаємо групування без ризику помилки
    wbSrc.Activate
    wbSrc.Sheets(arrSheets(1)).Activate

    ' -- 4. Основний цикл ---------------------------------------------------
    exportCount = 0

    For i = LBound(arrSheets) To UBound(arrSheets)

        Set ws = wbSrc.Sheets(arrSheets(i))      ' < wbSrc
        wsName = ws.Name

        Set wbNew = Workbooks.Add
        ws.Copy Before:=wbNew.Sheets(1)

        Do While wbNew.Sheets.Count > 1
            wbNew.Sheets(wbNew.Sheets.Count).Delete
        Loop

        Do While wbNew.Queries.Count > 0
            wbNew.Queries(1).Delete
        Loop

        Do While wbNew.Connections.Count > 0
            wbNew.Connections(1).Delete
        Loop

        ' Таблиці, завантажені з запитів, лишаються прив'язаними до QueryTable і дають
        ' попередження при відкритті файла. QueryTable.Delete від'єднує таблицю від запиту,
        ' а сама таблиця (ListObject) з даними й форматом лишається (перенесено з modExport).
        On Error Resume Next
        For Each lo In wbNew.Sheets(1).ListObjects
            If lo.SourceType = xlSrcQuery Then lo.QueryTable.Delete
        Next lo
        On Error GoTo CleanUp

        wbNew.SaveAs fileName:=UniqueFilePath(exportPath, wsName & extraText), _
                     FileFormat:=xlOpenXMLWorkbook, _
                     CreateBackup:=False
        wbNew.Close SaveChanges:=False

        exportCount = exportCount + 1

    Next i

CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If Err.Number <> 0 Then
        MsgBox "Помилка: " & Err.Description, vbCritical
    ElseIf exportCount > 0 Then
        MsgBox "Готово! Експортовано аркушів: " & exportCount & vbLf & _
               "Папка: " & exportPath, vbInformation
    End If

End Sub

Function UniqueFilePath(folder As String, BaseName As String) As String
    Dim fullPath As String
    Dim counter  As Long

    fullPath = folder & "\" & BaseName & ".xlsx"
    counter = 1

    Do While Dir(fullPath) <> ""
        fullPath = folder & "\" & BaseName & "_" & counter & ".xlsx"
        counter = counter + 1
    Loop

    UniqueFilePath = fullPath
End Function
