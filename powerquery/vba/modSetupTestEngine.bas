Attribute VB_Name = "modSetupTestEngine"
Option Explicit

' ============================================================================
' modSetupTestEngine: перетворення КОПІЇ книги Engine на тестову (етап C)
'
' Запуск: імпортуйте модуль у тестову книгу, запустіть SetupTestEngine і виберіть
' розпаковану теку пакета (у ній мають бути теки engine і test_engine).
'
' Що робить (повторний запуск безпечний, дані тестів просто перезаписуються):
'   1. додає тестові призначення в tblTargetType (TestCsv, TestXlsx, TestNoSchema);
'   2. замінює рядки tblSourceRegistry, tblSchemaRegistry, tblRowFilters рядками
'      з test_config.xlsx і очищає tblDefaultHeaders; Path = тека fixtures пакета;
'      відсутні колонки реєстру (наприклад, ErrorThreshold) додає;
'   3. створює або оновлює tblTestExpected і tblTestLoad;
'   4. імпортує всі запити з engine\powerquery (Engine v4.5 і тести);
'   5. завантажує test_All, test_Summary, test_Load на аркуш «Тести» і оновлює їх.
' НЕ запускайте в робочій книзі: реєстр і схема стануть тестовими.
' ============================================================================

Private Const CFG_FILE   As String = "test_engine\test_config.xlsx"
Private Const FIX_DIR    As String = "test_engine\fixtures"
Private Const PQ_DIR     As String = "engine\powerquery"
Private Const TEST_SHEET As String = "Тести"
Private Const NOTE_COL   As String = "Опис (не копіювати)"

Public Sub SetupTestEngine()
    Dim root As String, fd As FileDialog, fso As Object, wbCfg As Workbook
    Dim report As String, nNew As Long, nUpd As Long

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "Виберіть розпаковану теку пакета (з теками engine і test_engine)"
    If fd.Show <> -1 Then Exit Sub
    root = fd.SelectedItems(1)
    If Right$(root, 1) <> "\" Then root = root & "\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(root & CFG_FILE) Or Not fso.FolderExists(root & FIX_DIR) _
       Or Not fso.FolderExists(root & PQ_DIR) Then
        MsgBox "У вибраній теці немає " & CFG_FILE & ", " & FIX_DIR & " або " & PQ_DIR & ".", _
               vbCritical, "Тестовий Engine"
        Exit Sub
    End If

    If MsgBox("Книгу «" & ThisWorkbook.Name & "» буде перетворено на тестову:" & vbCrLf & vbCrLf & _
              "• оновлено всі запити Power Query з " & PQ_DIR & ";" & vbCrLf & _
              "• замінено рядки tblSourceRegistry, tblSchemaRegistry, tblRowFilters, очищено tblDefaultHeaders;" & vbCrLf & _
              "• додано тестові призначення в tblTargetType і таблиці tblTestExpected, tblTestLoad;" & vbCrLf & _
              "• на аркуш «" & TEST_SHEET & "» завантажено test_All, test_Summary, test_Load." & vbCrLf & vbCrLf & _
              "Тестові файли: " & root & FIX_DIR & vbCrLf & _
              "(не переміщуйте цю теку після налаштування)" & vbCrLf & vbCrLf & _
              "НЕ запускайте це в робочій книзі. Продовжити?", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Тестовий Engine") <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wbCfg = Workbooks.Open(root & CFG_FILE, ReadOnly:=True, UpdateLinks:=False)

    report = report & "tblTargetType: додано " & _
          AddMissingValues(NeedTable("tblTargetType"), CfgRange(wbCfg, "Призначення"), "TargetType") & vbCrLf
    report = report & "tblSourceRegistry: " & _
          ReplaceRows(NeedTable("tblSourceRegistry"), CfgRange(wbCfg, "Реєстр"), True, root & FIX_DIR) & " рядків" & vbCrLf
    report = report & "tblSchemaRegistry: " & _
          ReplaceRows(NeedTable("tblSchemaRegistry"), CfgRange(wbCfg, "Схема"), False) & " рядків" & vbCrLf
    report = report & "tblRowFilters: " & _
          ReplaceRows(NeedTable("tblRowFilters"), CfgRange(wbCfg, "Фільтри"), False) & " рядків" & vbCrLf
    ClearRows FindTable("tblDefaultHeaders")
    report = report & "tblDefaultHeaders: очищено" & vbCrLf
    report = report & EnsureTestTable("tblTestExpected", wbCfg.Worksheets("Очікування")) & vbCrLf
    report = report & EnsureTestTable("tblTestLoad", wbCfg.Worksheets("Навантаження")) & vbCrLf

    wbCfg.Close SaveChanges:=False
    Set wbCfg = Nothing

    ImportQueries root & PQ_DIR, nNew, nUpd
    report = report & "Запити Power Query: додано " & nNew & ", оновлено " & nUpd & vbCrLf

    Application.ScreenUpdating = True
    Application.StatusBar = "Завантаження тестових запитів..."
    Dim ws As Worksheet
    Set ws = EnsureSheet(TEST_SHEET)
    report = report & LoadQueryToSheet("test_All", ws, "A1") & vbCrLf
    report = report & LoadQueryToSheet("test_Summary", ws, "H1") & vbCrLf
    report = report & LoadQueryToSheet("test_Load", ws, "O1") & vbCrLf
    Application.StatusBar = False
    Application.EnableEvents = True

    MsgBox "Тестовий Engine налаштовано." & vbCrLf & vbCrLf & report & vbCrLf & _
           "Далі: Дані > Оновити все, потім аркуш «" & TEST_SHEET & "»: у test_All має бути Result = PASS." & vbCrLf & _
           "Збережіть книгу.", vbInformation, "Тестовий Engine"
    Exit Sub

Fail:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    If Not wbCfg Is Nothing Then wbCfg.Close SaveChanges:=False
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Налаштування перервано: " & msg & vbCrLf & vbCrLf & "Виконано до помилки:" & vbCrLf & report, _
           vbCritical, "Тестовий Engine"
End Sub

' ================== ТАБЛИЦІ ==================
Private Function CfgRange(wb As Workbook, sheetName As String) As Range
    Set CfgRange = wb.Worksheets(sheetName).Range("A1").CurrentRegion
End Function

Private Function NeedTable(tblName As String) As ListObject
    Set NeedTable = FindTable(tblName)
    If NeedTable Is Nothing Then Err.Raise vbObjectError + 513, , "У книзі немає таблиці " & tblName
End Function

' Замінює рядки таблиці рядками діапазону src (перший рядок src = заголовки).
' Колонки зіставляються за назвою; колонки src, яких немає в таблиці, пропускаються
' (або додаються, якщо addCols). pathValue, якщо задано, записується в колонку Path.
Private Function ReplaceRows(lo As ListObject, src As Range, addCols As Boolean, _
                             Optional pathValue As String = "") As Long
    Dim c As Long, r As Long, h As String, lr As ListRow
    If addCols Then
        For c = 1 To src.Columns.Count
            h = Trim$(CStr(src.Cells(1, c).value))
            If h <> "" And h <> NOTE_COL And Not ColumnExists(lo, h) Then lo.ListColumns.Add.Name = h
        Next c
    End If
    ClearRows lo
    For r = 2 To src.Rows.Count
        Set lr = lo.ListRows.Add
        For c = 1 To src.Columns.Count
            h = Trim$(CStr(src.Cells(1, c).value))
            If h <> "" And ColumnExists(lo, h) Then
                lr.Range(1, lo.ListColumns(h).Index).value = src.Cells(r, c).value
            End If
        Next c
        If pathValue <> "" And ColumnExists(lo, "Path") Then
            lr.Range(1, lo.ListColumns("Path").Index).value = pathValue
        End If
    Next r
    ReplaceRows = src.Rows.Count - 1
End Function

Private Sub ClearRows(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
End Sub

' Додає рядки src, значення keyCol яких ще немає в таблиці (без урахування регістру)
Private Function AddMissingValues(lo As ListObject, src As Range, keyCol As String) As Long
    Dim r As Long, c As Long, k As Long, h As String, v As String, lr As ListRow, found As Boolean
    For c = 1 To src.Columns.Count
        If Trim$(CStr(src.Cells(1, c).value)) = keyCol Then k = c
    Next c
    If k = 0 Then Exit Function
    For r = 2 To src.Rows.Count
        v = Trim$(CStr(src.Cells(r, k).value))
        found = False
        If Not lo.DataBodyRange Is Nothing Then
            Dim cell As Range
            For Each cell In lo.ListColumns(keyCol).DataBodyRange.Cells
                If StrComp(Trim$(CStr(cell.value)), v, vbTextCompare) = 0 Then found = True: Exit For
            Next cell
        End If
        If Not found And v <> "" Then
            Set lr = lo.ListRows.Add
            For c = 1 To src.Columns.Count
                h = Trim$(CStr(src.Cells(1, c).value))
                If h <> "" And ColumnExists(lo, h) Then lr.Range(1, lo.ListColumns(h).Index).value = src.Cells(r, c).value
            Next c
            AddMissingValues = AddMissingValues + 1
        End If
    Next r
End Function

' Таблиця очікувань: копія аркуша з конфігурації, якщо таблиці ще немає, інакше заміна рядків
Private Function EnsureTestTable(tblName As String, wsCfg As Worksheet) As String
    Dim lo As ListObject
    Set lo = FindTable(tblName)
    If lo Is Nothing Then
        wsCfg.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        EnsureTestTable = tblName & ": аркуш «" & wsCfg.Name & "» скопійовано"
    Else
        EnsureTestTable = tblName & ": " & ReplaceRows(lo, wsCfg.Range("A1").CurrentRegion, True) & " рядків"
    End If
End Function

' ================== ЗАПИТИ ==================
Private Sub ImportQueries(dirPath As String, ByRef nNew As Long, ByRef nUpd As Long)
    Dim fso As Object, f As Object, nm As String, code As String, q As WorkbookQuery
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(dirPath).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "pq" Then
            nm = fso.GetBaseName(f.Name)
            code = ReadUtf8(CStr(f.path))
            Set q = Nothing
            On Error Resume Next
            Set q = ThisWorkbook.Queries(nm)
            On Error GoTo 0
            If q Is Nothing Then
                ThisWorkbook.Queries.Add Name:=nm, formula:=code
                nNew = nNew + 1
            Else
                q.formula = code
                nUpd = nUpd + 1
            End If
        End If
    Next f
End Sub

' Завантаження запиту на аркуш таблицею (якщо ще не завантажений) і оновлення
Private Function LoadQueryToSheet(queryName As String, ws As Worksheet, topLeft As String) As String
    Dim lo As ListObject
    Set lo = FindQueryTable(queryName)
    On Error GoTo Bad
    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(SourceType:=0, _
            Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & _
                    ";Extended Properties=""""", Destination:=ws.Range(topLeft))
        With lo.QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [" & queryName & "]")
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = False
            .RefreshStyle = xlInsertDeleteCells
            .SaveData = True
            .AdjustColumnWidth = True
            .PreserveColumnInfo = True
        End With
        On Error Resume Next
        lo.Name = "tbl_" & queryName
        On Error GoTo Bad
    End If
    lo.QueryTable.Refresh BackgroundQuery:=False
    LoadQueryToSheet = queryName & ": завантажено на «" & ws.Name & "»"
    Exit Function
Bad:
    LoadQueryToSheet = queryName & ": ПОМИЛКА завантаження (" & Err.Description & "); оновіть вручну"
End Function

' ================== ДОПОМІЖНІ ==================
Private Function EnsureSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        EnsureSheet.Name = sheetName
    End If
End Function

Private Function FindTable(tblName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tblName, vbTextCompare) = 0 Then Set FindTable = lo: Exit Function
        Next lo
    Next ws
End Function

Private Function FindQueryTable(queryName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject, cmd As String
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            cmd = ""
            On Error Resume Next
            If lo.SourceType = xlSrcQuery Then cmd = CStr(lo.QueryTable.CommandText)
            On Error GoTo 0
            If InStr(1, cmd, "[" & queryName & "]", vbTextCompare) > 0 Then Set FindQueryTable = lo: Exit Function
        Next lo
    Next ws
End Function

Private Function ColumnExists(lo As ListObject, colName As String) As Boolean
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then ColumnExists = True: Exit Function
    Next lc
End Function

Private Function ReadUtf8(path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function
