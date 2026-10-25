Attribute VB_Name = "modRunLog"
Option Explicit

' ============================================================================
' modRunLog: оновлення книги з журналом запусків (етап D; v4.13)
' Потребує modEngineCommon.
'
' RefreshWithRunLog:
'   0. якщо pq_QualitySummary не завантажено на аркуш, пропонує завантажити його на аркуш «Якість»;
'   1. синхронно оновлює всі запити («Оновити все») і вимірює тривалість;
'   2. для кожного рядка pq_QualitySummary дописує рядок у таблицю tblRunLog
'      (аркуш «RunLog», створюється при першому запуску): номер і час запуску, тривалість,
'      помилку оновлення, SourceID, призначення, період, Status, QualityFlag, Message, кількість рядків
'      і помилок, шлях до файла, дату зміни й розмір файла на момент запуску, користувача;
'   3. показує підсумок QualityFlag.
' Історія накопичується; старі рядки можна видаляти вручну. Самого Power Query для історії
' недостатньо: результати запитів перезаписуються при кожному оновленні.
' ============================================================================

Private Const LOG_TABLE As String = "tblRunLog"
Private Const LOG_SHEET As String = "RunLog"
Private Const SUMMARY_QUERY As String = "pq_QualitySummary"
Private Const QUALITY_SHEET As String = "Якість"

Public Sub RefreshWithRunLog()
    Dim note As String, t0 As Double, dur As Double, refreshErr As String, runAt As Date
    Dim qs As ListObject, reg As ListObject, lg As ListObject, i As Long, runId As Long, n As Long
    Dim sid As String, rr As Long, filePath As String, flag As String
    Dim nOK As Long, nWarn As Long, nFail As Long, failIds As String

    note = engLoadSettings()
    If note <> "" Then
        If MsgBox(note & vbCrLf & vbCrLf & "Продовжити?", vbYesNo + vbExclamation, "Журнал запусків") = vbNo Then Exit Sub
    End If

    On Error GoTo Fail
    ' Зведення якості має бути на аркуші ДО оновлення: інакше після «Оновити все» журнал
    ' не буде звідки взяти (у тестовій книзі pq_QualitySummary - лише підключення)
    Set qs = engFindQueryTable(SUMMARY_QUERY)
    If qs Is Nothing Then
        If MsgBox("Запит " & SUMMARY_QUERY & " не завантажено на аркуш, а журнал будується з нього." & _
                  vbCrLf & vbCrLf & "Завантажити його на аркуш «" & QUALITY_SHEET & "» і продовжити?" & _
                  vbCrLf & "(одноразово; таблиця не оновлюється з «Оновити все» - її оновлює цей макрос)", _
                  vbYesNo + vbQuestion, "Журнал запусків") = vbNo Then Exit Sub
        Set qs = engLoadQueryToSheet(SUMMARY_QUERY, EnsureSheet(QUALITY_SHEET), "A3", False)
    End If

    ' Звіти якості з v4.13 вимкнено з «Оновити все» (оновлюються на вимогу). Позначку підключення
    ' макрос не змінює: після «Оновити все» він сам оновлює зведення, якщо воно вимкнене.

    runAt = Now
    t0 = Timer
    Application.Cursor = xlWait
    Application.StatusBar = "Оновлення всіх запитів (синхронно; може тривати кілька хвилин)..."
    refreshErr = engRefreshAllSync()
    dur = Timer - t0
    If dur < 0 Then dur = dur + 86400      ' оновлення через північ
    Application.StatusBar = False
    Application.Cursor = xlDefault

    Set qs = engFindQueryTable(SUMMARY_QUERY)       ' після оновлення таблицю могли перебудувати
    If qs Is Nothing Then
        MsgBox "Оновлення завершено за " & Format$(dur, "0.0") & " с, але таблицю " & SUMMARY_QUERY & _
               " на аркуші не знайдено: журнал не записано.", vbExclamation, "Журнал запусків"
        Exit Sub
    End If

    ' Зведення вимкнено з «Оновити все» - оновлюємо його окремо (одне читання файлів-джерел)
    Dim sumErr As String
    Application.StatusBar = "Оновлення " & SUMMARY_QUERY & "..."
    sumErr = engRefreshIfExcluded(SUMMARY_QUERY)
    Application.StatusBar = False
    dur = Timer - t0
    If dur < 0 Then dur = dur + 86400
    If sumErr <> "" Then refreshErr = refreshErr & IIf(refreshErr = "", "", "; ") & sumErr
    Set qs = engFindQueryTable(SUMMARY_QUERY)
    Set reg = engFindTable(engTableName("Registry"))
    Set lg = EnsureLogTable()
    runId = NextRunId(lg)

    Application.ScreenUpdating = False
    If Not qs.DataBodyRange Is Nothing Then
        For i = 1 To qs.ListRows.Count
            sid = Trim$(CStr(engCell(qs, i, "SourceID")))
            flag = UCase$(Trim$(CStr(engCellOpt(qs, i, "QualityFlag", ""))))
            filePath = ""
            rr = 0
            If Not reg Is Nothing Then rr = RegistryRow(reg, sid)
            If rr > 0 Then filePath = engSourcePath(reg, rr)

            AppendLog lg, runId, runAt, dur, refreshErr, qs, i, filePath
            n = n + 1

            Select Case flag
                Case "OK": nOK = nOK + 1
                Case "WARN": nWarn = nWarn + 1
                Case "FAIL", "BAD_THRESHOLD"
                    nFail = nFail + 1
                    If nFail <= 10 Then failIds = failIds & vbCrLf & "  * " & sid & " (" & flag & ")"
            End Select
        Next i
    End If
    Application.ScreenUpdating = True

    MsgBox "Запуск № " & runId & ": оновлення " & Format$(dur, "0.0") & " с" & _
           IIf(refreshErr <> "", vbCrLf & "Помилка оновлення: " & refreshErr, "") & vbCrLf & vbCrLf & _
           "Джерел у журналі: " & n & " (OK " & nOK & ", WARN " & nWarn & ", FAIL " & nFail & ")" & _
           IIf(failIds <> "", vbCrLf & vbCrLf & "Завантаження зупиниться для:" & failIds, "") & vbCrLf & vbCrLf & _
           "Журнал: аркуш «" & lg.Parent.Name & "», таблиця " & LOG_TABLE, _
           IIf(nFail > 0 Or refreshErr <> "", vbExclamation, vbInformation), "Журнал запусків"
    Exit Sub
Fail:
    Dim errMsg As String: errMsg = Err.Description
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    MsgBox "Журнал запусків перервано: " & errMsg, vbCritical, "Журнал запусків"
End Sub

' Рядок журналу для i-го рядка зведення
Private Sub AppendLog(lg As ListObject, runId As Long, runAt As Date, dur As Double, refreshErr As String, _
                      qs As ListObject, i As Long, filePath As String)
    Dim lr As ListRow, fm As Variant, fs As Variant
    fm = Empty: fs = Empty
    If filePath <> "" Then
        If engFileExists(filePath) Then
            On Error Resume Next
            fm = FileDateTime(filePath)
            fs = Round(FileLen(filePath) / 1024, 1)
            On Error GoTo 0
        End If
    End If

    Set lr = lg.ListRows.Add
    SetLogCell lg, lr, "RunID", runId
    SetLogCell lg, lr, "RunAt", runAt
    SetLogCell lg, lr, "DurationSec", Round(dur, 1)
    SetLogCell lg, lr, "RefreshError", IIf(refreshErr = "", Empty, refreshErr)
    SetLogCell lg, lr, "SourceID", engCellOpt(qs, i, "SourceID", Empty)
    SetLogCell lg, lr, "TargetType", engCellOpt(qs, i, "TargetType", Empty)
    SetLogCell lg, lr, "PeriodType", engCellOpt(qs, i, "PeriodType", Empty)
    SetLogCell lg, lr, "PeriodNumber", engCellOpt(qs, i, "PeriodNumber", Empty)
    SetLogCell lg, lr, "Status", engCellOpt(qs, i, "Status", Empty)
    SetLogCell lg, lr, "QualityFlag", engCellOpt(qs, i, "QualityFlag", Empty)
    SetLogCell lg, lr, "Message", engCellOpt(qs, i, "Message", Empty)
    SetLogCell lg, lr, "Rows", engCellOpt(qs, i, "Rows", Empty)
    SetLogCell lg, lr, "ErrorRows", engCellOpt(qs, i, "ErrorRows", Empty)
    SetLogCell lg, lr, "RowsDropped", engCellOpt(qs, i, "RowsDropped", Empty)
    SetLogCell lg, lr, "FilePath", IIf(filePath = "", Empty, filePath)
    SetLogCell lg, lr, "FileModified", fm
    SetLogCell lg, lr, "FileSizeKB", fs
    SetLogCell lg, lr, "User", Application.UserName
End Sub

' Колонки, яких немає в таблиці журналу (наприклад, її змінили вручну), пропускаються
Private Sub SetLogCell(lg As ListObject, lr As ListRow, col As String, v As Variant)
    Dim c As Range
    If Not engColumnExists(lg, col) Then Exit Sub
    Set c = lr.Range(1, lg.ListColumns(col).Index)
    c.value = v
    If VarType(v) = vbDate Then c.NumberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Private Function EnsureSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

' Таблиця журналу; створюється на аркуші RunLog при першому запуску
Private Function EnsureLogTable() As ListObject
    Dim lo As ListObject, ws As Worksheet, hdr As Variant
    Set lo = engFindTable(LOG_TABLE)
    If Not lo Is Nothing Then Set EnsureLogTable = lo: Exit Function

    hdr = Array("RunID", "RunAt", "DurationSec", "RefreshError", "SourceID", "TargetType", "PeriodType", _
                "PeriodNumber", "Status", "QualityFlag", "Message", "Rows", "ErrorRows", "RowsDropped", _
                "FilePath", "FileModified", "FileSizeKB", "User")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = LOG_SHEET
    End If
    ' шапка в рядку 3: рядок 1 - підпис таблиці (UpdateHelp), рядок 2 порожній (hlpProject)
    ws.Range("A3").Resize(1, UBound(hdr) + 1).value = hdr
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A3").Resize(2, UBound(hdr) + 1), , xlYes)
    lo.Name = LOG_TABLE
    ' таблиця створена з одним порожнім рядком даних: прибираємо його
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    Set EnsureLogTable = lo
End Function

Private Function NextRunId(lg As ListObject) As Long
    Dim c As Range, m As Long
    If Not lg.DataBodyRange Is Nothing Then
        For Each c In lg.ListColumns("RunID").DataBodyRange.Cells
            If IsNumeric(c.value) And Not IsEmpty(c.value) Then
                If CLng(c.value) > m Then m = CLng(c.value)
            End If
        Next c
    End If
    NextRunId = m + 1
End Function

' Номер рядка реєстру з указаним SourceID (0 = не знайдено)
Private Function RegistryRow(reg As ListObject, sid As String) As Long
    Dim i As Long
    If reg.DataBodyRange Is Nothing Or sid = "" Then Exit Function
    For i = 1 To reg.ListRows.Count
        If StrComp(Trim$(CStr(engCell(reg, i, "SourceID"))), sid, vbTextCompare) = 0 Then
            RegistryRow = i
            Exit Function
        End If
    Next i
End Function
