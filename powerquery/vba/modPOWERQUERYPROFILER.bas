Attribute VB_Name = "modPOWERQUERYPROFILER"
Option Explicit

' ============================================================================
' modPOWERQUERYPROFILER: час оновлення кожного підключення Power Query (етап D; v4.12)
' Потребує modEngineCommon.
'
' ProfilePowerQueries послідовно оновлює OLEDB-підключення книги, які щось завантажують
' (таблиця на аркуші або модель даних), і записує тривалість кожного на аркуш PQ_Profiler.
' Підсумок запуску - окремим рядком на аркуші PQ_ProfilerSummary.
'   * ім'я запиту береться з параметра Location рядка підключення, тож воно правильне і в
'     англійському («Query - X»), і в українському («Запит – X») Excel;
'   * сортування: новіші запуски зверху, у межах запуску - найдовші запити зверху;
'   * під час заміру обчислення вручну, фоновий режим підключень потім відновлюється;
'   * необов'язковий прогрівальний прохід; колонка Notes: «після прогріву» і джерело живлення
'     (мережа / батарея: на батареї ноутбук сповільнює обчислення, заміри непорівнянні);
'   * колонка Seq (v4.12): порядковий номер оновлення в запуску - видно, чи збігся стрибок
'     тривалості з тим, що оновлювалося перед ним.
' Правила порівняння (architecture, 6.12): лише від мережі в режимі найкращої продуктивності;
' розкид між запусками +-20-30%, тому варіанти порівнювати в одному запуску (два запити поруч)
' або за медіаною 3+ запусків.
'
' Важливо для читання результатів: кожне підключення оновлюється окремо, тому його час
' включає обчислення всіх запитів, від яких воно залежить (кешу між оновленнями немає).
' Сума тривалостей тому більша за час «Оновити все».
' ============================================================================

Private Type SYSTEM_POWER_STATUS
    ACLineStatus As Byte
    BatteryFlag As Byte
    BatteryLifePercent As Byte
    SystemStatusFlag As Byte
    BatteryLifeTime As Long
    BatteryFullLifeTime As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function GetSystemPowerStatus Lib "kernel32" (ByRef lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
Private Declare Function GetSystemPowerStatus Lib "kernel32" (ByRef lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

Private Const DATA_SHEET As String = "PQ_Profiler"
Private Const SUMMARY_SHEET As String = "PQ_ProfilerSummary"

' Джерело живлення: "мережа", "батарея" або "невідомо"
Private Function PowerSource() As String
    Dim ps As SYSTEM_POWER_STATUS
    PowerSource = "невідомо"
    If GetSystemPowerStatus(ps) = 0 Then Exit Function
    Select Case ps.ACLineStatus
        Case 1: PowerSource = "мережа"
        Case 0: PowerSource = "батарея"
    End Select
End Function

Private Function GetPreciseTime() As Double
    Static frequency As Double
    Dim counter As Currency, freqCurrency As Currency
    If frequency = 0 Then
        QueryPerformanceFrequency freqCurrency
        frequency = CDbl(freqCurrency)
    End If
    QueryPerformanceCounter counter
    GetPreciseTime = CDbl(counter) / frequency
End Function

Public Sub ProfilePowerQueries()
    Dim wb As Workbook, conn As WorkbookConnection, ws As Worksheet
    Dim elapsed As Double, runNumber As Long, nextRow As Long, startDateTime As Date
    Dim qName As String, cName As String, status As String, errorText As String, notes As String
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim processed As Long, total As Long, errCount As Long, sumTime As Double
    Dim slowName As String, slowTime As Double

    Set wb = ThisWorkbook
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts

    On Error GoTo FatalError
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set ws = GetDataSheet(wb)
    runNumber = GetNextRunNumber(ws)

    ' Підключення, які щось завантажують (таблиця на аркуші або модель даних). «Лише підключення»
    ' під час Refresh не обчислюються, їхні 0,01-0,2 с нічого не вимірюють (аудит, v4.10)
    Dim measured As New Collection, skipped As Long, warm As VbMsgBoxResult, c As Variant, pwr As String
    For Each conn In wb.Connections
        If IsOLEDBConnection(conn) Then
            If engFindQueryTable(QueryNameOf(conn)) Is Nothing And Not InModel(conn) Then
                skipped = skipped + 1
            Else
                measured.Add conn
            End If
        End If
    Next conn
    total = measured.Count
    If total = 0 Then
        MsgBox "У книзі не знайдено підключень, що завантажують дані. Профілювати нічого.", vbExclamation, "Power Query Profiler"
        GoTo CleanExit
    End If

    pwr = PowerSource()                  ' один раз для вікна запуску; у рядках - на момент заміру
    warm = MsgBox("Підключень для заміру: " & total & " (лише підключення пропущено: " & skipped & ")." & vbCrLf & vbCrLf & _
                  "Зробити прогрівальний прохід без запису?" & vbCrLf & _
                  "Перше оновлення платить за запуск Power Query і читання файлів з диска; прогрів робить" & vbCrLf & _
                  "заміри порівнянними, але подвоює час." & vbCrLf & vbCrLf & _
                  "Живлення: " & pwr & IIf(pwr = "батарея", _
                  " - на батареї заміри непорівнянні; підключіть мережу і виберіть режим найкращої продуктивності.", ""), _
                  vbYesNoCancel + IIf(pwr = "батарея", vbExclamation, vbQuestion), "Power Query Profiler")
    If warm = vbCancel Then GoTo CleanExit

    ' ручний режим обчислення: інакше кожне завантаження таблиці запускає перерахунок книги,
    ' і його час потрапляє в замір (аудит: чисті тести в пам'яті займали 7-9 с)
    Application.Calculation = xlCalculationManual

    If warm = vbYes Then
        processed = 0
        For Each c In measured
            processed = processed + 1
            Application.StatusBar = "Power Query Profiler: прогрів " & processed & "/" & total
            DoEvents
            Set conn = c
            RefreshTimed conn, status, errorText
        Next c
    End If

    startDateTime = Now
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    processed = 0
    For Each c In measured
        Set conn = c
        processed = processed + 1
        cName = conn.Name
        qName = QueryNameOf(conn)
        notes = IIf(warm = vbYes, "після прогріву; ", "") & "живлення: " & PowerSource()

        Application.StatusBar = "Power Query Profiler: " & processed & "/" & total & " - " & qName
        DoEvents

        elapsed = RefreshTimed(conn, status, errorText)
        If status = "ERROR" Then errCount = errCount + 1

        WriteResult ws, nextRow, runNumber, startDateTime, qName, cName, elapsed, status, errorText, notes, processed
        nextRow = nextRow + 1
        sumTime = sumTime + elapsed
        If elapsed > slowTime Then slowTime = elapsed: slowName = qName
    Next c

    FormatDataSheet ws
    SortDataSheet ws
    WriteSummary wb, runNumber, startDateTime, processed, errCount, sumTime, slowName, slowTime

    Application.StatusBar = False
    MsgBox "Профілювання завершено." & vbCrLf & vbCrLf & _
           "Запуск № " & runNumber & vbCrLf & _
           "Підключень: " & processed & IIf(errCount > 0, " (помилок: " & errCount & ")", "") & _
           " (лише підключення пропущено: " & skipped & ")" & vbCrLf & _
           "Обчислення під час заміру: вручну" & IIf(warm = vbYes, "; з прогрівом", "") & vbCrLf & _
           "Сума тривалостей: " & Format$(sumTime, "0.0") & " с" & vbCrLf & _
           "Найдовший: " & slowName & " (" & Format$(slowTime, "0.0") & " с)" & vbCrLf & _
           vbCrLf & _
           "Аркуші: " & DATA_SHEET & ", " & SUMMARY_SHEET, _
           IIf(errCount > 0, vbExclamation, vbInformation), "Power Query Profiler"
    GoTo CleanExit

FatalError:
    Application.StatusBar = False
    MsgBox "Профілювання перервано." & vbCrLf & vbCrLf & "Помилка " & Err.Number & ":" & vbCrLf & _
           Err.Description, vbCritical, "Power Query Profiler"

CleanExit:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = False
End Sub

' ================== ДОПОМІЖНІ ==================
' Синхронне оновлення одного підключення з виміром часу. Фоновий режим вимикається лише на час
' оновлення і потім відновлюється (раніше лишався вимкненим назавжди; аудит, v4.10)
Private Function RefreshTimed(ByVal conn As WorkbookConnection, ByRef status As String, ByRef errorText As String) As Double
    Dim bg As Boolean, hasBg As Boolean, t0 As Double
    On Error Resume Next
    bg = conn.OLEDBConnection.BackgroundQuery
    hasBg = (Err.Number = 0)
    Err.Clear
    If hasBg Then conn.OLEDBConnection.BackgroundQuery = False
    Err.Clear
    t0 = GetPreciseTime()
    conn.Refresh
    RefreshTimed = GetPreciseTime() - t0
    If Err.Number <> 0 Then
        status = "ERROR"
        errorText = Err.Number & ": " & Err.Description
    Else
        status = "OK"
        errorText = ""
    End If
    Err.Clear
    If hasBg Then conn.OLEDBConnection.BackgroundQuery = bg
    Err.Clear
    On Error GoTo 0
End Function


Private Function IsOLEDBConnection(ByVal conn As WorkbookConnection) As Boolean
    On Error Resume Next
    IsOLEDBConnection = (conn.Type = xlConnectionTypeOLEDB)
    On Error GoTo 0
End Function

Private Function InModel(ByVal conn As WorkbookConnection) As Boolean
    On Error Resume Next
    InModel = conn.InModel
    On Error GoTo 0
End Function

' Ім'я запиту: Location рядка підключення; якщо його немає - назва підключення без префікса
Private Function QueryNameOf(ByVal conn As WorkbookConnection) As String
    Dim s As String, p As Long
    s = engQueryNameOf(conn)
    If s = "" Then
        s = conn.Name
        ' «Query - X» або «Запит – X» (середнє тире; Excel ставить перед ним нерозривний пробіл)
        s = Replace(s, ChrW(160), " ")
        p = InStr(s, " - ")
        If p = 0 Then p = InStr(s, " " & ChrW(8211) & " ")
        If p > 0 Then s = Mid$(s, p + 3)
    End If
    QueryNameOf = s
End Function

Private Function GetDataSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = DATA_SHEET
        ws.Range("A1:I1").value = Array("Run", "Date/Time", "Query", "Connection", "Duration (sec)", _
                                        "Status", "Error", "Duration (min)", "Notes")
    End If
    If IsEmpty(ws.Range("J1").value) Then ws.Range("J1").value = "Seq"     ' з v4.12; старі рядки без номера
    Set GetDataSheet = ws
End Function

Private Function GetNextRunNumber(ByVal ws As Worksheet) As Long
    Dim lastRow As Long, i As Long, m As Long, v As Variant
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        v = ws.Cells(i, 1).value
        If Not IsError(v) And Not IsEmpty(v) Then
            If IsNumeric(v) Then If CLng(v) > m Then m = CLng(v)
        End If
    Next i
    GetNextRunNumber = m + 1
End Function

Private Sub WriteResult(ByVal ws As Worksheet, ByVal r As Long, ByVal runNumber As Long, _
                        ByVal startDateTime As Date, ByVal qName As String, ByVal cName As String, _
                        ByVal elapsed As Double, ByVal status As String, ByVal errorText As String, _
                        ByVal notes As String, ByVal seq As Long)
    ws.Cells(r, 1).value = runNumber
    ws.Cells(r, 2).value = startDateTime
    ws.Cells(r, 3).value = qName
    ws.Cells(r, 4).value = cName
    ws.Cells(r, 5).value = elapsed
    ws.Cells(r, 6).value = status
    ws.Cells(r, 7).value = errorText
    ws.Cells(r, 8).value = elapsed / 60
    ws.Cells(r, 9).value = notes
    ws.Cells(r, 10).value = seq
End Sub

Private Sub FormatDataSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Rows(1).Font.Bold = True
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:J" & Application.Max(lastRow, 2)).AutoFilter
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 35
    ws.Columns("D").ColumnWidth = 45
    ws.Columns("E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G").ColumnWidth = 50
    ws.Columns("H").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 32
    ws.Columns("J").ColumnWidth = 6
    If lastRow >= 2 Then
        ws.Range("B2:B" & lastRow).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        ws.Range("E2:E" & lastRow).NumberFormat = "0.000"
        ws.Range("H2:H" & lastRow).NumberFormat = "0.000"
    End If
End Sub

' Новіші запуски зверху; у межах запуску - найдовші запити зверху
Private Sub SortDataSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 2 Then Exit Sub
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending
        .SortFields.Add key:=ws.Range("E2:E" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending
        .SetRange ws.Range("A1:J" & lastRow)
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

' Один рядок на запуск на окремому аркуші
Private Sub WriteSummary(ByVal wb As Workbook, ByVal runNumber As Long, ByVal startDateTime As Date, _
                         ByVal n As Long, ByVal errCount As Long, ByVal sumTime As Double, _
                         ByVal slowName As String, ByVal slowTime As Double)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = wb.Worksheets(SUMMARY_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SUMMARY_SHEET
        ws.Range("A1:H1").value = Array("Run", "Date/Time", "Connections", "Errors", _
                                        "Total (sec)", "Total (min)", "Slowest query", "Slowest (sec)")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A").ColumnWidth = 8
        ws.Columns("B").ColumnWidth = 20
        ws.Columns("C:F").ColumnWidth = 12
        ws.Columns("G").ColumnWidth = 35
        ws.Columns("H").ColumnWidth = 14
    End If
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).value = runNumber
    ws.Cells(r, 2).value = startDateTime
    ws.Cells(r, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(r, 3).value = n
    ws.Cells(r, 4).value = errCount
    ws.Cells(r, 5).value = sumTime
    ws.Cells(r, 5).NumberFormat = "0.000"
    ws.Cells(r, 6).value = sumTime / 60
    ws.Cells(r, 6).NumberFormat = "0.000"
    ws.Cells(r, 7).value = slowName
    ws.Cells(r, 8).value = slowTime
    ws.Cells(r, 8).NumberFormat = "0.000"
End Sub
