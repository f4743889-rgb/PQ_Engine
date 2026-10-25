Attribute VB_Name = "modSelectSourceFiles"
Option Explicit

' ============================================================================
' modSelectSourceFiles: вибір файлів-джерел для рядків tblSourceRegistry
'
' ПЕРЕДУМОВА: у tblSourceRegistry уже є стільки рядків, скільки файлів потрібно
'   для роботи: SourceID, TargetType, PeriodType, PeriodNumber, IsRequired,
'   Enabled заповнені. Макрос рядків не додає й не видаляє.
'
' Що робить:
'   * проходить активні рядки (Enabled порожнє або TRUE; FALSE пропускається);
'   * для кожного пропонує вибрати файл (розширення з cfgSettings[Extensions]);
'   * для CSV перевіряє наявність обов'язкових колонок схеми (tblSchemaRegistry,
'     IsRequired = TRUE, Use <> FALSE) з тим самим роздільником і кодуванням,
'     що й Engine (BOM файла, колонки Delimiter/Encoding реєстру, інакше cfgSettings);
'   * показує підсумок; таблиця змінюється ЛИШЕ після підтвердження;
'   * записує повний шлях у FullPath і очищає Path та FileName цього рядка;
'   * пропонує оновити запити й показує підсумок QualityFlag з pq_QualitySummary.
'
' Потребує modEngineCommon (етап D: спільні процедури замість власних копій).
'
' PeriodType і PeriodNumber макрос не змінює: перевірте їх навпроти кожного файлу.
' Ручне редагування FullPath (або Path + FileName) у таблиці так само допустиме.
' ============================================================================

' Імена таблиць, параметри CSV, логічні значення й розширення файлів беруться з cfgSettings
' через modEngineCommon (engTableName, engSetting, engToBool, engIsExtension).
Private Const SUMMARY_QUERY  As String = "pq_QualitySummary"
Private Const SUMMARY_MAX_ROWS As Long = 15      ' MsgBox вміщує близько 1000 символів

Private Enum PickResult
    prOK = 0
    prSkipped = 1
    prAbort = 2
End Enum

' ================== ПУБЛІЧНИЙ ВХІД ==================
Public Sub SelectSourceFiles()
    Dim note As String
    note = engLoadSettings()
    If note <> "" Then
        If MsgBox(note & vbCrLf & vbCrLf & "Роздільник """ & CfgDelimiter() & """, кодування " & CfgEncoding() & "." & _
                  vbCrLf & "Продовжити?", vbYesNo + vbExclamation, ENG_SETTINGS_QUERY) = vbNo Then Exit Sub
    End If

    Dim lo As ListObject
    Set lo = engFindTable(RegistryTable())
    If lo Is Nothing Then
        MsgBox "Таблицю " & RegistryTable() & " не знайдено.", vbCritical
        Exit Sub
    End If
    If Not HasColumns(lo, Array("SourceID", "TargetType", "FullPath", "Path", "FileName", _
                                "PeriodType", "PeriodNumber", "IsRequired", "Enabled")) Then Exit Sub

    Dim rowsIdx As Collection
    Set rowsIdx = ActiveRows(lo)
    If rowsIdx.Count = 0 Then
        MsgBox "У " & RegistryTable() & " немає активних рядків." & vbCrLf & vbCrLf & _
               "Спершу додайте рядки призначень (SourceID, TargetType, період, " & _
               "IsRequired, Enabled), стільки, скільки файлів потрібно, і запустіть макрос знову.", _
               vbExclamation, "Немає рядків"
        Exit Sub
    End If

    Dim n As Long: n = rowsIdx.Count
    Dim picks() As String, states() As String
    Dim reselectAll As Boolean, k As Long, ans As VbMsgBoxResult

    On Error GoTo Fail

    Do
        ReDim picks(1 To n): ReDim states(1 To n)

        reselectAll = True
        If AnyValidPath(lo, rowsIdx) Then
            ans = MsgBox("Для частини рядків файли вже задано." & vbCrLf & vbCrLf & _
                         "[Так]  вибрати файли заново для всіх активних рядків" & vbCrLf & _
                         "[Ні]  вибрати лише для рядків без дійсного файла" & vbCrLf & _
                         "[Скасувати]  вийти без змін", _
                         vbYesNoCancel + vbQuestion, "Вибір файлів")
            If ans = vbCancel Then Exit Sub
            reselectAll = (ans = vbYes)
        End If

        For k = 1 To n
            Dim r As Long, cur As String, res As PickResult, newPath As String
            r = rowsIdx(k)
            cur = engSourcePath(lo, r)

            If Not reselectAll And cur <> "" And engFileExists(cur) Then
                picks(k) = cur: states(k) = "без змін"
            Else
                Application.StatusBar = "Вибір файлу " & k & " з " & n
                res = PickFile(lo, r, k, n, cur, newPath)
                Select Case res
                    Case prAbort
                        Application.StatusBar = False
                        MsgBox "Вибір скасовано. Таблицю не змінено.", vbInformation, "Скасовано"
                        Exit Sub
                    Case prSkipped
                        picks(k) = cur
                        states(k) = IIf(cur = "", "не вибрано", "пропущено, лишився попередній")
                    Case prOK
                        picks(k) = newPath
                        states(k) = IIf(LCase$(newPath) = LCase$(cur), "без змін", "новий")
                End Select
            End If
        Next k
        Application.StatusBar = False

        ans = MsgBox(BuildSummary(lo, rowsIdx, picks, states) & vbCrLf & vbCrLf & _
                     "[Так]  записати вибір у " & RegistryTable() & vbCrLf & _
                     "[Ні]  вибрати заново" & vbCrLf & _
                     "[Скасувати]  вийти без змін", _
                     vbYesNoCancel + vbQuestion, "Підсумок вибору")
        If ans = vbCancel Then
            MsgBox "Таблицю не змінено.", vbInformation, "Скасовано"
            Exit Sub
        End If
    Loop While ans = vbNo

    WritePicks lo, rowsIdx, picks

    If MsgBox("Шляхи записано у " & RegistryTable() & "." & vbCrLf & vbCrLf & _
              "Перевірте PeriodType і PeriodNumber навпроти кожного файлу: макрос їх не змінює." & _
              vbCrLf & vbCrLf & "Оновити запити зараз?", _
              vbYesNo + vbQuestion, "Готово") = vbYes Then
        RefreshAndReport
    End If
    Exit Sub

Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Помилка макросу: " & Err.Description, vbCritical
End Sub

' ================== ВИБІР І ПЕРЕВІРКА ОДНОГО ФАЙЛУ ==================
Private Function PickFile(lo As ListObject, r As Long, k As Long, n As Long, _
                          cur As String, ByRef newPath As String) As PickResult
    Dim req As Boolean, dlgTitle As String, f As Variant, missing As String, ans As VbMsgBoxResult
    req = engToBool(engCell(lo, r, "IsRequired"), True)

    dlgTitle = "Файл " & k & " з " & n & ": " & RowLabel(lo, r) & IIf(req, " [обов'язковий]", "")
    If cur <> "" And Not engFileExists(cur) Then dlgTitle = dlgTitle & " (збережений шлях недійсний)"

    GoToFolderOf cur

    Do
        f = Application.GetOpenFilename(FileFilter(), 1, dlgTitle)

        If VarType(f) = vbBoolean Then          ' діалог скасовано
            If Not req Then
                PickFile = prSkipped: Exit Function
            End If
            ans = MsgBox(RowLabel(lo, r) & " є обов'язковим." & vbCrLf & vbCrLf & _
                         "[Повторити]  вибрати файл" & vbCrLf & _
                         "[Скасувати]  вийти без змін", _
                         vbRetryCancel + vbExclamation, "Обов'язковий файл")
            If ans = vbCancel Then PickFile = prAbort: Exit Function
        ElseIf Not engIsExtension(CStr(f)) Then
            MsgBox "Формат не підтримується Engine: " & CStr(f) & vbCrLf & _
                   "Дозволено (cfgSettings[Extensions]): " & AllowedList() & ".", vbExclamation
        Else
            missing = MissingRequiredColumns(lo, r, CStr(f))
            If missing = "" Then
                newPath = CStr(f): PickFile = prOK: Exit Function
            End If
            ans = MsgBox("Файл: " & FileNameOf(CStr(f)) & vbCrLf & _
                         "Рядок: " & RowLabel(lo, r) & vbCrLf & vbCrLf & missing & vbCrLf & vbCrLf & _
                         "[Так]  вибрати інший файл" & vbCrLf & _
                         "[Ні]  залишити цей файл (Engine покаже помилку у звітах)" & vbCrLf & _
                         "[Скасувати]  вийти без змін", _
                         vbYesNoCancel + vbExclamation, "Структура файлу")
            If ans = vbNo Then newPath = CStr(f): PickFile = prOK: Exit Function
            If ans = vbCancel Then PickFile = prAbort: Exit Function
        End If
    Loop
End Function

' Проблеми заголовка CSV: відсутні обов'язкові колонки схеми і повторювані назви.
' "" = усе гаразд або перевірка неможлива.
Private Function MissingRequiredColumns(lo As ListObject, r As Long, filePath As String) As String
    If Not engIsExtension(filePath, "Csv") Then Exit Function   ' Excel перевіряє Engine після оновлення

    Dim reqCols As Collection
    Set reqCols = RequiredSchemaColumns(Trim$(CStr(engCell(lo, r, "TargetType"))))
    If reqCols.Count = 0 Then Exit Function

    Dim delim As String, enc As Long, hdrLine As String, ok As Boolean
    delim = Trim$(CStr(engCellOpt(lo, r, "Delimiter", "")))
    If delim = "" Then delim = CfgDelimiter()
    enc = CfgEncoding()
    If IsNumeric(engCellOpt(lo, r, "Encoding", "")) Then enc = CLng(engCellOpt(lo, r, "Encoding", ""))

    Dim cs As String
    cs = DetectBom(filePath)                               ' BOM має пріоритет, як у Power Query
    If cs = "" Then cs = CharsetOf(enc)
    hdrLine = ReadFirstLine(filePath, cs, ok)
    If Not ok Then
        MissingRequiredColumns = "Файл не вдалося прочитати (зайнятий іншою програмою або пошкоджений)."
        Exit Function
    End If

    Dim miss As String, nCols As Long, dupl As String
    miss = MissingList(hdrLine, delim, reqCols, nCols)
    dupl = DuplicateNames(hdrLine, delim)
    If miss = "" And dupl = "" Then Exit Function

    Const DUP_NOTE As String = "Engine візьме першу колонку з такою назвою, а повтори отримають суфікс _1, _2."
    If miss = "" Then
        MissingRequiredColumns = "Повторювані назви колонок у заголовку:" & dupl & vbCrLf & DUP_NOTE
        Exit Function
    End If

    ' діагностика: інше кодування, інший роздільник, прочитаний заголовок
    Dim hint As String, alt As String, altHeader As String, ok2 As Boolean, dummy As Long
    alt = IIf(cs = "utf-8", "windows-1251", "utf-8")
    altHeader = ReadFirstLine(filePath, alt, ok2)
    If ok2 Then
        If MissingList(altHeader, delim, reqCols, dummy) = "" Then
            hint = vbCrLf & "З кодуванням " & alt & " усі колонки знайдено: задайте Encoding = " & _
                   IIf(alt = "utf-8", "65001", "1251") & " у рядку реєстру або CsvEncoding у cfgSettings."
        End If
    End If
    If hint = "" And nCols <= 1 Then
        hint = vbCrLf & "У заголовку знайдено лише одну колонку: ймовірно, інший роздільник " & _
               "(задайте Delimiter у рядку реєстру)."
    End If

    MissingRequiredColumns = "Відсутні обов'язкові колонки (роздільник """ & _
        IIf(delim = vbTab, "Tab", delim) & """, кодування " & cs & "):" & miss & vbCrLf & hint & _
        vbCrLf & "Прочитаний заголовок: " & Left$(hdrLine, 150) & IIf(Len(hdrLine) > 150, "...", "")
    If dupl <> "" Then
        MissingRequiredColumns = MissingRequiredColumns & vbCrLf & vbCrLf & _
            "Повторювані назви колонок:" & dupl & vbCrLf & DUP_NOTE
    End If
End Function

' Перелік обов'язкових колонок, яких немає в заголовку; nCols = кількість колонок заголовка
Private Function MissingList(hdrLine As String, delim As String, reqCols As Collection, _
                             ByRef nCols As Long) As String
    Dim cols As Collection, c As Variant, miss As String
    Set cols = SplitCsvLine(hdrLine, delim)
    nCols = cols.Count
    For Each c In reqCols
        If Not InCollection(CStr(c), cols) Then miss = miss & vbCrLf & "  * " & CStr(c)
    Next c
    MissingList = miss
End Function

' Назви, що трапляються в заголовку кілька разів (точний збіг, як у Engine)
Private Function DuplicateNames(hdrLine As String, delim As String) As String
    Dim cnt As Object, c As Variant, k As Variant, res As String
    Set cnt = CreateObject("Scripting.Dictionary")
    cnt.CompareMode = vbBinaryCompare
    For Each c In SplitCsvLine(hdrLine, delim)
        If CStr(c) <> "" Then cnt(CStr(c)) = cnt(CStr(c)) + 1
    Next c
    For Each k In cnt.Keys
        If cnt(k) > 1 Then res = res & vbCrLf & "  * " & CStr(k) & " (" & cnt(k) & " рази)"
    Next k
    DuplicateNames = res
End Function

' Кодування за BOM: "utf-8", "unicode" (UTF-16 LE) або "" (BOM немає)
Private Function DetectBom(filePath As String) As String
    Dim st As Object, b() As Byte, n As Long
    On Error GoTo Done
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.LoadFromFile filePath
    n = st.Size
    If n > 3 Then n = 3
    If n >= 2 Then b = st.Read(n)
    st.Close
    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then DetectBom = "utf-8": Exit Function
    End If
    If n >= 2 Then
        If b(0) = &HFF And b(1) = &HFE Then DetectBom = "unicode": Exit Function
    End If
Done:
End Function

Private Function RequiredSchemaColumns(targetType As String) As Collection
    Dim res As New Collection, lo As ListObject, i As Long
    Set RequiredSchemaColumns = res
    Set lo = engFindTable(engTableName("Schema"))
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Not HasColumns(lo, Array("TargetType", "SourceColumnName", "IsRequired"), True) Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(engCell(lo, i, "TargetType"))), targetType, vbTextCompare) = 0 Then   ' правило регістру
            If engToBool(engCell(lo, i, "IsRequired"), False) And engToBool(engCellOpt(lo, i, "Use", ""), True) Then
                If Trim$(CStr(engCell(lo, i, "SourceColumnName"))) <> "" Then
                    res.Add Trim$(CStr(engCell(lo, i, "SourceColumnName")))
                End If
            End If
        End If
    Next i
End Function

' ================== ПІДСУМОК І ЗАПИС ==================
Private Function BuildSummary(lo As ListObject, rowsIdx As Collection, _
                              picks() As String, states() As String) As String
    Dim k As Long, r As Long, txt As String, shown As Long, noFile As Long, reqNoFile As Long
    Dim noPeriod As Long, dup As String

    For k = 1 To rowsIdx.Count
        r = rowsIdx(k)
        If picks(k) = "" Then
            noFile = noFile + 1
            If engToBool(engCell(lo, r, "IsRequired"), True) Then reqNoFile = reqNoFile + 1
        End If
        If Trim$(CStr(engCell(lo, r, "PeriodType"))) = "" Or Trim$(CStr(engCell(lo, r, "PeriodNumber"))) = "" Then
            noPeriod = noPeriod + 1
        End If
        If shown < SUMMARY_MAX_ROWS Then
            txt = txt & RowLabel(lo, r) & ": " & _
                  IIf(picks(k) = "", "(не вибрано)", FileNameOf(picks(k))) & _
                  " [" & states(k) & "]" & vbCrLf
            shown = shown + 1
        End If
    Next k
    If rowsIdx.Count > shown Then txt = txt & "... і ще " & (rowsIdx.Count - shown) & " рядків" & vbCrLf

    dup = DuplicateFiles(lo, rowsIdx, picks)

    txt = txt & String(40, "-") & vbCrLf & _
          "Активних рядків: " & rowsIdx.Count & ", без файла: " & noFile
    If reqNoFile > 0 Then txt = txt & vbCrLf & "УВАГА: обов'язкових без файла: " & reqNoFile & _
                               " (fn_ReadPurpose зупиниться з помилкою)"
    If noPeriod > 0 Then txt = txt & vbCrLf & "Рядків без PeriodType або PeriodNumber: " & noPeriod & _
                              " (перевірте, якщо період потрібен)"
    If dup <> "" Then txt = txt & vbCrLf & "УВАГА: той самий файл (і лист) у кількох рядках призначення: " & _
                         dup & " (помилка ДублюванняФайлу)"
    txt = txt & vbCrLf & "Перевірка CSV: роздільник """ & IIf(CfgDelimiter() = vbTab, "Tab", CfgDelimiter()) & _
          """, кодування " & CfgEncoding() & " (з " & ENG_SETTINGS_QUERY & "; BOM файла має пріоритет)"
    BuildSummary = txt
End Function

' Той самий шлях + TableSheetName у кількох рядках одного TargetType (як ДублюванняФайлу в Engine)
Private Function DuplicateFiles(lo As ListObject, rowsIdx As Collection, picks() As String) As String
    Dim seen As Object, k As Long, r As Long, keyName As String, res As String
    Set seen = CreateObject("Scripting.Dictionary")
    For k = 1 To rowsIdx.Count
        If picks(k) <> "" Then
            r = rowsIdx(k)
            keyName = LCase$(Trim$(CStr(engCell(lo, r, "TargetType")))) & "|" & LCase$(picks(k)) & "|" & _
                  LCase$(Trim$(CStr(engCellOpt(lo, r, "TableSheetName", ""))))
            If seen.Exists(keyName) Then
                res = res & IIf(res = "", "", "; ") & seen(keyName) & " і " & CStr(engCell(lo, r, "SourceID"))
            Else
                seen.Add keyName, CStr(engCell(lo, r, "SourceID"))
            End If
        End If
    Next k
    DuplicateFiles = res
End Function

Private Sub WritePicks(lo As ListObject, rowsIdx As Collection, picks() As String)
    Dim k As Long, r As Long
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    For k = 1 To rowsIdx.Count
        r = rowsIdx(k)
        If picks(k) <> "" And LCase$(picks(k)) <> LCase$(engSourcePath(lo, r)) Then
            engSetCell lo, r, "FullPath", picks(k)
            engSetCell lo, r, "Path", Empty
            engSetCell lo, r, "FileName", Empty
        End If
    Next k
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ================== ОНОВЛЕННЯ І ЗВЕДЕННЯ ЯКОСТІ ==================
Private Sub RefreshAndReport()
    Dim errText As String, msg As String, sErr As String
    Application.Cursor = xlWait
    Application.StatusBar = "Оновлення запитів..."
    errText = engRefreshAllSync()
    ' зведення якості вимкнено з «Оновити все» (v4.13): оновлюємо його для підсумку QualityFlag
    Application.StatusBar = "Оновлення " & SUMMARY_QUERY & "..."
    sErr = engRefreshIfExcluded(SUMMARY_QUERY)
    If sErr <> "" Then errText = errText & IIf(errText = "", "", "; ") & sErr
    Application.Cursor = xlDefault
    Application.StatusBar = False
    If errText <> "" Then msg = "Оновлення завершилося з помилкою: " & errText & vbCrLf & vbCrLf
    MsgBox msg & QualityReport(), IIf(msg = "", vbInformation, vbExclamation), "Оновлення"
End Sub

' Підсумок QualityFlag з таблиці pq_QualitySummary на аркуші
Private Function QualityReport() As String
    Dim lo As ListObject
    Set lo = engFindQueryTable(SUMMARY_QUERY)
    If lo Is Nothing Then
        QualityReport = "Таблицю " & SUMMARY_QUERY & " на аркуші не знайдено. Перегляньте звіти якості вручну."
        Exit Function
    End If
    If Not HasColumns(lo, Array("SourceID", "QualityFlag"), True) Then
        QualityReport = "У " & SUMMARY_QUERY & " немає колонки QualityFlag. Перегляньте звіти вручну."
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        QualityReport = SUMMARY_QUERY & ": немає рядків."
        Exit Function
    End If

    Dim i As Long, flag As String, nOK As Long, nWarn As Long, nFail As Long, nOther As Long, failIds As String
    For i = 1 To lo.ListRows.Count
        flag = UCase$(Trim$(CStr(engCell(lo, i, "QualityFlag"))))
        Select Case flag
            Case "OK": nOK = nOK + 1
            Case "WARN": nWarn = nWarn + 1
            Case "FAIL", "BAD_THRESHOLD"
                nFail = nFail + 1
                If nFail <= 10 Then failIds = failIds & vbCrLf & "  * " & CStr(engCell(lo, i, "SourceID")) & " (" & flag & ")"
            Case Else: nOther = nOther + 1
        End Select
    Next i

    QualityReport = SUMMARY_QUERY & ": OK " & nOK & ", WARN " & nWarn & ", FAIL " & nFail & _
                    IIf(nOther > 0, ", без даних " & nOther, "") & _
                    IIf(failIds <> "", vbCrLf & vbCrLf & "Завантаження зупиниться для:" & failIds, "")
End Function

' ================== ДОПОМІЖНІ: ТАБЛИЦІ ==================
Private Function HasColumns(lo As ListObject, colNames As Variant, Optional silent As Boolean = False) As Boolean
    Dim nm As Variant, miss As String
    For Each nm In colNames
        If Not engColumnExists(lo, CStr(nm)) Then miss = miss & " " & CStr(nm)
    Next nm
    HasColumns = (miss = "")
    If miss <> "" And Not silent Then
        MsgBox "У таблиці " & lo.Name & " немає колонок:" & miss, vbCritical
    End If
End Function

' Активні рядки: TargetType заповнений і Enabled порожнє або TRUE (як fn_ActiveSources)
Private Function ActiveRows(lo As ListObject) As Collection
    Dim res As New Collection, i As Long
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            If Trim$(CStr(engCell(lo, i, "TargetType"))) <> "" Then
                If engToBool(engCell(lo, i, "Enabled"), True) Then res.Add i
            End If
        Next i
    End If
    Set ActiveRows = res
End Function

Private Function AnyValidPath(lo As ListObject, rowsIdx As Collection) As Boolean
    Dim v As Variant, p As String
    For Each v In rowsIdx
        p = engSourcePath(lo, CLng(v))
        If p <> "" Then
            If engFileExists(p) Then AnyValidPath = True: Exit Function
        End If
    Next v
End Function

' ================== ПАРАМЕТРИ (з cfgSettings через modEngineCommon) ==================
Private Function RegistryTable() As String
    RegistryTable = engTableName("Registry")
End Function

Private Function CfgDelimiter() As String
    CfgDelimiter = CStr(engSetting("CsvDelimiter"))
End Function

Private Function CfgEncoding() As Long
    CfgEncoding = CLng(engSetting("CsvEncoding"))
End Function

' "csv, txt, xlsx, ..." з cfgSettings[Extensions]
Private Function AllowedList() As String
    AllowedList = Replace(Mid$(CStr(engSetting("Extensions.Csv")) & Mid$(CStr(engSetting("Extensions.Excel")), 2), 2), "|", ", ")
    If Right$(AllowedList, 2) = ", " Then AllowedList = Left$(AllowedList, Len(AllowedList) - 2)
End Function

' Фільтр діалогу вибору: «Дані (*.csv;*.txt;...)»
Private Function FileFilter() As String
    Dim masks As String
    masks = "*." & Replace(AllowedList(), ", ", ";*.")
    FileFilter = "Дані (" & masks & ")," & masks & ",Усі файли (*.*),*.*"
End Function

Private Function RowLabel(lo As ListObject, r As Long) As String
    Dim pt As String, pn As String
    pt = Trim$(CStr(engCell(lo, r, "PeriodType")))
    pn = Trim$(CStr(engCell(lo, r, "PeriodNumber")))
    RowLabel = Trim$(CStr(engCell(lo, r, "SourceID"))) & " (" & Trim$(CStr(engCell(lo, r, "TargetType"))) & _
               IIf(pt & pn <> "", ", " & Trim$(pt & " " & pn), "") & ")"
End Function

' ================== ДОПОМІЖНІ: ЗНАЧЕННЯ, ФАЙЛИ, CSV ==================
Private Function FileNameOf(filePath As String) As String
    FileNameOf = Mid$(filePath, InStrRev(filePath, "\") + 1)
End Function

' Відкрити діалог у теці поточного файла (лише локальні диски)
Private Sub GoToFolderOf(filePath As String)
    Dim p As Long: p = InStrRev(filePath, "\")
    If p < 3 Then Exit Sub
    On Error Resume Next
    If Mid$(filePath, 2, 1) = ":" Then ChDrive Left$(filePath, 1)
    ChDir Left$(filePath, p - 1)
    On Error GoTo 0
End Sub

Private Function CharsetOf(enc As Long) As String
    Select Case enc
        Case 65001: CharsetOf = "utf-8"
        Case 1200: CharsetOf = "unicode"
        Case 1251: CharsetOf = "windows-1251"
        Case 866: CharsetOf = "cp866"
        Case Else: CharsetOf = "windows-1251"
    End Select
End Function

Private Function ReadFirstLine(filePath As String, csName As String, ByRef ok As Boolean) As String
    Dim st As Object, txt As String, p As Long
    On Error GoTo Bad
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.charset = csName
    st.Open
    st.LoadFromFile filePath
    txt = st.ReadText(16384)
    st.Close
    p = InStr(txt, vbLf)
    If p > 0 Then txt = Left$(txt, p - 1)
    txt = Replace(txt, vbCr, "")
    If Len(txt) > 0 Then
        If AscW(Left$(txt, 1)) = &HFEFF Then txt = Mid$(txt, 2)
    End If
    ok = True
    ReadFirstLine = txt
    Exit Function
Bad:
    ok = False
End Function

' Розбір рядка CSV з урахуванням лапок (як QuoteStyle.Csv)
Private Function SplitCsvLine(txtLine As String, delim As String) As Collection
    Dim res As New Collection, i As Long, ch As String, cur As String, inQ As Boolean
    i = 1
    Do While i <= Len(txtLine)
        ch = Mid$(txtLine, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid$(txtLine, i + 1, 1) = """" Then cur = cur & """": i = i + 1 Else inQ = False
            Else
                cur = cur & ch
            End If
        ElseIf ch = """" Then
            inQ = True
        ElseIf Mid$(txtLine, i, Len(delim)) = delim Then
            res.Add cur: cur = "": i = i + Len(delim) - 1
        Else
            cur = cur & ch
        End If
        i = i + 1
    Loop
    res.Add cur
    Set SplitCsvLine = res
End Function

' Точний збіг, як у Engine (регістр і пробіли важливі)
Private Function InCollection(item As String, col As Collection) As Boolean
    Dim v As Variant
    For Each v In col
        If CStr(v) = item Then InCollection = True: Exit Function
    Next v
End Function
