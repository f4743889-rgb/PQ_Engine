Attribute VB_Name = "modImportPowerQueryFromFolder"
Option Explicit

' ============================================================================
' modImportPowerQueryFromFolder: імпорт або оновлення запитів Power Query з файлів .pq (v4.16)
' Самодостатній (без modEngineCommon): ним збирають книгу з нуля.
'
' ImportPQQueriesFromFolder: ім'я файла = ім'я запиту, вміст - код M у UTF-8.
' Шукає .pq у вибраній теці; якщо там їх немає - у відомих підтеках: powerquery\queries
' (Git-репозиторій), engine\powerquery і for_workbook\powerquery (пакет). Запити пишуться
' в книгу з цим модулем (ThisWorkbook), а не в активну.
' Перед записом показує підсумок (створити / змінити / без змін / є лише в книзі) і просить
' підтвердження; записує лише створені й змінені запити, нічого не видаляє. Резервна копія - Git.
' ============================================================================

Public Sub ImportPQQueriesFromFolder()
    Dim fd As FileDialog, fso As Object, file As Object, pqQuery As WorkbookQuery
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

    ' 1. Попередній підсумок (рецензія коду 25.10.2026): що буде створено, змінено, лишиться без
    '    змін; які запити книги відсутні в теці (макрос їх не чіпає і нічого не видаляє)
    Dim codes As Object, q As WorkbookQuery, k As Variant, nNew As Long, nChg As Long, nSame As Long
    Dim newList As String, chgList As String, extraList As String, nExtra As Long
    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Set codes = CreateObject("Scripting.Dictionary")
    codes.CompareMode = vbTextCompare
    For Each file In fso.GetFolder(folderPath).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "pq" Then
            codes(fso.GetBaseName(file.Name)) = ReadUtf8File(file.path)
        End If
    Next file
    For Each k In codes.Keys
        Set pqQuery = Nothing
        On Error Resume Next
        Set pqQuery = ThisWorkbook.Queries(CStr(k))
        On Error GoTo Fail
        If pqQuery Is Nothing Then
            nNew = nNew + 1: If nNew <= 15 Then newList = newList & vbCrLf & "    + " & k
        ElseIf SameCode(pqQuery.formula, CStr(codes(k))) Then
            nSame = nSame + 1
        Else
            nChg = nChg + 1: If nChg <= 15 Then chgList = chgList & vbCrLf & "    ~ " & k
        End If
    Next k
    For Each q In ThisWorkbook.Queries
        If Not codes.Exists(q.Name) Then
            nExtra = nExtra + 1: If nExtra <= 15 Then extraList = extraList & vbCrLf & "    = " & q.Name
        End If
    Next q
    If nNew + nChg = 0 Then
        MsgBox "Тека: " & folderPath & vbCrLf & vbCrLf & "Усі " & nSame & " запитів збігаються з книгою - змінювати нічого.", _
               vbInformation, "Power Query Import"
        Exit Sub
    End If
    If MsgBox("Книга: " & ThisWorkbook.Name & vbCrLf & "Тека: " & folderPath & vbCrLf & vbCrLf & _
              "Буде створено: " & nNew & newList & IIf(nNew > 15, vbCrLf & "    ...", "") & vbCrLf & _
              "Буде змінено: " & nChg & chgList & IIf(nChg > 15, vbCrLf & "    ...", "") & vbCrLf & _
              "Без змін: " & nSame & vbCrLf & _
              IIf(nExtra > 0, "Є в книзі, немає в теці (не чіпаються): " & nExtra & extraList & _
                  IIf(nExtra > 15, vbCrLf & "    ...", "") & vbCrLf, "") & vbCrLf & _
              "Резервна копія - Git: якщо в книзі є зміни, яких ще немає в репозиторії, спершу" & vbCrLf & _
              "зробіть ExportPowerQueriesToRepo і коміт. Продовжити?", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Power Query Import") <> vbYes Then Exit Sub

    ' 2. Запис лише створених і змінених запитів
    Application.ScreenUpdating = False

    For Each k In codes.Keys
        queryName = CStr(k)
        mCode = CStr(codes(k))
        Set pqQuery = Nothing
        On Error Resume Next
        Set pqQuery = ThisWorkbook.Queries(queryName)
        On Error GoTo Fail

        If pqQuery Is Nothing Then
            ThisWorkbook.Queries.Add Name:=queryName, formula:=mCode
            countImported = countImported + 1
        ElseIf Not SameCode(pqQuery.formula, mCode) Then
            pqQuery.formula = mCode
            countUpdated = countUpdated + 1
        End If
    Next k

    Application.ScreenUpdating = oldScreen
    MsgBox "Імпорт завершено." & vbCrLf & vbCrLf & _
           "Тека: " & folderPath & vbCrLf & _
           "• Створено нових запитів: " & countImported & vbCrLf & _
           "• Змінено: " & countUpdated & vbCrLf & _
           "• Без змін: " & nSame, _
           vbInformation, "Power Query Import"
    Exit Sub
Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "Імпорт перервано на «" & queryName & "»: " & Err.Description & vbCrLf & _
           "Встигли: створено " & countImported & ", змінено " & countUpdated & ".", _
           vbCritical, "Power Query Import"
End Sub

Private Function ReadUtf8File(filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                                          ' adTypeText
    stream.charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8File = stream.ReadText
    stream.Close
End Function

' Код однаковий, якщо збігається без урахування BOM, закінчень рядків і пробілів у кінці
Private Function SameCode(a As String, b As String) As Boolean
    SameCode = (NormCode(a) = NormCode(b))
End Function

Private Function NormCode(ByVal t As String) As String
    t = Replace(t, ChrW(&HFEFF), "")
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    Do While Len(t) > 0 And (Right$(t, 1) = vbLf Or Right$(t, 1) = " " Or Right$(t, 1) = vbTab)
        t = Left$(t, Len(t) - 1)
    Loop
    NormCode = t
End Function

Private Function CountPq(fso As Object, dirPath As String) As Long
    Dim f As Object
    If Not fso.FolderExists(dirPath) Then Exit Function
    For Each f In fso.GetFolder(dirPath).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "pq" Then CountPq = CountPq + 1
    Next f
End Function
