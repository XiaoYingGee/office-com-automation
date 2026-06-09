Attribute VB_Name = "ExcelEditorBridge"
Option Explicit

' ============================================================
' ExcelEditorBridge.bas — VBA bridge for Excel COM automation
' Runs inside Excel process (zero IPC for object model access)
' Called via Application.Run from Python
' ============================================================

Private Const BUF_INIT As Long = 8192

' --- StringBuilder for O(n) JSON construction ---
Private Type StringBuilder
    buf As String
    pos As Long
End Type

Private Sub SbInit(sb As StringBuilder, Optional initSize As Long = 0)
    If initSize <= 0 Then initSize = BUF_INIT
    sb.buf = Space$(initSize)
    sb.pos = 0
End Sub

Private Sub SbAppend(sb As StringBuilder, s As String)
    Dim sLen As Long
    sLen = Len(s)
    If sLen = 0 Then Exit Sub
    If sb.pos + sLen > Len(sb.buf) Then
        Dim newSize As Long
        newSize = Len(sb.buf) * 2
        If newSize < sb.pos + sLen Then newSize = sb.pos + sLen + BUF_INIT
        sb.buf = sb.buf & Space$(newSize - Len(sb.buf))
    End If
    Mid$(sb.buf, sb.pos + 1, sLen) = s
    sb.pos = sb.pos + sLen
End Sub

Private Function SbToString(sb As StringBuilder) As String
    SbToString = Left$(sb.buf, sb.pos)
End Function

' --- JSON helpers ---
Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, Chr(34), "\" & Chr(34))
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

Private Function JsonString(ByVal s As String) As String
    JsonString = Chr(34) & JsonEscape(s) & Chr(34)
End Function

Private Function JsonValue(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        JsonValue = "null"
    ElseIf VarType(v) = vbString Then
        JsonValue = JsonString(CStr(v))
    ElseIf VarType(v) = vbBoolean Then
        If v Then JsonValue = "true" Else JsonValue = "false"
    ElseIf IsNumeric(v) Then
        JsonValue = CStr(v)
    Else
        JsonValue = JsonString(CStr(v))
    End If
End Function

' ============================================================
' PUBLIC API: InspectWorkbookJson
' Returns workbook structure as JSON string
' ============================================================
Public Function InspectWorkbookJson() As String
    Dim sb As StringBuilder
    SbInit sb, 4096

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    SbAppend sb, "{" & Chr(34) & "sheets" & Chr(34) & ":["

    Dim ws As Worksheet
    Dim i As Long
    i = 0
    For Each ws In wb.Worksheets
        If i > 0 Then SbAppend sb, ","
        SbAppend sb, "{"
        SbAppend sb, Chr(34) & "index" & Chr(34) & ":" & CStr(ws.index)
        SbAppend sb, "," & Chr(34) & "name" & Chr(34) & ":" & JsonString(ws.Name)

        Dim used As Range
        Set used = ws.UsedRange
        If Not used Is Nothing Then
            SbAppend sb, "," & Chr(34) & "used_range" & Chr(34) & ":" & JsonString(used.Address)
            SbAppend sb, "," & Chr(34) & "rows" & Chr(34) & ":" & CStr(used.Rows.Count)
            SbAppend sb, "," & Chr(34) & "cols" & Chr(34) & ":" & CStr(used.Columns.Count)
        End If
        SbAppend sb, "}"
        i = i + 1
    Next ws

    SbAppend sb, "]}"
    InspectWorkbookJson = SbToString(sb)
End Function

' ============================================================
' PUBLIC API: InspectSheetJson(paramsJson)
' Returns sheet data as JSON. Params: {"sheet":..., "max_rows":50, "max_cols":26}
' ============================================================
Public Function InspectSheetJson(ByVal paramsJson As String) As String
    Dim sb As StringBuilder
    SbInit sb, 16384

    Dim ws As Worksheet
    Dim maxRows As Long, maxCols As Long
    Dim sheetRef As Variant

    ' Simple JSON parse for params
    sheetRef = JsonExtractValue(paramsJson, "sheet")
    maxRows = CLng(JsonExtractNumeric(paramsJson, "max_rows", 50))
    maxCols = CLng(JsonExtractNumeric(paramsJson, "max_cols", 26))

    If IsNumeric(sheetRef) Then
        Set ws = ActiveWorkbook.Worksheets(CLng(sheetRef))
    Else
        Set ws = ActiveWorkbook.Worksheets(CStr(sheetRef))
    End If

    SbAppend sb, "{" & Chr(34) & "name" & Chr(34) & ":" & JsonString(ws.Name)

    Dim used As Range
    Set used = ws.UsedRange
    If used Is Nothing Then
        SbAppend sb, "," & Chr(34) & "data" & Chr(34) & ":[]}"
        InspectSheetJson = SbToString(sb)
        Exit Function
    End If

    SbAppend sb, "," & Chr(34) & "used_range" & Chr(34) & ":" & JsonString(used.Address)

    Dim rows As Long, cols As Long
    Dim startRow As Long, startCol As Long
    startRow = used.Row
    startCol = used.Column
    rows = used.Rows.Count
    cols = used.Columns.Count
    If rows > maxRows Then rows = maxRows
    If cols > maxCols Then cols = maxCols

    ' Batch read entire range at once (fast!)
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + rows - 1, startCol + cols - 1))
    Dim vals As Variant
    If rows = 1 And cols = 1 Then
        ReDim vals(1 To 1, 1 To 1)
        vals(1, 1) = dataRange.Value2
    Else
        vals = dataRange.Value2
    End If

    SbAppend sb, "," & Chr(34) & "data" & Chr(34) & ":["

    Dim r As Long, c As Long
    For r = 1 To rows
        If r > 1 Then SbAppend sb, ","
        SbAppend sb, "["
        For c = 1 To cols
            If c > 1 Then SbAppend sb, ","
            SbAppend sb, JsonValue(vals(r, c))
        Next c
        SbAppend sb, "]"
    Next r

    SbAppend sb, "]}"
    InspectSheetJson = SbToString(sb)
End Function

' ============================================================
' PUBLIC API: ExecuteActionsJson(actionsJson)
' Execute array of actions, return JSON results array
' ============================================================
Public Function ExecuteActionsJson(ByVal actionsJson As String) As String
    Dim sb As StringBuilder
    SbInit sb, 4096
    SbAppend sb, "["

    Dim actions As Collection
    Set actions = JsonParseArray(actionsJson)

    Dim i As Long
    For i = 1 To actions.Count
        If i > 1 Then SbAppend sb, ","
        Dim result As String
        result = ExecuteSingleAction(actions(i))
        SbAppend sb, result
    Next i

    SbAppend sb, "]"
    ExecuteActionsJson = SbToString(sb)
End Function

' ============================================================
' PRIVATE: Execute a single action
' ============================================================
Private Function ExecuteSingleAction(ByVal actionJson As String) As String
    On Error GoTo ErrHandler

    Dim act As String
    act = JsonExtractString(actionJson, "action")

    Dim ws As Worksheet
    Dim sheetRef As Variant
    sheetRef = JsonExtractValue(actionJson, "sheet")
    If Not IsEmpty(sheetRef) And Len(CStr(sheetRef)) > 0 Then
        If IsNumeric(sheetRef) Then
            Set ws = ActiveWorkbook.Worksheets(CLng(sheetRef))
        Else
            Set ws = ActiveWorkbook.Worksheets(CStr(sheetRef))
        End If
    Else
        Set ws = ActiveWorkbook.ActiveSheet
    End If

    Select Case act
        Case "write_cell"
            Dim cellAddr As String, cellVal As Variant, kind As String
            cellAddr = JsonExtractString(actionJson, "cell")
            cellVal = JsonExtractValue(actionJson, "value")
            kind = JsonExtractString(actionJson, "kind")
            If kind = "formula" Then
                ws.Range(cellAddr).Formula = CStr(cellVal)
            Else
                ws.Range(cellAddr).Value2 = cellVal
            End If
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "read_cell"
            Dim readAddr As String, prop As String
            readAddr = JsonExtractString(actionJson, "cell")
            prop = JsonExtractString(actionJson, "property")
            If prop = "" Then prop = "value2"
            Dim readVal As Variant
            If prop = "formula" Then
                readVal = ws.Range(readAddr).Formula
            ElseIf prop = "text" Then
                readVal = ws.Range(readAddr).Text
            Else
                readVal = ws.Range(readAddr).Value2
            End If
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "," & Chr(34) & "value" & Chr(34) & ":" & JsonValue(readVal) & "}"

        Case "write_range"
            Dim wRangeAddr As String
            wRangeAddr = JsonExtractString(actionJson, "range")
            Dim wVals As Variant
            wVals = JsonExtract2DArray(actionJson, "values")
            ws.Range(wRangeAddr).Value2 = wVals
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "read_range"
            Dim rRangeAddr As String
            rRangeAddr = JsonExtractString(actionJson, "range")
            Dim rVals As Variant
            rVals = ws.Range(rRangeAddr).Value2
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "," & Chr(34) & "values" & Chr(34) & ":" & Variant2DToJson(rVals) & "}"

        Case "clear_range"
            Dim clrAddr As String, clrMode As String
            clrAddr = JsonExtractString(actionJson, "range")
            clrMode = JsonExtractString(actionJson, "mode")
            If clrMode = "" Then clrMode = "contents"
            If clrMode = "all" Then
                ws.Range(clrAddr).Clear
            Else
                ws.Range(clrAddr).ClearContents
            End If
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "merge_cells"
            Dim mrgAddr As String
            mrgAddr = JsonExtractString(actionJson, "range")
            ws.Range(mrgAddr).Merge
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "unmerge_cells"
            Dim umrgAddr As String
            umrgAddr = JsonExtractString(actionJson, "range")
            ws.Range(umrgAddr).UnMerge
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "set_format"
            Dim fmtAddr As String
            fmtAddr = JsonExtractString(actionJson, "range")
            Dim rng As Range
            Set rng = ws.Range(fmtAddr)
            ApplyFormat rng, actionJson
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "insert_rows"
            Dim insRow As Long, insCount As Long
            insRow = CLng(JsonExtractNumeric(actionJson, "row", 1))
            insCount = CLng(JsonExtractNumeric(actionJson, "count", 1))
            Dim ir As Long
            For ir = 1 To insCount
                ws.Rows(insRow).Insert
            Next ir
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "delete_rows"
            Dim delRow As Long, delCount As Long
            delRow = CLng(JsonExtractNumeric(actionJson, "row", 1))
            delCount = CLng(JsonExtractNumeric(actionJson, "count", 1))
            ws.Rows(CStr(delRow) & ":" & CStr(delRow + delCount - 1)).Delete
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "insert_cols"
            Dim insCol As Long, insColCnt As Long
            insCol = CLng(JsonExtractNumeric(actionJson, "col", 1))
            insColCnt = CLng(JsonExtractNumeric(actionJson, "count", 1))
            Dim ic As Long
            For ic = 1 To insColCnt
                ws.Columns(insCol).Insert
            Next ic
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "delete_cols"
            Dim delCol As Long, delColCnt As Long
            delCol = CLng(JsonExtractNumeric(actionJson, "col", 1))
            delColCnt = CLng(JsonExtractNumeric(actionJson, "count", 1))
            ws.Columns(CStr(delCol) & ":" & CStr(delCol + delColCnt - 1)).Delete
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "autofit_columns"
            Dim afAddr As String
            afAddr = JsonExtractString(actionJson, "range")
            If afAddr <> "" Then
                ws.Range(afAddr).Columns.AutoFit
            Else
                ws.UsedRange.Columns.AutoFit
            End If
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "add_sheet"
            Dim newName As String, afterSheet As String
            newName = JsonExtractString(actionJson, "name")
            afterSheet = JsonExtractString(actionJson, "after")
            Dim newWs As Worksheet
            If afterSheet <> "" Then
                Set newWs = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(afterSheet))
            Else
                Set newWs = ActiveWorkbook.Worksheets.Add
            End If
            If newName <> "" Then newWs.Name = newName
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "," & Chr(34) & "name" & Chr(34) & ":" & JsonString(newWs.Name) & "}"

        Case "delete_sheet"
            ws.Delete
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "rename_sheet"
            Dim rnNew As String
            rnNew = JsonExtractString(actionJson, "new_name")
            ws.Name = rnNew
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "calculate"
            Application.Calculate
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "sort_range"
            Dim sortAddr As String, keyCol As String, sortOrder As String
            sortAddr = JsonExtractString(actionJson, "range")
            keyCol = JsonExtractString(actionJson, "key_col")
            sortOrder = JsonExtractString(actionJson, "order")
            Dim xlOrd As Long
            If sortOrder = "desc" Then xlOrd = 2 Else xlOrd = 1
            ws.Range(sortAddr).Sort Key1:=ws.Range(keyCol), Order1:=xlOrd, Header:=1
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "find_replace"
            Dim frAddr As String, frFind As String, frRepl As String
            frAddr = JsonExtractString(actionJson, "range")
            frFind = JsonExtractString(actionJson, "find")
            frRepl = JsonExtractString(actionJson, "replace")
            ws.Range(frAddr).Replace What:=frFind, Replacement:=frRepl
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case "export_pdf"
            Dim pdfPath As String
            pdfPath = JsonExtractString(actionJson, "path")
            ActiveWorkbook.ExportAsFixedFormat Type:=0, Filename:=pdfPath
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":true," & Chr(34) & "action" & Chr(34) & ":" & JsonString(act) & "}"

        Case Else
            ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":false," & Chr(34) & "error" & Chr(34) & ":" & JsonString("Unknown action: " & act) & "}"
    End Select
    Exit Function

ErrHandler:
    ExecuteSingleAction = "{" & Chr(34) & "ok" & Chr(34) & ":false," & Chr(34) & "error" & Chr(34) & ":" & JsonString("VBA Error " & Err.Number & ": " & Err.Description) & "}"
End Function

' ============================================================
' PRIVATE: Apply formatting from JSON params
' ============================================================
Private Sub ApplyFormat(rng As Range, ByVal actionJson As String)
    Dim v As Variant

    v = JsonExtractValue(actionJson, "bold")
    If Not IsEmpty(v) Then rng.Font.Bold = CBool(v)

    v = JsonExtractValue(actionJson, "italic")
    If Not IsEmpty(v) Then rng.Font.Italic = CBool(v)

    v = JsonExtractValue(actionJson, "font_size")
    If Not IsEmpty(v) Then rng.Font.Size = CDbl(v)

    v = JsonExtractValue(actionJson, "font_name")
    If Not IsEmpty(v) Then rng.Font.Name = CStr(v)

    v = JsonExtractValue(actionJson, "font_color")
    If Not IsEmpty(v) Then rng.Font.Color = CLng(v)

    v = JsonExtractValue(actionJson, "bg_color")
    If Not IsEmpty(v) Then rng.Interior.Color = CLng(v)

    v = JsonExtractValue(actionJson, "number_format")
    If Not IsEmpty(v) Then rng.NumberFormat = CStr(v)

    v = JsonExtractValue(actionJson, "h_align")
    If Not IsEmpty(v) Then rng.HorizontalAlignment = CLng(v)

    v = JsonExtractValue(actionJson, "v_align")
    If Not IsEmpty(v) Then rng.VerticalAlignment = CLng(v)

    v = JsonExtractValue(actionJson, "merge")
    If Not IsEmpty(v) Then
        If CBool(v) Then rng.Merge
    End If

    v = JsonExtractValue(actionJson, "wrap_text")
    If Not IsEmpty(v) Then rng.WrapText = CBool(v)
End Sub

' ============================================================
' PRIVATE: Minimal JSON parsing (no external dependency)
' ============================================================
Private Function JsonExtractString(ByVal json As String, ByVal key As String) As String
    Dim pattern As String
    pattern = Chr(34) & key & Chr(34) & ":"
    Dim pos As Long
    pos = InStr(1, json, pattern)
    If pos = 0 Then Exit Function
    pos = pos + Len(pattern)
    ' skip whitespace
    Do While pos <= Len(json) And Mid$(json, pos, 1) = " "
        pos = pos + 1
    Loop
    If Mid$(json, pos, 1) = Chr(34) Then
        ' string value
        pos = pos + 1
        Dim endPos As Long
        endPos = pos
        Do While endPos <= Len(json)
            If Mid$(json, endPos, 1) = Chr(34) And Mid$(json, endPos - 1, 1) <> "\" Then Exit Do
            endPos = endPos + 1
        Loop
        JsonExtractString = Mid$(json, pos, endPos - pos)
        JsonExtractString = Replace(JsonExtractString, "\\" & Chr(34), Chr(34))
        JsonExtractString = Replace(JsonExtractString, "\\\\", "\")
    End If
End Function

Private Function JsonExtractNumeric(ByVal json As String, ByVal key As String, Optional defaultVal As Double = 0) As Double
    Dim v As Variant
    v = JsonExtractValue(json, key)
    If IsEmpty(v) Or Len(CStr(v)) = 0 Then
        JsonExtractNumeric = defaultVal
    Else
        JsonExtractNumeric = CDbl(v)
    End If
End Function

Private Function JsonExtractValue(ByVal json As String, ByVal key As String) As Variant
    Dim pattern As String
    pattern = Chr(34) & key & Chr(34) & ":"
    Dim pos As Long
    pos = InStr(1, json, pattern)
    If pos = 0 Then
        JsonExtractValue = Empty
        Exit Function
    End If
    pos = pos + Len(pattern)
    Do While pos <= Len(json) And Mid$(json, pos, 1) = " "
        pos = pos + 1
    Loop
    Dim ch As String
    ch = Mid$(json, pos, 1)
    If ch = Chr(34) Then
        ' string
        pos = pos + 1
        Dim ePos As Long
        ePos = pos
        Do While ePos <= Len(json)
            If Mid$(json, ePos, 1) = Chr(34) And Mid$(json, ePos - 1, 1) <> "\" Then Exit Do
            ePos = ePos + 1
        Loop
        Dim sv As String
        sv = Mid$(json, pos, ePos - pos)
        sv = Replace(sv, "\\" & Chr(34), Chr(34))
        sv = Replace(sv, "\\\\", "\")
        JsonExtractValue = sv
    ElseIf ch = "t" Then
        JsonExtractValue = True
    ElseIf ch = "f" Then
        JsonExtractValue = False
    ElseIf ch = "n" Then
        JsonExtractValue = Null
    Else
        ' number
        Dim numEnd As Long
        numEnd = pos
        Do While numEnd <= Len(json)
            Dim nc As String
            nc = Mid$(json, numEnd, 1)
            If nc = "," Or nc = "}" Or nc = "]" Or nc = " " Then Exit Do
            numEnd = numEnd + 1
        Loop
        Dim ns As String
        ns = Mid$(json, pos, numEnd - pos)
        If InStr(ns, ".") > 0 Then
            JsonExtractValue = CDbl(ns)
        Else
            JsonExtractValue = CLng(ns)
        End If
    End If
End Function

Private Function JsonParseArray(ByVal json As String) As Collection
    ' Minimal: parse top-level JSON array of objects, return each object as raw JSON string
    Dim result As New Collection
    Dim pos As Long
    pos = InStr(1, json, "[")
    If pos = 0 Then
        Set JsonParseArray = result
        Exit Function
    End If
    pos = pos + 1

    Dim depth As Long
    Dim startPos As Long
    Dim inString As Boolean

    Do While pos <= Len(json)
        ' skip whitespace
        Do While pos <= Len(json) And InStr(1, " " & vbCr & vbLf & vbTab, Mid$(json, pos, 1)) > 0
            pos = pos + 1
        Loop
        If pos > Len(json) Then Exit Do
        If Mid$(json, pos, 1) = "]" Then Exit Do
        If Mid$(json, pos, 1) = "," Then
            pos = pos + 1
        Else
            ' start of object/value
            startPos = pos
            depth = 0
            inString = False
            Do While pos <= Len(json)
                Dim c As String
                c = Mid$(json, pos, 1)
                If inString Then
                    If c = "\" Then
                        pos = pos + 1 ' skip escaped char
                    ElseIf c = Chr(34) Then
                        inString = False
                    End If
                Else
                    If c = Chr(34) Then
                        inString = True
                    ElseIf c = "{" Or c = "[" Then
                        depth = depth + 1
                    ElseIf c = "}" Or c = "]" Then
                        If depth <= 1 Then
                            If depth = 1 Then pos = pos + 1
                            Exit Do
                        End If
                        depth = depth - 1
                    ElseIf c = "," And depth = 0 Then
                        Exit Do
                    End If
                End If
                pos = pos + 1
            Loop
            result.Add Mid$(json, startPos, pos - startPos)
        End If
    Loop

    Set JsonParseArray = result
End Function

Private Function JsonExtract2DArray(ByVal json As String, ByVal key As String) As Variant
    ' Extract "values":[[...],[...]] and return as VBA 2D array (1-based)
    Dim pattern As String
    pattern = Chr(34) & key & Chr(34) & ":"
    Dim pos As Long
    pos = InStr(1, json, pattern)
    If pos = 0 Then
        JsonExtract2DArray = Empty
        Exit Function
    End If
    pos = pos + Len(pattern)
    Do While pos <= Len(json) And Mid$(json, pos, 1) = " "
        pos = pos + 1
    Loop

    ' Parse [[val,...],[val,...],...]
    ' First pass: count rows and cols
    Dim rows As Collection
    Set rows = New Collection

    If Mid$(json, pos, 1) <> "[" Then Exit Function
    pos = pos + 1

    Do While pos <= Len(json)
        Do While pos <= Len(json) And InStr(1, " " & vbCr & vbLf, Mid$(json, pos, 1)) > 0
            pos = pos + 1
        Loop
        If Mid$(json, pos, 1) = "]" Then Exit Do
        If Mid$(json, pos, 1) = "," Then pos = pos + 1

        If Mid$(json, pos, 1) = "[" Then
            ' parse one row
            pos = pos + 1
            Dim rowVals As New Collection
            Do While pos <= Len(json)
                Do While pos <= Len(json) And InStr(1, " " & vbCr & vbLf, Mid$(json, pos, 1)) > 0
                    pos = pos + 1
                Loop
                If Mid$(json, pos, 1) = "]" Then
                    pos = pos + 1
                    Exit Do
                End If
                If Mid$(json, pos, 1) = "," Then pos = pos + 1

                ' parse value
                Dim val As Variant
                Dim ch2 As String
                ch2 = Mid$(json, pos, 1)
                If ch2 = Chr(34) Then
                    pos = pos + 1
                    Dim se As Long
                    se = InStr(pos, json, Chr(34))
                    val = Mid$(json, pos, se - pos)
                    pos = se + 1
                ElseIf ch2 = "t" Then
                    val = True: pos = pos + 4
                ElseIf ch2 = "f" Then
                    val = False: pos = pos + 5
                ElseIf ch2 = "n" Then
                    val = Null: pos = pos + 4
                Else
                    Dim ne As Long
                    ne = pos
                    Do While ne <= Len(json)
                        Dim nc2 As String
                        nc2 = Mid$(json, ne, 1)
                        If nc2 = "," Or nc2 = "]" Then Exit Do
                        ne = ne + 1
                    Loop
                    val = CDbl(Mid$(json, pos, ne - pos))
                    pos = ne
                End If
                rowVals.Add val
            Loop
            rows.Add rowVals
            Set rowVals = Nothing
        End If
    Loop

    If rows.Count = 0 Then Exit Function

    Dim nRows As Long, nCols As Long
    nRows = rows.Count
    nCols = rows(1).Count

    Dim arr() As Variant
    ReDim arr(1 To nRows, 1 To nCols)
    Dim ri As Long, ci As Long
    For ri = 1 To nRows
        For ci = 1 To rows(ri).Count
            If ci <= nCols Then arr(ri, ci) = rows(ri)(ci)
        Next ci
    Next ri

    JsonExtract2DArray = arr
End Function

Private Function Variant2DToJson(v As Variant) As String
    Dim sb As StringBuilder
    SbInit sb, 2048

    If IsNull(v) Or IsEmpty(v) Then
        Variant2DToJson = "null"
        Exit Function
    End If

    If Not IsArray(v) Then
        Variant2DToJson = "[[" & JsonValue(v) & "]]"
        Exit Function
    End If

    Dim r As Long, c As Long
    SbAppend sb, "["
    For r = LBound(v, 1) To UBound(v, 1)
        If r > LBound(v, 1) Then SbAppend sb, ","
        SbAppend sb, "["
        For c = LBound(v, 2) To UBound(v, 2)
            If c > LBound(v, 2) Then SbAppend sb, ","
            SbAppend sb, JsonValue(v(r, c))
        Next c
        SbAppend sb, "]"
    Next r
    SbAppend sb, "]"
    Variant2DToJson = SbToString(sb)
End Function

' ============================================================
' Health check
' ============================================================
Public Function Ping() As String
    Ping = "ok"
End Function
