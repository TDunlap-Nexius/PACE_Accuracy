Attribute VB_Name = "misc"
Option Explicit
Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long


Function getListObjectFromRange(cell As Excel.Range) As Scripting.Dictionary
    Dim d As Scripting.Dictionary, ws As Worksheet, lo As ListObject, _
        lr As ListRow, lc As ListColumn
    Set ws = cell.Parent
    For Each lo In ws.ListObjects
        If Not Intersect(cell, lo.DataBodyRange) Is Nothing Then
            If lo.ShowHeaders Then
                Set lr = lo.ListRows(cell.Row - lo.Range(1).Row)
            Else
                Set lr = lo.ListRows(cell.Row - lo.Range(1).Row - 1)
            End If
            Set lc = lo.ListColumns(cell.Column - lo.Range(1).Column + 1)
            
            Set d = New Scripting.Dictionary
            d.Add key:="lo", Item:=lo
            d.Add key:="lr", Item:=lr
            d.Add key:="lc", Item:=lc
            Set getListObjectFromRange = d
            Exit Function
        End If
    Next lo
End Function


Function getLoValue(lr As ListRow, lc_name As String) As Variant
    Dim lo As ListObject, lc As ListColumn
    Set lo = lr.Parent
    Set lc = lo.ListColumns(lc_name)
    getLoValue = lc.DataBodyRange(lr.Index).value
End Function


Function setLoValue(lr As ListRow, lc_name As String, value As Variant) As Variant
    Dim lo As ListObject, lc As ListColumn
    Set lo = lr.Parent
    Set lc = lo.ListColumns(lc_name)
    lc.DataBodyRange(lr.Index).value = value
End Function


Private Sub concatSelection()
    Dim cell As Excel.Range, s As String
    Const concat As String = "; "
    For Each cell In Selection
        If cell.value <> "" Then
            s = s & cell.value & concat
        End If
    Next cell
    s = Left(s, Len(s) - Len(concat))
    copyToClipboard s
End Sub


Private Function concatDictionary(d As Scripting.Dictionary, _
                                  concat As String) As String
    Dim i As Long, s As String
    For i = 0 To d.Count - 2
        s = s & d.Keys(i) & concat
    Next i
    s = s & d.Keys(d.Count - 1)
    concatDictionary = s
End Function


Sub BubbleSortArray(ary())
'   Sorts an array using bubble sort algorithm
    Dim First As Integer, Last As Long
    Dim i As Long, j As Long
    Dim temp As Long
    
    First = LBound(ary)
    Last = UBound(ary)
    For i = First To Last - 1
        For j = i + 1 To Last
            If ary(i) > ary(j) Then
                temp = ary(j)
                ary(j) = ary(i)
                ary(i) = temp
            End If
        Next j
    Next i
End Sub


Public Sub SetClipboard(sUniText As String)
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

Public Function GetClipboard() As String
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Dim sUniText As String
    Const CF_UNICODETEXT As Long = 13&
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
            lstrcpy StrPtr(sUniText), iLock
            GlobalUnlock iStrPtr
        End If
        GetClipboard = sUniText
    End If
    CloseClipboard
End Function


Sub showUberFilter()
Attribute showUberFilter.VB_ProcData.VB_Invoke_Func = "Q\n14"
    frm_UberFilter.Show
End Sub


Private Function selectFolder() As Object
    Dim picker As Object, fso As Object, fol As Object, fol_path As String
    Set picker = Application.FileDialog(msoFileDialogFolderPicker)
    With picker
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then Exit Function
        fol_path = .SelectedItems(1)
    End With
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fol = fso.getFolder(fol_path)
    Set selectFolder = fol
End Function


Private Function selectFile() As Object
    Dim picker As Object, fso As Object, fil As Object, fil_path As String
    Set picker = Application.FileDialog(msoFileDialogFilePicker)
    With picker
        .Title = "Select a File"
        .AllowMultiSelect = False
        .InitialFileName = "C:\Users\" & Environ("username") & "\Downloads\"
        .Filters.Add "CSV files", "*.csv"
        If .Show <> -1 Then Exit Function
        fil_path = .SelectedItems(1)
    End With
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fil = fso.GetFile(fil_path)
    Set selectFile = fil
End Function


Private Sub addListObjectSubtotals()
    Dim lo As ListObject, cell As Range
    Set lo = misc.getListObjectFromRange(ActiveCell.Offset(2, 0))("lo")
    For Each cell In Selection
        cell.Formula = "=SUBTOTAL(9, " & lo.Name & "[" & cell.Offset(1, 0).value & "])"
        cell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Next cell
End Sub


Private Sub test()
    Dim cell As Range
    For Each cell In Selection
        cell.Offset(0, -1).value = Format(cell.value, "MMM 'YY")
    Next cell
End Sub


Private Sub getPaceLock()
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, tasks As Object, _
        lo_pace As ListObject, key As Variant, lc_af As ListColumn, _
        lr As ListRow, lr_new As ListRow
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Cells(1, 1).value = "PACE Job Number"
    ws.Cells(1, 2).value = "Task"
    ws.Cells(1, 3).value = "Forecast"
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                Source:=ws.Range(ws.Cells(1, 1), ws.Cells(1, 3)), _
                                XlListObjectHasHeaders:=xlYes)
    With ActiveSheet
        If .ListObjects.Count > 0 Then
            Set lo_pace = .ListObjects(1)
        Else
            Set lo_pace = .ListObjects.Add( _
                SourceType:=xlSrcRange, _
                Source:=.Range(.Cells(1, 1), _
                               .Cells.SpecialCells(xlCellTypeLastCell)), _
                XlListObjectHasHeaders:=xlYes)
            lo_pace.TableStyle = ""
        End If
    End With
    For Each key In tasks.Keys
        Set lc_af = lo_pace.ListColumns(tasks(key) & "_AF")
        For Each lr In lo_pace.ListRows
            
        Next lr
    Next key
End Sub


Private Function getTaskList() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add key:="RE020", Item:="RE020"
    d.Add key:="CI025", Item:="CI025"
    d.Add key:="CI128", Item:="CI128"
    d.Add key:="CI020", Item:="CI020"
    d.Add key:="CI050", Item:="CI050"
    d.Add key:="CL100", Item:="CL100"
    Set getTaskList = d
End Function


Private Sub combineAccuV()
    Dim ws_from As Worksheet, ws_to As Worksheet, col_map As Object, i As Long, _
        rng As Excel.Range, r_max As Long, key As Variant
    
    Set ws_from = ActiveSheet
    Set ws_to = ActiveSheet
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Set col_map = CreateObject("Scripting.Dictionary")
    i = 2
    Do While ws_from.Cells(1, i).value <> Empty
        Set rng = ws_to.Rows(1).Find(what:=ws_from.Cells(1, i).value, _
                                     LookIn:=xlFormulas, lookat:=xlWhole)
        If Not rng Is Nothing Then col_map.Add key:=rng.Column, Item:=i
        i = i + 1
    Loop
    i = 2
    Do While ws_from.Cells(i, 1).value <> Empty
        r_max = ws_to.Cells(1000000, 1).End(xlUp).Row + 1
        ws_to.Cells(r_max, 1).value = ws_from.Cells(i, 1).value
        For Each key In col_map.Keys
            ws_to.Cells(r_max, key).value = ws_from.Cells(i, col_map(key)).value
        Next key
        'ws_to.Rows(r_max).WrapText = False
        i = i + 1
        Application.StatusBar = i
        DoEvents
    Loop
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
