Attribute VB_Name = "codeage"
Sub getPaceLock()
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, tasks As Object, _
        lo_pace As ListObject, key As Variant, lc_af As ListColumn, _
        lr As ListRow, lr_new As ListRow, cur_month As String
    
    cur_month = Format(Date, "MMMYY")
'    With ActiveSheet
'        If .ListObjects.Count > 0 Then
'            Set lo_pace = .ListObjects(1)
'        Else
'            Set lo_pace = .ListObjects.Add( _
'                SourceType:=xlSrcRange, _
'                Source:=.Range(.Cells(1, 1), _
'                               .Cells.SpecialCells(xlCellTypeLastCell)), _
'                XlListObjectHasHeaders:=xlYes)
'            lo_pace.TableStyle = ""
'        End If
'    End With
    Set lo_pace = getListObj
    Set lo = ThisWorkbook.Worksheets("Lock Summary").ListObjects(1)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    Set tasks = getTaskList
    
    For Each key In tasks.Keys
        Set lc_af = lo_pace.ListColumns(key & "_AF")
        For Each lr In lo_pace.ListRows
            If Format(lc_af.DataBodyRange(lr.Index).value, "MMMYY") = cur_month Then
                Set lr_new = lo.ListRows.Add
                setLoValue lr_new, "PACE Job Number", getLoValue(lr, "JOB #")
                setLoValue lr_new, "Task", key
                setLoValue lr_new, "Forecast", lc_af.DataBodyRange(lr.Index).value
            End If
        Next lr
        Application.StatusBar = "Getting " & key
        DoEvents
    Next key
    Application.StatusBar = False
    
'    With lo.Parent.Parent.Worksheets("PACE Lock File")
'        .Cells.Delete
'        copyLockFile lo_pace, .Cells(1, 1)
'    End With
    
    lo_pace.Parent.Parent.Close False
End Sub


Private Function getTaskList() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add key:="SS014", Item:="CYCLE"
    d.Add key:="SS076", Item:="CYCLE"
    d.Add key:="RE007", Item:="CYCLE"
    d.Add key:="RE020", Item:="ALL"
    d.Add key:="CI025", Item:="ALL"
    d.Add key:="CI128", Item:="ALL"
    d.Add key:="CI020", Item:="ALL"
    d.Add key:="CI032", Item:="ALL"
    d.Add key:="CI050", Item:="ALL"
    d.Add key:="CL001", Item:="ALL"
    d.Add key:="CL100", Item:="ALL"
    Set getTaskList = d
End Function

Sub getCurrent()
    Dim lo As ListObject, tasks As Object, lo_acc As ListObject, _
        proj_map As Object, lo_pace As ListObject, _
        lc_af As ListColumn, lc_f As ListColumn, lc_a As ListColumn, _
        lr As ListRow, lr_acc As ListRow, re_datetime As Object, ma As Object, _
        lr_new As ListRow, key As Variant, proj As Variant, _
        datetime As String, pace As String, status As String, _
        action As String, dt_ini As Date, dt_f As Date, dt_af As Date

'    With ActiveSheet
'        If .ListObjects.Count > 0 Then
'            Set lo_acc = .ListObjects(1)
'        Else
'            Set lo_acc = .ListObjects.Add( _
'                SourceType:=xlSrcRange, _
'                Source:=.Range(.Cells(1, 1), _
'                               .Cells.SpecialCells(xlCellTypeLastCell)), _
'                XlListObjectHasHeaders:=xlYes)
'            lo_acc.TableStyle = ""
'        End If
'    End With
    Set lo_acc = getListObj
    Set re_datetime = CreateObject("VBScript.RegExp")
    re_datetime.Pattern = "NDPATT_MT_[^_]+_[^_]+_(\d{2})_(\d{2})_(\d{4})_(\d{2})_(\d{2})_(\d{2})"
    Set ma = re_datetime.Execute(lo_acc.Parent.Parent.Name)
    datetime = CDate(ma(0).submatches(0) & "/" & ma(0).submatches(1) & "/" & _
                     ma(0).submatches(2) & " " & ma(0).submatches(3) & ":" & _
                     ma(0).submatches(4))
'    lo_trk.Range(1).Offset(-2, 2).value = "Last Updated: " & Format(DateTime, "M/D/YY HH:MM")
    
    Set lo = ThisWorkbook.Worksheets("Forecast Accuracy Data").ListObjects(1)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    
    Set lo_pace = ThisWorkbook.Worksheets("Lock Summary").ListObjects(1)
    Set tasks = getTaskList
    Set proj_map = getPaceMap(lo_acc)
    
    For Each key In tasks.Keys
        If tasks(key) = "ALL" Then
            Set lc_af = lo_acc.ListColumns(key & "_AF")
            Set lc_f = lo_acc.ListColumns(key & "_F")
            Set lc_a = lo_acc.ListColumns(key & "_A")
            For Each lr In lo_pace.ListRows
                If getLoValue(lr, "Task") = key Then
                    Set lr_new = lo.ListRows.Add
                    pace = getLoValue(lr, "PACE Job Number")
                    dt_ini = getLoValue(lr, "Forecast")
                    If pace <> Empty And proj_map.Exists(pace) Then
                        Set lr_acc = proj_map(pace)
                        dt_af = lc_af.DataBodyRange(lr_acc.Index)
                        dt_f = lc_f.DataBodyRange(lr_acc.Index)
                        dt_a = lc_a.DataBodyRange(lr_acc.Index)
                        If dt_a = Empty Then ' Not Actual
                            If dt_af > dt_ini + 14 Then ' AF Pushed
                                status = "Pushed 14+"
                                If dt_f < dt_ini + 14 Then ' F Not Pushed
                                    If dt_f > Date Then ' F Not Pushed / Pending approval
                                        action = "Pending Pull In Approval"
                                    Else ' F Not Pushed / Can't approve because PD
                                        action = "Pull In Not Approved"
                                    End If
                                Else ' F Pushed
                                    If Date < dt_ini + 14 Then ' F Pushed / Can Pull In
                                        action = "Need to Pull In"
                                    Else ' F Pushed / Can't Pull In
                                        action = "Can't Pull In"
                                    End If
                                End If
                            ElseIf dt_f > dt_ini + 14 Then ' AF Not Pushed / F Pushed
                                status = "Within 14"
                                action = "Pending Push, Pull back in"
                            Else ' AF Not Pushed / F Not Pushed
                                status = "Within 14"
                                action = "None"
                            End If
                        Else ' Actual
                            action = "Actualized"
                            If dt_a > dt_ini + 14 Then
                                status = "Pushed 14+"
                            Else
                                status = "Within 14"
                            End If
                        End If
                        setLoValue lr_new, "PACE Job Number", pace
                        setLoValue lr_new, "Task", key
                        setLoValue lr_new, "Initial FCST", Format(dt_ini, "mm/dd/yy")
                        setLoValue lr_new, "Cur FCST", Format(dt_f, "mm/dd/yy")
                        setLoValue lr_new, "App FCST", Format(dt_af, "mm/dd/yy")
                        setLoValue lr_new, "Actual", Format(dt_a, "mm/dd/yy")
                        setLoValue lr_new, "Push Status", status
                        setLoValue lr_new, "Action Required", action
                    Else
                        setLoValue lr_new, "PACE Job Number", pace
                        setLoValue lr_new, "Task", key
                        setLoValue lr_new, "Initial FCST", Format(dt_ini, "mm/dd/yy")
                        setLoValue lr_new, "Cur FCST", "N/A"
                        setLoValue lr_new, "App FCST", "N/A"
                        setLoValue lr_new, "Actual", "N/A"
                        setLoValue lr_new, "Push Status", "No Longer In PACE"
                        setLoValue lr_new, "Action Required", "None"
                    End If
                End If
            Next lr
        End If
    Next key
'        For Each lr In lo_pace.ListRows
'            If Format(lc_af.DataBodyRange(lr.Index).value, "MMMYY") = cur_month Then
'                Set lr_new = lo.ListRows.Add
'                setLoValue lr_new, "PACE Job Number", getLoValue(lr, "PACE_NUMBER")
'                setLoValue lr_new, "Task", key
'                setLoValue lr_new, "Forecast", lc_af.DataBodyRange(lr.Index).value
'            End If
'        Next lr
'        Application.StatusBar = "Getting " & key
'        DoEvents
'    Next proj
    Application.StatusBar = False
    lo_acc.Parent.Parent.Close False
End Sub


Private Function getPaceMap(lo As ListObject) As Object
    Dim lr As ListRow, d0 As Object
    Set d0 = CreateObject("Scripting.Dictionary")
    For Each lr In lo.ListRows
        d0.Add key:=getLoValue(lr, "Job #"), Item:=lr
    Next lr
    Set getPaceMap = d0
End Function


Private Function getListObj() As ListObject
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, r_max As Long, _
        c_max As Long, i As Long, fd As Object, wb_filename As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Select Import File"
        .Filters.Add "CSV Files", "*.csv"
        .InitialFileName = "C:\Users\" & Environ("username") & "\Downloads\"
        If .Show = -1 Then
            ws_filename = .SelectedItems(1)
        Else
            Exit Function
        End If
    End With
    
    Set wb = Workbooks.Open(ws_filename, False, True)
    Set ws = wb.Sheets(1)
    If ws.ListObjects.Count > 0 Then
        Set getListObj = ws.ListObjects(1)
        Exit Function
    End If
    For i = 1 To 5
        If ws.Cells(i, 16000).End(xlToLeft).Column > c_max Then
            c_max = ws.Cells(i, 16000).End(xlToLeft).Column
        End If
    Next i
    For i = 1 To c_max
        If ws.Cells(1000000, i).End(xlUp).Row > r_max Then
            r_max = ws.Cells(1000000, i).End(xlUp).Row
        End If
    Next i
    
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                Source:=ws.Range(ws.Cells(1, 1), _
                                                 ws.Cells(r_max, c_max)), _
                                XlListObjectHasHeaders:=xlYes)
    lo.TableStyle = ""
    Set getListObj = lo
End Function


Private Sub copyLockFile(lo As ListObject, target As Range)
    Dim tasks As Object, rng_copy As Range, key_str As String
    Set tasks = getTaskList
    tasks.Add key:="Job #", Item:=Empty
    tasks.Add key:="Location-Code", Item:=Empty
    tasks.Add key:="StrategicYear", Item:=Empty
    tasks.Add key:="JobStatus", Item:=Empty
    tasks.Add key:="Market/Cost Center", Item:=Empty
    tasks.Add key:="Project #", Item:=Empty
    tasks.Add key:="Technology", Item:=Empty
    tasks.Add key:="SiteAcquisition-Vendor", Item:=Empty
    tasks.Add key:="Job-Vendor", Item:=Empty
    tasks.Add key:="Civil-Vendor", Item:=Empty
    tasks.Add key:="MarketStrategicYear", Item:=Empty
    tasks.Add key:="Site-Name", Item:=Empty
    tasks.Add key:="MOD Code", Item:=Empty


    
    For Each key In tasks.Keys
        If tasks(key) = Empty Then
            key_str = key
        Else
            key_str = key & "_AF"
        End If
        If rng_copy Is Nothing Then
            Set rng_copy = lo.ListColumns(key_str).Range
        Else
            Set rng_copy = Union(rng_copy, lo.ListColumns(key_str).Range)
        End If
    Next key
    
    rng_copy.Copy target
    
End Sub
