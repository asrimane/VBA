Attribute VB_Name = "RESET_FILTERS"

Sub ResetFilters_Exclude_WorkBook_By_Name()
Attribute ResetFilters_Exclude_WorkBook_By_Name.VB_ProcData.VB_Invoke_Func = "m\n14"
    On Error GoTo err
    Dim dontResetFiltersOn
    dontResetFiltersOn = Array("Etat par géomaticiens", "Cercle_autocad", "evolution", "13 graphique", "#72 Armoire recap")
      
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim tmpIndex
    Dim listObj As ListObject
    Set wb = ThisWorkbook
    For Each ws In wb.Worksheets
        tmpIndex = Application.Match(ws.Name, dontResetFiltersOn, 0)
        If IsError(tmpIndex) Then
            If ws.FilterMode Then
                ws.ShowAllData
            End If
            For Each listObj In ws.ListObjects
                If listObj.ShowHeaders Then
                    If FilterIsOn(listObj) Then
                        listObj.AutoFilter.ShowAllData
                        listObj.Sort.SortFields.Clear
                    End If
                End If
            Next listObj
        End If
    Next
endOfSub:
    Exit Sub
err:
    MsgBox err + vbCrLf + _
            err.Description + vbCrLf + _
            err.Number + vbCrLf + _
            err.Source, vbCritical + vbOKOnly
End Sub

Private Function FilterIsOn(lo As ListObject) As Boolean
    Dim returnValue As Boolean
    returnValue = False
    On Error Resume Next
    If lo.AutoFilter.Filters.Count > 0 Then
        If err.Number = 0 Then returnValue = True
    End If
    On Error GoTo 0
    FilterIsOn = returnValue
End Function
Sub ResetFiltersExcludWorkBookByIndex()
On Error GoTo err
    Dim dontResetFiltersOn, dontResetFiltersOnSize As Integer
    dontResetFiltersOn = Array(1, 3, 6)
     
    dontResetFiltersOnSize = UBound(dontResetFiltersOn)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim listObj As ListObject
     Set wb = ThisWorkbook
     'Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
        tmpIndex = Application.Match(ws.Index, dontResetFiltersOn, 0)
        If IsError(tmpIndex) Then
            If ws.FilterMode Then
                ws.ShowAllData
            End If
            
            For Each listObj In ws.ListObjects
                If listObj.ShowHeaders Then
                    listObj.AutoFilter.ShowAllData
                    listObj.Sort.SortFields.Clear
                End If
            Next listObj
        End If
    Next
endOfSub:
    Exit Sub
err:
    MsgBox Error$, vbCritical + vbOKOnly
    Resume endOfSub
End Sub

