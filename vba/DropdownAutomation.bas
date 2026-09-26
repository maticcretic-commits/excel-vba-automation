Attribute VB_Name = "DropdownAutomation"
Option Explicit

'======================================================================
' Dropdown Automation
' ---------------------------------------------------------------------
' Builds dependent (cascading) dropdowns using named ranges plus data
' validation — no formulas typed by hand, no INDIRECT hacks left for
' the user to maintain. Re-runnable: rebuilds named ranges from the
' Lists sheet and reapplies validation on the Form sheet.
'
' Expected workbook layout:
'   Lists : row 1 = master category names; each column below a header
'           lists the allowed values for that category.
'           e.g.  A: Department      B: Role
'                 Operations         Analyst
'                 Finance            Manager
'   Form  : B2 = master choice (e.g. Department)
'           B3 = dependent choice (e.g. Role within the department)
'======================================================================

'----------------------------------------------------------------------
' Main entry point: rebuild named ranges from Lists, then wire the
' dropdowns on Form.
'----------------------------------------------------------------------
Public Sub BuildCascadingDropdowns()
    Dim listsWs As Worksheet, formWs As Worksheet
    Set listsWs = ThisWorkbook.Worksheets("Lists")
    Set formWs = ThisWorkbook.Worksheets("Form")

    Application.ScreenUpdating = False

    CreateNamedRanges listsWs
    ApplyMasterDropdown formWs, listsWs
    ApplyDependentDropdown formWs

    Application.ScreenUpdating = True
    MsgBox "Cascading dropdowns are ready. Pick a value in B2, then choose in B3.", _
           vbInformation, "Done"
End Sub

'----------------------------------------------------------------------
' Clears the Form inputs (and their validation) so a user can start
' over. Named ranges are kept — use RefreshLists to rebuild those.
'----------------------------------------------------------------------
Public Sub ResetForm()
    Dim formWs As Worksheet
    Set formWs = ThisWorkbook.Worksheets("Form")

    On Error Resume Next
    formWs.Range("B2:B3").Validation.Delete
    formWs.Range("B2:B3").ClearContents
    On Error GoTo 0

    MsgBox "Form cleared.", vbInformation, "Done"
End Sub

'----------------------------------------------------------------------
' Re-reads the Lists sheet and recreates the named ranges. Use after
' adding new categories or values to Lists.
'----------------------------------------------------------------------
Public Sub RefreshLists()
    CreateNamedRanges ThisWorkbook.Worksheets("Lists")
    MsgBox "Named ranges refreshed from the Lists sheet.", vbInformation, "Done"
End Sub

'----------------------------------------------------------------------
' For each header in row 1 of Lists, creates a workbook-scoped named
' range covering the non-blank values beneath it. Also creates a
' "Masters" range holding all headers (for the master dropdown).
' Old ranges created by this tool are deleted first.
'----------------------------------------------------------------------
Private Sub CreateNamedRanges(ByVal listsWs As Worksheet)
    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        If Left(nm.Name, 6) = "DDMAP_" Then nm.Delete
    Next nm

    Dim lastCol As Long, c As Long
    lastCol = listsWs.Cells(1, listsWs.Columns.Count).End(xlToLeft).Column

    Dim masterList As String
    masterList = ""

    For c = 1 To lastCol
        Dim header As String, safeName As String
        header = Trim(CStr(listsWs.Cells(1, c).Value))
        If header <> "" Then
            safeName = SafeRangeName(header)

            Dim lastItem As Long
            lastItem = listsWs.Cells(listsWs.Rows.Count, c).End(xlUp).Row
            If lastItem >= 2 Then
                ThisWorkbook.Names.Add _
                    Name:="DDMAP_" & safeName, _
                    RefersTo:="='" & listsWs.Name & "'!$" & _
                              ColLetter(c) & "$2:$" & ColLetter(c) & "$" & lastItem
            End If

            If masterList <> "" Then masterList = masterList & ","
            masterList = masterList & "DDMAP_" & safeName
        End If
    Next c

    If masterList <> "" Then
        ThisWorkbook.Names.Add Name:="DDMAP_Masters", RefersTo:="=""" & masterList & """"
    End If
End Sub

' Master dropdown: a plain list of the category headers in Form!B2.
Private Sub ApplyMasterDropdown(ByVal formWs As Worksheet, ByVal listsWs As Worksheet)
    Dim lastCol As Long, c As Long, listStr As String
    lastCol = listsWs.Cells(1, listsWs.Columns.Count).End(xlToLeft).Column
    listStr = ""
    For c = 1 To lastCol
        If Trim(CStr(listsWs.Cells(1, c).Value)) <> "" Then
            If listStr <> "" Then listStr = listStr & ","
            listStr = listStr & listsWs.Cells(1, c).Value
        End If
    Next c

    With formWs.Range("B2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=listStr
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Choose"
        .InputMessage = "Pick a category from the list."
    End With
End Sub

' Dependent dropdown: Form!B3 looks up the named range that matches
' the current Form!B2 choice (spaces stripped, prefixed DDMAP_).
Private Sub ApplyDependentDropdown(ByVal formWs As Worksheet)
    With formWs.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=INDIRECT(""DDMAP_""&SUBSTITUTE($B$2,"" "",""""))"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Choose"
        .InputMessage = "Pick a value — the list depends on your B2 choice."
    End With
End Sub

' Turns a header into a valid Excel name: strip spaces and anything
' that is not a letter, digit or underscore.
Private Function SafeRangeName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            out = out & ch
        End If
    Next i
    SafeRangeName = out
End Function

' Column index -> letter (works past Z, e.g. 28 -> "AB").
Private Function ColLetter(ByVal colIndex As Long) As String
    ColLetter = Split(Cells(1, colIndex).Address, "$")(1)
End Function
