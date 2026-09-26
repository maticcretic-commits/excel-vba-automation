Attribute VB_Name = "TimesheetPayroll"
Option Explicit

'======================================================================
' Timesheet & Payroll Automation
' ---------------------------------------------------------------------
' Builds a monthly timesheet from the Employees roster, calculates
' regular vs overtime hours, computes gross pay and tax, and produces
' a payroll summary table ready for review / export.
'
' Expected workbook layout:
'   Employees  : A=Emp ID | B=Name | C=Department | D=Role | E=Hourly Rate
'   Timesheet  : A=Emp ID | B=Name | C=Date | D=Day | E=Hours Worked
'   Payroll    : A=Emp ID | B=Name | C=Regular Hrs | D=OT Hrs
'                E=Gross Pay | F=Tax | G=Net Pay
'
' Daily regular-hours cap: 8 (anything above is overtime at 1.5x).
'======================================================================

Private Const DAILY_REGULAR_HOURS As Double = 8
Private Const OVERTIME_MULTIPLIER As Double = 1.5

'----------------------------------------------------------------------
' Generates a blank monthly timesheet: one row per employee per
' weekday of the chosen month. Run first, fill in hours, then run
' CalculatePayroll.
'----------------------------------------------------------------------
Public Sub GenerateMonthlyTimesheet()
    Dim yr As Long, mth As Long
    yr = InputBox("Enter year (e.g. 2026):", "Generate Timesheet", Year(Date))
    If yr = 0 Then Exit Sub
    mth = InputBox("Enter month (1-12):", "Generate Timesheet", Month(Date))
    If mth < 1 Or mth > 12 Then Exit Sub

    Dim empWs As Worksheet, tsWs As Worksheet
    Set empWs = ThisWorkbook.Worksheets("Employees")
    Set tsWs = ThisWorkbook.Worksheets("Timesheet")

    Application.ScreenUpdating = False

    tsWs.Cells.Clear
    tsWs.Range("A1:E1").Value = Array("Emp ID", "Name", "Date", "Day", "Hours Worked")

    Dim lastRow As Long, r As Long, outRow As Long
    lastRow = empWs.Cells(empWs.Rows.Count, "A").End(xlUp).Row
    outRow = 2

    Dim d As Long, dt As Date, wd As Long
    For d = 1 To Day(DateSerial(yr, mth + 1, 0))
        dt = DateSerial(yr, mth, d)
        wd = Weekday(dt, vbMonday)
        If wd <= 5 Then ' weekdays only
            For r = 2 To lastRow
                If empWs.Cells(r, "A").Value <> "" Then
                    tsWs.Cells(outRow, "A").Value = empWs.Cells(r, "A").Value
                    tsWs.Cells(outRow, "B").Value = empWs.Cells(r, "B").Value
                    tsWs.Cells(outRow, "C").Value = dt
                    tsWs.Cells(outRow, "D").Value = Format(dt, "ddd")
                    tsWs.Cells(outRow, "E").Value = DAILY_REGULAR_HOURS
                    outRow = outRow + 1
                End If
            Next r
        End If
    Next d

    tsWs.Columns("A:E").AutoFit
    tsWs.Range("C2:C" & outRow - 1).NumberFormat = "dd-mmm-yyyy"
    Application.ScreenUpdating = True

    MsgBox "Timesheet skeleton created for " & MonthName(mth) & " " & yr & _
           " (" & outRow - 2 & " rows). Fill in Hours Worked, then run CalculatePayroll.", _
           vbInformation, "Done"
End Sub

'----------------------------------------------------------------------
' Reads Timesheet + Employees, aggregates hours per employee, and
' writes the Payroll summary. Handles empty timesheets gracefully.
'----------------------------------------------------------------------
Public Sub CalculatePayroll()
    Dim empWs As Worksheet, tsWs As Worksheet, payWs As Worksheet
    Set empWs = ThisWorkbook.Worksheets("Employees")
    Set tsWs = ThisWorkbook.Worksheets("Timesheet")
    Set payWs = ThisWorkbook.Worksheets("Payroll")

    Dim empLast As Long, tsLast As Long
    empLast = empWs.Cells(empWs.Rows.Count, "A").End(xlUp).Row
    tsLast = tsWs.Cells(tsWs.Rows.Count, "A").End(xlUp).Row
    If tsLast < 2 Then
        MsgBox "Timesheet is empty. Run GenerateMonthlyTimesheet first.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    payWs.Cells.Clear
    payWs.Range("A1:G1").Value = _
        Array("Emp ID", "Name", "Regular Hrs", "OT Hrs", "Gross Pay", "Tax", "Net Pay")

    Dim r As Long, t As Long, outRow As Long
    Dim empId As Variant, rate As Double
    Dim regHrs As Double, otHrs As Double, gross As Double, tax As Double
    outRow = 2

    For r = 2 To empLast
        empId = empWs.Cells(r, "A").Value
        If empId <> "" Then
            rate = CDbl(empWs.Cells(r, "E").Value)
            regHrs = 0: otHrs = 0
            For t = 2 To tsLast
                If tsWs.Cells(t, "A").Value = empId Then
                    regHrs = regHrs + RegularHours(CDbl(tsWs.Cells(t, "E").Value))
                    otHrs = otHrs + OvertimeHours(CDbl(tsWs.Cells(t, "E").Value))
                End If
            Next t

            gross = (regHrs * rate) + (otHrs * rate * OVERTIME_MULTIPLIER)
            tax = MonthlyTax(gross)

            payWs.Cells(outRow, "A").Value = empId
            payWs.Cells(outRow, "B").Value = empWs.Cells(r, "B").Value
            payWs.Cells(outRow, "C").Value = Round(regHrs, 2)
            payWs.Cells(outRow, "D").Value = Round(otHrs, 2)
            payWs.Cells(outRow, "E").Value = Round(gross, 2)
            payWs.Cells(outRow, "F").Value = Round(tax, 2)
            payWs.Cells(outRow, "G").Value = Round(gross - tax, 2)
            outRow = outRow + 1
        End If
    Next r

    ' Totals row
    payWs.Cells(outRow, "B").Value = "TOTAL"
    payWs.Cells(outRow, "B").Font.Bold = True
    payWs.Range("C" & outRow & ":G" & outRow).FormulaR1C1 = _
        "=SUM(R2C:R[-1]C)"
    payWs.Range("C" & outRow & ":G" & outRow).Font.Bold = True

    With payWs.Range("A1:G" & outRow)
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    payWs.Range("E2:G" & outRow).NumberFormat = "#,##0.00"

    Application.ScreenUpdating = True
    MsgBox "Payroll calculated for " & outRow - 2 & " employees.", vbInformation, "Done"
End Sub

'----------------------------------------------------------------------
' Splits a day's hours into regular (<= 8) and overtime (> 8) portions.
'----------------------------------------------------------------------
Public Function RegularHours(ByVal totalHours As Double) As Double
    If totalHours <= DAILY_REGULAR_HOURS Then
        RegularHours = totalHours
    Else
        RegularHours = DAILY_REGULAR_HOURS
    End If
End Function

Public Function OvertimeHours(ByVal totalHours As Double) As Double
    If totalHours > DAILY_REGULAR_HOURS Then
        OvertimeHours = totalHours - DAILY_REGULAR_HOURS
    Else
        OvertimeHours = 0
    End If
End Function

'----------------------------------------------------------------------
' Monthly tax on a given monthly gross: annualize, apply the slab
' schedule (simplified India-style slabs), then divide by 12.
' Also usable directly as a worksheet formula: =MonthlyTax(50000)
'----------------------------------------------------------------------
Public Function MonthlyTax(ByVal monthlyGross As Double) As Double
    Dim annualGross As Double
    annualGross = monthlyGross * 12
    MonthlyTax = AnnualTax(annualGross) / 12
End Function

Public Function AnnualTax(ByVal annualIncome As Double) As Double
    Dim tax As Double
    tax = 0

    tax = tax + SlabPortion(annualIncome, 400000, 800000, 0.05)
    tax = tax + SlabPortion(annualIncome, 800000, 1200000, 0.1)
    tax = tax + SlabPortion(annualIncome, 1200000, 1600000, 0.15)
    tax = tax + SlabPortion(annualIncome, 1600000, 2000000, 0.2)
    tax = tax + SlabPortion(annualIncome, 2000000, 2400000, 0.25)
    tax = tax + SlabPortion(annualIncome, 2400000, 1E+18, 0.3)

    AnnualTax = tax
End Function

Private Function SlabPortion(ByVal income As Double, _
                             ByVal lower As Double, _
                             ByVal upper As Double, _
                             ByVal rate As Double) As Double
    Dim taxable As Double
    taxable = Application.WorksheetFunction.Max(0, _
              Application.WorksheetFunction.Min(income, upper) - lower)
    SlabPortion = taxable * rate
End Function
