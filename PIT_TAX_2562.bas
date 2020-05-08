Option Explicit

Sub CalculateTax()
    Dim lastrow As Long
    Dim d As Date
    Dim wb As Workbook
    Dim dsh As Worksheet
    Dim csh As Worksheet
    Dim rsh As Worksheet

    d = Now()
    Set wb = ThisWorkbook
    Set dsh = wb.Sheets("DATA")
    Set csh = wb.Sheets("TaxCalculate")
    Set rsh = wb.Sheets("TaxReport")
    Call ClearReportArea(wb, rsh)
    
    lastrow = dsh.UsedRange.Rows.Count

    Call PreStateReport(wb, lastrow)

   
    MsgBox ("Report run complete " & DateDiff("s", d, Now()) & " Seconds")
End Sub

Sub PreStateReport(ByRef wb As Workbook, ByVal R As Long)
    Dim datarng As Range
    Dim c As Long
    Dim rsh As Worksheet, csh As Worksheet
    Dim sql As String
    Dim constr As String
    constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Macro;HDR=YES"";"
    'constr = "Excel File=" & ThisWorkbook.FullName & ";Offline=true;Query Passthrough=true;Cache Location= " & ThisWorkbook.Path & "\.cache.db;"
    Dim cn As Object
    Dim rs As Object
    Set rsh = wb.Sheets("TaxReport")
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    Set csh = wb.Sheets("TaxCalculate")
    sql = "SELECT d.TaxID, d.Name, sum(d.Free) from [" & "DATA$A5:D" & R & "] as d GROUP BY d.TaxID, d.Name ORDER BY d.Name, d.TaxID "
    
    cn.Open constr
    rs.Open sql, cn
    rsh.Range("A6").CopyFromRecordset rs
    rs.movefirst
    c = 6
     Do Until rs.EOF
        Debug.Print rs(0), rs(1), rs(2)
        csh.Activate
        Range("INP_NETREV").Value = rs(2)
        rsh.Activate
        Range("D" & c).Value = csh.Range("OUTPUT_TAXPAY").Value
        c = c + 1
        rs.movenext
    Loop

End Sub

Sub ClearReportArea(ByRef wb As Workbook, ByRef sh As Worksheet)
    sh.Activate
    Range("A6:D" & sh.UsedRange.Rows.Count).ClearContents
End Sub
