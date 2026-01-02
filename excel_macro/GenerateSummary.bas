Sub GenerateSummary()
    '
    ' GenerateSummary Macro
    ' Creates or updates the SUMMARY sheet with MRP and Offer Price totals from all sheets
    '
    
    Dim wb As Workbook
    Dim summarySheet As Worksheet
    Dim dataSheet As Worksheet
    Dim lastRow As Long
    Dim totalRow As Long
    Dim sheetName As String
    Dim mrpValue As Double
    Dim offerValue As Double
    Dim summaryRow As Integer
    Dim i As Integer
    
    ' Disable screen updating for faster execution
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    
    Set wb = ActiveWorkbook
    
    ' Create or get SUMMARY sheet
    On Error Resume Next
    Set summarySheet = wb.Sheets("SUMMARY")
    On Error GoTo 0
    
    If summarySheet Is Nothing Then
        ' Create new SUMMARY sheet at the beginning
        Set summarySheet = wb.Sheets.Add(Before:=wb.Sheets(1))
        summarySheet.Name = "SUMMARY"
    Else
        ' Clear existing content
        summarySheet.Cells.Clear
    End If
    
    ' Set up headers
    With summarySheet
        .Range("A1").Value = "Sr. No"
        .Range("B1").Value = "Sheet Name"
        .Range("C1").Value = "MRP"
        .Range("D1").Value = "OFFER PRICE"
        
        ' Format headers
        With .Range("A1:D1")
            .Font.Bold = True
            .Font.Name = "Bookman Old Style"
            .Font.Size = 12
            .Interior.Color = RGB(255, 199, 206) ' Pink fill
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Set column widths
        .Columns("A:A").ColumnWidth = 10
        .Columns("B:B").ColumnWidth = 25
        .Columns("C:C").ColumnWidth = 15
        .Columns("D:D").ColumnWidth = 15
    End With
    
    summaryRow = 2
    
    ' Loop through all sheets (except SUMMARY)
    For Each dataSheet In wb.Sheets
        sheetName = dataSheet.Name
        
        ' Skip SUMMARY sheet itself
        If UCase(sheetName) <> "SUMMARY" Then
            ' Find TOTAL VALUE row
            totalRow = FindTotalRow(dataSheet)
            
            If totalRow > 0 Then
                ' Get MRP (column G = 7) and Offer Price (column I = 9)
                On Error Resume Next
                mrpValue = dataSheet.Cells(totalRow, 7).Value
                offerValue = dataSheet.Cells(totalRow, 9).Value
                On Error GoTo 0
                
                ' If values are not numbers, try to convert
                If Not IsNumeric(mrpValue) Then mrpValue = 0
                If Not IsNumeric(offerValue) Then offerValue = 0
                
                ' Write to summary sheet
                With summarySheet
                    .Cells(summaryRow, 1).Value = summaryRow - 1 ' Sr. No
                    .Cells(summaryRow, 1).HorizontalAlignment = xlCenter
                    .Cells(summaryRow, 1).Font.Name = "Cambria"
                    
                    .Cells(summaryRow, 2).Value = sheetName
                    .Cells(summaryRow, 2).Font.Name = "Cambria"
                    
                    .Cells(summaryRow, 3).Value = mrpValue
                    .Cells(summaryRow, 3).NumberFormat = "#,##0"
                    .Cells(summaryRow, 3).HorizontalAlignment = xlRight
                    .Cells(summaryRow, 3).Font.Name = "Cambria"
                    
                    .Cells(summaryRow, 4).Value = offerValue
                    .Cells(summaryRow, 4).NumberFormat = "#,##0"
                    .Cells(summaryRow, 4).HorizontalAlignment = xlRight
                    .Cells(summaryRow, 4).Font.Name = "Cambria"
                    
                    ' Add borders
                    .Range(.Cells(summaryRow, 1), .Cells(summaryRow, 4)).Borders.LineStyle = xlContinuous
                End With
                
                summaryRow = summaryRow + 1
            End If
        End If
    Next dataSheet
    
    ' Add TOTAL MRP row
    If summaryRow > 2 Then
        With summarySheet
            .Cells(summaryRow, 2).Value = "TOTAL MRP"
            .Cells(summaryRow, 2).Font.Bold = True
            .Cells(summaryRow, 2).Font.Name = "Bookman Old Style"
            .Cells(summaryRow, 2).Font.Size = 12
            .Cells(summaryRow, 2).Interior.Color = RGB(255, 199, 206)
            
            .Cells(summaryRow, 3).Formula = "=SUM(C2:C" & summaryRow - 1 & ")"
            .Cells(summaryRow, 3).NumberFormat = "#,##0"
            .Cells(summaryRow, 3).Font.Bold = True
            .Cells(summaryRow, 3).Font.Name = "Bookman Old Style"
            .Cells(summaryRow, 3).Font.Size = 12
            .Cells(summaryRow, 3).Interior.Color = RGB(255, 199, 206)
            .Cells(summaryRow, 3).HorizontalAlignment = xlRight
            
            .Cells(summaryRow, 4).Formula = "=SUM(D2:D" & summaryRow - 1 & ")"
            .Cells(summaryRow, 4).NumberFormat = "#,##0"
            .Cells(summaryRow, 4).Font.Bold = True
            .Cells(summaryRow, 4).Font.Name = "Bookman Old Style"
            .Cells(summaryRow, 4).Font.Size = 12
            .Cells(summaryRow, 4).Interior.Color = RGB(255, 199, 206)
            .Cells(summaryRow, 4).HorizontalAlignment = xlRight
            
            .Range(.Cells(summaryRow, 1), .Cells(summaryRow, 4)).Borders.LineStyle = xlContinuous
        End With
        
        summaryRow = summaryRow + 2
        
        ' Add FINAL OFFER VALUE row
        With summarySheet
            .Cells(summaryRow, 2).Value = "FINAL OFFER VALUE ( INCL GST )"
            .Cells(summaryRow, 2).Font.Bold = True
            .Cells(summaryRow, 2).Font.Name = "Bookman Old Style"
            .Cells(summaryRow, 2).Font.Size = 12
            .Cells(summaryRow, 2).Interior.Color = RGB(255, 199, 206)
            
            .Cells(summaryRow, 3).Formula = "=D" & summaryRow - 2
            .Cells(summaryRow, 3).NumberFormat = "#,##0"
            .Cells(summaryRow, 3).Font.Bold = True
            .Cells(summaryRow, 3).Font.Name = "Bookman Old Style"
            .Cells(summaryRow, 3).Font.Size = 12
            .Cells(summaryRow, 3).Interior.Color = RGB(255, 199, 206)
            .Cells(summaryRow, 3).HorizontalAlignment = xlRight
            
            .Range("B" & summaryRow & ":C" & summaryRow).Borders.LineStyle = xlContinuous
        End With
        
        ' Add Terms and Conditions section
        summaryRow = summaryRow + 3
        
        With summarySheet
            ' Terms and Conditions Header
            .Cells(summaryRow, 2).Value = "Terms and Conditions"
            .Cells(summaryRow, 2).Font.Bold = True
            .Cells(summaryRow, 2).Font.Name = "Tw Cen MT"
            .Cells(summaryRow, 2).Interior.Color = RGB(255, 199, 206)
            .Range("B" & summaryRow & ":D" & summaryRow).Merge True
            .Cells(summaryRow, 2).HorizontalAlignment = xlCenter
            
            summaryRow = summaryRow + 1
            
            ' Terms and Conditions content
            Dim terms(1 To 9) As String
            terms(1) = "1. Above Prices are inclusive of GST."
            terms(2) = "2. Prices are valid for 15 days from the date of the quotation."
            terms(3) = "3. Transportation, Unloading & Lifting Charges additional as per actual."
            terms(4) = "4. Any damage material have to reported within 3 days of the material delivery."
            terms(5) = "5. Payment Terms 100 % advance against the PO."
            terms(6) = "6. Once the purchase order released, any changes or modifications cannot be accepted."
            terms(7) = "7. Unloading of material at the site is buyer's scope and not our responsibility."
            terms(8) = "8. In case of any complaint, the matter should be taken up with the manufacturer directly."
            terms(9) = "9. For additional products the customer has to pay as per the present price."
            
            For i = 1 To 9
                .Cells(summaryRow, 2).Value = terms(i)
                .Cells(summaryRow, 2).Font.Name = "Tw Cen MT"
                .Cells(summaryRow, 2).Font.Size = 11
                .Range("B" & summaryRow & ":D" & summaryRow).Merge True
                .Cells(summaryRow, 2).WrapText = True
                summaryRow = summaryRow + 1
            Next i
            
            summaryRow = summaryRow + 1
            
            ' TRANSPORTATION NOT INCLUDED
            .Cells(summaryRow, 2).Value = "TRANSPORTATION NOT INCLUDED"
            .Cells(summaryRow, 2).Font.Bold = True
            .Cells(summaryRow, 2).Font.Name = "Tw Cen MT"
            .Cells(summaryRow, 2).Font.Size = 11
            .Cells(summaryRow, 2).Interior.Color = RGB(255, 199, 206)
            .Range("B" & summaryRow & ":D" & summaryRow).Merge True
            .Cells(summaryRow, 2).HorizontalAlignment = xlCenter
        End With
    End If
    
    ' Activate SUMMARY sheet
    summarySheet.Activate
    
    Application.ScreenUpdating = True
    
    MsgBox "Summary sheet generated successfully!", vbInformation, "Success"
    
End Sub

Function FindTotalRow(ws As Worksheet) As Long
    '
    ' Finds the row containing "TOTAL VALUE" or "TOTAL" in the sheet
    '
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    For i = 1 To lastRow
        cellValue = UCase(Trim(ws.Cells(i, 3).Value))
        
        If InStr(cellValue, "TOTAL") > 0 Then
            FindTotalRow = i
            Exit Function
        End If
    Next i
    
    FindTotalRow = 0
End Function
