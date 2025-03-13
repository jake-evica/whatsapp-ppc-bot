Option Explicit

Sub processData()
    Dim sheet_from As Worksheet
    Dim sheet_to As Worksheet
    Dim asinSheet As Worksheet
    Dim aovCollection As Collection
    Dim last_row As Long, last_row_campaign As Long
    Dim i As Long, j As Long
    Dim result As Variant
    Dim sales As Double, clicks As Double, impressions As Double, orders As Double, spend As Double
    Dim currentAOV As Double, adGroupID As String, asin As String
    Dim targetACOS As Double
    Dim increaseSpend As Integer
    Dim newWb As Workbook
    Dim newWs As Worksheet
   
    ' Popup for Target ACOS at the start
    Dim userInput As String
    userInput = InputBox("Please enter the Target ACOS (e.g., 30%)", "Target ACOS")
    If userInput = "" Then Exit Sub ' Cancel pressed
    If Right(userInput, 1) = "%" Then userInput = Left(userInput, Len(userInput) - 1)
    If Not IsNumeric(userInput) Then
        MsgBox "Invalid input. Please enter a numeric value (e.g., 30 or 30%).", vbCritical
        Exit Sub
    End If
    targetACOS = CDbl(userInput) / 100 ' Convert to decimal (e.g., "30" ? 0.3)
   
    ' Second popup for increasing spend on low spend keywords
    increaseSpend = MsgBox("Do you want to increase spend on low spend keywords?", vbYesNo, "Increase Spend")
   
    ' Set input sheet (second sheet)
    Set sheet_from = Sheets(2)
   
    ' Delete all sheets after the second one, except "ASIN Data" and "Config"
    Application.DisplayAlerts = False
    Debug.Print "Sheets before deletion:"
    For i = 1 To Sheets.Count
        Debug.Print "Sheet " & i & ": " & Sheets(i).Name
    Next i
    For i = Sheets.Count To 3 Step -1
        If Sheets(i).Name <> "ASIN Data" And Sheets(i).Name <> "Config" Then
            Debug.Print "Deleting sheet: " & Sheets(i).Name
            Sheets(i).Delete
        Else
            Debug.Print "Preserving sheet: " & Sheets(i).Name
        End If
    Next i
    Application.DisplayAlerts = True
   
    ' Create output sheet
    sheet_from.Copy After:=sheet_from
    Set sheet_to = ActiveSheet
    sheet_to.Name = "Output"
   
    ' Freeze panes at C2
    sheet_to.Range("C2").Select
    ActiveWindow.FreezePanes = True
   
    ' Insert new columns AC:AL (10 columns)
    sheet_to.Range("AC1:AL1").EntireColumn.Insert
    sheet_to.Range("AC1:AL1").Value = Array("New Bid.", "CPC", "tCPC", "ACOS", "Spend", "ACTC", "Orders", "AOV", "% of AOV", "RPC")
   
    ' Load ASIN Data into collection
    On Error Resume Next
    Set asinSheet = Sheets("ASIN Data")
    On Error GoTo 0
    If asinSheet Is Nothing Then
        MsgBox "Please create an 'ASIN Data' sheet with ASINs in column A and AOVs in column B.", vbCritical
        Exit Sub
    End If
   
    Set aovCollection = New Collection
    With asinSheet
        last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
        For i = 2 To last_row
            If Len(Trim(.Cells(i, "A").Value)) > 0 Then
                On Error Resume Next ' Handle duplicate ASINs by overwriting
                aovCollection.Add CDbl(Val(.Cells(i, "B").Value)), Trim(.Cells(i, "A").Value)
                On Error GoTo 0
            End If
        Next i
    End With
   
    ' Process data
    last_row = sheet_to.Range("A1048576").End(xlUp).Row
    last_row_campaign = 0
    ReDim result(2 To last_row, 1 To 10)
   
    For i = 2 To last_row
        With sheet_to
            ' Copy values from source (post-insertion columns)
            result(i, 2) = .Range("BE" & i).Value ' CPC (AU/47 ? BE/57)
            result(i, 4) = .Range("BD" & i).Value ' ACOS (AT/46 ? BD/56)
            result(i, 5) = .Range("AY" & i).Value ' Spend (AO/41 ? AY/51)
            result(i, 7) = .Range("BA" & i).Value ' Orders (AQ/43 ? BA/53)
           
            ' Track last campaign row
            If .Range("B" & i).Value = "Campaign" Then
                last_row_campaign = i
            End If
           
            ' Get source values for calculations (post-insertion columns)
            impressions = CDbl(Val(.Range("AV" & i).Value)) ' Impressions (AL/38 ? AV/48)
            clicks = CDbl(Val(.Range("AW" & i).Value))     ' Clicks (AM/39 ? AW/49)
            spend = CDbl(Val(.Range("AY" & i).Value))      ' Spend (AO/41 ? AY/51)
            sales = CDbl(Val(.Range("AZ" & i).Value))      ' Sales (AP/42 ? AZ/52)
            orders = CDbl(Val(.Range("BA" & i).Value))     ' Orders (AQ/43 ? BA/53)
           
            ' Calculate ACTC (Clicks / Orders)
            If orders <> 0 Then
                result(i, 6) = clicks / orders ' AW / BA
            ElseIf last_row_campaign > 0 Then
                result(i, 6) = result(last_row_campaign, 6)
            Else
                result(i, 6) = "N/A"
            End If
           
            ' Calculate AOV: Primary method (Sales / Orders), Fallback (ASIN lookup)
            If orders <> 0 And sales <> 0 Then
                currentAOV = sales / orders ' Primary method
            Else
                ' Fallback: Find ASIN via Ad Group ID
                adGroupID = .Range("E" & i).Value
                asin = ""
                For j = 2 To last_row
                    If .Range("E" & j).Value = adGroupID And Len(Trim(.Range("W" & j).Value)) > 0 Then
                        asin = Trim(.Range("W" & j).Value)
                        Exit For
                    End If
                Next j
                ' Lookup AOV from ASIN Data
                currentAOV = 0
                If Len(asin) > 0 Then
                    On Error Resume Next
                    currentAOV = aovCollection(asin)
                    On Error GoTo 0
                End If
            End If
            result(i, 8) = IIf(currentAOV > 0, currentAOV, "N/A") ' AOV
           
            ' Calculate % of AOV (Spend / AOV)
            If currentAOV > 0 Then
                result(i, 9) = spend / currentAOV ' Spend / AOV as a decimal (e.g., 0.5 for 50%)
            Else
                result(i, 9) = "N/A"
            End If
           
            ' Calculate RPC (Sales / Clicks)
            If clicks <> 0 Then
                result(i, 10) = sales / clicks ' AZ / AW
            ElseIf last_row_campaign > 0 Then
                result(i, 10) = result(last_row_campaign, 10)
            Else
                result(i, 10) = "N/A"
            End If
        End With
    Next i
   
    ' Output results to AC:AL
    With sheet_to
        .Range("AC2:AL" & last_row).Value = result
       
        ' Check AB post-insertion and write AA if empty
        For i = 2 To last_row
            If IsEmpty(.Range("AB" & i).Value) Or .Range("AB" & i).Value = "" Then
                .Range("AB" & i).Value = .Range("AA" & i).Value
            End If
        Next i
       
        .Range("AF2:AF" & last_row).NumberFormat = "0.00%" ' ACOS
        .Range("AK2:AK" & last_row).NumberFormat = "0.00%" ' % of AOV
       
        ' Sort by inserted ACOS column (AF)
        .Range("A1:BE" & last_row).Sort Key1:=.Range("AF1"), Order1:=xlDescending, Header:=xlYes
        .Range("A1:BE1").EntireColumn.AutoFit
       
        ' Apply filter on column B
        .Range("A1:BE" & last_row).AutoFilter Field:=2, Criteria1:=Array("Keyword", "Product Targeting"), Operator:=xlFilterValues
       
        ' Highlight rows based on Target ACOS and write "Update" in Column C and New Bid in AC/AE
        For i = 2 To last_row
            Dim acosValue As Double
            Dim ordersValue As Double
            Dim aovPercent As Double
            Dim impressionsValue As Double
            Dim isHighlighted As Boolean
           
            ' Get values (handle non-numeric cases)
            If IsNumeric(.Range("AF" & i).Value) Then acosValue = .Range("AF" & i).Value Else acosValue = 0
            If IsNumeric(.Range("AI" & i).Value) Then ordersValue = .Range("AI" & i).Value Else ordersValue = 0
            If IsNumeric(.Range("AK" & i).Value) Then aovPercent = .Range("AK" & i).Value Else aovPercent = 0
            If IsNumeric(.Range("AV" & i).Value) Then impressionsValue = .Range("AV" & i).Value Else impressionsValue = 0
           
            ' Reset highlight flag
            isHighlighted = False
           
            ' Condition 1: ACOS ≥ Target ACOS + 10% of Target ACOS (Light Orange)
            If acosValue >= (targetACOS + (targetACOS * 0.1)) Then
                .Range("A" & i & ":BE" & i).Interior.Color = RGB(255, 204, 153)
                If IsNumeric(.Range("AL" & i).Value) Then ' Check if RPC is numeric
                    .Range("AE" & i).Value = .Range("AL" & i).Value * targetACOS ' RPC x Target ACOS to AE
                    If IsNumeric(.Range("AE" & i).Value) And IsNumeric(.Range("AB" & i).Value) And IsNumeric(.Range("AD" & i).Value) And .Range("AD" & i).Value <> 0 Then
                        .Range("AC" & i).Value = .Range("AE" & i).Value * (.Range("AB" & i).Value / .Range("AD" & i).Value) ' AE * (AB / AD) to AC
                    End If
                End If
                isHighlighted = True
           
            ' Condition 2: ACOS ≤ Target ACOS - 10% of Target ACOS AND Orders > 1 (Light Green)
            ElseIf acosValue <= (targetACOS - (targetACOS * 0.1)) And ordersValue > 1 Then
                .Range("A" & i & ":BE" & i).Interior.Color = RGB(144, 238, 144)
                If IsNumeric(.Range("AB" & i).Value) Then ' Check if AB is numeric
                    If acosValue <= (targetACOS * 0.5) Then ' 50% below Target ACOS
                        .Range("AC" & i).Value = .Range("AB" & i).Value * 1.15 ' AB * 1.15
                    Else
                        .Range("AC" & i).Value = .Range("AB" & i).Value * 1.1  ' AB * 1.1
                    End If
                End If
                isHighlighted = True
           
            ' Condition 3: ACOS ≤ Target ACOS - 10% of Target ACOS AND Orders = 1 (Lighter Green)
            ElseIf acosValue <= (targetACOS - (targetACOS * 0.1)) And ordersValue = 1 Then
                .Range("A" & i & ":BE" & i).Interior.Color = RGB(50, 205, 50)
                If IsNumeric(.Range("AB" & i).Value) Then ' Check if AB is numeric
                    If acosValue <= (targetACOS * 0.5) Then ' 50% below Target ACOS
                        .Range("AC" & i).Value = .Range("AB" & i).Value * 1.06 ' AB * 1.06
                    Else
                        .Range("AC" & i).Value = .Range("AB" & i).Value * 1.05 ' AB * 1.05
                    End If
                End If
                isHighlighted = True
           
            ' Condition 4: ACOS = 0% AND % of AOV ≥ Target ACOS - 10% of Target ACOS (Darker Orange)
            ElseIf acosValue = 0 And aovPercent >= (targetACOS - (targetACOS * 0.1)) Then
                .Range("A" & i & ":BE" & i).Interior.Color = RGB(255, 140, 0)
                If IsNumeric(.Range("AJ" & i).Value) And IsNumeric(.Range("AH" & i).Value) And IsNumeric(.Range("AW" & i).Value) And (.Range("AH" & i).Value + .Range("AW" & i).Value) <> 0 Then
                    .Range("AE" & i).Value = .Range("AJ" & i).Value / (.Range("AH" & i).Value + .Range("AW" & i).Value) * targetACOS ' AJ / (AH + AW) * Target ACOS to AE
                    If IsNumeric(.Range("AE" & i).Value) And IsNumeric(.Range("AB" & i).Value) And IsNumeric(.Range("AD" & i).Value) And .Range("AD" & i).Value <> 0 Then
                        .Range("AC" & i).Value = .Range("AE" & i).Value * (.Range("AB" & i).Value / .Range("AD" & i).Value) ' AE * (AB / AD) to AC
                    End If
                End If
                isHighlighted = True
           
            ' Condition 5: ACOS = 0% AND % of AOV ≤ 10% AND Impressions ≥ 0.3% (Light Blue, conditional)
            ElseIf increaseSpend = vbYes And acosValue = 0 And aovPercent <= 0.1 And impressionsValue >= 0.003 Then
                .Range("A" & i & ":BE" & i).Interior.Color = RGB(173, 216, 230)
                If IsNumeric(.Range("AB" & i).Value) Then ' Check if AB is numeric
                    .Range("AC" & i).Value = .Range("AB" & i).Value * 1.05 ' AB * 1.05 to AC
                End If
                isHighlighted = True
            End If
           
            ' Write "Update" in Column C if row is highlighted
            If isHighlighted Then
                .Range("C" & i).Value = "Update"
            End If
        Next i
       
        ' Step 1: Delete rows without "Update" in Column C
        Application.DisplayAlerts = False
        For i = last_row To 2 Step -1 ' Work backwards to avoid skipping rows
            If .Range("C" & i).Value <> "Update" Then
                .Rows(i).Delete
            End If
        Next i
        Application.DisplayAlerts = True
       
        ' Update last_row after deletions and check if any rows remain
        last_row = .Range("A" & .Rows.Count).End(xlUp).Row
        If last_row <= 1 Then ' Only headers remain or sheet is empty
            MsgBox "No rows marked for update. Process completed with no changes.", vbInformation
            Exit Sub
        End If
       
        ' Step 2: Rename column AB to "Old Bid"
        .Range("AB1").Value = "Old Bid"
       
        ' Step 3: Rename column AC to "Bid"
        .Range("AC1").Value = "Bid"
       
        ' Ensure columns are still auto-fitted after changes
        .Range("A1:BE1").EntireColumn.AutoFit
       
        ' Step 4: Create new workbook and paste data (no save)
        Set newWb = Workbooks.Add ' Create new workbook
        Set newWs = newWb.Sheets(1) ' Use first sheet in new workbook
        .UsedRange.Copy Destination:=newWs.Range("A1") ' Copy all data from Output to new workbook
        newWs.Range("A1:BE1").EntireColumn.AutoFit ' Auto-fit columns in new workbook
        newWs.Name = "Bid Optimization " & Format(Now, "MMDDYY") ' Name the sheet (e.g., "Bid Optimization 030625")
    End With
   
    ' Final Step: Show completion popup
    MsgBox "Bids are optimized, please check and delete all new columns.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub dsda()
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = Sheets("Config")
    On Error GoTo 0
    If Not sheet Is Nothing Then
        sheet.Visible = xlSheetVisible
    Else
        MsgBox "Sheet 'Config' not found. Please create it if needed.", vbInformation
    End If
End Sub