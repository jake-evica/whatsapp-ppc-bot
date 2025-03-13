Sub GenerateAmazonKeywords()
    Dim wsSearch As Worksheet, wsSponsored As Worksheet, wsAsinList As Worksheet
    Dim wsHigh As Worksheet, wsHighReview As Worksheet
    Dim wsLow As Worksheet, wsLowReview As Worksheet
    Dim wsProductTargets As Worksheet, wsProductTargetsReview As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim acosThreshold As Double
    Dim searchTerm As String, keywordText As String, campaignId As String, sku As String
    Dim orders As Long
    Dim skuKeywords As Collection ' Regular Keywords
    Dim targetingKeywords As Collection ' b0 Terms
    Dim skuKeys As Collection     ' Keys for skuKeywords
    Dim targetingKeys As Collection ' Keys for targetingKeywords
    Dim skuKey As String
    Dim brandNames() As String
    Dim targetWs As Worksheet
    Dim campaignCount As Long
    Dim keywordCount As Long
    Dim highRow As Long, highReviewRow As Long, lowRow As Long, lowReviewRow As Long, b0Row As Long, b0ReviewRow As Long
    Dim rowCount As Long
    Dim newId As String
    Dim iKey As Long
    Dim targetingKeywordsColl As Collection
    Dim exclude As Boolean
    Dim keywordInfo As Variant
    Dim targetCollection As Collection
    Dim targetKeys As Collection
    Dim columnN As String
    Dim skuCollection As Collection
    Dim brandInput As String
    Dim currentDate As String
    Dim newWb As Workbook ' New workbook for output
    Dim asinLastRow As Long
    Dim isOwnProduct As Boolean
    Dim matchType As String
   
    ' Set worksheets from current workbook
    Set wsSearch = ThisWorkbook.Sheets("SP Search Term Report")
    Set wsSponsored = ThisWorkbook.Sheets("Sponsored Products Campaigns")
    Set wsAsinList = ThisWorkbook.Sheets("ASIN list")
    Set skuKeywords = New Collection
    Set targetingKeywords = New Collection
    Set skuKeys = New Collection
    Set targetingKeys = New Collection
   
    ' Get ACOS threshold
    acosThreshold = CDbl(InputBox("Enter max ACOS % (e.g., 30):", "ACOS Threshold")) / 100
   
    ' Get brand names to exclude
    brandInput = InputBox("Enter brand names to exclude (comma-separated, e.g., Nike, Adidas):", "Brand Exclusion")
    If brandInput = "" Then
        brandNames = Split("", ",")
    Else
        brandNames = Split(Trim(brandInput), ",")
        For i = 0 To UBound(brandNames)
            brandNames(i) = " " & LCase(Trim(brandNames(i))) & " "
        Next i
    End If
   
    ' Get match type
    Do
        matchType = InputBox("What match type do you want to create? (Exact, Phrase, or Broad):", "Match Type Selection")
        matchType = LCase(Trim(matchType)) ' Convert to lowercase for case-insensitive comparison
        If matchType = "exact" Or matchType = "phrase" Or matchType = "broad" Then Exit Do
        MsgBox "Please enter a valid match type: Exact, Phrase, or Broad.", vbExclamation
    Loop
   
    ' Get current date in YYYYMMDD format
    currentDate = Format(Date, "YYYYMMDD")
   
    ' Step 1: Collect keywords by SKU
    lastRow = wsSearch.Cells(wsSearch.Rows.count, "P").End(xlUp).Row
    asinLastRow = wsAsinList.Cells(wsAsinList.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
        searchTerm = Trim(wsSearch.Cells(i, "P").Value)
        orders = wsSearch.Cells(i, "V").Value
        If orders >= 1 And wsSearch.Cells(i, "Y").Value < acosThreshold Then
            keywordText = " " & LCase(wsSearch.Cells(i, "L").Value) & " "
            If InStr(1, keywordText, " " & LCase(searchTerm) & " ") = 0 Then
                exclude = False
                For j = 0 To UBound(brandNames)
                    If InStr(1, " " & LCase(searchTerm) & " ", brandNames(j)) > 0 Then
                        exclude = True
                        Exit For
                    End If
                Next j
                If Not exclude Then
                    campaignId = wsSearch.Cells(i, "B").Value
                    sku = GetSKU(wsSponsored, campaignId)
                    ' Include bid value from column Z and ad group name from column F
                    keywordInfo = Array(searchTerm, orders, wsSearch.Cells(i, "Z").Value, wsSearch.Cells(i, "F").Value)
                    If LCase(Left(searchTerm, 2)) = "b0" Then
                        ' Check if this is one of our own products in Column A of ASIN list (case-insensitive)
                        isOwnProduct = False
                        For j = 2 To asinLastRow
                            If UCase(Trim(wsAsinList.Cells(j, "A").Value)) = UCase(searchTerm) Then
                                isOwnProduct = True
                                Exit For
                            End If
                        Next j
                        If Not isOwnProduct Then
                            skuKey = sku
                            Set targetCollection = targetingKeywords
                            Set targetKeys = targetingKeys
                            On Error Resume Next
                            Set skuCollection = targetCollection(skuKey)
                            If Err.Number <> 0 Then
                                Set skuCollection = New Collection
                                targetCollection.Add skuCollection, skuKey
                                targetKeys.Add skuKey
                            End If
                            On Error GoTo 0
                            skuCollection.Add keywordInfo
                        End If
                    Else
                        skuKey = sku & "|" & orders
                        Set targetCollection = skuKeywords
                        Set targetKeys = skuKeys
                        On Error Resume Next
                        Set skuCollection = targetCollection(skuKey)
                        If Err.Number <> 0 Then
                            Set skuCollection = New Collection
                            targetCollection.Add skuCollection, skuKey
                            targetKeys.Add skuKey
                        End If
                        On Error GoTo 0
                        skuCollection.Add keywordInfo
                    End If
                End If
            End If
        End If
    Next i
   
    ' Create new workbook for output
    Set newWb = Workbooks.Add
    ' Set up all six sheets
    With newWb
        Set wsHigh = .Sheets(1)
        wsHigh.Name = "3+ Orders"
        Set wsHighReview = .Sheets.Add(After:=wsHigh)
        wsHighReview.Name = "3+ Orders - Review"
        Set wsLow = .Sheets.Add(After:=wsHighReview)
        wsLow.Name = "1-2 Orders"
        Set wsLowReview = .Sheets.Add(After:=wsLow)
        wsLowReview.Name = "1-2 Orders - Review"
        Set wsProductTargets = .Sheets.Add(After:=wsLowReview)
        wsProductTargets.Name = "Product Targets"
        Set wsProductTargetsReview = .Sheets.Add(After:=wsProductTargets)
        wsProductTargetsReview.Name = "Product Targets - Review"
    End With
   
    ' Set headers - Review tabs include "Review Needed", non-review tabs exclude it
    With wsHighReview
        .Cells(1, 1) = "Review Needed": .Cells(1, 2) = "Product": .Cells(1, 3) = "Entity"
        .Cells(1, 4) = "Operation": .Cells(1, 5) = "Campaign ID": .Cells(1, 6) = "Ad Group ID"
        .Cells(1, 7) = "Portfolio ID": .Cells(1, 8) = "Ad ID": .Cells(1, 9) = "Keyword ID"
        .Cells(1, 10) = "Product Targeting ID": .Cells(1, 11) = "Campaign Name": .Cells(1, 12) = "Ad Group Name"
        .Cells(1, 13) = "Start Date": .Cells(1, 14) = "End Date": .Cells(1, 15) = "Targeting Type"
        .Cells(1, 16) = "State": .Cells(1, 17) = "Daily Budget": .Cells(1, 18) = "SKU"
        .Cells(1, 19) = "Ad Group Default Bid": .Cells(1, 20) = "Bid": .Cells(1, 21) = "Keyword Text"
        .Cells(1, 22) = "Native Language Keyword": .Cells(1, 23) = "Native Language Locale"
        .Cells(1, 24) = "Match Type": .Cells(1, 25) = "Bidding Strategy": .Cells(1, 26) = "Placement"
        .Cells(1, 27) = "Percentage": .Cells(1, 28) = "Product Targeting Expression"
    End With
    wsLowReview.Range("A1:AB1") = wsHighReview.Range("A1:AB1").Value
    wsProductTargetsReview.Range("A1:AB1") = wsHighReview.Range("A1:AB1").Value
   
    With wsHigh
        .Cells(1, 1) = "Product": .Cells(1, 2) = "Entity"
        .Cells(1, 3) = "Operation": .Cells(1, 4) = "Campaign ID": .Cells(1, 5) = "Ad Group ID"
        .Cells(1, 6) = "Portfolio ID": .Cells(1, 7) = "Ad ID": .Cells(1, 8) = "Keyword ID"
        .Cells(1, 9) = "Product Targeting ID": .Cells(1, 10) = "Campaign Name": .Cells(1, 11) = "Ad Group Name"
        .Cells(1, 12) = "Start Date": .Cells(1, 13) = "End Date": .Cells(1, 14) = "Targeting Type"
        .Cells(1, 15) = "State": .Cells(1, 16) = "Daily Budget": .Cells(1, 17) = "SKU"
        .Cells(1, 18) = "Ad Group Default Bid": .Cells(1, 19) = "Bid": .Cells(1, 20) = "Keyword Text"
        .Cells(1, 21) = "Native Language Keyword": .Cells(1, 22) = "Native Language Locale"
        .Cells(1, 23) = "Match Type": .Cells(1, 24) = "Bidding Strategy": .Cells(1, 25) = "Placement"
        .Cells(1, 26) = "Percentage": .Cells(1, 27) = "Product Targeting Expression"
    End With
    wsLow.Range("A1:AA1") = wsHigh.Range("A1:AA1").Value
    wsProductTargets.Range("A1:AA1") = wsHigh.Range("A1:AA1").Value
   
    ' Step 2: Write regular keywords
    highRow = 2: highReviewRow = 2: lowRow = 2: lowReviewRow = 2: b0Row = 2: b0ReviewRow = 2
    For i = 1 To skuKeys.count
        skuKey = skuKeys(i)
        Set keywords = skuKeywords(skuKey)
        sku = Split(skuKey, "|")(0)
        orders = CLng(Split(skuKey, "|")(1))
        Set targetWs = IIf(orders >= 3, IIf(sku = "Multi ASIN", wsHighReview, wsHigh), IIf(sku = "Multi ASIN", wsLowReview, wsLow))
       
        campaignCount = 0
        keywordCount = 0
       
        For iKey = 1 To keywords.count
            If keywordCount Mod 10 = 0 Then
                campaignCount = campaignCount + 1
                keywordCount = 0
               
                rowCount = IIf(orders >= 3, IIf(sku = "Multi ASIN", highReviewRow, highRow), IIf(sku = "Multi ASIN", lowReviewRow, lowRow))
                targetWs.Rows(rowCount & ":" & rowCount + 2).Insert Shift:=xlDown
                newId = sku & " - SP " & matchType & " - " & campaignCount ' Update campaign name with match type
               
                With targetWs
                    If sku = "Multi ASIN" Then
                        ' Review tabs - keep Column A
                        .Cells(rowCount, 2) = "Sponsored Products"
                        .Cells(rowCount, 3) = "Campaign"
                        .Cells(rowCount, 4) = "Create"
                        .Cells(rowCount, 5) = newId
                        .Cells(rowCount, 11) = newId
                        .Cells(rowCount, 13) = currentDate
                        .Cells(rowCount, 15) = "MANUAL"
                        .Cells(rowCount, 16) = "enabled"
                        .Cells(rowCount, 17) = 10
                        .Cells(rowCount, 25) = "Dynamic bids - down only"
                        .Cells(rowCount, 1) = "Yes - " & keywords(iKey)(3)
                       
                        .Cells(rowCount + 1, 2) = "Sponsored Products"
                        .Cells(rowCount + 1, 3) = "Ad Group"
                        .Cells(rowCount + 1, 4) = "Create"
                        .Cells(rowCount + 1, 5) = newId
                        .Cells(rowCount + 1, 6) = newId
                        .Cells(rowCount + 1, 12) = newId
                        .Cells(rowCount + 1, 16) = "enabled"
                        .Cells(rowCount + 1, 19) = 1
                        .Cells(rowCount + 1, 1) = "Yes - " & keywords(iKey)(3)
                       
                        .Cells(rowCount + 2, 2) = "Sponsored Products"
                        .Cells(rowCount + 2, 3) = "Product Ad"
                        .Cells(rowCount + 2, 4) = "Create"
                        .Cells(rowCount + 2, 5) = newId
                        .Cells(rowCount + 2, 6) = newId
                        .Cells(rowCount + 2, 16) = "enabled"
                        .Cells(rowCount + 2, 18) = sku
                        .Cells(rowCount + 2, 1) = "Yes - " & keywords(iKey)(3)
                    Else
                        ' Non-review tabs - shift left, no Column A
                        .Cells(rowCount, 1) = "Sponsored Products"
                        .Cells(rowCount, 2) = "Campaign"
                        .Cells(rowCount, 3) = "Create"
                        .Cells(rowCount, 4) = newId
                        .Cells(rowCount, 10) = newId
                        .Cells(rowCount, 12) = currentDate
                        .Cells(rowCount, 14) = "MANUAL"
                        .Cells(rowCount, 15) = "enabled"
                        .Cells(rowCount, 16) = 10
                        .Cells(rowCount, 24) = "Dynamic bids - down only"
                       
                        .Cells(rowCount + 1, 1) = "Sponsored Products"
                        .Cells(rowCount + 1, 2) = "Ad Group"
                        .Cells(rowCount + 1, 3) = "Create"
                        .Cells(rowCount + 1, 4) = newId
                        .Cells(rowCount + 1, 5) = newId
                        .Cells(rowCount + 1, 11) = newId
                        .Cells(rowCount + 1, 15) = "enabled"
                        .Cells(rowCount + 1, 18) = 1
                       
                        .Cells(rowCount + 2, 1) = "Sponsored Products"
                        .Cells(rowCount + 2, 2) = "Product Ad"
                        .Cells(rowCount + 2, 3) = "Create"
                        .Cells(rowCount + 2, 4) = newId
                        .Cells(rowCount + 2, 5) = newId
                        .Cells(rowCount + 2, 15) = "enabled"
                        .Cells(rowCount + 2, 17) = sku
                    End If
                End With
               
                If orders >= 3 Then
                    If sku = "Multi ASIN" Then highReviewRow = highReviewRow + 3 Else highRow = highRow + 3
                Else
                    If sku = "Multi ASIN" Then lowReviewRow = lowReviewRow + 3 Else lowRow = lowRow + 3
                End If
            End If
           
            With targetWs
                rowCount = IIf(orders >= 3, IIf(sku = "Multi ASIN", highReviewRow, highRow), IIf(sku = "Multi ASIN", lowReviewRow, lowRow))
                If sku = "Multi ASIN" Then
                    ' Review tabs - keep Column A
                    .Cells(rowCount, 2) = "Sponsored Products"
                    .Cells(rowCount, 3) = "Keyword"
                    .Cells(rowCount, 4) = "Create"
                    .Cells(rowCount, 5) = newId
                    .Cells(rowCount, 6) = newId
                    .Cells(rowCount, 16) = "enabled"
                    .Cells(rowCount, 20) = keywords(iKey)(2) ' Use bid from column Z
                    .Cells(rowCount, 21) = keywords(iKey)(0)
                    .Cells(rowCount, 24) = matchType ' Use user-selected match type
                    .Cells(rowCount, 1) = "Yes - " & keywords(iKey)(3)
                Else
                    ' Non-review tabs - shift left, no Column A
                    .Cells(rowCount, 1) = "Sponsored Products"
                    .Cells(rowCount, 2) = "Keyword"
                    .Cells(rowCount, 3) = "Create"
                    .Cells(rowCount, 4) = newId
                    .Cells(rowCount, 5) = newId
                    .Cells(rowCount, 15) = "enabled"
                    .Cells(rowCount, 19) = keywords(iKey)(2) ' Use bid from column Z
                    .Cells(rowCount, 20) = keywords(iKey)(0)
                    .Cells(rowCount, 23) = matchType ' Use user-selected match type
                End If
                If orders >= 3 Then
                    If sku = "Multi ASIN" Then highReviewRow = highReviewRow + 1 Else highRow = highRow + 1
                Else
                    If sku = "Multi ASIN" Then lowReviewRow = lowReviewRow + 1 Else lowRow = lowRow + 1
                End If
            End With
            keywordCount = keywordCount + 1
        Next iKey
    Next i
   
    ' Step 3: Write targeting terms (b0)
    For i = 1 To targetingKeys.count
        skuKey = targetingKeys(i)
        Set targetingKeywordsColl = targetingKeywords(skuKey)
        sku = skuKey
        Set targetWs = IIf(sku = "Multi ASIN", wsProductTargetsReview, wsProductTargets)
       
        campaignCount = 0
        keywordCount = 0
       
        For iKey = 1 To targetingKeywordsColl.count
            If keywordCount Mod 10 = 0 Then
                campaignCount = campaignCount + 1
                keywordCount = 0
               
                rowCount = IIf(sku = "Multi ASIN", b0ReviewRow, b0Row)
                targetWs.Rows(rowCount & ":" & rowCount + 2).Insert Shift:=xlDown
                newId = sku & " - SP ASIN - " & campaignCount
               
                With targetWs
                    If sku = "Multi ASIN" Then
                        ' Review tabs - keep Column A
                        .Cells(rowCount, 2) = "Sponsored Products"
                        .Cells(rowCount, 3) = "Campaign"
                        .Cells(rowCount, 4) = "Create"
                        .Cells(rowCount, 5) = newId
                        .Cells(rowCount, 11) = newId
                        .Cells(rowCount, 13) = currentDate
                        .Cells(rowCount, 15) = "MANUAL"
                        .Cells(rowCount, 16) = "enabled"
                        .Cells(rowCount, 17) = 10
                        .Cells(rowCount, 25) = "Dynamic bids - down only"
                        .Cells(rowCount, 1) = "Yes - " & targetingKeywordsColl(iKey)(3)
                       
                        .Cells(rowCount + 1, 2) = "Sponsored Products"
                        .Cells(rowCount + 1, 3) = "Ad Group"
                        .Cells(rowCount + 1, 4) = "Create"
                        .Cells(rowCount + 1, 5) = newId
                        .Cells(rowCount + 1, 6) = newId
                        .Cells(rowCount + 1, 12) = newId
                        .Cells(rowCount + 1, 16) = "enabled"
                        .Cells(rowCount + 1, 19) = 1
                        .Cells(rowCount + 1, 1) = "Yes - " & targetingKeywordsColl(iKey)(3)
                       
                        .Cells(rowCount + 2, 2) = "Sponsored Products"
                        .Cells(rowCount + 2, 3) = "Product Ad"
                        .Cells(rowCount + 2, 4) = "Create"
                        .Cells(rowCount + 2, 5) = newId
                        .Cells(rowCount + 2, 6) = newId
                        .Cells(rowCount + 2, 16) = "enabled"
                        .Cells(rowCount + 2, 18) = sku
                        .Cells(rowCount + 2, 1) = "Yes - " & targetingKeywordsColl(iKey)(3)
                    Else
                        ' Non-review tabs - shift left, no Column A
                        .Cells(rowCount, 1) = "Sponsored Products"
                        .Cells(rowCount, 2) = "Campaign"
                        .Cells(rowCount, 3) = "Create"
                        .Cells(rowCount, 4) = newId
                        .Cells(rowCount, 10) = newId
                        .Cells(rowCount, 12) = currentDate
                        .Cells(rowCount, 14) = "MANUAL"
                        .Cells(rowCount, 15) = "enabled"
                        .Cells(rowCount, 16) = 10
                        .Cells(rowCount, 24) = "Dynamic bids - down only"
                       
                        .Cells(rowCount + 1, 1) = "Sponsored Products"
                        .Cells(rowCount + 1, 2) = "Ad Group"
                        .Cells(rowCount + 1, 3) = "Create"
                        .Cells(rowCount + 1, 4) = newId
                        .Cells(rowCount + 1, 5) = newId
                        .Cells(rowCount + 1, 11) = newId
                        .Cells(rowCount + 1, 15) = "enabled"
                        .Cells(rowCount + 1, 18) = 1
                       
                        .Cells(rowCount + 2, 1) = "Sponsored Products"
                        .Cells(rowCount + 2, 2) = "Product Ad"
                        .Cells(rowCount + 2, 3) = "Create"
                        .Cells(rowCount + 2, 4) = newId
                        .Cells(rowCount + 2, 5) = newId
                        .Cells(rowCount + 2, 15) = "enabled"
                        .Cells(rowCount + 2, 17) = sku
                    End If
                End With
               
                If sku = "Multi ASIN" Then b0ReviewRow = b0ReviewRow + 3 Else b0Row = b0Row + 3
            End If
           
            With targetWs
                rowCount = IIf(sku = "Multi ASIN", b0ReviewRow, b0Row)
                If sku = "Multi ASIN" Then
                    ' Review tabs - keep Column A
                    .Cells(rowCount, 2) = "Sponsored Products"
                    .Cells(rowCount, 3) = "Product Targeting"
                    .Cells(rowCount, 4) = "Create"
                    .Cells(rowCount, 5) = newId
                    .Cells(rowCount, 6) = newId
                    .Cells(rowCount, 16) = "enabled"
                    .Cells(rowCount, 20) = targetingKeywordsColl(iKey)(2) ' Use bid from column Z
                    .Cells(rowCount, 28) = "asin=""" & targetingKeywordsColl(iKey)(0) & """"
                    .Cells(rowCount, 1) = "Yes - " & targetingKeywordsColl(iKey)(3)
                Else
                    ' Non-review tabs - shift left, no Column A
                    .Cells(rowCount, 1) = "Sponsored Products"
                    .Cells(rowCount, 2) = "Product Targeting"
                    .Cells(rowCount, 3) = "Create"
                    .Cells(rowCount, 4) = newId
                    .Cells(rowCount, 5) = newId
                    .Cells(rowCount, 15) = "enabled"
                    .Cells(rowCount, 19) = targetingKeywordsColl(iKey)(2) ' Use bid from column Z
                    .Cells(rowCount, 27) = "asin=""" & targetingKeywordsColl(iKey)(0) & """"
                End If
                If sku = "Multi ASIN" Then b0ReviewRow = b0ReviewRow + 1 Else b0Row = b0Row + 1
            End With
            keywordCount = keywordCount + 1
        Next iKey
    Next i
   
    ' Activate the new workbook so user can see it
    newWb.Activate
    MsgBox "Keyword generation complete! Output is in a new workbook with separate review tabs.", vbInformation
End Sub

Function GetSKU(ws As Worksheet, campaignId As String) As String
    Dim lastRow As Long, i As Long, count As Long
    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
    count = 0
    For i = 2 To lastRow
        If ws.Cells(i, "D").Value = campaignId And ws.Cells(i, "B").Value = "Product Ad" Then
            count = count + 1
            If count = 1 Then GetSKU = ws.Cells(i, "V").Value
            If count > 1 Then
                GetSKU = "Multi ASIN"
                Exit For
            End If
        End If
    Next i
    If count = 0 Then GetSKU = "Not Found"
End Function