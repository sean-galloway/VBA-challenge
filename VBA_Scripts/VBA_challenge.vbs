Attribute VB_Name = "VBA_challenge"
Option Explicit

' Worksheet Globals

' Active Workbook Path
Dim gActiveWorkbookPath As String

' Sheet Name Globals
Const gSingle = "SingleSheet"
Const gMulti = "MultiSheet"

' File Related Cells
Const gPathName = "A2"
Const gFileName = "B2"
Const gTabName = "C2"
Const gOutputPathName = "D2"

' Column constants
Const gTickerCol = 1
Const gDateCol = 2
Const gOpenCol = 3
Const gHighCol = 4
Const gLowCol = 5
Const gCloseCol = 6
Const gVolCol = 7

Const gSumTickerCol = 9
Const gSumYearlyChangeCol = 10
Const gSumPercentChangeCol = 11
Const gSumTotalStockVolumeCol = 12

Const gBonusCol = 16

' Start Rows
Const gVBAParseStartRow = 2
Const gVBAWriteStartRow = 2

' This is the Master Routine for opening the selected sheet, parsing it, and returning results
Public Sub Open_Parse_Workbook_SingleSheet()

    Dim strPathName As String
    Dim strFileName As String
    Dim strTabName As String
    Dim strOutputPathName As String
    Dim strFullPath As String
    Dim strFullPathOut As String
    Dim wb As Workbook
    Dim wsActiveSheet As Worksheet
    Dim dctSummary As Scripting.Dictionary
    
    ' Gather the variables
    Sheets(gSingle).Select
    strPathName = Range(gPathName).Value
    strFileName = Range(gFileName).Value
    strTabName = Range(gTabName).Value
    strOutputPathName = Range(gOutputPathName)
    gActiveWorkbookPath = ActiveWorkbook.Path
    strFullPath = gActiveWorkbookPath & "/" & strPathName & "/" & strFileName
    strFullPathOut = gActiveWorkbookPath & "/" & strOutputPathName & "/" & strFileName

    ' Open the Workbook and Activate the Selected Sheet
    Set wb = Workbooks.Open(fileName:=strFullPath)
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=strFullPathOut
    Application.DisplayAlerts = True
    
    ' Set the Active sheet to the one selected in the drop down list
    Set wsActiveSheet = wb.Worksheets(strTabName)
    wsActiveSheet.Activate

    ' Log the Workbook information, this line comes from the following website with modifications by sean galloway
    ' https://www.exceltip.com/files-workbook-and-worksheets-in-vba/log-files-using-vba-in-microsoft-excel.html
    Debug.Print ThisWorkbook.Name & " opened by " & Application.UserName & " " & Format(Now, "yyyy-mm-dd hh:mm")

    ' Parse the Current Active Sheet
    Debug.Print "Parsing Sheet"
    Call Parse_The_Active_Sheet(wsActiveSheet, dctSummary)
    Debug.Print "Done Parsing Sheet"
    
    ' Summarize the Results
    Call Summarize_Result_To_The_Active_Sheet(wsActiveSheet, dctSummary)
    
    ' Save the Workbook to the designated results area
    wb.Save
End Sub

' This is the Master Routine for opening all sheets, parsing, and returning results
Public Sub Open_Parse_Workbook_MultiSheet()

    Dim strPathName As String
    Dim strFileName As String
    Dim strTabName As String
    Dim strOutputPathName As String
    Dim strFullPath As String
    Dim strFullPathOut As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsActiveSheet As Worksheet
    Dim dctSummary As Scripting.Dictionary
    
    ' Gather the variables
    Sheets(gMulti).Select
    strPathName = Range(gPathName).Value
    strFileName = Range(gFileName).Value
    strOutputPathName = Range(gOutputPathName)
    gActiveWorkbookPath = ActiveWorkbook.Path
    strFullPath = gActiveWorkbookPath & "/" & strPathName & "/" & strFileName
    strFullPathOut = gActiveWorkbookPath & "/" & strOutputPathName & "/" & strFileName

    ' Open the Workbook and Activate the Selected Sheet
    Set wb = Workbooks.Open(fileName:=strFullPath)
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=strFullPathOut
    Application.DisplayAlerts = True
     
    For Each ws In wb.Sheets
        strTabName = ws.Name
        ' Set the Active sheet to the one from the for loop
        Set wsActiveSheet = wb.Worksheets(strTabName)
        wsActiveSheet.Activate
    
        ' Log the Workbook information, this line comes from the following website with modifications by sean galloway
        ' https://www.exceltip.com/files-workbook-and-worksheets-in-vba/log-files-using-vba-in-microsoft-excel.html
        Debug.Print ThisWorkbook.Name & " opened by " & Application.UserName & " " & Format(Now, "yyyy-mm-dd hh:mm")
    
        ' Parse the Current Active Sheet
        Debug.Print "Parsing Sheet " & strTabName
        Call Parse_The_Active_Sheet(wsActiveSheet, dctSummary)
        Debug.Print "Done Parsing Sheet " & strTabName
        
        ' Summarize the Results
        Call Summarize_Result_To_The_Active_Sheet(wsActiveSheet, dctSummary)
    Next ws
    
    ' Save the Workbook to the designated results area
    wb.Save
End Sub

' This is the main routine that parse the Active Sheet and stores the results in a dictionary
Private Sub Parse_The_Active_Sheet(ByVal wsActiveSheet As Worksheet, ByRef dctSummary As Scripting.Dictionary)
    Dim lngI As Long
    Dim rngLastCell As Range
    Dim workIter As New VBAParseItem
    Dim key As String
    Dim workSummary As TickerSummaryItem
    
    ' Create a clean dictionary
    Set dctSummary = New Scripting.Dictionary
    
    ' Find the last cell
    Set rngLastCell = wsActiveSheet.Range("A1").SpecialCells(xlLastCell)
    
    ' Iterate over the sheet; create a dictionary of tickers to track the year long data
    With rngLastCell
        For lngI = gVBAParseStartRow To .Row
            
            ' Grab the info from line lngI
            Call ReadWorkLine(wsActiveSheet, lngI, workIter)
            key = workIter.strTicker
            
            ' Update the ticker info in the dictionary
            If dctSummary.Exists(key) Then
                Set workSummary = dctSummary.Item(key)
                workSummary.dblYearEnd = workIter.dblClose
                workSummary.lngVol = workSummary.lngVol + workIter.lngVol
            Else
                Set workSummary = New TickerSummaryItem
                workSummary.strTicker = key
                workSummary.dblYearStart = workIter.dblOpen
                workSummary.dblYearEnd = 0
                workSummary.lngVol = workIter.lngVol
                dctSummary.Add key:=key, Item:=workSummary
            End If
        Next
    End With
    
    ' Log all of the keys and the Summary Objects to the immediate window for debug
    Dim varKey As Variant
    For Each varKey In dctSummary.Keys
        Set workSummary = dctSummary(varKey)
        Debug.Print varKey & ", " & workSummary.strTicker & ", " & Str(workSummary.dblYearStart) & ", " & Str(workSummary.dblYearEnd) & ", " & Str(workSummary.lngVol)
    Next varKey
    
End Sub

' This parses the current work line and saves it into the VBAParseItem object
Private Sub ReadWorkLine(ByVal wsActiveSheet As Worksheet, ByVal lngI As Long, ByRef myLine As VBAParseItem)
    myLine.lngLine = lngI
    myLine.strTicker = wsActiveSheet.Cells(lngI, gTickerCol).Value
    myLine.strDate = wsActiveSheet.Cells(lngI, gDateCol).Value
    myLine.dblOpen = wsActiveSheet.Cells(lngI, gOpenCol).Value
    myLine.dblHigh = wsActiveSheet.Cells(lngI, gHighCol).Value
    myLine.dblLow = wsActiveSheet.Cells(lngI, gLowCol).Value
    myLine.dblClose = wsActiveSheet.Cells(lngI, gCloseCol).Value
    myLine.lngVol = wsActiveSheet.Cells(lngI, gVolCol).Value
End Sub

' Do the calculations on each record and create a table of the formatted results
Private Sub Summarize_Result_To_The_Active_Sheet(ByVal wsActiveSheet As Worksheet, ByRef dctSummary As Scripting.Dictionary)
    Dim varAnalysisHeaders As Variant
    Dim lngArrayLength As Long
    Dim lngI As Long
    Dim workSummary As TickerSummaryItem
    
    ' Analysis Headers
    varAnalysisHeaders = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    lngArrayLength = UBound(varAnalysisHeaders) - LBound(varAnalysisHeaders) + 1
    
    ' Draw the Analysis Headers
    For lngI = 0 To lngArrayLength - 1
        wsActiveSheet.Cells(1, gSumTickerCol + lngI).Value = varAnalysisHeaders(lngI)
    Next lngI
    
    ' Bonus Headers
    wsActiveSheet.Cells(1, gBonusCol).Value = "Ticker"
    wsActiveSheet.Cells(1, gBonusCol + 1).Value = "Value"
    wsActiveSheet.Cells(2, gBonusCol - 1).Value = "Greatest % Increase"
    wsActiveSheet.Cells(3, gBonusCol - 1).Value = "Greatest % Decrease"
    wsActiveSheet.Cells(4, gBonusCol - 1).Value = "Greatest Total Volume"
    
    With wsActiveSheet.Columns("O")
        .ColumnWidth = 20
    End With

    ' Fill in the Summary Table and do the Greatest* Comparisons
    Dim varKey As Variant
    Dim dblYearlyChange As Double
    Dim dblPercentChange As Double
    Dim clsGreatestPerInc As New GreatestItem
    Dim clsGreatestPerDec As New GreatestItem
    Dim clsGreatestTotVol As New GreatestItem

    lngI = 1
    clsGreatestPerInc.dblValue = 0
    clsGreatestPerDec.dblValue = 0
    clsGreatestTotVol.dblValue = 0

    For Each varKey In dctSummary.Keys
        ' Handle some local variables
        lngI = lngI + 1
        Set workSummary = dctSummary(varKey)
        
        ' Do Calculations
        dblYearlyChange = workSummary.dblYearEnd - workSummary.dblYearStart
        If workSummary.dblYearStart > 0 Then
            dblPercentChange = dblYearlyChange / workSummary.dblYearStart
        Else
            dblPercentChange = 0
        End If
        
        ' Do the Greatest* Comparisons
        If dblPercentChange > clsGreatestPerInc.dblValue Then
            clsGreatestPerInc.dblValue = dblPercentChange
            clsGreatestPerInc.strTicker = workSummary.strTicker
        End If
        
        If dblPercentChange < clsGreatestPerDec.dblValue Then
            clsGreatestPerDec.dblValue = dblPercentChange
            clsGreatestPerDec.strTicker = workSummary.strTicker
        End If
        
        If workSummary.lngVol > clsGreatestTotVol.dblValue Then
            clsGreatestTotVol.dblValue = workSummary.lngVol
            clsGreatestTotVol.strTicker = workSummary.strTicker
        End If
        
        ' Fill In the Cells
        wsActiveSheet.Cells(lngI, gSumTickerCol).Value = workSummary.strTicker
        wsActiveSheet.Cells(lngI, gSumYearlyChangeCol).Value = dblYearlyChange
        wsActiveSheet.Cells(lngI, gSumPercentChangeCol).Value = dblPercentChange
        wsActiveSheet.Cells(lngI, gSumTotalStockVolumeCol).Value = workSummary.lngVol
        
        ' Format the Cells
        If dblYearlyChange < 0 Then
            wsActiveSheet.Cells(lngI, gSumYearlyChangeCol).Interior.ColorIndex = 3
        Else
            wsActiveSheet.Cells(lngI, gSumYearlyChangeCol).Interior.ColorIndex = 4
        End If
        wsActiveSheet.Cells(lngI, gSumYearlyChangeCol).NumberFormat = "#,##0.00"
        wsActiveSheet.Cells(lngI, gSumPercentChangeCol).NumberFormat = "0.00%"
    Next varKey
    
    ' Fill in the Greatest* Cells
    wsActiveSheet.Cells(2, gBonusCol).Value = clsGreatestPerInc.strTicker
    wsActiveSheet.Cells(2, gBonusCol + 1).Value = clsGreatestPerInc.dblValue
    wsActiveSheet.Cells(3, gBonusCol).Value = clsGreatestPerDec.strTicker
    wsActiveSheet.Cells(3, gBonusCol + 1).Value = clsGreatestPerDec.dblValue
    wsActiveSheet.Cells(4, gBonusCol).Value = clsGreatestTotVol.strTicker
    wsActiveSheet.Cells(4, gBonusCol + 1).Value = clsGreatestTotVol.dblValue
    
    ' Format the Cells
    wsActiveSheet.Cells(2, gBonusCol + 1).NumberFormat = "0.00%"
    wsActiveSheet.Cells(3, gBonusCol + 1).NumberFormat = "0.00%"
    
End Sub

