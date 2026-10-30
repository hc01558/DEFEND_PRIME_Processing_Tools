Attribute VB_Name = "dMain_v9_8"
' To run this DEFEND PRIME value selection software, in the Microsoft Visual Basic for Applications window, please follow these steps:
' 1. Click the 'Tools' tab above, in the drop-down list click 'Macros...', then with 'AA_RunThisForDEFENDPRIME' selected, click 'Run'.
' 2. If any error messages are reported, please contact Hugh Claridge (h.claridge@surrey.ac.uk) from the DEFEND PRIME team.
' Programme Identifier: DEFEND PRIME data anonymisation Excel macro
' Name of programmer: Hugh Claridge
' Brief description of programme: This Excel VBA macro processes the DEFEND PRIME data extracted from EMIS Web by the provided XML file to further enhance the anonymisation by grouping ages, retaining only the prioritised results and removing the now extraneous event dates and results.
' Revision history for dMain
' Differs from v9_7 through ClearWorksheet now acting on named Sheet1 and not just the 1st positioned worksheet in the workbook and v9.7 in that ReadyApplication now sets to on/off due to helpermethods change
' Version number: v9_8
' Release Date: 24th October 2025

Option Explicit

'' Main DEFEND PRIME processing macro which takes EMIS format data present in Sheet1 and performs validation, prioritisation and anonymisation
Public Sub AA_RunThisForDEFENDPRIME()
    '' If something unexpected goes wrong like a VBA runtime error, jump to the error handler
    On Error GoTo ErrHandler
    
    '' Comment in Excel status bar so user knows what's going on
    Application.StatusBar = "Checking worksheet formatting and presence of EMIS-format data in preparation for transformation"
 
    '' Check that Sheet1 exists in the workbook being used, and if it does, ensure that it's the active worksheet so that people have it open
    If Not Evaluate("ISREF(Sheet1!A1)") Then
        MsgBox "Sorry, the DEFEND PRIME macro requires the data to be present in the worksheet entitled 'Sheet1' but it cannot be found. Please ensure the correct sheet is present and try again. If the problem persists, please contact the DEFEND PRIME team noting 'Error 1'.", vbExclamation
        Application.StatusBar = False '' Clear the status bar
        Exit Sub '' Stop before changing any data or Excel settings
    ElseIf ThisWorkbook.ActiveSheet.Name <> "Sheet1" Then
        MsgBox "Sorry, Sheet1 is not currently the active worksheet, so please ensure that that worksheet is open and it contains the extracted data before this processing macro is run. If the problem persists, please contact the DEFEND PRIME team noting 'Error 2'.", vbExclamation
        Application.StatusBar = False '' Clear the status bar
        Exit Sub '' Stop before changing any data or Excel settings
    End If
    
    '' Check if the DEFEND PRIME processing has already been run by looking for text indicating processing in cells C1 and F3 (as double pasting of headers can mean F2 is blank)
    '' Double processing removes the required values from the already-processed worksheet, thus this prevents accidental re-running
    If ThisWorkbook.Sheets("Sheet1").Range("C1").Value = "Baseline Height: Value" And _
       ThisWorkbook.Sheets("Sheet1").Range("F3").Value = "Present" Then
        MsgBox "Sorry, processing appears to have already been performed thus re-processing by the DEFEND PRIME macro has been stopped. If you believe this to be incorrect please contact the DEFEND PRIME team noting 'Error 3'.", vbExclamation
        Application.StatusBar = False '' Clear the status bar
        Exit Sub
    End If
    
    '' Check if someone is accidentally trying to run this on a blank worksheet, having perhaps forgotten to paste the data in
    If ThisWorkbook.Sheets("Sheet1").Range("A1").Value = "" Or _
        ThisWorkbook.Sheets("Sheet1").Range("A1").Value = "Please copy and paste your data into this cell" Then
        MsgBox "Sorry, the data do not appear to have been pasted into cell A1 of Sheet1. Please delete any data you have pasted in and repaste the original extract data into cell A1 of Sheet1. If this problem persists, please contact the DEFEND PRIME team noting 'Error 4'.", vbExclamation
        Application.StatusBar = False '' Clear the status bar
        Exit Sub
    End If
    
    '' Check if someone is accidentally trying to run this on TPP data prior to transformation
    If ThisWorkbook.Sheets("Sheet1").Range("A1").Value = "Organisation ID" Then
        MsgBox "Sorry, the TPP data do not appear to have been processed yet by Button 1: TRANSFORM TPP DATA. Please ensure you have already processed the TPP data using Button 1 before clicking Button 2. If this problem persists, please contact the DEFEND PRIME team noting 'Error 6'.", vbExclamation
        Application.StatusBar = False '' Clear the status bar
        Exit Sub
    End If
       
    '' Set Up Application/Globals
    Application.StatusBar = "Checking data" '' Let user know in Excel status bar that processing has started
    Call SetWorksheet '' From dGlobals sets dWS to Sheet1 for subsequent processing
    Call ReadyApplication(processingMode:=True) '' From dHelperMethods, turns off certain Excel settings for faster processing
    
'Err.Raise 12345, , "Deliberate test of EMIS error handler" '' Uncomment the start of this line to test VBA error handler
    
    Call SetHeaderValues '' From dGlobals, loads the standard DEFEND PRIME header names into gHeaderTexts
    
    '' Check and prepare pasted data before processing
    Call CheckData '' Removes extraneous EMIS export structure, standardises headers and establishes first and final data rows
    '' Stop cleanly if CheckData could not confirm the expected EMIS structure
    If gFirstRowOfData = 0 Then
        Application.StatusBar = False '' Clear processing message
        Call ReadyApplication(processingMode:=False) '' Turn on normal Excel settings after processing was stopped
        Call ClearGlobalData '' Clear globals ready for any subsequent run
        Exit Sub
    End If
    '' Continue preparing the pasted data
    Call ClearFormatting(dWS) '' From dHelperMethods, removes existing worksheet formatting before applying consistent output formatting later
    Call FixHyphenError(dWS) '' From dHelperMethods, reprocesses cells containing hyphens to address known date-recognition issues
    Call FormatDataSheet '' Apply consistent date formatting to the defined date cols
    Call LoadDataIntoArray '' Load the data into an array to process in memory rather than in the worksheet given the processing cost of doing this
       
    Application.StatusBar = "Removing identifiable features" '' Tell user the main anonymisation stage is underway
    Dim iRow As Long '' Row counter used to loop through each patient record in gDataSet
    For iRow = LBound(gDataSet, 1) To UBound(gDataSet, 1) '' Loop from first to last patient row held in the array
             
        Call G_CancerTestingDateReplacement(iRow) '' Replace pancreatic cancer diagnosis, CA19-9 and CEA dates with Before/On/After index labels
        Call C1_HbA1cDateValidation(iRow) '' Validate HbA1c 2a-2e dates against their required windows
        '' Xx prefix denotes weight, BMI and height dates and values processing but X9 is the exception as it acts on all including HbA1c
        Call X1_WeightBMIHeight2cDateValidation(iRow) '' Validate Weight/BMI/Height 2c dates
        Call X2_WeightBMIHeight2dDateValidation(iRow) '' Validate and classify Weight/BMI/Height 2d dates
        Call X3_WeightBMIHeight2eDateValidation(iRow) '' Same for 2e dates
        Call X4_WeightBMIHeight1a1b2a2bDateValidation(iRow) '' Validate 1a/1b/2a/2b dates and store day differences for later comparison
        Call X5_WeightBMIHeight1a1bDateDifferenceComparison(iRow) '' Compare 1a with 1b to find result closest to index
        Call X6_WeightBMIHeight2a2bDateDifferenceComparison(iRow) '' Compare 2a with 2b to find result closest to one year pre-index
        Call X7_WeightBMIHeightExcess1a1bValueRemoval(iRow) '' Blank the non-preferred 1a/1b Weight/BMI/Height measurement values
            'X8_ removed
        Call X9_HbA1cWeightBMIHeight2ExcessValueRemoval(iRow) '' Blank non-preferred 2a-2e HbA1c/Weight/BMI/Height values
        
        '' Mark rows calculated as under 50 at index for later deletion given mid-processing deletion slows processing
        If Len(gDataSet(iRow, DataCols.IndexDatePatientsAgeatEvent)) > 0 Then '' If a value is present (should always be)
            If IsAgeBelow50(gDataSet(iRow, DataCols.IndexDatePatientsAgeatEvent)) Then  '' Check age or existing age band against under-50 criteria
                gDataSet(iRow, DataCols.lastColumn) = True '' Mark the temporary Delete column for this row if yes
            End If
        End If
        
        '' Apply final anonymisation steps only to rows not marked for deletion
        If gDataSet(iRow, DataCols.lastColumn) <> True Then '' Skip this bit if the row's been marked for deletion
            '' Do the age range grouping into 5-year bands
            Call B04_ReplaceAgeWithAgeRange(iRow)
            '' Replace index date with 'Present' now that that date is no longer needed for calculations
            Call B05_ReplaceWithPresent(iRow, DataCols.IndexDateDate)
            '' Replace anomyised ID cells with 'Present'
            Call B06_RemoveAnonymisedID(iRow, DataCols.AnonymisedIdentifier)
        End If
    Next iRow '' Move to next data row until all have been processed
    
    '' Write the processed data back to Sheet1, format the output and restore normal Excel settings
    Call LoadDataIntoSheet '' Write gDataSet back to worksheet, delete marked rows and remove temporary Delete column
    Call ApplyBorders '' Apply consistent borders, widths and header formatting to the processed data
    Call ReadyApplication(processingMode:=False) '' Turn back on the Excel settings turned off at the start of processing
    Application.StatusBar = False '' Clear processing message from Excel status bar
    Call ClearGlobalData '' Clear module-level data ready for any subsequent run (shouldn't ever be needed in normal use but safer for repeat testing runs)
    
    '' Check contents in cells C1 and F3 indicating successful processing (i.e. 'Baseline Height: Value' and 'Present' are in defined cells), and display success message
    'ThisWorkbook.Sheets("Sheet1").Range("C1").Value = "TEST" '' Uncomment start of this line to test Error 5
    If ThisWorkbook.Sheets("Sheet1").Range("C1").Value = "Baseline Height: Value" And _
       ThisWorkbook.Sheets("Sheet1").Range("F3").Value = "Present" Then
        MsgBox "DEFEND PRIME processing complete, please now save and send the file to the DEFEND PRIME team.", vbInformation  '' Display success message to the user
        '' If 'Baseline Height: Value' or 'Present' are not in the defined cells, then there's a problem
    ElseIf ThisWorkbook.Sheets("Sheet1").Range("C1").Value <> "Baseline Height: Value" Or _
           ThisWorkbook.Sheets("Sheet1").Range("F3").Value <> "Present" Then
        '' Display Error message to the user saying processing looks to have gone wrong
        MsgBox "It looks like something may have gone wrong, so please check that the DEFEND PRIME processing has been run on the desired data by conducting a quick visual check to see if data are present as expected. If not, please repaste the original data, including all headers, in cell A1 of a worksheet with worksheet tab labelled 'Sheet1' and rerun the DEFEND PRIME macro. If the problem persists, please contact the DEFEND PRIME team noting 'Error 5'.", vbExclamation
    End If
    Exit Sub '' If all completes as normal, do not enter error handler
    
ErrHandler: '' This runs only if an unexpected VBA runtime error occurs and enables a clean exit
    Dim errNumber As Long '' Stores original VBA error number before clean-up code can change Err
    Dim errDescription As String '' Stores original VBA error description

    '' Save error details immediately before clean-up code can alter Err
    errNumber = Err.Number
    errDescription = Err.Description

    '' Restore the Excel settings changed by ReadyApplication
    On Error Resume Next '' Prevents a secondary error during clean-up from replacing the original error being reported
    Call ReadyApplication(processingMode:=False) '' Ensure normal Excel settings are turned on
    Application.StatusBar = False '' Clear any processing message left in status bar

    Call ClearGlobalData '' Clear arrays/global values which may contain partially processed data
    On Error GoTo 0 '' Return from permissive clean up error handling to standard VBA behaviour

    '' Report the original error details saved before clean-up
    MsgBox "DEFEND PRIME processing stopped with error " & errNumber & ":" & vbCrLf & _
           errDescription, vbCritical, "DEFEND PRIME processing"

    Exit Sub '' End macro after displaying the unexpected runtime error

End Sub

'' Developer utility only, it's not part of the actual DEFEND PRIME macro
'' To ease creation of the blank sheet for testing and provision to primary care practices, this clears the Sheet1 and puts back in the placeholder text in A1
Private Sub ClearWorksheet()
    Dim ws As Worksheet '' Worksheet variable used for the sheet being reset
    Set ws = ThisWorkbook.Worksheets("Sheet1") '' Set ws to the worksheet with the tab name Sheet1 so another worksheet cannot be cleared because of tab position
    ws.Cells.Clear '' Remove all contents and formatting from the worksheet
    ws.Columns("A:A").ColumnWidth = 41 '' Set Column A wide enough for the paste-in instruction
    With ws.Range("A1") '' Apply placeholder text and formatting to the first cell (i.e. top left)
        .value2 = "Please copy and paste your data into this cell"
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        .Interior.Color = vbYellow
        .Font.Bold = True
    End With
End Sub

'' This sub is only used by the G_CancerTestingDateReplacement sub below for pancreatic cancer diagnosis, CA 19-9 and CEA
'' The previous G_ processing was removing rows, even when run isolated, thus have removed the remFromSamp lines as unclear what their purpose was as the cases being removed did not meet the apparent criteria for removal
Private Sub CheckAgainstIndex(iRow As Long, iColumn As Long)
    Dim dateDifference As Long '' Stores number of days from Index Date to the result date
    dateDifference = DateDiff("d", gDataSet(iRow, DataCols.IndexDateDate), gDataSet(iRow, iColumn)) '' Positive if event after index, negative if before
    If dateDifference > 0 Then '' Result was after Index Date
        gDataSet(iRow, iColumn) = "After index" '' Replace date with label
    ElseIf dateDifference < 0 Then '' Result was before Index Date
        gDataSet(iRow, iColumn) = "Before index" '' Replace date with label
    Else '' If difference is zero means result date was same as Index Date
        gDataSet(iRow, iColumn) = "On index" '' Replace date with label
    End If
End Sub

'' Uses CheckAgainstIndex to replace pancreatic cancer diagnosis, CA19-9 and CEA dates with whether they occurred before, on or after Index Date
Private Sub G_CancerTestingDateReplacement(iRow As Long)
    If IsDate(gDataSet(iRow, DataCols.IndexDateDate)) Then '' If there's an Index Date
        If IsDate(gDataSet(iRow, DataCols.PancreaticcancerdiagnosisDate)) Then '' Check pancreatic cancer diagnosis date is a valid date
            Call CheckAgainstIndex(iRow, DataCols.PancreaticcancerdiagnosisDate) '' Replace with Before/On/After index
        End If
        If IsDate(gDataSet(iRow, DataCols.Cancerantigen199CA199Date)) Then '' Check CA19-9 result date is a valid date
            Call CheckAgainstIndex(iRow, DataCols.Cancerantigen199CA199Date) '' Replace with Before/On/After index
            
        End If
        If IsDate(gDataSet(iRow, DataCols.CarcinoembryonicantigenCEADate)) Then '' Check CEA result date is a valid date
            Call CheckAgainstIndex(iRow, DataCols.CarcinoembryonicantigenCEADate) '' Replace with Before/On/After index
        End If
    End If
End Sub

'Validates all HbA1c result dates other than the index HbA1c result (so validates 2a, 2b, 2c, 2d and 2e)
Private Sub ValidateHbA1cDates(iRow As Long, iCol As Long, iBottom As Long, iTop As Long)
    Dim dateDifference As Long '' Stores number of days between the result date and Index Date
    If IsDate(gDataSet(iRow, iCol)) Then '' If the result date cell contains a date
        dateDifference = DateDiff("d", gDataSet(iRow, iCol), gDataSet(iRow, DataCols.IndexDateDate)) '' Calculate days before Index Date
        
        If dateDifference >= iBottom And dateDifference <= iTop Then '' Date lies inside the expected window for that range
            gDataSet(iRow, iCol) = "Valid" '' Replace date with validity label
        Else
            Dim invalidLabel As String '' Stores label saying how far the date lies outside the expected window for that range
            If dateDifference < iBottom Then '' Result is later than the permitted window, i.e. closer to the Index Date than permitted
                invalidLabel = "Inval (+" & (iBottom - dateDifference) & " days)" '' Record number of days outside the later boundary of the valid window
            Else '' Result is earlier than the permitted window, i.e. further from the Index Date than permitted
                invalidLabel = "Inval (-" & (dateDifference - iTop) & " days)" '' Record number of days outside the earlier boundary of the valid window
            End If
            gDataSet(iRow, iCol) = invalidLabel '' Replace original date with calculated invalidity label
        End If
    End If
End Sub

'' Runs HbA1c date validation for all five pre-index HbA1c result windows
Private Sub C1_HbA1cDateValidation(iRow As Long)
    If IsDate(gDataSet(iRow, DataCols.IndexDateDate)) Then '' Only validate results dates where Index Date is available
        Call ValidateHbA1cDates(iRow, DataCols.hbA1c2aDate, 274, 455) '' HbA1c 2a must be 274-455 days before index
        Call ValidateHbA1cDates(iRow, DataCols.hbA1c2bDate, 456, 546) '' HbA1c 2b must be 456-546 days before index
        Call ValidateHbA1cDates(iRow, DataCols.hbA1c2cDate, 183, 273) '' HbA1c 2c must be 183-273 days before index
        Call ValidateHbA1cDates(iRow, DataCols.hbA1c2dDate, 92, 182) '' HbA1c 2d must be 92-182 days before index
        Call ValidateHbA1cDates(iRow, DataCols.hbA1c2eDate, 547, 728) '' HbA1c 2e must be 547-728 days before index
    End If
End Sub

'Validates the weight, BMI and height dates for 2c against expected 456-546 day pre-index window
Private Sub ValidateWeightBMIHeight2cDates(iRow As Long, iCol As Long, iBottom As Long, iTop As Long)
    Dim dateDifference As Long '' Stores number of days from result date to Index Date
    If IsDate(gDataSet(iRow, iCol)) Then '' Only process if result date field contains a date
        dateDifference = DateDiff("d", gDataSet(iRow, iCol), gDataSet(iRow, DataCols.IndexDateDate)) '' Calculate days before Index Date
        
        If dateDifference >= iBottom And dateDifference <= iTop Then '' Date lies inside the 2c window
            gDataSet(iRow, iCol) = "Valid" '' Replace original date with validity label
        Else
            Dim invalidLabel As String '' Stores how far result lies outside permitted window
            If dateDifference < iBottom Then '' Result is later than the permitted window, i.e. closer to the Index Date than permitted
                invalidLabel = "Inval (+" & (iBottom - dateDifference) & " days)" '' Record number of days outside the later boundary of the valid window
            Else '' Result is earlier than the permitted window, i.e. further from the Index Date than permitted
                invalidLabel = "Inval (-" & (dateDifference - iTop) & " days)" '' Record number of days outside the earlier boundary of the valid window
            End If
            gDataSet(iRow, iCol) = invalidLabel '' Replace original date with calculated invalidity label
        End If
    End If
End Sub

'' Runs 2c date validation for Weight, BMI and Height
Private Sub X1_WeightBMIHeight2cDateValidation(iRow As Long)
    If IsDate(gDataSet(iRow, DataCols.IndexDateDate)) Then '' Only validate results dates where Index Date is available
        Call ValidateWeightBMIHeight2cDates(iRow, DataCols.weight2cDate, 456, 546) '' Weight 2c: 456-546 days before Index Date
        Call ValidateWeightBMIHeight2cDates(iRow, DataCols.bMI2cDate, 456, 546) '' BMI 2c: 456-546 days before Index Date
        Call ValidateWeightBMIHeight2cDates(iRow, DataCols.height2cDate, 456, 546) '' Height 2c: 456-546 days before Index Date
    End If
End Sub

'Validates the weight, BMI and height dates for 2d against expected 92-273 day pre-index window
'' Also records whether a valid result falls 3-6 or 6-9 months before Index Date for more detailed reporting
Private Sub ValidateWeightBMIHeight2dDates(iRow As Long, iCol As Long, iBottom As Long, iTop As Long)
    Dim dateDifference As Long '' Stores number of days from result date to Index Date
    If IsDate(gDataSet(iRow, iCol)) Then '' Only process if result date field contains a date
        dateDifference = DateDiff("d", gDataSet(iRow, iCol), gDataSet(iRow, DataCols.IndexDateDate)) '' Calculate number of days before Index Date
        
        If dateDifference >= iBottom And dateDifference <= iTop Then '' Result date lies inside the overall 2d window
            '' Additional checks to ascertain the specific range within the valid range
            If dateDifference >= 92 And dateDifference <= 182 Then '' Result lies 3-6 months before Index Date
                gDataSet(iRow, iCol) = "Valid 3-6" '' Replace original date with specific validity label
            ElseIf dateDifference >= 183 And dateDifference <= 273 Then '' Result lies 6-9 months before Index Date
                gDataSet(iRow, iCol) = "Valid 6-9" '' Replace original date with specific validity label
            End If
        Else
            Dim invalidLabel As String '' Stores how far the result date lies outside the defined window
            If dateDifference < iBottom Then '' Result is later than the permitted window, i.e. closer to the Index Date than permitted
                invalidLabel = "Inval (+" & (iBottom - dateDifference) & " days)" '' Record number of days outside the later boundary of the valid window
            Else '' Result is earlier than the permitted window, i.e. further from the Index Date than permitted
                invalidLabel = "Inval (-" & (dateDifference - iTop) & " days)" '' Record number of days outside the earlier boundary of the valid window
            End If
            gDataSet(iRow, iCol) = invalidLabel '' Replace original date with calculated invalidity label
        End If
    End If
End Sub

'' Runs 2d date validation for Weight, BMI and Height
Private Sub X2_WeightBMIHeight2dDateValidation(iRow As Long)
    If IsDate(gDataSet(iRow, DataCols.IndexDateDate)) Then '' Only validate results dates where Index Date is available
        Call ValidateWeightBMIHeight2dDates(iRow, DataCols.weight2dDate, 92, 273) '' Weight 2d: 92-273 days before Index Date
        Call ValidateWeightBMIHeight2dDates(iRow, DataCols.bMI2dDate, 92, 273) '' BMI 2d: 92-273 days before Index Date
        Call ValidateWeightBMIHeight2dDates(iRow, DataCols.height2dDate, 92, 273) '' Height 2d: 92-273 days before Index Date
    End If
End Sub

'' Validates the weight, BMI and height dates for 2e against expected 547-3652 day pre-index window
'' Also records whether a valid result falls 18-24 months or 24 months-10 years before Index Date for more detailed reporting
Private Sub ValidateWeightBMIHeight2eDates(iRow As Long, iCol As Long, iBottom As Long, iTop As Long)
    Dim dateDifference As Long '' Stores number of days from result date to Index Date
    If IsDate(gDataSet(iRow, iCol)) Then '' Only process if result date field contains a date
        dateDifference = DateDiff("d", gDataSet(iRow, iCol), gDataSet(iRow, DataCols.IndexDateDate)) '' Calculate number of days before Index Date
        
        If dateDifference >= iBottom And dateDifference <= iTop Then '' Result date lies inside the overall 2e window
             '' Additional checks to ascertain the specific range within the valid range
            If dateDifference >= 547 And dateDifference <= 728 Then '' Result lies 18-24 months before Index Date
                gDataSet(iRow, iCol) = "Valid 18-24" '' Replace original date with specific validity label
            ElseIf dateDifference >= 729 And dateDifference <= 3652 Then '' Result lies 24 months-10 years before Index Date
                gDataSet(iRow, iCol) = "Valid 24-10yr" '' Replace original date with specific validity label
            End If
        Else
            Dim invalidLabel As String '' Stores how far the result date lies outside the defined window
            If dateDifference < iBottom Then '' Result is later than the permitted window, i.e. closer to the Index Date than permitted
                invalidLabel = "Inval (+" & (iBottom - dateDifference) & " days)" '' Record number of days outside the later boundary of the valid window
            Else '' Result is earlier than the permitted window, i.e. further from the Index Date than permitted
                invalidLabel = "Inval (-" & (dateDifference - iTop) & " days)" '' Record number of days outside the earlier boundary of the valid window
            End If
            gDataSet(iRow, iCol) = invalidLabel '' Replace original date with calculated invalidity label
        End If
    End If
End Sub

'' Runs 2e date validation for Weight, BMI and Height
Private Sub X3_WeightBMIHeight2eDateValidation(iRow As Long)
    If IsDate(gDataSet(iRow, DataCols.IndexDateDate)) Then '' Only validate results dates where Index Date is available
        Call ValidateWeightBMIHeight2eDates(iRow, DataCols.weight2eDate, 547, 3652) '' Weight 2e: 547-3652 days before Index Date
        Call ValidateWeightBMIHeight2eDates(iRow, DataCols.bMI2eDate, 547, 3652) '' BMI 2e: 547-3652 days before Index Date
        Call ValidateWeightBMIHeight2eDates(iRow, DataCols.height2eDate, 547, 3652) '' Height 2e: 547-3652 days before Index Date
    End If
End Sub

'' Validates the weight, BMI and height dates for 1a, 1b, 2a and 2b against expected windows
'' Stores the number of days from the result to Index Date for valid results so paired results can be compared and prioritised
Private Sub ValidateWeightBMIHeight1a1b2a2bDates(iRow As Long, iCol As Long, iBottom As Long, iTop As Long)
    Dim dateDifference As Long '' Stores number of days from result date to Index Date
    If IsDate(gDataSet(iRow, iCol)) Then '' Only process if result date field contains a date
        dateDifference = DateDiff("d", gDataSet(iRow, iCol), gDataSet(iRow, DataCols.IndexDateDate)) '' Calculate day difference - results after Index Date produce negative values
        
        If dateDifference >= iBottom And dateDifference <= iTop Then '' Result date lies inside the defined window
        'Instead of usual inserting 'Valid', if valid, insert the number of days to enable comparison between 1a and 1b, and 2a and 2b if present
            gDataSet(iRow, iCol) = dateDifference
        Else
            Dim invalidLabel As String '' Stores how far the result date lies outside the defined window
            If dateDifference < iBottom Then '' Result is later than permitted relative to the Index Date: closer to Index Date for pre-index windows, or further after Index Date for 1a
                invalidLabel = "Inval (+" & (iBottom - dateDifference) & " days)" '' Record number of days beyond the later boundary of the valid window
            Else '' Result is earlier than permitted relative to the Index Date: further before Index Date for pre-index windows, or before Index Date for 1a
                invalidLabel = "Inval (-" & (dateDifference - iTop) & " days)" '' Record number of days beyond the earlier boundary of the valid window
            End If
            gDataSet(iRow, iCol) = invalidLabel '' Replace original date with calculated invalidity label
        End If
    End If
End Sub

'' Runs 1a, 1b, 2a and 2b date validation for Weight, BMI and Height
Private Sub X4_WeightBMIHeight1a1b2a2bDateValidation(iRow As Long)
    If IsDate(gDataSet(iRow, DataCols.IndexDateDate)) Then '' Only validate results dates where Index Date is available
        '' 1a and 1b. CAREFUL as 1a is after Index Date, so DateDiff(result date, Index Date) produces negative values for valid 1a dates despite being after index, but is an exceptional circumstance compared to most which are before index
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.weight1aDate, -91, 0) '' Weight 1a: Index Date to 91 days after Index Date
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.weight1bDate, 1, 91) '' Weight 1b: 1-91 days before Index Date
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.bMI1aDate, -91, 0) '' BMI 1a: Index Date to 91 days after Index Date
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.bMI1bDate, 1, 91) '' BMI 1b: 1-91 days before Index Date
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.height1aDate, -91, 0) '' Height 1a: Index Date to 91 days after Index Date
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.height1bDate, 1, 91) '' Height 1b: 1-91 days before Index Date
        '' 2a and 2b
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.weight2aDate, 274, 364) '' Weight 2a: 274-364 days before Index Date
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.weight2bDate, 365, 455) '' Weight 2b: 365-455 days before Index Date
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.bMI2aDate, 274, 364) '' BMI 2a: 274-364 days before Index Date
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.bMI2bDate, 365, 455) '' BMI 2b: 365-455 days before Index Date
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.height2aDate, 274, 364) '' Height 2a: 274-364 days before Index Date
        Call ValidateWeightBMIHeight1a1b2a2bDates(iRow, DataCols.height2bDate, 365, 455) '' Height 2b: 365-455 days before Index Date
        
    End If
End Sub

'' Checks the magnitude of the difference between the 1a and 1b dates for weight, height and BMI and classifies these dates accordingly in relation to Index Date
Private Sub CompareWeightBMIHeight1a1bDateDifferences(iRow As Long, iCol1 As Long, iCol2 As Long)
    Dim value1 As Variant, value2 As Variant '' Variables to hold the processed 1a and 1b date-field contents
    value1 = gDataSet(iRow, iCol1) '' Get 1a day difference or validity label
    value2 = gDataSet(iRow, iCol2) '' Get 1b day difference or validity label
    
    '' Check first whether both possible results are blank or invalid
    If (value1 = "" Or InStr(value1, "Inval") > 0) And (value2 = "" Or InStr(value2, "Inval") > 0) Then '' Both missing/invalid
        Exit Sub '' Neither result is valid/present so nothing to compare
    End If

    '' Check and label when only one value is numeric and valid
    '' Check if only 1a is valid
    If IsNumeric(value1) And (Not IsNumeric(value2) Or value2 = "" Or InStr(value2, "Inval") > 0) Then '' Only 1a present and valid
        gDataSet(iRow, iCol1) = "Valid" '' Mark 1a as valid because there is no valid 1b result to compare against
        Exit Sub '' Exit given no further comparison is required
    '' Check if only 1b is valid
    ElseIf IsNumeric(value2) And (Not IsNumeric(value1) Or value1 = "" Or InStr(value1, "Inval") > 0) Then '' Only 1b present and valid
        gDataSet(iRow, iCol2) = "Valid" '' Mark 1b as valid because there is no valid 1a result to compare against
        Exit Sub '' Exit given no further comparison is required
    End If

    '' If both 1a and 1b results are valid and numeric, compare their distance from Index Date
    If IsNumeric(value1) And IsNumeric(value2) Then
        If Abs(value1) < Abs(value2) Then '' 1a is fewer days from Index Date than 1b
            gDataSet(iRow, iCol1) = "Closer" '' Mark 1a as preferred
            gDataSet(iRow, iCol2) = "Further" '' Mark 1b as not
        ElseIf Abs(value1) > Abs(value2) Then '' 1b is fewer days from Index Date than 1a
            gDataSet(iRow, iCol1) = "Further" '' Mark 1a as not preferred
            gDataSet(iRow, iCol2) = "Closer" '' Mark 1b as preferred
        Else '' Both results are the same number of days from Index Date
            gDataSet(iRow, iCol1) = "Equidistant" '' Mark 1a as equidistant
            gDataSet(iRow, iCol2) = "Equidistant" '' Mark 1b as equidistant
        End If
    End If
End Sub

'' Runs the 1a versus 1b date difference comparison for Weight, BMI and Height
Private Sub X5_WeightBMIHeight1a1bDateDifferenceComparison(iRow As Long)
    Call CompareWeightBMIHeight1a1bDateDifferences(iRow, DataCols.weight1aDate, DataCols.weight1bDate) '' Compare Weight 1a and 1b
    Call CompareWeightBMIHeight1a1bDateDifferences(iRow, DataCols.bMI1aDate, DataCols.bMI1bDate) '' Compare BMI 1a and 1b
    Call CompareWeightBMIHeight1a1bDateDifferences(iRow, DataCols.height1aDate, DataCols.height1bDate) '' Compare Height 1a and 1b
End Sub

'' Checks the magnitude of the difference between the 2a and 2b dates for weight, height and BMI and classifies these dates accordingly
'' Almost identical to the 1a and 1b date processing, but subtracts 364.5 days from the result to ascertain which is closer to index (smallest difference alone would incorrectly always result in 2a)
'' 364.5 is the midpoint between the 2a upper boundary of 364 days and the 2b lower boundary of 365 days
'' An alternative method would have been to originally calculate the date difference from the index date less 364.5 days
'' but then that would have required two processes for 1a1b and 2a2b date difference processing rather than the two date difference comparisons, so there's no meaningful difference
Private Sub CompareWeightBMIHeight2a2bDateDifferences(iRow As Long, iCol1 As Long, iCol2 As Long)
    Dim value1 As Variant, value2 As Variant '' Variables to hold the processed 2a and 2b date contents
    value1 = gDataSet(iRow, iCol1) '' Get 2a day difference or validity label
    value2 = gDataSet(iRow, iCol2) '' Get 2b day difference or validity label
    
    '' Check first whether both possible results are blank or invalid
    If (value1 = "" Or InStr(value1, "Inval") > 0) And (value2 = "" Or InStr(value2, "Inval") > 0) Then '' Both missing/invalid
        Exit Sub '' Neither result is valid/present so nothing to compare
    End If

    '' Check and label when only one value is numeric and valid
    '' Check if only 2a is valid
    If IsNumeric(value1) And (Not IsNumeric(value2) Or value2 = "" Or InStr(value2, "Inval") > 0) Then '' Only 2a present and valid
        gDataSet(iRow, iCol1) = "Valid" '' Mark 2a as valid because there is no valid 2b result to compare against
        Exit Sub ' Exit after processing to prevent further comparison
    ElseIf IsNumeric(value2) And (Not IsNumeric(value1) Or value1 = "" Or InStr(value1, "Inval") > 0) Then '' Only 2b present and valid
        gDataSet(iRow, iCol2) = "Valid" '' Mark 2b as valid because there is no valid 2a result to compare against
        Exit Sub ' Exit after processing to prevent further comparison
    End If

    '' If both 2a and 2b results are valid and numeric, compare their distance from the 364.5-day target
    If IsNumeric(value1) And IsNumeric(value2) Then
        If Abs(value1 - 364.5) < Abs(value2 - 364.5) Then '' 2a lies closer to the 364.5-day target than 2b
            gDataSet(iRow, iCol1) = "Closer" '' Mark 2a as preferred
            gDataSet(iRow, iCol2) = "Further" '' Mark 2b as not
        ElseIf Abs(value1 - 364.5) > Abs(value2 - 364.5) Then '' 2b lies closer to the 364.5-day target than 2a
            gDataSet(iRow, iCol1) = "Further" '' Mark 2a as not preferred
            gDataSet(iRow, iCol2) = "Closer" '' Mark 2b as preferred
        Else '' Both results are same distance from the 364.5-day target
            gDataSet(iRow, iCol1) = "Equidistant" '' Mark 2a as equidistant
            gDataSet(iRow, iCol2) = "Equidistant" '' Mark 2b as equidistant
        End If
    End If
End Sub

'' Runs the 2a versus 2b date-difference comparison for Weight, BMI and Height
Private Sub X6_WeightBMIHeight2a2bDateDifferenceComparison(iRow As Long)
        Call CompareWeightBMIHeight2a2bDateDifferences(iRow, DataCols.weight2aDate, DataCols.weight2bDate) '' Compare Weight 2a and 2b
        Call CompareWeightBMIHeight2a2bDateDifferences(iRow, DataCols.bMI2aDate, DataCols.bMI2bDate) '' Compare BMI 2a and 2b
        Call CompareWeightBMIHeight2a2bDateDifferences(iRow, DataCols.height2aDate, DataCols.height2bDate) '' Compare Height 2a and 2b
End Sub

'' Removes the excess/non-preferred weights, BMI and height values from 1a 1b using the preceding date comparison
'' If 1a/b are equidistant, preference is given to 1a
Private Sub RemoveExcessWeightBMIHeight1a1bValues(iRow As Long, OneA As Long, OneB As Long, OneAVal As Long, OneBVal As Long)
    If gDataSet(iRow, OneA) = "Closer" Or gDataSet(iRow, OneA) = "Equidistant" Or InStr(gDataSet(iRow, OneA), "Valid") > 0 Then '' 1a is preferred, equidistant, or the only valid result
        gDataSet(iRow, OneBVal) = "" '' Remove the 1b value
    ElseIf gDataSet(iRow, OneA) = "" Or gDataSet(iRow, OneA) = "Further" Or InStr(gDataSet(iRow, OneA), "Inval") > 0 Then '' 1a is absent, further away or invalid
        If gDataSet(iRow, OneB) = "Closer" Or InStr(gDataSet(iRow, OneB), "Valid") > 0 Then '' Confirm 1b is preferred or the only valid result
            gDataSet(iRow, OneAVal) = "" '' Remove the 1a value
        End If
    End If
End Sub

'' Runs 1a/1b excess result removal for Weight, BMI and Height
Private Sub X7_WeightBMIHeightExcess1a1bValueRemoval(iRow As Long)
    Call RemoveExcessWeightBMIHeight1a1bValues(iRow, DataCols.weight1aDate, DataCols.weight1bDate, DataCols.Weight1aValue, DataCols.Weight1bValue) '' Retain preferred Weight result
    Call RemoveExcessWeightBMIHeight1a1bValues(iRow, DataCols.bMI1aDate, DataCols.bMI1bDate, DataCols.BMI1aValue, DataCols.BMI1bValue) '' Retain preferred BMI result
    Call RemoveExcessWeightBMIHeight1a1bValues(iRow, DataCols.height1aDate, DataCols.height1bDate, DataCols.Height1aValue, DataCols.Height1bValue) '' Retain preferred height result
End Sub

'' Removes the excess/non-preferred HbA1c, weights, BMI and height values from 2a, 2b, 2c, 2d and 2e using the preceding date validation and prioritisation
'' If 2a/b are equidistant, preference is given to 2a
Private Sub RemoveExcessHbA1cWeightBMIHeight2Values(iRow As Long, TwoA As Long, TwoB As Long, TwoC As Long, TwoD As Long, TwoE As Long, _
                                              TwoAVal As Long, TwoBVal As Long, TwoCVal As Long, TwoDVal As Long, TwoEVal As Long)
    If gDataSet(iRow, TwoA) = "Closer" Or gDataSet(iRow, TwoA) = "Equidistant" Or gDataSet(iRow, TwoA) = "Valid" Then '' 2a is preferred or the only valid 2a/2b result
        '' Clear all later values since TwoA is valid or closest
        gDataSet(iRow, TwoBVal) = ""
        gDataSet(iRow, TwoCVal) = ""
        gDataSet(iRow, TwoDVal) = ""
        gDataSet(iRow, TwoEVal) = ""
    ElseIf gDataSet(iRow, TwoB) = "Closer" Or gDataSet(iRow, TwoB) = "Valid" Then '' 2b is preferred or the only valid 2a/2b result
        '' Clear all values except TwoB as it's the preferred one
        gDataSet(iRow, TwoAVal) = ""
        gDataSet(iRow, TwoCVal) = ""
        gDataSet(iRow, TwoDVal) = ""
        gDataSet(iRow, TwoEVal) = ""
    ElseIf gDataSet(iRow, TwoC) = "Valid" Then '' Neither 2a nor 2b is retained and 2c is valid
        '' Clear all values except TwoC given it's the preferred one
        gDataSet(iRow, TwoAVal) = ""
        gDataSet(iRow, TwoBVal) = ""
        gDataSet(iRow, TwoDVal) = ""
        gDataSet(iRow, TwoEVal) = ""
    ElseIf InStr(gDataSet(iRow, TwoD), "Valid") > 0 Then '' No prior preferred result exists and 2d contains a valid 3-6 or 6-9 month result
        '' Clear all values except TwoD given it's the preferred one
        gDataSet(iRow, TwoAVal) = ""
        gDataSet(iRow, TwoBVal) = ""
        gDataSet(iRow, TwoCVal) = ""
        gDataSet(iRow, TwoEVal) = ""
    ElseIf InStr(gDataSet(iRow, TwoE), "Valid") > 0 Then '' No earlier preferred result exists and 2e contains a valid 18-24 month or 24 month-10 year result
        '' Clear all earlier values as TwoE is the remaining preferred one
        gDataSet(iRow, TwoAVal) = ""
        gDataSet(iRow, TwoBVal) = ""
        gDataSet(iRow, TwoCVal) = ""
        gDataSet(iRow, TwoDVal) = ""
    End If
End Sub

'Private Sub X8_WeightBMIHeight2ExcessValueRemoval(iRow As Long)
'    Call RemoveExcessWeightBMIHeight2Values(iRow, DataCols.weight2aDate, DataCols.weight2bDate, DataCols.weight2cDate, DataCols.weight2dDate, DataCols.weight2eDate, _
'    DataCols.Weight2aValue, DataCols.Weight2bValue, DataCols.Weight2cValue, DataCols.Weight2dValue, DataCols.Weight2eValue)
'    Call RemoveExcessWeightBMIHeight2Values(iRow, DataCols.bMI2aDate, DataCols.bMI2bDate, DataCols.bMI2cDate, DataCols.bMI2dDate, DataCols.bMI2eDate, _
'    DataCols.BMI2aValue, DataCols.BMI2bValue, DataCols.BMI2cValue, DataCols.BMI2dValue, DataCols.BMI2eValue)
'    Call RemoveExcessWeightBMIHeight2Values(iRow, DataCols.height2aDate, DataCols.height2bDate, DataCols.height2cDate, DataCols.height2dDate, DataCols.height2eDate, _
'    DataCols.Height2aValue, DataCols.Height2bValue, DataCols.Height2cValue, DataCols.Height2dValue, DataCols.Height2eValue)
'End Sub

'' Runs the common 2a-2e excess-result removal process for HbA1c, Weight, BMI and Height
Private Sub X9_HbA1cWeightBMIHeight2ExcessValueRemoval(iRow As Long)
    Call RemoveExcessHbA1cWeightBMIHeight2Values(iRow, DataCols.hbA1c2aDate, DataCols.hbA1c2bDate, DataCols.hbA1c2cDate, DataCols.hbA1c2dDate, DataCols.hbA1c2eDate, _
    DataCols.HbA1c2aValue, DataCols.HbA1c2bValue, DataCols.HbA1c2cValue, DataCols.HbA1c2dValue, DataCols.HbA1c2eValue) '' Retain the preferred HbA1c 2a-2e result
    Call RemoveExcessHbA1cWeightBMIHeight2Values(iRow, DataCols.weight2aDate, DataCols.weight2bDate, DataCols.weight2cDate, DataCols.weight2dDate, DataCols.weight2eDate, _
    DataCols.Weight2aValue, DataCols.Weight2bValue, DataCols.Weight2cValue, DataCols.Weight2dValue, DataCols.Weight2eValue) '' Retain the preferred Weight 2a-2e result
    Call RemoveExcessHbA1cWeightBMIHeight2Values(iRow, DataCols.bMI2aDate, DataCols.bMI2bDate, DataCols.bMI2cDate, DataCols.bMI2dDate, DataCols.bMI2eDate, _
    DataCols.BMI2aValue, DataCols.BMI2bValue, DataCols.BMI2cValue, DataCols.BMI2dValue, DataCols.BMI2eValue) '' Retain the preferred BMI 2a-2e result
    Call RemoveExcessHbA1cWeightBMIHeight2Values(iRow, DataCols.height2aDate, DataCols.height2bDate, DataCols.height2cDate, DataCols.height2dDate, DataCols.height2eDate, _
    DataCols.Height2aValue, DataCols.Height2bValue, DataCols.Height2cValue, DataCols.Height2dValue, DataCols.Height2eValue) '' Retain the preferred Height 2a-2e result
End Sub

'' Checks whether a patient's age at Index Date is below the project inclusion criterion of 50 years
'' Handles both numeric ages and 5-year age-band text if either is encountered
Private Function IsAgeBelow50(ageGroup As Variant) As Boolean
    Select Case ageGroup '' Assess the present age or age group value
        Case Is < 50 '' Numeric age is below 50 years
            IsAgeBelow50 = True
        Case "0-4", "5-9", "10-14", "15-19", "20-24", "25-29", "30-34", "35-39", "40-44", "45-49" '' Present age band is below 50 years
            IsAgeBelow50 = True
        Case Else '' Age is 50 years or older
            IsAgeBelow50 = False
    End Select
End Function

'' Replaces numeric ages with 5-year age bands to further anonymise the data
Private Sub B04_ReplaceAgeWithAgeRange(iRow As Long)
    Dim column As Integer '' Stores the column number of the age field currently being processed

    '' Replace age with age range
    '' Index Date Age
    column = DataCols.IndexDatePatientsAgeatEvent '' Set target column to Age at Index Date
    If Len(gDataSet(iRow, column)) > 0 And IsNumeric(gDataSet(iRow, column)) Then '' Only process if a numeric age is present
        gDataSet(iRow, column) = dHelperMethods_v3.GetAgeRange(CInt(gDataSet(iRow, column))) '' Convert whole-year age to corresponding 5-year age band
    End If
    '' Pancreatic Cancer diagnosis age
    column = DataCols.PancreaticcancerdiagnosisPatientsAgeatEvent '' Set target column to Age at pancreatic cancer diagnosis
    If Len(gDataSet(iRow, column)) > 0 And IsNumeric(gDataSet(iRow, column)) Then '' Only process if a numeric age is present
        gDataSet(iRow, column) = dHelperMethods_v3.GetAgeRange(CInt(gDataSet(iRow, column))) '' Convert whole-year age to corresponding 5-year age band
    End If
    
End Sub

'' Further anonymises data by replacing remaining dates with 'Present' after all date-dependent processing is complete - actually now acts only on Index Date col
Private Sub B05_ReplaceWithPresent(iRow As Long, ByVal iColumn As Long)

If IsDate(gDataSet(iRow, iColumn)) Then '' Only replace the value where a date remains
    gDataSet(iRow, iColumn) = "Present" '' Remove the actual date whilst marking that a date was present
End If

End Sub

'' Further anonymises data by replacing Anonymised Identifiers with 'Present'
Private Sub B06_RemoveAnonymisedID(iRow As Long, ByVal iColumn As Long)

If VarType(gDataSet(iRow, iColumn)) = vbString Then '' Only replace the value when the anon identifier is there
    gDataSet(iRow, iColumn) = "Present" '' Remove the actual anon ID whilst marking that it was present
End If

End Sub

'' Loads the data into an array for processing
Private Sub LoadDataIntoArray()
    Dim rng As Range '' Range variable for the complete patient data block
    Set rng = dWS.Range(dWS.Cells(gFirstRowOfData, DataCols.organisationCode), dWS.Cells(gFinalRow, DataCols.lastColumn)) '' Set range from first data cell to final data cell
    gDataSet = rng '' Load the defined worksheet range into the gDataSet array
    rng.ClearContents '' Clear the data from the spreadsheet whilst processing is underway in memory
End Sub

'' Loads the processed array back into Sheet1, removes rows marked for deletion and removes the temporary Delete column
Private Sub LoadDataIntoSheet()
    Dim rng As Range '' Range variable for the complete patient data block
    Set rng = dWS.Range(dWS.Cells(gFirstRowOfData, DataCols.organisationCode), dWS.Cells(gFinalRow, DataCols.lastColumn)) '' Set range from first data cell to final data cell
    rng = gDataSet '' Write the processed gDataSet array back into the defined worksheet range
    
    '' Sort the temporary Delete column so all rows marked True bubble up together at the top
    '' This allows all marked patient rows to be removed with one worksheet row deletion rather than deleting them individually
    Dim sortRange As String '' Stores the header cell reference of the temporary Delete column
    sortRange = Split((Columns(DataCols.lastColumn).Address(, 0)), ":")(0) & CStr(1) '' Convert temporary Delete column number to its column letter and row 1 address
    Dim header As Range
    Set header = dWS.Range(dWS.Cells(1, DataCols.organisationCode), dWS.Cells(1, DataCols.lastColumn)) '' Define header range including temporary Delete column
    header.AutoFilter '' Turn on AutoFilter to enable sorting
    dWS.AutoFilter.Sort.SortFields.Clear '' Clear any existing sort criteria
    dWS.AutoFilter.Sort.SortFields.Add Key:= _
        Range(sortRange), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal '' Sort Delete column descending so True values are moved to the top
    With dWS.AutoFilter.Sort '' Configure and run the sort
        .header = xlYes '' Tell Excel that row 1 is a header
        .MatchCase = False '' Do not use case sensitive sorting
        .Orientation = xlTopToBottom '' Sort rows vertically
        .SortMethod = xlPinYin '' Use standard Excel sort method
        .Apply '' Apply the defined sort
    End With
    
    '' Determine the final row of the contiguous deletion block by counting the patients marked True
    Dim i As Long '' Counter used to loop through patient rows in gDataSet
    Dim rowsToDel As Long '' Used to determine the last worksheet row of the sorted deletion block
    rowsToDel = 1 '' Start with the row 1 header offset; each marked patient increases this by one
    
    For i = LBound(gDataSet, 1) To UBound(gDataSet, 1) '' Loop through every processed patient row
        If gDataSet(i, DataCols.lastColumn) Then '' If temporary Delete field is True for this row
            rowsToDel = rowsToDel + 1 '' Add one to the number of rows to delete
        End If
    Next i '' Move to next patient row
    
    '' Delete the single block containing all rows marked True for deletion after the preceding sort
    If rowsToDel > 1 Then '' Only delete if at least one patient row was marked
        dWS.Rows(CStr(gFirstRowOfData) & ":" & CStr(rowsToDel)).Delete Shift:=xlUp '' Delete marked rows and shift remaining rows upwards
    End If
    
    '' Delete the temporary Delete column given no longer required
    dWS.Columns(DataCols.lastColumn).EntireColumn.Delete
    
    '' Switch off AutoFilter used for sorting
    dWS.AutoFilterMode = False
    
    '' Add a blank-space header at the very end of the worksheet structure - useful marker given deletion mechanism troubles
    dWS.Cells(1, DataCols.lastColumn).value2 = " "
    
    '' Recalculate gFinalRow after deleting patient rows
    gFinalRow = dWS.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row '' Store the row number of the last remaining populated row
    
End Sub

'' Checks copied and pasted EMIS/already-transformed TPP data, removes export artefacts and applies the standard DEFEND PRIME headings
Private Sub CheckData()
    Dim i As Integer '' Counter used for row and column loops

    '' Initially assume patient data begin in worksheet row 1, but if a header is found or inserted this is changed to row 2
    gFirstRowOfData = 1
    
    Dim hasHeaders As Boolean
    hasHeaders = True ''' Assume EMIS metadata/header rows may be present above the standard headings unless Organisation Code is already in A1
    
    If dWS.Cells(1, 1).value2 = "Organisation Code" Then '' Organisation Code already in A1 means the standard column-header row is already row 1
        '' There is no full higher-level EMIS header, either because it was exported without it or only the bottom header row was selected when copy/pasting
        hasHeaders = False '' No metadata/header rows above the standard headings need to be removed
        gFirstRowOfData = 2 '' Patient data begin immediately beneath the header row
    End If
    
    If hasHeaders Then
        '' If Organisation Code was not already present in A1, search for it lower down
        For i = 1 To 30
            If dWS.Cells(i, 1).value2 = "Organisation Code" Then
                Exit For
            End If
        Next i
        
        '' Only delete preceding rows if Organisation Code was actually found
        If i <= 30 And dWS.Cells(i, 1).value2 = "Organisation Code" Then
            dWS.Rows("1:" & CStr(i - 1)).EntireRow.Delete
            gFirstRowOfData = 2
        End If
    End If
    
    '' Stop before altering the worksheet further if the expected EMIS header was not identified
    If gFirstRowOfData = 1 Then
        MsgBox "Sorry, the pasted data do not appear to have the expected DEFEND PRIME EMIS Web structure, so processing has been stopped. Please delete all data from Sheet1 and repaste the original extracted data, including all headers, into cell A1. If the problem persists, please contact the DEFEND PRIME team noting 'Error 7'.", vbExclamation
        gFirstRowOfData = 0
        Exit Sub
    End If
    
    '' Next issue is if data was exported in xls format we may have some cells merged together this bug is a bit unpredictable.
    '' Not sure if the cause is tied to the data or the name of the column. Either way it could change with EMIS updates.
    '' Suggested fix is to loop backwards across the header and delete any columns where the header is vbnullstring.
    '' This does mean that if the data are pasted without any headers at all and there are blank columns (e.g. pancreatic cancer diagnosis)
    '' then the subsequent columns are knocked. This should not happen with the main data columns.
    '' Possible remedy is relational checks with all blank columns and contents of those in close proximity, but that would fail if more than 1 blank column.
    For i = 90 To 1 Step -1 '' Check possible columns from right to left
        If dWS.Cells(1, i).value2 = vbNullString Then '' If the header cell for this column is blank
            dWS.Cells(1, i).EntireColumn.Delete '' Delete the blank structural column
      End If
    Next i '' Move one column to the left
   
    
    '' Validate the prepared worksheet structure before applying the standard DEFEND PRIME headers
    Dim validHeader As Boolean
    Dim validRawEMISHeader As Boolean
    Dim validDEFENDPRIMEHeader As Boolean
    
    '' Check a few positions across the pasted columns against the expected raw EMIS export headings
    '' These checks detect unexpected populated columns or other structural issues - should be unlikely given the XML defines the output
    validRawEMISHeader = True
    If dWS.Cells(1, 82).value2 <> "Date" Then validRawEMISHeader = False
    If dWS.Cells(1, 70).value2 <> "Unit of Measure" Then validRawEMISHeader = False
    If dWS.Cells(1, 15).value2 <> "Value" Then validRawEMISHeader = False
    If dWS.Cells(1, 7).value2 <> "Patient's Age at Event" Then validRawEMISHeader = False
    If dWS.Cells(1, 2).value2 <> "Anonymised Identifier" Then validRawEMISHeader = False
    
    '' Also allow data already using the DEFEND PRIME headers, such as transformed TPP data
    validDEFENDPRIMEHeader = True
    If dWS.Cells(1, 82).value2 <> gHeaderTexts(82) Then validDEFENDPRIMEHeader = False
    If dWS.Cells(1, 70).value2 <> gHeaderTexts(70) Then validDEFENDPRIMEHeader = False
    If dWS.Cells(1, 15).value2 <> gHeaderTexts(15) Then validDEFENDPRIMEHeader = False
    If dWS.Cells(1, 7).value2 <> gHeaderTexts(7) Then validDEFENDPRIMEHeader = False
    If dWS.Cells(1, 2).value2 <> gHeaderTexts(2) Then validDEFENDPRIMEHeader = False
    
    '' Accept either supported header form
    validHeader = validRawEMISHeader Or validDEFENDPRIMEHeader
    
    '' Stop before processing if the expected contents couldn't be confirmed
    If validHeader = False Then
        MsgBox "Sorry, the pasted data do not appear to have the expected DEFEND PRIME EMIS Web structure, so processing has been stopped. Please delete all data from Sheet1 and repaste the original extracted data, including all headers, into cell A1. If the problem persists, please contact the DEFEND PRIME team noting 'Error 7'.", vbExclamation
        gFirstRowOfData = 0 '' Signal to the main procedure that CheckData failed
        Exit Sub
    End If
    
    '' Replace existing headings with the standard DEFEND PRIME header array after structural preparation
    With dWS
        .Range(.Cells(1, DataCols.organisationCode), .Cells(1, DataCols.lastColumn - 1)).value2 = gHeaderTexts '' Write the standard project headers across all output columns
    End With
    
    '' Find the final populated row after the input structure has been prepared
    gFinalRow = dWS.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    '' Add temporary final column used to mark rows for deletion during processing
    dWS.Cells(1, DataCols.lastColumn).value2 = "Delete"

End Sub

'' Applies borders and standard column/header formatting to the final processed dataset
Private Sub ApplyBorders()
    '' Format Data Range
    With dWS.Range(dWS.Cells(gFirstRowOfData, DataCols.organisationCode), dWS.Cells(gFinalRow, DataCols.lastColumn - 1))
        With .Borders(xlEdgeLeft) '' Format left outer border
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop) '' Format top outer border
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom) '' Format bottom outer border
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight) '' Format right outer border
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlInsideVertical) '' Format borders between data columns
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlInsideHorizontal) '' Format borders between patient rows
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    '' Apply final formatting to header row
    With dWS.Range(dWS.Cells(1, DataCols.organisationCode), dWS.Cells(1, DataCols.lastColumn - 1))
        '' Fix column widths
        .Columns.ColumnWidth = 9.78 '' Apply consistent width across output columns
        '' Format header text
        .Font.Bold = True '' Make headings bold
        With .Borders(xlEdgeLeft) '' Format left outer header border
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop) '' Format top outer header border
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom) '' Format bottom outer header border
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight) '' Format right outer header border
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlInsideVertical) '' Format borders between header cells
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End With
End Sub

'' Clears any active AutoFilter and applies consistent visual date formatting to all date columns used in processing
Private Sub FormatDataSheet()
    
    '' Clear any active AutoFilter so no rows remain hidden during processing
    If dWS.AutoFilterMode Then
        dWS.AutoFilterMode = False
    End If

    '' Apply dd/mm/yyyy formatting to date columns; @ leaves text values unchanged
    With dWS
        .Cells(1, DataCols.bMI1aDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.bMI1bDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.bMI2aDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.bMI2bDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.bMI2cDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.bMI2dDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.bMI2eDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.Cancerantigen199CA199Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.CarcinoembryonicantigenCEADate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.hbA1c2aDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.hbA1c2bDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.hbA1c2cDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.hbA1c2dDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.hbA1c2eDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.height1aDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.height1bDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.height2aDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.height2bDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.height2cDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.height2dDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.height2eDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.weight1aDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.weight1bDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.weight2aDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.weight2bDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.weight2cDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.weight2dDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.weight2eDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.IndexDateDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, DataCols.PancreaticcancerdiagnosisDate).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
    End With
End Sub





