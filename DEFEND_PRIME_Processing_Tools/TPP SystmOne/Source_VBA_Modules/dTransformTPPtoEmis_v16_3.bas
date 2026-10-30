Attribute VB_Name = "dTransformTPPtoEmis_v16_3"
'' To run this DEFEND PRIME TPP SystmOne extract to EMIS Web format transformation software, in the Microsoft Visual Basic for Applications window, please follow these steps:
'' 1. Click the 'Tools' tab above, in the drop-down list click 'Macros...', then with 'A_TransformTPPtoEMIS' selected, click 'Run'.
'' 2. If any error messages are reported, please contact Hugh Claridge (h.claridge@surrey.ac.uk) from the DEFEND PRIME team.
'' Programme Identifier: DEFEND PRIME TPP to EMIS transformation Excel macro
'' Name of programmer: Hugh Claridge
'' Brief description of programme: This Excel VBA macro processes the DEFEND PRIME data extracted from TPP SystmOne by the provided extraction file into a similar format as the extract from EMIS Web to enable subsequent processing.
'' Revision history for dTransformTPPtoEmis
'' Differs from v16_2 due to implementing ReadyApplication from dHelperMethods and improving order of checks
'' Version number: v16_3 or v16.3
'' Release Date: 24th October 2025

Option Explicit

'' Converts TPP SystmOne data in Sheet1 into EMIS Web format
'' Stores the first row number for entire module use after removal of e.g. metadata rows from Sheet1 (not encountered yet but TPP format may change in future)
Private tppFirstRowOfData As Long

'' Enum to reference TPP columns by name instead of hardcoding indices - more flexible if col nums change in future, and easier to read code
Private Enum TPPCols
    OrganisationID = 1
    DoB = 2
    Gender = 3
    HbA1c1 = 4
    HbA1c1Date = 5
    HbA1c2 = 6
    HbA1c2Date = 7
    HbA1c3 = 8
    HbA1c3Date = 9
    HbA1c4 = 10
    HbA1c4Date = 11
    HbA1c5 = 12
    HbA1c5Date = 13
    HbA1c6 = 14
    HbA1c6Date = 15
    HbA1c7 = 16
    HbA1c7Date = 17
    HbA1c8 = 18
    HbA1c8Date = 19
    HbA1c9 = 20
    HbA1c9Date = 21
    Weight1 = 22
    Weight1Date = 23
    Weight2 = 24
    Weight2Date = 25
    Weight3 = 26
    Weight3Date = 27
    Weight4 = 28
    Weight4Date = 29
    Weight5 = 30
    Weight5Date = 31
    Weight6 = 32
    Weight6Date = 33
    Weight7 = 34
    Weight7Date = 35
    Weight8 = 36
    Weight8Date = 37
    Weight9 = 38
    Weight9Date = 39
    BMI1 = 40
    BMI1Date = 41
    BMI2 = 42
    BMI2Date = 43
    BMI3 = 44
    BMI3Date = 45
    BMI4 = 46
    BMI4Date = 47
    BMI5 = 48
    BMI5Date = 49
    BMI6 = 50
    BMI6Date = 51
    BMI7 = 52
    BMI7Date = 53
    BMI8 = 54
    BMI8Date = 55
    BMI9 = 56
    BMI9Date = 57
    Height1 = 58
    Height1Date = 59
    Height2 = 60
    Height2Date = 61
    Height3 = 62
    Height3Date = 63
    Height4 = 64
    Height4Date = 65
    Height5 = 66
    Height5Date = 67
    Height6 = 68
    Height6Date = 69
    Height7 = 70
    Height7Date = 71
    Height8 = 72
    Height8Date = 73
    Height9 = 74
    Height9Date = 75
    CA199_1 = 76
    CA199_1_Date = 77
    CA199_2 = 78
    CA199_2_Date = 79
    CEA_1 = 80
    CEA_1_Date = 81
    CEA_2 = 82
    CEA_2_Date = 83
    PCancer_1 = 84
    PCancer_1_Date = 85
    PCancer_2 = 86
    PCancer_2_Date = 87
    PCancer_3 = 88
    PCancer_3_Date = 89
End Enum


Sub A_TransformTPPtoEMIS()
    
    '' If something goes wrong, jump to ErrHandler to prevent silent fails
    On Error GoTo ErrHandler
    
    '' From dGlobals, cannot trust anything left in the shared global variables from a previous run; force the current run to establish its own worksheet, row positions, dataset and headers.
    Call ClearGlobalData
    
    '' Comment in Excel status bar so user knows what's going on
    Application.StatusBar = "Checking worksheet formatting and presence of TPP data in preparation for transformation"
    
    ''Check that Sheet1 exists in the workbook being used, and if it does, ensure that it's the active worksheet so that people have Sheet1 open
    If Not Evaluate("ISREF(Sheet1!A1)") Then '' If Sheet1 doesn't exist
        MsgBox "Sorry, the DEFEND PRIME TPP processing macro requires the data to be present in the worksheet entitled 'Sheet1' but this worksheet cannot be found. Please ensure the correct sheet is present and try again. If the problem persists, please contact the DEFEND PRIME team noting 'Error 1a'.", vbExclamation
        Application.StatusBar = False '' Clear the status bar before stopping
        Exit Sub '' Stop before changing worksheet data, Excel performance settings or user controls
    ElseIf ThisWorkbook.ActiveSheet.Name <> "Sheet1" Then '' If Sheet1 isn't the active worksheet (basically the one clicked on and being looked at)
        MsgBox "Sorry, Sheet1 is not currently the active worksheet, so please ensure that that worksheet is open and it contains the extracted data before this TPP processing macro is run. If the problem persists, please contact the DEFEND PRIME team noting 'Error 2a'.", vbExclamation
        Application.StatusBar = False '' Clear the status bar before stopping
        Exit Sub '' Stop before changing worksheet data, Excel performance settings or user controls
    End If
    
    '' Check if the DEFEND PRIME TPP processing has already been run by looking for text indicating processing in cells A1 and B1
    '' Double processing removes the required values from the already-processed worksheet, thus this prevents accidental re-running
    If ThisWorkbook.Sheets("Sheet1").Range("A1").Value = "Organisation Code" And _
       ThisWorkbook.Sheets("Sheet1").Range("B1").Value = "Anonymised Identifier" Then
        MsgBox "Sorry, TPP processing appears to have already been performed as processed data appear to be present in Sheet1 thus re-processing by the DEFEND PRIME macro has been stopped. If you believe this to be incorrect please contact the DEFEND PRIME team noting 'Error 3a'.", vbExclamation
        Application.StatusBar = False '' Clear the status bar before stopping
        Exit Sub '' Stop before changing worksheet data, Excel performance settings or user controls
    End If
    
    '' Check if someone is accidentally trying to run this on a blank worksheet, having perhaps forgotten to paste the data in, so check if A1 is blank or contains stock message
     If ThisWorkbook.Sheets("Sheet1").Range("A1").Value = "" Or _
        ThisWorkbook.Sheets("Sheet1").Range("A1").Value = "Please copy and paste your data into this cell" Then
        MsgBox "Sorry, the expected TPP extract data do not appear to have been pasted into cell A1 of Sheet1. Please delete any data you have pasted in and repaste the original extract data into cell A1 of Sheet1. If this problem persists, please contact the DEFEND PRIME team noting 'Error 4a'.", vbExclamation
        Application.StatusBar = False '' Clear the status bar before stopping
        Exit Sub '' Stop before changing worksheet data, Excel performance settings or user controls
    End If
    
    '' Make it clear to user that something is happening
    Application.Cursor = xlWait '' Change cursor from standard to hour glass
    SafeDisableButton ActiveSheet, "btnRunTPP" '' Prevent the macro run button from being reclickable whilst macro is running
    SafeDisableButton ActiveSheet, "btnRunDP" '' Prevent the second macro run button from being clickable whilst macro is running

    Application.StatusBar = "Running TPP transformation macro, this may take up to a minute to complete" '' Comment in Excel status bar
    
    Dim sourceSheet As Worksheet '' Worksheet containing the raw TPP input data
    Set sourceSheet = ThisWorkbook.Sheets("Sheet1") '' Set sourceSheet to Sheet1 in this workbook
    Call SetWorksheet '' From dGlobals for EMIS Web output in Sheet1, setting dWS as Sheet1 thus the TPP transformed to EMIS output sheet
    Call SetHeaderValues '' From dGlobals for EMIS Web output in Sheet1, populating gHeaderTexts(1 To 82) array with EMIS Web output column headers
    Call ReadyApplication(processingMode:=True) '' From dHelperMethods, explicitly sets Excel to reduced-overhead settings for transformation

'9999     Err.Raise 12345, , "Deliberate test of TPP error handler" '' Uncomment the start of this line to test error handler

    tppFirstRowOfData = 0 '' Set to 0 before checking TPP data
    Call CheckTPPData '' Modified version of CheckData from dMain checking and tidying up Sheet1 before processing
    If tppFirstRowOfData = 0 Then GoTo ExitHere '' CheckTPPData failed, so cleanly stop processing
    
    '' Check that patient data are present before altering or formatting the validated TPP data - uses DoB given that's one we can be most sure of people having recorded
    Dim lastRow As Long '' Stores the final TPP patient row identified from the DoB column
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 2).End(xlUp).Row '' Find last non-empty row in Col B of TPP data using xlUp upward scan starting from total number of rows in worksheet
    If lastRow < tppFirstRowOfData Then '' If it finds the last row of data to be less than row 2, then trigger message and exit
        MsgBox "Sorry, the expected TPP extract data do not appear to have been pasted into cell A1 of Sheet1 as no Dates of Birth are present in the DoB column. Please delete any data you have pasted in and repaste the original extract data into cell A1 of Sheet1. If this problem persists, please contact the DEFEND PRIME team noting 'Error 6a'.", vbExclamation
        GoTo ExitHere
    End If
    
    '' Continue preparing the validated TPP data for processing
    Call FixHyphenError(sourceSheet) '' From dHelperMethods to help tidy up any problematic hyphens
    Call FormatTPPInputSheet '' Clears autofilters so all data are visible to macro and formats date cells to be consistent for subsequent processing
    
    '' Load TPP data into memory
    Dim inputData As Variant '' Array holding the raw TPP input data loaded from Sheet1
    inputData = sourceSheet.Range(sourceSheet.Cells(tppFirstRowOfData, 1), sourceSheet.Cells(lastRow, 89)).Value '' Loads block of data into memory from Sheet1 (essentially row 2 to last row found spanning 89 cols)
    Dim outputData() As Variant '' Array used to store the transformed 82 col output data
    ReDim outputData(1 To lastRow - 1, 1 To UBound(gHeaderTexts)) '' Now size the output array for all patient rows and all DEFEND PRIME output columns

    Dim i As Long '' Declare variable i as row loop counter
    Dim outputRowIndex As Long '' Tracks the row to use when writing transformed results to outputData
    outputRowIndex = 1 '' Initialise the output row counter to the first row to start writing data here
    
    '' Shared variables for HbA1c, height, weight, BMI windows
    Dim rangeStart As Long, rangeEnd As Long '' Variables for start and end points for windows post- but mostly pre-index
    Dim dateStrategy As String '' Variable for selection strategy of highest, earliest or latest
    Dim outputValCol As Long, outputDateCol As Long, outputUnitCol As Long '' Variables to write extracted values/dates/UoMs to

    '' We know that the dataset contains results (and their qualifying inclusion result) for those with an HbA1c >= 48 mmol/mol for the first time in the last three years
    '' But we don't know which of the up to 9 results and dates this is. We therefore need to find these.
    '' We also need to get dates of birth to calculate age at index date, when found.
    
    For i = 1 To UBound(inputData, 1) '' Loop through the possible rows of TPP input data
        If i Mod 10 = 0 Then DoEvents '' Maintains Excel responsiveness every 10 rows processed (increase row number for faster processing but less responsiveness)
        
        '' Extract date of birth and prepare index variables
        Dim dateOfBirth As Variant '' Stores the DoB for the current row
        dateOfBirth = inputData(i, TPPCols.DoB) '' Get the DoB from the input DoB col and assign it to the dateOfBirth variable for this row
        Dim indexDate As Date, indexHbA1cValue As Double, numericHbA1c As Double, indexHbA1cRaw As String '' Variables for index HbA1c date and result and general HbA1c numeric result
        Dim allHbA1cValues(1 To 9) As Variant, allHbA1cDates(1 To 9) As Variant '' Declare arrays to hold the 9 possible HbA1c results and their dates

        '' Collect all 9 HbA1c values and their dates from input
        Dim hba1cResultNumCounter As Long  '' Counter to loop through the HbA1c result and date pairs
        For hba1cResultNumCounter = 1 To 9 '' Nine loops as nine result pairs
            allHbA1cValues(hba1cResultNumCounter) = inputData(i, TPPCols.HbA1c1 + (hba1cResultNumCounter - 1) * 2) '' Dynamically calculate col number for value (first HbA1c result col: HbA1c1 = 4+(1-1)*2= 4+0 = 4), second result is 4+(2-1)*2= 4+2 = 6
            allHbA1cDates(hba1cResultNumCounter) = inputData(i, TPPCols.HbA1c1Date + (hba1cResultNumCounter - 1) * 2) '' Dynamically calculate col number for date as above
        Next hba1cResultNumCounter '' Onto next result/date pair

        '' Identify index date, being the earliest qualifying HbA1c >=48 between Dec 2021 and Dec 2024
        indexDate = 0 '' Set earliest date to zero, i.e. clear it fresh for the loops
        For hba1cResultNumCounter = 1 To 9 '' Nine loops as nine result pairs
            numericHbA1c = ExtractNumeric(allHbA1cValues(hba1cResultNumCounter)) '' Get the numerical result out of the cell given also present with UoM
            If numericHbA1c >= 48 And IsDate(allHbA1cDates(hba1cResultNumCounter)) Then '' Checks if there's a numerical result meeting the inclusion criteria and there's a date to go with it, if not, skip
                If allHbA1cDates(hba1cResultNumCounter) >= DateSerial(2021, 12, 1) And allHbA1cDates(hba1cResultNumCounter) <= DateSerial(2024, 12, 1) Then '' If in the three-year window for inclusion
                    If indexDate = 0 Or allHbA1cDates(hba1cResultNumCounter) < indexDate Then '' If indexDate remains 0 then this one could be the first qualifying result thus far, or check if earlier than an already-stored qualifying date
                        indexDate = allHbA1cDates(hba1cResultNumCounter) '' If yes then save this date as the Index Date for that row
                        indexHbA1cValue = numericHbA1c '' If yes then save this result as the Index HbA1c Result for that row for comparisons
                        indexHbA1cRaw = CStr(allHbA1cValues(hba1cResultNumCounter)) '' And keep this for printing (in case a comparator was present, want that visually maintained if not for comparisons)
                    End If
                End If
            End If
        Next hba1cResultNumCounter '' Onto next result

        '' If no index HbA1c found, mark and skip further processing for this row given everything relates to Index HbA1c date - should be impossible given extraction criteria
        If indexDate = 0 Then '' If the indexDate remains as 0 after the previous section that means no qualifying result and date were found
            outputData(outputRowIndex, 1) = "No valid Index HbA1c >= 48 in date window" '' Marks the row with the issue
            GoTo FinaliseRow '' Jumps past further processing of this row given no index date means nothing can be done with the row
        End If

        '' Assign the saved Index Date and Index HbA1c Result to the defined columns for this row
        outputData(outputRowIndex, DataCols.IndexDateDate) = indexDate '' Write index date to output column in array
        outputData(outputRowIndex, DataCols.IndexDateValue) = FormatMeasuredValue(indexHbA1cRaw) '' Write index result to output column in array

        
        '' Calculate age at index
        If IsDate(indexDate) And IsDate(dateOfBirth) And CDbl(dateOfBirth) <> 0 Then '' Ensure both are dates and DoB is not 0 serial (IsDate takes blanks as 0 = 30th Dec 1899)
            Dim ageAtIndex As Long '' Store age in whole years
            ageAtIndex = AgeYears(dateOfBirth, indexDate)  '' Calculates whole-year age at index
        
            If ageAtIndex >= 0 Then '' An age above 0 is valid for these purposes
                outputData(outputRowIndex, DataCols.IndexDatePatientsAgeatEvent) = ageAtIndex '' Output the calculated age
            Else '' But if the AgeYears function has returned -1 as per defined error process
                outputData(outputRowIndex, DataCols.IndexDatePatientsAgeatEvent) = "DoB or Index dates missing or invalid format" '' Mark the DoB cell as such
            End If
        Else '' If other checks fail then they are either not dates or DoB was blank/0 serial
            outputData(outputRowIndex, DataCols.IndexDatePatientsAgeatEvent) = "DoB or Index dates missing or invalid format" '' Marks the DoB cell as such
        End If

        '' Define pre-index date ranges and target columns for HbA1c 2a to 2e using array and search through the 9 possible HbA1c results and dates for 'best' (highest) results in each ENDPAC window
        Dim hba1cWindowSettings(1 To 5) As Variant '' Array to store the 5 ENDPAC HbA1c time window settings
        hba1cWindowSettings(1) = Array(274, 455, DataCols.HbA1c2aValue, DataCols.hbA1c2aDate)  '' Window 1 (2a): rangeStart lower signed day-difference bound (+274 = 274 days before index), rangeEnd upper signed day-difference bound (+455 = 455 days before index), two cols to store the best result and associated date
        hba1cWindowSettings(2) = Array(456, 546, DataCols.HbA1c2bValue, DataCols.hbA1c2bDate) '' Window 2 (2b) etc.
        hba1cWindowSettings(3) = Array(183, 273, DataCols.HbA1c2cValue, DataCols.hbA1c2cDate)
        hba1cWindowSettings(4) = Array(92, 182, DataCols.HbA1c2dValue, DataCols.hbA1c2dDate)
        hba1cWindowSettings(5) = Array(547, 728, DataCols.HbA1c2eValue, DataCols.hbA1c2eDate)

        '' For each pre-index HbA1c time window, find the highest result within the time frame
        Dim hba1cWindowCounter As Long
        For hba1cWindowCounter = 1 To 5 '' Loop over the five defined windows
            Dim bestHba1cValue As Variant '' Stores the best numeric HbA1c value found for the current window
            bestHba1cValue = "" '' Clear numeric value for picking
            Dim bestHbA1cRaw As String '' Stores the original result text so any comparator can be preserved in the output
            bestHbA1cRaw = vbNullString '' Clear the stored result ready for this window
            Dim bestHba1cDate As Date '' Stores the date associated with the best HbA1c value
            bestHba1cDate = 0 '' Clear best value's date
        
            '' Extract settings from each window from the above hba1cWindowSettings array
            rangeStart = hba1cWindowSettings(hba1cWindowCounter)(0) '' Element 0 being the first item in the settings array, i.e. 274 for Window 1 (2a) thus rangeStart
            rangeEnd = hba1cWindowSettings(hba1cWindowCounter)(1) '' Element 1 being the second item in the settings array, i.e. 455 for Window 1 (2a) thus rangeEnd
            outputValCol = hba1cWindowSettings(hba1cWindowCounter)(2)
            outputDateCol = hba1cWindowSettings(hba1cWindowCounter)(3)

        
            For hba1cResultNumCounter = 1 To 9 '' Loop through the 9 possible HbA1c results for each row
                numericHbA1c = ExtractNumeric(allHbA1cValues(hba1cResultNumCounter)) '' Get the numerical result out of the cell given also present with UoM and no need to retain UoM for Hba1c for EMIS transformation
                If numericHbA1c > 0 And IsDate(allHbA1cDates(hba1cResultNumCounter)) Then '' Checks if there's a numerical result present and a date to go with it, if not, skip
                    Dim daysDiff As Long '' Stores the number of days between the HbA1c result and Index Date
                    daysDiff = DateDiff("d", allHbA1cDates(hba1cResultNumCounter), indexDate) '' Calculates how many days before the index date the result is (indexDate minus result date, reverse of what's shown because VBA)
                    If daysDiff >= rangeStart And daysDiff <= rangeEnd Then '' Does the calculated day difference fall into the valid range for the window
                        If bestHba1cValue = "" Or numericHbA1c > bestHba1cValue Then '' If result is the first valid one (otherwise would already be populated) or it's better than populated one
                            bestHba1cValue = numericHbA1c '' If yes then store the result
                            bestHbA1cRaw = CStr(allHbA1cValues(hba1cResultNumCounter))  '' Keep original raw result (may include comparator)
                            bestHba1cDate = allHbA1cDates(hba1cResultNumCounter) '' And store the result's date
                        End If
                    End If
                End If
            Next hba1cResultNumCounter '' Onto next result
        
            '' After looping through the 9 possible results against the defined window, if a 'best' result was found then save it and its date in the output columns for that window
            If bestHba1cValue <> "" Then
                outputData(outputRowIndex, outputValCol) = FormatMeasuredValue(bestHbA1cRaw)  '' Write the formatted HbA1c result (e.g. '> 120') to the output array, preserving any comparator
                outputData(outputRowIndex, outputDateCol) = bestHba1cDate '' Write the date to output column in array
            End If
        Next hba1cWindowCounter '' Onto the next defined window until all cycled through

        '' Extract gender
        Dim genderText As String '' Stores gender for the current row
        genderText = Trim(CStr(inputData(i, TPPCols.Gender))) '' Get gender for current row, convert to string and trim any trailing spaces
        If Len(genderText) > 0 Then '' Check cell not blank
            outputData(outputRowIndex, DataCols.Gender) = genderText '' If not blank then store gender in the defined col of the output array
        End If

        '' Extract Organisation ID/code/GP practice identifier
        Dim organisationCode As String '' Stores organisation code/ID for the current row
        organisationCode = Trim(CStr(inputData(i, TPPCols.OrganisationID))) '' Get org code/ID for current row, convert to string and trim
        If Len(organisationCode) > 0 Then '' Check cell not blank
            outputData(outputRowIndex, DataCols.organisationCode) = organisationCode '' If not blank then store org ID in defined col of output array
        End If

        '' Extract CA19-9 values, choose earliest valid date, and extract value and unit for EMIS Web
        Dim ca199ValueRaw As String '' Stores the raw CA19-9 value with UoM, if any
        Dim ca199EarliestDate As Date '' Stores the earliest CA19-9 date found
        Dim ca199ResultNumCounter As Long '' Counter to loop through the CA19-9 result and date pairs
        ca199EarliestDate = 0 '' Set earliest date to zero, i.e. clear it fresh for the two loops
        For ca199ResultNumCounter = 1 To 2 '' Two loops as two result pairs
            Dim ca199ValCol As Long '' Stores the dynamically calculated CA19-9 value column
            ca199ValCol = TPPCols.CA199_1 + (ca199ResultNumCounter - 1) * 2 '' Dynamically calculate col number for value (first ca19-9 result col: CA199_1 = 76+(1-1)*2= 76+0 = 76), second result is 76+(2-1)*2= 76+2 = 78
            Dim ca199DateCol As Long '' Stores the dynamically calculated CA19-9 date column
            ca199DateCol = TPPCols.CA199_1_Date + (ca199ResultNumCounter - 1) * 2 '' Dynamically calculate col number for date as above
            If IsDate(inputData(i, ca199DateCol)) And HasMeasurement(inputData(i, ca199ValCol)) Then '' Ensure date cell contains a date and accompanying result (given issue of dates recorded when test ordered not actually performed) - different to HbA1c as 0 is legitimate result
                If ca199EarliestDate = 0 Or inputData(i, ca199DateCol) < ca199EarliestDate Then '' If result is the first valid one (otherwise would not be 0) or it's earlier than currently populated one
                    ca199EarliestDate = inputData(i, ca199DateCol) '' If yes then store the date
                    ca199ValueRaw = inputData(i, ca199ValCol) '' If yes then store the result
                End If
            End If
        Next ca199ResultNumCounter '' Onto the next result
        If ca199EarliestDate <> 0 And Len(ca199ValueRaw) > 0 Then '' Checks a date was found and the result for it isn't blank either
            outputData(outputRowIndex, DataCols.Cancerantigen199CA199Value) = FormatMeasuredValue(ca199ValueRaw) '' Write result even if no UoM present to output column
            outputData(outputRowIndex, DataCols.Cancerantigen199CA199UnitofMeasure) = ExtractUnit(ca199ValueRaw) '' Write UoM if present  e.g. u/mL to output column in array
            outputData(outputRowIndex, DataCols.Cancerantigen199CA199Date) = ca199EarliestDate '' Store associated valid date for the earliest result
        End If

        '' Extract CEA values, choose earliest valid date, and extract value and unit for EMIS Web
        Dim ceaValueRaw As String '' Stores the raw CEA value with UoM, if any
        Dim ceaEarliestDate As Date '' Stores the earliest CEA date found
        Dim ceaResultNumCounter As Long '' Counter to loop through the CEA result and date pairs
        ceaEarliestDate = 0 '' Set earliest date to zero, i.e. clear it fresh for the two loops
        For ceaResultNumCounter = 1 To 2 '' Two loops as two result pairs
            Dim ceaValCol As Long '' Stores the dynamically calculated CEA value column
            ceaValCol = TPPCols.CEA_1 + (ceaResultNumCounter - 1) * 2  '' Dynamically calculate col number for value (first CEA result col: CEA_1 = 80+(1-1)*2= 80+0 = 80), second result is 80+(2-1)*2= 80+2 = 82
            Dim ceaDateCol As Long '' Stores the dynamically calculated CEA date column
            ceaDateCol = TPPCols.CEA_1_Date + (ceaResultNumCounter - 1) * 2 '' Dynamically calculate col number for date as above
            If IsDate(inputData(i, ceaDateCol)) And HasMeasurement(inputData(i, ceaValCol)) Then '' Ensure date cell contains a date and accompanying result (given issue of dates recorded when test ordered not actually performed) - different to HbA1c as 0 is legitimate result
                If ceaEarliestDate = 0 Or inputData(i, ceaDateCol) < ceaEarliestDate Then '' If result is the first valid one (otherwise would not be 0) or it's earlier than currently populated one
                    ceaEarliestDate = inputData(i, ceaDateCol) '' If yes then store the date
                    ceaValueRaw = inputData(i, ceaValCol) '' If yes then store the result
                End If
            End If
        Next ceaResultNumCounter '' Onto the next result
        If ceaEarliestDate <> 0 And Len(ceaValueRaw) > 0 Then '' Checks a date was found and the result for it isn't blank either
            outputData(outputRowIndex, DataCols.CarcinoembryonicantigenCEAValue) = FormatMeasuredValue(ceaValueRaw) '' Write result even if no UoM present to output column
            outputData(outputRowIndex, DataCols.CarcinoembryonicantigenCEAUnitofMeasure) = ExtractUnit(ceaValueRaw) '' Write UoM if present e.g. ug/L to output column in array
            outputData(outputRowIndex, DataCols.CarcinoembryonicantigenCEADate) = ceaEarliestDate '' Write associated valid date for the earliest result
        End If

        '' Extract pancreatic cancer data, pick earliest valid date, no pancreatic cancer code to output into EMIS just date
        Dim pcancerText As String '' Stores the pancreatic cancer diagnosis text if any
        Dim pcancerDate As Date '' Stores the earliest pancreatic cancer result recorded (given diagnostic codes are often repeatedly input over time)
        Dim pcancerResultNumCounter As Long '' Counter to loop through the pancreatic cancer result and date pairs
        pcancerDate = 0 '' Set earliest date to zero, i.e. clear it fresh for the three loops
        For pcancerResultNumCounter = 1 To 3 '' Three loops as three result pairs
            Dim pcancerCol As Long '' Stores the dynamically calculated pancreatic cancer result column
            pcancerCol = TPPCols.PCancer_1 + (pcancerResultNumCounter - 1) * 2  '' Dynamically calculate col number for value (first pcancer result col: PCancer_1 = 84+(1-1)*2= 84+0 = 84), second result is 84+(2-1)*2= 84+2 = 86 etc.
            Dim pcancerDateCol As Long '' Stores the dynamically calculated pancreatic cancer date column
            pcancerDateCol = TPPCols.PCancer_1_Date + (pcancerResultNumCounter - 1) * 2 '' Same logic as above for three dates
            If IsDate(inputData(i, pcancerDateCol)) And Len(Trim(inputData(i, pcancerCol))) > 0 Then '' Ensure date cell contains a date and accompanying code/text is present too (given issue of dates recorded when investigation ordered not actually performed)
                If pcancerDate = 0 Or inputData(i, pcancerDateCol) < pcancerDate Then '' If result is the first valid one (otherwise would not be 0) or it's earlier than currently populated one
                    pcancerDate = inputData(i, pcancerDateCol) '' If yes then store the date (no result as not relevant for EMIS Web)
                End If
            End If
        Next pcancerResultNumCounter '' Onto next result
        If pcancerDate <> 0 Then '' Checks a date was found (presence of text associated with it verified previously but not used here given not relevant for EMIS Web)
            outputData(outputRowIndex, DataCols.PancreaticcancerdiagnosisDate) = pcancerDate '' If date (earliest) found then write to output column in array
            If IsDate(dateOfBirth) Then '' Check DoB is a valid date format
                outputData(outputRowIndex, DataCols.PancreaticcancerdiagnosisPatientsAgeatEvent) = AgeYears(dateOfBirth, pcancerDate) '' Calculates whole-year age at diagnosis
            End If
        End If

        '' Find and extract the baseline (very earliest) height measurement when aged 18+ for each person
        '' Keeps the original height result text so its comparator and comma decimal can be preserved in the output, but only accepts a measurement if its numeric part is > 0
        Dim earliestValidHeightDate As Date '' Stores the earliest valid height date
        earliestValidHeightDate = 0 '' Sets to 0 to clear fresh for the 9 loops
        Dim earliestValidHeightText As String '' Stores the earliest height recorded in full, i.e. measurement and UoM (e.g., "< 170,5 m")
        Dim heightResultNumCounter As Long '' Counter to loop through the height result and date pairs
        Dim numericHeightBaseline As Double  '' Stores the numeric part of the height result for checking it's greater than zero
        
        For heightResultNumCounter = 1 To 9 '' Nine loops as nine result pairs
            Dim heightVal As Variant '' Stores the current height result
            heightVal = inputData(i, TPPCols.Height1 + (heightResultNumCounter - 1) * 2)   '' Height1, Height2, ...
            Dim heightDate As Variant '' Stores date of current height result
            heightDate = inputData(i, TPPCols.Height1Date + (heightResultNumCounter - 1) * 2) '' Height1Date, Height2Date, ...
        
            '' Check date validity and that patient was aged 18 or older at measurement; then require a non-zero numeric value
            If IsDate(dateOfBirth) And IsDate(heightDate) Then
                numericHeightBaseline = ExtractNumeric(heightVal) '' Handles "." or "," decimals and ignores any comparator for numeric logic
                If numericHeightBaseline > 0 Then                  '' Excludes true zeros like "0 m" which are not useful
                    If AgeYears(dateOfBirth, heightDate) >= 18 Then '' Baseline height only if an adult
                        If earliestValidHeightDate = 0 Or heightDate < earliestValidHeightDate Then '' First seen or earlier than current earliest
                            earliestValidHeightDate = heightDate '' Save earliest valid date
                            earliestValidHeightText = Trim$(CStr(heightVal)) '' Store the original height result text so e.g. comparator and comma decimal can be preserved in the output
                        End If
                    End If
                End If
            End If
        Next heightResultNumCounter '' Onto next result
        
        If earliestValidHeightDate <> 0 And Len(earliestValidHeightText) > 0 Then '' Checks a date was found and height result with it
            outputData(outputRowIndex, DataCols.BaselineHeightValue) = FormatMeasuredValue(earliestValidHeightText) '' Write the formatted earliest height value (e.g. '> 180,5' but no cm) to the output array, preserving the comparator and original number text
            outputData(outputRowIndex, DataCols.BaselineHeightUnitofMeasure) = ExtractUnit(earliestValidHeightText) '' Write the UoM to the output array
        End If
        
        
        '' ENDPAC heights: extract 1a-2e height values per specified date ranges relative to index date
        '' Keeps the original height result text so e.g. comparator and comma decimal can be preserved in the output, but only considers results with a numeric value > 0
        '' Define the ENDPAC height result windows and selection criteria - differs from HbA1c process as all were positive as all were pre-index, but here is a mix
        Dim heightWindowSettings(1 To 7) As Variant '' Array of seven ENDPAC height window definitions (1a, 1b, 2a-2e)
        heightWindowSettings(1) = Array(0, 91, "earliest", DataCols.Height1aValue, DataCols.Height1aUnitofMeasure, DataCols.height1aDate)  '' Window 1 (1a): rangeStart lower signed day-difference bound (0 = Index Date), rangeEnd upper signed day-difference bound (+91 = 91 days after index), selection strategy, three cols to store best result, UoM and associated date
        heightWindowSettings(2) = Array(-91, -1, "latest", DataCols.Height1bValue, DataCols.Height1bUnitofMeasure, DataCols.height1bDate) '' Window 2 (1b): rangeStart lower signed day-difference bound (-91 = 91 days before index), rangeEnd upper signed day-difference bound (-1 = 1 day before index), selection strategy, three cols to store best result, UoM and associated date
        heightWindowSettings(3) = Array(-364, -274, "earliest", DataCols.Height2aValue, DataCols.Height2aUnitofMeasure, DataCols.height2aDate) '' 2a etc.
        heightWindowSettings(4) = Array(-455, -365, "latest", DataCols.Height2bValue, DataCols.Height2bUnitofMeasure, DataCols.height2bDate) '' 2b
        heightWindowSettings(5) = Array(-546, -456, "latest", DataCols.Height2cValue, DataCols.Height2cUnitofMeasure, DataCols.height2cDate) '' 2c
        heightWindowSettings(6) = Array(-273, -92, "earliest", DataCols.Height2dValue, DataCols.Height2dUnitofMeasure, DataCols.height2dDate) '' 2d
        heightWindowSettings(7) = Array(-3652, -547, "latest", DataCols.Height2eValue, DataCols.Height2eUnitofMeasure, DataCols.height2eDate) '' 2e
                
        Dim heightWindowCounter As Long '' Counter to loop through the windows
        For heightWindowCounter = 1 To 7 '' Seven windows so seven loops
            
            Dim bestHeightDate As Date '' Stores the best height date found for this window
            bestHeightDate = 0 '' Reset the best date for this window
            Dim bestHeightRaw As String '' Stores the full original height result text selected for this window
            bestHeightRaw = vbNullString '' Clear it for this window
            Dim numericHeight As Double '' Stores the numeric part of the height result for checking it's greater than zero
            Dim heightDayDiff As Long '' Days from index date to candidate measurement
            
            '' Extract window settings and output mappings from settings array
            rangeStart = heightWindowSettings(heightWindowCounter)(0)    '' Start day difference from index date
            rangeEnd = heightWindowSettings(heightWindowCounter)(1)      '' End day difference from index date
            dateStrategy = heightWindowSettings(heightWindowCounter)(2)  '' Strategy of earliest or latest
            outputValCol = heightWindowSettings(heightWindowCounter)(3)  '' Output column for measurement
            outputUnitCol = heightWindowSettings(heightWindowCounter)(4) '' Output column for UoM
            outputDateCol = heightWindowSettings(heightWindowCounter)(5) '' Output column for date of measurement
               
            '' Loop through the 9 possible height readings
            For heightResultNumCounter = 1 To 9
                Dim heightCandidate As Variant '' Stores the current candidate height result
                heightCandidate = inputData(i, TPPCols.Height1 + (heightResultNumCounter - 1) * 2) '' Get the current height result from the appropriate TPP input col
                Dim heightCandidateDate As Variant '' Stores the date of the current candidate height result
                heightCandidateDate = inputData(i, TPPCols.Height1Date + (heightResultNumCounter - 1) * 2) '' Get the current height result date from the appropriate TPP input col
        
                '' Check date validity and that numeric part is > 0 (exclude true zeros)
                If IsDate(heightCandidateDate) Then
                    numericHeight = ExtractNumeric(heightCandidate) '' Supports "," or "." decimals; comparator ignored for logic
                    If numericHeight > 0 Then
                        heightDayDiff = DateDiff("d", indexDate, heightCandidateDate) '' +ve after index, -ve before index
                        '' If height measurement is within the window, pick earliest or latest by strategy
                        If heightDayDiff >= rangeStart And heightDayDiff <= rangeEnd Then
                            If bestHeightDate = 0 _
                               Or (dateStrategy = "earliest" And heightCandidateDate < bestHeightDate) _
                               Or (dateStrategy = "latest" And heightCandidateDate > bestHeightDate) Then
                                bestHeightDate = heightCandidateDate
                                bestHeightRaw = Trim$(CStr(heightCandidate))  '' Store the original height result text so e.g. comparator and comma decimal can be preserved in the output
                            End If
                        End If
                    End If
                End If
            Next heightResultNumCounter '' Onto next result
        
            '' If a valid result was found then extract the result and UoM and put into output array
            If bestHeightDate <> 0 And Len(bestHeightRaw) > 0 Then '' Checks a date was found and result with it
                outputData(outputRowIndex, outputValCol) = FormatMeasuredValue(bestHeightRaw)  '' Write the formatted height value (e.g. '> 180,5' but no cm) to the output array
                outputData(outputRowIndex, outputUnitCol) = ExtractUnit(bestHeightRaw)         '' Write the extracted height UoM to the output array
                outputData(outputRowIndex, outputDateCol) = bestHeightDate                     '' Write the selected height date to the output array
            End If
        Next heightWindowCounter '' Onto next defined window
        
        
        '' ENDPAC BMIs: extract 1a-2e BMI values per specified date ranges relative to index date (no UoM output needed)
        '' Keeps the original BMI result text so its comparator and comma decimal can be preserved in the output, but only considers results with a numeric value > 0
        '' Define the ENDPAC BMI result windows and selection criteria - differs from HbA1c process as all were positive as all were pre-index, but here is a mix
        Dim bmiWindowSettings(1 To 7) As Variant '' Array of seven ENDPAC BMI window definitions (1a, 1b, 2a-2e)
        bmiWindowSettings(1) = Array(0, 91, "earliest", DataCols.BMI1aValue, DataCols.bMI1aDate)  '' Window 1 (1a): rangeStart lower signed day-difference bound (0 = Index Date), rangeEnd upper signed day-difference bound (+91 = 91 days after index), selection strategy, three cols to store best result, UoM and associated date
        bmiWindowSettings(2) = Array(-91, -1, "latest", DataCols.BMI1bValue, DataCols.bMI1bDate) '' Window 2 (1b): rangeStart lower signed day-difference bound (-91 = 91 days before index), rangeEnd upper signed day-difference bound (-1 = 1 day before index), selection strategy, three cols to store best result, UoM and associated date
        bmiWindowSettings(3) = Array(-364, -274, "earliest", DataCols.BMI2aValue, DataCols.bMI2aDate) '' 2a etc.
        bmiWindowSettings(4) = Array(-455, -365, "latest", DataCols.BMI2bValue, DataCols.bMI2bDate) '' 2b
        bmiWindowSettings(5) = Array(-546, -456, "latest", DataCols.BMI2cValue, DataCols.bMI2cDate) '' 2c
        bmiWindowSettings(6) = Array(-273, -92, "earliest", DataCols.BMI2dValue, DataCols.bMI2dDate) '' 2d
        bmiWindowSettings(7) = Array(-3652, -547, "latest", DataCols.BMI2eValue, DataCols.bMI2eDate) '' 2e
        
        Dim bmiWindowCounter As Long '' Counter to loop through the windows
        For bmiWindowCounter = 1 To 7 '' Seven windows so seven loops
        
            Dim bestBmiDate As Date '' Stores the best BMI date found for this window
            bestBmiDate = 0 '' Reset the date for window
            Dim bestBmiRaw As String '' Stores the original BMI result text selected for this window
            bestBmiRaw = vbNullString  '' Clear result for this window
            Dim numericBMI As Double '' Stores the numeric part of the BMI result for checking it's greater than zero
            Dim bmiDayDiff As Long  '' Days from index date to measurement
            
            '' Obtain window parameters
            rangeStart = bmiWindowSettings(bmiWindowCounter)(0)    '' Start day difference from index date
            rangeEnd = bmiWindowSettings(bmiWindowCounter)(1)      '' End day difference from index date
            dateStrategy = bmiWindowSettings(bmiWindowCounter)(2)  '' Strategy of earliest or latest
            outputValCol = bmiWindowSettings(bmiWindowCounter)(3)  '' Output column for measurement (no UoM for BMI)
            outputDateCol = bmiWindowSettings(bmiWindowCounter)(4) '' Output column for date of measurement
        
            '' Loop through the 9 possible BMI readings
            Dim bmiResultNumCounter As Long
            For bmiResultNumCounter = 1 To 9
                Dim bmiCandidate As Variant '' Stores the current candidate BMI result
                bmiCandidate = inputData(i, TPPCols.BMI1 + (bmiResultNumCounter - 1) * 2) '' Get the current BMI result from the appropriate TPP input col
                Dim bmiCandidateDate As Variant '' Stores the date of the current candidate BMI result
                bmiCandidateDate = inputData(i, TPPCols.BMI1Date + (bmiResultNumCounter - 1) * 2) '' Get the current BMI result date from the appropriate TPP input column
        
                '' Check the dates are valid and numeric part > 0 (exclude true zeros)
                If IsDate(bmiCandidateDate) Then
                    numericBMI = ExtractNumeric(bmiCandidate) '' Comma/period decimals accepted; comparator ignored for logic
                    If numericBMI > 0 Then
                        bmiDayDiff = DateDiff("d", indexDate, bmiCandidateDate) '' Days from index date to measurement
                        '' If BMI measurement is within the time window, select earliest or latest depending on strategy
                        If bmiDayDiff >= rangeStart And bmiDayDiff <= rangeEnd Then
                            If bestBmiDate = 0 _
                               Or (dateStrategy = "earliest" And bmiCandidateDate < bestBmiDate) _
                               Or (dateStrategy = "latest" And bmiCandidateDate > bestBmiDate) Then
                                bestBmiDate = bmiCandidateDate
                                bestBmiRaw = Trim$(CStr(bmiCandidate))  '' Store the original BMI result text so its e.g. comparator and comma decimal can be preserved in the output
                            End If
                        End If
                    End If
                End If
            Next bmiResultNumCounter '' Onto next result
        
            '' If a valid result was found then put result and date into output array (no UoM as BMI)
            If bestBmiDate <> 0 And Len(bestBmiRaw) > 0 Then '' Checks a date was found and result with it
                outputData(outputRowIndex, outputValCol) = FormatMeasuredValue(bestBmiRaw)  '' Write the formatted BMI result (e.g. '> 32') to the output array, preserving the comparator and original number text
                outputData(outputRowIndex, outputDateCol) = bestBmiDate  '' Write the selected BMI date to the output array
            End If
        Next bmiWindowCounter '' Onto next defined window
        
        
        '' ENDPAC weights: extract 1a-2e weight values per specified date ranges relative to index date
        '' Keeps the original weight result text so its e.g. comparator and comma decimal can be preserved in the output, but only considers results with a numeric value > 0
        '' Define the ENDPAC weight result windows and selection criteria - differs from HbA1c process as all were positive as all were pre-index, but here is a mix
        Dim weightWindowSettings(1 To 7) As Variant '' Array of seven ENDPAC weight window definitions (1a, 1b, 2a-2e)
        weightWindowSettings(1) = Array(0, 91, "earliest", DataCols.Weight1aValue, DataCols.Weight1aUnitofMeasure, DataCols.weight1aDate) '' Window 1 (1a): rangeStart lower signed day-difference bound (0 = Index Date), rangeEnd upper signed day-difference bound (+91 = 91 days after index), selection strategy, three cols to store best result, UoM and associated date
        weightWindowSettings(2) = Array(-91, -1, "latest", DataCols.Weight1bValue, DataCols.Weight1bUnitofMeasure, DataCols.weight1bDate) '' Window 2 (1b): rangeStart lower signed day-difference bound (-91 = 91 days before index), rangeEnd upper signed day-difference bound (-1 = 1 day before index), selection strategy, three cols to store best result, UoM and associated date
        weightWindowSettings(3) = Array(-364, -274, "earliest", DataCols.Weight2aValue, DataCols.Weight2aUnitofMeasure, DataCols.weight2aDate) '' 2a etc.
        weightWindowSettings(4) = Array(-455, -365, "latest", DataCols.Weight2bValue, DataCols.Weight2bUnitofMeasure, DataCols.weight2bDate) '' 2b
        weightWindowSettings(5) = Array(-546, -456, "latest", DataCols.Weight2cValue, DataCols.Weight2cUnitofMeasure, DataCols.weight2cDate) '' 2c
        weightWindowSettings(6) = Array(-273, -92, "earliest", DataCols.Weight2dValue, DataCols.Weight2dUnitofMeasure, DataCols.weight2dDate) '' 2d
        weightWindowSettings(7) = Array(-3652, -547, "latest", DataCols.Weight2eValue, DataCols.Weight2eUnitofMeasure, DataCols.weight2eDate) '' 2e

        Dim weightWindowCounter As Long '' Counter to loop through the windows
        For weightWindowCounter = 1 To 7 '' Seven windows so seven loops
        
            Dim bestWeightDate As Date '' Stores the best weight date found for this window
            bestWeightDate = 0         '' Reset the best date for window
            Dim bestWeightRaw As String '' Stores the original weight result text selected for this window
            bestWeightRaw = vbNullString '' Clear ready for this window
            Dim numericWeight As Double '' Stores the numeric part of the weight result for checking it's greater than zero
            Dim weightDayDiff As Long  '' Days from index date to measurement
            
            '' Obtain window parameters
            rangeStart = weightWindowSettings(weightWindowCounter)(0)    '' Start day difference from index date
            rangeEnd = weightWindowSettings(weightWindowCounter)(1)      '' End day difference from index date
            dateStrategy = weightWindowSettings(weightWindowCounter)(2)  '' Strategy of earliest or latest
            outputValCol = weightWindowSettings(weightWindowCounter)(3)  '' Output column for measurement
            outputUnitCol = weightWindowSettings(weightWindowCounter)(4) '' Output column for UoM
            outputDateCol = weightWindowSettings(weightWindowCounter)(5) '' Output column for date of measurement
        
            '' Loop through the 9 possible weight readings
            Dim weightResultNumCounter As Long
            For weightResultNumCounter = 1 To 9
                Dim weightCandidate As Variant '' Stores the current candidate weight result
                weightCandidate = inputData(i, TPPCols.Weight1 + (weightResultNumCounter - 1) * 2) '' Get the current weight result from the appropriate TPP input col
                Dim weightDate As Variant '' Stores the date of the current candidate weight result
                weightDate = inputData(i, TPPCols.Weight1Date + (weightResultNumCounter - 1) * 2) '' Get the current weight result date from the appropriate TPP input col
        
                '' Check the dates are valid and numeric part > 0 (exclude true zeros)
                If IsDate(weightDate) Then
                    numericWeight = ExtractNumeric(weightCandidate) '' Comma/period decimals ok; comparator ignored for logic
                    If numericWeight > 0 Then
                        weightDayDiff = DateDiff("d", indexDate, weightDate) '' Days from index date to measurement
                        '' If weight measurement is within the time window, select earliest or latest depending on strategy
                        If weightDayDiff >= rangeStart And weightDayDiff <= rangeEnd Then
                            If bestWeightDate = 0 _
                               Or (dateStrategy = "earliest" And weightDate < bestWeightDate) _
                               Or (dateStrategy = "latest" And weightDate > bestWeightDate) Then
                                bestWeightDate = weightDate
                                bestWeightRaw = Trim$(CStr(weightCandidate))  '' Store the original weight result text so its e.g. comparator and comma decimal can be preserved in the output
                            End If
                        End If
                    End If
                End If
            Next weightResultNumCounter '' Onto next result
        
            '' If a valid result was found then extract the result and UoM and put into output array
            If bestWeightDate <> 0 And Len(bestWeightRaw) > 0 Then '' Checks a date was found and result with it
                outputData(outputRowIndex, outputValCol) = FormatMeasuredValue(bestWeightRaw)  '' Write the formatted weight value (e.g. '> 82,6' but no kg) to the output array
                outputData(outputRowIndex, outputUnitCol) = ExtractUnit(bestWeightRaw)         '' Write the extracted weight UoM to the output array
                outputData(outputRowIndex, outputDateCol) = bestWeightDate                     '' Write the selected weight date to the output array
            End If
        Next weightWindowCounter '' Onto next defined window


FinaliseRow: '' Either a row has been fully processed to get here or has jumped here after no valid Index date found
        outputRowIndex = outputRowIndex + 1 '' Moves output pointer to next row in output array for next record
    Next i '' Complete for this iteration and move on to next row for this loop (basically all the processing)

    '' Output the TPP to EMIS transformed data into Sheet1
    Application.StatusBar = "Writing transformed data to Sheet1" '' Comment in Excel status bar
    dWS.Cells.ClearContents '' Clear contents in output sheet
    Call ClearFormatting(dWS) '' From dHelperMethods
    dWS.Range(dWS.Cells(1, 1), dWS.Cells(1, UBound(gHeaderTexts))).Value = gHeaderTexts '' Writes the EMIS headers from gHeaderTexts array into the output worksheet Sheet1
    dWS.Range(dWS.Cells(2, 1), dWS.Cells(outputRowIndex, UBound(gHeaderTexts))).Value = outputData '' Write outputData array to sheet, from A2 (2, 1) down to defined row and col
    
    '' Update screen so people can see that the data have changed before acknowledging the pop-up messages
    Application.ScreenUpdating = True
    
    '' Check for correct baseline height header in cell C1 and HbA1c index date in F2 indicating successful TPP to EMIS transformation and display success message
    If ThisWorkbook.Sheets("Sheet1").Range("C1").Value = "Baseline Height: Value" And _
       IsDate(ThisWorkbook.Sheets("Sheet1").Range("F2").Value) Then
        Application.StatusBar = "TPP processing complete" '' Comment in Excel status bar
        MsgBox "DEFEND PRIME TPP to EMIS processing complete. Please now proceed with the processing and anonymisation macro - Button 2.", vbInformation
    Else
        Application.StatusBar = "TPP processing failed validation" '' If cell checks fail, comment in Excel status bar and add next line's message to pop-up message box
        MsgBox "It looks like something may have gone wrong. Please check Sheet1 visually to ensure the TPP to EMIS processing ran correctly. If it did not, please delete all data from Sheet1 and repaste the original TPP extract (including headers) into cell A1 of worksheet named 'Sheet1', and rerun this TPP macro. If the issue persists, please contact the DEFEND PRIME team noting 'Error 5a'.", vbExclamation
    End If
    
ExitHere: '' Clean exit after normal completion or encountered known processing failure once transformation has started
    ''' Turn back on the Excel settings turned off at the start of processing
    Call ReadyApplication(processingMode:=False) '' From dHelperMethods, restores normal Excel settings
    Application.StatusBar = False '' Clear status bar to prevent lingering messages from causing confusion
    Call ClearGlobalData '' Clear ready for next macro
    Application.Cursor = xlDefault '' Return cursor from hour glass to standard
    
    '' Re-enable button actions
    SafeSetOnAction ActiveSheet, "btnRunTPP", "'" & ThisWorkbook.Name & "'!A_TransformTPPtoEMIS" ''  Make the macro run button clickable again in case something went wrong
    SafeSetOnAction ActiveSheet, "btnRunDP", "'" & ThisWorkbook.Name & "'!AA_RunThisForDEFENDPRIME"   '' Make the second macro run button clickable ready for subsequent processing
    
    Exit Sub

ErrHandler: '' Runs only if there's a VBA runtime error
    Dim errNumber As Long '' Stores the VBA error number before any clean-up code can alter the Err object
    Dim errDescription As String '' Stores the VBA error description before any clean-up code can alter the Err object
    Dim errLine As Long '' Stores the numbered VBA code line where the error occurred, if line numbers are present
    errNumber = Err.Number '' Save the original VBA error number
    errDescription = Err.Description '' Save the original VBA error description
    errLine = Erl '' Save the line number where the error occurred; returns 0 if the relevant code has no line number
    Application.ScreenUpdating = True '' Reenable screen updating
    Application.StatusBar = False '' Clear status bar to prevent lingering messages from causing confusion
    Application.Cursor = xlDefault '' Return cursor from hour glass to standard

    '' Restore buttons even with error
    SafeSetOnAction ActiveSheet, "btnRunTPP", "'" & ThisWorkbook.Name & "'!A_TransformTPPtoEMIS" ''  Make the macro run button clickable again in case something went wrong
    SafeSetOnAction ActiveSheet, "btnRunDP", "'" & ThisWorkbook.Name & "'!AA_RunThisForDEFENDPRIME"   '' Make the second macro run button clickable ready for subsequent processing

    '' clear, specific message; Erl shows a line number if you temporarily number lines
    MsgBox "TPP Transformation stopped with error " & errNumber & _
           IIf(errLine > 0, " (line " & errLine & ")", "") & ":" & vbCrLf & _
           errDescription, vbCritical, "TPP to EMIS transformation"

    Resume ExitHere '' Even on error go back to normal tidy-up

End Sub

'' Extracts the first numeric value at the start of the cell (after an optional comparator)
'' Accepts either "." or "," as the decimal separator and returns the result as a VBA Double
'' Regex is Reg.ular Ex.pression
Function ExtractNumeric(cellValue As Variant) As Double
    Dim s As String '' Stores the supplied measurement converted to text
    Dim rx As Object '' RegExp object used to identify the numerical part of the measurement
    Dim m As Object   '' Stores the section of measurement text identified by the RegExp
    Dim numText As String '' Stores the extracted number as text before conversion to a VBA Double
    
    s = Trim$(CStr(cellValue)) '' Convert the input to text and remove leading/trailing spaces

    Set rx = CreateObject("VBScript.RegExp") '' Create the RegExp object used to search the measurement text
    
    rx.Pattern = "^\s*(?:[<>]=?)?\s*(\d+(?:[.,]\d+)?)" '' Find an optional comparator followed by a number at the start of the measurement and capture only the number
    rx.Global = False  '' Stop searching once the measurement value at the start of the cell text has been found

    If rx.Test(s) Then '' Check that the measurement contains a numerical value in the expected format
        Set m = rx.Execute(s)(0) '' Store the section of text found at the start of the measurement
        numText = m.SubMatches(0) '' Store only the captured number text, excluding comparator, spaces and UoM, e.g. '< 170,5 cm' gives '170,5'
        
        '' Convert a comma decimal separator to a period so the extracted text can be converted consistently to Double
        numText = Replace(numText, ",", ".")
        ExtractNumeric = CDbl(numText) '' Convert the extracted number text to a numerical Double and return it
    Else
        ExtractNumeric = 0 '' Return zero if no numerical measurement can be identified
    End If
End Function

'' Extracts the unit of measurement by excluding any leading comparator ("<", ">", "<=", ">="), number, and spaces
Function ExtractUnit(cellValue As Variant) As String
    Dim cellText As String '' Stores the supplied measurement converted to text
    Dim regexMatches As Object '' Stores the section of measurement text identified by the RegExp
    
    cellText = Trim(CStr(cellValue)) '' Convert the input to text and remove leading/trailing spaces
    
    With CreateObject("VBScript.RegExp") '' Create the RegExp object used to search the measurement text
        .Pattern = "^\s*(?:[<>]=?)?\s*\d+(?:[.,]\d+)?\s*" '' Find the optional comparator, number and surrounding spaces at the start of the measurement so the remaining text can be returned as the UoM
        .Global = False '' Stop searching once the measurement value at the start of the cell text has been found
        If .Test(cellText) Then '' Check that the measurement contains a numerical value in the expected format
            Set regexMatches = .Execute(cellText) '' Store the section of text found at the start of the measurement
            ExtractUnit = Trim(Mid(cellText, Len(regexMatches(0).Value) + 1)) '' Return the text remaining after the comparator, number and surrounding spaces removed, i.e. the UoM
        Else
            ExtractUnit = "" '' Return blank if no numerical measurement can be identified
        End If
    End With
End Function

'' Calculates whole years of age between Date of Birth and an event date - fix for the fringe case flawed arithmetic method of subtracting e.g. DoB from Index/365.2425
'' Inclusive of the birthday (i.e., on the exact birthday, returns the new age)
'' Handles leap years and leap-day births correctly by using DateAdd/DateDiff calendar arithmetic
'' Returns -1 if either input cannot be converted to a valid VBA Date
Private Function AgeYears(ByVal dateOfBirthInput As Variant, ByVal eventDateInput As Variant) As Long
    On Error GoTo Fail '' If anything goes wrong in conversion/arithmetic, jump to Fail to return -1

    Dim patientDateOfBirth As Date '' Stores the Date of Birth as a VBA Date
    patientDateOfBirth = CDate(dateOfBirthInput) '' Coerce input DoB (which may be Date/Double/String/Variant) to a VBA Date (Excel serial)
    Dim eventDate As Date '' Stores the event date as a VBA Date
    eventDate = CDate(eventDateInput) '' Coerce input event date to a VBA Date (Excel serial)

    '' DateDiff with "yyyy" gives the count of calendar year boundaries crossed between DoB and eventDate
    '' e.g. 01/01/1950 -> 31/12/1967 gives 17; 01/01/1950 -> 01/01/1968 gives 18
    Dim yrs As Long '' Stores the calculated whole-year difference
    yrs = DateDiff("yyyy", patientDateOfBirth, eventDate) '' Calculate the number of calendar year boundaries crossed

    '' If the event occurs before the birthday in the event year, subtract one because the birthday hasn't been reached yet
    '' DateAdd("yyyy", yrs, DoB) = DoB advanced by 'yrs' whole calendar years; if eventDate is earlier than that anniversary, reduce yrs by 1
    If eventDate < DateAdd("yyyy", yrs, patientDateOfBirth) Then
        yrs = yrs - 1
    End If

    AgeYears = yrs '' Normal return path with whole years of age
    Exit Function

Fail:
    AgeYears = -1 '' Signal an invalid input (e.g., CDate failed) to callers that check for this
End Function

'' True if text begins with an optional comparator then a number (unit may follow)
'' Accepts: "0", "0.01", "<0.01", "< 0.01", ">= 1.2", etc. (even when unit is stuck on)
Private Function HasMeasurement(cellValue As Variant) As Boolean '' Declares a function returning Boolean
    Dim cellText As String '' Stores the supplied measurement converted to text
    cellText = Trim(CStr(cellValue)) '' Converts input to String and removes leading/trailing spaces
    With CreateObject("VBScript.RegExp") '' Creates RegExp object to test the leading portion
        .Pattern = "^\s*(?:[<>]=?)?\s*\d+(?:[.,]\d+)?(?=\D|$)" '' Check for an optional comparator followed by an integer or decimal number at the start of the measurement
        .Global = False '' Stop searching once the measurement value at the start of the cell text has been found
        HasMeasurement = .Test(cellText) '' Return True if the expected measurement format is found at the start of the cell text; otherwise False
    End With
End Function

'' Returns a leading comparator if present like "<", ">", "<=", ">=" else empty string
Private Function ExtractComparator(ByVal cellValue As Variant) As String '' Declares a function returning String
    Dim cellText As String '' Stores the supplied measurement converted to text
    cellText = Trim(CStr(cellValue)) '' Convert the input to String and remove leading/trailing spaces
    Dim regexMatches As Object  '' Stores the section of measurement text identified by the Regex
    With CreateObject("VBScript.RegExp") '' Creates a RegExp object to run a regex on cellText
        .Pattern = "^\s*([<>]=?)?\s*\d+(?:[.,]\d+)?(?=\D|$)" '' Find an optional comparator followed by a number at the start of the measurement and capture only the comparator
        .Global = False '' Stop searching once the measurement value at the start of the cell text has been found
        If .Test(cellText) Then ''  '' Check that the measurement contains a numerical value in the expected format
            Set regexMatches = .Execute(cellText) '' Store the section of text found at the start of the measurement
            If regexMatches(0).SubMatches(0) <> "" Then '' Check whether the optional comparator was captured
                ExtractComparator = regexMatches(0).SubMatches(0) '' Return only the captured comparator
            Else
                ExtractComparator = "" '' No comparator captured; return blank
            End If
        Else
            ExtractComparator = "" '' No prefix match; return blank
        End If
    End With
End Function

'' Formats measurement value for output by keeping original number text and any comparator, but excludes the UoM
'' Inserts a space after comparator if present (e.g. "<0.01 mg/dL" becomes "< 0.01").
Private Function FormatMeasuredValue(ByVal cellValue As Variant) As String
    Dim s As String  '' Stores the supplied measurement converted to text
    Dim comp As String '' Stores any comparator at the start of the measurement, e.g. "<", ">", "<=" or ">="
    Dim rx As Object '' RegExp object used to identify the numerical part of the measurement
    Dim m As Object '' '' Stores the section of measurement text identified by the RegExp
    
    s = Trim$(CStr(cellValue)) '' Convert the input to text and remove leading/trailing spaces
    comp = ExtractComparator(s) '' Extract the comparator from the supplied measurement, or return blank if none present
    
    Set rx = CreateObject("VBScript.RegExp") '' Create the RegExp object used to search the measurement text
    
    '' Find an optional comparator followed by a number at the start of the measurement and capture only the number, retaining "." or "," decimals as originally written
    rx.Pattern = "^\s*(?:[<>]=?)?\s*(\d+(?:[.,]\d+)?)"
    rx.Global = False '' Stop searching once the measurement value at the start of the cell text has been found

    If rx.Test(s) Then '' Check that the measurement contains a numerical value in the expected format
        Set m = rx.Execute(s)(0) '' Store the section of text found at the start of the measurement
        
        If comp <> "" Then '' If the original measurement contained a comparator
            FormatMeasuredValue = comp & " " & m.SubMatches(0) '' Return the comparator and original number text separated by a single space
        Else
            FormatMeasuredValue = m.SubMatches(0) '' Return the original number text when no comparator is present
        End If
    Else
        FormatMeasuredValue = "" '' Return blank if no numerical measurement can be identified
    End If
End Function

'' Modified version of CheckData used in dMain for TPP sheet
'' Cleans and prepares the data on Sheet1 for processing

Private Sub CheckTPPData()
    Dim sourceSheet As Worksheet '' Worksheet containing the raw TPP input data
    Set sourceSheet = ThisWorkbook.Sheets("Sheet1") '' Sets sourceSheet to Sheet1 in this workbook
    Dim i As Long '' Variable i to loop through rows
    Dim c As Long '' Variable c to loop through columns
    Dim foundHeaderRow As Long '' Stores the row number where the TPP header is found
    foundHeaderRow = 0 '' Set to 0 to start i.e. not found yet
    Dim lastHeaderCol As Long '' Stores the col number of the final populated header column after blank column clean-up
    Dim validHeader As Boolean '' Tracks whether the prepared TPP header structure matches expectations
    
    '' Search for the header row by checking the first column
    For i = 1 To 30 '' Loop through rows 1-30 (seems a reasonable number if TPP adds lots of blurb in future for some reason)
        If Trim(sourceSheet.Cells(i, 1).value2) = "Organisation ID" Then '' If Organisation ID found (trimming any leading/trailing spaces)
            foundHeaderRow = i '' Save that row number to foundHeaderRow
            Exit For '' Exit the loop given it's been found
        End If
    Next i '' Move on to next row if not found yet

    If foundHeaderRow = 0 Then '' If still not found after 30 rows then pop-up below message to user
        MsgBox "Sorry, the expected TPP extract data do not appear to have been pasted into cell A1 of Sheet1. Please delete any data you have pasted in and repaste the original extract data into cell A1 of Sheet1. If this problem persists, please contact the DEFEND PRIME team noting 'Error 4b'", vbExclamation
        tppFirstRowOfData = 0 '' Leave failure flag for the caller, which checks for 0 immediately after CheckTPPData returns
        Exit Sub
    End If

    '' Delete all rows before the header row
    If foundHeaderRow > 1 Then '' If it's row 1, don't delete any rows, but if it's bigger than 1 i.e. if Organisation ID found below row 1 (2, 3 etc.)
        sourceSheet.Rows("1:" & (foundHeaderRow - 1)).Delete '' Delete the rows above Organisation ID (header row) - i.e. from row 1 up to the row before (-1) the header row, so if it's row 5, delete 1 to 5minus1=4)
    End If

    '' Set the global start row (used in data processing)
    tppFirstRowOfData = 2 '' Since headers (indexed by Organisation ID) are now on row 1, data are row 2 onwards

    '' Remove any merged or blank columns in header row (which is now row 1)
    For c = 89 To 1 Step -1 '' Loop backwards through cols 89 to 1 (otherwise loop would break upon deletion as col indices move left) in the header row deleting cols where cell is blank - issue caused by merged headers cells
        If sourceSheet.Cells(1, c).value2 = vbNullString Then '' Check if cell in row 1 of col i is blank
            sourceSheet.Cells(1, c).EntireColumn.Delete '' If it is then delete that column
        End If
    Next c '' Move to next column
    
    '' Validate the prepared TPP worksheet structure after removing potential blank column extraction artefacts
    validHeader = True '' Assume structure is valid unless the check below fails
    lastHeaderCol = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).column '' Find the final populated header column after blank col clean-up

    '' The expected TPP extract contains 89 populated columns after known blank structural columns have been removed
    '' An unexpected populated col would mean more than 89 columns, whilst a missing one would leave fewer than 89
    If lastHeaderCol <> 89 Then validHeader = False

    '' Stop before formatting or processing if the expected TPP structure hasn't been confirmed
    If validHeader = False Then
        MsgBox "Sorry, the pasted data do not appear to have the expected DEFEND PRIME TPP SystmOne structure, so processing has been stopped. Please delete all data from Sheet1 and repaste the original extracted data, including all headers, into cell A1. If the problem persists, please contact the DEFEND PRIME team noting 'Error 4c'.", vbExclamation
        tppFirstRowOfData = 0 '' Leave failure flag for the caller, which checks for 0 immediately after CheckTPPData returns
        Exit Sub
    End If
    
End Sub

Private Sub FormatTPPInputSheet()
    Dim sourceSheet As Worksheet '' Worksheet containing the raw TPP input data
    Set sourceSheet = ThisWorkbook.Sheets("Sheet1") '' Set sourceSheet to Sheet1 in this workbook

    ' Clear autofilter if active
    If sourceSheet.AutoFilterMode Then sourceSheet.AutoFilterMode = False '' Turns off filters which can cause rows being hidden from macro

    ' Apply consistent visual date formatting to all known date columns
    With sourceSheet '' With block avoids needing to repeat the sourceSheet object name on all following lines
        ' Date of Birth
        .Cells(1, TPPCols.DoB).EntireColumn.NumberFormat = "dd/mm/yyyy;@" '' if numeric format them into the specified date format ; if text, @, leave as is

        ' HbA1c dates
        .Cells(1, TPPCols.HbA1c1Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.HbA1c2Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.HbA1c3Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.HbA1c4Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.HbA1c5Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.HbA1c6Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.HbA1c7Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.HbA1c8Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.HbA1c9Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"

        ' Weight dates
        .Cells(1, TPPCols.Weight1Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Weight2Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Weight3Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Weight4Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Weight5Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Weight6Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Weight7Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Weight8Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Weight9Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"

        ' BMI dates
        .Cells(1, TPPCols.BMI1Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.BMI2Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.BMI3Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.BMI4Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.BMI5Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.BMI6Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.BMI7Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.BMI8Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.BMI9Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"

        ' Height dates
        .Cells(1, TPPCols.Height1Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Height2Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Height3Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Height4Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Height5Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Height6Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Height7Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Height8Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.Height9Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"

        ' Tumour marker dates
        .Cells(1, TPPCols.CA199_1_Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.CA199_2_Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.CEA_1_Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.CEA_2_Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"

        ' Pancreatic cancer diagnosis dates
        .Cells(1, TPPCols.PCancer_1_Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.PCancer_2_Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
        .Cells(1, TPPCols.PCancer_3_Date).EntireColumn.NumberFormat = "dd/mm/yyyy;@"
    End With
End Sub

'' Blanks a button’s OnAction so clicking does nothing while the macro is running
Private Sub SafeDisableButton(ByVal ws As Worksheet, ByVal btnName As String) '' Declares a Sub with worksheet and button name
    On Error Resume Next '' Suppresses errors if button missing or wrong type
    ws.Buttons(btnName).OnAction = "" '' Clears the macro action so clicks do nothing
    On Error GoTo 0 '' Restores normal error handling
End Sub

'' Safely assigns a workbook-qualified macro to a worksheet button’s OnAction
Private Sub SafeSetOnAction(ByVal ws As Worksheet, ByVal btnName As String, ByVal actionText As String) '' Declares a Sub with worksheet, button name, and macro text
    On Error Resume Next '' Suppresses errors if button missing
    ws.Buttons(btnName).OnAction = actionText '' Assigns the macro action (e.g. "'MyBook.xlsm'!MyMacro")
    On Error GoTo 0 '' Restores normal error handling
End Sub

