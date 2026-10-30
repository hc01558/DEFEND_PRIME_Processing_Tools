Attribute VB_Name = "dGlobals_v3"
' Revision history for dGlobals
' Differs from v2 through SetWorksheet change from 1st worksheet ordered in file to actually named Sheet1
' Version number: v3
' Release Date: 17th September 2024

Option Explicit 'Forces the declaration of all variables in the module before they are used to minimise errors as undeclared variables will trigger a compile error
Option Private Module 'Means the module can only be used within the specified project
Option Base 1  '' Sets the default lower bound of arrays to 1 (Default is 0) so array positions align with worksheet column numbers

Public dWS As Worksheet '' Public to allow other modules in this project to use the working worksheet reference
Public gFinalRow As Long   '' Public to store the final row of data for use across modules
Public gFirstRowOfData As Long  '' Public to store the first row of data for use across modules
Public gDataSet As Variant '' Public to hold the dataset in memory for use across modules

'' Definitions for the 82 DEFEND PRIME output columns
Public gHeaderTexts(82) As String  '' Stores the header text for output cols 1 to 82; populated by SetHeaderValues()

'' Descriptive names for output cols 1 to 82 so code can use purpose-indicating names instead of hard-coded column numbers, easier to work with
Public Enum DataCols  'Specify 1 for first column rather than default 0, the remainder automatically increase from there
    organisationCode = 1
    AnonymisedIdentifier
    BaselineHeightValue
    BaselineHeightUnitofMeasure
    IndexDateValue
    IndexDateDate
    IndexDatePatientsAgeatEvent
    Gender
    HbA1c2aValue
    hbA1c2aDate
    HbA1c2bValue
    hbA1c2bDate
    HbA1c2cValue
    hbA1c2cDate
    HbA1c2dValue
    hbA1c2dDate
    HbA1c2eValue
    hbA1c2eDate
    Weight1aValue
    Weight1aUnitofMeasure
    weight1aDate
    Weight1bValue
    Weight1bUnitofMeasure
    weight1bDate
    Weight2aValue
    Weight2aUnitofMeasure
    weight2aDate
    Weight2bValue
    Weight2bUnitofMeasure
    weight2bDate
    Weight2cValue
    Weight2cUnitofMeasure
    weight2cDate
    Weight2dValue
    Weight2dUnitofMeasure
    weight2dDate
    Weight2eValue
    Weight2eUnitofMeasure
    weight2eDate
    BMI1aValue
    bMI1aDate
    BMI1bValue
    bMI1bDate
    BMI2aValue
    bMI2aDate
    BMI2bValue
    bMI2bDate
    BMI2cValue
    bMI2cDate
    BMI2dValue
    bMI2dDate
    BMI2eValue
    bMI2eDate
    Height1aValue
    Height1aUnitofMeasure
    height1aDate
    Height1bValue
    Height1bUnitofMeasure
    height1bDate
    Height2aValue
    Height2aUnitofMeasure
    height2aDate
    Height2bValue
    Height2bUnitofMeasure
    height2bDate
    Height2cValue
    Height2cUnitofMeasure
    height2cDate
    Height2dValue
    Height2dUnitofMeasure
    height2dDate
    Height2eValue
    Height2eUnitofMeasure
    height2eDate
    PancreaticcancerdiagnosisDate
    PancreaticcancerdiagnosisPatientsAgeatEvent
    Cancerantigen199CA199Value
    Cancerantigen199CA199UnitofMeasure
    Cancerantigen199CA199Date
    CarcinoembryonicantigenCEAValue
    CarcinoembryonicantigenCEAUnitofMeasure
    CarcinoembryonicantigenCEADate
    lastColumn
End Enum

'' Populates the global gHeaderTexts array with the 82 standard DEFEND PRIME output column headers
Public Sub SetHeaderValues()
    ' Data Columns
    gHeaderTexts(1) = "Organisation Code"
    gHeaderTexts(2) = "Anonymised Identifier"
    gHeaderTexts(3) = "Baseline Height: Value"
    gHeaderTexts(4) = "Baseline Height: Unit of Measure"
    gHeaderTexts(5) = "Index Date: Value"
    gHeaderTexts(6) = "Index Date: Date"
    gHeaderTexts(7) = "Index Date: Patient's Age at Event"
    gHeaderTexts(8) = "Gender"
    gHeaderTexts(9) = "HbA1c (2a): Value"
    gHeaderTexts(10) = "HbA1c (2a): Date"
    gHeaderTexts(11) = "HbA1c (2b): Value"
    gHeaderTexts(12) = "HbA1c (2b): Date"
    gHeaderTexts(13) = "HbA1c (2c): Value"
    gHeaderTexts(14) = "HbA1c (2c): Date"
    gHeaderTexts(15) = "HbA1c (2d): Value"
    gHeaderTexts(16) = "HbA1c (2d): Date"
    gHeaderTexts(17) = "HbA1c (2e): Value"
    gHeaderTexts(18) = "HbA1c (2e): Date"
    gHeaderTexts(19) = "Weight (1a): Value"
    gHeaderTexts(20) = "Weight (1a): Unit of Measure"
    gHeaderTexts(21) = "Weight (1a): Date"
    gHeaderTexts(22) = "Weight (1b): Value"
    gHeaderTexts(23) = "Weight (1b): Unit of Measure"
    gHeaderTexts(24) = "Weight (1b): Date"
    gHeaderTexts(25) = "Weight (2a): Value"
    gHeaderTexts(26) = "Weight (2a): Unit of Measure"
    gHeaderTexts(27) = "Weight (2a): Date"
    gHeaderTexts(28) = "Weight (2b): Value"
    gHeaderTexts(29) = "Weight (2b): Unit of Measure"
    gHeaderTexts(30) = "Weight (2b): Date"
    gHeaderTexts(31) = "Weight (2c): Value"
    gHeaderTexts(32) = "Weight (2c): Unit of Measure"
    gHeaderTexts(33) = "Weight (2c): Date"
    gHeaderTexts(34) = "Weight (2d): Value"
    gHeaderTexts(35) = "Weight (2d): Unit of Measure"
    gHeaderTexts(36) = "Weight (2d): Date"
    gHeaderTexts(37) = "Weight (2e): Value"
    gHeaderTexts(38) = "Weight (2e): Unit of Measure"
    gHeaderTexts(39) = "Weight (2e): Date"
    gHeaderTexts(40) = "BMI (1a): Value"
    gHeaderTexts(41) = "BMI (1a): Date"
    gHeaderTexts(42) = "BMI (1b): Value"
    gHeaderTexts(43) = "BMI (1b): Date"
    gHeaderTexts(44) = "BMI (2a): Value"
    gHeaderTexts(45) = "BMI (2a): Date"
    gHeaderTexts(46) = "BMI (2b): Value"
    gHeaderTexts(47) = "BMI (2b): Date"
    gHeaderTexts(48) = "BMI (2c): Value"
    gHeaderTexts(49) = "BMI (2c): Date"
    gHeaderTexts(50) = "BMI (2d): Value"
    gHeaderTexts(51) = "BMI (2d): Date"
    gHeaderTexts(52) = "BMI (2e): Value"
    gHeaderTexts(53) = "BMI (2e): Date"
    gHeaderTexts(54) = "Height (1a): Value"
    gHeaderTexts(55) = "Height (1a): Unit of Measure"
    gHeaderTexts(56) = "Height (1a): Date"
    gHeaderTexts(57) = "Height (1b): Value"
    gHeaderTexts(58) = "Height (1b): Unit of Measure"
    gHeaderTexts(59) = "Height (1b): Date"
    gHeaderTexts(60) = "Height (2a): Value"
    gHeaderTexts(61) = "Height (2a): Unit of Measure"
    gHeaderTexts(62) = "Height (2a): Date"
    gHeaderTexts(63) = "Height (2b): Value"
    gHeaderTexts(64) = "Height (2b): Unit of Measure"
    gHeaderTexts(65) = "Height (2b): Date"
    gHeaderTexts(66) = "Height (2c): Value"
    gHeaderTexts(67) = "Height (2c): Unit of Measure"
    gHeaderTexts(68) = "Height (2c): Date"
    gHeaderTexts(69) = "Height (2d): Value"
    gHeaderTexts(70) = "Height (2d): Unit of Measure"
    gHeaderTexts(71) = "Height (2d): Date"
    gHeaderTexts(72) = "Height (2e): Value"
    gHeaderTexts(73) = "Height (2e): Unit of Measure"
    gHeaderTexts(74) = "Height (2e): Date"
    gHeaderTexts(75) = "Pancreatic cancer diagnosis: Date"
    gHeaderTexts(76) = "Pancreatic cancer diagnosis: Patient's Age at Event"
    gHeaderTexts(77) = "Cancer antigen 19-9 (CA 19-9): Value"
    gHeaderTexts(78) = "Cancer antigen 19-9 (CA 19-9): Unit of Measure"
    gHeaderTexts(79) = "Cancer antigen 19-9 (CA 19-9): Date"
    gHeaderTexts(80) = "Carcinoembryonic antigen (CEA): Value"
    gHeaderTexts(81) = "Carcinoembryonic antigen (CEA): Unit of Measure"
    gHeaderTexts(82) = "Carcinoembryonic antigen (CEA): Date"
End Sub

'' Establish once which worksheet dWS represents, so the whole processing pipeline can then refer to dWS
Public Sub SetWorksheet()
    Set dWS = ThisWorkbook.Worksheets("Sheet1") '' Set the global worksheet reference dWS to the worksheet with the tab name Sheet1
End Sub

'' Do not trust anything left in the shared global variables from the previous run; force the current run to establish its own worksheet, row positions, dataset and headers.
Public Sub ClearGlobalData()
    gFinalRow = 0 '' Clear any final row number retained from a previous run; 0 indicates that the final row has not yet been identified
    gFirstRowOfData = 0 '' Clear any first data row number retained from a previous run; 0 indicates that the first data row has not yet been identified
    Set dWS = Nothing '' Clear any worksheet reference retained from a previous run so it must be set again before use
    ReDim gDataSet(1) '' Clear any dataset retained from a previous run and leave gDataSet as a minimal one-element array
    Erase gHeaderTexts '' Clear any header values retained from a previous run so they must be populated again before use
End Sub
