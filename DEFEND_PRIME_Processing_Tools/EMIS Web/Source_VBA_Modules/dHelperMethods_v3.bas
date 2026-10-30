Attribute VB_Name = "dHelperMethods_v3"
' Revision history for dHelperMethods
' Differs from v2 in that ReadyApplication no longer relies on knowledge of the existing state, settings are now set to on or off
' Version number: v3
' Release Date: 19th December 2024

Option Explicit '' Forces the declaration of all variables in the module before they are used to minimise errors as undeclared variables will trigger a compile error
Option Private Module '' Means the module can only be used within the specified project

'' Sets Excel to either reduced-overhead settings for processing or restores normal operation
Public Sub ReadyApplication(processingMode As Boolean)
    Dim errDescription As String '' Stores any error description before restoration code can alter the Err object
    
    On Error GoTo RestoreSettings '' If changing an application setting causes an error, restore essential Excel settings before stopping
    
    If processingMode Then '' Set Excel to reduced-overhead settings whilst processing
        Application.ScreenUpdating = False '' Stop Excel updating the screen during processing
        On Error Resume Next '' Ignore an error if PrintCommunication is unavailable in the installed Excel version
        Application.PrintCommunication = False '' Suspend printer and page setup communication during processing
        Err.Clear '' Clear any ignored PrintCommunication error before restoring normal error handling
        On Error GoTo RestoreSettings '' Restore defined settings
        Application.Calculation = xlCalculationManual '' Prevent formula recalculation during processing
        Application.DisplayStatusBar = True '' Ensure display of the Excel status bar
        Application.EnableEvents = False '' Prevent Excel event procedures running during processing
    Else
         '' Restore normal Excel settings after processing
        Application.ScreenUpdating = True  '' Resume updating the Excel display
        On Error Resume Next '' Ignore errors if PrintCommunication is unavailable in the installed Excel version
        Application.PrintCommunication = True '' Resume communication with the printer and page setup
        Err.Clear '' Clear any ignored PrintCommunication error before restoring normal error handling
        On Error GoTo RestoreSettings '' Restore defined settings
        Application.Calculation = xlCalculationAutomatic '' Resume automatic formula calculation
        Application.DisplayStatusBar = True '' Ensure display of the Excel status bar
        Application.EnableEvents = True '' Resume Excel event procedures
    End If
    
    Exit Sub
    
RestoreSettings: '' Runs if an unexpected error occurs while running ReadyApplication
    errDescription = Err.Description '' Save the original error description before attempting restoration
    
    '' Restore normal Excel settings if changing the application settings itself causes an error
    On Error Resume Next
    Application.ScreenUpdating = True  '' Resume updating the Excel display
    Application.PrintCommunication = True '' Attempt to restore printer communication; any error (e.g. if incompatible Excel version) throws it's ignored during emergency clean-up
    Application.Calculation = xlCalculationAutomatic '' Resume automatic formula calculation
    Application.DisplayStatusBar = True '' Restore display of the Excel status bar
    Application.EnableEvents = True '' Resume Excel event procedures
    On Error GoTo 0
    
    MsgBox "An error occurred: " & errDescription
End Sub

'' Removes all cell formatting from the supplied worksheet without altering cell contents
Public Sub ClearFormatting(ws As Worksheet)
    ws.Cells.ClearFormats '' Remove formatting from all cells on the worksheet
End Sub

'' Force Excel to re-evaluate values containing hyphens by replacing each "-" with itself, correcting date interpretation in older Excel versions
'' Was running FormulaVersion from TH BIU VBA team but this is not compatible pre-MS365
Public Sub FixHyphenError(ws As Worksheet)
    ws.Cells.Replace What:="-", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False '' Re-enter hyphens throughout the worksheet so Excel re-evaluates affected values
End Sub

'' Converts an age into its corresponding 5-year age group, e.g. 52 becomes "50-54"
Function GetAgeRange(ageValue As Integer) As String
    GetAgeRange = CStr((Int(ageValue / 5) * 5)) & "-" & CStr((Int(ageValue / 5) * 5 + 4)) '' Round age down to the nearest multiple of 5 and use this as the lower bound, with four years added for the upper bound
End Function
