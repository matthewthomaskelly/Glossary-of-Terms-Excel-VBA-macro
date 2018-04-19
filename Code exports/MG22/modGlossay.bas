Attribute VB_Name = "modGlossay"
Option Explicit

Public Const g_sPROGRAM_NAME As String = "MTK22(b) TEST.docm"
Private Const m_sMODULE_NAME As String = "modGlossay"

' ******************************************************************
' ** Name:          CreateGlossaryOfSelectedText()
' ** Purpose:       Macro to take selected text and query glossary of terms spreadsheet
' **                 and produce a report containing relevant terms
' ** Returns:       Report - Glossary of Terms
' ** Parameters:    -
' ** Author:        Matthew KELLY
' ** Date:          15/03/2018
' ** Revisions:     20/03/2018 - added coding to check for empty string, request DFU reference
' **                              and remove whitespace including Carriage Returns (vbCR)
' ** NOTE:
' ******************************************************************
Public Sub CreateGlossaryOfSelectedText()

    On Error GoTo errorHandler

    Const sFUNCTION_NAME        As String = "CreateGlossaryOfSelectedText()"
    
    Dim sDfuReference           As String   ' String to store user input of DFU Reference
    Dim sSelectedText           As String   ' Selected text for checks and passing to Excel Spreadsheet
    Dim sExcelPath              As String   ' made up of current path and Excel spreadsheet name
    
    Dim Document                As Document ' To be assigned this Document and used to get current working directory
    Dim ExcelApp                As Object   ' Excel app -> Workbook -> queried
    
    ' get selected text, remove Carriage returns and whitespace
    sSelectedText = Selection.Text
    sSelectedText = Replace(sSelectedText, vbCr, "")
    sSelectedText = Trim$(sSelectedText)
    
    ' check there is some selected text
    If sSelectedText <> "" Then
    
        ' what is the DFU reference
        sDfuReference = InputBox("Please enter DFU reference: ", "DFU reference")
        ' Don't pass '\' as reference used to create new folder name
        sDfuReference = Replace(sDfuReference, "\", "-")

        ' Get current working directory and append for Worksheet name
        Set Document = ThisDocument
        sExcelPath = ThisDocument.Path
        sExcelPath = sExcelPath & "\Felis Silvestris Cattus - Glossary of Terms.xlsm"
        
        ' create Excel App object
        Set ExcelApp = CreateObject("Excel.Application")
        ' display
        ExcelApp.Visible = True
        ' open
        Call ExcelApp.WorkBooks.Open(sExcelPath)
        ' run macro, passing Selected Text and DFU reference
        Call ExcelApp.Run("g_create_glossary_of_terms_from_selection", sSelectedText, sDfuReference)
        ' close Excel Workbook
        Call ExcelApp.WorkBooks.Close
        'call excelapp
    
    Else
        
        ' Display error message if no text selected
        Call MsgBox("No text selected, please try again!", vbExclamation + vbOKOnly, "No selection")
    
    End If
    
ResumeProgram:
    ' tidy up
    Set Document = Nothing
    Set ExcelApp = Nothing
    Exit Sub

errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume ResumeProgram
    
End Sub
