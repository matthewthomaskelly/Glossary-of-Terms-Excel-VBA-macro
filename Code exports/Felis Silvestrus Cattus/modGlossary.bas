Attribute VB_Name = "modGlossary"
' ******************************************************************
' ** Name:          Glossary of Terms.xls
' ** Purpose:       Excel Spreadsheet that contains a list of Terms and Definitions
' ** Author:        Matthew KELLY
' ** Date:          12/03/2015
' ** Revisions:     First Version Release March 2018
' ******************************************************************
Option Explicit

Public Const g_sPROGRAM_NAME As String = "Glossary of Terms.xls"
Private Const m_sMODULE_NAME As String = "modGlossary"


' ******************************************************************
' ** Name:          PrintMe()
' ** Purpose:       Macro Called to Populate and then display main Form Interface
' ** Returns:       None
' ** Parameters:     Each Term in placed into a Class instance and then added to a Collection
' ** Author:        Matthew KELLY
' ** Date:          12/03/2015
' ** Revisions:     None
' ******************************************************************
Public Sub PrintMe()
Attribute PrintMe.VB_ProcData.VB_Invoke_Func = " \n14"

    On Error GoTo errorHandler

    Const sFUNCTION_NAME        As String = "PrintMe Macro"
    
    Dim objTermsCollection      As clsTermCollection    ' Collection for all Terms and Definitions Classes
    Dim frmObjTerms             As frmTerms     ' Class to gather information for Excel Data
    
    ' Populate Collection with Class data representing Excel Data
    Set objTermsCollection = objPopulateCollection("A2")
    
    ' Create instance of form...
    Set frmObjTerms = New frmTerms
    ' ... populate with data, and...
    Call frmObjTerms.Populate(objTermsCollection)
    ' ....display
    frmObjTerms.Show
    
ResumeProgram:
    ' tidy up
    Set frmObjTerms = Nothing
    Set objTermsCollection = Nothing
    Exit Sub

errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume ResumeProgram
    
End Sub


' ******************************************************************
' ** Name:          objPopulateCollection()
' ** Purpose:       Iterates over Excel Spreadsheet data and first populate class instance
' **                 for each Row and subsequently adds this to a collection for return
' ** Returns:       Collection of Classes containing Spreadsheet data
' ** Parameters:
' ** Author:        Matthew KELLY
' ** Date:          12/03/2015
' ** Revisions:
' ******************************************************************
Private Function objPopulateCollection(ByVal v_sStartCell As String) As clsTermCollection
    
    On Error GoTo errorHandler

    Const sFUNCTION_NAME        As String = "objPopulateCollection"
    
    Dim sMaxRowColumn           As String   ' Maximum rows in Excel spreadsheet
    Dim lTotalCompFilesCount    As Long     ' Total number of rows
    Dim lCounter                As Long     ' counter for loop for each row
    
    Dim objTermsCollection      As clsTermCollection    ' Class instance containing collection to store each object instance of class term
    Dim objTerm                 As clsTerm              ' these will be instantiated and added to collection
    
    ' set collection
    Set objTermsCollection = New clsTermCollection
    
    ' get max numbers of row/column in Spreadsheet
    sMaxRowColumn = m_sGetUsedCellRange("Definitions")
    ' then calculate Rows count
    lTotalCompFilesCount = Trim(Right(sMaxRowColumn, Len(sMaxRowColumn) - 1))
    
    ' loop for each row starting at row 2 as this is where data starts
    For lCounter = 2 To lTotalCompFilesCount
        
        ' Class instance for Spreadsheet data, select corrct Sheet and recover data, assigning to Class instance before
        Set objTerm = New clsTerm
        Sheets("Definitions").Select
        objTerm.lIndex = lCounter - 2
        objTerm.bSelected = False
        objTerm.sTerm = Range("A" & lCounter).Value
        objTerm.sType = Range("B" & lCounter).Value
        objTerm.sDefinition = Range("C" & lCounter).Value
        Call objTerm.setcolAlternative(Range("D" & lCounter).Value)
        ' ... adding to Collection
        Call objTermsCollection.Add(objTerm, Str(lCounter - 2))
        
    Next lCounter
    
ResumeProgram:
    ' return Collection as is...
    Set objPopulateCollection = objTermsCollection
    Exit Function

errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume ResumeProgram
    
End Function

' ******************************************************************
' ** Name:          m_sGetUsedCellRange()
' ** Purpose:       Finds and locates the last used data cell in specified sheet name
' ** Returns:       Used data cell range
' ** Parameters:     v_sSheetName = specified Sheet Name
' **                v_bRowOnly   = optional boolean value allowing rows used only to be returned; default is off
' ** Author:        Matthew KELLY
' ** Date:          10/03/2015
' ** Revisions:     None
' ******************************************************************
Private Function m_sGetUsedCellRange(ByVal v_sSheetName As String, _
                                     Optional ByVal v_bRowOnly As Boolean = False _
                                    ) As String

    On Error GoTo errorHandler

    Const sFUNCTION_NAME        As String = "m_sGetUsedCellRange"
    
    Dim lRowCount               As Long
    Dim lColumnCount            As Long
    Dim sMaxCellRange           As String
    Dim sSelectCellRange        As String
    
    ' Select specified Worksheet
    With ActiveWorkbook.Worksheets(v_sSheetName)
                      
        ' get used row/column, used row/column count and subtract one
        lRowCount = .UsedRange.Row + .UsedRange.Rows.Count - 1
        lColumnCount = .UsedRange.Column + .UsedRange.Columns.Count - 1
        ' convert range to specified format for Excel
        sMaxCellRange = m_sConvertToLetter(lColumnCount)
        
    End With

    ' Return rows if specified, otherwise convert to Excel format
    If v_bRowOnly = True Then
        m_sGetUsedCellRange = sMaxCellRange
    Else
        m_sGetUsedCellRange = sMaxCellRange & Trim$((Str(lRowCount)))
    End If

ResumeProgram:
    Exit Function

errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume ResumeProgram

End Function


' ******************************************************************
' ** Name:          m_sConvertToLetter()
' ** Purpose:       Function to convert column number to specified alphanumeric equivalent
' ** Returns:       String corresponding column name. e.g 1=A, 27=AA
' ** Parameters:     Long; lColumn - total number of columns
' ** Author:        Matthew KELLY
' ** Date:          26/02/2015
' ** Revisions:     27/02/2015 - Matthew KELLY, tightened coding, added comments and error handling.
' **                10/03/2015 - Matthew KELLY. Changed function from private to public and placed into modCommon
' ** NOTE:          THIS FUNCTION ASSUME CELLS WILL NOT EXCEED TWO LETTERS i.e - above 'ZZ'
' ******************************************************************
Private Function m_sConvertToLetter(ByVal lColumn As Long) As String

    On Error GoTo errorHandler

    Const sFUNCTION_NAME        As String = "m_sConvertToLetter"

    Dim lAlpha          As Long
    Dim lRemainder      As Long
    
    ' divide column number by 27 and force integer value. If number is above 27 then corresponding column will have two alphanumeric characters
    lAlpha = Int(lColumn / 26)
    ' there will be a remainder of the integer calculation if iAplha larger than 0 and more than two letters
    lRemainder = lColumn - (lAlpha * 26)
    
    If lAlpha > 0 Then
        m_sConvertToLetter = Chr(lAlpha + 64)
    End If
    
    If lRemainder > 0 Then
        m_sConvertToLetter = m_sConvertToLetter & Chr(lRemainder + 64)
    End If

ResumeProgram:
    Exit Function

errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume ResumeProgram

End Function

' ******************************************************************
' ** Name:          g_create_glossary_of_terms_from_selection()
' ** Purpose:       Macro callable from Outside world! (intended )
' ** Returns:       Report?!
' ** Parameters:    -
' ** Author:        Matthew KELLY
' ** Date:          19/03/2018
' ** Revisions:     20/03/2018 - Removed Report creation into new function sCreateGlossaryFromCollection() that is called from Form also
' ** NOTE:
' ******************************************************************
Public Sub g_create_glossary_of_terms_from_selection(ByVal v_sSelectedText As String, Optional v_sDfuReference As String)
    
    On Error GoTo errorHandler

    Const sFUNCTION_NAME        As String = "g_create_glossary_of_terms_from_selection"
    
    ' Variables for All Excel data
    Dim objAllTermsCollection           As clsTermCollection    ' Collection for all Terms and Definitions Classes
    Dim objTerm                         As clsTerm              ' Class instance to loop through each term in collection above
    
    ' Variables for Excel data to be included in report
    Dim varAltText                      As Variant              ' Used to loop through string Collection of Alternative terms (created in objPopulateCollection)
    Dim asTypes()                       As String               ' String array to inlude selected terms for passing to Glossary of Terms creation tool
    
    
    ' Populate Collection with Class data representing Excel Data
    Set objAllTermsCollection = objPopulateCollection("A2")

    ' loop for each Term data (i.e. each Row)
    For Each objTerm In objAllTermsCollection.TermCollection
    
        ' Is the currently examined term in the selected text to be considered
        If InStr(Trim$(LCase$(v_sSelectedText)), LCase$(objTerm.sTerm)) <> 0 _
        And objTerm.sTerm <> "" Then
            ' great, change status to selected
            objTerm.bSelected = True
        Else
            ' Examine each Alternative term to see if any matches (e.g Clone rather that Acquisiton)
            For Each varAltText In objTerm.colAlternative
                If InStr(Trim$(LCase$(v_sSelectedText)), LCase$(varAltText)) <> 0 _
                And varAltText <> "" Then
                    ' great, change status to selected
                    objTerm.bSelected = True
                End If
            Next varAltText
        End If
    
    Next objTerm
     
    ' Create Types string array
    asTypes = objAllTermsCollection.asBuildTypesArray(True)

    ' Try to create Glossary of Terms
    Call sCreateGlossaryFromCollection(objAllTermsCollection.SelectedTermsCollection, asTypes(), v_sDfuReference)

ResumeProgram:
    ' tidy up
    Set objAllTermsCollection = Nothing
    Exit Sub

errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume ResumeProgram
    
End Sub

' ******************************************************************
' ** Name:          sCreateGlossaryFromCollection()
' ** Purpose:       -
' ** Returns:       -
' ** Parameters:    -
' ** Author:        Matthew KELLY
' ** Date:          20/03/2018
' ** Revisions:     -
' ** NOTE:
' ******************************************************************
Public Function sCreateGlossaryFromCollection(ByRef r_colTermsCollection As Collection, _
                                                ByRef r_asTypes() As String, _
                                                ByVal v_sDfuReference As String) As String
                                                
    
    On Error GoTo errorHandler

    Const sFUNCTION_NAME        As String = "lCreateGlossaryFromCollection"

    Dim sReturnValue            As String

    Dim sWordPath               As String
    Dim sSavePath               As String

    Dim Workbook                As Workbook
    Dim WordApp                 As Object
        
 
    ' assume  errors
    sReturnValue = ""
        
    ' preparing to open word
    Set Workbook = ThisWorkbook
    sWordPath = ThisWorkbook.Path
    sWordPath = sWordPath & "\Terms and Definitions.docm"
    
    Set WordApp = CreateObject("Word.Application")
    
    WordApp.Visible = True
    
    Call WordApp.Documents.Open(sWordPath)
    
    Call WordApp.Run("g_CreateGlossaryFromCollection", r_colTermsCollection, v_sDfuReference, r_asTypes)
    
    sSavePath = WordApp.ActiveDocument.Path
    sSavePath = sSavePath & "\" & Date$ & " " & Replace(Time$, ":", "") & " " & v_sDfuReference
    MkDir (sSavePath)
    sSavePath = sSavePath & "\" & v_sDfuReference & ".doc"
    Call WordApp.ActiveDocument.SaveAs2(Filename:=sSavePath, FileFormat:=16) ' FileFormat:=wdFormatDocumentDefault)
    Call WordApp.Documents.Open(sSavePath)
    
    sSavePath = Replace(sSavePath, ".doc", ".pdf")
    Call WordApp.ActiveDocument.SaveAs2(Filename:=sSavePath, FileFormat:=17) ' FileFormat:=wdFormatDocumentDefault)

    sReturnValue = sSavePath

ResumeProgram:
    ' tidy up
    Set Workbook = Nothing
    Set WordApp = Nothing
    sCreateGlossaryFromCollection = sReturnValue
    Exit Function

errorHandler:
    ' unexpected error
    sReturnValue = ""
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume ResumeProgram

End Function
