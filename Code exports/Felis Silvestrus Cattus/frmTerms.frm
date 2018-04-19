VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTerms 
   Caption         =   "felis silvestrus cattus"
   ClientHeight    =   12975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22980
   OleObjectBlob   =   "frmTerms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' MTK 14/03/2018
' DLL declarations to be used if Form Resize to be utilised, as Excel forms does not support this inherintly

''Written: August 02, 2010
''Author:  Leith Ross
''Summary: Makes the UserForm resizable by dragging one of the sides. Place a call
''         to the macro MakeFormResizable in the UserForm's Activate event.
'
'Private Declare Function SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long) As Long
'
'Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
'
'Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'
'Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Private Const m_sMODULE_NAME  As String = "frmTerms"

Private m_objTermsCollection    As clsTermCollection
Private m_bUpdatingLstTerms     As Boolean

' ******************************************************************
' ** Name:          Populate()
' ** Purpose:       Subroutine called after form creation to populate
' **                 listboxes and initialise Form
' ** Returns:       None
' ** Parameters:     Collection of Term/Definition Objects
' ** Author:        Matthew KELLY
' ** Date:          12/03/2018
' ** Revisions:     None
' ******************************************************************
Public Sub Populate(ByRef r_objTermsCollection As clsTermCollection)

    On Error GoTo errorHandler

    Const sFUNC_NAME As String = "Populate"

    ' Parameters used to control Do Loop for each Type definitions
    Dim lCounter        As Long     ' Counter for loop
    Dim lLowTermBound   As Long
    Dim lUppTermBound   As Long
    
    Dim asTypes()       As String
    Dim objTerm         As clsTerm

    ' set number of colums and width for list box
    lstTerms.ColumnCount = 3
    lstTerms.ColumnWidths = "100;50"
    
    ' set reference parameter copy to modular copy of class collection of terms
    Set m_objTermsCollection = r_objTermsCollection
    
    ' From the lstTerms data, build an array of Types data
    'asTypes = asBuildTypesArray()
    asTypes = m_objTermsCollection.asBuildTypesArray()
    
    'Call RefreshListBoxTermsWithCollection
    'Call lstTerms.Clear

    ' loop for each instance of Term contained within holding Collection
    For Each objTerm In m_objTermsCollection.TermCollection
    
        ' add data to listbox for Terms (lstTerms)
        Call lstTerms.AddItem
        lstTerms.List(objTerm.lIndex, 0) = objTerm.sTerm
        lstTerms.List(objTerm.lIndex, 1) = objTerm.sType
        lstTerms.List(objTerm.lIndex, 2) = objTerm.sDefinition
        lstTerms.Selected(objTerm.lIndex) = objTerm.bSelected
        
    Next objTerm
    
    ' Then loop for each Type and populate Listbox for types
    lLowTermBound = LBound(asTypes)
    lUppTermBound = UBound(asTypes)
    lCounter = lLowTermBound
    Do
        Call lstTypes.AddItem
        lstTypes.List(lCounter - 1) = asTypes(lCounter)
        lCounter = lCounter + 1
    Loop Until lCounter > lUppTermBound

resumeFunc:
    ' tidy up
    Exit Sub

errorHandler:
    Call MsgBox("Error in function: " & sFUNC_NAME & vbCr & _
                " Module: " & m_sMODULE_NAME & vbCr & _
                " Spreadsheet: " & g_sPROGRAM_NAME & vbCr & _
                " Error message: " & Err.Number & ": " & Err.Description)
    GoTo resumeFunc

End Sub


Private Sub RefreshListBoxesWithCollection()

    On Error GoTo errorHandler

    Const sFUNC_NAME As String = "RefreshListBoxTermsWithCollection"

    Dim lObjNoAllTypesSelected          As Long
    Dim lObjAllTypesCounter             As Long
    Dim asObjAllTypesSelected()         As String
    
    Dim lLowTermBound                   As Long
    Dim lUppTermBound                   As Long
    Dim lCounter                        As Long
    
    Dim bMatch                          As Boolean
    
    Dim objTerm                         As clsTerm
    
    
    m_bUpdatingLstTerms = True

    ' loop for each instance of Term contained within holding Collection
    For Each objTerm In m_objTermsCollection.TermCollection
    
        ' add data to listbox for Terms (lstTerms)
        'Call lstTerms.AddItem
        lstTerms.List(objTerm.lIndex, 0) = objTerm.sTerm
        lstTerms.List(objTerm.lIndex, 1) = objTerm.sType
        lstTerms.List(objTerm.lIndex, 2) = objTerm.sDefinition
        lstTerms.Selected(objTerm.lIndex) = objTerm.bSelected
        
    Next objTerm
    

    For lCounter = 0 To lstTypes.ListCount - 1
        bMatch = False
        lObjNoAllTypesSelected = m_objTermsCollection.lAllSelectedTypesArray(asObjAllTypesSelected())
        For lObjAllTypesCounter = 1 To lObjNoAllTypesSelected
            ' if match found need to highlight
            If InStr(lstTypes.List(lCounter), asObjAllTypesSelected(lObjAllTypesCounter)) <> 0 Then
                lstTypes.Selected(lCounter) = True
                bMatch = True
                Exit For
            End If
        
        Next lObjAllTypesCounter
    
        If Not bMatch Then
            lstTypes.Selected(lCounter) = False
        End If
    
    Next lCounter
    
resumeFunc:
    ' tidy up
    m_bUpdatingLstTerms = False
    Exit Sub

errorHandler:
    Call MsgBox("Error in function: " & sFUNC_NAME & vbCr & _
                " Module: " & m_sMODULE_NAME & vbCr & _
                " Spreadsheet: " & g_sPROGRAM_NAME & vbCr & _
                " Error message: " & Err.Number & ": " & Err.Description)
    GoTo resumeFunc

End Sub

' ******************************************************************
' ** Name:          lBuildTypesArray()
' ** Purpose:       Function to loop through each item in ListBox Terms and
' **                 create an array containing Types
' ** Returns:       Number of items in Types Array
' ** Parameters:    String array by reference to contain all Types in lstTypes
' ** Author:        Matthew KELLY
' ** Date:          12/03/2018
' ** Revisions:     13/03/2018 - Added Optional parameter for use when creating Document with selected items only
' ******************************************************************
Private Function lBuildTypesArray(ByRef r_asTypes() As String) As Long
    
    On Error GoTo errorHandler

    Const sFUNC_NAME                As String = "Populate"
    
    ' Varibales to be used whilst looping for each List Term (Outer loop)
    Dim lLstTypesTotCount           As Long         ' The Total List Terms Count
    Dim lLstTypesCounter            As Long         ' Counter, and
    
    ' Varibales to be used whilst looping for each List Term (Inner loop)
    Dim lTypeCount                  As Long         ' Total number of current Types located
    Dim bNewTypeFound               As Boolean      ' Has the Current Type already been populated into array?
    
    
    ' No Matching types have been located yet
    lTypeCount = 0
    
    ' How many terms are there currently in List Box Terms
    lLstTypesTotCount = lstTypes.ListCount
    ' Loop for each (List Boxes start counting from Zero!)
    For lLstTypesCounter = 0 To lLstTypesTotCount - 1
        
        bNewTypeFound = False
        If lstTypes.Selected(lLstTypesCounter) Then
            bNewTypeFound = True
        End If
    
        If bNewTypeFound Then
            lTypeCount = lTypeCount + 1
            ReDim Preserve r_asTypes(1 To lTypeCount)
            r_asTypes(lTypeCount) = lstTypes.List(lLstTypesCounter)
        End If
    
    Next lLstTypesCounter
        
    
resumeFunc:
    lBuildTypesArray = lTypeCount
    Exit Function

errorHandler:
    Call MsgBox("Error in function: " & sFUNC_NAME & vbCr & _
                " Module: " & m_sMODULE_NAME & vbCr & _
                " Spreadsheet: " & g_sPROGRAM_NAME & vbCr & _
                " Error message: " & Err.Number & ": " & Err.Description)
    GoTo resumeFunc

End Function


Private Sub cmdExit_Click()
    Call Unload(Me)
End Sub

' ******************************************************************
' ** Name:          cmdPrint_Click()
' ** Purpose:
' ** Returns:
' ** Parameters:
' ** Author:        Matthew KELLY
' ** Date:          12/03/2018
' ** Revisions:
' ******************************************************************
Private Sub cmdPrint_Click()

    On Error GoTo errorHandler

    Const sFUNCTION_NAME        As String = "cmdPrint_Click"
    
    Dim sDfuReference           As String
    Dim asTypes()               As String
 
    Dim sCheckReturn            As String
    Dim sPrintMessage           As String
    Dim vbMessageBoxStyle       As VbMsgBoxStyle
 
    sDfuReference = txtDfuReference.Text
    
    ReDim asTypes(1 To 1)
    asTypes(1) = ""
    If chkPrintByCategory.Value = True Then
        asTypes = m_objTermsCollection.asBuildTypesArray(True)
    End If
    
    sCheckReturn = sCreateGlossaryFromCollection(m_objTermsCollection.TermCollection, asTypes(), sDfuReference)
    If sCheckReturn <> "" Then
        sPrintMessage = "Glossary of terms successfully generated as " & sCheckReturn & ". " & _
        vbCr & vbCr & "Please delete the relevant folder once saved to own report location."
        vbMessageBoxStyle = vbInformation
    Else
        sPrintMessage = "Error creating glossary of terms. " & _
        vbCr & vbCr & " Sorry!"
        vbMessageBoxStyle = vbCritical
    End If
    
    Call MsgBox(sPrintMessage, vbMessageBoxStyle, "Report status")
  
ResumeProgram:
    ' tidy up
    Exit Sub

errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume ResumeProgram
    
End Sub


' ******************************************************************
' ** Name:          lstTypes_Change()
' ** Purpose:
' ** Returns:
' ** Parameters:
' ** Author:        Matthew KELLY
' ** Date:          12/03/2018
' ** Revisions:
' ******************************************************************
Private Sub lstTerms_Change()

    On Error GoTo errorHandler

    Const sFUNCTION_NAME        As String = "lstTypes_Change"
    
    Dim lListTermCounter    As Long
    Dim objTerms            As clsTerm
    
    If Not m_bUpdatingLstTerms Then
        
        'Set objTerms = New clsTerm
        
        For Each objTerms In m_objTermsCollection.TermCollection
        
            lListTermCounter = objTerms.lIndex
            If lstTerms.Selected(lListTermCounter) Then
                objTerms.bSelected = True
            Else
                objTerms.bSelected = False
            End If
            
        Next objTerms
        
        Call RefreshListBoxesWithCollection
    
    End If

ResumeProgram:
    Exit Sub

errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume ResumeProgram
    
End Sub


' ******************************************************************
' ** Name:          lstTypes_Change()
' ** Purpose:       The modular collection is a memory copy of list box and the only thing
' **                 that needs updating for this instance of all items selected
' **                As such first thing that needs establising is what lstBoxType has changed, if any...
' **                 Then remove all relevant types if changed from selected to deselected
' **                OR add all relevant types if changed from deselected to selected
' **                Then refresh lstBoxTerms to reflect modular Collection.
' ** Returns:       None
' ** Parameters:    None
' ** Author:        Matthew KELLY
' ** Date:          12/03/2018
' ** Revisions:     16/03/2018 - amended coding to change modular collection object and to
' **                 subsequently update lstBoxTerms. THis would be far easier if VBA recorded
' **                 which lstBox item had been updated
' ******************************************************************
Private Sub lstTypes_Change()

    On Error GoTo errorHandler

    Const sFUNCTION_NAME        As String = "lstTypes_Change"
    
    Dim lObjNoTypesAllSelected              As Long
    Dim lObjCounter                         As Long
    Dim asObjAllTypesSelected()             As String
    
    Dim lLstBoxNoTypesSelected              As Long
    Dim lLstBoxCounter                      As Long
    Dim asLstBoxTypesSelected()             As String
    
    Dim bMatchFound                         As Boolean
    
    
    If Not m_bUpdatingLstTerms Then
    
        ' need to establish if any of the current selected items (before change) are ALL selected
        lObjNoTypesAllSelected = m_objTermsCollection.lAllSelectedTypesArray(asObjAllTypesSelected())
        
        ' then whether or not and how many lstTypes are selected
        lLstBoxNoTypesSelected = lBuildTypesArray(asLstBoxTypesSelected())
        
        ' from this, compare against new currently selected and only change those not listed
        If lLstBoxNoTypesSelected = 0 And lObjNoTypesAllSelected = 0 Then
            
            ' no changes necessary
            Call MsgBox("Nothing in memory, no lst box items selected?")
    
        ElseIf lLstBoxNoTypesSelected > 0 And lObjNoTypesAllSelected = 0 Then
            
            For lLstBoxCounter = 1 To lLstBoxNoTypesSelected
            
                Call m_objTermsCollection.lSelectItemsWithCorrespondingType(asLstBoxTypesSelected(lLstBoxCounter), True)
                
            Next lLstBoxCounter
            
        
        ElseIf lLstBoxNoTypesSelected = 0 And lObjNoTypesAllSelected > 0 Then
        
            For lObjCounter = 1 To lObjNoTypesAllSelected
            
                Call m_objTermsCollection.lSelectItemsWithCorrespondingType(asObjAllTypesSelected(lObjCounter), False)
            
            Next lObjCounter
        
        Else
        
            bMatchFound = False
            ' are the number of selected list box types more than those in memory, if so need to add some
            If lLstBoxNoTypesSelected > lObjNoTypesAllSelected Then
            
                ' else need to remove
                
                For lLstBoxCounter = 1 To lLstBoxNoTypesSelected
            
                    bMatchFound = False
                    For lObjCounter = 1 To lObjNoTypesAllSelected
                
                        ' find list box items selected that aren't in memory: select all
                        If InStr(asLstBoxTypesSelected(lLstBoxCounter), asObjAllTypesSelected(lObjCounter)) <> 0 Then
                            bMatchFound = True
                            Exit For
                        End If
                        
                    Next lObjCounter
                    
                    If Not bMatchFound Then
                        Call m_objTermsCollection.lSelectItemsWithCorrespondingType(asLstBoxTypesSelected(lLstBoxCounter), Not bMatchFound)
                    End If
                    
                Next lLstBoxCounter
                
            Else
            
                For lObjCounter = 1 To lObjNoTypesAllSelected
                
                    bMatchFound = False
                    For lLstBoxCounter = 1 To lLstBoxNoTypesSelected
                
                        ' find list box items selected that aren't in memory: select all
                        If InStr(asLstBoxTypesSelected(lLstBoxCounter), asObjAllTypesSelected(lObjCounter)) <> 0 Then
                            bMatchFound = True
                            Exit For
                        End If
                        
                    Next lLstBoxCounter
                    
                    If Not bMatchFound Then
                        Call m_objTermsCollection.lSelectItemsWithCorrespondingType(asObjAllTypesSelected(lObjCounter), bMatchFound)
                        Exit For
                    End If
                    
                Next lObjCounter
                
            ' find memory items that aren't selected: deselect all
            
            End If
            
        End If
        
        ' then refresh from modular collection
        Call RefreshListBoxesWithCollection
    
    End If
    
ResumeProgram:
    Exit Sub

errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume ResumeProgram
    
End Sub


' ******************************************************************
' ** Name:          txtDfuReference_Change()
' ** Purpose:       To monitor changes in DFU Reference textbox and enable Print command if text present
' ** Returns:       None
' ** Parameters:    Event!
' ** Author:        Matthew KELLY
' ** Date:          12/03/2018
' ** Revisions:     None
' ******************************************************************
Private Sub txtDfuReference_Change()
    If Trim$(txtDfuReference.Text) <> "" Then
        cmdPrint.Enabled = True
        txtDfuReference.Text = Replace(txtDfuReference.Text, "\", "-")
    Else
        cmdPrint.Enabled = False
    End If
End Sub



' MTK 14/03/2018
'  Subs to be used if Form Resize to be utilised, as Excel forms does not support this inherintly

'' ******************************************************************
'' ** Name:          UserForm_Resize()
'' ** Purpose:
'' ** Returns:
'' ** Parameters:
'' ** Author:        Matthew KELLY
'' ** Date:          12/03/2018
'' ** Revisions:
'' ******************************************************************
'Private Sub UserForm_Resize()
'
'    On Error GoTo errorHandler
'
'    Const sFUNCTION_NAME        As String = "UserForm_Resize"
'
'    Dim lFormWidth              As Long
'    Dim lFormHeight             As Long
'    Dim lFrame1Height           As Long
'
'
'    ' get height and width of form
'    lFormWidth = frmTerms.Width
'    lFormHeight = frmTerms.Height
'
'    fraFrame1.Left = 0
'    fraFrame1.Top = 0
'    fraFrame1.Width = lFormWidth
'    lFrame1Height = (2 * lFormHeight) / 3
'    fraFrame1.Height = lFrame1Height
'
'    fraFrame2.Top = lFrame1Height
'    fraFrame2.Left = 0
'    fraFrame2.Width = lFormWidth / 2
'    fraFrame2.Height = lFormHeight / 3
'
'    lstTerms.Width = fraFrame1.Width
'    lstTerms.Height = fraFrame1.Height
'
'    lstTypes.Width = fraFrame2.Width
'    lstTerms.Height = fraFrame2.Height
'
'    cmdExit.Top = lFormHeight - cmdExit.Height
'    cmdExit.Left = lFormWidth - cmdExit.Width
'
'    cmdPrint.Top = lFormHeight - cmdExit.Height - cmdPrint.Height
'    cmdExit.Left = lFormWidth - cmdPrint.Width
'
'
'ResumeProgram:
'    Exit Sub
'
'errorHandler:
'    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
'                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
'    Resume ResumeProgram
'
'End Sub
'
'
'
''Written: August 02, 2010
''Author:  Leith Ross
''Summary: Makes the UserForm resizable by dragging one of the sides. Place a call
''         to the macro MakeFormResizable in the UserForm's Activate event.
'
'Public Sub MakeFormResizable()
'
'    Dim lStyle As Long
'    Dim hWnd As Long
'    Dim RetVal
'
'    Const WS_THICKFRAME = &H40000
'    Const GWL_STYLE As Long = (-16)
'
'    hWnd = GetActiveWindow
'
'    'Get the basic window style
'    lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME
'
'    'Set the basic window styles
'    RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)
'
'    'Clear any previous API error codes
'    SetLastError 0
'
'    'Did the style change?
'    If RetVal = 0 Then
'        MsgBox "Unable to make UserForm Resizable."
'    End If
'
'End Sub
