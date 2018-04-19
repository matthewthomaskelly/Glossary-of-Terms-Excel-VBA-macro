VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTermCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const m_sMODULE_NAME    As String = "clsTermCollection"

Private m_objTermCollection     As Collection
Private m_lTermsCount           As Long


Private Sub Class_Initialize()
    Set m_objTermCollection = New Collection
    m_lTermsCount = 0
End Sub

Private Sub Class_Terminate()
    Set m_objTermCollection = Nothing
End Sub

Public Property Get TermCollection() As Collection
    Set TermCollection = m_objTermCollection
End Property


Public Property Get TermsCount() As Long
    TermsCount = m_lTermsCount
End Property

Public Sub Add(ByRef r_objTerm As clsTerm, ByVal v_sIndex As String)

    On Error GoTo errorHandler
    
    Const sFUNCTION_NAME As String = "AddTerm"
    
    Call m_objTermCollection.Add(r_objTerm, v_sIndex)
    
    m_lTermsCount = m_lTermsCount + 1
    
       
resumeFunc:
    ' tidy up
    Exit Sub
    
errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume resumeFunc

End Sub

Public Property Get SelectedTermsCollection() As Collection
    
    On Error GoTo errorHandler
    
    Const sFUNCTION_NAME As String = "ReturnSelectedTerms"

    Dim objReturnCollection     As Collection
    Dim objTerm                 As clsTerm
    
    Set objReturnCollection = New Collection
    
    For Each objTerm In m_objTermCollection
        If objTerm.bSelected Then
            Call objReturnCollection.Add(objTerm)
        End If
    Next objTerm
    
resumeFunc:
    ' tidy up
    Set SelectedTermsCollection = objReturnCollection
    Set objReturnCollection = Nothing
    Exit Property
    
errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume resumeFunc

End Property


Public Function lSelectItemsWithCorrespondingType(ByVal v_sType As String, ByVal v_bSelection As Boolean) As Long

    On Error GoTo errorHandler
    
    Const sFUNCTION_NAME As String = "lSelectItemsWithCorrespondingType"
    
    Dim lSelectedItemsCount     As Long
    Dim objTerm                As clsTerm
    
    Set objTerm = New clsTerm
    
    lSelectedItemsCount = 0
    
    For Each objTerm In m_objTermCollection
    
        If InStr(v_sType, objTerm.sType) <> 0 _
        And objTerm.sType <> "" Then
            lSelectedItemsCount = lSelectedItemsCount + 1
            objTerm.bSelected = v_bSelection
        End If
    
    Next objTerm
    
resumeFunc:
    ' tidy up
    'objTerm = Nothing
    lSelectItemsWithCorrespondingType = lSelectedItemsCount
    Exit Function
    
errorHandler:
    Call MsgBox(Err.Description & ":" & Str$(Err.Number) & ". in: " _
                & g_sPROGRAM_NAME & "; " & m_sMODULE_NAME & "; " & sFUNCTION_NAME & "()")
    Resume resumeFunc
    
End Function

' ******************************************************************
' ** Name:          asBuildTypesArray()
' ** Purpose:       Function to loop through each item in ListBox Terms and
' **                 create an array containing Types
' ** Returns:       Array String of Types
' ** Parameters:    Optional boolean specifying if ONLY Selected Types to be populated to array
' ** Author:        Matthew KELLY
' ** Date:          12/03/2018
' ** Revisions:     13/03/2018 - Added Optional parameter for use when creating Document with selected items only
' **                15/03/2018 - Moved into Class Object for holding collection of Terms
' ******************************************************************
Public Function asBuildTypesArray(Optional v_bSelectedOnly As Boolean = False) As String()
    
    On Error GoTo errorHandler

    Const sFUNC_NAME                As String = "asBuildTypesArray"
    
    ' Varibales to be used whilst looping for each List Term (Outer loop)
    Dim lLstTermsTotCount           As Long         ' The Total List Terms Count
    Dim lLstTermsCounter            As Long         ' Counter, and
    
    Dim objTerm                     As clsTerm
    Dim sLstCurrentTermType         As String       ' Current Term Type (definition) String Text
    
    ' Varibales to be used whilst looping for each List Term (Inner loop)
    Dim lTypeCount                  As Long         ' Total number of current Types located
    Dim lCounterTypes               As Long         ' Counter for each current Types in array
    
    Dim bNewTypeFound               As Boolean      ' Has the Current Type already been populated into array?
    
    ' String array to be populated with each List TYPE
    Dim asTypes()                   As String
    
    ' No Matching types have been located yet
    lTypeCount = 0
    
    ' Loop for each objTerm in collection
    For Each objTerm In m_objTermCollection
        
        ' Current Term being examined in List Box Terms
        sLstCurrentTermType = objTerm.sType
        
        ' assume new type located until proven otherwise
        bNewTypeFound = True
        
        ' find first item
        If lTypeCount = 0 Then
        
            ' loop for each and add to array if any located
            If v_bSelectedOnly Then
                If objTerm.bSelected = False Then
                    bNewTypeFound = False
                End If
            Else
                bNewTypeFound = True
            End If
        
        Else
        
            ' How many types are currently present
            lTypeCount = UBound(asTypes) - LBound(asTypes) + 1
            
            ' Loop for each Type currently in Array
            For lCounterTypes = 1 To lTypeCount
            
                If v_bSelectedOnly Then
                    If InStr(sLstCurrentTermType, asTypes(lCounterTypes)) <> 0 _
                    Or objTerm.bSelected = False _
                    Or sLstCurrentTermType = "" Then
                        ' item already present
                        bNewTypeFound = False
                        Exit For
                    End If
                Else
                    If InStr(asTypes(lCounterTypes), sLstCurrentTermType) <> 0 _
                    Or sLstCurrentTermType = "" Then
                        ' item already present
                        bNewTypeFound = False
                        Exit For
                    End If
              End If
              
            Next lCounterTypes
            
        End If
        
        If bNewTypeFound Then
            lTypeCount = lTypeCount + 1
            ReDim Preserve asTypes(1 To lTypeCount)
            asTypes(lTypeCount) = sLstCurrentTermType
        End If
        
    Next objTerm
    
    
resumeFunc:
    asBuildTypesArray = asTypes()
    Exit Function

errorHandler:
    Call MsgBox("Error in function: " & sFUNC_NAME & vbCr & _
                " Module: " & m_sMODULE_NAME & vbCr & _
                " Spreadsheet: " & g_sPROGRAM_NAME & vbCr & _
                " Error message: " & Err.Number & ": " & Err.Description)
    GoTo resumeFunc

End Function


' ******************************************************************
' ** Name:          asAllSelectedTypesArray()
' ** Purpose:       Function to loop through each item in ListBox Terms and
' **                 create an array containing Types where each and every item is selected
' ** Returns:       Array String of Types
' ** Parameters:    None
' ** Author:        Matthew KELLY
' ** Date:          16/03/2018
' ** Revisions:
' ******************************************************************
Public Function lAllSelectedTypesArray(ByRef r_asTypesAllSelected() As String) As Long
    
    On Error GoTo errorHandler

    Const sFUNC_NAME                As String = "asAllSelectedTypesArray()"
    
    ' Varibales to be used whilst looping for each List Term (Outer loop)
    Dim lLstTermsTotCount           As Long         ' The Total List Terms Count
    Dim lLstTermsCounter            As Long         ' Counter, and
    
    Dim objTerm                     As clsTerm
    Dim sLstCurrentTermType         As String       ' Current Term Type (definition) String Text
    
    ' Varibales to be used whilst looping for each List Term (Inner loop)
    Dim lTypeCount                  As Long         ' Total number of current Types located
    Dim lCounterTypes               As Long         ' Counter for each current Types in array
    
    Dim bAllSelected                As Boolean      ' Has the Current Type already been populated into array?
    
    ' String array to be populated with each List TYPE
    Dim asTypes()                   As String
    
    ' get all current types
    asTypes() = asBuildTypesArray()

    lTypeCount = 0
    
    lLstTermsTotCount = UBound(asTypes()) - LBound(asTypes()) + 1
    For lLstTermsCounter = 1 To lLstTermsTotCount
        
        bAllSelected = True
        ' Loop for each objTerm in collection
        For Each objTerm In m_objTermCollection
            
            If InStr(asTypes(lLstTermsCounter), objTerm.sType) <> 0 _
            And objTerm.bSelected = False _
            And objTerm.sType <> "" Then
                bAllSelected = False
                Exit For
            End If
            
        Next objTerm
    
        If bAllSelected Then
            lTypeCount = lTypeCount + 1
            ReDim Preserve r_asTypesAllSelected(1 To lTypeCount)
            r_asTypesAllSelected(lTypeCount) = asTypes(lLstTermsCounter)
        End If
    
    Next lLstTermsCounter
    
resumeFunc:
    lAllSelectedTypesArray = lTypeCount
    Exit Function

errorHandler:
    Call MsgBox("Error in function: " & sFUNC_NAME & vbCr & _
                " Module: " & m_sMODULE_NAME & vbCr & _
                " Spreadsheet: " & g_sPROGRAM_NAME & vbCr & _
                " Error message: " & Err.Number & ": " & Err.Description)
    GoTo resumeFunc

End Function

