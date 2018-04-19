Attribute VB_Name = "modGlossaryCreate"
Sub g_CreateGlossaryFromCollection(ByRef r_objCollectionTerms As Collection, ByVal v_sHeading As String, ByRef r_asTypes() As String)
       
    Call AddHeading(v_sHeading)
       
    Call AddTableTerms(r_objCollectionTerms, r_asTypes)
       
End Sub

Private Sub AddHeading(ByVal v_sHeading As String)

    Selection.GoTo What:=wdGoToBookmark, Name:="Heading2"
    With ActiveDocument.Bookmarks
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    Selection.Style = ActiveDocument.Styles("Heading 2")
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:=v_sHeading
   
End Sub


Private Sub AddSubHeadingTo(ByVal v_sSubHeading As String, ByVal v_sBookMark As String)

    Selection.GoTo What:=wdGoToBookmark, Name:=v_sBookMark
    With ActiveDocument.Bookmarks
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    Selection.Style = ActiveDocument.Styles("Heading 3")
    Selection.TypeText Text:=v_sSubHeading + vbCr
    
End Sub

Sub Edit()

    Call MsgBox("hello!", vbCritical + vbExclamation)
    
End Sub


Private Sub AddTableTerms(ByRef r_objCollectionTerms As Collection, ByRef r_asTypes() As String)
Attribute AddTableTerms.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Table"

    Dim lLowerBoundTypes        As Long
    Dim lUpperBoundTypes        As Long
    Dim lCounter                As Long

    Dim bContentCreated         As Boolean
    
    lLowerBoundTypes = LBound(r_asTypes)
    lUpperBoundTypes = UBound(r_asTypes)
    
    Call DropDownBookMarkAndSelectCurrent("End")
    
    For lCounter = lLowerBoundTypes To lUpperBoundTypes
         
        If bContentCreated Then
            Call DropDownBookMarkAndSelectCurrent("End")
        End If
        bContentCreated = False
        
        For Each objTerm In r_objCollectionTerms
        
            If InStr(r_asTypes(lCounter), objTerm.sType) <> 0 _
            And objTerm.sType <> "" _
            Or r_asTypes(lCounter) = "" Then
            
                If Not bContentCreated Then
                
                    Call ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:= _
                                            3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
                                            wdAutoFitFixed)
                    With Selection.Tables(1)
                        If .Style <> "Table Grid" Then
                            .Style = "Table Grid"
                        End If
                        .ApplyStyleHeadingRows = True
                        .ApplyStyleLastRow = False
                        .ApplyStyleFirstColumn = True
                        .ApplyStyleLastColumn = False
                        .ApplyStyleRowBands = True
                        .ApplyStyleColumnBands = False
                        .Columns(1).SetWidth ColumnWidth:=120.25, RulerStyle:=wdAdjustFirstColumn
                        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
                        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
                        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
                        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
                        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
                    End With
                End If
                
                bContentCreated = True
                
                Selection.Text = objTerm.sTerm
                Selection.MoveRight Unit:=wdCell
                Selection.Text = objTerm.sDefinition + vbCr
                Selection.MoveRight Unit:=wdCell
                ' MTK 22032018 - Alternative terms testing
                Selection.Text = "Alternative terms"
                Selection.MoveRight Unit:=wdCell
                
            End If
            
        Next objTerm
    
        If bContentCreated And r_asTypes(lCounter) <> "" Then
            Call AddSubHeadingTo(r_asTypes(lCounter), "Start")
        End If
    
    Next lCounter
    
End Sub


Private Sub DropDownBookMarkAndSelectCurrent(ByVal v_sBookMarkName As String)

    ' go to start
    Selection.GoTo What:=wdGoToBookmark, Name:=v_sBookMarkName
    With ActiveDocument.Bookmarks
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    ActiveDocument.Bookmarks(v_sBookMarkName).Delete
    
    ' create Start
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Start"
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    
    ' add paragraph. i.e. Carriage line feed, return
    Selection.TypeParagraph

    ' create Current
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Current"
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    
    ' add paragraph. i.e. Carriage line feed, return
    Selection.TypeParagraph
    
     'create new reference to End
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="End"
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    
    ' go to current
    Selection.GoTo What:=wdGoToBookmark, Name:="Current"

End Sub
