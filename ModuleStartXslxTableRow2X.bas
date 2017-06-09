Attribute VB_Name = "ModuleStartXslxTableRow2X"
Option Explicit

Private Sub ReplacePptxShapeText(thisShape As PowerPoint.Shape, searchText As String, replaceText As String)
    On Error Resume Next
    Dim childShape As PowerPoint.Shape
    Dim replaceResult As TextRange

    If thisShape.Type = msoGroup Then
        For Each childShape In thisShape.GroupItems
            ReplacePptxShapeText childShape, searchText, replaceText
        Next
    Else
        If thisShape.HasTextFrame = msoTrue Then
            Set replaceResult = thisShape.TextFrame.TextRange
            Do While Not replaceResult Is Nothing
                Set replaceResult = thisShape.TextFrame.TextRange.Replace(searchText, replaceText)
                DoEvents
            Loop
        End If
    End If
End Sub

Private Sub ReplaceXslxShapeText(thisShape As Shape, searchText As String, replaceText As String)
    On Error Resume Next
    Dim childShape As Shape
    Dim thisText As String

    If thisShape.Type = msoGroup Then
        For Each childShape In thisShape.GroupItems
            ReplaceXslxShapeText childShape, searchText, replaceText
        Next
    Else
        thisText = thisShape.TextFrame.Characters.Text
        If Err Then
            Call Err.Clear
        Else
            thisShape.TextFrame.Characters.Text = Replace(thisText, searchText, replaceText)
        End If
    End If
End Sub


Sub StartXslxTableRow2PptxSlide()
    Dim thisRange As Range
    Dim thisListObject As ListObject
    Dim thisListRow As ListRow
    Dim thisListRowRange As Range
    Dim thisListColumn As ListColumn
    Dim thisListColumnRange As Range
    Dim thisListHeaderRange As Range
    Dim thisKey As String
    Dim thisValue As String
    Dim powerPointApplication As PowerPoint.Application
    Dim powerPointPresentation As PowerPoint.Presentation
    Dim templateSlide As PowerPoint.Slide
    Dim thisSlideRange As PowerPoint.SlideRange
    Dim thisSlide As PowerPoint.Slide
    Dim thisShape As PowerPoint.Shape
    Dim replaceResult As TextRange

    On Error Resume Next
    For Each thisListObject In ActiveWorkbook.ActiveSheet.ListObjects
        Set thisRange = Intersect(thisListObject.Range, Application.Selection)
        If Application.Selection.Address = thisRange.Address Then
            Set thisListHeaderRange = thisListObject.HeaderRowRange
            Set powerPointApplication = New PowerPoint.Application
            Set powerPointPresentation = powerPointApplication.Presentations.Open(fileName:=Application.GetOpenFilename(), ReadOnly:
=msoTrue)
            Set templateSlide = powerPointPresentation.Slides(1)

            For Each thisListRow In thisListObject.ListRows
                Set thisListRowRange = thisListRow.Range
                If Not thisListRowRange.EntireRow.Hidden Then
                    Set thisSlideRange = templateSlide.Duplicate
                    thisSlideRange.MoveTo powerPointPresentation.Slides.Count
                    Set thisSlide = thisSlideRange(1)
                    thisSlide.Select
                    For Each thisListColumn In thisListObject.ListColumns
                        Set thisListColumnRange = thisListColumn.Range
                        If Not thisListColumnRange.EntireColumn.Hidden Then
                            thisKey = Intersect(thisListColumnRange, thisListHeaderRange).Value
                            thisValue = Intersect(thisListColumnRange, thisListRowRange).Value
                            For Each thisShape In thisSlide.Shapes
                                ReplacePptxShapeText thisShape, thisKey, thisValue
                                DoEvents
                            Next
                        End If
                        DoEvents
                    Next
                End If
                DoEvents
            Next

            templateSlide.Select
            templateSlide.Delete
        End If
        DoEvents
    Next
End Sub

Sub StartXslxTableRow2XslxSheet()
    Dim thisRange As Range
    Dim thisListObject As ListObject
    Dim thisListRow As ListRow
    Dim thisListRowRange As Range
    Dim thisListColumn As ListColumn
    Dim thisListColumnRange As Range
    Dim thisListHeaderRange As Range
    Dim thisKey As String
    Dim thisValue As String
    Dim targetWorkbook As Workbook
    Dim templateSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim thisShape As Shape
    Dim replaceResult As TextRange

    On Error Resume Next
    For Each thisListObject In ActiveWorkbook.ActiveSheet.ListObjects
        Set thisRange = Intersect(thisListObject.Range, Application.Selection)
        If Application.Selection.Address = thisRange.Address Then
            Set thisListHeaderRange = thisListObject.HeaderRowRange
            Set targetWorkbook = Workbooks.Open(fileName:=Application.GetOpenFilename(), ReadOnly:=msoTrue)
            Set templateSheet = targetWorkbook.Sheets(1)

            For Each thisListRow In thisListObject.ListRows
                Set thisListRowRange = thisListRow.Range
                If Not thisListRowRange.EntireRow.Hidden Then
                    Dim targetSheetTitleChanged As Boolean
                    templateSheet.Copy Before:=templateSheet
                    Set targetSheet = targetWorkbook.ActiveSheet
                    targetSheetTitleChanged = False

                    For Each thisListColumn In thisListObject.ListColumns
                        Set thisListColumnRange = thisListColumn.Range
                        If Not thisListColumnRange.EntireColumn.Hidden Then
                            thisKey = Intersect(thisListColumnRange, thisListHeaderRange).Value
                            thisValue = Intersect(thisListColumnRange, thisListRowRange).Value

                            If Not targetSheetTitleChanged Then
                                targetSheet.Name = thisValue
                                targetSheetTitleChanged = True
                            End If

                            targetSheet.Cells.Replace What:=thisKey, replacement:=thisValue, _
                                LookAt:=xlPart, SearchOrder:=xlByRows, _
                                MatchCase:=True, MatchByte:=True, _
                                SearchFormat:=False, ReplaceFormat:=False

                            For Each thisShape In targetSheet.Shapes
                                ReplaceXslxShapeText thisShape, thisKey, thisValue
                                DoEvents
                            Next
                        End If
                        DoEvents
                    Next
                End If
                DoEvents
            Next

            targetWorkbook.Worksheets(1).Activate
            Application.DisplayAlerts = False
            templateSheet.Delete
            Application.DisplayAlerts = True
        End If
        DoEvents
    Next
End Sub
