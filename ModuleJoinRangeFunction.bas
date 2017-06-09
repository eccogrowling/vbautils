Attribute VB_Name = "ModuleJoinRangeFunction"
Option Explicit

Private Function RangeWithoutBlank(targetRange As Range) As Range
    Dim resultRange As Range
    Dim targetCell
    If Not targetRange Is Nothing Then
        For Each targetCell In targetRange
            If WorksheetFunction.CountBlank(targetCell) = 0 Then
                If resultRange Is Nothing Then
                    Set resultRange = targetRange.Worksheet.Range(targetCell.address)
                Else
                    Set resultRange = Union(resultRange, targetCell)
                End If
            End If
        Next
    End If
    Set RangeWithoutBlank = resultRange
End Function

Function JoinRange(targetRange As Range, Optional ByVal delimiter As String = "") As String
    Dim resultValue As String
    Dim targetCell
    resultValue = vbNullString

    If Not targetRange Is Nothing Then
        For Each targetCell In targetRange
            If resultValue = vbNullString Then
                resultValue = targetCell.Text
            Else
                resultValue = resultValue & delimiter & targetCell.Text
            End If
        Next
    End If

    JoinRange = resultValue
End Function

Function JoinRangeLn(targetRange As Range) As String
    JoinRangeLn = JoinRange(targetRange, vbLf)
End Function

Function JoinRangeWithoutBlank(targetRange As Range, Optional ByVal delimiter As String = "") As String
    JoinRangeWithoutBlank = JoinRange(RangeWithoutBlank(targetRange), delimiter)
End Function

Function JoinRangeWithoutBlankLn(targetRange As Range) As String
    JoinRangeWithoutBlankLn = JoinRangeLn(RangeWithoutBlank(targetRange))
End Function
