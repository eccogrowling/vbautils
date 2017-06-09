Attribute VB_Name = "Quotation"
Sub Quote()
    Debug.Print "Quote"
    Dim Inspector As Inspector
    Dim Document As Word.Document
    Dim Selection As Word.Selection

    Set Inspector = Application.ActiveInspector
    If Inspector.IsWordMail And Inspector.EditorType = olEditorWord Then
        Set Document = Inspector.WordEditor

        If Document.ProtectionType <> wdNoProtection Then
            Call Beep
        Else
            Set Selection = Document.Application.Selection
            Call Selection.Expand(wdLine)

            With New RegExp
                Let .Pattern = "^(|(?!$).*?)$(?=.+)"
                Let .IgnoreCase = False
                Let .Global = True
                Let .MultiLine = True
                Let Selection.Text = .Replace(Selection.Text, "> $1")
            End With
        End If
    End If
End Sub

Sub Unquote()
    Debug.Print "Unquote"
    Dim Inspector As Inspector
    Dim Document As Word.Document
    Dim Selection As Word.Selection

    Set Inspector = Application.ActiveInspector
    If Inspector.IsWordMail And Inspector.EditorType = olEditorWord Then
        Set Document = Inspector.WordEditor

        If Document.ProtectionType <> wdNoProtection Then
            Call Beep
        Else
            Set Selection = Document.Application.Selection
            Call Selection.Expand(wdLine)

            With New RegExp
                Let .Pattern = "^> (|(?!$).*?)$(?=.+)"
                Let .IgnoreCase = False
                Let .Global = True
                Let .MultiLine = True
                Let Selection.Text = .Replace(Selection.Text, "$1")
            End With
        End If
    End If
End Sub
