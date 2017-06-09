Attribute VB_Name = "ModuleRegExpFunction"
Option Explicit

Private Function GetCachedRegExp(ByVal patternText As String, _
Optional ByVal ignoreCaseOption As Boolean = False, Optional ByVal globalOption As Boolean = False, Optional ByVal multiLineOption As Boolean = False) _
As Variant
    Static patternCaches As Variant
    Dim optionsCaches As Variant
    Dim optionsKey As String
    Dim resultRegExp As Variant

    optionsKey = "ignoreCaseOption" & ignoreCaseOption & "globalOption" & globalOption & "multiLineOption" & multiLineOption
    If IsEmpty(patternCaches) Then
        Set patternCaches = CreateObject("Scripting.Dictionary")
    End If
    If Not patternCaches.Exists(patternText) Then
        Set patternCaches.Item(patternText) = CreateObject("Scripting.Dictionary")
    End If
    Set optionsCaches = patternCaches.Item(patternText)
    If Not optionsCaches.Exists(optionsKey) Then
        Set resultRegExp = CreateObject("VBScript.RegExp")
        Set optionsCaches.Item(optionsKey) = resultRegExp
        With resultRegExp
            .Pattern = patternText
            .IgnoreCase = ignoreCaseOption
            .Global = globalOption
            .MultiLine = multiLineOption
        End With
    Else
        Set resultRegExp = optionsCaches.Item(optionsKey)
    End If

    Set GetCachedRegExp = resultRegExp
End Function

Private Function RegExpExecute(ByVal expression As String, ByVal patternText As String, _
Optional ByVal ignoreCaseOption As Boolean = False, Optional ByVal globalOption As Boolean = False, Optional ByVal multiLineOption As Boolean = False) _
As Variant
    Dim matches As Variant

    With GetCachedRegExp(patternText, ignoreCaseOption, globalOption, multiLineOption)
        Set matches = .Execute(expression)
    End With

    Set RegExpExecute = matches
End Function

Private Function RegExpExecuteGlobalMultiLine(ByVal expression As String, ByVal patternText As String, _
Optional ByVal ignoreCaseOption As Boolean = False) _
As Variant
    Set RegExpExecuteGlobalMultiLine = RegExpExecute(expression, patternText, ignoreCaseOption, True, True)
End Function

Function RegExpFirstSubMatchValue(ByVal expression As String, ByVal patternText As String, _
Optional ByVal ignoreCaseOption As Boolean = False, Optional ByVal globalOption As Boolean = False, Optional ByVal multiLineOption As Boolean = False) _
As String
    Dim resultSubMatchString As String
    Dim matchItem
    Dim subMatchItem

    For Each matchItem In RegExpExecute(expression, patternText, ignoreCaseOption, globalOption, multiLineOption)
        For Each subMatchItem In matchItem.SubMatches()
            resultSubMatchString = subMatchItem
            Exit For
        Next
        Exit For
    Next

    RegExpFirstSubMatchValue = resultSubMatchString
End Function

Function RegExpFirstSubMatchValueGlobalMultiLine(ByVal expression As String, ByVal patternText As String, _
Optional ByVal ignoreCaseOption As Boolean = False) _
As String
    RegExpFirstSubMatchValueGlobalMultiLine = RegExpFirstSubMatchValue(expression, patternText, ignoreCaseOption, True, True)
End Function

Function RegExpTest(ByVal expression As String, ByVal patternText As String, _
Optional ByVal ignoreCaseOption As Boolean = False, Optional ByVal globalOption As Boolean = False, Optional ByVal multiLineOption As Boolean = False) _
As Boolean
    With GetCachedRegExp(patternText, ignoreCaseOption, globalOption, multiLineOption)
        RegExpTest = .test(expression)
    End With
End Function

Function RegExpReplace(ByVal expression As String, ByVal patternText As String, ByVal replacement As String, _
Optional ByVal ignoreCaseOption As Boolean = False, Optional ByVal globalOption As Boolean = False, Optional ByVal multiLineOption As Boolean = False) _
As String
    With GetCachedRegExp(patternText, ignoreCaseOption, globalOption, multiLineOption)
        RegExpReplace = .Replace(expression, replacement)
    End With
End Function
