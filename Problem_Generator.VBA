Sub ChangeBoldToWhiteSelectedText()

    Dim s As Range
    For Each s In Selection.Characters
        If (s.Font.Bold) Then
            s.Font.Color = wdColorWhite
        End If
    Next s

End Sub

Sub UnderlineWhiteCharactersInBlack()
    Dim s As Range
    For Each s In Selection.Characters
        If s.Font.Color = wdColorWhite Then
            s.Font.UnderlineColor = wdColorBlack
            s.Font.Underline = wdUnderlineSingle
        End If
    Next s
End Sub

Sub ChangeWhiteToBlack()
    For Each c In Selection.Characters
        If c.Font.Color = wdColorWhite Then
            c.Font.Color = wdColorBlack
        End If
    Next c
End Sub

Sub FormatText()
    Selection.Font.Size = 8
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    Selection.ParagraphFormat.SpaceBefore = 0
    Selection.ParagraphFormat.SpaceAfter = 0
End Sub
            

Sub ChangeHighlightToBold_Word()
    For Each c In Selection.Characters
        If c.HighlightColorIndex <> wdNoHighlight Then
            c.HighlightColorIndex = wdNoHighlight
            c.Font.Bold = True
        End If
    Next c
End Sub


Sub MathType2Word()
    ' By Amin Yahyaabadi
    ' MathType2Word Macro: to convert Mathtype Equations to Microsoft Word Equations

    Dim found As Boolean

    Do
        With Selection.Find
            .ClearFormatting
            .Text = "$*$"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With

        found = Selection.Find.Execute

        If found Then
            Selection.Cut
            Selection.OMaths.Add Range:=Selection.Range
            Selection.Paste
            Selection.OMaths.BuildUp
        End If
    Loop While found
End Sub



