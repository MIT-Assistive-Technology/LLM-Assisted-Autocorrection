Option Explicit

Sub EnterAndCorrectLastParagraph()

    Dim finishedParagraph As Range
    Dim endOfParagraph As Range

    If Selection.Type <> wdSelectionIP Then
        Selection.TypeParagraph
        Exit Sub
    End If

    Set finishedParagraph = Selection.Paragraphs(1).Range.Duplicate
    finishedParagraph.End = finishedParagraph.End - 1   'Exclude paragraph mark

    CorrectUnknownWords finishedParagraph

    'Reset the cursor to the actual end of the corrected paragraph.
    Set endOfParagraph = finishedParagraph.Paragraphs(1).Range.Duplicate
    endOfParagraph.End = endOfParagraph.End - 1
    endOfParagraph.Collapse wdCollapseEnd
    endOfParagraph.Select

    Selection.TypeParagraph

End Sub

Private Sub CorrectUnknownWords(ByVal target As Range)

    Dim i As Long
    Dim wordRange As Range
    Dim originalText As String
    Dim wordText As String
    Dim trailingText As String
    Dim suggestions As SpellingSuggestions

    For i = target.Words.Count To 1 Step -1

        Set wordRange = target.Words(i).Duplicate
        originalText = wordRange.Text
        wordText = Trim$(originalText)

        'Preserve the space that Word includes after a word.
        trailingText = Mid$(originalText, Len(wordText) + 1)

        If Len(wordText) > 0 Then

            If Not Application.CheckSpelling(Word:=wordText) Then

                Set suggestions = _
                    Application.GetSpellingSuggestions(Word:=wordText)

                If suggestions.Count > 0 Then
                    wordRange.Text = suggestions(1).Name & trailingText
                End If

            End If

        End If

    Next i

End Sub

Sub InstallEnterKey()

    Dim kb As KeyBinding

    CustomizationContext = ActiveDocument

    Set kb = KeyBindings.Key(BuildKeyCode(wdKeyReturn))
    If Not kb Is Nothing Then kb.Clear

    Set kb = KeyBindings.Add( _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="EnterAndCorrectLastParagraph", _
        KeyCode:=BuildKeyCode(wdKeyReturn))

    MsgBox "Installed: Enter now runs " & kb.Command

End Sub