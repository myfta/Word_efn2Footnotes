Attribute VB_Name = "Efn2Footnotes"
Option Explicit

' Converts tagged text such as:
' [efn_note]This text becomes a footnote.[/efn_note]
' into a standard Word footnote reference.
'
' Designed for Microsoft Word 2010 and later.

Public Sub ConvertEfnTagsToFootnotes()

    Const OPEN_TAG As String = "[efn_note]"
    Const CLOSE_TAG As String = "[/efn_note]"

    Dim doc As Document
    Dim searchRange As Range
    Dim closeRange As Range
    Dim insertionRange As Range
    Dim noteText As String
    Dim openStart As Long
    Dim openEnd As Long
    Dim closeStart As Long
    Dim closeEnd As Long
    Dim converted As Long
    Dim missingClosingTags As Long

    Set doc = ActiveDocument

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Do
        'Always search the current main document story. The tagged block
        'is replaced as the document changes during the loop.
        Set searchRange = doc.Content

        With searchRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = OPEN_TAG
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = False
        End With

        If Not searchRange.Find.Execute Then Exit Do

        openStart = searchRange.Start
        openEnd = searchRange.End

        'Find the next closing tag after this opening tag.
        Set closeRange = doc.Range(Start:=openEnd, End:=doc.Content.End)

        With closeRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = CLOSE_TAG
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = False
        End With

        If Not closeRange.Find.Execute Then
            'Leave the unmatched opening tag in place and continue looking
            'after it, rather than deleting the rest of the document.
            missingClosingTags = missingClosingTags + 1
            Set searchRange = doc.Range(Start:=openEnd, End:=doc.Content.End)
            If Not searchRange.Find.Execute(FindText:=OPEN_TAG, _
                                            Forward:=True, _
                                            Wrap:=wdFindStop, _
                                            MatchWildcards:=False) Then
                Exit Do
            End If
            'No reliable further processing is possible if the first
            'opening tag has no closing tag; report it and stop safely.
            Exit Do
        End If

        closeStart = closeRange.Start
        closeEnd = closeRange.End

        'Read the text between the tags.
        noteText = doc.Range(Start:=openEnd, End:=closeStart).Text

        'Convert paragraph marks to spaces so tagged multi-line content does
        'not create unwanted empty paragraphs in the footnote.
        noteText = Replace(noteText, vbCr, " ")
        noteText = Replace(noteText, Chr(7), " ")
        noteText = Trim$(noteText)

        'Remove the complete tagged block.
        doc.Range(Start:=openStart, End:=closeEnd).Delete

        'Insert a standard Word footnote reference at the block's location.
        Set insertionRange = doc.Range(Start:=openStart, End:=openStart)
        doc.Footnotes.Add Range:=insertionRange, Text:=noteText

        converted = converted + 1
    Loop

    Application.ScreenUpdating = True

    If missingClosingTags > 0 Then
        MsgBox CStr(converted) & " footnote block(s) converted." & vbCrLf & _
               CStr(missingClosingTags) & " unmatched opening tag(s) found. " & _
               "The document was not changed at the unmatched tag.", _
               vbExclamation, "Conversion complete"
    Else
        MsgBox CStr(converted) & " footnote block(s) converted.", _
               vbInformation, "Conversion complete"
    End If

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "The macro stopped because of the following error:" & vbCrLf & _
           Err.Description, vbCritical, "Conversion error"

End Sub
