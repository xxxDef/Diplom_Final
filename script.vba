Sub ReplaceBracketCitationsWithCrossRefsAndKeepBrackets()

    Dim doc As Document
    Dim rngSearch As Range
    Dim rngFound As Range
    Dim findPattern As String
    Dim originalFoundText As String
    Dim citationContent As String
    Dim commaPosition As Long
    Dim refNumStr As String
    Dim refNum As Long
    Dim pageInfo As String
    Dim replacementsCount As Long
    Dim errorsCount As Long

    ' --- Settings ---
    ' Search pattern: \[ (opening bracket)
    ' [0-9]@ (one or more digits - reference number)
    ' * (any characters zero or more times - for pages, etc.)
    ' \] (closing bracket)
    findPattern = "\[[0-9]{1,3}"
    ' -----------------

    On Error GoTo ErrorHandler

    Set doc = ActiveDocument
    Set rngSearch = doc.Content ' Search the entire document

    Application.ScreenUpdating = False
    Application.StatusBar = "Searching for and replacing citations..."

    replacementsCount = 0
    errorsCount = 0

    With rngSearch.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findPattern
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop ' Stop at the end of the document

        Do While .Execute
            ' rngSearch now points to the found text
            Set rngFound = rngSearch.Duplicate ' Work with a copy of the found range
            
            If rngFound.Characters.Count >= 2 Then
                rngFound.MoveStart Unit:=wdCharacter, Count:=1
                ' Òåïåðü rngFound.Text ýòî, íàïðèìåð, "2" èëè "2, pp. 120-130"
            Else
                ' Åñëè rngFound áûë, íàïðèìåð, "[]", îí ñòàíåò ïóñòûì.
                ' Åñëè îí áûë "[" èëè "", îí íå èçìåíèòñÿ ýòèì áëîêîì,
                ' íî ñëåäóþùàÿ ïðîâåðêà íà Len(originalFoundText) > 2 åãî îòëîâèò
                ' èëè ìîæíî äîáàâèòü çäåñü GoTo NextIteration, åñëè ýòî îøèáêà.
                Debug.Print "Íàéäåííûé òåêñò ñëèøêîì êîðîòêèé äëÿ èçâëå÷åíèÿ ñîäåðæèìîãî: " & originalFoundText
                ' Åñëè ýòî êðèòè÷íî è òàêîé òåêñò íå äîëæåí îáðàáàòûâàòüñÿ äàëüøå:
                'rngSearch.Collapse wdCollapseEnd
                GoTo NextIteration ' (íî óáåäèòåñü, ÷òî rngSearch.Collapse wdCollapseEnd áóäåò âûçâàí)
            End If

            originalFoundText = rngFound.Text ' e.g., "[2]" or "[2, pp. 120-130]"
            citationContent = originalFoundText
            

            pageInfo = "" ' Reset page information
            commaPosition = InStr(citationContent, ",")

            If commaPosition > 0 Then
                refNumStr = Trim(Left(citationContent, commaPosition - 1))
                pageInfo = Trim(Mid(citationContent, commaPosition)) ' Includes the comma, e.g., ", pp. 120-130"
            Else
                refNumStr = Trim(citationContent)
            End If

            If IsNumeric(refNumStr) Then
                refNum = CLng(refNumStr)

                ' Store the original range to potentially restore text if cross-ref fails
                Dim originalRangeForRestore As Range
                Set originalRangeForRestore = rngFound.Duplicate

                ' Clear the original bracketed text
                rngFound.Text = "" ' Clears the content, rngFound is now a collapsed range at the start

                ' --- Start building the new content ---
                'rngFound.Text = "[" ' Insert opening bracket
                'rngFound.Collapse wdCollapseEnd ' Move cursor after "["

                ' Try to insert the cross-reference
                On Error Resume Next ' Enable error handling for this specific operation
                rngFound.InsertCrossReference _
                    ReferenceType:="Numbered item", _
                    ReferenceKind:=wdNumberNoContext, _
                    ReferenceItem:=refNum, _
                    InsertAsHyperlink:=True, _
                    IncludePosition:=False
                
                If Err.Number <> 0 Then
                    ' Error inserting cross-reference (e.g., number not found)
                    Debug.Print "Error inserting cross-reference for [" & refNumStr & "]: " & Err.Description & ". Restoring original: " & originalFoundText
                    originalRangeForRestore.Text = originalFoundText ' Restore the original text
                    errorsCount = errorsCount + 1
                    Err.Clear
                Else
                    ' Cross-reference inserted successfully
                    rngFound.Collapse wdCollapseEnd ' Move cursor after the inserted cross-reference

                    ' Add page information if it exists
                    If pageInfo <> "" Then
                        rngFound.InsertAfter pageInfo
                        rngFound.Collapse wdCollapseEnd ' Move cursor after page info
                    End If

                    'rngFound.InsertAfter "]" ' Insert closing bracket
                    replacementsCount = replacementsCount + 1
                End If
                On Error GoTo ErrorHandler ' Restore standard error handling
                
            Else
                Debug.Print "Skipped (non-numeric reference number): " & originalFoundText
                errorsCount = errorsCount + 1 ' Count as an error/skipped item
            End If
NextIteration:
            ' Crucial: Collapse the main search range to its end to continue searching AFTER the processed/skipped item
            rngSearch.Collapse wdCollapseEnd
        Loop
    End With

    Application.StatusBar = "Processing complete!"
    Application.ScreenUpdating = True

    MsgBox "Processing finished." & vbCrLf & _
           "Citations replaced: " & replacementsCount & vbCrLf & _
           "Skipped/Errors: " & errorsCount, vbInformation

    Exit Sub

ErrorHandler:
    Application.StatusBar = "An error occurred!"
    Application.ScreenUpdating = True
    MsgBox "A VBA error occurred: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

