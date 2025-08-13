Sub InsertHiddenKeywordsInFooters()
    Dim sec As Section
    Dim footerRange As Range
    Dim wordList As Variant
    Dim i As Integer
    Dim bgColor As Long
    Dim insertText As String
    Dim insertStart As Long
    Dim inputText As String
    Dim footerTypes As Variant
    Dim j As Integer

    ' Prompt user for keywords
    inputText = InputBox("Enter keywords separated by commas:", "Keyword Input")
    If Trim(inputText) = "" Then
        MsgBox "No keywords entered. Macro canceled.", vbExclamation
        Exit Sub
    End If

    ' Split and clean keywords
    wordList = Split(inputText, ",")
    For i = 0 To UBound(wordList)
        wordList(i) = Trim(wordList(i))
    Next i

    ' Default to white background
    bgColor = RGB(255, 255, 255)

    ' Build the hidden text block
    insertText = ""
    For i = 0 To UBound(wordList)
        insertText = insertText & wordList(i) & " "
    Next i
    insertText = Trim(insertText)

    ' Define footer types
    footerTypes = Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)

    ' Loop through each section and footer type
    For Each sec In ActiveDocument.Sections
        For j = 0 To UBound(footerTypes)
            Set footerRange = sec.Footers(footerTypes(j)).Range

            With footerRange
                insertStart = .Characters.Count + 1
                .InsertAfter vbCrLf & insertText

                With .Characters(insertStart).Duplicate
                    .MoveEnd Unit:=wdCharacter, Count:=Len(insertText)
                    .Font.Color = bgColor
                    .Font.Size = 1
                    .Font.Hidden = True
                End With
            End With
        Next j
    Next sec

    MsgBox "Hidden keywords inserted into all footers.", vbInformation
End Sub
