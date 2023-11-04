# VBA

This repo is about learning vba


```
Sub TestIfObjectExists()
  Dim i As Integer
  Dim docTables As Document
  Dim blnExists As Boolean

  For i = 1 to Documents.Count
    Set docTables = Documents(i)
    If LCase(docTables.Name) = "filename.docx" Then
      blnExists = True
      Exit For
    End If
  Next i

  If blnExists = False Then
    msgBox Prompt:="Document "filename.docx" not found."
...
```



### Add Header Text
```
Sub AddHeaderText()

With ActiveDocument.Sections(2).Headers(wdHeaderFooterPrimary)
  .LinkToPrevious = False
  .Range.Text = "Some sample text"
End With

End Sub

' Or with a specific document:

Sub AddHeaderText()

With Documents("Name of a Document.docm").Sections(2).Headers(wdHeaderFooterPrimary)
  .LinkToPrevious = False
  .Range.Text = "Some sample text"
End With

End Sub
```

### Checking if the font is Arial
```
Sub checkingFont()

Dim oPara As Paragraph

    For Each oPara In ActiveDocument.Paragraphs
        If oPara.Range.Font.Name = "Arial" Then
            MsgBox Prompt:="Yes, we are using the font called ""Arial"""
        End If
    Next

End Sub
```

Getting into vba needs a lot of training.

### Counting the characters in a paragraph
```
ActiveDocument.Paragraph(1).Range.Characters.Count
```
