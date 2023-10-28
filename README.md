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
