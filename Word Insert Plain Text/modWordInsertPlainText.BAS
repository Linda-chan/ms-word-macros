Option Explicit

'====================================================================
Sub InsertPlainText()
  On Error Resume Next
  
  Selection.PasteSpecial , , wdInLine, , wdPasteText
  
  If Err.Number = 5342 Then
    Err.Clear
    Selection.PasteSpecial , , wdInLine, , wdPasteBitmap
  End If
  
  If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
  End If
  
  'Selection.PasteAndFormat wdPasteDefault ' wdFormatPlainText
  'If Err Then MsgBox Err.Description, vbCritical
End Sub
