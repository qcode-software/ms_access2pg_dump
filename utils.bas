Public Sub ReplaceByRef(str As String, strFind As String, strReplace As String, start As Double, found As Boolean)
  ' Replace all occurences of strFind in str with strReplace
  Dim Position As Double
  Dim Flength As Double
  Dim Rlength As Double
  Flength = Len(strFind)
  Rlength = Len(strReplace)
  Position = InStr(start, str, strFind)
  If Position <> 0 Then
    found = True
    str = Left$(str, Position - 1) & strReplace & Right$(str, Len(str) - Position - Flength + 1)
    Call ReplaceByRef(str, strFind, strReplace, Position + Rlength, found)
  Else
    found = False
  End If
End Sub

Public Function replace(ByVal str As String, ByVal strFind As String, ByVal strReplace As String)
  ' Return string with all occurences of strFind replaces by strReplace
  Call ReplaceByRef(str, strFind, strReplace, 1, True)
  replace = str
End Function

Public Sub pushStr(arr() As String, value As String)
  ' Push string onto an variable length array
  On Error GoTo push_Err
  ReDim Preserve arr(UBound(arr) + 1)
  arr(UBound(arr)) = value

  push_Exit:
    Exit Sub
    
  push_Err:
    If Err.number = 9 Then ReDim arr(0)
    Resume Next
    
End Sub

Function strQuote(str As String, quoteChar As String)
  '-- Quote str using quoteChar escaping any quoteChar in string
  strQuote = quoteChar & replace(str, quoteChar, quoteChar & quoteChar) & quoteChar
End Function

Function strDeQuote(str As String, quoteChar As String)
  '-- Remove quoteChar quoting from a string unescaping any quoteChar in string
  Dim RE As Object
  Set RE = CreateObject("vbscript.regexp")
  Dim matches As Object
    
  RE.Pattern = "^" & quoteChar & "(.*?)" & quoteChar & "$"
  Set matches = RE.Execute(str)
    
  If (matches.count > 0) Then
    '-- Remove quoting and unescape any quoteChar in string
    strDeQuote = replace(matches.Item(0).SubMatches.Item(0), quoteChar & quoteChar, quoteChar)
  Else
    '-- String not quoted - just return original string
    strDeQuote = str
  End If
End Function
