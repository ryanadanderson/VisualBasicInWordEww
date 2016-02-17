Sub AutoOpen()

Dim blnHeyLook, blnOneFound As Boolean
Dim strFind, strDocNames As String
Dim myRange As Range




For i = 1 To 1
  
    Set myRange = ActiveDocument.Content
    'If i = 1 Then strFind = "DP-155-03-000"
    'If i = 2 Then strFind = "QF-150-03-000-00"
    'If i = 3 Then strFind = "QF-155-03-907-00"
    'If i = 4 Then strFind = ""
    'If i = 5 Then strFind = ""
    'If i = 6 Then strFind = ""
    'If i = 7 Then strFind = ""
    'If i = 8 Then strFind = ""
    'If i = 9 Then strFind = ""
    'If i = 10 Then strFind = ""
    'If i = 11 Then strFind = ""
    'If i = 12 Then strFind = ""
    'If i = 13 Then strFind = ""
    'If i = 14 Then strFind = ""
    'If i = 15 Then strFind = ""
    'If i = 16 Then strFind = ""
    'If i = 17 Then strFind = ""
    'If i = 18 Then strFind = ""
    'If i = 19 Then strFind = ""
    'If i = 20 Then strFind = ""
    'If i = 21 Then strFind = ""
    'If i = 22 Then strFind = ""
    'If i = 23 Then strFind = ""
    'If i = 24 Then strFind = ""
    'If i = 25 Then strFind = ""
    'If i = 26 Then strFind = ""
    'If i = 27 Then strFind = ""
    'If i = 28 Then strFind = ""
    'If i = 29 Then strFind = ""
    'If i = 30 Then strFind = ""
    'If i = 31 Then strFind = ""
    'If i = 32 Then strFind = ""
    'If i = 33 Then strFind = ""
    'If i = 34 Then strFind = ""
    'If i = 35 Then strFind = ""
    'If i = 36 Then strFind = ""
    'If i = 37 Then strFind = ""
    'If i = 38 Then strFind = ""
    'If i = 39 Then strFind = ""
    'If i = 40 Then strFind = ""
    'If i = 41 Then strFind = ""
    'If i = 42 Then strFind = ""
    'If i = 43 Then strFind = ""
    'If i = 44 Then strFind = ""
    'If i = 45 Then strFind = ""
    'If i = 46 Then strFind = ""
    'If i = 47 Then strFind = ""
    'If i = 48 Then strFind = ""
    'If i = 49 Then strFind = ""
    'If i = 50 Then strFind = ""

    myRange.Find.Text = strFind
    If strFind <> "" Then blnHeyLook = myRange.Find.Execute



If blnHeyLook = True Then
    strDocNames = strDocNames & strFind & ", "
    blnHeyLook = False
    blnOneFound = True
    Count = Count + 1
    strFind = ""
End If

Next

'If strDocNames <> "" Then
'    strDocNames = Left(strDocNames, Len(strDocNames) - 2) & "."
'    MsgBox (i & " docs were searched. " & Count & "were found. This document contains: " & strDocNames)
'    Else
'    MsgBox (i & " docs were searched. No outdated information as of %date%.")
'End If


End Sub

