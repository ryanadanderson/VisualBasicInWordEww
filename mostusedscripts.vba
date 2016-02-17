'might have made a mistake 2016-02-16:
    'changed "REFERENCE" to "REFERENCED DOCUMENTS"
            'issues - could mess up old McGill document tranfers
            'issues - could mess up copy between text, added 10 characters including space
            'issues - now is searching fo "REFERENCE DOCUMENTS" instead of "REFERENCE"
        'shoot...might need to fix that
Public strPrimaryDocLong


Sub StartUp()

'this sets the save to location


'set to "" if all locations are teh same
ALLPlaces = "C:\Users\raanderson\Desktop\Script1\Scripted\"
If Right(ALLPlaces, 1) <> "\" Then ALLPlaces = ALLPlaces & "\"


MSPF_SaveToLocation = "C:\Users\raanderson\Desktop\Script1\Scripted\"
If Right(MSPF_SaveToLocation, 1) <> "\" Then MSPF_SaveToLocation = MSPF_SaveToLocation & "\"

SaveToLocation = "C:\Users\raanderson\Desktop\Jared Email\fixed\"
If Right(SaveToLocation, 1) <> "\" Then SaveToLocation = SaveToLocation & "\"

ECNSaveToLocation = "C:\Users\raanderson\Desktop\Jared Email\fixed\"
If Right(ECNSaveToLocation, 1) <> "\" Then ECNSaveToLocation = ECNSaveToLocation & "\"


'Will overwright all locations to make it easier if need be
If ALLPlaces <> "" Then
    ECNSaveToLocation = ALLPlaces
    SaveToLocation = ALLPlaces
    ECNSaveToLocation = ALLPlaces
End If

End Sub

Sub BringToNewDoc()
'
' BringToNewDoc Macro
'

Call StartUp
Dim MyData As DataObject

'document/procedure title and description
Dim strPara
Dim subject As String

Dim blnFound As Boolean
Dim rng1 As Range
Dim rng2 As Range
Dim rngfound As Range
Dim strTheText As String 'string? or range?

    'goes to beginning of document
        
        Selection.WholeStory
        Selection.Font.Size = 10
        Selection.Font.Name = "Arial"
        
        Selection.HomeKey Unit:=wdStory
        
    'gets title of the document for TITLE portion
        Set MyData = New DataObject
        strMainDocLong = ActiveDocument.Name
        'strMainDoc = Left(strMainDocLong, Len(strMainDocLong) - 4) 'variable change for docx doc length document name
        
        strMainDoc = Left(strMainDocLong, 13) 'change back, one time use, onetime, fix,
    '/gets...TITLE portion
    
    'Find the subject of the document
        Selection.HomeKey Unit:=wdStory
        Selection.Find.ClearFormatting
        Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=4, Name:=""
        Selection.StartOf Unit:=wdParagraph
        Selection.MoveEnd Unit:=wdParagraph
        Selection.Copy
        MyData.GetFromClipboard
        strClipSubject = MyData.GetText
        'MsgBox (strClipSubject)
        If strClipSubject = "" Then
            Selection.HomeKey Unit:=wdStory
            Selection.Find.ClearFormatting
            Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=5, Name:=""
            Selection.StartOf Unit:=wdParagraph
            Selection.MoveEnd Unit:=wdParagraph
            Selection.Copy
            MyData.GetFromClipboard
            strClipSubject = MyData.GetText
        End If
              
    'opens template '2015-03-16
    Documents.Add Template:= _
        "G:\Common\Jerry\Controlled Documents\Document Conversion\Template QA14.dotx" _
        , NewTemplate:=False, DocumentType:=0
    '/open template
    
    With Dialogs(wdDialogFileSummaryInfo) 'makes title of new doc to the controlled doc title
     .Title = strMainDoc
     .subject = strClipSubject
     .Execute
    End With
    '****select title and change to black color

    strNewDoc = ActiveDocument.Name 'variable for the newer document, will be a .docx
    Documents(strMainDocLong).Activate
    
    'FIND IT - finds paragraphs
    Application.ScreenUpdating = False
    Selection.HomeKey wdStory
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.Find.Text = "Purpose"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "Revisions:"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start - 4, rng2.Start - 5)
                rngfound.Select
                Selection.Copy
            
        End If
    End If

        'insert text here
        Documents(strNewDoc).Activate
        'Selection.GoTo What:=wdGoToBookmark, Name:="BeginningOfDocument"
        Selection.PasteAndFormat (wdPasteDefault) 'wdFormatSurroundingFormattingWithEmphasis
    
    Application.ScreenUpdating = True
    Windows(strNewDoc).Activate
    
    Call FormatTitlesOldDocToNew
    
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Copy
    
    Call SaveDocument

    ActiveWindow.ActivePane.LargeScroll down:=1
    
    strNewDocLong = strMainDoc & ".docx"
    
    Call NewParagraphFixer
    
    Call Tabs
    Call ApplyHeadingStyles
     
    Call FindAndReplaceDoubleSpaces
          
    Call RevisionsTableCount
    'Call AddParagrahToTitles
    Call deleteEmptyParagraphs
    Call FixParagraphIndents
    Call EndOfOldEmersonForamtDocument
    
    ActiveDocument.Save
    
End Sub

Sub OldTitleAndSubject()
'
' OldTitleAndSubject Macro
' puts subject into new document and description
' make it save the document as the new title
'

Dim MyData As DataObject
Dim strClipSubject As String


    strPrimaryDoc = ActiveDocument.Name
    strPrimaryDocLong = ActiveDocument.Name
 
    Call StartUp
    Call NumbersToManualNumbers 'could change font to 10.
    Call TemplateOpen
    
    strNewDocLong = ActiveDocument.Name

        
    Documents(strPrimaryDoc).Activate
    
        'gets title of the document for TITLE portion
        Set MyData = New DataObject
        strNewDoc = Left(strPrimaryDocLong, Len(strPrimaryDocLong) - 4)
        strDocType = Left(strNewDoc, Len(strNewDoc) - (Len(strNewDoc) - 2))
        
        If strDocType = "WI" Then
        
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                With Selection.Find
                    .Text = "CHARACTERISTICS & MEASURING DEVICES"
                    .Replacement.Text = "CHARACTERISTICS AND MEASURING DEVICES"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
        End If

    '/gets...TITLE portion
    
    'open document varliable
    Selection.HomeKey Unit:=wdStory
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    
    'just changed here...maybe a goto in a table
    Selection.Copy
        MyData.GetFromClipboard
        strClipSubject = MyData.GetText
        'MsgBox (strClipSubject)
      
    'change to string from the opened template
    Documents(strNewDocLong).Activate
    With Dialogs(wdDialogFileSummaryInfo) 'makes title of new doc to the controlled doc title
     .Title = strClipSubject
     .subject = strNewDoc
    ' .ECN = "1234"
     .Execute
    End With
    
    'saving the Template to proper name
    strMainDoc = strNewDoc
    Call SaveDocument
    
    
''''''''''''''''''''''''''
' Before this comment this does the following:
' 1) opens a template .docx
' 2) gets the title and description from established doc (subject in MSWord form)
' 3) decides if it's a WI, and does some modifications if it is
' 4) saves the document at the above location
''''''''''''''''''''''''''
strNewDocLong = strNewDoc & ".docx"

'have these in order from farthest down template document to the top.
'based on line in each of the subs and this changes if you start at the top

Call Reference    'Reference
Call ProcedureBody   'Procedure Body
Call TermsAndDefinitions   'Terms and Definitions
Call Scope      'Scope
Call Purpose    'Purpose

Documents(strNewDocLong).Activate
Call ApplyHeadingStyles 'adjusts the numbering of the automatic numbering of document.
Call CapsForSetUpAndCharacteristics 'looks for titles of sections and makes them all caps
Call RevisionsTableCount 'changes the cell that cantains 00 to a numbered format
Call FindAndReplaceDoubleSpaces
Call deleteEmptyParagraphs
Call FixParagraphIndents
ActiveDocument.Save

End Sub

Sub NumbersToManualNumbers()
'used in OldTitleAndSubject()
'
'    'changes all font size to 10. Might be a mistake...
    
    'the below code replaces manual page breaks with nothing
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^m" 'page break code
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p^t" 'page break code
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^t" 'page break code
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    'the below code replaces 2 'enters' (paragraph symbols) and replces it with one
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p" 'two paragraph
        .Replacement.Text = "^p" 'single paragraph symbol
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.WholeStory
    Selection.Font.Size = 10
    Selection.Font.Name = "Arial"

    ActiveDocument.Range.ListFormat.ConvertNumbersToText

End Sub

Sub NewNumbersToManualNumbers()
'used in BringToNewDoc()
'
'    'changes all font size to 10. Might be a mistake...
    
    'the below code replaces manual page breaks with nothing
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^m" 'page break code
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^t" 'page break code
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.WholeStory
    Selection.Font.Size = 10
    Selection.Font.Name = "Arial"

    ActiveDocument.Range.ListFormat.ConvertNumbersToText

End Sub

Sub ApplyMultiLevelHeadingNumbers()
 Dim LT As ListTemplate, i As Long
 Set LT = ActiveDocument.ListTemplates.Add(OutlineNumbered:=True)
 For i = 1 To 9
   With LT.ListLevels(i)
     .NumberFormat = Choose(i, "%1", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4", "%1.%2.%3.%4.%5", "%1.%2.%3.%4.%5.%6", "%1.%2.%3.%4.%5.%6.%7", "%1.%2.%3.%4.%5.%6.%7.%8", "%1.%2.%3.%4.%5.%6.%7.%8.%9")
     .TrailingCharacter = wdTrailingTab
     .NumberStyle = wdListNumberStyleArabic
     .NumberPosition = CentimetersToPoints(0)
     .Alignment = wdListLevelAlignLeft
     .TextPosition = CentimetersToPoints(0.5 + i * 0.5)
     .ResetOnHigher = True
     .StartAt = 1
     .LinkedStyle = "Heading " & i
   End With
 Next
 End Sub

Sub ApplyHeadingStyles()

    Dim i As Integer
    Dim Para As Paragraph, Rng As Range, iLvl As Long
    Dim blnFound As Boolean
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rngfound As Range
       
    'put in range from Purpose: to Revisions:
    Selection.HomeKey wdStory
    
    'select all
    'numberthing
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    
    Selection.Find.Text = "PURPOSE:" 'can change to different title to change different areas
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "Revisions:"
        blnFound = Selection.Find.Execute
        
        If blnFound = False Then 'test
            Selection.Find.Text = "REVISIONS:" 'test
            blnFound = Selection.Find.Execute
         End If 'test
        
       
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start - 4, rng2.Start)
            rngfound.Select
        End If
      
    End If

If rngfound Is Nothing Then
    MsgBox ("highlight text")
    Set rngfound = Selection.Range
End If

With rngfound 'ActiveDocument.Range
    For Each Para In .Paragraphs
     If Para.Range.Information(wdWithInTable) = False Then
     Set Rng = Para.Range.Words.First
     With Rng
       If IsNumeric(.Text) Then
         While .Characters.Last.Next.Text Like "[0-9. " & vbTab & "]"
           .End = .End + 1
         Wend
         iLvl = UBound(Split(.Text, "."))
         If IsNumeric(Split(.Text, ".")(UBound(Split(.Text, ".")))) Then iLvl = iLvl + 1
         If iLvl < 10 Then
           If iLvl < 1 Then iLvl = 1
           .Text = vbNullString
           Para.Style = "Heading " & iLvl
         End If
         End If
     End With
     End If
   Next
End With
  
Call FixLevel1Headings

 
 
 End Sub

Sub TemplateOpen()
' called by OldTitleAndSubject
' TemplateOpen Macro
' Opens the "Template QA10.dotx"
'
    Documents.Add Template:= _
        "G:\Common\Jerry\Controlled Documents\Document Conversion\Template QA14.dotx" _
        , NewTemplate:=False, DocumentType:=0
   
    
End Sub

Sub TermsAndDefinitions()
'
'Finds the Definitions in the older McGill format and slapps it in the
'Terms and Definitions the new Emerson format. Let's do this!
'

Dim blnFound As Boolean
Dim rng1 As Range
Dim rng2 As Range
Dim rngfound As Range
'Dim strTheTermsAndDefsText As String


Documents(strPrimaryDocLong).Activate

    Application.ScreenUpdating = False
    Selection.HomeKey wdStory
    Selection.Find.Text = "DEFINITIONS"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "PROCEDURE STATEMENTS"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start + 12, rng2.Start - 5) '-2
            rngfound.Select
            Selection.Copy
        End If
    End If

If blnFound = True Then

    'insert text here
    Documents(strNewDocLong).Activate
    
    Selection.GoTo What:=wdGoToBookmark, Name:="TermsAndDefinitions"
    Selection.PasteAndFormat (wdPasteDefault)

End If


End Sub
Sub MSPF_TermsAndDefinitions()
'
'Finds the Definitions in the older McGill format and slapps it in the
'Terms and Definitions the new Emerson format. Let's do this!
'

Dim blnFound As Boolean
Dim rng1 As Range
Dim rng2 As Range
Dim rngfound As Range
'Dim strTheTermsAndDefsText As String


Documents(strPrimaryDocLong).Activate

    Application.ScreenUpdating = False
    Selection.HomeKey wdStory
    Selection.Find.Text = "DEFINITIONS"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "INSTRUCTIONS"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start + 12, rng2.Start - 6) '-5
            rngfound.Select
            Selection.Copy
        End If
    End If

If blnFound = True Then

    'insert text here
    Documents(strNewDocLong).Activate
    
    Selection.GoTo What:=wdGoToBookmark, Name:="TermsAndDefinitions"
    Selection.PasteAndFormat (wdPasteDefault)

End If


End Sub
Sub ProcedureBody()
'
'Finds the Procedure Statements in the older McGill format and slapps it in the
'Procedure Body the new Emerson format. Let's do this!
'

Dim blnFound As Boolean
Dim rng1 As Range
Dim rng2 As Range
Dim rngfound As Range



Documents(strPrimaryDocLong).Activate

    Application.ScreenUpdating = False
    Selection.HomeKey wdStory
    Selection.Find.Text = "PROCEDURE STATEMENTS"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "QUALITY RECORDS^p"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start + 21, rng2.Start - 5)
            rngfound.Select
            Selection.Copy

        End If
    End If
    
    If blnFound = False And strDocType = "WI" Then
        Selection.Find.Text = "CHARACTERISTICS AND MEASURING DEVICES"
        blnFound = Selection.Find.Execute
        If blnFound Then
        
            Selection.MoveLeft wdWord
            Set rng1 = Selection.Range
            Selection.Find.Text = "QUALITY RECORDS^p"
            blnFound = Selection.Find.Execute
                  
            If blnFound Then
                Set rng2 = Selection.Range
                Set rngfound = ActiveDocument.Range(rng1.Start - 5, rng2.Start - 5)
                rngfound.Select
                Selection.Copy
            End If
        End If
    End If
    
If blnFound Then

'insert text here
Documents(strNewDocLong).Activate
'Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=16, Name:=""
Selection.GoTo What:=wdGoToBookmark, Name:="ProcedureBody"
Selection.PasteAndFormat (wdPasteDefault)
'Selection.Paragraphs.Add

End If


End Sub


Sub MSPF_Body()
'
'Finds the Procedure Statements in the older McGill format and slapps it in the
'Procedure Body the new Emerson format. Let's do this!
'

Dim blnFound As Boolean
Dim rng1 As Range
Dim rng2 As Range
Dim rngfound As Range


Documents(strPrimaryDocLong).Activate

    Application.ScreenUpdating = False
    Selection.HomeKey wdStory
    Selection.Find.Text = "INSTRUCTIONS^p"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "REVISIONS^p"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start + 21, rng2.Start - 5)
            rngfound.Select
            Selection.Copy

        End If
    End If
    
    If blnFound = False And strDocType = "WI" Then
        Selection.Find.Text = "CHARACTERISTICS AND MEASURING DEVICES"
        blnFound = Selection.Find.Execute
        If blnFound Then
        
            Selection.MoveLeft wdWord
            Set rng1 = Selection.Range
            Selection.Find.Text = "QUALITY RECORDS^p"
            blnFound = Selection.Find.Execute
                  
            If blnFound Then
                Set rng2 = Selection.Range
                Set rngfound = ActiveDocument.Range(rng1.Start - 5, rng2.Start - 5)
                rngfound.Select
                Selection.Copy
            End If
        End If
    End If
    
If blnFound Then

'insert text here
Documents(strNewDocLong).Activate
'Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=16, Name:=""
Selection.GoTo What:=wdGoToBookmark, Name:="ProcedureBody"
Selection.PasteAndFormat (wdPasteDefault)
'Selection.Paragraphs.Add

End If


End Sub

Sub Reference()
'
'Finds the purpose in the older McGill format and slapps it in the purpose
'the new Emerson format. Let's do this!
'

Dim blnFound As Boolean
Dim rng1 As Range
Dim rng2 As Range
Dim rngfound As Range


Documents(strPrimaryDocLong).Activate

    Application.ScreenUpdating = False
    Selection.HomeKey wdStory
    Selection.Find.Text = "REFERENCED DOCUMENTS"
    blnFound = Selection.Find.Execute
    
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "DEFINITIONS"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start + 22, rng2.Start - 5) 'rng2 - 2
            rngfound.Select
            Selection.Copy
        End If
            If blnFound Then
                Documents(strNewDocLong).Activate
                Selection.GoTo What:=wdGoToBookmark, Name:="Reference"
                'Selection.Paragraphs.Add
                Selection.PasteAndFormat (wdPasteDefault)
                'Selection.Paragraphs.Add
            End If
            

        'WI Specific coding for collecting
        If blnFound = False And strDocType = "WI" Then
            Selection.Find.Text = "CHARACTERISTICS AND MEASURING DEVICES"
            blnFound = Selection.Find.Execute
            
            If blnFound Then
                Set rng2 = Selection.Range
                Set rngfound = ActiveDocument.Range(rng1.Start + 22, rng2.Start - 5) 'rng2 - 2
                rngfound.Select
                Selection.Copy
                
                Documents(strNewDocLong).Activate
                Selection.GoTo What:=wdGoToBookmark, Name:="Reference"
                Selection.PasteAndFormat (wdPasteDefault)
                
            End If

            Documents(strPrimaryDocLong).Activate
            Selection.HomeKey wdStory
            Selection.Find.Text = "QUALITY RECORDS"
            blnFound = Selection.Find.Execute
            
            If blnFound Then
                Selection.MoveLeft wdWord
                Set rng1 = Selection.Range
                Selection.Find.Text = "REVISIONS"
                blnFound = Selection.Find.Execute
                If blnFound Then
                    Set rng2 = Selection.Range
                    Set rngfound = ActiveDocument.Range(rng1.Start + 15, rng2.Start - 5) 'rng2 - 2
                    rngfound.Select
                    Selection.Copy
                End If
                Documents(strNewDocLong).Activate
                Selection.GoTo What:=wdGoToBookmark, Name:="Reference2forWI"
                'Selection.Paragraphs.Add
                Selection.PasteAndFormat (wdPasteDefault)
                'Selection.Paragraphs.Add


            End If
        End If

    End If


End Sub

Sub Purpose()
'
'Finds the purpose in the older McGill format and slapps it in the purpose
'the new Emerson format. Let's do this!
'

Dim blnFound As Boolean
Dim rng1 As Range
Dim rng2 As Range
Dim rngfound As Range
Dim strThePurposeText As String

  
Documents(strPrimaryDocLong).Activate

    Application.ScreenUpdating = False
    Selection.HomeKey wdStory
    Selection.Find.Text = "PURPOSE"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "SCOPE"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start + 8, rng2.Start - 5)
            rngfound.Select
            Selection.Copy

        End If
    End If

    If blnFound = True Then
    
        'insert text here
        Documents(strNewDocLong).Activate
        Selection.GoTo What:=wdGoToBookmark, Name:="Purpose"
        Selection.PasteAndFormat (wdPasteDefault)
    
    End If


End Sub

Sub Scope()
'
'
'
'
Dim blnFound As Boolean
Dim rng1 As Range
Dim rng2 As Range
Dim rngfound As Range
Dim strThePurposeText As String

Documents(strPrimaryDocLong).Activate

    Application.ScreenUpdating = False
    Selection.HomeKey wdStory
    Selection.Find.Text = "SCOPE"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "REFERENCED DOCUMENTS"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start + 6, rng2.Start - 5) '3
            rngfound.Select
            Selection.Copy
        End If
    End If

    If blnFound = True Then
    
        Documents(strNewDocLong).Activate
        Selection.GoTo What:=wdGoToBookmark, Name:="Scope"
        Selection.PasteAndFormat (wdPasteDefault)
    
    End If


End Sub

Sub RevisedFindIt()
' Purpose: display the text between (but not including)
' the words "Purpose" and "scope" if they both appear.
    
    Dim rng1 As Range
    Dim rng2 As Range
    Dim strPurpose As String
    Dim strScope As String
    Dim strTermsAndDef As String
    Dim strRelated As String
    Dim strDefinitions As String
    Dim strProcedureStatements As String
    Dim strQualityRecords As String
    
    'PURPOSE <- Purpose
    ' Purpose: display the text between (but not including)
    ' the words "Purpose" and "scope" if they both appear.
    Set rng1 = ActiveDocument.Range
    If rng1.Find.Execute(FindText:="Purpose") Then
        Set rng2 = ActiveDocument.Range(rng1.End, ActiveDocument.Range.End)
        If rng2.Find.Execute(FindText:="Scope") Then
            strPurpose = ActiveDocument.Range(rng1.End, rng2.Start - 5).Text
            'MsgBox strPurpose
        End If
    End If
    
    'SCOPE <- Scope
    Set rng1 = ActiveDocument.Range
    If rng1.Find.Execute(FindText:="Scope") Then
        Set rng2 = ActiveDocument.Range(rng1.End, ActiveDocument.Range.End)
        If rng2.Find.Execute(FindText:="Related Procedures and other documents") Then
            strScope = ActiveDocument.Range(rng1.End, rng2.Start - 5).Text
            'MsgBox strScope
        End If
    End If

    
    'TERMS AND DEFINITIONS <- from relatd procedures and other documents
    Set rng1 = ActiveDocument.Range
    If rng1.Find.Execute(FindText:="RELATED PROCEDURES AND OTHER DOCUMENTS") Then
        Set rng2 = ActiveDocument.Range(rng1.End, ActiveDocument.Range.End)
        If rng2.Find.Execute(FindText:="DEFINITIONS") Then
            strRelated = ActiveDocument.Range(rng1.End, rng2.Start - 5).Text
            'MsgBox strRelated
        End If
    End If

    'TERMS AND DEFINITIONS <- Definitions
    Set rng1 = ActiveDocument.Range
    If rng1.Find.Execute(FindText:="Definitions") Then
        Set rng2 = ActiveDocument.Range(rng1.End, ActiveDocument.Range.End)
        If rng2.Find.Execute(FindText:="Procedure Statements") Then
            strDefinitions = ActiveDocument.Range(rng1.End, rng2.Start - 5).Text
            'MsgBox strDefinitions
        End If
    End If

    'PROCEDURE BODY <- Procedure Statements
    Set rng1 = ActiveDocument.Range
    If rng1.Find.Execute(FindText:="Procedure Statements") Then
        Set rng2 = ActiveDocument.Range(rng1.End, ActiveDocument.Range.End)
        If rng2.Find.Execute(FindText:="Quality Records") Then
            strProcedureStatements = ActiveDocument.Range(rng1.End, rng2.Start - 5).Text
            'MsgBox strProcedureStatements
        End If
    End If
 

End Sub

Sub TextAndRevisionsTable()

' TextAndRevisionsTable Macro
' uses file "TableForScript.doc" to get the table to paste
' at the end of the document
' This will make all text in Arial and a revisions table at the end.


Dim strFirstWindow As String

strFirstWindow = ActiveDocument.Name

    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "7.0 revis"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.GoTo What:=wdGoToTable, Which:=wdGoToNext, Count:=1, Name:=""
    Selection.Tables(1).Select
    Selection.Tables(1).delete
    Windows("TableForScript.doc [Compatibility Mode]").Activate
    Selection.HomeKey Unit:=wdLine
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.Copy
    
    Documents(strFirstWindow).Activate
       
    Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
    Selection.MoveUp Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.Font.Shrink
    Selection.Font.Shrink
    ActiveDocument.Save
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(0.44)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
    End With
   
       Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "7.0 revis"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.GoTo What:=wdGoToTable, Which:=wdGoToNext, Count:=1, Name:=""
    Selection.MoveDown Unit:=wdLine, Count:=4, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Font.Size = 8
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

    Selection.WholeStory
    Selection.Font.Name = "Arial"
    
    ActiveDocument.Save
    ActiveDocument.Close
    
End Sub

Sub FormatTitlesOldDocToNew()
'
'
'
'
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    Selection.HomeKey Unit:=wdStory

    With Selection.Find
        .Text = "1.0*PURPOSE"
        .Replacement.Text = "PURPOSE"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Style = "Heading 1"
        .Find.Execute
    End With
    With Selection.Find
        .Text = "2.0*SCOPE:"
        .Replacement.Text = "SCOPE:"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With

    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Style = "Heading 1"
        .Find.Execute
    End With
    With Selection.Find
        .Text = "3.0 TERMS AND DEFINITIONS:"
        .Replacement.Text = "TERMS AND DEFINITIONS:"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With

    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Style = "Heading 1"
        .Find.Execute
    End With
    With Selection.Find
        .Text = "4.0 PROCEDURE BODY:"
        .Replacement.Text = "PROCEDURE BODY:"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Style = "Heading 1"
        .Find.Execute
    End With
    
    With Selection.Find
        .Text = "5.0*RESPONSIBILITIES:"
        .Replacement.Text = "RESPONSIBILITIES:"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With

    Selection.Find.Execute

    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Style = "Heading 1"
        .Find.Execute
    End With

    With Selection.Find
        .Text = "6.0*REFERENCE:"
        .Replacement.Text = "REFERENCE DOCUMENTS:"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Style = "Heading 1"
    Selection.Find.Execute
    
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        Selection.Style = "Heading 1"
        .Find.Execute
    End With
   
End Sub

Sub NEWDOHICKY()
'
'
' changes the styles of the headings to the proper quickstyles
'
 For i = 1 To 7
        If i = 1 Then strSection = "PURPOSE:"
        If i = 2 Then strSection = "SCOPE:"
        If i = 3 Then strSection = "TERMS AND DEFINITIONS"
        If i = 4 Then strSection = "PROCEDURE BODY:"
        If i = 5 Then strSection = "RESPONSIBILITIES:"
        If i = 6 Then strSection = "REFERENCE DOCUMENTS:"
        If i = 7 Then strSection = "REVISIONS:"
    
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("Heading 1")
        With Selection.Find.Replacement.ParagraphFormat
        
        End With
        
        With Selection.Find
            .Text = strSection
            .Replacement.Text = strSection
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        
        Selection.Find.Execute
        
        With Selection
            If .Find.Forward = True Then
                .Collapse Direction:=wdCollapseStart
            Else
                .Collapse Direction:=wdCollapseEnd
            End If
            .Find.Execute Replace:=wdReplaceOne
            If .Find.Forward = True Then
                .Collapse Direction:=wdCollapseEnd
            Else
                .Collapse Direction:=wdCollapseStart
            End If
            .Find.Execute
        End With
    
    Next

End Sub

Sub RevisionsTableCount()
'
' Changes the 00 non-formatted text in the revision table to a formatted number
' so that MSword can pick it up and use it at the bottom of the document.

    Selection.HomeKey Unit:=wdStory
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "REVISIONS:"
        .Replacement.Text = "REVISIONS:"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.GoTo What:=wdGoToTable, Which:=wdGoToNext, Count:=1, Name:=""
   
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.delete Unit:=wdCharacter, Count:=1
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0.25)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabicLZ
        .NumberPosition = InchesToPoints(0.25)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 0
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
        True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = InchesToPoints(1)
        .BottomMargin = InchesToPoints(1)
        .LeftMargin = InchesToPoints(1)
        .RightMargin = InchesToPoints(1)
        .Gutter = InchesToPoints(0)
        .HeaderDistance = InchesToPoints(0.63)
        .FooterDistance = InchesToPoints(0.37)
        .PageWidth = InchesToPoints(8.5)
        .PageHeight = InchesToPoints(11)
        .FirstPageTray = 1025
        .OtherPagesTray = 1025
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(0.38)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
    End With
End Sub

Sub CapsForSetUpAndCharacteristics()
'
' Makes the words "characteriscs and measuring devices" and "set up approval" all caps
' I think it's for WIs with gage stuff.
'
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "characteristics and measuring devices"
        .Replacement.Text = "CHARACTERISTICS AND MEASURING DEVICES"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "set up approval"
        .Replacement.Text = "SET UP APPROVAL"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub FindAndReplaceDoubleSpaces()
'
' Replaces double tabs, spaces with mroe than 4 in a row, and more than 2 in a row.
'
'
   
    'replaces multiple tabs in a row with one tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t{2,}"
        .Replacement.Text = "^t" 'new tab change
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'finds and replaces multiple spaces with one space
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " {2,}"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " ^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub NewParagraphFixer()

'for the new format to new format


Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p^t" 'page break code
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t ^t" 'page break code
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^t" 'page break code
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'tabs kinda?
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " ^t " 'page break code
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    

    
    'the below code replaces 2 'enters' (paragraph symbols) and replces it with one
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p" 'two paragraph
        .Replacement.Text = "^p" 'single paragraph symbol
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub

Sub AddParagrahToTitles()
'
'
'
'
 For i = 1 To 7
    If i = 1 Then strSection = "PURPOSE:"
    If i = 2 Then strSection = "SCOPE:"
    If i = 3 Then strSection = "TERMS AND DEFINITIONS"
    If i = 4 Then strSection = "PROCEDURE BODY:"
    If i = 5 Then strSection = "RESPONSIBILITIES:"
    If i = 6 Then strSection = "REFERENCE DOCUMENTS:"
    If i = 7 Then strSection = "REVISIONS:"

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Heading 1")
    With Selection.Find.Replacement.ParagraphFormat
    End With
    With Selection.Find
        .Text = strSection
        .Replacement.Text = strSection
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With

Next
End Sub

Sub deleteEmptyParagraphs()

'''  question submitted at below sites about this code:
'''  http://stackoverflow.com/questions/23411007/call-argument-error-with-asc-delete-unnecessary-lines-in-msword-2007
'''  http://www.vbaexpress.com/forum/showthread.php?10558-Solved-How-to-delete-all-empty-rows-in-Word&p=309032

' Goes through the document line by line and finds lines that have any combination of
' paragraph characters, tab characters, or space characters and deletes them.
' delete empty paragraph
' delete empty line, delete line, delete paragraph

    Dim oPara As Word.Paragraph
    Dim var
    Dim SpaceTabCounter As Long
    Dim oChar As Word.Characters
    Dim TabCounter As Long

    For Each oPara In ActiveDocument.Paragraphs
        If Len(oPara.Range) = 1 Then
            oPara.Range.delete
        Else
            TabCounter = 0
            Set oChar = oPara.Range.Characters
            For var = 1 To oChar.Count
                If oChar(var) <> "" Then ' stops Asc from throwing runtime error 5
                    Select Case Asc(oChar(var)) ' no more errrors!
                        Case 9, 32 '9 is tabs, 32 is spaces,
                            SpaceTabCounter = SpaceTabCounter + 1
                    End Select
                End If
            Next
            If SpaceTabCounter + 1 = Len(oPara.Range) Then
             ' paragraph contains ONLY spaces, tabs, and the paragraph.
                oPara.Range.delete
            End If
        End If
    Next

End Sub


Sub DeleteLastParaOfFooter()
'
' DeleteLastParaOfFooter Macro
' Deletes last paragraph of footer. This was used on "older" documents made by
' scripts that weren't up to par.

    WordBasic.ViewFooterOnly
    Selection.EndKey Unit:=wdStory
    Selection.delete Unit:=wdCharacter, Count:=1
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Save
    
End Sub

Sub Tabs()

    'MsgBox ("Tabs is being used. May not want to replace with 'null', maybe a space.")
    
    'replaces multiple tabs in a row with a tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t{1,}"
        .Replacement.Text = "^t" 'new tab change
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'find all rows that start with a tab and delete the tab
    With Selection.Find
        .Text = "^p^t"
        .Replacement.Text = "^p" 'new tab change
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
End Sub

Sub EndOfOldEmersonForamtDocument()
    'tries to get the tables from the end of the other document which is after
    'the revisions portion of the documents
    
    'activate orriginal document
    Documents(strMainDoc & ".docx").Activate 'change back to doc
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "revisions:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.GoTo What:=wdGoToTable, Which:=wdGoToNext, Count:=1, Name:=""
    Selection.EndOf Unit:=wdTable

    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    Selection.Copy
    'go to document that we're changing
    Documents(strNewDocLong).Activate
    'go to end of document
    Selection.EndKey Unit:=wdStory
    Selection.Paste

End Sub

Sub SaveDocument()
    
    ChangeFileOpenDirectory _
        SaveToLocation
    ActiveDocument.SaveAs FileName:= _
        SaveToLocation & strMainDoc & ".docx" _
        , FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False
        
    strNewDocLong = strMainDoc & ".docx"
    strNewDoc = strMainDoc
    
End Sub

Sub TabsMultiTwoOne()
'
' TablesFixer Macro
' does neat things!
' Finds any 2 or more tabs in a row and replaces it with 1 tab.
' Finds a line that starts with a tab and removes the tab.
'

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t{2,}"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^t"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue 'wdFindStop 'wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue 'wdFindStop 'wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


End Sub



Sub delete()



Call deleteEmptyParagraphs

End Sub

Sub NewDoc()
    Documents(strNewDocLong).Activate
End Sub

Sub FontFormatting()

        Selection.WholeStory
        Selection.Font.Size = 10
        Selection.Font.Name = "Arial"
        Selection.HomeKey Unit:=wdStory

End Sub

Sub OldDoc()
    Documents(strPrimaryDocLong).Activate
End Sub

Sub FixParagraphIndents()

'applies to quickstyles "Normal" an "List Paragraph", may need to add more
'
'This is a little sketchy

'It relies on having the correct "Heading 1"
'This should put the "nomral" and "list" styles indented after the
'line before it, as long as the previous line has a correcting quick style
'This script falls apart on some of the "older" documents before the
'quickstyles were standardized with the script.

Dim prePara As Paragraph
Dim curPara As Paragraph
Dim nextPara As Paragraph
Dim i As Integer
Dim DoTermsAndDef
Dim TermAndDefFix
Dim TandDLine

DoTermsAndDef = False

For i = 2 To ActiveDocument.Paragraphs.Count

    Set prePara = ActiveDocument.Paragraphs(i - 1)
    Set curPara = ActiveDocument.Paragraphs(i)
    If i <> ActiveDocument.Paragraphs.Count Then
        Set nextPara = ActiveDocument.Paragraphs(i + 1)
    End If
    
    
    If Left(curPara.Range, 9) = "TERMS AND" Then
        TermAndDefFix = True
        TandDLine = i
    End If
    If Left(curPara.Range, 11) = "PROCEDURE B" Then TermAndDefFix = False



    If curPara.Style = "Normal" And Selection.Information(wdWithInTable) = False Or curPara.Style = "List Paragraph" And prePara.Style = "Heading 1" And Selection.Information(wdWithInTable) = False Then
        If TermAndDefFix <> True Then
            ActiveDocument.Paragraphs(i).Range.Select
            If Selection.Information(wdWithInTable) = False Then
                curPara.LeftIndent = prePara.LeftIndent
                curPara.FirstLineIndent = 0
            End If
        End If
    End If



    If TermAndDefFix = True Then
        ActiveDocument.Paragraphs(i).Range.Select
        If Selection.Information(wdWithInTable) = False Then
            curPara.FirstLineIndent = -18
            curPara.LeftIndent = 36
            'this line corrects the Terms And Definitions line that fell victim to formatting
            If i = TandDLine Then ActiveDocument.Paragraphs(TandDLine).LeftIndent = 18
        End If
    End If

    
Next
    
    Selection.HomeKey Unit:=wdStory

End Sub

Sub RenameDocToTitle()
'
' RenameDocToTitle Macro
' ECN Naming the document
'
' When to use:
' After putting in the new doc title into the title line and
' after double checking the docuemnt.
'
' How to improve this script:
' Make it into a loop that gets new data from an exel file or range of some sort
' so that the title doesn't have to be put in manually.

Dim MyDataTitle As DataObject
Dim NextData As DataObject
Dim EndingNumber As String

Call StartUp 'Sub StartUp() has the location of saving the document

Set MyDataTitle = New DataObject
'ECN Saving to new place
Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
Selection.Copy
MyDataTitle.GetFromClipboard
strClipSubject = MyDataTitle.GetText

With Dialogs(wdDialogFileSummaryInfo) 'makes title of new doc to the controlled doc title
    .Title = strClipSubject
    .Execute
End With

strNewNameECN = Dialogs(wdDialogFileSummaryInfo).Title

    ChangeFileOpenDirectory ECNSaveToLocation
    ActiveDocument.SaveAs FileName:= _
        ECNSaveToLocation & strNewNameECN & ".docx" _
        , FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False
        
'EndingNumber = Right(strClipSubject, 2)
'strClipSubject = Left(strClipSubject, Len(strClipSubject) - 2)
'
'Set NextData = New DataObject
'
'NextData = strClipSubject & EndingNumber
'NextData.PutInClipboard

Call FixParagraphIndents

ActiveDocument.Save
ActiveDocument.Close


End Sub

Sub trythis() 'scripts after importing data manually

Call ApplyHeadingStyles 'adjusts the numbering of the automatic numbering of document.
'Call CapsForSetUpAndCharacteristics 'looks for titles of sections and makes them all caps
Call RevisionsTableCount 'changes the cell that cantains 00 to a numbered format
Call FindAndReplaceDoubleSpaces
Call deleteEmptyParagraphs
Call FixParagraphIndents
ActiveDocument.Save
End Sub
Sub ChangeTablesRevisionsNumber()

' This is used to change the ECN number in the revisions table in a single document.
'
'
'
    Dim RevisionNumber
    
    RevisionNumber = "98167"
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "revisions:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.GoTo What:=wdGoToTable, Which:=wdGoToNext, Count:=1, Name:=""
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText Text:=RevisionNumber
    Selection.MoveLeft Unit:=wdCell
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    ActiveDocument.Save
    ActiveDocument.Close
End Sub
Sub IndexFileNamesAndSubjectProperties()

' This sub goes through a directory and pulls out the "Title" and "Subject" of each word document
' and places it in a word document. After it runs, it will make it into a cute table!
' list lists index table files
' gets title gets subject files in folder


    Dim vDirectory As String
    Dim oDoc As Document
    Dim strSubject
    Dim strTitle
    Dim vFile
    Dim strAllData
    vDirectory = CurDir()
   
    If Right(vDirectory, 1) <> "\" Then vDirectory = vDirectory & "\"

    vFile = Dir(vDirectory & "*.*")

    Do While vFile <> ""
        Set oDoc = Documents.Open(FileName:=vDirectory & vFile)
            
            strSubject = ActiveDocument.BuiltInDocumentProperties("Subject")
            strTitle = ActiveDocument.BuiltInDocumentProperties("Title")
            ActiveDocument.Close SaveChanges:=False
                        
            strAllData = strAllData & "#p#" & strSubject & "#t#" & strTitle
            'MsgBox (strAllData)
        
        vFile = Dir
    Loop
    
    Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
    Selection.TypeText (strAllData)

    
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "#t#"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "#p#"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.WholeStory
    WordBasic.TextToTable ConvertFrom:=1, NumColumns:=3, _
        InitialColWidth:=wdAutoPosition, Format:=0, Apply:=1184, AutoFit:=0, _
        SetDefault:=0, Word8:=0, Style:="Table Grid"
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Style = ActiveDocument.Styles("No Spacing")
    

    ChangeFileOpenDirectory vDirectory
    ActiveDocument.SaveAs FileName:=vDirectory & "ListOfFiles.docx"
    ActiveDocument.Close

End Sub


Sub AddSpaceBeforeAndAfter()

Selection.ParagraphFormat.SpaceAfter = 6
Selection.ParagraphFormat.SpaceBefore = 6

End Sub

Sub AddSpaceAfter()

Selection.ParagraphFormat.SpaceAfter = 6

End Sub

Sub AddSpaceBefore()

Selection.ParagraphFormat.SpaceBefore = 6

End Sub
Sub Macro1()
'
' Save as .DOCX file and close document
'
'
Dim docName

docName = Left(ActiveDocument.Name, (Len(ActiveDocument.Name) - 4))

'MsgBox (docName)
    
    ChangeFileOpenDirectory _
        "G:\Common\Anderson (the better looking one)\Quality Document ECN Process\3 Rename these New Name\98171 - IBOP-830 to BOP-150\DoubleCheck\"
    ActiveDocument.SaveAs FileName:= _
        "G:\Common\Anderson (the better looking one)\Quality Document ECN Process\3 Rename these New Name\98171 - IBOP-830 to BOP-150\DoubleCheck\" & docName & ".docx" _
        , FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False
    ActiveDocument.Close

End Sub
Sub Macro2()
'
' ecn replace number at bottom of page
'
'
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "97167"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Paste
    ActiveDocument.Save
    ActiveDocument.Close
End Sub

Sub MSPUnformatted()

'Gets Subject for Document to be placed as Title/Name of new document
MSPU_NameLong = ActiveDocument.Name
MSPU_Title = Left(MSPU_NameLong, Len(MSPU_NameLong) - 4)

'Goes to beginning of document changes font type and size
Selection.WholeStory
Selection.Font.Size = 10
Selection.Font.Name = "Arial"

'Gets Description/Subject of the document
Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=1, Name:=""
Selection.StartOf Unit:=wdParagraph 'select paragraph start
Selection.MoveEnd Unit:=wdParagraph 'select paragraph end
MSPU_Subject = Selection 'make MSPU_Subject the selection
MsgBox (MSPU_Subject) 'test MSPU Subject

'Gets the body of the document
Selection.EndKey Unit:=wdStory, Extend:=wdExtend 'select paragraph end
Selection.Copy

'Open template
Documents.Add Template:= _
    "G:\Common\Jerry\Controlled Documents\Document Conversion\Template QA10 - Doc with Body.dotx" _
    , NewTemplate:=False, DocumentType:=0
    
    

'Go to body to paste
Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=8, Name:=""
 Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)

Call ApplyHeadingStyles 'adjusts the numbering of the automatic numbering of document.
Call RevisionsTableCount 'changes the cell that cantains 00 to a numbered format

Call FindAndReplaceDoubleSpaces
Call deleteEmptyParagraphs
Call FixParagraphIndents
Call NewParagraphFixer

End Sub

Sub MSPU_ApplyHeadingStyles()

    Dim i As Integer
    Dim Para As Paragraph, Rng As Range, iLvl As Long
    Dim blnFound As Boolean
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rngfound As Range
       
    'put in range from Purpose: to Revisions:
    Selection.HomeKey wdStory
    
    'select all
    'numberthing
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    
    Selection.Find.Text = "BODY:"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "REVISIONS:"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start - 4, rng2.Start)
            rngfound.Select
        End If
    End If
 
With rngfound 'ActiveDocument.Range
    For Each Para In .Paragraphs
     Set Rng = Para.Range.Words.First
     With Rng
       If IsNumeric(.Text) Then
         While .Characters.Last.Next.Text Like "[0-9. " & vbTab & "]"
           .End = .End + 1
         Wend
         iLvl = UBound(Split(.Text, "."))
         If IsNumeric(Split(.Text, ".")(UBound(Split(.Text, ".")))) Then iLvl = iLvl + 1
         If iLvl < 10 Then
           .Text = vbNullString
           Para.Style = "Heading " & iLvl
         End If
       End If
     End With
   Next
End With
  
 'Corrects the heading of the Section Titles to be Heading 1
 'because something in this sub is messing up :-(
  For i = 1 To 2
    If i = 1 Then strSection = "BODY:"
    If i = 2 Then strSection = "REVISIONS:"

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Heading 1")
    With Selection.Find.Replacement.ParagraphFormat
    End With
    With Selection.Find
        .Text = strSection
        .Replacement.Text = strSection
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With

   Next
   

 
 End Sub


Sub TabFinder()

Dim oPara As Word.Paragraph 'paragraph
Dim var 'a counter for a FOR loop for finding tabs in a paragraph, represents a single character
Dim i As Integer

Dim paraText As String 'represents the string of each paragraph

Call deleteEmptyParagraphs
Call FindAndReplaceDoubleSpaces
Call Tabs

For var = 1 To ActiveDocument.Paragraphs.Count 'for every paragraph in the document
    Set oPara = ActiveDocument.Paragraphs(var)
    For i = oPara.Range.Characters.Count To 1 Step -1
        If oPara.Range.Characters(i) = Chr(9) Then
            'As long as there's another tab to the left of this one, delete this one
            If InStr(Left(oPara.Range.Text, i - 1), Chr(9)) > 1 Then
                oPara.Range.Characters(i).delete
            End If
        End If
    Next
Next



End Sub

Sub MSPF_Translate()

Call StartUp
    
    Dim i As Integer
    Dim Para As Paragraph, Rng As Range, iLvl As Long
    Dim blnFound As Boolean
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rngfound As Range
       
    'put in range from Purpose: to Revisions:
    'Selection.HomeKey wdStory
    
    'select all
    'numberthing
    ActiveDocument.Range.ListFormat.ConvertNumbersToText

strMainDocLong = ActiveDocument.Name

strMainDoc = Left(strMainDocLong, Len(strMainDocLong) - 4) 'variable change for docx doc length document name
strPrimaryDocLong = strMainDocLong
strPrimaryDoc = strMainDoc
strNewDocLong = strMainDoc & ".docx"



Documents.Add Template:= _
    "G:\Common\Jerry\Controlled Documents\Document Conversion\Template QA10.dotx" _
    , NewTemplate:=False, DocumentType:=0
    
    ChangeFileOpenDirectory MSPF_SaveToLocation
    ActiveDocument.SaveAs (strNewDocLong)

    With Dialogs(wdDialogFileSummaryInfo) 'makes title of new doc to the controlled doc title
     .Title = strMainDoc
     .subject = "" 'figure out how to do this
     .Execute
    End With

 Documents(strMainDocLong).Activate


Call MSPF_Body
Call MSPF_TermsAndDefinitions
Call Scope
Call Purpose
Call Reference '123
    'put in range from Purpose: to Revisions:
    Selection.HomeKey wdStory
    
    'select all
    'numberthing
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    
    Selection.Find.Text = "PURPOSE:"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "Revisions:"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start - 4, rng2.Start)
            rngfound.Select
        End If
    End If

With rngfound 'ActiveDocument.Range
    For Each Para In .Paragraphs
     Set Rng = Para.Range.Words.First
     With Rng
       If IsNumeric(.Text) Then
         While .Characters.Last.Next.Text Like "[0-9. " & vbTab & "]"
           .End = .End + 1
         Wend
         iLvl = UBound(Split(.Text, "."))
         If IsNumeric(Split(.Text, ".")(UBound(Split(.Text, ".")))) Then iLvl = iLvl + 1
         If iLvl < 10 And iLvl > 0 Then
           .Text = vbNullString
           Para.Style = "Heading " & iLvl
         End If
       End If
     End With
   Next
End With


  
Call FixLevel1Headings
Call TabFinder
Call RevisionsTableCount
Call deleteEmptyParagraphs
Call FixFrontTypeSize

End Sub

Sub FixFrontTypeSize()

    Dim blnFound As Boolean
    Dim rng2 As Range
    Dim rng1 As Range
    Dim rngfound As Range
    
    Selection.Find.Text = "PURPOSE:"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.MoveLeft wdWord
        Set rng1 = Selection.Range
        Selection.Find.Text = "Revisions:"
        blnFound = Selection.Find.Execute
        If blnFound Then
            Set rng2 = Selection.Range
            Set rngfound = ActiveDocument.Range(rng1.Start, rng2.Start)
            rngfound.Select
            With rngfound 'ActiveDocument.Range
                Selection.Font.Size = 10
                Selection.Font.Name = "Arial"
            End With
        End If
    End If
    
Selection.HomeKey Unit:=wdStory
    
End Sub

Sub FixLevel1Headings()

Dim i As Integer


 'Corrects the heading of the Section Titles to be Heading 1
 'because something in this sub is messing up :-(

For i = 1 To 8
    If i = 1 Then strSection = "PURPOSE:"
    If i = 2 Then strSection = "SCOPE:"
    If i = 3 Then strSection = "TERMS AND DEFINITIONS"
    If i = 4 Then strSection = "PROCEDURE BODY:"
    If i = 5 Then strSection = "RESPONSIBILITIES:"
    If i = 6 Then strSection = "REFERENCE DOCUMENTS:"
    If i = 7 Then strSection = "FLOW CHART/TURTLE DIAGRAM:"
    If i = 8 Then strSection = "REVISIONS:"

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Heading 1")
    With Selection.Find.Replacement.ParagraphFormat
    End With
    With Selection.Find
        .Text = strSection
        .Replacement.Text = strSection
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With


Next
End Sub




Sub ClearLeadingSpacesAndTabs()
'
' Macro3 Macro
'
'
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p "
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^p^t"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p "
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub BasicFixing()
    Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText
    
    Call AddBlankSections
    
    Call FindAndReplaceDoubleSpaces

    Call deleteEmptyParagraphs
    Call ClearLeadingSpacesAndTabs
    Call ApplyHeadingStyles
    Call FixLevel1Headings
    Call FixParagraphIndents
    Call RevisionsTableCount
    Call deleteEmptyParagraphs
    Call SpaceingBeforeAfterPara
    
    Selection.WholeStory
    Selection.Font.Size = 10
    Selection.Font.Name = "Arial"
    Application.ScreenUpdating = True
End Sub
Sub SpaceingBeforeAfterPara()

Dim Para As Word.Paragraph


For Each Para In ActiveDocument.Paragraphs
If Para.SpaceBefore = 12 Then Para.SpaceBefore = 6
If Para.SpaceAfter = 12 Then Para.SpaceAfter = 6
If Para.Alignment = wdAlignParagraphJustify Then Para.Alignment = wdAlignParagraphLeft

Next

End Sub
Sub DeleteAst()

Dim objDoc As Document
Dim oCell As Cell
Dim oCol As Column
Dim objTable As Table
Dim bFlag As Boolean

Set objDoc = ActiveDocument
Set objTable = Selection.Tables(1)

'This may or may not be necessary, but I think it's a good idea.
'Tables with spans can not be accessed via the spanned object.
'Helper function below.
If IsColumnAccessible(objTable, 2) Then
    For Each oCell In objTable.Columns(2).Cells
        'This is the easiest way to check for an asterisk,
        'but it assumes you have decent control over your
        'content. This checks for an asterisk anywhere in the
        'cell. If you need to be more specific, keep in mind
        'that the cell will contain a paragraph return as well,
        'at a minimum.
        bFlag = (InStr(oCell.Range.Text, "*") > 0)
        'Delete the content of the cell; again, this assumes
        'the only options are blank or asterisk.
        oCell.Range.delete
        objDoc.FormFields.Add Range:=oCell.Range, Type:=wdFieldFormCheckBox
        'Set the value. I found some weird results doing this
        'any other way (such as setting the form field to a variable).
        'This worked, though.
        If bFlag Then
            oCell.Range.FormFields(1).CheckBox.Value = True
        End If
    Next oCell
End If
'Then do the same for column 5.

End Sub
Public Function IsColumnAccessible(ByRef objTable As Table, iColumn As Integer) As Boolean
Dim objCol As Column
'This is a little helper function that returns false if
'the column can't be accessed. If you know you won't have
'any spans, you can probably skip this.
On Error GoTo IsNotAccessible
IsColumnAccessible = True
Set objCol = objTable.Columns(iColumn)
Exit Function

IsNotAccessible:
IsColumnAccessible = False

End Function

Sub TableAdjusting()

Dim blnFound
    Selection.WholeStory
    Selection.Font.Size = 10
    Selection.Font.Name = "Arial"
    
    
    'replaces ECN#/ANNUAL MANAGEMENT REVIEW with ECN#
    Selection.HomeKey Unit:=wdStory
    Selection.Find.Text = "ECN#/ANNUAL MANAGEMENT REVIEW"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "ECN#/ANNUAL MANAGEMENT REVIEW"
            .Replacement.Text = "ECN#"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        With Selection
            If .Find.Forward = True Then
                .Collapse Direction:=wdCollapseStart
            Else
                .Collapse Direction:=wdCollapseEnd
            End If
            .Find.Execute Replace:=wdReplaceOne
            If .Find.Forward = True Then
                .Collapse Direction:=wdCollapseEnd
            Else
                .Collapse Direction:=wdCollapseStart
            End If
            .Find.Execute
        End With
        Selection.HomeKey Unit:=wdStory
    End If
    
    'goes to revisions to get to the table
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "REVISIONS:"
        .Replacement.Text = "REVISIONS:"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.GoTo What:=wdGoToTable, Which:=wdGoToNext, Count:=1, Name:=""
    
    
        
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    Selection.Rows.HeightRule = wdRowHeightExactly
    Selection.Rows.Height = InchesToPoints(0.15)
    Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    Selection.Columns.PreferredWidth = InchesToPoints(1.3)
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
   Selection.Tables(1).Columns.DistributeWidth
    
'    Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
'    Selection.MoveLeft Unit:=wdCell, Count:=1
'
'    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
'    Selection.Cells.DistributeWidth
    
    'Selection.HomeKey Unit:=wdStory
    
    ActiveDocument.Save
    MsgBox (":-)")
    
End Sub

Sub AddBlankSections()

Dim blnFound
Dim RefMSGBln
Dim strSection
Dim rng1
Dim rng2
Dim rngfound

Dim RefLin
Dim RefPag
Dim BodLin
Dim BodPag



Dim PurBln '1
Dim ScoBln '2
Dim TerBln '3
Dim ProBln '4
Dim BodBln '5 - part of procedure if it only says "body:" instead of "procedure body:"
Dim ResBln '6
Dim RefBln '7
Dim FloBln '8

Call TableAdjusting 'change this back when done


Dim i

    'replaces ECN#/ANNUAL MANAGEMENT REVIEW with ECN#
    
    For i = 1 To 8
        If i = 1 Then strSection = "PURPOSE:"
        If i = 2 Then strSection = "SCOPE:"
        If i = 3 Then strSection = "TERMS AND DEFINITIONS:"
        If i = 4 Then strSection = "PROCEDURE BODY:"
        If i = 5 Then strSection = "BODY:"
        If i = 6 Then strSection = "RESPONSIBILITIES:"
        If i = 7 Then strSection = "REFERENCE DOCUMENTS:"
        If i = 8 Then strSection = "FLOW CHART/TURTLE DIAGRAM:"
    
        Selection.HomeKey Unit:=wdStory
        Selection.Find.Text = strSection
        blnFound = Selection.Find.Execute
        If blnFound Then
            If strSection = "PURPOSE:" Then PurBln = True
            If strSection = "SCOPE:" Then ScoBln = True
            If strSection = "TERMS AND DEFINITIONS:" Then TerBln = True
            If strSection = "PROCEDURE BODY:" Then 'different to skip body
                ProBln = True
                i = i + 1 'skips over "body" since we know it exists already
            End If
            If strSection = "BODY:" Then BodBln = True
            If strSection = "RESPONSIBILITIES:" Then ResBln = True
            If strSection = "REFERENCE DOCUMENTS:" Then
                RefBln = True
                RefMSGBln = True
            End If
            If strSection = "FLOW CHART/TURTLE DIAGRAM:" Then FloBln = True
        End If
    Next
    
    If ProBln = False And BodBln = False Then
        MsgBox ("'PROCEDURE BODY:' is not located in this file. Please correct this.")
        Exit Sub
    End If
    
            'NO (purpose, scope, terms, reference)
            If PurBln = False And ScoBln = False And TerBln = False And ProBln = True Or PurBln = False And ScoBln = False And TerBln = False And BodBln = True Then
                    Selection.HomeKey Unit:=wdStory
                    Selection.Find.Text = "PROCEDURE BODY:"
                    blnFound = Selection.Find.Execute
                    If blnFound Then
                        Selection.HomeKey Unit:=wdLine
                        Selection.TypeText Text:="PURPOSE:"
                        PurBln = True
                        Selection.TypeParagraph
                        Selection.TypeText Text:="SCOPE:"
                        ScoBln = True
                        Selection.TypeParagraph
                        
                        If RefBln = True Then
                            RefMSGBln = True
                        End If
                        If RefBln = False Then
                            Selection.TypeText Text:="REFERENCE DOCUMENTS:"
                            Selection.TypeParagraph
                            RefBln = True
                        End If
                        Selection.TypeText Text:="TERMS AND DEFINITIONS:"
                        TerBln = True
                        Selection.TypeParagraph
                    End If
                    If blnFound = False And BodBln = True Then
                        Selection.Find.Text = "BODY:"
                        blnFound = Selection.Find.Execute
                        Selection.HomeKey Unit:=wdLine
                        Selection.TypeText Text:="PROCEDURE "
                        ProBln = True
                        BodBln = False
                        Selection.HomeKey Unit:=wdLine
                        Selection.TypeText Text:="PURPOSE:"
                        PurBln = True
                        Selection.TypeParagraph
                        Selection.TypeText Text:="SCOPE:"
                        ScoBln = True
                        Selection.TypeParagraph
                        
                        If RefBln = True Then
                            RefMSGBln = True
                        End If
                        If RefBln = False Then
                            Selection.TypeText Text:="REFERENCE DOCUMENTS:"
                            Selection.TypeParagraph
                            RefBln = True
                        End If
                        
                        Selection.TypeText Text:="TERMS AND DEFINITIONS:"
                        TerBln = True
                        Selection.TypeParagraph
                    End If
                    
            End If 'end NO (purpose, scope, terms) - takes into account ref. might be in different spot, needs manual change if so

            If FloBln = False Then
                    Selection.HomeKey Unit:=wdStory
                    Selection.Find.Text = "REVISIONS:"
                    blnFound = Selection.Find.Execute
                    If blnFound Then
                        Selection.HomeKey Unit:=wdLine 'goes to beginning of line
                        Selection.TypeText Text:="FLOW CHART/TURTLE DIAGRAM:" 'type this
                            FloBln = True 'set flowchart bool to true
                            Selection.TypeParagraph 'make new paragraph
                    End If
            End If

            'changes "BODY:" to "PROCEDURE BODY:"
            If BodBln = True Then
                    Selection.HomeKey Unit:=wdStory
                    Selection.Find.Text = "BODY:"
                    blnFound = Selection.Find.Execute
                        Selection.HomeKey Unit:=wdLine
                        Selection.TypeText Text:="PROCEDURE "
                        Selection.HomeKey Unit:=wdLine
                        BodBln = False
                        ProBln = True
            End If




            'Adds purpose if Scope is present *****check here
            If ScoBln = True And PurBln = False Then
                    Selection.HomeKey Unit:=wdStory
                    Selection.Find.Text = "SCOPE:"
                    blnFound = Selection.Find.Execute
                        Selection.HomeKey Unit:=wdLine
                        Selection.TypeText Text:="PURPOSE:"
                        Selection.TypeParagraph
                        
                        PurBln = True
            End If



            'Adds TER(rms and definitions) if PRO(cedure body) is present
            If ProBln = True And TerBln = False Then
                    Selection.HomeKey Unit:=wdStory
                    Selection.Find.Text = "PROCEDURE BODY:"
                    blnFound = Selection.Find.Execute
                        Selection.HomeKey Unit:=wdLine
                        Selection.TypeText Text:="TERMS AND DEFINITIONS:"
                        Selection.TypeParagraph
                        TerBln = True
            End If
            
            
            'Adds (Ref)
            If TerBln = True And RefBln = False Then
                    Selection.HomeKey Unit:=wdStory
                    Selection.Find.Text = "TERMS AND DEFINITIONS:"
                    blnFound = Selection.Find.Execute
                        Selection.HomeKey Unit:=wdLine
                        Selection.TypeText Text:="REFERENCE DOCUMENTS:"
                        Selection.TypeParagraph
                        RefBln = True
            End If
            
            
            'Adds TER(rms and definitions) if PRO(cedure body) is present
            If RefBln = True And ScoBln = False Then
                    Selection.HomeKey Unit:=wdStory
                    Selection.Find.Text = "REFERENCE DOCUMENTS:"
                    blnFound = Selection.Find.Execute
                        Selection.HomeKey Unit:=wdLine
                        Selection.TypeText Text:="SCOPE:"
                        Selection.TypeParagraph
                        ScoBln = True
            End If


            'puts in flow chart/turtle diagram line if not present
            If FloBln = False Then
                    Selection.HomeKey Unit:=wdStory
                    Selection.Find.Text = "REVISIONS:"
                    blnFound = Selection.Find.Execute
                    If blnFound Then
                        Selection.HomeKey Unit:=wdLine 'goes to beginning of line
                        Selection.TypeText Text:="FLOW CHART/TURTLE DIAGRAM:" 'type this
                            FloBln = True 'set flowchart bool to true
                            Selection.TypeParagraph 'make new paragraph
                    End If
            End If


            'if responsibilities is not present, then put it in!
            If ResBln = False And FloBln = True Then
                    Selection.HomeKey Unit:=wdStory
                    Selection.Find.Text = "FLOW CHART/TURTLE DIAGRAM:"
                    blnFound = Selection.Find.Execute
                    If blnFound Then
                        Selection.HomeKey Unit:=wdLine
                        Selection.TypeText Text:="RESPONSIBILITIES:"
                            ResBln = True
                            Selection.TypeParagraph
                    End If
            End If
            


    If RefMSGBln = True Then
        Selection.HomeKey Unit:=wdStory
        
        'the below finds the line number of Pro
        Selection.Find.Text = "PROCEDURE BODY:"
        blnFound = Selection.Find.Execute
        BodLin = Selection.Range.Information(wdFirstCharacterLineNumber)
        BodPag = Selection.Range.Information(wdActiveEndPageNumber)
        'the below find the line number of Ref
        Selection.HomeKey Unit:=wdStory
        Selection.Find.Text = "REFERENCE DOCUMENTS:"
        blnFound = Selection.Find.Execute
        RefLin = Selection.Range.Information(wdFirstCharacterLineNumber)
        RefPag = Selection.Range.Information(wdActiveEndPageNumber)
        'MsgBox (BodLin & " Body Line" & "..." & RefLin & "RefLin") 'this line displays a msgbox with the line numbers
            
            'If Ref is found and Body is before Ref then
            
            If BodPag < RefPag Then
                If blnFound Then
                    Set rng1 = Selection.Range
                    Selection.MoveLeft wdWord
                    Selection.Find.Text = "FLOW CHART/TURTLE DIAGRAM:"
                    
                    blnFound = Selection.Find.Execute
                    If blnFound Then
                        Set rng2 = Selection.Range
                        Set rngfound = ActiveDocument.Range(rng1.Start, rng2.Start - 0)
                            rngfound.Select
                            Selection.Cut
                                Selection.HomeKey Unit:=wdStory
                                Selection.Find.Text = "TERMS AND DEFINITIONS:"
                                blnFound = Selection.Find.Execute
                                Selection.HomeKey Unit:=wdLine
                                Selection.Paste
                    End If
                End If
            End If
            
            If BodPag = RefPag Then
                If blnFound And BodLin < RefLin Then
                    Selection.MoveLeft wdWord
                    Set rng1 = Selection.Range
                    Selection.Find.Text = "FLOW CHART/TURTLE DIAGRAM:"
                    blnFound = Selection.Find.Execute
                    If blnFound Then
                        Set rng2 = Selection.Range
                        Set rngfound = ActiveDocument.Range(rng1.Start, rng2.Start - 0)
                            rngfound.Select
                            Selection.Cut
                                Selection.HomeKey Unit:=wdStory
                                Selection.Find.Text = "TERMS AND DEFINITIONS:"
                                blnFound = Selection.Find.Execute
                                Selection.HomeKey Unit:=wdLine
                                Selection.Paste
                    End If
                End If
            End If
            
            
    End If
    
  
    Selection.HomeKey Unit:=wdStory
    Selection.Find.Text = "FLOW CHART/TURTLE DIAGRAM:"
    blnFound = Selection.Find.Execute
    If blnFound Then
        Selection.HomeKey Unit:=wdLine
        Selection.Range.ListFormat.ListOutdent
    End If
 
    
    
Selection.HomeKey Unit:=wdStory

End Sub
Sub samefont()

    Selection.WholeStory
    Selection.Font.Size = 10
    Selection.Font.Name = "Arial"
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " {2,}"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Call SpaceingBeforeAfterPara
    'ActiveDocument.Save
    'ActiveDocument.Close

End Sub

Sub SwitchTitleandSubject()

'-----first part works-----'
'switches the TITLE and SUBJECT of the document

Dim strNewTitle
Dim strOldTitle
Dim strNewSubject
Dim strOldSubject

    Documents(Fname).Activate

    strOldTitle = Dialogs(wdDialogFileSummaryInfo).Title
    strOldSubject = Dialogs(wdDialogFileSummaryInfo).subject
     
    strNewSubject = strOldTitle
    strNewTitle = strOldSubject
    'MsgBox (strNewSubject)
    'MsgBox (strNewTitle)
    

    With Dialogs(wdDialogFileSummaryInfo)
     .Title = strNewTitle
     .subject = strNewSubject
     .Execute
    End With

'-----second part testing--'
    Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=2, Name:=""
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend 'select line
    Selection.Cut 'cut text
    Selection.EndKey Unit:=wdLine 'go to end of line
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph 'make a new line
    Selection.PasteAndFormat (wdPasteDefault) 'paste text
    Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=4, Name:="" 'goto forth line
    Selection.delete Unit:=wdCharacter, Count:=1 'delete 1 character
    
    'add in hyperlinke
    Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=2, Name:=""
    
'---third part not working--'
    'here
    'Call hyperlinknewsubject 'tries to make the text a hyperlinnk

    
'---fourth part works---'
    'edit the footer to do the ol' switcheroo
    WordBasic.ViewFooterOnly
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Cut
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph
    Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
    Selection.delete Unit:=wdCharacter, Count:=1
    Selection.HomeKey Unit:=wdStory
    Selection.EndKey Unit:=wdLine
    Selection.MoveLeft Unit:=wdWord, Count:=4, Extend:=wdExtend
    Selection.Cut
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine
    Selection.TypeText Text:=vbTab
    Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
    Call FixWordDocTitleSubject
End Sub

Sub hyperlinknewsubject()
    'the text seemed to be too long and cause a memory error...here's the work around
    Dim First
    Dim Second
    Dim Third
    
    First = "http://ptsportal.emrsn.org/sites/qa/eptqa/docs/"
    Second = "PTS%20QA%20Documents%20Tier%202%20Only/Tier%20"
    Third = "2%20Quality%20Documents/BOP-110.docx"
    
    
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Hyperlinks.Add Anchor:=Selection.Range, Address:=First & Second & Third
    Selection.Style = ActiveDocument.Styles("QD_DocSubject")

End Sub
Sub ReplaceEntireHdr()

'COPY THE HEADER YOU WANT AND RUN THIS
'fix header, update header, replace header script
'**Action needed before running macro**
' *See Below*

    Dim Fname As String
    Dim wrd As Word.Application
    Set wrd = CreateObject("word.application")
    wrd.Visible = True
    ActiveDocument.Close False
    'AppActivate wrd.Name
     '*Change the directory to YOUR folder's path
    Fname = Dir("C:\Users\raanderson\Desktop\Test Script\*.docx")
    Do While (Fname <> "")
        With wrd
             '*Change the directory to YOUR folder's path
            .Documents.Open ("C:\Users\raanderson\Desktop\Test Script\" & Fname)
            If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
                .ActiveWindow.ActivePane.View.Type = wdPrintView
            Else
                .ActiveWindow.View.Type = wdPrintView
            End If
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
            .Selection.WholeStory
            .Selection.Paste

            Call FixFooter
            
            
            .ActiveDocument.Save
            .ActiveDocument.Close
        End With
        Fname = Dir
    Loop
    Set wrd = Nothing

End Sub
Sub FixWordDocTitleSubject()


    Dim Fname As String
    Dim wrd As Word.Application
    Dim Directory As String
    Dim SaveDirectory As String
    
    Dim strNewTitle
    Dim strOldTitle
    Dim strNewSubject
    Dim strOldSubject
    
    
    Directory = "C:\Users\raanderson\Desktop\Test Script\"
    SaveDirectory = "C:\Users\raanderson\Desktop\Test Script\new\"
    
    Set wrd = CreateObject("word.application")
    wrd.Visible = True
    AppActivate wrd.Name
    '*Change the directory to YOUR folder's path
    Fname = Dir(Directory)
    Do While (Fname <> "")
        With wrd
             '*Change the directory to YOUR folder's path
            .Documents.Open (Directory & Fname)
        End With

        ChangeFileOpenDirectory SaveDirectory
        
        
        strOldTitle = Dialogs(wdDialogFileSummaryInfo).Title
        strOldSubject = Dialogs(wdDialogFileSummaryInfo).subject
         
        strNewSubject = strOldTitle
        strNewTitle = strOldSubject
        'MsgBox (strNewSubject)
        'MsgBox (strNewTitle)
        

        With Dialogs(wdDialogFileSummaryInfo)
         .Title = strNewTitle
         .subject = strNewSubject
         .Execute
        End With

    '-----second part testing--'
        Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=2, Name:=""
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend 'select line
        Selection.Cut 'cut text
        Selection.EndKey Unit:=wdLine 'go to end of line
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.TypeParagraph 'make a new line
        Selection.PasteAndFormat (wdPasteDefault) 'paste text
        Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=4, Name:="" 'goto forth line
        Selection.delete Unit:=wdCharacter, Count:=1 'delete 1 character
        
        'add in hyperlinke
        Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst, Count:=2, Name:=""
        
    '---third part not working--'
        'here
        'Call hyperlinknewsubject 'tries to make the text a hyperlinnk
    
        
    '---fourth part works---'
        'edit the footer to do the ol' switcheroo
        WordBasic.ViewFooterOnly
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.Cut
        Selection.EndKey Unit:=wdLine
        Selection.TypeParagraph
        Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
        Selection.delete Unit:=wdCharacter, Count:=1
        Selection.HomeKey Unit:=wdStory
        Selection.EndKey Unit:=wdLine
        Selection.MoveLeft Unit:=wdWord, Count:=4, Extend:=wdExtend
        Selection.Cut
        Selection.MoveDown Unit:=wdLine, Count:=1
        Selection.EndKey Unit:=wdLine
        Selection.TypeText Text:=vbTab
        Selection.PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
        
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        
        Call FixWordDocTitleSubject
    
    
    
    
            
            
            
                ChangeFileOpenDirectory SaveDirectory
            With wrd
                .ActiveDocument.Save
                .ActiveDocument.Close
            End With
            Fname = Dir
    Loop
    Set wrd = Nothing

End Sub
Sub FixWordDocs2()


    Dim vDirectory As String
    Dim oDoc As Document
    Dim strSubject
    Dim strTitle
    Dim vFile
    Dim strAllData
    vDirectory = CurDir()
   
    If Right(vDirectory, 1) <> "\" Then vDirectory = vDirectory & "\"

    vFile = Dir(vDirectory & "*.*")

    Do While vFile <> ""
        Set oDoc = Documents.Open(FileName:=vDirectory & vFile)
            
            strSubject = ActiveDocument.BuiltInDocumentProperties("Subject")
            strTitle = ActiveDocument.BuiltInDocumentProperties("Title")
            ActiveDocument.Close SaveChanges:=False
                        
            strAllData = strAllData & "#p#" & strSubject & "#t#" & strTitle
            'MsgBox (strAllData)
        
        vFile = Dir
    Loop
    
End Sub

Sub EndOfDocumentBookmark()

    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "Annual management review"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.MoveDown Unit:=wdLine, Count:=2
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="EndOfDocument"
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    ActiveDocument.Save
End Sub


Sub Macro3()
'alt+j
    Dim strECN As String
    Dim strDate As String
    
    
    strECN = "19873"
    strDate = "12/22/2014"
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "12/08/2014"
        .Replacement.Text = strDate
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    With Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    
    
    With Selection.Find
        .Text = "ECNNUMBER"
        .Replacement.Text = strECN
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveDocument.Save
    ActiveDocument.Close
    
End Sub

Sub FixFooter()
'
' Used Above
'
'
    WordBasic.ViewFooterOnly
    Selection.EndKey Unit:=wdStory
    Selection.MoveLeft Unit:=wdWord, Count:=3, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveUp Unit:=wdParagraph, Count:=3, Extend:=wdExtend
    Selection.Font.Color = wdColorAutomatic
    Selection.HomeKey Unit:=wdStory
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    Application.WindowState = wdWindowStateNormal
End Sub
Sub Macro4()
'
' Macro4 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "/ANNUAL MANAGEMENT REVIEW"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

