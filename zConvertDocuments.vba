Option Explicit

Public strNewDoc As String 'new document without .doc (subtracts 4 for the extention)
Public strNewDocLong As String
Public strSection As String
Public strDocType As String
Public strClipSubject As String
Public strMainDoc As String
Public strMainDocLong As String
Public SaveToLocation As String
Public strNewNameECN As String
Public ECNSaveToLocation As String
Public DocTitle As String
Public DocNumber As String

'MSPU Names
Public MSPU_Subject As String
Public MSPU_Title As String
Public MSPU_NameLong As String
Public MSPF_SaveToLocation As String 'MSPFormatted_SaveToLocation

Sub LoopThroughFiles()

'run this but make a script to add a bookmark at the bottom of the files.
'This willhelp a bunch. use table as reference and go down?

    Dim StrFile As String
    StrFile = Dir("C:\Users\raanderson\Desktop\Script1\")
    
    Do While Len(StrFile) > 0
        Documents.Open FileName:=("C:\Users\raanderson\Desktop\Script1\" & StrFile)
        Application.ScreenUpdating = False
        If ActiveDocument.Bookmarks.Exists("EndOfDocument") = False Then
                'MsgBox ("EndOfDocument Bookmark doesn't exist")
                Call EndOfDocumentBookmark
                Call Run
                'ActiveDocument.Close
            Else
                Call Run
                'ActiveDocument.Close
        End If
        Application.ScreenUpdating = True
            StrFile = Dir
    Loop
    MsgBox ("Done!")
End Sub


Sub Run()


'If ActiveDocument.Bookmarks.Exists("EndOfDocument") = False Then
'    MsgBox ("EndOfDocument Bookmark doesn't exist")
'    End
'End If

Call NEWESTUpdateBringToNewDoc
Documents(strMainDocLong).Close SaveChanges:=wdDoNotSaveChanges
'MsgBox ("hi")

End Sub

Sub StartUp()
Dim ALLPlaces As String

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

Sub NEWESTUpdateBringToNewDoc()
'
' BringToNewDoc Macro
'

Call StartUp
Call DestroyTabThings

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
        
    'gets Document Numnber/subject of the document for TITLE portion
        Set MyData = New DataObject
        strMainDocLong = ActiveDocument.Name
        'strMainDoc = Left(strMainDocLong, Len(strMainDocLong) - 5) 'variable change for docx doc length document name
              'strMainDoc = Left(strMainDocLong, 13) 'change back, one time use, onetime, fix,
    '/gets...Document Numnber/subject portion
    
    'Get the Document Title/Title of the document
        DocTitle = Dialogs(wdDialogFileSummaryInfo).subject
        DocNumber = Dialogs(wdDialogFileSummaryInfo).Title
        
    '/get the Document title/Tile of the document
              
              
              
              
    'opens template
    Documents.Add Template:= _
        "G:\Common\Jerry\Controlled Documents\Document Conversion\Template QA14.dotx" _
        , NewTemplate:=False, DocumentType:=0
    '/open template
    
    With Dialogs(wdDialogFileSummaryInfo) 'makes title of new doc to the controlled doc title
     .Title = DocTitle
     .subject = DocNumber
     .Execute
    End With
    '****select title and change to black color

    strNewDoc = ActiveDocument.Name 'variable for the newer document, will be a .docx
    Documents(strMainDocLong).Activate
    
    Call TranscribeData
    

    Windows(strNewDoc).Activate
    
    Call NewApplyHeadingStyles
    Call NewSaveDocument
    
    Call NewDeleteEmptyParagraphs
    Call NewFixParagraphIndents
    Call SpaceingBeforeAfterPara
    
    ActiveDocument.Save
    ActiveDocument.Close
    
    
    
'    ActiveWindow.ActivePane.LargeScroll down:=1
'
'    strNewDocLong = DocNumber & ".docx"
'
'    Call NewParagraphFixer
'
'    Call Tabs

'
'    Call FindAndReplaceDoubleSpaces
'
'    Call RevisionsTableCount
'    'Call AddParagrahToTitles
'    Call deleteEmptyParagraphs
'    Call FixParagraphIndents
'    Call EndOfOldEmersonForamtDocument
'
'    ActiveDocument.Save
    
End Sub


Sub TranscribeData()
Dim strStartCopy
Dim strEndCopy
Dim i
Dim Bookmark
Dim blnFound As Boolean
Dim rng1 As Range
Dim rng2 As Range
Dim rngfound
Dim blnDoSeven As Boolean

blnDoSeven = True

Call RangedNumbersToText

For i = 1 To 8
    If i = 1 Then
        strStartCopy = "Purpose:"
        strEndCopy = "Scope:"
        Bookmark = "Purpose"
    End If
    If i = 2 Then
        strStartCopy = "Scope:"
        strEndCopy = "Terms And Definitions:"
        Bookmark = "Scope"
    End If
    If i = 3 Then
        strStartCopy = "Terms And Definitions:"
        strEndCopy = "Procedure Body:"
        Bookmark = "TermsAndDefinitions"
    End If
    If i = 4 Then
        strStartCopy = "Procedure Body:"
        strEndCopy = "Responsibilities:"
        Bookmark = "ProcedureBody"
    End If
    If i = 5 Then
        strStartCopy = "Reference:"
        strEndCopy = "Revisions:"
        Bookmark = "Reference"
    End If
    If i = 6 Then
        strStartCopy = "Responsibilities:"
        strEndCopy = "Reference:"
        Bookmark = "Responsibilities"
    End If
    
    
    'CHANGE HERE!!!!!!!!!
    If i = 7 Then
        Documents(strMainDocLong).Activate
        Selection.Find.Text = "Flow Chart/Turtle"
        blnFound = Selection.Find.Execute
        
        If blnFound Then
            
            strStartCopy = "Flow Chart:"
            strEndCopy = "Revisions:"
            Bookmark = "FlowChart"
            
            Else: blnDoSeven = False
        End If
       
        
        
        
        Documents(strNewDoc).Activate
    End If
    
    'Goto "endOfDocument" bookmark in the old template.
    'Selects to the last page, copies, and pastes into new document
    If i = 8 And ActiveDocument.Bookmarks.Exists("EndOfDocument") Or ActiveDocument.Bookmarks.Exists("EndOfDoc") Then
        blnDoSeven = True
        Documents(strMainDocLong).Activate
        If ActiveDocument.Bookmarks.Exists("EndOfDocument") Then
            Selection.GoTo What:=wdGoToBookmark, Name:="EndOfDocument"
        End If
        
        If ActiveDocument.Bookmarks.Exists("EndOfDoc") Then
            Selection.GoTo What:=wdGoToBookmark, Name:="EndOfDoc"
        End If
        
        Selection.EndKey Unit:=wdStory, Extend:=wdExtend
        Selection.Copy
        Documents(strMainDocLong).Activate
        
        If ActiveDocument.Bookmarks.Exists("EndOfDocument") Then
            Selection.GoTo What:=wdGoToBookmark, Name:="EndOfDocument"
        End If
        
        If ActiveDocument.Bookmarks.Exists("EndOfDoc") Then
            Selection.GoTo What:=wdGoToBookmark, Name:="EndOfDoc"
        End If
        
        'Selection.GoTo What:=wdGoToBookmark, Name:="EndOfDocument"
        
        Documents(strNewDoc).Activate
        Selection.GoTo What:=wdGoToBookmark, Name:="EndOfDoc"
        Selection.PasteAndFormat (wdPasteDefault)
        'Application.ScreenUpdating = True
    End If

    If i <> 8 And blnDoSeven = True Then
        'Application.ScreenUpdating = False
        Selection.HomeKey wdStory
        Selection.Find.Text = strStartCopy
        blnFound = Selection.Find.Execute
        If blnFound Then
            Selection.MoveRight wdWord 'was left
            Set rng1 = Selection.Range
                'If Selection.Range <> "" Then
                    Selection.Find.Text = strEndCopy
                    blnFound = Selection.Find.Execute
                    If blnFound Then
                        Set rng2 = Selection.Range
                        If rng2 <> "" Then
                            If (rng1.Start + 1) < (rng2.Start - 5) Then
                            Set rngfound = ActiveDocument.Range(rng1.Start + 1, rng2.Start - 5)
                                rngfound.Select
                                'Application.ScreenUpdating = True
                                Selection.Copy
                                Documents(strNewDoc).Activate
                                Selection.GoTo What:=wdGoToBookmark, Name:=Bookmark
                                Selection.PasteAndFormat (wdPasteDefault) 'wdFormatSurroundingFormattingWithEmphasis
                                Documents(strMainDocLong).Activate
                                'Application.ScreenUpdating = True
                            End If
                        End If
                    End If
                'End If
            'Was annoying not having nice lines
        End If
    End If
Next

End Sub

Sub NewApplyHeadingStyles()

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
    
    Call RangedNumbersToText
    
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
    If i = 6 Then strSection = "REFERENCE:"
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
Sub RangedNumbersToText()
Dim myDoc
Dim myRange As Object
Dim iParCount As Integer

'based on number of paragraphs in the template. This will change if
'you change footer or number of paragraphs after "revisions"
'section (including footer of last page)


iParCount = ActiveDocument.Paragraphs.Count - 11
'there seems to be 11 paragraphs before the
'footer at the end and "revisions section
Set myRange = _
ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(1).Range.Start, _
    End:=ActiveDocument.Paragraphs(iParCount).Range.End)
myRange.ListFormat.ConvertNumbersToText wdNumberParagraph

End Sub

Sub NewSaveDocument()
    Dim SaveName
    
    Call StartUp
    
    SaveName = Dialogs(wdDialogFileSummaryInfo).subject
    
    
    ChangeFileOpenDirectory _
        SaveToLocation
    ActiveDocument.SaveAs FileName:= _
        SaveToLocation & SaveName & ".docx" _
        , FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False

End Sub

Sub DestroyTabThings()

    Selection.WholeStory
    Selection.ParagraphFormat.TabStops.Add Position:=InchesToPoints(3.13), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = InchesToPoints(0.5)

End Sub

Sub NewFixParagraphIndents()

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



Sub NewDeleteEmptyParagraphs()

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

Sub SpaceingBeforeAfterPara()

Dim Para As Word.Paragraph


For Each Para In ActiveDocument.Paragraphs
If Para.SpaceBefore = 12 Then Para.SpaceBefore = 6
If Para.SpaceAfter = 12 Then Para.SpaceAfter = 6
If Para.Alignment = wdAlignParagraphJustify Then Para.Alignment = wdAlignParagraphLeft

Next

    Selection.Find.ClearFormatting
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
    Selection.Find.Execute
    Selection.MoveDown Unit:=wdLine, Count:=2
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="EndOfDocument"
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
End Sub
