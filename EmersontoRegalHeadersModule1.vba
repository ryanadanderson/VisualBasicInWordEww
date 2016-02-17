Option Explicit
Public EmersonToRegalDoc As Document
Public FindText
Public ReplaceText
Public doc As Document
Public myPath As String
Public myFile As String
Public myExtension As String
Public FldrPicker As FileDialog
Public hdr As HeaderFooter
Public ftr As HeaderFooter
Public wb As Word.Application
Public wdSeekCurrentFooter As String
Public revision
Public PlzSearchForText




Sub LoopAllExcelFilesInFolder2()

'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'TASK: 1) replace header with Regal header. 2) To replace one document with another

'SOURCE: www.TheSpreadsheetGuru.com

    'Set wb = CreateObject("word.application")
    Set EmersonToRegalDoc = ActiveDocument

'Optimize Macro Speed
  Application.ScreenUpdating = False
    FindText = ActiveDocument.FormFields(1).Result
    ReplaceText = ActiveDocument.FormFields(2).Result
'Unprotect the file
    If ActiveDocument.ProtectionType <> wdNoProtection Then
      ActiveDocument.Unprotect Password:=""
    End If

'Retrieve Target Folder Path From User
Beginning:
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  'myExtension = "*.docx" 'change here for docx files and comment the next line out
  myExtension = "*.doc"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
        'Set variable equal to opened document
        Set doc = Documents.Open(FileName:=myPath & myFile)
        
        'Do the script - put any repeating tasks below
        'Call ConvertFile
    '    Call GetHeader
    '    Call PutHeader
        'Call DocProperties 'gets revision, title, and doc number
        'Call GetFooter 'gets footer from "Emerson to Regal Headers.docm"
        'Call PutFooter
    '    Call FormatStuff
    '    Call ColorFooterExpDate
        If PlzSearchForText = True Then Call Replace
        doc.Save
        doc.Close SaveChanges:=True
    
        'Get next file name
        myFile = Dir
  Loop


ResetSettings:
  'Reset Macro Optimization Settings
    Application.ScreenUpdating = True
    
If EmersonToRegalDoc.ProtectionType = wdNoProtection Then
    EmersonToRegalDoc.Activate
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    EmersonToRegalDoc.Protect _
    Type:=wdAllowOnlyFormFields, NoReset:=True, Password:=""
End If

End Sub
Sub ConvertFile()


On Error Resume Next

doc.Activate
ActiveDocument.Convert

    ActiveDocument.Convert

    If Right(doc, 3) = "doc" Then
    ActiveDocument.SaveAs FileName:=Left(doc, Len(doc) - 4) _
        , FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False
    End If
    
End Sub

Sub PutHeader()
    With doc
    doc.Activate
            For Each hdr In ActiveDocument.Sections(1).Headers
                hdr.Range.Text = vbNullString
                Next hdr
                
                If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
                    .ActiveWindow.ActivePane.View.Type = wdPrintView
                Else
                    .ActiveWindow.View.Type = wdPrintView
                End If

                .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
                    
                Selection.Delete Unit:=wdCharacter, Count:=1
                Selection.Delete Unit:=wdCharacter, Count:=1
                Selection.Delete Unit:=wdCharacter, Count:=1
                
                .ActiveWindow.ActivePane.Selection.Paste
    End With
                Selection.Delete Unit:=wdCharacter, Count:=1
    
    'TabsStop Formatting
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = InchesToPoints(0.5)
    Selection.ParagraphFormat.TabStops.Add Position:=InchesToPoints(6.5), _
        Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        
    'Lower Boarder
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    '''END HEADER
End Sub

Sub GetFooter()

        '''
        EmersonToRegalDoc.Activate
        
        If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
            ActiveWindow.Panes(2).Close
        End If
        If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
            ActivePane.View.Type = wdOutlineView Then
            ActiveWindow.ActivePane.View.Type = wdPrintView
        End If
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
        Selection.WholeStory
        
        'Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        Selection.Copy
        'ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

End Sub

Sub Replace()
        'set find text
'        FindText = EmersonToRegalDoc.FormFields(1).Result
'        ReplaceText = EmersonToRegalDoc.FormFields(2).Result
        
        If FindText <> "" And ReplaceText <> "" And FindText <> ReplaceText Then
        
        'initiate the action
        'Replace ( string1, find, replacement, [start, [count, [compare]]] )
            With ActiveDocument.Range.Find
              .Text = FindText
              .Replacement.Text = ReplaceText
              .Replacement.ClearFormatting
              .Replacement.Font.Italic = False
              .Forward = True
              .Wrap = wdFindContinue
              .Format = False
              .MatchCase = False
              .MatchWholeWord = False
              .MatchWildcards = False
              .MatchSoundsLike = False
              .MatchAllWordForms = False
              .Execute Replace:=wdReplaceAll
            End With
        End If
End Sub

Sub GetHeader()
  'Get Header from main document
  'make sub to get header?
    EmersonToRegalDoc.Activate
    
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.WholeStory
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Copy
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
  'end sub to get header from main document
End Sub

Sub PutFooter()

Dim ftr As HeaderFooter

 With doc
            '''Footer
            doc.Activate
            For Each ftr In ActiveDocument.Sections(1).Footers
                ftr.Range.Text = vbNullString
                Next ftr
                
                If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
                    .ActiveWindow.ActivePane.View.Type = wdPrintView
                Else
                    .ActiveWindow.View.Type = wdPrintView
                End If

                .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
                    
                Selection.Delete Unit:=wdCharacter, Count:=1
                Selection.Delete Unit:=wdCharacter, Count:=1
                Selection.Delete Unit:=wdCharacter, Count:=1
                
                .ActiveWindow.ActivePane.Selection.Paste
            
                'Selection.Delete Unit:=wdCharacter, Count:=1
                    
                'Footer work for converting Emerson to Regal
                
                'Top border black
                
                'Document Number
                'Title ^t Page X of N
                'Effective Date: DD/MM/YYYY ^t Rev. NN ^t Expires on MM/DD/YYYY
                
                'espires on MM/DD/YYYY is in red.
                
                
                    Selection.HomeKey Unit:=wdStory
                    With Selection.Borders(wdBorderTop)
                        .LineStyle = Options.DefaultBorderLineStyle
                        .LineWidth = Options.DefaultBorderLineWidth
                        .Color = Options.DefaultBorderColor
                    End With
                
                'Title line and page number tabs
                    Selection.HomeKey Unit:=wdStory
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    Selection.ParagraphFormat.TabStops.ClearAll
                    ActiveDocument.DefaultTabStop = InchesToPoints(0.5)
                    Selection.ParagraphFormat.TabStops.Add Position:=InchesToPoints(6.5), _
                        Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
                        
                        
                'Last line in footer
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    Selection.ParagraphFormat.TabStops.ClearAll
                    ActiveDocument.DefaultTabStop = InchesToPoints(0.5)
                    
                    Selection.ParagraphFormat.TabStops.Add Position:=InchesToPoints(6.5), _
                        Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
                    Selection.ParagraphFormat.TabStops.Add Position:=InchesToPoints(3.25), _
                        Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
                        
                'format footer
                    'font size 8
                    Selection.WholeStory
                    Selection.Font.Size = 8
                    
                    'make expires on date red
                    Selection.EndKey Unit:=wdStory
                    Selection.MoveLeft Unit:=wdWord, Count:=2, Extend:=wdExtend
                    Selection.Font.Color = wdColorRed
                    Selection.MoveLeft Unit:=wdWord, Count:=4
                    'type revision here
                    Selection.TypeText (revision)
                    Selection.MoveLeft Unit:=wdWord, Count:=4
                    Selection.TypeText Format(Date, "MM/DD/YYYY")
    End With
End Sub

Sub DocProperties()
    Dim i As Integer
    Dim metaprop As MetaProperty
    'On Error Resume Next

    
'    MsgBox (Dialogs(wdDialogFileSummaryInfo).Title)
'    MsgBox (Dialogs(wdDialogFileSummaryInfo).Subject)

    'changes the revision up a number
    
    revision = ActiveDocument.ContentTypeProperties.Item(3).Value + 1
    
    If Len(revision) = 1 Then revision = "0" & revision

    
    For Each metaprop In ActiveDocument.ContentTypeProperties
        If metaprop.Name = "Title" Then
'        MsgBox ("Title")
        Dialogs(wdDialogFileSummaryInfo).Title = metaprop.Value
        End If
        
        If metaprop.Name = "Document No." Then
'        MsgBox ("Doc Number")
        Dialogs(wdDialogFileSummaryInfo).Subject = metaprop.Value
        End If
        
'        If metaprop.Name = "Description" Then
'        metaprop.Value
'        End If
        
    Next
    
    With Dialogs(wdDialogFileSummaryInfo) 'makes title of new doc to the controlled doc title
     .Subject = Left(doc, Len(doc) - 5)
     .Execute
    End With
    
    
End Sub

Sub FormatStuff()
Dim para


ActiveDocument.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

On Error Resume Next

'change space before paragraphs to size 6 to save space
For Each para In ActiveDocument.Paragraphs
    If para.SpaceBefore = 12 Then para.SpaceBefore = 6
    If para.SpaceAfter = 12 Then para.SpaceAfter = 6
    If para.Alignment = wdAlignParagraphJustify Then para.Alignment = wdAlignParagraphLeft
Next

'change text/typeface
    Selection.WholeStory
    Selection.Font.Size = 10
    Selection.Font.Name = "Arial"


'Remove Annual Mgt Review in ECN revision table
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

'Add next line in paragraph for revisions
'RYAN





ActiveDocument.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Selection.WholeStory
    Selection.Font.Size = 8
    Selection.Font.Name = "Arial"
    Selection.Font.Color = wdColorBlack
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With

End Sub

Sub NewTableStuff()
' activedocument.tables.count doesn't work. WHHHY stacked overflow issue submitted
 
 Dim tblNew As Table
 Dim rowNew As Row
 Dim celTable As Cell
 Dim intCount As Integer
 
 MsgBox (ActiveDocument.Tables.Count)
 
' intCount = 1
' Set tblNew = ActiveDocument.Tables(3)
' Set rowNew = tblNew.Rows.Add(BeforeRow:=tblNew.Rows(1))
' For Each celTable In rowNew.Cells
' celTable.Range.InsertAfter Text:="Cell " & intCount
' intCount = intCount + 1
' Next celTable

End Sub

Sub ColorFooterExpDate()
'
' Colors the footer expiration text to red
'
'
    WordBasic.ViewFooterOnly
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "expires on"
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
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    Selection.Font.Color = wdColorRed
    ActiveDocument.Save
End Sub
