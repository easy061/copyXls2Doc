Sub WorksheetLoop()

    Dim WS_Count As Integer
    Dim I As Integer

    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count

    'Dim tbl As Excel.Range
    Dim WordApp As Word.Application
    Dim myDoc As Word.Document
    Dim WordTable As Word.Table
    
'Create an Instance of MS Word
    On Error Resume Next
    
    'Is MS Word already opened?
      Set WordApp = GetObject(class:="Word.Application")
    
    'Clear the error between errors
      Err.Clear

    'If MS Word is not already open then open MS Word
      If WordApp Is Nothing Then Set WordApp = CreateObject(class:="Word.Application")
    
    'Handle if the Word Application is not found
      If Err.Number = 429 Then
        MsgBox "Microsoft Word could not be found, aborting."
        GoTo EndRoutine
      End If

    On Error GoTo 0

    'Optimize Code
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    'Create a New Document
    Set myDoc = WordApp.Documents.Add


' Begin the loop.
    For I = 1 To WS_Count

    ' Insert your code here.
    ' The following line shows how to reference a sheet within
    ' the loop by displaying the worksheet name in a dialog box.
    ' MsgBox ActiveWorkbook.Worksheets(I).Name
    

   
    'Copy Excel Used Range in WorkSheets[I]
    ActiveWorkbook.Worksheets(I).UsedRange.Copy
 
    'Make MS Word Visible and Active
    WordApp.Visible = True
    WordApp.Activate


   
    'Paste Table into MS Word
    myDoc.Paragraphs(I).Range.PasteExcelTable _
        LinkedToExcel:=False, _
        WordFormatting:=False, _
        RTF:=False

    'Autofit Table so it fits inside Word Document
    Set WordTable = myDoc.Tables(1)
    WordTable.AutoFitBehavior (wdAutoFitWindow)

    'myDoc.Paragraphs.Add _
    '    Range:=myDoc.Paragraphs(myDoc.Paragraphs.Count).Range
    'myDoc.Paragraphs(1).Range.MoveEnd Unit:=wdCharacter, Count:=-1
    'myDoc.Paragraphs(1).Range.InsertParagraphBefore
    'myDoc.Paragraphs(I).Range.MoveEnd Unit:=wdCharacter, Count:=-1
    'myDoc.Paragraphs(I).Range.InsertParagraphAfter
    'myDoc.Paragraphs(I).Range.InsertParagraphBefore
    myDoc.Paragraphs(I).Range.InsertParagraphBefore

    
    Next I

EndRoutine:
    'Optimize Code
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    'Clear The Clipboard
    Application.CutCopyMode = False
    
End Sub
