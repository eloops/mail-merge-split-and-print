Sub SplitMergeLetterToPrinter()
  ' Macro to split a mail merged document up and print by sections
  Dim sPrinter As String
  Dim Letters As Long
  Dim Counter As Long

  Letters = ActiveDocument.Sections.Count
  ' Because the target document actually has a continuous section break within it
  Counter = 0
  With Dialogs(wdDialogFilePrintSetup)
    ' Grab the currently selected printer
    sPrinter = .Printer
    ' Change to the target printer but don't set it as the default
    .Printer = "Print Queue Name Here"
    .DoNotSetAsSysDefault = True
    ' Turn off those stupid outside margin alerts
    Application.DisplayAlerts = wdAlertsNone
    .Execute
    
    ' Change the counter values if there are no existing sections within the document
    While Counter < Letters
      ActiveDocument.PrintOut Background:=False, _
      Range:=wdPrintFromTo, _
      From:="s" & Format(Counter - 1), To:="s" & Format(Counter)
      Counter = Counter + 1
    Wend

    ' Set everything back to normal
    .Printer = sPrinter
    Application.DisplayAlerts = wdAlertsAll
  End With
End Sub
