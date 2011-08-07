' Word Constants
Const wdFormatPDF = 17  ' PDF format. 
Const wdFormatXPS = 18  ' XPS format. 
Const WdDoNotSaveChanges = 0

'Excel Constants
Const xlTypePDF = 0 ' PDF format. 
Const xlTypeXPS = 1 ' XPS format. 

'Powerpoint Constants
Const ppFixedFormatTypePDF = 2 ' PDF format. 
Const ppFixedFormatTypeXPS = 1 ' XPS format. 
Const ppWindowMinimized = 2
Const ppWindowNormal = 1
Const ppSaveAsPdf = 32

' Global variables
Dim arguments
Set arguments = WScript.Arguments

' ***********************************************
' CHECKARGS
'
' Makes some preliminary checks of the arguments.
' Quits the application is any problem is found.
'
Function CheckArgs()
  ' Check that <doc-file> is specified
  If arguments.Unnamed.Count <> 1 Then
    WScript.Echo "Error: Obligatory <doc-file> parameter missing!"
    WScript.Quit 1
  End If

  bShowDebug = arguments.Named.Exists("debug") Or arguments.Named.Exists("d")

End Function


Function TOPDF( sSourceFile, sPDFFile )
  Dim fos : Set fso = CreateObject("Scripting.FileSystemObject")
  sSourceFile = fso.GetAbsolutePathName(sSourceFile)

  ' WScript.Echo "Source file = '" + sSourceFile + "'"
  
  sFolder = fso.GetParentFolderName(sSourceFile)

  sPDFFile = fso.GetBaseName(sSourceFile) + ".pdf"
  sPDFFile = sFolder + "\" + sPDFFile
  ' WScript.Echo "PDF file = '" + sPDFFile + "'"

  Dim strFileExt : strFileExt = UCase(fso.GetExtensionName(sSourceFile))
  Select Case strFileExt
    Case "DOC", "DOCX", "DOT", "DOTX", "RTF"
      Call doc2pdf( sSourceFile, sPDFFile )
    Case "XLS", "XLSX"
      Call xls2pdf( sSourceFile, sPDFFile )
    Case "PPT", "PPTX"
      Call ppt2pdf( sSourceFile, sPDFFile )
    Case Else
      WScript.Echo strFileExt
  End Select
  Set fso = Nothing
End Function

Function doc2pdf( sSourceFile, sPDFFile )
  Dim wApp ' As Word.Application
  Dim wDoc ' As Word.Document

  Set wApp = CreateObject("Word.Application")

  on error resume next
  Set wDoc = WApp.Documents.Open(sSourceFile, False, nil, nil, "?#nonsense@$")
    Select Case Err.Number
      Case 0
        ' try and loop through all fields in the document and unlink them (convert them into text)
        For Each sr In wdoc.StoryRanges
          For Each oField in sr.Fields
            oField.Unlink
          Next
        Next
        ' Let Word document save as PDF
        ' - for documentation of SaveAs() method,
        '   see http://msdn2.microsoft.com/en-us/library/bb221597.aspx 

        wDoc.ExportAsFixedFormat SPDFFile, wdFormatPDF, False
        ' wDoc.SaveAs sPDFFile, wdFormatPDF
        Select Case Err.Number
          Case 0
            wdoc.Close WdDoNotSaveChanges
          Case 4605
            wdoc.Close WdDoNotSaveChanges
          Case 5125
            wdoc.Close WdDoNotSaveChanges
          Case Else
            WScript.Echo "save_problem"
        End Select
      Case 5408
        WScript.Echo "password"
      Case 4198
        WScript.Echo "popup"
      Case Else
        Wscript.Echo "problem"
    End Select

  wApp.Quit WdDoNotSaveChanges
  Set wApp = Nothing
End Function

Function xls2pdf( sSourceFile, sPDFFile )
  Dim eApp ' As Excel.Application
  Dim eBook ' As Excel.workbook

  Set eApp = CreateObject("Excel.Application")
  eApp.DisplayAlerts = False

  ' Open the excel workbooks
  on error resume next
  Set eBook = eApp.Workbooks.Open(sSourceFile, 2, , , "?#nonsense@$")
    Select Case Err.Number
      Case 0
      ' Excel save as PDF
      ' - for documentation of SaveAs() method
      '   see http://msdn.microsoft.com/en-us/library/bb238907.aspx
        eBook.ExportAsFixedFormat xlTypePDF, sPDFFile, , , , , , False
        Select Case Err.Number
          Case 0
            eBook.Close False
          Case 4605
            eBook.Close False
          Case Else
            WScript.Echo "save_problem"
          End Select
      Case Else
        Wscript.Echo "problem"
    End Select

  eApp.Quit
  Set eApp = Nothing
end Function

Function ppt2pdf( sSourceFile, sPDFFile )
  Dim pApp ' As Powerpoint.Application
  Dim pPresentation ' As Powerpoint presentation
  Dim oApp ' As Object
  On Error Resume Next
  Set oApp = GetObject(, "PowerPoint.Application")
  If Err.Number = 0 Then
    ' it is running so kick out
    WScript.Echo "powerpoint_running"
  Else
    Err.Clear()

    Set pApp = CreateObject("Powerpoint.Application")
    pApp.Visible = True
    pApp.WindowState = ppWindowMinimized

    Set pPresentation = pApp.Presentations.Open(sSourceFile)
    on error resume next
    ' - for documentation of SaveAs() method,
    '   see http://msdn.microsoft.com/en-us/library/bb238907.aspx 
    'pPresentation.ExportAsFixedFormat sPDFFile, ppFixedFormatTypeP
    pPresentation.SaveAs sPDFFile, ppSaveAsPDF
    Select Case Err.Number
      Case 0
        pPresentation.Close False
      Case 4605
        pPresentation.Close False
      Case Else
        WScript.Echo "save_problem"
    End Select

    pApp.Quit
    Set pApp = Nothing
  End If
  oApp = Nothing

end Function

' *** MAIN **************************************

Call CheckArgs()
Call TOPDF( arguments.Unnamed.Item(0), arguments.Named.Item("o") )

Set arguments = Nothing