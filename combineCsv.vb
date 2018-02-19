'''''''''''COMBINING A BUNCH OF CSV FILES BY CLEANING AND VALIDATION INTO A SINGLE CSV FILE TO BE IMPORTED INTO NEO4J GRAPH DATABASE

Sub combineFiles()

ThisWorkbook.Activate
Dim rng1 As Range
Dim rng2 As Range

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Dim obj_fso As Object, obj_dir As Object, obj_file As Object

Dim s_directory As String, s_fileName As String, s_ticker As String

s_directory = "...N\ABC Fund\Residual" ''"C:\Users\name\Documents\Entity Stucture"
'
'
'
'

''''''OPEN THE COMB FILE  '''''''
Workbooks("CombinedEntities.csv").Activate


'
'ThisWorkbook.Sheets("FileName").Select
    
        '''''CHECK IF EXISTS WITH THE NAME OR TICKER

    
    Set obj_fso = CreateObject("Scripting.FileSystemObject")
    Set obj_dir = obj_fso.GetFolder(s_directory)
    ret = False
    
        For Each obj_file In obj_dir.Files
        
          s_fileName = Replace(obj_file, s_directory, "")
          
                If obj_fso.fileExists(s_directory & "" & s_fileName) = True Then
                  
                  
                  
                  s_ticker = Replace(Replace(obj_file, s_directory & "\", ""), ".xlsx", "")
                  
                  If InStr(s_ticker, "_") > 0 Then
                    s_ticker = Left(s_ticker, Application.WorksheetFunction.Find("_", s_ticker) - 1)
                  End If
                  
'
                If InStr(Replace(obj_file, s_directory, ""), "-") = 0 And InStr(Replace(obj_file, s_directory, ""), "_") = 0 Then '' no - or _ in the filename
                  
                  On Error GoTo DoTask
                  
                  '''''CAN error and go to the task else go to next file
                  a = Application.WorksheetFunction.VLookup(Trim(s_ticker), Workbooks("CombinedEntities.csv").Sheets(1).Range("F:F"), 1, False)
                  
                  GoTo ContinueNext
                  '''''''''''''CHECK THE IN TICKER WITH THE EXISTING TICKERS FOR DUPLICATE
                  '''''''''USING VLOOKUP''''''''''''''''''
              '    If "NOT FOUND" = Application.WorksheetFunction.IfError(Application.WorksheetFunction.VLookup(s_ticker, Workbooks("CombinedEntities.xlsx").Sheets(1).Range("F:F"), 1, False), "NOT FOUND") Then
DoTask:
                              ''''''OPEN THE FILES '''''''
                              Workbooks.Open obj_file
                              Workbooks(Replace(obj_file, s_directory & "\", "")).Activate
                              ''''''''''GET THE ROW COUNT
                              SourceCount = Application.WorksheetFunction.CountA(Range("A:A"))
                                
                            If InStr(Range("D2").Value, "-") > 0 Then  ' identifier vheck
                                  
                                    '''''SELECT THE SOURCE RANGE 1
                                    Set rng1 = Range("B2:D" & SourceCount)
                                    '''''SELECT THE SOURCE RANGE 2
                                    Set rng2 = Range("F2:G" & SourceCount)
                                            
                                    ''''''''IN THE OPENED FILE
                                    'Workbooks(Replace(obj_file, s_directory, "")).Activate
                                    Workbooks("CombinedEntities.csv").Activate
                                    DestCount = Application.WorksheetFunction.CountA(Range("C:C")) + 1
                                    '''''''PASTE RANGE 1
                                    rng1.Copy
                                    Workbooks("CombinedEntities.csv").Sheets(1).Range("A" & DestCount).PasteSpecial xlPasteValues
                                    '''''''PASTE RANGE 2
                                    rng2.Copy
                                    Workbooks("CombinedEntities.csv").Sheets(1).Range("D" & DestCount).PasteSpecial xlPasteValues
                                    
                                    '''''''''SAVE THE TICKER '''''''
                                    
                                    Range("F" & DestCount).Value = Trim(s_ticker)
                        
                        
                            End If
                            
                            '''CLOSE THE FILE W/O SAVING
                            Workbooks(Replace(obj_file, s_directory & "\", "")).Close False
                            
                            If Err.Number > 0 Then
                              Resume ContinueNext
                              Err.Number = 0
                            End If
                  
                   'Exit For
ContinueNext:
                   
                 End If
                 End If
                ' Workbooks("CombinedEntities.csv").Save
                 
        Next

    Set obj_fso = Nothing
       Set obj_dir = Nothing
    End Sub
