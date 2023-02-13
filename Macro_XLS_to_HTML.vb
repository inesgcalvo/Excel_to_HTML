Sub Macro_XLS_to_HTML()

'Define the variables.
Dim iRow As Long
Dim iPage As Integer
   
'Start on the 2nd row to avoid the 1st row, that contains the header.
iRow = 2

'Start the Loop
Do While WorksheetFunction.CountA(Rows(iRow)) > 0

'Create an .html file in the same directory as your active workbook
'Name the file as the first cell of the Row
Dim sFile As String
sFile = ActiveWorkbook.Path & "\" & Cells(iRow, 1) & ".html"
Close
   
'Print the HTML scaffolding
Open sFile For Output As #1
Print #1, "<html>"
Print #1, "<head>"
Print #1, "<style type=""text/css"">"
Print #1, "  body { font-size:12px;font-family:georgia } "
Print #1, "</style>"
Print #1, "</head>"
Print #1, "<body>"

'Translate the columns of the table
    If Not IsEmpty(Cells(iRow, 1)) Then
      Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 1).Value & "</text>"
            iPage = iPage + 1
    End If
 
'Each row is presented isolated in order to modify its characteristics individually if desired
    If Not IsEmpty(Cells(iRow, 2)) Then
        Print #1, "<br><h1 href=""" & iPage & ".html"">" & Cells(iRow, 2).Value & "</h1>"
            iPage = iPage + 1
    End If

    If Not IsEmpty(Cells(iRow, 3)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 3).Value & "</text>"
        iPage = iPage + 1
    End If
        
    If Not IsEmpty(Cells(iRow, 4)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 4).Value & "</text>"
        iPage = iPage + 1
    End If
            
    If Not IsEmpty(Cells(iRow, 5)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 5).Value & "</text>"
        iPage = iPage + 1
    End If

    If Not IsEmpty(Cells(iRow, 6)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 6).Value & "</text>"
            iPage = iPage + 1
    End If

    If Not IsEmpty(Cells(iRow, 7)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 7).Value & "</text>"
        iPage = iPage + 1
    End If
        
    If Not IsEmpty(Cells(iRow, 8)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 8).Value & "</text>"
        iPage = iPage + 1
    End If
            
    If Not IsEmpty(Cells(iRow, 9)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 9).Value & "</text>"
        iPage = iPage + 1
    End If
    
        If Not IsEmpty(Cells(iRow, 10)) Then
      Print #1, "<text href=""" & iPage & ".html"">" & Cells(iRow, 10).Value & "</text>"
            iPage = iPage + 1
    End If
      
    If Not IsEmpty(Cells(iRow, 11)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 11).Value & "</text>"
            iPage = iPage + 1
    End If

    If Not IsEmpty(Cells(iRow, 12)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 12).Value & "</text>"
        iPage = iPage + 1
    End If
        
    If Not IsEmpty(Cells(iRow, 13)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 13).Value & "</text>"
        iPage = iPage + 1
    End If
            
    If Not IsEmpty(Cells(iRow, 14)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 14).Value & "</text>"
        iPage = iPage + 1
    End If

    If Not IsEmpty(Cells(iRow, 15)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 15).Value & "</text>"
            iPage = iPage + 1
    End If

    If Not IsEmpty(Cells(iRow, 16)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 16).Value & "</text>"
        iPage = iPage + 1
    End If
        
    If Not IsEmpty(Cells(iRow, 17)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 17).Value & "</text>"
        iPage = iPage + 1
    End If
            
    If Not IsEmpty(Cells(iRow, 18)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 18).Value & "</text>"
        iPage = iPage + 1
    End If
            
    If Not IsEmpty(Cells(iRow, 19)) Then
        Print #1, "<br><text href=""" & iPage & ".html"">" & Cells(iRow, 19).Value & "</text>"
        iPage = iPage + 1
    End If
    
'Add ending HTML tags
    Print #1, "</body>"
    Print #1, "</html>"
    Close
'Add this line to the code if you want Excel to open all the files
'Shell "hh " & vbLf & sFile, vbMaximizedFocus
    
'Finish the Loop
iRow = iRow + 1
Loop

'Finish the Macro
End Sub
