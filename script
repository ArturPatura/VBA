Sub OpenTextFileRead()
TextFile_FindReplace

Dim fs, a, line, counter, letter, IloscSlow, slowo
Dim words() As String
counter = 3
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.OpenTextFile("C:\Users\Dell\Desktop\te.gcode", ForReading, False)
Do While a.AtEndOfStream <> True
    line = a.ReadLine
    FirstWord = Mid(line, 1, 2)
    If FirstWord = "G1" Then
        words() = Split(line)
        IloscSlow = UBound(words)
        For i = 2 To 6
             
            For slowo = 1 To IloscSlow
                If Cells(1, i) = Mid(words(slowo), 1, 1) Then
                    If Mid(words(slowo), 1, 1) = (Cells(1, 2)) Or Mid(words(slowo), 1, 1) = (Cells(1, 3)) Or Mid(words(slowo), 1, 1) = (Cells(1, 4)) Then
                            Cells(counter, i) = Mid(words(slowo), 2)
                    
                    ElseIf Mid(words(slowo), 1, 1) = (Cells(1, 5)) Or Mid(words(slowo), 1, 1) = (Cells(1, 6)) Then
                        If Mid(words(slowo), 2) < 0 Then
                            Cells(counter, i) = 0
                        ElseIf Mid(words(slowo), 2) >= 0 Then
                            Cells(counter, i) = Mid(words(slowo), 2)
                   
            
                        End If
                    
                        If i = 5 Then
                        Cells(counter, i) = Cells(counter, i) + Cells(counter - 1, i)
                        End If
                    End If
                End If
            
            Next slowo
            
            If IsEmpty(Cells(counter, i)) Then
             Cells(counter, i) = Cells(counter - 1, i)
            End If
        
          Next i
        
        
            If (Cells(counter, 2) <> Cells(counter - 1, 2)) Or (Cells(counter, 3) <> Cells(counter - 1, 3)) Or (Cells(counter, 4) <> Cells(counter - 1, 4)) Then
                'dl
                Cells(counter, 7) = ((Cells(counter, 2) - Cells(counter - 1, 2)) ^ 2 + (Cells(counter, 3) - Cells(counter - 1, 3)) ^ 2 + (Cells(counter, 4) - Cells(counter - 1, 4)) ^ 2) ^ (0.5)
               
                'time
                Cells(counter, 1) = (Cells(counter, 7) / Cells(counter, 6)) * 60 + Cells(16, 18) + Cells(counter - 1, 1)
                
                'przyrost E
                Cells(counter, 9) = Cells(counter, 5) - Cells(counter - 1, 5)
                'przyrost time
                Cells(counter, 11) = Cells(counter, 1) - Cells(counter - 1, 1)
                
            
                'objetosc V
                  
                If Cells(counter, 9) = 0 Then
                    Cells(counter, 10) = 0
                Else
                     Cells(counter, 10) = Cells(11, 18) * Cells(12, 18) * Cells(12, 18) / 4 * Cells(counter, 9) / 1000
                End If
             
                
                'power
                
                If Cells(counter, 11) = 0 Then
                   Cells(counter, 8) = 0
                Else
                  Cells(counter, 8) = Cells(counter, 10) * Cells(13, 18) * Cells(14, 18) * Cells(15, 18) / Cells(counter, 11) * 1000
                End If
                
                Cells(counter - 1, 8) = Cells(counter, 8)
                
                counter = counter + 1
            End If
    
    End If
           
Loop
a.Close
        If Cells(counter, 2) = Cells(counter - 1, 2) And Cells(counter, 3) = Cells(counter - 1, 3) And Cells(counter, 4) = Cells(counter - 1, 4) Then
                Rows(counter).EntireRow.Delete
        End If

End Sub


Sub TextFile_FindReplace()

Dim TextFile As Integer
Dim FilePath As String
Dim FileContent As String

  FilePath = "C:\Users\Dell\Desktop\te.gcode"
  TextFile = FreeFile
  Open FilePath For Input As TextFile
  
  FileContent = Input(LOF(TextFile), TextFile)


  Close TextFile
  
'Find/Replace
  FileContent = Replace(FileContent, ";", " ;")

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file in a Write State
  Open FilePath For Output As TextFile
  
'Write New Text data to file
  Print #TextFile, FileContent

'Close Text File
  Close TextFile

End Sub


Sub SaveTextToFile()

    Dim FilePath As String
    FilePath = "C:\Users\Dell\Desktop\test.txt"

    Dim LastRow As Long
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim fileStream As TextStream
    Dim WordStream As String
    
    LastRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row


    Set fileStream = fso.CreateTextFile(FilePath)


    For i = 2 To LastRow
        
        WordStream = CStr(Round(Cells(i, 1), 4)) + "," + CStr(Round(Cells(i, 2), 4)) + "," + CStr(Round(Cells(i, 3), 4)) + "," + CStr(Round(Cells(i, 4), 4)) + "," + CStr(Round(Cells(i, 8), 4))
        fileStream.WriteLine WordStream
    
    Next i

    fileStream.Close

    If fso.FileExists(FilePath) Then
        MsgBox "The file was created!"
    End If

End Sub

