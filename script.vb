Sub PrepareWorksheet()
TextFile_FindReplace

Dim fs, a, line, counter, letter, NumberOfTerms, term 
Dim words() As String

'licznik ustawiony od 3, ponieważ skrypt rozpoczyna szukanie od 3 linii
counter = 3				
Const ForReading = 1, ForWriting = 2, ForAppending = 8

'tworzenie obiektu fs- czyli file stream 
Set fs = CreateObject("Scripting.FileSystemObject")  

'utworzenie obiektu otwierającego kod maszynowy (GCODE) 
Set a = fs.OpenTextFile("C:\Users\Dell\Desktop\te.gcode", ForReading, False)

'tak długo, aż plik a nie bedzie pusty, to będzie sie wykonywać 
Do While a.AtEndOfStream <> True  

'sczytywanie linii
    line = a.ReadLine

'sprawdza czy jest w danej linii g1, (od pierwszej pozycji w linii, sprawdza dwie litery)
    FirstWord = Mid(line, 1, 2)
    If FirstWord = "G1" Then
   
'rozdziela wyrazy
	words() = Split(line) 

'przypisuje do zmiennej ilość słów      
	NumberOfTerms = UBound(words)  

'sprawdza kolumny od 2 do 6
        For i = 2 To 6

'sprawdza każde słowo po kolei         
            For term = 1 To NumberOfTerms  
                If Cells(1, i) = Mid(words(term), 1, 1) Then
                    If Mid(words(term), 1, 1) = (Cells(1, 2)) Or Mid(words(term), 1, 1) = (Cells(1, 3)) Or Mid(words(term), 1, 1) = (Cells(1, 4)) Then
                            Cells(counter, i) = Mid(words(term), 2)
                    ElseIf Mid(words(term), 1, 1) = (Cells(1, 5)) Or Mid(words(term), 1, 1) = (Cells(1, 6)) Then
'powyżej, jeśli pierwsze słowo danej linijki to współrzędne to są one wpisywane

               If Mid(words(term), 2) < 0 Then
                            Cells(counter, i) = 0
                        ElseIf Mid(words(term), 2) >= 0 Then
                            Cells(counter, i) = Mid(words(term), 2)
                        End If

'przyrost grubości filamentu (dodawane jest komórka(x) i komórka(x-1))                   
                        If i = 5 Then
                        Cells(counter, i) = Cells(counter, i) + Cells(counter - 1, i)
                        End If
                    End If
                End If
            
            Next term
'jeśli któraś komórka była pusta, to jest uzupełniana komórką wyżej
            If IsEmpty(Cells(counter, i)) Then
             Cells(counter, i) = Cells(counter - 1, i)
            End If
        
          Next i

'przy zmianie chociaż jednej ze współrzędnych(x,y,z) obliczane są konkretne kolumny: długość, czas, przyrost grubości, przyrost czasu, objętość i moc 
        
            If (Cells(counter, 2) <> Cells(counter - 1, 2)) Or (Cells(counter, 3) <> Cells(counter - 1, 3)) Or (Cells(counter, 4) <> Cells(counter - 1, 4)) Then

                'dl
                Cells(counter, 7) = ((Cells(counter, 2) - Cells(counter - 1, 2)) ^ 2 + (Cells(counter, 3) - Cells(counter - 1, 3)) ^ 2 + (Cells(counter, 4) - 
					Cells(counter - 1, 4)) ^ 2) ^ (0.5)
               
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
