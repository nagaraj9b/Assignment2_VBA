Attribute VB_Name = "Module1"
Sub Processtorunallsheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub

Public Sub RunCode()


For Each ws In Worksheets


' Define Variables volume and previousvolume as long
Dim volume, prevvol, i, opencounter As Long

' Define Variables tickers as string
Dim maxticker, minticker, maxvolticker As String

' Define variables as double
Dim openstock, Closestock, MaxYearChg, MinYearChg, MaxVol As Double

' Counter to calculate the unique tickers
Dim Counter, j As Integer


'Intializiation of variables
Counter = 0
volume = 0
prevvol = 0
openstock = 0
Closestock = 0#
opencounter = 0#

'Header Definition
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percentage Change"
Cells(1, 12) = "Total Stock Volume"
Cells(2, 16) = "Greatest % increase"
Cells(3, 16) = "Greatest % Decrease"
Cells(4, 16) = "Greatest Total Volume"
Cells(1, 17) = "Ticker"
Cells(1, 18) = "Value"

'Making Headers bold
Cells(1, 9).Font.Bold = True
Cells(1, 10).Font.Bold = True
Cells(1, 11).Font.Bold = True
Cells(1, 12).Font.Bold = True

' Making Autofit
Columns(10).AutoFit
Columns(11).AutoFit
Columns(12).AutoFit



' Determine the Last Row
lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        

        
'Create a loop to compare the tickers from one row  to next row
    For i = 2 To lastrow
       
    ' Comparing row1 to row2 if equals then checks condition to Total volume
    If (Cells(i + 1, 1).Value) = (Cells(i, 1).Value) Then
        volume = Cells(i, 7).Value + prevol
        prevol = volume
        opencounter = opencounter + 1
        If opencounter = 1 Then
        openstock = Cells(i, 3)
        End If
    End If
        
        
    ' Comparing row1 to row2 if not equal then checks condition to finalize the Total volume and keep counter on Tickers
    If (Cells(i + 1, 1).Value) <> (Cells(i, 1).Value) Then
         volume = Cells(i, 7).Value + prevol
        Counter = Counter + 1
        Closestock = Cells(i, 6)
        
        ' For loop is to write the totalvolums and corresponding tickers to right of data
          For j = Counter To Counter
          Cells(j + 1, 9).Value = Cells(i, 1)
          Cells(j + 1, 10).Value = Closestock - openstock
          
          'Color coding based on total volume
          If (Closestock - openstock) >= 0 Then
          Cells(j + 1, 10).Interior.ColorIndex = 4
          Else
          Cells(j + 1, 10).Interior.ColorIndex = 3
          End If
          
          
          'if the divisor is zero then show 0 value
          If (openstock = 0) Then
          Cells(j + 1, 11).Value = 0
          Else
          Cells(j + 1, 11) = ((Closestock - openstock) / openstock)
          End If
                    
          Cells(j + 1, 12).Value = volume
          Next j
    
    'reset Prevol counter
    prevol = 0
    opencounter = 0
    
    End If
    Next i
    
    MaxYearChg = 0
    MinYearChg = 0
    MaxVol = 0
    
    ' Loop to check maximum year change, Minimum year change and Maximum total volume
    For k = 2 To Counter + 1
     Cells(k, 11).NumberFormat = "0.00%"
     
     If Cells(k, 11).Value > MaxYearChg Then
     MaxYearChg = Cells(k, 11).Value
     maxticker = Cells(k, 9).Value
     End If
     
     If Cells(k, 11).Value < MinYearChg Then
     MinYearChg = Cells(k, 11).Value
     minticker = Cells(k, 9).Value
     End If
     
     If Cells(k, 12).Value > MaxVol Then
     MaxVol = Cells(k, 12).Value
     maxvolticker = Cells(k, 9).Value
     End If
     
      
     
     
     
    Next k
    
    'Hard - assigning max change, min change and max volume
    
    Range("q2") = maxticker
    Range("r2") = MaxYearChg
    
    Range("q3") = minticker
    Range("r3") = MinYearChg
    
    Range("q4") = maxvolticker
    Range("r4") = MaxVol
    
    
    Range("r2").NumberFormat = "0.00%"
    Range("r3").NumberFormat = "0.00%"
    
    
    
        
Next ws




End Sub

