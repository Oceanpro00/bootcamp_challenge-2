Attribute VB_Name = "Module1"
Sub SortOut()

' Declare Objects

Dim ws As Worksheet
Dim LastRow As Long
Dim LastRowSummary As Long

Dim summaryList As Integer
Dim opening As Double
Dim closing As Double
Dim total As Double

Dim MaxPercentage As Double
Dim MinPercentage As Double
Dim MaxTotal As Double

' Worksheet loop

For Each ws In Worksheets

' Summary Table Headers Placement

    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Quarterly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"

' Last Row Function for For loops

    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
' Setting Initial Values for Summary List Input and Quaterly Opening

    summaryList = 2
    opening = ws.Cells(2, 3).Value

' Data set For Loop

    For i = 2 To LastRow
    
' If statement checking for whether NEXT value is not the same
 
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then

' Set closing price point and add final total

            closing = ws.Cells(i, 6).Value
            total = total + ws.Cells(i, 7).Value
            
' Calculate and input each Column using the counter

    ' Ticker
            ws.Cells(summaryList, 10).Value = ws.Cells(i, 1).Value
    ' Quaterly Change
            ws.Cells(summaryList, 11).Value = closing - opening
    ' Percent Change
            ws.Cells(summaryList, 12).Value = ((closing - opening) / opening)
    ' Total Stock Volume
            ws.Cells(summaryList, 13).Value = total
            
' Number Format Designation

            ws.Cells(summaryList, 11).NumberFormat = "0.00"
            
            ws.Cells(summaryList, 12).NumberFormat = "0.00%"

' If statement to set a Color Scale formating for the Quaterly Change Column

            If (ws.Cells(summaryList, 11).Value > 0) Then
            
                ws.Cells(summaryList, 11).Interior.ColorIndex = 4
                
            ElseIf (ws.Cells(summaryList, 11).Value < 0) Then
            
                ws.Cells(summaryList, 11).Interior.ColorIndex = 3
                
            End If
            
' Adding +1 to SummaryList so that the next Stock can be inputted below
            
            summaryList = summaryList + 1
            
' Setting next Opening Value and reseting total

            opening = ws.Cells(i + 1, 3).Value
            total = 0

' adding total value on each line where initial if isnt executed

        Else
        
            total = total + ws.Cells(i, 7).Value
            
            
' End the if and For loop for the Orginial Data set
            
        End If
        
    Next i
    
' Summary Last Row Function for Summary Table For Loop
    
    LastRowSummary = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
' Max/ Min Table Headers
    
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    
' Max/ Min Initial Values
    
    MaxPercentage = ws.Cells(2, 12).Value
    MinPercentage = ws.Cells(2, 12).Value
    MaxTotal = ws.Cells(2, 13).Value
    
' For Loop cycling through Summary Table
    
    For j = 2 To LastRowSummary
    
' If statement looking for Max/ Min Values
    
        If (ws.Cells(j, 12).Value > MaxPercentage) Then
        
            MaxPercentage = ws.Cells(j, 12).Value
            ws.Cells(2, 17).Value = ws.Cells(j, 10).Value
            ws.Cells(2, 18).Value = MaxPercentage
            
        ElseIf (ws.Cells(j, 12).Value < MinPercentage) Then
        
            MinPercentage = ws.Cells(j, 12).Value
            ws.Cells(3, 17).Value = ws.Cells(j, 10).Value
            ws.Cells(3, 18).Value = MinPercentage
        
        ElseIf (ws.Cells(j, 13).Value > MaxTotal) Then
        
            MaxTotal = ws.Cells(j, 13).Value
            ws.Cells(4, 17).Value = ws.Cells(j, 10).Value
            ws.Cells(4, 18).Value = MaxTotal
            
' End if and for loops
            
        End If
        
    Next j
    
' Reformate Number Format for Min/ Max Table
        
    ws.Cells(2, 18).NumberFormat = "0.00%"
    
    ws.Cells(3, 18).NumberFormat = "0.00%"
    
' Autofit Columns

    ws.Columns("A:R").AutoFit
    
Next ws

' finalizing MsgBox

MsgBox ("All Done")
    
End Sub
