Sub compileIntoTable()
'T Osborne
'30/09/20 - v1.0.0 - First issue
'30/09/20 - v1.0.1 - Updated to add loop for array build

'Input data variables.
Dim firstCol As Long        'First column number (E.g Col C = column 3)
Dim firstRow As Long        'First row number
Dim nCol As Long            'Number of columns in the table
Dim nRow As Long            'Number of rows in the table
Dim lastCol As Long         'Last column
Dim lastRow As Long         'Last row
Dim inputNRow As Long       'Row containing input speed values
Dim outputNCol As Long      'Column containing output speed values
Dim nGears As Integer       'Number of gears and hence worksheets
Dim curGear As Integer      'Current gear

'Initialise input data variables.
firstCol = 3
firstRow = 7
nCol = 48
nRow = 48
lastCol = firstCol + nCol - 1
lastRow = firstRow + nRow - 1
inputNRow = firstRow - 1
outputNCol = firstCol - 1
nGears = 7
curGear = 1

'Output array variables.
Dim nShifts                 'Number of shifts
Dim inputN                  'Input speed
Dim outputN                 'Output speed
Dim SS                      'Sliding speed
Dim W                       'Friction work
Dim P                       'Friction power
Dim testPointArr(nCol)      'Output array
Dim outputTableRow As Long  'Current output row
Dim outputTableRow_ As Long 'First output row
Dim outputTableCol As Long  'Current output column
Dim outputTableCol_ As Long 'First output column

'Initialise output array variables
outputTableRow_ = 4
outputTableRow = outputTableRow_
outputTableCol_ = 1
outputTableCol = outputTableCol_

'Synchroniser Limits.
Dim SSLim As Long
Dim WLim As Long
Dim PLim As Long
Dim SSPercent As Long
Dim WPercent As Long
Dim PPercent As Long
Dim waitTime As Long
Dim thresholdPercent As Long
Dim thresholdSpeed As Long

'Loop through each gear dataset.
Dim k As Integer
For k = 1 To nGears
    
    'Define the correct worksheets to use.
    Set wsInput = Sheets(CStr(k))
    Set wsInputSS = Sheets(CStr(k) & " SS")
    Set wsInputW = Sheets(CStr(k) & " W")
    Set wsInputP = Sheets(CStr(k) & " P")
    Set wsTable = Sheets(CStr(k) & " T")    
    
    ' Loop through each column on input data.
    ' Looping through columns first as we want to order output data by the input speed set point.
    Dim i As Integer
    For i = firstCol To lastCol
    
        'Loop through each row within the column.
        Dim j As Integer
        For j = firstRow To lastRow

            'Only process the cells that contain data.
            nShifts = wsInput.Cells(j, i).Value
            If nShifts <> "NaN" Then
                
                'Find the relevant data and add to variables.
                SS = wsInputSS.Cells(j + 1, i + 1).Value
                W = wsInputW.Cells(j + 1, i + 1).Value
                P = wsInputP.Cells(j + 1, i + 1).Value
                inputN = wsInput.Cells(inputNRow, i).Value
                outputN = wsInput.Cells(j, outputNCol).Value
                
                'SS limit check.
                If curGear < 4 Then
                    SSLim = 3201.5 '89mm
                Else
                    SSLim = 3665.5 '78mm
                End If
                
                SSPercent = (SS / SSLim) * 100
                            
                'Friction work limit check.
                WLim = 1.5
                WPercent = (W / WLim) * 100
                
                'Friction Power Limit check.
                PLim = 10
                PPercent = (P / PLim) * 100
                
                'Wait time calculation.
                thresholdPercent = 80 'Percent.
                thresholdSpeed = 2000 'rpm 10s below limit, 7s above limit.
                
                If inputN <= thresholdSpeed Or SSPercent > thresholdPercent Or WPercent > thresholdPercent Or PPercent > thresholdPercent Then
                    waitTime = 10
                Else
                    waitTime = 7
                End If                                
                    
                'Build the array - probably a better way to do this part!
                testPointArr(0) = inputN
                testPointArr(1) = Round(outputN, 0)
                testPointArr(2) = nShifts
                testPointArr(3) = Round(SS, 2)
                testPointArr(4) = Round(W, 2)
                testPointArr(5) = Round(P, 2)
                testPointArr(6) = waitTime
                testPointArr(7) = SSPercent
                testPointArr(8) = WPercent
                testPointArr(9) = PPercent
                
                'File each array value into the output table
                Dim n As Integer
                For n = 0 To n = 9
                    wsTable.Cells(outputTableRow, outputTableCol).Value = testPointArr(n)
                    outputTableCol = outputTableCol + 1
                Next n
                            
                'reset column and select next row
                outputTableCol = outputTablecol_
                outputTableRow = outputTableRow + 1                
            
            End If
            
        Next j
    
    Next i

    'Show progress
    Debug.Print k & ", " & wsInput.Name & ", " & wsInputSS.Name & ", " & wsInputW.Name & ", " & wsInputP.Name & ", " & wsTable.Name
    
    'reset output table for next gear
    outputTableRow = outputTableRow_
    outputTableCol = outputTableCol_
    
Next k

MsgBox "Complete"
End Sub