Public Sub deltaSpeedAdjust()
' Macro Created by T Osborne on 12th November 2020

'File directory: -
'\\DSUK01\Company Shared Documents\Projects\1306\DATA EXCHANGE\HPP\Received\201007 - Gearchange Duty Heat Maps\

'File Name: -
' dataProcessing_deltaSpeedAdjust.xlsx

'Delta speed at each synchro
'1  syncDeltaSpeed = (InputN * (18/52)) - (OutputN * (35/11) * (39/28))
'       inputN = (syncDeltaSpeed / (18/52)) + (OutputN * (35/11) * (39/28))
'
'2  syncDeltaSpeed = (InputN * (22/46)) - (OutputN * (35/11) * (39/28))
'       inputN = (syncDeltaSpeed / (22/46)) + (OutputN * (35/11) * (39/28))
'
'3  syncDeltaSpeed = (InputN * (27/43)) - (OutputN * (35/11) * (39/28))
'       inputN = (syncDeltaSpeed / (27/43)) + (OutputN * (35/11) * (39/28))
'
'4  syncDeltaSpeed = InputN - (OutputN * (35/11) * (39/28) * (34/27))
'       inputN = syncDeltaSpeed + (OutputN * (35/11) * (39/28) * (34/27))
'
'5  syncDeltaSpeed = InputN - (OutputN * (35/11) * (39/28) * (35/34))
'       inputN = syncDeltaSpeed + (OutputN * (35/11) * (39/28) * (35/34))
'
'6  syncDeltaSpeed = InputN - (OutputN * (35/11) * (39/28) * (28/32))
'       inputN = syncDeltaSpeed + (OutputN * (35/11) * (39/28) * (28/32))
'
'7  syncDeltaSpeed = InputN - (OutputN * (35/11) * (39/28) * (31/40))
'       inputN = syncDeltaSpeed + (OutputN * (35/11) * (39/28) * (31/40))

' Structure Variables.
Dim nRows As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim nCols As Integer
Dim firstCol As Integer
Dim lastCol As Integer

nRows = Range("A2", Range("A2").End(xlDown)).Rows.Count
firstRow = 2
lastRow = firstRow + nRows
nCols = 6
firstCol = 1
lastCol = firstCol + nCols

' Data Variables.
Dim gear As Variant
Dim inputN As Long
Dim inputN_ As Long
Dim inputN_valid As Boolean
Dim outputN As Long
Dim outputN_ As Long
Dim nShifts As Integer
Dim waitTime As Integer
Dim shiftType As String
Dim shiftType_ As String
Dim syncDeltaSpeed As Long
Dim syncDeltaSpeed_ As Long
Dim gearN As Long
Dim shaftN As Long
Dim cellChanged As Boolean
Dim whatChanged As String
Dim nChanged As Integer
nChanged = 0

' Workbook variables
Dim wb As Workbook
Dim ws As Worksheet
Set wb = ThisWorkbook
Set ws = wb.Sheets("Sheet1")
Debug.Print wb.Name
Debug.Print ws.Name

Application.ScreenUpdating = False

' Loop through each row
For i = firstRow To lastRow

    gear = ws.Cells(i, 1).Value
    inputN = ws.Cells(i, 2).Value
    outputN = ws.Cells(i, 3).Value
    nShifts = ws.Cells(i, 4).Value
    waitTime = ws.Cells(i, 5).Value
    shiftType = ws.Cells(i, 6).Value
    
    ' clear cellChanged value
    ws.Cells(i, 7).Value = False
    
    'clear whatChanged value
    ws.Cells(i, 8).Value = ""

    If shiftType = "Track" Then
        
        'Edit track to Normal
        ws.Cells(i, 6).Value = "Normal"
        
        whatChanged = "shiftType Normal, was Track; "
        ws.Cells(i, 8).Value = whatChanged
        
        cellChanged = True
        ws.Cells(i, 7).Value = cellChanged

        'First calculate the syncDeltaSpeed for evaluation.
        If gear = 1 Then
            syncDeltaSpeed = (inputN * (18 / 52)) - (outputN * (35 / 11) * (39 / 28))
        
        ElseIf gear = 2 Then
            syncDeltaSpeed = (inputN * (22 / 46)) - (outputN * (35 / 11) * (39 / 28))

        ElseIf gear = 3 Then
            syncDeltaSpeed = (inputN * (27 / 43)) - (outputN * (35 / 11) * (39 / 28))

        ElseIf gear = 4 Then
            syncDeltaSpeed = inputN - (outputN * (35 / 11) * (39 / 28) * (34 / 27))

        ElseIf gear = 5 Then
            syncDeltaSpeed = inputN - (outputN * (35 / 11) * (39 / 28) * (35 / 34))

        ElseIf gear = 6 Then
            syncDeltaSpeed = inputN - (outputN * (35 / 11) * (39 / 28) * (28 / 32))

        ElseIf gear = 7 Then
            syncDeltaSpeed = inputN - (outputN * (35 / 11) * (39 / 28) * (31 / 40))
        
        End If
        
        'enter sync deltaSpeed into column 9
        ws.Cells(i, 9).Value = syncDeltaSpeed
                
        'Evaluate if syncDeltaSpeed falls within specified limits
        If syncDeltaSpeed < 800 And syncDeltaSpeed > -800 Then

            'Evaluate if positive or negative and set syncDeltaSpeed accordingly
            If syncDeltaSpeed >= 0 And syncDeltaSpeed < 800 Then
                syncDeltaSpeed_ = 800
            ElseIf syncDeltaSpeed < 0 And syncDeltaSpeed > -800 Then
                syncDeltaSpeed_ = -800
            End If
            
            'Set output speed to 1000 ready for re-calculation
            outputN_ = 1000

            'Calculate new inputN_
            If gear = 1 Then
                inputN_ = (syncDeltaSpeed_ / (18 / 52)) + (outputN_ * (35 / 11) * (39 / 28))
            
            ElseIf gear = 2 Then
                inputN_ = (syncDeltaSpeed_ / (22 / 46)) + (outputN_ * (35 / 11) * (39 / 28))
            
            ElseIf gear = 3 Then
                inputN_ = (syncDeltaSpeed_ / (27 / 43)) + (outputN_ * (35 / 11) * (39 / 28))
            
            ElseIf gear = 4 Then
                inputN_ = syncDeltaSpeed_ + (outputN_ * (35 / 11) * (39 / 28) * (34 / 27))
            
            ElseIf gear = 5 Then
                inputN_ = syncDeltaSpeed_ + (outputN_ * (35 / 11) * (39 / 28) * (35 / 34))
            
            ElseIf gear = 6 Then
                inputN_ = syncDeltaSpeed_ + (outputN_ * (35 / 11) * (39 / 28) * (28 / 32))
            
            ElseIf gear = 7 Then
                inputN_ = syncDeltaSpeed_ + (outputN_ * (35 / 11) * (39 / 28) * (31 / 40))

            End If
            
            'enter sync deltaSpeed into column 9
            ws.Cells(i, 10).Value = syncDeltaSpeed_
            
            'Check inputN_ is within rig limits
            If inputN_ < 9950 Then
                inputN_valid = True
                
                cellChanged = True
                ws.Cells(i, 7).Value = cellChanged
                
                whatChanged = whatChanged & "inputN now " & inputN_ & " was " & inputN & "; OutputN now " & outputN_ & " was " & outputN
                ws.Cells(i, 8).Value = whatChanged
                                
                'Output new values
                ws.Cells(i, 2).Value = inputN_
                ws.Cells(i, 3).Value = outputN_
                
            Else
                inputN_valid = False
                    
            End If
            Debug.Print inputN & "," & outputN & "," & syncDeltaSpeed & "," & inputN_ & "," & outputN_ & "," & syncDeltaSpeed_ & "," & inputN_valid & "," & nShifts
        End If
    
    End If
    whatChanged = ""
    cellChanged = False
Next i

Application.ScreenUpdating = True

End Sub