
' XT-13885 - Gearchange Duty Cycle
' Macro Created by T Osborne on 14th October 2020

'File directory: -
'\\DSUK01\Company Shared Documents\Projects\1306\DATA EXCHANGE\HPP\Received\201007 - Gearchange Duty Heat Maps\DATA PROCESSING\

'Spreadsheet names: -
'------------------------------
'Normal_Transmission_Heat_Maps_200518_AMG_Rural_02_06_11
'Normal_Transmission_Heat_Maps_200518_AMG_Rural_02_06_11_2
'Normal_Transmission_Heat_Maps_200518_AMGCity_5deg_02_07_00
'Normal_Transmission_Heat_Maps_200518_AMGCity_Hot_02_07_00
'Normal_Transmission_Heat_Maps_200518_BrixBrack_02_06_12
'Normal_Transmission_Heat_Maps_200518_BrixCamb_10deg_02_06_11
'Normal_Transmission_Heat_Maps_200518_FTP75_02_06_11
'Normal_Transmission_Heat_Maps_200518_Highway_02_06_11
'Normal_Transmission_Heat_Maps_200518_WLTC_02_06_12
'Normal_Transmission_Heat_Maps_200518_WLTC_02_06_12_2
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11_2
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11_3
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11_4
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11_5
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11_6
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11_7
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11_8
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11_9
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11_10
'Track_Transmission_Heat_Maps_200518_RDE_Agg_zero_deg_02_06_11
'Track_Transmission_Heat_Maps_200518_Roundabout_Sprint_02_06_11
'Track_Transmission_Heat_Maps_200518_Roundabout_Sprint_02_06_11_2
'Track_Transmission_Heat_Maps_200518_Roundabout_Sprint_02_06_11_3
'Track_Transmission_Heat_Maps_200518_Roundabout_Sprint_02_06_11_4
'Track_Transmission_Heat_Maps_200518_Roundabout_Sprint_02_06_11_5
'Track_Transmission_Heat_Maps_200518_Roundabout_Sprint_02_06_11_6
'Track_Transmission_Heat_Maps_200518_zero_250_02_06_11
'Track_Transmission_Heat_Maps_200518_zero_250_02_06_11_2

'Worksheet names: -
'1st
'2nd
'3rd
'4th
'5th
'6th
'7th

'Outputs required
'- Count No. of Shifts per gear, per drive cycle
'- Total number of shifts per gear
'- Total number of shifts per drive cycle
'- Total number of shifts overall
'- Compare to Overall duty cycle

' ----------------------------------------------------------
'| Gear | InputN | OutputN | nShifts | waitTime | shiftType |
' ----------------------------------------------------------

' Set target folder
' Loop through each workbook
'   Open file
'   set shiftType
'   Loop through each worksheet
'       set Gear
'       Loop through each cell in the matrix
'           set nShifts, InputN, OutputN & waitTime
'           add variables to Array
'           Select wbOutput > wsOutput
'           output Array to new row
'       Close loop
'       Sum totalShifts_Cycle_Gear
'   Close loop
'   Sum totalShifts_Cycle
' Close loop
' Sum totalShifts


Public Sub dutyCycleProcessing()

'Filepath & Directory variables.
Dim targetDir As String       ' Folder containing workbooks to be processed.
Dim fileType As String           ' File Extension .xlsx
Dim filePath As String           ' Concatenated string for exact file path
Dim fileCount As Integer         ' Number of files in the targetDir

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

Dim gearSpeed1 As Long      'Gear speed 1st
Dim gearSpeed2 As Long      'Gear speed 2nd
Dim gearSpeed3 As Long      'Gear speed 3rd
Dim gearSpeed4 As Long      'Gear speed 4th
Dim gearSpeed5 As Long      'Gear speed 5th
Dim gearSpeed6 As Long      'Gear speed 6th
Dim gearSpeed7 As Long      'Gear speed 7th

Dim shaftSpeed1 As Long     'Shaft speed 1st
Dim shaftSpeed2 As Long     'Shaft speed 2nd
Dim shaftSpeed3 As Long     'Shaft speed 3rd
Dim shaftSpeed4 As Long     'Shaft speed 4th
Dim shaftSpeed5 As Long     'Shaft speed 5th
Dim shaftSpeed6 As Long     'Shaft speed 6th
Dim shaftSpeed7 As Long     'Shaft speed 7th


'Output variables.
Dim nShifts      'Number of shifts
Dim totalShifts_Cycle_Gear  'total for worksheet
Dim totalShifts_Cycle       'total for workbook
Dim totalShifts             'total for all
Dim trackShift

'Initialise input data variables.
firstCol = 3
firstRow = 4
nCol = 48
nRow = 48
lastCol = firstCol + nCol - 1
lastRow = firstRow + nRow - 1
inputNRow = firstRow - 1
outputNCol = firstCol - 1
nGears = 7
curGear = 1


'initialise outputs
totalShifts_Cycle_Gear = 0
totalShifts_Cycle = 0
totalShifts = 0

'Table output array variables.
Dim inputN                  'Input speed
Dim outputN                 'Output speed
Dim SS                      'Sliding speed
Dim W                       'Friction work
Dim P                       'Friction power
Dim testPointArr(7)         'Output array
Dim outputTableRow As Long  'Current output row
Dim outputTableRow_ As Long 'First output row
Dim outputTableCol As Long  'Current output column
Dim outputTableCol_ As Long 'First output column

'Initialise output array variables
outputTableRow_ = 2
outputTableRow = outputTableRow_
outputTableCol_ = 1
outputTableCol = outputTableCol_



'------------------------------------
'Set Target Folder
'------------------------------------
targetDir = "\\DSUK01\Company Shared Documents\Projects\1306\DATA EXCHANGE\HPP\Received\201007 - Gearchange Duty Heat Maps\DATA PROCESSING\"
fileType = ".xlsx"
filePath = targetDir & "*" & fileType

ChDir targetDir
fileName = Dir(filePath)

wbOutputPath = targetDir & "\output\dataProcessing_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & fileType
Workbooks.Add.SaveAs fileName:=wbOutputPath
                        
'Loop through each file
While (fileName <> "")
    fileCount = fileCount + 1
    
    'Open a file and count some values
    Dim wbCur As Workbook
    Set wbCur = Workbooks.Open(targetDir & fileName) 'now the workbook is open
    wbCur.Activate
    Debug.Print "Active workbook is now " & wbCur.Name
    
        'Define input worksheet array
        Dim wsArr As Variant
        wsArr = Array("1st", "2nd", "3rd", "4th", "5th", "6th", "7th")

        'Set shift type
        trackShift = isTrackShift(CStr(fileName))
    
        If trackShift Then
            shiftType = "Track"
        Else
            shiftType = "Normal"
        End If
        Debug.Print "Shift type set to: " & shiftType
        
        'Loop through sheets
        Dim wsCur As Worksheet
        
        For k = 0 To 6
        Debug.Print "Working on... " & wsArr(k)
        
        Set wsCur = wbCur.Sheets(wsArr(k))
        curGear = k + 1

            'Now we have the worksheet selected, we need to loop through each cell and gather data

            'Loop through each column
            Dim i As Integer
            For i = firstCol To lastCol
            
                'Loop through each row within the column.
                Dim j As Integer
                For j = firstRow To lastRow

                    'Only process the cells that contain data.
                    nShifts = wsCur.Cells(j, i).Value * 58
                    If nShifts <> "NaN" Then
                        
                        ' Count shifts.
                        totalShifts_Cycle_Gear = totalShifts_Cycle_Gear + nShifts

                        
                        ' Compile data
                        ' Required Columns: - Input Speed, Output Speed, Gear, No. of Shifts, Wait Time, Shift Type

                        inputN = (wsCur.Cells(inputNRow, i).Value + wsCur.Cells(inputNRow, i + 1).Value) / 2
                        outputN = (wsCur.Cells(j, outputNCol).Value + wsCur.Cells(j + 1, outputNCol).Value) / 2

                        'Wait time calculation.
                        'toDo: Need to add SS, W & P calculations
                        
                        thresholdSpeed = 2000 'rpm 10s below limit, 7s above limit.
                        
                        If inputN <= thresholdSpeed Then
                            waitTime = 10
                        Else
                            waitTime = 7
                        End If

                        'Build the array - probably a better way to do this part!
                        testPointArr(0) = curGear
                        testPointArr(1) = inputN
                        testPointArr(2) = Round(outputN, 0)
                        testPointArr(3) = nShifts
                        testPointArr(4) = waitTime
                        testPointArr(5) = shiftType
                        testPointArr(6) = wbCur.Name
                        Debug.Print "Output Array: " & testPointArr(0) & ", " & testPointArr(1) & ", " & testPointArr(2) & ", " & testPointArr(3) & ", " & testPointArr(4) & ", " & testPointArr(5)
                        
                        'File each array value into the output table
                        Dim wbOutput As Workbook
                        Dim wsOutput As Worksheet
                                                
                        Set wbOutput = Workbooks.Open(wbOutputPath)
                        wbOutput.Activate
                        Set wsOutput = wbOutput.Sheets("Sheet1")
                        
                        Debug.Print "Output wb: " & ActiveWorkbook.Name

                        Dim n As Integer
                        For n = 0 To 5
                            wsOutput.Cells(outputTableRow, outputTableCol).Value = testPointArr(n)
                            outputTableCol = outputTableCol + 1
                            Debug.Print "Output Array Index: " & n & ", " & testPointArr(n)
                        Next n
                                    
                        'reset column and select next row
                        outputTableCol = outputTableCol_
                        outputTableRow = outputTableRow + 1
                        wbCur.Activate
                        Debug.Print "Active wb: " & wbCur.Name
                    End If
                    
                Next j
            
            Next i

            totalShifts_Cycle = totalShifts_Cycle + totalShifts_Cycle_Gear
            Debug.Print fileName & ": " & wsArr(k) & ": " & totalShifts_Cycle_Gear & " shifts."

            'Reset total count for each gear.
            totalShifts_Cycle_Gear = 0
        Next k
        Debug.Print fileName & ": " & totalShifts_Cycle & " shifts in cycle."

        totalShifts = totalShifts + totalShifts_Cycle
        
        'Reset total count for each cycle
        totalShifts_Cycle = 0

    fileName = Dir()
    ActiveWorkbook.Close
    
    'Do output file work here
    


Wend
Debug.Print fileCount & " files found."
Debug.Print totalShifts & " shifts counted."

totalShifts = totalShifts
Debug.Print "Total Shifts = " & totalShifts

End Sub

Public Sub outputData()

'Filepath & Directory variables.
Dim targetFileType As String     ' File Extension .xlsx
Dim wbOutput As Workbook    'Spreadhseet to store output data
Dim outputSheet As String: outputSheet = "output"

'Define file location and type.
targetFileType = ".xlsx"
outputDir = "\\DSUK01\Company Shared Documents\Projects\1306\DATA EXCHANGE\HPP\Received\201007 - Gearchange Duty Heat Maps\"
outputPath = outputDir & "dataProcessing" & targetFileType

Set wbOutput = Workbooks.Open(outputPath)

If DoesSheetExists(outputSheet) Then
    Debug.Print "sheet exists"
Else
    Sheets.Add.Name = outputSheet
End If

End Sub

Function DoesSheetExists(sh As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sh)
    On Error GoTo 0

    If Not ws Is Nothing Then DoesSheetExists = True
End Function

Public Function isTrackShift(mainString As String) As Boolean

Dim subString As String: subString = "track"

'Debug.Print "Checking " & mainString

i = InStr(LCase(mainString), LCase(subString))

If i = 1 Then
    isTrackShift = True
End If

End Function

Sub findShiftType()

'Filepath & Directory variables.
Dim targetDir As String          ' Folder containing workbooks to be processed.
Dim targetFileType As String     ' File Extension .xlsx
Dim filePath As String           ' Concatenated string for exact file path
Dim fileCount As Integer         ' Number of files in the targetDir
Dim shiftType As String

'Define file location and type.
targetDir = "\\DSUK01\Company Shared Documents\Projects\1306\DATA EXCHANGE\HPP\Received\201007 - Gearchange Duty Heat Maps\DATA PROCESSING\"
targetFileType = ".xlsx"
filePath = targetDir & "*" & targetFileType
' Debug.Print filePath

Dim fileName As String
fileName = Dir(filePath)

Dim trackShift As Boolean

While (fileName <> "")
    trackShift = isTrackShift(fileName)
    
    If trackShift Then
        shiftType = "Track"
    Else
        shiftType = "Normal"
    End If
    
    Debug.Print fileName & ": " & shiftType
    fileName = Dir()
Wend

End Sub



