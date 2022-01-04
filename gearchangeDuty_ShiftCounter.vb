
' XT-13885 - Gearchange Duty Cycle
' Macro Created by T Osborne on 14th October 2020

'File directory: -
'\\DSUK01\Company Shared Documents\Projects\1306\DATA EXCHANGE\HPP\Received\201007 - Gearchange Duty Heat Maps\DATA PROCESSING\

'Spreadsheet names: -
'------------------------------
'Normal_Transmission_Heat_Maps_200518_AMG_Rural_02_06_11
'Normal_Transmission_Heat_Maps_200518_AMGCity_5deg_02_07_00
'Normal_Transmission_Heat_Maps_200518_BrixBrack_02_06_12
'Normal_Transmission_Heat_Maps_200518_BrixCamb_10deg_02_06_11
'Normal_Transmission_Heat_Maps_200518_FTP75_02_06_11
'Normal_Transmission_Heat_Maps_200518_Highway_02_06_11
'Normal_Transmission_Heat_Maps_200518_WLTC_02_06_12
'Track_Transmission_Heat_Maps_200518_NBR_02_06_11
'Track_Transmission_Heat_Maps_200518_RDE_Agg_zero_deg_02_06_11
'Track_Transmission_Heat_Maps_200518_Roundabout_Sprint_02_06_11
'Track_Transmission_Heat_Maps_200518_zero_250_02_06_11

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

' Set target folder
' Loop through each spreadsheet
'   Open file
'   Loop through each worksheet
'       Loop through each cell in the matrix
'           Store Count no. of shifts (cell value)
'       Close loop
'       Sum totalShifts_Cycle_Gear
'   Close loop
'   Sum totalShifts_Cycle
' Close loop
' Sum totalShifts

' variables required: -
' totalShifts > totalShifts_Cycle > totalShifts_Cycle_Gear > noShifts

Public Sub dutyCycleProcessing()

'Filepath & Directory variables.
Dim targetDir As String          ' Folder containing workbooks to be processed.
'Dim fileName(nFiles)            ' Array containing filenames in the targetDir. Might require manual building.
Dim targetFileType As String     ' File Extension .xlsx
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

'Shift counter output variables.
Dim nShifts                 'Number of shifts
Dim totalShifts_Cycle_Gear  'total for worksheet
Dim totalShifts_Cycle       'total for workbook
Dim totalShifts             'total for all
Dim nShifts_1st             
Dim nShifts_2nd
Dim nShifts_3rd
Dim nShifts_4th
Dim nshifts_5th
Dim nShifts_6th

'initialise shift count outputs
totalShifts_Cycle_Gear = 0
totalShifts_Cycle = 0
totalShifts = 0

'Define file location and type.
targetDir = "\\DSUK01\Company Shared Documents\Projects\1306\DATA EXCHANGE\HPP\Received\201007 - Gearchange Duty Heat Maps\DATA PROCESSING\"
targetFileType = ".xlsx"
filePath = targetDir & "*" & targetFileType
' Debug.Print filePath

ChDir targetDir
'Count files in targetDir.
Filename = Dir(filePath)
Debug.Print Filename

'Loop through each file
While (Filename <> "")
    fileCount = fileCount + 1
    ' Debug.Print fileCount & " " & Filename
    
    'Open a file and count some values

    Dim wbCur As Workbook
    Set wbCur = Workbooks.Open(targetDir & Filename)
        'now the workbook is open, we can count shifts in each worksheet
        
        'Loop through sheets
        'Define worksheet array
        Dim wsArr(7)
        wsArr(0) = "1st"
        wsArr(1) = "2nd"
        wsArr(2) = "3rd"
        wsArr(3) = "4th"
        wsArr(4) = "5th"
        wsArr(5) = "6th"
        wsArr(6) = "7th"
        
        Dim wsCur As Worksheet
        
        'Loop through each worksheet
        For k = 0 To 6
        Set wsCur = wbCur.Sheets(wsArr(k))
        wsCur.Activate

            'Now we have the worksheet selected, we need to loop through each cell and sum the shifts
            ' totalShifts > totalShifts_Cycle > totalShifts_Cycle_Gear > nShifts

            'Loop through each column
            Dim i As Integer
            For i = firstCol To lastCol
            
                'Loop through each row within the column.
                Dim j As Integer
                For j = firstRow To lastRow

                    'Only process the cells that contain data.
                    nShifts = wsCur.Cells(j, i).Value
                    If nShifts <> "NaN" Then
                        
                        totalShifts_Cycle_Gear = totalShifts_Cycle_Gear + nShifts
                    
                    End If
                    
                Next j
            
            Next i

            totalShifts_Cycle = totalShifts_Cycle + totalShifts_Cycle_Gear
            Debug.Print Filename & ": " & wsArr(k) & ": " & totalShifts_Cycle_Gear & " shifts."

            'Reset total count for each gear.
            totalShifts_Cycle_Gear = 0
        Next k
        Debug.Print Filename & ": " & totalShifts_Cycle & " shifts in cycle."

        totalShifts = totalShifts + totalShifts_Cycle
        
        'Reset total count for each cycle
        totalShifts_Cycle = 0

    Filename = Dir()
    ActiveWorkbook.Close
    
Wend
Debug.Print fileCount & " files found."
Debug.Print totalShifts & " shifts counted."

totalShifts = totalShifts * 58
Debug.Print "Total Shifts = " & totalShifts

End Sub