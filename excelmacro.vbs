Option Explicit

Sub RunMacro()
 
 Dim tempWB As Workbook, rw As Range
 Dim tempWS, dataWS, outputWS As Worksheet
 Dim countValue, outputPrintRow As Long
 Dim dataWSCol, tempRNG As Variant
 Dim masterListLocation, lastName, firstName, login As String
 Dim iInMasterRow, iInDataListRow, PersonsMissing As Collection
 Dim x, y, z As Integer
 

 On Error GoTo ErrHandler
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    'SPECIFY LOCAL PATH OF MASTERLIST.XLSX FILE
    masterListLocation = "C:\Users\Herman.Toro\Desktop\Excel-Macro-Example\MasterList.xlsx"
                   
    'ASSIGN DATA WORKSHEET
    Set dataWS = ActiveSheet
    
    'ASSIGN OUTPUT WORKSHEET
    ActiveWorkbook.Worksheets.Add(After:=Worksheets(1)).Name = "output"
    Set outputWS = ActiveWorkbook.Sheets("output")
    
    'OPEN THE MASTER LIST
    On Error Resume Next
            Set tempWB = Workbooks.Open(masterListLocation, True, True)
            Set tempWS = tempWB.Sheets("Sheet1")
            If tempWS Is Nothing Then
                Call CloseAll
            MsgBox "Cannot open Master List file", vbCritical
                Exit Sub
            End If
    On Error GoTo ErrHandler
    
    'ONLY COPY ROWS WITH DATA FROM MASTER LIST TO TEMP RANGE
    tempRNG = tempWS.Range(tempWS.Range("A1"), tempWS.Range("A1").End(xlDown)).Cells
    
    'CREATE COLLECTION OF ROW NUMBERS FROM MASTER LIST THAT CONTAIN DATA
    Set iInMasterRow = GetMasterRows(tempRNG)
    Set tempRNG = Nothing
    
    'ITERATE THROUGH COLLECTION OF MASTER LIST ROW NUMBERS ONE BY ONE
    Dim r As Integer
    If iInMasterRow.Count > 0 Then
        
        'INITIALIZE EMPLOYEE MISSING VARS
        Set PersonsMissing = New Collection
        outputPrintRow = 1
    
        For x = 2 To iInMasterRow.Count
        
                lastName = ""
                firstName = ""
                login = ""
                              
                'GET LAST NAME FROM MASTER LIST ROW
                lastName = Trim(tempWS.Cells(iInMasterRow(x), 1).Value)
                 
                'GET FIRST NAME FROM MASTER LIST ROW
                firstName = Trim(tempWS.Cells(iInMasterRow(x), 2).Value)
                                 
                'GET LOGIN FROM MASTER LIST ROW
                login = Trim(tempWS.Cells(iInMasterRow(x), 3).Value)
                                                       
                'COPY COLUMN B OF DATA LIST SINCE IT CONTAINS LOGIN DATA
                dataWSCol = Range(dataWS.Range("B1"), dataWS.Range("B1").End(xlDown)).Cells
                
                'CREATE COLLECTION OF DATA LIST ROW NUMBERS THAT CONTAIN THE SAME LOGIN AS MASTER LIST ROW
                Set iInDataListRow = GetMatchingRows(dataWSCol, login)
                Set dataWSCol = Nothing
                
                'ITERATE THROUGH COLLECTION OF DATA LIST ROW NUMBERS AND SUM UP VALUES FROM COLUMN C
                If iInDataListRow.Count > 0 Then
                
                    countValue = 0
                               
                    For y = 1 To iInDataListRow.Count
                    
                       'Value from Column C
                        countValue = countValue + dataWS.Cells(iInDataListRow(y), 3).Value
                                               
                    Next y
                    
                    'PRINT NAME AND TOTAL FOR EACH PERSON
                    If (outputPrintRow = 1) Then
                       outputWS.Cells(outputPrintRow, 1).Value = "TOTALS FROM DATALIST"
                       outputPrintRow = outputPrintRow + 1
                    End If
                    
                    outputWS.Cells(outputPrintRow, 1).Value = firstName & Space(2) & lastName & Space(8) & CStr(countValue)
                    outputPrintRow = outputPrintRow + 1
                Else
                    'SAVE LIST OF PERSONS MISSING FROM DATA LIST
                    PersonsMissing.Add firstName & Space(2) & lastName
                End If
                               
                Set iInDataListRow = Nothing
        
        Next x
        
        Set iInMasterRow = Nothing
        
        'ITERATE THROUGH COLLECTION OF PERSONS MISSING FROM DATA LIST
        For z = 1 To PersonsMissing.Count
            If (z = 1) Then
                outputPrintRow = outputPrintRow + 1
                outputWS.Cells(outputPrintRow, 1).Value = "PERSONS MISSING FROM DATA LIST"
                outputPrintRow = outputPrintRow + 1
            End If
            
            'PRINT NAMES OF MISSING PERSONS
            outputWS.Cells(outputPrintRow, 1).Value = PersonsMissing.Item(z)
            outputPrintRow = outputPrintRow + 1
        Next z
        
        Set PersonsMissing = Nothing
        
    End If
      
    tempWB.Close False
    Set tempWS = Nothing
    Set tempWB = Nothing
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
    
    MsgBox "File Processed Successfully"
    
    Exit Sub
    
ErrHandler:
        Call CloseAll
        MsgBox "Unhandled Error, please contact Systems", vbCritical
End Sub


Sub CloseAll()
        'CLOSE ALL OPEN WORKBOOKS ON ERROR
        If Not tempWB Is Nothing Then
            tempWB.Close False
        End If
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
End Sub

Function GetMatchingRows(arr, v) As Collection
Dim lb As Long, ub As Long, e As Long
    Set GetMatchingRows = New Collection
    lb = LBound(arr)
    ub = UBound(arr)
    For e = lb To ub
        If LCase(arr(e, 1)) = LCase(v) Then
            GetMatchingRows.Add e
        End If
    Next e
End Function

Function GetMasterRows(arr) As Collection
Dim lb As Long, ub As Long, e As Long
    Set GetMasterRows = New Collection
    lb = LBound(arr)
    ub = UBound(arr)
    For e = lb To ub
            GetMasterRows.Add e
    Next e
End Function

