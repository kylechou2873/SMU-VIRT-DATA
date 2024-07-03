SMU VIRT DATA Assignment #2
Kyle Chou
Screen Shots of result and Separate VBA file in this folder

'Copy of the code in VBA file is below

Sub assig2():
    Dim ws As Worksheet
    'iter for each row
    Dim i As Integer
    'count for last row
    Dim rowC As Integer
    'loop for each worksheet
    For Each ws In Worksheets
        rowC = ws.Range("A1", ws.Range("A1").End(xlDown)).Rows.Count
        'setup column for result
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        
        'iter for results
        Dim res As Integer
        res = 2
        'store total for volume
        Dim vol As LongLong
        vol = 0
        'store open and close value
        Dim openQ As Double
        Dim closeQ As Double
        'store ticker
        Dim ticker As String
        'store greatest % increase
        Dim gTick As String
        Dim gVal As Double
        gVal = 0
        'store greatest % decrease
        Dim dTick As String
        Dim dVal As Double
        dVal = 0
        'store greatest Volume
        Dim vTick As String
        Dim vVal As LongLong
        vVal = 0
        For i = 2 To rowC
            'if last row ticker is different from this row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                'store open value
                openQ = ws.Cells(i, 3).Value
            End If
            vol = vol + ws.Cells(i, 7).Value
            'if next row ticker is different from this row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'store close value
                closeQ = ws.Cells(i, 6).Value
                'write result
                ticker = ws.Cells(i, 1).Value
                ws.Cells(res, 10).Value = ticker
                ws.Cells(res, 11).Value = closeQ - openQ
                ws.Cells(res, 11).NumberFormat = "$#,##0.00"
                If ws.Cells(res, 11).Value > 0 Then
                    ws.Cells(res, 11).Interior.ColorIndex = 4
                ElseIf ws.Cells(res, 11).Value < 0 Then
                    ws.Cells(res, 11).Interior.ColorIndex = 3
                End If
                ws.Cells(res, 12).Value = (closeQ - openQ) / openQ
                'if this % is bigger then the current greatest % increase then replace
                If ws.Cells(res, 12).Value > gVal Then
                    gTick = ticker
                    gVal = ws.Cells(res, 12).Value
                'if this % is less then the current greatest % decrease then replace
                ElseIf ws.Cells(res, 12).Value < dVal Then
                    dTick = ticker
                    dVal = ws.Cells(res, 12).Value
                End If
                ws.Cells(res, 12).NumberFormat = "0.00%"
                ws.Cells(res, 13).Value = vol
                'if this volume is greater then the current total volume then replace
                If vol > vVal Then
                    vTick = ticker
                    vVal = vol
                End If
                'reset vol and get next res
                vol = 0
                res = res + 1
            End If
        Next i
        'write result summary
        ws.Cells(2, 17).Value = gTick
        ws.Cells(2, 18).Value = gVal
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = dTick
        ws.Cells(3, 18).Value = dVal
        ws.Cells(3, 18).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = vTick
        ws.Cells(4, 18).Value = vVal
        
    Next ws
End Sub
'End Copy
