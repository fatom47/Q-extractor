Dim subgSize As Byte 'subgroup size for Q-das <2;25>
Dim choice As String '<daily base;fixed count>
Dim fixedCount As Long 'required count of Output samples
Dim groupCount As Long 'number of groups (fixedCount/sibgSize)
Dim rowCount As Long 'count of rows in Input sheet
Dim groupSize As Long 'size of each group (rowCount/groupCount)
Dim columnCount As Byte 'count of columns in Input sheet
Dim groupStart As Long 'starting row number of a group
Dim row As Integer 'counter of actual Output row
Dim dateTime As Byte 'column number of DateTime in Input sheet
Dim day As String 'currently processed day
Dim start As Long, finish As Long 'row number of start and finish production day
Dim random As Integer 'random row count of shift of groupStart
Dim auxInt As Long 'auxiliary variable of Long integer data type
Dim rRange As Integer 'maximal random range

Sub Run()
subgSize = [B2].Value
choice = [B3].Value
row = 2 '1 is header
rRange = [B6].Value

Range("A7:B8").Select
Worksheets("Output").Cells.Clear
Worksheets("Output").Cells.NumberFormat = "@" 'text format

' Gets the number of rows in Input sheet
rowCount = Worksheets("Input").Cells(Rows.Count, 1).End(xlUp).row
' Gets the number of columns in Input sheet
columnCount = Worksheets("Input").Cells(1, Columns.Count).End(xlToLeft).Column

' Checks if Input is not empty
If (rowCount = 1 And columnCount = 1) Then
    MsgBox ("Missing data in Input sheet")
    Exit Sub
End If

' Copies the header from Input to Output
For C = 1 To columnCount
    Worksheets("Output").Cells(1, C).Value = Worksheets("Input").Cells(1, C).Value
    Worksheets("Output").Cells(1, C).Font.Bold = True
Next C

' Evenly spaced interval
If (choice = "Fixed count") Then
    fixedCount = [B4].Value
    groupCount = fixedCount / subgSize
    groupSize = rowCount / groupCount
    groupStart = (groupSize / 2) - (subgSize / 2)

    ' Checks if N subgroups can consist requested fixed samples count
    If ((fixedCount Mod subgSize) <> 0) Then
        MsgBox ("The number of samples does not match the size of the subgroup size.")
        Exit Sub
    End If

    ' Checks if random range is larger than half of group size
    If (rRange > groupSize / 2) Then
        Cells(6, 2).Value = groupSize / 2
        rRange = groupSize / 2
        MsgBox ("Random range was too large and was lowered to the half of group size")
    End If
    
    ' Copies values of selected cells from Input to Output
    For G = 1 To groupCount
        
        ' Randomization of groupStart
        Do
            random = retRandom()
            auxInt = ((G - 1) * groupSize) + groupStart + random
        Loop Until auxInt > 1 And auxInt < (rowCount - subgSize)
        
        For S = 1 To subgSize
            For A = 1 To columnCount
                Worksheets("Output").Cells(row, A).Value = Worksheets("Input").Cells(((G - 1) * groupSize) + (groupStart + S - 1 + random), A).Value
            Next A
            row = row + 1
        Next S
    Next G
    
' Day by day
ElseIf (choice = "Daily base") Then
    dateTime = [B5].Value
    
    ' Checks if date value is sufficiently detailed
    If (Len(Worksheets("Input").Cells(row, dateTime)) >= 10) Then
        Worksheets("Input").Cells(1, columnCount + 1).Value = "Date"
        
        ' Makes an extract of timestamp
        For R = 2 To rowCount
            Worksheets("Input").Cells(R, columnCount + 1) = Left(Worksheets("Input").Cells(R, dateTime), 10)
        Next R
        
        start = 2
        finish = 2
        day = Worksheets("Input").Cells(start, columnCount + 1)
        
        Do While finish < rowCount
            ' Finds end of the same day
            Do While day = Worksheets("Input").Cells(finish, columnCount + 1)
                finish = finish + 1
            Loop
            
            ' Anticycling protection
            If (subgSize >= finish - start) Then
                MsgBox ("Block of data between rows " & start & " and " & finish & " is smaller than size of subgroup!")
                Exit Do
            End If
            
            ' Regular group start without random value
            groupStart = ((start + finish - 1) / 2) - (subgSize / 2)
            
            ' Randomization of groupStart within day limits
            Do
                random = retRandom()
                auxInt = groupStart + random
            Loop Until auxInt > (start - 1) And auxInt < (finish - subgSize)
            
            ' Copies values of selected cells from Input to Output
            For S = 1 To subgSize
                For A = 1 To columnCount
                    Worksheets("Output").Cells(row, A).Value = Worksheets("Input").Cells(groupStart + S - 1 + random, A).Value
                Next A
                row = row + 1
            Next S
            
            start = finish
            day = Worksheets("Input").Cells(start, columnCount + 1)
        Loop
        
        ' Clears the extract of timestamp
        Worksheets("Input").Columns(columnCount + 1).EntireColumn.Delete
    Else
        MsgBox ("Must be in at least 10 character format")
        Exit Sub
    End If
    
Else
    MsgBox ("Miracle!")
    Exit Sub
End If

' Uniform appearance
Worksheets("Input").Activate
ActiveWindow.ScrollRow = 1
ActiveWindow.ScrollColumn = 1
Worksheets("Input").Range("A1").Select
Worksheets("Output").Activate
ActiveWindow.ScrollRow = 1
ActiveWindow.ScrollColumn = 1
Worksheets("Output").Range("A2").Select
End Sub

' Returns random number
Public Function retRandom() As Integer
    Dim number As Integer
    number = Rnd() * rRange
    If Rnd() < 0.5 Then number = number * (-1)
    retRandom = number
End Function
