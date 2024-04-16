'By O.Zhadovets
'https://github.com/Alexthundergod/Robust-Statistics-Excel

Function RSD(ParamArray data_ranges() As Variant) As Single

Dim totalCells As Long
totalCells = 0

For Each data_range In data_ranges
    totalCells = totalCells + data_range.Cells.Count
Next data_range

Dim dataList() As Single
ReDim dataList(1 To totalCells)
Dim currentIndex As Long
currentIndex = 1

For Each data_range In data_ranges
    For Each cell In data_range
        dataList(currentIndex) = cell.Value
        currentIndex = currentIndex + 1
    Next cell
Next data_range

Dim medValue As Single
medValue = WorksheetFunction.Median(dataList)

Dim j As Long
Dim resultList() As Single
ReDim resultList(1 To totalCells)
For j = 1 To totalCells
    resultList(j) = Abs(dataList(j) - medValue)
Next j

RSD = WorksheetFunction.Median(resultList) * 1.4826
End Function
Function MedianAbsDev(ParamArray data_ranges() As Variant) As Single

Dim totalCells As Long
totalCells = 0

For Each data_range In data_ranges
    totalCells = totalCells + data_range.Cells.Count
Next data_range

Dim dataList() As Single
ReDim dataList(1 To totalCells)
Dim currentIndex As Long
currentIndex = 1

For Each data_range In data_ranges
    For Each cell In data_range
        dataList(currentIndex) = cell.Value
        currentIndex = currentIndex + 1
    Next cell
Next data_range

Dim medValue As Single
medValue = WorksheetFunction.Median(dataList)

Dim j As Long
Dim resultList() As Single
ReDim resultList(1 To totalCells)
For j = 1 To totalCells
    resultList(j) = Abs(dataList(j) - medValue)
Next j

MedianAbsDev = WorksheetFunction.Median(resultList)
End Function
Function MeanAbsDev(ParamArray data_ranges() As Variant) As Single

Dim totalCells As Long
totalCells = 0

For Each data_range In data_ranges
    totalCells = totalCells + data_range.Cells.Count
Next data_range

Dim dataList() As Single
ReDim dataList(1 To totalCells)
Dim currentIndex As Long
currentIndex = 1

For Each data_range In data_ranges
    For Each cell In data_range
        dataList(currentIndex) = cell.Value
        currentIndex = currentIndex + 1
    Next cell
Next data_range

Dim avgValue As Single
avgValue = WorksheetFunction.Average(dataList)

Dim j As Long
Dim resultList() As Single
ReDim resultList(1 To totalCells)
For j = 1 To totalCells
    resultList(j) = Abs(dataList(j) - avgValue)
Next j

MeanAbsDev = WorksheetFunction.Sum(resultList) / totalCells
End Function
Function PercentageActivity(signal, high_control, low_control As Single) As Single

PercentageActivity = ((signal - low_control) / (high_control - low_control)) * 100
End Function
Function PercentageInhibition(signal, high_control, low_control As Single) As Single

PercentageInhibition = (1 - ((signal - low_control) / (high_control - low_control))) * 100
End Function
Function RPercentageDrift(data_range1 As Range, data_range2 As Range, plate_range As Range) As Single

RPercentageDrift = (100 * (WorksheetFunction.Median(data_range1) - WorksheetFunction.Median(data_range2))) / WorksheetFunction.Median(plate_range)
End Function
Function PercentageDrift(data_range1 As Range, data_range2 As Range, plate_range As Range) As Single

PercentageDrift = (100 * (WorksheetFunction.Average(data_range1) - WorksheetFunction.Average(data_range2))) / WorksheetFunction.Average(plate_range)
End Function
Function RCV(ParamArray data_ranges() As Variant) As Single

Dim totalCells As Long
totalCells = 0

For Each data_range In data_ranges
    totalCells = totalCells + data_range.Cells.Count
Next data_range

Dim dataList() As Single
ReDim dataList(1 To totalCells)
Dim currentIndex As Long
currentIndex = 1

For Each data_range In data_ranges
    For Each cell In data_range
        dataList(currentIndex) = cell.Value
        currentIndex = currentIndex + 1
    Next cell
Next data_range

Dim medValue As Single
medValue = WorksheetFunction.Median(dataList)

Dim j As Long
Dim resultList() As Single
ReDim resultList(1 To totalCells)
For j = 1 To totalCells
    resultList(j) = Abs(dataList(j) - medValue)
Next j

RCV = WorksheetFunction.Median(resultList) / WorksheetFunction.Median(dataList)
End Function
Function CV(ParamArray data_ranges() As Variant) As Single

Dim totalCells As Long
totalCells = 0

For Each data_range In data_ranges
    totalCells = totalCells + data_range.Cells.Count
Next data_range

Dim dataList() As Single
ReDim dataList(1 To totalCells)
Dim currentIndex As Long
currentIndex = 1

For Each data_range In data_ranges
    For Each cell In data_range
        dataList(currentIndex) = cell.Value
        currentIndex = currentIndex + 1
    Next cell
Next data_range

CV = WorksheetFunction.stDev(dataList) / WorksheetFunction.Average(dataList)
End Function
Function ZPrime(high_control_sd, low_control_sd, high_control_mean, low_control_mean As Double) As Double

difference = high_control_mean - low_control_mean
If difference > 0 Then
    ZPrime = 1 - (3 * (high_control_sd + low_control_sd) / difference)
ElseIf difference < 0 Then
    ZPrime = 1 - (3 * (high_control_sd + low_control_sd) / ((-1) * difference))
Else
    ZPrime = 0
End If
End Function

