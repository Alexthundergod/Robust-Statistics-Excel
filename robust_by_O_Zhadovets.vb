Function MedianAbsDev(ParamArray data_ranges() As Variant) As Single
'By O.Zhadovets

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
        dataList(currentIndex) = cell.value
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
'By O.Zhadovets

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
        dataList(currentIndex) = cell.value
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

Function PercentageInhibition(signal, postive_control, negative_control As Single) As Single
'By O.Zhadovets

PercentageInhibition = ((signal - postive_control) / (negative_control - postive_control)) * 100
End Function