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
        If Not IsEmpty(cell.Value) Then
            dataList(currentIndex) = cell.Value
            currentIndex = currentIndex + 1
        End If
    Next cell
Next data_range

Dim improvedDataList() As Single
ReDim improvedDataList(1 To currentIndex - 1)
    
For I = 1 To currentIndex - 1
    improvedDataList(I) = dataList(I)
Next I

Dim medValue As Single
medValue = WorksheetFunction.Median(improvedDataList)

Dim j As Long
Dim resultList() As Single
ReDim resultList(1 To currentIndex - 1)
For j = 1 To currentIndex - 1
    resultList(j) = Abs(improvedDataList(j) - medValue)
Next j

RSD = WorksheetFunction.Median(resultList) / 0.6745
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
        If Not IsEmpty(cell.Value) Then
            dataList(currentIndex) = cell.Value
            currentIndex = currentIndex + 1
        End If
    Next cell
Next data_range

Dim improvedDataList() As Single
ReDim improvedDataList(1 To currentIndex - 1)
    
For I = 1 To currentIndex - 1
    improvedDataList(I) = dataList(I)
Next I

Dim medValue As Single
medValue = WorksheetFunction.Median(improvedDataList)

Dim j As Long
Dim resultList() As Single
ReDim resultList(1 To currentIndex - 1)
For j = 1 To currentIndex - 1
    resultList(j) = Abs(improvedDataList(j) - medValue)
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
        If Not IsEmpty(cell.Value) Then
            dataList(currentIndex) = cell.Value
            currentIndex = currentIndex + 1
        End If
    Next cell
Next data_range

Dim improvedDataList() As Single
ReDim improvedDataList(1 To currentIndex - 1)
    
For I = 1 To currentIndex - 1
    improvedDataList(I) = dataList(I)
Next I

Dim avgValue As Single
avgValue = WorksheetFunction.Average(improvedDataList)

Dim j As Long
Dim resultList() As Single
ReDim resultList(1 To currentIndex - 1)
For j = 1 To currentIndex - 1
    resultList(j) = Abs(improvedDataList(j) - avgValue)
Next j

MeanAbsDev = WorksheetFunction.Sum(resultList) / (currentIndex - 1)
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

Dim RSD As Single
RSD = WorksheetFunction.Median(resultList) / 0.6745

RCV = RSD / WorksheetFunction.Median(dataList)
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

CV = WorksheetFunction.StDev(dataList) / WorksheetFunction.Average(dataList)
End Function
Function RZPrime(high_control_rsd, low_control_rsd, high_control_median, low_control_median As Double) As Double

difference = high_control_median - low_control_median
If difference > 0 Then
    RZPrime = 1 - (3 * (high_control_rsd + low_control_rsd) / difference)
ElseIf difference < 0 Then
    RZPrime = 1 - (3 * (high_control_rsd + low_control_rsd) / ((-1) * difference))
Else
    RZPrime = 0
End If
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
Function RZPrimeSamples(high_control_rsd, low_control_rsd, high_control_median, low_control_median, reps_per_sample As Double) As Double

difference = high_control_median - low_control_median
If difference > 0 Then
    RZPrimeSamples = 1 - (3 * (high_control_rsd + low_control_rsd) / (difference * (reps_per_sample) ^ 0.5))
ElseIf difference < 0 Then
    RZPrimeSamples = 1 - (3 * (high_control_sd + low_control_sd) / ((-1) * difference * (reps_per_sample) ^ 0.5))
Else
    RZPrimeSamples = 0
End If
End Function
Function ZPrimeSamples(high_control_sd, low_control_sd, high_control_mean, low_control_mean, reps_per_sample As Double) As Double

difference = high_control_mean - low_control_mean
If difference > 0 Then
    ZPrimeSamples = 1 - (3 * (high_control_sd + low_control_sd) / (difference * (reps_per_sample) ^ 0.5))
ElseIf difference < 0 Then
    ZPrimeSamples = 1 - (3 * (high_control_sd + low_control_sd) / ((-1) * difference * (reps_per_sample) ^ 0.5))
Else
    ZPrimeSamples = 0
End If
End Function
Function RZPrime365(high_control, low_control As Range) As Double

Dim high_control_median As Double
high_control_median = WorksheetFunction.Median(high_control)

Dim low_control_median As Double
low_control_median = WorksheetFunction.Median(low_control)

difference = high_control_median - low_control_median

Dim totalCellsHigh As Long
totalCellsHigh = WorksheetFunction.Count(high_control)

Dim totalCellsLow As Long
totalCellsLow = WorksheetFunction.Count(low_control)

Dim resultListHigh() As Single
ReDim resultListHigh(1 To totalCellsHigh)

Dim resultListLow() As Single
ReDim resultListLow(1 To totalCellsLow)

For j = 1 To totalCellsHigh
    resultListHigh(j) = Abs(high_control(j) - high_control_median)
Next j

For j = 1 To totalCellsLow
    resultListLow(j) = Abs(low_control(j) - low_control_median)
Next j

Dim high_control_rsd As Double
high_control_rsd = WorksheetFunction.Median(resultListHigh) / 0.6745
Dim low_control_rsd As Double
low_control_rsd = WorksheetFunction.Median(resultListLow) / 0.6745

If difference > 0 Then
    RZPrime365 = 1 - (3 * (high_control_rsd + low_control_rsd) / difference)
ElseIf difference < 0 Then
    RZPrime365 = 1 - (3 * (high_control_rsd + low_control_rsd) / ((-1) * difference))
Else
    RZPrime365 = 0
End If
End Function
Function ZPrime365(high_control, low_control As Range) As Double

Dim high_control_sd As Double
high_control_sd = WorksheetFunction.StDev(high_control)

Dim low_control_sd As Double
low_control_sd = WorksheetFunction.StDev(low_control)

Dim high_control_mean As Double
high_control_mean = WorksheetFunction.Average(high_control)

Dim low_control_mean As Double
low_control_mean = WorksheetFunction.Average(low_control)

difference = high_control_mean - low_control_mean

If difference > 0 Then
    ZPrime365 = 1 - (3 * (high_control_sd + low_control_sd) / difference)
ElseIf difference < 0 Then
    ZPrime365 = 1 - (3 * (high_control_sd + low_control_sd) / ((-1) * difference))
Else
    ZPrime365 = 0
End If
End Function
Function RSW(high_control_rsd, low_control_rsd, high_control_median, low_control_median As Double) As Double

If low_control_rsd = 0 Then
    RSW = 0
Else
    RSW = (Abs(high_control_median - low_control_median) - 3 * (high_control_rsd + low_control_rsd)) / low_control_rsd
End If
End Function
Function SW(high_control_sd, low_control_sd, high_control_mean, low_control_mean As Double) As Double

If low_control_sd = 0 Then
    SW = 0
Else
    SW = (Abs(high_control_mean - low_control_mean) - 3 * (high_control_sd + low_control_sd)) / low_control_sd
End If
End Function
Function RZscore(signal As Double, plate_range As Range) As Double

Dim plate_median As Double
plate_median = WorksheetFunction.Median(plate_range)

Dim totalCellsPlate As Long
totalCellsPlate = WorksheetFunction.Count(plate_range)

Dim resultListPlate() As Single
ReDim resultListPlate(1 To totalCellsPlate)


For j = 1 To totalCellsPlate
    resultListPlate(j) = Abs(plate_range(j) - plate_median)
Next j


Dim plate_rsd As Double
plate_rsd = WorksheetFunction.Median(resultListPlate) / 0.6745

RZscore = (signal - plate_median) / plate_rsd
End Function
Function Zscore(signal As Double, plate_range As Range) As Double

Zscore = (signal - WorksheetFunction.Average(plate_range)) / WorksheetFunction.StDev(plate_range)
End Function

