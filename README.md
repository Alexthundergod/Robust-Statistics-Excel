![GitHub commit activity](https://img.shields.io/github/commit-activity/w/Alexthundergod/Robust-Statistics-Excel?style=flat&color=ff80ff)
![GitHub Repo stars](https://img.shields.io/github/stars/Alexthundergod/Robust-Statistics-Excel?style=flat&color=88E809)

# Robust Statistics Excel

A VBA module that allows you to find Median Absolute Deviation in Excel and, accordingly, calculate robust statistics. In the future, I plan to add a few more useful statistical functions.

## Installation

Save the <a href=https://github.com/Alexthundergod/Robust-Statistics-Excel/blob/main/robust_by_O_Zhadovets.xlam><i>robust_by_O_Zhadovets.xlam</i></a> file in

for Windows

```
C:\Users\USERNAME\AppData\Roaming\Microsoft\AddIns
```

for MacOS

```
/Users/USERNAME/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins
```

Next, be sure to unblock :white_check_mark: the downloaded file as follows 

```
robust_by_O_Zhadovets.xlam –> Properties –> General –> Security: Unblock
```
<img src="https://github.com/Alexthundergod/Robust-Statistics-Excel/blob/main/0.png"></img>

Then open Excel and Go to

```
Options –> Add-ins –> Manage: Excel Add-ins –> Go...
```
<img src="https://github.com/Alexthundergod/Robust-Statistics-Excel/blob/main/1.png"></img>
<img src="https://github.com/Alexthundergod/Robust-Statistics-Excel/blob/main/2.png"></img>
<img src="https://github.com/Alexthundergod/Robust-Statistics-Excel/blob/main/3.png"></img>

In the Add-ins dialogue box tick <i>Robust_by_O_Zhadovets</i> :white_check_mark:, and click OK. Voila! Now all the functions will be available in any Excel sheet.

<img src="https://github.com/Alexthundergod/Robust-Statistics-Excel/blob/main/4.png"></img>

## Functions

**Syntax:**
`RSD(Data_range1,Data_range2...)`

**Formula:**
$Median|X_i - Median(X)| • 1.4826$

---

**Syntax:**
`MedianAbsDev(Data_range1,Data_range2...)`

**Formula:**
$Median|X_i - Median(X)|$

---

**Syntax:**
`MeanAbsDev(Data_range1,Data_range2...)` 

**Formula:**
$Mean|X_i - Mean(X)|$

---

**Syntax:**
`PercentageActivity(Signal,High_control,Low_control)`

**Formula:**
$\dfrac{Signal - LowC}{HighC - LowC} • 100$

---

**Syntax:**
`PercentageInhibition(Signal,High_control,Low_control)` 

**Formula:**
$(1 - \dfrac{Signal - LowC}{HighC - LowC}) • 100$

---

**Syntax:**
`RPercentageDrift(Data_range1,Data_range2,Plate_range)` 

**Formula:**
$\dfrac{Median(X1) - Median(X2)}{Median(Plate)} • 100$

---

**Syntax:**
`PercentageDrift(Data_range1,Data_range2,Plate_range)` 

**Formula:**
$\dfrac{Mean(X1) - Mean(X2)}{Mean(Plate)} • 100$

---

**Syntax:**
`RCV(Data_range1,Data_range2...)` 

**Formula:**
$\dfrac{RSD(X)}{Median(X)}$

---

**Syntax:**
`CV(Data_range1,Data_range2...)` 

**Formula:**
$\dfrac{SD(X)}{Mean(X)}$

---

**Syntax:**
`RZprime(High_control_RSD,Low_control_RSD,High_control_Median,Low_control_Median)` 

**Formula:**
$1 - \dfrac{3(RSD(HighC) + RSD(LowC))}{Median(HighC) - Median(LowC)}$

---

**Syntax:**
`Zprime(High_control_SD,Low_control_SD,High_control_Mean,Low_control_Mean)` 

**Formula:**
$1 - \dfrac{3(SD(HighC) + SD(LowC))}{Mean(HighC) - Mean(LowC)}$

---

**Syntax:**
`ZprimeSamples(High_control_RSD,Low_control_RSD,High_control_Median,Low_control_Median, Samples_amount)` 

**Formula:**
$\dfrac{(Median(HighC) - 3(RSD(HighC)/\sqrt{samplesN})-(Median(LowC) - 3(RSD(LowC)/\sqrt{samplesN})}{Median(HighC) - Median(LowC)}$

---

**Syntax:**
`ZprimeSamples(High_control_SD,Low_control_SD,High_control_Mean,Low_control_Mean,Samples_amount)` 

**Formula:**
$\dfrac{(Mean(HighC) - 3(SD(HighC)/\sqrt{samplesN})-(Mean(LowC) - 3(SD(LowC)/\sqrt{samplesN})}{Mean(HighC) - Mean(LowC)}$ 

---

**Syntax:**
RSW(High_control_RSD,Low_control_RSD,High_control_Median,Low_control_Median) 

**Formula:**
$\dfrac{|Median(HighC) - Median(LowC)| - 3(RSD(HighC) + RSD(LowC))}{RSD(LowC)}$

---

**Syntax:**
`SW(High_control_SD,Low_control_SD,High_control_Mean,Low_control_Mean)` 

**Formula:**
$\dfrac{|Mean(HighC) - Mean(LowC)| - 3(SD(HighC) + SD(LowC))}{SD(LowC)}$

---

>NB!: The following functions require VSTACK() or HSTACK(), available **ONLY** in Office 365, to combine disparate ranges.
>
>Use it as follows: =Function(VSTACK(A1:B24;AU1:AW24);VSTACK(C1:D24;AX1:AY24))

---

**Syntax:**
`RZprime365(High_control,Low_control)` 

**Formula:**
$1 - \dfrac{3RSD(HighC) + 3RSD(LowC)}{Median(HighC) - Median(LowC)}$

---

**Syntax:**
`Zprime365(High_control,Low_control)` 

**Formula:**
$1 - \dfrac{3SD(HighC) + 3SD(LowC)}{Mean(HighC) - Mean(LowC)}$

---

  
## License

Distributed under the MIT License. See `LICENSE` for more information.
