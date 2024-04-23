![GitHub commit activity](https://img.shields.io/github/commit-activity/w/Alexthundergod/Robust-Statistics-Excel?style=flat&color=ff80ff)
![GitHub Repo stars](https://img.shields.io/github/stars/Alexthundergod/Robust-Statistics-Excel?style=flat&color=88E809)

<h1>Robust Statistics Excel</h1>

A VBA module that allows you to find Median Absolute Deviation in Excel and, accordingly, calculate robust statistics. In the future, I plan to add a few more useful statistical functions.

<h2>Installation</h2>

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

<h2>Functions</h2>

- RSD **(** *`Datarange1`*, *`Datarange2...`* **)** — $Median(Abs(X_i - Median(X))) • 1.4826$

- MedianAbsDev **(** *`Datarange1`*, *`Datarange2...`* **)** — $Median(Abs(X_i - Median(X)))$

- MeanAbsDev **(** *`Datarange1`*, *`Datarange2...`* **)** — $Mean(Abs(X_i - Mean(X)))$

- PercentageActivity **(** *`Signal`*, *`High_control`*, *`Low_control`* **)** — $\dfrac{Signal - LowC}{HighC - LowC} • 100$

- PercentageInhibition **(** *`Signal`*, *`High_control`*, *`Low_control`* **)** — $(1 - (Signal - LowC)/(HighC - LowC)) • 100$

- RPercentageDrift **(** *`Datarange1`*, *`Datarange2`*, *`Platerange`* **)** — $(Median(X1) - Median(X2))/Median(Plate)$

- PercentageDrift **(** *`Datarange1`*, *`Datarange2`*, *`Platerange`* **)** — $(Mean(X1) - Mean(X2))/Mean(Plate)$

- RCV **(** *`Datarange1`*, *`Datarange2...`* **)** — $RSD(X) / Median(X)$

- CV **(** *`Datarange1`*, *`Datarange2...`* **)** — $SD(X) / Mean(X)$

- RZprime **(** *`High_control_RSD`*, *`Low_control_RSD`*, *`High_control_Median`*, *`Low_control_Median`* **)** — $1 - \dfrac{3(RSD(HighC) + RSD(LowC))}{Median(HighC) - Median(LowC)}$

- Zprime **(** *`High_control_SD`*, *`Low_control_SD`*, *`High_control_Mean`*, *`Low_control_Mean`* **)** — $1 - \dfrac{3(SD(HighC) + SD(LowC))}{Mean(HighC) - Mean(LowC)}$

- RSW **(** *`High_control_RSD`*, *`Low_control_RSD`*, *`High_control_Median`*, *`Low_control_Median`* **)** — $\dfrac{Abs(Median(HighC) - Median(LowC)) - 3(RSD(HighC) + RSD(LowC))}{RSD(LowC)}$

- SW **(** *`High_control_SD`*, *`Low_control_SD`*, *`High_control_Mean`*, *`Low_control_Mean`* **)** — $\dfrac{Abs(Mean(HighC) - Mean(LowC)) - 3(SD(HighC) + SD(LowC))}{SD(LowC)}$
  
>NB!: The following functions require VSTACK() or HSTACK(), available **ONLY** in Office 365, to combine disparate ranges.
>
>Use it as follows: =Function(VSTACK(A1:B24;AU1:AW24);VSTACK(C1:D24;AX1:AY24))

- RZprime365 **(** *`High_control`*, *`Low_control`* **)** — $1 - \dfrac{3RSD(HighC) + 3RSD(LowC)}{Median(HighC) - Median(LowC)}$

- Zprime365 **(** *`High_control`*, *`Low_control`* **)** — $1 - \dfrac{3SD(HighC) + 3SD(LowC)}{Mean(HighC) - Mean(LowC)}$
  
<h2>License</h2>

Distributed under the MIT License. See `LICENSE` for more information.
