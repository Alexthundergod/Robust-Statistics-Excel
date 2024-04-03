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

Then open Excel and Go to

```
Options –> Add-ins –> Manage: Excel Add-ins –> Go...
```

In the Add-ins dialogue box tick <i>Robust_by_O_Zhadovets</i> :white_check_mark:, and click OK. Voila! Now all the functions will be available in any Excel sheet.

<h2>Functions</h2>

- MedianAbsDev **(** *`Datarange1`*, *`Datarange2...`* **)** — $Median(X_i - Median(X))$

- MeanAbsDev **(** *`Datarange1`*, *`Datarange2...`* **)** — $Mean(X_i - Mean(X))$

- PercentageInhibition **(** *`Signal`*, *`High_control`*, *`Low_control`* **)** — $(Signal - LowC)/(HighC - LowC)$

- RCV **(** *`Datarange1`*, *`Datarange2...`* **)** — $MedianAbsDev(X) / Median(X)$

- CV **(** *`Datarange1`*, *`Datarange2...`* **)** — $StdDev(X) / Mean(X)$
  
<h2>License</h2>

Distributed under the MIT License. See `LICENSE` for more information.
