# SydiServe_Excel_Fix
Fix for SydiServe Excel issue on Excel 2013 and higher


Sydi Server not working with Excel 2013 or Excel 2016
 
Sydiserver 2.4 released http://networklore.com/sydi/

When using the overview module sydi-overview.vbs I was getting errors;

sydi-overview2-4.vbs(596, 2) Microsoft VBScript runtime error: Subscript out of range

Great error message!

Turns out Excel 2013 opens 1 sheet by default, so you need to add in sheets 2 and 3 by copying in line 598 a few times.

Code lines for original are (590 to 601)
objExcel.Cells.Select
    objExcel.Cells.EntireColumn.AutoFit 
    objExcel.ActiveWindow.SplitRow = 0.8
    objExcel.ActiveWindow.FreezePanes = True
objExcel.Range("A1").Select
objExcel.Sheets(1).Name = "Computers"
objExcel.Sheets(2).Name = "WMI Programs"
objExcel.Sheets(3).Name = "Registry Programs"
objExcel.Sheets.Add ,objExcel.Sheets(3) ' Add a new sheet after the last one
objExcel.Sheets(4).Name = "Processes"
objExcel.Sheets.Add ,objExcel.Sheets(4) ' Add a new sheet after the last one
objExcel.Sheets(5).Name = "OS Distribution Data"

New lines need to be as follows (590 to 603)
objExcel.Cells.Select
    objExcel.Cells.EntireColumn.AutoFit 
    objExcel.ActiveWindow.SplitRow = 0.8
    objExcel.ActiveWindow.FreezePanes = True
objExcel.Range("A1").Select
objExcel.Sheets(1).Name = "Computers"
objExcel.Sheets.Add ,objExcel.Sheets(1) ' Add a new sheet after the last one
objExcel.Sheets(2).Name = "WMI Programs"
objExcel.Sheets.Add ,objExcel.Sheets(2) ' Add a new sheet after the last one
objExcel.Sheets(3).Name = "Registry Programs"
objExcel.Sheets.Add ,objExcel.Sheets(3) ' Add a new sheet after the last one
objExcel.Sheets(4).Name = "Processes"
objExcel.Sheets.Add ,objExcel.Sheets(4) ' Add a new sheet after the last one
objExcel.Sheets(5).Name = "OS Distribution Data"﻿

Thanks to Patrick Ogenstad  http://networklore.com/ for the tools
