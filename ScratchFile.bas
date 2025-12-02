Attribute VB_Name = "ScratchFile"
Option Explicit

Sub FormatasText()
Attribute FormatasText.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FormatasText Macro
'

'
    Selection.NumberFormat = "@"
    ActiveWorkbook.Save
    Range("A10:E27").Select
    Selection.ClearContents
    Columns("L:L").Select
    Selection.ClearContents
    Range("D11").Select
    Selection.Font.Bold = False
    Range("G23").Select
    Sheets("Test").Select
    ActiveWindow.SmallScroll Down:=-24
    Sheets("Cryptographic").Select
    Sheets.Add
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Cyber Roles"
    Columns("A:I").Select
    Selection.NumberFormat = "@"
    Range("A1").Select
    ActiveSheet.Paste
    Range("D9").Select
    Sheets("December 2022").Select
    ActiveWindow.SmallScroll Down:=-12
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Cyber Security Roles"
    ActiveWindow.SmallScroll Down:=-78
    Range("D2:D25").Select
    Selection.Copy
    Sheets("Cyber Roles").Select
    Range("K1").Select
    ActiveWindow.SmallScroll Down:=-21
    Columns("K:K").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "@"
    Range("K1").Select
    Sheets("December 2022").Select
    Selection.Copy
    Sheets("Cyber Roles").Select
    ActiveSheet.Paste
    Range("A7:A31").Select
    Selection.ClearContents
    Range("D7:D30").Select
    Selection.ClearContents
    Range("A1:G4").Select
    Application.Run "PERSONAL.XLSB!fourdigitControls"
    Range("A1:G4").Select
    ActiveCell.FormulaR1C1 = "' 0714"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "' 0724"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "' 0725"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "' 0726"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "' 0718"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "' 0732"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "'0717"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "'0731"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "' 0720"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "' 0734"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "'1618"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "'0027"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "' 0735"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "' 733"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "' 0733"
    Range("M1:M24").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-15
    Range("A1:G4").Select
    Application.Run "PERSONAL.XLSB!TrimSingleQuote"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "'0714"
    Range("A1:G4").Select
    Selection.ClearContents
    ActiveSheet.Paste
    Application.Run "PERSONAL.XLSB!fourdigitControls"
    Range("E13").Select
    ActiveWorkbook.Save
    Range("A1:G4").Select
    Application.Run "PERSONAL.XLSB!fourdigitControls"
    Range("A7:D31").Select
    Selection.ClearContents
    Range("M1:M27").Select
    Selection.ClearContents
    Range("A1:G4").Select
    Application.Run "PERSONAL.XLSB!fourdigitControls"
    Application.Run "PERSONAL.XLSB!fourdigitControls"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "714"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "724"
    Columns("A:G").Select
    Selection.NumberFormat = "@"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "714"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "0714"
    Range("A1:G4").Select
    Application.Run "PERSONAL.XLSB!fourdigitControls"
    Range("P24").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("A7:D30").Select
    Selection.ClearContents
    Range("M1:M25").Select
    Selection.ClearContents
    Range("R24").Select
    ActiveWindow.SmallScroll Down:=-21
    Sheets("Test").Select
    ActiveWindow.SmallScroll Down:=-24
    Range("A11:D26").Select
    Selection.ClearContents
    Columns("L:L").Select
    Selection.ClearContents
    Range("G7").Select
    Selection.Style = "Normal"
    Range("F17").Select
    Sheets("Cyber Roles").Select
    Range("A7:D17").Select
    Selection.ClearContents
    Range("M1:M24").Select
    Selection.ClearContents
    Range("M1:M28").Select
    Selection.ClearContents
    Range("A7:D34").Select
    Selection.ClearContents
    Range("K22").Select
    ActiveWindow.SmallScroll Down:=-27
    Columns("A:G").Select
    Selection.NumberFormat = "General"
    Range("A1").Select
    ActiveSheet.Paste
    Columns("K:K").Select
    Selection.NumberFormat = "General"
    Sheets("December 2022").Select
    Selection.Copy
    Sheets("Cyber Roles").Select
    Range("K1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.NumberFormat = "General"
    Columns("K:K").Select
    Sheets("December 2022").Select
    Selection.Copy
    Sheets("Cyber Roles").Select
    Range("K1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "0.00"
    Range("K2:K24").Select
    Selection.NumberFormat = "0.00"
    Range("M4").Select
    ActiveWorkbook.Save
    ActiveWindow.SmallScroll Down:=-30
    Range("T8").Select
    ActiveWindow.SmallScroll Down:=-60
    Sheets("Cryptographic").Select
    Sheets.Add
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "Cyber Incidents0576    1625    1626    0"
    Range("A1").Select
    ActiveSheet.Paste
    Range("D10").Select
    Sheets("Cyber Incidents0576 1625    1626    0").Select
    Sheets("Cyber Incidents0576 1625    1626    0").Name = "Cyber Incidents"
    Range("J27").Select
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Cyber Security Incidents"
    Range("D26:D69").Select
    Selection.Copy
    Sheets("Cyber Incidents").Select
    Range("J1").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("N8").Select
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Data Transfers"
    Sheets("Test").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("IRAP Core Applicability").Select
    Sheets.Add
    Sheets("Sheet4").Select
    Sheets("Sheet4").Name = "Data Transfers"
    Range("A1").Select
    ActiveSheet.Paste
    Range("G8").Select
    Sheets("December 2022").Select
    Range("D838:D851").Select
    Selection.Copy
    Sheets("Data Transfers").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("M9").Select
    ActiveWorkbook.Save
    Sheets("IRAP Core Applicability").Select
    Sheets.Add
    Sheets("Sheet5").Select
    Sheets("Sheet5").Name = "Network"
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Networking"
    ActiveWindow.SmallScroll Down:=-117
    Range("D639:D707").Select
    Selection.Copy
    Sheets("Network").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("R11").Select
    ActiveWindow.SmallScroll Down:=-24
    Sheets("IRAP Core Applicability").Select
    Sheets.Add
    Sheets("Sheet6").Select
    Sheets("Sheet6").Name = "Email"
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("December 2022").Select
    Selection.Copy
    Sheets("Email").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("M10").Select
    ActiveWindow.SmallScroll Down:=24
    Rows("69:69").Select
    ActiveWindow.SmallScroll Down:=-81
    Range("A6:F79").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-33
    Range("N6").Select
    ActiveWindow.SmallScroll Down:=-15
    Columns("J:L").Select
    Selection.Delete Shift:=xlToLeft
    Range("O14").Select
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Email"
    ActiveWindow.SmallScroll Down:=-78
    Range("D613:D638").Select
    Selection.Copy
    Sheets("Email").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("M9").Select
    ActiveWorkbook.Save
    Sheets("IRAP Core Applicability").Select
    Sheets.Add
    Sheets("Sheet7").Select
    Sheets("Sheet7").Name = "Outsourcing"
    Range("A1").Select
    ActiveSheet.Paste
    Range("J1").Select
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Outsourcing"
    ActiveWindow.SmallScroll Down:=-57
    Range("D43:D77").Select
    Selection.Copy
    Sheets("Outsourcing").Select
    ActiveSheet.Paste
    Range("N10:O11").Select
    Range("O11").Activate
    Sheets("IRAP Core Applicability").Select
    Sheets.Add
    Sheets("Sheet8").Select
    Sheets("Sheet8").Name = "Security Doco"
    Range("A1").Select
    ActiveSheet.Paste
    Range("F11").Select
    Sheets("December 2022").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Security Documentation"
    ActiveWindow.SmallScroll Down:=-15
    Range("D78:D87").Select
    Selection.Copy
    Sheets("Security Doco").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("C11").Select
    ActiveWorkbook.Save
    Sheets("December 2022").Select
    Range("G855").Select
    ActiveWindow.SmallScroll ToRight:=2
    Sheets("IRAP Core Applicability").Select
    Sheets.Add
    Sheets("Sheet9").Select
    Sheets("Sheet9").Name = "Physical Security"
    Range("A1").Select
    ActiveSheet.Paste
    Range("J1").Select
    Sheets("Cyber Roles").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("December 2022").Select
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Physical Security"
    ActiveWindow.SmallScroll Down:=-15
    Range("D88:D98").Select
    Selection.Copy
    Sheets("Physical Security").Select
    ActiveSheet.Paste
    Range("M13").Select
    ActiveWorkbook.Save
    ActiveWindow.TabRatio = 0.774
    Sheets("DEC New ISM Controls ").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("SEP New ISM Controls Debrief").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Epics Completed ").Select
    ActiveWindow.SelectedSheets.Delete
    ActiveWindow.SmallScroll Down:=-144
    Sheets("Support").Select
    ActiveWindow.SelectedSheets.Delete
    ActiveWindow.LargeScroll ToRight:=-1
    Sheets("SDE SoA Representation Options").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("IRAP Core Applicability").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("deleted controls").Select
    ActiveWindow.SelectedSheets.Delete
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets.Add after:=ActiveSheet
    Sheets("Sheet10").Select
    Sheets("Sheet10").Name = "Personnel Security"
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Personnel Security"
    ActiveWindow.SmallScroll Down:=-12
    Range("D99:D150").Select
    Selection.Copy
    Sheets("Personnel Security").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("M6").Select
    ActiveWindow.SmallScroll Down:=-3
    Sheets("December 2022").Select
    Sheets.Add after:=ActiveSheet
    Sheets("Sheet11").Select
    Sheets("Sheet11").Move after:=Sheets(12)
    Sheets("Sheet11").Select
    ActiveSheet.Paste
    Sheets("Sheet11").Select
    Sheets("Sheet11").Name = "Comms Infra"
    Range("J1").Select
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Communications Infrastructure"
    ActiveWindow.SmallScroll Down:=-12
    Range("D151:D202").Select
    Selection.Copy
    Sheets("Comms Infra").Select
    ActiveSheet.Paste
    Range("J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-33
    Range("O5").Select
    ActiveWindow.SmallScroll Down:=-45
    Sheets("Comms Infra").Select
    With ActiveWorkbook.Sheets("Comms Infra").Tab
        .Color = 5287936
        .TintAndShade = 0
    End With
    Range("T27").Select
    Sheets("Physical Security").Select
    With ActiveWorkbook.Sheets("Physical Security").Tab
        .Color = 5287936
        .TintAndShade = 0
    End With
    Range("S27").Select
    Sheets("Outsourcing").Select
    With ActiveWorkbook.Sheets("Outsourcing").Tab
        .Color = 5287936
        .TintAndShade = 0
    End With
    Range("N35").Select
    Sheets("Cyber Roles").Select
    ActiveWindow.SmallScroll Down:=-24
    Range("B6").Select
    ActiveWindow.SmallScroll Down:=-15
    Sheets("Cyber Roles").Select
    With ActiveWorkbook.Sheets("Cyber Roles").Tab
        .Color = 5287936
        .TintAndShade = 0
    End With
    Range("F32").Select
    Sheets("Data Transfers").Select
    ActiveWorkbook.Save
    Sheets("December 2022").Select
    Sheets.Add after:=ActiveSheet
    Sheets("Sheet12").Select
    Sheets("Sheet12").Move after:=Sheets(13)
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveSheet.Paste
    Sheets("Sheet12").Select
    Sheets("Sheet12").Name = "Comms Systems"
    Range("T14").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Communications Systems"
    ActiveWindow.SmallScroll Down:=-18
    Range("D203:D235").Select
    Selection.Copy
    Sheets("Comms Systems").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("O8").Select
    ActiveWorkbook.Save
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("December 2022").Select
    Sheets.Add after:=ActiveSheet
    Sheets("Sheet13").Select
    Sheets("Sheet13").Move after:=Sheets(14)
    ActiveSheet.Paste
    Sheets("Sheet13").Select
    Sheets("Sheet13").Name = "Evaluated Products"
    Range("U30").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Evaluated Products"
    ActiveWindow.SmallScroll Down:=-18
    Range("D273:D278").Select
    Selection.Copy
    ActiveWindow.TabRatio = 0.856
    Sheets("Evaluated Products").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("J11").Select
    Sheets("Personnel Security").Select
    Range("A10:E10").Select
    Selection.Copy
    Sheets("Evaluated Products").Select
    Range("A3").Select
    ActiveSheet.Paste
    Range("E3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "0"
    Range("E8").Select
    ActiveWorkbook.Save
    Sheets("Evaluated Products").Select
    With ActiveWorkbook.Sheets("Evaluated Products").Tab
        .Color = 5287936
        .TintAndShade = 0
    End With
    Range("T31").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets.Add after:=ActiveSheet
    Range("Y30").Select
    Sheets("Sheet14").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Sheet14").Select
    Sheets("Sheet14").Name = "ICT Equipment"
    Range("J5").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for ICT Equipment"
    ActiveWindow.SmallScroll Down:=-12
    Range("D279:D312").Select
    Selection.Copy
    Sheets("ICT Equipment").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("N5").Select
    ActiveWindow.SmallScroll Down:=-12
    ActiveWorkbook.Save
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("ICT Equipment").Select
    Sheets.Add after:=ActiveSheet
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Security Doco").Select
    ActiveCell.SpecialCells(xlLastCell).Select
    Sheets(Array("Security Doco", "Physical Security", "Personnel Security", _
        "Comms Infra", "Comms Systems", "Evaluated Products", "ICT Equipment", "Sheet15")). _
        Select
    Sheets("Sheet15").Activate
    ActiveSheet.Paste
    Sheets(Array("Security Doco", "Physical Security", "Personnel Security", _
        "Comms Infra", "Comms Systems", "Evaluated Products", "ICT Equipment", "Sheet15")). _
        Select
    Sheets("Sheet15").Activate
    Sheets("Sheet15").Name = "Media"
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Media"
    Range("D313:D366").Select
    Selection.Copy
    Sheets("Media").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("M5").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("F17").Select
    ActiveWindow.SmallScroll Down:=33
    Sheets.Add after:=ActiveSheet
    ActiveWindow.ScrollWorkbookTabs Sheets:=-3
    Sheets("Sheet16").Select
    ActiveSheet.Paste
    Sheets("Sheet16").Select
    Sheets("Sheet16").Name = "System Hardening"
    Range("J1").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-4
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for System Hardening"
    Range("D367:D493").Select
    Selection.Copy
    Sheets("System Hardening").Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("U4").Select
    ActiveWindow.SmallScroll Down:=12
    ActiveWorkbook.Save
    Sheets.Add after:=ActiveSheet
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-3
    Sheets("Sheet17").Select
    ActiveSheet.Paste
    Sheets("Sheet17").Select
    Sheets("Sheet17").Name = "System Mgt"
    Range("J1").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for System Management"
    Range("D494:D547").Select
    Selection.Copy
    Sheets("System Mgt").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-6
    Range("R21").Select
    ActiveWindow.SmallScroll Down:=-36
    ActiveWorkbook.Save
    ActiveWindow.SmallScroll Down:=9
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Evaluated Products").Select
    Sheets.Add after:=ActiveSheet
    ActiveSheet.Paste
    Sheets("Sheet18").Select
    Sheets("Sheet18").Name = "System Mon"
    Range("P5").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for System Monitoring"
    ActiveWindow.SmallScroll Down:=-6
    Range("D415:D611").Select
    ActiveWindow.SmallScroll Down:=-6
    Selection.Copy
    Sheets.Add after:=ActiveSheet
    Sheets("System Mon").Select
    Application.CutCopyMode = False
    Sheets("System Mon").Move after:=Sheets(20)
    Range("J1").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-7
    Sheets("December 2022").Select
    Selection.Copy
    Sheets("System Mon").Select
    ActiveSheet.Paste
    ActiveWorkbook.Save
    Sheets.Add after:=ActiveSheet
    ActiveSheet.Paste
    Sheets("Sheet20").Select
    Sheets("Sheet20").Name = "Software Dev"
    Sheets("System Mon").Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Software Development"
    Range("D557:D584").Select
    Selection.Copy
    Sheets(Array("December 2022", "Sheet19", "Cyber Roles")).Select
    Sheets("Cyber Roles").Activate
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    Sheets("Email").Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    Range("J1").Select
    ActiveSheet.Paste
    ActiveWorkbook.Save
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets.Add after:=ActiveSheet
    Sheets("Sheet21").Select
    Sheets.Add after:=ActiveSheet
    Sheets("Sheet21").Select
    Sheets("Sheet21").Name = "Database Sys1425  1269    1277    1270"
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Database Sys1425    1269    1277    1270").Select
    Sheets("Database Sys1425    1269    1277    1270").Name = "Database Systems"
    Range("A1").Select
    Range(Selection, Cells(1)).Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    Range("A571").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Database Systems"
    Range("D585:D612").Select
    Selection.Copy
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Previous.Select
    Range("J1").Select
    ActiveSheet.Paste
    Range("E24").Select
    ActiveWorkbook.Save
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    Sheets("Sheet22").Select
    Sheets("Sheet22").Name = "Gateways"
    Range("A1").Select
    ActiveSheet.Paste
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Gateways"
    ActiveWindow.SmallScroll Down:=-9
    Range("D775:D851").Select
    Selection.Copy
    Sheets("December 2022").Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    Range("J1").Select
    ActiveSheet.Paste
    Columns("L:L").Select
    Selection.ClearContents
    Range("A11:F81").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-18
    Range("J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-51
    Range("R11").Select
    ActiveWindow.SmallScroll Down:=-12
    ActiveWorkbook.Save
    ActiveWindow.ScrollWorkbookTabs Sheets:=-10
    Sheets("Comms Systems").Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    Sheets.Add after:=ActiveSheet
    Sheets("Sheet23").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-11
    Sheets("Evaluated Products").Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Paste
    Sheets("Sheet23").Select
    Sheets("Sheet23").Name = "Enterprise Mobility"
    Range("W32").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-10
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Enterprise Mobility"
    ActiveWindow.SmallScroll Down:=-87
    Range("D236:D272").Select
    Selection.Copy
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    ActiveSheet.Next.Select
    Range("J1").Select
    ActiveSheet.Paste
    Sheets("Enterprise Mobility").Select
    With ActiveWorkbook.Sheets("Enterprise Mobility").Tab
        .Color = 5287936
        .TintAndShade = 0
    End With
    Range("T28").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-12
    ActiveWorkbook.Save
    Sheets("Cyber Incidents").Select
    Range("V22").Select
    Sheets("Data Transfers").Select
    ActiveWindow.SmallScroll Down:=-18
    Sheets("Sheet19").Select
    ActiveWindow.SelectedSheets.Delete
    Range("K32").Select
    Sheets("Cyber Incidents").Select
    ActiveWindow.SmallScroll Down:=-9
    Sheets("Cryptographic").Select
    ActiveWindow.SmallScroll Down:=42
    Range("L68").Select
    ActiveWindow.SmallScroll Down:=-42
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("A11:F27").Select
    Selection.ClearContents
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.SmallScroll Down:=-12
    Range("J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-51
    Range("O15").Select
    ActiveWindow.SmallScroll Down:=-30
    Sheets("December 2022").Select
    ActiveSheet.Range("$A$1:$AE$851").AutoFilter Field:=1, Criteria1:= _
        "Guidelines for Cryptography"
    Sheets("Cryptographic").Select
    Columns("A:G").Select
    Selection.NumberFormat = "General"
    Columns("A:G").Select
    Selection.ClearContents
    Range("A1").Select
    ActiveSheet.Paste
    Range("D15").Select
    ActiveWindow.SmallScroll Down:=33
    Range("L67").Select
    ActiveWindow.SmallScroll Down:=-63
    Range("A11:D11").Select
    Selection.Font.Bold = False
    Range("G15").Select
    ActiveWindow.SmallScroll Down:=-18
    Sheets("Cryptographic").Select
    Sheets("Cryptographic").Name = "Cryptography"
    Range("G26").Select
    ActiveWindow.SmallScroll Down:=-90
    Sheets("Data Transfers").Select
    ActiveWindow.SmallScroll Down:=-27
    Sheets("December 2022").Select
    ActiveWorkbook.Save
End Sub
