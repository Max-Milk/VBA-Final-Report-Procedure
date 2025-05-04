# VBA-Final-Report-Procedure

# 📊 Automating Final Reports with Excel VBA

This repository demonstrates how to use Excel VBA to automate the generation and formatting of quarterly and yearly reports.

## 📁 Files Included

- `QuarterlyReport_RawData.xlsx`: Sample raw data for quarterly reports.
- `QuarterlyReport.xlsm`: Macro-enabled workbook with built-in VBA scripts.

## ⚙️ Features

In `QuarterlyReport.xlsm`, the following macros have been implemented:

- `AddHeaders`: Adds headers to each worksheet.
- `FormatData`: Applies formatting to the data range.
- `AutoSum`: Inserts a SUM formula at the bottom of the data.

These macros are then combined into a single automated workflow using a loop.

## 🔁 VBA Script to Automate All Sheets

The `FinalReportLoop` macro loops through all worksheets (except the last one: `"Yearly Report"`), runs the other macros (`AddHeaders`, `FormatData`, `AutoSum`), and consolidates all data into the `"Yearly Report"` sheet.

### 📜 VBA Code Snippet

```vba
Public Sub FinalReportLoop()
    Dim i As Integer
    i = 1
    Do While i <= Worksheets.Count - 1
        Worksheets(i).Select
        AddHeaders
        FormatData
        AutoSum

        ' Copy the current data
        Range("A1").Select
        Selection.CurrentRegion.Select
        Selection.Copy

        ' Select the final report worksheet
        Worksheets("Yearly Report").Select

        ' Find the next empty row
        Range("A30000").Select
        Selection.End(xlUp).Select
        ActiveCell.Offset(3, 0).Select

        ' Paste the new data in
        ActiveSheet.Paste

        i = i + 1
    Loop

    Columns("C:F").EntireColumn.AutoFit
End Sub
````

## ▶️ How to Run the Macro

1. Open `QuarterlyReport.xlsm` in Excel.
2. Go to **Developer** → **Visual Basic** → `Module1` to view or edit the VBA code.
3. Alternatively, press `F8` to step through the code line-by-line for learning.
4. To run the macro:

   * Go to **Developer** → **Macros**
   * Select `FinalReportLoop` and click **Run**

## 📝 Tip for Using the Raw Data File

You can download `QuarterlyReport_RawData.xlsx`, save it as a `.xlsm` file, paste in the VBA code, and run `FinalReportLoop` to test the automation on new data.

## 📧 Contact

* **LinkedIn**: [Max Nguyen Hoang Minh](https://www.linkedin.com/in/max-nguyen-hoang-minh)
* **Email**: [maxnguyenhoangminh@gmail.com](mailto:maxnguyenhoangminh@gmail.com)

```


