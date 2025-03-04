# Twileloop.SpreadSheet Documentation  

## Overview  

- **Twileloop.SpreadSheet** is a very simple & interesting library for reading and writing various spreadsheet formats effortlessly.  
- It is **not** a low-level implementation but a wrapper around **NPOI** and **Google Sheets API**, abstracting their complexities into a unified, easy-to-use API.  
- The goal is to provide a single, simplified interface for handling both Excel and Google Sheets without dealing with intricate details.  
- Designed for quick and hassle-free spreadsheet creation and access.  
- Supports seamless data transfer between different spreadsheet formats.  
- Currently, only basic editing and styling features are available; advanced elements like images and graphics are not yet supported.  

### **One Code --> Export to Excel and/or Google Sheets effortlessly.**

> With more driver additions, More and more spreadsheet kinds can be supported later
----

## Installation

To get started, install the **core package** and the **driver package(s)** for the spreadsheet formats you want to use.

### 1. Install the Core Package
```bash
dotnet add package Twileloop.SpreadSheet
```

### 2. Install Driver Packages (1 or More)
Install the driver(s) for the formats you plan to use:

| Driver | Format | Install Command |
|--------|--------|-----------------|
| <img src="https://iili.io/HUaMOEG.png" alt="Google Sheets Logo" height="30"> | Google Sheets | `dotnet add package Twileloop.SpreadSheet.GoogleSheet` |
| <img src="https://iili.io/HUaM8Yl.png" alt="Microsoft Excel Logo" height="30"> | Microsoft Excel | `dotnet add package Twileloop.SpreadSheet.MicrosoftExcel` |

---

## Getting Started

### 1. Initialize Drivers

Start by initializing the drivers for the spreadsheet formats you want to work with.

```csharp
using Twileloop.SpreadSheet.GoogleSheet;
using Twileloop.SpreadSheet.MicrosoftExcel;

// Initialize Microsoft Excel Driver
var excelDriver = new MicrosoftExcelDriver(new MicrosoftExcelOptions(filePath));

// Initialize Google Sheets Driver
var googleSheet = new GoogleSheetDriver(new GoogleSheetOptions(
    sheetsURI: new Uri("https://docs.google.com/spreadsheets/d/1YWqL4_jmGhtpj--ZBLRe598w7IXDCvzL0UWHU_wZMqU/edit?gid=0#gid=0"),
    sheetName: "MySheet",
    jsonCredentialContent: File.ReadAllText("secrets.json"),
    bulkUpdate: false
));
```

> **Note for Google Sheets**:
> - Create a service account in Google Cloud Platform (GCP).
> - Download the `secrets.json` credentials file from the GCP console.
> - Enable the Google Sheets API in your GCP console.
> - Share your Google Sheet with the service account's email and assign it as an "Editor" for write permissions.

---

# For GoogleSheetDriver, Setting `bulkUpdate` to `true`, Makes Writes To GoogleSheets Faster, But This Is Now Expirimental

### bulkUpdate: false 
![image](https://raw.githubusercontent.com/sangeethnandakumar/Twileloop.Spreadsheet/refs/heads/master/doc/PerfNonBulk.webp)

### bulkUpdate: true
![image](https://raw.githubusercontent.com/sangeethnandakumar/Twileloop.Spreadsheet/refs/heads/master/doc/PerfBulk.webp)
---


### 2. Create an Adapter

Once the drivers are initialized, create an adapter to interact with the spreadsheet.

```csharp
using Twileloop.SpreadSheet.Factory;

// Create adapters for Excel and Google Sheets
ISpreadSheetAdapter excelAdapter = SpreadSheetFactory.CreateAdapter(excelDriver);
ISpreadSheetAdapter gsheetAdapter = SpreadSheetFactory.CreateAdapter(googleSheet);
```

### 3. Initialize the Workbook

Before performing any operations, initialize the workbook.

```csharp
adapter.Controller.InitialiseWorkbook(); //Mandatory step
```

---

### 4. Create and Open Sheets

You can create new sheets and open them for reading or writing. Remember, You must open a sheet to do any operations

> Once a sheet is opened all below commands execute in that opened sheet. If you're dealing with multiple sheets, Call `adapter.Controller.OpenSheet("B");` again as needed whenever you need to change sheet and rest of code need to execute in new sheet

```csharp
// Create a new sheet
adapter.Controller.CreateSheets("A"); //Optional, If creating a new spreadsheet

// Open a sheet
adapter.Controller.OpenSheet("A"); //Mandatory
```

# READING FROM SPREADSHEET

# WRITING INTO SPREADSHEET

### 5. Writing Data

You can write data to cells, rows, columns, or tables.

> You can use `"A1"` notation to address a cell or `(row, col)` notation as well. Both works on your convenience on most functions

#### Write Individual Cells

> Use below if you need to tweek specific cells or write to few cells. For a list of data either use WriteRow or WriteColumn as they're bulk writes and more efficient than WriteCell

```csharp
adapter.Writer.WriteCell("A1", "Write");  //Address notation based access
adapter.Writer.WriteCell((1, 2), "Individual"); //(row, col) notation based access, This is not an index and starts with 1 not 0
adapter.Writer.WriteCell("C1", "Cells");
```

#### Write Rows

> Use below if you need to write rows or cols in one go. For a bigger set of data, preprare your data into a DataTable & use WriteTable instead as it's more efficient then WriteRow or WriteColumn

```csharp
adapter.Writer.WriteRow("A3", new[] { "Col 1", "Col 2", "Col 3", "Col 4" });
```

> You can also pass a style if needed

```csharp
var myStyle = new StyleBuilder()
    .WithFont("Arial")
    .WithTextColor(Color.AliceBlue)
    .WithTextAllignment(HorizontalTxtAlignment.CENTER, VerticalTxtAlignment.BOTTOM)
    .WithBackgroundColor(Color.Black)
    .Build();

adapter.Writer.WriteRow("A3", new[] { "Col 1", "Col 2", "Col 3", "Col 4" }, myStyle);
```

#### Write Columns

> Passing style is optional, You can simply avoid passing it if focusing only on data

```csharp
adapter.Writer.WriteColumn("A7", new[] { "Row 1", "Row 2", "Row 3", "Row 4" }, myStyle);
```

#### Write Tables

> If you need to write huge dataset, The best way is to convert your data into a .NETs built in DataTable and write. This is much more efficient than writing cell by cell or row/col wise

```csharp
var table = new DataTable();
table.Columns.Add("ID");
table.Columns.Add("Name");
table.Columns.Add("Age");
table.Columns.Add("City");
table.Columns.Add("Salary");
table.Rows.Add(1, "John Doe", 28, "New York", 55000);

adapter.Writer.WriteTable("A12", table, myStyle);
```

---

### 6. Formatting and Styling

Apply styles to cells, rows, columns, or tables.

#### Create Styles
```csharp
var headingStyle = new StyleBuilder()
    .Bold()
    .WithFontSize(18)
    .WithFont("Arial")
    .WithTextColor(Color.Blue)
    .WithTextAllignment(HorizontalTxtAlignment.LEFT, VerticalTxtAlignment.TOP)
    .WithBackgroundColor(Color.LightBlue)
    .Build();
```

#### Apply Styles

> Here `A1` to `E1` is a grid selection rectanglular area where this style applies.

```csharp
adapter.Writer.ApplyStyling("A1", "E1", headingStyle);
```

#### Apply Borders

> Here `A12` to `E21` is a grid selection rectanglular area where this border applies. (Outer border)
> 
```csharp
adapter.Writer.ApplyBorder("A12", "E21", new BorderStyling
{
    TopBorder = true,
    LeftBorder = true,
    RightBorder = true,
    BottomBorder = true,
    BorderType = BorderType.SOLID,
    BorderColor = Color.OrangeRed,
    Thickness = BorderThickness.Thick
});
```

---

### 7. Resizing Rows and Columns

Adjust the size of rows and columns.

```csharp
// Resize a row
adapter.Writer.ResizeRow("A1", 50);

// Resize a column
adapter.Writer.ResizeColumn("A1", 40);
```

---

### 8. Saving the Workbook

After making changes, save the workbook.

```csharp
adapter.Controller.SaveWorkbook(); 
```
