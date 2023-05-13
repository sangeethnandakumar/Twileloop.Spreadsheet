<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/sangeethnandakumar/Twileloop.SpreadSheet">
    <img src="https://iili.io/HUaMukB.png" alt="Logo" width="80" height="80">
  </a>

  <h2 align="center"> Twileloop SpreadSheet</h2>
  <h4 align="center"> Single API | Cross-Format | Free & Fast </h4>
</div>

## About
A cross format spreadsheet accessor that empowers you to effortlessly read, write, copy, and move data across popular spreadsheet formats like Google Sheets and Microsoft Excel.

## License
> Twileloop.SpreadSheet is licensed under the MIT License. See the LICENSE file for more details.

#### This library is absolutely free. If it gives you a smile, A small coffee would be a great way to support my work. Thank you for considering it!
[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/sangeethnanda)

# Usage
***To get started, You have to install atleast 2 packages:***

- The core `Twileloop.SpreadSheet` package
- A driver package for your desired spreadsheet (Microsoft Excel, Google Sheet etc...)

> **Note**
> ***In the backstage, Twileloop.SpreadSheet uses NPOI to connect with Excel files and Google.Apis.Sheets.v4 to connect with Google Sheets***

<hr/>

## 1. Install Core Package
```bash
dotnet add package Twileloop.SpreadSheet
```

## 2. Install Driver Packages (One or More)

> There is no need to install all these driver packages, If you only prefer to work with Microsoft Excel, ignore the Google Sheets driver package

| Driver | To Use | Install Package   
| :---: | :---:   | :---:
| <img src="https://iili.io/HUaMOEG.png" alt="Logo" height="30"> | Google Sheet | `dotnet add package Twileloop.SpreadSheet.GoogleSheet`  
| <img src="https://iili.io/HUaM8Yl.png" alt="Logo" height="30"> | Microsoft Excel | `dotnet add package Twileloop.SpreadSheet.MicrosoftExcel`  

### Supported Features

| Feature     | Microsoft Excel | Google Sheets
| ---      | ---       | ---
| Plan Text Reads | ✅ | ✅
| Plan Text Writes | ✅ | ✅
| Switch Sheets | ✅ | ✅
| Text Formatting | ✅ | ✅
| Cell Formatting | ✅ | ✅
| Border Formatting | 🚧 | 🚧
| Cell Merging | 🚧 | 🚧
| Image Reads | 🚧 | 🚧
| Image Writes | 🚧 | 🚧
| Formulas | ❌ | ❌
| Draw Graph | ❌ | ❌

✅ - Available &nbsp;&nbsp;&nbsp; 
🚧 - Work In Progress &nbsp;&nbsp;&nbsp; 
❌ - Not Available

## 3. Initialize Driver(s) 
Once installed packages, Initialize your drivers

```csharp
    using Twileloop.SpreadSheet.GoogleSheet;
    using Twileloop.SpreadSheet.MicrosoftExcel;

    //Step 1: Initialize your prefered spreadsheet drivers
    var excelDriver = new MicrosoftExcelDriver(new MicrosoftExcelOptions
    {
        FileLocation = "<YOUR_EXCEL_FILE_LOCATION>"
    });
    
    var googleSheetsDriver = new GoogleSheetDriver(new GoogleSheetOptions
    {
        SheetsURI = new Uri("<YOUR_GOOGLE_SHEETS_URL>"),
        Credential = "secrets.json" //Location of your credential file
    });
```
> **Warning**
> ***If planning to use Excel fomat, Avoid opening your spreadsheet in Microsoft Excel at the same time while Twileloop.SpreadSheet is using it***

> **Warning**
> ***If planning to use Google Sheets, You have to:***
1. Create a service account
1. Download the credentials `secrets.json` from GCP console
1. Enable Google Sheets API in your GCP portal
1. Then share your Google Sheet with that service account's email id with Editor permission (for write access)

The above process is out of scope to explain here in detail.

Here's a good video tutorial (Upto 3:07): https://www.youtube.com/watch?v=fxGeppjO0Mg

## 4. Get An Accessor
Once driver(s) are initialized, Create an accessor to access the spreadsheet

```csharp
    using Twileloop.SpreadSheet.Factory;

    //Step 2: Use that driver to build a spreadsheet accessor
    var excelAccessor = SpreadSheetFactory.CreateAccessor(excelDriver);
    var googleSheetAccessor = SpreadSheetFactory.CreateAccessor(googleSheetsDriver);
```

An accessor wil give you 3 handles:
- **Reader** => Use this handle to read from your spreadsheets
- **Writer** => Use this handle to write to your spreadsheets
- **Controller** => Use this handle to control your spreadsheets

## 5. Load WorkSheet
First step is to load your prefered sheet by controlling the spreadsheet
> You must load a worksheet using the Controller before reading or writing to your spreadsheet

```csharp
    //Step 3: Now this accessor can Read/Write and Control spreadsheet. Let's open Sheet1
    using (excelAccessor)
    {
        excelAccessor.Controller.LoadSheet("Sheet1");
    }
    
    using (googleSheetAccessor)
    {
        excelAccessor.Controller.LoadSheet("Sheet1");
    }
```

## 6. Read SpreadSheet
Reading is as simple as this

```csharp
    //Step 4: Different Ways To Read Data
    using (excelAccessor)
    {
        //Load prefered sheet
        excelAccessor.Controller.LoadSheet("Sheet1");
    
        //Read a single cell
        string data1 = excelAccessor.Reader.ReadCell(1, 1);
        string data2 = excelAccessor.Reader.ReadCell("A10");
    
        //Read a full row in bulk
        string[] data3 = excelAccessor.Reader.ReadRow(1);
        string[] data4 = excelAccessor.Reader.ReadRow("C9");
    
        //Read a full column in bulk
        string[] data5 = excelAccessor.Reader.ReadColumn(1);
        string[] data6 = excelAccessor.Reader.ReadColumn("D20");
    
        //Select an area and extract data in bulk
        DataTable data7 = excelAccessor.Reader.ReadSelection(1, 1, 10, 10);
        DataTable data8 = excelAccessor.Reader.ReadSelection("A1", "J10");
    }
```

> If you're using Google Sheet, It's recommended to use any bulk reads/writes operations, Because in case of Google Sheets calling `ReadCell()` multiple times is not efficient as it fires multiple API calls to Google to read cells.

> Bulk reads/writes will fire only once and get data in one go. If you just need to read a single cell, Feel free to use `ReadCell()` since it makes sense in a read and drop situation


## 7. Write SpreadSheet
Writing is as simple as this

```csharp
    //Step 5: Different Ways To Write Data
    using (googleSheetAccessor)
    {
        googleSheetAccessor.Controller.LoadSheet("Sheet1");
    
        //Write a single cell
        googleSheetAccessor.Writer.WriteCell(1, 1, "Country");
        googleSheetAccessor.Writer.WriteCell("C17", "Country");
    
        //Write a full row in bulk
        googleSheetAccessor.Writer.WriteRow(1, new string[] { "USA", "China", "Russia", "India" });
        googleSheetAccessor.Writer.WriteRow("A1", new string[] { "USA", "China", "Russia", "India" });
    
        //Write a full column in bulk
        googleSheetAccessor.Writer.WriteColumn(1, new string[] { "USA", "China", "Russia", "India" });
        googleSheetAccessor.Writer.WriteColumn("B22", new string[] { "USA", "China", "Russia", "India" });
    
        //Select an area and write a grid in bulk
        DataTable grid = new DataTable();
        grid.Columns.Add("Rank");
        grid.Columns.Add("Powerfull Militaries");
    
        grid.Rows.Add(1, "USA");
        grid.Rows.Add(2, "China");
        grid.Rows.Add(3, "Russia");
        grid.Rows.Add(4, "India");
        grid.Rows.Add(5, "France");
    
        googleSheetAccessor.Writer.WriteSelection(1, 1, grid);
        googleSheetAccessor.Writer.WriteSelection("D20", grid);
    }
```

## 8. Read/Write Multiple SpreadSheets In One Go
Open multiple spreadsheets in one go by cascading accessors then move data in between

```csharp
    //Read and write both spreadsheets at once
    using (excelAccessor)
    {
        using (googleSheetAccessor)
        {
            //Step 1: Open both spreadsheets
            excelAccessor.Controller.LoadSheet("Sheet1");
            googleSheetAccessor.Controller.LoadSheet("Sheet1");
    
            //Step 2: Read from excel
            DataTable excelData = excelAccessor.Reader.ReadSelection("A1", "D10");
    
            //Step 3: Then write it to Google Sheet
            googleSheetAccessor.Writer.WriteSelection("C1", excelData);                    
        }
    }
```

## 9. Sheets Controls
Create one or more sheets, Get all sheets or find active sheet name

```csharp
    //Create one or more new sheets
    excelAccessor.Controller.CreateSheets("Sheet1", "Sheet2", "Sheet3");
    googleSheetAccessor.Controller.CreateSheets("Sheet1", "Sheet2");

    //Get list of sheets
    var allExcelSheets = excelAccessor.Controller.GetSheets();
    var allGoogleSheetSheet = googleSheetAccessor.Controller.GetSheets();

    //Get active sheet name
    var activeExcelSheet = excelAccessor.Controller.GetActiveSheet();
    var googleSheetSheet = googleSheetAccessor.Controller.GetActiveSheet();
```

## 10. Styling And Formatting
Styling is easy as hell. Just define all your different styles/formatting globally and apply it for a selected cell range

A formatting can have 3 types
- Text Formatting
- Cell Formatting
- Border Formatting

> Keep `NULL` for whichever format type you don't want to change


```csharp

    //Define your formatting, Let's say for titles
    var titleFormat = new Formatting
    {
        //Text formatting
        TextFormating = new TextFormating
        {
            Bold = false,
            Italic = true,
            Underline = false,
            Size = 15,
            HorizontalAlignment = HorizontalAllignment.RIGHT,
            VerticalAlignment = VerticalAllignment.BOTTOM,
            Font = "Impact",
            Color = System.Drawing.Color.White,
        },
        //Cell formatting
        CellFormating = new CellFormating
        {
            BackgroundColor = System.Drawing.Color.IndianRed
        },
        //Border formatting
        BorderFormating = new BorderFormating
        {
            TopBorder = true,
            LeftBorder = true,
            RightBorder = true,
            BottomBorder = true,
            BorderType = BorderType.SOLID,
            Thickness = 5
        }
    };

    //Then simply apply it as needed for a cell range
    excelAccessor.Writer.ApplyFormatting(1, 1, 10, 4, titleFormat);
    googleSheetAccessor.Writer.ApplyFormatting(1, 1, 10, 4, titleFormat);
```
