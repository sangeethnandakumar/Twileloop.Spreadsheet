<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/sangeethnandakumar/Twileloop.SpreadSheet">
    <img src="https://iili.io/HUaMukB.png" alt="Logo" width="80" height="80">
  </a>

  <h2 align="center"> Twileloop SpreadSheet</h2>
  <h4 align="center"> Free | Cross-Format | Fast </h4>
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


## 1. Install Core Package
```bash
dotnet add package Twileloop.SpreadSheet --version <LATEST VERSION>
```

## 2. Install Driver Packages (One or More)

> There is no need to install all these driver packages, If you only prefer to work with Microsoft Excel, ignore the Google Sheets driver package


| To Use | Install Package   
| :---:   | :---:
| Google Sheet | `dotnet add package Twileloop.SpreadSheet.GoogleSheet --version <LATEST VERSION>`  
| Microsoft Excel | `dotnet add package Twileloop.SpreadSheet.MicrosoftExcel --version <LATEST VERSION>`  

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
        Credential = "secrets.json"
    });
```

> If you're planning to use Google Sheets, You have to:
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

## 6. Different Ways To - Read Data
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


## 8. Different Ways To - Write Data
Writing is as simple as this

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
