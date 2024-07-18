---
title: Fixed Excel 2003 VBA Runtime Error 9 Subscript Out of Range | Stellar
date: 2024-07-17T17:15:43.130Z
updated: 2024-07-18T17:15:43.130Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2003 VBA Runtime Error 9 Subscript Out of Range
excerpt: This article describes Fixed Excel 2003 VBA Runtime Error 9 Subscript Out of Range
keywords: repair .xltm,repair excel,repair .xlb,repair excel 2003,repair damaged .xls,repair excel 2013,repair damaged .xltx
thumbnail: https://thmb.techidaily.com/e086106512b2e9d753e081998408f58f33739fa1f0e7070951ebc50a7e45bfdb.jpg
---

## \[Fixed\] Excel VBA Runtime Error 9: Subscript Out of Range

**Summary:** The runtime error 9 in Excel usually occurs when you use different objects in a code or the object you are trying to use is not defined. This post will discuss the reasons behind the Excel VBA error "Subscript out of Range” and the solutions to resolve the issue. It will also mention an Excel repair tool that can help fix the error if it occurs due to corruption in worksheet.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Many users have reported encountering the error “Subscript out of range” (runtime error 9) when using VBA code in Excel. The error often occurs when the object you are referring to in a code is not available, deleted, or not defined earlier. Sometimes, it occurs if you have declared an array in code but forgot to specify the DIM or ReDIM statement to define the length of array.

<!-- affiliate ads begin -->
<a href="https://ursime.pxf.io/c/5597632/2048972/16384" target="_top" id="2048972"><img src="//a.impactradius-go.com/display-ad/16384-2048972" border="0" alt="" width="1200" height="900"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2048972/16384" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## **Causes of VBA Runtime Error 9: Subscript Out Of Range**

The error ‘Subscript out of range’ in Excel can occur due to several reasons, such as:

- Object you are trying to use in the VBA code is not defined earlier or is deleted.
- Entered a wrong declaration syntax of the array.
- Wrong spelling of the variable name.
- Referenced a wrong array element.
- Entered incorrect name of the worksheet you are trying to refer.
- Worksheet you trying to call in the code is not available.
- Specified an invalid element.
- Not specified the number of elements in an array.
- Workbook in which you trying to use VBA is corrupted.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=4715391&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/7f687767ccf20fcea1c9dc4a5adc2326/Digisigner_banner_728_x_90_color_version.png" border="0"></a>
<!-- affiliate ads end -->
## **Methods to Fix Excel VBA Error ‘Subscript out of Range’**

Following are some workarounds you can try to fix the runtime error 9 in Excel.

### **Method 1: Check the Name of Worksheet in the Code**

Sometimes, Excel throws the runtime error 9: Subscript out of range if the name of the worksheet is not defined correctly in the code. For example – When trying to copy content from one Excel sheet (emp) to another sheet (emp2) via VBA code, you have mistakenly mentioned wrong name of the worksheet (see the below code).

```
Private Sub CommandButton1_Click()
Worksheets("emp").Range("A1:E5").Select
Selection.Copy
Worksheets("emp3").Activate
Worksheets("emp3").Range("A1:E5").Select
ActiveSheet.Paste
Application.CutCopyMode = False
End Sub
```

<!-- affiliate ads begin -->
<a href="https://parisrhonecom.sjv.io/c/5597632/1896607/21553" target="_top" id="1896607"><img src="//a.impactradius-go.com/display-ad/21553-1896607" border="0" alt="" width="750" height="422"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1896607/21553" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![VBA Error Subscript Out Of Range-When Incorrect Name](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/VBA-error-subscript-out-of-range-when-incorrect-name.jpg)

When you run the above code, the Excel will throw the Subscript out of range error.

So, check the name of the worksheet and correct it. Here are the steps:

- Go to the **Design** tab in the **Developer** section.
- Double-click on the **Command** button.
- Check and modify the worksheet name (e.g. from “emp” to “emp2”).

![Modified Code From emp to emp2](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/modified-code-from-emp-to-emp2.jpg)

- Now run the code.
- The content in ‘emp’ worksheet will be copied to ‘emp2’ (see below).

![Content Copied From emp to emp2](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/content-copied-from-emp-to-emp2.jpg)

<!-- affiliate ads begin -->
<a href="https://ursime.pxf.io/c/5597632/2092236/16384" target="_top" id="2092236"><img src="//a.impactradius-go.com/display-ad/16384-2092236" border="0" alt="" width="1920" height="329"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2092236/16384" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### Method 2: Check the Range of the Array

The VBA error “Subscript out of range” also occurs if you have declared an array in a code but didn’t specify the number of elements. For example – If you have declared an array and forgot to declare the array variable with elements, you will get the error (see below):

<!-- affiliate ads begin -->
<a href="https://versadesk.pxf.io/c/5597632/1892108/21290" target="_top" id="1892108"><img src="//a.impactradius-go.com/display-ad/21290-1892108" border="0" alt="" width="1080" height="1080"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1892108/21290" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Runtime Error 9 When Not Declared Array](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/run-time-error-9-when-not-declared-array.jpg)

To fix this, specify the array variable:

```
Sub FillArray()
Dim curExpense(364) As Currency
Dim intI As Integer
For intI = 0 to 364
curExpense(intI) = 20
Next
End Sub
```

### **Method 3: Change Macro Security Settings**

The Runtime error 9: Subscript out of range can also occur if there is an issue with the macros or macros are disabled in the Macro Security Settings. In such a case, you can check and change the macro settings. Follow these steps:

- Open your Microsoft Excel.
- Navigate to **File > Options > Trust Center**.
- Under **Trust Center**, select **Trust Center Settings**.
- Click **Macro Settings**, select **Enable all macros**, and then click **OK**.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3851655&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
<!-- affiliate ads end -->
![Macro Settings In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/macro-settings-in-trust-center.jpg)

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BGeneral%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/general-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
### **Method 4: Repair your Excel File**

The name or format of the Excel file or name of the objects may get changed due to corruption in the file. When the objects are not identified in a VBA code, you may encounter the Subscript out of range error. You can use the Open and Repair utility in Excel to repair the corrupted file. To use this utility, follow these steps:

- In your MS Excel, click **File > Open**.
- Browse to the location where the affected file is stored.
- In the **Open** dialog box, select the corrupted workbook.
- In the **Open** dropdown, click on **Open and Repair**.
- You will see a prompt asking you to repair the file or extract data from it.
- Click on the **Repair** option to extract the data as much as possible. If **Repair** button fails, then click **Extract** button to recover data without formulas and values.

If the “Open and Repair” utility fails to repair the corrupted/damaged macro-enabled Excel file, then try an advanced Excel repair tool, such as Stellar Repair for Excel. It can easily repair severely corrupted Excel workbook and recover all the items, including macros, cell comments, table, charts, etc. with 100% integrity. The tool is compatible with all versions of Microsoft Excel.

## **Conclusion**

You may experience the “Subscript out of range” error while using VBA in Excel. You can follow the workarounds discussed in this blog to fix the error. If the Excel file is corrupt, then you can use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to repair the file. It’s a powerful software that can help fix all the issues that occur due to corruption in the Excel file. It helps to recover all the data from the corrupt Excel files (.xls, .xlsx, .xltm, .xltx, and .xlsm) without changing the original formatting. The tool supports Excel 2021, 2019, 2016, and older versions.


## Excel Repair Tool to Repair Corrupt Excel files (.XLS/.XLSX)

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3851691&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
<!-- affiliate ads end -->
## When to Use Stellar Excel File Repair Tool?

Unable to Open an Excel File Due to Invalid Extension?

![Unable-to-Open-an-Excel-File-Due-to-Invalid-Extension](https://www.stellarinfo.com/public/image/catalog/usecase/excel-repair/Unable-to-Open-an-Excel-File-Due-to-Invalid-Extension.jpg)

You may face an error - "Excel cannot open the file .xlsx” in Excel 2021, 2019, 2016, etc., leading to data loss. This error occurs when you try to open corrupt Excel file or an invalid file format. Using the correct extension can resolve the issue, if there is no corruption. However, you need an Excel repair tool if the file is corrupt. Stellar Repair for Excel can repair the corrupt file and recover all objects in intact form.

[_Learn More_ ![arrow](https://www.stellarinfo.com/public/image/catalog/v6/arrow.svg)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

Excel File Not Opening Due to Corruption?

<!-- affiliate ads begin -->
<a href="https://funwhole.sjv.io/c/5597632/1702887/17189" target="_top" id="1702887"><img src="//a.impactradius-go.com/display-ad/17189-1702887" border="0" alt="" width="1000" height="1000"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1702887/17189" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Is-MDF-File-Header-Corrupted](https://www.stellarinfo.com/public/image/catalog/usecase/excel-repair/Excel-File-Not-Opening-Due-to-Corruption.jpg)

You cannot open an Excel file if it is corrupted. For example, opening an Excel file created in a lower version like Excel 2007 in Excel 2010 or later version can throw a corruption error message. Or, the file may open in a ‘protected view,’ not allowing any write operations. The Excel repair tool from Stellar provides a comprehensive solution to fix corrupt Excel files across all versions, including Excel 2021, 2019, 2016, 2013, and older.

[_Learn More_ ![arrow](https://www.stellarinfo.com/public/image/catalog/v6/arrow.svg)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

Excel Found Unreadable Content?

<!-- affiliate ads begin -->
<a href="https://godlikehost.sjv.io/c/5597632/1920054/21774" target="_top" id="1920054"><img src="//a.impactradius-go.com/display-ad/21774-1920054" border="0" alt="" width="320" height="100"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1920054/21774" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Excel-Found-Unreadable-Content](https://www.stellarinfo.com/public/image/catalog/usecase/excel-repair/Excel-Found-Unreadable-Content.jpg)

You may encounter an error message – “Excel found unreadable content in filename.xls”, with a message to recover the contents of the workbook. Clicking ‘Yes’ to recover the contents may lead to loss of formatting, replacement of formulas, and inconsistencies. Stellar Phoenix Excel Repair software now Stellar Repair for Excel can scan the workbook and recover its contents.

[_Learn More_ ![arrow](https://www.stellarinfo.com/public/image/catalog/v6/arrow.svg)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

KEY FEATURES FOR REPAIR FOR EXCEL

### Software Important Capabilities

![Repair Large-sized Excel Files](https://www.stellarinfo.com/image/catalog/feature-icon/Excel/Repairs-Large-Size-Excel-Files.png)

### Repair Large-sized Excel Files

Stellar Repair for Excel software previously known as Stellar Phoenix Excel Repair can repair & fix corrupt Excel files of any size. It removes corruption from individual objects, fixes the damage, and restores the Excel file back to its original state. The Excel repair tool can repair multiple Excel files in a batch.  
[Learn More](https://www.stellarinfo.com/support/kb/index.php/article/repair-corrupt-excel-file)

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2082538/7443" target="_top" id="2082538"><img src="//a.impactradius-go.com/display-ad/7443-2082538" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2082538/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Resolves All Excel Corruption Errors](https://www.stellarinfo.com/image/catalog/feature-icon/Excel/Resolve-All-Excel-Corruption-Errors.png)

<!-- affiliate ads begin -->
<a href="https://engwe.pxf.io/c/5597632/2093504/25579" target="_top" id="2093504"><img src="//a.impactradius-go.com/display-ad/25579-2093504" border="0" alt="" width="1200" height="1200"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2093504/25579" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### Resolves All Excel Corruption Errors

This Excel file repair tool fixes all types of Excel corruption errors, such as unrecognizable format, Excel found unreadable content in name.xls, Excel cannot open the file filename.xlsx, file name is not valid, the Excel file is corrupt and cannot be opened, etc. It provides a comprehensive solution for fixing Excel file issues.  
[Learn More](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

<!-- affiliate ads begin -->
<a href="https://united.elfm.net/c/5597632/517826/4704" target="_top" id="517826"><img src="//a.impactradius-go.com/display-ad/4704-517826" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://united.elfm.net/i/5597632/517826/4704" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Preview the Repaired Excel File](https://www.stellarinfo.com/image/catalog/feature-icon/Excel/Preview-of-Excel-File.png)

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2075461/7443" target="_top" id="2075461"><img src="//a.impactradius-go.com/display-ad/7443-2075461" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2075461/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### Preview the Repaired Excel File

The software shows a preview of the repaired Excel file and its recoverable contents in the main interface. This functionality allows you to verify the data in your repaired Excel file, including all of its objects, before saving the file. The Excel File Recovery software helps in determining the final state of data you will receive after repairing the corrupted Excel file.  
[Learn More](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

<!-- affiliate ads begin -->
<a href="https://store.bitdefender.com/affiliate.php?ACCOUNT=BITLATIN&AFFILIATE=108875&PATH=http%3A%2F%2Fwww.bitdefender.com%2Fbusiness%3FAFFILIATE%3D108875%26RESOURCE%3D30%2525%2BOff%2Ball%2BGravityZone%2BProducts"><img src="https://www.bitdefender.com/content/dam/bitdefender/business/campaign/1200X628.png" border="0"></a>
<!-- affiliate ads end -->
![Recovers All Excel file Objects](https://www.stellarinfo.com/image/catalog/feature-icon/Excel/Recovers-All-Excel-Objects.png)

<!-- affiliate ads begin -->
<a href="https://printrendy.pxf.io/c/5597632/1453720/17020" target="_top" id="1453720"><img src="//a.impactradius-go.com/display-ad/17020-1453720" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1453720/17020" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### Recovers All Excel file Objects

The software repairs the corrupt Excel file and recovers all objects, including tables, charts, series trendline, conditional formatting rules, and properties of the worksheet. The software also recovers embedded functions, group & subtotal, engineering formulas, numbers, texts, rules, etc. It recovers Excel file data in its intact form.

Reviews & Feedback

### Recommendation by Microsoft MVPs

OTHER IMPORTANT FEATURES

### Know your Product Better

![Option to Find Excel Files](https://www.stellarinfo.com/image/catalog/feature-icon/Excel/Allows-to-Search-for-Excel-Files.png)

#### Option to Find Excel Files

Stellar Excel repair software helps users unaware of the Excel file location to search for all the Excel files on the computer. It provides ‘Find’ option to quickly locate and list all the Excel files for repair. You can select single or multiple files from the list that you want to repair.

![Stellar Toolkit for File Repair](https://www.stellarinfo.com/image/catalog/feature-icon/Excel/Stellar-Toolkit-for-File-Repair.png)

<!-- affiliate ads begin -->
<a href="https://store.iobit.com/order/checkout.php?PRODS=4596923&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/184260348236f9554fe9375772ff966e/ascscan_468X60.png" border="0"></a>
<!-- affiliate ads end -->
### Stellar Toolkit for File Repair

Stellar Toolkit for File Repair provides essential tools to repair corrupt Office files via a single interface. It comprises tools like MS Excel Repair, MS Word Repair, MS PowerPoint Repair, and PDF Repair to fix corrupt documents while maintaining the original file format.  
[Learn More](https://tools.techidaily.com/stellardata-recovery/file-repair-toolkit/)

HOW TO USE STELLAR REPAIR FOR EXCEL

### Screenshots & Video

<!-- affiliate ads begin -->
<a href="https://turtlebeachus.sjv.io/c/5597632/1988416/23719" target="_top" id="1988416"><img src="//a.impactradius-go.com/display-ad/23719-1988416" border="0" alt="" width="600" height="600"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1988416/23719" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![](https://www.stellarinfo.com/public/image/catalog/screenshot/excel-repair/stellar-repair-for-excel-main-interface.png)

PRICING

### Choose The Best Plan

Excel Repair

Repairs corrupted Excel files with 100% integrity.

- Repairs XLS, XLSX, XLTM, XLTX, and XLSM files
- Repairs multiple Excel files
- Previews the repaired file
- Supports Excel 2021 & older versions

File Repair Toolkit

Repairs corrupted Excel, Word, PowerPoint, & PDF files.

- Repairs XLS, XLSX, XLTM, XLTX, and XLSM files
- Repairs multiple files
- Previews the repaired file
- Supports Excel 2021 & older versions
- Repairs .DOC & .DOCX files
- Repairs .PPT, .PPTX, & .PPTM files
- Repairs corrupted PDF file

Best Seller

File Repair Toolkit Technician

Repairs corrupted Excel, Word, PowerPoint, & PDF files up to 3 systems.

- Repairs XLS, XLSX, XLTM, XLTX, and XLSM files
- Repairs multiple files
- Previews the repaired file
- Supports Excel 2021 & older versions
- Repairs .DOC & .DOCX files
- Repairs .PPT, .PPTX, & .PPTM files
- Repairs corrupted PDF file

CUSTOMER REVIEWS

### You're in Good Hands

![left quote](https://www.stellarinfo.com/public/image/catalog/v6/left-quote.png)

![right quote](https://www.stellarinfo.com/public/image/catalog/v6/right-quote.png)

AWARDS & REVIEWS

### Most tested. Most awarded

![q1](https://www.stellarinfo.com/images/v7/q1.png) ![q1](https://www.stellarinfo.com/images/v7/q2.png)

DATA SHEET

<!-- affiliate ads begin -->
<a href="https://turbotech.pxf.io/c/5597632/1450763/17212" target="_top" id="1450763"><img src="//a.impactradius-go.com/display-ad/17212-1450763" border="0" alt="" width="2560" height="1440"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1450763/17212" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### Technical Specifications

<!-- affiliate ads begin -->
<a href="https://thefitville.pxf.io/c/5597632/1526796/15852" target="_top" id="1526796"><img src="//a.impactradius-go.com/display-ad/15852-1526796" border="0" alt="" width="1200" height="628"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1526796/15852" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![product Icon](https://www.stellarinfo.com/image/catalog/feature-icon/Excel/excel-repair-product.svg)

About Product

**Stellar Repair for Excel**

<table><tbody><tr><td><strong>Version:</strong></td><td>6.0.0.7</td></tr><tr><td><strong>License:</strong></td><td>Single System</td></tr><tr><td><strong>Edition:</strong></td><td>Standard, Technician, &amp; Toolkit</td></tr><tr><td><strong>Language Supported:</strong></td><td>English</td></tr><tr><td><strong>Release Date:</strong></td><td>February, 2024</td></tr></tbody></table>

<table><tbody><tr><td><strong>Processor:</strong></td><td>Intel compatible (x64-based processor)</td></tr><tr><td><strong>Memory:</strong></td><td>4 GB minimum<span> (8 GB recommended)</span></td></tr><tr><td><strong>Hard Disk:</strong></td><td>250 MB of Free Space</td></tr><tr><td><strong>Operating System:
(64 Bit only)</strong></td><td>Windows 11, 10, 8.1, 8, 7</td></tr></tbody></table>

USEFUL ARTICLES

### Product Related Articles

How do I repair multiple Excel files by using Stellar Repair for Excel software?

After launching the software, click Select File button in the Home tab. Next, click Browse and select the checkbox against all the Excel files you need to repair. Then, click the Repair button to start repairing all the Excel files.

[_Learn More_](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

How do I see the Preview of repaired Excel file using the Demo version of the software?

Browse and select the file(s) to repair. The software will start scanning the Excel files once you click the Repair button. Next, it will display the files in the left pane. You can preview their contents in the right pane.

[_Learn More_](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

How do I find the recovered Excel file?

The software saves the repaired file with the prefix “Recovered” at the user-specified location. You can find the recovered file using the Search box utility in the taskbar.

[_Learn More_](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

![Stellar Repair for Excel](https://www.stellarinfo.com/image/boxshot/Stellar-Repair-for-Excel.png)

### Start Using Stellar Repair for Excel Today

- Trusted by Millions of Users
- Awarded by Top Tech Media
- 100% Safe & Secure to Use

Free download to scan and preview all recoverable Excel data.


## \[Solved\] : How to Fix MS Excel Crash Issue

Microsoft [Excel may stop responding](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), hang, freeze, or stop working due to several reasons, such as in compatible add-ins. In such a case, you may receive one of the following error messages.

- Excel has stopped working

<!-- affiliate ads begin -->
<a href="https://turtlebeacheu.sjv.io/c/5597632/1996818/23722" target="_top" id="1996818"><img src="//a.impactradius-go.com/display-ad/23722-1996818" border="0" alt="" width="600" height="600"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1996818/23722" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Excel has stopped working](https://www.stellarinfo.com/public/image/catalog//article/email-repair/exchange/excelnew1.jpg)

- Excel is not responding

![Excel is not responding](https://www.stellarinfo.com/image/catalog/article/excelnew2.jpg)

- A problem caused the program to stop working correctly. Windows will close the program and notify you if a solution is available.

![A problem caused the MS Excel to stop working correctly](https://www.stellarinfo.com/image/catalog/article/excelnew3.jpg)

## Why Does Excel Keep Crashing?

If Excel keeps crashing on your PC while opening a workbook, saving Excel file, scrolling or editing cells, etc., it indicates a problem with your Excel program or the Excel file.

Microsoft Excel may crash due to any one or more reasons given below,

-  Incompatible Add-Ins
- Outdated MS Excel program
- Conflict with other programs or antivirus tool
-  Excel file created by third party software
- Problem with network connection
-  Combination of Cell formatting and stylings
- Problem with MS Office installation
- Partially damaged or corrupt Excel file

## Problems Caused by Excel Crash Issue

Microsoft Excel crash may cause damage to Excel file and also lead to Excel (XLS/XLSX) file corruption.

Such corrupt Excel files can't be opened or accessed via MS Excel app. If you try to access a corrupt Excel file, MS Excel may fail to open the file or stop responding and crash. Additionally, you may receive the following or similar error message,

![Excel files can't be opened or accessed](https://www.stellarinfo.com/image/catalog/article/excelnew4.jpg)

In such a case, you should immediately try to recover the Excel file. You may do so by restoring the Excel file from backup or by using an [Excel File Repair software.](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) Otherwise, continue following this guide.

<!-- affiliate ads begin -->
<a href="https://martinic.evyy.net/c/5597632/1422856/4482" target="_top" id="1422856"><img src="//a.impactradius-go.com/display-ad/4482-1422856" border="0" alt="" width="580" height="309"/></a>
<!-- affiliate ads end -->
## How to Solve Excel Crash Issue?

Before heading to solutions, follow these troubleshooting steps to resolve the Excel Crash issue.

<!-- affiliate ads begin -->
<span id="1793213">
					<video width="1080" height="1620" style="cursor:pointer"
           poster="//a.impactradius-go.com/display-clicktoplayimage/1793213.jpeg"
           onclick="if(!this.playClicked){this.play();this.setAttribute('controls',true);this.playClicked=true;}">
	   <source src="//a.impactradius-go.com/display-ad/19135-1793213">
	   <img src="//a.impactradius-go.com/display-clicktoplayimage/1793213.jpeg" style="border: none; height: 100%; width: 100%; object-fit: contain">
	</video>
	<div style="width:1080px;text-align:center"><a href="javascript:window.open(decodeURIComponent('https%3A%2F%2Ftinyland.pxf.io%2Fc%2F5597632%2F1793213%2F19135'), '_blank');void(0);">Click here</a></div>
</span>
<img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793213/19135" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Step 1: Copy File to Local Drive**

If you are trying to access and edit or view an Excel file from a network drive, try moving the file to local drive. This will help you find if there is something wrong with the file or the network.

<!-- affiliate ads begin -->
<a href="https://shop.mondly.com/affiliate.php?ACCOUNT=ATISTUDI&AFFILIATE=108875&PATH=https%3A%2F%2Fwww.mondly.com%3FAFFILIATE%3D108875%26RESOURCE%3D%2BEducational%2B970x90%2B"><img src="https://secure.avangate.com/images/merchant/69c418c33ec2e1a4267fa9bb77fa1428/educational-970x90.gif" border="0"></a>
<!-- affiliate ads end -->
### **Step 2: Ensure Sufficient Memory**

Excel files can grow fairly large when you start adding lots of formatting and shapes. Make sure that your system has enough RAM to run the application.

![Ensure Sufficient Memory](https://www.stellarinfo.com/image/catalog/article/excelnew5.jpg)

If you often work with large Excel files and complex data values& formulas, then install 64-bit versions of MS Office. It will give you an advantage of larger processing capacities and prevent Excel from crash or freeze.

### **Step 3: Check If Excel is Open and In Use by Another Process**

Open **Task Manager** and close all processes or apps (tasks) that may be using or have access to your Excel file that you are working on. You can find this detail in status bar of Excel program at the bottom of program window.

![Task Manager](https://www.stellarinfo.com/image/catalog/article/excel6.jpg)

After closing the tasks, try to access the Excel file and check if this fixes the performance and crash problem in Excel.

### Step 4: Test and Repair Excel File

Create a copy of the Excel file and install **Stellar Repair for Excel** software. It's free to download. Scan and repair your Excel file using the software. After repair, save the Excel file at your desired location and then open the Excel file in the MS Excel program.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3727260&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
<!-- affiliate ads end -->
![Stellar Repair for Excel software](https://www.stellarinfo.com/image/catalog/article/excel7.jpg)

This should ideally fix all the issues with Excel.

However, if the Excel program still crashes, the problem lies within the system or program. Follow the solutions discussed in this guide to try to fix the Excel crash issue.

**NOTE:** To save repaired Excel file using the mentioned software, you must purchase the activation key and activate it.

## Solutions to Fix MS Excel Crash Issue

Following are some solutions to resolve problems with MS Excel such as,

- Excel not responding
- Excel won't open
- Excel keeps crashing

Follow these solutions in the given order. In case a method doesn't work, move to the next one.

### Solution 1: Restart Excel in Safe Mode

By starting MS Excel in safe mode, you can run the program without loading the Excel add-ins and with limited features. But COM add-ins are excluded.

To launch Excel in safe mode, close MS Excel and follow these steps,

- Create a shortcut of MS Excel (.exe) on Desktop
- Press and hold the Ctrl key while launching the program
- Click 'Yes' when a prompt appears to confirm

Alternatively, press Windows+R, type excel /safe and press 'Enter'. Use this to open Excel in safe mode on Windows 10, 8.1, 8, or 7 system.

![type excel /safe](https://www.stellarinfo.com/image/catalog/article/excel8.jpg)

Now try to open and access the Excel file and check if the issue is resolved. If it's not, head on to the next solution.

### Solution 2: Check and Remove Faulty Add-ins

In case Excel doesn't crash in Safe Mode, it's possible that some faulty add-ins are the culprit behind frequent Excel crash and freeze. These Excel add-ins may interfere or conflict with the Excel program.

<!-- affiliate ads begin -->
<a href="https://homestyler.sjv.io/c/5597632/2044747/22993" target="_top" id="2044747"><img src="//a.impactradius-go.com/display-ad/22993-2044747" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2044747/22993" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Check and Remove Faulty Add-ins](https://www.stellarinfo.com/image/catalog/article/excel9.jpg)

Find and remove the faulty add-in. It can resolve the issue. To do so, follow these steps,

- Restart Excel in normal mode and go to File> Options> Add-ins
- Choose COM Add-ins from the drop-down and click Go

![COM Add-ins](https://www.stellarinfo.com/image/catalog/article/excel10.jpg)

- Uncheck all the checkboxes and click OK

![Uncheck all the check boxes](https://www.stellarinfo.com/image/catalog/article/excel11.jpg)

- Restart Excel and check if the issue is resolved
-  If Excel doesn't crash or freeze anymore, open COM Add-ins and enable one add-in at a time followed by Excel restart. Then observe Excel for freeze or crash problem

This will help you find out the faulty add-in, which is causing the problem. Remove the add-in which is causing the problem to resolve the issue. If that doesn't fix, move to the next solution.

### Solution 3: Check and Install the Latest Updates

If you haven't set Windows to Download and Install Updates automatically, do it now.

Apart from updating the operating system, latest Windows updates sometimes fixes bugs for other applications installed on the system such as MS Office. Often installing an important update that you might have missed may correct the Excel crash problem.

You can also update MS Office manually. Follow these steps,

Go to File > Account

 Under Product Information, select Update Options and click Update Now

![Product Information](https://www.stellarinfo.com/image/catalog/article/excel12.jpg)

If you have installed MS Excel from Microsoft Store, open the store and update your Office applications.

NOTE: This also works if you can't open Excel file or Excel crashes after Windows upgrade from Windows 7 or Windows 8/8.1 to Windows 10.

After installing the latest MS Office updates, check if Excel works fine. If not, head to the next solution.

### Solution 4: Clear Conditional Formatting Rules

If a sheet is causing Excel to freeze or crash, there might be a problem with that particular sheet. In such a case, you may try clearing the Conditional Formatting rules. The steps are as follows,

- Under Home, click 'Conditional Formatting > Clear Rules\> Clear Rules from Entire Sheet'

<!-- affiliate ads begin -->
<a href="https://aligracehair.sjv.io/c/5597632/2087267/19272" target="_top" id="2087267"><img src="//a.impactradius-go.com/display-ad/19272-2087267" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2087267/19272" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Conditional Formatting](https://www.stellarinfo.com/image/catalog/article/excel13.jpg)

- You may repeat this step for all other sheets in the Excel workbook
- Then click File> Save as and save the Sheet as a new file at a different location

This avoids overwriting or making changes to the original Excel file. Once done, try working on the sheet.

If this doesn't work out, move to the next solution.

### Solution 5: Remove Multiple Cell Formatting and Styles

If a workbook is being shared and edited by others on different platforms then it's possible that many cells are formatted differently. This can cause issues with Excel such as crash and freeze. It can also lead to Excel file corruption. The problem mostly occurs when a workbook contains multiple worksheets using different formatting.

You can [follow this guide](https://docs.microsoft.com/en-gb/office/troubleshoot/excel/too-many-different-cell-formats-in-excel) to remove different cell formats and styles, and then open the Excel file.

### Solution 6: Disable Microsoft Excel Animation

Animations require additional processing power and resources. By disabling animations in Excel, you may resolve Excel freeze and crash issue. This also improves MS Excel performance.

To disable the animations in MS Excel, follow these steps:

- Go to File > Options
- Click 'Advanced' and check 'Disable hardware graphics acceleration'animation

![Disable hardware graphics acceleration](https://www.stellarinfo.com/image/catalog/article/excel14.jpg)

- Click 'OK' to close the window and then restart MS Excel

This has helped many users in fixing the Excel crash issue. If it doesn't work for you, head to the next solution.

### Solution 7: Check If Excel File is Generated by a Third-Party Application

There are applications which you may have used to generate Excel files to fetch data. For instance, downloading data from Google Analytics in Excel format.

Sometimes, these Excel files are not generated correctly by such third-party apps. Thus, some features in Excel may not work as intended when you access the files in MS Excel.

In such a case, you should get in touch with the app developer for help with the file or use Stellar Repair for Excel to repair such Excel files.

### Solution 8: Check If Antivirus or Other Apps are Conflicting with MS Excel

Ensure your antivirus is up-to-date and not conflicting with MS Excel. An outdated antivirus tool may conflict with Excel which can cause the application to hang, freeze, or crash.

- Update your antivirus
- Try disabling the add-in or integration between Excel and antivirus. See if it works

Alternatively, you may disable the anti-virus tool temporarily to check if it is the culprit behind Excel performance issue and crash. If that resolves the problem, get in touch with your antivirus vendor and report the problem.

They might provide you with a better solution or workaround to fix this problem without disabling the antivirus protection.

IMPORTANT NOTE: Disabling or altering antivirus protection makes your PC vulnerable to malicious attacks and virus or malware intrusion.

### Solution 9: Clean Boot Windows to Inspect the Cause Behind Excel Crash

When Windows boot, it starts several processes, services, and application during start up automatically, which runs in the background.

These startup apps and services can interfere with other applications such as MS Excel. To find out if that's the cause behind Excel crash, you can perform a Clean Boot.

This helps you identify processes, services, or applications that are conflicting with Excel. Steps to perform Clean Boot are as follows,

- Press Windows key + R, type MSConfig, and press 'Enter'
- In System Configuration window, click on the General tab and choose Selective startup

![System Configuration](https://www.stellarinfo.com/image/catalog/article/excel15.jpg)

Uncheck 'Load startup items' and click 'OK'

After this, close all running applications and restart your PC

Check if the crash problem with Excel is resolved. Uninstall the conflicting apps or update them. If your issue is not resolved, follow the next solution.

### Solution 10: Repair or Reinstall MS Office

Repairing Office programs may also resolve Excel crash issues if caused by damaged MS Excel program or MS Office files. The steps are as follow,

- Close all MS Office apps and open the Control Panel
- Click Uninstall a program under Programs

![Uninstall a program](https://www.stellarinfo.com/image/catalog/article/excel16.jpg)

- Click on Microsoft Office and then click on the Change option
- Choose 'Quick repair' and then select 'Repair'
- Click 'Continue' to repair MS Office installation

You may also try 'Online Repair' if this fails to fix the issue. After repair, if the Excel issue persists, reinstall MS Office.

## Need More Help?

If none of the above-mentioned solutions worked for you, it indicates that the problem is not with the Excel program but with the Excel file. If you haven't tried the Stellar Repair for Excel software, do it now.

Select the Excel file which is causing the problem and repair it with the software. It's a powerful Excel repair software that can fix all the problems with Excel files (XLS/XLSX). It repairs corrupt and severely damaged Excel files.

The software is compatible with all Excel files created using MS Excel 2019, 2016, 2013, 2010, 2007, 2003 or 2000.

After repairing and saving the Excel file, you can open it in your MS Excel program and work on it without any performance issue. To know more about this software, visit [this page.](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)




## How to Fix 'Sharing Violation Error' when Saving Excel?

**Summary:** You may encounter the sharing violation error in Excel when you repeatedly save changes in a workbook. The error can occur due to different reasons. In this blog, we will discuss the possible reasons behind this sharing violation error and some effective solutions to fix it. If the issue has occurred due to corruption in Excel file, you can try the advanced Excel repair tool mentioned in the post to repair the corrupted file.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

It is not uncommon to encounter errors while working with Excel application. One common error that Excel users face is the sharing violation error that occurs while saving an Excel spreadsheet. The complete error message says, “Your changes could not be saved to file because of a sharing violation.” When this error occurs, users won’t be able to save the changes in the file. So, it is important to fix this issue as soon as possible. But before fixing the error, let’s find out the causes behind this error.

## Causes of Excel Sharing Violation Error

This error may pop up due to the below reasons:

- The file you are trying to save is corrupted.
- The Excel file is not in the trusted location.
- Sharing Wizard is disabled.
- You do not have permission to modify the Excel file.
- The Excel file is not permitted to get indexed.

<!-- affiliate ads begin -->
<a href="https://ephamedtechinc.pxf.io/c/5597632/2097467/26400?prodsku=B700" target="_top" id="2097467"><img src="//a.impactradius-go.com/display-ad/26400-2097467" border="0" alt="" width="640" height="640"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2097467/26400" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## Methods to Fix the Sharing Violation Error in Excel

You can move the affected Excel file to a new folder and save it with a different name. Then, see if it fixes the error. If it doesn’t help, you can try the below methods.

### **Method 1: Check and Change the Excel File Properties**

You can get the sharing violation error in Excel if the file attribute options, such as “File is ready for archiving” and “Allow this file to have contents indexed in addition to file properties” are disabled. You can check the File Properties and enable these options to fix the issue. Here are the steps:

- Right-click on any Excel file and select **Properties**.

![Click On Properties Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/click-on-properties-option.jpg)

- In the **Properties** window, click on the **Advanced** option.

![Click On Advanced Button On Properties Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/click-on-advanced-button-on-properties-window.jpg)

- In the **Advanced Attributes** window, select the below options under **File attributes**:
- File is ready for archiving.
- Allow this file to have contents indexed in addition to file properties.

![Select File Is Ready For Archiving Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/select-file-is-ready-for-archiving-option.jpg)

- Click **OK**.

**_Note_**_: If these options are already selected, then unselect and re-select them._

### **Method 2: Enable Sharing Wizard Option**

The error “Your changes could not be saved to file because of a sharing violation” can also occur if the sharing wizard option is disabled on your system. You can check and enable the sharing wizard option using these steps:

- Go to your system’s **Documents** folder.
- Click **View > Options > Change folders** **and search options.**

<!-- affiliate ads begin -->
<a href="https://laganoo.pxf.io/c/5597632/1657399/16446" target="_top" id="1657399"><img src="//a.impactradius-go.com/display-ad/16446-1657399" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1657399/16446" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Click View Option In Documents](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/click-view-option-in-documents.jpg)

- In the **Folder Options** window, click **View**.

<!-- affiliate ads begin -->
<a href="https://tinyland.pxf.io/c/5597632/1793214/19135" target="_top" id="1793214"><img src="//a.impactradius-go.com/display-ad/19135-1793214" border="0" alt="" width="900" height="900"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793214/19135" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![In Folder Options Click On View](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/in-folder-options-click-on-view.jpg)

- Under the **View** section, search for the “**Use Sharing Wizard**” option in the **Advanced Settings**.

![Select Use Sharing Wizard](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/select-use-sharing-wizard-1.jpg)

- If the **Use Sharing Wizard** option is unselected, select it and then click **OK**.

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2082529/7443" target="_top" id="2082529"><img src="//a.impactradius-go.com/display-ad/7443-2082529" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2082529/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Method 3: Move the Excel File to a Trusted Location**

You can encounter the sharing violation error if the file you are trying to save is not in the trusted location. You can try moving the file to a trusted location by following these steps:

- In Excel, go to **File** and then click **Options.**
- Click **Trust Center** and then click **Trust Center Settings**.

<!-- affiliate ads begin -->
<a href="https://natural-cycles.sjv.io/c/5597632/2072199/17885" target="_top" id="2072199"><img src="//a.impactradius-go.com/display-ad/17885-2072199" border="0" alt="" width="300" height="300"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2072199/17885" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Click Trust Center Settings In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/click-trust-center-settings-in-trust-center.jpg)

- In the **Trust Center** window, click **Trusted Locations** and then click **Add new location**.

<!-- affiliate ads begin -->
<a href="https://vapordna.pxf.io/c/5597632/1496243/17238" target="_top" id="1496243"><img src="//a.impactradius-go.com/display-ad/17238-1496243" border="0" alt="" width="1000" height="1221"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1496243/17238" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Click On Add New Location Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/click-on-add-new-location-option.jpg)

- In the **Add new location** window, select **Browse** to locate and choose the folder, and then click **OK**.

### **Method 4: Open Excel in Safe Mode**

Incompatible add-ins can create issues in the Excel file. To check if the sharing violation issue has occurred due to add-ins, open Excel in safe mode. To do so, follow these steps:

- Open the **Run** window using **Windows + R**.

![Type Safe Mode Command In Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/type-safe-mode-command-in-excel.jpg)

- Type **excel /safe** and click on **OK**.
- Open the affected file and then try to save the changes.
- If you are able to save the changes without any error, then it indicates add-ins are causing the issue. To fix this, you can [remove the recently downloaded add-ins](https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) (if any).

<!-- affiliate ads begin -->
<a href="https://lightailing.sjv.io/c/5597632/1725213/17190" target="_top" id="1725213"><img src="//a.impactradius-go.com/display-ad/17190-1725213" border="0" alt="" width="1000" height="1000"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1725213/17190" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### Method 5: Repair the Excel File

Corruption in Excel file can also create issue while saving the changes. In such a case, you can repair the corrupted Excel file using the inbuilt utility in Excel, named Open and Repair. Follow these steps to use this utility:

- In Excel, navigate to **File > Open > Browse**.
- In the **Open** dialog box, click on the affected Excel file.
- Click the arrow next to the **Open** button and select **Open and Repair** from the dropdown.
- Click on the **Repair** option to recover as much data from the file as possible.

If the above utility fails to fix the corrupt Excel file, then you can use a more powerful [Excel repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), such as Stellar Repair for Excel. This Excel repair tool can repair even severely corrupted or damaged Excel files (xls, .xlsx, .xltm, .xltx, and .xlsm). This tool can recover all the data from the corrupted Excel file, including images, chart sheets, formulas, etc., without changing the original format. It can help in fixing common corruption-related errors in Excel. You can download the software’s demo version to scan the corrupt file.

## **To Conclude**

Above, we have discussed some effective methods to fix the sharing violation error in Excel. This error may also occur if you try to save the Excel file in an incompatible format. So, check the format and try saving the file in a compatible format. If the error occurs due to Excel file corruption, you can [repair corrupt Excel file](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) using the Open and Repair tool. If nothing works, then download a third-party Excel repair tool, such as Stellar Repair for Excel. It is an advanced tool that can fix severely corrupted Excel files. You can install this repair tool on any Windows system.


<!-- affiliate ads begin -->
<a href="https://mushroom-supplies.sjv.io/c/5597632/1692242/18134" target="_top" id="1692242"><img src="//a.impactradius-go.com/display-ad/18134-1692242" border="0" alt="" width="834" height="592"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1692242/18134" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## Excel Stuck at Opening File 0% - Resolve Performance Issues

**Summary:** If an Excel workbook is stuck at opening file 0%, it usually indicates a problem with the Excel file and its objects. This may happen due to Excel file corruption and a few other reasons. In this post, we have discussed these reasons along with the methods to fix and prevent ‘Excel stuck at opening file 0%’ issue.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

When you open an Excel file (XLS/XLSX) in MS Excel, the program reads and then loads the file data along with all its objects and properties. While opening and loading an Excel file, MS Excel displays an “_Opening percentage_.” You won’t usually notice or see this Excel file opening progress percentage while accessing smaller worksheets.

It’s more noticeable when you open a large Excel file or workbook with multiple objects, formulae, formatting, etc. However, after opening an Excel file with double-click, if it is stuck at _Splash Screen_ with a message “**Opening: FileName.xlsx (0%)”** for a while (say 15-30 minutes) and does not progress, it indicates a problem with the Excel file, MS Excel program, or the system.  

![excel stuck at 0 percent](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/excel-stuck-at-opening-1024x576.png)

<!-- affiliate ads begin -->
<span id="1993652">
					<video width="720" height="300" style="cursor:pointer"
           poster="//a.impactradius-go.com/display-clicktoplayimage/1993652.jpeg"
           onclick="if(!this.playClicked){this.play();this.setAttribute('controls',true);this.playClicked=true;}">
	   <source src="//a.impactradius-go.com/display-ad/22993-1993652">
	   <img src="//a.impactradius-go.com/display-clicktoplayimage/1993652.jpeg" style="border: none; height: 100%; width: 100%; object-fit: contain">
	</video>
	<div style="width:720px;text-align:center"><a href="javascript:window.open(decodeURIComponent('https%3A%2F%2Fhomestyler.sjv.io%2Fc%2F5597632%2F1993652%2F22993'), '_blank');void(0);">Click here</a></div>
</span>
<img height="0" width="0" src="https://imp.pxf.io/i/5597632/1993652/22993" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## Why Excel is Stuck at Opening File 0%?

If you have encountered this error, it may happen due to one of the following issues,

1. Damaged or corrupt Excel file
2. Incompatible or faulty Excel add-ins
3. Problem with the system’s display driver
4. Damaged MS Office (Excel) application

<!-- affiliate ads begin -->
<a href="https://unicoeye.pxf.io/c/5597632/2084396/18498" target="_top" id="2084396"><img src="//a.impactradius-go.com/display-ad/18498-2084396" border="0" alt="" width="1920" height="700"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2084396/18498" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## Methods to Fix ‘Excel Stuck at Opening File 0%’ Issue

Before fixing and troubleshooting the problem, check and confirm if the Excel file is working and not corrupt. For this, you can try opening it on another PC. Now there could be two scenarios,

<!-- affiliate ads begin -->
<a href="https://propmoneyinc.pxf.io/c/5597632/1803116/14559" target="_top" id="1803116"><img src="//a.impactradius-go.com/display-ad/14559-1803116" border="0" alt="" width="859" height="859"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1803116/14559" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Scenario 1:  Excel File Does Not Open**

If the Excel file doesn’t open on another PC also, it indicates Excel file corruption. In such cases, look for the backup copy of the file, if you have downloaded it from an email or a website.

However, if there’s no backup, then you need an Excel file repair software, such as **Stellar Repair for Excel** to repair the corrupt file. This software preserves Excel file properties, such as cell formatting, formula bar, freeze panes, gridlines, etc. and helps you restore the damaged or corrupt worksheets to its original state with 100% integrity.

[![free download](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

To repair Excel file, download and launch _Stellar Repair for Excel_ software on your PC, choose the corrupt Excel (XLS/XLSX) file and click ‘**Repair’**. You can see the preview of your Excel file with all data and then save the repaired file at your desired location on the system as a new Excel file.

<!-- affiliate ads begin -->
<a href="https://twopages.pxf.io/c/5597632/1873305/18544" target="_top" id="1873305"><img src="//a.impactradius-go.com/display-ad/18544-1873305" border="0" alt="" width="1080" height="1350"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1873305/18544" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![stellar repair excel file](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Stellar-repair-for-excel.jpg)

### **Scenario 2: Excel File Is Accessible on Another PC**

If the Excel file opens successfully on another PC, then follow the troubleshooting methods below to resolve the Excel file stuck opening at 0%.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=4940317&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/333ac5d90817d69113471fbb6e531bee/sps-partnership-728x90eng.png" border="0"></a>
<!-- affiliate ads end -->
## Method 1: Open MS Excel in Safe Mode

To check if an incompatible or faulty add-in or setting is causing the error, restart MS Excel in safe mode and then open the worksheet from the MS Excel ‘**File**’ options. The steps are as follows,

1. Press **Windows+R** and type **excel.exe /safe**
2. Hit **Enter** or press ‘**OK**’ to open MS Excel in safe mode

<!-- affiliate ads begin -->
<a href="https://shop.copernic.com/order/checkout.php?PRODS=41033101&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.2checkout.com/images/merchant/8d30aa96e72440759f74bd2306c1fa3d/Copernic-2023-Affiliate-728x90-Elite.png" border="0"></a>
<!-- affiliate ads end -->
![open excel in safe mode](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/run-excel-in-safe-mode.png)

- Go to **File > Open** and then choose the Excel file to open it
- If it opens, the problem is probably caused by the add-ins. Go to **File > Options > Add-ins > Manage > COM Add-ins** and disable all the third-party add-ins

![remove faulty add in from excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/check-and-remove-faulty-add-ins-1024x540.png)

- Restart MS Excel normally and then go to **File > Open** and open the same Excel file. If it opens, the problem is solved.

However, if you want to keep the add-ins, enable one add-in at a time and open the same file to find which add-in is causing the problem. When found, remove the faulty add-in.

If it doesn’t work, head to the next solution.

## Method 2. Disable Hardware Graphics Acceleration

If you’re using hardware graphics acceleration adapter to run an external monitor, you may encounter problems with the Excel application. If the adapter is plugged in but doesn’t work correctly, Excel will usually hang on the loading screen. To resolve this problem, you will need to disable the hardware graphics acceleration adapter by following these steps,

- Quit all running instances of Excel from **Task Manager**

<!-- affiliate ads begin -->
<a href="https://unicoeye.pxf.io/c/5597632/2084399/18498" target="_top" id="2084399"><img src="//a.impactradius-go.com/display-ad/18498-2084399" border="0" alt="" width="1125" height="600"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2084399/18498" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![task manager to close program](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Check-task-manager.png)

- Launch MS Excel directly, don’t _double-click_ on the faulty workbook file to open MS Excel as it won’t open
- Click on **File > Options > Advanced**

![Disable hardware graphics acceleration](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Disable-hardware-graphics-acceleration.png)

- Under the ‘**Display**’ options, check the box ‘**Disable hardware graphics acceleration**’
- Click on ‘**OK**’

Try to open the Excel file now. If it still doesn’t work, move to the next solution.

## Method 3. Repair MS Excel Application and Install the Latest Updates

Problems within MS Excel installation could also be a source of many unknown issues. Messed up registry settings, bugged updates, and even wrong user ‘**Preferences**’ can cause your Excel application to behave unusually. The fix for all such issues is to repair the Excel installation. To do so, follow these steps,

- Open **Control Panel**
- From **Category** view, under **Programs**, select **Uninstall a program**
- Click on the MS Office and then click ‘**Change**’

<!-- affiliate ads begin -->
<a href="https://coinrule.sjv.io/c/5597632/1958374/18409" target="_top" id="1958374"><img src="//a.impactradius-go.com/display-ad/18409-1958374" border="0" alt="" width="300" height="300"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1958374/18409" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![repair ms office](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Repair-MS-Office-installation-1024x577.png)

- When prompted, click on ‘**Repair’** and then follow the instructions to complete the repair process

<!-- affiliate ads begin -->
<a href="https://zebaoaffiliateprogram.pxf.io/c/5597632/1853659/21526" target="_top" id="1853659"><img src="//a.impactradius-go.com/display-ad/21526-1853659" border="0" alt="" width="1920" height="750"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1853659/21526" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![quick repair ms office](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Quick-repair-MS-Excel.png)

**To update the MS Excel,**

- Go to **File > Account** and click on **Update options**

<!-- affiliate ads begin -->
<a href="https://electronicx.pxf.io/c/5597632/1872496/14483" target="_top" id="1872496"><img src="//a.impactradius-go.com/display-ad/14483-1872496" border="0" alt="" width="750" height="625"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1872496/14483" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![check MS Excel updates](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Check-MS-Excel-updates-1024x539.png)

- Then click ‘**Update’**

<!-- affiliate ads begin -->
<a href="https://natural-cycles.sjv.io/c/5597632/2072200/17885" target="_top" id="2072200"><img src="//a.impactradius-go.com/display-ad/17885-2072200" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2072200/17885" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Download MS Excel updates](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Install-MS-Excel-Updates-1024x539.png)

- MS Excel will start downloading the latest updates and then apply it, which might fix this Excel error

![Apply MS Excel updates](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Applying-updates-to-Excel.png)

Still, Excel stuck on processing file at 0%? That means the Excel file you’re trying to open is severely corrupted. Thus, as mentioned earlier, use **Stellar Repair for Excel** software to repair corrupt or damaged Excel (XLS/XLSX) files and restore everything to a new Excel file. With the help of some best-in-class repair algorithms, this software enables you to fix problems within Excel files and recover tables, charts, cell comments, images, formulae, sorts, and filters. It is compatible with MS Excel 2019, 2016, 2013, 2010, 2007, and 2003.

<!-- affiliate ads begin -->
<a href="https://electronicx.pxf.io/c/5597632/1872456/14483" target="_top" id="1872456"><img src="//a.impactradius-go.com/display-ad/14483-1872456" border="0" alt="" width="500" height="375"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1872456/14483" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## Conclusion

Hopefully, one of the above-mentioned solutions has helped you overcome the “Excel stuck at Opening file 0%” error and Excel hangs on opening file issues. Also, you are able to access your MS Excel worksheet now. If you face any problems with your Excel workbooks in future, remember to get to the root of the issue first. Also, inculcate the habit of backing up your critical files regularly (if possible) and keep products like **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** in mind to save the day, when nothing else works.


<!-- affiliate ads begin -->
<a href="https://zonlipartnershipprogram.pxf.io/c/5597632/1821134/17882" target="_top" id="1821134"><img src="//a.impactradius-go.com/display-ad/17882-1821134" border="0" alt="" width="320" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1821134/17882" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## How to Repair Corrupted or Damaged Excel File with Ease?

**Summary:** The Excel file is prone to corruption. Users can face several issues related to corruption. So here in this infographic, I am discussing a professional tool,- Stellar Repair for Excel, to easily repair corrupted Excel files.

<!-- affiliate ads begin -->
<a href="https://ursime.pxf.io/c/5597632/2048963/16384" target="_top" id="2048963"><img src="//a.impactradius-go.com/display-ad/16384-2048963" border="0" alt="" width="1200" height="900"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2048963/16384" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Stellar Repair for Excel is among the top choices for repairing corrupt or damaged Excel (.XLS/.XLSX) files. This [Excel recovery software](https://www.stellarinfo.com//blog/top-10-best-excel-recovery-software/) restores everything from the corrupt file to a new blank Excel file. Incoming, the information graphics complete overview of the repair process is explained in step-by-step methodology. Explore and reap the benefits of recovering corrupt or damaged Excel files.

<!-- affiliate ads begin -->
<a href="https://zonlipartnershipprogram.pxf.io/c/5597632/1596691/17882" target="_top" id="1596691"><img src="//a.impactradius-go.com/display-ad/17882-1596691" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1596691/17882" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
[![Repair Corrupt Excel Files Infographic](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2024/02/Repair-Corrupt-Excel-Files-Infographic-2-scaled.jpg)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

Very much sure about the result of using the excel file recovery tool, share your experience with us.



<!-- affiliate ads begin -->
<a href="https://store.iobit.com/order/checkout.php?PRODS=1468905&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/184260348236f9554fe9375772ff966e/ascscan_728x90.png" border="0"></a>
<!-- affiliate ads end -->
## Solutions to Repair Corrupt Excel File

**Summary:** MS Excel can throw various errors due to corrupted Excel files. This blog discusses the error messages that indicate Excel file corruption and the methods to prevent data loss due to a corrupt file. It also discusses the reasons behind the corruption in Excel file and their solutions. It also mentions a “Stellar repair for Excel” tool that can help to repair the corrupt or damaged Excel file.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Is your Excel file corrupted? And you don’t have backup of your data? There is no need to worry. There are some simple solutions to repair Excel file 2019. But before heading towards the solutions, let’s discuss the possible reasons for Excel file corruption and how you can prevent losing your data.

## **Error Messages that Indicate Excel File Corruption**

**When an Excel file gets corrupted, different error messages appear. For example:**

- “Excel found unreadable content in <filename>. Do you want to recover the content of this workbook, click Yes.”
- “Can’t find project and library.”
- “The workbook cannot be opened or repaired by Microsoft Excel because it is corrupted.”
- “Microsoft Excel has stopped working.”

## **Reasons Behind Excel File Corruption**

**The reasons for corruption in Excel file could be any of the following:**

- Improper system shutdown
- Computer virus/malware attack/Hacker attack
- Outdated anti-virus definition
- Hardware failure
- Unintentional deletion of files
- Large Excel files
- Bad sectors on storage media

## **How to Avoid Data Loss Due to Excel File Corruption?**

**Excel users should follow the below precautionary measures to prevent data loss due to Excel file corruption:**

### **1\. Create an Automatic Backup Copy**

When you create an Excel spreadsheet, it is advised to **Save As** your document, as follows:

1. In **Save As** window, click **Tools** next to **Save** option.
2. Select **General Options** from the drop-down menu.
3. Then check the dialogue box **Always create back up** and click **OK.**

<!-- affiliate ads begin -->
<a href="https://atezr.pxf.io/c/5597632/2018605/18496" target="_top" id="2018605"><img src="//a.impactradius-go.com/display-ad/18496-2018605" border="0" alt="" width="798" height="807"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2018605/18496" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Enable automatic backup by clicking Tools next to Save in the Save As window, choosing General Options, checking the Always create backup box, and clicking OK.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/general-options-1024x576.png)

This will always create a backup of your Excel. If it’s deleted or corrupted at any time, it can be recovered.

<!-- affiliate ads begin -->
<a href="https://dhgate.sjv.io/c/5597632/1678785/12108" target="_top" id="1678785"><img src="//a.impactradius-go.com/display-ad/12108-1678785" border="0" alt="" width="300" height="250"/></a>
<!-- affiliate ads end -->
### **2\. Create Recovery File at Different Time Periods**

**Steps are as follows:**

1. Go to **File** and then click **Excel** **Options**.
2. Click **Save** and then select the **Save** **Auto Recover information every** checkbox
3. Add the required minutes and location. Ensure that **Disable AutoRecover for this workbook only** box is unchecked.

<!-- affiliate ads begin -->
<h3 id="200610"><a href="https://sentrypc.7eer.net/c/5597632/200610/3022">Parental Control Software</a></h3>
<span class="text-ad-content">
	#1 Rated Parental Control Software.<br/>
	Monitor & Control all PC Activity!<br/>
		<cite style="color:green">sentrypc.com/parental-controls/</cite>
	</span><img height="0" width="0" src="https://sentrypc.7eer.net/i/5597632/200610/3022" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Access Excel Options from the File menu, navigate to Save, enable Save AutoRecover with specified minutes and location, and ensure the Disable AutoRecover for this workbook only box is unchecked.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Disable-auto-recover-1024x576.png)

## **Methods to Repair Corrupted Excel 2019 File**

**Try using these 5 methods to restore your Excel file and recover data:**

<!-- affiliate ads begin -->
<a href="https://ancheer.sjv.io/c/5597632/1657301/17326" target="_top" id="1657301"><img src="//a.impactradius-go.com/display-ad/17326-1657301" border="0" alt="" width="1920" height="933"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1657301/17326" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Method 1: ‘Open and Repair’ Excel Files**

Excel automatically opens the corrupted file in Recovery Mode. If not, you can repair Excel file manually through the following steps:

- Click on the **File** and select **Open**.

<!-- affiliate ads begin -->
<a href="https://aspironcom.sjv.io/c/5597632/1941789/21554" target="_top" id="1941789"><img src="//a.impactradius-go.com/display-ad/21554-1941789" border="0" alt="" width="650" height="800"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1941789/21554" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![File and select Open](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/1-4.png)

- Go to the location where the corrupt workbook is stored. In the **Open** window, select the corrupt file.
- Click **Open** and then select **Open and Repair**.
- In the window that opens, click **Repair**.

If the Repair option doesn’t work, you can select **Extract Data** and try to extract the values and formulae safely from the corrupt file.

### Method 2: Recover Data from Open Workbook

If you face issues while working in an Excel file, you can choose to return to the last saved version of the Excel file. For this:

- Click **File**. Then select **Open**.
- Double click on the name of the workbook (the one that is open in your Excel).
- Click **Yes** to reopen it.

![Navigate to the File menu, select Open, double-click on the open workbook's name in Excel, and confirm by clicking Yes to reopen the workbook.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Excel-reopen-error.png)

- The workbook will now appear.

_**Please note that it will show the last saved version and changes made after that won’t be recovered.**_

<!-- affiliate ads begin -->
<a href="https://store.revouninstaller.com/order/checkout.php?PRODS=28010250&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/4282ec8de8c9be897e7aff4aa231b1a4/336__280a.jpg" border="0"></a>
<!-- affiliate ads end -->
### Method 3: Set Calculation Option as Manual

You can also recover data from Excel workbooks that you’re unable to open. For this, you need to configure the **calculation option** as **manual** in Excel. You can do this through the following steps:

- Click on **File**. Select **New** and open a **Blank** workbook.
- From File, select Excel Options.

![Microsoft Excel - Home Options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Options-2.png)

- From the **Formulas** category, under the section **Calculation options**, select **Manual. Now** click **OK**.

<!-- affiliate ads begin -->
<span id="1997795">
					<video width="250" height="250" style="cursor:pointer"
           poster="//a.impactradius-go.com/display-clicktoplayimage/1997795.jpeg"
           onclick="if(!this.playClicked){this.play();this.setAttribute('controls',true);this.playClicked=true;}">
	   <source src="//a.impactradius-go.com/display-ad/23621-1997795">
	   <img src="//a.impactradius-go.com/display-clicktoplayimage/1997795.jpeg" style="border: none; height: 100%; width: 100%; object-fit: contain">
	</video>
	<div style="width:250px;text-align:center"><a href="javascript:window.open(decodeURIComponent('https%3A%2F%2Fproteahair.pxf.io%2Fc%2F5597632%2F1997795%2F23621'), '_blank');void(0);">Click here</a></div>
</span>
<img height="0" width="0" src="https://imp.pxf.io/i/5597632/1997795/23621" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Access the Formulas category, go to Calculation options, choose Manual, and confirm the changes by clicking OK.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/formula-manual.png)

- Then click **File**, and select **Open** to open the corrupted or damaged Excel file.

### Method 4: Recover Content by Using External Links

You can also recover specifically the content (leaving formulas/calculated values) from the workbook by using external references (to link Excel workbook). For this:

- Click on **File**, Select **Open**.
- Navigate to the folder that contains the corrupted workbook.
- Now, right-click on the file name of the corrupted workbook and click **Copy**.
- Click **File** button. Then, select **New** and create another blank workbook.
- In the first cell (A1), type =!A1 and press Enter.
  - Select the corrupted workbook in the **Update Values** dialogue (if it appears). Then click **OK**.
  - Select the relevant sheet in the **Select Sheet** dialogue (if it appears). Then click **OK**.

<!-- affiliate ads begin -->
<a href="https://bluettius.sjv.io/c/5597632/2027209/17108" target="_top" id="2027209"><img src="//a.impactradius-go.com/display-ad/17108-2027209" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2027209/17108" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Microsoft Excel - Dialog box](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/formula.png)

- Again, select the cell A1, go to **Home** and select **Copy**.
- Now select (start from the cell A1) an area equal to that of the data in the original workbook.
- Go to Home now and select **Paste**.
- Again, go to Home, and Copy the data (the same selection of cells).
- Go to Home, and then click on the arrow below **Paste**. Then click on **Values**.

By pasting values, you removed the links to the corrupted workbook and only the data is left behind.

### Method 5: Excel Repair Software

**If the above-mentioned methods do not help in repairing the corrupt Excel file, try an Excel repair software.**

One of the most commonly used Excel repair tools is [**Stellar Repair for Excel**.](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/?utm_source=Site_Blog&utm_medium=Site_Blog_Excel_2019_Repair&utm_campaign=Site_Blog_Excel_2019_Repair) Its trial version is available for free download, which lets you scan and preview the repaired Excel files. Once you’ve ascertained the effectiveness of the software, you can save the file after activating the software.

Here’s the complete repairing process of the corrupt Excel file

<iframe width="560" height="315" title="YouTube video player" frameborder="0" allowfullscreen="" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" nitro-og-src="https://www.youtube.com/embed/VAeGzHnETu0?si=ksZ355zGrL1qxD9r&amp;autoplay=1" nitro-lazy-src="data:text/html;https://www.youtube.com/embed/VAeGzHnETu0?si=ksZ355zGrL1qxD9r&amp;autoplay=1;base64,PGJvZHkgc3R5bGU9J3dpZHRoOjEwMCU7aGVpZ2h0OjEwMCU7bWFyZ2luOjA7cGFkZGluZzowO2JhY2tncm91bmQ6dXJsKGh0dHBzOi8vaW1nLnlvdXR1YmUuY29tL3ZpL1ZBZUd6SG5FVHUwLzAuanBnKSBjZW50ZXIvMTAwJSBuby1yZXBlYXQnPjxzdHlsZT5ib2R5ey0tYnRuQmFja2dyb3VuZDpyZ2JhKDAsMCwwLC42NSk7fWJvZHk6aG92ZXJ7LS1idG5CYWNrZ3JvdW5kOnJnYmEoMCwwLDApO2N1cnNvcjpwb2ludGVyO30jcGxheUJ0bntkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7Y2xlYXI6Ym90aDt3aWR0aDoxMDBweDtoZWlnaHQ6NzBweDtsaW5lLWhlaWdodDo3MHB4O2ZvbnQtc2l6ZTo0NXB4O2JhY2tncm91bmQ6dmFyKC0tYnRuQmFja2dyb3VuZCk7dGV4dC1hbGlnbjpjZW50ZXI7Y29sb3I6I2ZmZjtib3JkZXItcmFkaXVzOjE4cHg7dmVydGljYWwtYWxpZ246bWlkZGxlO3Bvc2l0aW9uOmFic29sdXRlO3RvcDo1MCU7bGVmdDo1MCU7bWFyZ2luLWxlZnQ6LTUwcHg7bWFyZ2luLXRvcDotMzVweH0jcGxheUFycm93e3dpZHRoOjA7aGVpZ2h0OjA7Ym9yZGVyLXRvcDoxNXB4IHNvbGlkIHRyYW5zcGFyZW50O2JvcmRlci1ib3R0b206MTVweCBzb2xpZCB0cmFuc3BhcmVudDtib3JkZXItbGVmdDoyNXB4IHNvbGlkICNmZmY7fTwvc3R5bGU+PGRpdiBpZD0ncGxheUJ0bic+PGRpdiBpZD0ncGxheUFycm93Jz48L2Rpdj48L2Rpdj48c2NyaXB0PmRvY3VtZW50LmJvZHkuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBmdW5jdGlvbigpe3dpbmRvdy5wYXJlbnQucG9zdE1lc3NhZ2Uoe2FjdGlvbjogJ3BsYXlCdG5DbGlja2VkJ30sICcqJyk7fSk7PC9zY3JpcHQ+PC9ib2R5Pg=="></iframe>

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2087407/7443" target="_top" id="2087407"><img src="//a.impactradius-go.com/display-ad/7443-2087407" border="0" alt="" width="600" height="500"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087407/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
## **Conclusion**

This post shared the reasons behind Excel file corruption and precautionary measures to prevent data loss. It also outlined different methods to repair corrupt Excel file 2019. There are several in-built utilities in Microsoft Excel to repair corrupt workbooks and recover data from it. In case these methods didn’t work, you can use Stellar Repair for Excel – an easy-to-use DIY tool that can fix all Excel corruption errors and restore data with all original properties.


<ins class="adsbygoogle"
     style="display:block"
     data-ad-client="ca-pub-7571918770474297"
     data-ad-slot="8358498916"
     data-ad-format="auto"
     data-full-width-responsive="true"></ins>
<ins class="adsbygoogle"
    style="display:block"
    data-ad-format="autorelaxed"
    data-ad-client="ca-pub-7571918770474297"
    data-ad-slot="1223367746"></ins>



<span class="atpl-alsoreadstyle">Also read:</span>
<div><ul>
<li><a href="https://instagram-video-recordings.techidaily.com/new-exploring-the-eminent-on-niche-focused-instagram-groups/"><u>[New] Exploring the Eminent on Niche-Focused Instagram Groups</u></a></li>
<li><a href="https://vp-tips.techidaily.com/new-fast-cash-on-reddit-check-out-these-top-13-skillless-strategies-for-2024/"><u>[New] Fast Cash on Reddit? Check Out These Top 13 Skillless Strategies for 2024</u></a></li>
<li><a href="https://twitter-videos.techidaily.com/new-gifify-it-free-method-for-twitter-videos-for-2024/"><u>[New] Gifify It  Free Method for Twitter Videos for 2024</u></a></li>
<li><a href="https://instagram-videos.techidaily.com/new-in-2024-discover-these-eight-handy-apps-to-upload-reels-on-instagram/"><u>[New] In 2024, Discover These Eight Handy Apps to Upload Reels on Instagram</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/new-mastering-the-art-of-vlogging-your-ultimate-journey-begins-here/"><u>[New] Mastering The Art of Vlogging  Your Ultimate Journey Begins Here</u></a></li>
<li><a href="https://some-guidance.techidaily.com/new-unveiling-the-process-of-sharing-srt-text-in-social-digital-spaces/"><u>[New] Unveiling the Process of Sharing SRT Text in Social Digital Spaces</u></a></li>
<li><a href="https://screen-video-capture.techidaily.com/updated-2024-approved-mental-gymnastics-at-the-best-puzzle-halls/"><u>[Updated] 2024 Approved  Mental Gymnastics at the Best Puzzle Halls</u></a></li>
<li><a href="https://extra-hints.techidaily.com/updated-30plus-futuristic-metaverse-musings-arvr/"><u>[Updated] 30+ Futuristic Metaverse Musings [AR/VR]</u></a></li>
<li><a href="https://extra-approaches.techidaily.com/updated-navigate-subtitle-installation-in-wmp-easily/"><u>[Updated] Navigate Subtitle Installation in WMP Easily</u></a></li>
<li><a href="https://youtube-help.techidaily.com/2024-approved-how-to-optimize-video-engagement-through-resizing-thumbnails-on-youtube/"><u>2024 Approved  How to Optimize Video Engagement Through Resizing Thumbnails on YouTube</u></a></li>
<li><a href="https://fox-boxes.techidaily.com/2024-approved-masterclass-in-making-panzoids/"><u>2024 Approved  Masterclass in Making Panzoids</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-c02-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia C02 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-c110-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia C110 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-c12-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia C12 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-c12-plus-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia C12 Plus | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-c12-pro-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia C12 Pro | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-c210-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia C210 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-c22-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia C22 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-c300-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia C300 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-c32-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia C32 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-g22-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia G22 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-g310-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia G310 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-g42-5g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia G42 5G | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nokia-xr21-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nokia XR21 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nubia-red-magic-8s-pro-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nubia Red Magic 8S Pro | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nubia-red-magic-8s-proplus-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nubia Red Magic 8S Pro+ | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nubia-red-magic-9-pro-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nubia Red Magic 9 Pro | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nubia-red-magic-9-proplus-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nubia Red Magic 9 Pro+ | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nubia-z50-ultra-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nubia Z50 Ultra | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-nubia-z50s-pro-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Nubia Z50S Pro | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-oneplus-11-5g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset OnePlus 11 5G | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-oneplus-11r-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset OnePlus 11R | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-oneplus-12-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset OnePlus 12 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-oneplus-12r-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset OnePlus 12R | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-oneplus-ace-2-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset OnePlus Ace 2 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-oneplus-ace-2-pro-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset OnePlus Ace 2 Pro | Dr.fone</u></a></li>
<li><a href="https://tiktok-videos.techidaily.com/big-deals-in-tiktok-videos-oversized-character-techniques-for-2024/"><u>Big Deals in TikTok Videos  Oversized Character Techniques for 2024</u></a></li>
<li><a href="https://howto.techidaily.com/how-to-fix-it-vivo-x-flip-wont-turn-on-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>How to Fix It Vivo X Flip Wont Turn On | Dr.fone</u></a></li>
<li><a href="https://apple-account.techidaily.com/how-to-get-the-apple-id-verification-code-on-apple-iphone-8-plus-in-the-best-ways-by-drfone-ios/"><u>How To Get the Apple ID Verification Code On Apple iPhone 8 Plus in the Best Ways</u></a></li>
<li><a href="https://change-location.techidaily.com/how-to-stop-life360-from-tracking-you-on-poco-m6-pro-5g-drfone-by-drfone-virtual-android/"><u>How to Stop Life360 from Tracking You On Poco M6 Pro 5G? | Dr.fone</u></a></li>
<li><a href="https://some-techniques.techidaily.com/in-2024-fast-and-easy-gif-transformation-ultimate-list-of-no-download-services/"><u>In 2024, Fast and Easy GIF Transformation  Ultimate List of No-Download Services</u></a></li>
<li><a href="https://android-pokemon-go.techidaily.com/in-2024-how-to-come-up-with-the-best-pokemon-team-on-htc-u23-pro-drfone-by-drfone-virtual-android/"><u>In 2024, How to Come up With the Best Pokemon Team On HTC U23 Pro? | Dr.fone</u></a></li>
<li><a href="https://some-techniques.techidaily.com/in-2024-innovative-filters-to-make-your-tiktok-stand-out/"><u>In 2024, Innovative Filters to Make Your TikTok Stand Out</u></a></li>
<li><a href="https://extra-approaches.techidaily.com/mastering-media-effects-a-comprehensive-guide-to-device-filters-for-2024/"><u>Mastering Media Effects  A Comprehensive Guide to Device Filters for 2024</u></a></li>
<li><a href="https://network-issues.techidaily.com/mastering-youtubes-green-screen-problems/"><u>Mastering YouTube's Green Screen Problems</u></a></li>
<li><a href="https://video-creation-software.techidaily.com/new-the-ultimate-list-10-adobe-premiere-elements-alternatives-for-video-enthusiasts/"><u>New The Ultimate List 10 Adobe Premiere Elements Alternatives for Video Enthusiasts</u></a></li>
<li><a href="https://unlock-android.techidaily.com/top-10-fingerprint-lock-apps-to-lock-your-vivo-y200e-5g-phone-by-drfone-android/"><u>Top 10 Fingerprint Lock Apps to Lock Your Vivo Y200e 5G Phone</u></a></li>
<li><a href="https://animation-videos.techidaily.com/updated-2024-approved-top-15-2d-animation-software-free-that-worth-to-try/"><u>Updated 2024 Approved Top 15 2D Animation Software Free That Worth to Try</u></a></li>
</ul></div>
