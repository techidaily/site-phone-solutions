---
title: Fixed Excel 2021 VBA Runtime Error 9 Subscript Out of Range | Stellar
date: 2024-03-12 21:18:16
updated: 2024-03-14 13:41:10
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2021 VBA Runtime Error 9 Subscript Out of Range
excerpt: This article describes Fixed Excel 2021 VBA Runtime Error 9 Subscript Out of Range
keywords: repair damaged .xltx files,repair damaged .xls files,repair .xlsx files,repair damaged .csv,repair .xlsm files,repair damaged .xltx,repair corrupt excel file,repair corrupt .xltx files
thumbnail: https://www.lifewire.com/thmb/ptfhak0BFgk1HbWMQnlfEezMM8Q=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/kentuckyderby-5c7ed5d646e0fb00011bf3da.jpg
---

## \[Fixed\] Excel VBA Runtime Error 9: Subscript Out of Range

**Summary:** The runtime error 9 in Excel usually occurs when you use different objects in a code or the object you are trying to use is not defined. This post will discuss the reasons behind the Excel VBA error "Subscript out of Range” and the solutions to resolve the issue. It will also mention an Excel repair tool that can help fix the error if it occurs due to corruption in worksheet.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Many users have reported encountering the error “Subscript out of range” (runtime error 9) when using VBA code in Excel. The error often occurs when the object you are referring to in a code is not available, deleted, or not defined earlier. Sometimes, it occurs if you have declared an array in code but forgot to specify the DIM or ReDIM statement to define the length of array.

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

### Method 2: Check the Range of the Array

The VBA error “Subscript out of range” also occurs if you have declared an array in a code but didn’t specify the number of elements. For example – If you have declared an array and forgot to declare the array variable with elements, you will get the error (see below):

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

![Macro Settings In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/macro-settings-in-trust-center.jpg)

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


## Excel Stuck at Opening File 0% - Resolve Performance Issues

**Summary:** If an Excel workbook is stuck at opening file 0%, it usually indicates a problem with the Excel file and its objects. This may happen due to Excel file corruption and a few other reasons. In this post, we have discussed these reasons along with the methods to fix and prevent ‘Excel stuck at opening file 0%’ issue.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

When you open an Excel file (XLS/XLSX) in MS Excel, the program reads and then loads the file data along with all its objects and properties. While opening and loading an Excel file, MS Excel displays an “_Opening percentage_.” You won’t usually notice or see this Excel file opening progress percentage while accessing smaller worksheets.

It’s more noticeable when you open a large Excel file or workbook with multiple objects, formulae, formatting, etc. However, after opening an Excel file with double-click, if it is stuck at _Splash Screen_ with a message “**Opening: FileName.xlsx (0%)”** for a while (say 15-30 minutes) and does not progress, it indicates a problem with the Excel file, MS Excel program, or the system.  

![excel stuck at 0 percent](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/excel-stuck-at-opening-1024x576.png)

## Why Excel is Stuck at Opening File 0%?

If you have encountered this error, it may happen due to one of the following issues,

1. Damaged or corrupt Excel file
2. Incompatible or faulty Excel add-ins
3. Problem with the system’s display driver
4. Damaged MS Office (Excel) application

## Methods to Fix ‘Excel Stuck at Opening File 0%’ Issue

Before fixing and troubleshooting the problem, check and confirm if the Excel file is working and not corrupt. For this, you can try opening it on another PC. Now there could be two scenarios,

### **Scenario 1:  Excel File Does Not Open**

If the Excel file doesn’t open on another PC also, it indicates Excel file corruption. In such cases, look for the backup copy of the file, if you have downloaded it from an email or a website.

However, if there’s no backup, then you need an Excel file repair software, such as **Stellar Repair for Excel** to repair the corrupt file. This software preserves Excel file properties, such as cell formatting, formula bar, freeze panes, gridlines, etc. and helps you restore the damaged or corrupt worksheets to its original state with 100% integrity.

[![free download](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/09/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

To repair Excel file, download and launch _Stellar Repair for Excel_ software on your PC, choose the corrupt Excel (XLS/XLSX) file and click ‘**Repair’**. You can see the preview of your Excel file with all data and then save the repaired file at your desired location on the system as a new Excel file.

![stellar repair excel file](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Stellar-repair-for-excel.jpg)

### **Scenario 2: Excel File Is Accessible on Another PC**

If the Excel file opens successfully on another PC, then follow the troubleshooting methods below to resolve the Excel file stuck opening at 0%.

## Method 1: Open MS Excel in Safe Mode

To check if an incompatible or faulty add-in or setting is causing the error, restart MS Excel in safe mode and then open the worksheet from the MS Excel ‘**File**’ options. The steps are as follows,

1. Press **Windows+R** and type **excel.exe /safe**
2. Hit **Enter** or press ‘**OK**’ to open MS Excel in safe mode

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

![repair ms office](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Repair-MS-Office-installation-1024x577.png)

- When prompted, click on ‘**Repair’** and then follow the instructions to complete the repair process

![quick repair ms office](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Quick-repair-MS-Excel.png)

**To update the MS Excel,**

- Go to **File > Account** and click on **Update options**

![check MS Excel updates](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Check-MS-Excel-updates-1024x539.png)

- Then click ‘**Update’**

![Download MS Excel updates](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Install-MS-Excel-Updates-1024x539.png)

- MS Excel will start downloading the latest updates and then apply it, which might fix this Excel error

![Apply MS Excel updates](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/09/Applying-updates-to-Excel.png)

Still, Excel stuck on processing file at 0%? That means the Excel file you’re trying to open is severely corrupted. Thus, as mentioned earlier, use **Stellar Repair for Excel** software to repair corrupt or damaged Excel (XLS/XLSX) files and restore everything to a new Excel file. With the help of some best-in-class repair algorithms, this software enables you to fix problems within Excel files and recover tables, charts, cell comments, images, formulae, sorts, and filters. It is compatible with MS Excel 2019, 2016, 2013, 2010, 2007, and 2003.

## Conclusion

Hopefully, one of the above-mentioned solutions has helped you overcome the “Excel stuck at Opening file 0%” error and Excel hangs on opening file issues. Also, you are able to access your MS Excel worksheet now. If you face any problems with your Excel workbooks in future, remember to get to the root of the issue first. Also, inculcate the habit of backing up your critical files regularly (if possible) and keep products like **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** in mind to save the day, when nothing else works.


## \[Fixed\] Excel PivotTable Overlap Error | Troubleshooting Guide

In Excel, you need to refresh the pivot table data source after adding new data. However, sometimes, while refreshing the pivot table, you may experience an error “PivotTable Report cannot Overlap.” This issue usually appears when there are multiple pivot tables in a single worksheet. It often occurs when you try to place one pivot table on top of another or if you try to set a common cell range to multiple pivot tables. However, there are many other causes associated with the error.

## **Reasons for a pivot table report cannot overlap another pivot table report issue:**

- Merged cells in a pivot table may cause the overlap issue
- Using the same range of cells for multiple pivot tables
- Hidden columns
- Preserve formatting option is enabled
- Modifying the pivot table using a macro that is corrupted
- Using the workbook.RefreshAll method incorrectly
- Number of pivot items goes beyond the number of cells available
- Excel file is corrupt
- [Corrupted Pivot table](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)
- Some columns are labeled with the same name

## Methods to Fix Excel PivotTable Report Cannot Overlap Error

You can get the pivot table overlapping issue if the field in pivot table crossed the maximum items limit. According to the Microsoft guide, you can specify up to 1,048,576 items to return per field. Check the cell fields in your pivot table. Also, make sure each column’s label is unique. Sometimes, the hidden columns or hidden sheets can also prevent you from modifying the pivot tables. You can check for hidden columns in the Data view.

If the error still persists, then try the below-mentioned methods to fix the error.

### **1\. Move the Pivot Table to a New Worksheet**

The “PivotTable Report cannot Overlap” error can occur if there is an issue with the columns in the pivot table. In this case, you can try moving the pivot table to a new worksheet. Moving the pivot table to a different worksheet automatically resets the column width according to the new sheet and creates space that can help in preventing the overlapping issue. Here are the steps to do so:

### **2\. Disable the Background Refresh Option**

When the background refresh option is enabled, then Excel updates the pivot table in the background after every minor change. It may create issue if you have a large-sized Excel file with multiple pivot tables. You can try disabling the background refresh option. Here’s how:

- The **Connection Properties** dialog box is displayed. Unselect the “**Enable background refresh”** option and select the **“Refresh data when opening the file”**
- Click **OK.  

    ![enable background refresh in connection properties window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivottable/enable-background-refresh-in-connection-properties-window.jpg)

    **

### **3\. Disable Autofit Column Widths**

When the Autofit column widths option is enabled, Excel automatically resizes the pivot table whenever you make changes to it. These automatic adjustments can sometimes add or remove fields which can result in the PivotTable Report cannot Overlap issue. To fix this, you can disable the “Autofit column widths on update” option. To do this, follow these steps:

- Right-click on any field on the pivot table.
- Select **PivotTable Options.  

    ![Select Pivot Table](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivottable/select-pivot-table.jpg)

    **

- In the **PivotTable Options** window, unselect **Autofit column widths on update**.  

    ![select autofit column widths in pivot table options](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/pivottable/select-autofit-column-widths-in-pivottable-options.jpg)

- Click on the **OK.**

### **4\. Check the Workbook.RefreshAll Method**

Several users have reported experiencing the “Excel PivotTable Report cannot Overlap” error when using the Workbook.RefreshAll method. This method is used to refresh data ranges in the pivot report. Sometimes, the error can occur due to missing variable that is representing an object (workbook) in a query. So, make sure you’re using the Workbook.RefreshAll function correctly.

### **5\. Repair your Excel File**

You may also encounter the “A PivotTable Report cannot Overlap” error if the Excel file is corrupted. You can use the inbuilt utility in Excel - Open and Repair to repair the corrupt file. Here’s how:

- In your Excel application, click on the **File** tab and then click **Open**.
- Click **Browse** to select the desired file.
- In the **Open** dialog box, click on the corrupted file.
- Click on the arrow next to the **Open** button and then click **Open and Repair**.
- Click on the **Repair**
- In the displayed message, click **Close**.

If the “Open and Repair” utility fails to fix the issue, then it means there is high level of corruption in the Excel file. To tackle this, you can take the help of a professional Excel file repair tool, such as Stellar Repair for Excel. The tool can easily repair severely corrupted Excel file and recover all the objects of the file, such as pivot tables, macros, charts, etc. with 100% integrity. You can download the free trial version of the tool to check its functionality.

## **Conclusion**

In this article, we have discussed the possible reasons behind the “PivotTable Report cannot overlap” error in Excel. You can follow the methods mentioned above to fix the issue. The error may also occur if the Excel file gets corrupted. In this case, you can try repairing the corrupted Excel file using the Open and Repair utility or consider using [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). The tool makes the process of repairing the Excel file smooth and quick.


## How to Repair Corrupt Pivot Table of MS Excel File?

**Summary:** If you are not able to perform any action on the Pivot Table of MS Excel file, it indicates Excel Pivot Table corruption. In such a case, you must repair the corrupt Pivot Table of MS Excel file by using an Excel repair software or manual troubleshooting steps discussed in this post.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

MS Excel is equipped with several brilliant features and functions which make working with large volumes of data easy. In addition to helping users save data into well-organized cells and tables, the application helps users draw inferences from the data. Pivot Table is one such Excel feature that helps users extract the gist from a large number of rowed data. But often, the Pivot table may get corrupted and lead to unexpected errors or data loss.

**Corrupt Pivot Tables** can stop users from reopening previously saved Excel workbooks, raising the serious issue of data inaccessibility. Resolving such issues is an uphill task unless one gets to the actual root cause of the problem.

_However, with Stellar Repair for Excel software, you can **repair the corrupt Pivot table of MS Excel file** while keeping the Excel file data, formatting, layout, etc. intact._

![Repair Corrupt Pivot Table of MS Excel File](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/07/1-2.jpg)

## **Excel Pivot Tables & Associated Problems**

Pivot Tables in Microsoft Excel are created by applying an operation such as sorting, averaging, or summing to the data in certain tables. The results of the operation are saved as summarized data in other tables. Typically, working on the grouping of saved data, Pivot Tables are used in data processing and are found in data visualization programs, such as spreadsheets or business intelligence software.

Put simply, Pivot Tables in Excel allow you to extract the significance or the gist from a large, detailed data set by allowing you to slice-and-dice data, sort-and-filter data, or arrange it in any way you want.

## Frequently Encountered Problems with Pivot Tables in MS Excel

Take a look at the most frequently encountered **Pivot Table issues**:

- You add **new data into a pivot table** but it doesn’t show up when you refresh
- **Pivot Table contains Blanks** instead of Zeros for fields that have no source data
- **Automatic field names assigned** by the Pivot Table can be inappropriate
- It doesn’t directly **show the percentage of total**
- **Grouping** one pivot table affects another
- Your **number of formatting gets lost**
- Refreshing a pivot table **messes up column widths**
- Field headings make no sense and **add clutter**

While some of the above problems seem minute and can easily be resolved using a few tweaks, bigger issues like unexpected Pivot Table error messages that an Excel throws can be troublesome.

## **Pivot Table Errors & Their Reasons**

Excel users who have built new Pivot Tables in Excel often report the following errors when trying to reopen a previously saved workbook:

_**We found a problem with some content in <filename>. Do you want us to try to recover as much as we can? If you trust the source of this workbook, click Yes.**_

![Pivot Table Corruption error in Excel File](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/07/1-3.png)

Naturally, users are prompted to click on ‘**Yes**’. But when they do, they get another error message saying:

_**Removed Part: /xl/pivotCache/pivotCacheDefinition1.xml part with XML error**_

_**(PivotTable cache) Load error. Line 2, column 0**_

_**Removed Feature: PivotTable report from /xl/pivotTables/pivotTable1.xml part (PivotTable view)**_

Such errors are indicative of the fact that the **data within the Pivot Table still exists**, but the table itself isn’t functioning anymore.

There could be two primary reasons behind such behavior:

- You’ve **created the Pivot Table in an older version** of Excel but are trying to open-refresh-save it through a newer Excel version
- The **Pivot Table itself is corrupted**

## **How to Repair the Pivot Table Quickly?**

To solve the errors associated with Pivot Tables, you need to repair them. But Microsoft doesn’t offer any inbuilt technique or option to repair Pivot Tables. Thus, to fix the issue, you either need some sort of workaround or an [Excel file repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).

## **Methods to Fix Corrupt Pivot Table in MS Excel**

Though there aren’t many options to fix the Pivot Table, you can follow these workarounds to try and repair a corrupt Pivot Table of MS Excel. However, before following these steps, create a backup copy of your Excel file.

### **Method 1: Open MS Excel in Safe Mode**

First, try opening the [Excel file in safe mode](https://support.office.com/en-us/article/open-office-apps-in-safe-mode-on-a-windows-pc-dedf944a-5f4b-4afb-a453-528af4f7ac72) and then check if you can access the Pivot Table. If you can, save all its contents to a new Pivot Table in the latest version of Excel so that this problem doesn’t arise anymore.

### **Method 2: Use Pivot Table Options**

If, however, above method doesn’t work, follow the below-mentioned steps:

- Right-click on the **Pivot Table** and click on **Pivot Table Options**
- On the Display tab, clear the checkbox labeled “**Show Properties in ToolTips**”
- Save the file (.xls, .xlsx) with the new settings intact

### **Method 3: Make Changes to Pivot Table**

If the above method or steps didn’t work,

- Try opening the **Pivot Table Options** window by right-clicking on the Pivot Table within your Excel file
- Select Pivot Table Options from the pop-up menu and make appropriate changes to the options given there
- Then check if the issues go away

### **Method 4: Check and Set Data Source**

If the problem in the Pivot table is related to data refresh,

- Go to **Analyze > Change Data Source**
- Check if the data source is set properly
- Also, try reselecting the data source and check if the refresh option is working properly

If not, resorting to **Stellar Repair for Excel** software might be your only hope.

## **Excel Pivot Table Repair by Using Excel Repair Software**

When corruption strikes an Excel Pivot Table and no manual trick work, **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** is the best solution. This easy-to-use Excel Repair software repairs even the most severely corrupted Excel (XLS/XLSX) files to restore all data, properties, formatting, and preferences. It enables users to extract their saved data into new blank Excel files.

If you have this utility by your side, you don’t need to think twice about any Excel error.

[![Stellar](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/07/image-56.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **What customer says about the Excel Repair Software?**

[**Spiceworks**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

![Spiceworks review of Excel repair](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/07/1-3.jpg)

[**CNET**](https://cloud.cnet.com/Stellar-Phoenix-Excel-Repair/3000-2077_4-10620661.html)

![excel review](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/07/1-4.png)

## **Conclusion**

**Excel Pivot Table corruption** may occur due to any unexpected errors or reasons. This can lead to inaccurate observation in data analysis and also cause data loss if not fixed quickly. However, you can prevent data loss due to problems caused by **Pivot Table corruption** by keeping a backup of all your critical Excel files and fix the Pivot Table corruption by using proper tools, such as Excel file repair software, that can help you get over any Excel corruption and errors quickly.


## Excel File Corruption Warnings and Solutions

**Summary:** Many users reported error messages they receive when they try to save or open an Excel file. In this blog, you will learn about the warning messages that indicate your Excel file is corrupt and possible solutions to repair it. It also outlines the Stellar Repair for Excel to repair corrupt Excel files.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Excel users often report about receiving warning messages suggesting corruption in the workbook. This usually happens while opening an Excel file, ‘.xls’ or ‘.xlsx’ file created by earlier versions, or attempting to create a copy of the workbook.

Excel file corruption may occur due to several reasons including (but not limited to) virus infection, sudden system shutdown during write operation, and leaving excel file open on the shared network.

### **Occurrences of Excel File Corruption Warnings**

_Occurrence 1 – “Excel found unreadable content in <filename>. Do you want to recover the contents of this workbook? If you trust the source of this workbook, click Yes”._

![Image of Excel Found Unreadable Content error message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Excel-found-unreadable-content-error.png)

On clicking ‘Yes’, you will receive the following error:

 _“The file is corrupt and cannot be opened”._

![Image Of Excel File Corruption error Message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/The-file-is-corrupt-and-cannot-be-opened-error.png)

_Occurrence 2 – “Excel cannot open the file <filename>, because the file format or file extension is not valid. Verify that the file has not been corrupted and that the file extension matches the format of the file”._

![Image of Excel File Format Or Extension is Not Valid error message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Excel-file-extension-error.png)

Besides the warning messages outlined above, there are a few other tell-tale signs of Excel file corruption such as:

- Excel crashes or freezes, preventing you from accessing the workbook and information stored in it.
- Unexpected errors occur during the save operation listed as below:
  - _“An unexpected error has occurred. AutoRecover has been disabled for this session of Excel”._
  - _“Errors were detected while saving <filename>”._

### **Solutions to Fix Excel File Corruption Issue**

Follow the below-listed solutions to deal with corruption issues in Excel:

**NOTE:** If you encountered problem opening Excel files after upgrading to latest Windows Operating System (OS) and Office program, try updating your Office as well as Windows OS to latest patches provided on the Microsoft site. Microsoft frequently releases Office and Windows OS patches to help users’ correct known errors. Check if you can open the corrupt workbook after installing the update.

#### **Solution 1 – Use Open and Repair Utility**

Excel comes with a built-in recovery mechanism. It automatically starts ‘File Recovery Mode’ when a user opens a corrupt workbook, and attempts to open and repair the workbook. Sometimes, the recovery mode might not start automatically. In that case, you will need to repair the Excel file manually by using ‘[Open and Repair](https://support.office.com/en-us/article/repairing-a-corrupted-workbook-7abfc44d-e9bf-4896-8899-bd10ef4d61ab)’ utility.

**Steps to use Microsoft’s built-in repair utility are as follows:**

**Step 1:** Select **File** > **Open**.

**Step 2:** Click the folder containing the corrupt workbook, and then click **Browse**.

**Step 3:** In the **Open** window, select the corrupt workbook.

**Step 4:** Next, click the arrow in the **Open** button, and then click **Open and Repair**.

![Image of Open and Repair in-built utility](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Open-and-Repair.png)

**Step 5:** In the window that appears, click **Repair**.

![Image of Excel warning message after using open and repair in-built utility.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Open-and-Repair-Repair-option.png)

If  [‘Open and Repair’ doesn’t work in excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), select **Extract Data** to extract formulas and values from the corrupt workbook.

**_NOTE:_** _If you need a quick solution to salvage your data, use an Excel file repair tool._

Or else, attempt the following solutions to deal with [corruption in Excel file](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).

#### **Solution 2 – Uninstall and Re-install Office Installation**

**_NOTE:_** _Make sure to create a backup of your Excel file before uninstalling and re-installing your Office application._

Download the Office uninstall support tool to remove the application.

You can read: [Simple Ways to Open Corrupt Excel file Without any Backup](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

**To reinstall Microsoft Office, follow these steps:**

_**NOTE:** Before proceeding with Office re-installation process, make sure that you have license keys ready._

**Step 1:** Open the [Microsoft Office](http://www.office.com/) site.

**Step 2:** Select **Sign in**.

**_NOTE:_** _You may skip this step if you’re already signed in._

**Step 3:** After signing in, from the Office sign-in page, click **Install**/**Install Office**

Your Office application will get re-installed. Now open the backed-up Excel file and see if the problem is fixed.

#### **Solution 3 – Move Excel File to a Different Location**

Often moving a corrupt Excel file to a different location can help solve the corruption problem. Here’s how:

**Step 1:** Open the corrupt Excel file by navigating to the following path:

**C:\\Users\\User\_Name\\AppData\\Roaming\\Microsoft\\Excel**

_**NOTE**: Make sure to replace User\_Name with your user name. If you are unable to find the Excel file, you will have to search for the file manually in Program Files (x86)._

![Image of Moving Excel File to a Different Location](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/program-files.png)

**Step 2:**  Open the Excel folder, and move the corrupt file to some other location.

**Step 3:** Delete the files from the Excel folder.

Now try opening the Excel file you have moved and see if the issue is resolved.

#### **Solution 4 – Use Excel File Repair Software**

If none of the above solutions works for you, use **Stellar Repair for Excel**. It is a specialized [Excel file repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) that helps repair corrupt Excel file and recover workbook data in its original state.

Essentially, the software helps rebuild the corrupt file to restore every single object in the file. It can recover objects including user-defined charts, conditional formatting rules, formatting of the charts, properties of worksheet, engineering formulas, etc.

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/06/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

#### **Steps to use Stellar Repair for Excel are as follows:**

**Step 1:** Download, install and launch **Stellar Repair for Excel** software.

**Step 2:** In **Select File** window, click **Browse** to select the file you want to repair.

![Image of Stellar Excel Repair software start screen.
Click on Select File -> Browse](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/user-interface-1024x544.png)

_**NOTE:** If you are unaware of the Excel file location, click ‘Search’ in the Select File window to find the file._

**Step 3:** Once the files are selected, click **Repair** to initiate the repair process.

![Image of Repair Process window after selecting the files to be repaired](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/select-file.png)

**Step 4:** Preview the repaired file and select all or specific files you want to save.

![Image of Preview of Repaired File ](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Preview-1024x545.png)

**Step 5:** Click **Save File** on **Home** menu.

![Image of Save File Button on Home Menu.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Save-file.png)

**Step 6:** In **Save File** window, choose ‘Default Location’ or ‘Select New Folder’ to select the location where you wish to save the file. Click **OK**.

![Image of save File window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/02/Save-file.jpg)

The selected files will be saved at the specified location.

### **Conclusion**

You may experience Excel file corruption warning messages while opening or saving an Excel file. The file may become corrupt due to malware infection, sudden system shutdown, and forgetting to close workbook on a shared network. This post outlined occurrences of Excel file corruption warnings, and also described solutions to fix the issue.

You may try using Microsoft’s built-in ‘Open and Repair’ tool to repair corrupt workbook and recover data from it. If this solution doesn’t work, proceed with uninstalling and re-installing the Office application. Another solution is to move corrupt files to another location. But if the problem still persists, use **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** software to repair single or multiple Excel (.xls or .xlsx) files and restore data.


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

![Click View Option In Documents](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/click-view-option-in-documents.jpg)

- In the **Folder Options** window, click **View**.

![In Folder Options Click On View](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/in-folder-options-click-on-view.jpg)

- Under the **View** section, search for the “**Use Sharing Wizard**” option in the **Advanced Settings**.

![Select Use Sharing Wizard](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/select-use-sharing-wizard-1.jpg)

- If the **Use Sharing Wizard** option is unselected, select it and then click **OK**.

### **Method 3: Move the Excel File to a Trusted Location**

You can encounter the sharing violation error if the file you are trying to save is not in the trusted location. You can try moving the file to a trusted location by following these steps:

- In Excel, go to **File** and then click **Options.**
- Click **Trust Center** and then click **Trust Center Settings**.

![Click Trust Center Settings In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/click-trust-center-settings-in-trust-center.jpg)

- In the **Trust Center** window, click **Trusted Locations** and then click **Add new location**.

![Click On Add New Location Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/click-on-add-new-location-option.jpg)

- In the **Add new location** window, select **Browse** to locate and choose the folder, and then click **OK**.

### **Method 4: Open Excel in Safe Mode**

Incompatible add-ins can create issues in the Excel file. To check if the sharing violation issue has occurred due to add-ins, open Excel in safe mode. To do so, follow these steps:

- Open the **Run** window using **Windows + R**.

![Type Safe Mode Command In Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/06/type-safe-mode-command-in-excel.jpg)

- Type **excel /safe** and click on **OK**.
- Open the affected file and then try to save the changes.
- If you are able to save the changes without any error, then it indicates add-ins are causing the issue. To fix this, you can [remove the recently downloaded add-ins](https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) (if any).

### Method 5: Repair the Excel File

Corruption in Excel file can also create issue while saving the changes. In such a case, you can repair the corrupted Excel file using the inbuilt utility in Excel, named Open and Repair. Follow these steps to use this utility:

- In Excel, navigate to **File > Open > Browse**.
- In the **Open** dialog box, click on the affected Excel file.
- Click the arrow next to the **Open** button and select **Open and Repair** from the dropdown.
- Click on the **Repair** option to recover as much data from the file as possible.

If the above utility fails to fix the corrupt Excel file, then you can use a more powerful [Excel repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), such as Stellar Repair for Excel. This Excel repair tool can repair even severely corrupted or damaged Excel files (xls, .xlsx, .xltm, .xltx, and .xlsm). This tool can recover all the data from the corrupted Excel file, including images, chart sheets, formulas, etc., without changing the original format. It can help in fixing common corruption-related errors in Excel. You can download the software’s demo version to scan the corrupt file.

## **To Conclude**

Above, we have discussed some effective methods to fix the sharing violation error in Excel. This error may also occur if you try to save the Excel file in an incompatible format. So, check the format and try saving the file in a compatible format. If the error occurs due to Excel file corruption, you can [repair corrupt Excel file](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) using the Open and Repair tool. If nothing works, then download a third-party Excel repair tool, such as Stellar Repair for Excel. It is an advanced tool that can fix severely corrupted Excel files. You can install this repair tool on any Windows system.


## How Do I Repair and Restore Excel File?

When an Excel file turns corrupt, the file might become inaccessible or you might receive errors. You may encounter errors, such as ‘the file is corrupt and cannot be opened,’ ‘Excel found unreadable content in "filename>",’ ‘Excel cannot open "filename" because the file format or extension is not valid,’ etc.

## Common Reasons for Excel File Corruption

There are several reasons that can turn the file corrupt. The most common reason is a damaged hard drive. Other factors that can cause corruption in an Excel file are as follows:

- System crash or abrupt shutdown of the system while the file is still open
- Viruses infecting the file with malicious code
- Bug in the operating system
- Bad sectors on the drive where the file is stored
- Large spreadsheets with formulas and other components

Whatever be the reason, if your business is dependent on an Excel file, corruption in the file could hamper your business continuity. Also, you may lose crucial data. In such a situation, you could try to repair the file.

## Before We Begin

It is important to identify the root cause behind Excel file corruption. If the problem has occurred due to a faulty hard disk drive, contact your hardware vendor to get it fixed. Also, move the file to another local drive and check if it opens. If nothing works, proceed with the methods discussed below to repair and restore the file.

## Methods to Repair and Restore Excel File

Try the following methods to fix corruption in an Excel file and restore it.

### Method 1 – Use the Built-in ‘Open and Repair’ Tool

You can use the Excel built-in Open and Repair utility to repair the corrupt file. Follow these steps:

- Open your Excel application and click on Blank workbook.

![blank excel workbook](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/blank-excel-workbook-img-1.png)

- On the blank workbook screen, click on the File tab.

![file menu](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/file-menu-img-2.png)

- Click Open > Computer > Browse.

![select the open option](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/select-the-open-option-img-3.png)

- Select the file you want to repair and then click on Open and Repair from the Open dropdown box.

![open and repair excel file](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/open-and-repair-excel-file-img-4.png)

- Click Repair to fix corruption in the Excel file and recover maximum data.

![repair or extract excel data](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/repair-or-extract-data-img-5.png)

- If you get the following error message, click Yes to open the file.

![excel file format does not match error](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/excel-file-format-does-not-match-error-img-6.png)

- If clicking Yes opens the file with garbage entries (see the image below), perform Step 1 – 5 and click Extract Data. This will only help you recover data without formulas and values.

![excel file with garbage entries](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/excel-file-with-garbage-entries-img-7.png)

Note: You may also try to recover the data from a corrupted workbook by using the [methods suggested by Microsoft](https://support.microsoft.com/en-us/office/repairing-a-corrupted-workbook-7abfc44d-e9bf-4896-8899-bd10ef4d61ab).  

A better way to repair and restore an Excel file with complete data is to use a specialized [Excel file repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).

### Method 2 – Use Excel File Repair Tool

Stellar Repair for Excel is a powerful tool designed to help users fix corrupted .xls or .xlsx files without any technical assistance. Also, the tool recovers all the components from a corrupted workbook, including tables, pivot tables, cell values, formulas, charts, images, etc. You can preview the repaired file and its contents by downloading the free demo version from the link below. It is a useful feature that allows the user to validate the data before saving it.

[

![Free Download For Windows](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)

](<https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/> "Free Download For Windows")

Here’s the step-by-step instructions to repair a corrupt Excel file using the software:

- Run the software. The software main interface opens with an instruction to add some add-ins if you’ve engineering formulas in the file you want to repair.

![software main screen](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/software-main-screen-img-8.png)

- Click OK to proceed.

- Select the file you wish to repair by using the Browse option.

Note: If you’re not aware of the file location, choose the ‘Search’ option to locate the file.

![repair excel file](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/repair-excel-file-img-9.png)

- A screen showing progress of the Excel file repair process is displayed.

![progress of the repair process](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/progress-of-the-repair-process-10.png)

- Preview of the repaired Excel file and its recoverable data is displayed.

![preview repaired excel file](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/preview-repaired-excel-file-img-11.png)

- After verifying the data, click on the Save File button on the File menu to save the repaired file.

![save repaired excel file](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/save-repaired-excel-file-img-12.png)

- Select the location where you wish to save the repaired file on the Save File window and then click OK.

![save at default location](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/Recover-Excel-Files/save-at-default-location.png)

A confirmation message will pop-up after completion of the repair process. You can now try to open the file in your Excel program.

## End Note

Even if you’re taking preventive measures, you might still experience corruption in an Excel file. So, it’s crucial to take regular backups of your workbooks. For this, ensure that the 'Always create backup' option is enabled in Excel. You can find it in General Options by clicking on the Tools button in the Save As dialog box. Enabling it will ensure that the Excel backup file is updated with the changes made in a spreadsheet.

Additionally, ensure that the Excel ‘AutoRecover’ feature is set to save a version of your Excel file after every 10 minutes. You can increase or shorten the interval as per your requirement.




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


