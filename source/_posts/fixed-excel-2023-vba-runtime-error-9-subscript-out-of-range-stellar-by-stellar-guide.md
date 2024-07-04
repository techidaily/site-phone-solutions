---
title: Fixed Excel 2023 VBA Runtime Error 9 Subscript Out of Range | Stellar
date: 2024-03-11 19:55:50
updated: 2024-03-14 10:13:29
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2023 VBA Runtime Error 9 Subscript Out of Range
excerpt: This article describes Fixed Excel 2023 VBA Runtime Error 9 Subscript Out of Range
keywords: repair damaged .csv,repair damaged .xls files,repair excel 2010,repair .xlsx,repair corrupt .csv files,repair .xltm
thumbnail: https://www.lifewire.com/thmb/0FZf3k28kLauMvGO0aGhDI7aaYY=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/sb10069770n-003-56a104403df78cafdaa7dd48-ba41d70c51114343aaa38409d9cdfc3f.jpg
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


## Simple ways to Open Corrupt Excel file Without any Backup

**Summary:** The blog describes simple ways to open corrupt Excel file without any backup. It explains some manual workarounds that you can try to open the file. Also, it mentions about an Excel file repair tool that can quickly fix the corrupt file and recover data from it.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Do you have an Excel file that does not open because of corruption issue? And every time you try to open it, an error message ‘the file is corrupt and cannot be opened’ pops-up?

![Excel file is corrupt and cannot be opened message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/05/Excel-file-corruption-message-300x139.png)

Excel File Corruption Message

Also, you don’t have a healthy backup of the Excel file to restore the data? If so, you can try repairing the corrupt file by using a few simple yet effective manual workarounds mentioned below.

## **How to Open a Corrupt Excel File without Backup?**

Following are some manual methods that can help you open a corrupt Excel file:

### **Method 1: Repair Corrupt Excel File**

When attempting to open a corrupt file, Excel automatically starts ‘File Recovery’ mode to repair the file. But, if the recovery mode doesn’t start, try Microsoft Excel’s built-in ‘Open and Repair’ feature to manually repair the file.

To use this feature, perform the following steps:

**Step 1:** Open a **Blank workbook** in Excel, and then click **File > Open**.

**Step 2:** In the **Open** window, browse and select the corrupt file.

**Step 3:** Click the arrow that is beside the **Open** tab, and select **Open and Repair**.

![Open a blank workbook in Excel, navigate to File > Open, choose the corrupt file, and, in the Open window, click the arrow beside the Open tab, selecting Open and Repair for file recovery.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/05/Open-and-repair.png)

Open and Repair Option

**Step 4:** Implement one of the following:

- Click the **Repair** button. (This is to recover as much data as possible.)
- Click the **Extract Data** button. (This is to recover values and formulas from the Excel file if the repair process fails to recover the entire data.)

![Initiate file recovery by selecting the Repair tab, and if necessary, retrieve values and formulas using the Extract Data tab in Excel.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/04/repair-excel-file-1-768x158.jpg)

Excel Built-in Repair Options

If using [Open and Repair does not work](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), move to the next method.

### **Method 2: Disable the Protected View Feature**

Some Excel users have reported that turning off the ‘protected view’ feature in Excel helped them open the corrupt file. You can also try to disable this feature and open the file. To do so, follow these steps:

**Step 1:** Open a blank Excel file, click on **File** > **Options**.

**Step 2:** In the **Excel Options** window, select **Trust Center**, and then click **Trust Center Settings**.

![In the Trust Center tab, click on Trust Center Settings...](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/05/Excel-trust-center-settings.png)

Excel Trust Center Settings

**Step 3:** Click **OK.**

Now check if you can open the corrupt file. If not, try implementing the next method.

### **Method 3: Look For Automatically Recovered Excel File**

If you have Excel’s AutoRecover feature enabled, you’ll have access to a copy of the **Excel file corrupted** or lost due to application crash, power outage, or accidental deletion.

**The ‘AutoRecover’** feature saves Excel worksheets at a temporary location after a certain time interval. It saves the worksheets automatically and is turned on by default to reduce the chance of data loss.

Check if you can **[recover corrupted Excel file](https://support.microsoft.com/en-us/office/repair-a-corrupted-workbook-153a45f4-6cab-44b1-93ca-801ddcd4ea53)** by following these steps:

**Step 1:** In Excel, open a **Blank workbook**.

**Step 2:** Go to **File** and click **Options**.

![Open a new Excel workbook, then access additional settings by navigating to File and selecting Options.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/08/Select-options-in-Excel-2013.jpg)

Figure 5 – Excel Options

**Step 3:** In the **Excel Options** dialog box, click **Save**, and then copy the ‘AutoRecover file location’.

![Copy the 'AutoRecover file location' for configuration or backup purposes.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/05/Autorecover-excel-file.png)

Excel Options Window

**Step 4:** Open File Explorer window and paste the copied AutoRecover file location, and press **Enter**.

**Step 5:** A list of saved Excel files will be displayed. Choose the file you want to recover.


_**TIP:** Use Excel’s AutoBackup feature to reduce chances of data loss, by saving a previous version of your spreadsheet automatically._

## **Use an Excel File Repair Software**

If the above manual methods fail, repair the **corrupt Excel file** by using a third-party software, such as Stellar Repair for Excel**.** The software helps repair Excel (XLS and XLSX) files easily and effectively.

[![Free Download for windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/01/Free-download-for-windows-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

**Read this: [How to repair corrupt Excel file using Stellar Repair for Excel?](https://www.stellarinfo.com/support/kb/index.php/article/repair-corrupt-excel-file)**

Some key features of Excel Repair software are as follows:

- Fixes all errors in the MS Excel file.
- Repairs multiple damaged Excel files in a go.
- Recovers chart, chart sheet, table, cell comment, image, formula, and sort & filter.
- Preserves properties and cell formatting of Excel worksheets.
- Previews recoverable Excel file data before saving.
- Recovers all data components from the corrupt files and saves them in a new blank Excel file.
- Compatible with Excel 2019, 2016, 2013, 2010, 2007, and lower versions.

## Conclusion

You can try the workarounds discussed in the blog to open a corrupt Excel file without a backup. Disabling the protected view feature can help you open the file. If the issue persists then try repairing the corrupted Excel file using the Open and Repair utility. Although, it may not be able to fix a severely corrupted workbook. In such a case you can use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It is an advanced tool that can help you repair a corrupted Excel file with 100% integrity.


## How Can I Recover Corrupted Excel File 2016?

## Error Messages Indicating Corruption in Excel File

- When an Excel 2016 file turns corrupt, you’ll receive an error message that reads: **“[The file is corrupt and cannot be opened](https://www.stellarinfo.com/blog/file-is-corrupted-and-cannot-be-opened-excel-2010/).”**

![](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/the-file-is-corrupt-and-cannot-be-opened-error-img1.png)

- But sometimes, you encounter the **“Excel cannot open this file”** error message due to corruption in the file.

![Excel-cannot-open-this-file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/Excel-cannot-open-this-file-img2.png)

## Why does Excel File turn Corrupt?

Following are some common reasons that can turn an Excel file corrupt:

- Large size of the Excel file
- The file is virus infected
- Hard drive on which Excel file is stored has developed bad sectors
- Abrupt system shutdown while working on a worksheet

## Workarounds to Recover Data from Corrupt Excel

The workarounds to recover corrupted Excel file 2016 data will vary depending on whether you can open the file or not.

How to Recover Corrupted Excel File 2016 Data When You Can Open the File?

If the corrupt Excel file is open, try any of the following workarounds to retrieve the data:

### **Workaround 1 – Use the Recover Unsaved Workbooks Option**

If your Excel file gets corrupt while you are working on it and you haven’t saved the changes, you can try retrieving the file’s data by following these steps:

- Open your Excel 2016 application and click on the **Open Other Workbooks** option.

![open-other-workbooks](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/open-other-workbooks-img3.png)

- Click the **Recover Unsaved Workbooks** button at the bottom of the ‘Recent Workbooks’ section.

![recover-unsaved-workbook](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/recover-unsaved-workbook-img4.png)

- A window with list of unsaved Excel files will open. Click the corrupt file you want to open.

This will reopen your last saved version of the Excel workbook. If this method doesn’t work, proceed with the next workaround.

### **Workaround 2 – Revert to Last Saved Version of your Excel File**

If your Excel file gets corrupt in the middle of making any changes, you can recover the file’s data if the changes haven’t been saved. For this, you need to revert to the last saved version of your Excel file. Doing so will discard any changes that may have caused the file to turn corrupt. Here’s how to do it:

- In your Excel 2016 file, click **File** from the main menu.
- Click **Open**. From the list of workbooks under Recent workbooks, double-click the corrupt workbook that is already open in Excel.
- Click **Yes** when prompted to reopen the workbook.

Excel will revert the corrupt file to its last saved version. If it fails, skip to the next workaround.

### **Workaround 3 – Save the Corrupted Excel File in Symbolic Link (SYLK) Format**

Saving an Excel file in SYLK format might help you filter out corrupted elements from the file. Here are the steps to do so:

- From your Excel **File** menu, choose **Save As**.
- In ‘Save As’ window that pops-up, from the **Save as type** dropdown list, choose the **SYLK (Symbolic Link)** option, and then click **Save**.

![symbolic link format](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/save-as-symbolic-link-format-img5.png)

**_Note:_** _Only the active sheet will be saved in workbook on choosing the SYLK format._

- Click **OK** when prompted that “The selected file type does not support workbooks that contain multiple sheets”. This will only save the active sheet.

![Workbooks contain multiple sheets warning msg](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/Workbooks-contain-multiple-sheets-warning-msg-img6.png)

- Click **Yes** when the warning message appears - “Some features in your workbook might be lost if you save it as SYLK (Symbolic Link)”.

![](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/Excel-message-img7.png)  

- Click **File** > **Open**.
- **Browse** the corrupt workbook saved with SYLK format (.slk) and open it.
- After opening the file, select **File** > **Save As**.
- In ‘Save as type’ dialog box, select Excel workbook.
- Rename the workbook and hit the **Save** button.

After performing these steps, a copy of your original workbook will be saved at the specified location.

How to Recover Corrupted Excel File 2016 Data When You Cannot Open the File?

If you can’t access the Excel file, apply one of these workarounds to salvage the file’s data.

### **Workaround 1 – Open and Repair the Excel File**

Excel automatically initiates ‘File Recovery’ mode on opening a corrupt file. After starting the auto-recovery mode, it attempts to reopen and repair the corrupt Excel file at the same time. If the auto-recovery mode does not start automatically, you can try to fix corrupted Excel file 2016 manually by using ‘Open and Repair’. Follow these steps:

- Open a blank file, click the **File** tab and select **Open**.
- **Browse** the location where the corrupt 2016 Excel file is stored.
- When an ‘Open’ dialog box appears, select the file you want to repair.
- Once the file is selected, click the arrow next to the **Open** button, and then click the **Open and Repair** button.
- Do any of these actions:
- Click **Repair** to fix corrupted file and recover data from it.
- Click **Extract Data** if you cannot repair the file or only need to extract values and formulas.

![repair excel file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/repair-excel-file-img8.jpg)

If performing these actions doesn’t help you retrieve the data, proceed with the next workaround.

### **Workaround 2 – Disable the Protected View Settings**

Follow these steps to disable the protected view settings in an Excel file:

- Open a blank 2016 workbook.

![blank excel file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/blank-excel-file-img9.png)

- Click the **File** tab and then select **Options**.

![Excel file options](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/Excel-file-options-img10.png)

- When an **Excel Options** window opens, click **Trust Center** > **Trust Center Settings.**

![open excel trust center settings](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/open-excel-trust-center-settings-img11.png)

- In the window that pops-up, choose **Protected View** from the left side navigation. Under ‘Protected View’, uncheck all the checkboxes, and then hit **OK**.

![disable-protected-view-settings](https://www.stellarinfo.com/blog/wp-content/uploads/2021/04/disable-protected-view-settings.png)

Now, try opening your corrupt Excel 2016 file. If it won’t open, try the next workaround.

### **Workaround 3 – Link to the Corrupt Excel File using External References**

If you only need to extract Excel file data without formulas or calculated values, use external references to link to your corrupt Excel 2016 file. Here’s how you can do it:

- From your Excel file, click **File** > **Open**.
- From the window that opens, click **Computer** and then click **Browse** and copy the name of your corrupt Excel 2016 file. Click the **Cancel** button.

![browse corrupted excel file](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/browse-corrupted-excel-file-img13.png)

- Go back to your Excel file, click **File** > **New** > **Blank workbook**.

![new excel workbook](https://www.stellarinfo.com/public/image/catalog//article/file-repair/recover-corrupted-excel-file/new-excel-workbook-img14.png)

- In the new Excel workbook, type “=CorruptExcelFile Name!A1” in cell A1 to reference cell A1 of the corrupted file. Replace the ‘CorruptExcelFile Name’ with the name of the corrupt file that you have copied above. Hit **ENTER**.
- If ‘Update Values’ dialog box appears, select the corrupt 2016 Excel file, and then click **OK**.
- If ‘Select Sheet’ dialog box pops-up, select a corrupt sheet, and press the **OK** button.
- Select and drag cell A1 till the columns required to store the data of your corrupted Excel file.
- Next, copy **row A** and drag it down to the rows needed to save the file’s data.
- Select and copy the file’s data.
- From the **Edit** menu, choose the **Paste Special** option and then select **Values**. Click **OK** to paste values and remove the reference links to the corrupt file.

Check the new Excel file for recoverable data. If this didn’t work, consider using an [Excel file repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to retrieve data.

### **Alternative Solution to Recover Excel File Data**

Applying the above workarounds may take considerable time to recover corrupted Excel file 2016. Also, they may fail to extract data from a severely corrupted file. Using Stellar Repair for Excel software can help you overcome these limitations. The software helps repair severely corrupted XLS/XLSX file and retrieve all the file data in a few simple steps.

[![free download](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

Key benefits of using Stellar Repair for Excel are as follows:

- Recovers tables, pivot tables, images, charts, chartsheets, hidden sheets, etc.
- Maintains original spreadsheet properties and cell formatting
- Batch repair multiple Excel XLS/XLSX files in a single go
- Supports MS Excel 2019, 2016, 2013, and previous versions

Check out this video to know how the Excel file repair tool from Stellar® works:

<iframe width="560" height="315" src="https://www.youtube.com/embed/VAeGzHnETu0" title="YouTube video player" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen=""></iframe>

## Conclusion

Errors such as ‘the file is corrupt and cannot be opened’, ‘Excel cannot open this file’, etc. indicate corruption in an Excel file. Large-sized workbook, virus infection, bad sectors on hard disk drive, etc. are some reasons that may result in Excel file corruption. The workarounds discussed in this article can help you recover corrupted Excel file 2016 data. However, manual methods can be time-consuming and might fail to extract data from severely corrupted workbook. A better alternative is to use Stellar Repair for Excel software that is purpose-built to repair and recover data from damaged or corrupted Excel file.



## Ways to Fix Personal Macro Workbook not Opening Issue

Many users have reported encountering issues while accessing personal macro workbook, such as personal macro workbook not opening, personal macro workbook not loading automatically, Excel personal macro workbook keeps getting disabled, etc.

Such issues may arise due to a problem with the directory where the personal workbook is stored. However, there are various other reasons that may lead to such issues. Below, we’ll discuss the reasons behind the personal macro workbook not opening issue and the solutions to troubleshoot and fix the issue. But before proceeding, let’s understand why personal macro workbook is used.

## Why Personal Macro Workbook is used?

You can access macros in a specific Excel workbook. However, when you need to use the same macro in other Excel worksheets, then you can create a personal macro workbook. A personal macro workbook (Personal.xlsb) is a hidden workbook that is used to store all macros. It makes your macros available every time you open Excel.

## Causes of Personal Macro Workbook not Opening Issue

You may encounter personal macro workbook is not opening issue when attempting to record macros. Some possible causes behind such an issue are:

- Personal macro workbook is stored at an untrusted location
- Location of xlsb is changed
- Personal macro workbook is hidden
- Personal macro workbook becomes corrupted
- Disabled items in add-ins
- Workbook is Read-only

## Methods to Fix the “Personal Macro Workbook not Opening” Issue

 Follow the given methods to fix the personal macro workbook is not opening issue:

###  **Method 1: Check the Path of Personal.xlsb**

The personal macro workbook (Personal.xlsb) file is stored in XLStart folder. It opens automatically when you open your Excel application. However, sometimes it fails to load automatically. It usually occurs when you try to open the file from an incorrect path. You can check the path of Personal.xlsb by following these steps:

- Open the workbook.
- Click on the **Developer** tab.

![developer tab ](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/fix-personal-macro-workbook/developer-tab.png)

- Press **Alt + F11** to open Visual Basic Editor.
- Go to **View > Immediate Window.**

**![immediate window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/fix-personal-macro-workbook/immediate-window.png)**

- In **Immediate Window**, type the following code to know the location of the workbook:

?thisworkbook.path.

- Then, hit Enter.

![personal macro workbook window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/fix-personal-macro-workbook/personalmacro-workbook-window.png)

- You will see the path of the personal macro workbook.
- Copy the path and paste it into **Quick Access** field in **File Explorer**.

![File Explorer window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/fix-personal-macro-workbook/file-explorer-window.png)

### Method 2: Unhide Personal Macro Workbook

If personal macro workbook is hidden, you may unable to see and open the Personal.xlsb file. To unhide the personal Macro workbook, follow the below steps:

- In Microsoft Excel, go to **View** and then click **Unhide**

![unhide personal workbook window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/fix-personal-macro-workbook/unhide-personalworkbook.png)

- The **Unhide** dialog box is displayed. Click PERSONAL and then **OK**.

![unhide window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/fix-personal-macro-workbook/unhide-window.png)

### **Method 3: Enable the Macro Add-ins**

You may unable to open the previously recorded macros in your personal macro workbook if the macros are disabled. To check and enable the items, follow these steps:

- Go to **File > Options.**
- In **Excel Options**, click on the **Add-ins**
- Select **Disabled Items** from the **Manage** section and click on **Go**.

![Access Option to Disable Items](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/fix-personal-macro-workbook/access-option-to-disableitems.png)

- The **Disabled Items** dialog box appears. Click on the disabled item and then click **Enable**.

### **Method 4: Change the Trusted Location**

You may encounter the “personal macro workbook not opening” issue if the Personal.xlsb file is stored at an untrusted location. You can check and modify the path of **XLSTART** folder using the Trust Center window. Here are the steps:

- Open MS Excel. Go to **File > Options**.
- Click **Trust Center > Trust Center Settings**.
- In the **Trust Center Settings** dialog box, click on **Trusted Locations**.

![Trust Center Window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/fix-personal-macro-workbook/trust-center-window.png)

- Verify the path of the **XLSTART** If it is untrusted or there is any issue, then click **Modify** and then click **OK**.

### **Method 5: Repair your Excel File**

You may fail to open personal macro workbook if it is corrupted. To repair the corrupt workbook, you can use the built-in Open and Repair utility in MS Excel. To use this tool, follow these steps:

- Open your Excel application.
- Click **File > Open**.

![Go to Options window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/fix-personal-macro-workbook/go-to-options.png)

- Browse to the location where the corrupted file is stored.
- In the **Open** dialog box, select the corrupted workbook.
- From the **Open** dropdown list, click **Open and Repair**.

The dialog box appears with the Repair and Extract buttons. Click **Repair** to retrieve all possible data or the **Extract** option to recover the data without formulas and values.

If the [Open and Repair utility fails](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to repair the corrupted Excel workbook, then you can use a professional Excel repair tool, such as Stellar Repair for Excel. It can easily repair severely corrupted Excel (XLSX and XLS) files and recover all the components. You can download the free trial version of the tool to preview the recoverable data.

## **Closure**

This article discussed the ways to fix the personal macro workbook not opening issue. In case you are unable to open the personal macro workbook because of corruption in the workbook, you can use the Open and Repair utility in MS Excel. If it fails, then you can use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to fix corruption in the Excel file and recover all its data with complete integrity.




## How to Fix the Unable to Record Macro Error in Excel?

**Summary:** You may encounter the “Unable to record macro” error in MS Excel when using Personal Macro Workbooks. In this post, we’ll discuss the possible causes behind this error and the ways to fix it. We’ll also mention a professional Excel repair tool that can help fix the error if it occurs due to corrupted workbook.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

A personal macro workbook (Personal.xlsb file) is a hidden workbook in Excel that stores all macros in a single workbook. This allows you to automate processes while working in Excel. Sometimes, when recording macro codes in the personal macro workbook, you may get the message “**Personal Macro Workbook in a startup folder must stay open for recording**”. When you click on the **OK** button, it will show the “unable to record” error. This prevents you from recording the macros. Below, we’ll see the causes behind this error and discuss how to resolve this error.

## **Causes of Unable to Record Macro Error**

You may be unable to record macros in Excel due to several reasons. Let’s take a look at the possible causes that can lead to this issue.

- The location of personal.xlsb file is changed.
- Personal.xlsb file is corrupted.
- Macros are disabled.

## **Methods to Fix the “Unable to Record Macro” Error in Excel**

Here are some possible solutions that can help you resolve the unable to record macro error in Excel.

### Method 1: Check the Path of XLStart Folder

You may be unable to record macros if the path of XLStart folder is incorrect. It is a folder where the Personal.xlsb file is stored by default. Follow these steps to find out the path of this folder:

- Open MS Excel. Go to **File > Options**.
- Click **Trust Center > Trust Center Settings**.

![Excel Options Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/excel-options-window.png)

- In the **Trust Center Settings** window, click on **Trusted Locations**.

![Path Of XLStart Folder In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/path-of-xlstart-folder-in-trust-center.png)

- Verify the path of the **XLSTART** folder and modify it if there is an issue.
- Once you are done, click on **OK**.

### Method 2: Change Macro Security

The “Unable to record macro” error can occur if macros are disabled in the Macro Security settings. You can try changing the macro settings using the below steps:

- In MS Excel, go to **File > Options > Trust Center**.

![Excel Options To Locate Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/excel-options-to-locate-trust-center.png)

- Under **Trust Center,** click on **Trust Center Settings**.

![Change Macro Settings In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/change-macro-settings-in-trust-center.png)

- Select **“Enable all macros”** and then click **OK.**

### Method 3: Check Add-ins for Disabled Items

If there are any items in add-ins that are disabled, they may prevent Excel from functioning properly. You can check and enable the items in MS Excel using the below steps:

- Click **File > Options.**

![Go To Options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/go-to-options-1.png)

- In **Excel Options**, click on the **Add-ins** option.
- Select **Disabled Items** from the **Manage** section and click on **Go**.

![Add-ins In Excel Options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/add-ins-in-excel-options.png)

- The **Disabled Items** window is displayed.
- Click on the disabled item and then click **Enable**.
- Restart Excel for the changes to take place.

### Method 4: Repair your Excel File

You may fail to record macros if there is corruption in the workbook. In such a case, you can use the “Open and Repair” utility in MS Excel to repair the corrupt workbook. To use this tool, follow these steps:

- Open your Excel application.
- Click **File > Open**.
- Browse to the location where the corrupted file is stored.
- In the **Open** dialog box, choose the corrupted workbook.

![Open Dialog Box](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/open-dialog-box.png)

- From the **Open** dropdown list, click **Open and Repair**.

![Open And Repair Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/02/open-and-repair-window.png)

Excel will prompt you to repair the file or extract data from it. Click **Repair** to retrieve maximum data. If the Repair option fails, then click on the Extract Data option to recover the data without formulas and values.

If the Microsoft utility “Open and Repair” fails to repair the corrupted Excel workbook, then try a professional Excel repair tool such as Stellar Repair for Excel. It is an advanced tool that can easily repair severely corrupted Excel (XLSX and XLS) files. It can recover all the file items, including chart sheets, cell comments, tables, macros, formulas, etc. without impacting the properties and cell format of the Excel file.

## **Closure**

You may receive the “unable to record” error in Excel while creating or storing macros in Personal Macro Workbooks. There are several reasons that can lead to this error. You can try the methods covered in this post to resolve the error. If the error appears due to corruption in workbook, then try to repair it using the Open and Repair utility. Alternatively, you can use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) – a professional tool that can help you repair an Excel file with problematic macros. Also, it allows recovery of all the file components with complete integrity. The tool is compatible with Excel 2021, 2019, 2016, and older versions.


## \[Fixed\] "Microsoft Excel Cannot Access the File" Error

**Summary:** The “Microsoft Excel cannot access the file” error usually occurs when there is an issue with the Excel file you are trying to save. This post summarizes the causes behind the error and mentions some effective solutions to fix it. If you suspect the problem is encountered due to corruption in the Excel file, you can use the professional Excel repair tool mentioned in the post to repair the file.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

You may experience the “Microsoft Excel cannot access the file” error when saving the Excel file. This happens when the Excel application fails to read the file you are attempting to save. The error message indicates that there is an issue with the file name or its path. Sometimes, the error occurs if the file you are trying to access is already in use by another application. Some other reasons for the “Excel cannot access the file” error are:

- Faulty or incompatible Excel add-ins.
- The file is in Protected View.
- The Excel file is damaged or corrupted.
- You do not have the required permissions to access the file.
- The Excel file is not in a compatible format.

## **Methods to Fix “Microsoft Excel Cannot Access the File” Error**

Sometimes, changing the file location can fix the “Microsoft Excel cannot access the file” error. You can try changing the file location, if the location is incorrect. If moving the file to a different location didn’t work, then try the below troubleshooting methods.

### **Method 1: Check the File Name and Path**

You can get the “Microsoft Excel cannot access file” error if there is an issue with the file path – either the path does not exist or it is too lengthy, thus creating conflicts. Make sure the file path is correct. If the file name is too long, you can rename the file with a short name and also move the file to the parent folder instead of a subfolder. After that, remove the file from the **Recent** list that is created by Excel based on your recent activity. Follow the below steps:

- Open the Excel application.
- In the **Recent list**, right-click on the affected Excel file.
- Now, select **Remove from list**.

![Selecting the "remove from list" option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-remove-from-list.jpg)

- Close the Excel application.

Now, reopen the problematic file and check if the error exists. If yes, then follow the next solution.

### **Method 2: Try Clearing the Microsoft Office Cache**

Sometimes, clearing the Microsoft Office cache can help eliminate the “Excel cannot access the file” error. To clear the Microsoft Office cache, follow the given steps:

- First, close all the Office applications.
- Press **Windows+R** to open the **Run** window.
- Type %localappdata%\\Microsoft\\Office\\16.0\\OfficeFileCache and press the **Enter** key. You can change ‘16.0’ with your Office version.

![Clearing Microsoft Cache from officefilecache Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/clear-microsoft-cache-from-officefilecache-window-1024x311.jpg)

- In the **OfficeFileCache** window, clear all the temporary files.

### **Method 3: Check and Update Microsoft Excel**

You can try updating your Microsoft Excel application. The latest updates include bug fixes, security patches, and other improvements. Updating the application can help fix several issues that might be causing the error. Here are the steps to update Microsoft Excel:

- Open your Excel application.
- Go to **File** and then select **Account.**
- Under **Product information**, click **Update Options** and then click **Update Now**.

### **Method 4: Disable Protected View**

You may get the “Microsoft Excel cannot access the file” error if the [Protected View](https://support.microsoft.com/en-au/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653) option is enabled. You can try disabling the Protected View settings in Excel. This allows you to open the file without any restrictions. However, disabling the protected view can put your system at high risk. To disable the Protected View in Microsoft Excel, follow the below steps:

- In Excel, go to **File** and then click **Options**.
- In the **Excel Options** window, click **Trust Center** and then click **Trust Center Settings.**

![Go To Trust Center and Click on Trust Center Settings ](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/go-to-trust-center-and-click-trust-center-settings-1.jpg)

- Click **Protected View** from the left pane in the **Trust Center Settings** window.
- Unselect the options under **Protected View**. Click **OK.**

### **Method 5: Check and Disable Add-ins**

The “Excel cannot access the file” error can also occur due to faulty add-ins in Excel. To check if the error has occurred due to some faulty add-ins, open the application in **safe mode** (press Windows + R and typeexcel /safe in the Run window**)**. If you can save the file without any hiccups in safe mode, this indicates some problematic add-ins are behind the error. You can remove the Excel add-ins by following these steps:

- Open your Excel application and go to **File > Options.**

- In **Excel Options**, select **Trust Center** and then click **Trust Center Settings**.
- In Trust Center Settings, click **Add-ins** and thenselect “**Disable all applications Add-ins”.** Click **OK.**

![Go to 'Add ins' and select disable all application add ins](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-add-ins-and-select-disable-all-application-add-ins.jpg)

### **Method 6: Check File Permission**

You can get the “Excel cannot access the file” error if you don’t have sufficient permissions to modify the Excel file. You can check and provide the write permissions to fix the issue. Here’s how to do so:

- Open Windows Explorer.
- Find the affected Excel file, right-click on it, and click **Properties**.  

![Click Properties Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-properties-option.jpg)

- In the **Properties** window, click the **Securities** option and click **Edit**.

![Go to Security and then click Edit option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/go-to-security-and-click-edit-option.jpg)

- In the **Security** window, select the **user names** under **‘Group or users name’**.
- Check the file permissions and make sure the write option is enabled. If not, then grant the permission. Click **Apply** and then **OK.**

### **Method 7: Check External Links**

The “Excel cannot access the file” error can also occur due to broken external links in the Excel file. External links are references to the data or content in other files. The link usually breaks if the file has been moved to another location or the file name is changed. You can check and [change the source of link.](https://support.microsoft.com/en-gb/office/fix-broken-links-to-data-84f494f9-1da9-460a-aa83-aba07108bc97)

### **Method 8: Repair your Excel File**

Excel may fail to read the file if it is corrupted or damaged. If the error “Excel cannot access the file” has occurred due to file corruption, then try the Excel’s Open and Repair utility to repair the Excel file. Here are the steps:

- In the Excel application, click the **File** tab and then select **Open.**
- Click **Browse** to select the problematic workbook.
- The **Open** dialog box will appear. Click on the corrupted file.
- Click the arrow next to the Open button and then select **Open and Repair.**
- You will see a dialog box with three buttons – **Repair, Extract Data,** and **Cancel.**

![Click repair option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-repair-option.jpg)

- Click on the **Repair** button to recover as much of the data as possible.
- After repair, a message is displayed. Click **Close**.

If the [Open and Repair utility fails to work](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), it indicates the Excel file is severely corrupted. Use Stellar Repair for Excel to repair severely corrupt Excel file. It helps recover all the components of the corrupted Excel file, such as charts, formulas, etc. without making any changes to the original file. It can also fix all types of corruption-related errors. You can use Stellar Repair for Excel to repair Excel files created in all Excel versions – from 2007 to 2023.

## **Closure**

The “Microsoft Excel cannot access the file” error can occur due to numerous reasons. Follow the troubleshooting methods, such as checking file location, path, permissions, etc., as discussed above to fix this error. Sometimes, Excel throws this error if the file you are trying to save is corrupted. You can try repairing the file using the built-in utility – Open and Repair. If the file is severely corrupted, then you can use [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It can repair damaged Excel files (.xls, .xlsx, .xltm, .xltx, and .xlsm) with complete integrity.



## Quick Fixes to Repair Microsoft Excel 2013/2016 Content related error

**Summary:** The blog outlines some quick tips to fix ‘We found a problem with some content’ error in Microsoft Excel 2013/2016. It explains manual procedure to resolve the error and also suggests an automated tool to perform the repair process to retrieve all possible data from a corrupt workbook.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Sometimes, when opening an MS Excel file, you may receive an error message that reads:

“**We found a problem with some content in ‘filename.xlsx’. Do you want us to try to recover as much as we can? If you trust the source of this workbook, click Yes.**“

![Microsoft Excel Content Error](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/07/Microsoft-Excel-Content-Error.jpg)

Figure 1 – Excel ‘found a problem with some content’ Error Message

## **What Causes ‘We Found a Problem with Some Content’ Error?**

There is no clear answer as to what results in the Excel error – ‘**We found a problem with some content in <filename.xlsx>**’. However, based on some user experiences, it appears that the error occurs due to corruption in an Excel workbook. It may turn corrupt when:

- You try opening the Excel file saved on a network-shared drive.
- A string is added in a cell in Excel, instead of a numeric value.
- Text values in formulas exceed 255 characters.

## **How to Resolve ‘We Found a Problem with Some Content’ Error?**

**Follow these tips to fix the Excel error:**

**IMPORTANT!** Before you follow the tips to resolve the Excel error, keep these points in mind: Make sure you have closed all of the opened Excel workbooks. Try restoring Excel file data from the most recent backup copy. If you don’t have a backup copy, make a copy of the corrupt Excel file and perform repair and recovery procedures on that backup copy.

### **Tip #1: Repair Corrupt Excel File**

File Recovery mode is a native Excel recovery utility that automatically opens whenever any inconsistencies are found in the worksheet. If Microsoft doesn’t detect any issue or fails to open the File Recovery mode, you can start it manually to recover the corrupt Excel file. To do so, follow the steps below:

1. Click on the **File** menu, and then select **Open**.
2. In the **Open** dialog box, navigate to the folder location where the corrupt Excel file is saved.
3. Select the corrupt file, and then click on arrow sign available next to **Open** button to select **Open and Repair** option.

![Open and Repair](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/03/Open-and-Repair.png "MS Excel Content error")

Figure 2 – Open and Repair Feature in Excel

1. Next, click **Repair to recover maximum possible data**.
2. If the repair is not able to recover the data from the workbook, select **Extract Data** to extract all possible formulas and values from the workbook.

If repairing the corrupt Excel file doesn’t work, you can try an Excel file repair tool to fix corruption errors. You can also try to recover data from the corrupt file manually by following the next tips.

**Read this:** [What to do when Open and Repair doesn’t work?](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

### **Tip #2: Set Calculation Option to Manual**

To make the file accessible, try setting the calculation option in Excel from automatic to manual. As a result, the workbook will not be recalculated and may open in Excel. For this, perform the following:

1. Click **File,** and then click **New**.
2. Under **New**, click the **Blank workbook** option.
3. When a blank workbook opens, click **File** > **Options**.
4. Under the Formulas category, pick Manual in the **Calculation options** section, and then click **OK**.

![calculation options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/03/calculation-options.png "MS Excel Content error")

Figure 3 – Select Manual in Calculation options

1. Now, again click on the **File** menu and then click **Open**.
2. Navigate to the corrupt workbook, and double-click it.

When the workbook opens, check if it contains all the data. If not, proceed to the next tip.

### **Tip #3: Copy Excel Workbook Contents to a New Workbook**

Several users have reported that they were able to fix ‘_We found a problem with some content in <filename>’_ error message by copying contents from the corrupt workbook to a separate workbook. **Detailed steps are as follows**:

1. Open the Excel workbook in **‘read-only’** mode, and copy all its contents.
2. Create a blank new workbook and paste the copied contents from the corrupt file to the new file.

### **Tip #4: Use External References to Link to the Damaged Workbook**

Use external references to link to the corrupted workbook. By implementing this fix, data contents can be retrieved. However, it is not feasible to recover formulas or calculated values using this solution.

**Follow the steps below:**

1. In Excel 2013/2016, click **File** > **Open**.
2. Navigate to the **folder** where the corrupt file is **saved**.
3. Right click the file, select **Copy,** and then click on **Cancel**.
4. Again, click on **File** and then **New**.
5. Under **New** option, click on **Blank workbook**.
6. In the **cell A1** of new workbook, type **\=File Name!A1** (where File Name indicates the name of the damaged workbook being copied in **Step 3**).
7. If **Update Values** dialog box appears, click the corrupt workbook, and choose **OK**.
8. If **Select Sheet** dialog box appears, click the appropriate sheet, and then click **OK**.
9. Select cell **A1**.
10. Next, click **Home,** and then click **Copy** (or, press Ctrl +C).
11. Starting in **cell A1**, select area approximately the same size as that of the cell range that contains data in the damaged workbook.
12. Next, click **Home** and select **Paste** (or click Ctrl + V).
13. Keep the range of cells selected, click **Home** and then **Copy**.
14. Finally, click on **Home**, click on the arrow associated with **Paste** and under **Paste Values** click on **Values**.

This will remove the link to the corrupt workbook and will retrieve data. But, keep in mind, the recovered data will no longer contain formulas or calculated values.

## **Alternative Solution – Stellar Repair for Excel**

If the above manual methods fail to fix the ‘We found a problem with some content in Excel error’, try using the Stellar Repair for Excel software to resolve this error. The software helps repair and recover corrupt Excel files in just a few clicks. It can be used on a Windows 10/8/7/Vista/XP/NT machine to repair a corrupted workbook and recover every single bit of data from all the versions of the Excel workbook.

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/07/free-download-1-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

**Read this:** [How to repair corrupt Excel file using Stellar Repair for Excel?](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **Conclusion**

In this blog, we discussed some possible reasons behind Microsoft Excel 2013/2016 _‘We found a problem with some content’_ error. The error may occur when an Excel file becomes corrupt. You may try repairing the corrupted Excel file manually by using the built-in ‘Open and Repair’ feature. Or, try the manual workarounds to extract data from the corrupt file discussed in this post. If the manual solutions don’t work for you, using [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) can come in handy in repairing the corrupt Excel (.xls/.xlsx) file and recovering the complete file data.


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


