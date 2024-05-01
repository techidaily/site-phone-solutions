---
title: Fixed Excel 2019 VBA Runtime Error 9 Subscript Out of Range | Stellar
date: 2024-03-13 21:26:12
updated: 2024-03-14 15:38:56
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2019 VBA Runtime Error 9 Subscript Out of Range
excerpt: This article describes Fixed Excel 2019 VBA Runtime Error 9 Subscript Out of Range
keywords: repair .xls,repair excel file,repair damaged .xltx,repair corrupt .xltm,repair corrupt excel file,repair corrupt .xlb files,repair corrupt .xls files,repair excel 2013,repair damaged .xlb,repair corrupt .csv files,repair damaged .xlsm files,repair corrupt .xlsm
thumbnail: https://www.lifewire.com/thmb/IOefl4e_6Yr9ZAczXaVDoLd0G88=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/GettyImages-1365158534-d6952b03afff43f7a1c5f0405f31dc92.jpg
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



## File Format and Extension of \[filename\] don't Match in Excel File

**Summary:** The “File format and extension of \[filename\] don't match. The file could be corrupted or unsafe” error message indicates that the Excel file you’re trying to open is unsupported, unsafe, or corrupted. Read this article to learn more about this error and how to fix this error. It also mentions an advanced Excel recovery tool to repair the corrupted Excel file and retrieve all its data in a few clicks.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

You can encounter the **“File format and extension of \[filename\] don’t match. The file could be corrupted or unsafe”** error when the Excel application detects any issue with the file. This happens when you try to open an old version file format in a newer version or if the file is received from an unsafe destination. This can prevent you from opening the Excel file.

**As indicated from the error message, this error occurs due to the following reasons:**

- The file has incorrect file extension.
- The file is corrupted.
- The file you are trying to open is protected.

Now, let’s see how to resolve this Excel error.

## **Methods to Fix the “File format and extension of \[filename\] don’t match” Error**

Try the following methods to troubleshoot the “File format and extension don’t match” error in Excel.

### **Method 1: Rename the Excel File**

You can face the “File format and extension don’t match” issue if the file has incorrect extension. It can occur if the file extension has been altered or you’ve mistakenly saved the file with incorrect extension. To fix this, you can try renaming the Excel file with the correct file extension.  

### **Method 2: Check the Default Excel File Format**

Different versions of Microsoft Excel use different default file formats. For example, .xls is the default file format of older versions (2003 and lower) of Excel, whereas .xlsx format is used by the newer versions (2007 and later). Opening the Excel file with an incompatible extension can cause the “File format and extension don’t match” issue. You can check the Excel version you are using and ensure it’s compatible with the Excel file you are trying to open.

### **Method 3: Change the Protected View Settings**

You may receive the “**File format and extension of excel don’t match**” error if the Excel file is protected. You can check and try disabling the [Protected View settings](https://support.microsoft.com/en-au/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653).  

**Caution:** Changing the Protected View settings can put your system at risk. If the Excel file is being downloaded from the internet, it may contain viruses that can infect your system. So be careful before disabling the Protected View settings.

### **Steps to Change Protected View Settings in Excel:**

- In the Excel’s File menu, click on **Options.**
- Select **Trust Center > Trust Center Settings**.

![Go to Trust Center and then click Trust center settings.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/go-to-trust-center-and-then-trust-center-settings.jpg)

- Under **Trust Center**, select **Protected View** and disable the below three options:
- Enable Protected View for files originating from the internet.
- Enable Protected View for files located in potentially unsafe locations.
- Enable Protected View for Outlook attachments.

![In the Trust Center Window, go to the Protected View tab and Disable all Protected View Checkboxes.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/select-all-the-options-under-protected-view.jpg)

- Click **OK.** Then, try to open the Excel file.

### **Method 4: Check and Provide the Excel File Permissions**

Sometimes, you can get the error if you don’t have sufficient permissions to open the Excel file. This usually happens when you try to open the Excel file received from other sources. You can check and provide the desired permissions to fix the error. Here are the steps:

- Locate the affected Excel file, right-click on it, and select **Properties.**

![Right Click on Excel file and click on Properties](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/go-to-folder-and-select-properties.jpg)

- In the **Properties** window, click the **Securities** option and select **Edit.**

![In Properties, Go to the Security Tab and click on Edit.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-security-and-then-click-edit-option.jpg)

- In the **Security** window, under **‘Group or users name’**, select the user names. Check the file permissions and make sure **Full Control** is enabled. If not, then click on the **Add** option.

![click add option under permissions](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-add-option-under-permissions.jpg)

- Click on the **Advanced** option in the **Users, Computers, Service Accounts, or Groups** window**.**

![Under Users, Computers, Service Accounts, or Groups window, click on Advanced.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-advanced-option-under-user-and-object-type-option.jpg)

- Click the **Find Now** option. A list of all users and groups appears in the search field.

![In the Select Users, Computers, Service Accounts, or Groups window, click on Find Now.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-find-now-option.jpg)

- Select **“Everyone”** from the list and then click **OK.**

![Select Everyone from the Search Results and Click on OK.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/selecting-everyone-from-the-listed-objects.jpg)

- In the **object names** field, you will see ‘**Everyone’**. Click on **OK.**

![After the ‘Everyone’ username is entered in the object names field, click on OK.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/selecting-everyone-from-the-listed-objects-1.jpg)

- In the **Permissions** window, select **“Everyone”** and enable all options **(Full Control, Modify, Read & Execute, Read,** and **Write**) under **Permissions for Everyone**.

![Allow all Permissions for ‘Everyone’ by checking the boxes under Allow](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/allow-all-permissions-and-then-apply.jpg)

- Click **Apply** and then **OK.**

### **Method 4: Repair your Excel File**

As the error message indicates, corruption is one of the causes of the “File format and extension of \[filename\] don’t match” error. If your file is corrupted, you can repair it using Microsoft’s built-in Open and Repair tool. Here are the steps to run the Open and Repair tool to repair corrupted Excel file:

- In Excel, click on **File.**
- Click **Open** and then click on **Browse** to select the corrupted Excel file.
- In the **Open** dialog box, click the Excel workbook (in which you are facing the error).
- Click the arrow next to the **Open** button and select **Open and Repair**.
- Then, click **Repair** to recover as much data as possible.
- The Excel prompts a message after the repair process is complete. Click **Close.**

The [Open and Repair utility may fail](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to give the intended results. In such a case, you can repair the corrupted/damaged Excel file using a specialized [Excel repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). Stellar Repair for Excel is one such tool that can repair severely corrupted Excel files. With the help of this tool, you can quickly recover all the objects from the Excel file. The tool has a simple user interface that even a non-technical can use to repair the Excel files. The tool can also repair multiple Excel files at once. You can check the tool’s functionality by downloading its demo version.

## **Closure**

You can encounter the “File format and extension of \[filename\] don’t match” error due to different reasons. To resolve the issue, you can check the file extension, permissions, protected settings, etc. If you suspect the error has occurred due to corruption in the Excel file, you can try repairing the Excel file using the Open and Repair tool. If nothing works for you, then try [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It can repair highly damaged Excel files and recover all the data while preserving the file properties and cell formatting. The tool can help you fix all the common corruption-related errors quickly.



## Solved - The File is Corrupted and Cannot be Opened - Excel

**Summary:** Unable to open Excel file due to the error ‘The file is corrupted and cannot be opened’? Read this blog to find more details about the error, possible reasons behind it, and solutions to fix the error. In addition, the blog mentions about Stellar Repair for Excel software that can help fix the Excel error in a few clicks. Download the software now and see free preview of the file.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

## **About the Error**

**Microsoft Excel** is a widely used spreadsheet application that comes bundled with MS Office. Users tend to update the application with new security patches and features. Sometimes these updates can cause problems, and result in “**The file is corrupted and cannot be opened**” error.

![The File is Corrupt and Cannot be Opened Error Message](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/04/the-file-is-corrupt-and-cannot-be-opened.jpg)

Figure 1 – Excel File Corrupted Error Message

## **Other Possible Reasons behind ‘The File is Corrupt and Cannot Be Opened’ Excel Error**

- Opening an older Excel version file in a newer version of Excel. For instance, opening Excel 2013, 2010, or earlier versions in Excel 2016.
- When attempting to open a Microsoft Office (Excel) email attachment in Microsoft Outlook 2010, MS Office 2010 reports a problem with the file preventing it from opening.

## ****How to Fix the ‘Excel File is Corrupt and Cannot Be Opened’ Error?****

Here are a few possible solutions that you can try to fix the ‘Excel file is corrupt and cannot be opened’ issue and open your Excel file.

**Solution 1**: Changing Component Services Settings

**Solution 2**: Changing the Protected View Settings

**Solution 3**: Repair Excel Files using Excel Repair Software

### **Solution 1: Changing Component Services Settings**

**\[Caution\]** Changing Component Services settings requires making changes to the registry, and any mistake can harm your computer.

**Follow these steps to change ‘Component Services’ settings:**

- Click ‘**Start**’ or ‘**Win+R**’ and type ‘**dcomcnfg**’ and press ‘**Enter’.**

- In the navigation pane, expand the ‘**Component Services**’, and then expand ‘**Computers**’.

___

![Changing Component Services Settings](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/04/Changing-Component-Services-Settings.jpg)

Figure 2 – Component Services Settings

- Next, right-click on ‘**My Computer’**, and then click ‘**Properties**’.

**When the ‘My Computer Properties’ dialog box appears, click on the ‘Default Properties’ tab and then set the following values:**

- **Default Authentication Level**: Connect
- **Default Impersonation Level**: Identify

![My Computer Properties](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/04/My-Computer-Properties.jpg)

Figure 3 – Illustrates My Computer Properties

- Click ‘**OK**’ to change ‘**Default Properties**’

### **Solution 2: Changing the Protected View Settings**

**\[Caution\]** Disabling the ‘Protected View’ can put your system at high risk. Viruses attached to the Excel files can attack and infect your system. Be careful before using this option.

Excel 2010 file cannot open due to the ‘**Protected View**’ setting in Microsoft Outlook 2010. And so, changing the setting may help fix the error. For this, perform these steps:

- Open MS Excel 2010, go to the ‘**File’** menu and click **‘Options’.**

![Select Options in Excel 2010](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/02/Excel-options.jpg)

Figure 4 – Options

- When the ‘Excel Options’ window opens, click on ‘**Trust Center**’ and then on ‘**Trust Center Settings**’.

![Trust center settings in Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/02/Open-trust-center-settings.jpg)

Figure 5 – Open Trust Center Settings

- Next, choose **‘Protected View**’ and uncheck all the options including ‘**Enable Protected View for Outlook attachments’** if you use Outlook for email.

![change protected view settings ](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/02/Uncheck-protected-view-settings.jpg)

Figure 6 – Uncheck Protected View Settings

- Click ‘**OK’.** Restart the application and try opening the Excel file again.

If none of the above solutions works for you, your Excel file is likely severely corrupt. To repair corrupt Excel files, you need to use advanced options like Stellar Repair for Excel tool. It repairs corrupt and damaged Excel files and helps in retrieving lost data.

### **Solution 3: **Use Excel File Repair Tool****

Considering the risks associated with the above solutions, it’s better to use an **Excel repair tool** to repair **single** or **multiple** corrupt Excel files at once. The process is simple, and even a novice can use the Excel file repair tool to repair Excel files with the help of the following steps:

- Download Stellar Repair for Excel and install it.

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/01/Free-download-for-windows-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

- Launch the tool. In the tool’s main interface, click ‘**Browse**’ to select the file. If you don’t know the file location use the ‘**Search’** option.

![Browse and Search](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2018/04/Browse-and-Search.jpg)

Figure 7 – Illustrates Selecting Corrupt Excel File in Stellar Repair for Excel

- Select the file, and then click on **Repair**.

![select corrupt file and repair](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/08/2-select-file.jpg)

Figure 8 – Illustrates Initiating Excel File Repair in Stellar Repair for Excel

- The software scans and lists the Excel file in the left pane. Click on the file to preview its recoverable objects in the right pane.

![preview recoverable excel objects](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/08/4-preview.png)

Figure 9 – Illustrates Preview of Recoverable Excel File Objects

- Save the repaired file at either the default location or a user-specified location.

![select repaired file location](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/08/6-save-file.jpg)

Figure 10 – Illustrates Saving Repaired Excel File in Stellar Repair for Excel

- Click ‘**OK’** to save the repaired Excel file. After the repair process is completed, browse to the location and open it with MS Excel 2010 or any other version.

![repaired file saved Dialog Box](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/08/7-saving-complete.jpg)

Figure 11 – Illustrates Saving Complete Message in Stellar Repair for Excel

You will be able to access your Excel file from the selected location.

## **Conclusion**

You can use the first two possible solutions to fix the “The file is corrupted and cannot be opened” error. If you can access the file, save its data and restore the default settings. However, if the file is corrupt and the data retrieved using the first two solutions is inconsistent or incomplete, use Stellar Repair for Excel. This tool can help you recover Tables, Charts, Chart Sheets, cell comments, Images, and Formulas while preserving the worksheet properties and cell formatting. You can also preview the file and verify the data inside the file before saving it.



## Filter Not Working Error in Excel [Fix 2024]

**Summary:** The filter is not working issue in Excel can occur due to several reasons, like blank rows, hidden rows, merged cells, corrupted data, etc. In this post, we will mention the reasons why the filter is not working correctly in Excel and several fixes to resolve the issue. We will also mention an advanced Excel repair tool to repair the Excel file if corruption in file is the cause of the issue.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

You can use the Filter function in Excel to filter data in large-sized Excel files quickly. While using Excel filters, sometimes, you face a situation where the filter is disabled or may fail to function properly.

![Filter Option Disabled](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/filter-option-disabled-1024x112.jpg)

The Excel filter usually fails to work if you have not selected the complete and correct range of data. Let’s learn more about the “Sort and Filter not working in Excel” issue and look at the possible methods to fix it.

## **Why the Filter is not Working in Excel?**

You can face the “filter is not working” issue if you are applying the filter on a protected worksheet or trying to find the data from a hidden row. Besides this, there could be many other reasons contributing to this issue, such as:

- The data you are trying to filter is in merged cells.
- The Excel file automatically selected the data up to the first empty cell, excluding the remaining rows.
- Grouped sheets in Excel file.
- Blank row in the Excel sheet.
- You are trying to apply a filter on an invalid data range.
- The workbooks in which you’re facing the filter issues are corrupted.
- You are specifying incorrect criteria in the filter columns.

## **Solutions to Resolve the Filter is not Working Issue in Excel**

There might be two scenarios: the Excel filter option is disabled/grayed out or the filters fail to function properly. You can follow the given troubleshooting solutions to resolve the issue based on the scenario you’re facing.

## **Scenario 1 – Filter Option is Disabled or Grayed Out**

### **Method 1: Check and Un-group the Worksheet**

When you apply filters to a single sheet in a grouped set, Excel disables the filter option in other sheets within the group. You can check the grouped sheets and try ungrouping them to enable the filter option. Here’s how to do so:

- In the Excel file, go to the **Group** section.

![Excel file navigation: Accessing the Group section
](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/go-to-group-section-1024x114.jpg)

- Right-click on the **Ungroup Sheets.**

Alternatively, you can press the Shift + Alt + Left keys to ungroup the sheets.

### **Method 2: Unprotect Worksheet**

The “disabled Excel filter” issue can also occur if your worksheet is protected. You can unprotect the worksheet to enable the filter option. To do so, go to the **Review** tab and then select **Unprotect Sheet.**

![Excel file: Navigating to Group section, resolving 'disabled Excel filter' issue with worksheet protection, unprotecting sheet from Review tab for filter activation.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/go-to-review-and-select-unprotect-sheet-1024x115.jpg)

### **Method 3: Check and Uninstall Excel Add-ins**

Sometimes, the Excel filter gets disabled due to faulty or corrupted Excel add-ins. You can run the Excel in Safe mode to check whether the issue has occurred due to add-ins. To do this, type excel /safe in the Run window and click **OK.**

![Troubleshooting disabled Excel filter caused by add-ins: Running Excel in Safe mode with 'excel /safe' in Run window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/type-excel-safe-command.jpg)

In safe mode, if you see the filter option, it indicates some problematic Excel add-ins were causing the issue. In such a case, you can check and uninstall the faulty Excel add-ins to fix the issue.

## **Scenario 2 – Filter is not Working**

### **Method 1: Try Clearing Filters**

Sometimes, the Excel filter fails to work correctly if some filters from the previous sessions are still active. In such a case, you can clear the applied filters. Follow the below steps:

- In Excel file, click Sort & Filter option.
- Select clear.

![Excel: Clicking 'Sort & Filter' and selecting 'Clear' option.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-clear-option.jpg)

### **Method 2: Select Entire Data**

The filter not working issue in Excel can occur when the range selected for filtering is incomplete or incorrect. You need to make sure that you’ve selected the entire data range in Excel. You can use the Ctrl+A keys to select the entire content in the worksheet.

### **Method 3: Check and Delete Blank Cells from the Table’s Columns**

When you apply a filter to the data, Excel expects data to be in a continuous range. Excel filters do not consider the blank cells, thereby resulting in incorrect functioning of the filter. To resolve this issue, check and delete all blank cells. In case your Excel file is too large to delete the blank cells, then you can add a “Serial number” row as an alternative. Adding serial number row creates a data continuity, thus helping in fixing the filter-related issue.

### **Method 4: Unhide Hidden Rows and Columns**

Hidden rows or columns in worksheets can also affect the filter functionality. You can check and unhide rows/columns to troubleshoot the issue. Here is how to do so:

- In the affected Excel file, go to Home.
- Click on **Format > Hide & Unhide**.

![Excel file: Navigating to Home, accessing Format > Hide & Unhide.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-format-select-hide-or-unhide-option-1024x228.jpg)

- Click **Unhide Rows** or **Unhide Columns** (as required).

![Selective unhiding in Excel: 'Unhide Rows' or 'Unhide Columns' as needed.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/10/click-unhide-rows-unhide-columns.jpg)

### **Method 5: Unmerge Cells**

You can experience the filter in Excel is not working issue if you are using the filter to extract data from merged cells. Ensure to unmerge the “merged cells” before applying a filter in Excel. Follow the below steps to unmerge the merged cells in Excel:

- Navigate to the **Home** option.
- In the toolbar, select the **Merge & Center** option.
- Click **Unmerge Cells.**

### **Method 6: Repair the Workbook**

Sometimes, the **Filter Not Working in Excel** issue can occur due to inconsistencies in file structure. If these issues occurred due to corruption in the worksheet, you can repair it using the Open and Repair tool. It is an in-built tool in Excel that is used to repair corrupted Excel files. Here are the steps to use this tool:

- In the Excel application, navigate to the **File** option.
- Click **Open** and then click **Browse** to choose the Excel file.
- In the **Open** dialog box, click the problematic Excel file.
- Click the arrow next to the **Open** option and select **Open and Repair.**
- Click **Repair** to recover as much data as possible.
- The application prompts a message after the repair process is complete. Click **Close**.

In most cases, the Open and Repair tool can easily fix corruption issues in the Excel file. However, for any reason, if the [open and repair tool doesn’t work](https://www.stellarinfo.com/blog/ms-excel-open-and-repair-option-is-not-working/) you can consider repairing the file using a professional Excel Repair tool. Stellar Repair for Excel is one such advanced and secure tool to repair Excel files. With this tool’s powerful scanning capabilities, you can repair highly corrupted Excel files and recover all their objects with complete integrity. The tool is compatible with all Windows editions, including the latest Windows 11.

## **Closure**

Several reasons are associated with the **filter not working issue in Excel**. The filter option may not work as expected if you have not selected the complete and correct range of data or for many other reasons. You can follow the troubleshooting methods discussed above to fix the issue. If the filter fails to work due to corruption in the workbook, then try [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It is an advanced tool that can even repair severely damaged files. It also helps to recover all the data from corrupted files without changing the original formatting. You can check the tool’s functionality by downloading its demo version. It allows you to preview all the repairable objects in the corrupted Excel file.


## How to repair corrupt Excel file

[**Stellar Repair for Excel**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) is an excellent tool to repair corrupt or damaged MS Excel files. Mentioned below are the steps to perform Excel repair with this tool:

- Download & Run the Stellar Repair for Excel.

[![free download](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)](https://cloud.stellarinfo.com/StellarRepairforExcel-KB.exe)

- A dialog box appears on your screen, click 'OK' to proceed.

![Stellar Repair for Excel - Main Interface](https://www.stellarinfo.com/blog/wp-content/uploads/2019/08/1-user-interface.png)

- To select your corrupt .XLS or .XLSX file, click 'Browse' button. However, if you do not know the location of your .XLS or .XLSX file, the software provides you the option 'Search' to search for your corrupt Excel files.

![Select excel file](https://www.stellarinfo.com/blog/wp-content/uploads/2019/08/2-select-file.jpg)

- Select the checkboxes against the files that you want to repair and click 'Repair'. This starts the scanning process.

![repair process](https://www.stellarinfo.com/blog/wp-content/uploads/2019/08/3-repair-process.jpg)

- The list of all the files that the software has scanned is displayed in the tree-view in the left pane. Click on a file from this tree-view to see its preview in the middle pane. From this list, you can select the file that you want to recover.

![Preview](https://www.stellarinfo.com/blog/wp-content/uploads/2019/08/4-preview.png)

- You can either select the 'Default location of file' or 'Select New Folder' in the 'Save Document' dialog box to save the repaired files.

![Save file](https://www.stellarinfo.com/blog/wp-content/uploads/2019/08/6-save-file.jpg)

Stellar Repair for Excel Stellar Repair for Excel is the best choice for repairing corrupt or damaged Excel (.XLS/.XLSX) files. This Excel recovery software restores everything from corrupt file to a new blank Excel file.

[Learn More ![red arrow](https://www.stellarinfo.com/image/catalog/blacktheme/data-recovery-standard/red-arrow.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)




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

