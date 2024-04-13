---
title: Fixed Excel 2013 VBA Runtime Error 9 Subscript Out of Range | Stellar
date: 2024-03-13 15:55:19
updated: 2024-03-14 19:22:13
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2013 VBA Runtime Error 9 Subscript Out of Range
excerpt: This article describes Fixed Excel 2013 VBA Runtime Error 9 Subscript Out of Range
keywords: repair damaged .xlsx,repair corrupt .xlsx files,repair .xlsm files,repair corrupt .xlb files,repair damaged .xlb files,repair .xls,repair damaged .csv
thumbnail: https://www.lifewire.com/thmb/NgJyQBGFmsrhkEXhA5pREWNGnoI=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/photo-editor-working-on-computer-at-desk-in-office-946928370-5b85791746e0fb0025b40d84.jpg
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


## How to Fix "Errors were detected while saving Excel" Error?

When trying to save the Excel file, you might face unexpected errors. The “Errors were detected while saving Excel” is one such error. It can also occur when using VBA in Excel. The complete error message appears as:

**“Errors were detected while saving \[file name\]. Microsoft Excel may be able to save the file by removing or repairing some features. To make the repairs in a new file, click Continue. To cancel saving the file, click Cancel.”  
**

The error can occur if the features (Pivot tables, charts, macros) used in the [Excel file get corrupted](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). However, there could be several other reasons behind the occurrence of the error. Let’s discuss them.

## **What Causes the "Errors were detected while saving Excel" Error?**

There are various reasons why you encounter this error. Here are some of them:

- Incompatible pivot table in the Excel file
- Large or uncompressed images in the Excel file
- File-sharing properties are not allowing file saving
- Excel file is corrupted
- Large-sized Excel file
- File version incompatibility
- VBA code is corrupted

## **Ways to Fix the “Errors were detected while saving Excel” Error**

You’re not able to save the Excel file if there is no storage space on your hard drive. So, first check if your hard drive has sufficient storage space to save the file. If this is not the case, then it might happen that your antivirus program is interrupting the saving process. To check this, temporarily disable your antivirus program and then try to save the file. If still your Excel is throwing the “Errors were detected while saving Excel” error, then follow the below given methods to fix the error:

###  **Method 1: Open the Excel in Safe Mode and Disable the Add-ins**

When you open Excel in safe mode, it opens without the third-party add-ins. This helps in finding out if any add-ins are causing the error.

 Here’s how to open the Excel in safe mode:

- Open the Run window by pressing **Windows key + R**.
- Type **excel /safe** in the Run window.  

    ![Excel Save Mode Command](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/excel-save-mode-command.jpg)?

- Next, click on **OK**.
- It will open Excel in safe mode.
- Now, try to open and save the affected file.

If you are able to save the file without any issue, then this means that the error has occurred due to third-party add-ins or settings. You can try disabling the add-ins to fix the issue. To do this, follow these steps:

- First, open Excel.
- Then, go to the **File** tab and click **Options**.

![Go To Options Window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/go-to-options-window.jpg)

- In **Excel Options**, click on the **Add-ins**

![Select Add-ins](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/select-add-ins4.jpg)

- Under the **Manage** section, select **Excel Add-ins** and then click on the **Go**

![Excel Add-ins Drop-down](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/excel-add-ins-drop-down.jpg)

- In the **Add-ins** dialog box, unselect the **add-ins** under the **Add-ins available** option and click **OK.**  

    ![Add-ins Window](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/add-ins-window.jpg)

**_Note_**_: Disabling add-ins does not remove them from the system. To remove them permanently, you need to uninstall them._

### **Method 2: Check the Excel File Name**

Some users have observed this error when saving the Excel file with an invalid name. You can check the file name and ensure that it should not contain more than 218 characters. If the name exceeds the required limit, then try shortening the file name or move the file to a folder with a short path name.

###  **Method 3: Copy the Data from the Affected File to a New File**

If you are not able to save the Excel document, then try copying the data from the affected file to a new Excel file. Then, save the new file with a different name. This helps in resolving the issue.

### **Method 4: Check and Provide File Permissions**

You may experience the “Errors were detected while saving Excel" issue when you do not have desired permissions to modify the folder in which your Excel file is located. To modify the folder, you should have read, write, and create permissions. You can check and provide the desired permissions using the below steps:

- Navigate to the Windows **Program Files** and then find the desired folder (where the Excel file is saved).
- Right-click on the folder and then choose **Properties**.
- Select the **Security** tab and then click
- Click on **Change Permissions** in the **Advanced Settings**
- Click **Administrators** and then click **Edit**.
- Now set the **Apply to drop-down** button to **This Folder, Subfolder, and Files**.
- Click on the **Full Control** field and then click **Apply > OK**.

###  **Method 5: Check Pivot Tables in Excel Sheet**

You can review Pivot tables to see if they are causing the “Errors were detected while saving Excel” error. To do so, follow the below steps:

- Click **Power Pivot > Manage**.  

    ![Check Pivot Table In Excel](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/check-pivot-table-in-excel.jpg)

- Check the tabs in the **Power Pivot**
- Check if all the formulas used in the table are correct. Sometimes, even a small typo can create an issue in Excel.

### **Method 6: Repair Your Excel File**

The “Errors were detected while saving Excel” issue can also occur if the Excel file is corrupted. In such a case, you can take the help of the built-in utility in Excel – Open and Repair to repair your Excel file. Here’s how to use the tool:

- In Excel, click the **File** tab and then click **Open**.
- Click **Browse** to select the desired file.
- The Open dialog box is displayed. Click on the corrupted file.
- Click on the arrow next to the **Open** button and then click **Open and Repair**.
- Click on the **Repair**

![Click On Repair Button](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/click-on-repair-button.jpg)

- After repair, a message will appear (as shown in the below figure).  

    ![Message Appear After Repair](https://www.stellarinfo.com/public/image/catalog//article/Repair-Office-Documents/errorsweredetected/message-appear-after-repair.jpg)

- Click **Close**.

 However, sometimes, the Open and Repair utility fails to fix the file if it is severely corrupted or large-sized. In such a case, you can take the help of a third-party Excel repair software, such as **[Stellar Repair for Excel.](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** The tool performs a comprehensive scan of the corrupted Excel file to fix the issues and recover all the items from the file without changing the original formatting. It can recover pivot tables, charts, images, engineering formulas, etc. The tool is compatible with Windows 11/10/8.1/8/7. You can download the free trial version of the tool to evaluate its functionality.

##  **Closure**

Many Excel users reported facing the situation when they are saving the Excel file. You can check the file’s compatibility to fix the “Errors were detected while saving Excel” issue. If you are getting this error in a Macro-enabled file then you can try deleting the VBA project from a document to resolve the issue. However, deleting the entire VBA code cannot be a better solution as it can lead a data loss in the Project you are working on. In the above article, you have learned the reasons behind the issue and discovered how to fix the error. Follow the methods and if none of them works then try using Stellar Repair for Excel. It is an advanced tool that can quickly repair corruption in Excel worksheets at any level. It lets you restore the corrupted components from the corrupted file without removing the existing data.


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


## Excel AutoRecover not working, what is next?

## Consider a Scenario

A professional with Windows 10 computer had MS Office 2016 installed on it. For an official purpose, he worked on an Excel workbook and saved it as an XLSX file. After working for hours on it, which was saved with a file name, a power outage occurred in his building for quite some time. After the power was back, he reopened Excel to find a list of recovered files in ‘Document Recovery’ section on the screen’s left side. However, the file that he had worked on recently was the ‘Original version,’ i.e. the last version saved by him and not the auto-saved Excel file. This meant the Excel document did not have any new data that was entered since the last time he saved it. Consequently, he lost hours of work. According to him, this happened despite the fact that the ‘AutoRecover’ feature was enabled. (Still, this needs to be checked and ensured.)

The 'AutoRecover' feature might not work in any of these cases:

- **AutoRecover Feature is disabled -** With this feature disabled, the Excel files are not auto-saved if the document is closed without saving, or the document closes unexpectedly due to an untoward incidence. To check, see if ‘Save AutoRecover information every \* minutes’ and ‘Keep the last auto-saved version if I close without saving’ checkboxes are checked or unchecked. If either one is unchecked or both are unchecked, it signifies that the AutoRecover feature is disabled. Else, the AutoRecover is enabled.
- **Corruption in the Excel XLSX file –** If ‘AutoRecover’ is enabled, most probably the cause is ‘damaged Excel XLSX file.’

Before discussing solutions to resolve the ‘Excel AutoRecover not working’ issue, let's have an overview of the 'AutoRecover' and 'AutoSave' features.

## A Brief Overview of Excel AutoRecover and AutoSave Feature

AutoRecover is an inbuilt feature in MS Excel 2019, 2016, 2013, 2010, 2003, and 2007 that allows saving all of the ‘open Excel files’ at a fixed-interval in a user-specified location or AutoRecover file location. Besides, AutoSave is an add-in that exists in Excel 2002 and earlier versions to save all open Excel files in case of a crash, power outage, or accidental closure of Excel files without saving. Users can recover these files if Excel closes suddenly, for instance, at the time of power outage or failure. The auto-recovered Excel files are saved at a default location.

## Methods to Fix ‘Excel AutoRecover Not Working’ Issue

### Manual Methods

### Method 1 - Enable the 'AutoRecover' Feature if Disabled

Make sure that you have the ‘AutoRecover' feature enabled in your Excel application. If not, follow these steps to enable it:

- Open Excel with MS Excel 2016/2019
- Click on File and then on Options tab
- In ‘Excel Options’, click on Save tab
- Check ‘Save AutoRecover information every \* minutes’ and ‘Keep the last auto saved version if I close without saving’ box
- Set the time in ‘Save AutoRecover information every \* minutes’
- Click on the OK button

Note: With this method, it is not possible to recover data from the current Excel file. From next time onwards, the AutoRecover feature starts working following which Excel shall start auto-saving Excel files as per the time set in ‘Save AutoRecover information every \* minute’.

### Method 2 - Repair Corrupt Excel File

If corruption in Excel XLSX file has resulted in ‘Excel AutoRecover not working’ issue, you will need to [repair Excel file](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). Use **‘Open and Repair’** inbuilt utility to fix and repair the damaged (corrupt) Excel file and extract its data. To use the inbuilt utility, execute the following steps:

- Go to location “C:\\Users\\AppData\\Local\\Microsoft\\Office\\UnsavedFiles” to find Excel TMP files and save it as XLSX file

Note: In Windows 8, the location is the same as mentioned above that is for Windows 10. In Windows 7, the location is “C:\\Users\\name\\AppData\\Roaming\\Microsoft\\Excel\\”

- Open a blank **Excel** sheet; click **File >> Open**
- Go to the location and folder containing the damaged Excel file
- In the **Open** dialog box, choose the damaged Excel file and click the arrow next to the Open button, and then click **Open and Repair**
- In the window that appears, click **Repair** to recover as much data as possible

Now, open the Excel (XLSX) file to check if the Excel file is repaired and its data is recovered.

### Use a Professional Excel File Repair Tool

Using a third-party Excel file repair tool can help you repair damaged Excel XLSX file and recover all the data. **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** is one tool you can rely on to repair severely corrupt Excel files (XLSX or XLS).

Here are a few other reasons why you should choose Stellar Repair for Excel software:

- Repairs Excel file while keeping the worksheet properties and cell formatting same as before.
- Recovers all of the Excel file components like tables, forms, reports, charts, chart sheets, cell comments, formulas, images, etc.
- Can batch repair multiple Excel files simultaneously
- Supports Excel 2019 and earlier versions

[![Free download Stellar Repair for Excel](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## Conclusion

To help resolve the problem of ‘AutoRecover not working’, different methods have been discussed, depending on the cause of the problem. These solutions can be implemented to check the possibility of getting back maximum data added in last saved version of the Excel file. However, to fix corruption in the excel file, using a specialized tool such as Stellar Repair for Excel software recommended by MS Excel Experts and MVPs can help. The software can repair severely damaged Excel file easily and efficiently.


## Resolve Compile Error in Hidden Module in Excel: Causes & Solutions

The hidden module in Excel refers to a container with VBA codes, custom queries, and complex macros. The compile error in a hidden (protected) module in the Excel worksheet usually occurs when doing different activities on a macro-enabled sheet, such as merging .xls files. The error can result in macros execution failure. You need to quickly resolve this compile error to restore full functionality of the VBA code. Below, we’ll be discussing the solutions to fix this Excel error. But before that, let’s see why this error occurs.

You may encounter the Compile error in hidden module due to one of the following reasons:

- The code in the workbook is not compatible with the Excel application.
- Manual queries created in a previous version are no longer compatible with your current version of Excel.
- Missing references.
- Invalid .exe files (control information cache files) are automatically created with ActiveX control insertion in Excel file.
- Protected module is corrupted.
- The workbook with hidden module is damaged or corrupted.
- Incompatible add-ins.
- Incompatible Excel file version.
- The module is protected or password-protected.
- Missing or corrupted mscomctl.ocx file.

Excel can throw the compile error while compiling the code that exists in the protected module. So, first check the error and identify the hidden module that is creating the issue. You can unprotect the module. Also, ensure that you have permission to access the VBA code in the module. If the error still exists, follow the below troubleshooting methods.

### Method 1: Re-register ActiveX Control Files or mscomctl.ocx Files

You can get the compile error in the Excel file, containing the VBA code related to ActiveX controls or OCX files. The ActiveX control files and OCX files (mscomctl.ocx files) are the components of Microsoft’s standard controls library. The compile error in the hidden module can occur if these files are missing. In this case, you can use the Regsvr32 tool to re-register the OCX files. The [Regsvr32](https://support.microsoft.com/en-au/topic/how-to-use-the-regsvr32-tool-and-troubleshoot-regsvr32-error-messages-a98d960a-7392-e6fe-d90a-3f4e0cb543e5) is a command-line utility to register and unregister OLE controls in the Windows registry.

### Method 2: Delete .exd Files

 The .exd files are temporary files created by Excel when inserting ActiveX controls objects. These temporary files can lead to a compile error if they are corrupted. So, if this issue has occurred, particularly in the Excel file containing ActiveX controls, then deleting .exd files might fix the issue. To delete the .exd file, follow the below steps:

- First, open the **Run** window by pressing the Windows+R keys.

![Open The Run Window](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/open-the-run-window.jpg)

- In the **Run** window, type **%appdata%**.

![Type App Data Command](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/type-app-data-command.jpg)

- In the **Roaming** window, click on the **Microsoft** option.

![Click On Microsoft Option Under Roaming](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-on-microsoft-option-under-roaming.jpg)

- Under **Microsoft**, you will see a list of folders. Search and click on **Forms.**
- Right-click on a file with .exd extension and select **Delete**.
- Once you delete the .exd files, restart your Excel application.

### Method 3: Rollback the Office Updates

MS Office updates or upgrades may also cause the compile error in hidden module in Excel. If the error has occurred after downloading the recent Microsoft Office updates, try [reverting to the previous version](https://support.microsoft.com/en-us/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841) or uninstalling the recent updates to fix the issue.

### Method 4: Unselect Missing References

The compile error in hidden module determine path in Excel can also occur if your file contains a reference to object library/type library, which is labelled as Missing. You can locate, check, and uncheck the references marked as ‘Missing’ to fix the issue. Here are the steps:

- Open your **Excel** and press **Alt + F11** keys.
- The **Visual Basic Editor** is displayed.

![Visual Basic Editor](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/visual-basic-editor.jpg)

- Go to the **Tools** option and then click **References**.

![Click On References Under Tools Option](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-on-references-under-tools-option.jpg)

- In the **References-VBAProject** window, under **Available References**, search and unselect the references starting as “Missing”.

![Unselect Missing References](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/unselect-missing-references.jpg)

- Click **OK**.

### Method 5: Check the Code in Module

The compile error in hidden module can occur if there are issues in the code within the module. The problems include incorrect or missing syntaxes, missing parameters/references, or the code contains incompatible functions or a wrong name of the object. You can check and fix these issues in the code by opening the VBA editor.

### Method 6: Check and Remove Add-ins

In Excel, the compile error in macro-enabled files can also occur due to incompatible add-ins. You can check and disable the **add-ins** in Excel using the below steps:

- First, open the **Run** window and type excel /safe and then click **OK**. The Excel application will open in safe mode.
- Now try to open the affected Excel file. If it opens without the error, then check and remove the latest installed Excel add-ins.
- Navigate to the **File** option and then select **Options**.
- In the **Excel Options** window, click **Add-ins**.

![Click Addins Select Latest Addins](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-add-ins-select-latest-add-ins.jpg)

- Under **Add-ins**, search and select the latest add-ins, and then click on **Go**.
- In the **Add-ins** window, uncheck the add-ins and then click **OK**.

![Select  Analysis Toolpak](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/select-analysis-toolpak.jpg)

- Restart Excel and then check if the error is fixed or not.

### Method 7: Repair the Corrupt Excel File

Corruption in the Excel file can affect the macros in the hidden module, which may result in the compile error. In such a case, you can try repairing the Excel file using Microsoft’s inbuilt utility -Open and Repair. To use this tool, follow these steps:

- Open your Excel application.
- Click the **File** tab and then click **Open**.
- Click **Browse** to select the affected workbook.
- The **Open** dialog box will appear. Click on the corrupted file.
- Click the arrow next to the **Open** button and then **Open and Repair**.
- You will see a dialog box with three buttons - Repair, Extract Data, and Cancel.

![Click On Repair Option](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/click-on-repair-option-1.jpg)

- Click on the Repair button to recover as much of the data as possible.
- After repair, a message is displayed. Click **Close**.

![Message Appear After Repair](https://www.stellarinfo.com/blog/wp-content/uploads/2023/07/message-appear-after-repair.jpg)

## **What if None of the Above Solutions Works?**

If the above methods fail to get rid of the “compile error in hidden module” in Excel, then use an Excel repair tool such as Stellar Repair for Excel. This tool is specifically designed to repair the corrupted Excel file. It can recover all the components from corrupted Excel file (macros, queries, formulas, etc.) without changing their original formatting. The tool is compatible with all Excel versions and can be downloaded on a Windows system. You can download the free trial version of Stellar Repair for Excel to scan the corrupted Excel file and preview the data.

## **Closure**

You can get the “compile error in hidden module” when Excel detects any issue while compiling the code in a protected module. It can occur when there is an issue with the macro-enabled Excel workbook or Excel add-ins. You can follow the above-mentioned methods to fix the issue. If the error occurs due to corruption in the database file, then you can try [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It can repair severely corrupted Excel files. It also helps recover all the Excel workbook’s components, including macros and queries. The tool has a simple and user-friendly interface.


## \[Fixed\] Excel Found a Problem with One or more Formula

**Summary:** The error ‘Excel found a problem with one or more formula references in this worksheet’ may appear while saving the Excel workbook. It occurs when Excel found a problem with the formula used in the sheet. However, it may also occur when the Excel workbook gets damaged or corrupt. In this guide, we’ve explained the reasons that may lead to this Excel error and methods to resolve the error, by using various Excel options and a third-party Excel file repair software.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

If you are experiencing the ‘Excel found a problem with one or more formula references in this worksheet’ error message in the Excel workbook, it indicates that the [Excel file is corrupt](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) or partially damaged. However, it may also occur due to incorrect reference to a wrong cell or object linking, which is not working. The complete error message says,

_‘Excel found a problem with one or more formula references in this worksheet. Check that the cell references, range names, defined names, and links to other workbooks in your formulas are all correct.’_

![Excel found a problem with one or more formula references](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/11/MS-Excel-problem-with-formula-reference.png)

In any case, resolving the error is critical as it doesn’t let you save the file and may result in loss of information from the Excel workbook.

## Reasons for Excel Formula References Error

A few reasons that may lead to such error are as follows,

- Wrong formula or reference cell
- Incorrect object linking or link embedding OLE
- Empty or no values in named or range cells
- Multiple Excel files (not common)

## Methods to Resolve ‘Excel Found a Problem with One or More Formula References in this Worksheet’ Error

Following are a few methods that you can follow to [fix Excel file](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) that can’t be saved due to problems with one or more formula references in the worksheet.

### Method 1: Check Formulas

If the problem has occurred in a large Excel workbook with multiple sheets, it’s quite hard to pinpoint the problem cell. In such cases, you can use the Error Checking option that runs a scan and checks for a problem with formulas used in the worksheet.

To run Error Checking in the Excel sheet, follow these steps,

- Go to Formulas and click on the ‘Error Checking’ button

![Error Checking](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/11/MS-Excel-error-checking-1024x431.png)

- This runs a scan on the sheet and displays the issues, if any. If no issue is found, it displays the following message,

_The error check is completed for the entire sheet._

In such a case, you can try saving the Excel file again. If the error message persists, proceed to the next method.

### Method 2: Check Individual Sheet

The problem may also occur due to an issue with one of the sheets in the workbook. To find the faulty sheet and fix the problem, you can copy each sheet content in a new Excel file and then try to save the Excel file.

This will help you find the faulty sheet from the workbook that you can review. This method makes the entire [process of troubleshooting Excel formula](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) reference error quite easy and convenient.

In case the error is not fixed, you can back up the faulty sheet content and remove it from the workbook to save the Excel file.

### Method 3: Check Links

When the Excel file contains external links with errors, MS Excel may display such error messages. To check and confirm if external links are causing the error, follow these steps,

- Navigate to _Data Tab > Queries & Connections > Edit Links_
- Check the links. If you find any faulty link, remove it and then save the sheet

### Method 4: Review Charts

You can review the charts to check if they are causing the formula reference error in Excel. It may take a while based on the size of the Excel file. Sometimes, it’s not practically possible to track down which Excel chart object is causing the error. Thus, you need to check specific locations, such as:

1. Check horizontal axis formula inside Select Data Source dialog box
2. Check Secondary Axis
3. Check linked Data Labels, Axis Labels, or Chart Title

### Method 5: Check Pivot Tables

To check Pivot Tables, follow these steps,

- Navigate to _PivotTable Tools > Analyze > Change Data Source > Change Data Source…_

![Edit links](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/11/MS-Excel-Edit-Links-1024x84.png)

- Check if any of the formula used is problematic. Sometimes small typo, such as misplaced comma, can lead to such problems in Excel. Thus, check each formula thoroughly and correct the formulas wherever needed.

### Method 6: Use Excel Repair Software

When none of the methods resolve the error, then you can rely on advanced [Excel repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), such as Stellar Repair for Excel. It’s a powerful tool that is recommended by several MVPs and IT administrators for resolving common Excel errors, such as ‘Excel found a problem with one or more formula references in this worksheet.’

![Stellar Repair for Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/public/image/catalog/screenshot/excel-repair/stellar-repair-for-excel-main-interface.png)

It repairs corrupt or damaged Excel (.xls/.xlsx) files, recovers Pivot tables, charts, etc., and save them in a new Excel worksheet. It helps Excel users, facing formula reference error, restore their Excel file without any risk of data loss, while preserving the sheet properties and formatting with 100% precision.

## Conclusion

Although the error ‘Excel found a problem with one or more formula references in this worksheet’ can be resolved by using various options in MS Excel, it may lead to a partial loss of information. Thus, you must perform these operations after taking a backup of the Excel worksheet. Also, if the MS Excel options fail to resolve the problem, you can use an Excel file repair software, such as Stellar Repair for Excel. The software helps fix Excel file corruption and restores the information and data from corrupt or damaged Excel files (.xls/.xlsx) to a new worksheet.




## How to Fix Excel File Couldn't Open in Protected View

**Summary:** This blog discusses about the Excel error 'The file couldn't open in Protected View' and workarounds to fix the error. If you cannot access the Excel file data, using an Excel file repair tool can help you quickly restore the data to its original state.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Protected View is a feature in Microsoft Excel that opens a file in read-only mode while disabling the editing functions. An Excel file opens in Protected View if the file is from an unsafe location. For example, the file may be opened from an insecure Internet location or downloaded as an Outlook attachment from suspicious senders. In that case, your Excel application won’t let you open the file to protect it against malware or virus attacks.

However, the Protected View feature sometimes prevents an Excel file from opening even if it is from a safe location. Also, opening an Excel file may return an error message ‘file couldn’t open in protected view’.

![the file couldn't open in Protected View excel error](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/03/file-couldnt-open-in-protected-view-error-1.png)

You may be able to resolve the issue by closing all the open Excel files, closing the Excel application, and then re-opening the problematic file. But this process can take a significant amount of time as you need to close and re-open the Excel application and the file continuously.

## **How to Fix the Excel File Couldn’t Open in Protected View Error?**

Here are some workarounds to fix the error:

### **Workaround 1 – Disable the Protected View Settings**

**Note:** Before disabling the protected view settings, ensure you have a reliable antivirus software installed on your computer to safeguard files against malware or virus attacks.

- Open the Excel file throwing the error, go to **File** and click the **Options** tab.
- From the Options dialog box, go to **Trust Center** and click on **Trust Center Settings**.
- From the Trust Center Settings dialog box, select the **Protected View** tab.
- Uncheck the appropriate option under the Protected View section, and then click **OK** to apply the changes. For instance, uncheck the “Enable Protected View for Outlook attachments” option if the file is a downloaded Outlook attachment.

![disable protected view settings ](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/03/disable-protected-view-2.png)

- Click **OK** again.

Now try opening the problematic Excel file. If it still doesn’t open, try the following workaround.

### **Workaround 2 – Repair the Office Installation**

Sometimes, corruption in the Office installation may cause Excel to return the ‘file could not open in protected view’ error. So, try repairing your Office installation and check if it solves the problem.

- Open the Control Panel and click the **Uninstall a program** option under **Programs**.
- Click **Uninstall a program** under the **Programs** tab.
- Right-click the Microsoft Office program you have installed and click the **Change** button.
- Select the **Repair** option and click **Continue**.

![repair office program](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/03/online-repair-microsoft-office-app-3-1024x387.png)

Once the repair operation is complete, check if you can open the Excel file without any issue.

### **Workaround 3 – Install Latest MS Office Updates**

Several Excel users have reported about experiencing [issues](https://answers.microsoft.com/en-us/msoffice/forum/all/cannot-open-some-office-files-after-upgrading-from/c103ee49-2e87-4421-ad96-f433ec77ec54) (such as the Protected View issue) on opening an Office document like Excel, Word, or others – after upgrading Windows to a newer version. This problem may occur due to incompatibility between the Office app and the updated Windows version. You can try to fix the problem by upgrading the Office app, and here’s how:

- Click on the **File** tab in the Excel file that throws an error.
- Click the **Account** tab at the left-hand bottom of the screen.
- From Office Updates, expand **Update Options**, and then click **Update Now**.

![office updates](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/03/update-excel-4.png)

After installing the updates, restart your PC and open the Excel file. If this doesn’t work, use the following workaround.

## **Workaround 4 – Convert the File and Rename It**

Some Excel users resolved the ‘Protected view’ issue by renaming their Excel files. But before that, ensure that the file you are trying to open is saved with .xlsx format. If the file type is .xls, you need to convert it to .xlsx format and then rename it. Before converting the file, save a copy of the original file in the desired location.

If the file type is .xlsx, copy the problematic file. Next, open a blank Excel document and paste the copied file. The file will appear as an icon, double-click on it. If it opens without any error, save it with a different name. Hopefully, reaming the Excel file may help resolve the problem.

## **Conclusion**

There is no fixed solution to resolve the Excel error ‘The file could not open in Protected View’. However, the workarounds discussed in the blog have helped users resolve the ‘Protected View’ issue. Make sure to use the workarounds in the same order as explained in the blog.

If you’re having an issue loading the Excel file data due to the ‘Protected View’ error, consider restoring the data using an [Excel file repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/03/free-download-windows-2.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)




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
