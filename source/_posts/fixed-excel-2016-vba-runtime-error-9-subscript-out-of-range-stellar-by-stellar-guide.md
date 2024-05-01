---
title: Fixed Excel 2016 VBA Runtime Error 9 Subscript Out of Range | Stellar
date: 2024-03-13 12:58:41
updated: 2024-03-14 13:47:34
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2016 VBA Runtime Error 9 Subscript Out of Range
excerpt: This article describes Fixed Excel 2016 VBA Runtime Error 9 Subscript Out of Range
keywords: repair corrupt .xlsx,repair corrupt .xls files,repair .xlsm,repair damaged .xltx,repair .xltx,repair excel 2016,repair damaged .xls,repair excel 2000,repair damaged .xltm files,repair corrupt .xltm,repair damaged .xlsm files
thumbnail: https://www.lifewire.com/thmb/gOgqwLvt0rf3-WdwEBSByMeqIHo=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/GettyImages-1353420724-65161751b9924195880d3273e327cb54.jpg
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



## How to Fix a Corrupted .xls File? The Everything Guide

Undoubtedly, Excel is so powerful that it can help you to process, analysis, and store data, in masses.

That’s the reason it has been there for years and helping this world in data.

But…

With all those powers comes some nasty problems which no Excel users like to face. Can you guess what I’m talking about?

Think about a Corrupted Excel File. Nightmare? Isn’t it?

And do you remember that last time when you have opened a workbook and you got a message that this workbook is might corrupt?

The TRUTH is, this is something which you cannot avoid, but, you can [prepare yourself in the best way](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) and deal with it like a PRO.

So today, in this post, I’d like to share with you to everything you need to know about a corrupt Excel file (.xls), why it happens, how to fix it like a PRO, and much more.

...let’s get started.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/introduction.jpg)

**Note**: In this post, we’ll be covering the .xls version (which is the extension for the file which is created in Excel 2007 or the earlier versions) and if you want to know about the new version, [here’s the quick fix](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) for that.

## Why My Excel File Got Corrupted?

There can be one or multiple reasons for an Excel file to get corrupted. Below I have detailed about some of the major of them.

### 1\. Large Excel File

You can store data in a workbook the way you want but sometimes using excessive thing can make an Excel file bigger in size.

And that kind of data files can crash at any point in time. Here are a few things which make the Excel files heavy, like

- Conditional Formatting.
- Colors formatting.
- Using merged cells in place of text alignment.
- Volatile functions: Formulae that iterate every time you open or change a cell value; OFFSET, NOW.
- Using a complete column or row as a reference than the data set range.
- Using complex formulas; VLOOKUP in place of Index/Match, Nested If in place of MAXIFS, MINIFS.
- Calculations or reference across workbooks.

**Related:** [How to Fix Formatting Issues in Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

### 2\. Abrupt System Shutdown

Shutting down the system without following the procedure can corrupt your data file.

This shut down can be due to a power failure or any other unexpected technical challenges.

So it is always important to follow the procedures and shut down your system properly to avoid data losses.

### 3\. Infected Excel File (Virus Attack)

This is the most common and obvious reason for Excel file corruption.

Although we always keep our system safe using various Antiviruses, still there is always a probability of virus attacks and loss of important files.

It is always advised to use a safe and strong antivirus compatible with your system requirements.

## What are the Signs to Know When an Excel File is Corrupted?

In this section, we will discuss what are the signs which you can get when an Excel file is corrupted, let’s dig into it.

### 1\. The File is Corrupt and Cannot Be Opened

This is one of the most common messages you can see when your workbook is corrupted.

But there is also a chance that it is just because of the version compatibility where you have a .xls file but you are using the latest version of Excel check out this detailed post by Priyanka

### 2\. We Found a Problem with some Content in this File…

There’s another error message which you can get while opening a file:

We Found a Problem with some content in Do you want us to recover as much as we can? If you trust the source of this workbook, click yes.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/we-found-a-problem-with-some-content-in-this-file.jpg)

There are a lot of applications out there (I think almost every) which exports the data as a .xls format. Those files have a greater chance of having this kind of error.

### 3\. “Filename.xls” cannot be accessed

There can also be a situation where you get the error:

_“Filename.xls” cannot be accessed. The file may be corrupted, located on a server that is not responding._

Well, this message is a bit misleading.

You won't be able to decide that your file is actually corrupted or just not on the location.

## My Excel File Got Corrupted, now What Should I Do?

There are many ways to recover the data from the corrupt excel files. But before you start, it is always advised to create a copy of the corrupted file.

You can save a lot of time with [**Stellar Repair for Excel,**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) which make data recovery just with few clicks.

But before you go for a data recovery software, let's try out some manual steps which can help.

When a workbook get corrupted the first thing comes to the mind is to recover data from it…

...and you what there’s a simple option there in the Excel which you can use to do this. Below are the steps you need to follow:

- First of all, open the Excel and click on the office icon.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/office-icon.jpg)

- After that, go to the “Open” and select the file which is corrupted.

- Now, click on the open drop-down and select “Open and Repair”.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/open-repair.jpg)

- At this point, you have two options:

1. **Repair File**
2. **Extract Data**

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/extract-data.jpg)

Let’s get into both of these options one by one...

### 1\. Repair File

This option helps you to repair the file and the moment you click on it it takes a few seconds afterward and shows you the result with a message box and also provide you a log file.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/repair-file.jpg)

And once it is done with repairing, you'll get your file opened and you can save that file as a new copy.

Yes, that’s it.

### 2\. Extract Data

If somehow you aren’t able to get your file repaired, you can also extract data from that file using “Extract Data” option.

Even in this option, you can get data in two ways.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/open-repair-options.png)

1. **As Values**
2. **With Formulas**

In the first option, Excel simply extracts data as value ignoring all the formulas driving those value (which is **the best way if you just need to have that data back** ).

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/open-repai-values.jpg)

But in the second option, Excel tries to recover the formulas as much as possible.

Check out this [**smart technique by Jyoti**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) which you can use it you aren’t able to recover data from the file.

## Preventions to Not to have any Excel File Go Corrupt in Future

Future is fragile, what I’m trying to say is the more you work in Excel and process data there could be a chance that your workbook goes corrupt.

If there’s no security then what an EXCEL POWER user should do?

Well, there are few things which you can do or take care of while working with Excel so that you won’t have to worry about corruption of Excel workbooks.

Let’s see what you can do…

### 1\. Change Recalculation Option

Now here’s the thing when you work with a hell lot of data, there a common thing that you gotta using formulas. Right?

But, the thing these formulas are something which makes your Excel file slows down sometimes make them go corrupt.

There’s one small tweak you can do in your workbook is change the calculation method.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/change-recalculation-option.jpg)

Now with the manual calculation, you just need to whenever you open your file it won’t recalculate all the formulas.

And when you update your data you can simply click on the “Calculate Now” and it will calculate all the formulas again.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/calculate-now.jpg)

**Quick Tip:** Beware of Volatile Functions and use them with caution as recalculates them every time you change something in the worksheet.

### 2\. Use VBA Codes Instead of Formulas

Now, this is what I do when I need to use complex formulas in a workbook.

Here’s how you can do this: Let’s say you have a formula in the cell A1, like below, which calculates the age.

\="You age is "& DATEDIF(Date-of-Birth,TODAY(),"y") &" Year(s), "& DATEDIF(Date-of-Birth,TODAY(),"ym")& " Month(s) & "& DATEDIF(Date-of-Birth,TODAY(),"md")& " Day(s)."

Now, instead of simply entering it into the cell A1 which I would write a macro code which inserts this formula into the cell A1 and then convert it into the a value.

**Here’s the code:**

Sub CalculateAge()  
Range("B1").Value = \_  
"=""Your age is """ & \_  
"&DATEDIF(A1,TODAY(),""y"")" & \_  
"&"" Year(s), """ & \_  
"&DATEDIF(A1,TODAY(),""ym"")" & \_  
"&"" Month(s), and """ & \_  
"&DATEDIF(A1,TODAY(),""md"")" & \_  
"&"" Days(s)."""  
Range("B1") = Range("B1").Value  
End Sub

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/VBA-codes.jpg)

**Note:** To write these code you need to have basic understading of VBA (make sure [check out this guide](https://excelchamps.com/learn-vba/) for this).

### 3\. Use a File Recovery Application

Recently we asked a quick question to our readers on ExcelChamps that if they have ever faced a situation where they got a corruption message in Excel.

You’ll be astonied to hear that 50% percent of the people said “YES” they faced this thing in the past.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/yesr.jpg)

Now, this is alarming, if you are heading a team or you have a bunch of people in your company who use Excel…

…there’s a high probability that half of them gonna face this issue. So the best way to deal with this to have an App FIX your Excel file for you.

With **STELLAR REPAIR FOR EXCEL,** you just need a few clicks, yes that’s right. Let me show you with the below steps:

- First of all, download the app and install it (it’s simple).

[![download](https://www.stellarinfo.com/blog/wp-content/uploads/2017/02/free-download-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

- After that, open the app and click on the “Browse” and simply select the file which is corrupted.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/stellar-home.jpg)

- In the end, click on the REPAIR to let the Excel repair software fix your file (it takes a few seconds).

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/stellar-log-report.jpg)

Once you complete repairing your file, you’ll get a message in your on the status bar and after that, you can open your file.

## Final Thoughts

If you are a POWER Excel user then there’s a must for you to have known how to deal with a situation where you got a corrupt Excel file.

But I must recommend you to TRY OUT Stellar Repair for Excel so that’s you don’t have to worry about your Excel files anymore.

![](https://www.stellarinfo.com/image/catalog/Mac-Regional/stellar-review.jpg)

I’m sure you found this post helpful, and please don’t forget to share this tip with your colleagues, I’m sure they’ll appreciate it.


## Recover Excel Files from Virus-Infected Pen Drives for Free

**Summary:** Imagine you lost your important Excel file on which you had been working since the morning and in the next moment you realized that the file was not saved and you just lost hours of work. Wondering how to deal with this situation? Read this blog to know how Stellar free data recovery software can help you.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

From making annual reports to business growth representation figures, excel is a commonly used program for organizing data, creating pivot tables, charts etc. People from all walks of life, know the importance of Excel and the part it plays. Although it is a common file, there is a probability that you may accidentally delete excel files while working or are unable to access it due to unexpected errors. In addition, one of the major issues users face is to recover excel files from a virus infected pen drive.

Pen drives have made it possible to store and carry our important files such as excel, word document, photos, videos, etc. with us day in and day out. They just fit perfectly in our pockets and are compatible with almost every device; hence, they are widely used for transferring data from one system to another. But what if your pen drive is infected by a virus and due to it you end up losing your excel files, how will you recover your excel files for free?

A user reported that his pen drive got virus-infected and to remove the virus from it, he ran an antivirus program which removed the virus but also deleted excel files stored on it.

When your pen drive is infected by a virus, the first thing you ought to do is stop using it, even not for removing virus as an antivirus utility may remove your files as well. Further, if you have a backup, then you can recover your excel files from it, else you can use these free data recovery methods to recover your excel files.

**1\. [Free File Recovery Software Approach](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)**

Stellar Windows Data Recovery – Free Edition is an easy to use tool to recover files from a virus-infected pen drive. The software is equipped with powerful utilities to recover lost and deleted files for free. Further, it supports a wide range of file systems and is efficient in recovering files such as Excel, emails, word files, photos, audio and video files.

Using **Stellar Windows Data Recovery – Free Edition**, you can recover your files from all storage devices for free. Here’s how the software works:

- From the website, [**download**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) Stellar Windows Data Recovery – Free Edition. Connect your pen drive to your system and launch the software
- On ‘**Select What to Recover**’ screen, select file types from the given option that you wish to recover. For instance, if you want to recover photos, then under **Multimedia Files**, select ‘**Photos**’ and click on ‘**Next**’

![Stellar](https://www.stellarinfo.com/blog/wp-content/uploads/2021/02/FDR1-2-1024x721.png)

- From ‘**Select Location**’ screen, select the connected pen drive and click ‘**Scan**’

![Stellar](https://www.stellarinfo.com/blog/wp-content/uploads/2021/02/FDR2-1-1024x717.png)

- The scanning process starts and once the process is complete, software lists all the recoverable files

![Stellar](https://www.stellarinfo.com/blog/wp-content/uploads/2021/02/FDR3-2-1024x714.png)

- Select the files from the list and click on ‘**Recover**’ to save the files

**2\. Restore Excel File from the Previous Version**

If excel files are deleted from your pen drive or from your system; then you can recover them from the previous version. This feature works when Windows Backup option is enabled, else, it will not work.

Follow these steps to recover excel files:

- Connect your pen drive to your system, go to This PC and navigate to the folder of excel files
- Select the folder, right-click on it and select ‘Restore previous versions’
- From the available version of excel files, select the required one and click on ‘Restore’

**3\. Use Command Line to Recover Excel Files**

The Command prompt should be your first choice to recover excel files from the virus-infected pen drive. Here’s how command prompt recovers your files:

- Connect your virus-infected pen drive to your system and then in the search box type ‘CMD’ and hit ‘Enter’
- In the command window, type in attrib –h-r-s /s/ drive letter:\\\*.\*”, for example, “attrib -h -r -s /s /d G:\\\*.\*” and hit ‘Enter’

![attrib command](https://www.stellarinfo.com/blog/wp-content/uploads/2017/10/attrib-command.png)

- Windows starts repairing the virus-infected pen drive and once the process is complete, you can access your pen drive and recover excel files.

Even after following the above-mentioned steps you’re unable to recover your excel files, then try a Home approach i.e. a data recovery tool.

**To Sum Up**

It is always a good idea to create a backup of important files since no one can anticipate what might go wrong. The scenario presented in the blog paints a clear picture of how you can recover your Microsoft excel files for free from a virus-infected pen drive. For quick and better results, you can always go with Stellar Windows Data Recovery – Free Edition.


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

