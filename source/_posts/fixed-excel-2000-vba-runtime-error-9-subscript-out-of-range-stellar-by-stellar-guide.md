---
title: Fixed Excel 2000 VBA Runtime Error 9 Subscript Out of Range | Stellar
date: 2024-03-12 13:36:28
updated: 2024-03-14 16:13:10
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2000 VBA Runtime Error 9 Subscript Out of Range
excerpt: This article describes Fixed Excel 2000 VBA Runtime Error 9 Subscript Out of Range
keywords: repair corrupt .xlsm,repair damaged .xlsm,repair excel 2013,repair .xlb,repair excel file,repair corrupt .xlsx files,repair damaged .xls,repair damaged .xltx files,repair excel 2023,repair damaged .xlsx files,repair corrupt .xltm files,repair .xlsx
thumbnail: https://www.lifewire.com/thmb/td12YpTj-yP7iyxYOnjHj9dHoA4=/400x300/filters:no_upscale():max_bytes(150000):strip_icc():format(webp)/streamfreesportsonline-5b1994178023b900364a81e5.jpg
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


## Get Rid of corrupt Excel File

**Summary:** What to do when an Excel file is corrupted? This is a common question that is often asked by Microsoft Excel users. If you too are seeking an answer to this question, read the blog to learn about a few manual workarounds and a specialized Excel file repair tool to resolve the Excel file corruption issue.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

An Excel file gets corrupted due to various reasons such as a virus/malware attack, sudden system shutdown when the Excel file is still open, power failure while working with an Excel spreadsheet, etc.

When Microsoft Excel detects corruption in a workbook, it attempts to repair the workbook by starting _‘File Recovery mode’._

**Tip!** If the file recovery mode doesn’t start, you may use the manual repair process or an Excel repair tool, such as Stellar Repair for Excel to repair a corrupted Excel file. The software can help you quickly retrieve contents from a damaged, corrupt, or inaccessible Excel file and restore the file to its original state.

There even exist a few manual tips that can be used to recover data from damaged MS Office Excel files.

## **Workarounds to Use When an Excel File is Corrupted**

**_Note:_** _Before carrying out any of the repair and recovery workarounds, it is advised that you must save a backup copy of the damaged file. This is to prevent your files from turning completely inaccessible in case the methods fail to give desired results._

### **Workaround 1: Use the Open and Repair Method**

If MS Excel cannot repair a corrupted workbook automatically, you can try to do it manually. To do so, perform the following:

- Open the corrupt file, like you normally open any file, by clicking **File** > **Open**.
- Browse and locate the folder containing the corrupted document.
- When the Open dialog box is displayed:
- Select the Excel document.
- Click on the arrow present to the right side of the Open button and select **Open and Repair** option.

![open and repair excel file](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2020/07/open-and-repair.jpg)

Figure 1 – Open and Repair Feature

If this doesn’t help repair the broken Excel file or you encounter  [Open and Repair does not work](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) issue, proceed with the next workaround.

**_Tip!_** _Try an alternative solution, i.e._ **_Stellar Repair for Excel software_** _to repair and recover corrupt Excel files (.xlsx or .xls) when the ‘Open and Repair’ method won’t work._  

### **Workaround 2: Restore an Excel File with a Shadow Copy**

If you’re a **Windows 7 or Vista user,** you can try restoring the corrupted spreadsheet by using a shadow copy (or a previous version). [Shadow copy](<https://en.wikipedia.org/wiki/Shadow_Copy#:~:text=Shadow%20Copy%20(also%20known%20as,the%20Volume%20Shadow%20Copy%20service>.) is basically a snapshot (backup copy) of computer files or volumes. The snapshot may contain an older version of your Excel file that has become damaged now. To find out, do the following:

- **Launch File Explorer**, and right-click the folder in which the **file is saved.**
- Choose **Properties.**
- Look for and click the **Previous Versions** tab. This will display a list of entries under **Folder versions** or **File versions**, going back a few days or weeks.
- Double-click one with a date when the file was accessible and could be read. Then, try to open its older version. If it opens, save the older version with a new name and execute the procedure with new file/folder entries.

![Excel file is corrupted](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/09/Shadow-copy.png)

Figure 2 – Volume Shadow Copy

You would have to repeat the process until you reach the point where the file became damaged. With this, you will get a baseline version of the file, but data may still have been lost.

### **Workaround 3: Test your Assumptions**

If you receive a message saying “[**Excel file corrupted and cannot be opened**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)”, you would probably believe it. However, there could be other reasons besides corruption that may cause Excel to throw this error message.

Your Office suite, which Excel is a part of, maybe having some primary issues in it causing problems while opening one Excel document. So, try opening another Excel file to check if the problem exists with all the files or just one.

If other Excel documents work correctly, it means that only the particular document is corrupt. On the contrary, if the issue is with your Office suite, repairing the current Office installation may help fix the issue. For this, perform these steps:

- Go to **Control Panel and click Uninstall** **the Program**.
- Choose **Office.**
- Click Change, and hit the **Repair button.**

You can **reinstall** the entire Office package. Once reinstalled, try to open the file to check if the issue has been fixed and the **Excel file repaired.**

![Excel file is corrupted](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2017/09/Repair-MS-Office.jpg)

Figure 3 – MS Office Repair

### **Workaround 4: Use Excel File Repair Tool**

If the above manual solutions fail, use [Excel repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) to successfully repair your damaged Excel workbook and recover all its data. Essentially, the software rebuilds damaged Excel workbook data at a granular level to recover every single object & all the original properties of the workbook.

**Suggested Read:** [**How to repair corrupt Excel files using Stellar Repair for Excel?**](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **Why Use Stellar Repair for Excel Software?**

- Repairs severely corrupted XLSX and XLS files.
- Can handle corrupt Excel files of any size.
- Demo version allows previewing recoverable Excel file items for free.
- **Supports Microsoft Excel 2019** and all lower versions.
- Compatible with **Windows 10 and lower versions.**
- Tested and recommended by **Microsoft Excel MVPs.**

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/01/Free-download-for-windows-1.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **Final Word**

When an Excel file is corrupted, it won’t open at all or you won’t be able to access all the file data. Such a situation can lead to unnecessary halts, impacting work productivity.

There are manual workarounds that may help fix the corrupt Excel file and recover its data, such as the ones covered in this blog. However, these solutions might not work in severe corruption cases and may require technical assistance. Also, they may result in some data loss.

To overcome the limitations of manual workarounds, it is recommended to go for a professional Excel file repair tool such as [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It helps **repair corrupt Excel** (XLS or XLSX) files and restores all worksheet data, such as the table, chart, chart sheet, cell comment, sort and filter, image, formula, etc. in a few simple clicks. Moreover, the software provides a free preview of the recoverable data with its demo version. You can check the preview to evaluate how the software works.


## 'Unable to Save Excel Workbook' Issue [Fix 2024]

**Summary:** You may unable to save your Excel Workbooks due to several reasons. Many users have reported this issue on the Tech Forums. This blog will discuss a few instances when users cannot save their Excel files. It lists the causes behind the issue and their possible solutions. It also mentions the Stellar Repair for Excel to fix the saving error if it is due to corruption in the Excel file.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

It is easy to work with Microsoft Excel but sometimes, the application may create issues thereby hampering the smooth functioning of the workbook. One such issue is “unable to Save Excel Workbook”.

## Let’s take a look at the issue of Unable to Save Excel Workbook

### **Instance 1:**

In an organization, users connected to one of the servers (Windows 2008 R2) using Citrix – a Terminal Server configured with Windows 2008 R2 –and accessed their data through a File Server, also configured with Windows 2008R2. Since the connectivity to Shared Drive was established through a Terminal server, any conflict amongst the server configuration may create conflict in shared file.

This issue was discussed at length at one of the [Tech Forums](https://community.spiceworks.com/topic/371563-excel-document-not-saved-issue), where the users were unable to access their workbooks stored on the shared drive. The File menu did not work. As a result, the users were forced to save the workbook by creating quick access shortcuts or locally on the desktop. In many cases, the saving option was ruled out completely.

### **Instance 2:**

A [similar problem](https://www.dell.com/support/article/in/en/indhs1/sln308103/unable-to-save-workbook-in-microsoft-excel?lang=en) was reported, wherein the users received an error when saving an Excel workbook after inserting a chart in an existing workbook (previously saved) or copying values from an existing workbook. A system is configured with Windows 7 and Microsoft Office 10 configuration. The issue arises when the user is unable to save the changes after editing in a saved spreadsheet. The following message displays on the screen:

![Image of Error message while trying to save excel file](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/05/Excel-Error.png)

Figure: Unable to Save Excel WorkBook Issue

**Further, if the user clicks ‘Continue’, the following error message is received:**

“Excel encountered errors during save. However, Excel was able to minimally save your file to <**filename.xlsx**\>”.

**Note:** This issue impacts build Version 1707 (Build 8326.2086) and later, and also only occurs with files that are stored locally, such as on the desktop. This problem does not occur if you manually enter values or insert a chart in a newly created workbook.

## Plausible reasons for the ‘Unable to save Excel workbook’ Issue

1. The issue was detected in Microsoft Office Professional Plus 2010 32-bit, Service Pack 14.0.6029.1000.
2. Excel version on the user system may or may not match with Excel version on File server.
3. The issue of ‘Unable to Save Excel Workbook’ impacts only the Build Version 1707 (Build 8326.2086) and later.
4. In case of Issue 2, the problem surfaces when the user adds files, tables or charts in the locally saved excel files, such as on the desktop.

## Methods to fix the ‘Unable to Save Excel Workbook’ Issue

There may be an issue with the Build version or the Registry Values settings may not be appropriate, which does not allow the Excel workbooks to save.

**But, before starting to resolve the issue, verify the following:**

1. **The location where the file is to be saved may not have enough space to save the Excel file**: Check the available space and save again. You may also use the option of ‘Save As’ to save the file at a new location.
2. **Excel file may be a shared one where edits are not allowed by a specific user**: There are restrictions attached to documents and other files shared over the network. Check for these restrictions.
3. **Antivirus may interrupt in during file saving**: Antivirus in the system may not allow saving of the files. Request the system administrator to uninstall the antivirus and reinstall after saving.
4. **The file is not saved within 218 characters**: If the file is not saved due to the naming issue, then check the character length and try again.
5. **Differences in Windows versions** of the local system and those on network drive may cause excel not saved issues. Check that all the systems have the same configuration and are updated to the recently available versions.
6. **Excel spreadsheet is corrupt**: If none of the above factors have not caused hindrance in saving the file, then there may be a probability of [corruption in the Excel spreadsheet](https://www.stellarinfo.com/blog/simple-way-to-open-corrupt-excel-file-without-any-backup/).

Once verified, look for a healthy and restorable backup. If backup is missing, resolve the issue of “Unable to open Excel File” with manual settings on local system or through a reliable Excel repair software.

### Method 1: Modify Registry Entries

If multiple users are unable to access their workbooks stored on the shared drive and facing unable to save Excel file problem (see Instance 1 above), then follow the below steps:

1. Go to ‘Registry Entry’. To do this, type ‘regedit’ in the Start Search box, and press ENTER

![Image of Run window with the command 'regedit" that is to be Run](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/05/Registry-Entry.jpg)

Figure: Edit Registry

2. You are prompted for the administrator password or for a confirmation, type the password, or click Continue
3. Locate the following registry subkey, and right-click it: **HKEY\_LOCAL\_MACHINE\\System\\CurrentControlSet\\Services\\CSC**

![Image of Registry Editor window, locating the registry subkey](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/05/CSC-Location.jpg)

Figure: CSC Location

4. Point the cursor to New, and click Key

![Image of Registry Window, Right clicking on the subkey 'CSC', hover over "New" and clicking on "Key"](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/05/New-Key.jpg)

Figure: Create new key

5. Type ‘File Parameters’ in the available box

!['Type ‘File Parameters’ in the available box' ](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/05/File-Parameter.jpg)

Figure: File parameters

6. Right-click Parameters, point the cursor to New, and click DWORD (32-bit) Value

![Image of selecting DWORD (32-bit) Value under "New" by right clicking on "File Parameters" in the Registry Editor](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/05/DWORD32bit.jpg)

Figure: File parameter (DWORD – 32 bit) value

7. Type ‘FormatDatabase’, and press ‘ENTER’. Right-click ‘FormatDatabase’, and click ‘Modify’

![Image of clicking on "Modify..." by right-clicking on FormatDatabase that was entered](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/05/Modify.jpg)

Figure: Modify format database

8. In the Value data box, type ‘1’, and click ‘OK’

![Image of Value Data set as 1 in the Value data box 1 after clicking on "Modify..."](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/05/value-data.jpg)

Figure: Value data

9. Exit ‘Registry Editor’
10. Restart the system and verify if the files can be saved now

### Method 2: Try Google Uploads

If the user is unable to save the changes after editing in a locally saved spreadsheet (see Instance 2 above), then follow these steps:

1. Upload the unsaved Excel file to Google Docs. Ensure that the file gets converted to Google Sheets format.
2. Check if all the formulae are active and working.
3. Make changes to the Google Sheet and verify that all the changes are working fine.
4. Use the Google Sheets export feature to download the file in Excel format.

### Method 3: Resolve manually with Open and Repair

**If the Excel file is found to have corruption, try out the Excel Open and Repair utility:**

1. Open a blank Excel File. Go to **File** and Click **Open**.
2. Go to **Computers** and click **Browse**.
3. Access the **Location and Folder** and click the arrow icon beside **Open** followed by **Open and Repair.**

![Image of Open and Repair built-in utility in the browse window.](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/05/Open-and-Repair-1024x549.png)

Figure: Illustrates Steps to use ‘Open and Repair’ method

The Open and Repair utility is not competitive enough and may not fix corruption in severely corrupted files. Hence, if you are unable to save Excel workbook after applying the manual methods, then you can search for a useful software-based repair utility.

### Method 4: Excel File Repair Software

Specifically meant to resolve Excel file corruption. **[Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)** helps you to repair every single object including charts, tables, their formatting, shared formulae and rules and more.

1. **Install** and **Open** the software and **select** the corrupt Excel File. You can also click the **Find** option if the file location is not known.
2. Click **Scan** and allow the software to **scan and repair** the corrupt Excel file.
3. Once repaired, the software displays the fixed file components to verify its content.
4. Click **Save** to save the file data in a blank new file as **‘Recovered\_abc.xls’**, where abc.xls is the name of the original file.

See the working of the software which has been declared as a tool that provides **100% integrity and precision**.

<iframe title="How to Repair and Recover Corrupted Excel Files?" width="750" height="422" frameborder="0" allowfullscreen="" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" nitro-og-src="https://www.youtube.com/embed/3SiJqmP7iMU?feature=oembed&amp;autoplay=1" nitro-lazy-src="data:text/html;https://www.youtube.com/embed/3SiJqmP7iMU?feature=oembed&amp;autoplay=1;base64,PGJvZHkgc3R5bGU9J3dpZHRoOjEwMCU7aGVpZ2h0OjEwMCU7bWFyZ2luOjA7cGFkZGluZzowO2JhY2tncm91bmQ6dXJsKGh0dHBzOi8vaW1nLnlvdXR1YmUuY29tL3ZpLzNTaUpxbVA3aU1VLzAuanBnKSBjZW50ZXIvMTAwJSBuby1yZXBlYXQnPjxzdHlsZT5ib2R5ey0tYnRuQmFja2dyb3VuZDpyZ2JhKDAsMCwwLC42NSk7fWJvZHk6aG92ZXJ7LS1idG5CYWNrZ3JvdW5kOnJnYmEoMCwwLDApO2N1cnNvcjpwb2ludGVyO30jcGxheUJ0bntkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2p1c3RpZnktY29udGVudDpjZW50ZXI7Y2xlYXI6Ym90aDt3aWR0aDoxMDBweDtoZWlnaHQ6NzBweDtsaW5lLWhlaWdodDo3MHB4O2ZvbnQtc2l6ZTo0NXB4O2JhY2tncm91bmQ6dmFyKC0tYnRuQmFja2dyb3VuZCk7dGV4dC1hbGlnbjpjZW50ZXI7Y29sb3I6I2ZmZjtib3JkZXItcmFkaXVzOjE4cHg7dmVydGljYWwtYWxpZ246bWlkZGxlO3Bvc2l0aW9uOmFic29sdXRlO3RvcDo1MCU7bGVmdDo1MCU7bWFyZ2luLWxlZnQ6LTUwcHg7bWFyZ2luLXRvcDotMzVweH0jcGxheUFycm93e3dpZHRoOjA7aGVpZ2h0OjA7Ym9yZGVyLXRvcDoxNXB4IHNvbGlkIHRyYW5zcGFyZW50O2JvcmRlci1ib3R0b206MTVweCBzb2xpZCB0cmFuc3BhcmVudDtib3JkZXItbGVmdDoyNXB4IHNvbGlkICNmZmY7fTwvc3R5bGU+PGRpdiBpZD0ncGxheUJ0bic+PGRpdiBpZD0ncGxheUFycm93Jz48L2Rpdj48L2Rpdj48c2NyaXB0PmRvY3VtZW50LmJvZHkuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBmdW5jdGlvbigpe3dpbmRvdy5wYXJlbnQucG9zdE1lc3NhZ2Uoe2FjdGlvbjogJ3BsYXlCdG5DbGlja2VkJ30sICcqJyk7fSk7PC9zY3JpcHQ+PC9ib2R5Pg=="></iframe>

The **Excel repair software** takes care to save the repaired data in a new file to minimize the chances of further corruption.

## **Conclusion**

‘Unable to save Excel file’ is a generic problem that may appear due to various reasons. In this blog post, we presented some of the actual instances reported by users on community forums.

Windows updates, the Build versions, the Service Packs of the local systems and those on the network drive must be either similar or in sync with each other. Any deviation may cause issues in accessing or saving the Microsoft files, as reported in Instance 1 is caused where user is unable to save Microsoft Excel file on the Network Drive. In case, the user is unable to save the file on network drive then the problem lies with the Registry value.

Another case is when the users receive an error while saving an Excel workbook after they insert a chart in an existing workbook or copying values from an existing workbook. This issue is known to affect build Version 1707 (Build 8326.2086) and later, and only occurs with locally stored files.

When a user is unable to save a specific Excel file, then the problem can be resolved using the manual methods or the software based utility. The mode of repair depends upon the level of corruption in Excel file.

Hence, it is suggested to analyze the nature of the problem and decide an appropriate resolution method.




## Ways to Fix the “Failed to Parse the Corrupted Excel File” Error

**Summary:** While parsing an Excel file, you may experience the “Failed to parse the corrupted Excel file” error. It usually occurs if the Excel file is corrupt. This blog covers some easy fixes to fix this Excel error. It also mentions an advanced Excel repair tool that can help fix the issue by repairing the corrupted file.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

You may encounter the “Failed to parse the corrupted Excel file” error when parsing (extracting or converting) the data in an Excel file. This error usually occurs if the Excel file is corrupted – either completely or partially (some of its objects, like formulas, macros, and values are corrupt). However, there are various other reasons that may lead to this Excel error. Let’s know the possible causes behind the “Failed to parse the corrupted Excel file” error and how to fix it.

## **Causes of the “Failed to parse the corrupted Excel file” Error**

This error could occur due to any of the following causes:

- Damaged or corrupted macros
- Excel document is closed suddenly
- Issues with Excel application installation
- Virus or malware attack
- Hardware failure
- Bad sectors on the hard drive

## **Ways to Fix the “Failed to parse the corrupted Excel file” Error**

Here are some possible ways to fix this error and recover the Excel file.

### **Method 1: Try to Save the File in a Different Format**

Sometimes, the error can occur if the file format is not compatible with your Excel version. You can try saving the file to a different format. To do this, follow the below steps:

- Open the Excel file.
- Click on the **File** option and then select **Save As**.
- Click on Browse option and then click **Save as Type**.

![Save As Window](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/05/save-as-window.jpg)

- Select the desired Excel file format and click **Save**.

### **Method 2: Revert to the Previously Saved Version of the Excel File**

You can try reverting to the previously saved version of the Excel file. Here’s how:

- Open your Excel application.
- Navigate to **File > Info > Manage Workbook**.

![Go To Manage Workbooks](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/05/go-to-manage-workbooks.jpg)

- Click **Recover unsaved workbooks**.

![Click Recover Unsaved Workbooks](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/05/click-recover-unsaved-workbooks.jpg)

- Click on the **Restore** option you see at the top of the file.

### **Method 3: Set the Calculation Option to Manual**

When the calculation mode is set to automatic, Excel automatically recalculates all the formulas in the file even if you make a minor change. It can take time to load the file and stop you from parsing data in your Excel file. You can change the calculation option to manual so that Excel only recalculates when you explicitly tell the application to do it. Here are the steps:

- Open your Excel application.
- Navigate to **File** and then click **Options**.
- In **Excel Options**, select **Formulas**.
- Under **Workbook Calculation**, select **Manual** and click **OK.**

![Select Manual On Excel Options](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/05/select-manual-on-excel-options.jpg)

### **Method 4: Use Open and Repair Utility**

You can try repairing the corrupted Excel file with Excel’s inbuilt tool – [Open and Repair](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). To use the tool, follow the steps cited below:

- In Excel, go to the **File** tab and then click **Open**.
- Click **Browse** to select the corrupted file.
- The Open dialog box will appear. Click on the corrupted file.
- Click on the arrow next to the **Open** button and then click **Open and Repair**.
- You will see a dialog box with three buttons **Repair, Extract Data, and Cancel**.

![Click Repair Button](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/05/click-repair-button.jpg)

- Click on the **Repair** button to recover as much of your work as possible.
- After repair, a message will appear (as shown in the below figure).

![Message After Repair](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/05/message-after-repair.jpg)

- Click **Close**.

## **What If None of the Above Methods Works?**

The above methods may fail to work if the Excel file is severely damaged. In such a case, we recommend using an efficient [Excel repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), such as Stellar Repair for Excel. It can quickly repair corrupted Excel (.XLS/.XLSX) files. It can recover all the objects from the file, including charts, tables, formulas, etc. You can download the free trial version of the tool to preview the recoverable data.

## **Conclusion**

Above, we have discussed some tried and tested methods to fix the “Failed to parse the corrupted Excel file” error. If the Excel file is severely damaged or corrupted, you can try repairing it using a third-party Excel repair software, like [Stellar Repair for Excel](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/). It can repair multiple Excel files without affecting the original formatting. The tool is compatible with all Excel versions.


## [Fix] Excel formula not showing result

**Summary:** Is your Excel spreadsheet showing text of a formula you’ve entered and not its result? This blog explains the possible reasons behind such an issue. Also, it describes solutions to fix the ‘Excel formula not showing result’ error. You can try Stellar Repair for Excel software to recover engineering and shared formulas.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Sometimes, when you type a formula in a cell of worksheet and press Enter, instead of showing the calculated result, it returns the formula as text. For instance, Excel cell shows:

![Excel not Showing Formula](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/02/Excel-not-showing-formula-1.jpg)

But you should get the result as:

![Excel Formula Working Sample](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2021/02/Excel-not-showing-formula-2.jpg)

## **Why Does Excel Show or Display the Formula Not the Result?**

**Following are the possible reasons that may lead to the ‘Excel showing formula not result’ issue:**

1. You accidentally enabled “**Show Formulas**” in Excel.
2. The cell format in a spreadsheet is set to text.
3. ‘Automatic calculation’ feature in Excel is set to manual.
4. Excel thinks your formula is text (Syntax are not followed).
5. You type numbers in a cell with unnecessary formatting.

## **How to Fix ‘Excel Showing Formula Not Result’ Issue?**

### **Solution 1 – Disable Show Formulas**

If only the formula shows in Excel not result, check if you have accidentally or intentionally enabled ‘show formula’ feature of Excel. Instead of applying calculations and then showing results, this feature displays the actual text written by you.

You can use the ‘Show Formulas’ feature to quickly view all formulas, but if you are not aware of this feature, and enabled it accidentally, it can be a headache. To disable this mode, go to ‘**Formulas**’ and click on ‘**Show formula enabled**.’ If it’s previously enabled, it will be disabled by just clicking on it.

![Show Formula Enabled](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/06/show-formulas-1.jpg)

### **Solution 2 – Cell Format Set to Text**

Another possible reason that only formula shows in Excel not result could be that the cell format is set to text. This means that anything written in any format in that cell will be treated as regular text. If so, change the format to General or any other. To get Excel to recognize the change in the format, you may need to enter cell edit mode by clicking into the formula bar or just press F2.

![Enter Cell Edit Mode by Clicking into the Formula Bar](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/06/formula-f2.jpg)

### **Solution 3 – Change Calculation Options from ‘Manual’ to ‘Automatic’**

There is an “automatic calculation” feature in Excel, which tells Excel to do calculations automatically or manually. If ‘Excel formula is not showing results’, it may be because the automatic calculations feature is set to manual. This issue is not easily detected because it results in calculating formula in one cell but if you copy it to some other cell, it will retain the first calculation and will not recalculate on the base of the new location. To fix this, follow these steps:

- In Excel, click on the ‘**File**’ tab on the top left corner of the screen.
- In the window that opens, click on ‘**Options**’ from the left menu bar.
- From ‘Excel Options’ dialog box, select ‘**Formulas**’ from the left side menu and then change the ‘**Calculation options**’ to ‘**Automatic**’ if it’s currently set as ‘Manual’.

![Automatic Calculations Feature](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2019/06/calculation-options.jpg)

- Click on ‘**OK**’. This will redirect you to your sheet.

### **Solution 4 – Type Formula in the Right Format**

There is a proper way to tell Excel that your text is a formula. If you don’t write the formula in a particular format, Excel considers it as simple text and hence no calculations are performed according to it. For this reason, keep the following in mind when typing a formula:

- **Equal sign:** Every formula in Excel should start with an equal sign (=). If you miss it, Excel will mistake your formula as regular text.

- **Space before equal sign:** You are not supposed to enter any space before equal sign. Maybe a single space will be hard for us to detect, but it breaks the rule of writing formulas for Excel.

- **Formula wrapped in quotes:** You need to make sure that your formula is not wrapped in quotes. People usually make this mistake of writing a formula in quotes, but in Excel, quotes are used to signify text. So your formula won’t be evaluated. But you can add quotes inside formula if required, for example: =SUMIFS(F5:F9,G5:G9,”>30″).

- **Match all parentheses in a formula:** Arguments of Excel functions are entered in parenthesis. In complex cases, you may need to enter more sets of parenthesis. If those parentheses are not paired/closed properly, Excel may not be able to evaluate the entered formula.

- **Nesting limit:** If you are nesting two or more Excel functions into each other, for example using nested IF loop, remember the following rules:
  - Excel 2019, 2016, 2013, 2010, and 2007 versions only allow to use up to 64 nested functions.
  - Excel 2003 and lower versions only allow up to 7 nested functions.

### **Solution 5 – Enter Numbers without any Formatting**

When you use a number in the formula, make sure you don’t enter any decimal separator or currency sign, e.g. $, etc. In an Excel formula, a comma is used to separate arguments of a function and a dollar sign makes an absolute cell reference. Most of these special characters have built-in functions so avoid using them unnecessarily.

## **What to Do If the Manual Solutions Don’t Work?**

If you’ve tried out the manual solutions mentioned above but still unable to resolve the ‘Excel formula not showing result’ issue, you can try repairing your Excel file with the help of an automated [Excel repair software](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/), such as **Stellar Repair for Excel**.

This reliable and competent software scans and repairs Excel files (.XLSX and .XLS). It also helps recover all the file components, like formulas, cell formatting, etc. Armed with an interactive GUI, this software is extremely easy to work with, and its advanced algorithms allow it to fend off Excel errors with ease.

[![Free Download for windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/01/Free-download-for-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)

## **Conclusion**

This blog outlined the possible reasons that may cause ‘Excel not showing formula results’ issue. Check out these reasons and implement the manual fixes, depending on what resulted in the problem in the first place. If none of these fixes help resolve the issue, corruption in the Excel file might be preventing the formulas from showing the actual results. In that case, using **Stellar Repair for Excel** tool might help.


## How to Fix Excel Formulas Not Working Properly | Step-by-Step Guide

**Summary:** Excel formulas sometimes fail to function correctly and even return an error. This article explains what you might be doing wrong that prevents Excel formulas from working properly and solutions to resolve the issue. If your formulas have disappeared from the Excel spreadsheet and you are having trouble recovering them, you can use an Excel repair tool to recover the formulas.

[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

When working with Excel formulas, situations may arise when the formula doesn’t calculate or update automatically. Or, you may receive errors by clicking on a formula.

## Problems Causing the ‘Excel Formulas not Working Properly’ Issue and Solutions

Let’s check out the possible reasons that cause Excel formulas to work properly and solutions to resolve the issue.

### Problem 1 – Switching Automatic to Manual Calculation Mode

Automatic and manual are the two modes of calculation in Microsoft Excel.

By default, Excel is set to automatic calculation mode. Everything is recalculated automatically when any changes are made in a worksheet in this mode. You may switch from automatic to manual mode to disable the recalculation of formulas, particularly when working with a large Excel file with too many formulas.

Excel will not calculate automatically when set to manual calculation mode. And this may make you think that the Excel formula is not working properly.

### Solution – Change Calculation Mode from Manual to Automatic

To do so, perform these steps:

- Click on the column with problematic formulas.
- Go to the **Formulas** tab, click the **Calculation Options** drop-down, and select **Automatic**.

![Automatic to Manual Calculation Mode](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/switch-manual-to-automatic-calculation-1.png)

### Problem 2 – Missing or Mismatched Parentheses

It’s easy to miss or incorrectly place parentheses or include extra parentheses in a complex formula. If a parenthesis is missing or mismatched and you click Enter after entering a formula, Excel displays a message window suggesting to fix the issue (refer to the screenshot below).

![Missing or Mismatched Parentheses](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/missing-or-mismatched-parantheses-2.png)

Clicking ‘Yes’ might help fix the issue. But Excel might not fix the parentheses properly, as it tends to add the missing parentheses at the end of a formula which won’t always be the case.

### Solution – Check for Visual Cues When Typing or Editing a Formula with Parentheses

When typing a formula or editing one, Excel provides visual cues to determine if there’s an issue with the parentheses inserted in a formula. Checking for these visual cues can help you fix missing/mismatched parentheses.

- Excel helps identify parenthesis pairs by highlighting them in different colors. For instance, the pair of parenthesis outside is black.
- Excel does not make the opening parentheses bold. So, if you’ve inserted the last closing parentheses in a formula, you can determine if your parentheses are mismatched.
- Excel helps identify parentheses pairs by highlighting and formatting them with the same color once you cross over them.

### Problem 3 – Formatting Cells in an Excel Formula

When adding a number in an Excel formula, don’t add any decimal separator or special characters like $ or €. You may use a comma to separate a function’s argument in an Excel formula or use a currency sign like $ or € as part of cell references. Formatting the numbers may prevent the formula from functioning correctly.

### Solution – Use Format Cells Option for Formatting

Use Format Cells instead of using a comma or currency signs for formatting a number in the formula. For instance, rather than entering a value of $10,000 in your formula, insert 10000, and click the ‘Ctrl+1’ keys together to open the Format Cells dialog box.

![Format Cells Option](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/format-cells-3-1.png)

### Problem 4 – Formatting Numbers as Text

Numbers are displayed as left-aligned in a sheet in a worksheet, and text formatted numbers are right-aligned in cells. Excel considers numbers formatted as text to be text strings. Thus, it leaves those numbers out of calculations. As a result, a formula won’t work as intended. For example, in the following screenshot, you can see that the SUM formula works correctly for normal numbers. But, when the SUM formula is applied to numbers formatted as text, the formula doesn’t return the correct value.

![Cells Formatted as Text](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/cells-formatted-as-text-4.png)

Sometimes, you may also see an apostrophe in the cells or green triangles in the top-left corner of all the cells when numbers in those cells are formatted as Text.

### Solution – Do Not Format Numbers as Text

To fix the issue, do the following:

- Select the cells with numbers stored as text, right-click on them, and click Format Cells.
- From the Format Cells window, click on Number and then press OK.

![Format Cells in Excel](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2022/02/format-cells-5.png)

### Problem 5 – Double Quotes to Enclose Numbers

Avoid enclosing numbers in a formula in double-quotes, as the numbers are interpreted as a string value.

Meaning if you enter a formula like =IF(A1>B1, “1”), Excel will consider the output one as a string and not a number. So, you won’t be able to use 1’s in calculations.

### Solution – Don’t Enclose Numbers in Double Quotes

Remove any double quotes around a number in your formula unless you want that number to be treated as text. For example, you can write the formula mentioned above as “1” =IF(A1>B1, 1).

### Problem 6 – Extra Space at Beginning of the Formula

When entering a formula, you may end up adding an extra space before the equal (=) sign. You may also add an apostrophe (‘) in the formula at times. As a result, the calculation won’t be performed and may return an error. This usually happens when you use a formula copied from the web.

### Solution – Remove Extra Space from the Formula

The fix to this issue is pretty simple. You need to look for extra space before the equal sign and remove it. Also, ensure there is an additional apostrophe added in the formula.

## Other Things to Consider to Fix the ‘Excel Formulas not Working Properly’ Issue

- If your Excel formula is not showing the result as intended, see this [blog](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).
- When you refer to other worksheets with spaces or any non-alphabetical character in their names, enclose the names in ‘single quotation marks’. For example, an external 5reference to cell A2 in a sheet named Data enclose the name in single quotes: **‘Data’!A1**.
- You may see the formula instead of the result if you have accidentally clicked the ‘Show Formulas’ option. So, click on the problematic cell, click on the Formula tab, and then click Show Formulas.
- If you’re getting an error “Excel found a problem with one or more formula references in this worksheet”, find solutions to fix the error [here](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/).

## Conclusion

This blog discussed some problems you might make causing an Excel formula to stop working properly. Read about these common problems and solutions to fix them. If a problem doesn’t apply in your case, move to the next one. If you cannot retrieve formulas in your Excel sheet, using an [Excel file repair tool](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/) like Stellar Repair for Excel can help you restore all the formulas. It does so by repairing the Excel file (XLS/XLSX) and recovering all the components, including formulas.

[![Free Download for Windows](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/03/free-download-windows-2.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/)


<ins class="adsbygoogle"
     style="display:block"
     data-ad-client="ca-pub-7571918770474297"
     data-ad-slot="8358498916"
     data-ad-format="auto"
     data-full-width-responsive="true"></ins>