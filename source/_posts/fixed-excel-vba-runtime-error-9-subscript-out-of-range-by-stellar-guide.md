---
title: Fixed Excel VBA Runtime Error 9 Subscript Out of Range
date: 2024-06-18T01:29:27.520Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel VBA Runtime Error 9 Subscript Out of Range
keywords: repair .xlb,repair .csv,repair excel 2007
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
<li><a href="https://phone-solutions.techidaily.com/in-2024-how-to-use-snapchat-location-spoofer-to-protect-your-privacy-on-xiaomi-13t-drfone-by-drfone-virtual-android/"><u>In 2024, How to use Snapchat Location Spoofer to Protect Your Privacy On Xiaomi 13T? | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/5-ways-to-reset-realme-gt-neo-5-without-volume-buttons-drfone-by-drfone-reset-android-reset-android/"><u>5 Ways to Reset Realme GT Neo 5 Without Volume Buttons | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-12-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus 12 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/fix-the-too-many-different-cell-formats-error-in-excel-2021-by-stellar-guide/"><u>Fix the Too many different cell formats Error in Excel 2021?</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/easy-steps-to-recover-deleted-photos-from-xperia-5-v-by-fonelab-android-recover-photos/"><u>Easy steps to recover deleted photos from Xperia 5 V.</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-ispoofer-is-not-working-on-infinix-hot-40-fixed-drfone-by-drfone-virtual-android/"><u>In 2024, iSpoofer is not working On Infinix Hot 40? Fixed | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-xiaomi-redmi-note-12-4g-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Xiaomi Redmi Note 12 4G Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-how-to-use-special-features-virtual-location-on-lava-yuva-3-drfone-by-drfone-virtual-android/"><u>In 2024, How To Use Special Features - Virtual Location On Lava Yuva 3? | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-life360-circle-everything-you-need-to-know-on-realme-c67-5g-drfone-by-drfone-virtual-android/"><u>In 2024, Life360 Circle Everything You Need to Know On Realme C67 5G | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/how-can-i-recover-corrupted-pdf-v10-file-stellar-by-stellar-guide/"><u>How Can I Recover Corrupted PDF v1.0 File | Stellar</u></a></li>
<li><a href="https://fake-location.techidaily.com/how-to-change-google-play-location-on-nokia-105-classic-drfone-by-drfone-virtual-android/"><u>How to Change Google Play Location On Nokia 105 Classic | Dr.fone</u></a></li>
<li><a href="https://youtube-clips.techidaily.com/ultimate-list-of-16-starter-videos-for-more-watchers/"><u>Ultimate List of 16 Starter Videos for More Watchers</u></a></li>
<li><a href="https://location-fake.techidaily.com/6-ways-to-change-spotify-location-on-your-nokia-c110-drfone-by-drfone-virtual-android/"><u>6 Ways to Change Spotify Location On Your Nokia C110 | Dr.fone</u></a></li>
<li><a href="https://ai-live-streaming.techidaily.com/updated-how-to-start-a-private-live-stream-on-youtube-for-2024/"><u>Updated How To Start a Private Live Stream on YouTube for 2024</u></a></li>
<li><a href="https://change-location.techidaily.com/how-to-use-life360-on-windows-pc-for-infinix-smart-8-hd-drfone-by-drfone-virtual-android/"><u>How to Use Life360 on Windows PC For Infinix Smart 8 HD? | Dr.fone</u></a></li>
<li><a href="https://youtube-videos.techidaily.com/in-2024-easy-guide-to-youtube-livestreaming-from-google-meet/"><u>In 2024, Easy Guide to YouTube Livestreaming From Google Meet</u></a></li>
<li><a href="https://extra-skills.techidaily.com/updated-simplest-way-win-clownfish-vocal-alteration-tool/"><u>[Updated] Simplest Way  Win Clownfish Vocal Alteration Tool</u></a></li>
<li><a href="https://android-transfer.techidaily.com/in-2024-how-to-transfer-contacts-from-tecno-spark-20-to-other-android-devices-devices-drfone-by-drfone-transfer-from-android-transfer-from-android/"><u>In 2024, How to Transfer Contacts from Tecno Spark 20 to Other Android Devices Devices? | Dr.fone</u></a></li>
<li><a href="https://some-techniques.techidaily.com/new-hand-tracking-insights-detailed-exploration/"><u>[New] Hand Tracking Insights  Detailed Exploration</u></a></li>
<li><a href="https://youtube-help.techidaily.com/2024-approved-inside-track-on-calculating-your-youtube-impact-and-income/"><u>2024 Approved  Inside Track on Calculating Your YouTube Impact and Income</u></a></li>
</ul></div>
