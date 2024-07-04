---
title: Fixed Excel 2007 VBA Runtime Error 9 Subscript Out of Range
date: 2024-06-16T22:45:56.626Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2007 VBA Runtime Error 9 Subscript Out of Range
keywords: repair excel 2016,repair .xls,repair excel 2019
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
<li><a href="https://phone-solutions.techidaily.com/how-to-add-a-digital-signature-block-to-jpeg-by-ldigisigner-sign-a-jpg-sign-a-jpg/"><u>How to add a digital signature block to .jpeg</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/in-2024-how-to-use-special-features-virtual-location-on-oppo-a78-drfone-by-drfone-virtual-android/"><u>In 2024, How To Use Special Features - Virtual Location On Oppo A78? | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/does-galaxy-s24-has-native-mkv-support-by-aiseesoft-video-converter-play-mkv-on-android/"><u>Does Galaxy S24 has native MKV support?</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/best-android-data-recovery-undelete-lost-music-from-motorola-razr-40-by-fonelab-android-recover-music/"><u>Best Android Data Recovery - Undelete Lost Music from Motorola Razr 40</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/how-to-delete-all-photos-from-iphone-13-beyond-scope-of-recovery-stellar-by-stellar-data-recovery-ios-iphone-data-recovery/"><u>How to Delete All Photos from iPhone 13 Beyond Scope of Recovery? | Stellar</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/easy-steps-to-recover-deleted-music-from-infinix-zero-30-5g-by-fonelab-android-recover-music/"><u>Easy steps to recover deleted music from Infinix Zero 30 5G</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-infinix-hot-40-pro-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Infinix Hot 40 Pro Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/failed-to-play-mp4-movies-with-galaxy-a54-5g-by-aiseesoft-video-converter-play-mp4-on-android/"><u>Failed to play MP4 movies with Galaxy A54 5G</u></a></li>
<li><a href="https://facebook-clips.techidaily.com/updated-in-2024-fb-song-savers-entry-level-access/"><u>[Updated] In 2024, FB Song Savers  Entry Level Access</u></a></li>
<li><a href="https://instagram-video-files.techidaily.com/new-tactics-to-determine-igtv-viewership-success/"><u>[New] Tactics to Determine IGTV Viewership Success</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/updated-translate-videos-and-subtitles-with-veedio/"><u>Updated Translate Videos and Subtitles with Veed.io</u></a></li>
<li><a href="https://audio-editing.techidaily.com/a-beginners-guide-to-selecting-premium-mac-mp3-label-editors/"><u>A Beginners Guide to Selecting Premium Mac MP3 Label Editors</u></a></li>
<li><a href="https://youtube-stream.techidaily.com/from-viewers-to-profits-an-authoritative-guide-on-youtube-revenue-and-popularity-analysis-for-2024/"><u>From Viewers to Profits  An Authoritative Guide on YouTube Revenue and Popularity Analysis for 2024</u></a></li>
<li><a href="https://android-unlock.techidaily.com/in-2024-5-solutions-for-samsung-galaxy-a14-4g-unlock-without-password-by-drfone-android/"><u>In 2024, 5 Solutions For Samsung Galaxy A14 4G Unlock Without Password</u></a></li>
<li><a href="https://facebook-video-share.techidaily.com/updated-building-a-stronger-presence-with-effective-youtube-pairings-for-2024/"><u>[Updated] Building a Stronger Presence with Effective YouTube Pairings for 2024</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/conjoin-video-streams-for-queue-curation/"><u>Conjoin Video Streams for Queue Curation</u></a></li>
<li><a href="https://extra-hints.techidaily.com/updated-80s-family-flick-the-vhs-edition-of-the-goofys/"><u>[Updated] '80S Family Flick  The VHS Edition of The Goofys</u></a></li>
</ul></div>
