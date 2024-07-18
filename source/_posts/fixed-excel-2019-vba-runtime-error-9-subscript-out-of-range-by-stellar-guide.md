---
title: Fixed Excel 2019 VBA Runtime Error 9 Subscript Out of Range
date: 2024-07-17T17:16:04.807Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2019 VBA Runtime Error 9 Subscript Out of Range
keywords: repair .xls,repair excel 2013,repair excel 2007,repair .xlsx,repair excel 2016
---

## \[Fixed\] Excel VBA Runtime Error 9: Subscript Out of Range

**Summary:** The runtime error 9 in Excel usually occurs when you use different objects in a code or the object you are trying to use is not defined. This post will discuss the reasons behind the Excel VBA error "Subscript out of Range” and the solutions to resolve the issue. It will also mention an Excel repair tool that can help fix the error if it occurs due to corruption in worksheet.

<!-- affiliate ads begin -->
<a href="https://twopages.pxf.io/c/5597632/1873313/18544" target="_top" id="1873313"><img src="//a.impactradius-go.com/display-ad/18544-1873313" border="0" alt="" width="1080" height="1263"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1873313/18544" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<a href="https://united.elfm.net/c/5597632/748964/4704" target="_top" id="748964"><img src="//a.impactradius-go.com/display-ad/4704-748964" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://united.elfm.net/i/5597632/748964/4704" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://shop.copernic.com/order/checkout.php?PRODS=41033095&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.2checkout.com/images/merchant/8d30aa96e72440759f74bd2306c1fa3d/Copernic-2023-Affiliate-728x90-Advanced-3YR.png" border="0"></a>
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

<!-- affiliate ads begin -->
<a href="https://vapordna.pxf.io/c/5597632/1494880/17238" target="_top" id="1494880"><img src="//a.impactradius-go.com/display-ad/17238-1494880" border="0" alt="" width="728" height="90"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1494880/17238" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Content Copied From emp to emp2](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/content-copied-from-emp-to-emp2.jpg)

<!-- affiliate ads begin -->
<a href="https://zonlipartnershipprogram.pxf.io/c/5597632/1611407/17882" target="_top" id="1611407"><img src="//a.impactradius-go.com/display-ad/17882-1611407" border="0" alt="" width="300" height="485"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1611407/17882" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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

<!-- affiliate ads begin -->
<iframe id="iframe_672" src="//a.impactradius-go.com/gen-ad-code/5597632/1959812/17834/" width="720" height="300" scrolling="no" frameborder="0" marginheight="0" marginwidth="0"></iframe>
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
<li><a href="https://some-techniques.techidaily.com/new-how-to-build-a-career-in-graphic-design/"><u>[New] How to Build A Career In Graphic Design</u></a></li>
<li><a href="https://tiktok-videos.techidaily.com/updated-2024-approved-the-ultimate-guide-to-stellar-tiktok-content-through-template-utilization/"><u>[Updated] 2024 Approved  The Ultimate Guide to Stellar TikTok Content Through Template Utilization</u></a></li>
<li><a href="https://extra-support.techidaily.com/updated-mastering-morphvox-transformation-top-techniques-revealed/"><u>[Updated] Mastering MorphVOX Transformation  Top Techniques Revealed</u></a></li>
<li><a href="https://facebook-video-share.techidaily.com/updated-techniques-for-perfect-chroma-key-compositing/"><u>[Updated] Techniques for Perfect Chroma Key Compositing</u></a></li>
<li><a href="https://youtube-stream.techidaily.com/2024-approved-steps-for-inspirational-daily-living-visual-stories/"><u>2024 Approved  Steps for Inspirational Daily Living Visual Stories</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-samsung-galaxy-s23-tactical-edition-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Samsung Galaxy S23 Tactical Edition Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-samsung-galaxy-s23-ultra-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Samsung Galaxy S23 Ultra Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-samsung-galaxy-s23plus-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Samsung Galaxy S23+ Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-samsung-galaxy-s24-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Samsung Galaxy S24 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-samsung-galaxy-s24-ultra-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Samsung Galaxy S24 Ultra Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-samsung-galaxy-s24plus-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Samsung Galaxy S24+ Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-samsung-galaxy-xcover-6-pro-tactical-edition-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Samsung Galaxy XCover 6 Pro Tactical Edition Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-samsung-galaxy-xcover-7-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Samsung Galaxy XCover 7 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-samsung-galaxy-z-flip-5-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Samsung Galaxy Z Flip 5 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-samsung-galaxy-z-fold-5-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Samsung Galaxy Z Fold 5 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-sony-xperia-1-v-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Sony Xperia 1 V Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-sony-xperia-10-v-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Sony Xperia 10 V Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-sony-xperia-5-v-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Sony Xperia 5 V Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-tecno-camon-20-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Tecno Camon 20 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-tecno-camon-20-premier-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Tecno Camon 20 Premier 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-tecno-camon-20-pro-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Tecno Camon 20 Pro 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-tecno-camon-30-pro-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Tecno Camon 30 Pro 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-tecno-phantom-v-flip-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Tecno Phantom V Flip Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-tecno-phantom-v-fold-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Tecno Phantom V Fold Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-tecno-pop-7-pro-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Tecno Pop 7 Pro Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-tecno-pop-8-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Tecno Pop 8 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-tecno-pova-5-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Tecno Pova 5 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://mondly-stories.techidaily.com/bridging-borders-optimal-language-learning-for-companies/"><u>Bridging Borders: Optimal Language Learning for Companies</u></a></li>
<li><a href="https://network-issues.techidaily.com/fixing-flaky-video-drivers-prevents-mc-crash/"><u>Fixing Flaky Video Drivers Prevents MC Crash</u></a></li>
<li><a href="https://fox-info.techidaily.com/in-2024-convert-subtitles-effortlessly-top-8-best-converters-from-sub-to-srt-format/"><u>In 2024, Convert Subtitles Effortlessly - Top 8 Best Converters From SUB to SRT Format</u></a></li>
<li><a href="https://facebook-videos.techidaily.com/in-2024-echovideo-capture-toolkit-for-fb/"><u>In 2024, EchoVideo Capture Toolkit for FB</u></a></li>
<li><a href="https://some-knowledge.techidaily.com/in-2024-how-to-create-an-affordable-customizable-virtual-reality-gear/"><u>In 2024, How to Create an Affordable, Customizable Virtual Reality Gear</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/in-2024-how-to-reset-a-realme-c51-phone-that-is-locked-by-drfone-android/"><u>In 2024, How to Reset a Realme C51 Phone that is Locked?</u></a></li>
<li><a href="https://tiktok-video-files.techidaily.com/jest-jamboree-exploring-the-newest-laughter-on-tiktok/"><u>Jest Jamboree  Exploring the Newest Laughter on TikTok</u></a></li>
<li><a href="https://youtube-videos.techidaily.com/making-an-impression-professional-pc-editing-strategies-for-youtube/"><u>Making an Impression  Professional PC Editing Strategies for YouTube</u></a></li>
<li><a href="https://ai-driven-video-production.techidaily.com/new-the-ultimate-guide-to-running-windows-and-mac-os-video-editing-apps-on-chrome-os-for-2024/"><u>New The Ultimate Guide to Running Windows and Mac OS Video Editing Apps on Chrome OS for 2024</u></a></li>
<li><a href="https://android-location-track.techidaily.com/top-5-tracking-apps-to-track-oneplus-ace-3-without-them-knowing-drfone-by-drfone-virtual-android/"><u>Top 5 Tracking Apps to Track OnePlus Ace 3 without Them Knowing | Dr.fone</u></a></li>
<li><a href="https://fake-location.techidaily.com/wondering-the-best-alternative-to-hola-on-poco-c50-here-is-the-answer-drfone-by-drfone-virtual-android/"><u>Wondering the Best Alternative to Hola On Poco C50? Here Is the Answer | Dr.fone</u></a></li>
</ul></div>
