---
title: Fixed Excel 2016 VBA Runtime Error 9 Subscript Out of Range
date: 2024-07-17T17:15:59.810Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2016 VBA Runtime Error 9 Subscript Out of Range
keywords: repair .xlb,repair excel 2000,repair excel 2007,repair .xlsm,repair excel 2023,repair .xlsx
---

## \[Fixed\] Excel VBA Runtime Error 9: Subscript Out of Range

**Summary:** The runtime error 9 in Excel usually occurs when you use different objects in a code or the object you are trying to use is not defined. This post will discuss the reasons behind the Excel VBA error "Subscript out of Range” and the solutions to resolve the issue. It will also mention an Excel repair tool that can help fix the error if it occurs due to corruption in worksheet.

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=3851655&QTY=1&AFFILIATE=108875&CART=1"><img src="http://www.aiseesoft.com/avangate/30p/banner.jpg" border="0"></a>
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
<a href="https://unicoeye.pxf.io/c/5597632/2084399/18498" target="_top" id="2084399"><img src="//a.impactradius-go.com/display-ad/18498-2084399" border="0" alt="" width="1125" height="600"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2084399/18498" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![VBA Error Subscript Out Of Range-When Incorrect Name](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/VBA-error-subscript-out-of-range-when-incorrect-name.jpg)

When you run the above code, the Excel will throw the Subscript out of range error.

So, check the name of the worksheet and correct it. Here are the steps:

- Go to the **Design** tab in the **Developer** section.
- Double-click on the **Command** button.
- Check and modify the worksheet name (e.g. from “emp” to “emp2”).

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2087407/7443" target="_top" id="2087407"><img src="//a.impactradius-go.com/display-ad/7443-2087407" border="0" alt="" width="600" height="500"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087407/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Modified Code From emp to emp2](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/modified-code-from-emp-to-emp2.jpg)

- Now run the code.
- The content in ‘emp’ worksheet will be copied to ‘emp2’ (see below).

<!-- affiliate ads begin -->
<a href="https://tinyland.pxf.io/c/5597632/1793214/19135" target="_top" id="1793214"><img src="//a.impactradius-go.com/display-ad/19135-1793214" border="0" alt="" width="900" height="900"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1793214/19135" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Content Copied From emp to emp2](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/content-copied-from-emp-to-emp2.jpg)

<!-- affiliate ads begin -->
<a href="https://modlily.sjv.io/c/5597632/2072819/17059" target="_top" id="2072819"><img src="//a.impactradius-go.com/display-ad/17059-2072819" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2072819/17059" style="position:absolute;visibility:hidden;" border="0" />
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

<!-- affiliate ads begin -->
<a href="https://caperobbin.sjv.io/c/5597632/2006123/18460" target="_top" id="2006123"><img src="//a.impactradius-go.com/display-ad/18460-2006123" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2006123/18460" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<li><a href="https://youtube-zero.techidaily.com/ools-and-methods-how-to-resize-youtube-thumbnail/"><u>[New] [Tools & Methods] How To Resize YouTube Thumbnail</u></a></li>
<li><a href="https://remote-screen-capture.techidaily.com/updated-2024-approved-creative-constructs-top-6-minimalist-mini-mansions/"><u>[Updated] 2024 Approved  Creative Constructs  Top 6 Minimalist Mini-Mansions</u></a></li>
<li><a href="https://screen-activity-recording.techidaily.com/updated-2024-approved-full-examination-evaluating-the-gecata-analyzer/"><u>[Updated] 2024 Approved  Full Examination  Evaluating the Gecata Analyzer</u></a></li>
<li><a href="https://facebook-video-files.techidaily.com/updated-2024-approved-unlocking-the-power-of-facebook-slideshows-in-digital-marketing/"><u>[Updated] 2024 Approved  Unlocking the Power of Facebook Slideshows in Digital Marketing</u></a></li>
<li><a href="https://fox-hovers.techidaily.com/updated-from-still-to-dynamic-the-xs-action-camera-unboxed-for-2024/"><u>[Updated] From Still to Dynamic  The XS Action Camera Unboxed for 2024</u></a></li>
<li><a href="https://remote-screen-capture.techidaily.com/updated-innovating-facial-capture-the-future-of-video-recording-for-2024/"><u>[Updated] Innovating Facial Capture  The Future of Video Recording for 2024</u></a></li>
<li><a href="https://video-screen-grab.techidaily.com/2024-approved-capture-and-store-nine-excellent-windows-tools-for-animation-gifs/"><u>2024 Approved  Capture and Store  Nine Excellent Windows Tools for Animation GIFs</u></a></li>
<li><a href="https://screen-mirror.techidaily.com/3-methods-to-mirror-realme-10t-5g-to-roku-drfone-by-drfone-android/"><u>3 Methods to Mirror Realme 10T 5G to Roku | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-11-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus 11 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-11r-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus 11R Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-12-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus 12 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-12r-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus 12R Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-ace-2-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus Ace 2 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-ace-2-pro-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus Ace 2 Pro Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-ace-2v-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus Ace 2V Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-ace-3-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus Ace 3 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-nord-3-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus Nord 3 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-nord-ce-3-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus Nord CE 3 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-nord-ce-3-lite-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus Nord CE 3 Lite 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-nord-n30-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus Nord N30 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-nord-n30-se-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus Nord N30 SE Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oneplus-open-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset OnePlus Open Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oppo-a1-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Oppo A1 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oppo-a18-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Oppo A18 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oppo-a1x-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Oppo A1x 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oppo-a2-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Oppo A2 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oppo-a38-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Oppo A38 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oppo-a56s-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Oppo A56s 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oppo-a58-4g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Oppo A58 4G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oppo-a59-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Oppo A59 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oppo-a78-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Oppo A78 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-oppo-a78-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Oppo A78 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/full-tutorial-to-bypass-your-honor-magic-5-pro-face-lock-by-drfone-android/"><u>Full Tutorial to Bypass Your Honor Magic 5 Pro Face Lock?</u></a></li>
<li><a href="https://blog-min.techidaily.com/how-to-restore-missing-music-files-from-motorola-razr-40-ultra-by-fonelab-android-recover-music/"><u>How To  Restore Missing Music Files from Motorola Razr 40 Ultra</u></a></li>
<li><a href="https://location-social.techidaily.com/how-to-sharefake-location-on-whatsapp-for-nubia-red-magic-8s-proplus-drfone-by-drfone-virtual-android/"><u>How to Share/Fake Location on WhatsApp for Nubia Red Magic 8S Pro+ | Dr.fone</u></a></li>
<li><a href="https://bypass-frp.techidaily.com/in-2024-easy-guide-how-to-bypass-honor-magic5-ultimate-frp-android-10111213-by-drfone-android/"><u>In 2024, Easy Guide How To Bypass Honor Magic5 Ultimate FRP Android 10/11/12/13</u></a></li>
<li><a href="https://tiktok-videos.techidaily.com/in-2024-unveiling-the-secrets-of-effective-video-captioning-for-tiktoks/"><u>In 2024, Unveiling the Secrets of Effective Video Captioning for TikToks</u></a></li>
<li><a href="https://tiktok-clips.techidaily.com/in-2024-zen-like-jams-top-20-relaxing-country-anthems-for-stress-free-dancing-tiktok/"><u>In 2024, Zen-Like Jams  Top 20 Relaxing Country Anthems for Stress-Free Dancing (TikTok)</u></a></li>
<li><a href="https://extra-lessons.techidaily.com/maximizing-video-capacity-in-64128gb-drives/"><u>Maximizing Video Capacity in 64/128GB Drives</u></a></li>
<li><a href="https://fake-location.techidaily.com/prevent-cross-site-tracking-on-honor-play-7t-and-browser-drfone-by-drfone-virtual-android/"><u>Prevent Cross-Site Tracking on Honor Play 7T and Browser | Dr.fone</u></a></li>
</ul></div>
