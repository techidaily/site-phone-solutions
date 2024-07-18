---
title: Fixed Excel 2023 VBA Runtime Error 9 Subscript Out of Range
date: 2024-07-17T17:16:13.592Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2023 VBA Runtime Error 9 Subscript Out of Range
keywords: repair excel,repair .xltm,repair excel 2013
---

## \[Fixed\] Excel VBA Runtime Error 9: Subscript Out of Range

**Summary:** The runtime error 9 in Excel usually occurs when you use different objects in a code or the object you are trying to use is not defined. This post will discuss the reasons behind the Excel VBA error "Subscript out of Range” and the solutions to resolve the issue. It will also mention an Excel repair tool that can help fix the error if it occurs due to corruption in worksheet.

<!-- affiliate ads begin -->
<a href="https://store.iobit.com/order/checkout.php?PRODS=4596923&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/184260348236f9554fe9375772ff966e/ascscan_468X60.png" border="0"></a>
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
<a href="https://atezr.pxf.io/c/5597632/2018605/18496" target="_top" id="2018605"><img src="//a.impactradius-go.com/display-ad/18496-2018605" border="0" alt="" width="798" height="807"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2018605/18496" style="position:absolute;visibility:hidden;" border="0" />
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

<!-- affiliate ads begin -->
<a href="https://bluettieu.pxf.io/c/5597632/2042323/17091" target="_top" id="2042323"><img src="//a.impactradius-go.com/display-ad/17091-2042323" border="0" alt="BLUETTI NEW LAUNCH AC180T" width="3840" height="1600"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2042323/17091" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<a href="https://ship7com.pxf.io/c/5597632/1509856/17634" target="_top" id="1509856"><img src="//a.impactradius-go.com/display-ad/17634-1509856" border="0" alt="" width="730" height="383"/></a>
<!-- affiliate ads end -->
### **Method 3: Change Macro Security Settings**

The Runtime error 9: Subscript out of range can also occur if there is an issue with the macros or macros are disabled in the Macro Security Settings. In such a case, you can check and change the macro settings. Follow these steps:

- Open your Microsoft Excel.
- Navigate to **File > Options > Trust Center**.
- Under **Trust Center**, select **Trust Center Settings**.
- Click **Macro Settings**, select **Enable all macros**, and then click **OK**.

![Macro Settings In Trust Center](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/macro-settings-in-trust-center.jpg)

<!-- affiliate ads begin -->
<a href="https://godlikehost.sjv.io/c/5597632/1920047/21774" target="_top" id="1920047"><img src="//a.impactradius-go.com/display-ad/21774-1920047" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1920047/21774" style="position:absolute;visibility:hidden;" border="0" />
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

<!-- affiliate ads begin -->
<a href="https://caperobbin.sjv.io/c/5597632/2006123/18460" target="_top" id="2006123"><img src="//a.impactradius-go.com/display-ad/18460-2006123" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2006123/18460" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<li><a href="https://youtube-docs.techidaily.com/024-approved-stream-video-smoothly-save-pcmac-screens-to-youtube/"><u>[New] 2024 Approved  Stream Video Smoothly  Save PC/Mac Screens to YouTube</u></a></li>
<li><a href="https://youtube-video-recordings.techidaily.com/new-addressing-missing-thumbnails-in-youtube-shorts/"><u>[New] Addressing Missing Thumbnails in YouTube Shorts</u></a></li>
<li><a href="https://digital-screen-recording.techidaily.com/new-seamless-transitions-from-real-to-reel-for-2024/"><u>[New] Seamless Transitions From Real to Reel for 2024</u></a></li>
<li><a href="https://extra-guidance.techidaily.com/new-sky-high-insights-comprehensive-phantom-4-breakdown/"><u>[New] Sky-High Insights  Comprehensive Phantom 4 Breakdown</u></a></li>
<li><a href="https://screen-mirroring-recording.techidaily.com/updated-digital-dive-essential-recording-tactics-for-shows-for-2024/"><u>[Updated] Digital Dive  Essential Recording Tactics for Shows for 2024</u></a></li>
<li><a href="https://fox-cloud.techidaily.com/updated-in-2024-innovative-8th-generation-photographic-fusion-app/"><u>[Updated] In 2024, Innovative 8Th Generation Photographic Fusion App</u></a></li>
<li><a href="https://facebook-record-videos.techidaily.com/updated-speedy-tactics-for-disorganized-youtube-song-listings/"><u>[Updated] Speedy Tactics for Disorganized YouTube Song Listings</u></a></li>
<li><a href="https://some-guidance.techidaily.com/2024-approved-tailored-titles-for-your-youtube-success/"><u>2024 Approved  Tailored Titles for Your YouTube Success</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-itel-s23plus-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Itel S23+ Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-agni-2-5g-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Agni 2 5G Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-blaze-2-5g-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Blaze 2 5G Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-blaze-2-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Blaze 2 Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-blaze-2-pro-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Blaze 2 Pro Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-blaze-curve-5g-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Blaze Curve 5G Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-blaze-pro-5g-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Blaze Pro 5G Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-storm-5g-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Storm 5G Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-yuva-2-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Yuva 2 Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-yuva-2-pro-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Yuva 2 Pro Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-yuva-3-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Yuva 3 Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lava-yuva-3-pro-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lava Yuva 3 Pro Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-lenovo-thinkphone-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Lenovo ThinkPhone Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-meizu-21-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Meizu 21 Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-meizu-21-pro-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Meizu 21 Pro Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-motorola-defy-2-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Motorola Defy 2 Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-motorola-edge-2023-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Motorola Edge 2023 Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-motorola-edge-40-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Motorola Edge 40 Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-motorola-edge-40-neo-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Motorola Edge 40 Neo Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-motorola-edge-40-pro-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Motorola Edge 40 Pro Hard Reset | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/4-easy-ways-for-your-motorola-edgeplus-2023-hard-reset-drfone-by-drfone-reset-android-reset-android/"><u>4 Easy Ways for Your Motorola Edge+ (2023) Hard Reset | Dr.fone</u></a></li>
<li><a href="https://graphic-issues.techidaily.com/direct3d-initialization-mastered/"><u>Direct3D Initialization Mastered</u></a></li>
<li><a href="https://android-unlock.techidaily.com/full-guide-to-unlock-your-oppo-find-n3-flip-by-drfone-android/"><u>Full Guide to Unlock Your Oppo Find N3 Flip</u></a></li>
<li><a href="https://location-social.techidaily.com/how-to-change-gps-location-on-samsung-galaxy-s23-ultra-easily-and-safely-drfone-by-drfone-virtual-android/"><u>How to Change GPS Location on Samsung Galaxy S23 Ultra Easily & Safely | Dr.fone</u></a></li>
<li><a href="https://smart-video-creator.techidaily.com/in-2024-full-guide-on-making-a-big-budget-cinematic-look-video-with-final-cut-pro-x/"><u>In 2024, Full Guide on Making a Big-Budget Cinematic Look Video with Final Cut Pro X</u></a></li>
<li><a href="https://iphone-unlock.techidaily.com/in-2024-reset-itunes-backup-password-of-apple-iphone-6s-plus-prevention-and-solution-drfone-by-drfone-ios/"><u>In 2024, Reset iTunes Backup Password Of Apple iPhone 6s Plus Prevention & Solution | Dr.fone</u></a></li>
<li><a href="https://snapchat-videos.techidaily.com/in-2024-top-humorous-snapchat-emojis-for-record-breaking-engagement/"><u>In 2024, Top Humorous Snapchat Emojis for Record-Breaking Engagement</u></a></li>
<li><a href="https://techidaily.com/repair-damaged-unplayable-video-files-of-vivo-y100i-power-5g-on-windows-by-stellar-video-repair-mobile-video-repair/"><u>Repair damaged, unplayable video files of Vivo Y100i Power 5G on Windows</u></a></li>
<li><a href="https://snapchat-videos.techidaily.com/the-ultimate-snapshot-experience-mastering-snapchat-filters-for-2024/"><u>The Ultimate Snapshot Experience  Mastering Snapchat Filters for 2024</u></a></li>
<li><a href="https://youtube-clips.techidaily.com/uncomplicated-online-music-conversion-top-6-free-ios-apps-for-video-and-audio/"><u>Uncomplicated Online Music Conversion  Top 6 Free iOS Apps for Video & Audio</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/updated-2024-approved-top-6-celebrity-text-to-speech-ai-voice-generators-you-may-like/"><u>Updated 2024 Approved Top 6 Celebrity Text to Speech AI Voice Generators You May Like</u></a></li>
<li><a href="https://ai-voice-clone.techidaily.com/updated-in-2024-installation-and-registration/"><u>Updated In 2024, Installation and Registration</u></a></li>
</ul></div>
