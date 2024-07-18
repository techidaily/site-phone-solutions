---
title: Fixed Excel 2013 VBA Runtime Error 9 Subscript Out of Range
date: 2024-07-17T17:15:55.977Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2013 VBA Runtime Error 9 Subscript Out of Range
keywords: repair excel 2023,repair .xltm,repair .xltx,repair excel,repair .csv,repair .xlb
---

## \[Fixed\] Excel VBA Runtime Error 9: Subscript Out of Range

**Summary:** The runtime error 9 in Excel usually occurs when you use different objects in a code or the object you are trying to use is not defined. This post will discuss the reasons behind the Excel VBA error "Subscript out of Range” and the solutions to resolve the issue. It will also mention an Excel repair tool that can help fix the error if it occurs due to corruption in worksheet.

<!-- affiliate ads begin -->
<a href="https://funwhole.sjv.io/c/5597632/1702887/17189" target="_top" id="1702887"><img src="//a.impactradius-go.com/display-ad/17189-1702887" border="0" alt="" width="1000" height="1000"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1702887/17189" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
[![Free Download for Windows](https://www.stellarinfo.com/images/free-download-windows.png)](https://tools.techidaily.com/stellardata-recovery/repaire-for-excel/ "Free Download for Windows")

Many users have reported encountering the error “Subscript out of range” (runtime error 9) when using VBA code in Excel. The error often occurs when the object you are referring to in a code is not available, deleted, or not defined earlier. Sometimes, it occurs if you have declared an array in code but forgot to specify the DIM or ReDIM statement to define the length of array.

<!-- affiliate ads begin -->
<a href="https://imp.i110150.net/c/5597632/924299/11305" target="_top" id="924299"><img src="//a.impactradius-go.com/display-ad/11305-924299" border="0" alt="" width="520" height="100"/></a>
<!-- affiliate ads end -->
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
<a href="https://lightailing.sjv.io/c/5597632/1725213/17190" target="_top" id="1725213"><img src="//a.impactradius-go.com/display-ad/17190-1725213" border="0" alt="" width="1000" height="1000"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1725213/17190" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://shop.copernic.com/order/checkout.php?PRODS=41033101&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.2checkout.com/images/merchant/8d30aa96e72440759f74bd2306c1fa3d/Copernic-2023-Affiliate-728x90-Elite.png" border="0"></a>
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

<!-- affiliate ads begin -->
<a href="https://appsumo.8odi.net/c/5597632/2087389/7443" target="_top" id="2087389"><img src="//a.impactradius-go.com/display-ad/7443-2087389" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087389/7443" style="position:absolute;visibility:hidden;" border="0" />
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
<li><a href="https://some-knowledge.techidaily.com/updated-guide-to-choosing-best-free-srt-translators-online/"><u>[Updated] Guide to Choosing Best FREE SRT Translators Online</u></a></li>
<li><a href="https://facebook-record-videos.techidaily.com/updated-in-2024-boost-video-appeal-with-these-20-best-thumbnail-fonts/"><u>[Updated] In 2024, Boost Video Appeal with These 20 Best Thumbnail Fonts</u></a></li>
<li><a href="https://facebook-video-content.techidaily.com/updated-joining-friends-and-family-online-facebook-setup-guide/"><u>[Updated] Joining Friends and Family Online (Facebook Setup Guide)</u></a></li>
<li><a href="https://twitter-videos.techidaily.com/updated-snapshare-success-the-videoviral-story-for-2024/"><u>[Updated] SnapShare Success  The #VideoViral Story for 2024</u></a></li>
<li><a href="https://twitter-clips.techidaily.com/updated-twitching-online-highlighting-todays-trending-video-takes/"><u>[Updated] Twitching Online  Highlighting Today's Trending Video Takes</u></a></li>
<li><a href="https://youtube-video-recordings.techidaily.com/2024-approved-empower-your-video-creation-blending-youtube-and-imovie-for-impressive-results/"><u>2024 Approved  Empower Your Video Creation  Blending YouTube and iMovie for Impressive Results</u></a></li>
<li><a href="https://ai-editing-video.techidaily.com/2024-approved-transforming-a-slow-motion-video-to-normal-mastering-the-skill/"><u>2024 Approved Transforming a Slow-Motion Video to Normal Mastering the Skill</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-hot-40i-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Hot 40i Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-note-30-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Note 30 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-note-30-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Note 30 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-note-30-pro-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Note 30 Pro Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-note-30-vip-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Note 30 VIP Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-note-30-vip-racing-edition-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Note 30 VIP Racing Edition Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-note-30i-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Note 30i Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-smart-7-hd-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Smart 7 HD Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-smart-7-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Smart 7 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-smart-8-hd-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Smart 8 HD Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-smart-8-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Smart 8 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-smart-8-plus-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Smart 8 Plus Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-smart-8-pro-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Smart 8 Pro Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-zero-30-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Zero 30 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-infinix-zero-5g-2023-turbo-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Infinix Zero 5G 2023 Turbo Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-a05s-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel A05s Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-a60-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel A60 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-a60s-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel A60s Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-a70-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel A70 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-p40-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel P40 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-p40plus-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel P40+ Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-p55-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel P55 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-p55-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel P55 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-p55plus-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel P55+ Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-p55t-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel P55T Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-s23-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel S23 Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-itel-s23plus-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Itel S23+ Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-lava-agni-2-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Lava Agni 2 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-solutions-to-hard-reset-lava-blaze-2-5g-phone-using-pc-drfone-by-drfone-reset-android-reset-android/"><u>3 Solutions to Hard Reset Lava Blaze 2 5G Phone Using PC | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/bypassing-google-account-with-vnrom-bypass-for-xiaomi-redmi-k70-pro-by-drfone-android/"><u>Bypassing Google Account With vnROM Bypass For Xiaomi Redmi K70 Pro</u></a></li>
<li><a href="https://easy-unlock-android.techidaily.com/in-2024-how-to-fix-oem-unlock-missing-on-nubia-red-magic-8s-proplus-by-drfone-android/"><u>In 2024, How To Fix OEM Unlock Missing on Nubia Red Magic 8S Pro+?</u></a></li>
<li><a href="https://extra-resources.techidaily.com/journey-through-high-definition-with-lgs-31mu97-b-screen-review/"><u>Journey Through High Definition with LG’s 31MU97-B Screen Review</u></a></li>
<li><a href="https://youtube-videos.techidaily.com/masterful-mythos-youtubes-finest-magic-tricks/"><u>Masterful Mythos  YouTube's Finest Magic Tricks</u></a></li>
<li><a href="https://screen-video-capture.techidaily.com/professional-obs-edits-the-ultimate-top-5-guide-for-2024/"><u>Professional OBS Edits  The Ultimate Top 5 Guide for 2024</u></a></li>
<li><a href="https://some-skills.techidaily.com/the-universal-watcher-exclusive-local-and-live-streaming-for-2024/"><u>The Universal Watcher  Exclusive Local and Live Streaming for 2024</u></a></li>
<li><a href="https://mondly-stories.techidaily.com/viking-verbal-legacy-the-echoes-of-old-norse-speech/"><u>Viking Verbal Legacy: The Echoes of Old Norse Speech</u></a></li>
<li><a href="https://graphic-issues.techidaily.com/win11-nightmeow-shadow-issue-solved-with-fall-patch/"><u>Win11 Nightmeow: Shadow Issue Solved with Fall Patch</u></a></li>
</ul></div>
