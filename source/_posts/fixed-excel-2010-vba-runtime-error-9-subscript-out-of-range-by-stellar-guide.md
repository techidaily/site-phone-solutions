---
title: Fixed Excel 2010 VBA Runtime Error 9 Subscript Out of Range
date: 2024-07-17T17:15:50.919Z
tags: 
  - repair
  - repair excel
  - fix excel
categories: 
  - apps
  - windows
description: This article describes Fixed Excel 2010 VBA Runtime Error 9 Subscript Out of Range
keywords: repair excel 2013,repair excel 2023,repair excel 2007,repair .xls
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

<!-- affiliate ads begin -->
<a href="https://secure.2checkout.com/order/checkout.php?PRODS=4940317&QTY=1&AFFILIATE=108875&CART=1"><img src="https://secure.avangate.com/images/merchant/333ac5d90817d69113471fbb6e531bee/sps-partnership-728x90eng.png" border="0"></a>
<!-- affiliate ads end -->
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
<a href="https://bluettius.sjv.io/c/5597632/2027209/17108" target="_top" id="2027209"><img src="//a.impactradius-go.com/display-ad/17108-2027209" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2027209/17108" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://appsumo.8odi.net/c/5597632/2087484/7443" target="_top" id="2087484"><img src="//a.impactradius-go.com/display-ad/7443-2087484" border="0" alt="" width="1200" height="600"/></a><img height="0" width="0" src="https://appsumo.8odi.net/i/5597632/2087484/7443" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
![Content Copied From emp to emp2](https://cdn-cmlep.nitrocdn.com/DLSjJVyzoVcUgUSBlgyEUoGMDKLbWXQr/assets/images/optimized/rev-2658c43/www.stellarinfo.com/blog/wp-content/uploads/2023/04/content-copied-from-emp-to-emp2.jpg)

### Method 2: Check the Range of the Array

The VBA error “Subscript out of range” also occurs if you have declared an array in a code but didn’t specify the number of elements. For example – If you have declared an array and forgot to declare the array variable with elements, you will get the error (see below):

<!-- affiliate ads begin -->
<a href="https://modlily.sjv.io/c/5597632/2072819/17059" target="_top" id="2072819"><img src="//a.impactradius-go.com/display-ad/17059-2072819" border="0" alt="" width="300" height="250"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/2072819/17059" style="position:absolute;visibility:hidden;" border="0" />
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
<a href="https://parisrhonecom.sjv.io/c/5597632/1922358/21553" target="_top" id="1922358"><img src="//a.impactradius-go.com/display-ad/21553-1922358" border="0" alt="" width="1080" height="1080"/></a><img height="0" width="0" src="https://imp.pxf.io/i/5597632/1922358/21553" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
### **Method 3: Change Macro Security Settings**

The Runtime error 9: Subscript out of range can also occur if there is an issue with the macros or macros are disabled in the Macro Security Settings. In such a case, you can check and change the macro settings. Follow these steps:

- Open your Microsoft Excel.
- Navigate to **File > Options > Trust Center**.
- Under **Trust Center**, select **Trust Center Settings**.
- Click **Macro Settings**, select **Enable all macros**, and then click **OK**.

<!-- affiliate ads begin -->
<a href="https://sentrypc.7eer.net/c/5597632/398453/3022" target="_top" id="398453"><img src="//a.impactradius-go.com/display-ad/3022-398453" border="0" alt="www.sentrypc.com" width="580" height="400"/></a><img height="0" width="0" src="https://sentrypc.7eer.net/i/5597632/398453/3022" style="position:absolute;visibility:hidden;" border="0" />
<!-- affiliate ads end -->
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
<li><a href="https://tiktok-videos.techidaily.com/new-in-2024-elevating-tiktok-experience-infusing-videos-with-audio-insights/"><u>[New] In 2024, Elevating TikTok Experience  Infusing Videos with Audio Insights</u></a></li>
<li><a href="https://facebook-video-footage.techidaily.com/new-in-2024-exploring-the-financial-framework-of-youtubes-burgeoning-shorts-fund/"><u>[New] In 2024, Exploring the Financial Framework of YouTube's Burgeoning Shorts Fund</u></a></li>
<li><a href="https://facebook-video-content.techidaily.com/new-in-2024-the-path-to-perfected-video-covers-update-guide-and-tips/"><u>[New] In 2024, The Path to Perfected Video Covers  Update Guide & Tips</u></a></li>
<li><a href="https://twitter-videos.techidaily.com/new-uniting-twitter-and-facebook-through-shared-content-for-2024/"><u>[New] Uniting Twitter and Facebook Through Shared Content for 2024</u></a></li>
<li><a href="https://tiktok-videos.techidaily.com/new-viral-anime-themed-videos-for-tiktok-success-for-2024/"><u>[New] Viral Anime-Themed Videos for TikTok Success for 2024</u></a></li>
<li><a href="https://tiktok-video-recordings.techidaily.com/updated-conquer-tiktok-with-proficient-voiceovers-for-2024/"><u>[Updated] Conquer TikTok with Proficient Voiceovers for 2024</u></a></li>
<li><a href="https://extra-resources.techidaily.com/2024-approved-bypass-samsungs-vr-camera-here-are-your-top-alternatives/"><u>2024 Approved  Bypass Samsung’s VR Camera - Here Are Your Top Alternatives</u></a></li>
<li><a href="https://youtube-help.techidaily.com/2024-approved-the-mini-video-medley-where-does-your-audience-flock-youtubes-or-tiktoks/"><u>2024 Approved  The Mini-Video Medley  Where Does Your Audience Flock? YouTubes or TikToks?</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-v30-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo V30 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-v30-lite-5g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo V30 Lite 5G | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-x-flip-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo X Flip | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-x-fold-2-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo X Fold 2 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-x100-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo X100 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-x100-pro-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo X100 Pro | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-x90s-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo X90S | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y02t-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y02T | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y100-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y100 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y100-5g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y100 5G | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y100a-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y100A | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y100i-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y100i | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y100i-power-5g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y100i Power 5G | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y100t-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y100t | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y17s-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y17s | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y200-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y200 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y200e-5g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y200e 5G | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y27-4g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y27 4G | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y27-5g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y27 5G | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y27s-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y27s | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y28-5g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y28 5G | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y36-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y36 | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y36i-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y36i | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y55s-5g-2023-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y55s 5G (2023) | Dr.fone</u></a></li>
<li><a href="https://phone-solutions.techidaily.com/3-easy-solutions-to-hard-reset-vivo-y56-5g-drfone-by-drfone-reset-android-reset-android/"><u>3 Easy Solutions to Hard Reset Vivo Y56 5G | Dr.fone</u></a></li>
<li><a href="https://unlock-android.techidaily.com/downloading-samfw-frp-tool-30-for-honor-90-lite-by-drfone-android/"><u>Downloading SamFw FRP Tool 3.0 for Honor 90 Lite</u></a></li>
<li><a href="https://fake-location.techidaily.com/fake-the-location-to-get-around-the-mlb-blackouts-on-itel-p40plus-drfone-by-drfone-virtual-android/"><u>Fake the Location to Get Around the MLB Blackouts on Itel P40+ | Dr.fone</u></a></li>
<li><a href="https://techidaily.com/how-to-get-out-of-recovery-on-apple-iphone-13-pro-max-drfone-by-drfone-ios-system-repair-ios-system-repair/"><u>How To Get Out of Recovery on Apple iPhone 13 Pro Max? | Dr.fone</u></a></li>
<li><a href="https://desktop-recording.techidaily.com/in-2024-essential-methods-for-documenting-overwatch-games/"><u>In 2024, Essential Methods for Documenting Overwatch Games</u></a></li>
<li><a href="https://instagram-videos.techidaily.com/in-2024-master-the-art-of-fast-instagram-videos/"><u>In 2024, Master the Art of Fast Instagram Videos</u></a></li>
<li><a href="https://some-approaches.techidaily.com/in-2024-timing-duration-for-a-20mb-high-definition-video/"><u>In 2024, Timing Duration for a 20Mb High-Definition Video</u></a></li>
<li><a href="https://win11.techidaily.com/master-keyboard-flair-with-typingaid-techniques/"><u>Master Keyboard Flair with TypingAid Techniques</u></a></li>
<li><a href="https://audio-shaping.techidaily.com/new-enhance-your-audio-remove-background-noise-with-wondershare-filmora/"><u>New Enhance Your Audio Remove Background Noise With Wondershare Filmora</u></a></li>
<li><a href="https://ai-video-apps.techidaily.com/new-top-explainer-video-tools-to-boost-your-brand/"><u>New Top Explainer Video Tools to Boost Your Brand</u></a></li>
<li><a href="https://fake-location.techidaily.com/prevent-cross-site-tracking-on-lava-blaze-curve-5g-and-browser-drfone-by-drfone-virtual-android/"><u>Prevent Cross-Site Tracking on Lava Blaze Curve 5G and Browser | Dr.fone</u></a></li>
<li><a href="https://fix-guide.techidaily.com/proven-ways-to-fix-there-was-a-problem-parsing-the-package-on-nubia-red-magic-9-pro-drfone-by-drfone-fix-android-problems-fix-android-problems/"><u>Proven Ways to Fix There Was A Problem Parsing the Package on Nubia Red Magic 9 Pro | Dr.fone</u></a></li>
<li><a href="https://tiktok-video-recordings.techidaily.com/sizzling-social-hot-food-trends-on-tiktok/"><u>Sizzling Social  Hot Food Trends on TikTok</u></a></li>
<li><a href="https://tiktok-clips.techidaily.com/snapchats-trendsetting-boomerangs-explained/"><u>SnapChat's Trendsetting Boomerangs Explained</u></a></li>
<li><a href="https://extra-information.techidaily.com/top-iosandroid-apps-for-creative-photo-captions/"><u>Top iOS/Android Apps for Creative Photo Captions</u></a></li>
<li><a href="https://unlock-android.techidaily.com/universal-unlock-pattern-for-itel-a05s-by-drfone-android/"><u>Universal Unlock Pattern for Itel A05s</u></a></li>
<li><a href="https://smart-video-editing.techidaily.com/updated-pro-video-editor-secrets-adding-awesome-effects-to-your-videos/"><u>Updated Pro Video Editor Secrets Adding Awesome Effects to Your Videos</u></a></li>
</ul></div>
