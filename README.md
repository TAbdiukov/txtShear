# HWZ
**CLI TEXT SKEWER**

A small but very fast Windows app to skew desired text to a desired angle, ultimately emulating handwriting

## Backstory

A few days ago I felt particularly curious about my pre-uni projects. And to my surprise, among trashy stuff I found something very peculiar. 

I found an extremely unoptimised (even by VB6 standards), crashy **but working** VB6 app... to simulate handwriting. It made use of:

* Different angles
* Different fonts
* Vertical and horizontal offsets
* Different canvas sizes
* Italics

I find it really charming and naive that I thought I could do something like that... in VB6. Nevertheless, the idea is really unique, and the app made use of some really tricky&sneaky WinAPI combinations (which I likely copied from somewhere, the sources likely lost forever). So since the idea is pretty unique, I decided to rewrite the project, with the command-line in mind, so that one could write a nice and easy Python/Java/(paste your fav lingua here) wrapper for it. 

Many old functions I had to rid off/merge/rewrite completely. I spent a few days (so far), implementing, cutting and rewriting, and by far I'm really happy with the result.

Ironically, the biggest obstacle turned out to be... CLI input/output. But on the bright side I now know how to CLI in VB6.

#### Naming
The original project's name roughly said *"Hand-Writer-Z"* (what a name heh).So in honour of the old code, and at the same time for convenient command-line usage, I trunkated the name to *hwz*

## Usage:
Unfortunately, since we deal with Visual Basic, there is no (known) effecient way to handle args, so all args are required. Which is frankly to biggie for this kind of software

    hwz <out_mode> <font_size> <font_col> <form_x> <form_y>
    <form_bg_col> <ang> "<font>" "<text>"
For help information, just run:

	hwz

### For example:
      hwz 1 14 FF0000 500 500 FFFFFF 90 "Arial" "Text"

### Manual:
    <out_mode> - Output mode. 4 modes currently supported
            * 1 - Use VB6 inbuilt form -> image functions. Outputs .bmp file
            * 2 - Use WinAPI effecient form -> image workarounds. Experimental
            * 3 - Print out. Use in combination w/ virt. printer, e.g. doPDF
			* 4 - Do operations and then wait utill form_click (or until you kill the process). Use w/ automation tool combinations, e.g. AHK+PicPick

    <font_size> - Font size. 1-1368
    <font_col> - Font colour. HEX notation, 000000-FFFFFF
    <form_x> - Canvas width
    <form_y> - Canvas height
    <form_bg_col> - Canvas background colour. HEX notation, 000000-FFFFFF
    <ang> - Angle in degrees. -359 - 359
    <font> - Font name. Must be TrueType. To list TrueType fonts, run 'hwz list'
    <text> - Text to print
    
## How to compile?
1. *[Recommended for compatibility]* Get a Windows XP VM
2. Get a **Microsoft Visual Basic 6.0** 

***Tip:** I unofficially recommend a portable version sticking around on BT, as you won't have to mess around with the installation and registry. Plus, it's only a few megabytes. Check out **Portable Microsoft Visual Basic 6.0 SP6***

3. Fire up **Microsoft Visual Basic 6.0**, open up the project.
4. Go to File -> Make *.exe -> Save
5. Patch the app for CLI use:
* You can use my [AMC patcher](https://github.com/TAbdiukov/AMC_patcher-CLI). For example,
	amc C:\Projects\HWZ\hwz.exe 3

* Or you can use the original Nirsoft's [Application Mode Changer](http://www.nirsoft.net/vb/console.zip) ([info](http://www.nirsoft.net/vb/console.html)), unpack the archive and then run the **appmodechange.exe**

**Tip:** On my VM, for whatever reason the  **appmodechange.exe** fails to launch. As a workaround, you can run another **Microsoft Visual Basic 6.0** window, open up the **appmodechange.vbp** project, and then right from IDE, go to Run -> Run With Full Compile*

6. Browse to your compiled copy of HWZ, pick *Console Application*, and then click on *Change Mode*.
7. Done!

## Example TrueType fonts
	Courier Arial   Arial CYR       Arial Cyr       Courier New     Courier New CYR
	Courier Courier New Cyr Lucida Console  Lucida Sans Unicode     Times New Roman
	Times New Roman CYR     Times New Roman Cyr     Symbol  Verdana Arial Black
	Comic Sans MS   Impact  Georgia Franklin Gothic Medium  Palatino Linotype
	Tahoma  Trebuchet MS    Sylfaen Microsoft Sans Serif
	
	