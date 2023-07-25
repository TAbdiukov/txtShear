# txtShear
![icon](icons8-insulin-pen-96.png) Fast engine to skew (or shear) text by a desired angle, emulating handwriting.

[![Download](https://img.shields.io/badge/download-success?style=for-the-badge&logo=github&logoColor=white)](https://github.com/TAbdiukov/txtShear/releases/download/2.00/txtShear.exe)

## Usage
All arguments are required due to VB6 limitations.

    txtShear <out_mode> <font_size> <font_col> <form_x> <form_y>
    <form_bg_col> <ang> "<font>" "<text>"
For help information, just run:

	txtShear

### For example:
      txtShear 1 14 FF0000 500 500 FFFF00 90 "Arial" "Text"

### Manual:
    <out_mode> - Output mode. 4 modes currently supported
            * 1 - Use VB6 inbuilt form -> image functions. Outputs .bmp file
            * 2 - Use WinAPI effecient form -> image workarounds. Experimental
            * 3 - Print out. Use in combination w/ virt. printer, e.g. doPDF
            * 4 - Do operations and then wait utill form_click (or until you kill the process). Can be automated, e.g. AHK+PicPick

    <font_size> - Font size. 1-1368
    <font_col> - Font colour. HEX notation, 000000-FFFFFF
    <form_x> - Canvas width
    <form_y> - Canvas height
    <form_bg_col> - Canvas background colour. HEX notation, 000000-FFFFFF
    <ang> - Angle in degrees. -359 - 359
    <font> - Font name. Must be TrueType. To list TrueType fonts, run 'txtShear list'
    <text> - Text to print
 
## How to compile
1. *[Recommended for compatibility]* Get a Windows XP VM
2. Get **Microsoft Visual Basic 6.0** 

	* **Tip:** There is is a portable build, only a few megabytes. Look up <ins>Portable Microsoft Visual Basic 6.0 SP6</ins>

3. Start **Microsoft Visual Basic 6.0**, open up the project.
4. Go to File → Make *.exe → Save
5. Patch the app for CLI use:
	* You can use my [AMC patcher](https://github.com/TAbdiukov/AMC_patcher-CLI). For example,

		```
		amc C:\Projects\txtShear\txtShear.exe 3
		```
		
	* Or you can use the original Nirsoft's [Application Mode Changer](http://www.nirsoft.net/vb/console.zip) ([docs](http://www.nirsoft.net/vb/console.html)), unpack the archive and then run **appmodechange.exe**

6. Done!


## Example TrueType fonts

 1 | 2 | 3 | 4 
---|---|---|---
Courier Arial | Arial CYR | Arial Cyr | Courier New | Courier New CYR
Courier | Courier New Cyr | Lucida Console |  Lucida Sans Unicode 
Times New Roman | Times New Roman CYR | Times New Roman Cyr | Symbol
Verdana | Arial Black | Comic Sans MS |  Impact
Georgia | Franklin | Gothic Medium |  Palatino Linotype
Tahoma  | Trebuchet MS | Sylfaen | Microsoft Sans Serif


## Backstory

A few days ago I felt particularly curious about my pre-uni projects. And to my surprise, among trashy stuff I found something very peculiar. 

I found an extremely unoptimised and unstable (even by VB6 standards), but functional VB6 app... to simulate handwriting. It made use of:

* Different angles
* Different fonts
* Vertical and horizontal offsets
* Different canvas sizes
* Italics

I find it really charming and naive that I thought I could do something like that... in VB6. Nevertheless, the idea is really unique, and the app made use of some really tricky&sneaky WinAPI combinations (which I likely copied from somewhere, the sources likely lost forever). So since the idea is pretty unique, I decided to rewrite the project, with the command-line in mind. 

Many old functions I had to rid off/merge/rewrite completely. I spent a few days (so far), implementing, refactoring and rewriting, and by far I'm really happy with the result.

## See also
*My other small Windows tools,*  

* [AMC_patcher-CLI](https://github.com/TAbdiukov/AMC_patcher-CLI) – (CLI) Patches app's SUBSYSTEM flag to modify app's behavior.
* [exe2wordsize](https://github.com/TAbdiukov/exe2wordsize) – (CLI) Detects Windows-compatible application bitness, without ever running it.
* [SCAPTURE.EXE](https://github.com/TAbdiukov/SCAPTURE.EXE) – (GUI) Simple screen-capturing tool for embedded systems.
* **<ins>txtShear</ins>** – (CLI) Fast engine to skew (or shear) text by a desired angle, emulating handwriting.
