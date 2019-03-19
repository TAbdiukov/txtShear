Attribute VB_Name = "API"
Option Explicit

' Consts
'' In use by LOGFONT object
Public Const LF_FACESIZE As Long = 32&

'' In use by DrawStr
Public Const OBJ_FONT As Long = 6&

'' In use by IsFontTrueType
Public Const TMPF_TRUETYPE = &H4

'' Return codes
Public Const ERR_SUCCESS As Long = 0
Public Const ERR_ARGS As Long = 1
Public Const ERR_NonTT As Long = 2
Public Const ERR_OUT As Long = 3
Public Const ERR_VB As Long = 127

'' Output modes
Public Const OUT_ERR As Integer = 0
Public Const OUT_IMG As Integer = 1
Public Const OUT_WINAPI As Integer = 2
Public Const OUT_PRN As Integer = 3

'' Pseudo consts
Public APP_NAME As String
Public DEBUGGER As Boolean
Public VER As String

' Declare Windows API functions
''In use by DrawStr
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

'' In use by DrawStr
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'' In use by DrawStr
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long ' GetObject bp API переименована, чтобы не затеняла GetObject из VB.

'' In use by DrawStr
Public Declare Function GetObjectDC Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

'' In use by IsFontTrueType
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

'' In use by DrawStr
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'' In use by DrawStr
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

'' In generic use
'' https://social.msdn.microsoft.com/Forums/sqlserver/en-US/d6e76731-8e3b-465f-9d5a-12c6498d6b6c/how-to-return-exit-code-from-vb6-form?forum=winforms
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

'' In use by SaveFormImageToFile
Private Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long


' Declare objects
'' In use by DrawStr + IsFontTrueType
Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

'' In use by IsFontTrueType
Public Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

' Is font True type?
' Non true type fonts are UNsupported
Public Function IsFontTrueType(f As Form, sFontName As String) As Boolean
        Dim lf As LOGFONT
        Dim tm As TEXTMETRIC
        Dim oldfont As Long, newfont As Long
        Dim tmpArray() As Byte
        Dim dummy As Long
        Dim Z As Integer
        
        'need to convert font name to byte array...
        tmpArray = StrConv(sFontName & vbNullString, vbFromUnicode)
        For Z = 0 To UBound(tmpArray)
                lf.lfFaceName(Z + 1) = tmpArray(Z)
        Next
        
        'create the font object
        newfont = CreateFontIndirect(lf)
        'save the current font object and use the new font object
        oldfont = SelectObject(f.hdc, newfont)
        'get the new font object's info
        dummy = GetTextMetrics(f.hdc, tm)
        'determine whether new font object is TrueType
        IsFontTrueType = (tm.tmPitchAndFamily And TMPF_TRUETYPE)
        'restore the original font object - !!!THIS IS IMPORTANT!!!
        dummy = SelectObject(f.hdc, oldfont)
End Function

' Pretty much the main function
' Src: unknown
' Usage:
' DrawStr (
'   <your_form>.hdc (whatever that means)
'   string_to_inprint
'   X pos on the form
'   Y pos on the form
'   (apparently)angle_in_degrees, from -359 to +359
' )
Public Sub DrawStr( _
        ByVal hdc As Long, _
        txt As String, _
        ByVal x As Long, ByVal y As Long, _
        ByVal Angle As Long)

        Dim hfnt As Long, hfntPrev As Long, lfont As LOGFONT
        
        hfntPrev = GetCurrentObject(hdc, OBJ_FONT)
        GetObjectDC hfntPrev, Len(lfont), lfont
        lfont.lfEscapement = Angle
        lfont.lfOrientation = Angle
        hfnt = CreateFontIndirect(lfont)
        hfntPrev = SelectObject(hdc, hfnt)
        TextOut hdc, x, y, txt, Len(txt)
        SelectObject hdc, hfntPrev
        DeleteObject hfnt
End Sub

' My wrapping of DrawStr
Public Function DrawWrap( _
    ByRef f As Form, _
    FormX As Integer, _
    FormY As Integer, _
    FormBG As Long, _
    FontName As String, _
    FontSize As Integer, _
    FontColor As Long, _
    Angle As Long, _
    txt As String _
)

    f.Width = FormX
    f.Height = FormY
    f.BackColor = FormBG
    
    f.Font.Name = FontName
    f.Font.Size = FontSize
    f.ForeColor = FontColor
    
    DrawStr f.hdc, txt, 0, 0, Angle
End Function

Public Function UnixTime() As Long
    ' approach 1: https://stackoverflow.com/a/2259363
    ' UnixTime = DateDiff("S", "1/1/1970", now())
    ' approach 2: https://stackoverflow.com/a/52406421
    ' CLng(Format(Now(), "ms"))
    UnixTime = UnixTime = DateDiff("S", "1/1/1970", Now())
End Function

' https://www.codeproject.com/Articles/23234/VB6-Save-Form-Image-To-File
' Used when output mode is OUT_WINAPI
Public Sub SaveFormImageToFile(ByRef ContainerForm As Form, _
                               ByRef PictureBoxControl As PictureBox, _
                               ByVal ImageFileName As String)
                               
  Dim FormInsideWidth As Long
  Dim FormInsideHeight As Long
  Dim PictureBoxLeft As Long
  Dim PictureBoxTop As Long
  Dim PictureBoxWidth As Long
  Dim PictureBoxHeight As Long
  Dim FormAutoRedrawValue As Boolean
  
  With PictureBoxControl
    'Set PictureBox properties
    .Visible = False
    .AutoRedraw = True
    .Appearance = 0 ' Flat
    .AutoSize = False
    .BorderStyle = 0 'No border
    
    'Store PictureBox Original Size and location Values
    PictureBoxHeight = .Height: PictureBoxWidth = .Width
    PictureBoxLeft = .Left: PictureBoxTop = .Top
    
    'Make PictureBox to size to inside of form.
    .Align = vbAlignTop: .Align = vbAlignLeft
    DoEvents
    
    FormInsideHeight = .Height: FormInsideWidth = .Width
    
    'Restore PictureBox Original Size and location Values
    .Align = vbAlignNone
    .Height = FormInsideHeight: .Width = FormInsideWidth
    .Left = PictureBoxLeft: .Top = PictureBoxTop
    
    FormAutoRedrawValue = ContainerForm.AutoRedraw
    ContainerForm.AutoRedraw = False
    DoEvents
    
    'Copy Form Image to Picture Box
    BitBlt .hdc, 0, 0, _
    FormInsideWidth / Screen.TwipsPerPixelX, _
    FormInsideHeight / Screen.TwipsPerPixelY, _
    ContainerForm.hdc, 0, 0, _
    vbSrcCopy
    
    DoEvents
    SavePicture .Image, ImageFileName
    DoEvents
    
    ContainerForm.AutoRedraw = FormAutoRedrawValue
    DoEvents
  End With
End Sub

Public Function OutputForm( _
 ByRef f As Form, ByVal mode As Integer, ByVal seed As Long)

    ' switch-case: https://stackoverflow.com/a/51016198
    Dim outName As String
    outName = APP_NAME + "_" + CStr(seed)
    
    Select Case mode
        Case OUT_IMG:
            SavePicture f.Image, outName + ".bmp"
        Case OUT_WINAPI:
            SaveFormImageToFile f, f.Picture1, outName + ".experimental"
        Case OUT_PRN:
            f.PrintForm
        Case Else:
            CLI.Send "ERROR: invalid output mode"
            quit API.ERR_OUT
    End Select
End Function

' https://stackoverflow.com/a/9068210
Public Function GetRunningInIDE() As Boolean
   Dim x As Long
   Debug.Assert Not TestIDE(x)
   GetRunningInIDE = x = 1
End Function

' https://stackoverflow.com/a/9068210
Private Function TestIDE(x As Long) As Boolean
    x = 1
End Function

Public Sub setup()
    '' Generic
    
    APP_NAME = "hwz"
    DEBUGGER = GetRunningInIDE()
    VER = CStr(App.Major) + "." + CStr(App.Minor)
    
End Sub

Public Function quit(code As Integer)
    On Error Resume Next

    CLI.Send vbNewLine
    CLI.Send vbNewLine
    
    If DEBUGGER Then
        Debug.Print "End"
    Else
        ExitProcess code
    End If
End Function

' http://www.freevbcode.com/ShowCode.asp?ID=6324
' (slightly modified for performance)
Public Function HEXCOL2RGB(ByVal HexColor As String) As Long

    'The input at this point could be HexColor = "#00FF1F"
    Dim Red As String
    Dim Green As String
    Dim Blue As String
    
    'Color = Replace(HexColor, "#", "")
        'Here HexColor = "00FF1F"
    
    Red = Val("&H" & Mid(HexColor, 1, 2))
        'The red value is now the long version of "00"
    
    Green = Val("&H" & Mid(HexColor, 3, 2))
        'The red value is now the long version of "FF"
    
    Blue = Val("&H" & Mid(HexColor, 5, 2))
        'The red value is now the long version of "1F"
    
    
    HEXCOL2RGB = RGB(Red, Green, Blue)
        'The output is an RGB value

End Function


Public Sub goto_VB_err()
    CLI.Sendln "VB6 Error: " + CStr(Err.Number)
    CLI.Sendln Err.Description
    CLI.Sendln "    Dll Error:" + CStr(Err.LastDllError)
    CLI.Sendln "    At" + Err.Source
    API.quit API.ERR_VB
End Sub

