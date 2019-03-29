VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "hwz"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3885
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   43.921
   ScaleMode       =   0  'User
   ScaleWidth      =   68.527
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' malloc variables
Public OutMode_local As Integer ' 0
Public FontSize_local As Integer ' 1
Public FontColor_local As Long ' 2
Public FormX_local As Integer ' 3
Public FormY_local As Integer  ' 4
Public FormBG_local As Long ' 5
Public Angle_local As Long ' 6
Public FontName_local As String ' 7
Public txt_local As String ' 8

Private Sub setup()
    ' Initialise form
    Me.AutoRedraw = True
    Me.Picture1.Visible = False
End Sub

Private Sub Form_Click()
    If (API.is_waiting_mode(OutMode_local)) Then
        API.quit API.ERR_SUCCESS
    End If
End Sub

Private Sub Form_Load()
    ' Initialise everything
    setup
    API.setup
    CLI.setup
    
    ' Initialise args
    Dim argw() As String
    Dim argc As Integer
    Dim TrimArg As String

    TrimArg = Trim(Command)
    argw = Split(TrimArg, " ")
    argc = UBound(argw) - LBound(argw) + 1 ' https://forums.windowssecrets.com/showthread.php/28214-counting-array-elements-(vb6)
        
    ' If number of arguments suffices
    If (argc >= 9) Then
        ' INPUT
        'MsgBox "1"
        'Dim i As Integer
        'For i = 0 To 7
        '    MsgBox "argw[" + CStr(i) + "]: " + argw(i)
        'Next
        
        ' and then set them accordingly
        OutMode_local = CInt(argw(0))
        FontSize_local = CInt(argw(1))
        FontColor_local = API.HEXCOL2RGB(argw(2))
        FormX_local = CInt(argw(3))
        FormY_local = CInt(argw(4))
        FormBG_local = API.HEXCOL2RGB(argw(5))
        Angle_local = CLng(argw(6))
        'MsgBox "2"
        
        'soupy get strings
        Dim soup As String
        Dim soup_len As Integer
        Dim soup_arr() As String
        
        soup_len = Len(TrimArg) - Len(argw(0)) - Len(argw(1)) - Len(argw(2)) - Len(argw(3)) - Len(argw(4)) - Len(argw(5)) - Len(argw(6)) - 7
        'MsgBox "3 Debug, soup_len: " + CStr(soup_len)
        soup = Right(TrimArg, soup_len)
        'MsgBox "4 Debug, soup: " + soup
        
        soup_arr = Split(soup, Chr(34))
        'Dim i As Integer
        'For i = 0 To 3
        '    MsgBox "soup_arr[" + CStr(i) + "]: |" + soup_arr(i) + "|"
        'Next
        
        FontName_local = soup_arr(1)
        txt_local = soup_arr(3)
        'MsgBox "5"
        
        ' OUTPUT
        If (API.IsFontTrueType(Me, FontName_local)) Then
            ' Is true type = good
            'MsgBox "6"
            DrawWrap Me, _
             FormX_local, FormY_local, FormBG_local, _
             FontName_local, FontSize_local, FontColor_local, _
             Angle_local, txt_local
            'MsgBox "7"
            If (Err.Number > 0) Then
                API.printErr
            Else
                OutputForm Me, OutMode_local, API.UnixTime()
            End If
            'MsgBox "8"
        Else
            CLI.Send "ERROR: Font is not a TrueType font (got: " _
             + FontName_local + ")"
            API.quit API.ERR_NonTT
        End If
            
        ' FIN
        CLI.Send "Success"
        API.quit API.ERR_SUCCESS
    Else
        CLI.Send "ERROR: Invalid number of args (got " + CStr(argc) + ")"
        showHelp
        
        API.quit API.ERR_ARGS
    End If
        
End Sub

Public Sub showHelp()
    CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_BLUE Or CLI.FOREGROUND_INTENSITY
    CLI.Sendln "HWZ - TEXT SKEWER" + " v" + VER
    CLI.Sendln ""
    
    CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_INTENSITY
    CLI.Sendln "USAGE:"
    CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
    CLI.Sendln "hwz <out_mode> <font_size> <font_col> " & _
     "<form_x> <form_y> <frm_bg_col> <ang> " & _
     Chr(34) & "<font>" & Chr(34) & " " & Chr(34) & "<text>" & Chr(34)
    CLI.Sendln ""
    
    CLI.SetTextColour CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_INTENSITY
    CLI.Sendln "FOR EXAMPLE:"
    CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
    CLI.Sendln "hwz 1 14 FF0000 500 500 FFFFFF 90 " & _
     Chr(34) & "Arial" & Chr(34) & " " & Chr(34) & "Text" & Chr(34)
    CLI.Sendln ""
    
    CLI.SetTextColour CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE Or CLI.FOREGROUND_INTENSITY
    CLI.Sendln "MANUAL:"
    CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
    CLI.Sendln "<out_mode> - Output mode. 4 modes currently supported"
    CLI.Sendln vbTab + "* 1: Use VB6 inbuilt form -> image functions. Outputs .bmp file"
    CLI.Sendln vbTab + "* 2: Use WinAPI effecient form -> image workarounds. Experimental"
    CLI.Sendln vbTab + "* 3: Print out. Use in combination w/ virt. printer, e.g. doPDF"
    CLI.Sendln vbTab + "* 4: Do&wait till form_click. Use w/ automation combo, e.g. AHK+doPDF"
    CLI.Sendln ""
    CLI.Sendln "<font_size> - Font size. 1-1368"
    CLI.Sendln "<font_col> - Font colour. HEX notation, 000000-FFFFFF"
    CLI.Sendln "<form_x> - Canvas width"
    CLI.Sendln "<form_y> - Canvas height"
    CLI.Sendln "<form_bg_col> - Canvas background colour. HEX notation, 000000-FFFFFF"
    CLI.Sendln "<ang> - Angle in degrees. -359 - 359"
    CLI.Sendln "<font> - Font name. Must be TrueType"
    CLI.Sendln "<text> - Text to print"
    API.quit API.ERR_SUCCESS
End Sub

