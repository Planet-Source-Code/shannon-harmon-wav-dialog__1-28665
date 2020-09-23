Attribute VB_Name = "MGetWavFile"
'--------------------------------------------------------------------------------
'Author :   Shannon Harmon - shannonh@theharmonfamily.com
'Date   :   11/05/2001
'Notes  :   Should be converted to an ActiveX to keep the images in frmMain together
'           with the code, etc...and to allow for easy configuration of properties
'           by the user.  Made this mainly for testing...
'TODO   :   Make the play/stop buttons enable and disable as needed.
'           Test, look for bugs, overlooked stuff.
'           Bitblt the images from frmMain so the backgrounds are transparent.
'--------------------------------------------------------------------------------
Option Explicit

'Window Messages
Private Const WM_DESTROY = &H2
Private Const WM_GETFONT = &H31
Private Const WM_INITDIALOG = &H110
Private Const WM_KEYUP = &H101
Private Const WM_KILLFOCUS = &H8
Private Const WM_LBUTTONUP = &H202
Private Const WM_NOTIFY = &H4E
Private Const WM_SETFONT = &H30
Private Const WM_USER = &H400

'Common Dialog Notification Messages
Private Const CDN_INITDONE = -601
Private Const CDN_SELCHANGE = -602
Private Const CDN_FOLDERCHANGE = -603
Private Const CDN_TYPECHANGE = -607
Private Const CDN_SHAREVIOLATION = -604
Private Const CDN_HELP = -605
Private Const CDN_FILEOK = -606

'Common Dialog Messages
Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_GETSPEC = (CDM_FIRST + &H0)
Private Const CDM_GETFILEPATH = (CDM_FIRST + &H1)
Private Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)
Private Const CDM_GETFOLDERIDLIST = (CDM_FIRST + &H3)
Private Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
Private Const CDM_HIDECONTROL = (CDM_FIRST + &H5)
Private Const CDM_SETDEFEXT = (CDM_FIRST + &H6)

'Open/Save FileName Flags
Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_CREATEPROMPT As Long = &H2000
Private Const OFN_ENABLEHOOK As Long = &H20
Private Const OFN_ENABLETEMPLATE As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_EXTENSIONDIFFERENT As Long = &H400
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_LONGNAMES As Long = &H200000
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_NODEREFERENCELINKS As Long = &H100000
Private Const OFN_NOLONGNAMES As Long = &H40000
Private Const OFN_NONETWORKBUTTON As Long = &H20000
Private Const OFN_NOREADONLYRETURN As Long = &H8000&
Private Const OFN_NOTESTFILECREATE As Long = &H10000
Private Const OFN_NOVALIDATE As Long = &H100
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_READONLY As Long = &H1
Private Const OFN_SHAREAWARE As Long = &H4000
Private Const OFN_SHAREFALLTHROUGH As Long = 2
Private Const OFN_SHAREWARN As Long = 0
Private Const OFN_SHARENOWARN As Long = 1
Private Const OFN_SHOWHELP As Long = &H10
Private Const OFS_MAXPATHNAME As Long = 260

'sndPlaySound Flags
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2

'Class Styles
Private Const BS_PUSHBUTTON = &H0
Private Const BS_BITMAP = &H80
Private Const BS_TEXT = &H0
Private Const SS_LEFT = &H0
Private Const WS_CHILD = &H40000000
Private Const WS_GROUP = &H20000
Private Const WS_TABSTOP = &H10000
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILDWINDOW = WS_CHILD

'Extended Window Styles
Private Const WS_EX_NOPARENTNOTIFY = &H4
Private Const WS_EX_LEFT = &H0
Private Const WS_EX_LTRREADING = &H0
Private Const WS_EX_RIGHTSCROLLBAR = &H0

'Button Class Messages
Private Const BM_SETIMAGE = &HF7

'Misc Constants
Private Const GWL_WNDPROC = (-4)
Private Const SW_NORMAL = 1
Private Const MAX_PATH = 260
Private Const MAX_LENGTH = 1024
Private Const VK_SPACE = &H20
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const FW_NORMAL = 400
Private Const ANSI_CHARSET = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const VARIABLE_PITCH = 2
Private Const FF_SWISS = 32
Private Const LF_FACESIZE = 32

'Rect Structure
Private Type RECT
    Left              As Long
    Top               As Long
    Right             As Long
    Bottom            As Long
End Type

'OS Version Structure
Private Type OSVERSIONINFO
    OSVSize           As Long
    dwVerMajor        As Long
    dwVerMinor        As Long
    dwBuildNumber     As Long
    PlatformID        As Long
    szCSDVersion      As String * 128
End Type

'Open FileName Structure
Private Type OPENFILENAME
    nStructSize       As Long
    hWndOwner         As Long
    hInstance         As Long
    sFilter           As String
    sCustomFilter     As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    sFile             As String
    nMaxFile          As Long
    sFileTitle        As String
    nMaxTitle         As Long
    sInitialDir       As String
    sDialogTitle      As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    sDefFileExt       As String
    nCustData         As Long
    fnHook            As Long
    sTemplateName     As String
End Type

'OpenFileName Structure (Win2k)
Private Type OPENFILENAME2000
    nStructSize       As Long
    hWndOwner         As Long
    hInstance         As Long
    sFilter           As String
    sCustomFilter     As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    sFile             As String
    nMaxFile          As Long
    sFileTitle        As String
    nMaxTitle         As Long
    sInitialDir       As String
    sDialogTitle      As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    sDefFileExt       As String
    nCustData         As Long
    fnHook            As Long
    sTemplateName     As String
    pvReserved        As Long
    dwReserved        As Long
    FlagsEx           As Long
End Type

'Notification Message Structure
Private Type NMHDR
    hwndFrom          As Long
    idfrom            As Long
    code              As Long
End Type

'LogFont Structure
Private Type LOGFONT
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
    lfFaceName(LF_FACESIZE) As Byte
End Type

'Windows API Functions
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As Any) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    
'Local Variables
Private hWndPlay As Long
Private hWndStop As Long
Private hWndPreview As Long
Private lpPlayWndProc As Long
Private lpStopWndProc As Long
Private strCurrentFile As String


Private Function FARPROC(ByVal pfn As Long) As Long
    'Returns the adddress of the File Open/Save dialog callback proc.
    FARPROC = pfn
End Function


Private Sub StopSound()
    'Stops currently playing file, if any
    Dim lngReturn As Long
    lngReturn = sndPlaySound(vbNullString, SND_ASYNC Or SND_NODEFAULT)
End Sub


Private Sub PlaySound()
    'Attempts to play the currently selected file
    Dim lngReturn As Long
    lngReturn = sndPlaySound(strCurrentFile, SND_ASYNC Or SND_NODEFAULT)
End Sub


Private Function PlayWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Callback for Play button created
    
    If uMsg = WM_DESTROY Then
        'Unsubclass this window
        SetWindowLong hWndPlay, GWL_WNDPROC, lpPlayWndProc
    ElseIf (uMsg = WM_LBUTTONUP) Or (uMsg = WM_KEYUP And wParam = VK_SPACE) Then
        'If the left mouse button up or the space key pressed
        PlaySound
        SendMessage hw, WM_KILLFOCUS, 0&, 0&
        PlayWindowProc = CallWindowProc(lpPlayWndProc, hw, uMsg, wParam, lParam)
    Else
        PlayWindowProc = CallWindowProc(lpPlayWndProc, hw, uMsg, wParam, lParam)
    End If
End Function


Private Function StopWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Callback for Stop button created
    
    If uMsg = WM_DESTROY Then
        'Unsubclass this window
        SetWindowLong hWndStop, GWL_WNDPROC, lpStopWndProc
    ElseIf (uMsg = WM_LBUTTONUP) Or (uMsg = WM_KEYUP And wParam = VK_SPACE) Then
        'If the left mouse button up or the space key pressed
        StopSound
        SendMessage hw, WM_KILLFOCUS, 0&, 0&
        StopWindowProc = CallWindowProc(lpStopWndProc, hw, uMsg, wParam, lParam)
    Else
        StopWindowProc = CallWindowProc(lpStopWndProc, hw, uMsg, wParam, lParam)
    End If
End Function


Private Function DialogHookProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Callback function for dialog hook
    
    Dim hWndParent As Long
            
    hWndParent = GetParent(hWnd)
            
    Select Case uMsg
        Case WM_INITDIALOG
            If hWndParent <> 0 Then
                Dim rc As RECT
                Dim lngLeft As Long
                Dim lngTop As Long
                Dim lngWidth As Long
                Dim lngHeight As Long
                Dim lngStyle As Long
                Dim lngExStyle As Long
                Dim lngReturn As Long
            
                'Get the screen position and size of the dialog window
                'so it can be centered and the height increased to allow
                'for our added windows
                Call GetWindowRect(hWndParent, rc)
                lngWidth = rc.Right - rc.Left
                lngHeight = rc.Bottom - rc.Top + 39
                lngLeft = ((Screen.Width \ Screen.TwipsPerPixelX) - lngWidth) \ 2
                lngTop = ((Screen.Height \ Screen.TwipsPerPixelY) - lngHeight) \ 2
                
                'Move the dialog window into it's new position
                Call MoveWindow(hWndParent, lngLeft, lngTop, lngWidth, lngHeight, True)
                
                'Style bits for the Preview static window
                lngStyle = WS_CHILDWINDOW Or WS_VISIBLE Or WS_GROUP Or SS_LEFT
                lngExStyle = WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR Or WS_EX_NOPARENTNOTIFY
                
                'Create the Preview window
                hWndPreview = CreateWindowEx(ByVal lngExStyle, "STATIC", "Preview:", lngStyle, 8, lngHeight - 65, 60, 20, hWndParent, ByVal 0&, App.hInstance, ByVal 0&)
                
                'Make sure the font is not bold
                Dim lFont As LOGFONT
                Dim hDlgFont As Long
                
                'Get the current font, if null, not set, using system default
                hDlgFont = SendMessage(hWndPreview, WM_GETFONT, 0&, ByVal 0&)
                
                If hDlgFont <> 0& Then
                    'It had a font set, get it and update the weight
                    If GetObject(hDlgFont, Len(lFont), lFont) <> 0& Then
                        lFont.lfWeight = FW_NORMAL
                        hDlgFont = CreateFontIndirect(lFont)
                        If hDlgFont <> 0& Then
                            SendMessage hWndPreview, WM_SETFONT, hDlgFont, ByVal 1&
                        End If
                    End If
                Else
                    'No font set, these values should match enough
                    With lFont
                        .lfHeight = 14
                        .lfWidth = 0
                        .lfEscapement = 0
                        .lfOrientation = 0
                        .lfWeight = FW_NORMAL
                        .lfItalic = 0
                        .lfUnderline = 0
                        .lfStrikeOut = 0
                        .lfCharSet = ANSI_CHARSET
                        .lfOutPrecision = OUT_DEFAULT_PRECIS
                        .lfClipPrecision = CLIP_DEFAULT_PRECIS
                        .lfQuality = DEFAULT_QUALITY
                        .lfPitchAndFamily = VARIABLE_PITCH Or FF_SWISS
                        .lfFaceName(0) = 0&
                    End With
                    
                    hDlgFont = CreateFontIndirect(lFont)
                    If hDlgFont <> 0& Then
                        SendMessage hWndPreview, WM_SETFONT, hDlgFont, ByVal 1&
                    End If
                End If
                
                'Style bits for the Play/Stop button windows
                lngStyle = WS_CHILDWINDOW Or WS_VISIBLE Or WS_TABSTOP Or BS_PUSHBUTTON Or BS_TEXT Or BS_BITMAP
                lngExStyle = WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR Or WS_EX_NOPARENTNOTIFY
                
                'Create the Play button
                hWndPlay = CreateWindowEx(ByVal lngExStyle, "BUTTON", "Play", lngStyle, 81, lngHeight - 65, 23, 21, hWndParent, ByVal 0&, App.hInstance, ByVal 0&)
                'Subclass the Play button
                lpPlayWndProc = SetWindowLong(hWndPlay, GWL_WNDPROC, AddressOf PlayWindowProc)
                'Set the image from our picture on the main form
                lngReturn = PostMessage(hWndPlay, BM_SETIMAGE, 0&, frmMain.picPlay.Picture.Handle)
                
                'Create the Stop button
                hWndStop = CreateWindowEx(ByVal 0&, "BUTTON", "Stop", lngStyle, 104, lngHeight - 65, 23, 21, hWndParent, ByVal 0&, App.hInstance, ByVal 0&)
                'Subclass the Stop button
                lpStopWndProc = SetWindowLong(hWndStop, GWL_WNDPROC, AddressOf StopWindowProc)
                'Set the image from our picture on the main form
                lngReturn = PostMessage(hWndStop, BM_SETIMAGE, 0&, frmMain.picStop.Picture.Handle)
                
                DialogHookProc = 1
            End If
          
        Case WM_NOTIFY
            Dim NMH As NMHDR

            CopyMemory NMH, ByVal lParam, LenB(NMH)
             
            Select Case NMH.code
                Case CDN_INITDONE
                    '
                Case CDN_SELCHANGE
                    'Selection changed, update our current file local variable
                    'and call StopSound to stop any currently playing sound
                    Dim strBuffer As String
                    Dim lngNullPos As Long
                    
                    strBuffer = String$(MAX_PATH, 0)
                    lngNullPos = SendMessage(hWndParent, CDM_GETFILEPATH, MAX_PATH, ByVal strBuffer)
                    strBuffer = Left$(strBuffer, lngNullPos - 1)
                    strCurrentFile = strBuffer
                    
                    StopSound
                
                Case Else:
            End Select
                  
        Case WM_DESTROY
            'Dialog window was destroyed, stop playing any currently playing sound
            StopSound
            
        Case Else:
    End Select
End Function


Private Function GetOpenWavFileName2000(Optional ByVal hWndOwner As Long = 0&, _
                                        Optional ByVal strInitDir As String = "", _
                                        Optional ByVal strFile As String = "", _
                                        Optional ByVal strDialogTitle As String = "Open Wav File") As String
    
    'Open filen dialog with the Windows 2k style Places Bar on the left
    
    Dim OFN As OPENFILENAME2000

    With OFN
       .nStructSize = Len(OFN)
       .hWndOwner = hWndOwner
       .sFilter = "Wav Files (*.wav)" & vbNullChar & "*.wav" & vbNullChar & vbNullChar
       .nFilterIndex = 2
       .sFile = strFile & String$(MAX_LENGTH - Len(strFile), 0)
       .nMaxFile = MAX_LENGTH
       .sDefFileExt = "wav" & ""
       .sFileTitle = String$(MAX_LENGTH, 0)
       .nMaxTitle = MAX_LENGTH
       .sInitialDir = strInitDir
       .sDialogTitle = strDialogTitle
       .flags = OFN_EXPLORER Or OFN_ENABLEHOOK Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES
       .fnHook = FARPROC(AddressOf DialogHookProc)
    End With
   
    If GetOpenFileName(OFN) Then
        If InStr(OFN.sFile, Chr$(0)) Then
            GetOpenWavFileName2000 = Left$(OFN.sFile, InStr(OFN.sFile, Chr$(0)))
        Else
            GetOpenWavFileName2000 = OFN.sFile
        End If
    End If
End Function

Private Function GetOpenWavFileNameStd(Optional ByVal hWndOwner As Long = 0&, _
                                       Optional ByVal strInitDir As String = "", _
                                       Optional ByVal strFile As String = "", _
                                       Optional ByVal strDialogTitle As String = "Open Wav File") As String
    
    'Standard open file dialog
    'Same exact function as GetOpenWavFileName2000 except the
    'OFN variable is a different structure
    
    Dim OFN As OPENFILENAME

    With OFN
       .nStructSize = Len(OFN)
       .hWndOwner = hWndOwner
       .sFilter = "Wav Files (*.wav)" & vbNullChar & "*.wav" & vbNullChar & vbNullChar
       .nFilterIndex = 2
       .sFile = strFile & String$(MAX_LENGTH - Len(strFile), 0)
       .nMaxFile = MAX_LENGTH
       .sDefFileExt = "wav" & ""
       .sFileTitle = String$(MAX_LENGTH, 0)
       .nMaxTitle = MAX_LENGTH
       .sInitialDir = strInitDir
       .sDialogTitle = strDialogTitle
       .flags = OFN_EXPLORER Or OFN_ENABLEHOOK Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES
       .fnHook = FARPROC(AddressOf DialogHookProc)
    End With
   
    If GetOpenFileName(OFN) Then
        If InStr(OFN.sFile, Chr$(0)) Then
            GetOpenWavFileNameStd = Left$(OFN.sFile, InStr(OFN.sFile, Chr$(0)))
        Else
            GetOpenWavFileNameStd = OFN.sFile
        End If
    End If
End Function

Public Function GetOpenWavFileName(Optional ByVal hWndOwner As Long = 0&, _
                                   Optional ByVal strInitDir As String = "", _
                                   Optional ByVal strFile As String = "", _
                                   Optional ByVal strDialogTitle As String = "Open Wav File", _
                                   Optional ByVal fShowPlaces As Boolean = True) As String
    
    'Public accessible function to get the filename and call the proper
    'function to allow for the Places Bar or not, will safely call the
    'correct function even if fShowPlaces is true and the OS is not Win2k
    
    Dim fWin2k As Boolean
    
    'If ShowPlaces was requested, make sure the OS is Win2k
    If fShowPlaces Then
        #If Win32 Then
            'Get windows version
            Dim OSV As OSVERSIONINFO
            OSV.OSVSize = Len(OSV)
            
            If GetVersionEx(OSV) = 1 Then
                fWin2k = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And (OSV.dwVerMajor = 5)
            End If
        #End If
    End If
    
    If fWin2k Then
        GetOpenWavFileName = GetOpenWavFileName2000(hWndOwner, strInitDir, strFile, strDialogTitle)
    Else
        GetOpenWavFileName = GetOpenWavFileNameStd(hWndOwner, strInitDir, strFile, strDialogTitle)
    End If
End Function
