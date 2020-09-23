VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wav Dialog w/Preview"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraWavFileName 
      Caption         =   "Select Wav File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   5655
      Begin VB.PictureBox picPlay 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   4800
         Picture         =   "frmMain.frx":0000
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   5
         Top             =   990
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picStop 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   5085
         Picture         =   "frmMain.frx":0073
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   4
         Top             =   1005
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CheckBox chkPlaces 
         Caption         =   "&Show Places (Win2k Only)"
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   930
         Width           =   2235
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   4620
         TabIndex        =   2
         Top             =   390
         Width           =   900
      End
      Begin VB.TextBox txtWavFileName 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   405
         Width           =   4395
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    Const cstrCaption As String = "Wav Dialog w/Preview"
    Static strLastFile As String
    Dim fShowPlaces As Boolean
    
    fShowPlaces = (chkPlaces.Value = 1)
    txtWavFileName.Text = GetOpenWavFileName(Me.hWnd, strLastFile, strLastFile, "Select Wav", fShowPlaces)
    
    If Trim$(txtWavFileName.Text) <> "" Then
        strLastFile = txtWavFileName.Text
    
        Dim lngPos As Long
        
        lngPos = InStrRev(strLastFile, "\")
        
        If lngPos Then
            Me.Caption = cstrCaption & " - " & Mid$(strLastFile, lngPos + 1)
        Else
            Me.Caption = cstrCaption & strLastFile
        End If
    Else
        Me.Caption = cstrCaption
    End If
End Sub
