VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Win32 Developer Test Pad"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin FlatButton.Flat cmdHorizontal 
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1402
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Flat Button Horizontal"
   End
   Begin FlatButton.Flat cmdVertical 
      Height          =   4965
      Left            =   120
      TabIndex        =   1
      Top             =   990
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   8758
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Flat Button Horizontal"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Main.frx":3452
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   1050
      TabIndex        =   2
      Top             =   1050
      Width           =   4635
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHorizontal_Click()
MsgBox "Hi, I'm a Flat Button with Horizontal text", vbInformation, "Flat Button"
End Sub

Private Sub cmdVertical_Click()
MsgBox "Hi, I'm a Flat Button with Vertical text", vbInformation, "Flat Button"
End Sub

Private Sub Form_Load()
cmdVertical.Caption = "F" & vbCrLf & "l" & vbCrLf & "a" & vbCrLf & "t" & vbCrLf & vbCrLf & "B" & _
                                 vbCrLf & "u" & vbCrLf & "t" & vbCrLf & "t" & vbCrLf & "o" & vbCrLf & "n" & _
                                 vbCrLf & vbCrLf & "V" & vbCrLf & "e" & vbCrLf & "r" & vbCrLf & "t" & vbCrLf & _
                                 "i" & vbCrLf & "c" & vbCrLf & "a" & vbCrLf & "l"
End Sub
