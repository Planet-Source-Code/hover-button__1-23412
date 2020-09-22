VERSION 5.00
Object = "*\AHButton.vbp"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HoverButton"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   330
   StartUpPosition =   3  'Windows Default
   Begin HoverButton.HButton HButton8 
      Height          =   315
      Left            =   60
      TabIndex        =   20
      Top             =   4260
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      Caption         =   "For comment mail to twanvl@hotmail.com"
      ForeColor       =   0
      HoverColor      =   16711680
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NoFocus         =   -1  'True
      hoverBorder     =   0   'False
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H8000000C&
      Height          =   4155
      Left            =   60
      ScaleHeight     =   4095
      ScaleWidth      =   1335
      TabIndex        =   17
      Top             =   60
      Width           =   1395
      Begin HoverButton.HButton HButton17 
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "Outlook bar"
         ForeColor       =   0
         HoverColor      =   0
         FontSize        =   8,25
         FontName        =   "MS Sans Serif"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoFocus         =   -1  'True
         darkBorder      =   -1  'True
      End
      Begin HoverButton.HButton HButton5 
         Height          =   1095
         Left            =   60
         TabIndex        =   18
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1931
         Caption         =   "Outlook bar"
         BackColor       =   -2147483636
         ForeColor       =   0
         HoverColor      =   0
         FontSize        =   8,25
         FontName        =   "MS Sans Serif"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoFocus         =   -1  'True
         darkBorder      =   -1  'True
         smallBorder     =   -1  'True
         textPos         =   4
         Picture         =   "Hbutton.frx":0000
         OverPicture     =   "Hbutton.frx":0FD2
      End
      Begin HoverButton.HButton HButton6 
         Height          =   1095
         Left            =   60
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1931
         Caption         =   "Outlook bar2"
         BackColor       =   -2147483636
         ForeColor       =   0
         HoverColor      =   0
         FontSize        =   8,25
         FontName        =   "MS Sans Serif"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoFocus         =   -1  'True
         darkBorder      =   -1  'True
         smallBorder     =   -1  'True
         textPos         =   4
         Picture         =   "Hbutton.frx":2454
      End
      Begin HoverButton.HButton HButton7 
         Default         =   -1  'True
         Height          =   1095
         Left            =   60
         TabIndex        =   21
         Top             =   2580
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1931
         Caption         =   "More outlook bar"
         BackColor       =   -2147483636
         ForeColor       =   0
         HoverColor      =   0
         FontSize        =   8,25
         FontName        =   "MS Sans Serif"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoFocus         =   -1  'True
         darkBorder      =   -1  'True
         smallBorder     =   -1  'True
         textPos         =   4
         Picture         =   "Hbutton.frx":355E
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Menu"
      Height          =   1095
      Left            =   3300
      TabIndex        =   13
      Top             =   1380
      Width           =   1575
      Begin HoverButton.HButton HButton3 
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Caption         =   "Menu item1"
         ForeColor       =   0
         HoverColor      =   -2147483634
         FontSize        =   9,75
         FontName        =   "MS Sans Serif"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoFocus         =   -1  'True
         smallBorder     =   -1  'True
         menuStyle       =   -1  'True
      End
      Begin HoverButton.HButton HButton3 
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   15
         Top             =   750
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Caption         =   "Menu item3"
         ForeColor       =   0
         HoverColor      =   -2147483634
         FontSize        =   9,75
         FontName        =   "MS Sans Serif"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoFocus         =   -1  'True
         smallBorder     =   -1  'True
         menuStyle       =   -1  'True
      End
      Begin HoverButton.HButton HButton3 
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   16
         Top             =   490
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Caption         =   "Menu item2"
         ForeColor       =   0
         HoverColor      =   -2147483634
         FontSize        =   9,75
         FontName        =   "MS Sans Serif"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NoFocus         =   -1  'True
         smallBorder     =   -1  'True
         menuStyle       =   -1  'True
      End
   End
   Begin HoverButton.HButton HButton4 
      Height          =   375
      Left            =   3300
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Out"
      ForeColor       =   0
      HoverColor      =   16711680
      FontSize        =   9,75
      FontName        =   "MS Sans Serif"
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NoFocus         =   -1  'True
   End
   Begin HoverButton.HButton HButton1 
      Height          =   615
      Left            =   1560
      TabIndex        =   9
      ToolTipText     =   "SmallBorder=True;  HoverColor=Blue;   Font=Arial;     TextPosition=2"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Test"
      ForeColor       =   -2147483635
      FontSize        =   14,25
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      smallBorder     =   -1  'True
      textPos         =   2
      Picture         =   "Hbutton.frx":49E0
   End
   Begin HoverButton.HButton HButton11 
      Height          =   495
      Left            =   3300
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "TEST"
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      smallBorder     =   -1  'True
      hoverBorder     =   0   'False
   End
   Begin HoverButton.HButton HButton10 
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   3180
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      BackColor       =   -2147483636
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NoFocus         =   -1  'True
      darkBorder      =   -1  'True
      smallBorder     =   -1  'True
      Picture         =   "Hbutton.frx":4E32
      OverPicture     =   "Hbutton.frx":4F8C
      DownPicture     =   "Hbutton.frx":50E6
   End
   Begin HoverButton.HButton HButton9 
      Height          =   375
      Left            =   3420
      TabIndex        =   5
      Top             =   3180
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NoFocus         =   -1  'True
      hoverBorder     =   0   'False
      Picture         =   "Hbutton.frx":55B8
      OverPicture     =   "Hbutton.frx":5712
   End
   Begin HoverButton.HButton HButton13 
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   1620
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Ok"
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HoverButton.HButton HButton14 
      Height          =   555
      Left            =   1560
      TabIndex        =   1
      Top             =   2280
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
      Caption         =   ""
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Hbutton.frx":586C
   End
   Begin HoverButton.HButton HButton15 
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Hi"
      BackColor       =   -2147483636
      HoverColor      =   -2147483635
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      darkBorder      =   -1  'True
   End
   Begin HoverButton.HButton HButton16 
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   3540
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Caption         =   "Disabled"
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin HoverButton.HButton HButton12 
      Height          =   675
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1191
      Caption         =   ""
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NoFocus         =   -1  'True
      smallBorder     =   -1  'True
      hoverBorder     =   0   'False
      Picture         =   "Hbutton.frx":712E
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000C&
      Height          =   495
      Left            =   3300
      ScaleHeight     =   435
      ScaleWidth      =   975
      TabIndex        =   8
      Top             =   3120
      Width           =   1035
   End
   Begin HoverButton.HButton HButton2 
      Height          =   915
      Left            =   3420
      TabIndex        =   10
      Top             =   240
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1614
      Caption         =   "Link"
      ForeColor       =   12582912
      HoverColor      =   16711680
      FontUnderline   =   -1  'True
      FontSize        =   27,75
      FontName        =   "Arial"
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NoFocus         =   -1  'True
      smallBorder     =   -1  'True
      Picture         =   "Hbutton.frx":89F0
      OverPicture     =   "Hbutton.frx":E226
      DownPicture     =   "Hbutton.frx":13A5C
      MousePointer    =   99
      MouseIcon       =   "Hbutton.frx":19292
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   3300
      Picture         =   "Hbutton.frx":193F4
      ScaleHeight     =   1125
      ScaleWidth      =   1500
      TabIndex        =   11
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub HButton3_MouseEnter(Index As Integer)
    HButton3(Index).BackColor = vbHighlight
End Sub

Private Sub HButton3_MouseExit(Index As Integer)
    HButton3(Index).BackColor = vbButtonFace
End Sub

Private Sub HButton3_MouseUp(Index As Integer, ByVal Button As Integer)
    HButton3(Index).Enabled = True
End Sub

Private Sub HButton4_MouseEnter()
    HButton4.Caption = "Over"
End Sub

Private Sub HButton4_MouseExit()
    HButton4.Caption = "Out"
End Sub

Private Sub HButton8_Click()
    ShellExecute Me.hwnd, "open", "mailto:twanvl@hotmail.com", vbNullString, vbNullString, 0
End Sub
