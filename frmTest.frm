VERSION 5.00
Object = "*\AprjFrameXP.vbp"
Begin VB.Form frmTest 
   Caption         =   "FrameXP Custom"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2265
      ScaleWidth      =   7545
      TabIndex        =   14
      Top             =   4440
      Width           =   7575
      Begin prjFrameXP.FrameXP FrameXP8 
         Height          =   1335
         Left            =   3240
         TabIndex        =   17
         Top             =   840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2355
         Caption         =   "FrameXP8"
         BorderColor     =   1747188
         CaptionColor    =   1747188
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check3"
            Height          =   255
            Left            =   2640
            TabIndex        =   21
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check2"
            Height          =   255
            Left            =   1200
            TabIndex        =   20
            Top             =   600
            Width           =   1335
         End
      End
      Begin prjFrameXP.FrameXP FrameXP7 
         Height          =   795
         Left            =   3240
         TabIndex        =   16
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1402
         Caption         =   "FrameXP7"
         BorderColor     =   12632064
         CaptionColor    =   8421376
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "BackColor was missing in first posting. Now it's added"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   3855
         End
      End
      Begin prjFrameXP.FrameXP FrameXP6 
         Height          =   2175
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3836
         Caption         =   "FrameXP6"
         BorderColor     =   16744703
         CaptionColor    =   16711680
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox Text3 
            Height          =   1455
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Text            =   "frmTest.frx":0000
            Top             =   480
            Width           =   2655
         End
      End
   End
   Begin prjFrameXP.FrameXP FrameXP5 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1085
      Caption         =   "How do i exit from here ?"
      BorderColor     =   16777215
      CaptionColor    =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Exit is right here :)  ..........................   >"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   270
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin prjFrameXP.FrameXP FrameXP4 
      Height          =   1575
      Left            =   4200
      TabIndex        =   3
      Top             =   2160
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2778
      Caption         =   "FrameXP4"
      BorderColor     =   12632256
      CaptionColor    =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
   End
   Begin prjFrameXP.FrameXP FrameXP3 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2778
      Caption         =   "Curve size = 25"
      BorderColor     =   49152
      CaptionColor    =   1747188
      CornerCurveSize =   25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Frame Frame1 
         Caption         =   "I'm the original MS frame"
         Height          =   975
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3495
      End
   End
   Begin prjFrameXP.FrameXP FrameXP2 
      Height          =   2055
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3625
      Caption         =   "Curve size = 0"
      BorderColor     =   255
      CaptionColor    =   16711680
      CornerCurveSize =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Font changeable..."
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
   End
   Begin prjFrameXP.FrameXP FrameXP1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3625
      Caption         =   "Caption"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   840
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub
