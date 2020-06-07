VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormBW 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Black Winter II : Final Assault"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmOneFrame 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   4200
   End
   Begin VB.Timer tmDelay 
      Left            =   960
      Top             =   4200
   End
   Begin VB.Timer tmMenu 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   4200
   End
   Begin VB.Timer tmOneSecond 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   4200
   End
   Begin VB.Timer tmIntro 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   4200
   End
   Begin VB.PictureBox Pause_Layer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   55
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
      Begin VB.Shape Curr 
         BorderColor     =   &H00C0FFC0&
         Height          =   375
         Index           =   18
         Left            =   3000
         Top             =   120
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H0080FF80&
         Height          =   375
         Index           =   19
         Left            =   3000
         Top             =   240
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H0000FF00&
         Height          =   375
         Index           =   20
         Left            =   3000
         Top             =   360
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H0000C000&
         Height          =   375
         Index           =   21
         Left            =   3000
         Top             =   480
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00008000&
         Height          =   375
         Index           =   22
         Left            =   3000
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00004000&
         Height          =   375
         Index           =   23
         Left            =   3000
         Top             =   720
         Width           =   2175
      End
      Begin VB.Line Dummy_Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         Visible         =   0   'False
         X1              =   2880
         X2              =   2880
         Y1              =   0
         Y2              =   2040
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "QUIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   29
         Left            =   360
         TabIndex        =   59
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "ABORT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   28
         Left            =   360
         TabIndex        =   58
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "RESUME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   27
         Left            =   360
         TabIndex        =   57
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Pause_Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GAME PAUSED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   56
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox Game_Layer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   495
      ScaleWidth      =   480
      TabIndex        =   33
      Top             =   5640
      Visible         =   0   'False
      Width           =   480
      Begin VB.PictureBox Weapon_Box 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   2480
         ScaleHeight     =   2175
         ScaleWidth      =   2250
         TabIndex        =   45
         Top             =   3375
         Width           =   2255
         Begin VB.Label WBox 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BackStyle       =   0  'Transparent
            Caption         =   "DUMMT TEXT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   51
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label WBox 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BackStyle       =   0  'Transparent
            Caption         =   "DUMMT TEXT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   50
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label WBox 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BackStyle       =   0  'Transparent
            Caption         =   "DUMMT TEXT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   49
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label WBox 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BackStyle       =   0  'Transparent
            Caption         =   "DUMMT TEXT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label WBox 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BackStyle       =   0  'Transparent
            Caption         =   "DUMMT TEXT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label WBox 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            BackStyle       =   0  'Transparent
            Caption         =   "DUMMT TEXT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0FF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "ENGINE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   26
         Left            =   360
         TabIndex        =   43
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "GENERATOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   25
         Left            =   360
         TabIndex        =   42
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "SHIELD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   24
         Left            =   360
         TabIndex        =   41
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "SPECIAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   23
         Left            =   360
         TabIndex        =   40
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Game_Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIGURE SHIP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   495
         Index           =   1
         Left            =   840
         TabIndex        =   39
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "SIDE GUN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   22
         Left            =   360
         TabIndex        =   38
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "MAIN GUN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   21
         Left            =   360
         TabIndex        =   37
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "ABORT MISSION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   20
         Left            =   360
         TabIndex        =   36
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "START MISSION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   19
         Left            =   360
         TabIndex        =   35
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Game_Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MISSION LAUNCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   495
         Index           =   0
         Left            =   840
         TabIndex        =   34
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label remap_label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Use Left and Right keys to make changes)"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   44
         Top             =   5520
         Width           =   3615
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00C0FFC0&
         Height          =   375
         Index           =   12
         Left            =   5040
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H0080FF80&
         Height          =   375
         Index           =   13
         Left            =   5040
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H0000FF00&
         Height          =   375
         Index           =   14
         Left            =   5040
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H0000C000&
         Height          =   375
         Index           =   15
         Left            =   5040
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00008000&
         Height          =   375
         Index           =   16
         Left            =   5040
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00004000&
         Height          =   375
         Index           =   17
         Left            =   5040
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Line Dummy_Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         Visible         =   0   'False
         X1              =   4920
         X2              =   4920
         Y1              =   0
         Y2              =   6600
      End
   End
   Begin VB.PictureBox Remap_Layer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   480
      Begin VB.Label remap_label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Use Left and Right keys to make changes)"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   26
         Top             =   4680
         Width           =   3615
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   18
         Left            =   840
         TabIndex        =   25
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "ACCEPT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   17
         Left            =   840
         TabIndex        =   24
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "BGM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   16
         Left            =   840
         TabIndex        =   23
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label remap_label 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "SOUND CONTROLS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   22
         Top             =   3480
         Width           =   3735
      End
      Begin VB.Label remap_label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Press Enter to remap new buttons)"
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   21
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "SE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   15
         Left            =   840
         TabIndex        =   20
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "SPECIAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   14
         Left            =   840
         TabIndex        =   19
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "FIRE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   13
         Left            =   840
         TabIndex        =   18
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "RIGHT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   12
         Left            =   840
         TabIndex        =   17
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "LEFT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   11
         Left            =   840
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "DOWN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   10
         Left            =   840
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "UP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   9
         Left            =   840
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label remap_label 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "KEYBOARD CONTROLS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label BGMm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label BGMp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label SEp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   30
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label SEm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         Top             =   3960
         Width           =   255
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00C0C0FF&
         Height          =   375
         Index           =   6
         Left            =   5040
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H008080FF&
         Height          =   375
         Index           =   7
         Left            =   5040
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H000000FF&
         Height          =   375
         Index           =   8
         Left            =   5040
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H000000C0&
         Height          =   375
         Index           =   9
         Left            =   5040
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00000080&
         Height          =   375
         Index           =   10
         Left            =   5040
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Line Dummy_Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         Visible         =   0   'False
         X1              =   4920
         X2              =   4920
         Y1              =   0
         Y2              =   6480
      End
      Begin VB.Label LabelBGM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "80"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   28
         Top             =   4305
         Width           =   735
      End
      Begin VB.Label LabelSE 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "80"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "SPECIAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Index           =   8
         Left            =   2760
         TabIndex        =   13
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "FIRE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Index           =   7
         Left            =   2760
         TabIndex        =   12
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "RIGHT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Index           =   6
         Left            =   2760
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "LEFT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "DOWN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Index           =   4
         Left            =   2760
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "UP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Index           =   3
         Left            =   2760
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00000040&
         Height          =   375
         Index           =   11
         Left            =   5040
         Top             =   5880
         Width           =   2175
      End
   End
   Begin VB.PictureBox Menu_Layer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5640
      Width           =   480
      Begin VB.PictureBox I_Box 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   4335
         TabIndex        =   87
         Top             =   2640
         Width           =   4335
         Begin VB.Shape Intro_Line 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   60
            Index           =   0
            Left            =   120
            Top             =   120
            Width           =   4095
         End
      End
      Begin VB.PictureBox I_Box 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   4335
         TabIndex        =   86
         Top             =   1200
         Width           =   4335
         Begin VB.Shape Intro_Line 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   60
            Index           =   1
            Left            =   120
            Top             =   120
            Width           =   4095
         End
      End
      Begin VB.PictureBox I_Box 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   1
         Left            =   3840
         ScaleHeight     =   1455
         ScaleWidth      =   855
         TabIndex        =   81
         Top             =   1200
         Width           =   855
         Begin VB.Label Title 
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   72
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1455
            Index           =   2
            Left            =   0
            TabIndex        =   82
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox I_Box 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   855
         Index           =   0
         Left            =   960
         ScaleHeight     =   855
         ScaleWidth      =   3015
         TabIndex        =   80
         Top             =   1200
         Width           =   3015
         Begin VB.Label Title 
            BackStyle       =   0  'Transparent
            Caption         =   "BLACK"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Index           =   0
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   2895
         End
      End
      Begin VB.PictureBox I_Box 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   2
         Left            =   480
         ScaleHeight     =   975
         ScaleWidth      =   3495
         TabIndex        =   84
         Top             =   1800
         Width           =   3495
         Begin VB.Label Title 
            BackStyle       =   0  'Transparent
            Caption         =   "WINTER"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   855
            Index           =   1
            Left            =   0
            TabIndex        =   85
            Top             =   0
            Width           =   3375
         End
      End
      Begin VB.Label Wait_Label 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Loading ..."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "START"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   78
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Created By Yat Seng 2003 (ruby@starpulse.com)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   69
         Top             =   6240
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label Title 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Final Assault"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   3
         Left            =   1080
         TabIndex        =   54
         Top             =   2880
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Line Dummy_Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         Visible         =   0   'False
         X1              =   4920
         X2              =   4920
         Y1              =   0
         Y2              =   6480
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "QUIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   2
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label menu 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "OPTIONS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   5040
         TabIndex        =   1
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00C0C0FF&
         Height          =   375
         Index           =   0
         Left            =   5040
         Top             =   120
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H008080FF&
         Height          =   375
         Index           =   1
         Left            =   5040
         Top             =   240
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   5040
         Top             =   360
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H000000C0&
         Height          =   375
         Index           =   3
         Left            =   5040
         Top             =   480
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00000080&
         Height          =   375
         Index           =   4
         Left            =   5040
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Curr 
         BorderColor     =   &H00000040&
         Height          =   375
         Index           =   5
         Left            =   5040
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.PictureBox Play_Layer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   53
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
      Begin VB.PictureBox BossBar 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   1215
         TabIndex        =   114
         Top             =   1080
         Width           =   1215
         Begin VB.Shape Boss_Bar_Top 
            BackColor       =   &H000000FF&
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   105
            Left            =   0
            Top             =   240
            Width           =   1215
         End
         Begin VB.Shape Boss_Bar_Bottom 
            BackColor       =   &H000000FF&
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   100
            Left            =   0
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Boss_Name 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "TRANSPORT"
            ForeColor       =   &H0000FF00&
            Height          =   210
            Left            =   0
            TabIndex        =   115
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox Bufferbox 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   3840
         ScaleHeight     =   255
         ScaleWidth      =   1695
         TabIndex        =   101
         Top             =   1080
         Width           =   1695
         Begin VB.Label BufferText 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "WILD CAT MISSILES"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   102
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.PictureBox Bufferbox 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   840
         ScaleHeight     =   225
         ScaleWidth      =   1455
         TabIndex        =   98
         Top             =   1080
         Width           =   1455
         Begin VB.Label BufferText 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "000000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   100
            Top             =   0
            Width           =   855
         End
         Begin VB.Label BufferText 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "SCORE"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox Bufferbox 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   6000
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   96
         Top             =   2880
         Width           =   1095
         Begin VB.Label BufferText 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   ">Sample Text"
            ForeColor       =   &H0000FF00&
            Height          =   210
            Index           =   4
            Left            =   0
            TabIndex        =   97
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.PictureBox Bufferbox 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   6000
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   91
         Top             =   2640
         Width           =   1095
         Begin VB.Label BufferText 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   ">Sample Text"
            ForeColor       =   &H0000FF00&
            Height          =   210
            Index           =   3
            Left            =   0
            TabIndex        =   95
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.PictureBox Bufferbox 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   6000
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   90
         Top             =   2400
         Width           =   1095
         Begin VB.Label BufferText 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   ">Sample Text"
            ForeColor       =   &H0000FF00&
            Height          =   210
            Index           =   2
            Left            =   0
            TabIndex        =   94
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.PictureBox Bufferbox 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   6000
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   89
         Top             =   2160
         Width           =   1095
         Begin VB.Label BufferText 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   ">Sample Text"
            ForeColor       =   &H0000FF00&
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   93
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.PictureBox Bufferbox 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   6000
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   88
         Top             =   1920
         Width           =   1095
         Begin VB.Label BufferText 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   ">Sample Text"
            ForeColor       =   &H0000FF00&
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   92
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.PictureBox Stats_Bar 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFF00&
         Height          =   975
         Left            =   720
         ScaleHeight     =   975
         ScaleWidth      =   4935
         TabIndex        =   60
         Top             =   7920
         Visible         =   0   'False
         Width           =   4935
         Begin VB.PictureBox Stt_SGUN_R 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3720
            ScaleHeight     =   255
            ScaleWidth      =   615
            TabIndex        =   76
            Top             =   720
            Width           =   615
            Begin VB.Label lbl_SGUN 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "SGUN"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Left            =   0
               TabIndex        =   77
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.PictureBox Stt_MGUN_R 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3720
            ScaleHeight     =   255
            ScaleWidth      =   615
            TabIndex        =   74
            Top             =   480
            Width           =   615
            Begin VB.Label lbl_MGUN 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "MGUN"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Left            =   0
               TabIndex        =   75
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.PictureBox Stt_ENG_R 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4320
            ScaleHeight     =   255
            ScaleWidth      =   615
            TabIndex        =   72
            Top             =   480
            Width           =   615
            Begin VB.Label lbl_ENG 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "ENG"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Left            =   0
               TabIndex        =   73
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.PictureBox Stt_PWR_R 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4320
            ScaleHeight     =   255
            ScaleWidth      =   615
            TabIndex        =   70
            Top             =   720
            Width           =   615
            Begin VB.Label lbl_PWR 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               Caption         =   "PWR"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Left            =   0
               TabIndex        =   71
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Shape BAR_SHD_C 
            FillColor       =   &H00FFFF00&
            FillStyle       =   0  'Solid
            Height          =   75
            Left            =   540
            Top             =   30
            Width           =   2175
         End
         Begin VB.Shape BAR_SPE_R 
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   75
            Left            =   540
            Top             =   330
            Width           =   2175
         End
         Begin VB.Shape BAR_SPE_C 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   75
            Left            =   540
            Top             =   270
            Width           =   2175
         End
         Begin VB.Shape BAR_SHD_R 
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   75
            Left            =   540
            Top             =   90
            Width           =   2175
         End
         Begin VB.Label Stt_SPE_Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "SPE"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   0
            TabIndex        =   62
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Stt_SHD_Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "SHD"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   0
            TabIndex        =   61
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Stt_PWR_F 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            Caption         =   "PWR"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4320
            TabIndex        =   66
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Stt_ENG_F 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            Caption         =   "ENG"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4320
            TabIndex        =   65
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Stt_SGUN_F 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            Caption         =   "SGUN"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3720
            TabIndex        =   68
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Stt_MGUN_F 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            Caption         =   "MGUN"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3720
            TabIndex        =   67
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Stt_LVL_SG 
            BackColor       =   &H00000000&
            Caption         =   "SG   1.0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   2880
            TabIndex        =   64
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Stt_LVL_MG 
            BackColor       =   &H00000000&
            Caption         =   "MG  1.0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   2880
            TabIndex        =   63
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Line Dummy_Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         Visible         =   0   'False
         X1              =   5640
         X2              =   720
         Y1              =   7920
         Y2              =   7920
      End
      Begin VB.Line Dummy_Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         Visible         =   0   'False
         X1              =   5640
         X2              =   720
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Dummy_Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         Visible         =   0   'False
         X1              =   720
         X2              =   720
         Y1              =   1440
         Y2              =   7920
      End
      Begin VB.Label Wait_label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Loading ..."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5880
         TabIndex        =   79
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Line Dummy_Line 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         Visible         =   0   'False
         X1              =   5655
         X2              =   5655
         Y1              =   1440
         Y2              =   7920
      End
   End
   Begin ComctlLib.ImageList Img 
      Index           =   31
      Left            =   4200
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   82
      ImageHeight     =   116
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   23
      Left            =   1800
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   161
      ImageHeight     =   154
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":70B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   30
      Left            =   1200
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   61
      ImageHeight     =   57
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1942C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1BD76
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1E6C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2100A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   29
      Left            =   600
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   41
      ImageHeight     =   38
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":23954
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":24C0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":27182
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   28
      Left            =   0
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   19
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2843C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2894E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":28E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":29372
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   3
      Left            =   2400
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   256
      ImageHeight     =   192
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   22
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":29884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":35CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":42128
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":4E57A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":5A9CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":66E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":73270
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":7F6C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":8BB14
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":97F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":A43B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":B080A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":BCC5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":C90AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":D5500
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":E1952
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":EDDA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":FA1F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":106648
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":112A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":11EEEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":12B33E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   27
      Left            =   1800
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   160
      ImageHeight     =   360
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":137790
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   26
      Left            =   1200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   34
      ImageHeight     =   25
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   15
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":161AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":16255C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":162FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":163A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1644CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":164F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1659BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":166438
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":166EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":16792C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1683A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":168E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":16989A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":16A314
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":16AD8E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MediaPlayerCtl.MediaPlayer MP6 
      Height          =   375
      Left            =   2400
      TabIndex        =   113
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MP4 
      Height          =   375
      Left            =   1440
      TabIndex        =   112
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MPS 
      Height          =   375
      Index           =   7
      Left            =   3360
      TabIndex        =   111
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MPS 
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   110
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MPS 
      Height          =   375
      Index           =   5
      Left            =   2400
      TabIndex        =   109
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MPS 
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   108
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MPS 
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   107
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MPS 
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   106
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MPS 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   103
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MPS 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   105
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MP5 
      Height          =   375
      Left            =   1920
      TabIndex        =   104
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin ComctlLib.ImageList Img 
      Index           =   25
      Left            =   3600
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   37
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":16B808
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   24
      Left            =   3000
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   158
      ImageHeight     =   145
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":16C9B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   22
      Left            =   2400
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   55
      ImageHeight     =   69
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":17D7A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   21
      Left            =   1800
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   70
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":18053A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   20
      Left            =   1200
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   49
      ImageHeight     =   68
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":182F1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   19
      Left            =   600
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   144
      ImageHeight     =   103
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1856BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   18
      Left            =   0
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   125
      ImageHeight     =   125
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1904E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   17
      Left            =   1200
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   60
      ImageHeight     =   60
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":19BCCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":19E74C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A11CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A3C50
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A4EB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   16
      Left            =   600
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A7934
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A7E36
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   15
      Left            =   0
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A8338
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A84CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   14
      Left            =   2400
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   40
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A865C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   13
      Left            =   1200
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A8E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A9150
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A9472
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A9794
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A9AB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   12
      Left            =   1800
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1A9DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1AA0FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   11
      Left            =   1200
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   30
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1AA41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1AAF36
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1ABA50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1AC56A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   10
      Left            =   600
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   130
      ImageHeight     =   130
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1AD084
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1B97E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1C5F48
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1D26AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1DEE0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1EB56E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":1F7CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":204432
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":210B94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   9
      Left            =   0
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   45
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":21D2F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   8
      Left            =   1800
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   120
      ImageHeight     =   70
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":21D8E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":223BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":229E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":23012E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2363F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":23C6B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":242974
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":248C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":24EEF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2551BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   7
      Left            =   1200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   24
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25B47C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25B67E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25B880
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25BA82
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25BC84
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25BE86
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25C088
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25C28A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25C48C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25C68E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25C890
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25CA92
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25CC94
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25CE96
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25D098
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25D29A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25D49C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25D69E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25D8A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25DAA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25DCA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25DEA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25E0A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25E2AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   6
      Left            =   600
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   50
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25E4AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25F17E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":25FE50
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":260B22
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2617F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2624C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":263198
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":263E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":264B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":26580E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   5
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   60
      ImageHeight     =   60
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2664E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":268F62
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":26B9E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":26E466
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   4
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":270EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2710A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2712C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2714EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2717A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":271A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":271C64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   2
      Left            =   600
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   256
      ImageHeight     =   192
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   30
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":271E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":27E2B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":28A70A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":296B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2A2FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2AF400
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2BB852
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2C7CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2D40F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2E0548
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2EC99A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":2F8DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":30523E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":311690
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":31DAE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":329F34
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":336386
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3427D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":34EC2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":35B07C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3674CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":373920
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":37FD72
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":38C1C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":398616
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3A4A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3B0EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3BD30C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3C975E
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3D5BB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   1
      Left            =   0
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   71
      ImageHeight     =   100
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3E2002
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3E4074
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3E60E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3E8158
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3EA1CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3EC23C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3EE2AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3F0320
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3F2392
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3F4404
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3F6476
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3F84E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3FA55A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3FC5CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":3FE63E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":4006B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Img 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   49
      ImageHeight     =   38
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":402722
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":403D6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FormBW.frx":4053B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MediaPlayerCtl.MediaPlayer MP3 
      Height          =   375
      Left            =   960
      TabIndex        =   52
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MP2 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MP1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "FormBW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'*********                                           ********
'*********      BLACK WINTER II : Final Assault      ********
'*********         Created by : Wong Yat Seng        ********
'*********                                           ********
'************************************************************

'If you've downloaded this from Planet-Source-Code, votes
'would be appreciated. =)

'Credits:
'   Credits are extended to the following for certain elements
'   borrowed from their submissions on Planet Source Code
'       -   Johnathan Roach (Space Project Demo)
'           -   Aircraft Sprite
'           -   Starfield Animation
'           -   ImageList Usage (masking transparency)
'
'       -   Martin Castaneda (Air-X 32)
'           -   Aircraft Movement Controls

'Note:
'   The following is a tutorial guide on how to make
'   an intermediate space shooter game. If you do not wish
'   to learn about space shooters, then just play the game
'   and dont bother yourself reading the 5000+ lines of codes
'   below.

'   I've created Black Winter2(BW2) using the simplest of
'   elements found in vb. It is meant to teach ppl how to
'   make a space shooter, without all the complex stuffs.
'   Everything 's done in one single form: No external calls,
'   no global variables (just public), no API calls, no
'   references to any .dll files, blah blah blah...

'   It should've been done in DirectDraw, but learning how
'   to reference to DD would kill beginners outright, so we'll
'   have to stick with the laggy graphic time being.

'Tutorial:
'   So anyhow, before you begin to create a space shooter,
'   you'll need a background, a picture box where all you
'   spaceships gonna blow each other up.

'   In BW2, i've created 5 such picture boxes, all are named
'   different layers. Most of the layers are meant for menus.
'   The layer for the background is named Play_Layer.

'   Enlarge the the form, look for one of the small black boxes
'   (there should be 5 of em, one of them is red) named Play_Layer
'   and enlarge it. Dun worry about messing with it's size because
'   i've set the form resize in the Form_Load sub.

'   You'll notice that this Layer has a white color box around
'   it. It wont be visible during play, but the area within the
'   white box is the actual play area visible to the player.
'   All the other bars outside the white lines are the statsbar
'   of the player representing his score, shields, vital stats,
'   etc... These bars themselves are contained within picture
'   boxes themselve. Moving the picture box will affect all the
'   items and bars inside it. So you must take note that when
'   u want to put a control on to a layer, click on the control
'   and DRAG (draw) on the picture box, NOT double-click from
'   the toolbar. This way, the controls belongs to the picture
'   box, not the form.

'   Okay... now that you've got a background, u'll need sprites.
'   Take note that all the sprites can be found inside the
'   ImageList controls name Img(index). There's a lot of these
'   on my form, u cant miss that. The reason why i use imagelist
'   to store my picture instead of the regular Image control is
'   because i need to 'paint' the picture ONTO the picturebox
'   background, rather just setting the Image control's top and
'   left properties to its coordinates.

'   Reason for using Imagelist compared to Image control is:
'       1.  Color Masking are possible using imagelist. You
'           can make the background of your sprite invisible
'           by setting the mask color in the ImageList property
'           to match it.
'       2.  It is MUCH easier to use compared to the BitBlt method
'           commonly used because there's no need to declare and
'           referencing to gdi32 (much easier for beginners to learn).
'       3.  Painting of pictures are a LOT faster compared to moving
'           an Image control.
'       4.  Hmm... there maybe many other benefits but i'm lazy
'           to think now... i guess the ones above would suffice.

'   After collecting all your sprites and putting them into
'   imglists, the next thing u need to do is to craft your objects =)

'   First thing u need is a ship (YOUR ship). I hope u've already
'   got a sprite for that. U must determine what kind of
'   characteristics your ship should have. Look at the variables the
'   i've delared below (look for *** SHIP STATS ***).

'   I've listed ShipX and ShipY. These two variable will store
'   the X and Y coordinates of your ship. It is important for
'   the imagelist to know exactly where to paint the sprite
'   on the background.

'   Ship_Speed is the amount of twips (15 twips = 1 pixel) a
'   ship will move per frame. the higher this value, the faster
'   your ship move (see below).

'   Next is Ship_Moving(3). This array has 4 boolean values.
'   each value represents the 4 directions a ship can move, e.g.
'   0-Left, 1-Up, 2-Right, 3-Down. A true value to a certain
'   direction will change the X and Y coordinates to move to
'   that direction every turn. E.g:
'               Ship_Moving(0) = True
'               Ship_Moving(3) = True           'Ship is moving Left and Down
'
'               ShipX = ShipX - Ship_Speed
'                                               '(Ship moves left, X coordinate gets closer to 0)
'               ShipY = ShipY + Ship_Speed
'                                               '(Ship moves down, Y coordinate gets further from 0)

'   Thus the result would be the coordinate of the ship moving
'   diagonally in a 225 degrees direction. Simple enuff?

'   For coding, u'll need to set the Ship_Moving boolean value
'   to true when a KeyDown event for the direction is presses,
'   and set it to false when a KeyUp for its button is release.
'   This will help eliminate the short lag when moving if u
'   use the Keypress event for your movements.
'       REFERENCE: see Private Sub Play_Layer_Keydown / KeyUp

'   Same goes for the Ship_Firing, if its true then bullets
'   will be shot from the ship. If KeyUp detected, firing will stop.

'   There are many other stats i've made in BW2 (shields, special
'   weapons, ship parts damage, repair speed, etc...) but i'll
'   explain only the most basic elements a ship needs. Feel
'   free to look up on my codes to learn about other stats.

'   Now that you've got your basic ship configured, its time to
'   configure its bullets. A bullet object is created.
'   (see *** BULLET STATS *** below)

'   I've declared an array of 300 Bullet objects. Actually, this
'   is a bad practice as these bullet objects already takes up
'   memory space before being used. A better way to do it would
'   be declaring a new object array everytime u need a new array.
'   Example: Dim a new bullet object into an upper bound (UBound)
'   of an exisitng array.
'   All this might be a bit confusing for programmers not familiar
'   with such methods, as of now we'll stick to our Dim bull(300)
'   in this tutorial

'   The bullet object contains a Used boolean attribute to indicate
'   whether the bullet should be painted on the Background. By
'   default, all bullets are Used=False.

'   The PosX ans PosY are similar to the ShipX and ShipY used
'   for our ship just now. Spd (Speed) is the same. Dir (Direction)
'   is the direction the bullet flies when fired.
'   Dama (Damage) is how much life an enemy gets deducted when
'   hit by this bullet. Kind determine wat the bullet looks like.
'   It is actually a reference to an ImageList index from whence its
'   bmp picture derive from. Span is a counter to keep track of
'   the duration a bullet has been in play since it's fired. This
'   is useful for certain bullets that may explode after a certain
'   duration.

'   Now that u've got a bullet, u need to animate the bullet flying from
'   the ship. Usually, u'll need a timer to achieve this.
'   Notice that i've create a very VERY big procedure called
'   tmOneFrame way way below. This timer draws out all the
'   sprites in the game in every single frame of its interval.
'   In other words, it is the most IMPORTANT control in the
'   entire game as there wont be any movements at all if this
'   timer is stopped (usually used only in paused mode)

'   At the top of tmOneFrame is a command called Play_Layer.Cls.
'   This command empties out the entire playing area. New pictures
'   are then painted on the background in different coordinates,
'   thus depicting a movement is perform in succession.
'   the syntax to paint a sprite on the background is:

'           ImgLst.ListImages(1).Draw PicBox.hDC, ShipX, ShipY, 1

'   ImgLst      = ImageList name
'   (1)         = The index of the picture in the list.
'   Picbox      = Background Picturebox name
'   ShipX/Y     = The coordinate of the picture to be painted.
'   Dont change the rest of the syntax as they are necessary.

'   So within this tmOneFrame sub that u insert the coordinate
'   calculations for moving ships and bullets and paint syntax
'   above. Do a loop that modifies the bullet's PosX and PosY
'   inside the timer and you'll see a moving bullet painted
'   in the background.

'   Your next step is to create enemies for your bullets to fry =).
'   Your'll to create an enemy object (see *** ENEMY SHIPS *** below)
'   Most of the stats are similar as your ships' stats, with exceptions
'   for:
'           Patt        :   Because Enemy ships movements are
'                           not controlled by the player, they
'                           need to move in a predetermined order
'                           fixed by its Patt property defined
'                           under tmOneFrame
'           DestX/Y     :   If Patt above is not used, the enemy
'                           ship will move towards its Destination
'                           X and Y coordinates.
'           Life        :   This is the amount of damage a ship
'                           receives before it is destroyed.
'           Firing_Patt :   As Patt above, firing patterns are
'                           also predetermined by the program.
'           Has_Bonus   :   This will indicate whether a weapon
'                           upgrade will be create in the position
'                           of the enemy destroyed.
'           Boss        :   If this boolean value is true, certain
'                           events will be trigger, such as Bossbar,
'                           pause incoming enemies, etc.

'   With an enemy ship created, u'll need to move it according to
'   its Movement Pattern. Coordinate changes are similar to player's
'   ship except that movements are preset rather than based on
'   KeyDown events. Refer to tmOneFrame and look under moving Enemy ship.

'   Next you'll need to perform collision detecion of your bullets
'   with your enemies.
'   U'll 1st need to know the rectangular border of your bullet
'   image, and the same for the enemy.
'   assuming you've Dim your Bullet object as Bull, and its picture
'   is stored in ImgLstBull frame 1.

'   Bull.PosY                               is your top border.
'   Bull.PosX                               is your left border.
'   Bull.PosY + ImgLstBull.ImageHeight      is your bottom border.
'   Bull.PosX + ImgLstBull.ImageWidth       is your right border.

'   The same borders of the enemy's ship is also needed to be
'   identified, thus

'   EShip.PosY
'   EShip.PosX
'   EShip.PosY + ImgLstEnemy.ImageHeight
'   EShip.PosX + ImgLstEnemy.ImageWidth

'   Now with all the borders, you can now perform an image
'   collision detection to check whether either of these
'   rectabgles are overlapping each other.
          
'       If Bull.PosX < EShip.PosX + ImgLstEnemy.ImageWidth AND_     'Checks for left side of Bullet and right side of enemy
'          Bull.PosX + ImgLstBull.ImageWidth > EShip.PosX AND_      'Checks for right side of Bullet and left side of enemy
'          Bull.PosY < EShip.PosY + ImgLstEnemy.ImageHeight AND_    'Checks for top side of Bullet and bottom side of enemy
'          Bull.PosY + ImgLstBull.ImageHeight > Eship.PosY THEN ... 'Checks for bottom side of Bullet and top side of enemy

'   If all the conditions is satisfied, then a collision is detected.
'   The neccesary effect below will be inserted in the THEN ... part.

'       1.  Reduce the enemy's health based on bullet's damage.
'           E.g:    EShip.Life = EShip.Life - Bull.Dama

'       2.  Make the Bullet dissapear from the game after a hit
'           E.g:    Bull.Used = False
'           Note - You only paint the bullets in your timer if
'                  the Used value of your bullet is true.

'       3.  Make a check on whether enemy's life fall below zero.
'           E.g:    If Eship.Life <= 0 Then EShip.Used = False

'   With the collision algorithm above, you should be able to
'   do the same for enemy bullets hitting your ship and collision
'   between 2 ships.

'   Now that you should have a basic game engine done, u'll need
'   to design and create some game levels. As mentioned earlier,
'   tmOneFrame is the timer that controls almost everything, so
'   u will just need to add a counter to keep track of the number
'   of times the timer has loop. Specifying events at fixed
'   counter intervals will help make your level more interesting.
'   With careful design and placement of ships, u'll get to
'   deploy enemy formations, their firing pattern and movement
'   pattern in accordance to the time counter used to keep
'   track of the events.

'   In BW2, a Game Event Counter (GEC) is used to determined
'   when events will take place.
'   Game Level designs are located in the GEC coding section located
'   near the end of the tmOneFrame sub.

'   Well... that basically wraps everything else this tutorial
'   needs to cover. Feel free to mess around with my codes to
'   learn more (highly recommended).

'   I do apologize for some badly structured codes and the
'   mass amount of public declarations i made. Kinda confusing
'   huh?

'   Anyhow, should there be further enquiries about my code,
'   please email me at:     ruby@starpulse.com
'   please do include your name.




'   PS:     Hehe... If you've downloaded this source from
'           Planet Source Code, please do vote for me if
'           you find this tutorial useful

Option Explicit

'-----------MENU MODE VARIABLES---------------

Dim Curr_Pos As Integer         'Current Position
Dim Curr_Start As Integer       'Cursor array start index
Dim Curr_End As Integer         'Cursor array end index
Dim Curr_From As Integer        'Menu item start array index
Dim Curr_To As Integer          'Menu item end array index

Dim Bal(23) As Integer          'Cursor's tracking usage (modifies coordinates)

Dim R_Color As Integer          'RED       Intro's RGB Colors
Dim G_Color As Integer          'GREEN
Dim B_Color As Integer          'BLUE

Dim KYB(5) As Integer           'Keyboard Controls      (0-Up  1-Down  2-Left  3-Right  4-Fire  5-Special)
Dim SHP(5) As Integer           'Ships Configuration    (0-Maingun  1-Sidegun  2-Special  3-Shield  4-Generator  5-Engine)
Dim Vol(1) As Integer           'Volume Controls        (0-Sound Effect  1-Background Music)

Dim MenuMode_Enabled As Boolean     'Mainscreen selection
Dim GameMode_Enabled As Boolean     'Game Mode (configure ship)
Dim PlayMode_Enabled As Boolean     'Playing Mode
Dim RemapMode_Enabled As Boolean    'Remap Keyboard/options
Dim PauseMode_Enabled As Boolean    'Pause During Game

Dim WaitForRemap As Boolean         'Detect Key Press during remap buttons
Dim MenuSelectLR As Boolean         'Used for Horinzontal selection

Dim Intro_Counter As Integer        'For Intro

Dim Delay_Counter As Integer        'Delay Count    (Used in WAIT procedure)
Dim Delay_Occasion As Integer       'Event for Delay

Dim Elasped As Long                 'Tracked Elasped time in FPS
Dim FPS As Integer                  'Frames Per Second
Dim TotFPS As Long                  'Total Accumulated Frames
    
'-----------PLAY MODE VARIABLES---------------

'   *** BULLET STATS ***
Private Type Bullet     'THE BULLET OBJECT
    Used As Boolean     'Availability
    Kind As Integer     'Type of Bullet
    PosX As Integer     'X Coordinate
    PosY As Integer     'Y Coordinate
    Dire As Integer     'Type of direction
    Span As Integer     'Time span bullet since shot
    Dama As Integer     'Damage it does
    Spd As Integer      'Bullet Speed
End Type

Private Type Explosion          'THE EXPLOSION OBJECT
    Used As Boolean             'Availability
    Kind As Integer             'Type of explosion
    PosX As Integer             'X Coordinate
    PosY As Integer             'Y Coordinate
    Current_Frame As Integer    'Displaying Current Frame
    Last_Frame As Integer       'Last frame in sequence
End Type

Private Type Bonuses        'BONUS OBJECT
    Used As Boolean         'Availability
    PosX As Integer         'X Coord
    PosY As Integer         'Y Coord
    Frame As Integer        'Frames
End Type

Private Type Starfield          'Credit goes to Johnathan Roach
    PosX As Integer             'Coord X
    PosY As Integer             'Coord Y
    Spd As Integer              'Speed
    Color As Integer            'Dimness
End Type

Const Num_Stars = 1000
Dim Star(Num_Stars) As Starfield                   'Stars
Dim Detail_Level As Integer                        'Custom detail level
Dim PowerUp(2) As Bonuses                          'Bonus

Dim Lock_All As Boolean                'Ship LOCK
Dim Lock_MainGun As Boolean            'Main Gun LOCK
Dim Lock_SideGun As Boolean            'Side Gun LOCK
Dim Lock_Generator As Boolean          'Generator LOCK
Dim Lock_Engine As Boolean             'Engine LOCK

'   *** SHIP STATS ***
Dim ShipX As Integer                 'Ship X Coordinate
Dim ShipY As Integer                 'Ship Y Coordinate
Dim Ship_Dir As Integer              'Direction
Dim Ship_Speed As Integer            'Engine Speed
Dim Ship_Moving(3) As Boolean        'Moving towards direction
Dim Ship_Firing As Boolean           'Firing Weapon

Dim Repair_Speed As Integer         'Rate of repair
Dim SHD_Charge_Speed As Integer     'Shield charge rate
Dim SHD_Charge_Base As Integer      'Base Charge Rate
Dim SPE_Charge_Speed As Integer     'Special charge rate
Dim SHD_Avail As Boolean            'Availability of shield
Dim SPE_Avail As Boolean            'Availability of special
Dim SHD_Charging As Boolean         'Charge ON/OFF
Dim SPE_Charging As Boolean         'Charge ON/OFF
Dim SHD_Width As Integer            'Shield Level
Dim SPE_Width As Integer            'Special Level

Dim SHIELD_ON_OFF As Integer        'Display SHD Visual
Dim Special_Running As Boolean      'Runs Special in OneFrameTimer

Dim MG_Lvl As Integer   'Main Gun Bullet Tech Level
Dim SG_Lvl As Integer   'Side Gun Bullet Tech Level

Dim Bull(300) As Bullet             'Bullets (Objects)
Dim Bull_Type_Limit As Integer      'Max bullet type on screen
Dim Bull_Delay As Integer           'Delay between MG
Dim Bull_Delay2 As Integer          'Delay between SG

Dim Special_Delay As Integer        'Delay for special's display

Dim Explo(40) As Explosion          'Explosion (Objects)
Dim ShakeIT As Integer              'Shake Screen During Major/Critical HIT
Dim Giga_Count As Integer           'Count for Giga Storm

Dim Moving_SpecialBar As Integer    'Indicator for special attack bar moving
Dim Moving_StatsBar As Integer      'Indicator for stat's bar moving
Dim Moving_BossBar As Integer       'Indicator for boss bar moving
Dim Buffering As Boolean            'To inform buffering session
Dim Buffer_Count As Integer         'To signal Mover
Dim Buffer_Mover As Boolean         'To Move text while trye
Const Buff_Speed = 30               'Text Buffer Speed

Dim KillCount As Integer            'Total enemies killed
Dim KilledBoss As Boolean           'Whether last boss is killed
Dim GameScore As Long               'Total Game Score
Dim GameStarted As Boolean          'See if Game is in progress
Dim GEC As Long                     'Game Counter Event
Dim C_GEC(4) As Long                'Cumulative GEC
Dim OVERIDE As Boolean              'Overide Total Controls

'-----------ENEMY VARIABLES---------------

Private Type EBullet        'THE ENEMY BULLET OBJECT
    Used As Boolean         'Availability
    Kind As Integer         'Type of Bullet (Either 11, 12, 21, 22)
    StartX As Integer       'Start X Coord
    StartY As Integer       'Start Y Coord
    PosX As Integer         'X Coordinate
    PosY As Integer         'Y Coordinate
    Dire As Integer         'Type of direction
    Span As Long            'Time span bullet since shot
    Spd As Integer          'Bullet Speed
    DestX As Integer        'Destination X
    DestY As Integer        'Destination Y
End Type

'   *** ENEMY SHIPS ***
Private Type EnemyShip      'THE ENEMY SHIP OBJECT
    Used As Boolean         'Availability
    Kind As Integer         'Image List Index Number
    Bul_Kind As Integer     'Bullet Kind
    Frame As Integer        'Image Index
    BullType As Integer     'Bullets Shot
    PosX As Integer         'Current X Coordinate
    PosY As Integer         'Current Y Coordinate
    DestX As Integer        'Destination X Coordiante
    DestY As Integer        'Destination Y Coordiante
    Span As Integer         'For Event counter
    Patt As Integer         'Movement Pattern
    Spd As Integer          'Move Speed
    Life As Integer         'Hit Points
    Max_Life As Integer     'For BOSS's Life Bar Calculation
    Score As Integer        'Worth of Points
    Has_Bonus As Boolean    'Whether Ship has BoNUS =)
    Firing_Patt As Integer  'Firing Pattern
    Boss As Boolean         'Whether it's a boss or not
End Type

Dim EBull(200) As EBullet
Dim EShip(40) As EnemyShip      '38,39,40 for BOSS

Private Sub Form_GotFocus()         'Returns Focus to picboxes
'If the player accidentally clicked the mouse on the form,
'focus is returned to the current layer, so that keydown events
'are still acknowledged.
If MenuMode_Enabled = True Then
    Menu_Layer.SetFocus
End If
If RemapMode_Enabled = True Then
    Remap_Layer.SetFocus
End If
If GameMode_Enabled = True Then
    Game_Layer.SetFocus
End If
If PlayMode_Enabled = True Then
    Play_Layer.SetFocus
End If
If PauseMode_Enabled = True Then
    Pause_Layer.SetFocus
End If
End Sub

Private Sub Form_Load()
'This sub basically set's the layers, menus, labels, bars,
'to their correct position. Dun mess this part.
Dim i As Integer
    Wait_label2.Visible = True
    Wait_label2.Top = 4320
    Wait_label2.Left = 2520
    
    FormBW.BackColor = &H80000012
    
    FormBW.Width = 5010         'Set Form Size
    FormBW.Height = 6990        '
    Menu_Layer.Width = 4920     '
    Menu_Layer.Height = 6490    '
    Menu_Layer.Left = 0         '
    Menu_Layer.Top = 0          '
    Remap_Layer.Width = 4920    '
    Remap_Layer.Height = 6490   '
    Remap_Layer.Left = 0        '
    Remap_Layer.Top = 0         '
    Game_Layer.Width = 4920     '
    Game_Layer.Height = 6490    '
    Game_Layer.Left = 0         '
    Game_Layer.Top = 0          '
    Play_Layer.Width = 6375     '
    Play_Layer.Height = 8655    '
    Play_Layer.Left = -720      '
    Play_Layer.Top = -1440      '
    Pause_Layer.Width = 2895    '
    Pause_Layer.Height = 2055   '
    Pause_Layer.Left = 1080     '
    Pause_Layer.Top = 2040      '
    
    'Set the speed optimization
    tmMenu.Interval = FormOptimize.Sldr_ms.Value
    tmOneFrame.Interval = FormOptimize.Sldr_ms.Value
    tmIntro.Interval = FormOptimize.Sldr_ms.Value * 10
    tmMenu.Interval = FormOptimize.Sldr_ms.Value
    If FormOptimize.Opt_Detail(0).Value = True Then Detail_Level = 50
    If FormOptimize.Opt_Detail(1).Value = True Then Detail_Level = 400
    If FormOptimize.Opt_Detail(2).Value = True Then Detail_Level = 1000
    
    'Sets the currently enabled mode to MENU
    MenuMode_Enabled = True
    GameMode_Enabled = False
    RemapMode_Enabled = False
    PlayMode_Enabled = False
    PauseMode_Enabled = False
    
    For i = 0 To 5              'Initialize ship's Configuration
        SHP(i) = 0
        Call LoadWbox(i, 0)
    Next i
    
    Call Load_Keys
    Call Load_Sounds
    Call Load_Graphics
    Call Wait(1, 1)           'Call main menu
    Call Load_Intro
End Sub
Private Sub MainMenu()          '>>>MAIN MENU<<<
Dim i As Integer
    
    'Aligns all the cursor to recognize the options available
    'on the main menu. there are 3 items (start, options, quit),
    'so the array begins with 0 and end with 2.
    'The third number is the cursor's current position,
    'which is default on 'Start Game' (0)
    'There are six coloured cursors, so the last 2 number
    'represents the start index(0) and end index(5)
    'also related to tracking sub later.
    Call Set_Curr(0, 2, 0, 0, 5)
    
    'Positions the menu in the correct places.
    For i = Curr_Start To Curr_End      'Align menus
        menu(i).Left = 1370
        menu(i).Top = 4840 + (i * 400)
    Next i
    
    Call Align_Curr
    
    Wait_Label.Visible = False
    
    tmMenu.Enabled = True
    Call PLAY(2)                           'Load BGM
End Sub
Private Sub RemapMenu()             '>>>REMAP MENU<<<
    WaitForRemap = False
    
    'Same as above, 10 options available (9-18)
    'Default on 9
    'Another six coloured cursor (6-11)
    Call Set_Curr(9, 18, 9, 6, 11)
    Call Align_Curr
    Call Load_Keys
End Sub

Private Sub GameMenu()             '>>>GAME MENU<<<
    'Left and right selection not available
    'Default only recognize up and down
    MenuSelectLR = False
        
    Call Set_Curr(19, 26, 19, 12, 17)
    Call Align_Curr
    
    'Aligns weapon display boxes
    Weapon_Box.Width = 2255
    Weapon_Box.Height = 2175
    Weapon_Box.Left = 2480
    Weapon_Box.Top = 3375
    Weapon_Box.BorderStyle = 0
End Sub

Private Sub PlayMenu()              '>>>PLAY MENU<<<
'This is a large sub here. Basically when a new game starts,
'all the bars and HUDs needs to be realigned.
'The following codes basically sets their top, left, width, height...

Dim i As Integer
Const LH = 255      'Label Height
Const LW = 615      'Label Width
    
    'A new game ...
    If GameStarted = False Then
        Elasped = 0
        TotFPS = 0
    
        BossBar.Top = 1080
        Stats_Bar.Top = 7920
        Stats_Bar.Left = 720
        Stats_Bar.Height = 975
        Stats_Bar.Width = 4935
        Stats_Bar.Visible = False
        
        Stt_SHD_Lbl.Top = 0
        Stt_SHD_Lbl.Left = 0
        Stt_SHD_Lbl.Height = LH
        Stt_SHD_Lbl.Width = 495
        Stt_SHD_Lbl.BackColor = &H0&
        Stt_SHD_Lbl.ForeColor = &HFF00&
        Stt_SPE_Lbl.Top = 240
        Stt_SPE_Lbl.Left = 0
        Stt_SPE_Lbl.Height = LH
        Stt_SPE_Lbl.Width = 495
        Stt_SPE_Lbl.BackColor = &H0&
        Stt_SPE_Lbl.ForeColor = &HFF00&
        Stt_LVL_MG.Top = 0
        Stt_LVL_MG.Left = 2880
        Stt_LVL_MG.Height = LH
        Stt_LVL_MG.Width = 855
        Stt_LVL_SG.Top = 240
        Stt_LVL_SG.Left = 2880
        Stt_LVL_SG.Height = LH
        Stt_LVL_SG.Width = 855
        
        Stt_MGUN_F.Top = 0
        Stt_MGUN_F.Left = 3720
        Stt_MGUN_F.Height = LH
        Stt_MGUN_F.Width = LW
        Stt_SGUN_F.Top = 240
        Stt_SGUN_F.Left = 3720
        Stt_SGUN_F.Height = LH
        Stt_SGUN_F.Width = LW
        Stt_ENG_F.Top = 0
        Stt_ENG_F.Left = 4320
        Stt_ENG_F.Height = LH
        Stt_ENG_F.Width = LW
        Stt_PWR_F.Top = 240
        Stt_PWR_F.Left = 4320
        Stt_PWR_F.Height = LH
        Stt_PWR_F.Width = LW
  
        Stt_MGUN_R.Top = 0
        Stt_MGUN_R.Left = 3720
        Stt_MGUN_R.Height = LH
        Stt_MGUN_R.Width = LW
        Stt_SGUN_R.Top = 240
        Stt_SGUN_R.Left = 3720
        Stt_SGUN_R.Height = LH
        Stt_SGUN_R.Width = LW
        Stt_ENG_R.Top = 0
        Stt_ENG_R.Left = 4320
        Stt_ENG_R.Height = LH
        Stt_ENG_R.Width = LW
        Stt_PWR_R.Top = 240
        Stt_PWR_R.Left = 4320
        Stt_PWR_R.Height = LH
        Stt_PWR_R.Width = LW
  
        BAR_SHD_C.Top = 30
        BAR_SHD_C.Left = 540
        BAR_SHD_C.Height = 75
        BAR_SHD_C.Width = 2175
        BAR_SHD_R.Top = 90
        BAR_SHD_R.Left = 540
        BAR_SHD_R.Height = 75
        BAR_SHD_R.Width = 2175
        BAR_SPE_C.Top = 270
        BAR_SPE_C.Left = 540
        BAR_SPE_C.Height = 75
        BAR_SPE_C.Width = 2175
        BAR_SPE_R.Top = 330
        BAR_SPE_R.Left = 540
        BAR_SPE_R.Height = 75
        BAR_SPE_R.Width = 2175
    
        For i = 1 To 4
            Bufferbox(i).Top = 2880 + ((i - 1) * 360)
            Bufferbox(i).Left = 1080
            Bufferbox(i).Width = 15
            BufferText(i).Caption = ""
        Next i
        
        Bufferbox(0).Top = Bufferbox(1).Top
        Bufferbox(0).Left = Bufferbox(1).Left
        Bufferbox(0).Width = Bufferbox(1).Width
        BufferText(0).Caption = BufferText(1).Caption
        Bufferbox(6).Top = 1080
        Bufferbox(6).Left = 3840
        Bufferbox(5).Top = 1080
        Bufferbox(5).Left = 840
        BufferText(6).Caption = "000000"
        
        Buffer_Count = 0
        
        Lock_All = True
        Lock_Engine = False
        Lock_Generator = False
        Lock_MainGun = False
        Lock_SideGun = False
        
        ShipX = 2720
        ShipY = 8440
        Ship_Dir = 0
        
        For i = 0 To 3
            Ship_Moving(i) = False  'All directions null
        Next i
        
        For i = 0 To 300
            Bull(i).Used = False    'Initialized bullets to unused
        Next i
        
        For i = 0 To 200
            EBull(i).Used = False   'Initialize bullets to unused
        Next i
        
        MG_Lvl = 1
        SG_Lvl = 1
        SHD_Avail = True
        SPE_Avail = True
        
        Select Case (SHP(2))        'Setting Special
        Case 0:
            SPE_Width = 435
            BAR_SPE_C.Width = 435
            BAR_SPE_R.Width = 435
        Case 1:
            SPE_Width = 870
            BAR_SPE_C.Width = 870
            BAR_SPE_R.Width = 870
        Case 2:
            SPE_Width = 1305
            BAR_SPE_C.Width = 1305
            BAR_SPE_R.Width = 1305
        Case 3:
            SPE_Width = 1740
            BAR_SPE_C.Width = 1740
            BAR_SPE_R.Width = 1740
        Case 4:
            SPE_Width = 2175
            BAR_SPE_C.Width = 2175
            BAR_SPE_R.Width = 2175
        End Select
        
        Select Case (SHP(3))        'Setting Shields
        Case 0:
            SHD_Width = 435
            BAR_SHD_C.Width = 435
            BAR_SHD_R.Width = 435
            SHD_Charge_Base = 7
        Case 1:
            SHD_Width = 870
            BAR_SHD_C.Width = 870
            BAR_SHD_R.Width = 870
            SHD_Charge_Base = 6
        Case 2:
            SHD_Width = 1305
            BAR_SHD_C.Width = 1305
            BAR_SHD_R.Width = 1305
            SHD_Charge_Base = 5
        Case 3:
            SHD_Width = 1740
            BAR_SHD_C.Width = 1740
            BAR_SHD_R.Width = 1740
            SHD_Charge_Base = 4
        Case 4:
            SHD_Width = 2175
            BAR_SHD_C.Width = 2175
            BAR_SHD_R.Width = 2175
            SHD_Charge_Base = 3
        End Select
        
        Repair_Speed = 2                'BASE VALUE
        SHD_Charging = False
        SPE_Charging = False
        
        'Sets the speed of the ship based on the engines chosen
        Select Case (SHP(5))            'ENGINE - Affects Repair
        Case 0:
            Ship_Speed = 80
            Repair_Speed = Repair_Speed + 10
        Case 1:
            Ship_Speed = 100
            Repair_Speed = Repair_Speed + 8
        Case 2:
            Ship_Speed = 130
            Repair_Speed = Repair_Speed + 6
        Case 3:
            Ship_Speed = 160
            Repair_Speed = Repair_Speed + 4
        End Select
  
        For i = 1 To 40
            EShip(i).Used = False
        Next i
        
        Stt_LVL_MG.Caption = "MG  1.0"
        Stt_LVL_SG.Caption = "SG   1.0"
  
        'This is for initialization of the stars in the background
        For i = 0 To Detail_Level     'For stars
            Star(i).PosX = (Rnd * 4920 + 720)
            Star(i).PosY = (Rnd * 6480 + 1440)
            Star(i).Spd = (Rnd * 40 + 10)
            Star(i).Color = (Rnd * 2 + 1)
        Next i
        
        For i = 0 To 2
            PowerUp(i).Used = False
        Next i
        
        'This are all (mostly) the variables that will be used
        'during playing mode. Description is above in the
        'declaration section
        Play_Layer.BackColor = &H0&
        MG_Lvl = 1
        SG_Lvl = 1
        Bull_Delay = 0
        Bull_Delay2 = 0
        Ship_Firing = False
        GameStarted = True
        GameScore = 0
        SHIELD_ON_OFF = 0
        ShakeIT = 0
        Wait_label2.Visible = False
        Moving_StatsBar = 0
        Moving_SpecialBar = 0
        Moving_BossBar = 0
        Special_Running = False
        Special_Delay = 0
        GEC = 0
        KillCount = 0
        KilledBoss = False
        
        Giga_Count = 0

        'Resets all stage counters to 0
        For i = 0 To 4
            C_GEC(i) = 0
        Next i
        'Load explosion sounds into media players
        For i = 0 To 3
            MPS(i).Open (App.Path & "\sound\splode.wav")
        Next i
        'Load explosion sounds into media players
        For i = 4 To 7
            MPS(i).Open (App.Path & "\sound\bplode.wav")
        Next i
    End If
    
    TotFPS = 0
    Elasped = 0
    tmOneFrame.Enabled = True       'Officially Starts the Game!
    tmOneSecond = True              'This is for FPS counting
End Sub

Private Sub PauseMenu()             '>>>PAUSE MENU<<<
Dim i As Integer
    Call Set_Curr(27, 29, 27, 18, 23)
    Call Align_Curr
    
    'HALTS the ship during pause
    For i = 0 To 3
        Ship_Moving(i) = False
    Next i
    
    'Halts firing too.
    Ship_Firing = False
End Sub

Private Sub Set_Curr(C_Start As Integer, C_End As Integer, C_Pos As Integer, C_From As Integer, C_To As Integer)
    'This basically tells the tmMenu_Timer what
    'are the menus that will be used in the current layer.
    'Explained above during the first call of Set_Curr
    Curr_Start = C_Start
    Curr_End = C_End
    Curr_Pos = C_Pos
    Curr_From = C_From
    Curr_To = C_To
End Sub

Private Sub Align_Curr()                'Aligns Cursor
'Aligns the cursors to the option selection boxes
Dim i As Integer
    For i = Curr_From To Curr_To
        Curr(i).Left = menu(Curr_Start).Left
        Curr(i).Width = menu(Curr_Start).Width
    Next i
    
    'Call the cursor to move to the top selection box
    Call TrackMenu(menu(Curr_Start).Top, Curr_From, Curr_From)  'Track to 1st item
    'Changes color of the text
    Call Change_Text(Curr_Start)                                'Change color of 1st item
End Sub

Private Sub I_Box_GotFocus(Index As Integer)
    Menu_Layer.SetFocus
End Sub

Private Sub Bufferbox_GotFocus(Index As Integer)
    Play_Layer.SetFocus
End Sub

Private Sub Stats_Bar_GotFocus()
    Play_Layer.SetFocus
End Sub

Private Sub tmMenu_Timer()  'Highlight Selected Menu

'REFERENCE: For a more theoritical tutorial on tracking,
'           refer to my submission titled 'Sprite Tracking' on
'           Planet Source Code

Dim i As Integer
   For i = Curr_From To Curr_To
            Curr(i).Top = Curr(i).Top - Bal(i)
            If i = Curr_From Then
                Call TrackMenu(menu(Curr_Pos).Top, i, Curr_From)
            Else
                Call TrackMenu(Curr(i - 1).Top, i, Curr_From)
            End If
    Next i
End Sub

Private Sub TrackMenu(Dest As Single, i As Integer, Init As Integer)

'REFERENCE: For a more theoritical tutorial on tracking,
'           refer to my submission titled 'Sprite Tracking' on
'           Planet Source Code

Const SP = 1    'Move Speed
Const TK = 1    'Track Speed
    If i = Init Then
        Bal(i) = (Curr(i).Top - Dest) / SP
    Else:
        Bal(i) = (Curr(i).Top - Dest) / TK
    End If
End Sub

Private Sub Change_Text(Choice As Integer)  'Change Highlighted Text Color
'This is a simple sub that changes the colour of the text
'when it is the currently selected
Dim i As Integer
Dim x As Integer
    For i = Curr_Start To Curr_End
        If i = Choice Then
            If i = 19 Or i = 20 Then        'Start/Abort Color
                menu(i).ForeColor = &HFFFF&
            Else:
                menu(i).ForeColor = &HFF00&
            End If
            If i >= 19 And i <= 26 Then
                For x = 21 To 26            'Wbox Color
                    If x = i Then
                        WBox(x - 21).ForeColor = &HFFFF80
                    Else:
                        WBox(x - 21).ForeColor = &HC0C0FF
                    End If
                Next x
            End If
        Else:
            menu(i).ForeColor = &HFFFFFF
        End If
    Next i
End Sub
Private Sub Menu_Layer_KeyDown(Key As Integer, Shift As Integer)
    
    'Detects the keypresses that the user made in the Menu Layer
    If MenuMode_Enabled = True And Wait_Label.Visible = False Then
        Select Case (Key)
        Case 27:                                'ESC
            Call Quit
        Case 38:                                'Menu KEYS
            Curr_Pos = Curr_Pos - 1     'UP BUTTON
            If Curr_Pos < Curr_Start Then Curr_Pos = Curr_End
            Call PLAY(1)            'Play the 'beep' sound =)
        Case 40:
            Curr_Pos = Curr_Pos + 1     'DOWN BUTTON
            If Curr_Pos > Curr_End Then Curr_Pos = Curr_Start
            Call PLAY(1)
        Case 13:
            Select Case (Curr_Pos)
            Case 0:                     'START
                Call Show_Hide(3, 1, 0.01, 3)
            Case 1:                     'REMAP
                Call Show_Hide(2, 1, 0.01, 2)
            Case 2:                     'QUIT
                Call Quit
            End Select
            Call PLAY(1)
        End Select
    Call Change_Text(Curr_Pos)
    End If
    If Key = 115 And Shift = 2 Then Call Quit  '(Alt + F4)
End Sub

Private Sub Remap_Layer_KeyDown(Key As Integer, Shift As Integer)
    
    'Detects the keypresses that the user made in the Option Layer
    If WaitForRemap = True Then     'See if it is waiting for user to enter a new remap button
        Call LoadList(Key, (Curr_Pos - 6))  'Load the Keyboard button Name
        KYB(Curr_Pos - 9) = Key        'Sets the saved button position to memory
        Call PLAY(1)
        WaitForRemap = False        'Program no longer waiting for user to enter a new keyboard button for remap
    Else:
        If RemapMode_Enabled = True Then
            Select Case (Key)
            Case 27:                                'ESC
                Call PLAY(1)
                Call Show_Hide(1, 2, 0.01, 1)
            Case 38:                                'Menu KEYS
                Curr_Pos = Curr_Pos - 1 'UP
                If Curr_Pos < Curr_Start Then Curr_Pos = Curr_End
                Call PLAY(1)
            Case 40:
                Curr_Pos = Curr_Pos + 1 'DOWN
                If Curr_Pos > Curr_End Then Curr_Pos = Curr_Start
                Call PLAY(1)
            
            'VOLUME CONTROLS
            Case 39:                        'INCREASE
                If Curr_Pos = 15 Then   'RIGHT
                    Vol(0) = Vol(0) + 1
                    If Vol(0) > 100 Then
                        Vol(0) = 100
                    End If
                    LabelSE.Caption = Str(Vol(0))
                    Call PLAY(1)
                End If
                If Curr_Pos = 16 Then   'LEFT
                    Vol(1) = Vol(1) + 1
                    If Vol(1) > 100 Then
                        Vol(1) = 100
                    End If
                    LabelBGM.Caption = Str(Vol(1))
                End If
                Call Set_Volumes
            Case 37:                'DECREASE
                If Curr_Pos = 15 Then
                    Vol(0) = Vol(0) - 1
                    If Vol(0) < 0 Then
                        Vol(0) = 0
                    End If
                    LabelSE.Caption = Str(Vol(0))
                    Call PLAY(1)
                End If
                If Curr_Pos = 16 Then
                    Vol(1) = Vol(1) - 1
                    If Vol(1) < 0 Then
                        Vol(1) = 0
                    End If
                    LabelBGM.Caption = Str(Vol(1))
                End If
                Call Set_Volumes
            
            'Remap new keyboard buttons
            Case 13:    'ENTER KEY
                Select Case (Curr_Pos)
                Case 9:                     'REMAP UP
                    Call Remap_Keys
                Case 10:                    'REMAP DOWN
                    Call Remap_Keys
                Case 11:                    'REMAP LEFT
                    Call Remap_Keys
                Case 12:                    'REMAP RIGHT
                    Call Remap_Keys
                Case 13:                    'REMAP FIRE
                    Call Remap_Keys
                Case 14:                    'REMAP SPECIAL
                    Call Remap_Keys
                Case 17:                     'ACCEPT
                    Call Show_Hide(1, 2, 0.1, 1)
                    Call Save_Keys  'Saved new keys
                Case 18:                      'CANCEL
                    Call Show_Hide(1, 2, 0.1, 1)
                    Call Load_Keys  '  Restore default keys
                End Select
                Call PLAY(1)
            End Select
        Call Change_Text(Curr_Pos)
        End If
    End If
    If Key = 115 And Shift = 2 Then Call Quit  '(Alt + F4)
End Sub

Private Sub Remap_Keys()
    'Sets the program to wait for the user to key in a button
    WaitForRemap = True
    menu(Curr_Pos - 6).Caption = "Press a Key"
End Sub

Private Sub Game_Layer_KeyDown(Key As Integer, Shift As Integer)
    'This is for ship configuration before u start the game
    If GameMode_Enabled = True Then
        Select Case (Key)
        Case 27:                                'ESC
            Call Show_Hide(1, 3, 0.1, 1)
            Call PLAY(1)
        Case 38:                                'Menu KEYS
            Curr_Pos = Curr_Pos - 1     'SCROLL through options
            If Curr_Pos < Curr_Start Then Curr_Pos = Curr_End
            If Curr_Pos >= 21 And Curr_Pos <= 26 Then
                MenuSelectLR = True     'Selected menu can scroll left and right
            Else:
                MenuSelectLR = False    'Selected menu is dormant
            End If
            Call PLAY(1)
        Case 40:
            Curr_Pos = Curr_Pos + 1     'SCROLL through options
            If Curr_Pos > Curr_End Then Curr_Pos = Curr_Start
            If Curr_Pos >= 21 And Curr_Pos <= 26 Then
                MenuSelectLR = True
            Else:
                MenuSelectLR = False
            End If
            Call PLAY(1)
        Case 39:                    'INCREASE WEAPON
           If MenuSelectLR = True Then
            If Curr_Pos >= 21 And Curr_Pos <= 26 Then
                'Scrolls through the ship's configurations
                SHP(Curr_Pos - 21) = SHP(Curr_Pos - 21) + 1
                Select Case (Curr_Pos - 21)
                Case 0:
                    If SHP(0) > 4 Then SHP(0) = 0   'After last option, returns to first option
                Case 1:
                    If SHP(1) > 3 Then SHP(1) = 0
                Case 2:
                    If SHP(2) > 4 Then SHP(2) = 0
                Case 3:
                    If SHP(3) > 4 Then SHP(3) = 0
                Case 4:
                    If SHP(4) > 2 Then SHP(4) = 0
                Case 5:
                    If SHP(5) > 3 Then SHP(5) = 0
                End Select
                'Put the name of the item selected in the wbox text
                Call LoadWbox(Curr_Pos - 21, SHP(Curr_Pos - 21))
            End If
            Call PLAY(3)
           End If
        Case 37:                    'DECREASE WEAPON
           If MenuSelectLR = True Then
            'Same as above...
            If Curr_Pos >= 21 And Curr_Pos <= 26 Then
                SHP(Curr_Pos - 21) = SHP(Curr_Pos - 21) - 1
                Select Case (Curr_Pos - 21)
                Case 0:
                    If SHP(0) < 0 Then SHP(0) = 4
                Case 1:
                    If SHP(1) < 0 Then SHP(1) = 3
                Case 2:
                    If SHP(2) < 0 Then SHP(2) = 4
                Case 3:
                    If SHP(3) < 0 Then SHP(3) = 4
                Case 4:
                    If SHP(4) < 0 Then SHP(4) = 2
                Case 5:
                    If SHP(5) < 0 Then SHP(5) = 3
                End Select
                Call LoadWbox(Curr_Pos - 21, SHP(Curr_Pos - 21))
            End If
            Call PLAY(3)
           End If
        Case 13:
            Select Case (Curr_Pos)
            Case 19:                     'START MISSION
                Call PLAY(1)
                MP2.Mute = True
                Play_Layer.Cls
                Wait_label2.Visible = True
                Call Show_Hide(4, 3, 0.3, 4)    '>goto play menu
            Case 20:                      'ABORT
                Call PLAY(1)
                Call Show_Hide(1, 3, 0.1, 1)    '>goto main menu
            End Select
            Call PLAY(1)
        End Select
        Call Change_Text(Curr_Pos)
    End If
    If Key = 115 And Shift = 2 Then Call Quit  '(Alt + F4)
End Sub

Private Sub Play_Layer_KeyDown(Key As Integer, Shift As Integer)
    
    'These buttons follows the buttons customly remapped by the player.
    If PlayMode_Enabled = True Then
        Select Case (Key)
        Case 27:                                'ESC
            Call Show_Hide(5, 4, 0.01, 5)       'Shows the PAUSE Menu
            Call PLAY(1)
        End Select
    
        If OVERIDE = False Then 'Lost control of ship
            If Lock_All = False Then
                Select Case (Key)
                'Ship moving in a certain direction is true
                Case KYB(0):                                'UP
                    If Lock_Engine = False Then Ship_Moving(1) = True
                Case KYB(1):                                'DOWN
                    If Lock_Engine = False Then Ship_Moving(3) = True
                Case KYB(2):                                'LEFT
                    If Lock_Engine = False Then Ship_Moving(0) = True
                Case KYB(3):                                'RIGHT
                    If Lock_Engine = False Then Ship_Moving(2) = True
                Case KYB(4):                                'FIRE
                    Ship_Firing = True
                Case KYB(5):                                'SPECIAL
                    Call Use_Special(SHP(2))
                End Select
            End If
        End If
    End If
End Sub

Private Sub Play_Layer_KeyUp(Key As Integer, Shift As Integer)
    'The ship stops performing that action if button is depressed
    If OVERIDE = False Then
            Select Case (Key)
            Case KYB(0):                                'UP
                Ship_Moving(1) = False
            Case KYB(1):                                'DOWN
                Ship_Moving(3) = False
            Case KYB(2):                                'LEFT
                Ship_Moving(0) = False
            Case KYB(3):                                'RIGHT
                Ship_Moving(2) = False
            Case KYB(4):                                'FIRE
                Ship_Firing = False
            End Select
    End If
End Sub

Private Sub Pause_Layer_KeyDown(Key As Integer, Shift As Integer)
        If PauseMode_Enabled = True Then
            Select Case (Key)
            Case 27:                                'ESC
                Call Show_Hide(4, 5, 0.01, 4)
            Case 38:                                'Menu KEYS
                Curr_Pos = Curr_Pos - 1
                If Curr_Pos < Curr_Start Then Curr_Pos = Curr_End
                Call PLAY(1)
            Case 40:
                Curr_Pos = Curr_Pos + 1
                If Curr_Pos > Curr_End Then Curr_Pos = Curr_Start
                Call PLAY(1)
            Case 13:
                Select Case (Curr_Pos)
                Case 27:                   'RESUME
                    Call Show_Hide(4, 5, 0.1, 4)
                Case 28:                   'ABORT
                    GameStarted = False
                    Call Show_Hide(1, 5, 0.2, 1)
                Case 29:                   'QUIT
                    Call Quit
                End Select
            End Select
        Call Change_Text(Curr_Pos)
        End If
    If Key = 115 And Shift = 2 Then Call Quit  '(Alt + F4)
End Sub

Private Sub Quit()      'END PROGRAM
    tmMenu.Enabled = False          'Kill Timers
    tmDelay.Enabled = False
    tmOneFrame.Enabled = False
    tmOneSecond = False
    tmIntro.Enabled = False
    Set FormBW = Nothing            'Kill Variables
    Unload Me                       'Kill Form
    End                             'Kill Program
End Sub

Private Sub Show_Hide(Show As Integer, Hide As Integer, Del_A As Single, Del_E As Integer)
'This sub performs 3 function, in the parameter list,
'the 1st number represents which Layer to show,
'the 2nd number represents which Layer to hide,
'the 3rd is the time to delay an event if there's an event after the show/hide is done.
'the 4th is the event number, each event bearing different outcome

Dim i As Integer
    Select Case (Show)      'Shows and enables the selected Layer
    Case 1:
        MenuMode_Enabled = True
        For i = 0 To 5
            Curr(i).Top = 120
        Next i
        Menu_Layer.Visible = True
        Menu_Layer.SetFocus
    Case 2:
        RemapMode_Enabled = True
        Remap_Layer.Visible = True
        Remap_Layer.SetFocus
    Case 3:
        GameMode_Enabled = True
        Game_Layer.Visible = True
        Game_Layer.SetFocus
    Case 4:
        PlayMode_Enabled = True
        Play_Layer.Visible = True
        Play_Layer.Enabled = True
        Play_Layer.SetFocus
        tmMenu.Enabled = False
    Case 5:
        PauseMode_Enabled = True
        Pause_Layer.Visible = True
        Pause_Layer.SetFocus
    End Select
    
    Select Case (Hide)      'Hides and disable the selected Layer
    Case 1:
        MenuMode_Enabled = False
        Menu_Layer.Visible = False
    Case 2:
        RemapMode_Enabled = False
        Remap_Layer.Visible = False
    Case 3:
        GameMode_Enabled = False
        Game_Layer.Visible = False
    Case 4:
        PlayMode_Enabled = False
        Play_Layer.Enabled = False
        tmOneFrame.Enabled = False
        tmOneSecond = False
        tmMenu.Enabled = True
    Case 5:
        PauseMode_Enabled = False
        Pause_Layer.Visible = False
    End Select
    
    Call Wait(Del_A, Del_E)     'Call an event to happen
    
End Sub

Private Sub Load_Sounds()
    'Sounds loaded that will not be changed anymore after loaded
    MP1.Open (App.Path & "\sound\menu.wav")
    MP3.Open (App.Path & "\sound\config.wav")
    MP4.Open (App.Path & "\sound\powerup.wav")
    Call Set_Volumes
End Sub

Private Sub PLAY(i As Integer)
'This sub performs nothing more that calling the
'media player with the loaded sound to play their files

If FormOptimize.Chk_Mute.Value = 1 Then Exit Sub
Static Curr_Player As Integer
On Error Resume Next
    Select Case (i)
    Case 1:                         'Menu Beep
        MP1.PLAY
    Case 2:                         'Intro BG
        MP2.Mute = False
        MP2.Open (App.Path & "\sound\bgmenu.mid")
    Case 3:                         'Config Ship
        MP3.PLAY
    Case 4:
        MP2.Stop
        MP2.Mute = False
        MP2.Open (App.Path & "\sound\bgplay.mid")
    Case 5:
        MP5.Open (App.Path & "\sound\warning.wav")
    Case 6:
        MP5.Open (App.Path & "\sound\startup.wav")
    Case 7:
        MP4.PLAY
    Case 99:
        MP2.PLAY
    Case Else:
        
            Curr_Player = Curr_Player + 1
            If Curr_Player > 3 Then Curr_Player = 0
            
            Select Case (i)
            Case 100:
                MPS(Curr_Player).PLAY
            Case 101:
                MPS(Curr_Player + 4).PLAY
            End Select
        
    End Select
End Sub

Private Sub Set_Volumes()
'Set the volume levels of the media players
Dim i As Integer
    MP1.Volume = (1 - (Vol(0) / 100)) * -3100
    MP2.Volume = (1 - (Vol(1) / 100)) * -3100
    MP3.Volume = (1 - (Vol(0) / 100)) * -3100
    MP5.Volume = (1 - (Vol(1) / 100)) * -3100
    For i = 0 To 7
        MPS(i).Volume = (1 - (Vol(0) / 100)) * -3100
    Next i
End Sub

Private Sub Save_Keys()

'If the user selects Accept in the options layer, the
'keys are saved into a file.
Dim File1
    File1 = FreeFile
Dim Filename As String
    Filename = App.Path & "\" & "custom.cfg"
Open (Filename) For Output As File1
    Write #File1, KYB(0)
    Write #File1, KYB(1)
    Write #File1, KYB(2)
    Write #File1, KYB(3)
    Write #File1, KYB(4)
    Write #File1, KYB(5)
    Write #File1, Vol(0)
    Write #File1, Vol(1)
Close File1
End Sub

Private Sub Load_Keys()
'When the game loads of the user select cancel from the options layer, then the default keys are loaded
Dim i As Integer
Dim File1
    File1 = FreeFile
Dim Filename As String
    Filename = App.Path & "\" & "custom.cfg"
Open (Filename) For Input As File1
    Input #File1, KYB(0)
    Input #File1, KYB(1)
    Input #File1, KYB(2)
    Input #File1, KYB(3)
    Input #File1, KYB(4)
    Input #File1, KYB(5)
    Input #File1, Vol(0)
    Input #File1, Vol(1)
Close File1

    For i = 0 To 5
        Call LoadList(KYB(i), (i + 3))
        LabelSE.Caption = Str(Vol(0))
        LabelBGM.Caption = Str(Vol(1))
    Next i
End Sub

Private Sub Load_Graphics()
    Menu_Layer.Picture = LoadPicture(App.Path & "\graphic\menubg.gif")
    Remap_Layer.Picture = LoadPicture(App.Path & "\graphic\menubg.gif")
    Game_Layer.Picture = LoadPicture(App.Path & "\graphic\menubg.gif")
End Sub


Private Sub Wait(MiliSecs As Single, Event_No As Integer)  'Wait Function
'This sub delays the computer for a number of seconds and
'calls an even to happen.
    tmDelay.Interval = MiliSecs * 1000
    tmDelay.Enabled = True
    Delay_Counter = 0
    Delay_Occasion = Event_No
End Sub

Private Sub Load_Intro()
'This sub just moves the titles around during the title screen startup.
Dim i As Integer
    Intro_Counter = 0
    tmIntro.Enabled = True
    
    For i = 0 To 2
        Title(i).ForeColor = RGB(5, 5, 5)
    Next i
    
    Title(3).Top = 3000
    Title(3).Left = 960
    
    I_Box(0).Top = 0
    I_Box(0).Left = 960
    I_Box(2).Top = 3000
    I_Box(2).Left = 480
    I_Box(1).Top = 1200
    I_Box(1).Left = 3840
    I_Box(3).Top = 1220
    I_Box(3).Left = -4460
    I_Box(4).Top = 2600
    I_Box(4).Left = 5140
    
    R_Color = 5
    G_Color = 5
    B_Color = 5
End Sub

Private Sub tmIntro_Timer()     'INTRODUCTION
'As the title screen loads...
Dim i As Integer
    Intro_Counter = Intro_Counter + 1
   
    '... title words are being moved around ...
    If Intro_Counter < 68 Then
        I_Box(0).Top = I_Box(0).Top + 18
        I_Box(2).Top = I_Box(2).Top - 18
    End If

    '... more title words are being moved around ...
    If Intro_Counter > 60 And Intro_Counter <= 100 Then
        I_Box(3).Left = I_Box(3).Left + 120
        I_Box(4).Left = I_Box(4).Left - 120
    End If

    '... and their text if fading black to white ...
    R_Color = R_Color + 5
    G_Color = G_Color + 5
    B_Color = B_Color + 5
    
    For i = 0 To 2
        Title(i).ForeColor = RGB(R_Color, G_Color, B_Color)
    Next i

    '... until they all are nicely set, then timer is killed.
    If Intro_Counter = 100 Then
        Title(3).Visible = True
        Title(4).Visible = True
        tmIntro.Enabled = False
    End If
End Sub

Private Sub tmDelay_Timer()         'Delay Counter + Occasions
Dim i As Integer
    'Pauses the game for a short period and make a call to intialize
    'a layer's variables and settings
    Delay_Counter = Delay_Counter + 1
    If Delay_Counter = 2 Then
        tmDelay.Enabled = False
        
        Select Case (Delay_Occasion)
        Case 1:
            Call MainMenu
        Case 2:
            Call RemapMenu
        Case 3:
            Call GameMenu
        Case 4:
            Call PlayMenu
        Case 5:
            Call PauseMenu
        End Select
    Delay_Counter = 0
    End If
End Sub

Private Sub LoadWbox(Weapon_Slot As Integer, Weapon_Name As Integer)
    
    'As you scroll around the ship's configurations, this sub
    'changes the text of the weapons selected.
    Select Case (Weapon_Slot)
    Case 0:
        Select Case (Weapon_Name)       'MAIN GUN
        Case 0:
            WBox(Weapon_Slot).Caption = "Vulcan-X"
        Case 1:
            WBox(Weapon_Slot).Caption = "Light Beam"
        Case 2:
            WBox(Weapon_Slot).Caption = "Particle Burst"
        Case 3:
            WBox(Weapon_Slot).Caption = "Plasma Wave"
        Case 4:
            WBox(Weapon_Slot).Caption = "Vulcan Blue"
        End Select
    Case 1:                             'SIDE GUN
        Select Case (Weapon_Name)
        Case 0:
            WBox(Weapon_Slot).Caption = "Micro Missiles"
        Case 1:
            WBox(Weapon_Slot).Caption = "Pulse Ring"
        Case 2:
            WBox(Weapon_Slot).Caption = "Crescent Blades"
        Case 3:
            WBox(Weapon_Slot).Caption = "Creep Mine"
        End Select
    Case 2:                             'SPECIAL
        Select Case (Weapon_Name)
        Case 0:
            WBox(Weapon_Slot).Caption = "Inferno Disc"
        Case 1:
            WBox(Weapon_Slot).Caption = "Matrix Globe"
        Case 2:
            WBox(Weapon_Slot).Caption = "Giga Storm"
        Case 3:
            WBox(Weapon_Slot).Caption = "Proximity Burn"
        Case 4:
            WBox(Weapon_Slot).Caption = "Armageddon"
        End Select
    Case 3:                             'SHIELD
        Select Case (Weapon_Name)
        Case 0:
            WBox(Weapon_Slot).Caption = "Matrix SHD MK I"
        Case 1:
            WBox(Weapon_Slot).Caption = "Matrix SHD MK II"
        Case 2:
            WBox(Weapon_Slot).Caption = "Matrix SHD MK III"
        Case 3:
            WBox(Weapon_Slot).Caption = "Matrix SHD MK IV"
        Case 4:
            WBox(Weapon_Slot).Caption = "Matrix SHD MK V"
        End Select
    Case 4:                             'GENERATOR
        Select Case (Weapon_Name)
        Case 0:
            WBox(Weapon_Slot).Caption = "Dual-Core Charger"
        Case 1:
            WBox(Weapon_Slot).Caption = "Fusion Charger"
        Case 2:
            WBox(Weapon_Slot).Caption = "Matrix Charger"
        End Select
    Case 5:                             'ENGINE
        Select Case (Weapon_Name)
        Case 0:
            WBox(Weapon_Slot).Caption = "4 Cylinders"
        Case 1:
            WBox(Weapon_Slot).Caption = "6 Cylinders"
        Case 2:
            WBox(Weapon_Slot).Caption = "8 Cylinders"
        Case 3:
            WBox(Weapon_Slot).Caption = "12 Cylinders"
        End Select
    End Select
    
End Sub

Private Sub LoadList(Key As Integer, Label As Integer)
    'Similar to above, as a new Keyboard button is being remapped,
    'it will be displayed to the user the selected button
    Select Case (Key)
    Case 8:
        menu(Label).Caption = "Backspace"
    Case 13:
        menu(Label).Caption = "Enter"
    Case 16:
        menu(Label).Caption = "Shift"
    Case 17:
        menu(Label).Caption = "Ctrl"
    Case 18:
        menu(Label).Caption = "Alt"
    Case 19:
        menu(Label).Caption = "Pause"
    Case 20:
        menu(Label).Caption = "Caps Lock"
    Case 32:
        menu(Label).Caption = "Spacebar"
    Case 33:
        menu(Label).Caption = "Page Up"
    Case 34:
        menu(Label).Caption = "Page Down"
    Case 35:
        menu(Label).Caption = "End"
    Case 36:
        menu(Label).Caption = "Home"
    Case 37:
        menu(Label).Caption = "Left"
    Case 38:
        menu(Label).Caption = "Up"
    Case 39:
        menu(Label).Caption = "Right"
    Case 40:
        menu(Label).Caption = "Down"
    Case 45:
        menu(Label).Caption = "Insert"
    Case 46:
        menu(Label).Caption = "Delete"
    Case 48:
        menu(Label).Caption = "0"
    Case 49:
        menu(Label).Caption = "1"
    Case 50:
        menu(Label).Caption = "2"
    Case 51:
        menu(Label).Caption = "3"
    Case 52:
        menu(Label).Caption = "4"
    Case 53:
        menu(Label).Caption = "5"
    Case 54:
        menu(Label).Caption = "6"
    Case 55:
        menu(Label).Caption = "7"
    Case 56:
        menu(Label).Caption = "8"
    Case 57:
        menu(Label).Caption = "9"
    Case 65:
        menu(Label).Caption = "A"
    Case 66:
        menu(Label).Caption = "B"
    Case 67:
        menu(Label).Caption = "C"
    Case 68:
        menu(Label).Caption = "D"
    Case 69:
        menu(Label).Caption = "E"
    Case 70:
        menu(Label).Caption = "F"
    Case 71:
        menu(Label).Caption = "G"
    Case 72:
        menu(Label).Caption = "H"
    Case 73:
        menu(Label).Caption = "I"
    Case 74:
        menu(Label).Caption = "J"
    Case 75:
        menu(Label).Caption = "K"
    Case 76:
        menu(Label).Caption = "L"
    Case 77:
        menu(Label).Caption = "M"
    Case 78:
        menu(Label).Caption = "N"
    Case 79:
        menu(Label).Caption = "O"
    Case 80:
        menu(Label).Caption = "P"
    Case 81:
        menu(Label).Caption = "Q"
    Case 82:
        menu(Label).Caption = "R"
    Case 83:
        menu(Label).Caption = "S"
    Case 84:
        menu(Label).Caption = "T"
    Case 85:
        menu(Label).Caption = "U"
    Case 86:
        menu(Label).Caption = "V"
    Case 87:
        menu(Label).Caption = "W"
    Case 88:
        menu(Label).Caption = "X"
    Case 89:
        menu(Label).Caption = "Y"
    Case 90:
        menu(Label).Caption = "Z"
    Case 96:
        menu(Label).Caption = "Numpad 0"
    Case 97:
        menu(Label).Caption = "Numpad 1"
    Case 98:
        menu(Label).Caption = "Numpad 2"
    Case 99:
        menu(Label).Caption = "Numpad 3"
    Case 100:
        menu(Label).Caption = "Numpad 4"
    Case 101:
        menu(Label).Caption = "Numpad 5"
    Case 102:
        menu(Label).Caption = "Numpad 6"
    Case 103:
        menu(Label).Caption = "Numpad 7"
    Case 104:
        menu(Label).Caption = "Numpad 8"
    Case 105:
        menu(Label).Caption = "Numpad 9"
    Case 106:
        menu(Label).Caption = "Numpad *"
    Case 107:
        menu(Label).Caption = "Numpad +"
    Case 109:
        menu(Label).Caption = "Numpad -"
    Case 110:
        menu(Label).Caption = "Numpad ."
    Case 111:
        menu(Label).Caption = "Numpad /"
    Case 112:
        menu(Label).Caption = "F1"
    Case 113:
        menu(Label).Caption = "F2"
    Case 114:
        menu(Label).Caption = "F3"
    Case 115:
        menu(Label).Caption = "F4"
    Case 116:
        menu(Label).Caption = "F5"
    Case 117:
        menu(Label).Caption = "F6"
    Case 118:
        menu(Label).Caption = "F7"
    Case 119:
        menu(Label).Caption = "F8"
    Case 120:
        menu(Label).Caption = "F9"
    Case 121:
        menu(Label).Caption = "F10"
    Case 144:
        menu(Label).Caption = "Num Lock"
    Case 145:
        menu(Label).Caption = "Scroll Lock"
    Case 186:
        menu(Label).Caption = ";"
    Case 187:
        menu(Label).Caption = "-"
    Case 188:
        menu(Label).Caption = ","
    Case 189:
        menu(Label).Caption = "="
    Case 190:
        menu(Label).Caption = "."
    Case 191:
        menu(Label).Caption = "/"
    Case 192:
        menu(Label).Caption = "`"
    Case 219:
        menu(Label).Caption = "["
    Case 220:
        menu(Label).Caption = "\"
    Case 221:
        menu(Label).Caption = "]"
    Case 222:
        menu(Label).Caption = "'"
    Case Else:
        menu(Label).Caption = "Undefined"
    End Select
End Sub

Private Sub EXPLODE(x As Integer, y As Integer, Style As Integer)
'This sub creates an explosion at a fixed point, determines
'how many frames are there in the explosion sequence,
'and set's the explosions X-offset and Y-offset to make
'their explosion point more correct
Dim i As Integer
    For i = 0 To 40     'Locates an unsed explosion object and create a new explosion
        If Explo(i).Used = False Then
            Explo(i).Used = True
            Explo(i).Kind = Style
            Explo(i).PosX = x       'Axis coordinates for explosion point
            Explo(i).PosY = y
            Explo(i).Current_Frame = 1  'Sets to first frame of imagelist
            Select Case (Style)
                Case 1:
                    Explo(i).Last_Frame = 16    'Determine how many frame in an explosion
                    Explo(i).PosX = Explo(i).PosX - 1065 / 2    'X Offset
                    Explo(i).PosY = Explo(i).PosY - 1500 / 2    'Y Offset
                Case 2:
                    Explo(i).Last_Frame = 30
                    Explo(i).PosX = Explo(i).PosX - 3840 / 2    'X Offset
                    Explo(i).PosY = Explo(i).PosY - 2880 / 2    'Y Offset
                Case 3:
                    Explo(i).Last_Frame = 22
                    Explo(i).PosX = Explo(i).PosX - 3840 / 2    'X Offset
                    Explo(i).PosY = Explo(i).PosY - 2880 / 2    'Y Offset
                Case 17:
                    Explo(i).Last_Frame = 5
                    Explo(i).PosX = Explo(i).PosX - 840 / 2    'X Offset
                    Explo(i).PosY = Explo(i).PosY - 880 / 2    'Y Offset
                Case 13:
                    Explo(i).Last_Frame = 5
                    Explo(i).PosX = Explo(i).PosX - 200 / 2    'X Offset
                    Explo(i).PosY = Explo(i).PosY - 200 / 2    'Y Offset
            End Select
            Exit For
        End If
    Next i
End Sub

Private Sub Use_Special(Kind As Integer)
'When a player uses a Special ...
Dim i As Integer
Dim x As Integer
    'Bar reaches zero, chargin starts again
    If SPE_Charging = False Then
        BAR_SPE_R.Width = 0
    End If

    If SPE_Avail = True Then
        Moving_SpecialBar = True    'Moves the special weapon textbar to appear briefly
        BufferText(7).Caption = UCase(WBox(2).Caption)
        BAR_SPE_C.Width = 0
        Stt_SPE_Lbl.BackColor = &HC0&
        Stt_SPE_Lbl.ForeColor = &HFFFFFF
        SPE_Avail = False
        Special_Running = True
        Moving_SpecialBar = 2
        
        Select Case (Kind)
        Case 0:     'Inferno Disc
            For i = 290 To 300
                'Creates a bullet
                If Bull(i).Used = False Then
                    Bull(i).Used = True
                    Bull(i).Kind = 1
                    Bull(i).Dire = 8
                    Bull(i).Span = 0
                    Bull(i).Dama = 50
                    Bull(i).PosX = ShipX - 90
                    Bull(i).PosY = ShipY - 0
                    Exit For
                End If
            Next i
        Case 1:     'Matrix Globe
            For i = 290 To 300
                'Creates a bullet
                If Bull(i).Used = False Then
                    Bull(i).Used = True
                    Bull(i).Kind = 1
                    Bull(i).Dire = 8
                    Bull(i).Span = 0
                    Bull(i).Dama = 1
                    Bull(i).PosX = ShipX - 90
                    Bull(i).PosY = ShipY - 0
                    Bull(i).Spd = 40
                    Exit For
                End If
            Next i
        Case 2:     'Giga Storm
            Giga_Count = 1
            'Mass Damage
            Img(27).ListImages(1).Draw Play_Layer.hDC, 1000, 2000, 1
            Img(27).ListImages(1).Draw Play_Layer.hDC, 3000, 2000, 1
            
            For i = 0 To 40
                If EShip(i).Used = True Then EShip(i).Life = EShip(i).Life - 500
                If EShip(i).Life < 0 Then EShip(i).Life = 0
                If EShip(i).Boss = True Then Boss_Bar_Top.Width = EShip(i).Life / EShip(i).Max_Life * Boss_Bar_Bottom.Width
                Call Check_EDeath(i)
                Call EXPLODE(EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) / 2, 1)
            Next i
            
        Case 3:     'Proximity Burn
            For i = 290 To 300
                'Create a bullet
                If Bull(i).Used = False Then
                    Bull(i).Used = True
                    Bull(i).Kind = 0
                    Bull(i).Dire = 8
                    Bull(i).Span = 0
                    Bull(i).Dama = 100
                    Bull(i).PosX = ShipX - 3840 / 2
                    Bull(i).PosY = ShipY - 2880 / 2
                    Bull(i).Spd = 1
                    Exit For
                End If
            Next i
        Case 4:     'Armageddon
            For i = 2000 To 4000 Step 1000
                For x = 2000 To 6000 Step 2000
                    Call EXPLODE(i, x, 2)
                Next x
            Next i
            For i = 0 To 40
                'Mass Damage
                If EShip(i).Used = True Then EShip(i).Life = EShip(i).Life - 4000
                If EShip(i).Life < 0 Then EShip(i).Life = 0
                If EShip(i).Boss = True Then Boss_Bar_Top.Width = EShip(i).Life / EShip(i).Max_Life * Boss_Bar_Bottom.Width
                Call Check_EDeath(i)
                Call EXPLODE(EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) / 2, 1)
                Call PLAY(101)
                Call PLAY(101)
                Call PLAY(101)
                Call PLAY(101)
            Next i
        End Select
    End If
    SPE_Charging = True
End Sub

Private Sub Send_MSG(Message As String, Color_0White_1Green_2Red_3Yellow As Integer, BColor_0Black_1Red_2Yellow As Integer)
'By calling this sub, u can send a message text to the player
'on the screen.
'Parameter list: The message to be sent, the forecolor, the backcolor
    Buffering = True
    BufferText(0).Caption = Message
    Select Case (Color_0White_1Green_2Red_3Yellow)
    Case 0:  'Friendly Units(BLUE)
        BufferText(0).ForeColor = &HFFFFFF
    Case 1:  'Power-ups     (GREEN)
        BufferText(0).ForeColor = &HFF00&
    Case 2:  'Enemy Comm    (RED)
        BufferText(0).ForeColor = &HFF&
    Case 3:  'Ship Reports  (YELLOW)
        BufferText(0).ForeColor = &HFFFF&
    End Select
    Select Case (BColor_0Black_1Red_2Yellow)
    Case 1:                 '(RED)
        BufferText(0).BackColor = &HFF&
    Case 2:                 '(YELLOW)
        BufferText(0).BackColor = &HFFFF&
    Case Else:              '(BLACK)
        BufferText(0).BackColor = &H0&
    End Select
    Bufferbox(0).Width = BufferText(0).Width
End Sub

Private Sub Critical_HIT()
'When a ship is critically shot and destroyed...
Dim i As Integer
ShakeIT = 1 'Shakes the screen

'Create an explosion at the ship's coordinates
Call EXPLODE(ShipX + (Img(0).ImageWidth * 15) / 2, ShipY + (Img(0).ImageHeight * 15) / 2, 1)
Call Send_MSG("> Your Ship has been DESTROYED ", 0, 0)
Call PLAY(101)

Moving_StatsBar = 1 'Hides all the bars
Lock_All = True     'Lock all controls
OVERIDE = True      'Program overtake controls

Ship_Moving(0) = False
Ship_Moving(1) = False
Ship_Moving(2) = False
Ship_Moving(3) = False
Ship_Firing = False
ShipX = 0   'Hides the ship
ShipY = 0   'Hides the ship

Call Check_Score

GEC = -300  'Playmode enters the death sequence before going back to main menu
End Sub

Private Sub Check_Score()
Dim i As Integer
'   High Score file Creation
Dim file2
    file2 = FreeFile
Dim Filename As String
    Filename = App.Path & "\" & "HiScore.txt"
Dim Hiscore As String

'Opens the Hiscore.txt file to retrieve existing hiscore
Open (Filename) For Input As file2
    For i = 1 To 3
        Input #file2, Hiscore       'Loads hiscore from file
    Next i
Close file2

'Extract highscore from sentence
Hiscore = Right(Hiscore, Len(Hiscore) - 13)

'Process storing of highscore if new score exceeds the old score
    If GameScore > Val(Hiscore) Then    'Checks if current score is higher than the saved hiscore
    Dim Temp_String As String
        
        Open (Filename) For Output As file2 'Store all the player records into the file
            Temp_String = "Send this file along with your name to:  ruby@starpulse.com"
            Write #file2, Temp_String
            
            Temp_String = "View Top 10 Pilots on http://www.planet-source-code.com/vb"
            Write #file2, Temp_String
            
            Temp_String = "High Score : " & GameScore
            Write #file2, Temp_String
            
            Temp_String = "Main Gun   : " & WBox(0).Caption
            Write #file2, Temp_String
            
            Temp_String = "Side Gun   : " & WBox(1).Caption
            Write #file2, Temp_String
            
            Temp_String = "Special    : " & WBox(2).Caption
            Write #file2, Temp_String
        
            Temp_String = "Shield     : " & WBox(3).Caption
            Write #file2, Temp_String
            
            Temp_String = "Generator  : " & WBox(4).Caption
            Write #file2, Temp_String
            
            Temp_String = "Engine     : " & WBox(5).Caption
            Write #file2, Temp_String
        
            Temp_String = "Duration   : " & GEC & " Frame(s)"
            Write #file2, Temp_String
    
            Temp_String = "Total Kill : " & KillCount & " Ship(s)"
            Write #file2, Temp_String
    
            If KilledBoss = True Then
                Temp_String = "Last Boss  : Destroyed"
            Else:
                Temp_String = "Last Boss  : Failed"
            End If
            Write #file2, Temp_String
    
            If Elasped = 0 Then Elasped = 1
            Temp_String = "Ave. FPS   : " & Trim(Str(Round(TotFPS / Elasped, 2)))
            Write #file2, Temp_String
                   
        Close file2
    End If
End Sub

Private Sub Major_HIT()
'Ship is hit without any shields

Dim i As Integer
Dim All_Gone As Boolean
Dim Dam_Inflicted As Boolean
    
    'Randomly lowers the sidegun/maingun weapon level by 1
    If GEC Mod 2 = 1 Then Call Alter_G(0, 0)
    If GEC Mod 2 = 0 Then Call Alter_G(1, 0)
    
    Call Send_MSG(" DANGER! ", 2, 2)
    ShakeIT = 1
    All_Gone = True     'All parts of the ship has been damaged
    Dam_Inflicted = False
    
    Select Case (FPS Mod 10)
    Case 9:                     '10% Chance of Direct HIT and get destroyed
        Call Critical_HIT
    Case Else:
        'If any of the ship's components are still intact, the all_gone is false
        If Lock_MainGun = False Then
            All_Gone = False
        End If
        If Lock_SideGun = False Then
            All_Gone = False
        End If
        If Lock_Engine = False Then
            All_Gone = False
        End If
        If Lock_Generator = False Then
            All_Gone = False
        End If
        
        'If all of the ship's components is damaged, then ship destroyed (death).
        If All_Gone = True Then
            Call Critical_HIT
        End If
        
        'If any part of the ship is still intact, perform
        'component failure on a random part.
        i = Int(FPS) Mod 4
        Do While Dam_Inflicted = False
            Select Case (i)
            Case 0:     'Fail Main Guns
                If Lock_MainGun = False Then
                    Lock_MainGun = True
                    Stt_MGUN_R.Width = 0
                    Dam_Inflicted = True
                End If
            Case 1:     'Fail Side Guns
                If Lock_SideGun = False Then
                    Lock_SideGun = True
                    Stt_SGUN_R.Width = 0
                    Dam_Inflicted = True
                End If
            Case 2:     'Fail Engines (cannot move)
                If Lock_Engine = False Then
                    Lock_Engine = True
                    Stt_ENG_R.Width = 0
                    Dam_Inflicted = True
                End If
            Case 3:     'Fail generator (stops charging)
                If Lock_Generator = False Then
                    Lock_Generator = True
                    Stt_PWR_R.Width = 0
                    Dam_Inflicted = True
                End If
            End Select
            
            'If component of the ship is already destroyed,
            'perform check on the component
            i = i + 1
            If i > 3 Then
                i = 0
            End If
            
            'If ship is already destroyed...
            If GEC < 0 Then Exit Do
        Loop
    End Select
End Sub

Private Sub Minor_HIT()
'Minor hit occur when the ship is inflicted damage
'while the shield is still intact. A bar of shiled will
'be deducted and shield charging will increase, if not already started
    If SHD_Charging = False Then
        BAR_SHD_R.Width = 0
    End If
    
    'If no more shields left, a major hit occurs
    If SHD_Avail = False Then
        Call Major_HIT
    Else
        'Minus one bar of shield
        BAR_SHD_C.Width = BAR_SHD_C.Width - 435
        SHIELD_ON_OFF = 1
        If BAR_SHD_C.Width <= 15 Then
            BAR_SHD_C.Width = 15
            Stt_SHD_Lbl.BackColor = &HC0&
            Stt_SHD_Lbl.ForeColor = &HFFFFFF
            SHD_Avail = False
        End If
    End If
    
    'Continue to charge shield
    SHD_Charging = True
End Sub

Private Sub Alter_G(MG_SG As Integer, Plus_Minus As Integer)
'This sub alter the Maingun and Sidegun levels.
'Parameterlist : 0/1 = Maingun/Sidegun
'                0/1 = Minus/Add Level
    If MG_SG = 0 Then
        If Plus_Minus = 0 Then
            MG_Lvl = MG_Lvl - 1
            If MG_Lvl <= 1 Then MG_Lvl = 1
        Else:
            MG_Lvl = MG_Lvl + 1
            If MG_Lvl >= 10 Then
                MG_Lvl = 10
                GameScore = GameScore + 2000    'Bonus points
            End If
            Call Send_MSG("MainGun Upgrade", 1, 0)
        End If
    Else:
        If Plus_Minus = 0 Then
            SG_Lvl = SG_Lvl - 1
            If SG_Lvl <= 1 Then SG_Lvl = 1
        Else:
            SG_Lvl = SG_Lvl + 1
            If SG_Lvl > 5 Then
                SG_Lvl = 5
                GameScore = GameScore + 2000    'Bonus points
                Call Alter_G(0, 1)      'Surplus sideguns goes to maingun
            Else:
                Call Send_MSG("SideGun Upgrade", 1, 0)
            End If
        End If
    End If
    
    'Changes the statsbar captions
    Stt_LVL_MG.Caption = "MG  " & MG_Lvl & ".0"
    Stt_LVL_SG.Caption = "SG   " & SG_Lvl & ".0"
End Sub

Private Sub Check_EDeath(x As Integer)
'This sub checks the lifepoints of an enemy after being
'inflicted damage. If the lifepoints is 0 or lower, then
'it will be destroyed.

Dim y As Integer
Dim temp As String

    If EShip(x).Life <= 0 And EShip(x).Used = True Then
        EShip(x).Used = False           'Destroyed
        KillCount = KillCount + 1       'Add to total kills
        If EShip(x).Has_Bonus = True Then Call Create_Bonus(EShip(x).PosX + (Img(EShip(x).Kind).ImageWidth * 15) / 2, EShip(x).PosY + (Img(EShip(x).Kind).ImageWidth * 15) / 2)     'If the ship has a powerup...
        Call PLAY(100)
        
        'If the ship destroyed was a BOSS
        If EShip(x).Boss = True And EShip(x).Life <= 0 Then
            Call PLAY(101)
            Call PLAY(101)
            Call PLAY(101)
            Call PLAY(101)
            
            'Proceed to next stage
            For y = 0 To 4
                If C_GEC(y) = 0 Then        'Next Level GEC counter initialized
                    C_GEC(y) = GEC
                    Exit For
                End If
            Next y
            If EShip(x).Boss = True Then
                Moving_BossBar = 2      'Hides bossbar
                Call EXPLODE(EShip(x).PosX + (Img(EShip(x).Kind).ImageWidth * 15) / 2, EShip(x).PosY + (Img(EShip(x).Kind).ImageHeight * 15) / 2, 2)
            End If
            EShip(x).Boss = False
        End If
        
        GameScore = GameScore + EShip(x).Score
        temp = ""
                                
        'Adds the score to the display on screen
        For y = 0 To ((6 - Len(Trim(Str(GameScore)))) - 1)
            temp = temp & "0"
        Next y
        temp = temp & Trim(Str(GameScore))
        BufferText(6).Caption = temp
        
        Call EXPLODE(EShip(x).PosX + (Img(EShip(x).Kind).ImageWidth * 15) / 2, EShip(x).PosY + (Img(EShip(x).Kind).ImageHeight * 15) / 2, 1)
    End If
End Sub

Private Sub tmOneSecond_Timer()     'Frames Per Second Indicator
'This sub is used to keep track of the player's average
'FRAME per second. It just merely shows the performance of the
'game on the player's PC
    Elasped = Elasped + 1
    TotFPS = TotFPS + FPS
    FPS = 0
    FormBW.Caption = "Black Winter II : Final Assault" & "    Fps : " & Str(Round(TotFPS / Elasped, 2))
End Sub

Private Sub tmOneFrame_Timer()
'This is the Most important sub, where every single
'frame of gameplay is drawn and animated frame by frame here.
'We'll go through this bit by bit ...

Dim B_Speed As Integer  'Bullet Speed
Dim x As Integer
Dim i As Integer
Dim y As Integer
    'This is to supply the information needed in the
    'previous sub to calculate the frames per second
    If PlayMode_Enabled = True Then
        Play_Layer.Cls  'CLS + FPS
        FPS = FPS + 1
        GEC = GEC + 1
        
        'If the background music is stopped, then is replayed
        If GEC Mod 10 = 1 Then Call PLAY(99)
        
        '-----------------------------------------
                                    'DEATH
    'If the ship is destroyed, the following messages will be displayed
    If GEC = -380 Then Call Send_MSG(" > Command Base: Congratulations, Black Winter.", 0, 0)
    If GEC = -330 Then Call Send_MSG(" MISSION COMPLETED ", 2, 2)
    If GEC = -260 Then Call Send_MSG(" > Check HiScore.txt ", 0, 0)
    If GEC = -220 Then Call Send_MSG(" > Compete with the BEST Pilots Worldwide!", 0, 0)
    If GEC = -180 Then Call Send_MSG(" > Send your HiScore files to: ", 0, 0)
    If GEC = -140 Then Call Send_MSG(" > ruby@starpulse.com ", 3, 0)
    If GEC = -100 Then Call Send_MSG(" GAME OVER ", 0, 1)
    If GEC = -1 Then
        'Game no longer running, returns to main menu
        GameStarted = False
        Call Show_Hide(1, 4, 0.2, 1)
    End If
        '-----------------------------------------
                                    'DRAWING A STARFIELD
        'This part of the code is modified from
        'Johnathan Roach's starfield codes.
        'Refer to his submission Project Space Demo
        'on planet source code
        For i = 0 To Detail_Level
            Star(i).PosY = Star(i).PosY + Star(i).Spd
            If Star(i).PosY > 7920 Then
                Star(i).PosY = 1440
                Star(i).PosX = (Rnd * 4920 + 720)
                Star(i).Spd = (Rnd * 40 + 20)
                Star(i).Color = (Rnd * 2 + 1)
            End If
            Select Case (Star(i).Color)
            Case 1:
                Play_Layer.PSet (Star(i).PosX, Star(i).PosY), &HFFFFFF
            Case 1:
                Play_Layer.PSet (Star(i).PosX, Star(i).PosY), &HE0E0E0
            Case Else:
                Play_Layer.PSet (Star(i).PosX, Star(i).PosY), &HC0C0C0
            End Select
        Next i
        
        '-----------------------------------------
                                    'CREATING BULLET OBJECT
        'If gun is firing, a new bullet is created
        If Lock_MainGun = False And Lock_All = False And Ship_Firing = True Then
        
        'The following codes are for weapons.
        'Most of the have common variables.
        'Bull_Delay is the delay between 2 bullets
        '   fired in succession
        'B_Speed is the speed of the bullet moving
        '   across their X and Y corrdinates
        'Bull_Type_Limit is the maximum number of
        '   bullets that can appear on the screen.
        '   Old bullets need to be 'hit' or flies
        '   out of the playing area before new ones can be
        '   created.
        'As you'll notice, as the MG_Lvl (Maingun Level)
        '   and SG_Lvl (Sideguun Level) increases, the
        '   number of bullets, bull_delay, b_speed,
        '   bull_type_limit all changes to increase
        '   the weapons' efficiency.
        '   Certain weapons have their own custom
        '   variables to define their custom pattern.
        '   Most of these are self-explainary and can
        '   be understood by studying the weapon's code.
        
            If SHP(0) = 0 Then          'VULCAN
                Const Vulcan_Red_Damage = 15
                If Bull_Delay = 0 Then
                    B_Speed = 340
                    If MG_Lvl >= 1 Then '----
                        Bull_Delay = 7
                        Bull_Type_Limit = 2
                        If MG_Lvl >= 2 Then '----
                            Bull_Type_Limit = 5
                            If MG_Lvl >= 3 Then '----
                                Bull_Delay = 6
                                Bull_Type_Limit = 19
                                If MG_Lvl >= 4 Then '----
                                    If MG_Lvl >= 5 Then '----
                                        Bull_Delay = 5
                                        Bull_Type_Limit = 29
                                        If MG_Lvl >= 6 Then '----
                                            If MG_Lvl >= 7 Then '----
                                                Bull_Type_Limit = 39
                                                If MG_Lvl >= 8 Then '----
                                                    Bull_Delay = 4
                                                    Bull_Type_Limit = 49
                                                    If MG_Lvl >= 9 Then '----
                                                        Bull_Type_Limit = 59
                                                        If MG_Lvl >= 10 Then
                                                            Bull_Delay = 3
                                                            Bull_Type_Limit = 99
                                                        End If '-End 10
                                                        For i = 0 To Bull_Type_Limit
                                                            If Bull(i).Used = False Then
                                                                Bull(i).Used = True
                                                                Bull(i).Kind = 1
                                                                Bull(i).Dire = 8
                                                                Bull(i).Span = 0
                                                                Bull(i).Dama = Vulcan_Red_Damage
                                                                Bull(i).PosX = ShipX + -100
                                                                Bull(i).PosY = ShipY + 320
                                                                Bull(i).Spd = B_Speed
                                                                Exit For
                                                            End If
                                                        Next i
                                                        For i = 0 To Bull_Type_Limit
                                                            If Bull(i).Used = False Then
                                                                Bull(i).Used = True
                                                                Bull(i).Kind = 1
                                                                Bull(i).Dire = 8
                                                                Bull(i).Span = 0
                                                                Bull(i).Dama = Vulcan_Red_Damage
                                                                Bull(i).PosX = ShipX + 620
                                                                Bull(i).PosY = ShipY + 320
                                                                Bull(i).Spd = B_Speed
                                                                Exit For
                                                            End If
                                                        Next i
                                                    End If '-End 9
                                                End If
                                                For i = 0 To Bull_Type_Limit
                                                    If Bull(i).Used = False Then
                                                        Bull(i).Used = True
                                                        Bull(i).Kind = 6
                                                        Bull(i).Dire = 73
                                                        Bull(i).Span = 0
                                                        Bull(i).Dama = Vulcan_Red_Damage
                                                        Bull(i).PosX = ShipX + 100
                                                        Bull(i).PosY = ShipY + 220
                                                        Bull(i).Spd = B_Speed
                                                        Exit For
                                                    End If
                                                Next i
                                                For i = 0 To Bull_Type_Limit
                                                    If Bull(i).Used = False Then
                                                        Bull(i).Used = True
                                                        Bull(i).Kind = 7
                                                        Bull(i).Dire = 93
                                                        Bull(i).Span = 0
                                                        Bull(i).Dama = Vulcan_Red_Damage
                                                        Bull(i).PosX = ShipX + 420
                                                        Bull(i).PosY = ShipY + 220
                                                        Bull(i).Spd = B_Speed
                                                        Exit For
                                                    End If
                                                Next i
                                            End If  '-End 7
                                            For i = 0 To Bull_Type_Limit
                                                If Bull(i).Used = False Then
                                                    Bull(i).Used = True
                                                    Bull(i).Kind = 4
                                                    Bull(i).Dire = 72
                                                    Bull(i).Span = 0
                                                    Bull(i).Dama = Vulcan_Red_Damage
                                                    Bull(i).PosX = ShipX + 50
                                                    Bull(i).PosY = ShipY + 320
                                                    Bull(i).Spd = B_Speed
                                                    Exit For
                                                End If
                                            Next i
                                            For i = 0 To Bull_Type_Limit
                                                If Bull(i).Used = False Then
                                                    Bull(i).Used = True
                                                    Bull(i).Kind = 5
                                                    Bull(i).Dire = 92
                                                    Bull(i).Span = 0
                                                    Bull(i).Dama = Vulcan_Red_Damage
                                                    Bull(i).PosX = ShipX + 470
                                                    Bull(i).PosY = ShipY + 320
                                                    Bull(i).Spd = B_Speed
                                                    Exit For
                                                End If
                                            Next i
                                        End If  '-End 6
                                    End If  '-End 5
                                    For i = 0 To Bull_Type_Limit
                                        If Bull(i).Used = False Then
                                            Bull(i).Used = True
                                            Bull(i).Kind = 3
                                            Bull(i).Dire = 7
                                            Bull(i).Span = 0
                                            Bull(i).Dama = Vulcan_Red_Damage
                                            Bull(i).PosX = ShipX + 50
                                            Bull(i).PosY = ShipY + 320
                                            Bull(i).Spd = B_Speed
                                            Exit For
                                        End If
                                    Next i
                                    For i = 0 To Bull_Type_Limit
                                        If Bull(i).Used = False Then
                                            Bull(i).Used = True
                                            Bull(i).Kind = 2
                                            Bull(i).Dire = 9
                                            Bull(i).Span = 0
                                            Bull(i).Dama = Vulcan_Red_Damage
                                            Bull(i).PosX = ShipX + 470
                                            Bull(i).PosY = ShipY + 320
                                            Bull(i).Spd = B_Speed
                                            Exit For
                                        End If
                                    Next i
                                End If      '-End 4
                            End If          '-End 3
                            For i = 0 To Bull_Type_Limit
                                If Bull(i).Used = False Then
                                    Bull(i).Used = True
                                    Bull(i).Kind = 1
                                    Bull(i).Dire = 8
                                    Bull(i).Span = 0
                                    Bull(i).Dama = Vulcan_Red_Damage
                                    Bull(i).PosX = ShipX + 400
                                    Bull(i).PosY = ShipY + 120
                                    Bull(i).Spd = B_Speed
                                    Exit For
                                End If
                            Next i
                        End If              '-End 2
                        For i = 0 To Bull_Type_Limit
                            If Bull(i).Used = False Then
                                Bull(i).Used = True
                                Bull(i).Kind = 1
                                Bull(i).Dire = 8
                                Bull(i).Span = 0
                                Bull(i).Dama = Vulcan_Red_Damage
                                If MG_Lvl = 1 Then
                                    Bull(i).PosX = ShipX + 260
                                Else:
                                    Bull(i).PosX = ShipX + 140
                                End If
                                Bull(i).PosY = ShipY + 120
                                Bull(i).Spd = B_Speed
                                Exit For
                            End If
                        Next i
                    End If                  '-End 1
                End If
            End If  '-END VULCAN
            
               
            If SHP(0) = 1 Then                  'LIGHT BEAM
            Dim Laser_OffSet As Integer
            Dim Laser_Style As Integer
            Dim Laser_Damage As Integer
            Dim Laser_PosX As Integer
                If Bull_Delay = 0 Then
                    If MG_Lvl >= 1 Then '----
                        B_Speed = 650
                        Laser_Style = 1
                        Laser_OffSet = 200
                        Laser_Damage = 3
                        Bull_Delay = 1
                        Bull_Type_Limit = 10
                        If MG_Lvl >= 2 Then '----
                            Laser_Style = 2
                            Laser_Damage = 5
                        End If
                        If MG_Lvl >= 3 Then '----
                            Laser_Style = 3
                            Laser_Damage = 7
                        End If
                        If MG_Lvl >= 4 Then '----
                            Laser_Style = 4
                            Laser_Damage = 10
                        End If
                        If MG_Lvl >= 5 Then '----
                            Laser_Style = 5
                            Laser_Damage = 12
                        End If
                        If MG_Lvl >= 6 Then '----
                            Laser_Style = 6
                            Laser_Damage = 15
                        End If
                        If MG_Lvl >= 7 Then '----
                            Laser_Style = 7
                            Laser_Damage = 18
                        End If
                        If MG_Lvl >= 8 Then '----
                            Laser_Style = 8
                            Laser_Damage = 20
                        End If
                        If MG_Lvl >= 9 Then '----
                            Laser_Style = 9
                            Laser_Damage = 25
                        End If
                        If MG_Lvl >= 10 Then '----
                            Laser_Style = 10
                            Laser_Damage = 35
                        End If
                        For i = 0 To Bull_Type_Limit
                            If Bull(i).Used = False Then
                                Bull(i).Used = True
                                Bull(i).Kind = Laser_Style
                                Bull(i).Dire = 8
                                Bull(i).Span = 0
                                Bull(i).Dama = Laser_Damage
                                Bull(i).PosX = ShipX + Laser_OffSet
                                Bull(i).PosY = ShipY + 0
                                Bull(i).Spd = B_Speed
                                
                                Laser_PosX = Bull(i).PosX
                                Exit For
                            End If
                        Next i
                    End If  '-End 1
                End If
            End If  '-END LIGHT BEAM
            
            If SHP(0) = 2 Then                  'PARTICLE BURST
                Static Random_Bound As Integer
                Static Particle_StartX(150) As Integer 'Start posX of particle weapon
                Static Particle_StartY(150) As Integer 'Start posY
                Static Particle_EndX(150) As Integer   'Destination PosX
                Static Particle_EndY(150) As Integer   'Destination PosY
                Static Particle_Range As Integer       'Determine Span of particle weapons
                If Bull_Delay = 0 Then
                    Random_Bound = 3
                    If MG_Lvl >= 1 Then '----
                        B_Speed = 0 'Not Used
                        Bull_Delay = 3
                        Bull_Type_Limit = 150
                        Particle_Range = 13
                        If MG_Lvl >= 2 Then '----
                            Random_Bound = 9
                        End If
                        If MG_Lvl >= 3 Then '----
                            Bull_Delay = 2
                        End If
                        If MG_Lvl >= 4 Then '----
                            For i = 0 To Bull_Type_Limit
                                If Bull(i).Used = False Then
                                    Bull(i).Used = True
                                    Bull(i).Kind = (Rnd * Random_Bound) + 1 + (23 - Random_Bound)
                                    Bull(i).Dire = 79
                                    Bull(i).Span = 0
                                    If Bull(i).Kind >= 20 Then Bull(i).Dama = 6
                                    If Bull(i).Kind >= 16 And Bull(i).Kind <= 19 Then Bull(i).Dama = 8
                                    If Bull(i).Kind >= 11 And Bull(i).Kind <= 15 Then Bull(i).Dama = 12
                                    If Bull(i).Kind >= 5 And Bull(i).Kind <= 10 Then Bull(i).Dama = 16
                                    If Bull(i).Kind >= 1 And Bull(i).Kind <= 4 Then Bull(i).Dama = 24
                                    Bull(i).PosX = ShipX + 270
                                    Bull(i).PosY = ShipY + 120
                                    Bull(i).Spd = B_Speed
                                
                                    Particle_StartX(i) = Bull(i).PosX
                                    Particle_StartY(i) = Bull(i).PosY
                                    Particle_EndX(i) = Bull(i).PosX - 1500 + (Rnd * 3000)
                                    Particle_EndY(i) = Bull(i).PosY - 6200 + (Rnd * 2000)
                                
                                    Exit For
                                End If
                            Next i
                        End If  '-End 3
                        If MG_Lvl >= 5 Then '----
                            Random_Bound = 14
                        End If
                        If MG_Lvl >= 6 Then '----
                            For i = 0 To Bull_Type_Limit
                                If Bull(i).Used = False Then
                                    Bull(i).Used = True
                                    Bull(i).Kind = (Rnd * Random_Bound) + 1 + (23 - Random_Bound)
                                    Bull(i).Dire = 79
                                    Bull(i).Span = 0
                                    If Bull(i).Kind >= 20 Then Bull(i).Dama = 6
                                    If Bull(i).Kind >= 16 And Bull(i).Kind <= 19 Then Bull(i).Dama = 8
                                    If Bull(i).Kind >= 11 And Bull(i).Kind <= 15 Then Bull(i).Dama = 12
                                    If Bull(i).Kind >= 5 And Bull(i).Kind <= 10 Then Bull(i).Dama = 16
                                    If Bull(i).Kind >= 1 And Bull(i).Kind <= 4 Then Bull(i).Dama = 24
                                    Bull(i).PosX = ShipX + 270
                                    Bull(i).PosY = ShipY + 120
                                    Bull(i).Spd = B_Speed
                                
                                    Particle_StartX(i) = Bull(i).PosX
                                    Particle_StartY(i) = Bull(i).PosY
                                    Particle_EndX(i) = Bull(i).PosX - 1500 + (Rnd * 3000)
                                    Particle_EndY(i) = Bull(i).PosY - 6200 + (Rnd * 2000)
                                
                                    Exit For
                                End If
                            Next i
                        End If      '-End 6
                        If MG_Lvl >= 7 Then '----
                            Random_Bound = 19
                        End If
                        If MG_Lvl >= 8 Then '----
                            For i = 0 To Bull_Type_Limit
                                If Bull(i).Used = False Then
                                    Bull(i).Used = True
                                    Bull(i).Kind = (Rnd * Random_Bound) + 1 + (23 - Random_Bound)
                                    Bull(i).Dire = 79
                                    Bull(i).Span = 0
                                    If Bull(i).Kind >= 20 Then Bull(i).Dama = 6
                                    If Bull(i).Kind >= 16 And Bull(i).Kind <= 19 Then Bull(i).Dama = 8
                                    If Bull(i).Kind >= 11 And Bull(i).Kind <= 15 Then Bull(i).Dama = 12
                                    If Bull(i).Kind >= 5 And Bull(i).Kind <= 10 Then Bull(i).Dama = 16
                                    If Bull(i).Kind >= 1 And Bull(i).Kind <= 4 Then Bull(i).Dama = 24
                                    Bull(i).PosX = ShipX + 270
                                    Bull(i).PosY = ShipY + 120
                                    Bull(i).Spd = B_Speed
                                
                                    Particle_StartX(i) = Bull(i).PosX
                                    Particle_StartY(i) = Bull(i).PosY
                                    Particle_EndX(i) = Bull(i).PosX - 1500 + (Rnd * 3000)
                                    Particle_EndY(i) = Bull(i).PosY - 6200 + (Rnd * 2000)
                                
                                    Exit For
                                End If
                            Next i
                        End If  '-End 8
                        If MG_Lvl >= 9 Then '----
                            Random_Bound = 23
                            Bull_Delay = 1
                        End If
                        If MG_Lvl >= 10 Then '----
                            For i = 0 To Bull_Type_Limit
                                If Bull(i).Used = False Then
                                    Bull(i).Used = True
                                    Bull(i).Kind = (Rnd * Random_Bound) + 1 + (23 - Random_Bound)
                                    Bull(i).Dire = 79
                                    Bull(i).Span = 0
                                    If Bull(i).Kind >= 20 Then Bull(i).Dama = 6
                                    If Bull(i).Kind >= 16 And Bull(i).Kind <= 19 Then Bull(i).Dama = 8
                                    If Bull(i).Kind >= 11 And Bull(i).Kind <= 15 Then Bull(i).Dama = 12
                                    If Bull(i).Kind >= 5 And Bull(i).Kind <= 10 Then Bull(i).Dama = 16
                                    If Bull(i).Kind >= 1 And Bull(i).Kind <= 4 Then Bull(i).Dama = 24
                                    Bull(i).PosX = ShipX + 270
                                    Bull(i).PosY = ShipY + 120
                                    Bull(i).Spd = B_Speed
                                
                                    Particle_StartX(i) = Bull(i).PosX
                                    Particle_StartY(i) = Bull(i).PosY
                                    Particle_EndX(i) = Bull(i).PosX - 1500 + (Rnd * 3000)
                                    Particle_EndY(i) = Bull(i).PosY - 6200 + (Rnd * 2000)
                                
                                    Exit For
                                End If
                            Next i
                        End If  '-End 10
                        For i = 0 To Bull_Type_Limit
                            If Bull(i).Used = False Then
                                Bull(i).Used = True
                                Bull(i).Kind = (Rnd * Random_Bound) + 1 + (23 - Random_Bound)
                                Bull(i).Dire = 79
                                Bull(i).Span = 0
                                If Bull(i).Kind >= 20 Then Bull(i).Dama = 6
                                If Bull(i).Kind >= 16 And Bull(i).Kind <= 19 Then Bull(i).Dama = 8
                                If Bull(i).Kind >= 11 And Bull(i).Kind <= 15 Then Bull(i).Dama = 12
                                If Bull(i).Kind >= 5 And Bull(i).Kind <= 10 Then Bull(i).Dama = 16
                                If Bull(i).Kind >= 1 And Bull(i).Kind <= 4 Then Bull(i).Dama = 24
                                Bull(i).PosX = ShipX + 270
                                Bull(i).PosY = ShipY + 120
                                Bull(i).Spd = B_Speed
                                
                                Particle_StartX(i) = Bull(i).PosX
                                Particle_StartY(i) = Bull(i).PosY
                                Particle_EndX(i) = Bull(i).PosX - 1500 + (Rnd * 3000)
                                Particle_EndY(i) = Bull(i).PosY - 6200 + (Rnd * 2000)
                                
                                Exit For
                            End If
                        Next i
                    End If  '-End 1
                End If
            End If  '-END PARTICLE BURST
            
            If SHP(0) = 3 Then                  'PLASMA WAVE
            Dim Wave_Pattern As Integer
                If Bull_Delay = 0 Then
                    If MG_Lvl >= 1 Then '----
                        B_Speed = 0 'Not Used
                        Bull_Delay = 6
                        Bull_Type_Limit = 2
                        Wave_Pattern = 1
                        If MG_Lvl >= 2 Then
                            Wave_Pattern = 2
                        End If
                        If MG_Lvl >= 3 Then
                            Wave_Pattern = 3
                            Bull_Delay = 5
                        End If
                        If MG_Lvl >= 4 Then
                            Wave_Pattern = 4
                            Bull_Delay = 4
                        End If
                        If MG_Lvl >= 5 Then
                            Wave_Pattern = 5
                            Bull_Type_Limit = 3
                        End If
                        If MG_Lvl >= 6 Then
                            Wave_Pattern = 6
                            Bull_Type_Limit = 4
                        End If
                        If MG_Lvl >= 7 Then
                            Wave_Pattern = 7
                            Bull_Delay = 3
                        End If
                        If MG_Lvl >= 8 Then
                            Wave_Pattern = 8
                            Bull_Delay = 2
                        End If
                        If MG_Lvl >= 9 Then
                            Wave_Pattern = 9
                            Bull_Delay = 1
                        End If
                        If MG_Lvl >= 10 Then
                            Wave_Pattern = 10
                            Bull_Type_Limit = 5
                        End If
                        For i = 0 To Bull_Type_Limit
                            If Bull(i).Used = False Then
                                Bull(i).Used = True
                                Bull(i).Kind = Wave_Pattern
                                Bull(i).Dire = 8
                                Bull(i).Span = 0
                                Bull(i).Dama = 4 * MG_Lvl
                                Bull(i).PosX = ShipX + -540
                                Bull(i).PosY = ShipY - 200
                                Bull(i).Spd = B_Speed
                                Exit For
                            End If
                        Next i
                    End If  '-End 1
                End If
            End If  '-End PLASMA WAVE
            
            If SHP(0) = 4 Then                  'VULCAN BLUE
                If Bull_Delay = 0 Then
                    Const Vulcan_Blue_Damage = 20
                    If MG_Lvl >= 1 Then '----
                        B_Speed = 300 'Not Used
                        Bull_Delay = 8
                        Bull_Type_Limit = 80
                        If MG_Lvl >= 2 Then
                            Bull_Delay = 6
                        End If
                        If MG_Lvl >= 3 Then
                            B_Speed = 350
                        End If
                        If MG_Lvl >= 4 Then
                            If MG_Lvl >= 5 Then
                                Bull_Delay = 4
                            End If
                            If MG_Lvl >= 6 Then
                                B_Speed = 400
                            End If
                            If MG_Lvl >= 7 Then
                                Bull_Delay = 3
                            End If
                            If MG_Lvl >= 8 Then
                                B_Speed = 450
                            End If
                            If MG_Lvl >= 9 Then
                                If MG_Lvl >= 10 Then
                                    B_Speed = 800
                                End If
                                For i = 0 To Bull_Type_Limit
                                    If Bull(i).Used = False Then
                                        Bull(i).Used = True
                                        Bull(i).Kind = 1
                                        Bull(i).Dire = 8
                                        Bull(i).Span = 0
                                        Bull(i).Dama = Vulcan_Blue_Damage
                                        Bull(i).PosX = ShipX + 940
                                        Bull(i).PosY = ShipY + 600
                                        Bull(i).Spd = B_Speed
                                        Exit For
                                    End If
                                Next i
                                For i = 0 To Bull_Type_Limit
                                    If Bull(i).Used = False Then
                                        Bull(i).Used = True
                                        Bull(i).Kind = 1
                                        Bull(i).Dire = 8
                                        Bull(i).Span = 0
                                        Bull(i).Dama = Vulcan_Blue_Damage
                                        Bull(i).PosX = ShipX - 460
                                        Bull(i).PosY = ShipY + 600
                                        Bull(i).Spd = B_Speed
                                        Exit For
                                    End If
                                Next i
                            End If  '-End 9
                            For i = 0 To Bull_Type_Limit
                                If Bull(i).Used = False Then
                                    Bull(i).Used = True
                                    Bull(i).Kind = 1
                                    Bull(i).Dire = 8
                                    Bull(i).Span = 0
                                    Bull(i).Dama = Vulcan_Blue_Damage
                                    Bull(i).PosX = ShipX + 580
                                    Bull(i).PosY = ShipY + 300
                                    Bull(i).Spd = B_Speed
                                    Exit For
                                End If
                            Next i
                            For i = 0 To Bull_Type_Limit
                                If Bull(i).Used = False Then
                                    Bull(i).Used = True
                                    Bull(i).Kind = 1
                                    Bull(i).Dire = 8
                                    Bull(i).Span = 0
                                    Bull(i).Dama = Vulcan_Blue_Damage
                                    Bull(i).PosX = ShipX - 100
                                    Bull(i).PosY = ShipY + 300
                                    Bull(i).Spd = B_Speed
                                    Exit For
                                End If
                            Next i
                        End If  '-End 4
                        For i = 0 To Bull_Type_Limit
                            If Bull(i).Used = False Then
                                Bull(i).Used = True
                                Bull(i).Kind = 1
                                Bull(i).Dire = 8
                                Bull(i).Span = 0
                                Bull(i).Dama = Vulcan_Blue_Damage
                                Bull(i).PosX = ShipX + 240
                                Bull(i).PosY = ShipY - 100
                                Bull(i).Spd = B_Speed
                                Exit For
                            End If
                        Next i
                    End If  '-End 1
                End If
            End If  '-End Vulcan Blue



        End If  '-END FOR ALL MG's
        
        
        If Lock_SideGun = False And Lock_All = False And Ship_Firing = True Then
            If SHP(1) = 0 Then                  'MICRO MISSILES
                If Bull_Delay2 = 0 Then
                Static Micro_LR As Integer
                    If SG_Lvl >= 1 Then '----
                        B_Speed = 0 'Not Used
                        Bull_Delay2 = 1
                        Bull_Type_Limit = 1
                        If SG_Lvl >= 2 Then '----
                            Bull_Type_Limit = 3
                        End If
                        If SG_Lvl >= 3 Then '----
                            Bull_Type_Limit = 5
                        End If
                        If SG_Lvl >= 4 Then '----
                            Bull_Type_Limit = 7
                        End If
                        If SG_Lvl >= 5 Then '----
                            Bull_Type_Limit = 9
                        End If
                        
                        If Micro_LR = 0 Then
                            For i = 200 To (Bull_Type_Limit + 200)
                                If Bull(i).Used = False Then
                                    Bull(i).Used = True
                                    Bull(i).Kind = 1
                                    Bull(i).Dire = 8
                                    Bull(i).Span = 0
                                    Bull(i).Dama = 5 + (5 * SG_Lvl)
                                    Bull(i).PosX = ShipX + 0
                                    Bull(i).PosY = ShipY + 100
                                    Bull(i).Spd = B_Speed
                                    Exit For
                                End If
                            Next i
                            Micro_LR = 1
                        Else:
                            Micro_LR = 0
                            For i = 200 To (Bull_Type_Limit + 200)
                                If Bull(i).Used = False Then
                                    Bull(i).Used = True
                                    Bull(i).Kind = 1
                                    Bull(i).Dire = 8
                                    Bull(i).Span = 0
                                    Bull(i).Dama = 5 + (5 * SG_Lvl)
                                    Bull(i).PosX = ShipX + 560
                                    Bull(i).PosY = ShipY + 100
                                    Bull(i).Spd = B_Speed
                                    Exit For
                                End If
                            Next i
                        End If
                    End If '-End 1
                End If
            End If  '-END FOR MICRO MISSILES
            
            If SHP(1) = 1 Then                  'PULSE RING
                If Bull_Delay2 = 0 Then
                    Static Pulse_Angle As Integer
                    If SG_Lvl >= 1 Then '----
                        B_Speed = 0 'Not Used
                        Bull_Delay2 = 1
                        Bull_Type_Limit = 0
                        Pulse_Angle = Pulse_Angle + 1
                        If SG_Lvl >= 1 And SG_Lvl <= 2 Then
                            If Pulse_Angle > 3 Then Pulse_Angle = 1
                        End If
                        If SG_Lvl >= 3 And SG_Lvl <= 4 Then
                            If Pulse_Angle > 6 Then Pulse_Angle = 4
                        End If
                        If SG_Lvl >= 5 Then
                            If Pulse_Angle > 9 Then Pulse_Angle = 7
                        End If
                        For i = 200 To 300
                            If Bull(i).Used = False Then
                                Bull(i).Used = True
                                Bull(i).Kind = Pulse_Angle
                                Bull(i).Dire = 8
                                Bull(i).Span = 0
                                Bull(i).Dama = 5 * SG_Lvl
                                Bull(i).PosX = ShipX - 600
                                Bull(i).PosY = ShipY - 700
                                Bull(i).Spd = B_Speed
                                Exit For
                            End If
                        Next i
                    End If  '-End 1
                End If
            End If      'END FOR PULSE RING
            
            If SHP(1) = 2 Then                  'CRESCENT BLADES
                If Bull_Delay2 = 0 Then
                Dim Cres_Angle As Integer
                Static CresLR(100) As Integer
                    If SG_Lvl >= 1 Then '----
                        B_Speed = 0 'Not Used
                        Bull_Delay2 = 5
                        Bull_Type_Limit = 1
                        Cres_Angle = Cres_Angle + 1
                        If SG_Lvl >= 2 Then
                            Bull_Type_Limit = 3
                        End If
                        If SG_Lvl >= 3 Then
                            Bull_Type_Limit = 5
                        End If
                        If SG_Lvl >= 4 Then
                            Bull_Type_Limit = 7
                        End If
                        If SG_Lvl >= 5 Then
                            Bull_Type_Limit = 9
                        End If
                        For i = 200 To (200 + Bull_Type_Limit)
                            If Bull(i).Used = False Then
                                CresLR(i - 200) = i Mod 2
                                Bull(i).Used = True
                                Bull(i).Kind = Cres_Angle
                                Bull(i).Dire = 8
                                Bull(i).Span = 0
                                Bull(i).Dama = 5 + (SG_Lvl * 15)
                                Bull(i).PosX = ShipX + 150
                                Bull(i).PosY = ShipY - 0
                                Bull(i).Spd = B_Speed
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If      '-END FOR CRESCENT BLADES
            
            If SHP(1) = 3 Then                  'CREEP MINE
                If Bull_Delay2 = 0 Then
                    If SG_Lvl >= 1 Then '----
                        B_Speed = 50 'Not Used
                        Bull_Delay2 = 14
                        Bull_Type_Limit = 2
                        If SG_Lvl >= 2 Then
                            Bull_Type_Limit = 3
                            Bull_Delay2 = 12
                        End If
                        If SG_Lvl >= 3 Then
                            Bull_Type_Limit = 5
                            Bull_Delay2 = 10
                        End If
                        If SG_Lvl >= 4 Then
                            Bull_Type_Limit = 7
                            Bull_Delay2 = 8
                        End If
                        If SG_Lvl >= 5 Then
                            Bull_Type_Limit = 10
                            Bull_Delay2 = 6
                        End If
                        For i = 200 To (200 + Bull_Type_Limit)
                            If Bull(i).Used = False Then
                                Bull(i).Used = True
                                Bull(i).Kind = 1
                                Bull(i).Dire = 8
                                Bull(i).Span = 0
                                Bull(i).Dama = 25 + (25 * SG_Lvl)
                                Bull(i).PosX = ShipX + 240
                                Bull(i).PosY = ShipY - 0
                                Bull(i).Spd = B_Speed
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If      '-END FOR CREEP MINE
            
        End If  '-END FOR ALL SIDE GUNS
        
        
        'Decrease the counter for the delay for the next bullet to appear
        Bull_Delay = Bull_Delay - 1
        Bull_Delay2 = Bull_Delay2 - 1
        If Bull_Delay <= 0 Then Bull_Delay = 0
        If Bull_Delay2 <= 0 Then Bull_Delay2 = 0
        
        
                                    'MANIPULATE EXISTING BULLET
        'This sub modifies the existing bullets created in the codes above.
        'Modifications are on their X and Y coordinates, and their availability
        For i = 0 To 300
            'If bullet is created and still in play field...
            If Bull(i).Used = True Then
                
                'Increases their time in play ...
                Bull(i).Span = Bull(i).Span + 1
                
                'These codes are for specialized weapons and their
                'behaviors. it doesn't affect other general weapons
                If (SHP(0) = 2 And i < 200) Then
                    If Bull(i).Span >= ((Rnd * 10) + Particle_Range) Then Bull(i).Used = False      'Particle Burst
                End If
                If (SHP(1) = 2 And i >= 200 And i < 290) Then
                    If Bull(i).Span < 100 Then Bull(i).PosY = Bull(i).PosY + (15 * Bull(i).Span)    'Crescent Blade
                End If
                If (SHP(1) = 3 And (i >= 200 And i < 290)) And (Bull(i).Span Mod 15 = 1) Then    'Creep Mine
                    Bull(i).Kind = 2
                Else:
                    If SHP(1) = 3 And i >= 200 And i < 290 Then Bull(i).Kind = 1
                End If
                                
                'These codes changes the Bullet's X/Y coordinates
                'based on their bullet's direction (Bull.dire)
                If Bull(i).Dire = 7 Then
                    Bull(i).PosY = Bull(i).PosY - Bull(i).Spd / 3 * 2
                    Bull(i).PosX = Bull(i).PosX - Bull(i).Spd / 3 * 2
                    If Bull(i).PosY <= 1200 Or Bull(i).PosX < 720 Then Bull(i).Used = False
                End If
                If Bull(i).Dire = 72 Then
                    Bull(i).PosY = Bull(i).PosY - Bull(i).Spd / 3 * 2
                    Bull(i).PosX = Bull(i).PosX - Bull(i).Spd / 3
                    If Bull(i).PosY <= 1200 Or Bull(i).PosX < 720 Then Bull(i).Used = False
                End If
                If Bull(i).Dire = 73 Then
                    Bull(i).PosY = Bull(i).PosY - Bull(i).Spd / 3 * 2
                    Bull(i).PosX = Bull(i).PosX - Bull(i).Spd / 7
                    If Bull(i).PosY <= 1200 Or Bull(i).PosX < 720 Then Bull(i).Used = False
                End If
                If Bull(i).Dire = 79 And i < 200 Then
                    Bull(i).PosX = Bull(i).PosX + ((Particle_EndX(i) - Particle_StartX(i)) / 20)
                    Bull(i).PosY = Bull(i).PosY - ((Particle_StartY(i) - Particle_EndY(i)) / 20)
                    If Bull(i).PosY <= 1200 Or Bull(i).PosX > 5655 Or Bull(i).PosX < 720 Then Bull(i).Used = False
                End If
                If Bull(i).Dire = 8 Then
                    If (SHP(0) = 3 And i < 200) Or (SHP(1) = 0 And i >= 200 And i < 290) Or (SHP(1) = 2 And i >= 200 And i < 290) Or (SHP(2) = 0 And i >= 290 And i <= 300) Then Bull(i).Spd = Bull(i).Span * Bull(i).Span 'Plasma Wave, Micro Missile, Crescent Blade, Inferno Disc
                    If (SHP(1) = 2 And i >= 200 And i < 290) Then
                        If CresLR(i - 200) = 0 Then Bull(i).PosX = Bull(i).PosX - ((-15 + Bull(i).Span) * ((i - 200) + 10)) 'cresent blades
                        If CresLR(i - 200) = 1 Then Bull(i).PosX = Bull(i).PosX + ((-15 + Bull(i).Span) * ((i - 200) + 10))
                    End If
                    Bull(i).PosY = Bull(i).PosY - Bull(i).Spd
                    If (SHP(0) = 1 And i < 200) Then Bull(i).PosX = Laser_PosX  'FOR LASER ONLY
                    If Bull(i).PosY <= 1200 Or Bull(i).PosX > 5655 Or Bull(i).PosX < 720 Then Bull(i).Used = False
                End If
                If Bull(i).Dire = 9 Then
                    Bull(i).PosY = Bull(i).PosY - Bull(i).Spd / 3 * 2
                    Bull(i).PosX = Bull(i).PosX + Bull(i).Spd / 3 * 2
                    If Bull(i).PosY <= 1200 Or Bull(i).PosX > 5655 Then Bull(i).Used = False
                End If
                If Bull(i).Dire = 92 Then
                    Bull(i).PosY = Bull(i).PosY - Bull(i).Spd / 3 * 2
                    Bull(i).PosX = Bull(i).PosX + Bull(i).Spd / 3
                    If Bull(i).PosY <= 1200 Or Bull(i).PosX < 720 Then Bull(i).Used = False
                End If
                If Bull(i).Dire = 93 Then
                    Bull(i).PosY = Bull(i).PosY - Bull(i).Spd / 3 * 2
                    Bull(i).PosX = Bull(i).PosX + Bull(i).Spd / 7
                    If Bull(i).PosY <= 1200 Or Bull(i).PosX < 720 Then Bull(i).Used = False
                End If
            End If
            
                                    'PAINT BULLET
            'If the bullet is still in play field, it
            'will be painted and show to the player.
            'Different ship configurations will show
            'different bullet sprites from their associated
            'imagelist...
            If Bull(i).Used = True Then
                If i < 200 Then
                    If SHP(0) = 0 Then Img(4).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                    If SHP(0) = 1 Then Img(6).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                    If SHP(0) = 2 Then Img(7).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                    If SHP(0) = 3 Then Img(8).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                    If SHP(0) = 4 Then Img(14).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                End If
                If i >= 200 And i < 290 Then
                    If SHP(1) = 0 Then Img(9).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                    If SHP(1) = 1 Then Img(10).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                    If SHP(1) = 2 Then
                        Bull(i).Kind = Bull(i).Kind + 1
                        If Bull(i).Kind > 4 Then Bull(i).Kind = 1
                        Img(11).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                    End If
                    If SHP(1) = 3 Then Img(12).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                End If
                If i >= 290 Then
                    If SHP(2) = 0 Then
                        Bull(i).Kind = Bull(i).Kind + 1
                        If Bull(i).Kind > 4 Then Bull(i).Kind = 1
                        Img(17).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                    End If
                    If SHP(2) = 1 Then
                        Bull(i).Kind = Bull(i).Kind + 1
                        If Bull(i).Kind > 4 Then Bull(i).Kind = 1
                        Img(5).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                    End If
                    If SHP(2) = 3 Then
                        Bull(i).Kind = Bull(i).Kind + 1
                        Img(3).ListImages(Bull(i).Kind).Draw Play_Layer.hDC, Bull(i).PosX, Bull(i).PosY, 1
                        If Bull(i).Kind > 21 Then
                            Bull(i).Used = False
                            Bull(i).Kind = 0
                        End If
                    End If
                End If
            End If
    
        Next i
        
        '-----------------------------------------
                                'MOVING ENEMY
        'This portion of code is similar to the moving patterns
        'of the player's bullet above.
        'At different time (ship's span), these enemy ships
        'will move with different speed and in different direction
        For i = 1 To 40
            If EShip(i).Used = True Then
                EShip(i).Span = EShip(i).Span + 1
                
                If EShip(i).Patt = 4646 Then  'Boss Left Right - Left Right
                    Static BossLR As Boolean
                    If EShip(i).PosY < 1560 Then EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                    If EShip(i).PosY >= 940 Then
                        If BossLR = True Then
                            EShip(i).PosX = EShip(i).PosX - EShip(i).Spd
                            If EShip(i).PosX < 980 Then BossLR = False
                        Else:
                            EShip(i).PosX = EShip(i).PosX + EShip(i).Spd
                            If EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) > 5380 Then BossLR = True
                        End If
                    End If
                    If EShip(i).Span Mod 10 = 1 Then Call Check_Fire(1, i)
                End If
                
                If EShip(i).Patt = 555 Then  'Burst-Burst
                    If EShip(i).Span Mod 40 < 10 Then Call Check_Fire(17, i)
                    EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                End If
                If EShip(i).Patt = 88 Then  'Straight Down
                    EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                    For x = 1 To 16
                        If EShip(i).Firing_Patt = x And EShip(i).Span Mod 25 = 20 Then Call Check_Fire(x, i)
                    Next x
                    If EShip(i).Firing_Patt = 32147 And EShip(i).Span Mod 7 = 1 Then Call Check_Fire(Rnd * 6 + 3, i)
                    If EShip(i).Firing_Patt = 12369 And EShip(i).Span Mod 7 = 1 Then Call Check_Fire(Rnd * 6 + 10, i)
                End If
                
                If EShip(i).Patt = 86 Then  'Down Curve right
                    If EShip(i).Span >= 20 And EShip(i).Span <= 39 Then EShip(i).Spd = 30
                    If EShip(i).Span >= 70 And EShip(i).Span <= 99 Then EShip(i).Spd = 15
                    If EShip(i).Span > 40 Then EShip(i).PosX = EShip(i).PosX + 30
                    EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                    If EShip(i).Span Mod 40 = 39 Then Call Check_Fire(EShip(i).Firing_Patt, i)
                End If
                
                If EShip(i).Patt = 84 Then  'Down Curve left
                    If EShip(i).Span >= 20 And EShip(i).Span <= 39 Then EShip(i).Spd = 30
                    If EShip(i).Span >= 70 And EShip(i).Span <= 99 Then EShip(i).Spd = 15
                    If EShip(i).Span > 40 Then EShip(i).PosX = EShip(i).PosX - 30
                    If EShip(i).Span Mod 40 = 39 Then Call Check_Fire(EShip(i).Firing_Patt, i)
                    EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                End If
                              
                If EShip(i).Patt = 821 Then  'Down then slow
                    If EShip(i).Span >= 20 And EShip(i).Span <= 59 Then EShip(i).Spd = 40
                    If EShip(i).Span >= 60 And EShip(i).Span <= 99 Then EShip(i).Spd = 20
                    If EShip(i).Span >= 100 Then EShip(i).Spd = 10
                    If EShip(i).Span Mod 35 = 34 Then Call Check_Fire(EShip(i).Firing_Patt, i)
                    EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                End If
                
                If EShip(i).Patt = 83 Then  'Left Diagonal
                    If EShip(i).Span >= 30 Then EShip(i).PosX = EShip(i).PosX + 35
                    If EShip(i).Span Mod 35 = 34 Then Call Check_Fire(EShip(i).Firing_Patt, i)
                    EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                End If
                If EShip(i).Patt = 81 Then  'Right Diagonal
                    If EShip(i).Span >= 30 Then EShip(i).PosX = EShip(i).PosX - 35
                    If EShip(i).Span Mod 35 = 34 Then Call Check_Fire(EShip(i).Firing_Patt, i)
                    EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                End If
                
                If EShip(i).Patt = 31313 Then  'Meteor Jiggly
                    EShip(i).PosX = EShip(i).PosX - 5
                    EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                    If EShip(i).Span Mod (Rnd * 40 + 1) = 1 Then EShip(i).Frame = EShip(i).Frame + 1
                    If EShip(i).Frame > 4 Then EShip(i).Frame = 1
                End If
                
                If EShip(i).Patt = 3311 Then  'Meteor Boss
                    Static M_Dir As Integer
                    If EShip(i).Span < 25 Then EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                    If EShip(i).Span Mod 30 = 1 Then M_Dir = (Rnd * 2)
                    If EShip(i).Span Mod 100 < 25 Then EShip(i).PosX = EShip(i).PosX - 75
                    If EShip(i).Span Mod 100 > 50 And EShip(i).Span Mod 100 < 75 Then EShip(i).PosX = EShip(i).PosX + 75
                    If EShip(i).Span Mod 100 < 25 Then Call Check_Fire(M_Dir + 17, i)
                    If EShip(i).Span Mod 100 > 50 And EShip(i).Span Mod 100 < 75 Then Call Check_Fire(M_Dir + 17, i)
                    If EShip(i).Span Mod 125 = 1 Then
                        For x = 1 To 20
                            Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                        Next x
                    End If
                End If
                            
                If EShip(i).Patt = 2288 Then    'Mother Bubble
                    If EShip(i).PosY < 1800 Then EShip(i).PosY = EShip(i).PosY + EShip(i).Spd
                    If EShip(i).Span Mod 200 < 25 Then
                        EShip(i).PosY = EShip(i).PosY + 60
                        EShip(i).PosX = EShip(i).PosX - 60
                    End If
                    If EShip(i).Span Mod 200 > 50 And EShip(i).Span Mod 200 < 75 Then
                        EShip(i).PosY = EShip(i).PosY + 60
                        EShip(i).PosX = EShip(i).PosX + 60
                    End If
                    If EShip(i).Span Mod 200 > 100 And EShip(i).Span Mod 200 < 125 Then
                        EShip(i).PosY = EShip(i).PosY - 60
                        EShip(i).PosX = EShip(i).PosX + 60
                    End If
                    If EShip(i).Span Mod 200 > 150 And EShip(i).Span Mod 200 < 175 Then
                        EShip(i).PosY = EShip(i).PosY - 60
                        EShip(i).PosX = EShip(i).PosX - 60
                    End If
                    If EShip(i).Span > 25 And EShip(i).Span < 50 Then
                        If EShip(i).Span Mod 5 = 0 Then
                            For x = 0 To 7
                                Call Check_Fire(x + 10, i)
                            Next x
                        End If
                    End If
                    If EShip(i).Span > 75 And EShip(i).Span < 100 Then
                        Call Check_Fire(1, i)
                    End If
                    If EShip(i).Span > 125 And EShip(i).Span < 150 Then
                        If EShip(i).Span Mod 5 = 0 Then
                            For x = 0 To 7
                                Call Check_Fire(x + 3, i)
                            Next x
                        End If
                    End If
                    If EShip(i).Span > 175 And EShip(i).Span < 200 Then
                        Call Check_Fire(17, i)
                    End If
                    If EShip(i).Span = 200 Then EShip(i).Span = 0
                End If
                
                If EShip(i).Patt = 7979 Then    'Mother Scyth
                    If EShip(i).PosY < 1500 Then EShip(i).PosY = EShip(i).PosY + 40
                    If EShip(i).Span Mod 240 > 110 And EShip(i).Span Mod 240 < 120 Then EShip(i).PosX = EShip(i).PosX - EShip(i).Spd
                    If EShip(i).Span Mod 240 > 230 And EShip(i).Span Mod 240 < 240 Then EShip(i).PosX = EShip(i).PosX + EShip(i).Spd
                    If EShip(i).PosX < 900 Then EShip(i).PosX = 900
                    If EShip(i).PosX > 2800 Then EShip(i).PosX = 2800
                    If EShip(i).Span Mod 20 = 1 Then Call Check_Fire(1, i)
                    If EShip(i).Span Mod 30 = 1 Then Call Check_Fire(2, i)
                    If EShip(i).Span Mod 150 = 1 Then
                        For x = 0 To 10
                            Call Check_Fire(17, i)
                        Next x
                    End If
                    If EShip(i).Span Mod 70 > 60 And EShip(i).Span Mod 70 < 70 Then
                        Call Check_Fire(17, i)
                    End If
                End If
                If EShip(i).Patt = 17931 Then    'BLACK DISC
                    If EShip(i).PosY < 1500 Then EShip(i).PosY = EShip(i).PosY + 50
                    If EShip(i).Span Mod 500 > 25 And EShip(i).Span Mod 500 < 100 Then EShip(i).PosX = EShip(i).PosX - 80
                    If EShip(i).Span Mod 500 > 150 And EShip(i).Span Mod 500 < 250 Then EShip(i).PosY = EShip(i).PosY + 80
                    If EShip(i).Span Mod 500 > 300 And EShip(i).Span Mod 500 < 400 Then EShip(i).PosX = EShip(i).PosX + 80
                    If EShip(i).Span Mod 500 > 450 And EShip(i).Span Mod 500 < 550 Then EShip(i).PosY = EShip(i).PosY - 80
                    
                    If EShip(i).Span Mod 60 = 1 Then Call Check_Fire(2, i)
                    If EShip(i).Span Mod 25 = 1 Then Call Check_Fire(1, i)
                    If EShip(i).Span Mod 500 > 25 And EShip(i).Span Mod 500 < 60 Then
                        If EShip(i).Span Mod 40 = 0 Then Call Check_Fire(19, i)
                    End If
                    If EShip(i).Span Mod 500 = 60 Or EShip(i).Span Mod 500 = 90 Or EShip(i).Span Mod 500 = 120 Then
                        For x = 0 To 3
                            Call Check_Fire(17, i)
                        Next x
                    End If
                    If EShip(i).Span Mod 500 = 401 Or EShip(i).Span Mod 500 = 475 Or EShip(i).Span Mod 500 = 499 Then
                        For x = 0 To 3
                            Call Check_Fire(17, i)
                        Next x
                    End If
                    If EShip(i).Span Mod 500 > 300 And EShip(i).Span Mod 500 < 400 And EShip(i).Span Mod 3 = 1 Then Call Check_Fire(1, i)
                    If EShip(i).Span Mod 5 = 0 And EShip(i).Span Mod 500 > 150 And EShip(i).Span Mod 500 < 250 Then
                        If EShip(i).Span Mod 20 = 0 Then Call Check_Fire(16, i)
                        If EShip(i).Span Mod 25 = 0 Then Call Check_Fire(Rnd * 6 + 10, i)
                    End If
                    If EShip(i).Span Mod 200 = 199 Then Call Create_Enemy(21, 11, (Rnd * 3800 + 800), 0, 88, 30, 160, 180, 0, False, True)
                    If EShip(i).Span Mod 200 = 199 Then
                        For x = 1 To 20
                            Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                        Next x
                    End If
                    If EShip(i).Span Mod 5 = 0 And EShip(i).Span Mod 500 > 450 And EShip(i).Span Mod 500 < 550 Then
                        If EShip(i).Span Mod 20 = 0 Then Call Check_Fire(9, i)
                        If EShip(i).Span Mod 30 = 0 Then Call Check_Fire(Rnd * 6 + 3, i)
                    End If
                    If EShip(i).Span Mod 30 = 1 Then Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    
                    If EShip(i).PosX < 800 Then EShip(i).PosX = 800
                    If EShip(i).PosY > 5325 Then EShip(i).PosY = 5325
                    If EShip(i).PosX > 3705 Then EShip(i).PosX = 3705
                    If EShip(i).Span > 100 Then
                        If EShip(i).PosY < 1500 Then EShip(i).PosY = 1500
                    End If
                End If
                
                'Kill ships
                If EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) >= 8640 Or EShip(i).PosY < 0 Or EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) >= 6360 Or EShip(i).PosX <= 0 Then EShip(i).Used = False
            
            End If
        Next i
        
        For i = 1 To 40
            If EShip(i).Used = True Then
                Img(EShip(i).Kind).ListImages(EShip(i).Frame).Draw Play_Layer.hDC, EShip(i).PosX, EShip(i).PosY, 1
            End If
        Next i
        
        '-----------------------------------------
                                'MOVING E_BULLET
        'Again, similar to player's bullet, but with
        'less complexity
        For i = 0 To 200
            If EBull(i).Used = True Then
                If EBull(i).Dire = 1 Then
                    Dim Bull_Distance As Double
                    Dim Temp_Division As Integer
                    Bull_Distance = (Abs(EBull(i).DestX - EBull(i).StartX)) + (Abs(EBull(i).DestY - EBull(i).StartY))
                    Temp_Division = (Bull_Distance / 200)
                    EBull(i).PosX = EBull(i).PosX + ((EBull(i).DestX - EBull(i).StartX) / (EBull(i).Spd + Abs(Temp_Division)))
                    EBull(i).PosY = EBull(i).PosY + ((EBull(i).DestY - EBull(i).StartY) / (EBull(i).Spd + Abs(Temp_Division)))
                End If
                If EBull(i).Dire = 2 Then
                End If
            
            If EBull(i).PosX < 560 Or EBull(i).PosX > 5760 Or EBull(i).PosY < 1280 Or EBull(i).PosY > 8040 Then EBull(i).Used = False
            
            Select Case (EBull(i).Kind)
            Case 11:
                If (EBull(i).PosX - 75 < ShipX + (Img(0).ImageWidth * 15)) And (EBull(i).PosX + 75 > ShipX) And (EBull(i).PosY - 75 < ShipY + (Img(0).ImageHeight * 15)) And (EBull(i).PosY + 75 > ShipY) Then
                    EBull(i).Used = False
                    Call Minor_HIT
                End If
            Case 12:
                If (EBull(i).PosX - 75 < ShipX + (Img(0).ImageWidth * 15)) And (EBull(i).PosX + 75 > ShipX) And (EBull(i).PosY - 75 < ShipY + (Img(0).ImageHeight * 15)) And (EBull(i).PosY + 75 > ShipY) Then
                    EBull(i).Used = False
                    Call Minor_HIT
                End If
            Case 21:
                If (EBull(i).PosX - 150 < ShipX + (Img(0).ImageWidth * 15)) And (EBull(i).PosX + 150 > ShipX) And (EBull(i).PosY - 150 < ShipY + (Img(0).ImageHeight * 15)) And (EBull(i).PosY + 150 > ShipY) Then
                    EBull(i).Used = False
                    Call Minor_HIT
                End If
            Case 22:
                If (EBull(i).PosX - 150 < ShipX + (Img(0).ImageWidth * 15)) And (EBull(i).PosX + 150 > ShipX) And (EBull(i).PosY - 150 < ShipY + (Img(0).ImageHeight * 15)) And (EBull(i).PosY + 150 > ShipY) Then
                    EBull(i).Used = False
                    Call Minor_HIT
                End If
            End Select
            
            If SHP(2) = 1 Then          'For Matrix Globe
                For x = 290 To 300
                    If Bull(x).Used = True Then
                        'To check the collision of the matrix globe with the enemy's bullets
                        If EBull(i).Kind = 11 Or EBull(i).Kind = 12 Then
                            If (EBull(i).PosX - 75 < Bull(x).PosX + (Img(5).ImageWidth * 15)) And (EBull(i).PosX + 75 > Bull(x).PosX) And (EBull(i).PosY - 75 < Bull(x).PosY + (Img(5).ImageHeight * 15)) And (EBull(i).PosY + 75 > Bull(x).PosY) Then EBull(i).Used = False
                        Else:
                            If (EBull(i).PosX - 150 < Bull(x).PosX + (Img(5).ImageWidth * 15)) And (EBull(i).PosX + 150 > Bull(x).PosX) And (EBull(i).PosY - 150 < Bull(x).PosY + (Img(5).ImageHeight * 15)) And (EBull(i).PosY + 150 > Bull(x).PosY) Then EBull(i).Used = False
                        End If
                    End If
                Next x
            End If
            
            'similar to other sprites, enemy bullets are painted
            Select Case (EBull(i).Kind)
                Case 11:
                    Img(15).ListImages(1).Draw Play_Layer.hDC, EBull(i).PosX - 75, EBull(i).PosY - 75, 1
                Case 12:
                    Img(15).ListImages(2).Draw Play_Layer.hDC, EBull(i).PosX - 75, EBull(i).PosY - 75, 1
                Case 21:
                    Img(16).ListImages(1).Draw Play_Layer.hDC, EBull(i).PosX - 150, EBull(i).PosY - 150, 1
                Case 22:
                    Img(16).ListImages(2).Draw Play_Layer.hDC, EBull(i).PosX - 150, EBull(i).PosY - 150, 1
                End Select
            End If
        Next i
        
        '-----------------------------------------
                                'POWER-UP
        'If a powerup is in play, it'll be animated ...
        For i = 0 To 2
            If PowerUp(i).Used = True Then
                PowerUp(i).PosY = PowerUp(i).PosY + 30
                If PowerUp(i).PosY > 7920 Then PowerUp(i).Used = False
                PowerUp(i).Frame = PowerUp(i).Frame + 1
                If PowerUp(i).Frame > 15 Then PowerUp(i).Frame = 1
                Img(26).ListImages(PowerUp(i).Frame).Draw Play_Layer.hDC, PowerUp(i).PosX, PowerUp(i).PosY, 1
            End If
        Next i
        
        '... and moved downwards to the screen bottom.
        For i = 0 To 2
            If PowerUp(i).Used = True Then
                If PowerUp(i).PosX + (Img(26).ImageWidth * 15) > ShipX And PowerUp(i).PosX < ShipX + (Img(0).ImageWidth * 15) And PowerUp(i).PosY + (Img(26).ImageHeight * 15) > ShipY And PowerUp(i).PosY < ShipY + (Img(0).ImageHeight * 15) Then
                    Call Alter_G(GEC Mod 2, 1)
                    Call PLAY(7)
                    PowerUp(i).Used = False
                    Exit For
                End If
            End If
        Next i

        
        '-----------------------------------------
                                'SHIELD CIRCLE
        'If the player's ship is hit with an enemy bullet,
        'there will be a brief flicker on the player's shield
        If GEC > 0 Then
        If SHIELD_ON_OFF >= 1 Then
            SHIELD_ON_OFF = SHIELD_ON_OFF + 1
            Select Case (SHIELD_ON_OFF)
                Case 2:
                    Img(5).ListImages(1).Draw Play_Layer.hDC, ShipX - 100, ShipY - 150, 1
                Case 4:
                    Img(5).ListImages(2).Draw Play_Layer.hDC, ShipX - 100, ShipY - 150, 1
                Case 6:
                    Img(5).ListImages(3).Draw Play_Layer.hDC, ShipX - 100, ShipY - 150, 1
                Case 8:
                    Img(5).ListImages(4).Draw Play_Layer.hDC, ShipX - 100, ShipY - 150, 1
                    SHIELD_ON_OFF = 0
            End Select
        End If
        End If
        
        
        '-----------------------------------------
                                                
                                'MOVING SHIP
        If (Lock_Engine = False And Lock_All = False) Or (OVERIDE = True) Then
            If Ship_Moving(0) = True Then           'Move LEFT
                ShipX = ShipX - Ship_Speed
                If ShipX <= 720 Then
                    ShipX = 720
                End If
            End If
            If Ship_Moving(1) = True Then           'Move UP
                ShipY = ShipY - Ship_Speed
                If ShipY <= 1440 Then
                    ShipY = 1440
                End If
            End If
            If Ship_Moving(2) = True Then           'Move RIGHT
                ShipX = ShipX + Ship_Speed
                If ShipX >= 4920 Then
                    ShipX = 4920
                End If
            End If
            If Ship_Moving(3) = True Then           'Move DOWN
                ShipY = ShipY + Ship_Speed
                If ShipY >= 6860 Then
                    ShipY = 6860
                End If
            End If
        End If
                            'SHIP PAINTING
        Img(0).ListImages((FPS Mod 3) + 1).Draw Play_Layer.hDC, (ShipX), (ShipY), 1
        
        '-----------------------------------------
                                        'MOVE STATSBAR UP DOWN
        If Moving_StatsBar = 1 Then
            Stats_Bar.Top = Stats_Bar.Top + 48
            Bufferbox(5).Top = Bufferbox(5).Top - 48
            If Stats_Bar.Top >= 7920 Then
                Stats_Bar.Top = 7920
                Bufferbox(5).Top = 1080
                Moving_StatsBar = 0
                Stats_Bar.Visible = False
            End If
        End If
        
        If Moving_StatsBar = 2 Then
            Stats_Bar.Visible = True
            Stats_Bar.Top = Stats_Bar.Top - 48
            Bufferbox(5).Top = Bufferbox(5).Top + 48
            If Stats_Bar.Top <= 7440 Then
                Stats_Bar.Top = 7440
                Bufferbox(5).Top = 1560
                Moving_StatsBar = 0
            End If
        End If
        
        If Moving_BossBar = 1 Then     'MOVE BOSSBAR UP DOWN
            BossBar.Top = BossBar.Top + 48
            If BossBar.Top >= 1560 Then
                BossBar.Top = 1560
                Moving_BossBar = 0
            End If
        End If
        
        If Moving_BossBar = 2 Then
            BossBar.Top = BossBar.Top - 48
            If BossBar.Top <= 1080 Then
                BossBar.Top = 1080
                Moving_BossBar = 0
            End If
        End If
        
        If Moving_SpecialBar = 1 Then   'MOVE SPECIALBAR UP DOWN
            Bufferbox(6).Top = Bufferbox(6).Top - 48
            If Bufferbox(6).Top <= 1080 Then
                Bufferbox(6).Top = 1080
                Moving_SpecialBar = 0
            End If
        End If
        
        If Moving_SpecialBar = 2 Then
            Bufferbox(6).Top = Bufferbox(6).Top + 48
            If Bufferbox(6).Top >= 1560 Then
                Bufferbox(6).Top = 1560
                Moving_SpecialBar = 0
                Special_Delay = 1
            End If
        End If
        
        If Special_Delay >= 1 Then Special_Delay = Special_Delay + 1
        If Special_Delay = 35 Then Moving_SpecialBar = 1
                
                                
                '-----------------------------------------
                                        'COLLISION DETECTION Bull > EShip
        'These codes perform collision detection between
        'the player's bullets with an enemy's ship.
        'Since the bullet and ships above are painted,
        'It will be calculated to see if any portion of
        'them is overlapping each other, thus causing a hit.
        For i = 0 To 300
            If Bull(i).Used = True Then
                For x = 1 To 40
                    If EShip(x).Used = True Then
                        Dim Temp_Img As Integer
                        Dim Img_SWidth As Integer
                        Dim Img_SHeight As Integer
                        Dim Bul_SWidth As Integer
                        Dim Bul_SHeight As Integer
                        Dim Img_EWidth As Integer
                        Dim Img_EHeight As Integer
                        Dim Bul_EWidth As Integer
                        Dim Bul_EHeight As Integer
                        
                            If SHP(0) = 0 And i < 200 Then Temp_Img = 4
                            If SHP(0) = 1 And i < 200 Then Temp_Img = 6
                            If SHP(0) = 2 And i < 200 Then Temp_Img = 7
                            If SHP(0) = 3 And i < 200 Then Temp_Img = 8
                            If SHP(0) = 4 And i < 200 Then Temp_Img = 14
                            If SHP(1) = 0 And i >= 200 And i < 290 Then Temp_Img = 9
                            If SHP(1) = 1 And i >= 200 And i < 290 Then Temp_Img = 10
                            If SHP(1) = 2 And i >= 200 And i < 290 Then Temp_Img = 11
                            If SHP(1) = 3 And i >= 200 And i < 290 Then Temp_Img = 12
                            If SHP(2) = 0 And i >= 290 Then Temp_Img = 17
                            If SHP(2) = 1 And i >= 290 Then Temp_Img = 5
                            If SHP(2) = 3 And i >= 290 Then Temp_Img = 3
                            
                            Bul_SHeight = Bull(i).PosY
                            Bul_SWidth = Bull(i).PosX
                            Img_SHeight = EShip(x).PosY
                            Img_SWidth = EShip(x).PosX
                            
                            Bul_EHeight = Bull(i).PosY + Img(Temp_Img).ImageHeight * 15
                            Bul_EWidth = Bull(i).PosX + Img(Temp_Img).ImageWidth * 15
                            Img_EHeight = EShip(x).PosY + Img(EShip(x).Kind).ImageHeight * 15
                            Img_EWidth = EShip(x).PosX + Img(EShip(x).Kind).ImageWidth * 15
                   
                            'Correctly adjust the detection width of weapon if using light beam
                            If SHP(0) = 1 And i < 200 Then      'All Level Light Beam
                                Bul_SWidth = Bul_SWidth + (10 - MG_Lvl) * 10
                                Bul_EWidth = Bul_EWidth - (10 - MG_Lvl) * 10
                            End If
                            
                            'The following 10 portions are to correctly detection the
                            'plasma wave weapon based on its varied hegiht and width
                            If SHP(0) = 3 And i < 200 And MG_Lvl >= 1 Then     'Wave 1
                                Bul_SWidth = Bul_SWidth + 48
                                Bul_EWidth = Bul_EWidth - 48
                                Bul_SHeight = Bul_SHeight + 17
                                Bul_EHeight = Bul_EHeight - 35
                            End If
                            If SHP(0) = 3 And i < 200 And MG_Lvl >= 2 Then     'Wave 2
                                Bul_SWidth = Bul_SWidth + 43
                                Bul_EWidth = Bul_EWidth - 43
                                Bul_SHeight = Bul_SHeight + 14
                                Bul_EHeight = Bul_EHeight - 35
                            End If
                            If SHP(0) = 3 And i < 200 And MG_Lvl >= 3 Then     'Wave 3
                                Bul_SWidth = Bul_SWidth + 32
                                Bul_EWidth = Bul_EWidth - 32
                                Bul_SHeight = Bul_SHeight + 14
                                Bul_EHeight = Bul_EHeight - 35
                            End If
                            If SHP(0) = 3 And i < 200 And MG_Lvl >= 4 Then     'Wave 4
                                Bul_SWidth = Bul_SWidth + 26
                                Bul_EWidth = Bul_EWidth - 26
                                Bul_SHeight = Bul_SHeight + 14
                                Bul_EHeight = Bul_EHeight - 35
                            End If
                            If SHP(0) = 3 And i < 200 And MG_Lvl >= 5 Then     'Wave 5
                                Bul_SWidth = Bul_SWidth + 26
                                Bul_EWidth = Bul_EWidth - 26
                                Bul_SHeight = Bul_SHeight + 14
                                Bul_EHeight = Bul_EHeight - 21
                            End If
                            If SHP(0) = 3 And i < 200 And MG_Lvl >= 6 Then     'Wave 6
                                Bul_SWidth = Bul_SWidth + 21
                                Bul_EWidth = Bul_EWidth - 21
                                Bul_SHeight = Bul_SHeight + 14
                                Bul_EHeight = Bul_EHeight - 21
                            End If
                            If SHP(0) = 3 And i < 200 And MG_Lvl >= 7 Then     'Wave 7
                                Bul_SWidth = Bul_SWidth + 14
                                Bul_EWidth = Bul_EWidth - 14
                                Bul_SHeight = Bul_SHeight + 9
                                Bul_EHeight = Bul_EHeight - 21
                            End If
                            If SHP(0) = 3 And i < 200 And MG_Lvl >= 8 Then     'Wave 8
                                Bul_SWidth = Bul_SWidth + 10
                                Bul_EWidth = Bul_EWidth - 10
                                Bul_SHeight = Bul_SHeight + 9
                                Bul_EHeight = Bul_EHeight - 16
                            End If
                            If SHP(0) = 3 And i < 200 And MG_Lvl >= 9 Then     'Wave 9
                                Bul_SWidth = Bul_SWidth + 7
                                Bul_EWidth = Bul_EWidth - 7
                                Bul_SHeight = Bul_SHeight + 6
                                Bul_EHeight = Bul_EHeight - 10
                            End If
                        
                        'If bullet overlaps enemy's ship ...
                        If Bul_SWidth < Img_EWidth And Bul_EWidth > Img_SWidth And Bul_SHeight < Img_EHeight And Bul_EHeight > Img_SHeight Then
                            'Reduce their hit points by bullet's damage
                            EShip(x).Life = EShip(x).Life - Bull(i).Dama
                            
                            'For bosses ...
                            If EShip(x).Boss = True Then
                                ' ... they cannot have life less than zero because ...
                                If EShip(x).Life < 0 Then EShip(x).Life = 0
                                ' ... width of their lifebar cannot be < 0.
                                Boss_Bar_Top.Width = EShip(x).Life / EShip(x).Max_Life * Boss_Bar_Bottom.Width
                            End If
                            
                            'Generate a small ricochet explosion on the enemy aircraft
                            Call EXPLODE(EShip(x).PosX + (Rnd * Img(EShip(x).Kind).ImageWidth * 15), EShip(x).PosY + (Rnd * Img(EShip(x).Kind).ImageHeight * 15), 13)
                            
                            'Check for their life points if it is below 0
                            Call Check_EDeath(x)
                            
                            'For plasma wave and matrix globe, the bullet remains after hitting enemy
                            If (SHP(0) = 3 And i < 200) Or (SHP(2) = 0 And i >= 290) Or (SHP(2) = 1 And i >= 290) Or (SHP(2) = 3 And i >= 290) Then
                                Bull(i).Used = True
                            Else:
                                'All other bullets will be removed (unused)
                                Bull(i).Used = False
                            End If
                        End If
                    End If
                Next x
            End If
        Next i
                
                '-----------------------------------------
                                        'COLLISION DETECTION SHIP VS SHIP
        'Similar to bullets with eship, this time, detection is
        'between the player's ship and the enemy's ship
        For i = 0 To 40
            If EShip(i).Used = True Then
                'Collision detection
                If EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2 > ShipX And EShip(i).PosX < ShipX + (Img(0).ImageWidth * 15) / 2 And EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2 > ShipY And EShip(i).PosY < ShipY + (Img(0).ImageHeight * 15) / 2 Then
                    'If collide, enemy deduct 150 life points
                    EShip(i).Life = EShip(i).Life - 150
                    'Player loses 1 shield
                    Call Minor_HIT
                    'Checks for enemy's lifepoints
                    Call Check_EDeath(i)
                End If
            End If
        Next i
        
                '-----------------------------------------
                                        'ANIMATE EXPLOSION
        'This sub paints the explosion picture, if any of it is in play.
        'It will keep track of their frames and slowly
        'increment it one by one. after all of its frame
        ' has been displayed, the explosion object will be
        ' removed (unused)
        For i = 0 To 40
            If Explo(i).Used = True Then
                If Explo(i).PosX < 0 Then Explo(i).PosX = 0
                If Explo(i).PosY < 0 Then Explo(i).PosY = 0
                Img(Explo(i).Kind).ListImages(Explo(i).Current_Frame).Draw Play_Layer.hDC, Explo(i).PosX, Explo(i).PosY, 1
                Explo(i).Current_Frame = Explo(i).Current_Frame + 1
                If Explo(i).Current_Frame >= Explo(i).Last_Frame Then
                    Explo(i).Used = False
                End If
            End If
        Next i
        
                '-----------------------------------------
                                        'SHAKEIT EFFECT
        'When a player gets a major hit, the screen
        'will shake and flash white. These codes does that.
        
        If ShakeIT = 3 Then
            Play_Layer.BackColor = &H0&     'Changes backcolor to BLACK
            Stats_Bar.Top = 7440
            ShakeIT = 0
        End If
        
        If ShakeIT >= 1 And ShakeIT <= 2 Then
            Play_Layer.BackColor = &HFFFFFF 'Change to WHITE
            Stats_Bar.Top = Stats_Bar.Top - 480     'SHAKE status bar
            ShakeIT = ShakeIT + 1
        End If
        
        If Giga_Count >= 1 Then             'This is for giga storm only ...
            Giga_Count = Giga_Count + 1
            'Shows the lightning image
            Img(27).ListImages(1).Draw Play_Layer.hDC, 1000, 2000, 1
            Img(27).ListImages(1).Draw Play_Layer.hDC, 3000, 2000, 1
            If Giga_Count > 5 Then Giga_Count = 0
        End If

                '-----------------------------------------
                                        'REPAIRS
        'If any component of the ship is damaged,
        'repairs will be perform.
        If GEC > 0 Then
        'Repair main gun if damaged
            If Lock_MainGun = True Then
                'Slowly increase the repaired bar
                Stt_MGUN_R.Width = Stt_MGUN_R.Width + Repair_Speed
                'If repairs completed
                If Stt_MGUN_R.Width >= 615 Then
                    Stt_MGUN_R.Width = 615
                    Lock_MainGun = False
                    Call Send_MSG("MainGun Repaired", 3, 0)
                End If
            End If
            
            'Repair side gun if damaged
            If Lock_SideGun = True Then
                Stt_SGUN_R.Width = Stt_SGUN_R.Width + Repair_Speed
                If Stt_SGUN_R.Width >= 615 Then
                    Stt_SGUN_R.Width = 615
                    Lock_SideGun = False
                    Call Send_MSG("SideGun Repaired", 3, 0)
                End If
            End If
            
            'Repair engine if damaged
            If Lock_Engine = True Then
                Stt_ENG_R.Width = Stt_ENG_R.Width + Repair_Speed
                If Stt_ENG_R.Width >= 615 Then
                    Stt_ENG_R.Width = 615
                    Lock_Engine = False
                    Call Send_MSG("Engines Repaired", 3, 0)
                End If
            End If
            
            'Repair generator if damaged
            If Lock_Generator = True Then
                Stt_PWR_R.Width = Stt_PWR_R.Width + Repair_Speed
                If Stt_PWR_R.Width >= 615 Then
                    Stt_PWR_R.Width = 615
                    Lock_Generator = False
                    Call Send_MSG("Generator Repaired", 3, 0)
                End If
            End If
        End If
        
                '-----------------------------------------
                                        'BUFFER LINES
        'This portion moves the on-screen msg sent to the
        'player and slowly scroll it downwards
        Buffering = False
        'If there's any text on screen, text scrolling is performed
        For i = 0 To 4
            If BufferText(i).Caption <> "" Then
                Buffering = True
            End If
        Next i
        
        'If the delay count of the lines moving matches
        'the delay speed set, the text is moved.
        'Note that there is a slight pause after the
        'lines have been moved, line after line...
        If Buffering = True Then
            Buffer_Count = Buffer_Count + 1
            If Buffer_Count = Buff_Speed Then
                Buffer_Count = 0
                Buffer_Mover = True
            End If
            
            'If moving is true (not during the slight pause)
            If Buffer_Mover = True Then
                If Buffer_Count Mod 2 = 0 Then
                    'Always remove the last line of text
                    Bufferbox(4).Width = 15
                    For i = 1 To 3
                        'From top line to bottom line, move the text downwards
                        Bufferbox(i).Top = Bufferbox(i).Top + 72
                        'If the last line reaches the bottom line...
                        If Bufferbox(3).Top >= 3960 Then
                            'All the text will inherit its attributes
                            'to their previous line, thus original
                            'attributes of the last line is lost.
                            'Attributes: Forecolor, Backcolor, Text
                            For x = 4 To 1 Step -1
                                BufferText(x).Caption = BufferText(x - 1).Caption
                                BufferText(x).ForeColor = BufferText(x - 1).ForeColor
                                Bufferbox(x).Width = Bufferbox(x - 1).Width
                                BufferText(x).BackColor = BufferText(x - 1).BackColor
                            Next x
                            'All the text line will take over the previous
                            'line's position, thus making it seem like there's
                            'continous shift of line, no sudden jump.
                            For x = 1 To 3
                                Bufferbox(x).Top = Bufferbox(x).Top - 360
                            Next x
                            Buffer_Mover = False
                            Bufferbox(0).Width = 15
                            BufferText(0).Caption = ""
                            BufferText(0).BackColor = &H0&
                        End If
                    Next i
                End If
            End If
        End If

                '-----------------------------------------
                
        'These codes Charges the shield bar and special bar
        'of the player's ship.
        Const n = 2 'Bonus
        Select Case (SHP(4))                'CHARGE SPEED (GENERATOR)
        Case 0: 'DUAL CORE
            If SHD_Charging = True And SPE_Charging = True Then
                SHD_Charge_Speed = SHD_Charge_Base + n
                SPE_Charge_Speed = n
            End If
            If SHD_Charging = True And SPE_Charging = False Then
                SHD_Charge_Speed = SHD_Charge_Base + n * 2
            End If
            If SHD_Charging = False And SPE_Charging = True Then
                SPE_Charge_Speed = n * 2
            End If
        Case 1: 'FUSION
            If SHD_Charging = True And SPE_Charging = True Then
                SHD_Charge_Speed = 0
                SPE_Charge_Speed = n * 3
            End If
            If SHD_Charging = True And SPE_Charging = False Then
                SHD_Charge_Speed = SHD_Charge_Base + n
            End If
            If SHD_Charging = False And SPE_Charging = True Then
                SPE_Charge_Speed = n * 8
            End If
        Case 2: 'MATRIX
            If SHD_Charging = True And SPE_Charging = True Then
                SHD_Charge_Speed = (SHD_Charge_Base + 4) * 3
                SPE_Charge_Speed = 0
            End If
            If SHD_Charging = True And SPE_Charging = False Then
                SHD_Charge_Speed = (SHD_Charge_Base + n) * 3
            End If
            If SHD_Charging = False And SPE_Charging = True Then
                SPE_Charge_Speed = n
            End If
            
        End Select
                               
                '-----------------------------------------
                                        'CHARGE SHIELD
        'If the player's shield is not full,
        'its bar will charge
        If SHD_Charging = True And Lock_Generator = False Then
            BAR_SHD_R.Width = BAR_SHD_R.Width + SHD_Charge_Speed
            If BAR_SHD_R.Width >= SHD_Width Then
                SHD_Charging = False
                SHD_Avail = True
                BAR_SHD_C.Width = SHD_Width
                BAR_SHD_R.Width = SHD_Width
                Stt_SHD_Lbl.BackColor = &H0&
                Stt_SHD_Lbl.ForeColor = &HFF00&
            End If
        End If
                '-----------------------------------------
                                        'CHARGE SPECIAL
        'If the player's special bar is not full,
        'it will charge
        If SPE_Charging = True And Lock_Generator = False Then
            BAR_SPE_R.Width = BAR_SPE_R.Width + SPE_Charge_Speed
            If BAR_SPE_R.Width >= SPE_Width Then
                SPE_Charging = False
                SPE_Avail = True
                BAR_SPE_C.Width = SPE_Width
                BAR_SPE_R.Width = SPE_Width
                Stt_SPE_Lbl.BackColor = &H0&
                Stt_SPE_Lbl.ForeColor = &HFF00&
            End If
        End If
        
        
                '-----------------------------------------
    End If
    
    Call GEC_Coding     'Create enemies at specified time (see below)
    
    For i = 200 To 290  'Remove Pulse Ring!
        If SHP(1) = 1 And i >= 200 And i < 290 Then Bull(i).Used = False
    Next i

End Sub

Private Sub GEC_Coding()
Dim i As Integer
                '-----------------------------------------
                '               GEC CODING
                '               GEC CODING
                '               GEC CODING
                '-----------------------------------------

    'Ships  :   18  Big Ufo
    '           19  Ferry
    '           20  Grey Fat Ship
    '           21  Bubble Yellow
    '           22  Army Fighter
    '           23  Meteor Boss
    '           24  Big Scyth
    '           25  Small Scyth
    '           28  Small Meteors
    '           29  Medium Meteors
    '           30  Big Meteors
    '           31  Mother Bubble

'-FOR TESTING ONLY-
'Uncomment lines below until the ----- only for
'testing purposes.
'The first line skips the intro and let the player jumo directly into play.
'The following line each will allow the player to
'skip a single stage.

'-Start uncomment below this line-
'If GEC = 1 Then GEC = 150
'If GEC = 151 Then
'    GEC = (2700 + 60)     'Modify the 2nd number in the bracket ONLY!
'    C_GEC(0) = 1250
'    C_GEC(1) = 2500
'    C_GEC(2) = 2600
'    C_GEC(3) = 2700
'End If
'------------

    Select Case (GEC)
        Case 1:
            Lock_All = True
            ShipX = 2800
            ShipY = 8000
        Case 20:
            Call Send_MSG(" ALERT ! APPROACHING ENEMY BASE ! ", 0, 1)
            OVERIDE = True
            Ship_Moving(1) = True
            Call PLAY(5)
        Case 45:
            Ship_Moving(1) = False
        Case 90:
            Call Send_MSG("> Command Base: Good Luck, Black Winter! ", 0, 0)
        Case 110:
            Call PLAY(6)
        Case 150:
            Stats_Bar.Visible = True
            Moving_StatsBar = 2
            Lock_All = False
            OVERIDE = False
            Call PLAY(4)
        Case 170:
        Call Send_MSG(" Stage 1: The Breach ", 0, 0)
            Call Create_Enemy(25, 11, 1760, 0, 88, 50, 40, 50, 1, False, False)
        Case 180:
            Call Create_Enemy(25, 11, 2760, 0, 88, 50, 40, 50, 1, False, False)
        Case 190:
            Call Create_Enemy(25, 11, 3760, 0, 88, 50, 40, 50, 1, False, False)
        Case 270:
            Call Create_Enemy(25, 11, 3760, 0, 88, 50, 40, 50, 1, False, False)
        Case 280:
            Call Create_Enemy(25, 11, 2760, 0, 88, 50, 40, 50, 1, False, False)
        Case 290:
            Call Create_Enemy(25, 11, 1760, 0, 88, 50, 40, 50, 1, False, False)
        Case 370:
            Call Create_Enemy(25, 11, 2260, 0, 88, 50, 40, 50, 1, False, False)
            Call Create_Enemy(25, 11, 4260, 0, 88, 50, 40, 50, 1, False, False)
        Case 390:
            Call Create_Enemy(25, 11, 1260, 0, 88, 50, 40, 50, 1, False, False)
            Call Create_Enemy(25, 11, 3260, 0, 88, 50, 40, 50, 1, False, False)
        Case 500:
            Call Create_Enemy(21, 11, 4660, 0, 88, 25, 100, 150, 0, False, True)
            Call Create_Enemy(25, 11, 1360, 0, 86, 40, 50, 50, 1, False, False)
        Case 530:
            Call Create_Enemy(25, 11, 1560, 0, 86, 40, 50, 50, 1, False, False)
        Case 560:
            Call Create_Enemy(25, 11, 1760, 0, 86, 40, 50, 50, 1, False, False)
        Case 600:
            Call Create_Enemy(21, 11, 1660, 0, 88, 25, 100, 150, 0, False, True)
        Case 610:
            Call Create_Enemy(22, 11, 2760, 0, 88, 35, 100, 100, 1, False, False)
            Call Create_Enemy(22, 11, 3760, 0, 88, 35, 100, 100, 1, False, False)
        Case 700:
            Call Create_Enemy(25, 11, 4960, 0, 88, 70, 60, 50, 2, False, False)
        Case 730:
            Call Create_Enemy(25, 11, 4960, 0, 88, 70, 60, 50, 3, False, False)
        Case 760:
            Call Create_Enemy(25, 11, 4960, 0, 88, 70, 60, 50, 4, False, False)
        Case 860:
            Call Create_Enemy(25, 11, 2060, 0, 88, 30, 40, 50, 1, False, False)
            Call Create_Enemy(25, 11, 4060, 0, 88, 30, 40, 50, 1, False, False)
        Case 880:
            Call Create_Enemy(25, 11, 1060, 0, 88, 30, 40, 50, 1, False, False)
            Call Create_Enemy(25, 11, 5060, 0, 88, 30, 40, 50, 1, False, False)
        Case 900:
            Call Create_Enemy(21, 11, 3060, 0, 88, 35, 100, 150, 0, False, True)
        Case 980:
            Call Create_Enemy(22, 11, 1760, 0, 88, 35, 100, 100, 6, False, False)
        Case 1010:
            Call Create_Enemy(22, 11, 4760, 0, 88, 35, 100, 100, 6, False, False)
        Case 1040:
            Call Create_Enemy(22, 11, 4760, 0, 88, 35, 100, 100, 6, False, False)
        Case 1100:
            Call Create_Enemy(22, 11, 5060, 0, 81, 35, 100, 100, 1, False, False)
        Case 1130:
            Call Create_Enemy(25, 11, 1160, 0, 86, 40, 40, 50, 1, False, False)
            Call Create_Enemy(25, 11, 4860, 0, 84, 40, 40, 50, 1, False, False)
        Case 1160:
            Call Create_Enemy(25, 11, 1160, 0, 86, 40, 40, 50, 1, False, False)
            Call Create_Enemy(25, 11, 4860, 0, 84, 40, 40, 50, 1, False, False)
        Case 1130:
            Call Create_Enemy(25, 11, 1160, 0, 86, 40, 40, 50, 1, False, False)
            Call Create_Enemy(25, 11, 4860, 0, 84, 40, 40, 50, 1, False, False)
        Case 1150:
            Call Create_Enemy(21, 11, 4860, 0, 88, 35, 100, 150, 0, False, True)
        Case 1170:
            Call Send_MSG(" INCOMING ENEMY ", 0, 1)
        Case 1250:
            Boss_Name.Caption = "TRANSPORT"
            Call Create_Enemy(19, 11, 2400, 0, 4646, 40, 1500, 400, 123, True, False)
        Case Else:
            If C_GEC(0) > 1 And C_GEC(1) = 0 And C_GEC(2) = 0 And C_GEC(3) = 0 And C_GEC(4) = 0 Then
                Select Case (GEC - C_GEC(0))
                Case 50:
                Call Send_MSG(" Stage 2: Space Debris ", 0, 0)
                    Call Create_Enemy(25, 11, 1200, 0, 821, 100, 60, 75, 2, False, False)
                    Call Create_Enemy(25, 11, 1700, 0, 821, 100, 60, 75, 2, False, False)
                    Call Create_Enemy(25, 11, 2200, 0, 821, 100, 60, 75, 2, False, False)
                Case 80:
                    Call Create_Enemy(25, 11, 2200, 0, 821, 100, 60, 75, 2, False, False)
                    Call Create_Enemy(25, 11, 2700, 0, 821, 100, 60, 75, 2, False, False)
                    Call Create_Enemy(25, 11, 3200, 0, 821, 100, 60, 75, 2, False, False)
                Case 110:
                    Call Create_Enemy(25, 11, 3200, 0, 821, 100, 60, 75, 2, False, False)
                    Call Create_Enemy(25, 11, 3700, 0, 821, 100, 60, 75, 2, False, False)
                    Call Create_Enemy(25, 11, 4200, 0, 821, 100, 60, 75, 2, False, False)
                Case 120:
                    Call Create_Enemy(21, 11, 4700, 0, 88, 35, 120, 150, 0, False, True)
                Case 140:
                    Call Create_Enemy(22, 11, 1860, 0, 88, 45, 120, 120, 1, False, False)
                Case 150:
                    Call Create_Enemy(22, 11, 2660, 0, 88, 45, 120, 120, 1, False, False)
                Case 180:
                    Call Create_Enemy(25, 11, 4200, 0, 81, 45, 60, 75, 1, False, False)
                Case 190:
                    Call Create_Enemy(25, 11, 4200, 0, 81, 45, 60, 75, 1, False, False)
                Case 200:
                    Call Create_Enemy(25, 11, 4200, 0, 81, 45, 60, 75, 1, False, False)
                Case 230:
                    'Scyth
                    Call Create_Enemy(25, 11, 900, 0, 821, 100, 60, 75, 1, False, False)
                    Call Create_Enemy(25, 11, 1500, 0, 821, 100, 60, 75, 1, False, False)
                Case 270:
                    'Yellow Bubble
                    Call Create_Enemy(21, 11, 1500, 0, 88, 30, 120, 150, 0, False, True)
                Case 280:
                    'Fat Grey Ship
                    Call Create_Enemy(20, 12, 4000, 0, 88, 30, 230, 150, 32147, False, False)
                Case 300:
                    Call Create_Enemy(20, 12, 4600, 0, 88, 30, 230, 150, 32147, False, False)
                Case 370:
                    Call Send_MSG(" ALERT! ASTEROID BELT AHEAD! ", 0, 1)
                Case 420:
                    For i = 1 To 5
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 20, Rnd * 10 + 10, 10, 0, False, False)
                    Next i
                Case 470:
                    For i = 1 To 5
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 20, Rnd * 10 + 10, 10, 0, False, False)
                    Next i
                Case 520:
                    For i = 1 To 5
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 20, Rnd * 10 + 10, 10, 0, False, False)
                    Next i
                Case 570:
                    For i = 1 To 5
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 30, Rnd * 20 + 10, 10, 0, False, False)
                    Next i
                Case 620:
                    For i = 1 To 5
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 30, Rnd * 20 + 10, 10, 0, False, False)
                    Next i
                Case 670:
                    For i = 1 To 5
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 30, Rnd * 20 + 10, 10, 0, False, False)
                    Next i
                Case 720:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 40, Rnd * 20 + 20, 15, 0, False, False)
                    Next i
                Case 770:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 40, Rnd * 20 + 20, 15, 0, False, False)
                    Next i
                Case 780:
                    Call Send_MSG("> Command Base: It's going to get rough! ", 0, 0)
                Case 820:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 50, Rnd * 20 + 30, 15, 0, False, False)
                    Next i
                Case 860:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 50, Rnd * 20 + 30, 15, 0, False, False)
                    Next i
                Case 900:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 50, Rnd * 20 + 30, 15, 0, False, False)
                    Next i
                Case 940:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 980:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1020:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1150:
                    Call Send_MSG("> Command Base: HailStorm!! ", 0, 0)
                Case 1170:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1190:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1210:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1220:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1230:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1240:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1250:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1260:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1270:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1280:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1290:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1300:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1310:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1320:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1330:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1340:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1350:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1360:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1370:
                    For i = 1 To 10
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 40 + 80, Rnd * 20 + 60, 30, 0, False, False)
                    Next i
                Case 1420:
                    Boss_Name.Caption = "GRANITE"
                    Call Create_Enemy(23, 11, 3050, 0, 3311, 80, 4500, 1700, 999, True, False)
                End Select
            End If
            
            If C_GEC(1) > 1 And C_GEC(2) = 0 And C_GEC(3) = 0 And C_GEC(4) = 0 Then
                Select Case (GEC - C_GEC(1))
                Case 60:
                    Call Send_MSG(" Stage 3: The Big Gift ", 0, 0)
                    'Yellow Bubble
                    Call Create_Enemy(21, 11, 2400, 0, 88, 30, 160, 180, 0, False, True)
                    Call Create_Enemy(21, 11, 3500, 0, 88, 30, 160, 180, 0, False, True)
                Case 80:
                    'Army Type
                    Call Create_Enemy(22, 11, 1000, 0, 88, 55, 200, 160, 16, False, False)
                Case 100:
                    Call Create_Enemy(22, 11, 5000, 0, 88, 55, 200, 160, 9, False, False)
                Case 120:
                    Call Create_Enemy(22, 11, 1000, 0, 88, 55, 200, 160, 16, False, False)
                Case 140:
                    Call Create_Enemy(22, 11, 5000, 0, 88, 55, 200, 160, 9, False, False)
                Case 160:
                    Call Create_Enemy(22, 11, 1000, 0, 88, 55, 200, 160, 16, False, False)
                Case 180:
                    Call Create_Enemy(22, 11, 5000, 0, 88, 55, 200, 160, 9, False, False)
                Case 210:
                    'Fat Grey Ship
                    Call Create_Enemy(20, 12, 4800, 0, 88, 30, 250, 170, 32147, False, False)
                    Call Create_Enemy(20, 12, 1100, 0, 88, 30, 250, 170, 12369, False, False)
                Case 240:
                    Call Create_Enemy(25, 11, 900, 0, 86, 70, 150, 100, 1, False, False)
                Case 250:
                    Call Create_Enemy(25, 11, 900, 0, 86, 70, 150, 100, 1, False, False)
                Case 260:
                    Call Create_Enemy(25, 11, 900, 0, 86, 70, 150, 100, 1, False, False)
                Case 270:
                    Call Create_Enemy(25, 11, 5000, 0, 84, 70, 150, 100, 1, False, False)
                Case 280:
                    Call Create_Enemy(25, 11, 5000, 0, 84, 70, 150, 100, 1, False, False)
                Case 290:
                    Call Create_Enemy(25, 11, 5000, 0, 84, 70, 150, 100, 1, False, False)
                Case 340:
                    Call Create_Enemy(21, 11, 900, 0, 88, 30, 160, 180, 0, False, True)
                    Call Create_Enemy(21, 11, 4900, 0, 88, 30, 160, 180, 0, False, True)
                Case 360:
                    Call Create_Enemy(22, 11, 1800, 0, 88, 50, 250, 190, (Rnd * 6 + 3), False, False)
                    Call Create_Enemy(22, 11, 3000, 0, 88, 50, 250, 190, (Rnd * 6 + 10), False, False)
                Case 390:
                    Call Create_Enemy(22, 11, 1800, 0, 88, 50, 250, 190, (Rnd * 6 + 3), False, False)
                    Call Create_Enemy(22, 11, 3000, 0, 88, 50, 250, 190, (Rnd * 6 + 10), False, False)
                Case 420:
                    Call Create_Enemy(22, 11, 1800, 0, 88, 50, 250, 190, (Rnd * 6 + 3), False, False)
                    Call Create_Enemy(22, 11, 3000, 0, 88, 50, 250, 190, (Rnd * 6 + 10), False, False)
                Case 450:
                    Call Create_Enemy(22, 11, 1800, 0, 88, 50, 250, 190, (Rnd * 6 + 3), False, False)
                    Call Create_Enemy(22, 11, 3000, 0, 88, 50, 250, 190, (Rnd * 6 + 10), False, False)
                Case 480:
                    Call Create_Enemy(22, 11, 1800, 0, 88, 50, 250, 190, (Rnd * 6 + 3), False, False)
                    Call Create_Enemy(22, 11, 3000, 0, 88, 50, 250, 190, (Rnd * 6 + 10), False, False)
                Case 510:
                    Call Create_Enemy(22, 11, 1800, 0, 88, 50, 250, 190, (Rnd * 6 + 3), False, False)
                    Call Create_Enemy(22, 11, 3000, 0, 88, 50, 250, 190, (Rnd * 6 + 10), False, False)
                Case 540:
                    Call Create_Enemy(22, 11, 1800, 0, 88, 50, 250, 190, (Rnd * 6 + 3), False, False)
                    Call Create_Enemy(22, 11, 3000, 0, 88, 50, 250, 190, (Rnd * 6 + 10), False, False)
                Case 610:
                    Call Create_Enemy(21, 11, 4500, 0, 88, 50, 160, 180, 0, False, True)
                Case 630:
                    Call Create_Enemy(21, 11, 4500, 0, 88, 50, 160, 180, 0, False, True)
                Case 660:
                    Call Create_Enemy(22, 11, 800, 0, 821, 9, 200, 160, 2, False, False)
                Case 680:
                    Call Create_Enemy(22, 11, 1600, 0, 821, 9, 200, 160, 2, False, False)
                Case 700:
                    Call Create_Enemy(22, 11, 2400, 0, 821, 9, 200, 160, 2, False, False)
                Case 720:
                    Call Create_Enemy(22, 11, 3200, 0, 821, 9, 200, 160, 2, False, False)
                Case 740:
                    Call Create_Enemy(22, 11, 4000, 0, 821, 9, 200, 160, 2, False, False)
                Case 760:
                    Call Create_Enemy(22, 11, 4800, 0, 821, 9, 200, 160, 2, False, False)
                Case 780:
                    Call Send_MSG("> Command Base: We've got Weird readings ahead ... ", 0, 0)
                Case 800:
                    Call Create_Enemy(20, 12, 1100, 0, 88, 30, 250, 170, 12369, False, False)
                Case 860:
                    Call Send_MSG(" INCOMING ENEMY ", 0, 1)
                Case 900:
                    Call Create_Enemy(21, 11, 9000, 0, 88, 30, 160, 180, 0, False, True)
                    Call Create_Enemy(21, 11, 4700, 0, 88, 30, 160, 180, 0, False, True)
                Case 950:   'Mother bubble
                    Boss_Name.Caption = "GIANT POD"
                    Call Create_Enemy(31, 21, 2550, 0, 2288, 100, 7500, 3000, 12321, True, True)
                End Select
            End If
            
            If C_GEC(2) > 1 And C_GEC(3) = 0 And C_GEC(4) = 0 Then
                Select Case (GEC - C_GEC(2))
                Case 20:
                    Call Send_MSG(" Stage 4: Barricade of Doom ", 0, 0)
                    Call Create_Enemy(21, 11, 900, 0, 88, 30, 160, 180, 0, False, True)
                    Call Create_Enemy(21, 11, 1800, 0, 88, 30, 160, 180, 0, False, True)
                Case 100:
                    'Scyth Squadron
                    Call Send_MSG("> Command base: Scyth Squadron coming ahead ! ", 0, 0)
                    Call Create_Enemy(25, 11, 900, 0, 88, 90, 90, 60, 1, False, False)
                Case 105:
                    Call Create_Enemy(25, 11, 1400, 0, 88, 90, 90, 60, 1, False, False)
                Case 110:
                    Call Create_Enemy(25, 11, 1000, 0, 88, 90, 90, 60, 1, False, False)
                Case 115:
                    Call Create_Enemy(25, 11, 4000, 0, 88, 90, 90, 60, 1, False, False)
                Case 120:
                    Call Create_Enemy(25, 11, 3000, 0, 88, 90, 90, 60, 1, False, False)
                Case 125:
                    Call Create_Enemy(25, 11, 3900, 0, 88, 90, 90, 60, 1, False, False)
                Case 130:
                    Call Create_Enemy(25, 11, 2600, 0, 88, 90, 90, 60, 1, False, False)
                Case 135:
                    Call Create_Enemy(25, 11, 1000, 0, 88, 90, 90, 60, 1, False, False)
                Case 140:
                    Call Create_Enemy(25, 11, 1800, 0, 88, 90, 90, 60, 1, False, False)
                Case 145:
                    Call Create_Enemy(25, 11, 4700, 0, 88, 90, 90, 60, 1, False, False)
                Case 150:
                    Call Create_Enemy(25, 11, 2600, 0, 88, 90, 90, 60, 1, False, False)
                Case 155:
                    Call Create_Enemy(25, 11, 3500, 0, 88, 90, 90, 60, 1, False, False)
                Case 160:
                    Call Create_Enemy(25, 11, 1600, 0, 88, 90, 90, 60, 1, False, False)
                Case 165:
                    Call Create_Enemy(25, 11, 2900, 0, 88, 90, 90, 60, 1, False, False)
                Case 170:
                    Call Create_Enemy(25, 11, 4700, 0, 88, 90, 90, 60, 1, False, False)
                Case 175:
                    Call Create_Enemy(25, 11, 4000, 0, 88, 90, 90, 60, 1, False, False)
                Case 180:
                    Call Create_Enemy(25, 11, 3000, 0, 88, 90, 90, 60, 1, False, False)
                Case 185:
                    Call Create_Enemy(25, 11, 2000, 0, 88, 90, 90, 60, 1, False, False)
                Case 190:
                    Call Create_Enemy(25, 11, 1000, 0, 88, 90, 90, 60, 1, False, False)
                Case 195:
                    Call Create_Enemy(25, 11, 3500, 0, 88, 90, 90, 60, 1, False, False)
                Case 200:
                    Call Create_Enemy(25, 11, 2300, 0, 88, 90, 90, 60, 1, False, False)
                Case 205:
                    Call Create_Enemy(25, 11, 3300, 0, 88, 90, 90, 60, 1, False, False)
                Case 225:
                    Call Create_Enemy(21, 11, 4800, 0, 88, 30, 160, 180, 0, False, True)
                Case 250:
                    Call Send_MSG("> Command base: More Scyths Detected ! ", 0, 0)
                Case 305:
                    Call Create_Enemy(25, 11, 1200, 0, 88, 90, 90, 60, 1, False, False)
                Case 310:
                    Call Create_Enemy(25, 11, 1800, 0, 88, 90, 90, 60, 1, False, False)
                Case 315:
                    Call Create_Enemy(25, 11, 3000, 0, 88, 90, 90, 60, 1, False, False)
                Case 320:
                    Call Create_Enemy(25, 11, 2000, 0, 88, 90, 90, 60, 1, False, False)
                Case 325:
                    Call Create_Enemy(25, 11, 2900, 0, 88, 90, 90, 60, 1, False, False)
                Case 330:
                    Call Create_Enemy(25, 11, 3600, 0, 88, 90, 90, 60, 1, False, False)
                Case 335:
                    Call Create_Enemy(25, 11, 1000, 0, 88, 90, 90, 60, 1, False, False)
                Case 340:
                    Call Create_Enemy(25, 11, 1800, 0, 88, 90, 90, 60, 1, False, False)
                Case 345:
                    Call Create_Enemy(25, 11, 3000, 0, 88, 90, 90, 60, 1, False, False)
                Case 350:
                    Call Create_Enemy(25, 11, 3900, 0, 88, 90, 90, 60, 1, False, False)
                Case 355:
                    Call Create_Enemy(25, 11, 2200, 0, 88, 90, 90, 60, 1, False, False)
                Case 360:
                    Call Create_Enemy(25, 11, 4400, 0, 88, 90, 90, 60, 1, False, False)
                Case 365:
                    Call Create_Enemy(25, 11, 3300, 0, 88, 90, 90, 60, 1, False, False)
                Case 370:
                    Call Create_Enemy(25, 11, 1100, 0, 88, 90, 90, 60, 1, False, False)
                Case 375:
                    Call Create_Enemy(25, 11, 3000, 0, 88, 90, 90, 60, 1, False, False)
                Case 380:
                    Call Create_Enemy(25, 11, 2700, 0, 88, 90, 90, 60, 1, False, False)
                Case 385:
                    Call Create_Enemy(25, 11, 800, 0, 88, 90, 90, 60, 1, False, False)
                Case 390:
                    Call Create_Enemy(25, 11, 3800, 0, 88, 90, 90, 60, 1, False, False)
                Case 395:
                    Call Create_Enemy(25, 11, 4700, 0, 88, 90, 90, 60, 1, False, False)
                Case 410:
                    Call Create_Enemy(25, 11, 900, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1200, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1500, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1800, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2100, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3600, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3900, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4200, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4500, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4800, 10, 88, 90, 120, 80, 1, False, False)
                Case 410:
                    Call Create_Enemy(25, 11, 2100, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2400, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2700, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3000, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3300, 0, 88, 90, 120, 80, 1, False, False)
                Case 450:
                    Call Create_Enemy(25, 11, 1500, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1800, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2100, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2400, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2700, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3000, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3300, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3600, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3900, 0, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4200, 0, 88, 90, 120, 80, 1, False, False)
                Case 480:
                    Call Create_Enemy(21, 11, 2500, 0, 88, 30, 160, 180, 0, False, True)
                Case 510:
                    Call Create_Enemy(25, 11, 900, 0, 86, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4700, 0, 84, 90, 120, 80, 1, False, False)
                Case 520:
                    Call Create_Enemy(25, 11, 900, 0, 86, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4700, 0, 84, 90, 120, 80, 1, False, False)
                Case 530:
                    Call Create_Enemy(25, 11, 900, 0, 86, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4700, 0, 84, 90, 120, 80, 1, False, False)
                Case 540:
                    Call Create_Enemy(25, 11, 900, 0, 86, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4700, 0, 84, 90, 120, 80, 1, False, False)
                Case 550:
                    Call Create_Enemy(25, 11, 900, 0, 86, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4700, 0, 84, 90, 120, 80, 1, False, False)
                Case 560:
                    Call Create_Enemy(25, 11, 900, 0, 86, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4700, 0, 84, 90, 120, 80, 1, False, False)
                Case 570:
                    Call Create_Enemy(25, 11, 900, 0, 86, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4700, 0, 84, 90, 120, 80, 1, False, False)
                Case 580:
                    Call Create_Enemy(25, 11, 900, 0, 86, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4700, 0, 84, 90, 120, 80, 1, False, False)
                Case 740:
                    Call Create_Enemy(19, 21, 1600, 100, 555, 30, 800, 600, 123, False, False)
                Case 880:
                    Call Create_Enemy(19, 21, 3100, 100, 555, 30, 800, 600, 123, False, False)
                Case 1000:
                    Call Send_MSG("> Command Base: Barricade ahead ! ", 0, 0)
                Case 1080:
                    Call Create_Enemy(25, 11, 900, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1200, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1500, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1800, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2100, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2400, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2700, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3000, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3600, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3600, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3900, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4200, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4500, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4800, 10, 88, 90, 120, 80, 1, False, False)
                Case 1160:
                    Call Create_Enemy(25, 11, 900, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1200, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1500, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1800, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2100, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2400, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2700, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3000, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3600, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3600, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3900, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4200, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4500, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4800, 10, 88, 90, 120, 80, 1, False, False)
                Case 1220:
                    Call Create_Enemy(25, 11, 900, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1200, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1500, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 1800, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2100, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2400, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 2700, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3000, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3600, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3600, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 3900, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4200, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4500, 10, 88, 90, 120, 80, 1, False, False)
                    Call Create_Enemy(25, 11, 4800, 10, 88, 90, 120, 80, 1, False, False)
                Case 1260:
                    Call Create_Enemy(21, 11, 1000, 0, 88, 30, 160, 180, 0, False, True)
                    Call Send_MSG(" INCOMING ENEMY ", 0, 1)
                Case 1330:      'Mother Scyth
                    Boss_Name.Caption = "SCYTH LORD"
                    Call Create_Enemy(24, 22, 2000, 0, 7979, 250, 12500, 4500, 3333, True, False)
                End Select
            End If
            
            If C_GEC(3) > 1 And C_GEC(4) = 0 Then
                Select Case (GEC - C_GEC(3))
                Case 60:
                Call Send_MSG(" Final Stage: The Chaotic End ", 0, 0)
                Case 130:
                    Call Send_MSG(" > Command Base: Becareful, Black Winter ", 0, 0)
                Case 160:
                    'ARMY TYPE
                    Call Create_Enemy(22, 11, 1300, 0, 83, 65, 220, 170, (Rnd * 13 + 3), False, False)
                Case 170:
                    Call Create_Enemy(22, 11, 1200, 0, 83, 65, 220, 170, (Rnd * 13 + 3), False, False)
                Case 180:
                    Call Create_Enemy(22, 11, 1100, 0, 83, 65, 220, 170, (Rnd * 13 + 3), False, False)
                Case 190:
                    Call Create_Enemy(22, 11, 1000, 0, 83, 65, 220, 170, (Rnd * 13 + 3), False, False)
                Case 200:
                    Call Create_Enemy(22, 11, 900, 0, 83, 65, 220, 170, (Rnd * 13 + 3), False, False)
                Case 250:
                    'Fat Grey Ship
                    Call Create_Enemy(21, 11, 2400, 0, 88, 30, 160, 180, 0, False, True)
                    Call Create_Enemy(20, 12, 4800, 0, 88, 30, 250, 170, 32147, False, False)
                    Call Create_Enemy(20, 12, 1100, 0, 88, 30, 250, 170, 12369, False, False)
                Case 320:
                    'Fat Grey Ship
                    Call Create_Enemy(20, 12, 4800, 0, 88, 30, 250, 170, 32147, False, False)
                    Call Create_Enemy(20, 12, 1100, 0, 88, 30, 250, 170, 12369, False, False)
                Case 390:
                    Call Create_Enemy(19, 21, 1600, 100, 555, 30, 2000, 800, 123, False, False)
                Case 440:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 470:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 490:
                    'Fat Grey Ship
                    Call Create_Enemy(20, 12, 4800, 0, 88, 30, 250, 170, 32147, False, False)
                    Call Create_Enemy(20, 12, 1100, 0, 88, 30, 250, 170, 12369, False, False)
                Case 520:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 530:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 540:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 570:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 2, False, False)
                Case 580:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 2, False, False)
                Case 590:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 2, False, False)
                Case 600:
                    Call Create_Enemy(25, 11, 1000, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 5000, 0, 88, 80, 140, 90, 2, False, False)
                Case 610:
                    Call Create_Enemy(25, 11, 1200, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 4800, 0, 88, 80, 140, 90, 2, False, False)
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 615:
                    Call Create_Enemy(25, 11, 1300, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 4700, 0, 88, 80, 140, 90, 2, False, False)
                Case 620:
                    Call Create_Enemy(21, 11, 2400, 0, 88, 30, 160, 180, 0, False, True)
                    Call Create_Enemy(25, 11, 1400, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 4600, 0, 88, 80, 140, 90, 2, False, False)
                Case 630:
                    Call Create_Enemy(25, 11, 1500, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 4500, 0, 88, 80, 140, 90, 2, False, False)
                Case 640:
                    Call Create_Enemy(25, 11, 1600, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 4400, 0, 88, 80, 140, 90, 2, False, False)
                Case 650:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                    Call Create_Enemy(25, 11, 1700, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 4300, 0, 88, 80, 140, 90, 2, False, False)
                Case 660:
                    Call Create_Enemy(25, 11, 1800, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 4200, 0, 88, 80, 140, 90, 2, False, False)
                Case 670:
                    Call Create_Enemy(25, 11, 1900, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 4100, 0, 88, 80, 140, 90, 2, False, False)
                Case 680:
                    Call Create_Enemy(25, 11, 2000, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 4000, 0, 88, 80, 140, 90, 2, False, False)
                Case 690:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                    Call Create_Enemy(25, 11, 2100, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 3900, 0, 88, 80, 140, 90, 2, False, False)
                Case 700:
                    Call Create_Enemy(25, 11, 2200, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 3800, 0, 88, 80, 140, 90, 2, False, False)
                Case 710:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                    Call Create_Enemy(25, 11, 2300, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 3700, 0, 88, 80, 140, 90, 2, False, False)
                Case 720:
                    Call Create_Enemy(25, 11, 2400, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 3600, 0, 88, 80, 140, 90, 2, False, False)
                Case 730:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                    Call Create_Enemy(25, 11, 2500, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 3500, 0, 88, 80, 140, 90, 2, False, False)
                Case 740:
                    Call Create_Enemy(25, 11, 2600, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 3400, 0, 88, 80, 140, 90, 2, False, False)
                Case 750:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                    Call Create_Enemy(25, 11, 2700, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 2900, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 3100, 0, 88, 80, 140, 90, 2, False, False)
                    Call Create_Enemy(25, 11, 3300, 0, 88, 80, 140, 90, 2, False, False)
                Case 770:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 775:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 780:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 785:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 810:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 815:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 820:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 825:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 830:
                    Call Create_Enemy(22, 11, 1300, 0, 83, 65, 220, 170, (Rnd * 13 + 3), False, False)
                    Call Create_Enemy(22, 11, 2300, 0, 83, 65, 220, 170, (Rnd * 13 + 3), False, False)
                    Call Create_Enemy(22, 11, 3300, 0, 83, 65, 220, 170, (Rnd * 13 + 3), False, False)
                    Call Create_Enemy(22, 11, 4300, 0, 83, 65, 220, 170, (Rnd * 13 + 3), False, False)
                Case 880:
                    Call Create_Enemy(21, 11, 1000, 0, 88, 30, 160, 180, 0, False, True)
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 920:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 940:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                
                Case 960:
                    Call Create_Enemy(19, 21, 1600, 100, 555, 30, 2200, 800, 123, False, False)
                Case 970:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 990:
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1000:
                    Call Create_Enemy(19, 21, 2600, 100, 555, 30, 2200, 800, 123, False, False)
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1030:
                    Call Create_Enemy(19, 21, 1600, 100, 555, 30, 2200, 800, 123, False, False)
                Case 1070:
                    Call Create_Enemy(19, 21, 2600, 100, 555, 30, 2200, 800, 123, False, False)
                Case 1090:
                    Call Create_Enemy(19, 21, 1600, 100, 555, 30, 2200, 800, 123, False, False)
                Case 1150:
                    Call Create_Enemy(25, 11, 3600, 0, 88, 80, 140, 90, 2, False, False)
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1160:
                    Call Create_Enemy(25, 11, 3600, 0, 88, 80, 140, 90, 2, False, False)
                Case 1170:
                    Call Create_Enemy(25, 11, 3600, 0, 88, 80, 140, 90, 2, False, False)
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1180:
                    Call Create_Enemy(25, 11, 3600, 0, 88, 80, 140, 90, 2, False, False)
                Case 1190:
                    Call Create_Enemy(25, 11, 3600, 0, 88, 80, 140, 90, 2, False, False)
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1290:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1300:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1310:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1320:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    For i = 1 To 8
                        Call Create_Enemy(Rnd * 2 + 28, Rnd * 4, Rnd * 4820 + 820, 0, 31313, Rnd * 30 + 60, Rnd * 20 + 40, 15, 0, False, False)
                    Next i
                Case 1330:
                    'Fat Grey Ship
                    Call Create_Enemy(21, 11, 2800, 0, 88, 30, 160, 180, 0, False, True)
                    Call Create_Enemy(20, 12, 4800, 0, 88, 30, 250, 170, 32147, False, False)
                    Call Create_Enemy(20, 12, 1100, 0, 88, 30, 250, 170, 12369, False, False)
                Case 1380:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1390:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1400:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1410:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1420:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1430:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1440:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1450:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1460:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1470:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1480:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1490:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1500:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1510:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1520:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1530:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1540:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1550:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1555:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1560:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                Case 1565:
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                    Call Create_Enemy(25, 11, (Rnd * 3800 + 800), 0, 88, 80, 140, 90, 1, False, False)
                
                Case 1570:
                    Call Send_MSG(" ALERT! INCOMING MOTHERSHIP ! ", 0, 1)
                Case 1650:
                    Boss_Name.Caption = "BLACK DISC"
                    Call Create_Enemy(18, 22, 2400, 0, 17931, 100, 15000, 8000, 666, True, False)
                End Select
            End If
            
            If C_GEC(4) > 1 Then
                Select Case (GEC - C_GEC(4))
                Case 1:
                    KilledBoss = True
                    Call Check_Score
                    GEC = -400
                End Select
            End If
            
    End Select
End Sub

Private Sub Check_Fire(Fire_Pattern As Integer, i As Integer)
'The firing patterns of the bullets
    Select Case (Fire_Pattern)
    Case 1: 'Homing
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 10, ShipX + (Img(0).ImageWidth * 15) / 2, ShipY + (Img(0).ImageHeight * 15) / 2)
    Case 2: 'Down
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 160, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, 7920)
    Case 3: 'Diagonal Left 1
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 720, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 6000)
    Case 4: 'Diagonal Left 2
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 720, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 5000)
    Case 5: 'Diagonal Left 3
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 720, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 4000)
    Case 6: 'Diagonal Left 4
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 720, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 3000)
    Case 7: 'Diagonal Left 5
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 720, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 2000)
    Case 8: 'Diagonal Left 6
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 720, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 1000)
    Case 9: 'Left
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 720, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15))
    Case 10: 'Diagonal Right 1
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 5640, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 6000)
    Case 11: 'Diagonal Right 2
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 5640, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 5000)
    Case 12: 'Diagonal Right 3
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 5640, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 4000)
    Case 13: 'Diagonal Right 4
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 5640, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 3000)
    Case 14: 'Diagonal Right 5
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 5640, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 2000)
    Case 15: 'Diagonal Right 6
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 5640, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15) + 1000)
    Case 16: 'Right
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 70, 5640, EShip(i).PosY + (Img(EShip(i).Kind).ImageWidth * 15))
    Case 17:
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 40, (Rnd * 4920 + 720), 9000)
    Case 18:
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 40, ShipX, ShipY)
    Case 19:
        Call Create_EBullet(EShip(i).Bul_Kind, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, EShip(i).PosY + (Img(EShip(i).Kind).ImageHeight * 15) / 2, 1, 30, EShip(i).PosX + (Img(EShip(i).Kind).ImageWidth * 15) / 2, 9000)
    End Select
End Sub

Private Sub Create_EBullet(Kind As Integer, StartX As Integer, StartY As Integer, Direction As Integer, Speed As Integer, DestinationX As Integer, DestinationY As Integer)
'Creates an enemy bullet when called
'Parameter list:
'                   imagelist index for bullet
'                   starting X coord
'                   starting Y coord
'                   bullet speed
'                   direction of the bullet
'                   ending X coord
'                   ending Y coord

Dim i As Integer
If GEC > 0 Then
    For i = 0 To 200
        If EBull(i).Used = False Then
            EBull(i).Used = True
            EBull(i).Kind = Kind
            EBull(i).StartX = StartX
            EBull(i).StartY = StartY
            EBull(i).PosX = StartX
            EBull(i).PosY = StartY
            EBull(i).Dire = Direction
            EBull(i).Spd = Speed
            EBull(i).DestX = DestinationX
            EBull(i).DestY = DestinationY
            Exit For
        End If
    Next i
End If
End Sub

Private Sub Create_Enemy(Kind As Integer, Bullet_Kind As Integer, StartX As Integer, StartY As Integer, Pattern As Integer, Speed As Integer, Hit_Points As Integer, Worth As Integer, Firing_Pattern As Integer, Boss As Boolean, Bonus As Boolean)
'This sub creates and enemy when called and place
'it on the play field
'Parameter list:    imagelist index for ship sprite
'                   imagelist index for bullet sprite
'                   starting X coord
'                   starting Y coord
'                   movement pattern
'                   ship movement speed
'                   life points
'                   the amount of score awarded if destroyed
'                   the pattern of the bullets it fires
'                   whether it is a boss
'                   whether is carries bonus

Dim i As Integer
    For i = 1 To 40
        If EShip(i).Used = False Then
            EShip(i).Used = True
            EShip(i).Kind = Kind
            EShip(i).Bul_Kind = Bullet_Kind
            EShip(i).PosX = StartX
            EShip(i).PosY = StartY
            EShip(i).Patt = Pattern
            EShip(i).Max_Life = Hit_Points
            EShip(i).Life = Hit_Points
            EShip(i).Score = Worth
            EShip(i).Span = 0
            EShip(i).Frame = 1
            EShip(i).Spd = Speed
            EShip(i).Has_Bonus = Bonus
            EShip(i).Firing_Patt = Firing_Pattern
            EShip(i).Boss = False
            
            'Drop down the boss bar if a boss is created
            If Boss = True Then
                EShip(i).Boss = True
                Moving_BossBar = 1
                Boss_Bar_Top.Width = Boss_Bar_Bottom.Width
            End If
            Exit For
        End If
    Next i
End Sub

Private Sub Create_Bonus(PosX As Integer, PosY As Integer)
'This sub creates a powerup object and allows it to be
'placed on the screen based on its called X and Y coord
Dim i As Integer
    For i = 0 To 2
        If PowerUp(i).Used = False Then
            PowerUp(i).Used = True
            PowerUp(i).PosX = PosX
            PowerUp(i).PosY = PosY
            PowerUp(i).Frame = 1
            Exit For
        End If
    Next i
End Sub
