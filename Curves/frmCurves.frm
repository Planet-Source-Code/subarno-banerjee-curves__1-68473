VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCurves 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conic Sections"
   ClientHeight    =   11040
   ClientLeft      =   -195
   ClientTop       =   195
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tab1 
      Height          =   1935
      Left            =   50
      TabIndex        =   0
      Top             =   45
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Circle"
      TabPicture(0)   =   "frmCurves.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCCen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCrad"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCx"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCy"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCrad"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDC"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCCl"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Parabola"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPcl"
      Tab(1).Control(1)=   "cmdPD"
      Tab(1).Control(2)=   "txtPa"
      Tab(1).Control(3)=   "lblP2"
      Tab(1).Control(4)=   "lblP1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Ellipse"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblE1"
      Tab(2).Control(1)=   "lblE2(0)"
      Tab(2).Control(2)=   "lblE2(1)"
      Tab(2).Control(3)=   "lblE2(2)"
      Tab(2).Control(4)=   "lblE2(3)"
      Tab(2).Control(5)=   "linE1"
      Tab(2).Control(6)=   "linE2"
      Tab(2).Control(7)=   "lblE3"
      Tab(2).Control(8)=   "txtEa"
      Tab(2).Control(9)=   "txtEb"
      Tab(2).Control(10)=   "cmdED"
      Tab(2).Control(11)=   "cmdEcl"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Hyperbola"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdHCl"
      Tab(3).Control(1)=   "cmdHD"
      Tab(3).Control(2)=   "txtHb"
      Tab(3).Control(3)=   "txtHa"
      Tab(3).Control(4)=   "lblH3"
      Tab(3).Control(5)=   "linH4"
      Tab(3).Control(6)=   "linH1"
      Tab(3).Control(7)=   "lblH2(7)"
      Tab(3).Control(8)=   "lbH2(6)"
      Tab(3).Control(9)=   "lblH2(5)"
      Tab(3).Control(10)=   "lbH2(4)"
      Tab(3).Control(11)=   "lblH1"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "General"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdDrw"
      Tab(4).Control(1)=   "cmdClear"
      Tab(4).Control(2)=   "txtGc"
      Tab(4).Control(3)=   "txtGf"
      Tab(4).Control(4)=   "Text5"
      Tab(4).Control(5)=   "txtGg"
      Tab(4).Control(6)=   "txtGh"
      Tab(4).Control(7)=   "txtGb"
      Tab(4).Control(8)=   "txtGa"
      Tab(4).Control(9)=   "lblG2"
      Tab(4).Control(10)=   "lblGeq"
      Tab(4).ControlCount=   11
      TabCaption(5)   =   "Customise"
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblAcol"
      Tab(5).Control(1)=   "Label2"
      Tab(5).Control(2)=   "lblbcol"
      Tab(5).Control(3)=   "picCcol"
      Tab(5).Control(4)=   "picACol"
      Tab(5).Control(5)=   "cdb"
      Tab(5).Control(6)=   "picbcol"
      Tab(5).ControlCount=   7
      Begin VB.PictureBox picbcol 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -72000
         MousePointer    =   99  'Custom
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   49
         ToolTipText     =   "Click here to change..."
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdDrw 
         Caption         =   "&Draw"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73440
         TabIndex        =   46
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72840
         TabIndex        =   45
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtGc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -71040
         MaxLength       =   5
         MousePointer    =   99  'Custom
         TabIndex        =   44
         ToolTipText     =   "Enter value for 'c'..."
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtGf 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -71650
         MaxLength       =   5
         MousePointer    =   99  'Custom
         TabIndex        =   43
         ToolTipText     =   "Enter value for 'f'..."
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -69120
         MousePointer    =   99  'Custom
         TabIndex        =   42
         ToolTipText     =   "Enter x-coordinate of the centre..."
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox txtGg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72420
         MaxLength       =   5
         MousePointer    =   99  'Custom
         TabIndex        =   41
         ToolTipText     =   "Enter value for 'g'..."
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtGh 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -73390
         MaxLength       =   5
         MousePointer    =   99  'Custom
         TabIndex        =   40
         ToolTipText     =   "Enter value for 'h'..."
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtGb 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -74160
         MaxLength       =   5
         MousePointer    =   99  'Custom
         TabIndex        =   39
         ToolTipText     =   "Enter value for 'b'..."
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtGa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -74760
         MaxLength       =   5
         MousePointer    =   99  'Custom
         TabIndex        =   38
         ToolTipText     =   "Enter value for 'a'..."
         Top             =   960
         Width           =   255
      End
      Begin MSComDlg.CommonDialog cdb 
         Left            =   -70800
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox picACol 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   -72000
         MousePointer    =   99  'Custom
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   37
         ToolTipText     =   "Click here to change..."
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox picCcol 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   -72000
         MousePointer    =   99  'Custom
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   36
         ToolTipText     =   "Click here to change..."
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdHCl 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         TabIndex        =   32
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdEcl 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         TabIndex        =   31
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdPcl 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         TabIndex        =   30
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdCCl 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   29
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdHD 
         Caption         =   "&Draw"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         TabIndex        =   28
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdED 
         Caption         =   "&Draw"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         TabIndex        =   27
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdPD 
         Caption         =   "&Draw"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         TabIndex        =   26
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdDC 
         Caption         =   "&Draw"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   25
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtHb 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   -73200
         MousePointer    =   99  'Custom
         TabIndex        =   23
         ToolTipText     =   "Enter value of 'b'..."
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtHa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   -74040
         MousePointer    =   99  'Custom
         TabIndex        =   22
         ToolTipText     =   "Enter value of 'a'..."
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtEb 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   -73200
         MousePointer    =   99  'Custom
         TabIndex        =   15
         ToolTipText     =   "Enter value of 'b'..."
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtEa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   -74040
         MousePointer    =   99  'Custom
         TabIndex        =   14
         ToolTipText     =   "Enter value of 'a'..."
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtPa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   -72840
         MaxLength       =   5
         MousePointer    =   99  'Custom
         TabIndex        =   8
         ToolTipText     =   "Enter value of 'a'..."
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtCrad 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         MousePointer    =   99  'Custom
         TabIndex        =   5
         ToolTipText     =   "Enter radius of the circle..."
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtCy 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2400
         MaxLength       =   5
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "Enter y-coordinate of the centre..."
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox txtCx 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "Enter x-coordinate of the centre..."
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblbcol 
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Color :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   -73800
         TabIndex        =   48
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblG2 
         BackStyle       =   0  'Transparent
         Caption         =   "2         2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   -74400
         TabIndex        =   47
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Curve Color :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   -73800
         TabIndex        =   35
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblAcol 
         BackStyle       =   0  'Transparent
         Caption         =   "Axis Color :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   -73800
         TabIndex        =   34
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblGeq 
         BackStyle       =   0  'Transparent
         Caption         =   "   x +     y + 2     x y + 2     x + 2     y +       = 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   -74640
         TabIndex        =   33
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lblH3 
         BackStyle       =   0  'Transparent
         Caption         =   "-      =  1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   495
         Left            =   -73440
         TabIndex        =   24
         Top             =   915
         Width           =   1575
      End
      Begin VB.Line linH4 
         BorderColor     =   &H00004080&
         BorderWidth     =   2
         X1              =   -73200
         X2              =   -72720
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line linH1 
         BorderColor     =   &H00004080&
         BorderWidth     =   2
         X1              =   -74040
         X2              =   -73560
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label lblH2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   7
         Left            =   -72810
         TabIndex        =   21
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label lbH2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   6
         Left            =   -73650
         TabIndex        =   20
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label lblH2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   5
         Left            =   -72840
         TabIndex        =   19
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbH2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   4
         Left            =   -73680
         TabIndex        =   18
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblH1 
         BackStyle       =   0  'Transparent
         Caption         =   "x      y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   495
         Left            =   -73920
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblE3 
         BackStyle       =   0  'Transparent
         Caption         =   "+      =  1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   495
         Left            =   -73500
         TabIndex        =   16
         Top             =   915
         Width           =   1575
      End
      Begin VB.Line linE2 
         BorderColor     =   &H00004080&
         BorderWidth     =   2
         X1              =   -73200
         X2              =   -72720
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line linE1 
         BorderColor     =   &H00004080&
         BorderWidth     =   2
         X1              =   -74040
         X2              =   -73560
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label lblE2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   3
         Left            =   -72810
         TabIndex        =   13
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label lblE2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   2
         Left            =   -73650
         TabIndex        =   12
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label lblE2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   1
         Left            =   -72840
         TabIndex        =   11
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblE2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   0
         Left            =   -73680
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblE1 
         BackStyle       =   0  'Transparent
         Caption         =   "x      y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   495
         Left            =   -73920
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblP2 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   -73800
         TabIndex        =   7
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblP1 
         Caption         =   "y   =  4     x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   495
         Left            =   -74040
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblCrad 
         BackStyle       =   0  'Transparent
         Caption         =   "Radius :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblCCen 
         BackStyle       =   0  'Transparent
         Caption         =   "Centre :   (      ,      )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   840
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmCurves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim px, py, qx, qy, y, a, b, h, g, f, c, mx, my, y1 As Long
Dim gen As Boolean
Dim acol, ccol, bcol As String
Private Sub cmdCCl_Click()
txtCx.Text = ""
txtCy.Text = ""
txtCrad.Text = ""
txtCx.SetFocus
End Sub
Private Sub cmdClear_Click()
txtGa.Text = ""
txtGb.Text = ""
txtGh.Text = ""
txtGg.Text = ""
txtGf.Text = ""
txtGc.Text = ""
txtGa.SetFocus
clear
End Sub
Private Sub cmdDC_Click()
On Error GoTo CEr
If gen = False Then
Me.Circle ((7000 + (Int(Val(txtCx.Text)) * 10)), (6000 - (Int(Val(txtCy.Text))) * 10)), (Int(Val(txtCrad.Text)) * 10), picCcol.BackColor
Else
Me.Circle ((Int(7000 - (g / a))), (Int(6000 - (f / a)))), (Int((((g ^ 2) / (a ^ 2)) + ((f ^ 2) / (a ^ 2)) - (c / a)) ^ 0.5)), picCcol.BackColor
gen = False
End If
CEr:
If Err.Number = 6 Then
MsgBox ("Too large values to handle!")
Call cmdCCl_Click
Exit Sub
End If
End Sub
Private Sub cmdDrw_Click()
On Error GoTo GEr
a = Val(txtGa.Text)
b = Val(txtGb.Text)
h = Val(txtGh.Text)
g = Val(txtGg.Text)
f = Val(txtGf.Text)
c = Val(txtGc.Text)
For x = 0 To 15000
y = Int((((-2) * ((h * x) + (2 * f))) + ((((2 * h * x) + (2 * f)) ^ 2) - 4 * b * ((a * (x ^ 2)) + (2 * g * x) + c) ^ 0.5)) / (2 * b))
y1 = Int((((-2) * ((h * x) + (2 * f))) - ((((2 * h * x) + (2 * f)) ^ 2) - 4 * b * ((a * (x ^ 2)) + (2 * g * x) + c) ^ 0.5)) / (2 * b))
Me.Line (7000 + x, 6000 - y)-(px, qx), picCcol.BackColor
Me.Line (7000 + x, 6000 - y1)-(px, my), picCcol.BackColor
px = 7000 + x
qx = 6000 - y
my = 6000 - y1
Next
GEr:
If Err.Number = 6 Then
MsgBox ("Too large values to handle!")
Call cmdClear_Click
Exit Sub
End If
End Sub
Private Sub cmdEcl_Click()
txtEa.Text = ""
txtEb.Text = ""
txtEa.SetFocus
clear
End Sub
Private Sub cmdED_Click()
On Error GoTo EEr
If gen = False Then
a = Val(txtEa.Text)
b = Val(txtEb.Text)
px = 7000 - a
qx = 6000
mx = 0
my = 0
End If
For x = -a To 0
If y < 0 Then
Exit For
End If
y = Int(((b ^ 2) - (((b * x) / a) ^ 2)) ^ 0.5)
Me.Line (px, qx)-(x + 7000 + mx, 6000 - y + my), picCcol.BackColor
Me.Line ((px + (2 * (7000 - px))), qx)-((x * (-1)) + 7000 + mx, 6000 - y + my), picCcol.BackColor
Me.Line (px, (qx + ((6000 - qx) * 2)))-(x + 7000 + mx, 6000 + y + my), picCcol.BackColor
Me.Line ((px + (2 * (7000 - px))), (qx + ((6000 - qx) * 2)))-((x * (-1)) + 7000 + mx, 6000 + y + my), picCcol.BackColor
px = x + 7000 + mx
qx = 6000 - y + my
Next
EEr:
If Err.Number = 6 Then
MsgBox ("Too large values to handle!")
Call cmdEcl_Click
Exit Sub
End If
End Sub
Private Sub cmdHCl_Click()
txtHa.Text = ""
txtHb.Text = ""
txtHa.SetFocus
clear
End Sub
Private Sub cmdHD_Click()
On Error GoTo HEr
If gen = False Then
a = Val(txtHa.Text)
b = Val(txtHb.Text)
px = 7000 + a
qx = 6000
mx = 0
my = 0
End If
For x = a To 10000
If y > 10000 Then
Exit For
End If
y = Int(((((b * x) / a) ^ 2) - (b ^ 2)) ^ 0.5)
Me.Line (px, qx)-(x + 7000 + mx, 6000 - y + my), picCcol.BackColor
Me.Line (px, (qx + ((6000 - qx) * 2)))-(x + 7000 + mx, 6000 + y + my), picCcol.BackColor
Me.Line ((14000 - px), qx)-(7000 - x + mx, 6000 - y + my), picCcol.BackColor
Me.Line ((14000 - px), (qx + ((6000 - qx) * 2)))-(7000 - x + mx, 6000 + y + my), picCcol.BackColor
px = x + 7000 + mx
qx = 6000 - y + my
Next
HEr:
If Err.Number = 6 Then
MsgBox ("Too large values to handle!")
Call cmdHCl_Click
Exit Sub
End If
End Sub
Private Sub cmdPcl_Click()
txtPa.Text = ""
txtPa.SetFocus
clear
End Sub
Private Sub cmdPD_Click()
On Error GoTo PEr
If gen = False Then
a = Val(txtPa.Text)
mx = 0
my = 0
End If
For x = 0 To 10000
If y > 7000 Then
Exit For
End If
y = (4 * a * x) ^ 0.5
Me.Line (px, py)-(x + 7000 + mx, y + 6000 + my), picCcol.BackColor
px = x + 7000 + mx
py = y + 6000 + my
Me.Line (qx, qy)-(x + 7000 + mx, 6000 - y + my), picCcol.BackColor
qx = x + 7000 + mx
qy = 6000 - y + my
Next
PEr:
If Err.Number = 6 Then
MsgBox ("Too large values to handle!")
Call cmdPcl_Click
Exit Sub
End If
End Sub
Private Sub Form_Load()
Call LoadSettings
Me.picACol.BackColor = acol
Me.picCcol.BackColor = ccol
Me.picbcol.BackColor = bcol
Me.BackColor = bcol
Me.tab1.BackColor = bcol
Me.Line (7000, 6000)-(1000, 6000), picACol.BackColor
Me.Line (7000, 6000)-(13000, 6000), picACol.BackColor
Me.Line (7000, 1000)-(7000, 6000), picACol.BackColor
Me.Line (7000, 6000)-(7000, 13000), picACol.BackColor
px = 7000
py = 6000
qx = 7000
qy = 6000
mx = 0
my = 0
End Sub
Private Sub picACol_Click()
cdb.ShowColor
picACol.BackColor = cdb.Color
Call SaveSettings
clear
End Sub
Private Sub picbcol_Click()
cdb.ShowColor
picbcol.BackColor = cdb.Color
Call SaveSettings
clear
End Sub
Private Sub picCcol_Click()
cdb.ShowColor
picCcol.BackColor = cdb.Color
Call SaveSettings
clear
End Sub
Private Sub tab1_Click(PreviousTab As Integer)
clear
End Sub
Private Sub txtCrad_Change()
If IsNumeric(txtCrad.Text) = False And txtCrad.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtCrad.Text = ""
txtCrad.SetFocus
End If
End If
If txtCx.Text = "" Or txtCy.Text = "" Or txtCrad.Text = "" Then
cmdDC.Enabled = False
clear
Else
cmdDC.Enabled = True
End If
End Sub
Private Sub txtCx_Change()
If txtCx.Text = "" Or txtCy.Text = "" Or txtCrad.Text = "" Then
cmdDC.Enabled = False
clear
Else
cmdDC.Enabled = True
End If
End Sub
Private Sub txtCx_LostFocus()
If IsNumeric(txtCx.Text) = False And txtCx.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtCx.Text = ""
txtCx.SetFocus
End If
End If
End Sub
Private Sub txtCy_Change()
If txtCx.Text = "" Or txtCy.Text = "" Or txtCrad.Text = "" Then
cmdDC.Enabled = False
clear
Else
cmdDC.Enabled = True
End If
End Sub
Private Sub txtCy_LostFocus()
If IsNumeric(txtCy.Text) = False And txtCy.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtCy.Text = ""
txtCy.SetFocus
End If
End If
End Sub
Private Sub txtEa_Change()
If IsNumeric(txtEa.Text) = False And txtEa.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtEa.Text = ""
txtEa.SetFocus
End If
End If
If txtEa.Text = "" Or txtEb.Text = "" Then
cmdED.Enabled = False
Else
cmdED.Enabled = True
End If
End Sub
Private Sub txtEb_Change()
If IsNumeric(txtEb.Text) = False And txtEb.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtEb.Text = ""
txtEb.SetFocus
End If
End If
If txtEa.Text = "" Or txtEb.Text = "" Then
cmdED.Enabled = False
Else
cmdED.Enabled = True
End If
End Sub
Private Sub txtGa_Change()
If txtGa.Text <> "" And txtGb.Text <> "" And txtGh.Text <> "" And txtGg.Text <> "" And txtGf.Text <> "" And txtGc.Text <> "" Then
cmdDrw.Enabled = True
Else
cmdDrw.Enabled = False
End If
End Sub
Private Sub txtGa_LostFocus()
If IsNumeric(txtGa.Text) = False And txtGa.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtGa.Text = ""
txtGa.SetFocus
End If
End If
End Sub
Private Sub txtGb_Change()
Call txtGa_Change
End Sub
Private Sub txtGb_LostFocus()
If IsNumeric(txtGb.Text) = False And txtGb.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtGb.Text = ""
txtGb.SetFocus
End If
End If
End Sub
Private Sub txtGc_Change()
Call txtGa_Change
End Sub
Private Sub txtGc_LostFocus()
If IsNumeric(txtGc.Text) = False And txtGc.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtGc.Text = ""
txtGc.SetFocus
End If
End If
End Sub
Private Sub txtGf_Change()
Call txtGa_Change
End Sub
Private Sub txtGf_LostFocus()
If IsNumeric(txtGf.Text) = False And txtGf.Text <> "" Then
If MsgBox("Input mustbe a numeric value.", vbCritical, "Error") = vbOK Then
txtGf.Text = ""
txtGf.SetFocus
End If
End If
End Sub
Private Sub txtGg_Change()
Call txtGa_Change
End Sub
Private Sub txtGg_LostFocus()
If IsNumeric(txtGg.Text) = False And txtGg.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtGg.Text = ""
txtGg.SetFocus
End If
End If
End Sub
Private Sub txtGh_Change()
Call txtGa_Change
End Sub
Private Sub txtGh_LostFocus()
If IsNumeric(txtGh.Text) = False And txtGh.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtGh.Text = ""
txtGh.SetFocus
End If
End If
End Sub
Private Sub txtHa_Change()
If IsNumeric(txtHa.Text) = False And txtHa.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtHa.Text = ""
txtHa.SetFocus
End If
End If
If txtHa.Text = "" Or txtHb.Text = "" Then
cmdHD.Enabled = False
Else
cmdHD.Enabled = True
End If
End Sub
Private Sub txtHb_Change()
If IsNumeric(txtHb.Text) = False And txtHb.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtHb.Text = ""
txtHb.SetFocus
End If
End If
If txtHa.Text = "" Or txtHb.Text = "" Then
cmdHD.Enabled = False
Else
cmdHD.Enabled = True
End If
End Sub
Private Sub txtPa_Change()
If IsNumeric(txtPa.Text) = False And txtPa.Text <> "" Then
If MsgBox("Input must be a numeric value.", vbCritical, "Error") = vbOK Then
txtPa.Text = ""
txtPa.SetFocus
End If
End If
If txtPa.Text = "" Then
cmdPD.Enabled = False
Else
cmdPD.Enabled = True
End If
End Sub
Function clear()
Me.Cls
Call Form_Load
End Function
Function Pos(x As Integer)
If x < 0 Then
Pos = x * (-1)
Else
Pos = x
End If
End Function
Private Sub SaveSettings()
Call SaveSetting("Conics", "Settings", "ACol", Me.picACol.BackColor)
Call SaveSetting("Conics", "Settings", "BCol", Me.picbcol.BackColor)
Call SaveSetting("Conics", "Settings", "CCol", Me.picCcol.BackColor)
End Sub
Private Sub LoadSettings()
acol = GetSetting("Conics", "Settings", "Acol", 0)
bcol = GetSetting("Conics", "Settings", "Bcol", vbWhite)
ccol = GetSetting("Conics", "Settings", "Ccol", 255)
End Sub
