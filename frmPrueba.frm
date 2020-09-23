VERSION 5.00
Begin VB.Form frmPrueba 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MiTextBox Form Test"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmPrueba.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ProyectoTextBox.TextBox TextBox1 
      Height          =   315
      Left            =   510
      TabIndex        =   0
      Top             =   1710
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      BackColor_OnGotFocus=   12648384
      Alignment       =   2
      FontSize        =   8,25
      ForeColor       =   -2147483640
   End
   Begin ProyectoTextBox.TextBox TextBox2 
      Height          =   315
      Left            =   510
      TabIndex        =   1
      Top             =   2070
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      BackColor_OnGotFocus=   16761024
      Appearance_OnGotFocus=   1
      Alignment       =   2
      FontSize        =   8,25
      ForeColor       =   -2147483640
   End
   Begin ProyectoTextBox.TextBox TextBox3 
      Height          =   315
      Left            =   510
      TabIndex        =   2
      Top             =   2430
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      BackColor_OnGotFocus=   16761024
      Appearance_OnGotFocus=   1
      Alignment       =   2
      FontSize        =   8,25
      ForeColor       =   -2147483640
   End
   Begin ProyectoTextBox.TextBox TextBox4 
      Height          =   315
      Left            =   510
      TabIndex        =   3
      Top             =   2790
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      BackColor_OnGotFocus=   16761024
      Appearance_OnGotFocus=   1
      Alignment       =   2
      FontSize        =   8,25
      ForeColor       =   -2147483640
   End
   Begin ProyectoTextBox.TextBox TextBox5 
      Height          =   315
      Left            =   510
      TabIndex        =   4
      Top             =   3150
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      BackColor_OnGotFocus=   16761024
      Appearance_OnGotFocus=   1
      Alignment       =   2
      FontSize        =   8,25
      ForeColor       =   -2147483640
   End
   Begin ProyectoTextBox.TextBox TextBox6 
      Height          =   315
      Left            =   510
      TabIndex        =   5
      Top             =   3510
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      BackColor_OnGotFocus=   16761024
      Appearance_OnGotFocus=   1
      Alignment       =   2
      FontSize        =   8,25
      ForeColor       =   -2147483640
   End
   Begin ProyectoTextBox.TextBox TextBox7 
      Height          =   315
      Left            =   510
      TabIndex        =   6
      Top             =   3870
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      BackColor_OnGotFocus=   16761024
      Appearance_OnGotFocus=   1
      Alignment       =   2
      FontSize        =   8,25
      ForeColor       =   -2147483640
   End
   Begin ProyectoTextBox.TextBox TextBox8 
      Height          =   315
      Left            =   510
      TabIndex        =   7
      Top             =   4230
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      BackColor_OnGotFocus=   16761024
      Appearance_OnGotFocus=   1
      Alignment       =   2
      FontSize        =   8,25
      ForeColor       =   -2147483640
   End
   Begin ProyectoTextBox.TextBox TextBox9 
      Height          =   315
      Left            =   510
      TabIndex        =   8
      Top             =   4590
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      BackColor_OnGotFocus=   16761024
      Appearance_OnGotFocus=   1
      Alignment       =   2
      FontSize        =   8,25
      ForeColor       =   -2147483640
   End
   Begin ProyectoTextBox.TextBox TextBox10 
      Height          =   315
      Left            =   510
      TabIndex        =   9
      Top             =   4950
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      BackColor_OnGotFocus=   16761024
      Appearance_OnGotFocus=   1
      Alignment       =   2
      FontSize        =   8,25
      ForeColor       =   -2147483640
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "See the color and/or Appareance change when the control Got the focus."
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1050
      TabIndex        =   12
      Top             =   210
      Width           =   2595
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To move Up:  left or up arrows"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2400
      TabIndex        =   11
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To move Down: press [Enter], right or down arrow"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   540
      TabIndex        =   10
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "See the color and/or Appareance change when the control Got the focus."
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1080
      TabIndex        =   15
      Top             =   240
      Width           =   2595
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To move Down: press [Enter], right or down arrow"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2430
      TabIndex        =   14
      Top             =   930
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To move Down: press [Enter], right or down arrow"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   570
      TabIndex        =   13
      Top             =   930
      Width           =   1695
   End
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

