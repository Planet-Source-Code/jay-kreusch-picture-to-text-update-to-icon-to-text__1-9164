VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About....."
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   405
      Left            =   1140
      TabIndex        =   1
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Modified by Jay Kreusch"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ver 2.0"
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by DreamVb"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   195
      TabIndex        =   2
      Top             =   720
      Width           =   2880
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2295
      Picture         =   "Form2.frx":0000
      Top             =   165
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon to Text"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1005
      TabIndex        =   0
      Top             =   300
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "Form2.frx":030A
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

