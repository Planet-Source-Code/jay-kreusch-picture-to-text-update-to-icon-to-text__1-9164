VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Cool-Art 1.0"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpenRTF 
      Caption         =   "Open File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   6000
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar pbrMain 
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   5400
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin RichTextLib.RichTextBox txtMain 
      Height          =   4815
      Left            =   1800
      TabIndex        =   12
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8493
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   3.40282e38
      TextRTF         =   $"Form1.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   405
      Left            =   9600
      TabIndex        =   9
      Top             =   6000
      Width           =   1365
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1245
      Left            =   6240
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Text            =   "Ascii art by DreamVb"
      Top             =   6030
      Width           =   1935
   End
   Begin VB.TextBox txtDestination 
      Height          =   300
      Left            =   870
      TabIndex        =   5
      Top             =   5520
      Width           =   3825
   End
   Begin VB.TextBox txtSource 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   5085
      Width           =   3825
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00E0E0E0&
      Caption         =   "....."
      Height          =   330
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5100
      Width           =   405
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convet and Save"
      Height          =   420
      Left            =   90
      TabIndex        =   0
      Top             =   6000
      Width           =   1470
   End
   Begin VB.Label lblLegend 
      Caption         =   "Legend:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblProgress 
      Caption         =   "Conversion Progress"
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblCaption 
      Caption         =   "Label4"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Label4"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Picture Title"
      Height          =   195
      Left            =   1830
      TabIndex        =   6
      Top             =   6075
      Width           =   840
   End
   Begin VB.Label lblSave 
      AutoSize        =   -1  'True
      Caption         =   "Save To:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   5580
      Width           =   660
   End
   Begin VB.Label lblPicture 
      AutoSize        =   -1  'True
      Caption         =   "Picture"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strMain As Variant      'Variant used because string is limited in size


Private Sub cmdConvert_Click()
    On Error GoTo errtrap
    Dim X, X1 As Integer        'X coordinate, X counter
    Dim Y, Y1 As Integer        'Y coordinate, y counter
    Dim jColor As Long          'Color Variable
    Dim lngColors() As Long     'Colors Array
    Dim intX As Long            'Loop Counter
    Dim Repeat As Boolean       'Indicator
    
    'Make sure There is a picture
    If Len(txtSource) = 0 Then
        MsgBox "You must select a picture"
        Exit Sub
    End If
    
    'Large picture warning
    MsgBox "Large pictures may not display properly. Please open in an RTF viewer and adjust page size and margins to view properly."
    
    'Give the colors palate space
    ReDim lngColors(0)
    
    'set up coordinate system
    Y = picMain.ScaleWidth - 1
    X = picMain.ScaleHeight - 1
           
    'Set Upper value of Progress Bar
    pbrMain.Max = X
    lblProgress.Caption = "Converting to text..."
    
    
    For X1 = 1 To X Step 1 'Loop Through Rows
        For Y1 = 1 To Y Step 1 'Loop Through Columns
            'Free the processor
            DoEvents
            
            'Set the current pixel color
            jColor = picMain.Point(Y1, X1)
            
            'If the color array is new, set the first color
            If UBound(lngColors()) = 0 Then
                lngColors(0) = jColor
                'Add space for the next color, will always be 0 for black (Logic causes skipping of 1)
                ReDim Preserve lngColors(1)
                'Start at chr(33) Less than 33 values include odd characters like line feeds
                strMain = strMain & Chr(33)
            Else
                'Assume the next color is not a repeat of the previous
                Repeat = False
                'but loop through the colors array to make sure
                For intX = 0 To UBound(lngColors())
                    DoEvents
                    'if you find a match
                    If lngColors(intX) = jColor Then
                        'add it to the string
                        strMain = strMain & Chr(33 + intX)
                        'disprove your assumption of non repeat
                        Repeat = True
                    End If
                Next
                'if you haven't been disproved, new color is used
                If Repeat = False Then
                    'make room for it
                    ReDim Preserve lngColors(UBound(lngColors()) + 1)
                    'set its value
                    lngColors(UBound(lngColors())) = jColor
                    'add it to the string
                    strMain = strMain & Chr(33 + UBound(lngColors()))
                End If
            End If
        Next
        'give the string a new row
        strMain = strMain & vbCrLf
        'increase the progressbar up to, but not beyond the max
        If pbrMain.Value <> pbrMain.Max Then
            pbrMain.Value = pbrMain.Value + 1
        End If
    Next
    
    'add the title to the string
    strMain = strMain & vbNewLine & txtTitle & vbCrLf & vbCrLf
    'add the color chart label
    strMain = strMain & "COLOR CHART:" & vbCrLf & vbCrLf

    'for each color in the palate
    For intX = 0 To UBound(lngColors())
        'create a new color palate indicator and legend guide
        Load lbl(intX + 1)
        Load lblCaption(intX + 1)
        lbl(intX + 1).Caption = ""
        'color the indicator
        lbl(intX + 1).BackColor = lngColors(intX)
        lbl(intX + 1).Visible = True
        'set the legend guide
        lblCaption(intX + 1).Caption = " = " & Chr(33 + intX)
        'position the legend items
        Select Case intX
            Case Is <= 14
                lblCaption(intX + 1).Move 350, 400 + (300 * intX), 500, 290
                lbl(intX + 1).Move 20, 400 + (300 * intX), 290, 290
            Case Is <= 27
                lblCaption(intX + 1).Move 1200, 400 + (300 * (intX - 15)), 500, 290
                lbl(intX + 1).Move 900, 400 + (300 * (intX - 15)), 290, 290
            Case Else
                lblCaption(intX + 1).Move 900, 400 + (300 * 14), 500, 290
                lblCaption(intX + 1).Caption = "More..."
                lbl(intX + 1).Move 900, 400 + (300 * 14), 290, 290
                lbl(intX + 1).BackColor = Me.BackColor
        End Select
        lblCaption(intX + 1).Visible = True
        'add the legend item to the string
        strMain = strMain & Chr(33 + intX) & " = " & lngColors(intX) & vbCrLf
    Next
    'code places the label to the front in case of more than 27 different colors
    lblCaption(intX).ZOrder
    
    'set the richtextbox's text property to the string
    txtMain.Text = strMain
    
    'code loops through the richtextbox and changes the colors of the symbols
    'TO ELIMINATE THE PAINTING.. COMMENT OUT THE FOLLOWING SECTION:
    lblProgress.Caption = "Painting Colors..."
    pbrMain.Value = pbrMain.Min
    pbrMain.Max = InStr(txtMain.Text, txtTitle)

    For intX = 0 To InStr(txtMain.Text, txtTitle)
        DoEvents
        txtMain.SelStart = intX
        txtMain.SelLength = 1
        If txtMain.SelText <> "" Then
            If (Asc(txtMain.SelText) - 33 <= UBound(lngColors())) And (Asc(txtMain.SelText) - 33 >= 0) Then
                txtMain.SelColor = lngColors(Asc(txtMain.SelText) - 33)
            End If
        End If

        DoEvents
        If pbrMain.Value <> pbrMain.Max Then
            pbrMain.Value = pbrMain.Value + 1
        End If
    Next
    'END SECTION
    
    lblProgress = "Complete!"
    SavePictureText txtDestination
    MsgBox "Your file has been saved to: " & txtDestination.Text
    cmdOpenRTF.Enabled = True
    
Exit Sub
errtrap:
Select Case Err.Number
    Case 5 'invalid proceedure
        If UBound(lngColors()) = 223 Then
            MsgBox "Your image has too many colors to convert, you must limit the palate to 223 colors"
        End If
    Case Else
        MsgBox Err.Number & ": " & Err.Description
End Select
End Sub

Private Sub cmdFind_Click()
    txtSource = Main.OpenFile
    picMain.Picture = LoadPicture(txtSource)
    If Len(txtSource) = 0 Then
        txtDestination = ""
    Else
        txtDestination = Left(txtSource, Len(txtSource) - 3) + "rtf"
    End If
   
End Sub

Sub SavePictureText(Filename As String)
    Open Filename For Output As #1
    Print #1, txtMain.TextRTF;
    Close #1
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdOpenRTF_Click()
    Call ShellExecute(0&, vbNullString, txtDestination, vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


