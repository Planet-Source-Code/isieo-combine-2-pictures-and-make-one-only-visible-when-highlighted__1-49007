VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PICTURE HIDER!!"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5685
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   3840
      ScaleHeight     =   1275
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   2280
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdLoad2 
      Caption         =   "&Load The invisible Picture "
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton CmdLoad1 
      Caption         =   "&Load The Visible Picture "
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   2895
      Left            =   2880
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton CmdMix 
      Caption         =   "&Embed In!!!"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdMix_Click()
Dim X As Long
Dim y As Long
Dim R1 As Integer
Dim G1 As Integer
Dim B1 As Integer
Dim R2 As Integer
Dim G2 As Integer
Dim B2 As Integer
Dim Fileloc As String
cdg1.ShowSave 'shows where to save the new picture
Fileloc = cdg1.FileName 'load the file name into the string
If cdg1.CancelError Then Exit Sub ' when user press cancel then exit sub.
Picture3.Cls  'clear the picture box
Picture3.Height = Picture1.Height 'resize th picture box
Picture3.Width = Picture1.Width 'resize
Me.Enabled = False 'disable form
    For y = 0 To Picture1.ScaleHeight 'loop! until y is same as the height of the picture
        DoEvents
        For X = 0 To Picture1.ScaleWidth 'another loop but now its x.
                ' when the x loop finishes, x will become 0 again and y will add 1 which means,
                  'we are filling in the pixels line by line
'***###IMPOTRANT!!###*** ALL THIS WILL ONLY WORK IF THE PICTURE BOX IS SET AUTOREDRAW = TRUE
                ' *Below* Read the rgb value of the picture from the
                ReadColours Picture1.Point(X, y), R1, G1, B1 ' first picture
                ReadColours Picture2.Point(X, y), R2, G2, B2 ' second picture
            If X Mod 2 <> 0 Then ' this is to alternate the mix of the pictures in the x axis
                If y Mod 2 <> 0 Then ' this is to alternate the mix of the pictures in the y axis
                Picture3.PSet (X, y), RGB(R1, G1, B1) ' set the pixel with the selected colours
                Else
                Picture3.PSet (X, y), RGB(R2, G2, B2) ' set the pixel with the selected colours
                End If
            Else
                If y Mod 2 = 0 Then ' this is to alternate the mix of the pictures in the y axis
                Picture3.PSet (X, y), RGB(R1, G1, B1) ' set the pixel with the selected colours
                Else
                Picture3.PSet (X, y), RGB(R2, G2, B2) ' set the pixel with the selected colours
                End If
            End If
        Next X
        Picture3.Picture = Picture3.Image 'convert the image into the picture
        Image1.Picture = Picture3.Picture ' show the preview
        Me.Caption = "PICTURE HIDER!! " & Round((y / Picture1.ScaleHeight) * 100, 2) & "% Done" ' Display the percentage
    Next y
Me.Enabled = True ' renable the form
Me.Caption = "PICTURE HIDER!!" ' Change the form caption
Picture3.Picture = Picture3.Image ' to reasure that the image is converted to the picture
SavePicture Picture3.Image, Fileloc 'save the picture in to a file
CmdLoad1.Enabled = True
CmdMix.Enabled = False
End Sub

Private Sub ReadColours(Colour As Long, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
    Dim lgMix As Long
    lgMix = (Colour And 255)
    R = lgMix And 255
    lgMix = Int(Colour / 256)
    G = lgMix And 255
    lgMix = Int(Colour / 65536)
    B = lgMix And 255
' not sure how this work... all i know is that this is to saperate the rgb of colours
' I Learned this code from Mauricio Castelazo Gamboa
' you  can get the original code from
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=10616&lngWId=1
End Sub

Private Sub CmdLoad1_Click()
On Error GoTo Oho
cdg1.ShowOpen ' show the open dilog box
If cdg1.CancelError Then Exit Sub ' if cancel then exit sub
Picture1.Picture = LoadPicture(cdg1.FileName) 'loads the picture
Form_Resize 'resize the form
CmdLoad1.Enabled = False 'disable the commandbutton
CmdLoad2.Enabled = True 'enable the command button
Exit Sub
Oho:
Beep
MsgBox "Error! Not a valid picture or not supported format!"
End Sub

Private Sub CmdLoad2_Click()
On Error GoTo Oho
cdg1.ShowOpen ' show the open dilog box
If cdg1.CancelError Then Exit Sub ' if cancel then exit sub
Picture2.Picture = LoadPicture(cdg1.FileName) 'loads the picture
'**##IMPORTANT!!##** THIS WILL ONLY WORK WITH AUTOREDRAW = TRUE
Picture2.PaintPicture Picture2.Picture, 0, 0, Picture2.Width, Picture2.Height, 0, 0 'Resize the picture
Picture2.Picture = Picture2.Image 'loads the image into the picture
Form_Resize 'resize the form
CmdLoad2.Enabled = False 'disable the commandbutton
CmdMix.Enabled = True 'enable the command button
Exit Sub
Oho:
Beep
MsgBox "Error! Not a valid picture or not supported format!"
End Sub




Private Sub Form_Resize()
' calculation to resize the form
CmdLoad1.Left = (Picture1.Width / 2) + Picture1.Left - (CmdLoad1.Width / 2)
CmdLoad2.Left = (Picture2.Width / 2) + Picture2.Left - (CmdLoad2.Width / 2)
Picture2.Width = Picture1.Width
Picture2.Height = Picture1.Height
Picture2.Left = Picture1.Left + Picture1.Width + 100
Me.Width = Picture1.Width + Picture2.Width + 200
Me.Height = Picture1.Height + Picture1.Top + 500
CmdMix.Left = (Me.Width / 2) - (CmdMix.Width / 2)
End Sub

