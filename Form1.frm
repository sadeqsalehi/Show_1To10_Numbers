VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00292929&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print: Numbers 1 To 10"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13320
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   21.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0062C5FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   13320
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Â‰—” «‰ ⁄·«„Â Ã⁄›—Ì"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   3750
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   6885
      Width           =   5820
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Â‰—¬„Ê“:’«·ÕÌ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   5430
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   7215
      Width           =   2460
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0062C5FF&
      Height          =   450
      Left            =   11235
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1410
      Width           =   2085
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0062C5FF&
      Height          =   450
      Left            =   11235
      MouseIcon       =   "Form1.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   855
      Width           =   2085
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   11235
      Picture         =   "Form1.frx":02A4
      Top             =   840
      Width           =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000D44EA&
      Caption         =   "»—‰«„Â «Ì »‰ÊÌ”Ìœ òÂ «⁄œ«œ 1  « 10 —« ‰„«Ì‘ œÂœ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   13470
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   11235
      Picture         =   "Form1.frx":33A3
      Top             =   1395
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MouseOver As Boolean
Private Sub Command1_Click()
Dim I As Integer
Me.Cls
Print
For I = 1 To 10
 Print I;
Next I
End Sub
Private Sub Command2_Click()
End
End Sub

'òœÂ«Ì “Ì— «Œ Ì«—Ì »ÊœÂ Ê »—«Ì  €ÌÌ— —‰ê œò„Â »ò«— „Ì —Ê‰œ
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseOver = True Then
Image1.Picture = LoadPicture(App.Path + "\green.jpg")
Image2.Picture = LoadPicture(App.Path + "\green.jpg")
Command1.ForeColor = &H62C5FF
Command2.ForeColor = &H62C5FF
MouseOver = False
End If
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseOver = False Then
Image1.Picture = LoadPicture(App.Path + "\red.jpg")
Command1.ForeColor = &H3A48F8
MouseOver = True
End If
End Sub
Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseOver = False Then
Image2.Picture = LoadPicture(App.Path + "\red.jpg")
Command2.ForeColor = &H3A48F8

MouseOver = True
End If
End Sub

