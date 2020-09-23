VERSION 5.00
Begin VB.Form frmCaptionWindow 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "2ND"
   ClientHeight    =   1935
   ClientLeft      =   2820
   ClientTop       =   6225
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   ScaleHeight     =   1935
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Drag window anywhere around the Screen."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Image ImgClose 
      Height          =   300
      Left            =   5730
      Picture         =   "frmCaptionWindow.frx":0000
      Top             =   0
      Width           =   225
   End
   Begin VB.Image ImgMax 
      Height          =   300
      Left            =   5505
      Picture         =   "frmCaptionWindow.frx":0402
      Top             =   0
      Width           =   225
   End
   Begin VB.Image ImgMin 
      Height          =   300
      Left            =   5280
      Picture         =   "frmCaptionWindow.frx":0804
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "RSM Window 1.0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image ImgTitlebar 
      Height          =   390
      Left            =   0
      Picture         =   "frmCaptionWindow.frx":0C06
      Top             =   0
      Width           =   6000
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1545
      Left            =   20
      Top             =   380
      Width           =   5955
   End
End
Attribute VB_Name = "frmCaptionWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'----------------------------------------------------
' Project Name : Drag Window
' Author: Mohit Soam
' Email : mohit.soam@rediff.com
' Address: III/10,Hydel Colony
'          Veerbhadra
'          Rishikesh (U.A)
'          India 249202
'
' Date of Creation:  July 7,2005
' Last Update:  August 1,2005
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.
'
'----------------------------------------------------
'
Dim CurrX, CurrY As Single

Private Sub cmdAbout_Click()
frmAbout.Show
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub ImgTitlebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurrX = X
    CurrY = Y
End Sub
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurrX = X
    CurrY = Y
End Sub
Private Sub ImgTitlebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Me.Left = Me.Left + (X - CurrX)
    Me.Top = Me.Top + (Y - CurrY)
End If
End Sub
Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Me.Left = Me.Left + (X - CurrX)
    Me.Top = Me.Top + (Y - CurrY)
End If
End Sub
