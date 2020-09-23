VERSION 5.00
Begin VB.Form frmWindow 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "2ND"
   ClientHeight    =   1935
   ClientLeft      =   1590
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form2"
   ScaleHeight     =   1935
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1320
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
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmWindow"
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurrX = X
    CurrY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DragWindow(Button, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurrX = X
    CurrY = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DragWindow(Button, X, Y)
End Sub

Private Function DragWindow(Button As Integer, X As Single, Y As Single)
If Button = 1 Then
    Me.Left = Me.Left + (X - CurrX)
    Me.Top = Me.Top + (Y - CurrY)
End If
End Function

