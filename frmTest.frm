VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTest 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   75
      Width           =   4515
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjAutoScroll As AutoScroll

Private Sub Form_Load()

    Set mobjAutoScroll = New AutoScroll
    
    mobjAutoScroll.Window = txtTest.hwnd

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If WindowState = vbMinimized Then Exit Sub

    With txtTest
        .Width = ScaleWidth - 2 * .Left
        .Height = ScaleHeight - 2 * .Top
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mobjAutoScroll = Nothing

End Sub
