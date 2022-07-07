VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   ScaleHeight     =   255
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Exposed = False
'Begin Code

Option Explicit

Private Sub Form_Activate()
Call FormPrimerPlano(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FormNoPrimerPlano(Me)
End Sub

Private Sub Timer1_Timer()
Me.Top = RenderTop * 9
Me.Left = RenderLeft * 15
End Sub

