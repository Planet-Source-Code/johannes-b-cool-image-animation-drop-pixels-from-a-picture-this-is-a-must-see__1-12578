VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COOL ANIMATION"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Drop pixles"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4260
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   280
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   374
      TabIndex        =   0
      Top             =   720
      Width           =   5670
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim A, B As Integer

Dim BC
Private Sub Command1_Click()
Picture1.FontSize = "8"
Picture1.FontName = "Arial"
Picture1.Cls
Picture1.Print ""
Picture1.Print ""

Picture1.ForeColor = vbBlack
Picture1.Print "      THIS COOL ANIMATION WAS MADE BY JOHANNES BOHMAN"
Picture1.Print ""
Picture1.Print ""
Picture1.Print ""
Picture1.Print ""
Picture1.Print ""
Picture1.FontSize = "35"
Picture1.ForeColor = vbRed
Picture1.Print " PLEASE VOTE!"

A = Picture1.ScaleWidth
B = Picture1.ScaleHeight


Do

A = A + 1

If A >= Picture1.ScaleWidth Then
A = 0
B = B - 1
Picture1.Refresh
End If

BC = Picture1.Point(A, B)

Picture1.ForeColor = BC

Picture1.Line (A, B)-(A, Picture1.ScaleHeight)


Loop Until B = 0
Picture1.Refresh
End Sub


