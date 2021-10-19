VERSION 5.00
Begin VB.Form frmCool 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alpha Blend"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmCool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCool.frx":000C
   ScaleHeight     =   4530
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5340
      Left            =   120
      Picture         =   "frmCool.frx":AEEA
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   568
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   8580
   End
End
Attribute VB_Name = "frmCool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Option Explicit
 
 Private Declare Function AlphaBlend _
  Lib "msimg32" ( _
  ByVal hDestDC As Long, _
  ByVal x As Long, ByVal y As Long, _
  ByVal nWidth As Long, _
  ByVal nHeight As Long, _
  ByVal hSrcDC As Long, _
  ByVal xSrc As Long, _
  ByVal ySrc As Long, _
  ByVal widthSrc As Long, _
  ByVal heightSrc As Long, _
  ByVal dreamAKA As Long) _
  As Boolean 'only Windows 98 or Latter
 Dim Num As Byte, nN%, nBlend&
 
Private Sub Run_Blending()
 Num = CByte(255)
 nN = 5
Do
 DoEvents
 '***********************************************
  nBlend = vbBlue - CLng(Num) * (vbYellow + 1)
 'It's Magic Formula is
 'Alchemical Mixture of Elements of Gold & Sky
 'It's obtained by an almost mystical way
 '***********************************************
 Num = CByte(Num) - nN
 If Num = CByte(0) Then
   nN = -5
 ElseIf Num = CByte(255) Then
   nN = 5
 End If
 Me.Cls
 AlphaBlend Me.hDC, 0, 0, picSrc.ScaleWidth, picSrc.ScaleHeight, picSrc.hDC, 0, 0, picSrc.ScaleWidth, picSrc.ScaleHeight, nBlend
Loop
End Sub


Private Sub Form_Activate()
 Call Run_Blending
End Sub

Private Sub Form_Unload(Cancel As Integer)
 End ' STOP Do Loop
End Sub
