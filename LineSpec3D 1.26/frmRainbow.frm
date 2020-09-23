VERSION 5.00
Begin VB.Form frmRainbow 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rainbow Chooser based upon Sunero Four Colour Gradient"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   7
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.PictureBox picFive 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   6720
      ScaleHeight     =   271
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox picTwo 
      AutoRedraw      =   -1  'True
      Height          =   4095
      Left            =   3600
      ScaleHeight     =   4035
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   360
      Width           =   3075
   End
   Begin VB.PictureBox picOne 
      AutoRedraw      =   -1  'True
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblInstruction 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label lblStyle1 
      Caption         =   "Design by Sunero Technologies"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   2355
   End
   Begin VB.Label lblHue 
      Caption         =   "Normal"
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmRainbow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim LngAlpha As Long
    Dim blnMouseDown As Boolean
    Dim BGRTemp As Long
    Dim BGRTmp2 As Long
    Dim r2F As Byte, r2 As Long
    Dim g2F As Byte, g2 As Long
    Dim b2F As Byte, b2 As Long
    Dim xf As Long, yf As Long

Private Sub Form_Load()

    picOne.ScaleMode = 3 'vbPixels
    picTwo.ScaleMode = 3
    picFive.ScaleMode = 3

    lblInstruction.Font = 1
    'lblInstruction.Caption = " Use bar to change Value (HSV)"

    '' Draw Three Colour Linear Gradient
    DrawGradient picFive.hdc, 0, 0, picFive.ScaleWidth, picFive.ScaleHeight / 2, vbWhite, vbWhite, BGR2, BGR2
    DrawGradient picFive.hdc, 0, picFive.ScaleHeight / 2, picFive.ScaleWidth, picFive.ScaleHeight / 2, BGR2, BGR2, vbBlack, vbBlack
    
    LngAlpha = vbWhite
    
    picFiveDownMoveCommon
    
End Sub
Private Sub picFiveDownMoveCommon()
    
    'Hue Whiteness Model
    DrawGradient picTwo.hdc, 10, 10, 30, 250, vbRed, vbYellow, LngAlpha, LngAlpha
    DrawGradient picTwo.hdc, 40, 10, 30, 250, vbYellow, vbGreen, LngAlpha, LngAlpha
    DrawGradient picTwo.hdc, 70, 10, 30, 250, vbGreen, vbCyan, LngAlpha, LngAlpha
    DrawGradient picTwo.hdc, 100, 10, 30, 250, vbCyan, vbBlue, LngAlpha, LngAlpha
    DrawGradient picTwo.hdc, 130, 10, 30, 250, vbBlue, vbMagenta, LngAlpha, LngAlpha
    DrawGradient picTwo.hdc, 160, 10, 30, 250, vbMagenta, vbRed, LngAlpha, LngAlpha
    
    ' Topleft
    DrawGradient picOne.hdc, 10, 10, 100, 100, vbRed, vbYellow, LngAlpha, LngAlpha
    'BottomRight
    DrawGradient picOne.hdc, 110, 110, 100, 100, LngAlpha, vbCyan, vbMagenta, vbBlue
    'Top Right
    DrawGradient picOne.hdc, 110, 10, 100, 100, vbYellow, vbGreen, LngAlpha, vbCyan
    'Bottom Left
    DrawGradient picOne.hdc, 10, 110, 100, 100, LngAlpha, LngAlpha, vbRed, vbMagenta
        
    picTwo.Refresh
    picOne.Refresh
    
    ExtractRGBComponents
    
    Caption = BGRTemp
        
End Sub
Private Sub OneTwoDownMoveCommon()
    If BGRTemp <> -1 Then
        Caption = BGRTemp
        DrawGradient picFive.hdc, 0, 0, picFive.ScaleWidth, picFive.ScaleHeight / 2, vbWhite, vbWhite, BGRTemp, BGRTemp
        DrawGradient picFive.hdc, 0, picFive.ScaleHeight / 2, picFive.ScaleWidth, picFive.ScaleHeight / 2, BGRTemp, BGRTemp, vbBlack, vbBlack
        picFive.Refresh
    End If
End Sub
Private Sub ExtractRGBComponents()
    b2& = ((BGR2& And &HFF0000) / &H10000) And &HFF
    g2& = ((BGR2& And &HFF00) / &H100) And &HFF
    r2& = BGR2& And &HFF
End Sub
Private Sub picFive_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnMouseDown = False
End Sub
Private Sub picFive_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnMouseDown = True
    Call picFive_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub picFive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnMouseDown Then
        BGRTemp = picFive.Point(2, Y)
        ExtractRGBComponents
        If BGRTemp <> -1 Then
            If Y > 0 And Y < 255 Then
                Y = 255 - Y
            ElseIf Y < 1 Then
                Y = 255
            Else
                Y = 0
            End If
            xf = X: yf = Y
            LngAlpha = RGB(Y, Y, Y)
            picFiveDownMoveCommon
        End If
    End If
End Sub
Private Sub picTwo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnMouseDown = True
    Call picTwo_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub picOne_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnMouseDown = True
    Call picOne_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub picTwo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnMouseDown Then
        xf = X: yf = Y
        BGRTemp = picTwo.Point(X, Y)
        ExtractRGBComponents
        OneTwoDownMoveCommon
    End If
End Sub
Private Sub picOne_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnMouseDown Then
        xf = X: yf = Y
        BGRTemp = picOne.Point(X, Y)
        ExtractRGBComponents
        OneTwoDownMoveCommon
    End If
End Sub
Private Sub picTwo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnMouseDown Then blnMouseDown = False
End Sub
Private Sub picOne_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnMouseDown Then blnMouseDown = False
End Sub
Private Sub cmdOK_Click()
    If BGRTemp <> -1 Then
        BGR2 = BGRTemp
        Unload Me
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
