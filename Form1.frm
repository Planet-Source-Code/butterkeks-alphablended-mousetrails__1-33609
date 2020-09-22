VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Alphablended Mousetrails"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   540
      Index           =   0
      Left            =   3240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Private Type TrailType
    X As Long
    Y As Long
    Alpha As Byte
    Frame As Integer
End Type
Const AC_SRC_OVER = &H0
Dim Trail(0 To 20) As TrailType
Dim MX
Dim MY
Dim BF As BLENDFUNCTION, lBF As Long

Private Sub Form_Load()
For i = UBound(Trail) To 0 Step -1
    With Trail(i)
        .Alpha = Int((100 / 255) * 20) * Invert(20, i)
    End With
Next i

Me.Show
Do
Me.Cls
Me.Print MX
Me.Print MY
For i = 0 To UBound(Trail)
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = Trail(i).Alpha
        .AlphaFormat = 0
    End With
    
    With Trail(i)
        If i = 0 Then
            .X = (.X - MX) / 2
            .Y = (.Y - MY) / 2
            .Frame = .Frame + 1
            If .Frame > 4 Then .Frame = 0
        End If
        If i <> 0 Then
            .X = (.X - Trail(i - 1).X) / 2
            .Y = (.Y - Trail(i - 1).Y) / 2
            .Frame = Trail(i - 1).Frame + 1
            If .Frame > 4 Then .Frame = 0
        End If
        RtlMoveMemory lBF, BF, 4
        AlphaBlend Me.hdc, .X, .Y, 32, 32, Picture2(0).hdc, 0, 0, Picture2(0).ScaleWidth, Picture2(0).ScaleHeight, lBF
    End With
Next i
DoEvents
Loop
End Sub

Function Invert(Max, Val)
If Val < Max / 2 Then
    Invert = Max - Val
End If
If Val > Max / 2 Then
    Invert = -(Val - Max)
End If
If Val = Max Then Invert = Val
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MX = X
MY = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

