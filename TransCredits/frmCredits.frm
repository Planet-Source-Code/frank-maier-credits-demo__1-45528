VERSION 5.00
Begin VB.Form frmCredits 
   AutoRedraw      =   -1  'True
   Caption         =   "Credits"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picCreditsTemp 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   720
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'Stellt dar wie man ein About Formular machen könnte.
'Source von Frank Maier.
'Bei Anregungen: FrankMr@gmx.de
'****************************************************
'An example for an About formular
'Source from Frank Maier
'Questions to FrankMr@gmx.de
'****************************************************

Option Explicit

'Farbverlauf
'Gradient
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long  'UNSIGNED Long
    LowerRight As Long 'UNSIGNED Long
End Type

'Tranparenz-Funktion
'Transparent
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BL As Long) As Long
Private Type BLENDTYPE
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Private Const AC_SRC_OVER = &H0

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim ExitCredits As Boolean

Private Sub Form_Activate()
Call Credits
End Sub

Private Sub Form_Resize()
'Erstellt Farbverlauf
'Creates gradient
Call FarbverlaufAPI(Me, vbMagenta, vbBlack, Me.ScaleWidth, Me.ScaleHeight / 2)
Call FarbverlaufAPI(Me, vbBlack, vbRed, Me.ScaleWidth, Me.ScaleHeight, 0, Me.ScaleHeight / 2)
frmCredits.Picture = frmCredits.Image
End Sub

Private Sub Form_Unload(Cancel As Integer)
ExitCredits = True
'End
End Sub





'*********************************************************************************
'Zweck: Erstellt Farbverlauf, Gibt Text auf die PicBox aus
'Use: Creates gradient, draws the text on the PicBox's
'*********************************************************************************
Private Sub Credits()
Dim NumLines As Integer
Dim PositionProzent As Integer
Dim BT As BLENDTYPE
Dim zBT As Long
Dim txtX As Long
Dim txtY As Long
Dim txtWidth As Long
Dim txtHeight As Long
Dim Txt As String
Dim CreditsText(100) As String
Dim Zahl As Integer
Dim lX As Long
Dim lY As Long
Dim Trans As Long

CreditsText(1) = "Titel Ihres Programms"
CreditsText(2) = ""
CreditsText(3) = "****************************"
CreditsText(4) = ""
CreditsText(5) = "Programmierung:"
CreditsText(6) = "Name(n)"
CreditsText(7) = ""
CreditsText(8) = "Grafiken:"
CreditsText(9) = "Name(n)"
CreditsText(10) = ""
CreditsText(11) = "****************************"
CreditsText(12) = ""
CreditsText(13) = "Ein kleines Beispiel zum"
CreditsText(14) = ""
CreditsText(15) = "erstellen eines"
CreditsText(16) = ""
CreditsText(17) = "About Formulars"
NumLines = 17

With frmCredits

lX = .ScaleLeft
lY = .ScaleHeight

'Erstellt Farbverlauf
'Creates Gradient
Call FarbverlaufAPI(Me, vbMagenta, vbBlack, .ScaleWidth, .ScaleHeight / 2)
Call FarbverlaufAPI(Me, vbBlack, vbRed, .ScaleWidth, .ScaleHeight, 0, .ScaleHeight / 2)
.Picture = .Image


Do
    .Cls
    For Zahl = 1 To NumLines
        
        'Text speichern
        'Saves the text
        Txt = CreditsText(Zahl)
        
        'Größe des Textes erfassen
        'Get the width and height of the text
        txtWidth = .picCreditsTemp.TextWidth(Txt) + 10
        txtHeight = .picCreditsTemp.TextHeight(Txt) + 10
        
        'Position neu berechnen
        'Calculates the position
        txtX = Int((.ScaleWidth / 2) - (txtWidth / 2))
        txtY = Int(lY + (Zahl * .FontSize + (12 * Zahl)))  'Y-Position und Abstand
        
        ' Hintegrundzwischenspeicher an die Textgröße anpassen
        ' Resize the Backgroundpuffer to the width and height of the Text
        .picCreditsTemp.Width = txtWidth
        .picCreditsTemp.Height = txtHeight
        ' und Hintegrund sichern
        ' saves the Background
        BitBlt .picCreditsTemp.hdc, 0, 0, txtWidth, txtHeight, .hdc, txtX, txtY, vbSrcCopy
        
        ' Text ausgeben
        ' Print Text
        TextOut .picCreditsTemp.hdc, 0, 0, Txt, Len(Txt)

        
        'Falls der Text über die Hälfte ist, wird er unsichtbarer
        'If the Text is over the half, it's get his tranparent and becomes unvisible
        If txtY < .ScaleHeight / 2 Then
            PositionProzent = (txtY / (.ScaleHeight / 2)) * 255
            Trans = PositionProzent
            If PositionProzent > 255 Then
                Trans = 255
            ElseIf PositionProzent < 0 Then
                Trans = 0
            End If
        'sonst sichtbar
        'else visible
        Else
            Trans = 255
        End If
       
        ' und Type initialisieren
        With BT
            .BlendOp = AC_SRC_OVER
            .BlendFlags = 0
            .SourceConstantAlpha = Trans
            .AlphaFormat = 0
        End With
        
        ' BT in zBT kopieren (weil ByVal)
        ' copy BT to zBT
        RtlMoveMemory zBT, BT, 4
    
    
        ' und Grafik Transparent ausgeben
        ' makes graphics transparent
        AlphaBlend .hdc, txtX, txtY, txtWidth, txtHeight, .picCreditsTemp.hdc, 0, 0, txtWidth, txtHeight, zBT
    
        If Zahl = NumLines And txtY < -25 Then
            lY = .ScaleHeight
            Zahl = 0
        End If
    
    Next Zahl
    
    lY = lY - 1
    'Pause
    Call Wait(10)
    
    If ExitCredits Then Exit Do
Loop

End With
End Sub

'*********************************************************************************
'Zweck: Ermöglicht eine Pause
'Use:   Pause
'*********************************************************************************
Private Function Wait(ByVal Delay As Long)
Dim TickCount As Long

    TickCount = GetTickCount
    
    While (TickCount + Delay) > GetTickCount
        DoEvents
    Wend
    
End Function

' ======================================================================================
' -----------------------------Erstellt Farbverläufe über API---------------------------
' -------------------------------Creates gradient over API------------------------------
' ======================================================================================

Private Sub FarbverlaufAPI(Objekt As Object, StartFarbe As Long, _
                           EndFarbe As Long, EndX As Integer, _
                           EndY As Integer, _
                          Optional StartX As Integer = 0, _
                          Optional StartY As Integer = 0, _
                          Optional Richtung As Byte = 1)
'*** abColor
Dim abColor(0 To 3) As Byte
'*** audtGradientTrivertex
Dim audtGradientTrivertex(0 To 1) As TRIVERTEX
Dim gRect As GRADIENT_RECT

'*** Copy lBackcolor to abColor
Call RtlMoveMemory(abColor(0), StartFarbe, 4)
'*** Set audtGradientTrivertex(0)
audtGradientTrivertex(0).X = StartX
audtGradientTrivertex(0).Y = StartY
Call RtlMoveMemory(audtGradientTrivertex(0).Red, (CLng(abColor(0)) * CLng(256)), 2)
Call RtlMoveMemory(audtGradientTrivertex(0).Green, (CLng(abColor(1)) * CLng(256)), 2)
Call RtlMoveMemory(audtGradientTrivertex(0).Blue, (CLng(abColor(2)) * CLng(256)), 2)
audtGradientTrivertex(0).Alpha = 0&

'*** Copy lGradientBackColor to abColor
Call RtlMoveMemory(abColor(0), EndFarbe, 4)
'*** Set audtGradientTrivertex(1)
audtGradientTrivertex(1).X = EndX
audtGradientTrivertex(1).Y = EndY
Call RtlMoveMemory(audtGradientTrivertex(1).Red, (CLng(abColor(0)) * CLng(256)), 2)
Call RtlMoveMemory(audtGradientTrivertex(1).Green, (CLng(abColor(1)) * CLng(256)), 2)
Call RtlMoveMemory(audtGradientTrivertex(1).Blue, (CLng(abColor(2)) * CLng(256)), 2)
audtGradientTrivertex(1).Alpha = 0&

gRect.UpperLeft = 0
gRect.LowerRight = 1

Call GradientFillRect(Objekt.hdc, audtGradientTrivertex(0), 2, gRect, 1, Richtung)
End Sub

