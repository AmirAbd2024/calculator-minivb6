VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator_VB6"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command23 
      Caption         =   "><"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3540
      TabIndex        =   23
      Top             =   3375
      Width           =   555
   End
   Begin VB.CommandButton Command22 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   3540
      TabIndex        =   22
      Top             =   1680
      Width           =   555
   End
   Begin VB.CommandButton Command20 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3540
      TabIndex        =   21
      Top             =   2820
      Width           =   555
   End
   Begin VB.CommandButton Command19 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2850
      TabIndex        =   20
      Top             =   3375
      Width           =   555
   End
   Begin VB.CommandButton Command18 
      Caption         =   "CLR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2850
      TabIndex        =   19
      Top             =   1680
      Width           =   555
   End
   Begin VB.CommandButton Command17 
      Caption         =   "pow"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2850
      TabIndex        =   18
      Top             =   2250
      Width           =   555
   End
   Begin VB.CommandButton Command16 
      Caption         =   "sqr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2850
      TabIndex        =   17
      Top             =   2835
      Width           =   555
   End
   Begin VB.CommandButton Command15 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   16
      Top             =   3375
      Width           =   555
   End
   Begin VB.CommandButton Command14 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   15
      Top             =   1680
      Width           =   555
   End
   Begin VB.CommandButton Command13 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   14
      Top             =   2250
      Width           =   555
   End
   Begin VB.CommandButton Command11 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   13
      Top             =   2820
      Width           =   555
   End
   Begin VB.CommandButton Command12 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2175
      TabIndex        =   12
      Top             =   3375
      Width           =   555
   End
   Begin VB.CommandButton Command10 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   870
      TabIndex        =   11
      Top             =   3375
      Width           =   555
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000008&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2175
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   10
      Top             =   1680
      Width           =   555
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000008&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1515
      TabIndex        =   9
      Top             =   1680
      Width           =   555
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000008&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   870
      TabIndex        =   8
      Top             =   1680
      Width           =   555
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000008&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2175
      TabIndex        =   7
      Top             =   2250
      Width           =   555
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000008&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1515
      TabIndex        =   6
      Top             =   2250
      Width           =   555
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000008&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   870
      TabIndex        =   5
      Top             =   2250
      Width           =   555
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000008&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2175
      TabIndex        =   4
      Top             =   2820
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000008&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1515
      TabIndex        =   3
      Top             =   2820
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000008&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   870
      TabIndex        =   2
      Top             =   2820
      Width           =   555
   End
   Begin VB.CommandButton Command0 
      BackColor       =   &H80000008&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1515
      TabIndex        =   1
      Top             =   3375
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   165
      TabIndex        =   0
      Text            =   "0123456789"
      Top             =   570
      Width           =   3915
   End
   Begin VB.Label Labelsds 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   ""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   165
      TabIndex        =   25
      Top             =   3855
      Width           =   690
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   210
      TabIndex        =   24
      Top             =   135
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim d1 As Double
Dim d2 As Double
Dim cleard As Boolean
Dim pointa As Boolean


Dim tempkey As Byte
Dim tempstr As String
Dim oprator As String

Private Sub Command1_Click()
getBUTT "1"
End Sub

Private Sub Command10_Click()
getBUTT "."
End Sub

Private Sub Command11_Click()
getBUTT "*"
End Sub

Private Sub Command13_Click()
getBUTT "-"
End Sub

Private Sub Command14_Click()
getBUTT "+"
End Sub

Private Sub Command15_Click()
getBUTT "/"
End Sub

Private Sub Command16_Click()
getBUTT "sqr"
End Sub

Private Sub Command17_Click()
getBUTT "^"
End Sub

Private Sub Command18_Click()
Form_Load
End Sub

Private Sub Command19_Click()
getBUTT "="
End Sub

Private Sub Command2_Click()
getBUTT "2"
End Sub

Private Sub Command22_Click()
If Text1.Text <> "" Then
Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
SendKeys "{END}"
End If
End Sub

Private Sub Command3_Click()
getBUTT "3"
End Sub

Private Sub Command4_Click()
getBUTT "4"
End Sub

Private Sub Command5_Click()
getBUTT "5"
End Sub

Private Sub Command6_Click()
getBUTT "6"
End Sub

Private Sub Command7_Click()
getBUTT "7"
End Sub

Private Sub Command8_Click()
getBUTT "8"
End Sub

Private Sub Command9_Click()
getBUTT "9"
End Sub

Private Sub Command0_Click()
getBUTT "0"
End Sub

Private Sub Command12_Click()
getBUTT "00"
End Sub

Private Sub Form_Load()
Label0.Caption = ""
Text1.Text = ""
tempkey = 0
tempstr = ""
d1 = 0
d2 = 0
cleard = True
pointa = True
End Sub


Private Function getBUTT(ByVal s As String)
tempkey = 0
If Text1.Text = "Cannot divide by zero" Then
Text1.Text = ""
End If

If s = "1" Then
Text1.Text = Text1.Text + "1"
ElseIf s = "2" Then
Text1.Text = Text1.Text + "2"
ElseIf s = "3" Then
Text1.Text = Text1.Text + "3"
ElseIf s = "4" Then
Text1.Text = Text1.Text + "4"
ElseIf s = "5" Then
Text1.Text = Text1.Text + "5"
ElseIf s = "6" Then
Text1.Text = Text1.Text + "6"
ElseIf s = "7" Then
Text1.Text = Text1.Text + "7"
ElseIf s = "8" Then
Text1.Text = Text1.Text + "8"
ElseIf s = "9" Then
Text1.Text = Text1.Text + "9"
ElseIf s = "0" And Text1.Text <> "" Then
Text1.Text = Text1.Text + "0"
ElseIf s = "00" And Text1.Text <> "" Then
Text1.Text = Text1.Text + "00"

ElseIf s = "." Then
If checkPoint() <> False And Text1.Text = "" Then
Text1.Text = "0."
ElseIf checkPoint() <> False And Text1.Text <> "" Then
Text1.Text = Text1.Text + "."
End If
pointa = False

ElseIf s = "+" Then
If cleard = True Then
d1 = 0
cleard = False
End If
oprator = "+"

If Text1.Text <> "" Then
d1 = d1 + Val(Text1.Text)
End If
Label0.Caption = Str(d1) + " +"
Text1.Text = ""

ElseIf s = "-" Then

oprator = "-"
If Text1.Text <> "" And Text1.Text <> "-" Then


If cleard = True Then
d1 = Val(Text1.Text)
cleard = False
Else
d1 = d1 - Val(Text1.Text)
End If

Else
Text1.Text = "-"
End If
Label0.Caption = Str(d1) + " -"
Text1.Text = ""


ElseIf s = "*" Then
If cleard = True Then
d1 = 1
cleard = False
End If

oprator = "*"
If Text1.Text <> "" Then
d1 = d1 * Val(Text1.Text): End If
Label0.Caption = Str(d1) + " x"
Text1.Text = ""

ElseIf s = "/" Then

oprator = "/"
If Text1.Text <> "" Then

If cleard = True Then
d1 = Val(Text1.Text)
cleard = False
Else
d1 = d1 / Val(Text1.Text)
End If

End If
Label0.Caption = Str(d1) + " /"
Text1.Text = ""


ElseIf s = "^" Then
oprator = "^"

If Text1.Text <> "" Then

If cleard = True Then
d1 = Val(Text1.Text)
cleard = False
Else
d1 = d1 ^ Val(Text1.Text)
End If

Label0.Caption = Str(d1) + " ^"
Text1.Text = ""

End If


ElseIf s = "sqr" Then
oprator = "sqr"

If Text1.Text <> "" Then

d1 = Sqr(Val(Text1.Text))

Label0.Caption = "ans ="
Text1.Text = Str(d1)

End If



ElseIf s = "=" And Text1.Text <> "" Then
Label0.Caption = "ans ="
d2 = Val(Text1.Text)
If oprator = "+" Then
Text1.Text = Str(d1 + d2)
ElseIf oprator = "-" Then
Text1.Text = Str(d1 - d2)
ElseIf oprator = "*" Then
Text1.Text = Str(d1 * d2)
ElseIf oprator = "/" Then
If d2 <> 0 Then
Text1.Text = Str(d1 / d2)
Else
Text1.Text = "Cannot divide by zero"
End If
ElseIf oprator = "^" Then
Text1.Text = Str(d1 ^ d2)
End If
If oprator = "+" Or oprator = "-" Then
d1 = 0
ElseIf oprator = "*" Or oprator = "/" Or oprator = "^" Then
d1 = 1
End If
'd1 = Val(Text1.Text)
cleard = True



End If


End Function

Private Sub Text1_Change()
Dim x As Long
Dim res As String
x = 1
res = ""
While x <= Len(Text1.Text)
If Mid(Text1.Text, x, 1) <> " " Then
res = res + Mid(Text1.Text, x, 1)
End If
x = x + 1
Wend
Text1.Text = res
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'MsgBox (keyASCII)
If Text1.Text = "Cannot divide by zero" Then
Text1.Text = ""
End If

tempkey = KeyAscii
tempstr = Text1.Text
If tempkey = 46 Then
KeyAscii = 0
getBUTT (".")
SendKeys "{END}"
ElseIf tempkey = 8 Then
KeyAscii = 0
Command22_Click
ElseIf tempkey = 43 Then
KeyAscii = 0
tempstr = "+"
getBUTT ("+")
ElseIf tempkey = 45 Then
KeyAscii = 0
tempstr = "-"
getBUTT ("-")
SendKeys "{END}"
ElseIf tempkey = 42 Then
KeyAscii = 0
tempstr = "8"
getBUTT ("*")
ElseIf tempkey = 47 Then
KeyAscii = 0
tempstr = "/"
getBUTT ("/")
ElseIf tempkey = 13 Then
KeyAscii = 0
getBUTT ("=")
SendKeys "{END}"
ElseIf tempkey = 27 Then
KeyAscii = 0
Form_Load
ElseIf tempkey >= 48 And tempkey <= 57 Then
'Skip blocking
ElseIf tempkey = 46 And tempkey = 0 And tempkey <> 32 Then
'Skip blocking
Else
KeyAscii = 0

End If
End Sub



Private Function allowed() As Boolean

If tempkey = 46 Or tempkey = 0 Or (tempkey >= 48 And tempkey <= 57) Then
allowed = True
Else
allowed = False
End If

End Function


Private Function checkPoint() As Boolean
Dim x As Byte
Dim res As Boolean
x = 1
res = True

If Text1.Text = "" And tempkey <> 46 Then
Text1.Text = "0."
checkPoint = False
End If

While x <= Len(Text1.Text) And res <> False

If Mid(Text1.Text, x, 1) = "." Then
res = False
End If

x = x + 1
Wend

checkPoint = res

End Function







