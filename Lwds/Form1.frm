VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "写字符"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4665
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1200
      TabIndex        =   2
      Text            =   "A,Z"
      Top             =   600
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "编写域："
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   660
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GetC(sValue As String) As String
If sValue = "" Then Exit Function
Picture1.Cls
Picture1.Print sValue
Dim X%, Y%
For X = 0 To Picture1.Width / 15
For Y = 0 To Picture1.Height / 15
If Picture1.Point(X * 15, Y * 15) = 0 Then
GetC = GetC & X & "," & Y & "|" '以缩小15
End If
Next Y
Next X
If GetC <> "" Then GetC = Left(GetC, Len(GetC) - 1)
End Function
Function BeS(X As String) As String
Dim A() As String, I%, SX As Double, SY As Double, x2 As Double, y2 As Double
A = Split(X, "|")
SX = 999999999
SY = 999999999
For I = 0 To UBound(A)
X = Val(Split(A(I), ",")(0))
Y = Val(Split(A(I), ",")(1))
 If X < SX Then SX = X
 If Y < SY Then SY = Y
Next I
For I = 0 To UBound(A)
x2 = Val(Split(A(I), ",")(0)) - SX
y2 = Val(Split(A(I), ",")(1)) - SY
A(I) = x2 & "," & y2
BeS = BeS & A(I) & "|"
Next I
If BeS <> "" Then BeS = Left(BeS, Len(BeS) - 1)
End Function
Function WH(X As String) As String
Dim I%, x2 As Double, y2 As Double
Dim Bw As Double, Bh As Double
Dim BSx As Double, BBx As Double, BSy As Double, BBy As Double
Dim B() As String
B = Split(X, "|")

For I = 0 To UBound(B)
 x2 = Split(B(I), ",")(0)
 y2 = Split(B(I), ",")(1)
 If x2 < BSx Then BSx = x2
 If x2 > BBx Then BBx = x2
 If y2 < BSy Then BSy = y2
 If y2 > BBy Then BBy = y2
Next I
Bw = BBx - BSx
Bh = BBy - BSy
 WH = "!" & Bw & "~" & Bh
End Function

Private Sub Command1_Click()
Dim I As Integer
Dim A%, B%, l$, Wg$, X$
Dim C%
If Text1 <> "" Then
A = Asc(Split(Text1.Text, ",")(0))
B = Asc(Split(Text1.Text, ",")(1))
Wg = BeS(GetC("?"))
If A > B Then
C = A
A = B
B = C
End If
Open "WDS.txt" For Output As #1
For I = A To B
l = BeS(GetC(Chr(I)))
If l <> "" Then
If l <> Wg Or (l = Wg And I = 63) Then
 If I = B Then Print #1, Chr(I) & "*" & l & WH(l); Else Print #1, Chr(I) & "*" & l & WH(l)
 End If
End If
Next I
Close #1
ElseIf Text2 <> "" Then

Open "WDS.txt" For Output As #1
For I = 1 To Len(Text2.Text)
X = Mid(Text2, I, 1)

l = BeS(GetC(X))
If l <> "" Then
If l <> Wg Or (l = Wg And I = 63) Then
 If I = Len(Text2.Text) Then Print #1, X & "*" & l & WH(l); Else Print #1, X & "*" & l & WH(l)
 End If
End If
Next I

Close #1



End If
MsgBox "完成！"
End
End Sub


