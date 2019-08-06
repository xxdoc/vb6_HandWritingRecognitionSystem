VERSION 5.00
Begin VB.Form LkP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "手写辨别系统"
   ClientHeight    =   4185
   ClientLeft      =   6255
   ClientTop       =   6960
   ClientWidth     =   5265
   Icon            =   "LkP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5265
   Begin VB.TextBox Bbd 
      Height          =   270
      Left            =   840
      TabIndex        =   9
      Text            =   "3"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Cs 
      Caption         =   "清空"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "应用"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Dw 
      Height          =   270
      Left            =   600
      TabIndex        =   4
      Text            =   "3"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Gok 
      Caption         =   "确定"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.PictureBox Ptg 
      DrawWidth       =   4
      Height          =   2550
      Left            =   120
      ScaleHeight     =   2550
      ScaleMode       =   0  'User
      ScaleWidth      =   2550
      TabIndex        =   0
      Top             =   120
      Width           =   2550
   End
   Begin VB.Label Label2 
      Caption         =   "辨别度："
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3540
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "线宽："
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3180
      Width           =   735
   End
   Begin VB.Label Rt 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1050
   End
End
Attribute VB_Name = "LkP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ping As Boolean '手写中/或不
Dim Xl%, Yl% '上次的点
Dim Xz() As Integer, Yz() As Integer
Dim Rut As String
Dim ReadyP As Boolean
Dim BBD2 As Integer

Private Sub Command1_Click()
Ptg.DrawWidth = Val(Dw)
End Sub


Private Sub Command2_Click()

End Sub

Private Sub Cs_Click()
Ptg.Cls
Rut = ""
 ReDim Xz(1 To 1) As Integer
 ReDim Yz(1 To 1) As Integer
End Sub

Private Sub Gok_Click()
If ReadyP = True Then
Rut = BeS(Left(Rut, Len(Rut) - 1))
Dim Bz() As String
Dim Ru() As String
Dim I%, I2%
Dim Lxs As String, Lxs2 As String
Dim Zx() As Double
Dim Sll As Double
Dim Sok As String
Dim Lst2() As String
Dim Wnum As Integer
Dim kp As Integer, KP2$
Dim x2 As Double, y2 As Double
Dim Aw As Double, Ah As Double
Dim Bw As Double, Bh As Double
Dim ASx As Double, ABx As Double, ASy As Double, ABy As Double
Dim BK As String


kp = 0
' Open "G:\手写辨别系统" & "\wds\wds.txt" For Input As #1
Open App.Path + "\Wds\WDS.txt" For Input As #1
Do While Not EOF(1)
kp = kp + 1
Line Input #1, KP2
Loop
Close #1

Sll = 999999999
Ru = Split(Rut, "|")
For I = 0 To UBound(Ru)
 x2 = Split(Ru(I), ",")(0)
 y2 = Split(Ru(I), ",")(1)
 If x2 < ASx Then ASx = x2
 If x2 > ABx Then ABx = x2
 If y2 < ASy Then ASy = y2
 If y2 > ABy Then ABy = y2
Next I
Aw = ABx - ASx
Ah = ABy - ASy

ReDim Zx(1 To kp) As Double '26字母
ReDim Lst2(1 To kp) As String
'定义
List1.Clear '清除上次的
Open App.Path & "\wds\wds.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, Lxs
Wnum = Wnum + 1
BK = Right(Lxs, 3)
Bw = Val(Left(BK, 1))
Bh = Val(Right(BK, 1))

Bz = Split(Mid(Lxs, 3, Len(Lxs) - 6), "|") '标准的变为数组
Bz = Split(BeS2(Ru, Bz, Aw, Ah, Bw, Bh), "|")

'****************************
 For I = 0 To UBound(Ru)
 Zx(Wnum) = Zx(Wnum) + WIsn(Ru(I), Bz)  '累积算法
   For I2 = 0 To UBound(Bz)
   Zx(Wnum) = Zx(Wnum) + WIsn(Bz(I2), Ru)
   Next I2
 Next I
 Zx(Wnum) = Zx(Wnum) / (UBound(Ru) + UBound(Bz) + 2)
'****************************
Lst2(Wnum) = Left(Lxs, 1) & "：" & Zx(Wnum)
If Zx(Wnum) < Sll Then Sll = Zx(Wnum): Sok = Lxs
Loop
Close #1
Call PLst(Lst2) '排列输出
Rt.Caption = "结果是：" & Left(Sok, 1)
ReadyP = False
End If
End Sub

Sub PLst(Sp() As String)
For I = 1 To UBound(Sp)
For j = I To UBound(Sp)
If Val(Split(Sp(I), "：")(1)) > Val(Split(Sp(j), "：")(1)) Then
tmp = Sp(I)
Sp(I) = Sp(j)
Sp(j) = tmp
End If
Next j
Next I
For I = 1 To UBound(Sp)
List1.AddItem Sp(I)
Next I
End Sub


Function WIsn(A As String, B() As String) As Double  '辨别点最近        定点  定线
Dim I As Integer
Dim Ax As Double, Ay As Double, Bx As Double, By As Double
Dim Dx As Double, Dy As Double
Dim Lg() As Double
ReDim Lg(0 To UBound(B)) As Double
WIsn = 10000
Ax = Val(Split(A, ",")(0))
Ay = Val(Split(A, ",")(1)) '获得A坐标
For I = 0 To UBound(B)
 Bx = Val(Split(B(I), ",")(0))
 By = Val(Split(B(I), ",")(1)) '获得坐标
  Dx = Abs(Ax - Bx)
  Dy = Abs(Ay - By) '平面距离
  Lg(I) = Sqr(Dx ^ 2 + Dy ^ 2) '算出距离长度
  If Lg(I) < WIsn Then WIsn = Lg(I)
Next I
End Function


Private Sub Ptg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Gok_Click: KeyAscii = 0
End Sub

Private Sub Ptg_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
ReDim Preserve Xz(1 To 1) As Integer, Yz(1 To 1) As Integer
Ping = True: Xl = X: Yl = y: Xz(1) = X: Yz(1) = y
End Sub

Private Sub Ptg_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Ping = True And X >= 0 And y >= 0 And X < Ptg.ScaleWidth And y < Ptg.ScaleHeight Then
ReadyP = True
 Ptg.Line (Xl, Yl)-(X, y)
 Xl = X: Yl = y
 ReDim Preserve Xz(1 To (UBound(Xz) + 1)) As Integer
 ReDim Preserve Yz(1 To (UBound(Yz) + 1)) As Integer
 Xz(UBound(Xz)) = X
 Yz(UBound(Yz)) = y
 BBD2 = BBD2 + 1
 If BBD2 = Bbd Then Ptg.DrawWidth = Dw + 1: Rut = Rut & Xz(UBound(Yz)) / 210 & "," & Yz(UBound(Yz)) / 210 & "|": BBD2 = 0: Ptg.PSet (X, y), vbRed: Ptg.DrawWidth = Dw
End If
End Sub

Private Sub Ptg_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Ping = False
End Sub

Private Function GetC(sValue As String) As String
If sValue = "" Then Exit Function
Picture1.Cls
Picture1.Print sValue
Dim X%, y%
For X = 0 To Picture1.Width / 15
For y = 0 To Picture1.Height / 15
If Picture1.Point(X * 15, y * 15) = 0 Then
GetC = GetC & X & "," & y & "|" '以缩小15
End If
Next y
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
y = Val(Split(A(I), ",")(1))
 If X < SX Then SX = X
 If y < SY Then SY = y
Next I
For I = 0 To UBound(A)
x2 = Val(Split(A(I), ",")(0)) - SX
y2 = Val(Split(A(I), ",")(1)) - SY
A(I) = x2 & "," & y2
BeS = BeS & A(I) & "|"
Next I
If BeS <> "" Then BeS = Left(BeS, Len(BeS) - 1)
End Function

Function BeS2(A() As String, B() As String, Aw As Double, Ah As Double, Bw As Double, Bh As Double) As String
Dim I%, x2 As Double, y2 As Double
Dim BwD As Double, BhD As Double
'B()变
If Bw = 0 Then Bw = 0.000000001
If Bh = 0 Then Bh = 0.000000001
BwD = Aw / Bw
BhD = Ah / Bh
For I = 0 To UBound(B)
 x2 = Split(B(I), ",")(0)
 y2 = Split(B(I), ",")(1)
 x2 = x2 * BwD
 y2 = y2 * BhD
 BeS2 = BeS2 & x2 & "," & y2 & "|"
Next I
BeS2 = Left(BeS2, Len(BeS2) - 1)
End Function


















