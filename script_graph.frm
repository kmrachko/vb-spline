VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Построитель графиков"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Записать точки в БД"
      Height          =   615
      Left            =   9360
      TabIndex        =   20
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ввод точек"
      Height          =   1695
      Left            =   6720
      TabIndex        =   17
      Top             =   1800
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "Из БД"
         Height          =   615
         Left            =   360
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Вручную"
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   5
      Left            =   10920
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   10080
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   4
      Left            =   10920
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   3
      Left            =   10920
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   10920
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   10920
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   10080
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   10080
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   10080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   10080
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      FillStyle       =   0  'Solid
      Height          =   6000
      Left            =   360
      ScaleHeight     =   5940
      ScaleWidth      =   5940
      TabIndex        =   1
      Top             =   360
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Построить график функции"
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "P5:"
      Height          =   255
      Left            =   9360
      TabIndex        =   14
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "P4:"
      Height          =   255
      Left            =   9360
      TabIndex        =   5
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "P3:"
      Height          =   255
      Left            =   9360
      TabIndex        =   4
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "P2:"
      Height          =   255
      Left            =   9360
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "P1:"
      Height          =   255
      Left            =   9360
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w As Double, X(1 To 10000) As Double, Y(1 To 10000) As Double, i As Double, n As Long, o As Boolean, x1 As Double, h As Integer
Dim c(1 To 100) As Double, x3 As Double, l As Boolean
Dim db As Database
Dim rs As Recordset
Private Sub Command1_Click()

n = 5

o = True
Call draw
Call def
Call sort

If o = False Then
X(n) = 0
Y(n) = 0
n = n - 1
Exit Sub
End If


Call lines
Call coef


For x1 = X(1) + 0.0001 To X(n) + 0.0001 Step 0.001
Picture1.PSet (x1, -spl(x1)), vbRed
Next x1

End Sub

Private Sub draw()
Picture1.Cls
Picture1.ScaleMode = 0
Picture1.ScaleHeight = w
Picture1.ScaleWidth = w
Picture1.ScaleLeft = -(w / 2)
Picture1.ScaleTop = -(w / 2)

'osi
Picture1.Line (-(w / 2), 0)-((w / 2), 0)
Picture1.Line (0, -(w / 2))-(0, (w / 2))

'metki po x
Picture1.ScaleHeight = w * 6
Picture1.ScaleTop = -(3 * w)
For i = -(w / 2) To (w / 2)
Picture1.Line (i, 1)-(i, -1)
Picture1.PSet (i, 1)
Picture1.Print i
Next i
i = 0
Picture1.ScaleHeight = w
Picture1.ScaleTop = -(w / 2)

'metki po y
Picture1.ScaleWidth = w * 6
Picture1.ScaleLeft = -(3 * w)
For i = -(w / 2) To (w / 2)
Picture1.Line (-1, i)-(1, i)
If i <> 0 Then
Picture1.PSet (1, i)
Picture1.Print -i
End If
Next i
i = 0
Picture1.ScaleWidth = w
Picture1.ScaleLeft = -(w / 2)



End Sub

Private Sub Command2_Click()
Call sort

rs.MoveFirst
Do While Not rs.EOF

rs.Delete
rs.MoveNext

Loop



'Print n
For i = 1 To n
'Print X(i)
'Print Y(i)
'Print Chr(13)
rs.AddNew
rs.Fields("x_coor").Value = X(i)
rs.Fields("y_coor").Value = Y(i)
rs.Update

Next

End Sub

Private Sub Form_Load()
w = 16

n = 5

h = 1

For i = 1 To 5

Randomize Timer

Text1(i).Text = i
Text2(i).Text = CInt(Rnd() * 6)

X(i) = Text1(i).Text
Y(i) = Text2(i).Text

Next

Set db = OpenDatabase(App.Path & "\db_spline.mdb")

Set rs = db.OpenRecordset("Select * from points", dbOpenDynaset)
Option1.Value = True
End Sub
Private Sub def()
'o = True
'zadanie textbox

If Option1.Value = True Then
For i = 1 To 5

If Not IsNumeric(Text1(i).Text) Or Not IsNumeric(Text2(i).Text) Then
MsgBox "Вы ввели не цифры"
o = False
Exit Sub
End If

X(i) = Text1(i).Text

Y(i) = Text2(i).Text


Next

'zadanie bd

ElseIf Option2.Value = True Then
rs.MoveLast
n = rs.RecordCount
rs.MoveFirst
If n < 3 Then
MsgBox "Меньше трёх точек"
o = False
Exit Sub
End If

For i = 1 To n

If Not IsNumeric(rs.Fields("x_coor").Value) Or Not IsNumeric(rs.Fields("y_coor").Value) Then
MsgBox "Вы ввели не цифры"
o = False
Exit Sub
End If



X(i) = rs.Fields("x_coor").Value
Y(i) = rs.Fields("y_coor").Value


rs.MoveNext

'Print Chr(13)
'Print X(i)
'Print Y(i)
Next

If i < 5 Then o = False
End If
End Sub
Private Sub lines()

'postroenie lineynogo grafika

For i = 1 To n - 1
Picture1.DrawStyle = 2
Picture1.Line (X(i), -Y(i))-(X(i + 1), -Y(i + 1))
Next

Picture1.DrawStyle = 0
Picture1.FillColor = vbRed

'postroenie krugov
For i = 1 To n
Picture1.Circle (X(i), -Y(i)), 0.15, vbRed
Next

End Sub
Private Sub coef()
'?

ReDim k(1 To 100) As Double

Dim g As Double, j As Double, m As Double, a As Double, b As Double, r As Double, d As Double

k(1) = 0
c(1) = 0

For g = 3 To n
j = g - 1
m = j - 1
a = X(g) - X(j)
b = X(j) - X(m)
r = 2 * (a + b) - b * c(j)
c(g) = a / r
k(g) = (3 * ((Y(g) - Y(j)) / a - (Y(j) - Y(m)) / b) - b * k(j)) / r
Next

c(n) = k(n)

For g = n - 1 To 2 Step -1
c(g) = k(g) - c(g) * c(g + 1)
Next

End Sub
Private Function spl(x1 As Double) As Double

Dim p As Double, q As Double, d As Double, a As Double, b As Double, aa As Double, bb As Double, dd As Double

g = 1
Do While x1 > X(g) And g <> n
g = g + 1
Loop
j = g - 1
a = Y(j)
b = X(j)
q = X(g) - b
r = x1 - b
p = c(g)
d = c(g + 1)
b = (Y(g) - a) / q - (d + 2 * p) * q / 3
d = (d - p) / q * r
spl = a + r * (b + r * (p + d / 3))


End Function

Private Sub Picture1_Click()
Call draw
Call sort
Call lines
Call coef

For x1 = X(1) + 0.001 To X(n) - 0.001 Step 0.005
Picture1.PSet (x1, -spl(x1)), vbRed
Next x1
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
l = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X4 As Single, Y4 As Single)

If l And Shift Then

X(cl_pt(X4)) = X4
Y(cl_pt(X4)) = -Y4

Call draw
Call sort
Call lines
Call coef

For x1 = X(1) + 0.001 To X(n) - 0.001 Step 0.005
Picture1.PSet (x1, -spl(x1)), vbRed
Next x1

End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X2 As Single, Y2 As Single)
l = False
If Button = 2 Then
o = True

If n > 5 Then


For h = 1 To n
If h = cl_pt(X2) Or h > cl_pt(X2) Then
X(h) = X(h + 1)
Y(h) = Y(h + 1)
End If
Next h

n = n - 1

End If

Call draw
Call sort
Call lines
Call coef

For x1 = X(1) + 0.001 To X(n) - 0.001 Step 0.005
Picture1.PSet (x1, -spl(x1)), vbRed
Next x1


ElseIf Button = 1 Then

o = True
n = n + 1
Call draw
X(n) = X2
Y(n) = -Y2
Call sort

If o = False Then
X(n) = 0
Y(n) = 0
n = n - 1
Exit Sub
End If
Call lines
Call coef

For x1 = X(1) + 0.001 To X(n) - 0.001 Step 0.005
Picture1.PSet (x1, -spl(x1)), vbRed
Next x1

End If
End Sub

Private Sub sort()

Dim q As Integer, t As Double, e As Integer

For q = 1 To n - 1
    For e = q + 1 To n
    If X(q) = X(e) Then o = False
    Next e
Next q
If o = False Then Exit Sub

For q = 1 To n - 1

    For e = 1 To n - q
    If X(e) > X(e + 1) Then
    
    t = X(e)
    X(e) = X(e + 1)
    X(e + 1) = t
    
    t = Y(e)
    Y(e) = Y(e + 1)
    Y(e + 1) = t
    
    End If
    Next e
    
Next q

'Form1.Cls

'For q = 1 To n
'Print X(q)
'Next q


End Sub
Private Function cl_pt(x1 As Single) As Integer
Dim xx1 As Double, xx2 As Double, xx3 As Double

    If x1 <= X(1) Then
    cl_pt = 1
    Exit Function
    ElseIf x1 >= X(n) Then
    cl_pt = n
    Exit Function
    End If
    
For i = 1 To n
 
    If X(i) <= x1 Then
    xx1 = X(i)
    xx2 = X(i + 1)
    End If
    
Next i


If (x1 - xx1) < (xx2 - x1) Then
xx3 = xx1
ElseIf (x1 - xx1) > (xx2 - x1) Then
xx3 = xx2
Else
xx3 = xx1
End If

For i = 1 To n
    If X(i) = xx3 Then cl_pt = i
Next i

End Function
