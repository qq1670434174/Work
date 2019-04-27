VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "2kl"
   ClientHeight    =   13905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   24030
   LinkTopic       =   "Form1"
   ScaleHeight     =   13905
   ScaleWidth      =   24030
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "由附图说明生成标号"
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   12480
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "生成附图说明"
      Height          =   735
      Left            =   17880
      TabIndex        =   9
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   2655
      Left            =   17760
      TabIndex        =   8
      Top             =   1680
      Width           =   5415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清空"
      Height          =   735
      Left            =   6360
      TabIndex        =   7
      Top             =   12480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "仅添加数字"
      Height          =   615
      Left            =   2760
      TabIndex        =   6
      Top             =   12480
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   7815
      Left            =   9240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   6360
      Width           =   16335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加括号和数字"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   12480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   4695
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   6480
      Width           =   8055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_click()
Dim a As String
Dim b As String
Dim e As String
Dim c(99) As String
Dim d(99) As String
Dim f(99) As String
f(0) = Text2.Text
For j = 0 To 90
If j Mod 10 = 0 Then
c(j) = Text1(j).Text & "（" & (j / 10 + 1) & "）"
Else: c(j) = Text1(j).Text & "（" & (j + 10) & "）"
End If
'Text1(j).Text = c(j)
d(j) = Text1(j).Text
f(j + 1) = Replace(f(j), d(j), c(j))
Next
Text3 = f(91)
End Sub


Private Sub Command2_Click()
Dim a As String
Dim b As String
Dim e As String
Dim c(99) As String
Dim d(99) As String
Dim f(99) As String
Dim g(99) As String
f(0) = Text2.Text
For j = 0 To 90
If Text1(j).Text <> "" Then
If j Mod 10 = 0 Then
c(j) = Text1(j).Text & (j / 10 + 1)
Else: c(j) = Text1(j).Text & (j + 10)
End If
End If
d(j) = Text1(j).Text
f(j + 1) = Replace(f(j), c(j), d(j))

Next

g(0) = f(91)
For j = 0 To 90
If j Mod 10 = 0 Then
c(j) = Text1(j).Text & (j / 10 + 1)
Else: c(j) = Text1(j).Text & (j + 10)
End If
d(j) = Text1(j).Text

g(j + 1) = Replace(g(j), d(j), c(j))
Next

Text3 = g(91)
End Sub

Private Sub Command3_Click()
Text2.Text = ""
End Sub

Private Sub Command4_Click()
Dim x(99) As String
Dim y(99) As String
For j = 0 To 90
If Text1(j).Text <> "" Then
If j Mod 10 = 0 Then
x(j) = (j / 10 + 1) & "、" & Text1(j).Text & "，"
Else: x(j) = (j + 10) & "、" & Text1(j).Text & "，"

End If
End If
y(j + 1) = y(j) & x(j)
Next
Text4 = y(91)
End Sub

Private Sub Form_load()
For k = 1 To 9
 For i = 1 To 9
  Load Text1(i + (k - 1) * 10)
  Text1(i + (k - 1) * 10).Top = Text1(i + (k - 1) * 10 - 1).Top + Text1(i + (k - 1) * 10 - 1).Height + 100

  If k > 1 Then Text1(i + (k - 1) * 10).Left = Text1(k * 10 - 20).Left + Text1(k * 10 - 20).Width + 100
  Text1(i + (k - 1) * 10).Visible = True
  'Print Text1(i + (k - 1) * 10).Left
 Next
Load Text1(k * 10)
Text1(k * 10).Left = Text1(k * 10 - 10).Left + Text1(k * 10 - 10).Width + 100
Text1(k * 10).Visible = True
Next
For m = 0 To 90
Text1(m).Visible = True
'Text1(m).Text = m
Next

For i = 1 To 9
Load Label1(i)
Label1(i).Left = Label1(i - 1).Left + Label1(i - 1).Width + 100
Label1(i).Visible = True
Label1(i).Caption = i + 1
Next

For i = 1 To 9
Load Label2(i)
Label2(i).Top = Label2(i - 1).Top + Label2(i - 1).Height + 100
Label2(i).Visible = True
Label2(i).Caption = i
Next
End Sub

