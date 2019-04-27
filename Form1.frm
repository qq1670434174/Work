VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20910
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   20910
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   2280
      TabIndex        =   0
      Text            =   ",1、手套外层，2、手套内层，3、腔体，4、散热孔，5、注射阀，6、减压阀，7、限流阀，8、退热凝胶。"
      Top             =   840
      Width           =   12015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_click()
Dim a(90) As String
Dim b(90) As Integer
Dim c(90) As String
Dim d(90) As Integer
Dim e(90) As Integer
For i = 0 To 90
If i Mod 10 = 0 Then
a(i) = "，" & i / 10 + 1 & "、"
Else
a(i) = "，" & i + 10 & "、"
End If
Next

For i = 0 To 90
b(i) = InStrRev(Text1.Text, a(i))
If b(i) <> 0 Then Print b(i)
Next
 
 j = 0
For i = 0 To 90
If b(i) <> 0 Then
d(j) = b(i)
If j > 0 Then
c(i) = Mid(Text1.Text, d(j - 1) + 3, d(j) - d(j - 1) - 3)
j = j + 1
Else
j = j + 1
End If

Print i & c(i)
End If
Next

End Sub

