Attribute VB_Name = "Module1"
Dim l As Integer

Public Sub load()
With Form1
.Label2.Left = 0
.Label3.Left = 0
.Label4.Left = 0
.Label5.Left = 0
End With
l = 1
End Sub

Public Sub vlabel()
With Form1
.Label2.Visible = Not .Label2.Visible
.Label3.Visible = .Label2.Visible
.Label4.Visible = .Label2.Visible
.Label5.Visible = .Label2.Visible
End With
End Sub

Public Sub label2mm()
With Form1
.Label2.ForeColor = vbBlue
.Label3.ForeColor = vbCyan
.Label4.ForeColor = vbCyan
.Label5.ForeColor = vbCyan
.Label2.FontSize = 36
.Label3.FontSize = 20
.Label4.FontSize = 20
.Label5.FontSize = 20
End With
End Sub

Public Sub label3mm()
With Form1
.Label3.ForeColor = vbBlue
.Label2.ForeColor = vbCyan
.Label4.ForeColor = vbCyan
.Label5.ForeColor = vbCyan
.Label2.FontSize = 20
.Label3.FontSize = 36
.Label4.FontSize = 20
.Label5.FontSize = 20
End With
End Sub



Public Sub label4mm()
With Form1
.Label4.ForeColor = vbBlue
.Label3.ForeColor = vbCyan
.Label2.ForeColor = vbCyan
.Label5.ForeColor = vbCyan
.Label4.FontSize = 36
.Label3.FontSize = 20
.Label2.FontSize = 20
.Label5.FontSize = 20
End With
End Sub




Public Sub label5mm()
With Form1
.Label5.ForeColor = vbBlue
.Label3.ForeColor = vbCyan
.Label4.ForeColor = vbCyan
.Label2.ForeColor = vbCyan
.Label2.FontSize = 20
.Label5.FontSize = 36
.Label4.FontSize = 20
.Label3.FontSize = 20
End With
End Sub
Public Sub aks()

Unload Form1
Form1.Show
On Error GoTo 1
If Form2.Caption <= Form2.Image5.Count Then
Form1.Picture = Form2.Image5(Form2.Caption)
Else
p = MsgBox("Shoma Barandeh Shdid.", vbOKOnly, "Barandeh")
Form2.Caption = 0
Unload Form1
Form1.Show
End If
Form1.PSet (2500, 4500), vbWhite
Form1.Print "S"
Form1.PSet (1600, 360), vbWhite
Form1.Print "p"
Module1.load
Form2.Caption = Form2.Caption + 1
1:
End Sub
