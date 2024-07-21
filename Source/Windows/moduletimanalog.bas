Attribute VB_Name = "Module"
Dim q As Single, X As String
Public Sub fl()
With Form1
.Label12.Width = .Pro3.Width
.Image7.Left = .Pro3.Width - 500
.Frame.Move 120, .Frame.Top, .Pro3.Width - 540, .Pro3.Height - 340 - .Frame.Top
.VScroll1.Move .Frame.Width + 120, .VScroll1.Top, .VScroll1.Width, .Frame.Height
.HScroll1.Move 120, .Frame.Top + .Frame.Height, .Frame.Width
For w = 5 To 60 Step 5
.Combo1.AddItem (w)
Next
.Combo1 = .Combo1.List(0)
For w = 1 To 100
.border.AddItem w
Next
.border = .border.List(0)
End With

End Sub
Public Sub PPaint_m_u(Button As Integer, X As Single, Y As Single, xx As Single, yy As Single)
With Form1
If Button = 1 Then
If .Optionline.Value = True Then .Picturepaint.Line (xx, yy)-(X, Y)
If .Optionmo.Value = True Then .Picturepaint.Line (xx, yy)-(X, Y), .cmdcolor.BackColor, B
If .Optionda.Value = True Then
If X > Y Then .Picturepaint.Circle (xx, yy), Abs(xx - X)
If X <= Y Then .Picturepaint.Circle (xx, yy), Abs(yy - Y)
End If
End If
End With
End Sub
Public Sub PPaint_m_m(Button As Integer, X As Single, Y As Single, xx As Single, yy As Single)
If Button = 1 And Form1.Optionpen.Value = True Then
Form1.Picturepaint.Line (xx, yy)-(X, Y)
xx = X
yy = Y
End If
End Sub


Public Sub Ppaint()
If Form1.Picturepaint.Height > Form1.Frame.Height Then Form1.VScroll1.Max = Int((Form1.Picturepaint.Height - Form1.Frame.Height + 120) / 200)
If Form1.Picturepaint.Width > Form1.Frame.Width Then Form1.HScroll1.Max = Int((Form1.Picturepaint.Width - Form1.Frame.Width + 120) / 200)
Form1.VScroll1.Move Form1.Frame.Width + 120, Form1.VScroll1.Top, Form1.VScroll1.Width, Form1.Frame.Height
Form1.HScroll1.Move 120, Form1.Frame.Top + Form1.Frame.Height, Form1.Frame.Width
End Sub
Public Sub Col()
With Form1
.Textc.Move 120, .Textc.Top, .Pro2.ScaleWidth - 240
.FrameCOL.Move .Pro2.ScaleWidth / 2 - .FrameCOL.Width / 2, (.Pro2.ScaleHeight / 2 - .FrameCOL.Height / 2) + 270
End With
End Sub
Public Sub Cmdc_Click(Index As Integer)
Form1.Textc = Form1.Textc & Form1.Cmdc(Index).Caption
End Sub
Public Sub Cmdc2_Click(Index As Integer)
On Error Resume Next
X = Form1.Cmdc2(Index).Caption
q = Form1.Textc
Form1.Textc = ""
End Sub
Public Sub CmdC3_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then Form1.Textc = Sin(3.14 * Form1.Textc / 180)
If Index = 1 Then Form1.Textc = Cos(3.14 * Form1.Textc / 180)
If Index = 2 Then Form1.Textc = Tan(3.14 * Form1.Textc / 180)
End Sub
Public Sub C_Click()
On Error Resume Next
If X = "*" Then Form1.Textc = q * Form1.Textc
If X = "/" Then Form1.Textc = q / Form1.Textc
If X = "-" Then Form1.Textc = q - Form1.Textc
If X = "+" Then Form1.Textc = q + Form1.Textc
End Sub
Public Sub Command12_Click()
Form1.Textc = ""
End Sub
Public Sub C3_Click()
If Len(Form1.Textc) <> 0 Then Form1.Textc = Left(Form1.Textc, Len(Form1.Textc) - 1)
End Sub

Public Sub NewTextDocument()
On Error Resume Next
Randomize Timer
r = FreeFile
Open "C:\Documents and Settings\All Users\Desktop\New Text Document " & Rnd * 10000 & ".txt" For Output As #r
Close #r

End Sub

Public Sub AllFolder()
With Form1
.Folders.Path = "C:\Documents and Settings\" & .lblusername & "\Desktop"
.Folders.Refresh
.ListFolders.Clear
.ImageFolder.ToolTipText = "Number Folder On The Desktop Is " & .Folders.ListCount
If .Folders.ListCount <> 0 Then
    For w = 0 To .Folders.ListCount - 1
    .ListFolders.AddItem (Mid(.Folders.List(w), Len(.Folders.Path) + 2, Len(.Folders.List(w)) - Len(.Folders.Path)))
    Next
End If
End With
End Sub
