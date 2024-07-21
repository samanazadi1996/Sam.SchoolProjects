Attribute VB_Name = "Moduleppro"
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

Public Sub LISTPRO()
On Error GoTo 1
Select Case Form1.ppro

Case "Calculator"
If Form1.Pro2.AutoRedraw = True Then Exit Sub
Form1.Pro2.AutoRedraw = True
    Form1.Pro2.Visible = True
    Form1.Pro2.Move 1200, 700
Form1.TPro = Form1.TPro + 1
Case "Show IP"
If Form1.pmsgbox.AutoRedraw = True Then Exit Sub
Form1.pmsgbox.AutoRedraw = True
    Form1.pmsgbox.Visible = True
    Form1.Lblmb.FontSize = 12
    Form1.Lblmb.Caption = "IP is:"
''''''''''''''''''''''''''''
Dim Ret As Long, Tel As Long
Dim MyByte(3) As Byte
Dim Listing As Long
GetIpAddrTable ByVal 0&, Ret, True
ReDim bbytes(o To Ret - 1) As Byte
GetIpAddrTable bbytes(0), Ret, False
CopyMemory Listing, 1, 4
CopyMemory Listing, bbytes(4), Len(Listing)
CopyMemory MyByte(0), Listing, 4
Form1.textipmsgbox = CStr(MyByte(0)) + "." + CStr(MyByte(1)) + "." + CStr(MyByte(2)) + "." + CStr(MyByte(3))
''''''''''''''''''''''''''

    Form1.pmsgbox.Move 1500, 1500
    Form1.TPro = Form1.TPro + 1
Case "Show Computer name"
If Form1.pmsgbox.AutoRedraw = True Then Exit Sub
    Form1.pmsgbox.AutoRedraw = True
    Form1.pmsgbox.Visible = True
    Form1.Lblmb.FontSize = 8
    Form1.Lblmb.Caption = "Computer Name as:"
    Moduleother.computername
    Form1.pmsgbox.Move 1500, 1500
    Form1.TPro = Form1.TPro + 1
Case "Paint"
If Form1.Pro3.AutoRedraw = True Then Exit Sub
Form1.Pro3.AutoRedraw = True
Form1.Pro3.Visible = True
Form1.Pro3.Move 1400, 900
Form1.TPro = Form1.TPro + 1

Case "Note Pad"
If Form1.PRONOTEPAD.AutoRedraw = True Then Exit Sub
Form1.PRONOTEPAD.AutoRedraw = True
Form1.PRONOTEPAD.Visible = True
Form1.TextNOTAPAD.Text = ""
Form1.PRONOTEPAD.Move 1600, 1100
Form1.TPro = Form1.TPro + 1
End Select
1:
End Sub

Public Sub Proper()
    If Form1.pro1.AutoRedraw = True Then Exit Sub
    Form1.pro1.AutoRedraw = True
    Form1.pro1.Visible = True
    Form1.pro1.Move 1000, 500
    Form1.TPro = Form1.TPro + 1
End Sub

Public Sub AboutPro()
If Form1.hlp.AutoRedraw = True Then Exit Sub
Form1.hlp.AutoRedraw = True
Form1.hlp.Visible = True
Form1.hlp.Move 1800, 1300
Form1.TPro = Form1.TPro + 1
End Sub
