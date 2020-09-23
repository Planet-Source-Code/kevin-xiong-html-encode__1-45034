VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Encrypter"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConvert 
      Caption         =   "En&code"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1440
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   240
      Width           =   6135
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "&Open File"
      Height          =   285
      Left            =   7680
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   120
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.html;*.htm"
   End
   Begin VB.TextBox Text2 
      Height          =   2805
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   840
      Width           =   7455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   "Copy and paste to new file."
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Encrypted File:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "File to Encrypt:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function convertToHex()
     Dim vAscii As Long
     Dim sString As String
     Dim vHex As String
         
     Dim xOutput As String
     
     Dim x As Long
     Dim y As Long
     
     Text2.Text = Text2.Text & "<! This page is protected by HTML Encoding !>" & vbCrLf
     Text2.Text = Text2.Text & "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
     Text2.Text = Text2.Text & "<html>" & vbCrLf
     Text2.Text = Text2.Text & "<head>" & vbCrLf
     Text2.Text = Text2.Text & "<meta name=" & Chr(34) & "generator" & Chr(34) & " content=" & Chr(34) & "HTML Encoding" & Chr(34) & ">" & vbCrLf
     Text2.Text = Text2.Text & "<meta http-equiv=content-type content=" & Chr(34) & "text/html; chrset=iso-8859-1" & Chr(34) & ">" & vbCrLf
     Text2.Text = Text2.Text & "<META HTTP-EQUIV=" & Chr(34) & "Pragma" & Chr(34) & " CONTENT=" & Chr(34) & "no-cache" & Chr(34) & ">" & vbCrLf
     Text2.Text = Text2.Text & "<META HTTP-EQUIV=" & Chr(34) & "expires" & Chr(34) & " CONTENT=" & Chr(34) & "1909-01-27T00:35:44+00:00" & Chr(34) & ">" & vbCrLf
     Text2.Text = Text2.Text & "<META HTTP-EQUIV=" & Chr(34) & "imagetoolbar" & Chr(34) & " CONTENT=" & Chr(34) & "no" & Chr(34) & ">" & vbCrLf
     Text2.Text = Text2.Text & "</head>" & vbCrLf
     Text2.Text = Text2.Text & "<body>" & vbCrLf
     Text2.Text = Text2.Text & "<script language=" & Chr(34) & "JavaScript" & Chr(34) & ">" & vbCrLf
     Text2.Text = Text2.Text & "<!-- Hide" & vbCrLf

     Text2.Text = Text2.Text & "var encrypt = " & Chr(34) & "eT69Hwg6^S1GFueT6*(&^&*@))#)$#(@@#(@@))!)_#__#9Hwg^df1GFuaeeT^69Hwg6SeafasdfOUI*&$^^&#$(#)#*$*%&%&$$(#)#1GFueafasdfOUI*&$^^&#$(#)#*$*%&%&$$(#)#1GFu#(@@))!)_#__#9Hwg^df1GFuaeeT^69Hwg6SeafasdfOUI*&$^^&#$(#)#*$*%&%&$$(#)#1GFu#(@@))!)_#__#9Hwg^df1GFuaeeT^69Hwg6SeafasdfOUI*&$^^&#$(#)#*$*%&%&$$(#)#1GFu" & Chr(34) & ";"
     Text2.Text = Text2.Text & "eval(unescape(" & Chr(34)
     sString = Text1.Text
     
     x = Len(sString)
     
     For y = 1 To x
          vAscii = Asc(Mid(sString, y, 1))
          vHex = Hex(vAscii)
          
          If Len(vHex) = 1 Then
               vHex = "0" + vHex
          End If
          
          xOutput = xOutput + "%" + CStr(vHex)
     Next
     
     Text2.Text = Text2.Text & xOutput
     Text2.Text = Text2.Text & Chr(34) & "));" & vbCrLf
     Text2.Text = Text2.Text & "// end hide -->" & vbCrLf
     Text2.Text = Text2.Text & "</script" & vbCrLf
     Text2.Text = Text2.Text & "<center><noscript>Page protected by HTML Encoding.  This page requires a JavaScript-enabled web browser.</noscript></center>" & vbCrLf
     Text2.Text = Text2.Text & "</body>" & vbCrLf
     Text2.Text = Text2.Text & "</html>" & vbCrLf
     
End Function

Private Sub cmdConvert_Click()
     Text1.Text = ""
     Text2.Text = ""
     If getFile = False Then
          Exit Sub
     End If
     convertToHex
End Sub

Function getFile() As Boolean
     Dim strBuffer As String
     
     On Error GoTo ErrorhHandler
     
     If txtFilePath.Text = "" Then
          MsgBox "Please select a valid HTML file!", vbInformation, "No File Selected"
          getFile = False
          Exit Function
     Else
          Open txtFilePath.Text For Input As #1
     
          Do Until EOF(1)
          
               Line Input #1, strBuffer
               strBuffer = Replace(strBuffer, Chr(34), "\" + Chr(34))
               strBuffer = "document.writeln(" + Chr(34) + strBuffer + Chr(34) + ");"
               
               Debug.Print strBuffer
               Text1.Text = Text1.Text + strBuffer + vbCrLf
          Loop
          
          Close #1
          
          getFile = True
          
          Exit Function
     End If
     
ErrorhHandler:
     
     getFile = False
     
End Function

Private Sub cmdExit_Click()
     End
End Sub

Private Sub cmdOpenFile_Click()

     cdlg1.ShowOpen
     
     If InStr(1, cdlg1.FileName, "htm") > 0 Or InStr(1, cdlg1.FileName, "html") > 0 Then
          txtFilePath.Text = cdlg1.FileName
     ElseIf cdlg1.FileName <> "" Then
          MsgBox "Invalid file type!" & vbCrLf & "Only html file can be encrypted with this tool", vbInformation, "Invalid file Type"
          txtFilePath.Text = ""
          Exit Sub
     End If
     
End Sub

