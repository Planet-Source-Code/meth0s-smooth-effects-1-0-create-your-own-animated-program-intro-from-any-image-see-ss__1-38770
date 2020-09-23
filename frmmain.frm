VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smooth Effects www.xgeek.org"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "frmmain"
   MaxButton       =   0   'False
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   786
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Help!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   9240
      TabIndex        =   18
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      Caption         =   "No Color"
      Height          =   1050
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   9480
      TabIndex        =   12
      Text            =   "10"
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9480
      TabIndex        =   10
      Text            =   "Timer1"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9480
      TabIndex        =   8
      Text            =   "Picture1"
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dont Generate Source Code (Preview Only)"
      Height          =   1050
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.PictureBox picactive 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   1425
      TabIndex        =   6
      Top             =   3195
      Width           =   1455
   End
   Begin VB.PictureBox picpick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   1425
      TabIndex        =   5
      Top             =   3795
      Width           =   1455
   End
   Begin VB.CommandButton cmdselpic 
      Caption         =   "Select Picture"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   3195
      Width           =   1455
   End
   Begin VB.CommandButton cmdcode 
      Caption         =   "Start"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   3795
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7800
      Top             =   2160
   End
   Begin VB.PictureBox picpreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   3000
      Left            =   3240
      ScaleHeight     =   196
      ScaleMode       =   0  'User
      ScaleWidth      =   557
      TabIndex        =   2
      Top             =   120
      Width           =   8415
   End
   Begin VB.TextBox txtcode 
      Height          =   3000
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox pictrace 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2490
      Left            =   120
      Picture         =   "frmmain.frx":0442
      ScaleHeight     =   162
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   4440
      Width           =   3810
   End
   Begin VB.Label Label6 
      Caption         =   $"frmmain.frx":10C6
      Height          =   615
      Left            =   4080
      TabIndex        =   17
      Top             =   5280
      Width           =   4935
   End
   Begin VB.Label Label5 
      Caption         =   "Hey do you like this or atleast think its a little cool? Vote for me and my work! at www.planet-source-code.com !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   16
      Top             =   6600
      Width           =   6615
   End
   Begin VB.Label Label4 
      Caption         =   $"frmmain.frx":1159
      Height          =   735
      Left            =   4080
      TabIndex        =   15
      Top             =   4440
      Width           =   4935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Timer Control Interval:"
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Timer Control Name:"
      Height          =   255
      Left            =   7800
      TabIndex        =   11
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Picture Box Name:"
      Height          =   255
      Left            =   7920
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X1(0 To 5000) As Long
Dim Y1(0 To 5000) As Long
Dim C1(0 To 5000) As Long

Dim nomore(0 To 5000) As Boolean
Dim ismore(0 To 5000) As Boolean

Dim totaltracecount As Long
Dim reboundlength As Long

Dim imagetrace1 As Long
Dim imagetrace2 As Long

Function traceimage(PicSrc As PictureBox, bColor As OLE_COLOR)
Dim looper As Integer
Dim looper2 As Integer
Dim bColor2 As OLE_COLOR
If Check1.Value = vbUnchecked Then
    txtcode.Text = txtcode.Text & "Public Function DrawGraph()" & vbCrLf
End If
For looper = 0 To PicSrc.Height - 8
    'If Right(looper, 1) = 2 Then
    If looper = imagetrace1 Then
        imagetrace1 = looper + 4
        For looper2 = 0 To PicSrc.Width - 4
           If Right(looper2, 1) = 2 Then
            'If looper2 = imagetrace2 Then
                'imagetrace2 = looper2 + 4
                bColor2 = PicSrc.Point(looper2, looper)
                If bColor2 <> bColor Then
                    If Check2.Value = vbChecked Then
                        bColor2 = vbBlack
                    Else
                        bColor2 = PicSrc.Point(looper2, looper)
                    End If
                    If Check1.Value = vbUnchecked Then
                        txtcode.Text = txtcode.Text & "   X1(" & totaltracecount & ") = " & looper2 & vbCrLf
                        txtcode.Text = txtcode.Text & "   Y1(" & totaltracecount & ") = " & looper & vbCrLf
                        txtcode.Text = txtcode.Text & "   C1(" & totaltracecount & ") = " & bColor2 & vbCrLf
                    End If
                    X1(totaltracecount) = looper2
                    Y1(totaltracecount) = looper
                    C1(totaltracecount) = bColor2
                    totaltracecount = totaltracecount + 1
                End If
            End If
        Next looper2
    End If
    Me.Caption = looper & " //~1/ - " & totaltracecount & " / " & imagetrace1 & " * " & imagetrace2
    DoEvents
Next looper
If Check1.Value = vbUnchecked Then
    txtcode.Text = txtcode.Text & "End Function" & vbCrLf & vbclrf
End If
End Function

Private Sub cmdcode_Click()
    For i = 0 To 5000
        X1(i) = 0
        Y1(i) = 0
        C1(i) = 0
        ismore(i) = False
        nomore(i) = False
    Next i
    If picpick.BackColor = &HC0C0FF Then
        MsgBox "Make sure you move your mouse over the image you selected and select a color to block out!", vbCritical, "Hey!"
    End If
    txtcode.Text = ""
    imagetrace1 = 4
    imagetrace2 = 4
    totaltracecount = 0
    Timer1.Enabled = False
    traceimage pictrace, picpick.BackColor
    If Check1.Value = vbUnchecked Then
        writedec
        writeform
        writetimer
        MsgBox "Great job everything appears to be done!" & vbCrLf & vbCrLf & "Now copy all the source code in the text box and copy it into your new project!", vbExclamation, "HEY!"
        testab = MsgBox("Would you like to copy the source code to your clipboard?" & vbCrLf & vbCrLf & "If you dont know what to choose just hit no!", vbYesNo Or vbInformation, "Hey!")
        If testab = vbYes Then
            Clipboard.Clear
            Clipboard.SetText txtcode.Text
        End If
    End If
    Timer1.Enabled = True
End Sub

Private Sub cmdselpic_Click()
On Error Resume Next
Dim FileName As SelectedFile
    FileDialog.sDlgTitle = "Import Picture - GRAPHIC FILES ONLY!"
    FileDialog.sFilter = "All Files (*.*)" & Chr(0) & "*.*"
    FileName = ShowOpen(hWnd)
    pictrace.Picture = LoadPicture(FileName.sLastDirectory & FileName.sFiles(1))
End Sub

Private Sub Command1_Click()
    MsgBox "Hello welcome to Smooth Effects 1.0 by meth0s" & vbCrLf & vbCrLf & "this program will actually write the entire first part of your program intro! Smooth Effects will take any image you want and transvert it into little tiny layers write the code and everything for you! all you have to do is create a new project and add a picture box and a timer! then paste the source code!", vbInformation, "Help!!"
End Sub

Private Sub Form_Load()
    picactive.AutoRedraw = True
    picpick.AutoRedraw = True
    picactive.CurrentX = picactive.ScaleWidth / 2 - picactive.TextWidth("Active Color") / 2
    picactive.CurrentY = picactive.ScaleHeight / 2 - picactive.TextHeight("Active Color") / 2
    picactive.Print "Active Color"
    picpick.CurrentX = picactive.ScaleWidth / 2 - picactive.TextWidth("Selected Color") / 2
    picpick.CurrentY = picactive.ScaleHeight / 2 - picactive.TextHeight("Selected Color") / 2
    picpick.Print "Selected Color"
    MsgBox "Hello welcome to Smooth Effects 1.0 by meth0s" & vbCrLf & vbCrLf & "this program will actually write the entire first part of your program intro! Smooth Effects will take any image you want and transvert it into little tiny layers write the code and everything for you! all you have to do is create a new project and add a picture box and a timer! then paste the source code!", vbInformation, "Hey!"
End Sub

Private Sub pictrace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picpick.BackColor = pictrace.Point(X, Y)
    picpick.CurrentX = picactive.ScaleWidth / 2 - picactive.TextWidth("Selected Color") / 2
    picpick.CurrentY = picactive.ScaleHeight / 2 - picactive.TextHeight("Selected Color") / 2
    picpick.Print "Selected Color"
End Sub

Private Sub pictrace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picactive.BackColor = pictrace.Point(X, Y)
    picactive.CurrentX = picactive.ScaleWidth / 2 - picactive.TextWidth("Active Color") / 2
    picactive.CurrentY = picactive.ScaleHeight / 2 - picactive.TextHeight("Active Color") / 2
    picactive.Print "Active Color"
End Sub

Public Function writeform()
    txtcode.Text = txtcode.Text & "Private Sub Form_Load()" & vbCrLf
    txtcode.Text = txtcode.Text & "   " & Text2.Text & ".Enabled = False" & vbCrLf
    txtcode.Text = txtcode.Text & "   reboundlength = 2" & vbCrLf
    txtcode.Text = txtcode.Text & "   " & Text1.Text & ".ScaleMode = 3" & vbCrLf
    txtcode.Text = txtcode.Text & "   " & Text1.Text & ".AutoRedraw = True" & vbCrLf
    txtcode.Text = txtcode.Text & "   Me.ScaleMode = 3" & vbCrLf
    txtcode.Text = txtcode.Text & "   Call DrawGraph" & vbCrLf
    txtcode.Text = txtcode.Text & "   " & Text2.Text & ".Interval = " & CStr(Text3.Text) & vbCrLf
    txtcode.Text = txtcode.Text & "   " & Text2.Text & ".Enabled = True" & vbCrLf
    txtcode.Text = txtcode.Text & "End Sub" & vbCrLf & vbCrLf
End Function

Public Function writedec()
    txtcode.Text = "Dim X1(0 To " & totaltracecount & ") As Long" & vbCrLf & vbCrLf & txtcode.Text
    txtcode.Text = "Dim Y1(0 To " & totaltracecount & ") As Long" & vbCrLf & txtcode.Text
    txtcode.Text = "Dim C1(0 To " & totaltracecount & ") As Long" & vbCrLf & txtcode.Text

    txtcode.Text = "Dim nomore(0 To " & totaltracecount & ") As Boolean" & vbCrLf & txtcode.Text
    txtcode.Text = "Dim ismore(0 To " & totaltracecount & ") As Boolean" & vbCrLf & txtcode.Text
    
    txtcode.Text = "Dim totaltracecount As Long" & vbCrLf & txtcode.Text
    
    txtcode.Text = "Dim imagetrace1 As Long" & vbCrLf & txtcode.Text
    txtcode.Text = "Dim imagetrace2 As Long" & vbCrLf & txtcode.Text
    txtcode.Text = "Dim reboundlength As Long" & vbCrLf & txtcode.Text
End Function

Public Function writetimer()
    txtcode.Text = txtcode.Text & "Private Sub " & Text2.Text & "_Timer()" & vbCrLf
    txtcode.Text = txtcode.Text & "   " & Text1.Text & ".Visible = False" & vbCrLf
    txtcode.Text = txtcode.Text & "   " & Text1.Text & ".Cls" & vbCrLf
    txtcode.Text = txtcode.Text & "For i = 0 To " & totaltracecount - 1 & vbCrLf
    txtcode.Text = txtcode.Text & "    If nomore(i) = True Then" & vbCrLf
    txtcode.Text = txtcode.Text & "        X1(i) = X1(i) - reboundlength" & vbCrLf
    txtcode.Text = txtcode.Text & "        If X1(i) <= 0 Then" & vbCrLf
    txtcode.Text = txtcode.Text & "            nomore(i) = False" & vbCrLf
    txtcode.Text = txtcode.Text & "        End If" & vbCrLf
    txtcode.Text = txtcode.Text & "    Else" & vbCrLf
    txtcode.Text = txtcode.Text & "        If X1(i) >= " & Text1.Text & ".Width Then" & vbCrLf
    txtcode.Text = txtcode.Text & "            X1(i) = X1(i) - reboundlength" & vbCrLf
    txtcode.Text = txtcode.Text & "            nomore(i) = True" & vbCrLf
    txtcode.Text = txtcode.Text & "        Else" & vbCrLf
    txtcode.Text = txtcode.Text & "            X1(i) = X1(i) + reboundlength" & vbCrLf
    txtcode.Text = txtcode.Text & "        End If" & vbCrLf
    txtcode.Text = txtcode.Text & "    End If" & vbCrLf
    txtcode.Text = txtcode.Text & "    If ismore(i) = True Then" & vbCrLf
    txtcode.Text = txtcode.Text & "        Y1(i) = Y1(i) - reboundlength" & vbCrLf
    txtcode.Text = txtcode.Text & "        If Y1(i) <= 0 Then" & vbCrLf
    txtcode.Text = txtcode.Text & "            ismore(i) = False" & vbCrLf
    txtcode.Text = txtcode.Text & "        End If" & vbCrLf
    txtcode.Text = txtcode.Text & "    Else" & vbCrLf
    txtcode.Text = txtcode.Text & "        If Y1(i) >= " & Text1.Text & ".Height Then" & vbCrLf
    txtcode.Text = txtcode.Text & "            Y1(i) = Y1(i) - reboundlength" & vbCrLf
    txtcode.Text = txtcode.Text & "            ismore(i) = True" & vbCrLf
    txtcode.Text = txtcode.Text & "        Else" & vbCrLf
    txtcode.Text = txtcode.Text & "            Y1(i) = Y1(i) + reboundlength" & vbCrLf
    txtcode.Text = txtcode.Text & "        End If" & vbCrLf
    txtcode.Text = txtcode.Text & "    End If" & vbCrLf
    txtcode.Text = txtcode.Text & "    " & Text1.Text & ".PSet (X1(i), Y1(i)), C1(i)" & vbCrLf
    txtcode.Text = txtcode.Text & "Next i" & vbCrLf
    txtcode.Text = txtcode.Text & "   " & Text1.Text & ".Visible = True" & vbCrLf
    txtcode.Text = txtcode.Text & "End Sub" & vbCrLf & vbCrLf
End Function

Private Sub Timer1_Timer()
    picpreview.Cls
    For i = 0 To totaltracecount - 1
        If nomore(i) = True Then
            X1(i) = X1(i) - 2
            If X1(i) = 0 Then
                nomore(i) = False
            End If
        Else
            If X1(i) >= picpreview.Width Then
                X1(i) = X1(i) - 2
                nomore(i) = True
            Else
                X1(i) = X1(i) + 2
            End If
        End If
        
        If ismore(i) = True Then
            Y1(i) = Y1(i) - 2
            If Y1(i) = 0 Then
                ismore(i) = False
            End If
        Else
            If Y1(i) >= picpreview.Height Then
                Y1(i) = Y1(i) - 2
                ismore(i) = True
            Else
                Y1(i) = Y1(i) + 2
            End If
        End If
        picpreview.PSet (X1(i), Y1(i)), C1(i)
    Next i
End Sub
