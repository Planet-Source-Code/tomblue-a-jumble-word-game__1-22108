VERSION 5.00
Begin VB.Form frmword 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Word Game 2001"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "frmgame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8535
      Begin VB.CommandButton Command3 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   7320
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRandomize 
         Caption         =   "Randomize"
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtinput 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Start"
         Height          =   375
         Left            =   6120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Chec&k"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   3360
         Width           =   45
      End
   End
   Begin VB.Image Image2 
      Height          =   1125
      Left            =   2970
      Picture         =   "frmgame.frx":0442
      Top             =   120
      Width           =   2850
   End
End
Attribute VB_Name = "frmword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Number As Long
Dim flag As Boolean
Dim mainvalue As String

' to interact in the game
' on the click event of the click1(0) the design time control check the validation
' else use the click event to interact with the end user
Private Sub Check1_Click(Index As Integer)
On Error GoTo err:

If flag = False Then
Else
Exit Sub
End If

If Index = 0 Then

For i = 1 To 2 * Number
    If Check1(i).Caption = "__" Then
    just = True
    Else
        If i = 2 * Number And just = False Then
        
        For m = Number + 1 To 2 * Number
        answer = answer & Check1(m).Caption
        Next m
            If answer = UCase(mainvalue) Then
            MsgBox "Successfull", 48 + vbOKOnly, "Jumble 1.1 -Word Game 2001"
            If mainvalue = "RAJESHLAL" Then Label1.Caption = "Word Game 2001 - By Rajesh Lal (razesh@hotmail.com) www.developerscolumn.com "

            Command2.Enabled = False
            flag = True
            Check1(0).Value = 0
            flag = False
            Else
            MsgBox "Click refresh or try again", 16 + vbOKOnly, "Jumble 1.1 -Word Game 2001"
            Command2.Enabled = True
            flag = True
            Check1(0).Value = 0
            flag = False
            End If
        End If
    End If
Next i


Else

If Index < Number + 1 Then
    If Check1(Index).Value = 1 Then
        For i = Number + 1 To 2 * Number
            If Check1(i).Caption = "__" Then
            Check1(i).Caption = Check1(Index).Caption
            flag = True
            Check1(i).Value = 0
            flag = False
            Exit Sub
            End If
        Next i
    Else
    
        For i = Number + 1 To 2 * Number
            If Check1(i).Caption = Check1(Index).Caption Then
            Check1(i).Caption = "__"
            flag = True
            Check1(i).Value = 1
            Check1(Index).Value = 0
            flag = False
            Exit Sub
            End If
        Next i
    
    End If

Else

If Check1(Index).Caption = "__" Then
flag = True
Check1(Index).Value = 1
flag = False
Exit Sub
End If
For ij = 1 To Number
If Check1(ij).Caption = Check1(Index).Caption Then
Check1(Index).Caption = "__"
flag = True
Check1(ij).Value = 0
flag = False
Exit Sub
End If
Next ij
End If
End If
Exit Sub
err:
MsgBox "Start the game and arrange the words before checking Error :" & err.Description, 16 + vbOKOnly, "Jumble 1.1"
End Sub
' generates the randomely jumbled words.
Private Sub cmdRandomize_Click()
txtinput.Text = ""
mainvalue = UCase(Text2.Text)

lbl:
Do While Len(Trim(Text2.Text)) <> 0

Randomize
r = Int((Len(Text2.Text)) * Rnd)

If r = 0 Then r = 1
txtinput.Text = Trim(txtinput.Text) & UCase(Mid(Text2.Text, r, 1))
Text2.Text = Replace(Text2.Text, Mid(Text2.Text, r, 1), "", 1, 1)
Loop

If txtinput.Text = mainvalue Then
Text2.Text = mainvalue
txtinput.Text = ""
GoTo lbl
Else
Text2.Text = mainvalue
End If

End Sub
' generates runtime check box graphical style according to the number of characters in the word
Private Sub Command1_Click()
On Error GoTo err:
Label1.Caption = "Word Game 2001 - By Rajesh Lal (razesh@hotmail.com) www.developerscolumn.com "
txtinput.Text = ""
Text2.Text = ""
opendatabase (App.Path & "/words.mdb")

Randomize
rno = Int(rswords.RecordCount * Rnd)
rswords.Move rno
Text2.Text = rswords.Fields(1)
cmdRandomize_Click
Number = Len(Trim(txtinput.Text))
If Command1.Caption = "&Reset" Then
For i = Check1.Count - 1 To 1 Step -1
Unload Check1(i)
Next i
End If

Command1.Caption = "&Reset"
For i = 1 To Number
Load Check1(Check1.Count)

'Check1(i).Caption = "A"
Check1(i).Width = 400

If i = 1 Then
Check1(i).Left = Check1(0).Left
Else
Check1(i).Left = Check1(i - 1).Left + Check1(i - 1).Width
End If
Check1(i).Top = Check1(0).Top + 700
Check1(i).Visible = True
Check1(i).Caption = ""
Next i

j = 1
For i = Number + 1 To 2 * Number

Load Check1(Check1.Count)
Check1(i).Width = 800
Check1(i).Height = Check1(0).Height * 2

Check1(i).Visible = True
Check1(i).Caption = "__"

If i = Number + 1 Then
Check1(i).Left = Check1(0).Left '+ Check1(i).Width
Else
Check1(i).Left = Check1(i - 1).Left + Check1(i - 1).Width
End If

j = j + 1

If i >= Number + 11 Then
Check1(i).Top = Check1(0).Top + Check1(i - 1).Height + 1400
If i = Number + 11 Then Check1(i).Left = Check1(0).Left
Else
Check1(i).Top = Check1(0).Top + 1400
End If
Check1(i).BackColor = vbWhite
DoEvents
Next i


flag = True
For i = Number + 1 To Number * 2
Check1(i).Value = 1
Check1(i).FontBold = True
Check1(i).FontSize = 20
DoEvents
Next i


random = Int((Rnd * Len(txtinput.Text)) + 1)
lblWord = lblWord & Mid(txtinput.Text, random, 1)


For i = 1 To Number

Check1(i).Caption = UCase(Mid(Trim(txtinput.Text), i, 1))

Check1(i).Value = 0
DoEvents
Next i


flag = False
Exit Sub
err:
MsgBox " Error Caused due to the following: " & err.Description & "  start the game again ", 16 + vbOKOnly, "Jumble1.1"
End Sub
' Refresh button to try again
Private Sub Command2_Click()
If Command1.Caption = "&Start" Then Exit Sub

flag = True

For i = 1 To Number
Check1(i).Caption = UCase(Mid(txtinput.Text, i, 1))
Check1(i).Value = 0
DoEvents
Next i

For i = Number + 1 To Number * 2
Check1(i).Value = 1
Check1(i).Caption = "__"
Check1(i).FontBold = True
DoEvents
Next i

flag = False
End Sub
' Exit the program
Private Sub Command3_Click()
Unload Me
End Sub
'Check whether to exit
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
i = MsgBox("Quit the Jumble game", 32 + vbYesNo, "Jumble1.1")
If i = vbNo Then
Cancel = True
End If
End Sub
