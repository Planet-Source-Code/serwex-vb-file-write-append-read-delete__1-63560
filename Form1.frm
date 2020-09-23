VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "4 Operation in File, By_ SerweX"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Delete File"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Append"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   "15"
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   "5"
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read -->"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "To Integer#:"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "From Interger#:"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   4200
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   4200
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmMain
' DateTime  : 07/12/2005 09:51
' Author    : Shahin Noursalhi
' Contact   : admin@MixofTix.net
' Purpose   : Demonstrate 4 normal operation in file!
'---------------------------------------------------------------------------------------

Private Sub Command1_Click()
'Write method (this will over write)
Dim intFnum As Integer
intFnum = FreeFile
Open "C:\MixofTix.txt" For Output As #intFnum
For i = CSng(Text2.Text) To CSng(Text3.Text)
Print #intFnum, i
Next
Close #intFnum
End Sub

Private Sub Command2_Click()
'Read method (this will read data from file)
Dim intFnum As Integer
If (Dir("c:\MixofTix.txt")) = "MixofTix.txt" Then
intFnum = FreeFile
Open "C:\MixofTix.txt" For Input As #intFnum
Text1.Text = ""
Do
Input #intFnum, jj
Text1.Text = Text1.Text & jj & vbCrLf
Loop Until EOF(intFnum)
Close #intFnum
Else
a = MsgBox("File Not Found!", vbExclamation, "Alert")
End If
End Sub

Private Sub Command3_Click()
'Append method (this will attach to the end of file)
Dim intFnum As Integer
intFnum = FreeFile
Open "C:\MixofTix.txt" For Append As #intFnum
For i = CSng(Text2.Text) To CSng(Text3.Text)
Print #intFnum, i
Next
Close #intFnum
End Sub

Private Sub Command4_Click()
'Delete method (this will delete the file!!!)
If (Dir("c:\MixofTix.txt")) = "MixofTix.txt" Then
Kill ("c:\MixofTix.txt")
Else
a = MsgBox("File Not Found!", vbExclamation, "Alert")
End If
End Sub

