VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Magic getallen raden"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   Icon            =   "FORM1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdNee 
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdJa 
      Caption         =   "&Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdAgain 
      Caption         =   "&Again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblVragen 
      Caption         =   "           Think at an number between 0 and  34."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Vraag As Integer
Dim Totaal As Integer


Private Sub cmdAgain_Click()
cmdJa.Visible = False
cmdNee.Visible = False
cmdAgain.Visible = False
cmdExit.Visible = False
cmdOK.Visible = True

lblVragen.Caption = "Think at an number between 0 and 34"
Vraag = 0
Totaal = 0




End Sub

Private Sub cmdExit_Click()
MsgBox "    Thanks for using my program." & vbCrLf & "             Also see my site " & vbCrLf & "         http://hsberry.cjb.net", , "Berry* says:"
End
End Sub


Private Sub cmdJa_Click()
Select Case Vraag
    Case 1
        Totaal = 1
        Call Vraag2
    Case 2
        Totaal = Totaal + 2
        Call Vraag3
    Case 3
        Totaal = Totaal + 3
        Call Vraag4
    Case 4
        Totaal = Totaal + 4
        Call Vraag5
    Case 5
        Totaal = Totaal + 8
        Call Vraag6
    Case 6
        Totaal = Totaal + 16
        lblVragen.Caption = "Your number was was: " & Totaal
        cmdJa.Visible = False
        cmdNee.Visible = False
        cmdAgain.Visible = True
        cmdExit.Visible = True
       
End Select
End Sub

Private Sub cmdNee_Click()
Select Case Vraag
    Case 1
        
        Call Vraag2
    Case 2
        
        Call Vraag3
    Case 3
        
        Call Vraag4
    Case 4
       
        Call Vraag5
    Case 5
       
        Call Vraag6
    Case 6
       lblVragen.Caption = "Your number was: " & Totaal
        cmdJa.Visible = False
        cmdNee.Visible = False
        cmdAgain.Visible = True
        cmdExit.Visible = True
        
       
   End Select
End Sub

Private Sub cmdOK_Click()
cmdOK.Visible = False
cmdJa.Visible = True
cmdNee.Visible = True


Call Vraag1
End Sub

Private Sub Vraag1()
Vraag = 1
lblVragen.Caption = ("Is you number in this row?" & vbCrLf & "1 3 4 9 17 21 25 29 32 34")
End Sub

Private Sub Vraag2()
Vraag = 2
lblVragen.Caption = ("Is you number in this row?" & vbCrLf & "2 3 5 6 10 13 14 18 22 26 30 33 34")
End Sub

Private Sub Vraag3()
Vraag = 3
lblVragen.Caption = ("Is you number in this row?" & vbCrLf & "4 5 7 11 13 15 19 23 27 31 32 33 34")

End Sub

Private Sub Vraag4()
Vraag = 4
lblVragen.Caption = ("Is you number in this row?" & vbCrLf & "6 7 12 14 15 20 21 22 23 28 29 30 31 32 33 34")
End Sub

Private Sub Vraag5()
lblVragen.Caption = ("Is you number in this row?" & vbCrLf & "8 9 10 11 12 13 14 15 24 25 26 27 28 29 30 31 32 33 34")
Vraag = 5
End Sub

Private Sub Vraag6()
Vraag = 6
lblVragen.Caption = ("Is you number in this row?" & vbCrLf & "16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 ")
End Sub

Private Sub Form_Load()
MsgBox "    Thanks for using my program." & vbCrLf & "             Also see my site " & vbCrLf & "         http://hsberry.cjb.net", , "Berry* says:"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "    Thanks for using my program." & vbCrLf & "             Also see my site " & vbCrLf & "         http://hsberry.cjb.net", , "Berry* says:"
End Sub
