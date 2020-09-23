VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Delete a file PERMANENTLY"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Destroy"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select File"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6015
   End
   Begin VB.Label Label7 
      Caption         =   "v1.2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   7095
   End
   Begin VB.Label Label6 
      Caption         =   "I am not responsible for any data loss or bad use of this source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   5655
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":00AE
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   " Peace"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7440
      TabIndex        =   5
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   " If you use it handle with care. You cannot recover file containts. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Source code submitted by GreekThought.  GreekThought@yahoo.gr"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This source is (C)opyrighted and submitted by GreekThought
'Mail :  GreekThought@yahoo.gr
'If you like/use this source please leave a comment for me too if you like

Private Sub Command1_Click()

With CommonDialog1
 .DialogTitle = "Please select a file to destroy"
 .InitDir = "c:"
 .ShowOpen
 Text1.Text = .FileName
End With

Command2.Visible = True

 
End Sub

Private Sub Command2_Click()

z_Recheck:

If Trim$(Text1.Text) = "" Then
    MsgBox "Please select a valid file!"
    Exit Sub
End If

If FileExist(Text1.Text) Then

 If ValidFileAttributes(Text1.Text) Then
 
    If Not DestroyFile(Text1.Text) Then
           MsgBox "Couldn't destroy file : " & vbCrLf & Text1.Text
           Exit Sub
       Else
           MsgBox "File : " & vbCrLf & Text1.Text & vbCrLf & "sucessfully destroyed."
                
                      If RenameAndKill(Text1.Text) Then
                            MsgBox "File was renamed and deleted succesfully"
                          Else
                            MsgBox "Coulnd't rename and delete the file"
                      End If

     End If

  Else
         NormalAttributes Text1.Text
         GoTo z_Recheck

  End If

Else
  MsgBox "File doesn't exist!"

End If



End Sub

