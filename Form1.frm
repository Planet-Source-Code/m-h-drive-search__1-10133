VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DriveSearch"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Get Drives"
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   3465
      Width           =   4005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   3885
      Width           =   4005
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   4005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "D r i v e s   Found"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   4005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    End
End Sub

Private Sub Command2_Click()
Dim nDrive, cDrive, DriveType As String
    List1.Clear
    List1.AddItem "Drive Name" & vbTab & "Drive Type"
    List1.AddItem ""
    For i = 65 To 90
    nDrive = Chr(i) & ":\"
    cDrive = GetDriveType(nDrive)
    Select Case cDrive
    Case 2
        DriveType = "Floppy Disk Drive"
    Case 3
        DriveType = "Primary/Logical Disk Drive"
    Case 4
        DriveType = "Network Drive"
    Case 5
        DriveType = "CD-ROM"
    End Select
    If cDrive <> 1 Then
        List1.AddItem nDrive & vbTab & vbTab & DriveType
    Else: End If
    Next i
End Sub

