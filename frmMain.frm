VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ListBox Example"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "File Contents:"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3375
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2143
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":0000
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load File List Into ListBox"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click() 'Loads the file list into the ListBox
    Dim X As Integer
    Dim FileName As String, tempStr As String

    Call Command2_Click 'Reset everything
    
    FileName = App.Path & "\FileList.txt"
    X = 0
    
    Open FileName For Input As #1 'Open the list of files
        While Not EOF(1)
            Input #1, tempStr
            If Len(tempStr) > 0 Then
                List1.AddItem tempStr, X 'Add entries to the listbox
                X = X + 1
            End If
        Wend
    Close #1
End Sub

Private Sub Command2_Click() 'Resets the form
    RichTextBox1.FileName = ""
    RichTextBox1.Text = "<no file loaded>"
    List1.Clear
End Sub

Private Sub List1_Click()
    Dim FileName As String

    On Error GoTo ErrorHandler
    
    If Len(List1.Text) > 0 Then
        FileName = App.Path & "\" & List1.Text & ".txt"
        RichTextBox1.FileName = FileName 'Load file into RichTextBox
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Could not open file: '" & FileName & "'.", vbCritical + vbOKOnly, "Error"
End Sub
