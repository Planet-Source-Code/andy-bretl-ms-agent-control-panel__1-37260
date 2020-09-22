VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form AgentForm 
   Caption         =   "Agent User Control"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbCharacters 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox cmbActions 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSpeak 
      Caption         =   "Speak"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtSpeak 
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label lblSpeak 
      Caption         =   "Say the following:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblAvailActions 
      Caption         =   "available Actions:"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblAvailCharacters 
      Caption         =   "available characters:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin AgentObjectsCtl.Agent msA 
      Left            =   6960
      Top             =   120
   End
End
Attribute VB_Name = "AgentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public activeC As IAgentCtlCharacter
Public strWinPath As String
Public bolIsLoaded As Boolean


Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
    "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long

Public Function getWinDir() As String

    Dim sBuffer As String
    Dim lRet As Long
    sBuffer = Space(255)
    lRet = GetWindowsDirectory(sBuffer, 255)
    getWinDir = Left$(sBuffer, lRet)

End Function

Private Sub cmbActions_Change()
    If Not Me.cmbActions.Text = "" Then
        Me.activeC.Play Me.cmbActions.Text
    End If
End Sub

Private Sub cmbActions_Click()
    If Not Me.cmbActions.Text = "" Then
        Me.activeC.Play Me.cmbActions.Text
    End If
End Sub

Private Sub cmbCharacters_Click()
    Me.loadCharacter (Me.cmbCharacters.Text)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub



Public Sub loadCharacter(CName As String)
    If Me.bolIsLoaded Then
        Me.removeCharacter
    End If
    msA.Characters.Load "C1", Me.strWinPath & CName & ".acs"
    Set activeC = msA.Characters("C1")
    showCharacter
    initList
    Me.bolIsLoaded = True
End Sub

Public Sub showCharacter()
    activeC.MoveTo 320, 320
    activeC.Show
End Sub

Public Sub removeCharacter()
    activeC.Play "Hide"
    Set activeC = Nothing
    msA.Characters.Unload "C1"
End Sub
Public Sub initList()
    cmbActions.Clear
    For Each AnimationName In msA.Characters("C1").AnimationNames
        Me.cmbActions.AddItem AnimationName
    Next
End Sub

Private Sub cmdSpeak_Click()
    If Not Me.txtSpeak.Text = "" Then
        Me.activeC.Speak "" & Me.txtSpeak.Text
    End If
End Sub

Private Sub Form_Load()
    Me.strWinPath = Me.getWinDir() & "\msagent\chars\"
    initCharacterList
    Me.bolIsLoaded = False
End Sub

Private Sub initCharacterList()
    Dim fso As New FileSystemObject
    Dim CharacterFolder As Folder
    Dim CharacterFile As File
    Set CharacterFolder = fso.GetFolder("" & Me.strWinPath)

    For Each CharacterFile In CharacterFolder.Files
        Me.cmbCharacters.AddItem Left(CharacterFile.Name, (Len(CharacterFile.Name) - 4))
    Next

End Sub
