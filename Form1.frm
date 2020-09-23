VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Delete Key Value"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Create Key"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete Key"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create Key Value"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Registry Value"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Registry Value"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cReg As clsRegistry

Private Sub Command1_Click()
Dim i As Long
Set cReg = New clsRegistry
i = cReg.QueryKeyValue("Network", "RestoreDiskChecked", HKEY_CURRENT_USER)
If i = 0 Then 'found!
   MsgBox "Found: " & Str(cReg.LongKeyValue) & vbCrLf & _
     "Key Handle: " & cReg.KeyHandle
Else
   MsgBox "Not found."
End If
Set cReg = Nothing
End Sub

Private Sub Command2_Click()
Dim i As Long
Set cReg = New clsRegistry
i = cReg.SetKeyValue("Network", "RestoreDiskChecked", 1, REG_DWORD, HKEY_CURRENT_USER)
If i = 0 Then
   MsgBox "Write was successful."
Else
   MsgBox "Write not successful."
End If
Set cReg = Nothing
End Sub

Private Sub Command3_Click()
Dim i As Long
Set cReg = New clsRegistry
i = cReg.QueryKeyValue("Network", "Temp", HKEY_CURRENT_USER)
If i = 0 Then 'found!
   MsgBox "Key already exists."
   Exit Sub
End If
i = cReg.SetKeyValue("Network", "Temp", 1, REG_DWORD, HKEY_CURRENT_USER)
If i = 0 Then
   MsgBox "New key value created."
Else
   MsgBox "Couldn't create key value."
End If
Set cReg = Nothing
End Sub

Private Sub Command4_Click()
Dim i As Long
Set cReg = New clsRegistry
i = cReg.DeleteKey(HKEY_CURRENT_USER, "Network\Temp")
If i = 0 Then
   MsgBox "Key deleted."
Else
   MsgBox "Key not found."
End If
Set cReg = Nothing
End Sub

Private Sub Command5_Click()
Dim i As Long
Set cReg = New clsRegistry
i = cReg.CreateKey("Network\", "Temp", HKEY_CURRENT_USER)
If i = 0 Then
   MsgBox "New key created."
Else
   MsgBox "Couldn't create key."
End If
Set cReg = Nothing
End Sub

Private Sub Command6_Click()
Dim i As Long
Dim n As Long
Set cReg = New clsRegistry
i = cReg.QueryKeyValue("Network", "Temp", HKEY_CURRENT_USER)
If i <> 0 Then
   MsgBox "Key doesn't exist."
   Exit Sub
End If
n = cReg.DeleteKeyValue("Network", "Temp", HKEY_CURRENT_USER)
If n = 0 Then
   MsgBox "Key value deleted."
Else
   MsgBox "Couldn't delete key value."
End If
Set cReg = Nothing
End Sub

