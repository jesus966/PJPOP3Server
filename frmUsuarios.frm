VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuarios 
   Caption         =   "Lista de Usuarios de el servidor POP3"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "frmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "Regresar "
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdEraseUser 
      Caption         =   "Borrar Usuario"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "Añadir Usuario"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Text            =   "\"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de Usuarios POP3"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin MSComctlLib.TreeView trwUsuarios 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   5318
         _Version        =   393217
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsuarios.frx":014A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Directorio del Buzón de correo:"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Usuario:"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblUserName 
      Caption         =   "Nombre de Usuario:"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim US As Usuarios
Private Sub cmdAddUser_Click()
Dim tNodo As Node
Dim sP As String, sH As String
Dim i As Long
Set tNodo = trwUsuarios.SelectedItem
sP = tNodo.Key
i = tNodo.Children
On Error Resume Next
Do
i = i + 1
sH = sP & "-" & CStr(i)
Err = 0
US.Pass = txtPass.Text
US.UserName = txtUserName.Text
US.MailBoxDir = txtDir.Text
If ExisteArchivo(txtDir.Text & "\Borrador") = True Then
Dim FF As Integer
FF = FreeFile
Open App.Path & "\Users\" & txtUserName.Text & ".usu" For Output As #FF
Print #FF, US.UserName
Print #FF, US.Pass
Print #FF, US.MailBoxDir
Close #FF
Else
MkDir txtDir.Text & "\Borrador"
End If
trwUsuarios.Nodes.Add sP, tvwChild, sH, txtUserName.Text, 1
If Err = 0 Then Exit Do
Loop
End Sub

Private Sub cmdEraseUser_Click()
If trwUsuarios.SelectedItem.Index = 1 Then Exit Sub
Kill App.Path & "\Users\" & trwUsuarios.SelectedItem.Text & ".usu"
trwUsuarios.Nodes.Clear
Call CargarUsuarios
End Sub


Private Sub cmdRefresh_Click()
Me.Refresh
End Sub

Private Sub cmdRegresar_Click()
Unload Me
End Sub

Private Sub cmdSaveUser_Click()
GuardarUsuarios
End Sub

Private Sub Form_Load()
trwUsuarios.LabelEdit = tvwAutomatic
trwUsuarios.Nodes.Add , , "Parent", "Usuarios", 1
CargarUsuarios
End Sub
Private Sub CargarUsuarios()
Dim A As String
trwUsuarios.Nodes.Clear
trwUsuarios.Nodes.Add , , "Parent", "Usuarios", 1
A = Dir(App.Path & "\Users\")
Do While A <> ""
trwUsuarios.Nodes.Add "Parent", tvwChild, , Left(A, InStrRev(A, ".") - 1), 1
A = Dir
Loop
trwUsuarios.Nodes(trwUsuarios.Nodes.Count).EnsureVisible
trwUsuarios.Nodes(1).Selected = True
End Sub
Private Sub GuardarUsuarios()
If trwUsuarios.SelectedItem.Index = 1 Then Exit Sub
US.Pass = txtPass.Text
US.UserName = txtUserName.Text
US.MailBoxDir = txtDir.Text
Dim FF As Integer
FF = FreeFile
Open App.Path & "\Users\" & trwUsuarios.SelectedItem.Text & ".usu" For Random As #FF Len = Len(US)
Put #FF, , US
Close #FF
End Sub
