VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Opciones"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "Regresar"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdAcept 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "PJSMTPServer"
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   5655
      Begin VB.CheckBox chkActivarSMTP 
         Caption         =   "Activar el Servidor SMTP"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Value           =   1  'Checked
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PJPOP3Server"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox chkActive 
         Caption         =   "Activar el Servidor POP3"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin VB.CheckBox chkIniciar 
         Caption         =   "Iniciar el servidor al entrar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   5295
      End
      Begin VB.TextBox txtDomainName 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Text            =   "localhost"
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de Dominio:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAcept_Click()
FF = FreeFile
Open App.Path & "\" & "Opciones.ini" For Output As #FF
Print #1, txtDomainName.Text
Print #1, chkIniciar.Value
Print #1, chkActive.Value
Print #1, chkActivarSMTP.Value
Close #FF
End Sub

Private Sub cmdRegresar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Iniciar, Activar, IniciarSMTP, ActivarSMTP
If ExisteArchivo(App.Path & "\" & "Opciones.ini") = True Then
FF = FreeFile
Open App.Path & "\" & "Opciones.ini" For Input As #FF
Input #FF, NombreDominio
Input #FF, Iniciar
Input #FF, Activar
Input #FF, ActivarSMTP
Close #FF
txtDomainName.Text = NombreDominio
If Iniciar = "1" Then chkIniciar.Value = vbChecked
If Activar = "0" Then chkActive.Value = vbUnchecked
If ActivarSMTP = "0" Then chkActivarSMTP.Value = vbUnchecked
End If
End Sub
