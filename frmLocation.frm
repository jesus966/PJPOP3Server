VERSION 5.00
Begin VB.Form frmLocation 
   Caption         =   "Examinar"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmLocation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "Regresar"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.DirListBox dirDir 
      Height          =   2340
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRegresar_Click()
Unload Me
End Sub
Private Sub dirDir_Click()
frmUsuarios.txtDir.Text = dirDir.Parent
End Sub

Private Sub drvDrive_Change()
drvDrive.Drive = dirDir.Path
End Sub
