VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "PJPOP3Server (Personal Jesus POP3 Server)"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "PJSMTPServer"
      Height          =   2415
      Left            =   0
      TabIndex        =   8
      Top             =   3720
      Width           =   11055
      Begin MSWinsockLib.Winsock WS2 
         Left            =   10560
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   25
      End
      Begin VB.ListBox lstSMTP 
         Height          =   1620
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   11055
      End
      Begin MSWinsockLib.Winsock WS3 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   25
      End
      Begin VB.Label Label3 
         Caption         =   "Eventos:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   9000
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   110
   End
   Begin VB.CommandButton cmdClearEvents 
      Caption         =   "Limpiar Eventos"
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   11055
      Begin VB.ListBox lstEvents 
         Height          =   2205
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   8055
      End
      Begin MSComctlLib.TreeView trwUsuarios 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   4048
         _Version        =   393217
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8040
         Top             =   2280
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
               Picture         =   "frmMain.frx":0442
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Eventos:"
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios Registrados:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1111
      ButtonWidth     =   1349
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Iniciar"
            Description     =   "Iniciar el Servidor POP3"
            Object.ToolTipText     =   "Iniciar el Servidor POP3"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Usuarios"
            Description     =   "Ir a la pantalla de Usuarios "
            Object.ToolTipText     =   "Ir a la pantalla de Usuarios "
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Description     =   "Salir de PJPOP3Server"
            Object.ToolTipText     =   "Salir de PJPOP3Server"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Opciones"
            Description     =   "Opciones"
            Object.ToolTipText     =   "Opciones"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbEstado 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6225
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Estado: Parado"
            TextSave        =   "Estado: Parado"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "22:54"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "20/02/2008"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentConection As Long
Dim UserFileContent
Dim CurrentConectionPasa As Boolean
Event SendCommand(Command As String)
Dim UsuarioExiste As Boolean
Dim NombreDominio
Dim CurrentConection2
Dim Dats As String
Dim MailFrom As String
Dim Command, NombreUsuario, Password, MailBoxDirec, Usuario, Passw, Instruction, SeAcabo, i, IdCorreo, Eliminar, Coger, FF
Private Sub cmdClearEvents_Click()
lstEvents.Clear
End Sub

Private Sub Form_Load()
CargarUsuarios
AddEv "PJSMTPServer Iniciado Correctamente v: 1.0.0.0"
lstEvents.AddItem "Iniciado PJPOP3Server correctamente v: 1.0.0.0"
UsuarioExiste = False
CurrentConectionPasa = False
If ExisteArchivo(App.Path & "\" & "Opciones.ini") = True Then
FF = FreeFile
Dim Iniciar, Activar, IniciarSMTP, ActivarSMTP
Open App.Path & "\" & "Opciones.ini" For Input As #FF
Input #FF, NombreDominio
Input #FF, Iniciar
Input #FF, Activar
Input #FF, ActivarSMTP
Input #FF, IniciarSMTP
Close #FF
End If
If Iniciar = "1" Then
Toolbar1.Buttons(1).Caption = "Detener"
WS.LocalPort = 110
WS.Listen
stbEstado.Panels(1).Text = "Estado: Iniciado en " & WS.LocalIP
lstEvents.AddItem "Servidor POP3 Iniciado"
End If
If IniciarSMTP = "1" Then
If ActivarSMTP = "1" Then
Toolbar1.Buttons(1).Caption = "Detener"
WS2.Listen
stbEstado.Panels(1).Text = "Estado: Iniciado en " & WS.LocalIP
lstSMTP.AddItem "Servidor SMTP Iniciado"
End If
End If
If Activar = "0" Then
Frame1.Enabled = False
trwUsuarios.Enabled = False
Label1.Enabled = False
Label2.Enabled = False
cmdClearEvents.Enabled = False
lstEvents.Enabled = False
End If
If ActivarSMTP = "0" Then
Frame2.Enabled = False
Label3.Enabled = False
lstSMTP.Enabled = False
End If
If NombreDominio = "" Then NombreDominio = WS.LocalIP
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo errdetect
Select Case Button.Index
Case "1"
If lstEvents.Enabled = True Then
If Toolbar1.Buttons(1).Caption = "Iniciar" Then
Toolbar1.Buttons(1).Caption = "Detener"
WS.LocalPort = 110
WS.Listen
WS2.Listen
stbEstado.Panels(1).Text = "Estado: Iniciado en " & WS.LocalIP
lstEvents.AddItem "Servidor POP3 Iniciado"
AddEv "Servidor SMTP Iniciado"
Else
WS.Close
WS2.Close
WS.LocalPort = 0
stbEstado.Panels(1).Text = "Estado: Parado"
Toolbar1.Buttons(1).Caption = "Iniciar"
lstEvents.AddItem "Servidor POP3 Detenido"
AddEv "Servidor SMTP detenido"
End If
End If
Exit Sub
Case "2"
frmUsuarios.Show
Case "3"
End
Case "4"
frmOptions.Show
End Select
Exit Sub
errdetect:
Select Case Err.Number
Case "10048"
MsgBox "Error 10048: Los puertos 110 (POP3) o 25 (SMTP) Están siendos utilizados", vbCritical, "PJPOP3Server"
lstEvents.AddItem "Error 10048: El puerto 110 (POP3) Está siendo utilizado"
AddEv "Error 10048:El puerto 25 (SMTP) Está siendo utilizado"
Toolbar1.Buttons(1).Caption = "Iniciar"
Case Else
MsgBox Err.Description, vbCritical, Err.Number & " PJPOP3Server"
lstEvents.AddItem Err.Description
End Select
End Sub

Private Sub WS_ConnectionRequest(ByVal requestID As Long)
WS.Close
WS.Accept requestID
CurrentConection = requestID
lstEvents.AddItem "Solicitud de conexión " & requestID & " aceptada"
Dim Timedate
Timedate = Time + Date
Send "+OK PJPOP3Server Ready"
End Sub
Private Sub Send(Data As String)
If WS.State = 7 Then
    WS.SendData Data & vbCrLf
    RaiseEvent SendCommand(Data)
End If
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
Dim Datos As String
WS.GetData Datos 'Recoje todos los datos
Instruction = GetInstr(Datos)
Command = GetCommand(Datos) 'Recoje el comando POP3
Select Case UCase(Command) 'Busca si el comando enviado al servidor corresponde con algun comando POP3
Case "NOOP" 'Si el caso es NOOP mantener la conexion con prioridad a los demas eventos del servidor
If CurrentConectionPasa = False Then
lstEvents.AddItem "¡INTENTO DE USAR EL COMANDO NOOP SIN HABER COMPROBADO USUARIO Y CONTRASEÑA!"
Send "-ERR No has iniciado sesion"
Exit Sub
Else
DoEvents
AddEvent ("Usado Comando NOOP por " & CurrentConection)
Send "+OK"
End If
Case "AUTH" 'Muestra la informacion del programa servidor
Send "+OK PJPOP3Server version 1.0.0.0 (c) Personal Jesus"
AddEvent ("Usado Comando AUTH por" & CurrentConection)
Case "USER" 'Nombre de usuario
AddEvent (CurrentConection & " esta ingresando un nombre de Usuario...")
Usuario = Instruction
Dim D As String
If ExisteArchivo(App.Path & "\Users\" & Usuario & ".usu") = True Then
UsuarioExiste = True
D = Dir(App.Path & "\Users\" & Usuario & ".usu") 'Busca en la carpeta de usuarios
Dim FF As Integer
FF = FreeFile
AddEvent ("El usuario ingresado por " & CurrentConection & " existe, esperando a que ingrese el comando PASS")
Open App.Path & "\Users\" & Usuario & ".usu" For Input As #FF 'Recoje los datos del usuario y deja paso al comando PASS
Input #FF, NombreUsuario
Input #FF, Password
Input #FF, MailBoxDirec
Close #FF
Send "+OK " & Usuario & "ingresa tu password ahora"
Else 'En caso contrario el usuario no existe y envia un mensaje de error al programa cliente y cierra la conexión
AddEvent ("EL USUARIO INGRESADO POR " & CurrentConection & " NO EXISTE")
UsuarioExiste = False
Send "-ERR El Usuario No existe"
Exit Sub
End If
Case "PASS"
AddEvent (CurrentConection & " esta ingresando una contraseña...")
Passw = Instruction 'Recoje la contraseña enviada por el cliente POP3
If UsuarioExiste = True Then 'Si el Usuario existe comprobar contraseña en caso contrario mandar un mensaje de error y cerrar la conexión
If Password = Passw Then 'Si la contraseña es correcta permite utilizar todos los comandos de el servidor en caso contrario manda un mensaje de error y cierra la conexión
AddEvent ("Contraseña ingresada por " & CurrentConection & " aceptada, identificacion de usuario concluida")
CurrentConectionPasa = True
Send "+OK " & Usuario & "tu password es valido"
Else
AddEvent ("CONTRASEÑA INGRESADA POR " & CurrentConection & " NO ES VÁLIDA")
CurrentConectionPasa = False
Send "-ERR El password ingresado no es valido"
Exit Sub
End If
Else
AddEvent (CurrentConection & " NO A INGRESADO UN NOMBRE DE USUARIO")
CurrentConectionPasa = False
Send "-ERR Primero ingresa un nombre de usuario"
Exit Sub
End If
Case "LIST"
Dim Spa
Dim Tempv
Spa = Instruction
If CurrentConectionPasa = False Then
AddEvent ("INTENTO DE USAR LIST SIN HABER INGRESADO EN EL SISTEMA")
Send "-ERR Inicia la sesion primero"
Else
If Spa = "" Then
Send "+OK " & IdCorreo & " messages 0 bytes"
Send IdCorreo & " 0 bytes"
Send "."
Else
FF = FreeFile
If Spa > IdCorreo Then
Send "-ERR El identificador de correo no existe"
Else
Open MailBoxDirec & "\" & Spa For Input As #FF
Tempv = Input(LOF(FF), #FF)
Close #FF
Tempv = Len(Tempv)
Send "+OK " & Spa & " " & Tempv
Send "."
End If
End If
End If
Case "QUIT"
If CurrentConectionPasa = True Then
Kill (MailBoxDirec & "\Borrador\*.*")
Open MailBoxDirec & "\Borrador\index.dat" For Output As #1
Print #1, "Indentifiquer"
Close #1
AddEvent ("Conexión " & CurrentConection & " cerrada")
WS.Close
WS.Listen
Else
AddEvent ("Conexión " & CurrentConection & " cerrada")
WS.Close
WS.Listen
End If
Case "STAT"
If CurrentConectionPasa = False Then
AddEvent ("INTENTO DE USAR STAT SIN HABER INGRESADO EN EL SISTEMA")
Send "-ERR Inicia la sesion primero"
Else
If ExisteArchivo(MailBoxDirec & "\" & "1") = False Then
Send "0 0"
AddEvent "El usuario " & Usuario & " no tiene correo nuevo"
Exit Sub
Else
AddEvent "El usuario " & Usuario & " Tiene correo nuevo"
Dim Byt As Long
Dim cb As Long
Byt = 0
Do
For i = 1 To 100
If ExisteArchivo(MailBoxDirec & "\" & i) = True Then
FF = FreeFile
Open MailBoxDirec & "\" & i For Input As #FF
cb = Len(Input(LOF(FF), #FF))
Byt = Byt + cb
Close #FF
Else
SeAcabo = i
Exit Do
End If
Next i
Loop
IdCorreo = SeAcabo - 1
Send IdCorreo & " " & Byt
End If
End If
Case "HELP"
AddEvent CurrentConection & " a pedido uso del comando HELP"
Send "+OK       PJPOP3SERVER COPYRIGHT PERSONAL JESUS"
Send "+OK       Esta es la ayuda sobre los comandos de este servidor:"
Send "+OK       HELP Muestra la ayuda"
Send "+OK       NOOP Deja el servidor en modo de espera"
Send "+OK       USER Introduce el nombre de usuario"
Send "+OK       PASS Introduce la contrasena del usuario"
Send "+OK       AUTH Muestra informacion del sistema"
Send "+OK       STAT Muestra si tienes mensajes nuevos"
Send "+OK       LIST Pide informacion sobre los mensajes nuevos"
Send "+OK       QUIT Cierra la sesion con el servidor POP3"
Send "+OK       RETR Recibe el mensaje especificado"
Send "+OK       DELE Marca para Eliminar el mensaje especificado"
Send "+OK       RSET Recupera los mensajes marcados para eliminar"
Send "+OK       Para poder operar en el sistema deve ingresar un nombre de usuario mediante USER y PASS"
Case "RETR"
If CurrentConectionPasa = True Then
Coger = Instruction
FF = FreeFile
Dim arch As String
Open MailBoxDirec & "\" & Coger For Input As #FF
arch = Input(LOF(FF), #FF)
Close #FF
Dim tama
tama = Len(arch)
Send "+OK " & tama & " octets"
Send arch
Send "."
Else
Dim sFile
AddEvent "INTENTO DE USAR RETR SIN HABER INICIADO LA SESION"
Send "-ERR Inicia la sesion primero"
End If
Case "DELE"
If CurrentConectionPasa = True Then
Eliminar = Instruction
If ExisteArchivo(MailBoxDirec & "\" & Eliminar) = True Then
AddEvent Usuario & " esta eliminando un mensaje..."
If ExisteArchivo(MailBoxDirec & "\Borrador\Index.dat") = True Then
FileCopy MailBoxDirec & "\" & Eliminar, MailBoxDirec & "\Borrador\" & Eliminar
Kill MailBoxDirec & "\" & Eliminar
AddEvent Usuario & " a eliminado un mensaje"
Send "+OK Marcado para eliminar"
Else
AddEvent "No existe la carpeta Borrador"
Send "-ERR Error Interno en el Servidor"
End If
Else
AddEvent "El mensaje a eliminar no existe"
Send "-ERR El mensaje a eliminar no existe"
End If
Else
AddEvent "USO DE DELE SIN HABER INGRESADO EN EL SISTEMA"
Send "-ERR Inicia la sesion primero"
WS.Close
WS.Listen
End If
Case "RSET"
If CurrentConectionPasa = True Then
If ExisteArchivo(MailBoxDirec & "\Borrador\*.*") = True Then
FileCopy MailBoxDirec & "\Borrador\" & IdCorreo, MailBoxDirec & " \ """
RmDir (MailBoxDirec & "\Borrador")
MkDir (MailBoxDirec & "\Borrador")
Send "+OK Todos los mensajes recuperados"
AddEvent Usuario & " a recuperado los mensajes a eliminar"
Else
AddEvent "No tiene ningun mensaje marcado para eliminar"
Send "-ERR No tienes mensajes para eliminar"
End If

Else
AddEvent "INTENTO DE USAR RSET SIN HABER INICIADO LA SESION"
Send "-ERR Inicia la sesion primero"
End If
Case Else
AddEvent "NO EXISTE EL COMANDO ENVIADO POR " & CurrentConection
Send "-ERR El comando no existe, Para ayuda escriba HELP"
End Select
End Sub
Public Sub AddEvent(Descripcion As String)
lstEvents.AddItem Descripcion
End Sub
Public Function GetInstr(Data As String) As String
    On Local Error Resume Next
    Dim LenCommand As Integer
    LenCommand = InStr(Data, " ")
    If LenCommand Then
        GetInstr = Mid(Data, LenCommand + 1, InStr(Data, vbCrLf) - LenCommand - 1)
    End If
End Function
Public Function GetDirec(Data As String) As Variant

    On Local Error Resume Next
    Dim LenCommand As Integer
    LenCommand = InStr(Data, "<")
    If LenCommand Then
    Dim gtdirec
        gtdirec = Mid(Data, LenCommand + 1, InStr(Data, ">") - LenCommand - 1)
        GetDirec = vbCrLf + gtdirec + vbCrLf
    End If
End Function
Public Function SepDirec(Data As String) As Variant
On Local Error Resume Next
    Dim LenCommand As Integer
    LenCommand = InStr(Data, "@")
    If LenCommand Then
        SepDirec = Mid(Data, LenCommand + 1, InStr(Data, vbCrLf))
    End If
End Function

Public Function GetUsu(Data As String) As Variant
On Local Error Resume Next
    Dim LenCommand As Integer
    LenCommand = InStr(Data, vbCrLf)
    If LenCommand Then
        SpDirec = Mid(Data, LenCommand + 1, InStr(Data, "@"))
    End If
End Function
Private Function GetFile(Carpeta As String, Archivo As String)
Dim sArchivo As String
If Carpeta = "" Then
If Archivo = "/" Then GetFile = "/": Exit Function
        If ExisteArchivo(sArchivo) = True Then GetFile = sArchivo
     Else
        If ExisteArchivo(Carpeta & Archivo) = True Then
            GetFile = sArchivo
        Else
            If ExisteArchivo(sArchivo) = True Then GetFile = sArchivo
        End If
     End If
End Function
Public Function AddEv(texto As String)
lstSMTP.AddItem texto
End Function

Private Sub WS2_ConnectionRequest(ByVal requestID As Long)
WS2.Close
WS2.Accept requestID
CurrentConection2 = requestID
AddEv "Solicitud de conexión " & requestID & " aceptada"
Dim Timedate
Timedate = Time + Date
Sen "220 PJSMTPServer Preparado..."
End Sub
Private Sub Sen(Data As String)
If WS2.State = 7 Then
    WS2.SendData Data & vbCrLf
    RaiseEvent SendCommand(Data)
End If
End Sub

Private Sub WS2_DataArrival(ByVal bytesTotal As Long)
Dim instructions, commands
Dim CDom As String
WS2.GetData Dats
instructions = GetInstr(Dats)
commands = GetCommand(Dats)
Select Case UCase(commands)
Case "HELO"
Dim das As String
das = instructions
CDom = das
Sen "250 HOLA " & NombreDominio
AddEv "Uso del comando HELO, cliente identificado como " & das
Case "EHLO"
AddEv "Uso del comando EHLO, esta versión de servidor no soporta las extensiones SMTP"
Sen "500 Este servidor no soporta las extensiones SMTP"
Case "MAIL"
Dim MDir As String
Dim mdr
MDir = instructions
MailFrom = GetDirec(MDir)
Sen "250 OK, Datos almacenados, ingresa RCPT TO"
AddEv "Uso del comando MAIL, datos almacenados correctamente"
Case "NOOP"
Sen "250 OK, Entrado en modo de espera"
AddEv "Uso de NOOP"
Case "RSET"
MailFrom = ""
Sen "250 OK, datos reiniciados"
Case "RCPT"
If MailFrom = "" Then
Sen "500 Envia Mail From Primero"
AddEv "Uso de RCPT TO sin usar MAIL FROM primero"
End If
Dim UsuarioLocal As Boolean
Dim TDir As String
Dim TDirec As String
TDir = instructions
TDirec = GetDirec(TDir)
Dim Servidor
Servidor = SepDirec(TDirec)
If Servidor = NombreDominio Then
UsuarioLocal = True
Else
UsuarioLocal = False
End If

Case "QUIT"
Sen "220 Adios, tenga un buen dia"
WS2.Close
WS2.Listen
Case Else
AddEv "El cliente a enviado un comando que no se reconoce"
Sen "500 No se conoce este comando, para ver los comandos admitidos use HELP"
End Select
End Sub

