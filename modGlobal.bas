Attribute VB_Name = "modGlobal"
Public Type Usuarios
UserName As String * 50
Pass As String * 50
MailBoxDir As String
End Type
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
Public Function GetCommand(Data As String) As String
    On Local Error Resume Next
    Dim Delimiter As Integer
    Delimiter = InStr(Data, " ")
    If Delimiter Then
        GetCommand = Left(Data, Delimiter - 1)
    Else
        GetCommand = Left(Data, InStr(Data, vbCrLf) - 1)
    End If
End Function
Public Function ExisteArchivo(ByVal Archivo As String) As Boolean
Dim L As Integer
On Error Resume Next
L = Len(Dir(Archivo))
If Err Or L = 0 Then
ExisteArchivo = False
Else
ExisteArchivo = True
End If
End Function
Function pn(texto As String) As String
Dim i As Long, aux As String, s As String
aux = ""
For i = 1 To Len(texto)
s = Hex(Asc(Mid(texto, i, 1)))
If Len(s) = 1 Then s = "0" & s
aux = aux + s
Next i
pn = aux
End Function
Function pt(numeros As String) As String
Dim i As Long, aux As String
aux = ""
For i = 1 To Len(numeros) Step 2
aux = aux + Chr("&h" & Mid(numeros, i, 2))
Next i
pt = aux
End Function
