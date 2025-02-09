VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsPersona.cls
Private pId As Integer
Private pNombre As String
Private pApellido As String
Private pFechaNacimiento As Date

' Propiedades
Public Property Get Id() As Integer
    Id = pId
End Property

Public Property Let Id(value As Integer)
    pId = value
End Property

Public Property Get Nombre() As String
    Nombre = pNombre
End Property

Public Property Let Nombre(value As String)
    pNombre = value
End Property

Public Property Get Apellido() As String
    Apellido = pApellido
End Property

Public Property Let Apellido(value As String)
    pApellido = value
End Property

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = pFechaNacimiento
End Property

Public Property Let FechaNacimiento(value As Date)
    pFechaNacimiento = value
End Property

' Método para mostrar información
Public Sub MostrarInformacion()
    MsgBox "Nombre: " & pNombre & " " & pApellido & ", Fecha de Nacimiento: " & pFechaNacimiento
End Sub

