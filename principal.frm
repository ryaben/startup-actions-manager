VERSION 5.00
Object = "{C3967F87-FD47-4E87-B007-06264CBD1A36}#2.0#0"; "systray.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StartUp Actions Manager"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   Icon            =   "principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Acerca de StartUp Actions Manager"
      Height          =   5175
      Left            =   1680
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command3 
         Caption         =   "Información del sistema"
         Height          =   855
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   1695
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Text            =   "principal.frx":7EFA
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Rama Studios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   45
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 2.0.0 Multilenguaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   41
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "StartUp Actions Manager"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   40
         Top             =   480
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   1335
         Left            =   240
         Picture         =   "principal.frx":8029
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   5175
      Left            =   1680
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   2640
         TabIndex        =   31
         Top             =   720
         Width           =   2415
         Begin VB.CheckBox Check2 
            Caption         =   "¿Desea verificar en el inicio del programa que éste se abra con Windows?"
            Height          =   735
            Left            =   240
            TabIndex        =   42
            Top             =   840
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Verificar inicio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   240
         TabIndex        =   27
         Top             =   2520
         Width           =   2415
         Begin VB.Timer Timer2 
            Interval        =   100
            Left            =   120
            Top             =   840
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1680
            TabIndex        =   44
            Text            =   "1"
            Top             =   1320
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1680
            TabIndex        =   43
            Text            =   "0"
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   240
            TabIndex        =   35
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.OptionButton Option4 
            Caption         =   "English"
            Height          =   255
            Left            =   720
            TabIndex        =   29
            Top             =   1320
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Español"
            Height          =   255
            Left            =   720
            TabIndex        =   28
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Idioma de la interfaz"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1815
         Left            =   2640
         TabIndex        =   24
         Top             =   2520
         Width           =   2415
         Begin VB.CheckBox Check1 
            Height          =   255
            Left            =   1080
            TabIndex        =   25
            Top             =   1320
            Value           =   2  'Grayed
            Width           =   255
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Cerrar tras efectuar operaciones de inicio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   735
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   2415
         Begin VB.CommandButton Command2 
            Caption         =   "Borrar del registro"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Escribir en el registro"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Iniciar con Windows"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Configurar acciones"
      Height          =   5175
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   5295
      Begin MSComctlLib.ListView ListView1 
         Height          =   1095
         Left            =   2760
         TabIndex        =   46
         Top             =   3240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1931
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mensaje"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Estilo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2280
         Top             =   240
      End
      Begin VB.CommandButton XPButton15 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   2880
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton XPButton17 
         Caption         =   "Añadir"
         Height          =   495
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton XPButton11 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   360
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton XPButton13 
         Caption         =   "Añadir"
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton XPButton8 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   2880
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton XPButton10 
         Caption         =   "Añadir"
         Height          =   495
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton XPButton14 
         Caption         =   "Añadir"
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton XPButton6 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   360
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.ListBox List3 
         Height          =   1035
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   2175
      End
      Begin VB.ListBox List2 
         Height          =   1035
         ItemData        =   "principal.frx":9041
         Left            =   2760
         List            =   "principal.frx":9043
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   3000
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar mensajes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Abrir archivos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Abrir sitios web"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Abrir carpetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5115
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton XPButton4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cerrar SUAM"
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton XPButton3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Opciones"
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton XPButton2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Acerca de ..."
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton XPButton1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Configurar acciones"
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin IconSystray.sysTray sysTray1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1799
         _ExtentY        =   1799
         ToolTipText     =   "StartUp Actions Manager"
         IconPicture     =   "principal.frx":9045
      End
   End
End
Attribute VB_Name = "principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1

' Opciones de seguridad de clave del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos ROOT de clave del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Const APPLICATION As String = "StartUp Actions Manager"

Dim opcionespañol As Single
Dim opcioningles As Single
Dim verificaciondeinicio As String

Dim Path_Archivo_Ini As String

'Función api que recupera un valor-dato de un archivo Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Intentar obtener ruta de acceso y nombre del programa de Info. del sistema a partir del Registro...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Intentar obtener sólo ruta del programa de Info. del sistema a partir del Registro...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validar la existencia de versión conocida de 32 bits del archivo
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error: no se puede encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If
    ' Error: no se puede encontrar la entrada del Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "La información del sistema no está disponible en este momento", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' Código de retorno
    Dim hKey As Long                                        ' Controlador de una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de datos de una clave de Registro
    Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tamaño de variable de clave de Registro
    '------------------------------------------------------------
    ' Abrir clave de registro bajo KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Error de controlador...
    
    tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
    KeyValSize = 1024                                       ' Marcar tamaño de variable
    
    '------------------------------------------------------------
    ' Obtener valor de clave de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtener o crear valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agregar cadena terminada en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Encontrado valor nulo, se va a quitar de la cadena
    Else                                                    ' En WinNT las cadenas no terminan en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se ha encontrado valor nulo, sólo se va a extraer la cadena
    End If
    '------------------------------------------------------------
    ' Determinar tipo de valor de clave para conversión...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Buscar tipos de datos...
    Case REG_SZ                                             ' Tipo de datos String de clave de Registro
        KeyVal = tmpVal                                     ' Copiar valor de cadena
    Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
        For i = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor carácter a carácter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a cadena
    End Select
    
    GetKeyValue = True                                      ' Se ha devuelto correctamente
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Borrar después de que se produzca un error...
    KeyVal = ""                                             ' Establecer valor a cadena vacía
    GetKeyValue = False                                     ' Fallo de retorno
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
End Function

'Lee un dato _
-----------------------------
'Recibe la ruta del archivo, la clave a leer y _
 el valor por defecto en caso de que la Key no exista
Private Function Leer_Ini(Path_INI As String, Key As String, Default As Variant) As String

Dim bufer As String * 256
Dim Len_Value As Long

        Len_Value = GetPrivateProfileString(APPLICATION, _
                                         Key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
        
        Leer_Ini = Left$(bufer, Len_Value)

End Function

'Escribe un dato en el INI _
-----------------------------
'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave

Private Function Grabar_Ini(Path_INI As String, Key As String, Valor As Variant) As String

    WritePrivateProfileString APPLICATION, _
                                         Key, _
                                         Valor, _
                                         Path_INI

End Function

Public Sub EjecutarArchivos(ruta As String)
Dim ejecutarShell As Variant
On Error GoTo errsub
ejecutarShell = Shell("rundll32.exe url.dll,FileProtocolHandler " & (ruta), 1)
Exit Sub
errsub: MsgBox Err.Description, vbCritical
End Sub

Private Sub AbrirArchivo1()
Dim LineaTexto As String
Open App.Path & "\carpetas.lst" For Input As #1

While Not EOF(1)
Line Input #1, LineaTexto
List1.AddItem LineaTexto
Wend

Close #1
End Sub

Private Sub GuardarArchivo1()
Dim i As Integer

Open App.Path & "\carpetas.lst" For Output As #1

If List1.ListCount > -1 Then
For i = 0 To List1.ListCount - 1
Print #1, List1.List(i)
Next i
Close #1
End If
End Sub

Private Sub AbrirArchivo2()
Dim LineaTexto As String
Open App.Path & "\sitiosweb.lst" For Input As #1

While Not EOF(1)
Line Input #1, LineaTexto
List2.AddItem LineaTexto
Wend

Close #1
End Sub

Private Sub GuardarArchivo2()
Dim i As Integer

Open App.Path & "\sitiosweb.lst" For Output As #1

If List2.ListCount > -1 Then
For i = 0 To List2.ListCount - 1
Print #1, List2.List(i)
Next i
Close #1
End If
End Sub

Private Sub AbrirArchivo3()
Dim LineaTexto As String
Open App.Path & "\archivos.lst" For Input As #1

While Not EOF(1)
Line Input #1, LineaTexto
List3.AddItem LineaTexto
Wend

Close #1
End Sub

Private Sub GuardarArchivo3()
Dim i As Integer

Open App.Path & "\archivos.lst" For Output As #1

If List3.ListCount > -1 Then
For i = 0 To List3.ListCount - 1
Print #1, List3.List(i)
Next i
Close #1
End If
End Sub

Private Sub AbrirArchivo4()
Dim LineaTexto As String
Open App.Path & "\mensajes.lst" For Input As #1

While Not EOF(1)
Line Input #1, LineaTexto
List4.AddItem LineaTexto
Wend

Close #1
End Sub

Private Sub GuardarArchivo4()
Dim i As Integer

Open App.Path & "\mensajes.lst" For Output As #1

If List4.ListCount > -1 Then
For i = 0 To List4.ListCount - 1
Print #1, List4.List(i)
Next i
Close #1
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next

Dim El_Objeto As Object
Set El_Objeto = CreateObject("WScript.Shell")

El_Objeto.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\StartUp Actions Manager", App.Path & "\" & App.EXEName & ".exe"
End Sub

Private Sub Command2_Click()
On Error Resume Next

Dim El_Objeto As Object
Set El_Objeto = CreateObject("WScript.Shell")

El_Objeto.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\StartUp Actions Manager"
End Sub

Private Sub Command3_Click()
Call StartSysInfo
End Sub

Private Sub Form_Load()
On Error Resume Next

Dim El_Objeto As Object
Set El_Objeto = CreateObject("WScript.Shell")
Dim Resultado As String
Dim fso As Object
Dim X As Integer
Dim listtotal1 As Integer
Dim listtotal2 As Integer
Dim listtotal3 As Integer
Dim listtotal4 As Integer

'Path del fichero Ini
    Path_Archivo_Ini = App.Path & "\config.ini"
    
    ' Lee las Key y  Les envia el valor por defecto por si no existe
    opcionespañol = Leer_Ini(Path_Archivo_Ini, "Español", 0)
    opcioningles = Leer_Ini(Path_Archivo_Ini, "Ingles", 1)
    verificaciondeinicio = Leer_Ini(Path_Archivo_Ini, "Inicio", 1)
    
If App.PrevInstance = True Then
If opcionespañol = 1 Then
MsgBox "StartUp Actions Manager ya se encuentra abierto."
End
ElseIf opcioningles = 1 Then
MsgBox "StartUp Actions Manager is already open."
End
End If
End If
    
    'Posiciona el formulario con los valores del archivo Ini
    Text3.Text = opcionespañol
    Text4.Text = opcioningles
    If Text3.Text = "1" Then
    Option3.Value = True
    Else
    Option3.Value = False
    End If
    If Text4.Text = "1" Then
    Option4.Value = True
    Else
    Option4.Value = False
    End If
    Check2.Value = verificaciondeinicio

Resultado = El_Objeto.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\StartUp Actions Manager")  'nombrejecutable deve ser el nombre que recivira el ejecutable

If Resultado = "" Then
If Check2.Value = 1 Then
If Option3.Value = True Then
If MsgBox("StartUp Actions Manager no se inicia con Windows. ¿Desea escribir en el registro su inicio automático?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
El_Objeto.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\StartUp Actions Manager", App.Path & "\" & App.EXEName & ".exe" 'lo mismo con nombrejecutable
End If
End If
If Option4.Value = True Then
If MsgBox("StartUp Actions Manager doesn't start on Windows startup. Do you want to write an automatic startup registry key?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
El_Objeto.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\StartUp Actions Manager", App.Path & "\" & App.EXEName & ".exe" 'lo mismo con nombrejecutable
End If
End If
End If
End If
Set El_Objeto = Nothing

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(App.Path & "\carpetas.lst") Then
Call AbrirArchivo1
Else
Call GuardarArchivo1
End If
If fso.FileExists(App.Path & "\sitiosweb.lst") Then
Call AbrirArchivo2
Else
Call GuardarArchivo2
End If
If fso.FileExists(App.Path & "\archivos.lst") Then
Call AbrirArchivo3
Else
Call GuardarArchivo3
End If
If fso.FileExists(App.Path & "\mensajes.lst") Then
Call AbrirArchivo4
Else
Call GuardarArchivo4
End If

listtotal1 = List1.ListCount

If listtotal1 > 0 Then
For X = 0 To listtotal1 - 1
Call EjecutarArchivos(List1.List(X))
Next
End If

listtotal2 = List2.ListCount

If listtotal2 > 0 Then
For X = 0 To listtotal2 - 1
Dim Z
Z = ShellExecute(Me.hwnd, "Open", List2.List(X), &O0, &O0, SW_NORMAL)
Next
End If

listtotal3 = List3.ListCount

If listtotal3 > 0 Then
For X = 0 To listtotal3 - 1
Call EjecutarArchivos(List3.List(X))
Next
End If

listtotal4 = List4.ListCount

If listtotal4 > 0 Then
For X = 0 To listtotal4 - 1
MsgBox List4.List(X), vbInformation
Next
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
sysTray1.PonerSystray
Me.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Escribe en el archivo Ini
    
    'Posición del Form
    Call Grabar_Ini(Path_Archivo_Ini, "Español", Text3.Text)
    Call Grabar_Ini(Path_Archivo_Ini, "Ingles", Text4.Text)
    Call Grabar_Ini(Path_Archivo_Ini, "Inicio", Check2.Value)
End Sub

Private Sub Label12_Click()
Dim Z
Z = ShellExecute(Me.hwnd, "Open", "http://adf.ly/Kk0PI", &O0, &O0, SW_NORMAL)
End Sub

Private Sub sysTray1_DblClick(Button As Integer)
Me.WindowState = vbNormal
Me.Visible = True
sysTray1.RemoverSystray
End Sub

Private Sub Timer2_Timer()

If Option4.Value = True Then
Frame1.Caption = "Configure StartUp Actions"
Frame2.Caption = "Options"
Frame7.Caption = "About StartUp Actions Manager"
XPButton1.Caption = "Configure StartUp Actions"
XPButton3.Caption = "Options"
XPButton2.Caption = "About ..."
XPButton4.Caption = "Close SUAM"
Label5.Caption = "Windows Startup"
Command1.Caption = "Write key in registry"
Command2.Caption = "Delete key from registry"
Label8.Caption = "Verify startup"
Check2.Caption = "Do you want the program to check at start the automatic startup with Windows?"
Label7.Caption = "Interface language"
Label6.Caption = "Close program after startup actions resolved"
Label1.Caption = "Open folders"
Label2.Caption = "Open websites"
Label3.Caption = "Open files"
Label4.Caption = "Show messages"
XPButton6.Caption = "Delete"
XPButton8.Caption = "Delete"
XPButton11.Caption = "Delete"
XPButton15.Caption = "Delete"
XPButton14.Caption = "Add"
XPButton10.Caption = "Add"
XPButton13.Caption = "Add"
XPButton17.Caption = "Add"
Label11.Caption = "1.3.0 Version Multi-language"
Command3.Caption = "System Information"
Text2.Text = "Warning: StartUp Actions Manager© and all their private components are rights reserved, Ramiro Yaben© 2013 at Safe Creative©. This software its for free and open use, with extra unlockable functions that doesn't involve money of any kind of financial information."
Text3.Text = "0"
Text4.Text = "1"
End If
If Option3.Value = True Then
Frame1.Caption = "Configurar acciones"
Frame2.Caption = "Opciones"
Frame7.Caption = "Acerca de StartUp Actions Manager"
XPButton1.Caption = "Configurar acciones"
XPButton3.Caption = "Opciones"
XPButton2.Caption = "Acerca de ..."
XPButton4.Caption = "Cerrar SUAM"
Label5.Caption = "Iniciar con Windows"
Command1.Caption = "Escribir en el registro"
Command2.Caption = "Borrar del registro"
Label8.Caption = "Verificar inicio"
Check2.Caption = "¿Desea verificar en el inicio del programa que éste se abra con Windows?"
Label7.Caption = "Idioma de la interfaz"
Label6.Caption = "Cerrar tras efectuar operaciones de inicio"
Label1.Caption = "Abrir carpetas"
Label2.Caption = "Abrir sitios web"
Label3.Caption = "Abrir archivos"
Label4.Caption = "Mostrar mensajes"
XPButton6.Caption = "Eliminar"
XPButton8.Caption = "Eliminar"
XPButton11.Caption = "Eliminar"
XPButton15.Caption = "Eliminar"
XPButton14.Caption = "Añadir"
XPButton10.Caption = "Añadir"
XPButton13.Caption = "Añadir"
XPButton17.Caption = "Añadir"
Label11.Caption = "Versión 1.3.0 Multilenguaje"
Command3.Caption = "Información del sistema"
Text2.Text = "Advertencia: StartUp Actions Manager© y todos sus componentes privados son derechos reservados de Ramiro Yaben© 2013 en Safe Creative©. Este software es de uso libre y gratuito, con funcionalidades extras desbloqueables que no implican compra alguna ni descripción de datos financieros."
Text3.Text = "1"
Text4.Text = "0"
End If
End Sub

Private Sub XPButton1_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame7.Visible = False
End Sub

Private Sub XPButton10_Click()
Dim webstring As String
If Option3.Value = True Then
webstring = InputBox("Seleccione un sitio web para abrir en el inicio (con http://)", "Abrir sitios web")
If webstring <> "" Then
List2.AddItem webstring
Call GuardarArchivo2
End If
End If
If Option4.Value = True Then
webstring = InputBox("Choose a website to open at startup (with http://)", "Open websites")
If webstring <> "" Then
List2.AddItem webstring
Call GuardarArchivo2
End If
End If
End Sub

Private Sub XPButton11_Click()
If List3.ListIndex <> -1 Then
'Eliminamos el elemento que se encuentra seleccionado
List3.RemoveItem List3.ListIndex
Call GuardarArchivo3
End If
End Sub

Private Sub XPButton13_Click()
With CommonDialog1
If Option3.Value = True Then
.DialogTitle = "Abrir archivos"
.FileName = ""
.Filter = "Todos los archivos|*.*"
End If
If Option4.Value = True Then
.DialogTitle = "Open files"
.FileName = ""
.Filter = "All files|*.*"
End If
.ShowOpen

If .FileName <> "" Then
List3.AddItem .FileName
Call GuardarArchivo3
End If
End With
End Sub

Private Sub XPButton14_Click()
Dim carpetastring As String
If Option3.Value = True Then
carpetastring = InputBox("Introduzca una ruta de carpeta válida para ejecutar en el inicio", "Abrir carpetas")
If carpetastring <> "" Then
List1.AddItem carpetastring
Call GuardarArchivo1
End If
End If
If Option4.Value = True Then
carpetastring = InputBox("Copy a valid folder route to execute at startup", "Open folders")
If carpetastring <> "" Then
List1.AddItem carpetastring
Call GuardarArchivo1
End If
End If
End Sub

Private Sub XPButton15_Click()
If List4.ListIndex <> -1 Then
'Eliminamos el elemento que se encuentra seleccionado
List4.RemoveItem List4.ListIndex
Call GuardarArchivo4
End If
End Sub

Private Sub XPButton17_Click()
Dim mensajestring As String
If Option3.Value = True Then
mensajestring = InputBox("Seleccione el mensaje para mostrar en el inicio", "Mostrar mensajes")
List4.AddItem mensajestring
Call GuardarArchivo4
End If
If Option4.Value = True Then
mensajestring = InputBox("Write a message to show at startup", "Show messages")
List4.AddItem mensajestring
Call GuardarArchivo4
End If
End Sub

Private Sub XPButton2_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame7.Visible = True
End Sub

Private Sub XPButton3_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame7.Visible = False
End Sub

Private Sub XPButton4_Click()
End
End Sub

Private Sub XPButton6_Click()
If List1.ListIndex <> -1 Then
'Eliminamos el elemento que se encuentra seleccionado
List1.RemoveItem List1.ListIndex
Call GuardarArchivo1
End If
End Sub

Private Sub XPButton8_Click()
If List2.ListIndex <> -1 Then
'Eliminamos el elemento que se encuentra seleccionado
List2.RemoveItem List2.ListIndex
Call GuardarArchivo2
End If
End Sub
