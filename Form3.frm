VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmuso 
   BackColor       =   &H8000000B&
   Caption         =   "Uso de reactivos"
   ClientHeight    =   8280
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15525
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   15525
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtnombre 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5640
      TabIndex        =   8
      Top             =   6240
      Width           =   3255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   12960
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form3.frx":B2C0
      OLEDBString     =   $"Form3.frx":B348
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Doctores"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13200
      TabIndex        =   6
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Usar"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13200
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtcantidadutilizar 
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox txtcantidad 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del reactivo:"
      BeginProperty Font 
         Name            =   "Yu Gothic Light"
         Size            =   48
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   960
      TabIndex        =   9
      Top             =   360
      Width           =   10695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quien usará el reactivo:"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   6240
      Width           =   6375
   End
   Begin VB.Label lblnombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del reactivo"
      BeginProperty Font 
         Name            =   "Yu Gothic Light"
         Size            =   48
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   0
      Top             =   360
      Width           =   10695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad a utilizar:"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   4800
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   4575
   End
End
Attribute VB_Name = "frmuso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If txtcantidadutilizar.Text = "" Then
        MsgBox "Ingrese una cantidad", vbInformation, "Laboratorios el Puente"
    Else
        If txtnombre.Text = "" Then
                MsgBox "Seleccione un nombre", vbInformation, "Laboratorios el Puente"
        Else
            'TablaRegistro_Uso
            'With rsRegistro
                '.Requery
                '.AddNew
                '!Doctor_ID = Val(txtnombre.Text)
                '!Identificador = Val(rsReactivos!Identificador)
                'rsRegistro.Fields("Doctor_ID") = txtnombre.Text
                'rsRegistro ("Doctor_ID") = val(
                'rsRegistro("Identificador") = Val(rsReactivos!Identificador)
                '.UpdateBatch
            'End With
            If Not (IsNumeric(txtcantidadutilizar.Text)) Then
                MsgBox "Ingrese la cantidad en números", vbInformation, "Laboratorios el Puente"
                txtcantidadutilizar.Text = ""
                txtcantidadutilizar.SetFocus
            Else
                Resultado = Val(txtcantidad.Text) - Val(txtcantidadutilizar.Text)
                rsReactivos.Fields("NúmeroReactivos") = Resultado
                rsReactivos.Update
                TablaRegistro_Uso
                rsRegistro.AddNew
                rsRegistro("Doctor_ID") = txtnombre.Text
                rsRegistro("Identificador") = lblnombre.Caption
                'rsRegistro("Doctor_ID") = txtnombre.Text
                'rsRegistro("Identificador") = lblnombre.Caption
                rsRegistro("Cantidad") = txtcantidadutilizar.Text
                rsRegistro.Update
                
                If rsReactivos.State = 1 Or rsReactivos.State = 0 Then
                    MsgBox "Cambios realizados", vbInformation, "Laboratorios el Puente"
                    frmbuscar.Show
                    Unload Me
                    
                Else
                    MsgBox "Ha ocurrido un error", vbInformation, "Laboratorios el Puente"
                End If
            End If
        End If
    End If
End Sub
Private Sub Command2_Click()
    If MsgBox("¿Desea volver a la selección de reactivos?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        Unload Me
        frmbuscar.Show
    End If
End Sub
Private Sub Form_Load()
    Label1.ForeColor = RGB(69, 110, 174)
    label2.ForeColor = RGB(69, 110, 174)
    Label3.ForeColor = RGB(69, 110, 174)
    txtnombre.Text = Usuario
    txtcantidad.Text = Cantidad
    lblnombre.Caption = Nombre
End Sub


