VERSION 5.00
Begin VB.Form frmdatos 
   BackColor       =   &H8000000E&
   Caption         =   "Administrar usuarios"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   Picture         =   "frmdatos.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      ItemData        =   "frmdatos.frx":B2C0
      Left            =   4080
      List            =   "frmdatos.frx":B2CA
      TabIndex        =   10
      Top             =   6720
      Width           =   3495
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
      Height          =   570
      Left            =   9480
      TabIndex        =   9
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Añadir"
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
      Left            =   9480
      TabIndex        =   8
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtedad 
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4080
      TabIndex        =   7
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox txtcontraseña 
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4080
      TabIndex        =   6
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox txtnombre 
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4080
      TabIndex        =   5
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   1725
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   2430
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de cuenta:"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   240
      TabIndex        =   3
      Top             =   6720
      Width           =   3105
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edad:"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administrar usuarios"
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
      Height          =   1230
      Left            =   6240
      TabIndex        =   0
      Top             =   240
      Width           =   8445
   End
End
Attribute VB_Name = "frmdatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If txtnombre.Text = "" Then
        MsgBox "Ingrese un nombre"
        txtnombre.SetFocus
    Else
        If txtcontraseña.Text = "" Then
            MsgBox "Ingrese una contraseña"
            txtcontraseña.SetFocus
        Else
            If txtedad.Text = "" Then
                MsgBox "Ingrese una edad"
                txtedad.SetFocus
            Else
                    If Not (IsNumeric(txtedad.Text)) Then
                        MsgBox "Ingrese la edad con números"
                        txtedad.Text = ""
                        txtedad.SetFocus
                    Else
                        If Combo1.Text = "" Then
                            MsgBox "Seleccione un tipo de cuenta"
                            Combo1.SetFocus
                        Else
                            If Combo1.Text = "Administrador" Then
                                'llenar tabla administrador
                                Administrador
                                rsAdministrador.AddNew
                                rsAdministrador("Nombre") = txtnombre.Text
                                rsAdministrador("Contraseña") = txtcontraseña.Text
                                rsAdministrador("Edad") = txtedad.Text
                                rsAdministrador.Update
                                Usuarios
                                rsUsuarios.AddNew
                                rsUsuarios("Nombre") = txtnombre.Text
                                rsUsuarios("Contraseña") = txtcontraseña.Text
                                rsUsuarios("Edad") = txtedad.Text
                                rsUsuarios.Update
                                MsgBox "Administrador registrado"
                            Else
                                'llenar tabla doctores
                                Usuarios
                                rsUsuarios.AddNew
                                rsUsuarios("Nombre") = txtnombre.Text
                                rsUsuarios("Contraseña") = txtcontraseña.Text
                                rsUsuarios("Edad") = txtedad.Text
                                rsUsuarios.Update
                                MsgBox "Doctor registrado"
                                txtnombre.Text = ""
                                txtedad.Text = ""
                                txtcontraseña.Text = ""
                                Combo1.Text = ""
                            End If
                        End If
                    End If
            End If
        End If
    End If
End Sub

Private Sub Command2_Click()
    If MsgBox("¿Desea regresar a la selección de reactivos?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
        frmbuscar.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Label2.ForeColor = RGB(69, 110, 174)
    Label3.ForeColor = RGB(69, 110, 174)
    Label4.ForeColor = RGB(69, 110, 174)
    Label5.ForeColor = RGB(69, 110, 174)
End Sub
