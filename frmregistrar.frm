VERSION 5.00
Begin VB.Form frmregistrar 
   Caption         =   "Registrar reactivos"
   ClientHeight    =   10365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17745
   LinkTopic       =   "Form1"
   Picture         =   "frmregistrar.frx":0000
   ScaleHeight     =   10365
   ScaleWidth      =   17745
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   13800
      TabIndex        =   12
      Top             =   6000
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   13800
      TabIndex        =   11
      Top             =   4680
      Width           =   3015
   End
   Begin VB.TextBox txtcantidadregistrar 
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
      Left            =   9960
      TabIndex        =   10
      Top             =   2520
      Width           =   8655
   End
   Begin VB.TextBox txtmarca 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   660
      Left            =   7320
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   7560
      Width           =   3015
   End
   Begin VB.TextBox txtfecha 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   660
      Left            =   7320
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   6120
      Width           =   3015
   End
   Begin VB.TextBox txtcantidad 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   660
      Left            =   7320
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox txtnombre 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   660
      Left            =   7320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese la cantidad de reactivos a registrar:"
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
      Left            =   11760
      TabIndex        =   9
      Top             =   3120
      Width           =   9735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   7680
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de expiración:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   6240
      Width           =   4575
   End
   Begin VB.Label label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad restante:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   4800
      Width           =   4335
   End
   Begin VB.Label lblnombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de reactivos"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmregistrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'txtcantidad
If Verificar = 0 Then
    
        'If txtcantidadregistrar.Text = "" Then
            'MsgBox "Ingrese una cantidad a registrar", vbInformation, "Laboratorios el Puente"
        'End If
        
    
        If txtnombre.Text = "" Then
                MsgBox "Ingrese un nombre", vbInformation, "Laboratorios el Puente"
        Else
            If txtcantidad.Text = "" Then
                MsgBox "Ingrese una cantidad"
            Else
                If txtfecha.Text = "" Then
                    MsgBox "Ingrese una fecha", vbInformation, "Laboratorios el Puente"
                Else
                    If txtmarca.Text = "" Then
                        MsgBox "Ingrese la marca", vbInformation, "Laboratorios el Puente"
                    Else
                        'If Verificar = 1 Then
                            'guardar nuevo reactivo
                           ' rsReactivos.AddNew
                           ' rsReactivos("NombreReactivos") = txtnombre.Text
                            'rsReactivos("NúmeroReactivos") = txtcantidad.Text
                            'rsReactivos("FechaExpiración") = txtfecha.Text
                            'rsReactivos("Marca") = txtmarca.Text
                            'rsReactivos.Update
                            'rsReactivos.Fields("NombreReactivos") = txtnombre.Text
                            'MsgBox "Guardar datos nuevos"
                            'Unload Me
                            'frmbuscar.Show
                            'Exit Sub
                        'Else
                            If Not (IsNumeric(txtcantidad.Text)) Then
                                MsgBox "Ingrese la cantidad en números", vbInformation, "Laboratorios el Puente"
                                txtcantidad.Text = ""
                                txtcantidad.SetFocus
                            Else
                                
                                    'Resultado = Val(txtcantidad.Text) + Val(txtcantidadregistrar.Text)
                                rsReactivos.Fields("NúmeroReactivos") = txtcantidad.Text
                                rsReactivos.Fields("NombreReactivos") = txtnombre.Text
                                rsReactivos.Fields("FechaExpiración") = txtfecha.Text
                                rsReactivos.Fields("Marca") = txtmarca.Text
                                rsReactivos.Update
                                If rsReactivos.State = 1 Or rsReactivos.State = 0 Then
                                    MsgBox "Reactivo modificado", vbInformation, "Laboratorios el Puente"
                                    Unload Me
                                    frmbuscar.Show
                                Else
                                    MsgBox "Ha ocurrido un error", vbInformation, "Laboratorios el Puente"
                                
                                End If
                            End If
                    End If
                End If
            End If
        End If
        
    
'registar reactivos
Else

        If txtnombre.Text = "" Then
                MsgBox "Ingrese un nombre", vbInformation, "Laboratorios el Puente"
        Else
            If txtcantidad.Text = "" Then
                MsgBox "Ingrese una cantidad"
            Else
                If txtfecha.Text = "" Then
                    MsgBox "Ingrese una fecha", vbInformation, "Laboratorios el Puente"
                Else
                    If txtmarca.Text = "" Then
                        MsgBox "Ingrese la marca", vbInformation, "Laboratorios el Puente"
                    Else
                        If Not (IsNumeric(txtcantidad.Text)) Then
                                MsgBox "Ingrese la cantidad en números", vbInformation, "Laboratorios el Puente"
                                txtcantidad.Text = ""
                                txtcantidad.SetFocus
                        Else
                            'guardar nuevo reactivo
                            rsReactivos.AddNew
                            rsReactivos("NombreReactivos") = txtnombre.Text
                            rsReactivos("NúmeroReactivos") = txtcantidad.Text
                            rsReactivos("FechaExpiración") = txtfecha.Text
                            rsReactivos("Marca") = txtmarca.Text
                            rsReactivos.Update
                            'rsReactivos.Fields("NombreReactivos") = txtnombre.Text
                            MsgBox "Reactivo registrado", vbInformation, "Laboratorios el Puente"
                            Unload Me
                            frmbuscar.Show
                            Exit Sub
                        End If
                    End If
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
Me.Icon = LoadPicture(App.Path & "/logo5.ico")
    If Verificar = 0 Then
        txtcantidad.Enabled = True
        Label5.Visible = False
        txtcantidadregistrar.Visible = False
    End If
    If Verificar = 1 Then
        Label5.Visible = False
        txtcantidadregistrar.Visible = False
        txtcantidad.Enabled = True
    End If
    lblnombre.ForeColor = RGB(69, 110, 174)
    Label2.ForeColor = RGB(69, 110, 174)
    Label3.ForeColor = RGB(69, 110, 174)
    Label4.ForeColor = RGB(69, 110, 174)
    Label5.ForeColor = RGB(69, 110, 174)
    txtnombre.Text = Nombre
    txtcantidad.Text = Cantidad
    txtfecha.Text = Fecha
    txtmarca.Text = Marca
End Sub

