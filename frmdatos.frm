VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
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
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14040
      TabIndex        =   16
      Top             =   7080
      Width           =   4095
   End
   Begin VB.CommandButton cmdguardar 
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
      Height          =   660
      Left            =   10200
      TabIndex        =   15
      Top             =   7080
      Width           =   3495
   End
   Begin VB.TextBox txtbuscartexto 
      Height          =   975
      Left            =   1080
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   16680
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
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
      Connect         =   $"frmdatos.frx":B2C0
      OLEDBString     =   $"frmdatos.frx":B348
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Administrador"
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton cmdmodificar 
      Caption         =   "Modificar usuario"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14040
      TabIndex        =   12
      Top             =   6240
      Width           =   4095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   18480
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
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
      Connect         =   $"frmdatos.frx":B3D0
      OLEDBString     =   $"frmdatos.frx":B458
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmdatos.frx":B4E0
      Height          =   3015
      Left            =   9000
      TabIndex        =   11
      Top             =   3000
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   35
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frmdatos.frx":B4F5
      Left            =   4800
      List            =   "frmdatos.frx":B4FF
      TabIndex        =   10
      Top             =   6720
      Width           =   3495
   End
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
      Height          =   690
      Left            =   12000
      TabIndex        =   9
      Top             =   8280
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Añadir usuario"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   8
      Top             =   6240
      Width           =   3495
   End
   Begin VB.TextBox txtedad 
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
      Height          =   690
      Left            =   4800
      TabIndex        =   7
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox txtcontraseña 
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
      Height          =   690
      Left            =   4800
      TabIndex        =   6
      Top             =   4200
      Width           =   3495
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
      Height          =   690
      Left            =   4800
      TabIndex        =   5
      Top             =   3000
      Width           =   3495
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmdatos.frx":B51A
      Height          =   3015
      Left            =   9000
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   35
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   540
      Left            =   960
      TabIndex        =   1
      Top             =   3120
      Width           =   1845
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   960
      TabIndex        =   2
      Top             =   4320
      Width           =   2625
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de cuenta:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   960
      TabIndex        =   3
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edad:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   960
      TabIndex        =   4
      Top             =   5520
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administrar usuarios"
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
      Height          =   1170
      Left            =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmdatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdañadir_Click()

End Sub

Private Sub cmdeliminar_Click()
If Len(Trim(txtbuscartexto.Text)) = 0 Then
    MsgBox "Seleccione el usuario a eliminar", vbInformation, "Laboratorios el Puente "
Else
    Usuarios
    Administrador
    rsUsuarios.Find "Nombre = '" & txtbuscartexto.Text & "'", , , 1
    If rsUsuarios!Tipo = "Administrador" Then
        
        rsAdministrador.Find "Nombre = '" & txtbuscartexto.Text & "'", , , 1
    End If
    If rsUsuarios.BOF = False And rsUsuarios.EOF = False Then
        If MsgBox("¿Desea eliminar el usuario seleccionado?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
                If rsUsuarios!Tipo = "Administrador" Then
                    rsAdministrador.Delete
                End If
                rsUsuarios.Delete
                'rsAdministrador.Delete
                MsgBox "Usuario eliminado", vbInformation, "Laboratorios el Puente "
                Unload Me
                frmdatos.Show
                DataGrid1.Refresh
           
        End If
    End If
End If
End Sub

Private Sub cmdguardar_Click()
If Combo1.Text = "Doctor" Then
    rsUsuarios.Fields("Nombre") = txtnombre.Text
    rsUsuarios.Fields("Edad") = txtedad.Text
    rsUsuarios.Fields("Contraseña") = txtcontraseña.Text
    rsUsuarios.Fields("Tipo") = Combo1.Text
    rsUsuarios.Update
    MsgBox "Usuario modificado"
    Unload Me
    frmdatos.Show
    Administrador
    rsAdministrador.Find "Nombre = '" & txtbuscartexto.Text & "'", , , 1
    If rsAdministrador.BOF = False And rsAdministrador.EOF = False Then
        rsAdministrador.Delete
        rsAdministrador.Update
        If rsUsuarios.State = 1 Or rsUsuarios.State = 0 Then
            MsgBox "Usuario modificado", vbInformation, "Laboratorios el Puente"
            'cmdmodificar.Caption = "Modificar usuario"
            Unload Me
            frmdatos.Show
        Else
            MsgBox "Ha ocurrido un error", vbInformation, "Laboratorios el Puente"
        
        End If
    End If
Else
    rsUsuarios.Fields("Nombre") = txtnombre.Text
    rsUsuarios.Fields("Edad") = txtedad.Text
    rsUsuarios.Fields("Contraseña") = txtcontraseña.Text
    rsUsuarios.Fields("Tipo") = Combo1.Text
    rsUsuarios.Update
    
    Administrador
    
   rsAdministrador.AddNew
        rsAdministrador("Nombre") = txtnombre.Text
        rsAdministrador("Contraseña") = txtcontraseña.Text
        rsAdministrador("Edad") = txtedad.Text
        rsAdministrador("Tipo") = Combo1.Text
        rsAdministrador.Update
        If rsUsuarios.State = 1 Or rsUsuarios.State = 0 Then
            MsgBox "Usuario modificado", vbInformation, "Laboratorios el Puente"
            'cmdmodificar.Caption = "Modificar usuario"
            Unload Me
            frmdatos.Show
        Else
            MsgBox "Ha ocurrido un error", vbInformation, "Laboratorios el Puente"
        
        End If
End If
End Sub

Private Sub cmdmodificar_Click()

If Len(Trim(txtbuscartexto.Text)) = 0 Then
    MsgBox "Seleccione el usuario a modificar", vbInformation, "Laboratorios el Puente "
Else
    Usuarios
    rsUsuarios.Find "Nombre = '" & txtbuscartexto.Text & "'", , , 1
    If rsUsuarios.BOF = False And rsUsuarios.EOF = False Then
        'cargar datos a las cajas de texto
        txtnombre.Text = rsUsuarios.Fields("Nombre")
        txtedad.Text = rsUsuarios.Fields("Edad")
        txtcontraseña.Text = rsUsuarios.Fields("Contraseña")
        Combo1.Text = rsUsuarios.Fields("Tipo")
        If MsgBox("¿Desea modificar el usuario seleccionado?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
            'cmdmodificar.Caption = "Guardar"
            cmdguardar.Enabled = True
            cmdmodificar.Enabled = False
        Else
            txtnombre.Text = ""
            txtedad.Text = ""
            txtcontraseña.Text = ""
            Combo1.Text = ""
            txtbuscartexto.Text = ""
        End If
        'Nombre = txtnombre.Text
        'Cantidad = txtcantidad.Text
        'Fecha = txtfecha.Text
        'Marca = txtmarca.Text
        'If MsgBox("¿Desea modificar el reactivo seleccionado?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
           ' Adodc2.RecordSource = "select * from Administrador where Nombre =  '" + txtusuario.Text + "'"
           ' Adodc2.Refresh
            'If Adodc2.Recordset.EOF Then
               ' MsgBox "Solo administradores pueden modificar reactivos", vbInformation, "Laboratorios el Puente "
           'Else
               ' Verificar = 0
                'frmregistrar.Show
                'Unload Me
            'End If
        'End If
    End If
End If












'If txtnombre.Text = "" Then
       ' MsgBox "Ingrese un nombre"
       ' txtnombre.SetFocus
 '  Else
     '   If txtcontraseña.Text = "" Then
       '     MsgBox "Ingrese una contraseña"
         '   txtcontraseña.SetFocus
      '  Else
           ' If txtedad.Text = "" Then
              '  MsgBox "Ingrese una edad"
               ' txtedad.SetFocus
          '  Else
                   '' If Not (IsNumeric(txtedad.Text)) Then
                       ' MsgBox "Ingrese la edad con números"
                      '  txtedad.Text = ""
                      '  txtedad.SetFocus
                  '  Else
                       ' If Combo1.Text = "" Then
                         '   MsgBox "Seleccione un tipo de cuenta"
                          '  Combo1.SetFocus
                       ' Else
                                'Usuarios
                               ' rsUsuarios.Find "Nombre = '" & txtbuscartexto.Text & "'", , , 1
                               ' If rsReactivos.BOF = False And rsReactivos.EOF = False Then
                                
                               ' End If
                            'El'se
                                'llenar tabla doctores
                               ' Usuarios
                               ' rsUsuarios.AddNew
                               ' rsUsuarios("Nombre") = txtnombre.Text
                               ' rsUsuarios("Contraseña") = txtcontraseña.Text
                               ' rsUsuarios("Edad") = txtedad.Text
                               ' rsUsuarios.Update
                               ' MsgBox "Doctor registrado"
                               ' txtnombre.Text = ""
                               ' txtedad.Text = ""
                               ' txtcontraseña.Text = ""
                               ' Combo1.Text = ""
                               ' Unload Me
                               ' frmdatos.Show
                            
                     '  End If
                   ' End If
           ' End If
     '  End If
  '  End If
 '   If Not (IsNumeric(txtedad.Text)) Then
      '  MsgBox "Ingrese la edad en números", vbInformation, "Laboratorios el Puente"
      '  txtedad.Text = ""
      '  txtcantidad.SetFocus
 '   Else
        
            'Resultado = Val(txtcantidad.Text) + Val(txtcantidadregistrar.Text)
        
            
      '  If Combo1.Text = "Doctor" Then
       '     Usuarios
         '   rsUsuarios.Fields("Nombre") = txtnombre.Text
        '    rsUsuarios.Fields("Edad") = txtedad.Text
          '  rsUsuarios.Fields("Contraseña") = txtcontraseña.Text
         '   rsUsuarios.Fields("Tipo") = Combo1.Text
            'rsUsuarios.Delete
           ' rsUsuarios.Update
           ' Administrador
           ' rsAdministrador.Delete
           ' rsAdministrador.Update
           ' If rsUsuarios.State = 1 Or rsUsuarios.State = 0 Then
           '     MsgBox "Usuario modificado", vbInformation, "Laboratorios el Puente"
           '     Unload Me
           '     frmdatos.Show
           ' Else
           '     MsgBox "Ha ocurrido un error", vbInformation, "Laboratorios el Puente"
           '
           ' End If
       ' Else
        '    Usuarios
         '   rsUsuarios.Fields("Nombre") = txtnombre.Text
          '  rsUsuarios.Fields("Edad") = txtedad.Text
           ' rsUsuarios.Fields("Contraseña") = txtcontraseña.Text
        '    rsUsuarios.Fields("Tipo") = Combo1.Text
         '   rsUsuarios.Update
          '  Administrador
         '   rsAdministrador.AddNew
          '  rsAdministrador("Nombre") = txtnombre.Text
           ' rsAdministrador("Contraseña") = txtcontraseña.Text
           ' rsAdministrador("Edad") = txtedad.Text
           ' rsAdministrador("Tipo") = Combo1.Text
          '  rsAdministrador.Update
         '   If rsUsuarios.State = 1 Or rsUsuarios.State = 0 Then
        '        MsgBox "Usuario modificado", vbInformation, "Laboratorios el Puente"
       '         Unload Me
      '          frmdatos.Show
     '       Else
    '            MsgBox "Ha ocurrido un error", vbInformation, "Laboratorios el Puente"
   '
  '          End If
 '   End If
'End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Cancelar" Then
    'cmdmodificar.Enabled = False
    txtnombre.Text = ""
    txtedad.Text = ""
    txtcontraseña.Text = ""
    Combo1.Text = ""
    Command1.Caption = "Añadir usuario"
    txtbuscartexto.Text = ""
    cmdguardar.Enabled = False
    cmdmodificar.Enabled = True
Else
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
                                rsAdministrador("Tipo") = Combo1.Text
                                rsAdministrador.Update
                                Usuarios
                                rsUsuarios.AddNew
                                rsUsuarios("Nombre") = txtnombre.Text
                                rsUsuarios("Contraseña") = txtcontraseña.Text
                                rsUsuarios("Edad") = txtedad.Text
                                rsUsuarios("Tipo") = Combo1.Text
                                rsUsuarios.Update
                                MsgBox "Administrador registrado"
                                Unload Me
                                frmdatos.Show
                            Else
                                'llenar tabla doctores
                                Usuarios
                                rsUsuarios.AddNew
                                rsUsuarios("Nombre") = txtnombre.Text
                                rsUsuarios("Contraseña") = txtcontraseña.Text
                                rsUsuarios("Edad") = txtedad.Text
                                rsUsuarios("Tipo") = Combo1.Text
                                rsUsuarios.Update
                                MsgBox "Doctor registrado"
                                txtnombre.Text = ""
                                txtedad.Text = ""
                                txtcontraseña.Text = ""
                                Combo1.Text = ""
                                Unload Me
                                frmdatos.Show
                            End If
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

Private Sub DataGrid1_Click()
    'If MsgBox("¿Desea modificar el usuario seleccionado?", vbInformation + vbYesNo, "Laboratorios el Puente") = vbYes Then
       txtbuscartexto.Text = DataGrid1.Columns(1).Text
        'cmdmodificar.Enabled = True
       ' cmdmodificar.Enabled = True
        Command1.Caption = "Cancelar"
        'cmdeliminar.Enabled = True
       ' txtnombre.Text = DataGrid1.Columns(1).Text
       ' txtedad.Text = DataGrid1.Columns(2).Text
       ' txtcontraseña.Text = DataGrid1.Columns(3).Text
       ' Combo1.Text = DataGrid1.Columns(4).Text
    'End If
    
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & "/logo5.ico")
    Label2.ForeColor = RGB(69, 110, 174)
    Label3.ForeColor = RGB(69, 110, 174)
    Label4.ForeColor = RGB(69, 110, 174)
    Label5.ForeColor = RGB(69, 110, 174)
    formato
    formato2
    'cmdeliminar.Enabled = False
    'cmdmodificar.Enabled = False
    cmdguardar.Enabled = False
End Sub
Sub formato()
    DataGrid1.Columns(0).Width = 0
    DataGrid1.Columns(1).Width = 2300
    DataGrid1.Columns(2).Width = 2300
    DataGrid1.Columns(3).Width = 2500
    DataGrid1.Columns(4).Width = 2500
    DataGrid1.ForeColor = RGB(69, 110, 174)
End Sub
Sub formato2()
    DataGrid2.Columns(0).Width = 0
    DataGrid2.Columns(1).Width = 2300
    DataGrid2.Columns(2).Width = 2300
    DataGrid2.Columns(3).Width = 2500
    DataGrid2.Columns(4).Width = 0
    DataGrid2.ForeColor = RGB(69, 110, 174)
End Sub
