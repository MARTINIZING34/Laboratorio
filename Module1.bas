Attribute VB_Name = "Module1"
Global Nombre, Fecha, Marca, Cantidad, Resultado, Verificar As Integer
Global Usuario As String
Global cn As New ADODB.Connection
Global rsReactivos As New ADODB.Recordset
Global rsRegistro As New ADODB.Recordset
Global rsUsuarios As New ADODB.Recordset

Sub main()
    With cn
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Bravo\Desktop\Git\G1\Laboratorio.mdb;Persist Security Info=False"
        frminicio.Show
    End With
End Sub
Sub TablaReactivos()
    With rsReactivos
        If .State = 1 Then .Close
        .Source = "Reactivos"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open "select * from Reactivos", cn
    End With
End Sub
Sub TablaRegistro_Uso()
    With rsRegistro
        If .State = 1 Then .Close
        .Source = "Reactivos"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open "select * from Registro_Uso", cn
    End With
End Sub
Sub Usuarios()
    With rsUsuarios
        If .State = 1 Then .Close
        .Source = "Reactivos"
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open "select * from Doctores", cn
    End With
End Sub
