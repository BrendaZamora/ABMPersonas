Imports System.ComponentModel
Imports System.Text

Public Class PersonasCollection

    Inherits BindingList(Of PersonasClass)

    Public Sub New()
        Me.TraerPersonas()
    End Sub
    Public Function TraerPersonas() As PersonasCollection

        Dim ObjBaseDatos As New BaseDatosClass
        Dim MiDataTable As New DataTable
        Dim MiPersona As PersonasClass

        ObjBaseDatos.objTabla = "personas"
        MiDataTable = ObjBaseDatos.Consultar

        For Each dr As DataRow In MiDataTable.Rows

            MiPersona = New PersonasClass

            MiPersona.id = CInt(dr("id"))

            MiPersona.nombre = dr("nombre")

            MiPersona.direccion = (dr("direccion"))

            MiPersona.codpostal = CInt(dr("codpostal"))

            MiPersona.idprovincia = CInt(dr("idprovincia"))

            MiPersona.tipodocumento = dr("tipodocumento")

            MiPersona.nrodocumento = CInt(dr("nrodocumento"))

            Me.Add(MiPersona)

        Next

        Return Me

    End Function
    Public Sub InsertarPersona(ByVal MiPersona As PersonasClass)

        Dim objBasedeDato As New BaseDatosClass
        objBasedeDato.objTabla = "personas"

        Dim vsql As New StringBuilder
        'vsql.Append("(id")
        vsql.Append("(nombre")
        vsql.Append(",direccion")
        vsql.Append(",nrodocumento")
        vsql.Append(",tipodocumento")
        vsql.Append(",Provincia")
        vsql.Append(",codpostal)")

        vsql.Append("Values")

        'vsql.Append("(" & MiPersona.id & "'")
        vsql.Append("('" & MiPersona.nombre & "'")
        vsql.Append(",'" & MiPersona.direccion & "'")
        vsql.Append(",'" & MiPersona.nrodocumento & "'")
        vsql.Append(",'" & MiPersona.tipodocumento & "'")
        vsql.Append(",'" & MiPersona.idprovincia & "'")
        vsql.Append(",'" & MiPersona.codpostal & "')")

        MiPersona.id = objBasedeDato.Insertar(vsql.ToString)
        Me.Add(MiPersona)


    End Sub
   
    Protected Overrides ReadOnly Property SupportsSearchingCore() As Boolean
        Get
            Return True
        End Get
    End Property

    Protected Overrides Function FindCore(ByVal prop As PropertyDescriptor, ByVal key As Object) As Integer
        For Each carrera In Me
            If prop.GetValue(carrera).Equals(key) Then
                Return Me.IndexOf(carrera)
            End If
        Next

        Return -1
    End Function


    Public Sub EliminarPersona(ByVal MiPersona As PersonasClass)

        'Instancia el Objeto BaseDatosClass para acceder a la base personas.
        Dim objBaseDatos As New BaseDatosClass
        objBaseDatos.objTabla = "personas"

        'Ejecuta el método base eliminar.
        Dim resultado As Boolean
        resultado = objBaseDatos.Eliminar(MiPersona.Id)

        If resultado Then

            'Creates a new collection and assign it the properties for modulo.
            Dim properties As PropertyDescriptorCollection = TypeDescriptor.GetProperties(MiPersona)

            'Sets an PropertyDescriptor to the specific property Id.
            Dim myProperty As PropertyDescriptor = properties.Find("Id", False)

            Me.RemoveAt(Me.FindCore(myProperty, MiPersona.Id))
        Else
            MessageBox.Show("No fue posible eliminar el registro.")
        End If

    End Sub

    Public Sub ActualizarPersona(ByVal MiPersona As PersonasClass)

        'Instancio el el Objeto BaseDatosClass para acceder al la base personas.
        Dim objBaseDatos As New BaseDatosClass
        objBaseDatos.objTabla = "personas"

        Dim vSQL As New StringBuilder
        Dim vResultado As Boolean = False

        vSQL.Append("id='" & MiPersona.id.ToString & "'")
        vSQL.Append(",tipodocumento='" & MiPersona.tipodocumento & "'")
        vSQL.Append(",nrodocumento='" & MiPersona.nrodocumento.ToString & "'")
        vSQL.Append(",nombre='" & MiPersona.nombre & "'")
        vSQL.Append(",direccion='" & MiPersona.direccion & "'")
        vSQL.Append(",idprovincia='" & MiPersona.idprovincia.ToString & "'")
        vSQL.Append(",codpostal='" & MiPersona.codpostal.ToString & "'")

        'Actualizo la tabla personas con el Id.
        Dim resultado As Boolean
        'El método actualizar es una función que devuelve True o False
        'Según como haya resultado la operación sobre la tabla SQL.
        resultado = objBaseDatos.Actualizar(vSQL.ToString, MiPersona.Id)

        If resultado Then
            'El código a continuación sirve para localizar el ítem en la lista
            'en este caso un persona.
            ' Creates a new collection and assign it the properties for modulo.
            Dim properties As PropertyDescriptorCollection = TypeDescriptor.GetProperties(MiPersona)

            'Sets an PropertyDescriptor to the specific property Id.
            Dim myProperty As PropertyDescriptor = properties.Find("Id", False)
            Me.Items.Item(Me.FindCore(myProperty, MiPersona.Id)) = MiPersona
        Else
            MessageBox.Show("No fue posible modificar el registro.")
        End If

    End Sub
End Class
