Public Class IngresarDatos

    Dim operacion_ As String

    Dim pos As Integer


    Dim MiPersona As New PersonasClass


    Public Property operacion() As String
        Get
            Return operacion_
        End Get
        Set(ByVal value As String)
            operacion_ = value
        End Set
    End Property


    Dim indice_ As Byte

    Public Property indice() As Byte
        Get
            Return indice_
        End Get
        Set(ByVal value As Byte)
            indice_ = value
        End Set
    End Property

    Private Sub PersonasForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ComboBox1.DataSource = MiPersona.provincias
        ComboBox2.DataSource = MiPersona.documentos
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        TextBox1.Enabled = False
        TextBox1.ReadOnly = True

    End Sub


    Private Sub Cancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancelar.Click
        Me.Close()
    End Sub


    Private Sub Aceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Aceptar.Click
        'MiPersona.id = Cint(TextBox5.Text)
        MiPersona.nombre = TextBox1.Text
        MiPersona.direccion = TextBox2.Text
        MiPersona.codpostal = CInt(TextBox3.Text)
        MiPersona.nrodocumento = CDec(TextBox4.Text)
        MiPersona.idprovincia = ComboBox1.SelectedIndex
        MiPersona.tipodocumento = ComboBox2.SelectedIndex


        Select Case operacion_
            Case "Agregar"
                lst.InsertarPersona(MiPersona)
            Case "Eliminar"
                lst.RemoveAt(indice_)
            Case "Modificar"
                lst.Item(indice_) = MiPersona

                PersonasGrid.DataGridView1.Refresh()


        End Select
        Me.Close()

    End Sub
End Class