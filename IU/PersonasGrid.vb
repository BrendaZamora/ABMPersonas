Public Class PersonasGrid

    
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        IngresarDatos.operacion = "Agregar"
        IngresarDatos.Text = "Agregar"
        IngresarDatos.Show()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        IngresarDatos.operacion = "Modificar"
        IngresarDatos.Text = "Modificar"
        IngresarDatos.indice = DataGridView1.CurrentRow.Index
        llenarform()
        IngresarDatos.Show()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        IngresarDatos.operacion = "Eliminar"
        IngresarDatos.Text = "Eliminar"
        IngresarDatos.indice = DataGridView1.CurrentRow.Index
        llenarForm()
        IngresarDatos.Show()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Close()
    End Sub
    Private Sub llenarform()
        IngresarDatos.TextBox1.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString
        IngresarDatos.TextBox2.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString
        IngresarDatos.TextBox3.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString
        IngresarDatos.TextBox4.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString
        IngresarDatos.TextBox5.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString
        IngresarDatos.ComboBox1.SelectedIndex = DataGridView1.CurrentRow.Cells(6).Value.ToString
        IngresarDatos.ComboBox2.SelectedIndex = DataGridView1.CurrentRow.Cells(5).Value.ToString
    End Sub
    Private Sub AlumnoGrid_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = Lst
    End Sub

  
End Class
