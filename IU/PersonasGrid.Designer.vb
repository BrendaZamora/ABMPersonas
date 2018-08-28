<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PersonasGrid
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PersonasGrid))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripButton4 = New System.Windows.Forms.ToolStripButton
        Me.IdDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NombreDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NrodocumentoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TipodocumentoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.IdprovinciaDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DireccionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CodpostalDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PersonasCollectionBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.PersonasCollectionBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(745, 63)
        Me.Panel1.TabIndex = 0
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(3, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(75, 50)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Algerian", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(116, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(268, 54)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Personas"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IdDataGridViewTextBoxColumn, Me.NombreDataGridViewTextBoxColumn, Me.NrodocumentoDataGridViewTextBoxColumn, Me.TipodocumentoDataGridViewTextBoxColumn, Me.IdprovinciaDataGridViewTextBoxColumn, Me.DireccionDataGridViewTextBoxColumn, Me.CodpostalDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.PersonasCollectionBindingSource
        Me.DataGridView1.Location = New System.Drawing.Point(0, 147)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(744, 197)
        Me.DataGridView1.TabIndex = 1
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2, Me.ToolStripButton3, Me.ToolStripButton4})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 63)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(745, 25)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never
        Me.ToolStripButton1.Size = New System.Drawing.Size(69, 22)
        Me.ToolStripButton1.Text = "Agregar"
        Me.ToolStripButton1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolStripButton1.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(78, 22)
        Me.ToolStripButton2.Text = "Modificar"
        Me.ToolStripButton2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ToolStripButton3.Image = CType(resources.GetObject("ToolStripButton3.Image"), System.Drawing.Image)
        Me.ToolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.Size = New System.Drawing.Size(70, 22)
        Me.ToolStripButton3.Text = "Eliminar"
        Me.ToolStripButton3.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ToolStripButton4
        '
        Me.ToolStripButton4.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripButton4.Image = CType(resources.GetObject("ToolStripButton4.Image"), System.Drawing.Image)
        Me.ToolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton4.Name = "ToolStripButton4"
        Me.ToolStripButton4.Size = New System.Drawing.Size(49, 22)
        Me.ToolStripButton4.Text = "Salir"
        Me.ToolStripButton4.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'IdDataGridViewTextBoxColumn
        '
        Me.IdDataGridViewTextBoxColumn.DataPropertyName = "id"
        Me.IdDataGridViewTextBoxColumn.HeaderText = "id"
        Me.IdDataGridViewTextBoxColumn.Name = "IdDataGridViewTextBoxColumn"
        Me.IdDataGridViewTextBoxColumn.ReadOnly = True
        '
        'NombreDataGridViewTextBoxColumn
        '
        Me.NombreDataGridViewTextBoxColumn.DataPropertyName = "nombre"
        Me.NombreDataGridViewTextBoxColumn.HeaderText = "nombre"
        Me.NombreDataGridViewTextBoxColumn.Name = "NombreDataGridViewTextBoxColumn"
        Me.NombreDataGridViewTextBoxColumn.ReadOnly = True
        '
        'NrodocumentoDataGridViewTextBoxColumn
        '
        Me.NrodocumentoDataGridViewTextBoxColumn.DataPropertyName = "nrodocumento"
        Me.NrodocumentoDataGridViewTextBoxColumn.HeaderText = "nrodocumento"
        Me.NrodocumentoDataGridViewTextBoxColumn.Name = "NrodocumentoDataGridViewTextBoxColumn"
        Me.NrodocumentoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'TipodocumentoDataGridViewTextBoxColumn
        '
        Me.TipodocumentoDataGridViewTextBoxColumn.DataPropertyName = "tipodocumento"
        Me.TipodocumentoDataGridViewTextBoxColumn.HeaderText = "tipodocumento"
        Me.TipodocumentoDataGridViewTextBoxColumn.Name = "TipodocumentoDataGridViewTextBoxColumn"
        Me.TipodocumentoDataGridViewTextBoxColumn.ReadOnly = True
        '
        'IdprovinciaDataGridViewTextBoxColumn
        '
        Me.IdprovinciaDataGridViewTextBoxColumn.DataPropertyName = "idprovincia"
        Me.IdprovinciaDataGridViewTextBoxColumn.HeaderText = "idprovincia"
        Me.IdprovinciaDataGridViewTextBoxColumn.Name = "IdprovinciaDataGridViewTextBoxColumn"
        Me.IdprovinciaDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DireccionDataGridViewTextBoxColumn
        '
        Me.DireccionDataGridViewTextBoxColumn.DataPropertyName = "direccion"
        Me.DireccionDataGridViewTextBoxColumn.HeaderText = "direccion"
        Me.DireccionDataGridViewTextBoxColumn.Name = "DireccionDataGridViewTextBoxColumn"
        Me.DireccionDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CodpostalDataGridViewTextBoxColumn
        '
        Me.CodpostalDataGridViewTextBoxColumn.DataPropertyName = "codpostal"
        Me.CodpostalDataGridViewTextBoxColumn.HeaderText = "codpostal"
        Me.CodpostalDataGridViewTextBoxColumn.Name = "CodpostalDataGridViewTextBoxColumn"
        Me.CodpostalDataGridViewTextBoxColumn.ReadOnly = True
        '
        'PersonasCollectionBindingSource
        '
        Me.PersonasCollectionBindingSource.DataSource = GetType(WindowsApplication1.PersonasCollection)
        '
        'PersonasGrid
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(745, 344)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "PersonasGrid"
        Me.Text = "PersonasGrid"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        CType(Me.PersonasCollectionBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton3 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton4 As System.Windows.Forms.ToolStripButton
    Friend WithEvents IdDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NombreDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NrodocumentoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TipodocumentoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdprovinciaDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DireccionDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CodpostalDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PersonasCollectionBindingSource As System.Windows.Forms.BindingSource

End Class
