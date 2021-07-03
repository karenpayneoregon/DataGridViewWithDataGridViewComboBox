<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cboColors = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.PersonId = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FirstNameColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColorsColumn = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cboColors)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 178)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(354, 70)
        Me.Panel1.TabIndex = 0
        '
        'cboColors
        '
        Me.cboColors.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboColors.FormattingEnabled = True
        Me.cboColors.Location = New System.Drawing.Point(160, 22)
        Me.cboColors.Name = "cboColors"
        Me.cboColors.Size = New System.Drawing.Size(182, 21)
        Me.cboColors.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Image = Global.Demo4.My.Resources.Resources.Add_8x_16x
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(19, 22)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(135, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Add color"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PersonId, Me.FirstNameColumn, Me.ColorsColumn})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(2)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(354, 178)
        Me.DataGridView1.TabIndex = 1
        '
        'PersonId
        '
        Me.PersonId.DataPropertyName = "ID"
        Me.PersonId.HeaderText = "ID"
        Me.PersonId.Name = "PersonId"
        '
        'FirstNameColumn
        '
        Me.FirstNameColumn.DataPropertyName = "FirstName"
        Me.FirstNameColumn.HeaderText = "First Name"
        Me.FirstNameColumn.Name = "FirstNameColumn"
        '
        'ColorsColumn
        '
        Me.ColorsColumn.HeaderText = "Color"
        Me.ColorsColumn.Name = "ColorsColumn"
        Me.ColorsColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColorsColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(354, 248)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Code sample"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As Button
    Friend WithEvents cboColors As ComboBox
    Friend WithEvents PersonId As DataGridViewTextBoxColumn
    Friend WithEvents FirstNameColumn As DataGridViewTextBoxColumn
    Friend WithEvents ColorsColumn As DataGridViewComboBoxColumn
End Class
