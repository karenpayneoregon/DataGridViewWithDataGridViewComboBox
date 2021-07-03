Imports Demo4.Classes
Imports Demo4.Extensions

Public Class Form1
    WithEvents bsPerson As New BindingSource
    WithEvents bsComboBox As New BindingSource
    Private bsColorInformation As New BindingSource

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim operation As New DataOperations
        operation.GetData()
        If Not operation.Exception.HasError Then

            DataGridView1.AutoGenerateColumns = False
            bsComboBox.DataSource = operation.ColorTable
            bsComboBox.Sort = "ColorText"

            ColorsColumn.DisplayMember = "ColorText"
            ColorsColumn.ValueMember = "ColorId"
            ColorsColumn.DataPropertyName = "ColorId"
            ColorsColumn.DataSource = bsComboBox
            ColorsColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing

            bsPerson.DataSource = operation.PersonsTable
            DataGridView1.DataSource = bsPerson

            bsColorInformation.DataSource = operation.InformationTable

        End If

        For Each col In GetColors()
            cboColors.Items.Add(col.ToString())
        Next

        cboColors.SelectedIndex = 0

        AddHandler DataGridView1.CellValueChanged, AddressOf DataGridView1_CellValueChanged

    End Sub

    ''' <summary>
    ''' Hook into changes while user is traversing items in the
    ''' DataGridViewComboBox column
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing

        If DataGridView1.CurrentCell.IsComboBoxCell Then
            If DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name = "ColorsColumn" Then
                Dim cb = TryCast(e.Control, ComboBox)
                RemoveHandler cb.SelectionChangeCommitted, AddressOf _SelectionChangeCommittedForColorColumn
                AddHandler cb.SelectionChangeCommitted, AddressOf _SelectionChangeCommittedForColorColumn
            End If
        End If

    End Sub
    Private Sub _SelectionChangeCommittedForColorColumn(sender As Object, e As EventArgs)
        Dim colorId As Integer = CType(CType(sender, DataGridViewComboBoxEditingControl).SelectedItem, DataRowView).Row.Field(Of Integer)("ColorId")
        UpdateTable(bsPerson.Identifier, colorId)
    End Sub
    Private Sub UpdateTable(pPersonId As Integer, pColorId As Integer)
        Dim operation As New DataOperations
        operation.UpdateCurrentPerson(pPersonId, pColorId)
        If operation.Exception.HasError Then
            MessageBox.Show(operation.Exception.Message)
        End If
    End Sub
    ''' <summary>
    ''' This event attempts to add a new color to the Colors table
    ''' which is also the data source of the DataGridViewComboBoxColumn.
    ''' 
    ''' If ColorsExists returns true the record is not added as it would
    ''' be a duplicate, if ColorsExists returns false the color is added
    ''' and shows up in the DataGridViewComboBoxColumn immediately.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim operations As New DataOperations
        Dim newColor As Color = Color.FromName(cboColors.Text)
        Dim newIdentifier As Integer = 0

        If Not operations.ColorExists(newColor) Then
            If operations.InsertNewColor(newColor, newIdentifier) Then
                CType(bsComboBox.DataSource, DataTable).Rows.Add(New Object() {newIdentifier, newColor.Name})
                MessageBox.Show($"The color {newColor.Name} has been added and ready now!")
            Else
                If operations.Exception.HasError Then
                    MessageBox.Show(operations.Exception.Message)
                End If
            End If
        Else
            MessageBox.Show($"The color {newColor.Name} not added as it exists already")
        End If
    End Sub
    ''' <summary>
    ''' Stub for providing access to updating, in this case FirstName. In a real application
    ''' there surely would be more properties and this idea here is not the best way too go.
    '''
    ''' A  better idea would be to use a BindingList and monitor ListChanged event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If DataGridView1.CurrentCell.NotComboBox Then
            Console.WriteLine(DataGridView1.Columns(e.ColumnIndex).Name)
        End If
    End Sub

    Private Function GetColors() As IEnumerable(Of KnownColor)
        Dim systemColorsType As Type = GetType(SystemColors)

        Return (From kc In [Enum].GetValues(GetType(KnownColor)).Cast(Of KnownColor)()
                Where kc <> KnownColor.Transparent AndAlso systemColorsType.GetProperty(kc.ToString()) Is Nothing Select kc).ToArray()

    End Function
End Class
