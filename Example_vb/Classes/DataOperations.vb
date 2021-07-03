Imports System.Data.SqlClient

Namespace Classes
    Public Class DataOperations
        Private ConnectionString As String = "Data Source=.\SQLEXPRESS;Initial Catalog=ExampleDataGridViewComboBox_1;Integrated Security=True"
        Public Property PersonsTable As New DataTable
        Public Property ColorTable As New DataTable
        Public Property InformationTable As New DataTable
        Public Property Exception As New ErrorInformation
        ''' <summary>
        ''' Load data, called in form load event
        ''' </summary>
        Public Sub GetData()

            Using cn As New SqlConnection With {.ConnectionString = ConnectionString}

                Using cmd As New SqlCommand With {.Connection = cn}

                    Try
                        cn.Open()
                    Catch ex As Exception
                        Exception.HasError = True
                        Exception.Message = ex.Message
                    End Try

                    cmd.CommandText = "SELECT ColorId,ColorText FROM Colors ORDER BY ColorText"
                    ColorTable.Load(cmd.ExecuteReader)

                    cmd.CommandText = "SELECT Id,[Text] FROM Information"
                    InformationTable.Load(cmd.ExecuteReader)
                    cmd.CommandText = "SELECT Id,ColorId ,FirstName FROM Person"
                    PersonsTable.Load(cmd.ExecuteReader)

                End Using

            End Using

        End Sub
        ''' <summary>
        ''' Add a new color to the Colors Table if it does not
        ''' exists in the database table.
        ''' </summary>
        ''' <param name="pColor">Color to add</param>
        ''' <param name="pNewIdentifier">Contains the new primary key on success</param>
        ''' <returns></returns>
        Public Function InsertNewColor(pColor As Color, ByRef pNewIdentifier As Integer) As Boolean

            Using cn As New SqlConnection With {.ConnectionString = ConnectionString}

                Using cmd As New SqlCommand With {.Connection = cn}

                    cmd.CommandText = "INSERT INTO dbo.Colors (ColorText) VALUES (@ColorText); " &
                                      "SELECT CAST(scope_identity() AS int);"

                    cmd.Parameters.AddWithValue("@ColorText", pColor.Name)
                    cn.Open()

                    Try
                        ' insert and get new id
                        pNewIdentifier = CInt(cmd.ExecuteScalar)
                        Return True

                    Catch ex As Exception
                        Exception.HasError = True
                        Exception.Message = ex.Message
                        Return False
                    End Try

                End Using

            End Using

        End Function
        ''' <summary>
        ''' Called before above method to see if the color currently
        ''' exists in the database table.
        ''' </summary>
        ''' <param name="pColor"></param>
        ''' <returns></returns>
        Public Function ColorExists(pColor As Color) As Boolean

            Using cn As New SqlConnection With {.ConnectionString = ConnectionString}

                Using cmd As New SqlCommand With {.Connection = cn}

                    cmd.CommandText = "SELECT [ColorId] FROM [dbo].[Colors] WHERE ColorText = @ColorText"
                    cmd.Parameters.AddWithValue("@ColorText", pColor.Name)

                    cn.Open()

                    Try
                        Dim reader As SqlDataReader = cmd.ExecuteReader
                        If reader.HasRows Then
                            Return True
                        Else
                            Return False
                        End If
                    Catch ex As Exception
                        Exception.HasError = True
                        Exception.Message = ex.Message
                        Return False
                    End Try

                End Using

            End Using

        End Function
        ''' <summary>
        ''' Invoked from SelectionChangeCommitted in the form to update
        ''' any changes to the color in the DataGridViewComboBoxColumn of
        ''' the current row in the DataGridView.
        ''' </summary>
        ''' <param name="pPersonId"></param>
        ''' <param name="pColorId"></param>
        ''' <returns></returns>
        Public Function UpdateCurrentPerson(pPersonId As Integer, pColorId As Integer) As Boolean

            Dim result As Integer = 0

            Using cn As New SqlConnection With {.ConnectionString = ConnectionString}

                Using cmd As New SqlCommand With {.Connection = cn}
                    cmd.CommandText = "UPDATE Person SET ColorId = @ColorId WHERE Id = @PersonId"
                    cmd.Parameters.AddWithValue("@ColorId", pColorId)
                    cmd.Parameters.AddWithValue("@PersonId", pPersonId)

                    Try
                        cn.Open()
                    Catch ex As Exception
                        Exception.HasError = True
                        Exception.Message = ex.Message
                    End Try

                    Try
                        result = cmd.ExecuteNonQuery
                    Catch ex As Exception
                        '
                        ' Decide how to deal with this
                        '
                        Console.WriteLine(ex.Message)
                    End Try

                End Using
            End Using

            Return result = 1

        End Function

    End Class
End Namespace