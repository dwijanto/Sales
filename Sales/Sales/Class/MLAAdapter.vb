Imports Sales.PublicClass
Public Class MLAAdapter
    Public Property DS As DataSet
    Public Property BS As New BindingSource

    Public Sub New()

    End Sub

    Public Function LoadData() As Boolean
        Dim myret As Boolean = True
        DS = New DataSet
        BS = New BindingSource
        Dim sqlstr = "select * from sales.mla order by id;"
        If Not DBAdapter1.getDataSet(sqlstr, DS) Then
            Return False
        End If
        'Set Primary Key
        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("id")
        DS.Tables(0).PrimaryKey = pk

        BS.DataSource = DS.Tables(0)
        Return myret
    End Function
    Public Function Save() As Boolean
        Dim myret As Boolean = False
        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If DbAdapter1.MLATX(Me, mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If
        Return myret
    End Function
End Class
