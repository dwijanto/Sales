Imports System.Windows.Forms

Public Class DialogProductFamily
    Private DRV As DataRowView

    Public Shared Event FinishTask()

    Public Sub New(ByVal drv As DataRowView)

        ' This call is required by the designer.
        InitializeComponent()
        Me.DRV = drv
        ' Add any initialization after the InitializeComponent() call.
        InitData()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Validate Then
            DRV.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            RaiseEvent FinishTask()
            Me.Close()
        End If
        
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        DRV.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub InitData()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()

        TextBox1.DataBindings.Add(New Binding("text", DRV, "id", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("text", DRV, "familyname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("text", DRV, "type", True, DataSourceUpdateMode.OnPropertyChanged, ""))
    End Sub

    Private Overloads Function validate() As Boolean
        ErrorProvider1.SetError(TextBox1, "")
        ErrorProvider1.SetError(TextBox2, "")
        ErrorProvider1.SetError(ComboBox1, "")
        Dim myret As Boolean = True
        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Value cannot be null.")
            myret = False
        End If
        If TextBox2.Text = "" Then
            ErrorProvider1.SetError(TextBox2, "Value cannot be null.")
            myret = False
        End If
        If IsNothing(ComboBox1.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from list.")
            myret = False
        End If
        Return myret
    End Function
End Class
