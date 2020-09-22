Imports System.Threading
Public Enum TxEnum
    NewRecord = 1
    CopyRecord = 2
    UpdateRecord = 3
End Enum
Public Class FormTefalShops
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As TefalShopsController

    Public Shared myForm As FormTefalShops
    Dim drv As DataRowView
    Dim ImportFile As String
    Private OpenFileDialog1 As New OpenFileDialog

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormTefalShops
        ElseIf myForm.IsDisposed Then
            myForm = New FormTefalShops
        End If
        Return myForm
    End Function

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub RefreshMYInterface()
        DataGridView1.Invalidate()
    End Sub

    Sub DoWork()
        myController = New TefalShopsController
        Try
            ProgressReport(1, "Loading...Please wait.")
            If myController.loaddata() Then
                ProgressReport(4, "Init Data")
            End If
            ProgressReport(1, String.Format("Loading...Done. Records {0}", myController.BS.Count))
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try

    End Sub
    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myController.BS
            End Select
        End If
    End Sub

    'Public Sub showTx(ByVal tx As TxEnum)

    '    If IsNothing(myController) Then
    '        MessageBox.Show("Refresh the query first.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '        Exit Sub
    '    End If
    '    If Not myThread.IsAlive And Not IsNothing(myController) Then
    '        Select Case tx
    '            Case TxEnum.NewRecord
    '                drv = myController.GetNewRecord
    '                drv.Row.Item("deptid") = 0
    '                drv.Row.Item("manufactureid") = 1
    '                drv.Row.Item("typeid") = 1
    '                drv.Row.Item("statusid") = 1
    '                drv.Row.Item("locationid") = 1


    '                Me.drv.BeginEdit()

    '            Case TxEnum.UpdateRecord
    '                drv = myController.GetCurrentRecord
    '                Me.drv.BeginEdit()
    '            Case TxEnum.CopyRecord
    '                Dim drvori = myController.GetCurrentRecord
    '                drv = myController.GetNewRecord

    '                For i = 1 To drv.Row.ItemArray.Length - 1
    '                    drv.Row.Item(i) = drvori.Row.Item(i)
    '                Next
    '                Me.drv.BeginEdit()
    '        End Select

    '        Dim myform = New DialogITAssets(drv, myController.Model.getManufactureBS,
    '                                          myController.Model.getTypeBS,
    '                                          myController.Model.getDeptBS,
    '                                          myController.Model.getStatusBS,
    '                                          myController.Model.getLocationBS)

    '        RemoveHandler myform.RefreshIntervace, AddressOf RefreshMYInterface
    '        AddHandler myform.RefreshIntervace, AddressOf RefreshMYInterface
    '        myform.ShowDialog()

    '    End If

    'End Sub

    Private Sub AddNewToolStripButton1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles AddNewToolStripButton1.Click
        myController.GetNewRecord()
    End Sub

    Private Sub DeleteToolStripButton2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles DeleteToolStripButton2.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myController.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub FormITAssets_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        LoadData()
    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub RefreshToolStripButton4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles RefreshToolStripButton4.Click
        LoadData()
    End Sub

  

    Private Sub CommitToolStripButton3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CommitToolStripButton3.Click
        myController.save()
    End Sub


    Private Sub ToolStripTextBox2_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ToolStripTextBox1.TextChanged
        Try
            myController.ApplyFilter = ToolStripTextBox1.Text
            ProgressReport(1, String.Format("Filter Done. Records {0} ", myController.BS.Count))
        Catch ex As Exception
            ProgressReport(1, String.Format("{0}", ex.Message))
        End Try
    End Sub



 
    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub
End Class