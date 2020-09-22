Imports System.Threading
Imports System.Text
Public Class FormMLA
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Public myAdapter As MLAAdapter

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        AddHandler DialogMLA.FinishTask, AddressOf RefreshDataGrid
    End Sub

    Private Sub FormProductFamily_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim abc = myAdapter.DS.GetChanges()
        If Not IsNothing(abc) Then

            Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                Case Windows.Forms.DialogResult.Yes
                    If Me.Validate Then
                        ToolStripButton4.PerformClick()
                    Else
                        e.Cancel = True
                    End If

                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub
    Private Sub FormUser_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        myAdapter = New MLAAdapter
        Try
            If myAdapter.LoadData() Then
                ProgressReport(4, "Init Data")
            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myAdapter.BS
            End Select
        End If
    End Sub

    'Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
    '    Dim myform = New DialogVendorInput(myAdapter.BS, TxRecord.UpdateRecord)
    '    myform.Show()
    'End Sub

    Private Sub ShowTx(ByVal tx As TxRecord)
        Dim drv As DataRowView = Nothing
        Select Case tx
            Case TxRecord.AddRecord
                drv = myAdapter.BS.AddNew()
            Case TxRecord.UpdateRecord
                drv = myAdapter.BS.Current
        End Select
        Dim myform As New DialogMLA(drv)
        myform.Show()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ToolStripButton1.Click
        ShowTx(TxRecord.AddRecord)
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ToolStripButton4.Click
        myAdapter.Save()
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ToolStripButton2.Click, DataGridView1.DoubleClick
        ShowTx(TxRecord.UpdateRecord)
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.loaddata()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(myAdapter.BS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myAdapter.BS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub RefreshDataGrid()
        DataGridView1.Invalidate()
    End Sub
End Class