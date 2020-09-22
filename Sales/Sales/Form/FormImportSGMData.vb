Imports System.Threading
Imports System.Text

Public Class FormImportSGMData
    Dim mythread As New Thread(AddressOf doWork)
    Dim errmsg As New StringBuilder
    Dim selectedfile As String
    Dim Stopwatch As New Stopwatch
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not mythread.IsAlive Then
            'Get file
            errmsg = New StringBuilder
            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""
            ToolStripStatusLabel3.Text = ""

            'OpenFileDialog1.InitialDirectory = "\\172.22.10.44\bonehk\gsmdata\"
            If openfiledialog1.ShowDialog = DialogResult.OK Then
                selectedfile = openfiledialog1.FileName
                mythread = New Thread(AddressOf doWork)
                mythread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Sub doWork()
        Stopwatch.Start()
        Dim ImportSGM = New ImportSGM(Me, OpenFileDialog1.FileName)
        ProgressReport(1, "Processing. Please wait..")
        ProgressReport(2, "Marque")
        If ImportSGM.ValidateFile Then

            If ImportSGM.DoImportFile Then
                Stopwatch.Stop()
                'Thread.Sleep(5000)
                ProgressReport(1, "Done.")
                ProgressReport(5, "Elapsed Time: " & Format(Stopwatch.Elapsed.Minutes, "00") & ":" & Format(Stopwatch.Elapsed.Seconds, "00") & "." & Stopwatch.Elapsed.Milliseconds.ToString)
            Else
                ProgressReport(1, String.Format("Error::{0}", ImportSGM.ErrorMsg))
            End If

            ProgressReport(3, "Continuous")
        Else
            ProgressReport(1, String.Format("Error::{0}", ImportSGM.ErrorMsg))
            ProgressReport(3, "Continuous")
        End If
        Stopwatch.Stop()
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

                Case 3
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous

                Case 5
                    ToolStripStatusLabel3.Text = message
            End Select

        End If

    End Sub
End Class