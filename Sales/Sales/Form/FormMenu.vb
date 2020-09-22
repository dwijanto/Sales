Imports System.Reflection
Imports Sales.PublicClass
Public Enum MyLocation
    HongKong = 1
    Taiwan = 2
End Enum

Public Class FormMenu
    Public Function GetMenuDesc() As String
        Label1.Text = "Welcome, " & HelperClass1.UserInfo.DisplayName
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & DbAdapter1.ConnectionStringDict.Item("HOST") & ", Database: " & DbAdapter1.ConnectionStringDict.Item("DATABASE") & ", Userid: " & HelperClass1.UserId & "-" & PublicClass.myLocation.ToString

    End Function

    Public Sub LoadMe()
        Me.FormMenu_Load(Me, New EventArgs)
    End Sub

    Private Sub FormMenu_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            HelperClass1 = New HelperClass
            DbAdapter1 = New DbAdapter
            Try
                PublicClass.myLocation = MyLocation.Taiwan
                'Don't Forget to Enable Import Sales Function For Taiwan
                'Import Sales Function Disabled for Hong Kong on purpose
                'PublicClass.myLocation = MyLocation.HongKong


                loglogin(DbAdapter1.userid)
                HelperClass1.UserInfo.IsSalesTeamHK = DbAdapter1.isSalesTeam(String.Format("{0}\{1}", HelperClass1.UserInfo.Domain, HelperClass1.UserInfo.userid))
            Catch ex As Exception
            End Try
            Me.Text = GetMenuDesc()
            Me.Location = New Point(300, 10)
            'dbtools1.Userid = DbAdapter1.userid
            'dbtools1.Password = DbAdapter1.password

            SalesReportToolStripMenuItem.Visible = IIf(PublicClass.myLocation = MyLocation.Taiwan, True, False)
            MasterToolStripMenuItem.Visible = IIf(PublicClass.myLocation = MyLocation.Taiwan, True, False)
            SalesReportHKToolStripMenuItem.Visible = IIf(PublicClass.myLocation = MyLocation.HongKong, True, False)
            ImportSGMDataToolStripMenuItem.Visible = IIf(PublicClass.myLocation = MyLocation.HongKong, True, False)
            SalesGMReportToolStripMenuItem.Visible = IIf(PublicClass.myLocation = MyLocation.HongKong, True, False)
            SalesReportTWToolStripMenuItem.Visible = IIf(PublicClass.myLocation = MyLocation.Taiwan, True, False)
            ImportDataToolStripMenuItem.Visible = IIf(PublicClass.myLocation = MyLocation.Taiwan, True, False)
            TefalShopsToolStripMenuItem.Visible = IIf(PublicClass.myLocation = MyLocation.HongKong, True, False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        End Try

    End Sub
    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "Sales"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
        'DbAdapter1.isSalesTeam(username)
    End Sub

    Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportDataToolStripMenuItem.Click, ConvertEncodingToolStripMenuItem.Click, ByCustomerToolStripMenuItem.Click, BySalesmanToolStripMenuItem.Click, ByCustomerBU.Click, ByDateRangeToolStripMenuItem.Click, SalesReportToolStripMenuItem.Click, ProductFamilyToolStripMenuItem.Click, MLAToolStripMenuItem.Click, SalesReportHKToolStripMenuItem.Click, SalesGMReportToolStripMenuItem.Click, TefalShopsToolStripMenuItem.Click
        Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim assembly1 As Assembly = Assembly.GetAssembly(GetType(FormMenu))
        Dim frm As Form = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Form)
        Dim inMemory As Boolean = False
        For i = 0 To My.Application.OpenForms.Count - 1
            If My.Application.OpenForms.Item(i).Name = frm.Name Then
                ExecuteForm(My.Application.OpenForms.Item(i))
                inMemory = True
            End If
        Next
        If Not inMemory Then
            ExecuteForm(frm)
        End If
    End Sub

    Private Sub ExecuteForm(ByVal obj As Windows.Forms.Form)
        With obj
            .WindowState = FormWindowState.Normal
            .StartPosition = FormStartPosition.CenterScreen
            .Show()
            .Focus()
        End With
    End Sub

    Private Sub FormMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not e.CloseReason = CloseReason.ApplicationExitCall Then
            If MessageBox.Show("Are you sure?", "Exit", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Me.CloseOpenForm()
                HelperClass1.fadeout(Me)
                DbAdapter1.Dispose()
                HelperClass1.Dispose()
                System.Diagnostics.Process.Start("ClearContentMSO.bat")
            Else
                e.Cancel = True
            End If
        End If
    End Sub
    Private Sub CloseOpenForm()
        For i = 1 To (My.Application.OpenForms.Count - 1)
            My.Application.OpenForms.Item(1).Close()
        Next
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Protected Friend Sub setBubbleMessage(ByVal title As String, ByVal message As String)
        NotifyIcon1.BalloonTipText = message
        NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
        NotifyIcon1.BalloonTipTitle = title
        NotifyIcon1.Visible = True
        NotifyIcon1.ShowBalloonTip(200)
        'ShowballonWindow(200)
    End Sub


    Private Sub ConvertEncodingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConvertEncodingToolStripMenuItem.Click

    End Sub

    Private Sub ByCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByCustomerToolStripMenuItem.Click

    End Sub


    Private Sub ByDateRangeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ByDateRangeToolStripMenuItem.Click

    End Sub


    Private Sub SalesReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesReportToolStripMenuItem.Click

    End Sub

    Private Sub ProductFamilyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductFamilyToolStripMenuItem.Click

    End Sub

    Private Sub MLAToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MLAToolStripMenuItem.Click

    End Sub

    Private Sub UserGuideToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserGuideToolStripMenuItem.Click
        Dim p As New System.Diagnostics.Process
        'p.StartInfo.FileName = "\\172.22.10.44\SharedFolder\PriceCMMF\New\template\Supplier Management Task User Guide-Admin.pdf"
        p.StartInfo.FileName = Application.StartupPath + "\help\Sales Report.docx"
        p.Start()
    End Sub



    Private Sub SalesReportHKToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesReportHKToolStripMenuItem.Click

    End Sub

    Private Sub ImportSGMDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportSGMDataToolStripMenuItem.Click
        Dim myform As New FormImportSGMData
        myform.ShowDialog()
    End Sub


    Private Sub SalesGMReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesGMReportToolStripMenuItem.Click

    End Sub

    Private Sub SalesReportTWToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesReportTWToolStripMenuItem.Click
        Dim myform As New FormSalesReportTW
        myform.ShowDialog()
    End Sub

    Private Sub TefalShopsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TefalShopsToolStripMenuItem.Click

    End Sub
End Class