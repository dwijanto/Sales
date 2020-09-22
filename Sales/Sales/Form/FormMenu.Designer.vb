<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMenu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMenu))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.DataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportSGMDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OrderStatusReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ByCustomerToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ByCustomerBU = New System.Windows.Forms.ToolStripMenuItem()
        Me.BySalesmanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ByDateRangeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SalesReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SalesReportHKToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SalesReportTWToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SalesGMReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AdminToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConvertEncodingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductFamilyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MLAToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UserGuideToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TefalShopsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DataToolStripMenuItem, Me.ReportsToolStripMenuItem, Me.ExitToolStripMenuItem, Me.AdminToolStripMenuItem, Me.MasterToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(662, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'DataToolStripMenuItem
        '
        Me.DataToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportDataToolStripMenuItem, Me.ImportSGMDataToolStripMenuItem, Me.TefalShopsToolStripMenuItem})
        Me.DataToolStripMenuItem.Name = "DataToolStripMenuItem"
        Me.DataToolStripMenuItem.Size = New System.Drawing.Size(43, 20)
        Me.DataToolStripMenuItem.Text = "Data"
        '
        'ImportDataToolStripMenuItem
        '
        Me.ImportDataToolStripMenuItem.Name = "ImportDataToolStripMenuItem"
        Me.ImportDataToolStripMenuItem.Size = New System.Drawing.Size(165, 22)
        Me.ImportDataToolStripMenuItem.Tag = "ImportSales"
        Me.ImportDataToolStripMenuItem.Text = "Import Data"
        '
        'ImportSGMDataToolStripMenuItem
        '
        Me.ImportSGMDataToolStripMenuItem.Name = "ImportSGMDataToolStripMenuItem"
        Me.ImportSGMDataToolStripMenuItem.Size = New System.Drawing.Size(165, 22)
        Me.ImportSGMDataToolStripMenuItem.Text = "Import SGM Data"
        '
        'ReportsToolStripMenuItem
        '
        Me.ReportsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OrderStatusReportToolStripMenuItem, Me.SalesGMReportToolStripMenuItem})
        Me.ReportsToolStripMenuItem.Name = "ReportsToolStripMenuItem"
        Me.ReportsToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.ReportsToolStripMenuItem.Text = "Reports"
        '
        'OrderStatusReportToolStripMenuItem
        '
        Me.OrderStatusReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ByCustomerToolStripMenuItem, Me.ByCustomerBU, Me.BySalesmanToolStripMenuItem, Me.ByDateRangeToolStripMenuItem, Me.SalesReportToolStripMenuItem, Me.SalesReportHKToolStripMenuItem, Me.SalesReportTWToolStripMenuItem})
        Me.OrderStatusReportToolStripMenuItem.Name = "OrderStatusReportToolStripMenuItem"
        Me.OrderStatusReportToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.OrderStatusReportToolStripMenuItem.Tag = ""
        Me.OrderStatusReportToolStripMenuItem.Text = "Sales Report"
        '
        'ByCustomerToolStripMenuItem
        '
        Me.ByCustomerToolStripMenuItem.Name = "ByCustomerToolStripMenuItem"
        Me.ByCustomerToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.ByCustomerToolStripMenuItem.Tag = "ReportSales"
        Me.ByCustomerToolStripMenuItem.Text = "By Customer"
        '
        'ByCustomerBU
        '
        Me.ByCustomerBU.Name = "ByCustomerBU"
        Me.ByCustomerBU.Size = New System.Drawing.Size(225, 22)
        Me.ByCustomerBU.Tag = "FormCustomerBU"
        Me.ByCustomerBU.Text = "By Customer-BU"
        '
        'BySalesmanToolStripMenuItem
        '
        Me.BySalesmanToolStripMenuItem.Name = "BySalesmanToolStripMenuItem"
        Me.BySalesmanToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.BySalesmanToolStripMenuItem.Tag = "ReportbySalesman"
        Me.BySalesmanToolStripMenuItem.Text = "By Salesman"
        '
        'ByDateRangeToolStripMenuItem
        '
        Me.ByDateRangeToolStripMenuItem.Name = "ByDateRangeToolStripMenuItem"
        Me.ByDateRangeToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.ByDateRangeToolStripMenuItem.Tag = "FormCustomerBUDateRange"
        Me.ByDateRangeToolStripMenuItem.Text = "By Customer-BU Date Range"
        Me.ByDateRangeToolStripMenuItem.Visible = False
        '
        'SalesReportToolStripMenuItem
        '
        Me.SalesReportToolStripMenuItem.Name = "SalesReportToolStripMenuItem"
        Me.SalesReportToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.SalesReportToolStripMenuItem.Tag = "FormTWSalesReport"
        Me.SalesReportToolStripMenuItem.Text = "Sales Report"
        '
        'SalesReportHKToolStripMenuItem
        '
        Me.SalesReportHKToolStripMenuItem.Name = "SalesReportHKToolStripMenuItem"
        Me.SalesReportHKToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.SalesReportHKToolStripMenuItem.Tag = "FormSalesReportHK"
        Me.SalesReportHKToolStripMenuItem.Text = "Sales Report HK"
        '
        'SalesReportTWToolStripMenuItem
        '
        Me.SalesReportTWToolStripMenuItem.Name = "SalesReportTWToolStripMenuItem"
        Me.SalesReportTWToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.SalesReportTWToolStripMenuItem.Tag = "FormSalesReportTW"
        Me.SalesReportTWToolStripMenuItem.Text = "Sales Report TW"
        '
        'SalesGMReportToolStripMenuItem
        '
        Me.SalesGMReportToolStripMenuItem.Name = "SalesGMReportToolStripMenuItem"
        Me.SalesGMReportToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.SalesGMReportToolStripMenuItem.Tag = "FormSGMReport"
        Me.SalesGMReportToolStripMenuItem.Text = "Sales && GM Report"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'AdminToolStripMenuItem
        '
        Me.AdminToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ConvertEncodingToolStripMenuItem})
        Me.AdminToolStripMenuItem.Name = "AdminToolStripMenuItem"
        Me.AdminToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.AdminToolStripMenuItem.Text = "Admin"
        Me.AdminToolStripMenuItem.Visible = False
        '
        'ConvertEncodingToolStripMenuItem
        '
        Me.ConvertEncodingToolStripMenuItem.Name = "ConvertEncodingToolStripMenuItem"
        Me.ConvertEncodingToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.ConvertEncodingToolStripMenuItem.Tag = "FormConvertEncoding"
        Me.ConvertEncodingToolStripMenuItem.Text = "Convert Encoding"
        '
        'MasterToolStripMenuItem
        '
        Me.MasterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProductFamilyToolStripMenuItem, Me.MLAToolStripMenuItem})
        Me.MasterToolStripMenuItem.Name = "MasterToolStripMenuItem"
        Me.MasterToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.MasterToolStripMenuItem.Text = "Master"
        '
        'ProductFamilyToolStripMenuItem
        '
        Me.ProductFamilyToolStripMenuItem.Name = "ProductFamilyToolStripMenuItem"
        Me.ProductFamilyToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.ProductFamilyToolStripMenuItem.Tag = "FormProductFamily"
        Me.ProductFamilyToolStripMenuItem.Text = "Product Family"
        '
        'MLAToolStripMenuItem
        '
        Me.MLAToolStripMenuItem.Name = "MLAToolStripMenuItem"
        Me.MLAToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.MLAToolStripMenuItem.Tag = "FormMLA"
        Me.MLAToolStripMenuItem.Text = "MLA"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UserGuideToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'UserGuideToolStripMenuItem
        '
        Me.UserGuideToolStripMenuItem.Name = "UserGuideToolStripMenuItem"
        Me.UserGuideToolStripMenuItem.Size = New System.Drawing.Size(131, 22)
        Me.UserGuideToolStripMenuItem.Text = "User Guide"
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Text = "NotifyIcon1"
        Me.NotifyIcon1.Visible = True
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label1.Size = New System.Drawing.Size(638, 23)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Label1"
        '
        'TefalShopsToolStripMenuItem
        '
        Me.TefalShopsToolStripMenuItem.Name = "TefalShopsToolStripMenuItem"
        Me.TefalShopsToolStripMenuItem.Size = New System.Drawing.Size(165, 22)
        Me.TefalShopsToolStripMenuItem.Tag = "FormTefalShops"
        Me.TefalShopsToolStripMenuItem.Text = "Tefal Shops"
        '
        'FormMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(662, 107)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FormMenu"
        Me.Text = "Sales"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents DataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents OrderStatusReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AdminToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ConvertEncodingToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByCustomerToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BySalesmanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByCustomerBU As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ByDateRangeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalesReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProductFamilyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MLAToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UserGuideToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalesReportHKToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportSGMDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalesGMReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalesReportTWToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TefalShopsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
