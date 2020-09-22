Imports System.Threading
Imports Microsoft.Office.Interop
Imports Sales.SharedClass
Imports System.Text
Imports Sales.PublicClass
Imports System.ComponentModel
Public Class FormTWSalesReport
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)
    Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)

    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim myQueryThread As New System.Threading.Thread(QueryDelegate)
    Dim startdate As Date
    Dim enddate As Date
    Dim saodate As String

    Dim Dataset1 As DataSet
    Dim Filename As String = String.Empty
    Dim exclude As Boolean = True
    Dim myYearWeek As String = String.Empty
    Dim salesmanbs As New BindingSource

    Dim mycriteria As String = String.Empty
    Dim SelectedPath As String = String.Empty
    Dim askmetoopen As Boolean = False

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not myThread.IsAlive Then
            'get Criteria
            'myYearWeek = String.Format("{0} {1:00}", TextBox1.Text, CInt(TextBox2.Text))
            startdate = DateTimePicker1.Value.Date
            enddate = DateTimePicker2.Value.Date
            askmetoopen = CheckBox1.Checked
            If Year(startdate) <> Year(enddate) Then
                MessageBox.Show("Year of startdate and endddate should be the same.")
                Exit Sub
            End If
            ProgressReport(5, "")
            Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
            DirectoryBrowser.Description = "Which directory do you want to use?"
            If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                SelectedPath = DirectoryBrowser.SelectedPath

                Try
                    myThread = New System.Threading.Thread(myThreadDelegate)
                    myThread.SetApartmentState(ApartmentState.MTA)
                    myThread.Start()
                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try
            End If

        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Sub DoWork()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        Dim status As Boolean = False
        Dim message As String = String.Empty
        sw.Start()

        'Dim chkstate As CheckState
        'chkstate = CheckedListBox1.GetItemChecked(0)

        'For Each item As Object In CheckedListBox1.CheckedItems

        '    ProgressReport(2, "Export To Excel..")
        '    ProgressReport(5, "")
        '    Dim dr As DataRowView = DirectCast(item, DataRowView)
        '    Dim myvalue = dr.Item(0)
        '    Filename = SelectedPath & ""
        '    If IsDBNull(myvalue) Then
        '        myvalue = ""
        '    End If
        '    If myvalue <> "Select All" Then
        '        mycriteria = ""
        '        'Filename = SelectedPath & "\" & String.Format("Sales {0} {1}-{2}.xlsx", myvalue, Format(startdate, "ddMMMyyyy"), Format(enddate, "ddMMMyyyy"))
        '        If myvalue <> "All Salesman" Then
        '            mycriteria = " and salesman is null"
        '            If myvalue <> "" Then
        '                mycriteria = " and salesman = " & escapestr(myvalue)
        '            End If

        '        End If

        '    End If

        'Next
        Filename = SelectedPath & "\" & String.Format("TW Sales {0} {1}-{2}.xlsx", "Report", Format(startdate, "ddMMMyyyy"), Format(enddate, "ddMMMyyyy"))
        'Select Case myvalue
        '    Case "Select All"
        '    Case "All Customer"
        '    Case Else

        '        mycriteria = " and customername=" & escapestr(myvalue)
        'End Select

        status = GenerateReport(message)
        If status Then
            sw.Stop()
            ProgressReport(5, String.Format("Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            'ProgressReport(2, TextBox2.Text & "Done.")
            'ProgressReport(5, "")
            If askmetoopen Then
                If MsgBox("File name: " & Filename & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                    Process.Start(Filename)
                End If
            End If

            'ProgressReport(5, "")
        Else
            ProgressReport(5, message)
        End If



        sw.Stop()
    End Sub

    'Private Function GenerateReport(ByRef FileName As String, ByRef errorMsg As String, ByVal dataset1 As DataSet) As Boolean
    Private Function GenerateReport(ByRef errmsg As String) As Boolean
        'Dim myCriteria As String = String.Empty
        Dim result As Boolean = False
        Dim hwnd As System.IntPtr
        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()
        'Cursor.Current = Cursors.WaitCursor


        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty

        Try
            'Create Object Excel 
            ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd

            oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(2, "Opening Template...")
            ProgressReport(2, "Generating records..")
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ExcelTemplate.xltx")

            Dim counter As Integer = 0
            ProgressReport(2, "Creating Worksheet...")
            'backOrder
            For i = 0 To 2
                oWb.Worksheets.Add(After:=(oWb.Worksheets(3 + i)))
            Next i

            Dim sqlstr As String = String.Empty
            Dim obj As New ThreadPoolObj

            'Get Filter

            obj.osheet = oWb.Worksheets(6)
            Dim myfilter As New System.Text.StringBuilder


            'obj.strsql = "(select invid,invdate,orderno,customerid,customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty as qty" & startdate.Year - 1 & ",totalsales as totalsales" & startdate.Year - 1 & ",totalcost as totalcost" & startdate.Year - 1 & ",null::integer as qty" & startdate.Year & ",null::numeric(13,2) as totalsales" & startdate.Year & ",null::numeric(15,5) as totalcost" & startdate.Year &
            '             " from sales.tx where invdate >= " & DateFormatyyyyMMdd(CDate(startdate.Year - 1 & "-" & startdate.Month & "-" & startdate.Day)) & " and invdate <= " & DateFormatyyyyMMdd(CDate(enddate.Year - 1 & "-" & enddate.Month & "-" & enddate.Day)) & mycriteria & " order by invdate) union all " &
            '             "(select invid,invdate,orderno,customerid,customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost" &
            '             " from sales.tx where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & " order by invdate)"

            If PublicClass.myLocation = MyLocation.HongKong Then
                obj.strsql = "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty as qty" & startdate.Year - 1 & ",totalsales as totalsales" & startdate.Year - 1 & ",totalcost as totalcost" & startdate.Year - 1 & ",null::integer as qty" & startdate.Year & ",null::numeric(13,2) as totalsales" & startdate.Year & ",null::numeric(15,5) as totalcost" & startdate.Year & ",qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                         " from sales.tx " &
                         " left join sales.customer c on c.customerid = tx.customerid " &
                         " where invdate >= " & DateFormatyyyyMMdd(CDate(startdate.Year - 1 & "-" & startdate.Month & "-" & startdate.Day)) & " and invdate <= " & DateFormatyyyyMMdd(CDate(enddate.Year - 1 & "-12-31")) & mycriteria & " order by invdate) union all " &
                         "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                         " from sales.tx " &
                         " left join sales.customer c on c.customerid = tx.customerid " &
                         " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & " order by invdate)"
            Else
                'obj.strsql = "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty as qty" & startdate.Year - 1 & ",totalsales as totalsales" & startdate.Year - 1 & ",totalcost as totalcost" & startdate.Year - 1 & ",null::integer as qty" & startdate.Year & ",null::numeric(13,2) as totalsales" & startdate.Year & ",null::numeric(15,5) as totalcost" & startdate.Year & ",qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                '             " from sales.txtw tx " &
                '             " left join sales.customer c on c.customerid = tx.customerid " &
                '             " where invdate >= " & DateFormatyyyyMMdd(CDate(startdate.Year - 1 & "-" & startdate.Month & "-" & startdate.Day)) & " and invdate <= " & DateFormatyyyyMMdd(CDate(enddate.Year - 1 & "-12-31")) & mycriteria & " order by invdate) union all " &
                '             "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                '             " from sales.txtw tx " &
                '             " left join sales.customer c on c.customerid = tx.customerid " &
                '             " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & " order by invdate)"
                'obj.strsql = "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location," &
                '             " date_part('month',invdate) as month,sales.get_salesreturn(invid) as salesreturn,sales.get_sales(invid,totalsales) as sales,sales.get_return(invid,totalsales) as return,sales.get_salesreversal(invid,totalsales) as salesreversal," &
                '             " custtype as channel,custname,f.type as ec,f.id as famlv1,merch,storename,mlacode,m.mlaname ,posid,od,invdate as filterdate1,invdate as filterdate2" &
                '             " from sales.txtw tx " &
                '             " left join sales.customer c on c.customerid = tx.customerid " &
                '             " left join sales.tw_family f on f.familyname = tx.productfamily" &
                '             " left join sales.mla m on m.id = tx.mlacode" &
                '             " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & " order by invdate)"
                'obj.strsql = "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location," &
                '             " date_part('month',invdate) as month,retur as salesreturn,sales.get_sales(retur,totalsales) as sales,sales.get_return(retur,totalsales) as return,sales.get_salesreversal(retur,totalsales) as salesreversal," &
                '             " custtype as channel,custname,f.type as ec,f.id as famlv1,merch,storename,mlacode,m.mlaname ,posid,od,invdate as filterdate1,invdate as filterdate2" &
                '             " from sales.txtw tx " &
                '             " left join sales.customer c on c.customerid = tx.customerid " &
                '             " left join sales.tw_family f on f.familyname = tx.productfamily" &
                '             " left join sales.mla m on m.id = tx.mlacode" &
                '             " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & " order by invdate)"
                obj.strsql = "with cmmf as (select distinct cmmf,first_value(materialdesc) over (partition by cmmf order by invdate desc,cmmf,materialdesc  )as materialdesc from sales.txtw) (select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,tx.cmmf,sbu,productfamily,brand,cmmf.materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location," &
                             " date_part('month',invdate) as month,retur as salesreturn,sales.get_sales(retur,totalsales) as sales,sales.get_return(retur,totalsales) as return,sales.get_salesreversal(retur,totalsales) as salesreversal," &
                             " custtype as channel,custname,f.type as ec,f.id as famlv1,merch,storename,mlacode,m.mlaname ,posid,od,invdate as filterdate1,invdate as filterdate2" &
                             " from sales.txtw tx " &
                             " left join cmmf on cmmf.cmmf = tx.cmmf" &
                             " left join sales.customer c on c.customerid = tx.customerid " &
                             " left join sales.tw_family f on f.familyname = tx.productfamily" &
                             " left join sales.mla m on m.id = tx.mlacode" &
                             " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & " order by invdate)"

            End If


            obj.osheet.Name = "DATA"

            FillWorksheet(obj.osheet, obj.strsql, DbAdapter1)
            Dim lastrow = obj.osheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row

            If lastrow > 1 Then
                ProgressReport(2, "Generating Pivot Tables..")
                'oXl.Visible = True
                CreatePivotTable(oXl, oWb, 1, enddate)
                'createchart(oWb, 1, errmsg)
            End If

            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()
            'Filename = ValidateFileName(Filename, Filename & "\" & String.Format("Sales-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day)))
            Filename = ValidateFileName(SelectedPath, Filename)

            ProgressReport(2, "Done ")
            ProgressReport(5, "Saving File ...")
            oWb.SaveAs(Filename)
            ProgressReport(5, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            ProgressReport(2, "")
            errmsg = ex.Message
        Finally
            'clear excel from memory
            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End Try
        Return result
    End Function

    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    'TextBox2.Text = message
                    Me.ToolStripStatusLabel1.Text = message
                Case 3
                    'TextBox3.Text = message
                    Me.ToolStripStatusLabel2.Text = message
                Case 4
                    'TextBox1.Text = message
                    Me.ToolStripStatusLabel3.Text = message
                Case 5
                    'ToolStripStatusLabel1.Text = message
                    'ComboBox1.DataSource = bs
                    'ComboBox1.DisplayMember = "typeofitem"
                    'ComboBox1.ValueMember = "typeofitemid"
                    Me.ToolStripStatusLabel3.Text = message
                Case 6
                    'CheckedListBox1.DataSource = salesmanbs
                    'CheckedListBox1.DisplayMember = "salesman"
                    'CheckedListBox1.ValueMember = "salesman"
                Case 7
                    getpropertyname()
            End Select

        End If

    End Sub

    Private Sub FormOrderStatusReport_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Application.DoEvents()

    End Sub

    Private Function CreateWorksheet(ByVal obj As Object) As Long
        Dim osheet = DirectCast(obj, ThreadPoolObj).osheet
        osheet.Name = DirectCast(obj, ThreadPoolObj).Name
        ProgressReport(5, "Waiting for the query to be executed..." & DirectCast(obj, ThreadPoolObj).osheet.Name)
        Dim sqlstr = DirectCast(obj, ThreadPoolObj).strsql
        FillWorksheet(osheet, sqlstr, DbAdapter1)
        Dim lastrow = osheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row
        Return lastrow
    End Function



    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Function getselected(ByVal sender As Object) As String
        Dim myobj = DirectCast(sender, CheckedListBox)


        Dim sb As New StringBuilder
        Return sb.ToString
    End Function

    Private Sub createchart(ByVal oWb As Excel.Workbook, ByVal sheetnum As Integer, ByVal message As String)
        'Create Name Range
        'Try
        '    Dim osheet = oWb.Worksheets(1)
        '    Dim myChart = osheet.ChartObjects(1).Chart
        '    myChart.SeriesCollection(1).XValues = "=PivotTables!miromyrange"
        '    myChart.SeriesCollection(1).Values = "=PivotTables!miroiplt"
        '    myChart.SeriesCollection(1).Name = "Average of Lead Time"
        '    myChart.SeriesCollection(2).XValues = "=PivotTables!miromyrange"
        '    myChart.SeriesCollection(2).Values = "=PivotTables!mirotargetiplt"
        '    myChart.SeriesCollection(2).Name = "Target 95%"
        '    myChart.SeriesCollection(3).XValues = "=PivotTables!miromyrange"
        '    myChart.SeriesCollection(3).Values = "=PivotTables!miroipltle7"
        '    myChart.SeriesCollection(3).Name = "%<=7 Days"


        '    myChart = osheet.ChartObjects(2).Chart
        '    myChart.SeriesCollection(1).XValues = "=PivotTables!miromyrangesao"
        '    myChart.SeriesCollection(1).Values = "=PivotTables!miroipltsao"
        '    myChart.SeriesCollection(1).Name = "Average of Lead Time"
        '    myChart.SeriesCollection(2).XValues = "=PivotTables!miromyrangesao"
        '    myChart.SeriesCollection(2).Values = "=PivotTables!mirotargetipltsao"
        '    myChart.SeriesCollection(2).Name = "Target 95%"
        '    myChart.SeriesCollection(3).XValues = "=PivotTables!miromyrangesao"
        '    myChart.SeriesCollection(3).Values = "=PivotTables!miroipltle7sao"
        '    myChart.SeriesCollection(3).Name = "%<=7 Days"
        'Catch ex As Exception
        '    message = ex.Message
        'End Try

    End Sub

    Private Sub CreatePivotTable(ByVal oxl As Excel.Application, ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal mydate As Date)
        Dim osheet As Excel.Worksheet
        'oWb.Names.Add("dbRange", RefersToR1C1:="=OFFSET('data'!R1C1,0,0,COUNTA('data'!C1),COUNTA('data'!R1))")
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        'oWb.Worksheets("Sheet1").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R6C15", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.Cells(1, 1) = "SEB ASIA LTD - TAIWAN BRANCH"
        osheet.Cells(2, 1) = "GROSS SALES (including sales return)"
        osheet.Cells(3, 1) = "Currency: NTD"
        With osheet.Range("A1:A2")
            .Font.Size = 20
        End With

        With osheet.Range("A3").Font
            .Size = 10
            .FontStyle = "Bold"
            .Color = -16776961
            .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
        End With

        osheet.PivotTables("PivotTable1").calculatedfields.add(" GrandTotal", "=sales + return + salesreversal", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add(" Proportion", "=return /(sales + salesreversal)", True)

        osheet.PivotTables("PivotTable1").Pivotfields("mlaname").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("mlacode").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesreturn").orientation = Excel.XlPivotFieldOrientation.xlPageField

        osheet.PivotTables("PivotTable1").Pivotfields("famlv1").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("storename").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("materialdesc").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("channel").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("merch").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesman").orientation = Excel.XlPivotFieldOrientation.xlPageField



        osheet.PivotTables("PivotTable1").Pivotfields("custname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales"), " Totals Sales ", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" Totals Sales ").numberformat = "#,##0_);(#,##0)"
        osheet.PivotTables("PivotTable1").Pivotfields("month").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        osheet.Range("A18").Select()
        oxl.ActiveWindow.FreezePanes = True

        osheet.Name = "Gross Sales"

        osheet.Cells.EntireColumn.AutoFit()

        ''Second PivotTable
        isheet = isheet + 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        'oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        oWb.Worksheets("Gross Sales").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.Cells(1, 1) = "SEB ASIA LTD - TAIWAN BRANCH"
        osheet.Cells(2, 1) = "Quantity (including sales return)"

        With osheet.Range("A1:A2")
            .Font.Size = 20
        End With


        osheet.PivotTables("PivotTable1").Pivotfields("mlaname").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("mlacode").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesreturn").orientation = Excel.XlPivotFieldOrientation.xlPageField

        osheet.PivotTables("PivotTable1").Pivotfields("famlv1").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("storename").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("materialdesc").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("channel").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("merch").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesman").orientation = Excel.XlPivotFieldOrientation.xlPageField



        osheet.PivotTables("PivotTable1").Pivotfields("custname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalqty"), " Totals QTY ", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" Totals QTY ").numberformat = "#,##0"
        osheet.PivotTables("PivotTable1").Pivotfields("month").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        osheet.Range("A18").Select()
        oxl.ActiveWindow.FreezePanes = True
        osheet.Name = "QTY"


        isheet = isheet + 1

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        'oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        oWb.Worksheets("Gross Sales").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R8C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.Cells(1, 1) = "SEB ASIA LTD - TAIWAN BRANCH"
        osheet.Cells(2, 1) = "Sales return on gross sales in % by customer"
        osheet.Cells(3, 1) = "Currency: NTD"
        With osheet.Range("A1:A2")
            .Font.Size = 20
        End With
        With osheet.Range("A3").Font
            .Size = 10
            .FontStyle = "Bold"
            .Color = -16776961
            .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
        End With
        osheet.PivotTables("PivotTable1").Pivotfields("month").orientation = Excel.XlPivotFieldOrientation.xlPageField

        osheet.PivotTables("PivotTable1").Pivotfields("custname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sales"), " Sales ", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("salesreversal"), " Sales Reversal ", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("return"), " Return ", Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").PivotFields(" Sales ").numberformat = "#,##0_);(#,##0)"
        osheet.PivotTables("PivotTable1").PivotFields(" Sales Reversal ").numberformat = "#,##0_);(#,##0)"
        osheet.PivotTables("PivotTable1").PivotFields(" Return ").numberformat = "#,##0_);(#,##0)"

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields(" GrandTotal"), " GrandTotal ", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields(" Proportion"), " Proportion on sales (in %)", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" GrandTotal ").numberformat = "#,##0_);(#,##0)"
        osheet.PivotTables("PivotTable1").PivotFields(" Proportion on sales (in %)").numberformat = "0%"

       
        osheet.Name = "Sales Return"


        osheet.Cells.EntireColumn.AutoFit()
        isheet = 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)

    End Sub


    Private Sub ReportSales_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Application.DoEvents()
        DateTimePicker1.Value = CDate(Year(Today) & "/01/01")
        DateTimePicker2.Value = Today
        'Me.ToolStripStatusLabel1.Text = "Populating Salesman. Please wait..."
        'myQueryThread.Start()
        'get propertyname

    End Sub

    Sub DoQuery()
        Dataset1 = New DataSet
        Dim sqlstr As String
        If PublicClass.myLocation = MyLocation.HongKong Then
            sqlstr = "select 'Select All' as salesman union all" &
                               " select 'All Salesman' as customerid union all" &
                               " (select distinct salesman from sales.tx order by salesman);"
        Else
            sqlstr = "select 'Select All' as salesman union all" &
                                           " select 'All Salesman' as customerid union all" &
                                           " (select distinct salesman from sales.txtw order by salesman);"
        End If




        If DbAdapter1.getDataSet(sqlstr, Dataset1) Then
            Dataset1.Tables(0).TableName = "Salesman"
            salesmanbs.DataSource = Dataset1.Tables(0)
            ProgressReport(6, "InitCheckedListBox")

            ProgressReport(2, "Populating Salesman. Done.")
            'ProgressReport(7, "Done")
        Else
            ProgressReport(5, "Error while loading Dataset.")
        End If
    End Sub

    'Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
    '    CheckedListBox_SelectedIndexChanged(sender, e)
    'End Sub

    Private Sub getpropertyname()
        Dim tl As ITypedList = TryCast(salesmanbs, ITypedList)
        Dim a = tl.GetItemProperties(Nothing)
        Dim dt As DataTable = TryCast(salesmanbs.DataSource, DataTable)
        MessageBox.Show(dt.Columns(0).DataType.Name)
    End Sub


    
End Class