Imports System.Threading
Imports Microsoft.Office.Interop
Imports Sales.SharedClass
Imports System.Text
Imports Sales.PublicClass
Public Class FormCustomerBU
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
    Dim customerbs As New BindingSource
    Dim sbubs As New BindingSource
    Dim mycriteria As String = String.Empty
    Dim mybuCriteria As String = String.Empty
    Dim SelectedPath As String = String.Empty
    Dim askmetoopen As Boolean = False
    Dim mergedReport As Boolean = False

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
        If mergedReport Then
            DoWorkMerged()
        Else
            DoWorkNotMerged()
        End If

        
    End Sub

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

            If PublicClass.myLocation = MyLocation.HongKong Then
                'obj.strsql = "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty as qty" & startdate.Year - 1 & ",totalsales as totalsales" & startdate.Year - 1 & ",totalcost as totalcost" & startdate.Year - 1 & ",null::integer as qty" & startdate.Year & ",null::numeric(13,2) as totalsales" & startdate.Year & ",null::numeric(15,5) as totalcost" & startdate.Year & ",qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                '         " from sales.tx " &
                '         " left join sales.customer c on c.customerid = tx.customerid " &
                '         " where invdate >= " & DateFormatyyyyMMdd(CDate(startdate.Year - 1 & "-" & startdate.Month & "-" & startdate.Day)) & " and invdate <= " & DateFormatyyyyMMdd(CDate(enddate.Year - 1 & "-12-31")) & mycriteria & mybuCriteria & " order by invdate) union all " &
                '         "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                '         " from sales.tx " &
                '         " left join sales.customer c on c.customerid = tx.customerid " &
                '         " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & mybuCriteria & " order by invdate)"
                If HelperClass1.UserInfo.IsSalesTeamHK Then
                    obj.strsql = "with cmmf as (" &
                          "select distinct cmmf,first_value(materialdesc) over (partition by cmmf order by invdate desc,cmmf,materialdesc  )as materialdesc  from sales.tx)" &
                          "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,tx.cmmf,sbu,productfamily,brand,cmmf.materialdesc,supplierid,qty as qty" & startdate.Year - 1 & ",totalsales as totalsales" & startdate.Year - 1 & ",totalcost as totalcost" & startdate.Year - 1 & ",null::integer as qty" & startdate.Year & ",null::numeric(13,2) as totalsales" & startdate.Year & ",null::numeric(15,5) as totalcost" & startdate.Year & ",qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                         " from sales.tx " &
                         " left join sales.customer c on c.customerid = tx.customerid " &
                         " left join cmmf on cmmf.cmmf = tx.cmmf " &
                         " where invdate >= " & DateFormatyyyyMMdd(CDate(startdate.Year - 1 & "-" & startdate.Month & "-" & startdate.Day)) & " and invdate <= " & DateFormatyyyyMMdd(CDate(enddate.Year - 1 & "-12-31")) & mycriteria & mybuCriteria & " order by invdate) union all " &
                         "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,tx.cmmf,sbu,productfamily,brand,cmmf.materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                         " from sales.tx " &
                         " left join sales.customer c on c.customerid = tx.customerid " &
                         " left join cmmf on cmmf.cmmf = tx.cmmf " &
                         " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & mybuCriteria & " order by invdate)"
                Else
                    obj.strsql = "with cmmf as (" &
                          "select distinct cmmf,first_value(materialdesc) over (partition by cmmf order by invdate desc,cmmf,materialdesc  )as materialdesc  from sales.tx)" &
                          "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,tx.cmmf,sbu,productfamily,brand,cmmf.materialdesc,supplierid,qty as qty" & startdate.Year - 1 & ",totalsales as totalsales" & startdate.Year - 1 & ",null::numeric(15,5) as totalcost" & startdate.Year - 1 & ",null::integer as qty" & startdate.Year & ",null::numeric(13,2) as totalsales" & startdate.Year & ",null::numeric(15,5) as totalcost" & startdate.Year & ",qty as totalqty ,totalsales as totalsales,null::numeric(15,5) as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                         " from sales.tx " &
                         " left join sales.customer c on c.customerid = tx.customerid " &
                         " left join cmmf on cmmf.cmmf = tx.cmmf " &
                         " where invdate >= " & DateFormatyyyyMMdd(CDate(startdate.Year - 1 & "-" & startdate.Month & "-" & startdate.Day)) & " and invdate <= " & DateFormatyyyyMMdd(CDate(enddate.Year - 1 & "-12-31")) & mycriteria & mybuCriteria & " order by invdate) union all " &
                         "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,tx.cmmf,sbu,productfamily,brand,cmmf.materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,null::numeric(15,5) as totalcost,qty as totalqty ,totalsales as totalsales,null::numeric(15,5) as totalcost ,region,location,invdate as filterdate1,invdate as filterdate2" &
                         " from sales.tx " &
                         " left join sales.customer c on c.customerid = tx.customerid " &
                         " left join cmmf on cmmf.cmmf = tx.cmmf " &
                         " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & mybuCriteria & " order by invdate)"

                End If
                

            Else
                'obj.strsql = "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty as qty" & startdate.Year - 1 & ",totalsales as totalsales" & startdate.Year - 1 & ",totalcost as totalcost" & startdate.Year - 1 & ",null::integer as qty" & startdate.Year & ",null::numeric(13,2) as totalsales" & startdate.Year & ",null::numeric(15,5) as totalcost" & startdate.Year & ",qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                '         " from sales.txtw tx" &
                '         " left join sales.customer c on c.customerid = tx.customerid " &
                '         " where invdate >= " & DateFormatyyyyMMdd(CDate(startdate.Year - 1 & "-" & startdate.Month & "-" & startdate.Day)) & " and invdate <= " & DateFormatyyyyMMdd(CDate(enddate.Year - 1 & "-12-31")) & mycriteria & mybuCriteria & " order by invdate) union all " &
                '         "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                '         " from sales.txtw tx" &
                '         " left join sales.customer c on c.customerid = tx.customerid " &
                '         " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & mybuCriteria & " order by invdate)"
                'obj.strsql = "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty as qty" & startdate.Year - 1 & ",totalsales as totalsales" & startdate.Year - 1 & ",totalcost as totalcost" & startdate.Year - 1 & ",null::integer as qty" & startdate.Year & ",null::numeric(13,2) as totalsales" & startdate.Year & ",null::numeric(15,5) as totalcost" & startdate.Year & ",qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location," &
                '    " date_part('month',invdate) as month,sales.get_salesreturn(invid) as salesreturn,sales.get_sales(invid,totalsales) as sales,sales.get_return(invid,totalsales) as return,sales.get_salesreversal(invid,totalsales) as salesreversal," &
                '    " custtype as channel,custname,f.type as ec,f.id as famlv1,merch,storename,mlacode,m.mlaname ,posid,od,invdate as filterdate1,invdate as filterdate2" &
                '    " from sales.txtw tx " &
                '    " left join sales.customer c on c.customerid = tx.customerid " &
                '    " left join sales.tw_family f on f.familyname = tx.productfamily" &
                '    " left join sales.mla m on m.id = tx.mlacode" &
                '    " where invdate >= " & DateFormatyyyyMMdd(CDate(startdate.Year - 1 & "-" & startdate.Month & "-" & startdate.Day)) & " and invdate <= " & DateFormatyyyyMMdd(CDate(enddate.Year - 1 & "-12-31")) & mycriteria & mybuCriteria & " order by invdate) union all " &
                '    "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location," &
                '    " date_part('month',invdate) as month,sales.get_salesreturn(invid) as salesreturn,sales.get_sales(invid,totalsales) as sales,sales.get_return(invid,totalsales) as return,sales.get_salesreversal(invid,totalsales) as salesreversal," &
                '    " custtype as channel,custname,f.type as ec,f.id as famlv1,merch,storename,mlacode,m.mlaname ,posid,od,invdate as filterdate1,invdate as filterdate2" &
                '    " from sales.txtw tx " &
                '    " left join sales.customer c on c.customerid = tx.customerid " &
                '    " left join sales.tw_family f on f.familyname = tx.productfamily" &
                '    " left join sales.mla m on m.id = tx.mlacode" &
                '    " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & mybuCriteria & " order by invdate)"
                obj.strsql = "with cmmf as (" &
                    " select distinct cmmf,first_value(materialdesc) over (partition by cmmf order by invdate desc,cmmf,materialdesc  )as materialdesc  from sales.txtw) " &
                    "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,tx.cmmf,sbu,productfamily,brand,cmmf.materialdesc,supplierid,qty as qty" & startdate.Year - 1 & ",totalsales as totalsales" & startdate.Year - 1 & ",totalcost as totalcost" & startdate.Year - 1 & ",null::integer as qty" & startdate.Year & ",null::numeric(13,2) as totalsales" & startdate.Year & ",null::numeric(15,5) as totalcost" & startdate.Year & ",qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location," &
                   " date_part('month',invdate) as month,retur as salesreturn,sales.get_sales(retur,totalsales) as sales,sales.get_return(retur,totalsales) as return,sales.get_salesreversal(retur,totalsales) as salesreversal," &
                   " custtype as channel,custname,f.type as ec,f.id as famlv1,merch,storename,mlacode,m.mlaname ,posid,od,invdate as filterdate1,invdate as filterdate2" &
                   " from sales.txtw tx " &
                   " left join sales.customer c on c.customerid = tx.customerid " &
                   " left join cmmf on cmmf.cmmf = tx.cmmf " &
                   " left join sales.tw_family f on f.familyname = tx.productfamily" &
                   " left join sales.mla m on m.id = tx.mlacode" &
                   " where invdate >= " & DateFormatyyyyMMdd(CDate(startdate.Year - 1 & "-" & startdate.Month & "-" & startdate.Day)) & " and invdate <= " & DateFormatyyyyMMdd(CDate(enddate.Year - 1 & "-12-31")) & mycriteria & mybuCriteria & " order by invdate) union all " &
                   "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,tx.cmmf,sbu,productfamily,brand,cmmf.materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location," &
                   " date_part('month',invdate) as month,retur as salesreturn,sales.get_sales(retur,totalsales) as sales,sales.get_return(retur,totalsales) as return,sales.get_salesreversal(retur,totalsales) as salesreversal," &
                   " custtype as channel,custname,f.type as ec,f.id as famlv1,merch,storename,mlacode,m.mlaname ,posid,od,invdate as filterdate1,invdate as filterdate2" &
                   " from sales.txtw tx " &
                   " left join sales.customer c on c.customerid = tx.customerid " &
                   " left join cmmf on cmmf.cmmf = tx.cmmf " &
                   " left join sales.tw_family f on f.familyname = tx.productfamily" &
                   " left join sales.mla m on m.id = tx.mlacode" &
                   " where invdate >= " & DateFormatyyyyMMdd(startdate) & " and invdate <= " & DateFormatyyyyMMdd(enddate) & mycriteria & mybuCriteria & " order by invdate)"

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
                    CheckedListBox1.DataSource = customerbs
                    CheckedListBox1.DisplayMember = "customername"
                    CheckedListBox1.ValueMember = "customername"

                    CheckedListBox2.DataSource = sbubs
                    CheckedListBox2.DisplayMember = "sbu"
                    CheckedListBox2.ValueMember = "sbu"
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
        Try
            Dim osheet = oWb.Worksheets(1)
            Dim myChart = osheet.ChartObjects(1).Chart
            myChart.SeriesCollection(1).XValues = "=PivotTables!miromyrange"
            myChart.SeriesCollection(1).Values = "=PivotTables!miroiplt"
            myChart.SeriesCollection(1).Name = "Average of Lead Time"
            myChart.SeriesCollection(2).XValues = "=PivotTables!miromyrange"
            myChart.SeriesCollection(2).Values = "=PivotTables!mirotargetiplt"
            myChart.SeriesCollection(2).Name = "Target 95%"
            myChart.SeriesCollection(3).XValues = "=PivotTables!miromyrange"
            myChart.SeriesCollection(3).Values = "=PivotTables!miroipltle7"
            myChart.SeriesCollection(3).Name = "%<=7 Days"


            myChart = osheet.ChartObjects(2).Chart
            myChart.SeriesCollection(1).XValues = "=PivotTables!miromyrangesao"
            myChart.SeriesCollection(1).Values = "=PivotTables!miroipltsao"
            myChart.SeriesCollection(1).Name = "Average of Lead Time"
            myChart.SeriesCollection(2).XValues = "=PivotTables!miromyrangesao"
            myChart.SeriesCollection(2).Values = "=PivotTables!mirotargetipltsao"
            myChart.SeriesCollection(2).Name = "Target 95%"
            myChart.SeriesCollection(3).XValues = "=PivotTables!miromyrangesao"
            myChart.SeriesCollection(3).Values = "=PivotTables!miroipltle7sao"
            myChart.SeriesCollection(3).Name = "%<=7 Days"
        Catch ex As Exception
            message = ex.Message
        End Try

    End Sub

    Private Sub CreatePivotTable(ByVal oxl As Excel.Application, ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal mydate As Date)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").calculatedfields.add("qtydif", "=qty" & mydate.Year & " - qty" & mydate.Year - 1, True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("qtydifpct", "=qty" & mydate.Year & " / qty" & mydate.Year - 1 & " - 1", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("amountdif", "=totalsales" & mydate.Year & " - totalsales" & mydate.Year - 1, True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("amountdifpct", "=totalsales" & mydate.Year & " / totalsales" & mydate.Year - 1 & " - 1", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("margin" & mydate.Year - 1 & "pct", "=(totalsales" & mydate.Year - 1 & " - totalcost" & mydate.Year - 1 & ")/ totalsales" & mydate.Year - 1, True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("margin" & mydate.Year & "pct", "=(totalsales" & mydate.Year & " - totalcost" & mydate.Year & ")/ totalsales" & mydate.Year, True)

        osheet.PivotTables("PivotTable1").Pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("A8").Group(True, True, Periods:={False, False, False, False, True, False, True})
        osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable1").pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlHidden

        osheet.PivotTables("PivotTable1").Pivotfields("filterdate1").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("A8").Group(True, True, Periods:={False, False, False, False, True, True, True})
        osheet.PivotTables("PivotTable1").pivotfields("Years2").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable1").pivotfields("Quarters").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable1").pivotfields("filterdate1").orientation = Excel.XlPivotFieldOrientation.xlHidden

        'For Each item As Object In osheet.PivotTables("PivotTable1").pivotfields("Years").pivotitems
        '    Dim obj = DirectCast(item, Excel.PivotItem)
        '    If obj.Value.ToString <> mydate.Year.ToString Then
        '        obj.Visible = False
        '    End If
        'Next

        osheet.PivotTables("PivotTable1").Pivotfields("location").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("customername").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesman").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("sbu").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlRowField

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year - 1), " Qty " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year - 1), " Total Sales " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year), " Qty " & mydate.Year, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year), " Totals Sales " & mydate.Year, Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydif"), " Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydifpct"), " %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdif"), " Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdifpct"), "%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year - 1 & "pct"), " %Margin " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year & "pct"), " %Margin " & mydate.Year, Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year - 1).NumberFormat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Total Sales " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year).numberformat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Totals Sales " & mydate.Year).numberformat = "#,##0.00"


        osheet.PivotTables("PivotTable1").PivotFields(" Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).NumberFormat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields("%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year).numberformat = "0.00%"

        osheet.Name = "YTD"

        osheet.Cells.EntireColumn.AutoFit()

        'Second PivotTable
        isheet = isheet + 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets("YTD").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").Pivotfields("location").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("customername").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesman").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("invdate").currentpage = Format(mydate, "MMM")
        'For Each item As Object In osheet.PivotTables("PivotTable2").pivotfields("Years").pivotitems
        '    Dim obj = DirectCast(item, Excel.PivotItem)
        '    If obj.Value.ToString <> mydate.Year.ToString Then
        '        obj.Visible = False
        '    End If
        'Next



        osheet.PivotTables("PivotTable1").Pivotfields("sbu").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlRowField

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year - 1), " Qty " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year - 1), " Total Sales " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year), " Qty " & mydate.Year, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year), " Totals Sales " & mydate.Year, Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydif"), " Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydifpct"), " %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdif"), " Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdifpct"), "%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year - 1 & "pct"), " %Margin " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year & "pct"), " %Margin " & mydate.Year, Excel.XlConsolidationFunction.xlSum)


        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year - 1).NumberFormat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Total Sales " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year).numberformat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Totals Sales " & mydate.Year).numberformat = "#,##0.00"

        osheet.PivotTables("PivotTable1").PivotFields(" Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).NumberFormat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields("%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year).numberformat = "0.00%"

        osheet.Name = "MTD"

        osheet.Cells.EntireColumn.AutoFit()

        'Third PivotTable
        isheet = isheet + 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets("YTD").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").Pivotfields("customername").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesman").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlPageField


        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("productid").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("productid").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        osheet.PivotTables("PivotTable1").Pivotfields("materialdesc").orientation = Excel.XlPivotFieldOrientation.xlRowField



        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year - 1), " Qty " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year - 1), " Total Sales " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year), " Qty " & mydate.Year, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year), " Totals Sales " & mydate.Year, Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydif"), " Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydifpct"), " %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdif"), " Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdifpct"), "%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year - 1 & "pct"), " %Margin " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year & "pct"), " %Margin " & mydate.Year, Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year - 1).NumberFormat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Total Sales " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year).numberformat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Totals Sales " & mydate.Year).numberformat = "#,##0.00"


        osheet.PivotTables("PivotTable1").PivotFields(" Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).NumberFormat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields("%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year).numberformat = "0.00%"

        osheet.Name = "Details"

        osheet.Cells.EntireColumn.AutoFit()

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)

        'Qty
        'Fourth Pivot Table
        isheet = isheet + 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets("YTD").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").Pivotfields("sbu").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("customername").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("Years2").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("Years2").Caption = "Filter Years"

        osheet.PivotTables("PivotTable1").pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").pivotfields("invdate").caption = "Months"
        'osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("Months").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'Dim mypivot As Excel.PivotItem
        'For Each mypivot In osheet.PivotTables("PivotTable1").pivotfields("Years").PivotItems
        '    mypivot.Value = "Qty " + mypivot.Value
        'Next

        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("productid").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("productid").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").Pivotfields("materialdesc").orientation = Excel.XlPivotFieldOrientation.xlRowField



        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalqty"), "Total Quantity", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year - 1), " Qty" & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year), " Qty" & mydate.Year, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" Qty" & mydate.Year - 1).NumberFormat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Qty" & mydate.Year).NumberFormat = "#,##0"

        osheet.PivotTables("PivotTable1").ShowTableStyleColumnStripes = True
        osheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight16"
        oWb.TableStyles("PivotStyleLight16").Duplicate("PivotStyleLight16 2" _
    )
        With oxl.ActiveWorkbook.TableStyles("PivotStyleLight16 2").TableStyleElements(Excel.XlTableStyleElementType.xlColumnSubheading1).Borders(Excel.XlBordersIndex.xlEdgeTop)
            .Weight = 2
            .LineStyle = 1
        End With
        With oxl.ActiveWorkbook.TableStyles("PivotStyleLight16 2").TableStyleElements(Excel.XlTableStyleElementType.xlColumnSubheading1).Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .Weight = 2
            .LineStyle = 1
        End With
        With oxl.ActiveWorkbook.TableStyles("PivotStyleLight16 2").TableStyleElements(Excel.XlTableStyleElementType.xlColumnSubheading1).Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .Weight = 2
            .LineStyle = 1
        End With
        With oxl.ActiveWorkbook.TableStyles("PivotStyleLight16 2").TableStyleElements(Excel.XlTableStyleElementType.xlColumnSubheading1).Borders(Excel.XlBordersIndex.xlEdgeRight)
            .Weight = 2
            .LineStyle = 1
        End With
        osheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight16 2"


        osheet.Name = "Quantity"

        osheet.Cells.EntireColumn.AutoFit()

        'Sales
        'Fifth Pivot Table
        isheet = isheet + 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets("YTD").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").Pivotfields("sbu").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("customername").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("Years2").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("Years2").Caption = "Filter Years"

        osheet.PivotTables("PivotTable1").pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").pivotfields("invdate").caption = "Months"
        'osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("Months").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}


        'For Each mypivot In osheet.PivotTables("PivotTable1").pivotfields("Years").PivotItems
        '    mypivot.Value = "Sales Amt " + mypivot.Value
        'Next

        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("productid").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("productid").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").Pivotfields("materialdesc").orientation = Excel.XlPivotFieldOrientation.xlRowField



        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales"), "Sales Amount", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year - 1), "Sales Amount" & DateTimePicker1.Value.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year), "Sales Amount" & DateTimePicker1.Value.Year, Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").PivotFields("Sales Amount" & mydate.Year - 1).NumberFormat = "#,##0,00"
        osheet.PivotTables("PivotTable1").PivotFields("Sales Amount" & mydate.Year).NumberFormat = "#,##0,00"

        osheet.PivotTables("PivotTable1").ShowTableStyleColumnStripes = True
        osheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight16 2"
        osheet.Name = "Sales Amt"

        osheet.Cells.EntireColumn.AutoFit()
        isheet = 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)

    End Sub


    Private Sub ReportSales_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Application.DoEvents()
        DateTimePicker1.Value = CDate(Year(Today) & "/01/01")
        DateTimePicker2.Value = Today
        Me.ToolStripStatusLabel1.Text = "Populating Customer. Please wait..."
        myQueryThread.Start()

    End Sub

    Sub DoQuery()
        Dataset1 = New DataSet
        Dim sqlstr As String
        If PublicClass.myLocation = MyLocation.HongKong Then
            sqlstr = "select 'Select All' as customername, '0' as customerid union all" &
                               " select 'All Customer', '1' as customerid union all" & _
                               " (select distinct customername,customerid from sales.tx order by customername);" &
                               "select 'Select All' as sbu union all" &
                               " (select distinct sbu from sales.tx order by sbu);"
        Else
            'sqlstr = "select 'Select All' as customername, '0' as customerid union all" &
            '                       " select 'All Customer', '1' as customerid union all" & _
            '                       " (select distinct customername,customerid from sales.txtw order by customername);" &
            '                       "select 'Select All' as sbu union all" &
            '                       " (select distinct sbu from sales.txtw order by sbu);"
            sqlstr = "select 'Select All' as customername  union all" &
                              " select 'All Customer'  union all" & _
                              " (select distinct custname from sales.txtw order by custname);" &
                                   "select 'Select All' as sbu union all" &
                                   " (select distinct sbu from sales.txtw order by sbu);"
        End If
        


        If DbAdapter1.getDataSet(sqlstr, Dataset1) Then
            Dataset1.Tables(0).TableName = "Customer"
            customerbs.DataSource = Dataset1.Tables(0)
            sbubs.DataSource = Dataset1.Tables(1)
            ProgressReport(6, "InitCheckedBoxList")
            ProgressReport(2, "Populating Customer. Done.")
        Else
            ProgressReport(5, "Error while loading Dataset.")
        End If
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged, CheckedListBox2.SelectedIndexChanged
        CheckedListBox_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            CheckedListBox1.SetItemChecked(1, Not (CheckBox2.Checked))
            CheckedListBox1.SetItemChecked(0, Not (CheckBox2.Checked))
        End If
        mergedReport = CheckBox2.Checked
    End Sub

    Private Sub DoWorkMerged()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        Dim status As Boolean = False
        Dim message As String = String.Empty
        sw.Start()

        Dim chkstate As CheckState
        chkstate = CheckedListBox1.GetItemChecked(0)


        mybuCriteria = getbucriteria()
        ProgressReport(2, "Export To Excel..")
        ProgressReport(5, "")
        mycriteria = ""
        For Each item As Object In CheckedListBox1.CheckedItems
            
            Dim dr As DataRowView = DirectCast(item, DataRowView)
            Dim myvalue = dr.Item(0)


            If myvalue <> "Select All" Then

                'Filename = SelectedPath & "\" & String.Format("Sales {0} {1}-{2}.xlsx", myvalue, Format(startdate, "ddMMMyyyy"), Format(enddate, "ddMMMyyyy"))
                If myvalue <> "All Customer" Then
                    mycriteria = mycriteria & If(mycriteria = "", "", ",") & escapestr(myvalue)

                End If
                
            End If

        Next
        If mycriteria <> "" Then
            'mycriteria = " and c.customername in (" & mycriteria & ")"
            If PublicClass.myLocation = MyLocation.HongKong Then
                mycriteria = " and c.customername in (" & mycriteria & ")"
            Else
                mycriteria = " and tx.custname in (" & mycriteria & ")"
            End If
        End If
        Filename = SelectedPath & ""
        Filename = SelectedPath & "\" & String.Format("Sales {0} {1}-{2}.xlsx", "Customer-BU", Format(startdate, "ddMMMyyyy"), Format(enddate, "ddMMMyyyy"))
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

    Private Sub DoWorkNotMerged()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        Dim status As Boolean = False
        Dim message As String = String.Empty
        sw.Start()

        Dim chkstate As CheckState
        chkstate = CheckedListBox1.GetItemChecked(0)


        mybuCriteria = getbucriteria
        For Each item As Object In CheckedListBox1.CheckedItems
            ProgressReport(2, "Export To Excel..")
            ProgressReport(5, "")
            Dim dr As DataRowView = DirectCast(item, DataRowView)
            Dim myvalue = dr.Item(0)
            Filename = SelectedPath & ""

            If myvalue <> "Select All" Then
                mycriteria = ""
                'Filename = SelectedPath & "\" & String.Format("Sales {0} {1}-{2}.xlsx", myvalue, Format(startdate, "ddMMMyyyy"), Format(enddate, "ddMMMyyyy"))
                If myvalue <> "All Customer" Then
                    'mycriteria = " and c.customername=" & escapestr(myvalue)
                    If PublicClass.myLocation = MyLocation.HongKong Then
                        mycriteria = " and c.customername=" & escapestr(myvalue)
                    Else
                        mycriteria = " and tx.custname=" & escapestr(myvalue)
                    End If
                End If
                Filename = SelectedPath & "\" & String.Format("Sales {0} {1}-{2}.xlsx", myvalue, Format(startdate, "ddMMMyyyy"), Format(enddate, "ddMMMyyyy"))
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
            End If

        Next




        sw.Stop()
    End Sub

    Private Function getbucriteria() As String
        Dim myret As String = String.Empty

        Dim i As Integer = 0
        For Each item As Object In CheckedListBox2.CheckedItems
            Dim dr As DataRowView = DirectCast(item, DataRowView)
            Dim myvalue = dr.Item(0)
            If myvalue <> "Select All" Then
                myret = myret + If(myret = "", "", ",") + "'" & myvalue & "'"
            Else
                Exit For
            End If

        Next
        If myret <> "" Then
            myret = " and sbu in (" & myret & ")"
        End If
        Return myret

    End Function

End Class