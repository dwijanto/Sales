Imports System.Threading
Imports Sales.PublicClass
Imports System.Text
Imports Sales.SharedClass
Imports Microsoft.Office.Interop
Imports System.IO

Public Class ImportSales

    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim errorsb As New StringBuilder
    Dim mydate As Date

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not mythread.IsAlive Then
            'Get file
            mydate = DateTimePicker1.Value.Date
            If openfiledialog1.ShowDialog = DialogResult.OK Then
                If PublicClass.myLocation = MyLocation.HongKong Then
                    mythread = New Thread(AddressOf doWork)
                Else
                    mythread = New Thread(AddressOf doWorktw)
                End If

                mythread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doWork()
        'get other table

        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim myMessage As String = String.Empty

        If Not filldataset(DS, myMessage) Then
            ProgressReport(1, myMessage)
            Exit Sub
        End If

        Dim sqlstr As String = String.Empty
        Dim mystr As New StringBuilder
        Dim cmmfsb As New StringBuilder
        Dim sbusb As New StringBuilder
        Dim prodfamilysb As New StringBuilder
        Dim brandsb As New StringBuilder
        Dim customersb As New StringBuilder



        sw.Start()
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        'Dim mycsvfile = convertfile(openfiledialog1.FileName)
        Dim mycsvfile = convertfileHK(openfiledialog1.FileName)
        Using objTFParser = New FileIO.TextFieldParser(mycsvfile)

            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                '.SetDelimiters(",")
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                ProgressReport(1, "Read Data")
                ProgressReport(2, "Continuous off")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count >= 1 Then

                        If CDate(myrecord(1)) >= DateTimePicker1.Value.Date Then
                            'po,companycode,createdon,createdby,vendorcode,termsofpayment,purchgroup,currency,docdate,incoterm1,incoterm2
                            'invid character varying,  invdate date,  orderno character varying,  customerid character varying,  customername character varying,  reportcode character varying,  saleforce character varying,  country character varying,  custtype character varying,  salesman character varying,  shipto character varying,  productid character varying,  cmmf bigint,  sbu character varying,  productfamily character varying,  brand character varying,  materialdesc character varying,  supplierid character varying,  qty integer,  totalsales numeric(13,2),  totalcost numeric(15,5),
                            'If myrecord(12) = 2100083280 Then
                            '    Debug.Print("")
                            'End If
                            Dim result As Object
                            If myrecord(13) <> "" Then
                                Dim pkey1(0) As Object
                                pkey1(0) = myrecord(13)
                                result = DS.Tables("sbu").Rows.Find(pkey1)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables("sbu").NewRow
                                    dr.Item(0) = myrecord(13)
                                    DS.Tables("sbu").Rows.Add(dr)
                                    sbusb.Append(validstr(myrecord(13)) & vbCrLf)
                                End If
                            End If

                            If myrecord(14) <> "" Then
                                Dim pkey2(0) As Object
                                pkey2(0) = myrecord(14)
                                result = DS.Tables("prodfamily").Rows.Find(pkey2)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables("prodfamily").NewRow
                                    dr.Item(0) = myrecord(14)
                                    DS.Tables("prodfamily").Rows.Add(dr)
                                    prodfamilysb.Append(validstr(myrecord(14)) & vbCrLf)
                                End If
                            End If

                            If myrecord(15) <> "" Then
                                Dim pkey3(0) As Object
                                pkey3(0) = myrecord(15)
                                result = DS.Tables("brand").Rows.Find(pkey3)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables("brand").NewRow
                                    dr.Item(0) = myrecord(15)
                                    DS.Tables("brand").Rows.Add(dr)
                                    brandsb.Append(validstr(myrecord(15)) & vbCrLf)
                                End If
                            End If


                            If myrecord(3) <> "" Then
                                Dim pkey4(0) As Object
                                pkey4(0) = myrecord(3)
                                result = DS.Tables("customer").Rows.Find(pkey4)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables("customer").NewRow
                                    dr.Item(0) = myrecord(3)
                                    DS.Tables("customer").Rows.Add(dr)
                                    customersb.Append(validstr(myrecord(3)) & vbTab &
                                                      validstr(myrecord(4)) & vbCrLf)
                                End If
                            End If

                            If myrecord(12) <> "" Then
                                Dim pkey0(0) As Object
                                pkey0(0) = myrecord(12)
                                result = DS.Tables("cmmf").Rows.Find(pkey0)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables("cmmf").NewRow
                                    dr.Item(0) = myrecord(12)
                                    DS.Tables("cmmf").Rows.Add(dr)
                                    'cmmf,productid,sbu,prodfamily,cdesc,brand
                                    cmmfsb.Append(validlong(myrecord(12)) & vbTab &
                                                  validstr(myrecord(11)) & vbTab &
                                                  validstr(myrecord(13)) & vbTab &
                                                  validstr(myrecord(14)) & vbTab &
                                                  validstr(myrecord(16)) & vbTab &
                                                  validstr(myrecord(15)) & vbCrLf)
                                End If
                            End If

                            Try
                                myInsert.Append(validstr(myrecord(0)) & vbTab &
                                                dateformatdmy(myrecord(1)) & vbTab &
                                                validstr(myrecord(2)) & vbTab &
                                                validstr(myrecord(3)) & vbTab &
                                                validstr(myrecord(4)) & vbTab &
                                                validstr(myrecord(5)) & vbTab &
                                                validstr(myrecord(6)) & vbTab &
                                                validstr(myrecord(7)) & vbTab &
                                                validstr(myrecord(8)) & vbTab &
                                                validstr(myrecord(9)) & vbTab &
                                                validstr(myrecord(10)) & vbTab &
                                                validstr(myrecord(11)) & vbTab &
                                                validlong(myrecord(12)) & vbTab &
                                                validstr(myrecord(13)) & vbTab &
                                                validstr(myrecord(14)) & vbTab &
                                                validstr(myrecord(15)) & vbTab &
                                                validstr(myrecord(16)) & vbTab &
                                                validstr(myrecord(17)) & vbTab &
                                                validint(myrecord(18)) & vbTab &
                                                validreal(myrecord(19)) & vbTab &
                                                validreal(myrecord(20)) & vbTab &
                                                validstr(myrecord(21)) & vbTab &
                                                validstr(myrecord(22)) & vbCrLf)

                            Catch ex As Exception                                
                                ProgressReport(1, String.Format("Error, Please check your Excel data. Wrong data found! Line {0}", count + 1))
                                ProgressReport(3, "Set Continuous Again")
                                Exit Sub
                            End Try

                        End If

                    End If
                    count += 1
                Loop
            End With
        End Using

        'update record
        If myInsert.Length > 0 Then

            ProgressReport(1, "Start Add New Records")
            Dim ra As Long = 0
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            Try
                If RadioButton1.Checked Then
                    ProgressReport(1, "Replace Record Please wait!")
                    If PublicClass.myLocation = MyLocation.HongKong Then
                        'mystr.Append("select sales.deletesales(" & DateFormatyyyyMMdd(DateTimePicker1.Value.Date) & ")")
                        mystr.Append("select sales.deletesales(" & DateFormatyyyyMMdd(mydate) & ")")
                    Else
                        'mystr.Append("select sales.deletesalestxtw(" & DateFormatyyyyMMdd(DateTimePicker1.Value.Date) & ")")
                        mystr.Append("select sales.deletesalestxtw(" & DateFormatyyyyMMdd(mydate) & ")")
                    End If

                    ra = DbAdapter1.ExNonQuery(mystr.ToString)
                End If

                If cmmfsb.Length > 0 Then
                    ProgressReport(1, "Add Record Please wait! Master CMMF...")
                    ' cmmf,productid,sbu,prodfamily,cdesc, brand
                    sqlstr = "copy sales.cmmf(cmmf,productid,sbu,prodfamily,cdesc, brand) from stdin with null as 'Null';"
                    errmessage = DbAdapter1.copy(sqlstr, cmmfsb.ToString, myret)
                    If Not myret Then
                        ProgressReport(1, errmessage)
                        ProgressReport(3, "Set Continuous Again")
                        Exit Sub
                    End If
                End If
                
                If sbusb.Length > 0 Then
                    ProgressReport(1, "Add Record Please wait! Master SBU...")
                    ' sbu
                    sqlstr = "copy sales.sbu(sbu) from stdin with null as 'Null';"
                    errmessage = DbAdapter1.copy(sqlstr, sbusb.ToString, myret)
                    If Not myret Then
                        ProgressReport(1, errmessage)
                        ProgressReport(3, "Set Continuous Again")
                        Exit Sub
                    End If
                End If

                If prodfamilysb.Length > 0 Then
                    ProgressReport(1, "Add Record Please wait! Master Prodfamily...")
                    ' sbu
                    sqlstr = "copy sales.prodfamily(prodfamily) from stdin with null as 'Null';"
                    errmessage = DbAdapter1.copy(sqlstr, prodfamilysb.ToString, myret)
                    If Not myret Then
                        ProgressReport(1, errmessage)
                        ProgressReport(3, "Set Continuous Again")
                        Exit Sub
                    End If
                End If

                If brandsb.Length > 0 Then
                    ProgressReport(1, "Add Record Please wait! Master Brand...")
                    ' sbu
                    sqlstr = "copy sales.brand(brand) from stdin with null as 'Null';"
                    errmessage = DbAdapter1.copy(sqlstr, brandsb.ToString, myret)
                    If Not myret Then
                        ProgressReport(1, errmessage)
                        ProgressReport(3, "Set Continuous Again")
                        Exit Sub
                    End If
                End If


                If customersb.Length > 0 Then
                    ProgressReport(1, "Add Record Please wait! Master Customer...")
                    ' sbu
                    sqlstr = "copy sales.customer(customerid,customername) from stdin with null as 'Null';"
                    errmessage = DbAdapter1.copy(sqlstr, customersb.ToString, myret)
                    If Not myret Then
                        ProgressReport(1, errmessage)
                        ProgressReport(3, "Set Continuous Again")
                        Exit Sub
                    End If
                End If
                ProgressReport(1, "Add Record Please wait! Sales Tx...")
                'invid ,invdate,orderno,customerid,customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty,totalsales,totalcost
                If PublicClass.myLocation = MyLocation.HongKong Then
                    sqlstr = "copy sales.tx(invid ,invdate,orderno,customerid,customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty,totalsales,totalcost,region,location) from stdin with null as 'Null';"
                Else
                    sqlstr = "copy sales.txtw(invid ,invdate,orderno,customerid,customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty,totalsales,totalcost,region,location) from stdin with null as 'Null';"
                End If

                errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                If myret Then
                    sw.Stop()
                    ProgressReport(1, "Add Records Done. Elapsed Time: " & Format(sw.Elapsed.Minutes, "00") & ":" & Format(sw.Elapsed.Seconds, "00") & "." & sw.Elapsed.Milliseconds.ToString)

                Else
                    ProgressReport(1, errmessage)

                    Dim errorFilename As String = Path.GetDirectoryName(openfiledialog1.FileName) & "\" & "Error.txt"
                    Using mystream As StreamWriter = File.CreateText(errorFilename)
                        mystream.WriteLine(myInsert.ToString)
                        mystream.Close()
                    End Using
                    Process.Start(errorFilename)
                End If
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        End If
        ProgressReport(3, "Set Continuous Again")
        'delete file
        Kill(mycsvfile)
        sw.Stop()
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
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
            End Select
        End If
 
    End Sub
    Private Function validstr(ByVal data As Object) As Object
        If IsDBNull(data) Then
            Return "Null"
        ElseIf data = "" Then
            Return "Null"
        End If
        data = CType(data, String).Replace("'", "")
        Return data
    End Function

    Private Function filldataset(ByVal DS As DataSet, ByRef myMessage As String) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = "select cmmf from sales.cmmf;" &
                                "select sbu from sales.sbu;" &
                                "select prodfamily from sales.prodfamily;" &
                                "select brand from sales.brand;" &
                                "select customerid from sales.customer"

        ProgressReport(1, "Reading Master Table")

        If Not DbAdapter1.getDataSet(sqlstr, DS, myMessage) Then

            ProgressReport(1, "Done with error. Error Reading Master Table")
            Return myret
        End If
        Try
            DS.Tables(0).TableName = "cmmf"
            DS.Tables(1).TableName = "sbu"
            DS.Tables(2).TableName = "prodfamily"
            DS.Tables(3).TableName = "brand"
            DS.Tables(4).TableName = "customer"

            Dim idx0(0) As DataColumn               'cmmf
            idx0(0) = DS.Tables(0).Columns(0)       'cmmf
            DS.Tables(0).PrimaryKey = idx0

            Dim idx1(0) As DataColumn               'sbu
            idx1(0) = DS.Tables(1).Columns(0)
            DS.Tables(1).PrimaryKey = idx1

            Dim idx2(0) As DataColumn               'prodfamily
            idx2(0) = DS.Tables(2).Columns(0)
            DS.Tables(2).PrimaryKey = idx2

            Dim idx3(0) As DataColumn               'brand
            idx3(0) = DS.Tables(3).Columns(0)
            DS.Tables(3).PrimaryKey = idx3

            Dim idx4(0) As DataColumn               'customer
            idx4(0) = DS.Tables(4).Columns(0)       'customer
            DS.Tables(4).PrimaryKey = idx4
            myret = True
        Catch ex As Exception
            myMessage = ex.Message
        End Try
        Return myret
    End Function

    Private Function convertfile(ByVal filename As String) As String
        Dim myret As String = String.Empty

        'openexcel saveas csv
        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Try
            'Create Object Excel 
            ProgressReport(1, "Preparing Data...")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            oWb = oXl.Workbooks.Open(filename)
            'myret = Replace(filename, "xlsx", "csv")
            myret = Path.GetDirectoryName(filename) & "\" & Path.GetFileNameWithoutExtension(filename) & ".csv"
            oWb.SaveAs(Filename:=myret, FileFormat:=Excel.XlFileFormat.xlCSV, CreateBackup:=False)

        Catch ex As Exception
        Finally
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

        End Try
        Return myret
    End Function

    Private Sub ImportSales_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If mythread.IsAlive Then
            If MessageBox.Show("Are you sure to stop the process?", "Stop the process", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                mythread.Abort()
            Else
                e.Cancel = True
            End If

        End If
    End Sub

    Private Sub doWorktw()
        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim myMessage As String = String.Empty

        If Not filldataset(DS, myMessage) Then
            ProgressReport(1, myMessage)
            Exit Sub
        End If

        Dim sqlstr As String = String.Empty
        Dim mystr As New StringBuilder
        Dim cmmfsb As New StringBuilder
        Dim sbusb As New StringBuilder
        Dim prodfamilysb As New StringBuilder
        Dim brandsb As New StringBuilder
        Dim customersb As New StringBuilder



        sw.Start()
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        Dim mycsvfile = convertfiletw(openfiledialog1.FileName)

        Using objTFParser = New FileIO.TextFieldParser(mycsvfile)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))

                '.SetDelimiters(",")
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                ProgressReport(1, "Read Data")
                ProgressReport(2, "Continuous off")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count >= 1 Then

                        If CDate(myrecord(1)) >= DateTimePicker1.Value.Date Then
                            'po,companycode,createdon,createdby,vendorcode,termsofpayment,purchgroup,currency,docdate,incoterm1,incoterm2
                            'invid character varying,  invdate date,  orderno character varying,  customerid character varying,  customername character varying,  reportcode character varying,  saleforce character varying,  country character varying,  custtype character varying,  salesman character varying,  shipto character varying,  productid character varying,  cmmf bigint,  sbu character varying,  productfamily character varying,  brand character varying,  materialdesc character varying,  supplierid character varying,  qty integer,  totalsales numeric(13,2),  totalcost numeric(15,5),
                            'If myrecord(12) = 2100083280 Then
                            '    Debug.Print("")
                            'End If
                            Dim result As Object
                            If myrecord(13) <> "" Then
                                Dim pkey1(0) As Object
                                pkey1(0) = myrecord(13)
                                result = DS.Tables("sbu").Rows.Find(pkey1)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables("sbu").NewRow
                                    dr.Item(0) = myrecord(13)
                                    DS.Tables("sbu").Rows.Add(dr)
                                    sbusb.Append(validstr(myrecord(13)) & vbCrLf)
                                End If
                            End If

                            If myrecord(14) <> "" Then
                                Dim pkey2(0) As Object
                                pkey2(0) = myrecord(14)
                                result = DS.Tables("prodfamily").Rows.Find(pkey2)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables("prodfamily").NewRow
                                    dr.Item(0) = myrecord(14)
                                    DS.Tables("prodfamily").Rows.Add(dr)
                                    prodfamilysb.Append(validstr(myrecord(14)) & vbCrLf)
                                End If
                            End If

                            If myrecord(15) <> "" Then
                                Dim pkey3(0) As Object
                                pkey3(0) = myrecord(15)
                                result = DS.Tables("brand").Rows.Find(pkey3)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables("brand").NewRow
                                    dr.Item(0) = myrecord(15)
                                    DS.Tables("brand").Rows.Add(dr)
                                    brandsb.Append(validstr(myrecord(15)) & vbCrLf)
                                End If
                            End If


                            If myrecord(3) <> "" Then
                                Dim pkey4(0) As Object
                                pkey4(0) = myrecord(3)
                                result = DS.Tables("customer").Rows.Find(pkey4)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables("customer").NewRow
                                    dr.Item(0) = myrecord(3)
                                    DS.Tables("customer").Rows.Add(dr)
                                    customersb.Append(validstr(myrecord(3)) & vbTab &
                                                      validstr(myrecord(4)) & vbCrLf)
                                End If
                            End If

                            If IsNumeric(myrecord(12)) Then
                                Dim pkey0(0) As Object
                                pkey0(0) = myrecord(12)
                                result = DS.Tables("cmmf").Rows.Find(pkey0)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables("cmmf").NewRow
                                    dr.Item(0) = myrecord(12)
                                    DS.Tables("cmmf").Rows.Add(dr)
                                    'cmmf,productid,sbu,prodfamily,cdesc,brand
                                    cmmfsb.Append(validlong(myrecord(12)) & vbTab &
                                                  validstr(myrecord(11)) & vbTab &
                                                  validstr(myrecord(13)) & vbTab &
                                                  validstr(myrecord(14)) & vbTab &
                                                  validstr(myrecord(16)) & vbTab &
                                                  validstr(myrecord(15)) & vbCrLf)
                                End If
                            Else

                                myrecord(11) = myrecord(12)
                                myrecord(12) = ""
                            End If

                            Try
                                myInsert.Append(validstr(myrecord(0)) & vbTab &
                                                dateformatdmy(myrecord(1)) & vbTab &
                                                validstr(myrecord(2)) & vbTab &
                                                validstr(myrecord(3)) & vbTab &
                                                validstr(myrecord(4)) & vbTab &
                                                validstr(myrecord(5)) & vbTab &
                                                validstr(myrecord(6)) & vbTab &
                                                validstr(myrecord(7)) & vbTab &
                                                validstr(myrecord(8)) & vbTab &
                                                validstr(myrecord(9)) & vbTab &
                                                validstr(myrecord(10)) & vbTab &
                                                validstr(myrecord(11)) & vbTab &
                                                validlong(myrecord(12)) & vbTab &
                                                validstr(myrecord(13)) & vbTab &
                                                validstr(myrecord(14)) & vbTab &
                                                validstr(myrecord(15)) & vbTab &
                                                validstr(myrecord(16)) & vbTab &
                                                validstr(myrecord(17)) & vbTab &
                                                validint(myrecord(18)) & vbTab &
                                                validreal(myrecord(19)) & vbTab &
                                                validreal(myrecord(20)) & vbTab &
                                                validstr(myrecord(22)) & vbTab &
                                                validstr(myrecord(25)) & vbTab &
                                                validstr(myrecord(28)) & vbTab &
                                                validstr(myrecord(29)) & vbTab &
                                                validstr(myrecord(30)) & vbTab &
                                                validstr(myrecord(32)) & vbTab &
                                                validstr(myrecord(33)) & vbCrLf)

                            Catch ex As Exception
                                ProgressReport(1, String.Format("Error, Please check your Excel data. Wrong data found! Line {0}", count + 1))
                                ProgressReport(3, "Set Continuous Again")
                                Exit Sub
                            End Try

                        End If

                    End If
                    count += 1
                Loop
            End With
        End Using

        'update record
        If myInsert.Length > 0 Then

            ProgressReport(1, "Start Add New Records")
            Dim ra As Long = 0
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            Try
                If RadioButton1.Checked Then
                    ProgressReport(1, "Replace Record Please wait!")
                    If PublicClass.myLocation = MyLocation.HongKong Then
                        'mystr.Append("select sales.deletesales(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & ")")
                        mystr.Append("select sales.deletesales(" & DateFormatyyyyMMdd(mydate) & ")")
                    Else
                        'mystr.Append("select sales.deletesalestxtw(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & ")")
                        mystr.Append("select sales.deletesalestxtw(" & DateFormatyyyyMMdd(mydate) & ")")
                    End If

                    ra = DbAdapter1.ExNonQuery(mystr.ToString)
                End If

                If PublicClass.myLocation = MyLocation.HongKong Then
                    If cmmfsb.Length > 0 Then
                        ProgressReport(1, "Add Record Please wait! Master CMMF...")
                        ' cmmf,productid,sbu,prodfamily,cdesc, brand
                        sqlstr = "copy sales.cmmf(cmmf,productid,sbu,prodfamily,cdesc, brand) from stdin with null as 'Null';"
                        errmessage = DbAdapter1.copy(sqlstr, cmmfsb.ToString, myret)
                        If Not myret Then
                            ProgressReport(1, errmessage)
                            ProgressReport(3, "Set Continuous Again")
                            Exit Sub
                        End If
                    End If
                End If
                

                If sbusb.Length > 0 Then
                    ProgressReport(1, "Add Record Please wait! Master SBU...")
                    ' sbu
                    sqlstr = "copy sales.sbu(sbu) from stdin with null as 'Null';"
                    errmessage = DbAdapter1.copy(sqlstr, sbusb.ToString, myret)
                    If Not myret Then
                        ProgressReport(1, errmessage)
                        ProgressReport(3, "Set Continuous Again")
                        Exit Sub
                    End If
                End If

                If prodfamilysb.Length > 0 Then
                    ProgressReport(1, "Add Record Please wait! Master Prodfamily...")
                    ' sbu
                    sqlstr = "copy sales.prodfamily(prodfamily) from stdin with null as 'Null';"
                    errmessage = DbAdapter1.copy(sqlstr, prodfamilysb.ToString, myret)
                    If Not myret Then
                        ProgressReport(1, errmessage)
                        ProgressReport(3, "Set Continuous Again")
                        Exit Sub
                    End If
                End If

                If brandsb.Length > 0 Then
                    ProgressReport(1, "Add Record Please wait! Master Brand...")
                    ' sbu
                    sqlstr = "copy sales.brand(brand) from stdin with null as 'Null';"
                    errmessage = DbAdapter1.copy(sqlstr, brandsb.ToString, myret)
                    If Not myret Then
                        ProgressReport(1, errmessage)
                        ProgressReport(3, "Set Continuous Again")
                        Exit Sub
                    End If
                End If


                If customersb.Length > 0 Then
                    ProgressReport(1, "Add Record Please wait! Master Customer...")
                    ' sbu
                    sqlstr = "copy sales.customer(customerid,customername) from stdin with null as 'Null';"
                    errmessage = DbAdapter1.copy(sqlstr, customersb.ToString, myret)
                    If Not myret Then
                        ProgressReport(1, errmessage)
                        ProgressReport(3, "Set Continuous Again")
                        Exit Sub
                    End If
                End If
                ProgressReport(1, "Add Record Please wait! Sales Tx...")
                'invid ,invdate,orderno,customerid,customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty,totalsales,totalcost
                If PublicClass.myLocation = MyLocation.HongKong Then
                    sqlstr = "copy sales.tx(invid ,invdate,orderno,customerid,customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty,totalsales,totalcost,region,location) from stdin with null as 'Null';"
                Else
                    sqlstr = "copy sales.txtw(invid ,invdate,orderno,customerid,customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,qty,totalsales,totalcost,retur,custname,merch,storename,mlacode,posid,od) from stdin with null as 'Null';"
                End If

                errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                If myret Then
                    sw.Stop()
                    ProgressReport(1, "Add Records Done. Elapsed Time: " & Format(sw.Elapsed.Minutes, "00") & ":" & Format(sw.Elapsed.Seconds, "00") & "." & sw.Elapsed.Milliseconds.ToString)

                Else
                    ProgressReport(1, errmessage)

                    Dim errorFilename As String = Path.GetDirectoryName(openfiledialog1.FileName) & "\" & "Error.txt"
                    Using mystream As StreamWriter = File.CreateText(errorFilename)
                        mystream.WriteLine(myInsert.ToString)
                        mystream.Close()
                    End Using
                    Process.Start(errorFilename)
                End If
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        End If
        ProgressReport(3, "Set Continuous Again")
        'delete file
        Kill(mycsvfile)
        sw.Stop()
    End Sub
    Private Function convertfileHK(ByVal filename As String) As String
        Dim myret As String = String.Empty

        'openexcel saveas csv
        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Try
            'Create Object Excel 
            ProgressReport(1, "Preparing Data...")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            oWb = oXl.Workbooks.Open(filename)
            'myret = Replace(filename, "xlsx", "csv")
            myret = Path.GetDirectoryName(filename) & "\" & Path.GetFileNameWithoutExtension(filename) & ".csv"
            'oWb.SaveAs(Filename:=myret, FileFormat:=Excel.XlFileFormat.xlCSV, CreateBackup:=False)
            oWb.SaveAs(Filename:=myret, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)

        Catch ex As Exception
        Finally
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

        End Try
        Return myret
    End Function
    Private Function convertfiletw(ByVal filename As String) As String
        Dim myret As String = String.Empty

        'openexcel saveas csv
        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Try
            'Create Object Excel 
            ProgressReport(1, "Preparing Data...")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            oWb = oXl.Workbooks.Open(filename)
            'myret = Replace(filename, "xlsx", "csv")
            myret = Path.GetDirectoryName(filename) & "\" & Path.GetFileNameWithoutExtension(filename) & ".txt"
            'oWb.SaveAs(Filename:=myret, FileFormat:=Excel.XlFileFormat.xlCSV, CreateBackup:=False)
            oWb.SaveAs(Filename:=myret, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)

        Catch ex As Exception
        Finally
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

        End Try
        Return myret
    End Function



End Class