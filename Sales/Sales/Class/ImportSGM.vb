Imports Microsoft.Office.Interop
Imports Sales.SharedClass
Imports System.Text

Public Class ImportSGM
    Private Enum SGMTypeEnum
        Actual = 1
        Budget = 2
    End Enum

    Protected Friend FileName As String
    Protected Friend myform As Object
    Protected Friend DS As DataSet
    Protected Friend mylist As New List(Of String())
    Protected Friend FirstDate As Date = Date.Today
    Protected Friend EndDate As Date = Date.Today
    Public mytype As Integer
    Public Property ErrorMsg As String


    Protected Friend myAdapter As DbAdapter = DbAdapter.getInstance
    'Dim SGMBrandController As New SGMBrandController
    'Dim OPLTCommentController As New OPLTCOMMENTController


    Public Sub New(ByVal myform As Object, ByVal Filename As String)
        Me.FileName = Filename
        Me.myform = myform
    End Sub
    Public Function ValidateFile() As Boolean

        Dim myret As Boolean = False
        If openExcelFile() Then
            myret = True
        End If
        Return myret
    End Function

    Private Function openExcelFile() As Boolean
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Dim myret As Boolean = False


        Try
            'Create Object Excel 
            myform.ProgressReport(1, "Validate File...")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            myform.ProgressReport(1, String.Format("Open... {0}", FileName))
            'MessageBox.Show(FileName)
            'oWb = oXl.Workbooks.Open(FileName, [ReadOnly]:=False)
            oWb = oXl.Workbooks.Open(FileName)
            myform.ProgressReport(1, String.Format("End Open... {0}", FileName))
            'Check FileType
            For i = 1 To oWb.Worksheets.Count
                If Not oWb.Worksheets(i).visible = False Then
                    oWb.Worksheets(i).select()
                    oSheet = oWb.Worksheets(i)
                    If oSheet.Cells(1, 1).value.Contains("SGMY064") Or oSheet.Cells(1, 1).value.Contains("BUD024") Then
                        'If IO.Path.GetExtension(FileName) = "xlsm" Then
                        oSheet.Columns("AI:AP").NumberFormat = "General"
                        myform.ProgressReport(1, "Savingfile File...")
                        oWb.SaveAs(Filename:=FileName.Replace(IO.Path.GetExtension(FileName), ".TXT"), FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
                        'Else
                        '   oWb.SaveAs(Filename:=FileName.Replace("xlsx", "TXT"), FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
                        'End If
                        myret = True
                        Exit For
                    End If
                End If
            Next
            If Not myret Then
                Throw New Exception("File is not valid.")
            End If
        Catch ex As Exception
            ErrorMsg = ex.Message
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

    Public Function DoImportFile() As Boolean
        DS = New DataSet
        Dim myret As Boolean
        mylist = New List(Of String())

        FirstDate = Date.Today
        EndDate = Date.Today

        Try
            Dim myrecord() As String
            Using objTFParser = New FileIO.TextFieldParser(FileName.Replace(IO.Path.GetExtension(FileName), ".TXT"))
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    myform.ProgressReport(1, "Read Data")
                    Dim mycountcheck As Integer
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count = 0 Then
                            If myrecord(0).Contains("SGMY064") Then
                                mycountcheck = 2
                                mytype = SGMTypeEnum.Actual
                            ElseIf myrecord(0).Contains("BUD024") Then
                                mycountcheck = 2
                                mytype = SGMTypeEnum.Budget
                            End If
                        End If
                        If count >= mycountcheck Then
                            mylist.Add(myrecord)
                        End If
                        count += 1
                    Loop
                End With
            End Using


            Dim sqlstr As String = "select * from sales.sgmtx;" &
                                   " Select * from sales.sgmbrand;" &
                                   " select * from sales.sgmcmmf;" &
                                   " select * from sales.sgmcountryiso;" &
                                   " select * from sales.sgmfamilylv1;" &
                                   " select * from sales.sgmfamilylv2;" &
                                   " select * from sales.sgmmla;" &
                                   " select * from sales.sgmmodel;" &
                                   " select * from sales.sgmpartnership;" &
                                   " select * from sales.sgmproductionsite;" &
                                   " select * from sales.sgmproductline;" &
                                   " select * from sales.sgmrange;" &
                                   " select * from sales.sgmrms;" &
                                   " select * from sales.sgmsalesforce;" &
                                   " select * from sales.sgmdistchannel;"


            If myAdapter.getDataSet(sqlstr, DS, ErrorMsg) Then
                DS.Tables(0).TableName = "SGMTX"
                DS.Tables(1).TableName = "SGMBRAND"
                Dim myIdx1(0) As Data.DataColumn
                myIdx1(0) = DS.Tables(1).Columns("brand")
                DS.Tables(1).PrimaryKey = myIdx1

                DS.Tables(2).TableName = "SGMCMMF"
                Dim myIdx2(0) As Data.DataColumn
                myIdx2(0) = DS.Tables(2).Columns("material")
                DS.Tables(2).PrimaryKey = myIdx2

                DS.Tables(3).TableName = "SGMCOUNTRYISO"
                Dim myIdx3(0) As Data.DataColumn
                myIdx3(0) = DS.Tables(3).Columns("countryiso")
                DS.Tables(3).PrimaryKey = myIdx3

                DS.Tables(4).TableName = "SGMFAMILYLV1"
                Dim myIdx4(0) As Data.DataColumn
                myIdx4(0) = DS.Tables(4).Columns("familylv1")
                DS.Tables(4).PrimaryKey = myIdx4

                DS.Tables(5).TableName = "SGMFAMILYLV2"
                Dim myIdx5(0) As Data.DataColumn
                myIdx5(0) = DS.Tables(5).Columns("familylv2")
                DS.Tables(5).PrimaryKey = myIdx5

                DS.Tables(6).TableName = "SGMMLA"
                Dim myIdx6(0) As Data.DataColumn
                myIdx6(0) = DS.Tables(6).Columns("mla")
                DS.Tables(6).PrimaryKey = myIdx6

                DS.Tables(7).TableName = "SGMMODEL"
                Dim myIdx7(0) As Data.DataColumn
                myIdx7(0) = DS.Tables(7).Columns("model")
                DS.Tables(7).PrimaryKey = myIdx7

                DS.Tables(8).TableName = "SGMPARTNERSHIP"
                Dim myIdx8(0) As Data.DataColumn
                myIdx8(0) = DS.Tables(8).Columns("partnership")
                DS.Tables(8).PrimaryKey = myIdx8

                DS.Tables(9).TableName = "SGMPRODUCTIONSITE"
                Dim myIdx9(0) As Data.DataColumn
                myIdx9(0) = DS.Tables(9).Columns("productionsite")
                DS.Tables(9).PrimaryKey = myIdx9

                DS.Tables(10).TableName = "SGMPRODUCTLINE"
                Dim myIdx10(0) As Data.DataColumn
                myIdx10(0) = DS.Tables(10).Columns("productline")
                DS.Tables(10).PrimaryKey = myIdx10

                DS.Tables(11).TableName = "SGMRANGE"
                Dim myIdx11(0) As Data.DataColumn
                myIdx11(0) = DS.Tables(11).Columns("range")
                DS.Tables(11).PrimaryKey = myIdx11

                DS.Tables(12).TableName = "SGMRMS"
                Dim myIdx12(0) As Data.DataColumn
                myIdx12(0) = DS.Tables(12).Columns("rms")
                DS.Tables(12).PrimaryKey = myIdx12

                DS.Tables(13).TableName = "SGMSalesForce"
                Dim myIdx13(0) As Data.DataColumn
                myIdx13(0) = DS.Tables(13).Columns("salesforce")
                DS.Tables(13).PrimaryKey = myIdx13

                DS.Tables(14).TableName = "SGMDistChannel"
                Dim myIdx14(0) As Data.DataColumn
                myIdx14(0) = DS.Tables(14).Columns("distchannel")
                DS.Tables(14).PrimaryKey = myIdx14

            Else

                Throw New DbAdapterExeption(ErrorMsg)
            End If

            'Prepare Record Budget & Actual
            Dim myImport As Object = Nothing
            Select Case mytype
                Case SGMTypeEnum.Actual
                    myImport = New ActualImport(Me)
                Case SGMTypeEnum.Budget
                    myImport = New BudgetImport(Me)
            End Select
            myImport.execute()
            myret = True
        Catch ex As Exception
            ErrorMsg = ex.Message
        End Try

        Return myret
    End Function

End Class


Public Interface ISGMImport
    Function Execute()
    Property DS As DataSet
    Property mylist As List(Of String())
    Property ImportSGM As ImportSGM

End Interface

Public Class ActualImport
    Implements ISGMImport


    Public Property DS As System.Data.DataSet Implements ISGMImport.DS
    Public Property mylist As System.Collections.Generic.List(Of String()) Implements ISGMImport.mylist
    Public Property ImportSGM As ImportSGM Implements ISGMImport.ImportSGM
    Dim myform As Object
    Dim myadapter As Object
    Dim firstdate As Date
    Dim enddate As Date
    Dim mytype As Object

    Public Sub New(ByVal ImportSGM As ImportSGM)
        Me.ImportSGM = ImportSGM
        Me.DS = ImportSGM.DS
        Me.mylist = ImportSGM.mylist
        Me.myform = ImportSGM.myform
        Me.myadapter = ImportSGM.myAdapter
        Me.firstdate = ImportSGM.FirstDate
        Me.enddate = CDate(String.Format("{0:yyyy}-1-1", ImportSGM.EndDate))
        Me.mytype = ImportSGM.mytype
    End Sub


    Public Function Execute() As Object Implements ISGMImport.Execute
        Dim BrandSB As New StringBuilder
        Dim CMMFSB As New StringBuilder
        Dim CountryISOSB As New StringBuilder
        Dim familylv1SB As New StringBuilder
        Dim familylv2SB As New StringBuilder
        Dim MLASB As New StringBuilder
        Dim modelSB As New StringBuilder
        Dim partnershipSB As New StringBuilder
        Dim productionSiteSB As New StringBuilder
        Dim productlineSB As New StringBuilder
        Dim rangeSB As New StringBuilder
        Dim rmsSB As New StringBuilder
        Dim SalesForceSB As New StringBuilder
        Dim TXSB As New StringBuilder
        Dim DistChannelSB As New StringBuilder
        Dim myret As Boolean
        Dim sqlstr As String


        For i = 0 To mylist.Count - 1




            'check brand
            Dim mykey1(0) As Object
            mykey1(0) = mylist(i)(19)
            Dim result As DataRow = DS.Tables(1).Rows.Find(mykey1)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(1).NewRow
                dr.Item("brand") = mykey1(0)
                dr.Item("brandname") = mylist(i)(20)
                DS.Tables(1).Rows.Add(dr)
                BrandSB.Append(mykey1(0) & vbTab & mylist(i)(20) & vbCrLf)
            End If


            Dim mykey2(0) As Object
            mykey2(0) = mylist(i)(7)
            result = DS.Tables(2).Rows.Find(mykey2)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(2).NewRow
                dr.Item("material") = mykey2(0)
                dr.Item("materialdesc") = mylist(i)(8)
                dr.Item("model") = mylist(i)(9)
                dr.Item("sbu") = mylist(i)(11)
                dr.Item("productline") = mylist(i)(13)
                dr.Item("vpm") = mylist(i)(15)
                dr.Item("brand") = mylist(i)(19)
                dr.Item("familylv1") = mylist(i)(23)
                dr.Item("familylv2") = mylist(i)(21)
                dr.Item("range") = mylist(i)(25)
                dr.Item("rms") = mylist(i)(27)
                dr.Item("source") = mylist(i)(29)
                dr.Item("productionsite") = mylist(i)(30)
                dr.Item("partnership") = mylist(i)(32)
                DS.Tables(2).Rows.Add(dr)
                CMMFSB.Append(mykey2(0) & vbTab &
                              mylist(i)(8) & vbTab &
                              mylist(i)(9) & vbTab &
                              mylist(i)(11) & vbTab &
                              mylist(i)(13) & vbTab &
                              mylist(i)(15) & vbTab &
                              mylist(i)(19) & vbTab &
                              mylist(i)(23) & vbTab &
                              mylist(i)(21) & vbTab &
                              mylist(i)(25) & vbTab &
                              mylist(i)(27) & vbTab &
                              mylist(i)(29) & vbTab &
                              mylist(i)(30) & vbTab &
                              mylist(i)(32) & vbCrLf)
            End If

            Dim mykey3(0) As Object
            mykey3(0) = mylist(i)(17)
            result = DS.Tables(3).Rows.Find(mykey3)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(3).NewRow
                dr.Item("countryiso") = mykey3(0)
                dr.Item("countryisodesc") = mylist(i)(18)
                DS.Tables(3).Rows.Add(dr)
                CountryISOSB.Append(mykey3(0) & vbTab & mylist(i)(18) & vbCrLf)
            End If

            Dim mykey4(0) As Object
            mykey4(0) = mylist(i)(23)
            result = DS.Tables(4).Rows.Find(mykey4)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(4).NewRow
                dr.Item("familylv1") = mykey4(0)
                dr.Item("familylv1desc") = mylist(i)(24)
                DS.Tables(4).Rows.Add(dr)
                familylv1SB.Append(mykey4(0) & vbTab & mylist(i)(24) & vbCrLf)
            End If

            Dim mykey5(0) As Object
            mykey5(0) = mylist(i)(21)
            result = DS.Tables(5).Rows.Find(mykey5)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(5).NewRow
                dr.Item("familylv2") = mykey5(0)
                dr.Item("familylv2desc") = mylist(i)(22)
                DS.Tables(5).Rows.Add(dr)
                familylv2SB.Append(mykey5(0) & vbTab & mylist(i)(22) & vbCrLf)
            End If

            Dim mykey6(0) As Object
            mykey6(0) = mylist(i)(3)
            result = DS.Tables(6).Rows.Find(mykey6)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(6).NewRow
                dr.Item("mla") = mykey6(0)
                dr.Item("mladesc") = mylist(i)(4)
                dr.Item("distchannel") = mylist(i)(6)
                dr.Item("countryiso") = mykey3(0)
                DS.Tables(6).Rows.Add(dr)
                MLASB.Append(mykey6(0) & vbTab &
                            mylist(i)(4) & vbTab &
                            mylist(i)(6) & vbTab &
                            mykey3(0) & vbCrLf)
            End If

            Dim mykey7(0) As Object
            mykey7(0) = mylist(i)(9)
            result = DS.Tables(7).Rows.Find(mykey7)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(7).NewRow
                dr.Item("model") = mykey7(0)
                dr.Item("modeldesc") = mylist(i)(10)
                DS.Tables(7).Rows.Add(dr)
                modelSB.Append(mykey7(0) & vbTab & mylist(i)(10) & vbCrLf)
            End If

            Dim mykey8(0) As Object
            mykey8(0) = mylist(i)(32)
            result = DS.Tables(8).Rows.Find(mykey8)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(8).NewRow
                dr.Item("partnership") = mykey8(0)
                dr.Item("partnershipdesc") = mylist(i)(33)
                DS.Tables(8).Rows.Add(dr)
                partnershipSB.Append(mykey8(0) & vbTab & mylist(i)(33) & vbCrLf)
            End If

            Dim mykey9(0) As Object
            mykey9(0) = mylist(i)(30)
            result = DS.Tables(9).Rows.Find(mykey9)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(9).NewRow
                dr.Item("productionsite") = mykey9(0)
                dr.Item("productionsitedesc") = mylist(i)(31)
                DS.Tables(9).Rows.Add(dr)
                productionSiteSB.Append(mykey9(0) & vbTab & mylist(i)(31) & vbCrLf)
            End If

            Dim mykey10(0) As Object
            mykey10(0) = mylist(i)(13)
            result = DS.Tables(10).Rows.Find(mykey10)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(10).NewRow
                dr.Item("productline") = mykey10(0)
                dr.Item("productlinedesc") = mylist(i)(14)
                DS.Tables(10).Rows.Add(dr)
                productlineSB.Append(mykey10(0) & vbTab & mylist(i)(14) & vbCrLf)
            End If

            Dim mykey11(0) As Object
            mykey11(0) = mylist(i)(25)
            result = DS.Tables(11).Rows.Find(mykey11)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(11).NewRow
                dr.Item("range") = mykey11(0)
                dr.Item("rangedesc") = mylist(i)(26)
                DS.Tables(11).Rows.Add(dr)
                rangeSB.Append(mykey11(0) & vbTab & mylist(i)(26) & vbCrLf)
            End If

            Dim mykey12(0) As Object
            mykey12(0) = mylist(i)(27)
            result = DS.Tables(12).Rows.Find(mykey12)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(12).NewRow
                dr.Item("rms") = mykey12(0)
                dr.Item("rmsdesc") = mylist(i)(28)
                DS.Tables(12).Rows.Add(dr)
                rmsSB.Append(mykey12(0) & vbTab & mylist(i)(28) & vbCrLf)
            End If

            Dim mykey13(0) As Object
            mykey13(0) = mylist(i)(1)
            result = DS.Tables(13).Rows.Find(mykey13)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(13).NewRow
                dr.Item("salesforce") = mykey13(0)
                dr.Item("salesforcedesc") = mylist(i)(2)
                DS.Tables(13).Rows.Add(dr)
                SalesForceSB.Append(mykey13(0) & vbTab & mylist(i)(2) & vbCrLf)
            End If

            Dim myfigures = mylist(i)(0).Split(Chr(10))
            Dim mydate = myfigures(1).Split(".")
            Dim period = String.Format("'{0}-{1:00}-1'", mydate(1), CInt(mydate(0)))
            Dim txdate = CDate(String.Format("{0}-{1}-1", mydate(1), CInt(mydate(0))))

            'If ImportSGM.FirstDate > txdate Then
            '    ImportSGM.FirstDate = txdate
            'End If
            'If ImportSGM.EndDate < txdate Then
            '    ImportSGM.EndDate = txdate
            'End If
            If firstdate > txdate Then
                firstdate = txdate
            End If
            If enddate < txdate Then
                enddate = txdate
            End If
            'keyfigures,salesforce,mla,material,rc,qty,reportingqty,netsales,costofsales,grossmargin,period,type
            TXSB.Append(mylist(i)(0).Replace(vbLf, " ") & vbTab &
                        mylist(i)(1) & vbTab &
                        mylist(i)(3) & vbTab &
                        mylist(i)(7) & vbTab &
                        mylist(i)(34) & vbTab &
                        validint(mylist(i)(35)) & vbTab &
                        validint(mylist(i)(36).Replace("$", "")) & vbTab &
                        validreal(mylist(i)(37).Replace("$", "")) & vbTab &
                        validreal(mylist(i)(38).Replace("$", "")) & vbTab &
                        validreal(mylist(i)(39).Replace("%", "")) & vbTab &
                        validreal(mylist(i)(40).Replace("$", "")) & vbTab &
                        validreal(mylist(i)(41).Replace("$", "")) & vbTab &
                        period & vbTab &
                        ImportSGM.mytype & vbCrLf)

            Dim mykey14(0) As Object
            mykey14(0) = mylist(i)(5)
            result = DS.Tables(14).Rows.Find(mykey14)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(14).NewRow
                dr.Item("distchannel") = mykey14(0)
                dr.Item("distchanneldesc") = mylist(i)(6)
                DS.Tables(14).Rows.Add(dr)
                DistChannelSB.Append(mykey14(0) & vbTab & mylist(i)(6) & vbCrLf)
            End If

        Next

        'Start Copy
        If BrandSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Brand...")
            sqlstr = "copy sales.sgmbrand(brand,brandname) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = ImportSGM.myAdapter.copy(sqlstr, BrandSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If CMMFSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! CMMF...")
            sqlstr = "copy sales.sgmcmmf(material,materialdesc,model,sbu,productline,vpm,brand,familylv1,familylv2,range,rms,source,productionsite,partnership) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, CMMFSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If CountryISOSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! CountryIso...")
            sqlstr = "copy sales.sgmcountryiso(countryiso,countryisodesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, CountryISOSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If


        If familylv1SB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! FamilyLV1...")
            sqlstr = "copy sales.sgmfamilylv1(familylv1,familylv1desc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, familylv1SB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If familylv2SB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! FamilyLV2...")
            sqlstr = "copy sales.sgmfamilylv2(familylv2,familylv2desc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, familylv2SB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If MLASB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! MLA...")
            sqlstr = "copy sales.sgmmla(mla,mladesc,distchannel,countryiso) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, MLASB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If


        If modelSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Model...")
            sqlstr = "copy sales.sgmmodel(model,modeldesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, modelSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If partnershipSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Partnership...")
            sqlstr = "copy sales.sgmpartnership(partnership,partnershipdesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, partnershipSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If


        If productionSiteSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Productionsite...")
            sqlstr = "copy sales.sgmproductionsite(productionsite,productionsitedesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, productionSiteSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If productlineSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! ProductLine...")
            sqlstr = "copy sales.sgmproductline(productline,productlinedesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, productlineSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If


        If rangeSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Range...")
            sqlstr = "copy sales.sgmrange(range,rangedesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, rangeSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If rmsSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! RMS...")
            sqlstr = "copy sales.sgmrms(rms,rmsdesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, rmsSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If SalesForceSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! SalesForce...")
            sqlstr = "copy sales.sgmsalesforce(salesforce,salesforcedesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, SalesForceSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If DistChannelSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Distribution Channel...")
            sqlstr = "copy sales.sgmdistchannel(distchannel,distchanneldesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, DistChannelSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If TXSB.Length > 0 Then
            myform.ProgressReport(1, "Clean Existing Record! TX...")

            Dim mystr As String = String.Empty
            'If mytype = SGMTypeEnum.Actual Then
            'select setval('sales.sgmtx_id_seq',1,false);
            mystr = String.Format("delete from sales.sgmtx where period >= '{0:yyyy-MM-dd}' and period <= '{1:yyyy-MM-dd}' and type = {2};", FirstDate, EndDate, mytype)
            'eIf mytype = SGMTypeEnum.Budget Then
            '  mystr = String.Format("delete from sales.sgmtx where type = {0};", mytype)
            'End If

            Dim ra = myadapter.ExNonQuery(mystr.ToString)
            myform.ProgressReport(1, "Add Record Please wait! TX...")
            sqlstr = "copy sales.sgmtx(keyfigures,salesforce,mla,material,rc,qty,reportingqty,netsales,costofsales,grossmargin,gmstandardcost,nsp,period,type) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, TXSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If
        myret = True
        Return myret
    End Function


    
End Class

Public Class BudgetImport
    Implements ISGMImport

    Public Property ImportSGM As ImportSGM Implements ISGMImport.ImportSGM

    Public Property DS As System.Data.DataSet Implements ISGMImport.DS
    Public Property mylist As System.Collections.Generic.List(Of String()) Implements ISGMImport.mylist
    Dim myform As Object
    Dim myadapter As Object
    Dim mytype As Object

    Public Sub New(ByVal importsgm As ImportSGM)
        Me.ImportSGM = ImportSGM
        Me.DS = ImportSGM.DS
        Me.mylist = ImportSGM.mylist
        Me.myform = ImportSGM.myform
        Me.myadapter = importsgm.myAdapter
        Me.mytype = importsgm.mytype
    End Sub

    Public Function Execute() As Object Implements ISGMImport.Execute
        Dim BrandSB As New StringBuilder
        Dim CMMFSB As New StringBuilder
        Dim CountryISOSB As New StringBuilder
        Dim familylv1SB As New StringBuilder
        Dim familylv2SB As New StringBuilder
        Dim MLASB As New StringBuilder
        Dim modelSB As New StringBuilder
        Dim partnershipSB As New StringBuilder
        Dim productionSiteSB As New StringBuilder
        Dim productlineSB As New StringBuilder
        Dim rangeSB As New StringBuilder
        Dim rmsSB As New StringBuilder
        Dim SalesForceSB As New StringBuilder
        Dim TXSB As New StringBuilder
        Dim DistChannelSB As New StringBuilder
        Dim myret As Boolean
        Dim sqlstr As String


        For i = 0 To mylist.Count - 1




            'check brand
            Dim mykey1(0) As Object
            mykey1(0) = mylist(i)(19)
            Dim result As DataRow = DS.Tables(1).Rows.Find(mykey1)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(1).NewRow
                dr.Item("brand") = mykey1(0)
                dr.Item("brandname") = mylist(i)(20)
                DS.Tables(1).Rows.Add(dr)
                BrandSB.Append(mykey1(0) & vbTab & mylist(i)(20) & vbCrLf)
            End If


            Dim mykey2(0) As Object
            mykey2(0) = mylist(i)(7)
            result = DS.Tables(2).Rows.Find(mykey2)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(2).NewRow
                dr.Item("material") = mykey2(0)
                dr.Item("materialdesc") = mylist(i)(8)
                dr.Item("model") = mylist(i)(9)
                dr.Item("sbu") = mylist(i)(11)
                dr.Item("productline") = mylist(i)(13)
                dr.Item("vpm") = mylist(i)(15)
                dr.Item("brand") = mylist(i)(19)
                dr.Item("familylv1") = mylist(i)(23)
                dr.Item("familylv2") = mylist(i)(21)
                dr.Item("range") = mylist(i)(25)
                dr.Item("rms") = mylist(i)(27)
                dr.Item("source") = mylist(i)(29)
                dr.Item("productionsite") = mylist(i)(30)
                dr.Item("partnership") = mylist(i)(32)
                DS.Tables(2).Rows.Add(dr)
                CMMFSB.Append(mykey2(0) & vbTab &
                              mylist(i)(8) & vbTab &
                              mylist(i)(9) & vbTab &
                              mylist(i)(11) & vbTab &
                              mylist(i)(13) & vbTab &
                              mylist(i)(15) & vbTab &
                              mylist(i)(19) & vbTab &
                              mylist(i)(23) & vbTab &
                              mylist(i)(21) & vbTab &
                              mylist(i)(25) & vbTab &
                              mylist(i)(27) & vbTab &
                              mylist(i)(29) & vbTab &
                              mylist(i)(30) & vbTab &
                              mylist(i)(32) & vbCrLf)
            End If

            Dim mykey3(0) As Object
            mykey3(0) = mylist(i)(17)
            result = DS.Tables(3).Rows.Find(mykey3)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(3).NewRow
                dr.Item("countryiso") = mykey3(0)
                dr.Item("countryisodesc") = mylist(i)(18)
                DS.Tables(3).Rows.Add(dr)
                CountryISOSB.Append(mykey3(0) & vbTab & mylist(i)(18) & vbCrLf)
            End If

            Dim mykey4(0) As Object
            mykey4(0) = mylist(i)(23)
            result = DS.Tables(4).Rows.Find(mykey4)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(4).NewRow
                dr.Item("familylv1") = mykey4(0)
                dr.Item("familylv1desc") = mylist(i)(24)
                DS.Tables(4).Rows.Add(dr)
                familylv1SB.Append(mykey4(0) & vbTab & mylist(i)(24) & vbCrLf)
            End If

            Dim mykey5(0) As Object
            mykey5(0) = mylist(i)(21)
            result = DS.Tables(5).Rows.Find(mykey5)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(5).NewRow
                dr.Item("familylv2") = mykey5(0)
                dr.Item("familylv2desc") = mylist(i)(22)
                DS.Tables(5).Rows.Add(dr)
                familylv2SB.Append(mykey5(0) & vbTab & mylist(i)(22) & vbCrLf)
            End If

            Dim mykey6(0) As Object
            mykey6(0) = mylist(i)(3)
            result = DS.Tables(6).Rows.Find(mykey6)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(6).NewRow
                dr.Item("mla") = mykey6(0)
                dr.Item("mladesc") = mylist(i)(4)
                dr.Item("distchannel") = mylist(i)(6)
                dr.Item("countryiso") = mykey3(0)
                DS.Tables(6).Rows.Add(dr)
                MLASB.Append(mykey6(0) & vbTab &
                            mylist(i)(4) & vbTab &
                            mylist(i)(6) & vbTab &
                            mykey3(0) & vbCrLf)
            End If

            Dim mykey7(0) As Object
            mykey7(0) = mylist(i)(9)
            result = DS.Tables(7).Rows.Find(mykey7)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(7).NewRow
                dr.Item("model") = mykey7(0)
                dr.Item("modeldesc") = mylist(i)(10)
                DS.Tables(7).Rows.Add(dr)
                modelSB.Append(mykey7(0) & vbTab & mylist(i)(10) & vbCrLf)
            End If

            Dim mykey8(0) As Object
            mykey8(0) = mylist(i)(32)
            result = DS.Tables(8).Rows.Find(mykey8)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(8).NewRow
                dr.Item("partnership") = mykey8(0)
                dr.Item("partnershipdesc") = mylist(i)(33)
                DS.Tables(8).Rows.Add(dr)
                partnershipSB.Append(mykey8(0) & vbTab & mylist(i)(33) & vbCrLf)
            End If

            Dim mykey9(0) As Object
            mykey9(0) = mylist(i)(30)
            result = DS.Tables(9).Rows.Find(mykey9)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(9).NewRow
                dr.Item("productionsite") = mykey9(0)
                dr.Item("productionsitedesc") = mylist(i)(31)
                DS.Tables(9).Rows.Add(dr)
                productionSiteSB.Append(mykey9(0) & vbTab & mylist(i)(31) & vbCrLf)
            End If

            Dim mykey10(0) As Object
            mykey10(0) = mylist(i)(13)
            result = DS.Tables(10).Rows.Find(mykey10)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(10).NewRow
                dr.Item("productline") = mykey10(0)
                dr.Item("productlinedesc") = mylist(i)(14)
                DS.Tables(10).Rows.Add(dr)
                productlineSB.Append(mykey10(0) & vbTab & mylist(i)(14) & vbCrLf)
            End If

            Dim mykey11(0) As Object
            mykey11(0) = mylist(i)(25)
            result = DS.Tables(11).Rows.Find(mykey11)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(11).NewRow
                dr.Item("range") = mykey11(0)
                dr.Item("rangedesc") = mylist(i)(26)
                DS.Tables(11).Rows.Add(dr)
                rangeSB.Append(mykey11(0) & vbTab & mylist(i)(26) & vbCrLf)
            End If

            Dim mykey12(0) As Object
            mykey12(0) = mylist(i)(27)
            result = DS.Tables(12).Rows.Find(mykey12)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(12).NewRow
                dr.Item("rms") = mykey12(0)
                dr.Item("rmsdesc") = mylist(i)(28)
                DS.Tables(12).Rows.Add(dr)
                rmsSB.Append(mykey12(0) & vbTab & mylist(i)(28) & vbCrLf)
            End If

            Dim mykey13(0) As Object
            mykey13(0) = mylist(i)(1)
            result = DS.Tables(13).Rows.Find(mykey13)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(13).NewRow
                dr.Item("salesforce") = mykey13(0)
                dr.Item("salesforcedesc") = mylist(i)(2)
                DS.Tables(13).Rows.Add(dr)
                SalesForceSB.Append(mykey13(0) & vbTab & mylist(i)(2) & vbCrLf)
            End If

            Dim myfigures = mylist(i)(0).Split(Chr(10))
            Dim mydate = myfigures(1).Split(".")
            Dim period = String.Format("'{0}-{1:00}-1'", mydate(1), CInt(mydate(0)))
            Dim txdate = CDate(String.Format("{0}-{1}-1", mydate(1), CInt(mydate(0))))

            If ImportSGM.FirstDate > txdate Then
                ImportSGM.FirstDate = txdate
            End If
            If ImportSGM.EndDate < txdate Then
                ImportSGM.EndDate = txdate
            End If
            'keyfigures,salesforce,mla,material,rc,qty,reportingqty,netsales,costofsales,grossmargin,period,type
            TXSB.Append(mylist(i)(0).Replace(vbLf, " ") & vbTab &
                        mylist(i)(1) & vbTab &
                        mylist(i)(3) & vbTab &
                        mylist(i)(7) & vbTab &                        
                        validint(mylist(i)(34)) & vbTab &
                        validint(mylist(i)(35).Replace("$", "")) & vbTab &
                        validreal(mylist(i)(36).Replace("$", "")) & vbTab &
                        validreal(mylist(i)(37).Replace("$", "")) & vbTab &
                        validreal(mylist(i)(38).Replace("$", "")) & vbTab &
                        validreal(mylist(i)(39).Replace("%", "")) & vbTab &
                        validreal(mylist(i)(40).Replace("$", "")) & vbTab &
                        validreal(mylist(i)(41).Replace("$", "")) & vbTab &
                        period & vbTab &
                        ImportSGM.mytype & vbCrLf)

            Dim mykey14(0) As Object
            mykey14(0) = mylist(i)(5)
            result = DS.Tables(14).Rows.Find(mykey14)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(14).NewRow
                dr.Item("distchannel") = mykey14(0)
                dr.Item("distchanneldesc") = mylist(i)(6)
                DS.Tables(14).Rows.Add(dr)
                DistChannelSB.Append(mykey14(0) & vbTab & mylist(i)(6) & vbCrLf)
            End If

        Next

        'Start Copy
        If BrandSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Brand...")
            sqlstr = "copy sales.sgmbrand(brand,brandname) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = ImportSGM.myAdapter.copy(sqlstr, BrandSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If CMMFSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! CMMF...")
            sqlstr = "copy sales.sgmcmmf(material,materialdesc,model,sbu,productline,vpm,brand,familylv1,familylv2,range,rms,source,productionsite,partnership) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, CMMFSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If CountryISOSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! CountryIso...")
            sqlstr = "copy sales.sgmcountryiso(countryiso,countryisodesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, CountryISOSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If


        If familylv1SB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! FamilyLV1...")
            sqlstr = "copy sales.sgmfamilylv1(familylv1,familylv1desc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, familylv1SB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If familylv2SB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! FamilyLV2...")
            sqlstr = "copy sales.sgmfamilylv2(familylv2,familylv2desc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, familylv2SB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If MLASB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! MLA...")
            sqlstr = "copy sales.sgmmla(mla,mladesc,distchannel,countryiso) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, MLASB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If


        If modelSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Model...")
            sqlstr = "copy sales.sgmmodel(model,modeldesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, modelSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If partnershipSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Partnership...")
            sqlstr = "copy sales.sgmpartnership(partnership,partnershipdesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, partnershipSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If


        If productionSiteSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Productionsite...")
            sqlstr = "copy sales.sgmproductionsite(productionsite,productionsitedesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, productionSiteSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If productlineSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! ProductLine...")
            sqlstr = "copy sales.sgmproductline(productline,productlinedesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, productlineSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If


        If rangeSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Range...")
            sqlstr = "copy sales.sgmrange(range,rangedesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, rangeSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If rmsSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! RMS...")
            sqlstr = "copy sales.sgmrms(rms,rmsdesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, rmsSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If SalesForceSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! SalesForce...")
            sqlstr = "copy sales.sgmsalesforce(salesforce,salesforcedesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, SalesForceSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If DistChannelSB.Length > 0 Then
            myform.ProgressReport(1, "Add Record Please wait! Distribution Channel...")
            sqlstr = "copy sales.sgmdistchannel(distchannel,distchanneldesc) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, DistChannelSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If

        If TXSB.Length > 0 Then
            myform.ProgressReport(1, "Clean Existing Record! TX...")

            Dim mystr As String = String.Empty
            'If mytype = SGMTypeEnum.Actual Then
            'select setval('sales.sgmtx_id_seq',1,false);
            'mystr = String.Format("delete from sales.sgmtx where period >= '{0:yyyy-MM-dd}' and period <= '{1:yyyy-MM-dd}' and type = {2};", FirstDate, EndDate, mytype)
            'eIf mytype = SGMTypeEnum.Budget Then
            mystr = String.Format("delete from sales.sgmtx where type = {0};", mytype)
            'End If

            Dim ra = myadapter.ExNonQuery(mystr.ToString)
            myform.ProgressReport(1, "Add Record Please wait! TX...")
            sqlstr = "copy sales.sgmtx(keyfigures,salesforce,mla,material,qty,reportingqty,netsales,costofsales,grossmargin,gmstandardcost,nsp,nsp2,period,type) from stdin with null as 'Null';"
            ImportSGM.ErrorMsg = myadapter.copy(sqlstr, TXSB.ToString, myret)
            If Not myret Then
                myform.ProgressReport(1, ImportSGM.ErrorMsg)
                myform.ProgressReport(3, "Set Continuous Again")
                Return myret
            End If
        End If
        myret = True
        Return myret
    End Function



       
End Class