Imports System.Text
Imports Microsoft.Office.Interop
Public Class SGMEventargs
    Inherits EventArgs

    Public Property yearperiod As Integer

    Public Sub New(ByVal yearperiod As Integer)
        Me.yearperiod = yearperiod
    End Sub

End Class
Public Class ReportSGM
    Public yearperiod As Integer
    Public Sub Generate(ByVal myForm As Object, ByVal e As SGMEventargs)
        yearperiod = e.yearperiod
        Dim sqlstr As String = String.Empty

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("SEBAsia_Sales&GM_Report_{0:yyyyMMdd}.xlsx", Date.Today)


        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 4

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
            Dim mystring As New StringBuilder

            mystring.Append(String.Format("(Select  '2. BUD{0}' as label,tx.period,case when tx.type = 1 then 'ACT' else 'BUD' end as type, tx.keyfigures,tx.salesforce,sf.salesforcedesc,tx.mla,m.mladesc,m.distchannel,dc.distchanneldesc ," &
                     " tx.material,c.materialdesc,c.model,md.modeldesc,c.sbu,c.productline,pl.productlinedesc,c.vpm,m.countryiso,ci.countryisodesc," &
                     " c.brand,b.brandname,c.familylv2,lv2.familylv2desc,c.familylv1,lv1.familylv1desc,c.range,r.rangedesc,c.rms,rm.rmsdesc," &
                     " c.source,c.productionsite,ps.productionsitedesc,c.partnership,pp.partnershipdesc,tx.rc,tx.qty,tx.reportingqty,tx.netsales,tx.netsales as bud{0},null::numeric as act{1},null::numeric as act{0},tx.costofsales,tx.grossmargin," &
                     " cf.commercialref," &
                     " substring(keyfigures,1,3) || ' VALUE' as value," &
                     " case brandname when 'SUPOR' then 'SUPOR' else " &
                     " case brandname when 'LAGOSTINA' then 'LAGO' else" &
                     " case brandname when 'LAGOSTINA CASA' then 'LAGO' else" &
                     " 'BASE' " & " End End End as category,cin.series,cin.range,cin.inductionproperty,cin.type,cin.size,cin.extmaterial,cin.intmaterial" &
                     " from sales.sgmtx tx" &
                     " left join sales.sgmsalesforce sf on sf.salesforce = tx.salesforce" &
                     " left join sales.sgmmla m on m.mla = tx.mla" &
                     " left join sales.sgmdistchannel dc on dc.distchannel = m.distchannel" &
                     " left join sales.sgmcmmf c on c.material = tx.material" &
                     " left join sales.sgmmodel md on md.model = c.model" &
                     " left join sales.sgmproductline pl on pl.productline = c.productline" &
                     " left join sales.sgmcountryiso ci on ci.countryiso = m.countryiso" &
                     " left join sales.sgmbrand b on b.brand = c.brand" &
                     " left join sales.sgmfamilylv1 lv1 on lv1.familylv1 = c.familylv1" &
                     " left join sales.sgmfamilylv2 lv2 on lv2.familylv2 = c.familylv2" &
                     " left join sales.sgmrange r on r.range = c.range" &
                     " left join sales.sgmrms rm on rm.rms = c.rms" &
                     " left join sales.sgmproductionsite ps on ps.productionsite = c.productionsite" &
                     " left join sales.sgmpartnership pp on pp.partnership = c.partnership" &
                     " left join cmmf cf on cf.cmmf = c.material" &
                     " left join sales.cmmfinfo cin on cin.cmmf = c.material where tx.type = 2 order by period) ", e.yearperiod, e.yearperiod - 1))
            mystring.Append(String.Format(" union all (Select  '1. ACT{0}' as label,tx.period,case when tx.type = 1 then 'ACT' else 'BUD' end as type, tx.keyfigures,tx.salesforce,sf.salesforcedesc,tx.mla,m.mladesc,m.distchannel,dc.distchanneldesc ," &
                     " tx.material,c.materialdesc,c.model,md.modeldesc,c.sbu,c.productline,pl.productlinedesc,c.vpm,m.countryiso,ci.countryisodesc," &
                     " c.brand,b.brandname,c.familylv2,lv2.familylv2desc,c.familylv1,lv1.familylv1desc,c.range,r.rangedesc,c.rms,rm.rmsdesc," &
                     " c.source,c.productionsite,ps.productionsitedesc,c.partnership,pp.partnershipdesc,tx.rc,tx.qty,tx.reportingqty,tx.netsales,null::numeric,tx.netsales,null::numeric,tx.costofsales,tx.grossmargin," &
                     " cf.commercialref," &
                     " substring(keyfigures,1,3) || ' VALUE' as value," &
                     " case brandname when 'SUPOR' then 'SUPOR' else " &
                     " case brandname when 'LAGOSTINA' then 'LAGO' else" &
                     " case brandname when 'LAGOSTINA CASA' then 'LAGO' else" &
                     " 'BASE' " & " End End End as category,cin.series,cin.range,cin.inductionproperty,cin.type,cin.size,cin.extmaterial,cin.intmaterial" &
                     " from sales.sgmtx tx" &
                     " left join sales.sgmsalesforce sf on sf.salesforce = tx.salesforce" &
                     " left join sales.sgmmla m on m.mla = tx.mla" &
                     " left join sales.sgmdistchannel dc on dc.distchannel = m.distchannel" &
                     " left join sales.sgmcmmf c on c.material = tx.material" &
                     " left join sales.sgmmodel md on md.model = c.model" &
                     " left join sales.sgmproductline pl on pl.productline = c.productline" &
                     " left join sales.sgmcountryiso ci on ci.countryiso = m.countryiso" &
                     " left join sales.sgmbrand b on b.brand = c.brand" &
                     " left join sales.sgmfamilylv1 lv1 on lv1.familylv1 = c.familylv1" &
                     " left join sales.sgmfamilylv2 lv2 on lv2.familylv2 = c.familylv2" &
                     " left join sales.sgmrange r on r.range = c.range" &
                     " left join sales.sgmrms rm on rm.rms = c.rms" &
                     " left join sales.sgmproductionsite ps on ps.productionsite = c.productionsite" &
                     " left join sales.sgmpartnership pp on pp.partnership = c.partnership" &
                     " left join cmmf cf on cf.cmmf = c.material" &
                     " left join sales.cmmfinfo cin on cin.cmmf = c.material where tx.type = 1 and date_part('Year',period) = {0} order by period)", e.yearperiod - 1))
            mystring.Append(String.Format(" union all (Select '3. ACT{0}' as label, tx.period,case when tx.type = 1 then 'ACT' else 'BUD' end as type, tx.keyfigures,tx.salesforce,sf.salesforcedesc,tx.mla,m.mladesc,m.distchannel,dc.distchanneldesc ," &
                     " tx.material,c.materialdesc,c.model,md.modeldesc,c.sbu,c.productline,pl.productlinedesc,c.vpm,m.countryiso,ci.countryisodesc," &
                     " c.brand,b.brandname,c.familylv2,lv2.familylv2desc,c.familylv1,lv1.familylv1desc,c.range,r.rangedesc,c.rms,rm.rmsdesc," &
                     " c.source,c.productionsite,ps.productionsitedesc,c.partnership,pp.partnershipdesc,tx.rc,tx.qty,tx.reportingqty,tx.netsales,null::numeric,null::numeric,tx.netsales,tx.costofsales,tx.grossmargin," &
                     " cf.commercialref," &
                     " substring(keyfigures,1,3) || ' VALUE' as value," &
                     " case brandname when 'SUPOR' then 'SUPOR' else " &
                     " case brandname when 'LAGOSTINA' then 'LAGO' else" &
                     " case brandname when 'LAGOSTINA CASA' then 'LAGO' else" &
                     " 'BASE' " & " End End End as category,cin.series,cin.range,cin.inductionproperty,cin.type,cin.size,cin.extmaterial,cin.intmaterial" &
                     " from sales.sgmtx tx" &
                     " left join sales.sgmsalesforce sf on sf.salesforce = tx.salesforce" &
                     " left join sales.sgmmla m on m.mla = tx.mla" &
                     " left join sales.sgmdistchannel dc on dc.distchannel = m.distchannel" &
                     " left join sales.sgmcmmf c on c.material = tx.material" &
                     " left join sales.sgmmodel md on md.model = c.model" &
                     " left join sales.sgmproductline pl on pl.productline = c.productline" &
                     " left join sales.sgmcountryiso ci on ci.countryiso = m.countryiso" &
                     " left join sales.sgmbrand b on b.brand = c.brand" &
                     " left join sales.sgmfamilylv1 lv1 on lv1.familylv1 = c.familylv1" &
                     " left join sales.sgmfamilylv2 lv2 on lv2.familylv2 = c.familylv2" &
                     " left join sales.sgmrange r on r.range = c.range" &
                     " left join sales.sgmrms rm on rm.rms = c.rms" &
                     " left join sales.sgmproductionsite ps on ps.productionsite = c.productionsite" &
                     " left join sales.sgmpartnership pp on pp.partnership = c.partnership" &
                     " left join cmmf cf on cf.cmmf = c.material" &
                     " left join sales.cmmfinfo cin on cin.cmmf = c.material where tx.type = 1 and date_part('Year',period) = {0} order by period)", e.yearperiod))
            sqlstr = mystring.ToString
            Dim myreport As New ExportToExcelFile(myForm, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")

            myreport.Run(myForm, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As ExportToExcelFileEventArgs)
        Dim osheet As Excel.Worksheet = DirectCast(sender, Excel.Worksheet)
        osheet.Name = "database"
        osheet.Cells.EntireColumn.AutoFit()



    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        Dim owb As Excel.Workbook = DirectCast(sender, Excel.Workbook)

        owb.Worksheets(1).select()
        Dim osheet As Excel.Worksheet = owb.Worksheets(1)

        owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "database!ExternalData_1").CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
            .RowGrand = False
        End With

        osheet.PivotTables("PivotTable1").calculatedfields.add("GM (%)", "=grossmargin/netsales", True)

        osheet.PivotTables("PivotTable1").PivotFields("distchanneldesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("distchanneldesc").Caption = "Distribution Channel Desc"
        osheet.PivotTables("PivotTable1").PivotFields("productlinedesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("productlinedesc").Caption = "Product Line Desc"
        osheet.PivotTables("PivotTable1").PivotFields("familylv1desc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("familylv1desc").Caption = "Family Level 1 Desc"
        osheet.PivotTables("PivotTable1").PivotFields("rangedesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("rangedesc").Caption = "Range Desc"
        osheet.PivotTables("PivotTable1").PivotFields("material").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("material").Caption = "Material"
        osheet.PivotTables("PivotTable1").PivotFields("materialdesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("materialdesc").Caption = "Material Description"
        osheet.PivotTables("PivotTable1").PivotFields("mladesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("mladesc").Caption = "MLA Desc"
        osheet.PivotTables("PivotTable1").PivotFields("brandname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("brandname").Caption = "Consumer Brand Desc"
        osheet.PivotTables("PivotTable1").PivotFields("model").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("model").Caption = "Model"
        osheet.PivotTables("PivotTable1").PivotFields("modeldesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("modeldesc").Caption = "Model desc"
        osheet.PivotTables("PivotTable1").PivotFields("commercialref").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("commercialref").Caption = "Commercial Reference"
        osheet.PivotTables("PivotTable1").Pivotfields("salesforcedesc").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesforcedesc").Caption = "Sales Force Desc"

        osheet.PivotTables("PivotTable1").PivotFields("familylv2desc").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("familylv2desc").Caption = "Family Level 2 Desc"
        osheet.PivotTables("PivotTable1").PivotFields("period").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("B17").Group(Start:=True, [End]:=True, Periods:={False, False, False, False, True, False, True})
        osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlHidden

        osheet.PivotTables("PivotTable1").PivotFields("label").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        With osheet.PivotTables("PivotTable1").PivotFields("period")
            .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
            .position = 2
        End With


        'osheet.PivotTables("PivotTable1").PivotFields("value").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        'osheet.PivotTables("PivotTable1").PivotFields("keyfigures").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        'osheet.PivotTables("PivotTable1").PivotFields("keyfigures").caption = "KEY FIGURES"

        ' osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty"), " Quantity", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("netsales"), " Net Sales (*1000)", Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("costofsales"), " Cost of Sales (*1000)", Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("grossmargin"), " Gross Margin on Standard Costs (*1000)", Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("GM (%)"), " GM (%)", Excel.XlConsolidationFunction.xlSum)

        'osheet.PivotTables("PivotTable1").PivotFields("value").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField

        osheet.PivotTables("PivotTable1").PivotFields(" Net Sales (*1000)").NumberFormat = "#,##0"
        'osheet.PivotTables("PivotTable1").PivotFields(" Net Sales (*1000)").NumberFormat = "#,##0.0"
        'osheet.PivotTables("PivotTable1").PivotFields(" Cost of Sales (*1000)").NumberFormat = "#,##0.0"
        'osheet.PivotTables("PivotTable1").PivotFields(" Gross Margin on Standard Costs (*1000)").NumberFormat = "#,##0.0"
        'osheet.PivotTables("PivotTable1").PivotFields(" GM (%)").NumberFormat = "0.0%"

        osheet.Cells.EntireColumn.AutoFit()
        osheet.Name = "Act Val"

        owb.Worksheets(2).select()

        osheet = owb.Worksheets(2)
        owb.Worksheets(1).PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True            
        End With

        osheet.PivotTables("PivotTable1").calculatedfields.add(String.Format("ACT{0}VS{1}", yearperiod, yearperiod - 1), String.Format("=iferror(act{0}/act{1}-1,0)", yearperiod, yearperiod - 1), True)
        osheet.PivotTables("PivotTable1").calculatedfields.add(String.Format("ACT{0}VSBUD{0}", yearperiod), String.Format("=iferror(act{0}/bud{0}-1,0)", yearperiod), True)

        osheet.PivotTables("PivotTable1").PivotFields("distchanneldesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("distchanneldesc").Caption = "Distribution Channel Desc"
        osheet.PivotTables("PivotTable1").PivotFields("productlinedesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("productlinedesc").Caption = "Product Line Desc"
        osheet.PivotTables("PivotTable1").PivotFields("familylv1desc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("familylv1desc").Caption = "Family Level 1 Desc"
        osheet.PivotTables("PivotTable1").PivotFields("rangedesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("rangedesc").Caption = "Range Desc"
        osheet.PivotTables("PivotTable1").PivotFields("material").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("material").Caption = "Material"
        osheet.PivotTables("PivotTable1").PivotFields("materialdesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("materialdesc").Caption = "Material Description"
        osheet.PivotTables("PivotTable1").PivotFields("mladesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("mladesc").Caption = "MLA Desc"
        osheet.PivotTables("PivotTable1").PivotFields("brandname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("brandname").Caption = "Consumer Brand Desc"
        osheet.PivotTables("PivotTable1").PivotFields("model").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("model").Caption = "Model"
        osheet.PivotTables("PivotTable1").PivotFields("modeldesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("modeldesc").Caption = "Model desc"
        osheet.PivotTables("PivotTable1").PivotFields("commercialref").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("commercialref").Caption = "Commercial Reference"
        osheet.PivotTables("PivotTable1").Pivotfields("salesforcedesc").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesforcedesc").Caption = "Sales Force Desc"

        osheet.PivotTables("PivotTable1").PivotFields("familylv2desc").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("familylv2desc").Caption = "Family Level 2 Desc"
        'osheet.PivotTables("PivotTable1").PivotFields("period").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        ' osheet.Range("B17").Group(Start:=True, [End]:=True, Periods:={False, False, False, False, True, False, True})
        'osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlHidden




        'osheet.PivotTables("PivotTable1").PivotFields("value").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        'osheet.PivotTables("PivotTable1").PivotFields("keyfigures").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        'osheet.PivotTables("PivotTable1").PivotFields("keyfigures").caption = "KEY FIGURES"

        ' osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty"), " Quantity", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields(String.Format("ACT{0}VS{1}", yearperiod, yearperiod - 1)), String.Format(" ACT{0}VS{1}", yearperiod, yearperiod - 1), Excel.XlConsolidationFunction.xlSum)        
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields(String.Format("ACT{0}VSBUD{0}", yearperiod)), String.Format(" ACT{0}VSBUD{0}", yearperiod), Excel.XlConsolidationFunction.xlSum)

      
        'osheet.PivotTables("PivotTable1").PivotFields("value").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField

        osheet.PivotTables("PivotTable1").PivotFields(String.Format(" ACT{0}VS{1}", yearperiod, yearperiod - 1)).NumberFormat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(String.Format(" ACT{0}VSBUD{0}", yearperiod)).NumberFormat = "0.00%"

        'osheet.PivotTables("PivotTable1").PivotFields(" Net Sales (*1000)").NumberFormat = "#,##0.0"
        'osheet.PivotTables("PivotTable1").PivotFields(" Cost of Sales (*1000)").NumberFormat = "#,##0.0"
        'osheet.PivotTables("PivotTable1").PivotFields(" Gross Margin on Standard Costs (*1000)").NumberFormat = "#,##0.0"
        'osheet.PivotTables("PivotTable1").PivotFields(" GM (%)").NumberFormat = "0.0%"
        'osheet.PivotTables("PivotTable1").PivotFields("label").Orientation = Excel.XlPivotFieldOrientation.xlColumnField

        With osheet.PivotTables("PivotTable1").PivotFields("period")
            .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
            .position = 2
        End With
        With osheet.PivotTables("PivotTable1")
            .DisplayErrorString = True
            .RowGrand = True
        End With
        osheet.Cells.EntireColumn.AutoFit()
        osheet.Name = "ACT % Diff"
        owb.Worksheets(3).select()

        osheet = owb.Worksheets(3)
        owb.Worksheets(1).PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").PivotFields("distchanneldesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("distchanneldesc").Caption = "Distribution Channel Desc"
        osheet.PivotTables("PivotTable1").PivotFields("productlinedesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("productlinedesc").Caption = "Product Line Desc"
        osheet.PivotTables("PivotTable1").PivotFields("familylv1desc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("familylv1desc").Caption = "Family Level 1 Desc"
        osheet.PivotTables("PivotTable1").PivotFields("rangedesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("rangedesc").Caption = "Range Desc"
        osheet.PivotTables("PivotTable1").PivotFields("material").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("material").Caption = "Material"
        osheet.PivotTables("PivotTable1").PivotFields("materialdesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("materialdesc").Caption = "Material Description"
        osheet.PivotTables("PivotTable1").PivotFields("mladesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("mladesc").Caption = "MLA Desc"
        osheet.PivotTables("PivotTable1").PivotFields("brandname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("brandname").Caption = "Consumer Brand Desc"
        osheet.PivotTables("PivotTable1").PivotFields("model").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("model").Caption = "Model"
        osheet.PivotTables("PivotTable1").PivotFields("modeldesc").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("modeldesc").Caption = "Model desc"
        osheet.PivotTables("PivotTable1").PivotFields("commercialref").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("commercialref").Caption = "Commercial Reference"
        osheet.PivotTables("PivotTable1").Pivotfields("salesforcedesc").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesforcedesc").Caption = "Sales Force Desc"

        osheet.PivotTables("PivotTable1").PivotFields("familylv2desc").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("familylv2desc").Caption = "Family Level 2 Desc"


        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("GM (%)"), " GM (%)", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" GM (%)").NumberFormat = "0,0%"

        osheet.PivotTables("PivotTable1").PivotFields("label").Orientation = Excel.XlPivotFieldOrientation.xlColumnField        
        With osheet.PivotTables("PivotTable1").PivotFields("period")
            .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
            .position = 2
        End With
        With osheet.PivotTables("PivotTable1")
            .DisplayErrorString = True
            .RowGrand = True
        End With
        osheet.Cells.EntireColumn.AutoFit()
        osheet.Name = "GM%"
    End Sub
End Class
