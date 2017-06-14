Imports MySql.Data.MySqlClient
Imports Microsoft.Reporting.WinForms

Public Class frmIncomeStatementChart
    Public trial As Boolean
    Dim header As String

    Dim otherIncome As Double
    Dim sales As Double
    Dim cost_of_sales As Double
    Dim general As Double
    Dim pomec As Double
    Dim net_income As Double
    Dim revenue As Double

    Dim dtChart As New DataTable("chart")
    Dim dtPie As New DataTable("pie")

    Dim original As DateTime
    Dim end_month As DateTime
    Dim lastOfMonth As DateTime

    Dim date1 As New Date
    Dim date2 As New Date
    Dim y As New Integer
    Dim month As String
    Dim id As New Integer

    Dim ds2 As New DataSet
    Dim ds3 As New DataSet

    Private Sub frmIncomeStatementChart_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        prd_depot = orig_prd_depot
    End Sub
    Private Sub escape(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then Me.Close()
    End Sub
    Private Sub form_activated(sender As Object, e As EventArgs) Handles Me.Activated
        Me.TopMost = False
    End Sub

    Private Sub frmIncomeStatementChart_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            sql = " SELECT id,period_name as period_name"
            sql = sql + " FROM accnt_period_h WHERE strt_fiscal_yr<='" & Format(Convert.ToDateTime(DATE_MIN), "yyyy-MM-dd") & "' ORDER BY strt_fiscal_yr DESC"

            ds2 = New DataSet
            da = New MySqlDataAdapter(sql, con)
            da.Fill(ds2, "acc_close")

            With ds2.Tables("acc_close")
                If .Rows.Count > 0 Then
                    id = .Rows(0)("id")
                End If
            End With

            For i = 0 To ds2.Tables(0).Rows.Count - 1
                AUTO_ACCOUNT_PERIOD.Add(ds2.Tables(0).Rows(i)("period_name").ToString())
            Next


        Catch ex As Exception
        End Try

        cmbYear.DataSource = ds2.Tables("acc_close")
        cmbYear.DisplayMember = "period_name"
        cmbYear.ValueMember = "id"

        cmbYear.AutoCompleteCustomSource = AUTO_ACCOUNT_PERIOD

        'MODIFIED
        Try
            If Convert.ToInt16(Format(serverNow(), "MM")) = 1 Then
                cm1.SelectedIndex = 11
                cm2.SelectedIndex = 11
                cmbYear.Text = Convert.ToInt32(cmbYear.Text) - 1
            Else
                cmbYear.Text = Convert.ToInt32(Format(serverNow(), "yyyy"))
                cm1.SelectedIndex = Convert.ToInt16(Format(serverNow(), "MM")) - 2
                cm2.SelectedIndex = Convert.ToInt16(Format(serverNow(), "MM")) - 2
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Me.ReportViewer1.RefreshReport()
        ReportViewer1.LocalReport.EnableExternalImages = True


        dtChart.Columns.Add("date") 'Date
        dtChart.Columns.Add("gross_income") 'summary of sales and cost of sales
        dtChart.Columns.Add("gp_otherincome") 'summary of gross income and other income
        dtChart.Columns.Add("gross_expenses") 'summary of expenses
        dtChart.Columns.Add("net_income") 'summary of gp otherincome minus gross expenses

        dtPie.Columns.Add("account_type") 'account type
        dtPie.Columns.Add("amount") 'amount summary

        'call loadPieChart()
        loadPieChart()

        Me.ReportViewer1.RefreshReport()

    End Sub

    Public Sub getScript()

        original = date1 ' The date you want to get the last day of the month for
        end_month = date2 ' The date you want to get the last day of the month for
        lastOfMonth = end_month.Date.AddDays(-(end_month.Day - 1)).AddMonths(1).AddDays(-1)


        sql = " SELECT class.sequence as sequence,  class.caption as class, accnt_type.caption as account_type,"
        sql = sql + " balance_sheet.code, balance_sheet.caption,"
        sql = sql + " IF (trial_debit > 0 , (trial_debit * -1),trial_credit) as amount"
        sql = sql + " FROM ("

        tbs_WHERE = " WHERE (t.post_date >='" & Format(original, "yyyy-MM") & "-01' AND  t.post_date <='" & Format(lastOfMonth, "yyyy-MM-dd") & "' ) "
        getTB()

        sql = sql + tbs

        sql = sql + " )"
        sql = sql + " overall"

        sql = sql + " GROUP BY overall.caption"

        sql = sql + " ) balance_sheet"

        sql = sql + " LEFT JOIN coa"
        sql = sql + " ON balance_sheet.code = coa.code "

        sql = sql + " LEFT JOIN accnt_type"
        sql = sql + " ON coa.accnt_type_id = accnt_type.id "

        sql = sql + " LEFT JOIN class"
        sql = sql + " ON class.id = accnt_type.class_id "

        sql = sql + " WHERE coa.reports='income_statement' "
        sql = sql + " ORDER BY class.sequence,  accnt_type.id"

        prd_depot = orig_prd_depot

    End Sub

    'Loads Line Chart RDLC
    Public Sub loadLineChart()
        dtChart.Clear()
        sql = "SELECT date_format(beg, '%c') as x FROM accnt_period_d WHERE accnt_priod_h_id =" & id
        da = New MySqlDataAdapter(sql, con)
        da.Fill(ds3, "x")

        With ds3.Tables("x")

            For w As Integer = 1 To .Rows.Count
                'System.Diagnostics.Debug.WriteLine("Row Count: " & .Rows.Count)
                If .Rows.Count > 0 Then
                    y = .Rows(w - 1)("x")
                End If
                Try
                    'Fill the date for getScript()
                    date1 = Convert.ToDateTime(y & "-01-" & cmbYear.Text)
                    date2 = Convert.ToDateTime(y & "-01-" & cmbYear.Text)
                Catch ex As Exception
                    MsgBox("Error 102: " & ex.Message)
                End Try

                'Make sure the variables are in 0 values
                otherIncome = 0
                sales = 0
                cost_of_sales = 0
                general = 0
                pomec = 0
                net_income = 0
                revenue = 0

                Try
                    'Call Script
                    getScript()
                    Dim ds As New DataSet
                    ds.Clear()

                    'System.Diagnostics.Debug.WriteLine(sql)

                    da = New MySqlDataAdapter(sql, con)
                    da.Fill(ds, "result")

                    Dim all As Decimal
                    'Get and filter all the necessary values in the datatable
                    With ds.Tables("result")

                        If .Rows.Count > 0 Then
                            For i As Integer = 0 To .Rows.Count - 1

                                Dim str As String = .Rows(i)("class").ToString.ToLower
                                Dim amount As Double = If(.Rows(i)("amount") < 0, (.Rows(i)("amount") * -1), .Rows(i)("amount"))

                                If .Rows(i)("sequence").ToString = "3" Then
                                    otherIncome = otherIncome + .Rows(i)("amount")

                                ElseIf .Rows(i)("sequence").ToString = "4" Then

                                    general = general + (.Rows(i)("amount") * -1)

                                    Dim amount_gen As Double
                                    amount_gen = (.Rows(i)("amount") * -1)

                                    all = all + amount_gen

                                ElseIf .Rows(i)("sequence").ToString = "5" Then

                                    pomec = pomec + .Rows(i)("amount")

                                Else

                                    Dim amount_sales As Double

                                    If .Rows(i)("class").ToString.ToLower.Contains("cost") Then
                                        amount_sales = (.Rows(i)("amount") * -1)
                                    Else
                                        If .Rows(i)("amount") < 0 And Not .Rows(i)("account_type").ToString.ToLower.Contains("disc") Then
                                            amount_sales = (.Rows(i)("amount") * -1)
                                        Else
                                            amount_sales = .Rows(i)("amount")
                                        End If
                                    End If

                                    If .Rows(i)("sequence").ToString = "1" Then
                                        sales = sales + amount_sales
                                    ElseIf .Rows(i)("sequence").ToString = "2" Then
                                        cost_of_sales = Math.Abs(cost_of_sales) + amount_sales
                                    End If

                                End If
                            Next
                        Else
                            ''Set the variables to 0 when there are no acquired data
                            otherIncome = 0
                            sales = 0
                            cost_of_sales = 0
                            general = 0
                            pomec = 0
                            net_income = 0
                            revenue = 0
                        End If

                    End With

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                'System.Diagnostics.Debug.WriteLine("Date: " & Format(date1, "MMMM"))
                'System.Diagnostics.Debug.WriteLine("Gross Income: (Sales - Cost of Sales)" & (Math.Abs(sales) - Math.Abs(cost_of_sales)))
                'System.Diagnostics.Debug.WriteLine("GP after Other Income (Sales - Cost of Sales + Other Income): " & ((Math.Abs(sales) - Math.Abs(cost_of_sales)) + otherIncome))
                'System.Diagnostics.Debug.WriteLine("Expenses (General + Pomec): " & Math.Abs(general) + Math.Abs(pomec))
                'System.Diagnostics.Debug.WriteLine("General: " & Math.Abs(general))
                'System.Diagnostics.Debug.WriteLine("pomec: " & Math.Abs(pomec))
                'System.Diagnostics.Debug.WriteLine("Net Income (GP Other Income - Gross_expenses) :" & (((Math.Abs(sales) - Math.Abs(cost_of_sales)) + Math.Abs(otherIncome)) - Math.Abs(general)) - Math.Abs(pomec))
                'System.Diagnostics.Debug.WriteLine("Revenue: " & revenue)
                'System.Diagnostics.Debug.WriteLine("----------------------------------")

                'Add gathered data to datatable
                dtChart.Rows.Add(Format(date1, "MMMM"),
                       (Math.Abs(sales) - Math.Abs(cost_of_sales)),
                       ((Math.Abs(sales) - Math.Abs(cost_of_sales)) + otherIncome),
                       (Math.Abs(general) + Math.Abs(pomec)),
                       (((Math.Abs(sales) - Math.Abs(cost_of_sales)) + Math.Abs(otherIncome)) - Math.Abs(general)) - Math.Abs(pomec))
            Next

           
        End With

        ReportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local
        ReportViewer1.LocalReport.ReportPath = System.Environment.CurrentDirectory & RDLC_PATH & "/Income Statement Line Graph.rdlc"
        ReportViewer1.LocalReport.DataSources.Clear()

        Try

            Dim test As New ReportParameter("trial", True)
            ReportViewer1.LocalReport.SetParameters(test)
            test = New ReportParameter("title", "Annual Income Statement Chart")
            header = COMPANY_NAME + vbCrLf & "Income Statement Chart"
            ReportViewer1.LocalReport.SetParameters(test)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ReportViewer1.LocalReport.DataSources.Add(New Microsoft.Reporting.WinForms.ReportDataSource("Line", dtChart))

        Try

            Dim test As New ReportParameter("Date1", cmbYear.Text)

            ReportViewer1.LocalReport.SetParameters(test)

            ReportViewer1.LocalReport.SetParameters(test)

            test = New ReportParameter("header", header)
            ReportViewer1.LocalReport.SetParameters(test)

            test = New ReportParameter("imageFile", LOGO_PATH & "/" & COMPANY_NAME & ".png")
            ReportViewer1.LocalReport.SetParameters(test)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ReportViewer1.DocumentMapCollapsed = True
        Me.ReportViewer1.RefreshReport()
        id = 0
    End Sub

    'Loads Pie Chart RDLC
    Public Sub loadPieChart()
        Try
            'Fill the date for getScript()
            date1 = Convert.ToDateTime(cm1.Text & "-01-" & cmbYear.Text)
            date2 = Convert.ToDateTime(cm2.Text & "-01-" & cmbYear.Text)
        Catch ex As Exception
            MsgBox("Error 102: " & ex.Message)
        End Try

        'Make sure the variables are in 0 values
        otherIncome = 0
        sales = 0
        cost_of_sales = 0
        general = 0
        pomec = 0
        net_income = 0

        Try
            'Call Script
            getScript()

            Dim ds As New DataSet
            ds.Clear()
            'Clear datatable
            dtPie.Clear()

            'System.Diagnostics.Debug.WriteLine(sql)

            da = New MySqlDataAdapter(sql, con)
            da.Fill(ds, "result")

            Dim all As Decimal

            'et and filter all the necessary values in the datatable
            With ds.Tables("result")

                If .Rows.Count > 0 Then
                    For i As Integer = 0 To .Rows.Count - 1

                        Dim str As String = .Rows(i)("class").ToString.ToLower
                        Dim amount As Double = If(.Rows(i)("amount") < 0, (.Rows(i)("amount") * -1), .Rows(i)("amount"))

                        If .Rows(i)("sequence").ToString = "3" Then
                            otherIncome = otherIncome + .Rows(i)("amount")


                        ElseIf .Rows(i)("sequence").ToString = "4" Then

                            general = general + (.Rows(i)("amount") * -1)

                            Dim amount_gen As Double
                            amount_gen = (.Rows(i)("amount") * -1)

                            all = all + amount_gen


                        ElseIf .Rows(i)("sequence").ToString = "5" Then

                            pomec = pomec + .Rows(i)("amount")

                        Else

                            Dim amount_sales As Double

                            Console.WriteLine(.Rows(i)("class").ToString)

                            If .Rows(i)("class").ToString.ToLower.Contains("cost") Then
                                amount_sales = (.Rows(i)("amount") * -1)
                            Else
                                If .Rows(i)("amount") < 0 And Not .Rows(i)("account_type").ToString.ToLower.Contains("disc") Then
                                    amount_sales = (.Rows(i)("amount") * -1)
                                Else
                                    amount_sales = .Rows(i)("amount")
                                End If
                            End If

                            If .Rows(i)("sequence").ToString = "1" Then
                                If .Rows(i)("class").ToString = "Revenue" Then
                                    revenue = revenue + amount_sales
                                Else
                                    sales = sales + amount_sales
                                End If

                            ElseIf .Rows(i)("sequence").ToString = "2" Then
                                cost_of_sales = Math.Abs(cost_of_sales) + amount_sales
                            End If

                        End If


                    Next

                    'Checks if the values are greater than 0
                    If sales > 0 Then
                        dtPie.Rows.Add("Sales",
                              sales)
                    End If

                    If revenue > 0 Then
                        dtPie.Rows.Add("Revenue",
                                 revenue)
                    End If

                    If cost_of_sales > 0 Then
                        dtPie.Rows.Add("Cost of Sales",
                                   cost_of_sales)
                    End If
                    If otherIncome > 0 Then
                        dtPie.Rows.Add("Other Income",
                                   otherIncome)
                    End If

                    If Math.Abs(general) - Math.Abs(pomec) > 0 Then
                        dtPie.Rows.Add("Expenses",
                                   Math.Abs(general) + Math.Abs(pomec))
                    End If

                    'System.Diagnostics.Debug.WriteLine("Sales: " & sales)
                    'System.Diagnostics.Debug.WriteLine("Cost of Sales:" & cost_of_sales)
                    'System.Diagnostics.Debug.WriteLine("Other Income:" & otherIncome)
                    'System.Diagnostics.Debug.WriteLine("Expenses: " & Math.Abs(general) - Math.Abs(pomec))

                Else
                    ReportViewer1.LocalReport.ReportPath = System.Environment.CurrentDirectory & RDLC_PATH & "/rptNoResults.rdlc"
                    Me.ReportViewer1.RefreshReport()
                    Exit Sub
                End If
            End With

            ReportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local
            ReportViewer1.LocalReport.ReportPath = System.Environment.CurrentDirectory & RDLC_PATH & "/Income Statement Pie Chart.rdlc"
            ReportViewer1.LocalReport.DataSources.Clear()

            Try
                Dim test As New ReportParameter("trial", True)
                ReportViewer1.LocalReport.SetParameters(test)
                test = New ReportParameter("title", "Current Month Income Statement Chart")
                header = COMPANY_NAME + vbCrLf & "Income Statement Chart"
                ReportViewer1.LocalReport.SetParameters(test)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            ReportViewer1.LocalReport.DataSources.Add(New Microsoft.Reporting.WinForms.ReportDataSource("PieChart", dtPie))

            Try

                Dim test As New ReportParameter("Date1", Format(lastOfMonth, "MMMM dd, " & cmbYear.Text))

                ReportViewer1.LocalReport.SetParameters(test)

                ReportViewer1.LocalReport.SetParameters(test)

                test = New ReportParameter("header", header)
                ReportViewer1.LocalReport.SetParameters(test)

                test = New ReportParameter("imageFile", LOGO_PATH & "/" & COMPANY_NAME & ".png")
                ReportViewer1.LocalReport.SetParameters(test)

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            ReportViewer1.DocumentMapCollapsed = True
            Me.ReportViewer1.RefreshReport()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub cmbYear_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            Dim cmd As New MySqlCommand("SELECT IF(close='1','depot_','') as prd_depot FROM accnt_period_h WHERE period_name='" & cmbYear.Text & "' LIMIT 1; ", con)
            prd_depot = cmd.ExecuteScalar()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnLoadChart_Click(sender As Object, e As EventArgs) Handles btnLoadChart.Click
        Try
            Dim cmd As New MySqlCommand("SELECT IF(close='1','depot_','') as prd_depot FROM accnt_period_h WHERE period_name='" & cmbYear.Text & "' LIMIT 1; ", con)
            prd_depot = cmd.ExecuteScalar()
        Catch ex As Exception
        End Try

        If btnLoadChart.Text = "Annual Chart" Then
            loadLineChart()
            btnLoadChart.Text = "Month Chart"
            cm1.Enabled = False
            cm2.Enabled = False

        Else
            loadPieChart()
            btnLoadChart.Text = "Annual Chart"
            cm1.Enabled = True
            cm2.Enabled = True
        End If

    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        If btnLoadChart.Text = "Annual Chart" Then
            loadPieChart()
        Else
            loadLineChart()
        End If

    End Sub
End Class
