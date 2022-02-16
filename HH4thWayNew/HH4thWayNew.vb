Imports System.Text
Imports System.Net.Mail
Imports System.Data.SqlClient
Imports System.IO
Module SIPP

    Dim LenderActivated(20000) As Integer

    Public Function fnDBStringField(ByVal sField) As String
        If IsDBNull(sField) Then
            fnDBStringField = " "
        Else
            fnDBStringField = Trim(CStr(sField))
        End If
    End Function
    Public Function fnDBIntField(ByVal sField) As String
        If IsDBNull(sField) Then
            fnDBIntField = 0
        Else
            fnDBIntField = CInt(sField)
        End If
    End Function

    Public Sub SendSimpleMail(sEmail As String, sSubject As String, sBody As String)

        Dim creds As emailcredentials = GetEmailCredentials("NoReply")

        Dim EmailList As List(Of String) = sEmail.Split(",").ToList()
        Dim Externals As Boolean = (From emailAddress In EmailList Where Not emailAddress.Contains("@investandfund.com")).Count() > 0
        Dim ToEmails As String = ""

        Dim nameOfPDF As String
        Dim path = Nothing

        'Dim sep = ;

        nameOfPDF = "4thWay" & DateTime.Now.Day.ToString() & DateTime.Now.Month.ToString() & DateTime.Now.Year.ToString() & ".csv"
        path = System.IO.Path.GetFullPath("C:\IandFApps\HH4thWayReporting\Reports\4thWay" & DateTime.Now.Day.ToString() & DateTime.Now.Month.ToString() & DateTime.Now.Year.ToString() & ".csv")

        Dim bytes As Byte() = System.IO.File.ReadAllBytes(path)

        Dim ms As MemoryStream = New MemoryStream(bytes)


        ToEmails = String.Join(",", EmailList)


        Dim MyMailMessage As New MailMessage() With {
            .From = New MailAddress(creds.User),
            .Subject = sSubject,
            .IsBodyHtml = True,
            .Body = "<table><tr><td>" + sBody + "</table></td></tr>"
        }
        MyMailMessage.To.Add(sEmail)

        MyMailMessage.Attachments.Add(New Attachment(ms, nameOfPDF, "text/plain"))

        Dim SMTPServer As New SmtpClient("smtp.office365.com") With {
            .Credentials = New System.Net.NetworkCredential(creds.User, creds.PW),
            .Port = 587,
            .EnableSsl = True
        }

        Try
            SMTPServer.Send(MyMailMessage)
        Catch ex As Exception
            SendErrorMessage(ex)
        End Try
        SMTPServer = Nothing
        MyMailMessage = Nothing
    End Sub


    Public Function ExecuteIT() As String
        Dim sUsers, MySQL, sHTML, s As String
        Dim iSIPPProvider, iThisUserid, iPrevUserid As Integer
        Dim sHowHear, sReportTitle As String
        Dim dt, dt2, dt3 As New DataTable
        Dim iRun As Integer
        Dim dt1 As New DataSet
        sUsers = Configuration.ConfigurationManager.AppSettings("EmailList")
        iSIPPProvider = Configuration.ConfigurationManager.AppSettings("SIPPProvider")
        sHowHear = Configuration.ConfigurationManager.AppSettings("HowHear").ToLower

        Dim dr As DataRow
        Dim startdate As Date = Now()
        Dim enddate As Date = Now()
        Dim accruedintamount As Integer
        Dim SendEmail As Boolean
        Dim fName As String = ""
        fName = "C:\IandFApps\4thWayReporting\Reports\4thWay" & DateTime.Now.Day.ToString() & DateTime.Now.Month.ToString() & DateTime.Now.Year.ToString() & ".csv"
        Dim csv As String = String.Empty

        SendEmail = True

        If iSIPPProvider > 0 Then
            MySQL = "select u.userid, a.accountid, u.firstname, u.lastname, u.datecreated
                from   USERS u, Accounts A
                where  u.USERID = a.USERID
                and u.USERTYPE = 0
                and a.accounttype = 1
		        and SIPP_Provider = @p1
					and u.ISACTIVE = 0
					and u.USERID not in (10,20,30)
                order by u.userid,  a.ACCOUNTID  "
            iRun = 1
            sReportTitle = "Automated SIPP Reporting"
        Else
            MySQL = "select u.userid, a.accountid, u.firstname, u.lastname, u.datecreated
                from   USERS u, Accounts A
                where  u.USERID = a.USERID
                and u.USERTYPE = 0
                and lower(Trim(HowHear)) = @p1
					and u.ISACTIVE = 0
					and u.USERID not in (10,20,30)
                order by u.userid,  a.ACCOUNTID  "
            iRun = 2
            sReportTitle = "Automated HowHear Reporting - " & sHowHear
        End If


        Try

            Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                Try
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                    Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                    con.Open()
                    cmd.Parameters.Clear()
                    If iRun = 1 Then
                        With cmd.Parameters
                            .Add(New SqlParameter("@p1", iSIPPProvider))
                        End With
                    Else
                        With cmd.Parameters
                            .Add(New SqlParameter("@p1", sHowHear))
                        End With
                    End If
                    adapter.SelectCommand = cmd
                    adapter.Fill(dt)

                    con.Close()
                    con.Dispose()
                Catch ex As Exception
                Finally

                End Try
            End Using

            Dim delimiter As String = "~"
            Dim sb As New StringBuilder()
            'create columnNames:

            'create columnNames headers:

            sb.Append("sep=~")
            sb.Append(Environment.NewLine)

            sb.Append("Name")
            'sb.Append("~AccountID")
            sb.Append("~DateJoined")
            sb.Append("~CashBalance")
            sb.Append("~Loans+PI")
            sb.Append("~CashPlusLoansTotal")
            sb.Append("~AccruedInterest")
            sb.Append("~ValueIncAccruedInterest")
            sb.Append("~~")
            sb.Append("~CashAvaiable")
            sb.Append("~CashSuspense")
            sb.Append("~LoanHoldings")
            sb.Append("~PurchasedInterest")
            sb.Append(Environment.NewLine)


            sHTML = "<html><body><head>
                <style>
                table {
                    font-family: arial, sans-serif;
                    border-collapse: collapse;
                    width: 100%;
                }

                td, th {
                    border: 1px solid #dddddd;
                    text-align: left;
                    padding: 8px;
                }

                tr:nth-child(even) {
                    background-color: #dddddd;
                }
                </style>
                </head>
                <table>
                  <tr>
                    <th style='font-size:30px' colspan=6>" & sReportTitle & "</th>
                  </tr>
                  <tr>
                    <th style='font-size:15px' align='center' colspan=6>Report run date " & startdate.ToString("dd/MM/yyyy") & "</th>
                  </tr>
                  <tr>
                    <th colspan=6></th>
                  </tr>
                  
                  <tr>
                    <th>Name</th>
                    <th>DateJoined</th>  
                    <th>Cash</th>
                    <th>Loans</th>
                    <th>Valuation Total</th>
                    <th>Accrued Interest</th>
                    <th>Value inc. Accrued Interest</th>
                    <th></th>
                    <th></th>
                    <th>Cash Available</th>
                    <th>Cash in Suspense</th>
                    <th>Actual Loan Holdings</th>
                    <th>Purchased Interest</th>
                  </tr>"


            For Each ThisRow As DataRow In dt.Rows
                'Check each account

                dt1 = New DataSet
                iThisUserid = fnDBIntField(ThisRow("userid"))
                If iThisUserid = 4468 Then
                    Dim a = 0

                End If
                If iThisUserid = iPrevUserid Then
                    'already processed figures for this userid
                Else
                    iPrevUserid = iThisUserid

                    '  MySQL = "select top 1 * from ACCRUEDINTAMOUNT
                    'WHERE ACCOUNTID = @I_ACC_ID
                    'order by RUNDATE desc "
                    Select Case iRun
                        Case 1
                            MySQL = "select sum(w.accruedintamount) as accruedinttotal
                        from 
                        (
                        SELECT        t.ACCOUNTID, t.accruedintamount
                        FROM           ( (SELECT        MAX(l.ACCRUEDINTID) AS maxaccruedintid, l.ACCOUNTID
                          FROM            accruedintamount l, accounts a
                          WHERE        l.accountid = a. accountid and a.userid = @I_USER_ID  and a.SIPP_PROVIDER = @p1
                          GROUP BY l.ACCOUNTID)  vt
        			  INNER JOIN
                         accruedintamount  t ON t.accruedintid = vt.maxaccruedintid)  ) w"
                        Case 2
                            MySQL = "select isnull(sum(w.accruedintamount),0) as accruedinttotal
                        from 
                        (
                        SELECT        t.ACCOUNTID, t.accruedintamount
                        FROM           ( (SELECT        MAX(l.ACCRUEDINTID) AS maxaccruedintid, l.ACCOUNTID
                          FROM            accruedintamount l, accounts a, users u
                          WHERE        l.accountid = a. accountid and a.userid = @I_USER_ID 
                             and a.userid = u.userid
                            and lower(Trim(HowHear)) = @p1 
                          GROUP BY l.ACCOUNTID)  vt
        			  INNER JOIN
                         accruedintamount  t ON t.accruedintid = vt.maxaccruedintid)  ) w"
                    End Select


                    Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                        Try
                            Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                            Dim cmd As SqlCommand = New SqlCommand(MySQL, con)
                            con.Open()
                            cmd.Parameters.Clear()
                            Select Case iRun
                                Case 1
                                    With cmd.Parameters
                                        '   .Add(New SqlParameter("@I_ACC_ID", fnDBIntField(ThisRow("accountid"))))
                                        .Add(New SqlParameter("@I_USER_ID", fnDBIntField(ThisRow("userid"))))
                                        .Add(New SqlParameter("@p1", iSIPPProvider))

                                    End With
                                Case 2
                                    With cmd.Parameters
                                        '   .Add(New SqlParameter("@I_ACC_ID", fnDBIntField(ThisRow("accountid"))))
                                        .Add(New SqlParameter("@I_USER_ID", fnDBIntField(ThisRow("userid"))))
                                        .Add(New SqlParameter("@p1", sHowHear))

                                    End With
                            End Select



                            adapter.SelectCommand = cmd
                            adapter.Fill(dt1)
                            con.Close()
                            con.Dispose()
                        Catch ex As Exception
                            s = ex.Message
                        Finally

                        End Try
                    End Using

                    dr = dt1.Tables(0).Rows(0)
                    '   accruedintamount = dr("accruedintamount")
                    accruedintamount = dr("accruedinttotal")

                    Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                        Dim cmd As New SqlCommand()
                        Dim Rdr As SqlDataReader

                        Try

                            cmd.CommandType = CommandType.StoredProcedure
                            cmd.Parameters.Add("@O_GROSS_YIELD", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_DATE_JOINED", SqlDbType.DateTime).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_CURRENT_INVESTMENTS", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_AMOUNT_AVAIL", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_LIVE_AUCTION_BIDS", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_TOTAL_FUNDS_BAL", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_FACILITY_FEES_TOTAL", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_GROSS_INTEREST_EARNED", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_ACE", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_ACCRUED_SOLD", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_ACCRUED_BOUGHT", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_NO_ACTIVE_INVESTMENTS", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_TRANSACTION_FEES_TOTAL", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_PURCHASED_BALANCES", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_LOAN_HOLDINGS", SqlDbType.Int).Direction = ParameterDirection.Output
                            Select Case iRun
                                Case 1
                                    With cmd.Parameters
                                        .Add(New SqlParameter("@I_USER_ID", fnDBIntField(ThisRow("userid"))))
                                        .Add(New SqlParameter("@I_SIPP_PROVIDER", iSIPPProvider))
                                    End With
                                Case 2
                                    With cmd.Parameters
                                        .Add(New SqlParameter("@I_USER_ID", fnDBIntField(ThisRow("userid"))))
                                    End With
                            End Select
                            cmd.Connection = con
                            con.Open()
                            Select Case iRun
                                Case 1
                                    cmd.CommandText = "INVESTOR_SUMMARY_SIPP"
                                Case 2
                                    cmd.CommandText = "INVESTOR_SUMMARY_HOWHEAR"
                            End Select

                            Rdr = cmd.ExecuteReader
                            If Rdr.Read Then

                                Dim ValIncAccruedInterest = Rdr("O_TOTAL_FUNDS_BAL")
                                ValIncAccruedInterest += accruedintamount
                                SendEmail = True

                                'sHTML &= "<td>" & fnDBIntField(ThisRow("userid")) & "</td>"
                                sHTML &= "<tr><td>" & fnDBStringField(ThisRow("firstname")) & " " & fnDBStringField(ThisRow("lastname")) & "</td>"
                                ' sHTML &= "<td>" & fnDBIntField(ThisRow("accountid")) & "</td>"
                                'sHTML &= "<td>" & fnDBStringField(Rdr("O_DATE_JOINED")) & "</td>"
                                sHTML &= "<td>" & fnDBDateField(ThisRow("datecreated")) & "</td>"
                                Dim iCash As Integer = Rdr("O_AMOUNT_AVAIL") +
                                                       Rdr("O_LIVE_AUCTION_BIDS")
                                sHTML &= "<td>" & PenceToCurrencyStringPounds(iCash) & "</td>"
                                sHTML &= "<td>" & PenceToCurrencyStringPounds(Rdr("O_CURRENT_INVESTMENTS")) & "</td>"
                                sHTML &= "<td>" & PenceToCurrencyStringPounds(Rdr("O_TOTAL_FUNDS_BAL")) & "</td>"
                                sHTML &= "<td>" & PenceToCurrencyStringPounds(accruedintamount) & "</td>"
                                sHTML &= "<td>" & PenceToCurrencyStringPounds(ValIncAccruedInterest) & "</td>"
                                sHTML &= "<td></td>"
                                sHTML &= "<td></td>"
                                sHTML &= "<td>" & PenceToCurrencyStringPounds(Rdr("O_AMOUNT_AVAIL")) & "</td>"
                                sHTML &= "<td>" & PenceToCurrencyStringPounds(Rdr("O_LIVE_AUCTION_BIDS")) & "</td>"
                                sHTML &= "<td>" & PenceToCurrencyStringPounds(Rdr("O_LOAN_HOLDINGS")) & "</td>"
                                sHTML &= "<td>" & PenceToCurrencyStringPounds(Rdr("O_PURCHASED_BALANCES")) & "</td></tr>"

                                sHTML &= vbNewLine

                                ' write row to csv
                                sb.Append(fnDBStringField(ThisRow("firstname")) & " " & fnDBStringField(ThisRow("lastname")))
                                ' sb.Append("~" & fnDBIntField(ThisRow("accountid")))
                                sb.Append("~" & fnDBDateField(ThisRow("datecreated")))
                                sb.Append("~" & PenceToCurrencyStringPounds(iCash))
                                sb.Append("~" & PenceToCurrencyStringPounds(Rdr("O_CURRENT_INVESTMENTS")))
                                sb.Append("~" & PenceToCurrencyStringPounds(Rdr("O_TOTAL_FUNDS_BAL")))
                                sb.Append("~" & PenceToCurrencyStringPounds(accruedintamount))
                                sb.Append("~" & PenceToCurrencyStringPounds(ValIncAccruedInterest))
                                sb.Append("~~")
                                sb.Append("~" & PenceToCurrencyStringPounds(Rdr("O_AMOUNT_AVAIL")))
                                sb.Append("~" & PenceToCurrencyStringPounds(Rdr("O_LIVE_AUCTION_BIDS")))
                                sb.Append("~" & PenceToCurrencyStringPounds(Rdr("O_LOAN_HOLDINGS")))
                                sb.Append("~" & PenceToCurrencyStringPounds(Rdr("O_PURCHASED_BALANCES")))
                                sb.Append(Environment.NewLine)
                            End If
                        Catch ex As Exception
                            s = ex.Message
                        Finally
                            ' below not required as using block closes and disposes
                            'If con IsNot Nothing AndAlso con.State = ConnectionState.Open Then
                            '    con.Close()
                            '    con.Dispose()
                            'End If
                        End Try
                    End Using
                    accruedintamount = 0
                End If
            Next

            'Download the CSV file.
            My.Computer.FileSystem.WriteAllText(fName, sb.ToString(), False, Encoding.GetEncoding("Windows-1252"))

            sHTML &= "</table></body></html>"
            If SendEmail = True Then
                SendSimpleMail(sUsers, "Automated 4th Way Reporting", sHTML)
            End If


            ExecuteIT = sHTML

            MySQL = " Insert into REPORTRUN (REPORTNAME,SUCCESS,EMAILSENT,MSGDATA) VALUES (@P1,@P2,@P3,@P4)"


            Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                con.Open()

                Dim command As SqlCommand = con.CreateCommand()
                command.Connection = con

                Try
                    command.CommandText = MySQL
                    With command.Parameters
                        .Add(New SqlParameter("@p1", "4thWayReporting"))
                        .Add(New SqlParameter("@p2", 1))
                        .Add(New SqlParameter("@p3", 1))
                        .Add(New SqlParameter("@p4", ""))

                    End With
                    command.ExecuteNonQuery()
                Catch ex As Exception

                End Try
            End Using

        Catch ex As Exception
            s = ex.Message
            SendErrorMessage(ex)
            ExecuteIT = 0
        End Try
    End Function

    Sub Main()
        ExecuteIT()
    End Sub

    Sub SendErrorMessage(ByVal ThisException As Exception)
        Dim errorPW As String = Configuration.ConfigurationManager.AppSettings("ErrorPW")
        Dim errorUSR As String = Configuration.ConfigurationManager.AppSettings("ErrorUSR")
        Dim mm As New MailMessage() With {
            .From = New MailAddress(errorUSR),
            .Subject = "An Error Has Occurred!",
            .IsBodyHtml = True,
            .Priority = MailPriority.High
        }
        mm.To.Add("web@investandfund.com")

        mm.Body =
            "<html>" & vbCrLf &
            "<body>" & vbCrLf &
            "<h1>An Error Has Occurred!</h1>" & vbCrLf &
            "<table cellpadding=""5"" cellspacing=""0"" border=""1"">" & vbCrLf &
            ItemFormat("Time of Error", DateTime.Now.ToString("dd/MM/yyyy HH:mm:sss"))

        Try
            mm.Body += ItemFormat("Exception Type", ThisException.GetType().ToString())
        Catch ex As Exception
            mm.Body += ItemFormat("Exception Type", "Could not get exception type")
        End Try

        Try
            mm.Body += ItemFormat("Message", ThisException.Message)
        Catch ex As Exception
            mm.Body += ItemFormat("Message", "Could not get message")
        End Try

        Try
            mm.Body += ItemFormat("File Name", "suspicioustransactions.vb")
        Catch ex As Exception
            mm.Body += ItemFormat("File Name", "Could not get file name")
        End Try

        Try
            mm.Body += ItemFormat("Line Number", New StackTrace(ThisException, True).GetFrame(0).GetFileLineNumber)
        Catch ex As Exception
            mm.Body += ItemFormat("Line Number", "Could not get line number")
        End Try

        mm.Body +=
            "</table>" & vbCrLf &
            "</body>" & vbCrLf &
            "</html>"

        Dim smtp As New SmtpClient("smtp.office365.com") With {
            .Credentials = New System.Net.NetworkCredential(errorUSR, errorPW),
            .EnableSsl = True,
            .Port = 587
        }
        smtp.Send(mm)
        Dim MySQL As String

        MySQL = " Insert into REPORTRUN (REPORTNAME,SUCCESS,EMAILSENT,MSGDATA) VALUES (@P1,@P2,@P3,@P4)"


        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            con.Open()

            Dim command As SqlCommand = con.CreateCommand()
            command.Connection = con

            Try
                command.CommandText = MySQL
                With command.Parameters
                    .Add(New SqlParameter("@p1", "CheckFinBals"))
                    .Add(New SqlParameter("@p2", 0))
                    .Add(New SqlParameter("@p3", 0))
                    .Add(New SqlParameter("@p4", ThisException.Message))

                End With
                command.ExecuteNonQuery()
            Catch ex As Exception

            End Try
        End Using

    End Sub

    Public Function ItemFormat(ByVal Title As String, ByVal Message As String) As String
        Return "  <tr>" & vbCrLf &
                "  <tdtext-align: right;font-weight: bold"">" & Title & ":</td>" & vbCrLf &
                "  <td>" & Message & "</td>" & vbCrLf &
                "  </tr>" & vbCrLf
    End Function
    Public Function PenceToCurrencyStringPounds(ByVal sField) As String
        Dim rVal As Double
        Dim iPence As Integer

        If IsDBNull(sField) Then
            rVal = 0.0
        Else
            Try
                iPence = CInt(sField)
                rVal = iPence / 100
            Catch ex As Exception
                rVal = 0.0
            End Try
        End If
        PenceToCurrencyStringPounds = "£" & Format(rVal, "###,###,##0.00")
    End Function
    Public Function fnDBDateField(ByVal sField) As DateTime
        If IsDBNull(sField) Then
            fnDBDateField = DateTime.MinValue
        Else
            fnDBDateField = CDate(sField)
        End If
    End Function

    Public Function GetEmailCredentials(ByVal sEmailName As String) As emailcredentials
        Dim ds As New DataSet

        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try



                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                Dim Cmd As New SqlCommand()
                con.Open()



                Cmd.Parameters.Clear()
                Cmd.Parameters.Add("@O_EMAIL_ADDRESS", SqlDbType.VarChar, 64).Direction = ParameterDirection.Output
                Cmd.Parameters.Add("@O_EMAIL_PASSWORD", SqlDbType.VarChar, 64).Direction = ParameterDirection.Output
                With Cmd.Parameters
                    .Add(New SqlParameter("@i_email_Name", sEmailName))
                End With
                Cmd.Connection = con
                Cmd.CommandText = "GET_EMAIL_PASSWORD"
                Cmd.CommandType = CommandType.StoredProcedure
                Cmd.ExecuteNonQuery()




                GetEmailCredentials = New emailcredentials
                GetEmailCredentials.User = fnDBStringField(Trim(Cmd.Parameters("@O_EMAIL_ADDRESS").Value))
                GetEmailCredentials.PW = Crypt.Decrypt(Trim(Cmd.Parameters("@O_EMAIL_PASSWORD").Value))



            Catch ex As Exception
                GetEmailCredentials = Nothing
            Finally
                con.Close()
                con.Dispose()
            End Try
        End Using
    End Function


End Module
