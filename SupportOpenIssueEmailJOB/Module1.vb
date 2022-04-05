Imports System.Configuration
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports log4net
Imports log4net.Config
Imports Microsoft.Office.Interop
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Module Module1
    Private ReadOnly log As ILog = LogManager.GetLogger(GetType(Module1))
    Sub Main()
        XmlConfigurator.Configure()
        If Date.Today.ToString("ddd").ToString().ToLower = ConfigurationManager.AppSettings("EmailSendDay").ToString().ToLower Then
            log.Info(DateTime.Now.ToString() & " - Its Friday.")
            GetSupportOpenDataAndSendEmail()
        Else
            log.Info(DateTime.Now.ToString() & " - Not Friday.")
        End If
    End Sub

    Private Sub GetSupportOpenDataAndSendEmail()
        Try
            log.Info("In GetSupportOpenDataAndSendEmail")
            Dim supportBal As SupportBAL = New SupportBAL()
            Dim dtData As DataTable = New DataTable()
            Dim condition As String = ""
            condition = " and SI.Status = 1"
            dtData = supportBal.GetSupportOpenDataAndSendEmail(condition)

            '.SupportItemId = p.Field(Of Integer)("SupportItemId"),
            '.CompanyId = p.Field(Of Integer)("CompanyId"),
            '.HubId = p.Field(Of Integer)("HubId"),
            '.SiteId = p.Field(Of Integer)("SiteId"),
            '.IsSendMail = p.Field(Of Integer)("IsSendMail"),
            '.AssignTicketTo = p.Field(Of Integer)("AssignTicketTo"),
            '.TicketOwnerPersonId = p.Field(Of Integer)("TicketOwnerPersonId"),
            '.TankId = p.Field(Of Integer)("TankId"),
            '.IsMassTickets = p.Field(Of Integer)("IsMassTickets"),
            '.VRIssueOrderReplacementParts = p.Field(Of Integer)("VRIssueOrderReplacementParts")
            '.Status = p.Field(Of String)("Status"),

            If (dtData IsNot Nothing) Then
                If (dtData.Rows.Count > 0) Then
                    Dim items = (From p In dtData.AsEnumerable()
                                 Select New With {
                                                  .SupportDateTime = p.Field(Of DateTime?)("SupportDateTime"),
                                                  .IssueTypeText = p.Field(Of String)("IssueTypeText"),
                                                  .IssueType = p.Field(Of String)("IssueType"),
                                                  .IssueDescription = p.Field(Of String)("IssueDescription"),
                                                  .Company = p.Field(Of String)("Company"),
                                                  .attachment = p.Field(Of String)("attachment"),
                                                  .HubSiteName = p.Field(Of String)("HubSiteName"),
                                                  .FsLink = p.Field(Of String)("FsLink"),
                                                  .CreatedBy = p.Field(Of String)("CreatedBy"),
                                                  .HUBName = p.Field(Of String)("HUBName"),
                                                  .LINKName = p.Field(Of String)("LINKName"),
                                                  .CreatedDate = p.Field(Of String)("CreatedDate"),
                                                  .CreatedTime = p.Field(Of String)("CreatedTime"),
                                                  .ResolutionDate = p.Field(Of String)("ResolutionDate"),
                                                  .ResolutionTime = p.Field(Of String)("ResolutionTime"),
                                                  .StatusText = p.Field(Of String)("StatusText"),
                                                  .Resolution = p.Field(Of String)("Resolution"),
                                                  .Contact = p.Field(Of String)("Contact"),
                                                  .selectedAdmins = p.Field(Of String)("selectedAdmins"),
                                                  .ResolutionDateTime = p.Field(Of DateTime?)("ResolutionDateTime"),
                                                  .ShortDescription = p.Field(Of String)("ShortDescription"),
                                                  .CaseOpenedBy = p.Field(Of String)("CaseOpenedBy"),
                                                  .CaseClosedBy = p.Field(Of String)("CaseClosedBy"),
                                                  .TicketOwner = p.Field(Of String)("TicketOwner"),
                                                  .TicketOwnerEmail = p.Field(Of String)("TicketOwnerEmail"),
                                                  .ReportedIssue = p.Field(Of String)("ReportedIssue"),
                                                  .Observations = p.Field(Of String)("Observations"),
                                                  .ActionsPerformed = p.Field(Of String)("ActionsPerformed"),
                                                  .Results = p.Field(Of String)("Results"),
                                                  .AdditionalComments = p.Field(Of String)("AdditionalComments"),
                                                  .MiscPartsOrder = p.Field(Of String)("MiscPartsOrder")}).ToList()

                    Dim emailSendTo As String() = ConfigurationManager.AppSettings("EmailSendTO").Split(",")

                    For index = 0 To emailSendTo.Count - 1
                        Try
                            Dim email As String = emailSendTo(index).Trim()
                            Dim data = items.Where(Function(x) x.TicketOwnerEmail?.ToString().Trim() = email).[Select](Function(fetch) New With {
                                                      .SupportDateTime = fetch.SupportDateTime,
                                                      .IssueTypeText = fetch.IssueTypeText,
                                                      .IssueType = fetch.IssueType,
                                                      .IssueDescription = fetch.IssueDescription,
                                                      .Company = fetch.Company,
                                                      .attachment = fetch.attachment,
                                                      .HubSiteName = fetch.HubSiteName,
                                                      .FsLink = fetch.FsLink,
                                                      .CreatedBy = fetch.CreatedBy,
                                                      .HUBName = fetch.HUBName,
                                                      .LINKName = fetch.LINKName,
                                                      .CreatedDate = fetch.CreatedDate,
                                                      .CreatedTime = fetch.CreatedTime,
                                                      .ResolutionDate = fetch.ResolutionDate,
                                                      .ResolutionTime = fetch.ResolutionTime,
                                                      .StatusText = fetch.StatusText,
                                                      .Resolution = fetch.Resolution,
                                                      .Contact = fetch.Contact,
                                                      .selectedAdmins = fetch.selectedAdmins,
                                                      .ResolutionDateTime = fetch.ResolutionDateTime,
                                                      .ShortDescription = fetch.ShortDescription,
                                                      .CaseOpenedBy = fetch.CaseOpenedBy,
                                                      .CaseClosedBy = fetch.CaseClosedBy,
                                                      .TicketOwner = fetch.TicketOwner,
                                                      .TicketOwnerEmail = fetch.TicketOwnerEmail,
                                                      .ReportedIssue = fetch.ReportedIssue,
                                                      .Observations = fetch.Observations,
                                                      .ActionsPerformed = fetch.ActionsPerformed,
                                                      .Results = fetch.Results,
                                                      .AdditionalComments = fetch.AdditionalComments,
                                                      .MiscPartsOrder = fetch.MiscPartsOrder
                                    }).ToList()

                            If data.Count > 0 Then
                                CreateReportAndSendEmail(data, emailSendTo(index).Trim())
                            End If
                        Catch ex As Exception
                            log.Error("Exception occurred in GetDataAndSendUnResolvedEmail. ex is :" & ex.ToString())
                        End Try
                    Next



                End If
            End If
            log.Info("Execution Ended")
        Catch ex As Exception
            log.Error("Exception occurred in GetDataAndSendUnResolvedEmail. ex is :" & ex.ToString())
        End Try

    End Sub

    Private Sub CreateReportAndSendEmail(ByVal data As Object, OwnerEmail As String)
        Try
            log.Info("In CreateReportAndSendEmail " & OwnerEmail)
            'Dim xlApp As Excel.Application = New Excel.Application()
            'If xlApp Is Nothing Then
            '    log.Info("Excel is not properly installed!!")
            '    Return
            'End If
            'Dim xlWorkBook As Excel.Workbook
            'Dim xlWorkSheet As Excel.Worksheet
            'Dim misValue As Object = System.Reflection.Missing.Value
            'xlWorkBook = xlApp.Workbooks.Add(misValue)
            'xlWorkSheet = xlWorkBook.Sheets("sheet1")

            Dim workbook As IWorkbook
            workbook = New HSSFWorkbook()
            Dim sheet1 As ISheet = workbook.CreateSheet("Sheet 1")
            Dim countRow = 0
            Dim row1 As IRow = sheet1.CreateRow(countRow)
            Dim cellMainHeader As ICell = row1.CreateCell(0)
            cellMainHeader.SetCellValue("Support Report for Open tickets")

            'xlWorkSheet.Cells(countRow, 1) = "Support Report for Open tickets"
            'Dim xlRangeHeader As Excel.Range
            'xlRangeHeader = xlWorkSheet.Range("A" + countRow.ToString(), "C" + countRow.ToString())
            'xlRangeHeader.Font.Bold = True
            'xlRangeHeader.TextToColumns("Support Report for Open tickets")
            countRow = countRow + 2

            Dim row2 As IRow = sheet1.CreateRow(countRow)
            'sheet1.CreateRow(countRow).CreateCell(0).SetCellValue("Support Report for Open tickets")

            'Dim xlRangeColHeader As Excel.Range
            'xlRangeColHeader = xlWorkSheet.Range("A" + countRow.ToString(), "T" + countRow.ToString())
            'xlRangeColHeader.Font.Bold = True

            row2.CreateCell(0).SetCellValue("Company")
            row2.CreateCell(1).SetCellValue("Issue Type")
            row2.CreateCell(2).SetCellValue("Short Description")
            row2.CreateCell(3).SetCellValue("Description Resolution")
            row2.CreateCell(4).SetCellValue("LINK Name")
            row2.CreateCell(5).SetCellValue("Site Name")
            row2.CreateCell(6).SetCellValue("Reported Issue")
            row2.CreateCell(7).SetCellValue("Observations")
            row2.CreateCell(8).SetCellValue("Actions")
            row2.CreateCell(9).SetCellValue("Results")
            row2.CreateCell(10).SetCellValue("Additional Comments")
            row2.CreateCell(11).SetCellValue("Misc Parts Order")
            row2.CreateCell(12).SetCellValue("Status")
            row2.CreateCell(13).SetCellValue("Created Date")
            row2.CreateCell(14).SetCellValue("Created Time")
            row2.CreateCell(15).SetCellValue("Created By")
            row2.CreateCell(16).SetCellValue("Resolution Date")
            row2.CreateCell(17).SetCellValue("Resolution Time")
            row2.CreateCell(18).SetCellValue("Closed By")
            row2.CreateCell(19).SetCellValue("Ticket Owner")

            countRow = countRow + 1

            For index = 0 To data.count - 1
                Try
                    Dim rowNew As IRow = sheet1.CreateRow(countRow)

                    Dim SupportDateTime As String = IIf(data(index).SupportDateTime Is Nothing, "", data(index).SupportDateTime)
                    Dim IssueTypeText As String = IIf(data(index).IssueTypeText Is Nothing, "", data(index).IssueTypeText)
                    Dim IssueType As String = IIf(data(index).IssueType Is Nothing, "", data(index).IssueType)
                    Dim IssueDescription As String = IIf(data(index).IssueDescription Is Nothing, "", data(index).IssueDescription)
                    Dim Company As String = IIf(data(index).Company Is Nothing, "", data(index).Company)
                    Dim attachment As String = IIf(data(index).attachment Is Nothing, "", data(index).attachment)
                    Dim HubSiteName As String = IIf(data(index).HubSiteName Is Nothing, "", data(index).HubSiteName)
                    Dim FsLink As String = IIf(data(index).FsLink Is Nothing, "", data(index).FsLink)
                    Dim CreatedBy As String = IIf(data(index).CreatedBy Is Nothing, "", data(index).CreatedBy)
                    Dim HUBName As String = IIf(data(index).HUBName Is Nothing, "", data(index).HUBName)
                    Dim LINKName As String = IIf(data(index).LINKName Is Nothing, "", data(index).LINKName)
                    Dim CreatedDate As String = IIf(data(index).CreatedDate Is Nothing, "", data(index).CreatedDate)
                    Dim CreatedTime As String = IIf(data(index).CreatedTime Is Nothing, "", data(index).CreatedTime)
                    Dim ResolutionDate As String = IIf(data(index).ResolutionDate Is Nothing, "", data(index).ResolutionDate)
                    Dim ResolutionTime As String = IIf(data(index).ResolutionTime Is Nothing, "", data(index).ResolutionTime)
                    Dim StatusText As String = IIf(data(index).StatusText Is Nothing, "", data(index).StatusText)
                    Dim Resolution As String = IIf(data(index).Resolution Is Nothing, "", data(index).Resolution)
                    Dim Contact As String = IIf(data(index).Contact Is Nothing, "", data(index).Contact)
                    Dim selectedAdmins As String = IIf(data(index).selectedAdmins Is Nothing, "", data(index).selectedAdmins)
                    Dim ResolutionDateTime As String = IIf(data(index).ResolutionDateTime Is Nothing, "", data(index).ResolutionDateTime)
                    Dim ShortDescription As String = IIf(data(index).ShortDescription Is Nothing, "", data(index).ShortDescription)
                    Dim CaseOpenedBy As String = IIf(data(index).CaseOpenedBy Is Nothing, "", data(index).CaseOpenedBy)
                    Dim CaseClosedBy As String = IIf(data(index).CaseClosedBy Is Nothing, "", data(index).CaseClosedBy)
                    Dim TicketOwner As String = IIf(data(index).TicketOwner Is Nothing, "", data(index).TicketOwner)
                    Dim TicketOwnerEmail As String = IIf(data(index).TicketOwnerEmail Is Nothing, "", data(index).TicketOwnerEmail)
                    Dim ReportedIssue As String = IIf(data(index).ReportedIssue Is Nothing, "", data(index).ReportedIssue)
                    Dim Observations As String = IIf(data(index).Observations Is Nothing, "", data(index).Observations)
                    Dim ActionsPerformed As String = IIf(data(index).ActionsPerformed Is Nothing, "", data(index).ActionsPerformed)
                    Dim Results As String = IIf(data(index).Results Is Nothing, "", data(index).Results)
                    Dim AdditionalComments As String = IIf(data(index).AdditionalComments Is Nothing, "", data(index).AdditionalComments)
                    Dim MiscPartsOrder As String = IIf(data(index).MiscPartsOrder Is Nothing, "", data(index).MiscPartsOrder)

                    rowNew.CreateCell(0).SetCellValue(Company)
                    rowNew.CreateCell(1).SetCellValue(IssueTypeText)
                    rowNew.CreateCell(2).SetCellValue(IssueDescription)
                    rowNew.CreateCell(3).SetCellValue(Resolution)
                    rowNew.CreateCell(4).SetCellValue(LINKName)
                    rowNew.CreateCell(5).SetCellValue(HubSiteName)
                    rowNew.CreateCell(6).SetCellValue(ReportedIssue)
                    rowNew.CreateCell(7).SetCellValue(Observations)
                    rowNew.CreateCell(8).SetCellValue(ActionsPerformed)
                    rowNew.CreateCell(9).SetCellValue(Results)
                    rowNew.CreateCell(10).SetCellValue(AdditionalComments)
                    rowNew.CreateCell(11).SetCellValue(MiscPartsOrder)
                    rowNew.CreateCell(12).SetCellValue(StatusText)
                    rowNew.CreateCell(13).SetCellValue(CreatedDate)
                    rowNew.CreateCell(14).SetCellValue(CreatedTime)
                    rowNew.CreateCell(15).SetCellValue(CreatedBy)
                    rowNew.CreateCell(16).SetCellValue(ResolutionDate)
                    rowNew.CreateCell(17).SetCellValue(ResolutionTime)
                    rowNew.CreateCell(18).SetCellValue(CaseClosedBy)
                    rowNew.CreateCell(19).SetCellValue(TicketOwner)

                    countRow = countRow + 1
                Catch ex As Exception
                    log.Error("Exception occurred in CreateReportAndSendEmail. ex is :" & ex.ToString())
                End Try
            Next
            Dim filePath As String = ConfigurationManager.AppSettings("PathForSaveOpenTicketReport").ToString()
            Dim fullPath = filePath & "SupportReport_" & DateTime.Now.ToString("MMddyyyyss") & ".xls" 'OwnerEmail.Substring(0, 3) & DateTime.Now.ToString("MMddyyyyHHss")"
            'xlWorkBook.SaveAs(fullPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
            'Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            'xlWorkBook.Close(True, misValue, misValue)
            'xlApp.Quit()

            'releaseObject(xlWorkSheet)
            'releaseObject(xlWorkBook)
            'releaseObject(xlApp)
            Using exportData = New MemoryStream()
                workbook.Write(exportData)
                Try
                    If (File.Exists(fullPath)) Then
                        File.Delete(fullPath)
                    End If
                Catch ex As Exception
                    log.Info("In WriteExcelWithNPOI => exception in delete file. filename: " & fullPath & "; exception is: " & ex.ToString())
                End Try


                log.Info("In WriteExcelWithNPOI => step 9")
                Dim bw As BinaryWriter = New BinaryWriter(File.Open(fullPath, FileMode.OpenOrCreate))

                bw.Write(exportData.GetBuffer())

                log.Info("In WriteExcelWithNPOI => step 10")
                bw.Close()
                workbook.Close()

            End Using

            Dim EmailSendCC As String = ConfigurationManager.AppSettings("EmailSendCC").ToString()



            Try

                Dim body As String = String.Empty
                Using sr As New StreamReader(ConfigurationManager.AppSettings("PathForUnResolvedIssueEmailTemplate"))
                    body = sr.ReadToEnd()
                End Using
                '------------------

                body = body.Replace("owneremail", OwnerEmail)

                Try
                    body = body.Replace("ImageSign", "<img src=""https://www.fluidsecure.net/Content/Images/FluidSECURELogo.png"" style=""width:200px""/>")
                    body = body.Replace("SupportTeamName", "FluidSecure Support Team")
                    body = body.Replace("supportemail", "support@fluidsecure.com")
                    body = body.Replace("SupportPhoneNumber", "1-850-878-4585")
                    body = body.Replace("SupportLine1", "Press ""0"" During Normal Business Hours:  Monday - Friday 8:00am - 5:00pm (EST)")
                    body = body.Replace("SupportLine2", "Press ""7"" After Normal Business Hours")
                    body = body.Replace("websiteURLHREF", "https://www.fluidsecure.com")
                    body = body.Replace("webisteURL", "www.fluidsecure.com")
                Catch ex As Exception
                    body = body.Replace("ImageSign", "")
                End Try

                Dim mailClient As New SmtpClient(ConfigurationManager.AppSettings("smtpServer"))
                mailClient.UseDefaultCredentials = False
                mailClient.Credentials = New NetworkCredential(ConfigurationManager.AppSettings("emailAccount"), ConfigurationManager.AppSettings("emailPassword"))
                mailClient.Port = Convert.ToInt32(ConfigurationManager.AppSettings("smtpPort"))

                Dim messageSend As New MailMessage()
                messageSend.Body = body
                messageSend.IsBodyHtml = True
                messageSend.Subject = "***Support Report for Open tickets.***"
                messageSend.From = New MailAddress(ConfigurationManager.AppSettings("FromEmail"))
                If (EmailSendCC IsNot OwnerEmail) Then
                    messageSend.CC.Add(New MailAddress(EmailSendCC)) '
                End If
                If FileExists(fullPath) Then
                    messageSend.Attachments.Add(New Attachment(fullPath))
                End If

                mailClient.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings("EnableSsl"))

                If OwnerEmail <> "" Then
                    messageSend.To.Add(OwnerEmail.Trim()) '
                    mailClient.Send(messageSend)
                    log.Info("Email send to: " + OwnerEmail)
                    messageSend.To.Remove(New MailAddress(OwnerEmail.Trim())) '
                    Try
                        System.IO.File.Delete(fullPath)
                    Catch ex As Exception
                        log.Error("When deleting file after email send : " + ex.ToString())
                    End Try
                End If

                Dim supportBal As SupportBAL = New SupportBAL()

            Catch ex As Exception
                log.Debug("Exception occurred in while sending unresolved email to " & OwnerEmail & " . ex is :" & ex.ToString())
            End Try


        Catch ex As Exception
            log.Error("Exception occurred in CreateReportAndSendEmail. ex is :" & ex.ToString())
        End Try
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            log.Error("Exception occurred in releaseObject. ex is :" & ex.ToString())
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Function FileExists(ByVal FileFullPath As String) _
     As Boolean
        Try
            If Trim(FileFullPath) = "" Then Return False

            Dim f As New IO.FileInfo(FileFullPath)
            Return f.Exists

        Catch ex As Exception
            log.Error("Exception occurred in FileExists. ex is :" & ex.ToString())
        End Try
    End Function
End Module
