Option Strict Off
Imports System
Imports System.Windows.Forms
Imports System.IO
Imports NXOpen
Imports System.Drawing
Imports System.Drawing.Text
Imports NXOpen.UF
Imports System.Security
Imports NXOpen.Utilities
Imports System.Threading
Imports System.Data.OleDb
Imports System.Collections
Imports System.Windows.Forms.MessageBox
Imports System.Collections.Generic
Imports System.Net.Mail
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports NXOpen.Assemblies
Imports NXOpenUI

Module IssueToCAMWorking

    Dim theSession As Session = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
    Dim theUI As UI = UI.GetUI()
    Dim lw As ListingWindow = theSession.ListingWindow()
    Dim ufs As UFSession = UFSession.GetUFSession()
    Dim workPart As Part = theSession.Parts.Work
    Dim displayPart As Part = theSession.Parts.Display
    Dim Form As New IssueForm
    Dim jobNum = theSession.Parts.Work.FullPath()
    Dim checkInArray(0) As String
    Dim tempArray As List(Of String)
    Dim i As Integer = 0
    Dim objWriter As New System.IO.StreamWriter("C:\Eng\IssueToCamCheckIn.txt")
    Dim logWriter As New System.IO.StreamWriter("C:\Eng\89005.txt")
    Dim specificationDrawingsList As New List(Of String)

    Sub Main()
        CreateUsageLog("IssueToCAM")

        'purge all termporary directories (containing email code, pictures)
        If System.IO.File.Exists("C:\eng\tempemail.txt") Then
            System.IO.File.Delete("C:\eng\tempemail.txt")
        End If

        If Directory.Exists("C:\eng\tempattachments") Then
            For Each filep As String In System.IO.Directory.EnumerateFiles("C:\eng\tempattachments")
                System.IO.File.Delete(filep)
            Next
        End If

        lw.Open()
        Form.ShowDialog()
    End Sub

    Public Sub CreateUsageLog(ByVal ProgramName As String)
        Dim username As String = System.Environment.UserName
        Dim UseDate As String = Now().Day & "-" & Now().Month & "-" & Now().Year
        Dim UsageLogFolderDir As String = "u:\logs\UG_Prog"

        If System.IO.Directory.Exists(UsageLogFolderDir) = False Then
            System.IO.Directory.CreateDirectory(UsageLogFolderDir)
        End If

        Dim UsageLogFileName As String = UsageLogFolderDir & "\" & ProgramName & ".log"

        If System.IO.File.Exists(UsageLogFileName) = True Then
            Dim objReader As New System.IO.StreamReader(UsageLogFileName)
            Dim TempContent As String = objReader.ReadToEnd
            objReader.Close()
            objReader.Dispose()

            Dim objWriter As New System.IO.StreamWriter(UsageLogFileName)
            objWriter.WriteLine(TempContent)
            objWriter.WriteLine(username & ";" & UseDate)
            objWriter.Close()
            objWriter.Dispose()
        Else
            Dim objWriter As New System.IO.StreamWriter(UsageLogFileName)
            objWriter.WriteLine("username;DD-MM-YY")
            objWriter.WriteLine(username & ";" & UseDate)
            objWriter.Close()
            objWriter.Dispose()
        End If
    End Sub

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function

    Public Class IssueForm

        Dim IssueEmail As New System.Net.Mail.MailMessage
        Dim Smtp_Server As New System.Net.Mail.SmtpClient
        Dim htmline As String = "<br>" & "Please see Attached Email Below for more details:" & "<br>" & "<hr>" & "<br>"
        Dim myPdfExporter As New NXJ_PdfExporter 'used for PDF creation
        Friend WithEvents DD As System.Diagnostics.Process
        Dim s As NXOpen.Session = NXOpen.Session.GetSession()
        'fyi, setting variables like this in a class is NOT good programming practice!
        Dim Filterstr As String = Nothing 'global for good reason
        Dim Revision As String = Nothing
        Dim ItemID As String = Nothing

        Dim IssueSuccessful As Boolean = False

        Private Sub IssueForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

            Me.AllowDrop = True
            MyBase.AllowDrop = True

            FromBox.Text = System.Environment.UserName & "@stackteck.com" 'from field, is editable
            'ToBox.Text = System.Environment.UserName & "@stackteck.com"  'to field 

            'exception for usernames
            '-------------------------------------------------------------------

            If System.Environment.UserName = "htang" Or System.Environment.UserName = "henryt" Then
                FromBox.Text = "htang@stackteck.com"
            End If

            If System.Environment.UserName = "JOHN" Or System.Environment.UserName = "john" Then
                FromBox.Text = "jbrtka@stackteck.com"
            End If

            If System.Environment.UserName = "susan" Or System.Environment.UserName = "SUSAN" Then
                FromBox.Text = "shiltz@stackteck.com"
            End If

            If System.Environment.UserName = "wei" Or System.Environment.UserName = "WEI" Then
                FromBox.Text = "wyuan@stackteck.com"
            End If
            If System.Environment.UserName.ToLower = "mandeep" Or System.Environment.UserName.ToLower = "mthandi" Then
                FromBox.Text = "mthandi@stackteck.com"
            End If
            '------------------------------

            ComboBox1.Items.Add("New Release")
            ComboBox1.Items.Add("Repeat Job")
            ComboBox1.Items.Add("-----------")
            ComboBox1.Items.Add("Molding Surface Update")
            ComboBox1.Items.Add("Mechanical Stack Update (3D)")
            ComboBox1.Items.Add("Mechanical (2D-NO CAM)")

            IssueEmail.IsBodyHtml = True

            Smtp_Server.EnableSsl = False
            Smtp_Server.Host = "192.168.0.47" 'server for email exchange

            DD_start()
            'for attaching emails

            Dim ComponentNames(0) As String
            Dim ComponentCounts As Integer = 0

            Dim strPartNo As String
            strPartNo = workPart.GetStringAttribute("DB_PART_NO")
            lw.WriteLine("part number: " & strPartNo)

            Filterstr = strPartNo 'part number string 

            Dim strRevision As String
            strRevision = workPart.GetStringAttribute("DB_PART_REV")
            lw.WriteLine("part revision: " & strRevision)
            Revision = strRevision

            ListBox1.Items.Add(Filterstr & " " & Revision)

            ItemID = Filterstr
            Dim tempstr As String = Filterstr.Remove(5, Filterstr.Length - 5)
            Filterstr = tempstr

            Try
                GetAssemblyTree(ComponentNames, ComponentCounts, Filterstr)
            Catch ex As Exception 'for the case that its a single part (no assy)
                MsgBox("175" + Environment.NewLine + ex.ToString)
            End Try

            Dim Count As Integer = 0

            Count = ListBox1.Items.Count
            'get rid of applications part 

            For i As Integer = 0 To Count
                Try
                    If ListBox1.Items.Item(i).ToString.Contains("S" & Filterstr) Then
                        ListBox1.Items.RemoveAt(i)
                        i = i - 1
                        Count = Count - 1
                    End If
                Catch ex As Exception
                End Try
            Next

            Dim emailtemplate As String = "Job Number:        " & Filterstr & vbNewLine & "Rework Number: XXXXX" & vbNewLine & "Reference Job: XXXXX " & vbNewLine & "Comments:"

            RichTextBox1.AppendText(emailtemplate)
        End Sub

        Private Sub DD_start()
            Me.DD = New System.Diagnostics.Process
            Me.DD.EnableRaisingEvents = True
            Me.DD.StartInfo.FileName = "U:\Programs\ReadOutlookFileReturnHTML\ReadOutlookFileReturnHTML\bin\Debug\ReadOutlookFileReturnHTML"
            Me.DD.Start()
        End Sub

        Private Sub DD_close(sender As Object, e As EventArgs) Handles DD.Exited
            Try
                Textbox4.DocumentText = Textbox4.DocumentText.Insert(0, System.IO.File.ReadAllText("C:\eng\tempemail.txt"))
            Catch ex As Exception
                Exit Sub
            End Try

            If DD.HasExited = False Then
                DD.Kill()
            End If
        End Sub

        Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
            'send email and issue button

            Dim RepeatJobValue As String = "XXXXX"
            Dim SearchSubjectBox As Boolean = SubjectBox.Text.Contains(RepeatJobValue)
            Dim RepeatJobNumberValue As String = "Repeat Job Number: XXXXX"
            Dim Searchemailbody As Boolean = RichTextBox1.Text.Contains(RepeatJobNumberValue)

            If (SearchSubjectBox Or Searchemailbody) Then
                Dim msg As String = "Please input the Repeat Job Number in the Subject line and email body by replacing XXXXX"
                MsgBox(msg)
                Exit Sub
            End If

            ProgressBar1.PerformStep()

            Dim releasestatus As String = "TCM RELEASED"

            IssueEmail.Body = "<br>"
            IssueEmail.Body = IssueEmail.Body.Insert(0, Textbox4.DocumentText) 'HTML email from Outlook
            IssueEmail.Body = IssueEmail.Body.Insert(0, ConvertToHTML(RichTextBox1) & htmline)

            ProgressBar1.PerformStep()

            'add recipients

            IssueEmail.To.Add(ToBox.Text)

            Try
                IssueEmail.To.Add(FromBox.Text)

            Catch ex As Exception
                MsgBox(" Error with adding username!")
            End Try

            IssueEmail.From = New System.Net.Mail.MailAddress(FromBox.Text) 'from field
            IssueEmail.Subject = SubjectBox.Text
            IssueEmail.CC.Add("jobfiler@stackteck.local")

            ProgressBar1.PerformStep()

            Try
                For Each filepath As String In System.IO.Directory.EnumerateFiles("C:\eng\tempattachments\")
                    IssueEmail.Attachments.Add(New System.Net.Mail.Attachment(filepath))
                    lw.WriteLine(filepath)
                Next
            Catch ex As Exception
            End Try

            ProgressBar1.PerformStep()

            lw.WriteLine(ItemID & " " & Revision)
            theSession.Parts.SaveAll(True, Nothing)

            ProgressBar1.PerformStep()

            For Each item As String In ListBox1.SelectedItems
                Try
                    CreatePDF(item)
                    'CreatePDF2(item)
                    lw.WriteLine("Creating a pdf of below:")
                    lw.WriteLine(item)
                Catch ex As Exception
                    lw.WriteLine("PDF not made for " & item)
                    lw.WriteLine("")
                    lw.WriteLine(ex.ToString + Environment.NewLine)
                End Try
            Next

            ProgressBar1.PerformStep()

            Try
                displayPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, Nothing)
            Catch ex As Exception
                MsgBox(ex.ToString + Environment.NewLine + "Could not close parent assembly")
                lw.WriteLine(ex.ToString)
            End Try

            For Each specificationDrawing As String In specificationDrawingsList
                Dim part1 As Part = CType(theSession.Parts.FindObject(specificationDrawing), Part)
                part1.Close(BasePart.CloseWholeTree.False, BasePart.CloseModified.UseResponses, Nothing)
            Next

            ProgressBar1.PerformStep()

            RunWorkflow(ItemID & " " & Revision)

            ProgressBar1.PerformStep()

            If IssueSuccessful = False Then
                lw.WriteLine("Error. The Part is Checked out, check it in through Teamcenter please")
                ProgressBar1.Value = 0
                OpenExistingPart(ItemID & "/" & Revision)
                MyBase.Close()
            Else

                Smtp_Server.Send(IssueEmail)
                lw.WriteLine("Email sent")
                ProgressBar1.Value = 100

                MyBase.Close()
            End If
        End Sub

        Private Function hasTCINStatus(thePart As BasePart, theStatus As String) As Boolean
            Dim encodedPath As String
            theUfSession.Part.AskPartName(workPart.Tag, encodedPath)
            Dim partNo As String
            Dim partRev As String
            Dim partName As String
            Dim partType As String
            theUfSession.Ugmgr.DecodePartFileName(encodedPath, partNo, partRev, partType, partName)

            Dim entries As String() = {"item_id", "object_type", "ItemRevision:IMAN_reference.item_revision_id", "release_status_list.name"}
            Dim values As String() = {partNo, "ItemRevision", partRev, theStatus}

            Dim mySearch As NXOpen.PDM.PdmSearch = theSession.PdmSearchManager.NewPdmSearch()
            Dim mySearchResult As NXOpen.PDM.SearchResult = mySearch.Advanced(entries, values)

            If mySearchResult.GetResultObjectNames().Length > 0 Then
                Return True
            End If

            Return False
        End Function

        Private Sub CreatePDF2(filename As String)
            myPdfExporter = New NXJ_PdfExporter

            ' 

        End Sub

        Private Sub CreatePDF(filename As String)

            myPdfExporter = New NXJ_PdfExporter
            Dim pdfname As String = filename
            pdfname = pdfname.Replace(" ", "_")
            pdfname = pdfname + "_PDF"
            'MsgBox(pdfname)
            Dim partloadstatus1 As NXOpen.PartLoadStatus = Nothing

            Dim tempy As String = filename.Replace(" ", "/")
            Dim tempfilename As String = "@DB/" + tempy

            Dim UGPartName As String = Nothing
            FindSpecDwgofMaster(UGPartName)

            Dim PDFPart As NXOpen.Part = Nothing

            Dim specexists As Boolean = False
            'try to open the specification drawing, otherwise, open the normal drawing 

            Try
                ' Soren's Way
                If UGPartName <> Nothing Then
                    OpenExistingSpecPart(tempy, UGPartName)
                    specexists = True
                End If

                ' Karan's Code
                '' theSession.Parts.OpenDisplay(specdwg, partloadstatus1)
                ''PDFPart = thesession.parts.Display
                'PDFPart = CType(theSession.Parts.FindObject(specdwg), Part)
                'theSession.Parts.setdisplay(PDFPart, False, False, partloadstatus1)
                PDFPart = s.Parts.Work

            Catch ex As Exception
                MsgBox("Unable to find specification drawing")
                specexists = False
                partloadstatus1 = Nothing
                PDFPart = Nothing
            End Try

            If specexists = False Then
                Try
                    theSession.Parts.OpenDisplay(tempfilename, partloadstatus1)
                    PDFPart = theSession.Parts.work
                Catch ex As Exception
                    lw.WriteLine(tempfilename & "Didnt create")
                End Try

                Try
                    PDFPart = CType(theSession.Parts.FindObject(tempfilename), Part)
                    theSession.Parts.setdisplay(PDFPart, False, False, partloadstatus1)
                Catch ex As Exception
                    lw.WriteLine(ex.ToString + Environment.NewLine + "Unable to create PDF Part Line 379")
                    Exit Sub
                End Try
            End If

            Try
                myPdfExporter.Part = PDFPart

                '$ set output folder
                myPdfExporter.OutputFolder = "C:\eng\"

                '$ set desired watermark text
                myPdfExporter.UseWatermark = True
                myPdfExporter.WatermarkText = "ISSUE COPY NAME:" & myPdfExporter.OutputPdfFileName.Replace(myPdfExporter.OutputFolder, "")
                myPdfExporter.WatermarkAddDatestamp = True

                '$ show confirmation dialog box on completion
                myPdfExporter.ShowConfirmationDialog = False

                If System.IO.File.Exists(myPdfExporter.OutputPdfFileName) = True Then
                    Try
                        My.Computer.FileSystem.DeleteFile("C:\eng\" + myPdfExporter.OutputPdfFileName)
                    Catch ex As Exception
                    End Try
                End If

                myPdfExporter.Commit()
                IssueEmail.Attachments.Add(New System.Net.Mail.Attachment(myPdfExporter.OutputPdfFileName))

                myPdfExporter = Nothing
            Catch ex As Exception
                lw.WriteLine(ex.ToString + "Problem on line 409")
            End Try
            Try
                '------------------------------------
                '----CREATE PDF IN TEAMCENTER!-------
                '-----------------------------------
                Dim theSession As Session = Session.GetSession()
                Dim workPart As Part = theSession.Parts.Work
                ' ----------------------------------------------
                '   Menu: File->Export->PDF...
                ' ----------------------------------------------
                Dim printPDFBuilder1 As PrintPDFBuilder
                printPDFBuilder1 = workPart.PlotManager.CreatePrintPdfbuilder()
                printPDFBuilder1.Relation = PrintPDFBuilder.RelationOption.Manifestation
                printPDFBuilder1.DatasetType = "PDF"
                printPDFBuilder1.NamedReferenceType = "PDF_Reference"
                printPDFBuilder1.Scale = 1.0
                printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.New
                printPDFBuilder1.DatasetName = pdfname
                printPDFBuilder1.Colors = PrintPDFBuilder.Color.BlackOnWhite
                printPDFBuilder1.Widths = PrintPDFBuilder.Width.CustomThreeWidths
                printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.English
                'printPDFBuilder1.XDimension = 8.5
                'printPDFBuilder1.YDimension = 11.0
                printPDFBuilder1.OutputText = PrintPDFBuilder.OutputTextOption.Polylines
                printPDFBuilder1.RasterImages = True

                Dim partInfo1 As PDM.PartBuilder.PartFileNameData
                ' No object found for the subject of the next call.
                ' This may be because of previous exceptions
                ' partInfo1 = <PDM.PartBuilder>.AssignPartFileName("80999_ASSY_CAV", "000", "manifestation", "80999_ASSY_CAV_000_PDF_2")
                printPDFBuilder1.Assign()
                printPDFBuilder1.Watermark = ""
                Dim sheets1(0) As NXObject
                Dim drawingSheet1 As Drawings.DrawingSheet = CType(workPart.DrawingSheets.FindObject("SHEET1"), Drawings.DrawingSheet)
                sheets1(0) = drawingSheet1
                printPDFBuilder1.SourceBuilder.SetSheets(sheets1)
                printPDFBuilder1.DatasetName = pdfname
                'add items here with extensive reasearch 
                Dim nXObject1 As NXObject
                nXObject1 = printPDFBuilder1.Commit()
                printPDFBuilder1.Destroy()

                ' ----------------------------------------------
                '   Menu: Tools->Journal->Stop Recording
                ' ----------------------------------------------



            Catch ex As Exception
                'MsgBox("PDF to TeamCenter Error:" + System.Environment.NewLine + ex.ToString)
            End Try

            If (specexists = True) Then
                specificationDrawingsList.Add(tempfilename + "/specification/" + UGPartName)
            End If
        End Sub

        Sub OpenExistingSpecPart(ByVal FileName As String, ByVal UGPartName As String)
            lw.WriteLine("Open Existing Spec Part Dwg ......")
            Dim basePart1 As BasePart
            Dim partLoadStatus1 As PartLoadStatus = Nothing

            lw.WriteLine("File Name: " & FileName)

            lw.WriteLine("@DB/" & FileName & "/specification/" & UGPartName)

            Try
                ' File already exists
                lw.WriteLine("@DB/" & FileName & "/specification/" & UGPartName)
                basePart1 = s.Parts.OpenBaseDisplay("@DB/" & FileName & "/specification/" & UGPartName, partLoadStatus1)
            Catch ex As NXException
                ex.AssertErrorCode(1020004)
            End Try

            Dim markId3 As Session.UndoMarkId
            markId3 = s.SetUndoMark(Session.MarkVisibility.Visible, "Change Display Part")

            Dim part1 As Part = CType(s.Parts.FindObject("@DB/" & FileName & "/specification/" & UGPartName), Part)

            Dim partLoadStatus2 As PartLoadStatus = Nothing
            Dim status1 As PartCollection.SdpsStatus
            status1 = s.Parts.SetDisplay(part1, True, True, partLoadStatus2)

            s.Parts.SetWork(part1)
        End Sub

        Public Sub FindSpecDwgofMaster(ByRef UGPartName As String)

            Dim myPartTag As Tag = ufs.Part.AskDisplayPart

            Dim EncodedName As String = Nothing
            ufs.Part.AskPartName(myPartTag, EncodedName)

            Dim PartNum As String = Nothing
            Dim PartRev As String = Nothing
            Dim PartFileType As String = Nothing
            Dim PartFileName As String = Nothing

            ufs.Ugmgr.DecodePartFileName(EncodedName, PartNum, PartRev, PartFileType, PartFileName)

            lw.WriteLine("Master Part Info....")
            lw.WriteLine("Part Name: " & PartNum)
            lw.WriteLine("Part Rev: " & PartRev)
            lw.WriteLine("Part File Type: " & PartFileType)
            lw.WriteLine("")

            Dim myDBTag As Tag
            ufs.Ugmgr.AskPartTag(PartNum, myDBTag)

            Dim NumOfRev As Integer
            Dim myRevTags As Tag() = Nothing

            ufs.Ugmgr.ListPartRevisions(myDBTag, NumOfRev, myRevTags)

            For i As Integer = 0 To myRevTags.Length - 1
                Dim TempRev As String = Nothing
                ufs.Ugmgr.AskPartRevisionId(myRevTags(i), TempRev)

                If TempRev = PartRev Then
                    Dim NumOfUGPart As Integer
                    Dim UGPartTypes As String() = Nothing
                    Dim UGPartNames As String() = Nothing
                    ufs.Ugmgr.ListPartRevFiles(myRevTags(i), NumOfUGPart, UGPartTypes, UGPartNames)

                    For j As Integer = 0 To UGPartNames.Length - 1
                        Select Case UGPartNames(j)
                            Case "dwg"
                                UGPartName = UGPartNames(j)
                                Exit For
                        End Select
                    Next

                    lw.WriteLine(" UGPart Name: " & UGPartName)
                    Exit Sub
                End If
            Next
        End Sub

        Private Sub RunWorkflow(name As String)

            Dim time As Integer = 0
            Dim ProgType As String = Nothing 'issue type in TC, sent as an argument to the batch script

            If PartIssue.Checked = True Then
                ProgType = "PART" 'this is not used anymore, so ignore this part of the code:
            ElseIf AssyIssue.Checked = True Then
                ProgType = "ASSY"
            Else
                ProgType = Nothing
            End If

            Dim checkInAllPartsExe As System.Diagnostics.Process = New System.Diagnostics.Process
            checkInAllPartsExe.StartInfo.FileName = "U:\Programs\Admin\checkInPart.bat"
            checkInAllPartsExe.Start()

            While checkInAllPartsExe.HasExited = False
            End While

            Dim WorkflowExe As System.Diagnostics.Process = New System.Diagnostics.Process
            WorkflowExe.StartInfo.FileName = "U:\Programs\Admin\RunIssueToCAM.bat"
            WorkflowExe.StartInfo.Arguments = name & " " & ProgType
            WorkflowExe.Start()

            'lw.writeline(WorkflowExe.ProcessName)
            'lw.writeline(WorkflowExe.MainWindowTitle)

            While WorkflowExe.HasExited = False
            End While

            If WorkflowExe.ExitCode = -1 Then
                IssueSuccessful = False
                lw.WriteLine("Exit Code is: " & WorkflowExe.ExitCode.ToString)
            Else
                lw.WriteLine("Exit Code is: " & WorkflowExe.ExitCode.ToString)
                IssueSuccessful = True
            End If

            'System.Diagnostics.Process.Start("U:\Programs\Admin\RunIssueToCAM.bat", name & " " & ProgType)
            'Dim Process As Process() = Nothing
        End Sub

        Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
            'issue type selection

            SubjectBox.ReadOnly = True
            Select Case ComboBox1.Text.ToString

                Case "Molding Surface Update"
                    'ToBox.Text = "rnaveed@stackteck.com"
                    ToBox.Text = "jngai@stackteck.com,rnaveed@stackteck.com,engcam@stackteck.local,pm@stackteck.local,jreis@stackteck.com,vtravaglini@stackteck.local,pcesario@stackteck.local,davidd@stackteck.com" & TeamLeader(System.Environment.UserName)

                    SubjectBox.Text = "ENG MOLDING SURFACE UPDATE/REWORK | " & Filterstr & " | " & DateAndTime.DateString & " e90" & Filterstr

                    Dim emailtemplate As String = "Job Number:        " & Filterstr & vbNewLine & "Rework Number: XXXXX" & vbNewLine & "Reference Job: XXXXX " & vbNewLine & "Comments:"
                    RichTextBox1.Text = emailtemplate

                Case "Mechanical Stack Update (3D)"
                    'ToBox.Text = "jngai@stackteck.com,SSarvestany@stackteck.com"
                    ToBox.Text = "jngai@stackteck.com,rnaveed@stackteck.com,engcam@stackteck.local,pm@stackteck.local,jreis@stackteck.com,vtravaglini@stackteck.local,pcesario@stackteck.local,davidd@stackteck.com" & TeamLeader(System.Environment.UserName)

                    SubjectBox.Text = "ENG STACK UPDATE - MECHANICAL - CAM | " & Filterstr & " | " & DateAndTime.DateString & " e90" & Filterstr

                    Dim emailtemplate As String = "Job Number:        " & Filterstr & vbNewLine & "Rework Number: XXXXX" & vbNewLine & "Reference Job: XXXXX " & vbNewLine & "Comments:"
                    RichTextBox1.Text = emailtemplate

                Case "Mechanical (2D-NO CAM)"
                    'ToBox.Text = "jngai@stackteck.com,SSarvestany@stackteck.com"
                    ToBox.Text = "jngai@stackteck.com,rnaveed@stackteck.com,engcam@stackteck.local,pm@stackteck.local,davidd@stackteck.com" & TeamLeader(System.Environment.UserName)

                    SubjectBox.Text = "ENG 2D UPDATE: FYI " & ItemID & " | " & DateAndTime.DateString & " e90" & Filterstr

                    Dim emailtemplate As String = "Job Number:        " & Filterstr & vbNewLine & "Rework Number: XXXXX" & vbNewLine & "Reference Job: XXXXX " & vbNewLine & "Comments:"
                    RichTextBox1.Text = emailtemplate

                Case "New Release"
                    'ToBox.Text = "rnaveed@stackteck.com"
                    ToBox.Text = "jngai@stackteck.com,rnaveed@stackteck.com,engcam@stackteck.local,pm@stackteck.local,davidd@stackteck.com" & TeamLeader(System.Environment.UserName)

                    SubjectBox.Text = "ENG NEW RELEASE | " & Filterstr & " | " & DateAndTime.DateString & " e90" & Filterstr

                    Dim emailtemplate As String = "Job Number:        " & Filterstr & vbNewLine & "Rework Number: XXXXX" & vbNewLine & "Reference Job: XXXXX " & vbNewLine & "Comments:"
                    RichTextBox1.Text = emailtemplate

                Case "Repeat Job"
                    'ToBox.Text = "davidd@stackteck.com,jngai@stackteck.com"
                    ToBox.Text = "jngai@stackteck.com,rnaveed@stackteck.com,engcam@stackteck.local,pm@stackteck.local,davidd@stackteck.com" & TeamLeader(System.Environment.UserName)

                    Dim emailtemplate As String = "Repeat Job Number: XXXXX" & vbNewLine & "Identical To:      " & Filterstr & vbNewLine & "Rework Number: XXXXX" & vbNewLine & "Reference Job: XXXXX " & vbNewLine & "Comments:"


                    RichTextBox1.Text = emailtemplate
                    SubjectBox.Text = "ENG REPEAT JOB | XXXXX | " & DateAndTime.DateString & " e90XXXXX"
                    SubjectBox.ReadOnly = False
                Case Else
                    MsgBox("Please select an issue type from the Drop Down", MsgBoxStyle.Information, "Error")
            End Select

            If (System.Environment.UserName.ToUpper.Contains("NAVEED")) Then
                ToBox.Text = "RNAVEED@STACKTECK.COM"
            End If
        End Sub

        Private Sub Button9_Click(sender As Object, e As EventArgs)
            Dim CopiedEmail As String = (System.Windows.Forms.Clipboard.GetData(System.Windows.Forms.DataFormats.Html))

            CopiedEmail.Insert(0, htmline)

            Textbox4.DocumentText = vbNewLine & Textbox4.DocumentText.Insert(Textbox4.DocumentText.Length, CopiedEmail)
        End Sub

        Public Function ConvertToHTML(ByVal Box As System.Windows.Forms.RichTextBox) _
                   As String
            ' Takes a RichTextBox control and returns a
            ' simple HTML-formatted version of its contents
            Dim strHTML As String
            Dim strColour As String
            Dim blnBold As Boolean
            Dim blnItalic As Boolean
            Dim strFont As String
            Dim shtSize As Short
            Dim lngOriginalStart As Long
            Dim lngOriginalLength As Long
            Dim intCount As Integer
            ' If nothing in the box, exit
            If Box.Text.Length = 0 Then Exit Function
            ' Store original selections, then select first character
            lngOriginalStart = 0
            lngOriginalLength = Box.TextLength
            Box.Select(0, 1)
            ' Add HTML header
            strHTML = "<html>"
            ' Set up initial parameters
            strColour = Box.SelectionColor.ToKnownColor.ToString
            blnBold = Box.SelectionFont.Bold
            blnItalic = Box.SelectionFont.Italic
            strFont = Box.SelectionFont.FontFamily.Name
            shtSize = Box.SelectionFont.Size
            ' Include first 'style' parameters in the HTML
            strHTML += "<span style=""font-family: " & strFont &
              "; font-size: " & shtSize & "pt; color: " _
                              & strColour & """>"
            ' Include bold tag, if required
            If blnBold = True Then
                strHTML += "<b>"
            End If
            ' Include italic tag, if required
            If blnItalic = True Then
                strHTML += "<i>"
            End If
            ' Finally, add our first character
            strHTML += Box.Text.Substring(0, 1)
            ' Loop around all remaining characters
            For intCount = 2 To Box.Text.Length
                ' Select current character
                Box.Select(intCount - 1, 1)
                ' If this is a line break, add HTML tag
                If Box.Text.Substring(intCount - 1, 1) =
                       Convert.ToChar(10) Then
                    strHTML += "<br>"
                End If
                ' Check/implement any changes in style
                If Box.SelectionColor.ToKnownColor.ToString <> strColour Or Box.SelectionFont.FontFamily.Name _
                   <> strFont Or Box.SelectionFont.Size <> shtSize Then
                    strHTML += "</span><span style=""font-family: " _
                      & Box.SelectionFont.FontFamily.Name &
                      "; font-size: " & Box.SelectionFont.Size &
                      "pt; color: " &
                      Box.SelectionColor.ToKnownColor.ToString & """>"
                End If
                ' Check for bold changes
                If Box.SelectionFont.Bold <> blnBold Then
                    If Box.SelectionFont.Bold = False Then
                        strHTML += "</b>"
                    Else
                        strHTML += "<b>"
                    End If
                End If
                ' Check for italic changes
                If Box.SelectionFont.Italic <> blnItalic Then
                    If Box.SelectionFont.Italic = False Then
                        strHTML += "</i>"
                    Else
                        strHTML += "<i>"
                    End If
                End If
                ' Add the actual character
                strHTML += Mid(Box.Text, intCount, 1)
                ' Update variables with current style
                strColour = Box.SelectionColor.ToKnownColor.ToString
                blnBold = Box.SelectionFont.Bold
                blnItalic = Box.SelectionFont.Italic
                strFont = Box.SelectionFont.FontFamily.Name
                shtSize = Box.SelectionFont.Size
            Next
            ' Close off any open bold/italic tags
            If blnBold = True Then strHTML += ""
            If blnItalic = True Then strHTML += ""
            ' Terminate outstanding HTML tags
            strHTML += "</span></html>"
            ' Restore original RichTextBox selection
            Box.Select(lngOriginalStart, lngOriginalLength)
            ' Return HTML
            Return strHTML
        End Function

        Sub GetAssemblyTree(ByRef ComponentNames As String(), ByRef ComponentCounts As Integer, ByVal JobNum As String)

            Dim theUI As NXOpen.UI = UI.GetUI()
            Dim ufs As UFSession = UFSession.GetUFSession()
            Dim lw As ListingWindow = s.ListingWindow

            Dim part1 As NXOpen.Part
            part1 = s.Parts.Work

            Dim c As Component

            Try
                c = part1.ComponentAssembly.RootComponent
            Catch ex As Exception
            End Try

            Dim count As Integer = 1
            Dim suppressCount As Integer = 1
            Dim TempComponentName(0) As String
            Dim suppressComponentNames(0) As String

            Try
                TempComponentName(0) = c.DisplayName
            Catch ex As Exception
                TempComponentName(0) = s.Parts.Work.FullPath()
            End Try

            Try
                ShowAssemblyTree(c, "", count, TempComponentName)
                createCheckInList(c, "", suppressCount, suppressComponentNames)
            Catch ex As Exception
                lw.WriteLine("This wasn't an assembly. No children to be found.")
            End Try

            ListBox2.Items.Add(s.Parts.Work.FullPath().Substring(0, s.Parts.Work.FullPath().Length - 4))

            objWriter.WriteLine("Y:\eng\IssueToCamLogs\" + JobNum + ".txt")
            objWriter.WriteLine(ListBox2.Items.Count)

            For Each item As String In ListBox2.Items
                objWriter.WriteLine(item)
            Next

            objWriter.Flush()
            objWriter.Close()
            objWriter.Dispose()

            lw.WriteLine("Total Count: " & count)
            lw.WriteLine("")

            Try
                ComponentNames(0) = c.DisplayName
            Catch ex As Exception
            End Try

            Dim i As Integer = 0

            For Each Tempstr As String In TempComponentName
                If Tempstr.StartsWith(JobNum) = True Then

                    Dim componentFound As Boolean = False

                    For Each componentname As String In ComponentNames
                        If Tempstr = componentname Then
                            componentFound = True
                            Exit For
                        End If
                    Next

                    If componentFound = False Then
                        i += 1
                        ReDim Preserve ComponentNames(i)
                        ComponentNames(i) = Tempstr
                        lw.WriteLine("Component Found: " & Tempstr)
                    End If
                End If
            Next
        End Sub

        Sub ShowAssemblyTree(ByVal c As Component, ByVal indent As String, ByRef count As Integer, ByRef TempComponentName As String())

            Dim children As Component() = c.GetChildren()
            Dim newIndent As String

            For Each child As Component In children

                If indent.Length = 0 Then
                    newIndent = " "
                Else
                    newIndent = indent & " "
                End If

                If child.IsSuppressed = False Then
                    count += 1
                    lw.WriteLine(newIndent & child.DisplayName)
                    ReDim Preserve TempComponentName(count - 1)
                    TempComponentName(count - 1) = child.DisplayName
                    ShowAssemblyTree(child, newIndent, count, TempComponentName)
                Else
                    lw.WriteLine(newIndent & child.DisplayName)
                End If
            Next

            For Each Name As String In TempComponentName

                Name = Name.Replace("/", " ")
                If ListBox1.Items.Contains(Name) Then
                    Continue For
                Else
                    Name = Name.Replace("/", " ")
                    ListBox1.Items.Add(Name)
                End If
            Next

            'Dim partload As partloadstatus
            'Dim p As basepart = CType(workPart.ComponentAssembly.RootComponent.FindObject("COMPONENT " & Name & " 1"), NXOpen.BasePart)
            'If p.isreadonly = True Then
            'Continue For
            ' End If
        End Sub

        Sub createCheckInList(ByVal c As Component, ByVal indent As String, ByRef suppressCount As Integer, ByRef suppressComponentNames As String())

            Dim children As Component() = c.GetChildren()
            Dim newIndent As String

            For Each child As Component In children

                If indent.Length = 0 Then
                    newIndent = " "
                Else
                    newIndent = indent & " "
                End If

                suppressCount += 1
                lw.WriteLine(newIndent & child.DisplayName)
                'ReDim Preserve suppressComponentNames(suppressCount - 1)
                'suppressComponentNames(suppressCount - 1) = child.DisplayName
                If (ListBox2.Items.Contains(child.DisplayName.Substring(0, child.DisplayName.Length - 4)) = False) Then
                    ListBox2.Items.Add(child.DisplayName.Substring(0, child.DisplayName.Length - 4))
                End If

                lw.WriteLine(newIndent & child.DisplayName & "is in suppress status")
                createCheckInList(child, newIndent, suppressCount, suppressComponentNames)
            Next

        End Sub
        Function OpenExistingPart(ByVal FileName As String) As Boolean
            Dim basePart1 As BasePart
            Dim partLoadStatus1 As PartLoadStatus = Nothing

            lw.WriteLine("File Name: " & FileName)
            Dim Testarray As String() = FileName.Split("/")
            lw.WriteLine("Test Array: " & Testarray.Length)
            If Testarray.Length > 2 Then
                MessageBox.Show("Part " & FileName & " will be skipped in this batch.  Please open this part and run this part individually.", "Invalid File Name", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                Return False
                Exit Function
            End If

            Try
                ' File already exists
                basePart1 = s.Parts.OpenBaseDisplay("@DB/" & FileName, partLoadStatus1)
            Catch ex As NXException
                ex.AssertErrorCode(1020004)
            End Try

            Dim markId3 As Session.UndoMarkId
            markId3 = s.SetUndoMark(Session.MarkVisibility.Visible, "Change Display Part")

            Dim part1 As Part = CType(s.Parts.FindObject("@DB/" & FileName), Part)

            Dim partLoadStatus2 As PartLoadStatus = Nothing
            Dim status1 As PartCollection.SdpsStatus
            status1 = s.Parts.SetDisplay(part1, True, True, partLoadStatus2)

            s.Parts.SetWork(part1)

            Return True
        End Function

        Private Sub PartIssue_CheckedChanged(sender As Object, e As EventArgs) Handles PartIssue.CheckedChanged
            If PartIssue.Checked = True Then
                If SubjectBox.Text.Contains("REPAIR") = True Then
                    SubjectBox.Text = SubjectBox.Text.Replace("REPAIR", "")
                    SubjectBox.Text = SubjectBox.Text.Insert(0, "ENG")
                End If
                AssyIssue.Checked = False
            End If
        End Sub

        Private Sub AssyIssue_CheckedChanged(sender As Object, e As EventArgs) Handles AssyIssue.CheckedChanged
            If AssyIssue.Checked = True Then
                If SubjectBox.Text.Contains("ENG") = True Then
                    SubjectBox.Text = SubjectBox.Text.Replace("ENG", "")
                    SubjectBox.Text = SubjectBox.Text.Insert(0, "REPAIR")
                End If
                PartIssue.Checked = False
            End If
        End Sub

        Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
            System.Diagnostics.Process.Start("\\cntfiler\_JobLib\EngProcedures\UG\ENGA04T - NX Stack Parts And Assembly Issue Procedure.pdf")
        End Sub

        Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
            theSession.Parts.SaveAll(True, Nothing)
            displayPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, Nothing)
            RunWorkflow(ItemID & " " & Revision)
            MyBase.Close()
        End Sub

        Private Function TeamLeader(ByRef Name As String) As String
            Dim LeaderEmail As String
            Dim Group As String

            lw.WriteLine(Name)

            If Name.Equals("tcretu") = True Or Name.Equals("theodor") = True Or Name.Equals("jjin") = True Or Name.Equals("davidd") = True Or Name.Equals("kshukla") = True Then
                Group = "Delta"
            ElseIf Name.Equals("twang") Or Name.Equals("tonym") = True Or Name = "wei" Or Name = "mabdellatif" Then
                Group = "Alpha"
            ElseIf Name = "henryt" Or Name = "mthandi" Or Name = "weiyu" Or Name = "mconnors" Then
                Group = "Omega"
            ElseIf Name.Equals("susan") = True Or Name = "ahall" Or Name = "kchiu" Or Name = "kevin" Or Name = "SUSAN" Then
                Group = "Helix"
            ElseIf Name = "jbrtka" Or Name = "skhattar" Or Name = "john" Or Name = "JOHN" Then
                Group = "Axis"
            Else
                lw.WriteLine("error finding group")
            End If

            lw.WriteLine(Group)

            Select Case Group
                Case "Delta"
                    LeaderEmail = ",tcretu@stackteck.com"
                Case "Alpha"
                    LeaderEmail = ",tmorrone@stackteck.com"
                Case "Omega"
                    LeaderEmail = ",mthandi@stackteck.com"
                Case "Helix"
                    LeaderEmail = ",kchiu@stackteck.com"
                Case "Axis"
                    LeaderEmail = ",jbrtka@stackteck.com"
                Case Else
                    lw.WriteLine("Error finding leader email")
            End Select

            lw.WriteLine(LeaderEmail)

            Return LeaderEmail
        End Function
    End Class

    Class NXJ_PdfExporter


#Region "information"

        'NXJournaling.com
        'Jeff Gesiakowski
        'December 9, 2013
        '
        'NX 8.5
        'class to export drawing sheets to pdf file
        '
        'Please send any bug reports and/or feature requests to: info@nxjournaling.com
        '
        'version 0.4 {beta}, initial public release
        '  special thanks to Mike H. and Ryan G.
        '
        'November 3, 2014
        'update to version 0.6
        '  added a public Sort(yourSortFunction) method to sort the drawing sheet collection according to a custom supplied function.
        '
        'November 7, 2014
        'Added public property: ExportSheetsIndividually and related code changes default value: False
        'Changing this property to True will cause each sheet to be exported to an individual pdf file in the specified export folder.
        '
        'Added new public method: New(byval thePart as Part)
        '  allows you to specify the part to use at the time of the NXJ_PdfExporter object creation
        '
        '
        'December 1, 2014
        'update to version 1.0
        'Added public property: SkipBlankSheets [Boolean] {read/write} default value: True
        '   If the drawing sheet contains no visible objects, it is not output to the pdf file.
        '   Checking the sheet for visible objects requires the sheet to be opened;
        '   display updating is suppressed while the check takes place.
        'Bugfix:
        '   If the PickExportFolder method was used and the user pressed cancel, a later call to Commit would still execute and send the output to a temp folder.
        '   Now, if cancel is pressed on the folder browser dialog a boolean flag is set (_cancelOutput), which the other methods will check before executing.
        '
        '
        'December 4, 2014
        'update to version 1.0.1
        'Made changes to .OutputPdfFileName property Set method: you can pass in the full path to the file (e.g. C:\temp\pdf-output\12345.pdf);
        'or you can simply pass in the base file name for the pdf file (e.g. 12345 or 12345.pdf). The full path will be built based on the
        '.OutputFolder property (default value - same folder as the part file, if an invalid folder path is specified it will default to the user's
        '"Documents" folder).
        '
        '
        'Public Properties:
        '  ExportSheetsIndividually [Boolean] {read/write} - flag indicating that the drawing sheets should be output to individual pdf files.
        '           cannot be used if ExportToTc = True
        '           default value: False
        '  ExportToTc [Boolean] {read/write} - flag indicating that the pdf should be output to the TC dataset, False value = output to filesystem
        '           default value: False
        '  IsTcRunning [Boolean] {read only} - True if NX is running under TC, false if native NX
        '  OpenPdf [Boolean] {read/write} - flag to indicate whether the journal should attempt to open the pdf after creation
        '           default value: False
        '  OutputFolder [String] {read/write} - path to the output folder
        '           default value (native): folder of current part
        '           default value (TC): user's Documents folder
        '  OutputPdfFileName [String] {read/write} - full file name of outut pdf file (if exporting to filesystem)
        '           default value (native): <folder of current part>\<part name>_<part revision>{_preliminary}.pdf
        '           default value (TC): <current user's Documents folder>\<DB_PART_NO>_<DB_PART_REV>{_preliminary}.pdf
        '  OverwritePdf [Boolean] {read/write} - flag indicating that the pdf file should be overwritten if it already exists
        '                                           currently only applies when exporting to the filesystem
        '           default value: True
        '  Part [NXOpen.Part] {read/write} - part that contains the drawing sheets of interest
        '           default value: none, must be set by user
        '  PartFilePath [String] {read only} - for native NX part files, the path to the part file
        '  PartNumber [String] {read only} - for native NX part files: part file name
        '                                    for TC files: value of DB_PART_NO attribute
        '  PartRevision [String] {read only} - for native NX part files: value of part "Revision" attribute, if present
        '                                      for TC files: value of DB_PART_REV
        '  PreliminaryPrint [Boolean] {read/write} - flag indicating that the pdf should be marked as an "preliminary"
        '                                       when set to True, the output file will be named <filename>_preliminary.pdf
        '           default value: False
        '  SheetCount [Integer] {read only} - integer indicating the total number of drawing sheets found in the file
        '  ShowConfirmationDialog [Boolean] {read/write} - flag indicating whether to show the user a confirmation dialog after pdf is created
        '                                                   if set to True and ExportToTc = False, user will be asked if they want to open the pdf file
        '                                                   if user chooses "Yes", the code will attempt to open the pdf with the default viewer
        '           default value: False
        '  SkipBlankSheets [Boolean] {read/write} - flag indicating if the user wants to skip drawing sheets with no visible objects.
        '           default value: True
        '  SortSheetsByName [Boolean] {read/write} - flag indicating that the sheets should be sorted by name before output to pdf
        '           default value: True
        '  TextAsPolylines [Boolean] {read/write} - flag indicating that text objects should be output as polylines instead of text objects
        '           default value: False        set this to True if you are using an NX font and the output changes the 'look' of the text
        '  UseWatermark [Boolean] {read/write} - flag indicating that watermark text should be applied to the face of the drawing
        '           default value: False
        '  WatermarkAddDatestamp [Boolean] {read/write} - flag indicating that today's date should be added to the end of the
        '                                                   watermark text
        '           default value: True
        '  WatermarkText [String] {read/write} - watermark text to use
        '           default value: "PRELIMINARY PRINT NOT TO BE USED FOR PRODUCTION"
        '
        'Public Methods:
        '  New() - initializes a new instance of the class
        '  New(byval thePart as Part) - initializes a new instance of the class and specifies the NXOpen.Part to use
        '  PickExportFolder() - displays a FolderPicker dialog box, the user's choice will be set as the output folder
        '  PromptPreliminaryPrint() - displays a yes/no dialog box asking the user if the print should be marked as preliminary
        '                               if user chooses "Yes", PreliminaryPrint and UseWatermark are set to True
        '  PromptWatermarkText() - displays an input box prompting the user to enter text to use for the watermark
        '                           if cancel is pressed, the default value is used
        '                           if Me.UseWatermark = True, an input box will appear prompting the user for the desired watermark text. Initial text = Me.WatermarkText
        '                           if Me.UseWatermark = False, calling this method will have no effect
        '  Commit() - using the specified options, export the given part's sheets to pdf
        '  SortDrawingSheets() - sorts the drawing sheets in alphabetic order
        '  SortDrawingSheets(ByVal customSortFunction As System.Comparison(Of NXOpen.Drawings.DrawingSheet)) - sorts the drawing sheets by the custom supplied function
        '    signature of the sort function must be: {function name}(byval x as Drawings.Drawingsheet, byval y as Drawings.DrawingSheet) as Integer
        '    a return value < 0 means x comes before y
        '    a return value > 0 means x comes after y
        '    a return value = 0 means they are equal (it doesn't matter which is first in the resulting list)
        '    after writing your custom sort function in the module, pass it in like this: myPdfExporter.Sort(AddressOf {function name})


#End Region


#Region "properties and private variables"

        Private Const Version As String = "1.0.1"

        Private _theSession As Session = Session.GetSession
        Private _theUfSession As UFSession = UFSession.GetUFSession
        Private lg As LogFile = _theSession.LogFile

        Private _cancelOutput As Boolean = False
        Private _drawingSheets As New List(Of Drawings.DrawingSheet)

        Private _exportFile As String = ""
        Private _partUnits As Integer
        Private _watermarkTextFinal As String = ""
        Private _outputPdfFiles As New List(Of String)

        Private _exportSheetsIndividually As Boolean = False
        Public Property ExportSheetsIndividually() As Boolean
            Get
                Return _exportSheetsIndividually
            End Get
            Set(ByVal value As Boolean)
                lg.WriteLine("Set Property ExportSheetsIndividually")
                _exportSheetsIndividually = value
                lg.WriteLine("  ExportSheetsIndividually: " & value.ToString)
                lg.WriteLine("exiting Set Property ExportSheetsIndividually")
                lg.WriteLine("")
            End Set
        End Property

        Private _exportToTC As Boolean = False
        Public Property ExportToTc() As Boolean
            Get
                Return _exportToTC
            End Get
            Set(ByVal value As Boolean)
                lg.WriteLine("Set Property ExportToTc")
                _exportToTC = value
                lg.WriteLine("  exportToTc: " & _exportToTC.ToString)
                Me.GetOutputName()
                lg.WriteLine("exiting Set Property ExportToTc")
                lg.WriteLine("")
            End Set
        End Property

        Private _isTCRunning As Boolean
        Public ReadOnly Property IsTCRunning() As Boolean
            Get
                Return _isTCRunning
            End Get
        End Property

        Private _openPdf As Boolean = False
        Public Property OpenPdf() As Boolean
            Get
                Return _openPdf
            End Get
            Set(ByVal value As Boolean)
                lg.WriteLine("Set Property OpenPdf")
                _openPdf = value
                lg.WriteLine("  openPdf: " & _openPdf.ToString)
                lg.WriteLine("exiting Set Property OpenPdf")
                lg.WriteLine("")
            End Set
        End Property

        Private _outputFolder As String = ""
        Public Property OutputFolder() As String
            Get
                Return _outputFolder
            End Get
            Set(ByVal value As String)
                lg.WriteLine("Set Property OutputFolder")
                If _cancelOutput Then
                    lg.WriteLine("  export pdf canceled")
                    Exit Property
                End If
                If Not Directory.Exists(value) Then
                    Try
                        lg.WriteLine("  specified directory does not exist, trying to create it...")
                        Directory.CreateDirectory(value)
                        lg.WriteLine("  directory created: " & value)
                    Catch ex As Exception
                        lg.WriteLine("  ** error while creating directory: " & value)
                        lg.WriteLine("  " & ex.GetType.ToString & " : " & ex.Message)
                        lg.WriteLine("  defaulting to: " & My.Computer.FileSystem.SpecialDirectories.MyDocuments)
                        value = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                    End Try
                End If
                _outputFolder = value
                _outputPdfFile = IO.Path.Combine(_outputFolder, _exportFile & ".pdf")
                lg.WriteLine("  outputFolder: " & _outputFolder)
                lg.WriteLine("  outputPdfFile: " & _outputPdfFile)
                lg.WriteLine("exiting Set Property OutputFolder")
                lg.WriteLine("")
            End Set
        End Property

        Private _outputPdfFile As String = ""
        Public Property OutputPdfFileName() As String
            Get
                Return _outputPdfFile
            End Get
            Set(ByVal value As String)
                lg.WriteLine("Set Property OutputPdfFileName")
                lg.WriteLine("  value passed to property: " & value)
                _exportFile = IO.Path.GetFileName(value)
                If _exportFile.Substring(_exportFile.Length - 4, 4).ToLower = ".pdf" Then
                    'strip off ".pdf" extension
                    _exportFile = _exportFile.Substring(_exportFile.Length - 4, 4)
                End If
                lg.WriteLine("  _exportFile: " & _exportFile)
                If Not value.Contains("\") Then
                    lg.WriteLine("  does not appear to contain path information")
                    'file name only, need to add output path
                    _outputPdfFile = IO.Path.Combine(Me.OutputFolder, _exportFile & ".pdf")
                Else
                    'value contains path, update _outputFolder
                    lg.WriteLine("  value contains path, updating the output folder...")
                    lg.WriteLine("  parent path: " & Me.GetParentPath(value))
                    Me.OutputFolder = Me.GetParentPath(value)
                    _outputPdfFile = IO.Path.Combine(Me.OutputFolder, _exportFile & ".pdf")
                End If
                '_outputPdfFile = value
                lg.WriteLine("  outputPdfFile: " & _outputPdfFile)
                lg.WriteLine("  outputFolder: " & Me.OutputFolder)
                lg.WriteLine("exiting Set Property OutputPdfFileName")
                lg.WriteLine("")
            End Set
        End Property

        Private _overwritePdf As Boolean = True
        Public Property OverwritePdf() As Boolean
            Get
                Return _overwritePdf
            End Get
            Set(ByVal value As Boolean)
                lg.WriteLine("Set Property OverwritePdf")
                _overwritePdf = value
                lg.WriteLine("  overwritePdf: " & _overwritePdf.ToString)
                lg.WriteLine("exiting Set Property OverwritePdf")
                lg.WriteLine("")
            End Set
        End Property

        Private _thePart As Part = Nothing
        Public Property Part() As Part
            Get
                Return _thePart
            End Get
            Set(ByVal value As Part)
                lg.WriteLine("Set Property Part")
                _thePart = value
                _partUnits = _thePart.PartUnits
                Me.GetPartInfo()
                Me.GetDrawingSheets()
                If Me.SortSheetsByName Then
                    Me.SortDrawingSheets()
                End If
                lg.WriteLine("exiting Set Property Part")
                lg.WriteLine("")
            End Set
        End Property

        Private _partFilePath As String
        Public ReadOnly Property PartFilePath() As String
            Get
                Return _partFilePath
            End Get
        End Property

        Private _partNumber As String
        Public ReadOnly Property PartNumber() As String
            Get
                Return _partNumber
            End Get
        End Property

        Private _partRevision As String = ""
        Public ReadOnly Property PartRevision() As String
            Get
                Return _partRevision
            End Get
        End Property

        Private _preliminaryPrint As Boolean = False
        Public Property PreliminaryPrint() As Boolean
            Get
                Return _preliminaryPrint
            End Get
            Set(ByVal value As Boolean)
                lg.WriteLine("Set Property PreliminaryPrint")
                _preliminaryPrint = value
                If String.IsNullOrEmpty(_exportFile) Then
                    'do nothing
                Else
                    Me.GetOutputName()
                End If
                lg.WriteLine("  preliminaryPrint: " & _preliminaryPrint.ToString)
                lg.WriteLine("exiting Set Property PreliminaryPrint")
                lg.WriteLine("")
            End Set
        End Property

        Public ReadOnly Property SheetCount() As Integer
            Get
                Return _drawingSheets.Count
            End Get
        End Property

        Private _showConfirmationDialog As Boolean = False
        Public Property ShowConfirmationDialog() As Boolean
            Get
                Return _showConfirmationDialog
            End Get
            Set(ByVal value As Boolean)
                _showConfirmationDialog = value
            End Set
        End Property

        Private _skipBlankSheets As Boolean = True
        Public Property SkipBlankSheets() As Boolean
            Get
                Return _skipBlankSheets
            End Get
            Set(ByVal value As Boolean)
                _skipBlankSheets = value
            End Set
        End Property

        Private _sortSheetsByName As Boolean
        Public Property SortSheetsByName() As Boolean
            Get
                Return _sortSheetsByName
            End Get
            Set(ByVal value As Boolean)
                lg.WriteLine("Set Property SortSheetsByName")
                _sortSheetsByName = value
                If _sortSheetsByName = True Then
                    'sort alphabetically by sheet name
                    Me.SortDrawingSheets()
                Else
                    'get original collection order of sheets
                    Me.GetDrawingSheets()
                End If
                lg.WriteLine("  sortSheetsByName: " & _sortSheetsByName.ToString)
                lg.WriteLine("exiting Set Property SortSheetsByName")
                lg.WriteLine("")
            End Set
        End Property

        Private _textAsPolylines As Boolean = False
        Public Property TextAsPolylines() As Boolean
            Get
                Return _textAsPolylines
            End Get
            Set(ByVal value As Boolean)
                lg.WriteLine("Set Property TextAsPolylines")
                _textAsPolylines = value
                lg.WriteLine("  textAsPolylines: " & _textAsPolylines.ToString)
                lg.WriteLine("exiting Set Property TextAsPolylines")
                lg.WriteLine("")
            End Set
        End Property

        Private _useWatermark As Boolean = False
        Public Property UseWatermark() As Boolean
            Get
                Return _useWatermark
            End Get
            Set(ByVal value As Boolean)
                lg.WriteLine("Set Property UseWatermark")
                _useWatermark = value
                lg.WriteLine("  useWatermark: " & _useWatermark.ToString)
                lg.WriteLine("exiting Set Property UseWatermark")
                lg.WriteLine("")
            End Set
        End Property

        Private _watermarkAddDatestamp As Boolean = True
        Public Property WatermarkAddDatestamp() As Boolean
            Get
                Return _watermarkAddDatestamp
            End Get
            Set(ByVal value As Boolean)
                lg.WriteLine("Set Property WatermarkAddDatestamp")
                _watermarkAddDatestamp = value
                lg.WriteLine("  watermarkAddDatestamp: " & _watermarkAddDatestamp.ToString)
                If _watermarkAddDatestamp Then
                    'to do: internationalization for dates
                    _watermarkTextFinal = _watermarkText & " " & Today
                Else
                    _watermarkTextFinal = _watermarkText
                End If
                lg.WriteLine("  watermarkTextFinal: " & _watermarkTextFinal)
                lg.WriteLine("exiting Set Property WatermarkAddDatestamp")
                lg.WriteLine("")
            End Set
        End Property

        Private _watermarkText As String = "PRELIMINARY PRINT NOT TO BE USED FOR PRODUCTION"
        Public Property WatermarkText() As String
            Get
                Return _watermarkText
            End Get
            Set(ByVal value As String)
                lg.WriteLine("Set Property WatermarkText")
                _watermarkText = value
                lg.WriteLine("  watermarkText: " & _watermarkText)
                lg.WriteLine("exiting Set Property WatermarkText")
                lg.WriteLine("")
            End Set
        End Property


#End Region

#Region "public methods"

        Public Sub New()

            Me.StartLog()

        End Sub

        Public Sub New(ByVal thePart As NXOpen.Part)

            Me.StartLog()
            Me.Part = thePart

        End Sub

        Public Sub PickExportFolder()

            'Requires:
            '    Imports System.IO
            '    Imports System.Windows.Forms
            'if the user presses OK on the dialog box, the chosen path is returned
            'if the user presses cancel on the dialog box, 0 is returned
            lg.WriteLine("Sub PickExportFolder")

            If Me.ExportToTc Then
                lg.WriteLine("  N/A when ExportToTc = True")
                lg.WriteLine("  exiting Sub PickExportFolder")
                lg.WriteLine("")
                Return
            End If

            Dim strLastPath As String

            'Key will show up in HKEY_CURRENT_USER\Software\VB and VBA Program Settings
            Try
                'Get the last path used from the registry
                lg.WriteLine("  attempting to retrieve last export path from registry...")
                strLastPath = GetSetting("NX journal", "Export pdf", "ExportPath")
                'msgbox("Last Path: " & strLastPath)
            Catch e As ArgumentException
                lg.WriteLine("  ** Argument Exception: " & e.Message)
            Catch e As Exception
                lg.WriteLine("  ** Exception type: " & e.GetType.ToString)
                lg.WriteLine("  ** Exception message: " & e.Message)
                'MsgBox(e.GetType.ToString)
            Finally
            End Try

            Dim FolderBrowserDialog1 As New FolderBrowserDialog

            ' Then use the following code to create the Dialog window
            ' Change the .SelectedPath property to the default location
            With FolderBrowserDialog1
                ' Desktop is the root folder in the dialog.
                .RootFolder = Environment.SpecialFolder.Desktop
                ' Select the D:\home directory on entry.
                If Directory.Exists(strLastPath) Then
                    .SelectedPath = strLastPath
                Else
                    .SelectedPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                End If
                ' Prompt the user with a custom message.
                .Description = "Select the directory to export .pdf file"
                If .ShowDialog = DialogResult.OK Then
                    ' Display the selected folder if the user clicked on the OK button.
                    Me.OutputFolder = .SelectedPath
                    lg.WriteLine("  selected output path: " & .SelectedPath)
                    ' save the output folder path in the registry for use on next run
                    SaveSetting("NX journal", "Export pdf", "ExportPath", .SelectedPath)
                Else
                    'user pressed 'cancel', keep original value of output folder
                    _cancelOutput = True
                    Me.OutputFolder = Nothing
                    lg.WriteLine("  folder browser dialog cancel button pressed")
                    lg.WriteLine("  current output path: {nothing}")
                End If
            End With

            lg.WriteLine("exiting Sub PickExportFolder")
            lg.WriteLine("")

        End Sub

        Public Sub PromptPreliminaryPrint()

            lg.WriteLine("Sub PromptPreliminaryPrint")

            If _cancelOutput Then
                lg.WriteLine("  output canceled")
                Return
            End If

            Dim rspPreliminaryPrint As DialogResult
            rspPreliminaryPrint = MessageBox.Show("Add preliminary print watermark?", "Add Watermark?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If rspPreliminaryPrint = DialogResult.Yes Then
                Me.PreliminaryPrint = True
                Me.UseWatermark = True
                lg.WriteLine("  this is a preliminary print")
            Else
                Me.PreliminaryPrint = False
                lg.WriteLine("  this is not a preliminary print")
            End If

            lg.WriteLine("exiting Sub PromptPreliminaryPrint")
            lg.WriteLine("")

        End Sub

        Public Sub PromptWatermarkText()

            lg.WriteLine("Sub PromptWatermarkText")
            lg.WriteLine("  useWatermark: " & Me.UseWatermark.ToString)

            Dim theWatermarkText As String = Me.WatermarkText

            If Me.UseWatermark Then
                theWatermarkText = InputBox("Enter watermark text", "Watermark", theWatermarkText)
                Me.WatermarkText = theWatermarkText
            Else
                lg.WriteLine("  suppressing watermark prompt")
            End If

            lg.WriteLine("exiting Sub PromptWatermarkText")
            lg.WriteLine("")

        End Sub

        Public Sub SortDrawingSheets()

            If _cancelOutput Then
                Return
            End If

            If Not IsNothing(_thePart) Then
                Me.GetDrawingSheets()
                _drawingSheets.Sort(AddressOf Me.CompareSheetNames)
            End If

        End Sub

        Public Sub SortDrawingSheets(ByVal customSortFunction As System.Comparison(Of NXOpen.Drawings.DrawingSheet))

            If _cancelOutput Then
                Return
            End If

            If Not IsNothing(_thePart) Then
                Me.GetDrawingSheets()
                _drawingSheets.Sort(customSortFunction)
            End If

        End Sub

        Public Sub Commit()

            If _cancelOutput Then
                Return
            End If

            lg.WriteLine("Sub Commit")
            lg.WriteLine("  number of drawing sheets in part file: " & _drawingSheets.Count.ToString)

            _outputPdfFiles.Clear()
            For Each tempSheet As Drawings.DrawingSheet In _drawingSheets
                If Me.PreliminaryPrint Then
                    _outputPdfFiles.Add(IO.Path.Combine(Me.OutputFolder, tempSheet.Name & "_preliminary.pdf"))
                Else
                    _outputPdfFiles.Add(IO.Path.Combine(Me.OutputFolder, tempSheet.Name & ".pdf"))
                End If
            Next

            'make sure we can output to the specified file(s)
            If Me.ExportSheetsIndividually Then
                'check each sheet
                For Each newPdf As String In _outputPdfFiles

                    If Not Me.DeleteExistingPdfFile(newPdf) Then
                        If _overwritePdf Then
                            'file could not be deleted
                            MessageBox.Show("The pdf file: " & newPdf & " exists and could not be overwritten." & ControlChars.NewLine &
                                            "PDF export exiting", "PDF export error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            'file already exists and will not be overwritten
                            MessageBox.Show("The pdf file: " & newPdf & " exists and the overwrite option is set to False." & ControlChars.NewLine &
                                            "PDF export exiting", "PDF file exists", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        Return

                    End If

                Next

            Else
                'check _outputPdfFile
                If Not Me.DeleteExistingPdfFile(_outputPdfFile) Then
                    If _overwritePdf Then
                        'file could not be deleted
                        MessageBox.Show("The pdf file: " & _outputPdfFile & " exists and could not be overwritten." & ControlChars.NewLine &
                                        "PDF export exiting", "PDF export error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        'file already exists and will not be overwritten
                        MessageBox.Show("The pdf file: " & _outputPdfFile & " exists and the overwrite option is set to False." & ControlChars.NewLine &
                                        "PDF export exiting", "PDF file exists", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    Return
                End If

            End If

            Dim sheetCount As Integer = 0
            Dim sheetsExported As Integer = 0

            Dim numPlists As Integer = 0
            Dim myPlists() As Tag

            _theUfSession.Plist.AskTags(myPlists, numPlists)
            For i As Integer = 0 To numPlists - 1
                _theUfSession.Plist.Update(myPlists(i))
            Next

            For Each tempSheet As Drawings.DrawingSheet In _drawingSheets

                sheetCount += 1

                lg.WriteLine("  working on sheet: " & tempSheet.Name)
                lg.WriteLine("  sheetCount: " & sheetCount.ToString)

                'update any views that are out of date
                lg.WriteLine("  updating OutOfDate views on sheet: " & tempSheet.Name)
                Me.Part.DraftingViews.UpdateViews(Drawings.DraftingViewCollection.ViewUpdateOption.OutOfDate, tempSheet)

            Next

            If Me._drawingSheets.Count > 0 Then

                lg.WriteLine("  done updating views on all sheets")

                Try
                    If Me.ExportSheetsIndividually Then
                        For Each tempSheet As Drawings.DrawingSheet In _drawingSheets
                            lg.WriteLine("  calling Sub ExportPdf")
                            lg.WriteLine("")
                            If Me.PreliminaryPrint Then
                                Me.ExportPdf(tempSheet, IO.Path.Combine(Me.OutputFolder, tempSheet.Name & "_preliminary.pdf"))
                            Else
                                Me.ExportPdf(tempSheet, IO.Path.Combine(Me.OutputFolder, tempSheet.Name & ".pdf"))
                            End If
                        Next
                    Else
                        lg.WriteLine("  calling Sub ExportPdf")
                        lg.WriteLine("")
                        Me.ExportPdf()
                    End If
                Catch ex As Exception
                    lg.WriteLine("  ** error exporting PDF")
                    lg.WriteLine("  " & ex.GetType.ToString & " : " & ex.Message)
                    'MessageBox.Show("Error occurred in PDF export" & vbCrLf & ex.Message, "Error exporting PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Throw ex
                End Try

            Else
                'no sheets in file
                lg.WriteLine("  ** no drawing sheets in file: " & Me._partNumber)

            End If

            If Me.ShowConfirmationDialog Then
                Me.DisplayConfirmationDialog()
            End If

            If (Not Me.ExportToTc) AndAlso (Me.OpenPdf) AndAlso (Me._drawingSheets.Count > 0) Then
                'open new pdf print
                lg.WriteLine("  trying to open newly created pdf file")
                Try
                    If Me.ExportSheetsIndividually Then
                        For Each newPdf As String In _outputPdfFiles
                            System.Diagnostics.Process.Start(newPdf)
                        Next
                    Else
                        System.Diagnostics.Process.Start(Me.OutputPdfFileName)
                    End If
                    lg.WriteLine("  pdf open process successful")
                Catch ex As Exception
                    lg.WriteLine("  ** error opening pdf **")
                    lg.WriteLine("  " & ex.GetType.ToString & " : " & ex.Message)
                End Try
            End If

            lg.WriteLine("  exiting Sub ExportSheetsToPdf")
            lg.WriteLine("")

        End Sub

#End Region


#Region "private methods"

        Private Sub GetPartInfo()

            lg.WriteLine("Sub GetPartInfo")

            If Me.IsTCRunning Then
                _partNumber = _thePart.GetStringAttribute("DB_PART_NO")
                _partRevision = _thePart.GetStringAttribute("DB_PART_REV")

                lg.WriteLine("  TC running")
                lg.WriteLine("  partNumber: " & _partNumber)
                lg.WriteLine("  partRevision: " & _partRevision)

            Else 'running in native mode

                _partNumber = IO.Path.GetFileNameWithoutExtension(_thePart.FullPath)
                _partFilePath = IO.Directory.GetParent(_thePart.FullPath).ToString

                lg.WriteLine("  Native NX")
                lg.WriteLine("  partNumber: " & _partNumber)
                lg.WriteLine("  partFilePath: " & _partFilePath)

                Try
                    _partRevision = _thePart.GetStringAttribute("REVISION")
                    _partRevision = _partRevision.Trim
                Catch ex As Exception
                    _partRevision = ""
                End Try

                lg.WriteLine("  partRevision: " & _partRevision)

            End If

            If String.IsNullOrEmpty(_partRevision) Then
                _exportFile = _partNumber
            Else
                _exportFile = _partNumber & "_" & _partRevision
            End If

            lg.WriteLine("")
            Me.GetOutputName()

            lg.WriteLine("  exportFile: " & _exportFile)
            lg.WriteLine("  outputPdfFile: " & _outputPdfFile)
            lg.WriteLine("  exiting Sub GetPartInfo")
            lg.WriteLine("")

        End Sub

        Private Sub GetOutputName()

            lg.WriteLine("Sub GetOutputName")

            _exportFile.Replace("_preliminary", "")
            _exportFile.Replace("_PDF_1", "")

            If IsNothing(Me.Part) Then
                lg.WriteLine("  Me.Part is Nothing")
                lg.WriteLine("  exiting Sub GetOutputName")
                lg.WriteLine("")
                Return
            End If

            If Not IsTCRunning And _preliminaryPrint Then
                _exportFile &= "_preliminary"
            End If

            If Me.ExportToTc Then      'export to Teamcenter dataset
                lg.WriteLine("  export to TC option chosen")
                If Me.IsTCRunning Then
                    lg.WriteLine("  TC is running")
                    _exportFile &= "_PDF_1"
                Else
                    'error, cannot export to a dataset if TC is not running
                    lg.WriteLine("  ** error: export to TC option chosen, but TC is not running")
                    'todo: throw error
                End If
            Else                    'export to file system
                lg.WriteLine("  export to filesystem option chosen")
                If Me.IsTCRunning Then
                    lg.WriteLine("  TC is running")
                    'exporting from TC to filesystem, no part folder to default to
                    'default to "MyDocuments" folder
                    _outputPdfFile = IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, _exportFile & ".pdf")
                Else
                    lg.WriteLine("  native NX")
                    'exporting from native to file system
                    'use part folder as default output folder
                    If _outputFolder = "" Then
                        _outputFolder = _partFilePath
                    End If
                    _outputPdfFile = IO.Path.Combine(_outputFolder, _exportFile & ".pdf")
                End If

            End If

            lg.WriteLine("  exiting Sub GetOutputName")
            lg.WriteLine("")

        End Sub

        Private Sub GetDrawingSheets()

            _drawingSheets.Clear()

            For Each tempSheet As Drawings.DrawingSheet In _thePart.DrawingSheets
                If _skipBlankSheets Then
                    _theUfSession.Disp.SetDisplay(UFConstants.UF_DISP_SUPPRESS_DISPLAY)
                    Dim currentSheet As Drawings.DrawingSheet = _thePart.DrawingSheets.CurrentDrawingSheet
                    tempSheet.Open()
                    If Not IsSheetEmpty(tempSheet) Then
                        _drawingSheets.Add(tempSheet)
                    End If
                    Try
                        currentSheet.Open()
                    Catch ex As NXException
                        lg.WriteLine("  NX current sheet error: " & ex.Message)
                    Catch ex As Exception
                        lg.WriteLine("  current sheet error: " & ex.Message)
                    End Try
                    _theUfSession.Disp.SetDisplay(UFConstants.UF_DISP_UNSUPPRESS_DISPLAY)
                    _theUfSession.Disp.RegenerateDisplay()
                Else
                    _drawingSheets.Add(tempSheet)
                End If
            Next

        End Sub

        Private Function CompareSheetNames(ByVal x As Drawings.DrawingSheet, ByVal y As Drawings.DrawingSheet) As Integer

            'case-insensitive sort
            Dim myStringComp As StringComparer = StringComparer.CurrentCultureIgnoreCase

            'for a case-sensitive sort (A-Z then a-z), change the above option to:
            'Dim myStringComp As StringComparer = StringComparer.CurrentCulture

            Return myStringComp.Compare(x.Name, y.Name)

        End Function

        Private Function GetParentPath(ByVal thePath As String) As String

            lg.WriteLine("Function GetParentPath(" & thePath & ")")

            Try
                Dim directoryInfo As System.IO.DirectoryInfo
                directoryInfo = System.IO.Directory.GetParent(thePath)
                lg.WriteLine("  returning: " & directoryInfo.FullName)
                lg.WriteLine("exiting Function GetParentPath")
                lg.WriteLine("")

                Return directoryInfo.FullName
            Catch ex As ArgumentNullException
                lg.WriteLine("  Path is a null reference.")
                Throw ex
            Catch ex As ArgumentException
                lg.WriteLine("  Path is an empty string, contains only white space, or contains invalid characters")
                Throw ex
            End Try

            lg.WriteLine("exiting Function GetParentPath")
            lg.WriteLine("")

        End Function

        Private Sub ExportPdf()

            lg.WriteLine("Sub ExportPdf")

            Dim printPDFBuilder1 As PrintPDFBuilder

            printPDFBuilder1 = _thePart.PlotManager.CreatePrintPdfbuilder()
            printPDFBuilder1.Scale = 1.0
            printPDFBuilder1.Colors = PrintPDFBuilder.Color.BlackOnWhite
            printPDFBuilder1.Size = PrintPDFBuilder.SizeOption.ScaleFactor
            printPDFBuilder1.RasterImages = True
            printPDFBuilder1.ImageResolution = PrintPDFBuilder.ImageResolutionOption.Medium

            If _thePart.PartUnits = BasePart.Units.Inches Then
                lg.WriteLine("  part units: English")
                printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.English
            Else
                lg.WriteLine("  part units: Metric")
                printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.Metric
            End If

            If _textAsPolylines Then
                lg.WriteLine("  output text as polylines")
                printPDFBuilder1.OutputText = PrintPDFBuilder.OutputTextOption.Polylines
            Else
                lg.WriteLine("  output text as text")
                printPDFBuilder1.OutputText = PrintPDFBuilder.OutputTextOption.Text
            End If

            lg.WriteLine("  useWatermark: " & _useWatermark.ToString)
            If _useWatermark Then
                printPDFBuilder1.AddWatermark = True
                printPDFBuilder1.Watermark = _watermarkTextFinal
            Else
                printPDFBuilder1.AddWatermark = False
                printPDFBuilder1.Watermark = ""
            End If

            lg.WriteLine("  export to TC? " & _exportToTC.ToString)
            If _exportToTC Then
                'output to dataset
                printPDFBuilder1.Relation = PrintPDFBuilder.RelationOption.Manifestation
                printPDFBuilder1.DatasetType = "PDF"
                printPDFBuilder1.NamedReferenceType = "PDF_Reference"
                'printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.Overwrite
                printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.New
                printPDFBuilder1.DatasetName = _exportFile & "_PDF_1"
                lg.WriteLine("  dataset name: " & _exportFile)

                Try
                    lg.WriteLine("  printPDFBuilder1.Assign")
                    printPDFBuilder1.Assign()
                Catch ex As NXException
                    lg.WriteLine("  ** error with printPDFBuilder1.Assign")
                    lg.WriteLine("  " & ex.ErrorCode & " : " & ex.Message)
                End Try

            Else
                'output to filesystem
                lg.WriteLine("  pdf file: " & _outputPdfFile)
                printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.Native
                printPDFBuilder1.Append = False
                printPDFBuilder1.Filename = _outputPdfFile

            End If

            printPDFBuilder1.SourceBuilder.SetSheets(_drawingSheets.ToArray)

            Dim nXObject1 As NXObject
            Try
                lg.WriteLine("  printPDFBuilder1.Commit")
                nXObject1 = printPDFBuilder1.Commit()

            Catch ex As NXException
                lg.WriteLine("  ** error with printPDFBuilder1.Commit")
                lg.WriteLine("  " & ex.ErrorCode & " : " & ex.Message)

                'If Me.ExportToTc Then

                '    Try
                '        lg.WriteLine("  trying new dataset option")
                '        printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.New
                '        printPDFBuilder1.Commit()
                '    Catch ex2 As NXException
                '        lg.WriteLine("  ** error with printPDFBuilder1.Commit")
                '        lg.WriteLine("  " & ex2.ErrorCode & " : " & ex2.Message)

                '    End Try

                'End If

            Finally
                printPDFBuilder1.Destroy()
            End Try

            lg.WriteLine("  exiting Sub ExportPdf")
            lg.WriteLine("")

        End Sub

        Private Sub ExportPdf(ByVal theSheet As Drawings.DrawingSheet, ByVal pdfFile As String)

            lg.WriteLine("Sub ExportPdf(" & theSheet.Name & ", " & pdfFile & ")")

            Dim printPDFBuilder1 As PrintPDFBuilder

            printPDFBuilder1 = _thePart.PlotManager.CreatePrintPdfbuilder()
            printPDFBuilder1.Scale = 1.0
            printPDFBuilder1.Colors = PrintPDFBuilder.Color.BlackOnWhite
            printPDFBuilder1.Size = PrintPDFBuilder.SizeOption.ScaleFactor
            printPDFBuilder1.RasterImages = True
            printPDFBuilder1.ImageResolution = PrintPDFBuilder.ImageResolutionOption.Medium

            If _thePart.PartUnits = BasePart.Units.Inches Then
                lg.WriteLine("  part units: English")
                printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.English
            Else
                lg.WriteLine("  part units: Metric")
                printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.Metric
            End If

            If _textAsPolylines Then
                lg.WriteLine("  output text as polylines")
                printPDFBuilder1.OutputText = PrintPDFBuilder.OutputTextOption.Polylines
            Else
                lg.WriteLine("  output text as text")
                printPDFBuilder1.OutputText = PrintPDFBuilder.OutputTextOption.Text
            End If

            lg.WriteLine("  useWatermark: " & _useWatermark.ToString)
            If _useWatermark Then
                printPDFBuilder1.AddWatermark = True
                printPDFBuilder1.Watermark = _watermarkTextFinal
            Else
                printPDFBuilder1.AddWatermark = False
                printPDFBuilder1.Watermark = ""
            End If

            lg.WriteLine("  export to TC? " & _exportToTC.ToString)
            'If _exportToTC Then
            '    'output to dataset
            '    printPDFBuilder1.Relation = PrintPDFBuilder.RelationOption.Manifestation
            '    printPDFBuilder1.DatasetType = "PDF"
            '    printPDFBuilder1.NamedReferenceType = "PDF_Reference"
            '    'printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.Overwrite
            '    printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.New
            '    printPDFBuilder1.DatasetName = _exportFile & "_PDF_1"
            '    lg.WriteLine("  dataset name: " & _exportFile)

            '    Try
            '        lg.WriteLine("  printPDFBuilder1.Assign")
            '        printPDFBuilder1.Assign()
            '    Catch ex As NXException
            '        lg.WriteLine("  ** error with printPDFBuilder1.Assign")
            '        lg.WriteLine("  " & ex.ErrorCode & " : " & ex.Message)
            '    End Try

            'Else
            'output to filesystem
            lg.WriteLine("  pdf file: " & pdfFile)
            printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.Native
            printPDFBuilder1.Append = False
            printPDFBuilder1.Filename = pdfFile

            'End If

            Dim outputSheets(0) As Drawings.DrawingSheet
            outputSheets(0) = theSheet
            printPDFBuilder1.SourceBuilder.SetSheets(outputSheets)

            Dim nXObject1 As NXObject
            Try
                lg.WriteLine("  printPDFBuilder1.Commit")
                nXObject1 = printPDFBuilder1.Commit()

            Catch ex As NXException
                lg.WriteLine("  ** error with printPDFBuilder1.Commit")
                lg.WriteLine("  " & ex.ErrorCode & " : " & ex.Message)

                'If Me.ExportToTc Then

                '    Try
                '        lg.WriteLine("  trying new dataset option")
                '        printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.New
                '        printPDFBuilder1.Commit()
                '    Catch ex2 As NXException
                '        lg.WriteLine("  ** error with printPDFBuilder1.Commit")
                '        lg.WriteLine("  " & ex2.ErrorCode & " : " & ex2.Message)

                '    End Try

                'End If

            Finally
                printPDFBuilder1.Destroy()
            End Try

            lg.WriteLine("  exiting Sub ExportPdf")
            lg.WriteLine("")

        End Sub

        Private Sub DisplayConfirmationDialog()

            Dim sb As New System.Text.StringBuilder

            If Me._drawingSheets.Count = 0 Then
                MessageBox.Show("No drawing sheets found in file.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            sb.Append("The following sheets were output to PDF:")
            sb.AppendLine()
            For Each tempSheet As Drawings.DrawingSheet In _drawingSheets
                sb.AppendLine("   " & tempSheet.Name)
            Next
            sb.AppendLine()

            If Not Me.ExportToTc Then
                If Me.ExportSheetsIndividually Then
                    sb.AppendLine("Open pdf files now?")
                Else
                    sb.AppendLine("Open pdf file now?")
                End If
            End If

            Dim prompt As String = sb.ToString

            Dim response As DialogResult
            If Me.ExportToTc Then
                response = MessageBox.Show(prompt, Me.OutputPdfFileName, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                response = MessageBox.Show(prompt, Me.OutputPdfFileName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
            End If

            If response = DialogResult.Yes Then
                Me.OpenPdf = True
            Else
                Me.OpenPdf = False
            End If

        End Sub

        Private Sub StartLog()

            lg.WriteLine("")
            lg.WriteLine("~ NXJournaling.com: Start of drawing to PDF journal ~")
            lg.WriteLine("  ~~ Version: " & Version & " ~~")
            lg.WriteLine("  ~~ Timestamp of run: " & DateTime.Now.ToString & " ~~")
            lg.WriteLine("PdfExporter Sub StartLog()")

            'determine if we are running under TC or native
            _theUfSession.UF.IsUgmanagerActive(_isTCRunning)
            lg.WriteLine("IsTcRunning: " & _isTCRunning.ToString)

            lg.WriteLine("exiting Sub StartLog")
            lg.WriteLine("")


        End Sub

        Private Function DeleteExistingPdfFile(ByVal thePdfFile As String) As Boolean

            lg.WriteLine("Function DeleteExistingPdfFile(" & thePdfFile & ")")

            If File.Exists(thePdfFile) Then
                lg.WriteLine("  specified PDF file already exists")
                If Me.OverwritePdf Then
                    Try
                        lg.WriteLine("  user chose to overwrite existing PDF file")
                        File.Delete(thePdfFile)
                        lg.WriteLine("  file deleted")
                        lg.WriteLine("  returning: True")
                        lg.WriteLine("  exiting Function DeleteExistingPdfFile")
                        lg.WriteLine("")
                        Return True
                    Catch ex As Exception
                        'rethrow error?
                        lg.WriteLine("  ** error while attempting to delete existing pdf file")
                        lg.WriteLine("  " & ex.GetType.ToString & " : " & ex.Message)
                        lg.WriteLine("  returning: False")
                        lg.WriteLine("  exiting Function DeleteExistingPdfFile")
                        lg.WriteLine("")
                        Return False
                    End Try
                Else
                    'file exists, overwrite option is set to false - do nothing
                    lg.WriteLine("  specified pdf file exists, user chose not to overwrite")
                    lg.WriteLine("  returning: False")
                    lg.WriteLine("  exiting Function DeleteExistingPdfFile")
                    lg.WriteLine("")
                    Return False
                End If
            Else
                'file does not exist
                Return True
            End If

        End Function

        Private Function IsSheetEmpty(ByVal theSheet As Drawings.DrawingSheet) As Boolean

            theSheet.Open()
            Dim sheetTag As NXOpen.Tag = theSheet.View.Tag
            Dim sheetObj As NXOpen.Tag = NXOpen.Tag.Null
            _theUfSession.View.CycleObjects(sheetTag, UFView.CycleObjectsEnum.VisibleObjects, sheetObj)
            If (sheetObj = NXOpen.Tag.Null) And (theSheet.GetDraftingViews.Length = 0) Then
                Return True
            End If

            Return False

        End Function

#End Region

    End Class

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class IssueForm
        Inherits System.Windows.Forms.Form

        'Form overrides dispose to clean up the component list.
        <System.Diagnostics.DebuggerNonUserCode()>
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
        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
            Me.ComboBox1 = New System.Windows.Forms.ComboBox()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.FromBox = New System.Windows.Forms.TextBox()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.ToBox = New System.Windows.Forms.TextBox()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.SubjectBox = New System.Windows.Forms.TextBox()
            Me.Button3 = New System.Windows.Forms.Button()
            Me.Button6 = New System.Windows.Forms.Button()
            Me.Button7 = New System.Windows.Forms.Button()
            Me.Label5 = New System.Windows.Forms.Label()
            Me.Label6 = New System.Windows.Forms.Label()
            Me.Textbox4 = New System.Windows.Forms.WebBrowser()
            Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
            Me.FileSystemWatcher1 = New System.IO.FileSystemWatcher()
            Me.ListBox1 = New System.Windows.Forms.ListBox()
            Me.PartIssue = New System.Windows.Forms.RadioButton()
            Me.AssyIssue = New System.Windows.Forms.RadioButton()
            Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
            Me.Label7 = New System.Windows.Forms.Label()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.ListBox2 = New System.Windows.Forms.ListBox()
            Me.Label8 = New System.Windows.Forms.Label()
            CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'OpenFileDialog1
            '
            Me.OpenFileDialog1.FileName = "OpenFileDialog1"
            '
            'ComboBox1
            '
            Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.ComboBox1.FormattingEnabled = True
            Me.ComboBox1.Location = New System.Drawing.Point(81, 13)
            Me.ComboBox1.Name = "ComboBox1"
            Me.ComboBox1.Size = New System.Drawing.Size(232, 21)
            Me.ComboBox1.TabIndex = 1
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Location = New System.Drawing.Point(13, 89)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(33, 13)
            Me.Label1.TabIndex = 2
            Me.Label1.Text = "From:"
            '
            'FromBox
            '
            Me.FromBox.Location = New System.Drawing.Point(64, 86)
            Me.FromBox.Name = "FromBox"
            Me.FromBox.Size = New System.Drawing.Size(249, 20)
            Me.FromBox.TabIndex = 3
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Location = New System.Drawing.Point(13, 129)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(23, 13)
            Me.Label2.TabIndex = 4
            Me.Label2.Text = "To:"
            '
            'ToBox
            '
            Me.ToBox.Location = New System.Drawing.Point(64, 126)
            Me.ToBox.Name = "ToBox"
            Me.ToBox.ReadOnly = True
            Me.ToBox.Size = New System.Drawing.Size(249, 20)
            Me.ToBox.TabIndex = 5
            '
            'Label3
            '
            Me.Label3.AutoSize = True
            Me.Label3.Location = New System.Drawing.Point(13, 170)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(43, 13)
            Me.Label3.TabIndex = 6
            Me.Label3.Text = "Subject"
            '
            'SubjectBox
            '
            Me.SubjectBox.Location = New System.Drawing.Point(64, 167)
            Me.SubjectBox.Name = "SubjectBox"
            Me.SubjectBox.ReadOnly = True
            Me.SubjectBox.Size = New System.Drawing.Size(249, 20)
            Me.SubjectBox.TabIndex = 7
            '
            'Button3
            '
            Me.Button3.BackColor = System.Drawing.SystemColors.Highlight
            Me.Button3.Location = New System.Drawing.Point(530, 657)
            Me.Button3.Name = "Button3"
            Me.Button3.Size = New System.Drawing.Size(155, 56)
            Me.Button3.TabIndex = 13
            Me.Button3.Text = "Send Email and Issue"
            Me.Button3.UseVisualStyleBackColor = False
            '
            'Button6
            '
            Me.Button6.Location = New System.Drawing.Point(430, 657)
            Me.Button6.Name = "Button6"
            Me.Button6.Size = New System.Drawing.Size(94, 56)
            Me.Button6.TabIndex = 14
            Me.Button6.Text = "Help"
            Me.Button6.UseVisualStyleBackColor = True
            '
            'Button7
            '
            Me.Button7.Location = New System.Drawing.Point(15, 657)
            Me.Button7.Name = "Button7"
            Me.Button7.Size = New System.Drawing.Size(55, 56)
            Me.Button7.TabIndex = 15
            Me.Button7.Text = "Lock Only"
            Me.Button7.UseVisualStyleBackColor = True
            '
            'Label5
            '
            Me.Label5.AutoSize = True
            Me.Label5.Location = New System.Drawing.Point(12, 196)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(253, 13)
            Me.Label5.TabIndex = 17
            Me.Label5.Text = "Your personal email body (will be appended to email)"
            '
            'Label6
            '
            Me.Label6.AutoSize = True
            Me.Label6.Location = New System.Drawing.Point(13, 16)
            Me.Label6.Name = "Label6"
            Me.Label6.Size = New System.Drawing.Size(62, 13)
            Me.Label6.TabIndex = 19
            Me.Label6.Text = "Issue Type:"
            '
            'Textbox4
            '
            Me.Textbox4.Location = New System.Drawing.Point(12, 353)
            Me.Textbox4.MinimumSize = New System.Drawing.Size(20, 20)
            Me.Textbox4.Name = "Textbox4"
            Me.Textbox4.ScriptErrorsSuppressed = True
            Me.Textbox4.Size = New System.Drawing.Size(673, 295)
            Me.Textbox4.TabIndex = 26
            '
            'RichTextBox1
            '
            Me.RichTextBox1.AcceptsTab = True
            Me.RichTextBox1.AutoWordSelection = True
            Me.RichTextBox1.EnableAutoDragDrop = True
            Me.RichTextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.RichTextBox1.Location = New System.Drawing.Point(12, 212)
            Me.RichTextBox1.Name = "RichTextBox1"
            Me.RichTextBox1.Size = New System.Drawing.Size(673, 122)
            Me.RichTextBox1.TabIndex = 27
            Me.RichTextBox1.Text = ""
            '
            'FileSystemWatcher1
            '
            Me.FileSystemWatcher1.EnableRaisingEvents = True
            Me.FileSystemWatcher1.NotifyFilter = System.IO.NotifyFilters.LastWrite
            Me.FileSystemWatcher1.Path = "C:\eng\"
            Me.FileSystemWatcher1.SynchronizingObject = Me
            '
            'ListBox1
            '
            Me.ListBox1.AllowDrop = True
            Me.ListBox1.FormattingEnabled = True
            Me.ListBox1.Location = New System.Drawing.Point(334, 33)
            Me.ListBox1.Name = "ListBox1"
            Me.ListBox1.ScrollAlwaysVisible = True
            Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
            Me.ListBox1.Size = New System.Drawing.Size(351, 160)
            Me.ListBox1.TabIndex = 30
            '
            'PartIssue
            '
            Me.PartIssue.AutoSize = True
            Me.PartIssue.Location = New System.Drawing.Point(16, 52)
            Me.PartIssue.Name = "PartIssue"
            Me.PartIssue.Size = New System.Drawing.Size(123, 17)
            Me.PartIssue.TabIndex = 32
            Me.PartIssue.TabStop = True
            Me.PartIssue.Text = "Engineering Release"
            Me.PartIssue.UseVisualStyleBackColor = True
            '
            'AssyIssue
            '
            Me.AssyIssue.AutoSize = True
            Me.AssyIssue.Location = New System.Drawing.Point(215, 52)
            Me.AssyIssue.Name = "AssyIssue"
            Me.AssyIssue.Size = New System.Drawing.Size(98, 17)
            Me.AssyIssue.TabIndex = 33
            Me.AssyIssue.TabStop = True
            Me.AssyIssue.Text = "Repair Release"
            Me.AssyIssue.UseVisualStyleBackColor = True
            '
            'ProgressBar1
            '
            Me.ProgressBar1.Location = New System.Drawing.Point(79, 674)
            Me.ProgressBar1.Name = "ProgressBar1"
            Me.ProgressBar1.Size = New System.Drawing.Size(345, 23)
            Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
            Me.ProgressBar1.TabIndex = 34
            '
            'Label7
            '
            Me.Label7.AutoSize = True
            Me.Label7.Location = New System.Drawing.Point(12, 337)
            Me.Label7.Name = "Label7"
            Me.Label7.Size = New System.Drawing.Size(76, 13)
            Me.Label7.TabIndex = 35
            Me.Label7.Text = "Email Preview:"
            '
            'Label4
            '
            Me.Label4.AutoSize = True
            Me.Label4.Location = New System.Drawing.Point(331, 13)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(272, 13)
            Me.Label4.TabIndex = 36
            Me.Label4.Text = "Select the Drawings that you would like to send PDFs of"
            '
            'ListBox2
            '
            Me.ListBox2.FormattingEnabled = True
            Me.ListBox2.Location = New System.Drawing.Point(658, 199)
            Me.ListBox2.Name = "ListBox2"
            Me.ListBox2.Size = New System.Drawing.Size(12, 4)
            Me.ListBox2.TabIndex = 37
            Me.ListBox2.Visible = False
            '
            'Label8
            '
            Me.Label8.AutoSize = True
            Me.Label8.Location = New System.Drawing.Point(331, 196)
            Me.Label8.Name = "Label8"
            Me.Label8.Size = New System.Drawing.Size(310, 13)
            Me.Label8.TabIndex = 38
            Me.Label8.Text = "For spc dwg's, run this program directly from the parent assy dwg"
            '
            'IssueForm
            '
            Me.AllowDrop = True
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(697, 725)
            Me.Controls.Add(Me.Label8)
            Me.Controls.Add(Me.ListBox2)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.Label7)
            Me.Controls.Add(Me.ProgressBar1)
            Me.Controls.Add(Me.AssyIssue)
            Me.Controls.Add(Me.PartIssue)
            Me.Controls.Add(Me.ListBox1)
            Me.Controls.Add(Me.RichTextBox1)
            Me.Controls.Add(Me.Textbox4)
            Me.Controls.Add(Me.Label6)
            Me.Controls.Add(Me.Label5)
            Me.Controls.Add(Me.Button7)
            Me.Controls.Add(Me.Button6)
            Me.Controls.Add(Me.Button3)
            Me.Controls.Add(Me.SubjectBox)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me.ToBox)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.FromBox)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.ComboBox1)
            Me.Name = "IssueForm"
            Me.Text = "Issue Email Program"
            CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
        Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents FromBox As System.Windows.Forms.TextBox
        Friend WithEvents Label2 As System.Windows.Forms.Label
        Friend WithEvents ToBox As System.Windows.Forms.TextBox
        Friend WithEvents Label3 As System.Windows.Forms.Label
        Friend WithEvents SubjectBox As System.Windows.Forms.TextBox
        Friend WithEvents Button3 As System.Windows.Forms.Button
        Friend WithEvents Button6 As System.Windows.Forms.Button
        Friend WithEvents Button7 As System.Windows.Forms.Button
        Friend WithEvents Label5 As System.Windows.Forms.Label
        Friend WithEvents Label6 As System.Windows.Forms.Label
        Friend WithEvents Textbox4 As System.Windows.Forms.WebBrowser
        Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
        Friend WithEvents FileSystemWatcher1 As System.IO.FileSystemWatcher
        Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
        Friend WithEvents AssyIssue As System.Windows.Forms.RadioButton
        Friend WithEvents PartIssue As System.Windows.Forms.RadioButton
        Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
        Friend WithEvents Label7 As System.Windows.Forms.Label
        Friend WithEvents Label4 As System.Windows.Forms.Label
        Friend WithEvents ListBox2 As System.Windows.Forms.ListBox
        Friend WithEvents Label8 As System.Windows.Forms.Label
    End Class

End Module