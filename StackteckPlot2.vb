'-Update-JUL 17 2017- Raem Naveed - Fixed PlotterChoice, wasnt reseting itself before' 

Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.UF
Imports nxopenui
Imports System.Windows.Forms
Imports NXOpen.Assemblies
Imports System.IO
Imports System.Data.OleDb
Imports NXOpen.Drawings
Imports System.Diagnostics
Imports System.Threading
Imports System.String

Module StackteckPlotV2
    Dim s As Session = Session.GetSession()
    Dim theUI As UI = UI.GetUI()
    Dim ufs As UFSession = UFSession.GetUFSession()
    Dim lw As ListingWindow = s.ListingWindow

    Dim SpecPltDwg As String = ""
    Dim dryRun As Boolean = False
    Dim batchPlot As Boolean = False

    Dim plotterString As String = Nothing
    Dim plotterChoice As String = Nothing
    Dim plotterName As String = Nothing
    Dim numCopies As Integer = Nothing

    Sub Main()
        ' This is where the program begins 
        CreateUsageLog("StackteckPlotV2")
        lw.Open()

        Dim dp As Part = s.Parts.Display
        Dim workPart As Part = s.Parts.Work
        Dim undoMark As Session.UndoMarkId = s.SetUndoMark(Session.MarkVisibility.Visible, "Plot")

        DeleteUserPrintSetting()

        Dim plotForm As StackteckPlotForm = New StackteckPlotForm
        plotForm.ShowDialog()
    End Sub
    Public Sub CreateUsageLog(ByVal ProgramName As String)
        Dim username As String = System.Environment.UserName
        Dim UseDate As String = Now().Day & "-" & Now().Month & "-" & Now().Year

        Dim UsageLogFolderDir As String = "u: \logs\UG_Prog"

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
    Sub DeleteUserPrintSetting()
        Dim MyDocDir As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        Dim PrintSubmitDir1 As String = MyDocDir & "\print_submit"
        Dim PrintSubmitDir2 As String = "Y:\eng\" & Environment.UserName & "\print_submit"

        Dim PrintSubmitDir As String = Nothing
        If Directory.Exists(PrintSubmitDir1) = True Then
            PrintSubmitDir = PrintSubmitDir1
        End If

        If Directory.Exists(PrintSubmitDir2) = True Then
            PrintSubmitDir = PrintSubmitDir2
        End If

        lw.WriteLine("")
        lw.WriteLine("User Print Submit Dir: " & PrintSubmitDir)

        If Directory.Exists(PrintSubmitDir) = True Then
            Directory.Delete(PrintSubmitDir, True)
            lw.WriteLine("Deleting " & PrintSubmitDir & "......")
        End If
    End Sub
    Private Sub GetJobNum(ByRef JobNum As String)
        Dim dispPart As Part = s.Parts.Display()
        JobNum = s.Parts.Work.FullPath().Substring(0, 5)
    End Sub

    Public Class StackteckPlotForm

        Function GetFileFromPlotFolder(ByVal JobNum As String, ByRef CGMFilenames As String(), ByVal NumOfProdDwgsPlot As Integer) As Boolean
            lw.WriteLine("Num Of Prod Plot: " & NumOfProdDwgsPlot)

            Dim PlotFolderPath As String = "c:\eng\plots\" & JobNum
            lw.WriteLine(PlotFolderPath)

            Dim myFiles As String() = Nothing
            Dim AttemptCount As Integer = 0

            Dim i As Integer

            Do
                myFiles = Directory.GetFiles(PlotFolderPath)
                i = UBound(myFiles)
            Loop Until i >= NumOfProdDwgsPlot - 1

            Dim NumOfCGM As Integer = 0

            For Each myFile As String In myFiles
                lw.WriteLine(myFile)
                If myFile.EndsWith(".cgm") = True Then
                    NumOfCGM += 1
                    ReDim Preserve CGMFilenames(NumOfCGM - 1)
                    CGMFilenames(CGMFilenames.Length - 1) = myFile
                End If
            Next
            Return True
        End Function
        Sub STMPlot(ByRef JobNum As String)
            If CheckOutOfDateDwg() = False Then
                If MessageBox.Show("The drawing is still out of date after few update attempt.  Do you still want to Print it out?", "Out Of Date Drawing", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    lw.WriteLine("Plot Skip.....")
                    Exit Sub
                End If
            End If
            PlotDwgToCGM()
        End Sub
        Function STMIssuePlot(ByRef JobNum As String) As Boolean
            STMPlot(JobNum)
            Dim CGMFileNames As String() = Nothing
            GetFileFromPlotFolder(JobNum, CGMFileNames, 1)

            For Each CGMFile As String In CGMFileNames
                Dim NumOfCopies As Integer = 1
                Dim TempStr As String = CGMFile.Replace(".cgm", "")
                lw.WriteLine(TempStr)
                If TempStr.EndsWith("_B") Then
                    plotterChoice = 2
                    plotterName = "ENG HP5200"
                Else
                    plotterChoice = 3
                    plotterName = "KIP7700"
                End If

                PrintCGMToPlotter(CGMFile, plotterChoice, plotterName, NumOfCopies, dryRun)

                plotterChoice = 2
                plotterName = "ENG HP5200"
                PrintCGMToPlotter(CGMFile, plotterChoice, plotterName, NumOfCopies, dryRun)
            Next
        End Function
        Sub STMBatchPlot(ByVal JobNum As String, ByRef SkipDwgsPlot As String(), ByRef ProdDwgsPlot As String(), ByRef OutOfDateDwgsPlot As String())

            Dim SkipDwgsPlotCount As Integer = 0
            Dim ProdDwgsPlotCount As Integer = 0
            Dim OutOfDateDwgsPlotCount As Integer = 0

            Dim ComponentNames(0) As String
            Dim ComponentCounts As Integer = 0

            If MessageBox.Show("Is this """ & JobNum & """ your Job Number?", "Job Number", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                JobNum = NXInputBox.GetInputString("Enter the job Number")
            End If

            Dim Filter As String = JobNum

            GetAssemblyTree(ComponentNames, ComponentCounts, Filter)

            Dim Count As Integer = 0
            lw.WriteLine("Number of Job Specific Drawing: " & ComponentNames.Length)

            Dim i As Integer = 0
            For i = 0 To ComponentNames.Length - 1

                If i <> 0 Then
                    If OpenExistingPart(ComponentNames(i)) = False Then
                        Continue For
                    End If
                    Dim TempworkPart As Part = s.Parts.Work
                End If

                lw.WriteLine("")
                lw.WriteLine("  Open: " & ComponentNames(i))

                Dim PartNum As String = Nothing
                If CheckPartNumForIssuePlot(PartNum) = True Then

                    Dim UGPartName As String = Nothing
                    FindSpecDwgofMaster(UGPartName)

                    If UGPartName <> Nothing Then
                        OpenExistingSpecPart(ComponentNames(i), UGPartName)
                    End If

                    Dim mySheet As DrawingSheet = Nothing

                    If FoundProdDwg(mySheet) = True Then
                        mySheet.Open()
                        If CheckOutOfDateDwg() = True Then
                            PlotDwgToCGM()
                            ProdDwgsPlotCount += 1
                            ReDim Preserve ProdDwgsPlot(ProdDwgsPlotCount - 1)
                            ProdDwgsPlot(ProdDwgsPlot.Length - 1) = ComponentNames(i)
                        End If
                    Else
                        OutOfDateDwgsPlotCount += 1
                        ReDim Preserve OutOfDateDwgsPlot(OutOfDateDwgsPlotCount - 1)
                        OutOfDateDwgsPlot(OutOfDateDwgsPlot.Length - 1) = ComponentNames(i)
                    End If

                Else
                    SkipDwgsPlotCount += 1
                    ReDim Preserve SkipDwgsPlot(SkipDwgsPlotCount - 1)
                    SkipDwgsPlot(SkipDwgsPlot.Length - 1) = ComponentNames(i)
                End If
            Next

            OpenExistingPart(ComponentNames(0))
        End Sub



        Function CheckOutOfDateDwg() As Boolean

            Dim OutOfDateStatus
            Dim mysheet As DrawingSheet = s.Parts.Display.DrawingSheets.CurrentDrawingSheet

            Try
                OutOfDateStatus = mysheet.IsOutOfDate
            Catch Ex As Exception
            End Try

            If OutOfDateStatus = True Then
                lw.WriteLine("Drawing Out of Date: YES")
            Else
                lw.WriteLine("Drawing Status: NO")
            End If

            Dim NumOfAttempt As Integer = 0

            While mysheet.IsOutOfDate = True
                NumOfAttempt += 1
                Try
                    lw.WriteLine("  Updating Sheet: " & mysheet.Name)
                    lw.WriteLine("  Attempt " & NumOfAttempt & " ..........")
                    s.Parts.Work.DraftingViews.UpdateViews(DraftingViewCollection.ViewUpdateOption.All, mysheet)
                Catch ex As Exception
                    lw.WriteLine("Error updating sheet: " & ex.Message)
                    MessageBox.Show("Please fix your drawing before run this program.", "Update Fail", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                    Exit Function
                End Try

                If NumOfAttempt > 2 Then
                    lw.WriteLine("Drawing Cannot Update! Skip Plot")
                    Return False
                    Exit Function
                End If
            End While

            Return True
        End Function

        Public Sub PlotDwgToCGM()
            ChangeLineThicknesses()

            Dim mysheet As DrawingSheet = s.Parts.Display.DrawingSheets.CurrentDrawingSheet

            Dim PaperSize As String = Nothing
            AskPlotPaperSize(mysheet, PaperSize)

            Dim myPlotBuilder As PlotBuilder

            Dim workPart As Part = s.Parts.Work
            myPlotBuilder = workPart.PlotManager.CreatePlotBuilder()

            Dim sheet1(0) As NXObject
            sheet1(0) = mysheet
            myPlotBuilder.SourceBuilder.SetSheets(sheet1)

            SetPlotBuilderStd(myPlotBuilder)

            myPlotBuilder.PlotterText = "CGM"
            myPlotBuilder.ProfileText = "<System Profile>"

            Dim JobName As String = Nothing
            GetJobName(workPart, JobName)

            myPlotBuilder.JobName = JobName

            Dim PlotDirName As String = "c:\eng\plots\" & JobName.Remove(5, JobName.Length - 5)

            Dim filenames(0) As String
            filenames(0) = PlotDirName & "\" & JobName & ".cgm"
            myPlotBuilder.SetFilenames(filenames)

            Dim cgmfilenames(0) As String
            cgmfilenames(0) = PlotDirName & "\" & JobName & "_" & PaperSize & ".cgm"
            lw.WriteLine(cgmfilenames(0))
            myPlotBuilder.SetGraphicFilenames(cgmfilenames)

            'DisplayPlotBuilderSetting(myPlotBuilder)

            Dim nXObject1 As NXObject
            nXObject1 = myPlotBuilder.Commit()

            myPlotBuilder.Destroy()
        End Sub

        Private Sub SetPlotBuilderStd(ByRef myPlotBuilder As PlotBuilder)
            myPlotBuilder.Copies = 1
            myPlotBuilder.Tolerance = 0.001
            myPlotBuilder.RasterImages = True
            myPlotBuilder.XOffset = 0.8
            myPlotBuilder.YOffset = 0.5
            myPlotBuilder.CharacterSize = 0.09
            myPlotBuilder.JobName = "UGPlot"
            myPlotBuilder.DisplayBanner = True
            myPlotBuilder.ColorsWidthsBuilder.Colors = PlotColorsWidthsBuilder.Color.AsDisplayed
            myPlotBuilder.ColorsWidthsBuilder.Widths = PlotColorsWidthsBuilder.Width.CustomThreeWidths
        End Sub



        Public Sub GetJobName(ByVal myPart As Part, ByRef JobName As String)

            Dim PartName As String = Nothing
            Dim PartRev As String = Nothing
            Dim PartNum As String = Nothing

            find_part_attr_by_name(myPart, "DB_PART_NO", PartName)
            find_part_attr_by_name(myPart, "DB_PART_REV", PartRev)
            find_part_attr_by_name(myPart, "STACKTECK_PARTN", PartNum)

            If PartNum = "" Then
                JobName = PartName & "-" & PartRev
            Else
                JobName = PartName & "_" & PartNum & "-" & PartRev
            End If
        End Sub

        Public Sub find_part_attr_by_name(ByVal thePart As Part, ByVal attrName As String, ByRef attrVal As String)
            Dim theAttr As Attribute = Nothing
            Dim attr_info() As NXObject.AttributeInformation
            attr_info =
          thePart.GetAttributeTitlesByType(NXObject.AttributeType.String)

            Dim title As String = ""
            Dim value As String = ""
            Dim inx As Integer = 0
            Dim count As Integer = attr_info.Length()

            If attr_info.GetLowerBound(0) < 0 Then
                Return
            End If

            Do Until inx = count
                Dim result As Integer = 0

                title = attr_info(inx).Title.ToString
                result = String.Compare(attrName, title)

                If result = 0 Then
                    attrVal = thePart.GetStringAttribute(title)
                    Return
                End If
                inx += 1
            Loop
            Return
        End Sub

        Sub GetPlotterProfile(ByVal FileName As String, ByVal PlotterChoice As Integer, ByRef PlotterProfile As String)

            Dim TempStr As String = FileName
            TempStr = TempStr.Replace(".cgm", "")

            If TempStr.EndsWith("_B") Then
                Select Case PlotterChoice
                    Case 1
                        PlotterProfile = "B Size (11 X 17)"
                    Case 2
                        PlotterProfile = "B Size"
                    Case 3
                        PlotterProfile = "D Size"
                    Case 4
                        PlotterProfile = "B Size"
                End Select
            End If

            If TempStr.EndsWith("_C") Then
                Select Case PlotterChoice
                    Case 1
                        PlotterProfile = "B Size (11 X 17)"
                    Case 2
                        PlotterProfile = "B Size"
                    Case 3
                        PlotterProfile = "D Size"
                    Case 4
                        PlotterProfile = "C Size"
                End Select
            End If

            If TempStr.EndsWith("_D") Then
                Select Case PlotterChoice
                    Case 1
                        PlotterProfile = "B Size (11 X 17)"
                    Case 2
                        PlotterProfile = "B Size"
                    Case 3
                        PlotterProfile = "D Size"
                    Case 4
                        PlotterProfile = "D Size"
                End Select
            End If

            If TempStr.EndsWith("_E") Then
                Select Case PlotterChoice
                    Case 1
                        PlotterProfile = "B Size (11 X 17)"
                    Case 2
                        PlotterProfile = "B Size"
                    Case 3
                        PlotterProfile = "E Size"
                    Case 4
                        PlotterProfile = "E Size"
                End Select
            End If

            If TempStr.EndsWith("_E+") Then
                Select Case PlotterChoice
                    Case 1
                        PlotterProfile = "B Size (11 X 17)"

                    Case 2
                        PlotterProfile = "B Size"

                    Case 3
                        PlotterProfile = "E+ Size"

                    Case 4
                        PlotterProfile = "E+ Size"
                End Select
            End If
        End Sub
        Sub PrintCGMToPlotter(ByVal FileName As String, ByVal PlotterChoice As Integer, ByVal PlotterName As String, ByVal Copies As Integer, ByVal DryRun As Boolean)

            'lw.WriteLine("FILENAME: " + FileName)
            'lw.WriteLine("Last Character: " + FileName(FileName.Length - 5))

            If (FileName(FileName.Length - 5) = "_") Then
                MsgBox("Filename must end with paper size. Not a standard paper size, print this part manually!")
                Exit Sub
            End If

            Dim PlotterProfile As String = Nothing
            GetPlotterProfile(FileName, PlotterChoice, PlotterProfile)

            Dim GroupDir As String = "U:\nxplot"

            lw.WriteLine("")
            lw.WriteLine("Print CGM To Plotter......")
            lw.WriteLine("FileName: " & FileName)
            lw.WriteLine("PlotterName: " & PlotterName)
            lw.WriteLine("PlotterProfile: " & PlotterProfile)
            lw.WriteLine("No. of Copies: " & Copies)

            Dim myArg As String = "-input=""" & FileName & """ -group_dir=""" & GroupDir & """ -printer=""" & PlotterName & """ -profile=""" & PlotterProfile & """ -copies=" & Copies.ToString
            lw.WriteLine("My Arg = " & myArg)

            Dim UGPath As String = System.Environment.GetEnvironmentVariable("UGII_BASE_DIR")
            lw.WriteLine(UGPath)

            Dim p As ProcessStartInfo = New ProcessStartInfo
            p.FileName = UGPath & "\nxplot\nxplot.exe"
            p.Arguments = myArg

            If DryRun = False Then
                Dim myP As Process = New Process()
                myP.StartInfo() = p
                myP.Start()
                myP.WaitForExit()
            End If
        End Sub
        Sub PrintCGMToPDF(ByVal FileName As String, ByVal PlotterChoice As Integer, ByVal PlotterName As String, ByVal Copies As Integer, ByVal DryRun As Boolean)

            PrintCGMToPlotter(FileName, PlotterChoice, PlotterName, Copies, DryRun)

            Dim TempSplitStr As String() = FileName.Split("\")
            Dim PDFFileName As String = "\\stungsrv\UNGSRV-C\eng\out\" & TempSplitStr(TempSplitStr.Length - 1) & ".pdf"
            PDFFileName = PDFFileName.Replace(".cgm", "")
            lw.WriteLine("PDF File Location: " & PDFFileName)

            If DryRun = False Then
                If (CheckPDFFileCreated(PDFFileName) = False) Then
                    Exit Sub
                End If

                Dim FinalPDFFileName As String = Nothing
                GetFinalDestinationFileName(PDFFileName, FinalPDFFileName)

                File.Copy(PDFFileName, FinalPDFFileName, True)
                File.Delete(PDFFileName)
            End If
        End Sub
        Sub GetFinalDestinationFileName(ByVal PDFFileName As String, ByRef FinalPDFFileName As String)
            Dim UserName As String = System.Environment.UserName
            Dim UserHomeDrivePath As String = "Y:\Eng\" & UserName

            Dim PDFFilePath As String = "\\stungsrv\UNGSRV-C\eng\out"
            Dim GeneralPDFFilePath As String = "C:\ENG\UG_Translation"

            If Directory.Exists(GeneralPDFFilePath) = False Then
                Directory.CreateDirectory(GeneralPDFFilePath)
            End If

            FinalPDFFileName = PDFFileName

            FinalPDFFileName = FinalPDFFileName.Replace("_B.cgm.pdf", ".pdf")
            FinalPDFFileName = FinalPDFFileName.Replace("_D.cgm.pdf", ".pdf")
            FinalPDFFileName = FinalPDFFileName.Replace("_E.cgm.pdf", ".pdf")
            FinalPDFFileName = FinalPDFFileName.Replace("_E+.cgm.pdf", ".pdf")

            If Directory.Exists(UserHomeDrivePath) = True Then
                FinalPDFFileName = FinalPDFFileName.Replace(PDFFilePath & "\", UserHomeDrivePath & "\")
            Else
                FinalPDFFileName = FinalPDFFileName.Replace(PDFFilePath & "\", GeneralPDFFilePath & "\" & UserName & "_")
            End If

            lw.WriteLine("Final PDF File Name: " & FinalPDFFileName)
        End Sub
        Private Function CheckPDFFileCreated(ByVal FileName As String) As Boolean

            Dim AttemptCount As Integer = 0

            While File.Exists(FileName) = False

                Dim AltFileName As String = FileName.Replace(".pdf", ".cgm.pdf")

                If File.Exists(AltFileName) = True Then
                    File.Copy(AltFileName, FileName, True)
                    File.Delete(AltFileName)
                End If

                If AttemptCount = 4 Then
                    lw.WriteLine("Attempt Waiting more than 20 sec.  Program Terminated...")
                    Return False
                    Exit Function
                End If

                AttemptCount += 1
                lw.WriteLine("Waiting For File Created.  Attempt #" & AttemptCount & "......")
                Thread.Sleep(5000)
            End While

            Return True
        End Function
        Sub GetAssemblyTree(ByRef ComponentNames As String(), ByRef ComponentCounts As Integer, ByVal JobNum As String)

            Dim part1 As Part
            part1 = s.Parts.Work

            Dim c As Component = part1.ComponentAssembly.RootComponent
            lw.WriteLine(c.DisplayName)

            Dim count As Integer = 1
            Dim TempComponentName(0) As String
            TempComponentName(0) = c.DisplayName

            ShowAssemblyTree(c, "", count, TempComponentName)

            lw.WriteLine("Total Count: " & count)
            lw.WriteLine("")

            ComponentNames(0) = c.DisplayName
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
                        lw.WriteLine(Tempstr)
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
                    lw.WriteLine(newIndent & child.DisplayName & "is in suppress status")
                End If
            Next
        End Sub
        Function FoundProdDwg(ByRef mysheet As DrawingSheet) As Boolean
            Dim myDrawingSheets As DrawingSheet()
            Dim dp As Part = s.Parts.Display

            myDrawingSheets = dp.DrawingSheets.ToArray

            Dim FoundSheet1 As Boolean = False

            For Each mysheet In myDrawingSheets
                lw.Open()
                lw.WriteLine(mysheet.Name)

                If mysheet.Name = "SHEET1" Then
                    FoundSheet1 = True
                    Exit For
                End If
            Next


            If FoundSheet1 = False Then
                lw.WriteLine("Productoin drawing shoud name as: SHEET1.  Please rename your drawing.")
                Return False
            Else
                Return True
            End If
        End Function
        Function CheckPartNumForIssuePlot(ByRef PartNum As String) As Boolean
            If ExceptionForIssuePlot() = True Then
                Return True
                Exit Function
            End If

            Dim dispPart As Part = s.Parts.Display()
            find_part_attr_by_name(dispPart, "STACKTECK_PARTN", PartNum)
            lw.WriteLine("PART NUM: " & PartNum)

            Dim VisDwgNum As String = Nothing
            Dim VisPartDesc As String = Nothing
            Dim VisComCode As String = Nothing

            If ReadVisData("Y:\eng\ENG_ACCESS_DATABASES\VisibPartAttributes.mdb", PartNum, VisPartDesc, VisComCode, VisDwgNum) = True Then
                Select Case VisDwgNum
                    Case "X"
                        Return True
                    Case "XX"
                        Return True
                    Case Else
                        Return False
                End Select
            Else
                Return False
            End If
        End Function
        Function ExceptionForIssuePlot() As Boolean
            Dim dispPart As Part = s.Parts.Display()
            Dim PartName As String = Nothing
            find_part_attr_by_name(dispPart, "DB_PART_NO", PartName)

            Dim ExceptLst(2) As String
            ExceptLst(0) = "ASSY_STA"
            ExceptLst(1) = "ASSY_STACK"
            ExceptLst(2) = "ASSY"

            For Each ExceptStr As String In ExceptLst
                If PartName.EndsWith(ExceptStr) = True Then
                    Return True
                    Exit Function
                End If
            Next

            Return False

        End Function
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
            Try
                Dim part1 As Part = CType(s.Parts.FindObject("@DB/" & FileName), Part)

                Dim partLoadStatus2 As PartLoadStatus = Nothing
                Dim status1 As PartCollection.SdpsStatus
                status1 = s.Parts.SetDisplay(part1, True, True, partLoadStatus2)

                s.Parts.SetWork(part1)
            Catch ex As Exception
                msgbox(ex.ToString)
                Return False
            End Try
            Return True

        End Function
        Sub OpenExistingSpecPart(ByVal FileName As String, ByVal UGPartName As String)
            lw.WriteLine("Open Existing Spec Part Dwg ......")
            Dim basePart1 As BasePart
            Dim partLoadStatus1 As PartLoadStatus = Nothing

            lw.WriteLine("File Name: " & FileName)
            Dim Testarray As String() = FileName.Split("/")
            'lw.WriteLine("Test Array: " & Testarray.Length)
            If Testarray.Length > 2 Then
                MessageBox.Show("Part " & FileName & " will be skipped in this batch.  Please open this part and run this part individually.", "Invalid File Name", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If

            Try
                ' File already exists
                lw.WriteLine("@DB/" & FileName & "/specification/" & UGPartName)
                basePart1 = s.Parts.OpenBaseDisplay("@DB/" & FileName & "/specification/" & UGPartName, partLoadStatus1)
            Catch ex As NXException
                'ex.AssertErrorCode(1020004)
            End Try

            Dim markId3 As Session.UndoMarkId
            markId3 = s.SetUndoMark(Session.MarkVisibility.Visible, "Change Display Part")
            Try
                Dim part1 As Part = CType(s.Parts.FindObject("@DB/" & FileName & "/specification/" & UGPartName), Part)

                Dim partLoadStatus2 As PartLoadStatus = Nothing
                Dim status1 As PartCollection.SdpsStatus
                status1 = s.Parts.SetDisplay(part1, True, True, partLoadStatus2)

                s.Parts.SetWork(part1)
            Catch ex As Exception
                msgbox("Spec Didnt open " + System.Environment.NewLine + ex.ToString)
            End Try

        End Sub
        Sub ListSkipDwgsPlot(ByVal SkipDwgsPlot As String())
            lw.WriteLine("")
            lw.WriteLine("############################")
            lw.WriteLine("    Dwgs Skip for Plot      ")
            lw.WriteLine("+###########################")

            If SkipDwgsPlot Is Nothing Then
                lw.WriteLine("Total File: 0")
                Exit Sub
            End If

            For Each Tempstr As String In SkipDwgsPlot
                lw.WriteLine(Tempstr)
            Next

            lw.WriteLine("Total File: " & SkipDwgsPlot.Length)

        End Sub
        Sub ListOutOfDateDwgsPlot(ByVal OutOfDateDwgsPlot As String())
            lw.WriteLine("")
            lw.WriteLine("###################################")
            lw.WriteLine("    Out Of Date Dwgs for Plot      ")
            lw.WriteLine("+##################################")

            If OutOfDateDwgsPlot Is Nothing Then
                lw.WriteLine("Total File: 0")
                Exit Sub
            End If

            For Each Tempstr As String In OutOfDateDwgsPlot
                lw.WriteLine(Tempstr)
            Next

            lw.WriteLine("Total File: " & OutOfDateDwgsPlot.Length)

        End Sub
        Sub PrintLogToTxt(ByVal JobNum As String, ByVal ProdDwgsPlot As String(), ByVal OutOfDateDwgsPlot As String(), ByVal SkipDwgsPlot As String())

            ListProdDwgsPlot(ProdDwgsPlot)
            ListOutOfDateDwgsPlot(OutOfDateDwgsPlot)
            ListSkipDwgsPlot(SkipDwgsPlot)

            Dim PlotDir As String = "c:\eng\plots\" & JobNum
            Dim myLogPath As String = PlotDir & "\" & JobNum & ".txt"

            Dim myWriter As StreamWriter = New StreamWriter(myLogPath)

            myWriter.WriteLine("Print it out for your reference if is necessary!!!!!")
            myWriter.WriteLine("")
            myWriter.WriteLine("############################################")
            myWriter.WriteLine("    Out Of Date Dwgs for Plot (Skipped)     ")
            myWriter.WriteLine("+###########################################")

            If OutOfDateDwgsPlot Is Nothing Then
                myWriter.WriteLine("Total File: 0")
            Else
                For Each Tempstr As String In OutOfDateDwgsPlot
                    myWriter.WriteLine(Tempstr)
                Next

                myWriter.WriteLine("Total File: " & OutOfDateDwgsPlot.Length)
            End If

            myWriter.WriteLine("")
            myWriter.WriteLine("")
            myWriter.WriteLine("############################")
            myWriter.WriteLine("    Prod Dwgs for Plot      ")
            myWriter.WriteLine("+###########################")

            If ProdDwgsPlot Is Nothing Then
                myWriter.WriteLine("Total File: 0")
            Else
                For Each Tempstr As String In ProdDwgsPlot
                    myWriter.WriteLine(Tempstr)
                Next

                myWriter.WriteLine("Total File for Plot: " & ProdDwgsPlot.Length)
            End If

            myWriter.WriteLine("")
            myWriter.WriteLine("")
            myWriter.WriteLine("############################")
            myWriter.WriteLine("    Dwgs Skip for Plot      ")
            myWriter.WriteLine("+###########################")

            If SkipDwgsPlot Is Nothing Then
                myWriter.WriteLine("Total File: 0")
            Else
                For Each Tempstr As String In SkipDwgsPlot
                    myWriter.WriteLine(Tempstr)
                Next

                myWriter.WriteLine("Total File: " & SkipDwgsPlot.Length)
            End If

            myWriter.Close()
            myWriter.Dispose()

            Dim p As ProcessStartInfo = New ProcessStartInfo()
            p.FileName = "notepad.exe"
            p.Arguments = myLogPath

            Dim myProcess As Process = New Process()
            myProcess.StartInfo = p
            myProcess.Start()

        End Sub
        Sub ListProdDwgsPlot(ByVal ProdDwgsPlot As String())
            lw.WriteLine("")
            lw.WriteLine("############################")
            lw.WriteLine("    Prod Dwgs for Plot      ")
            lw.WriteLine("+###########################")

            If ProdDwgsPlot Is Nothing Then
                lw.WriteLine("Total File: 0")
                Exit Sub
            End If


            For Each Tempstr As String In ProdDwgsPlot
                lw.WriteLine(Tempstr)
            Next

            lw.WriteLine("Total File for Plot: " & ProdDwgsPlot.Length)
        End Sub
        Public Function ReadVisData(ByVal FileName As String, ByVal PartNum As String, ByRef VisPartDesc As String, ByRef VisComCode As String, ByRef VisDwgNum As String) As Boolean

            'Define the connectors
            Dim cn As OleDbConnection
            Dim cmd As OleDbCommand
            Dim dr As OleDbDataReader
            Dim oConnect, oQuery As String
            Dim FoundStatus As Boolean = False

            'Define connection string
            'oConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName
            oConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName


            'Query String
            oQuery = "SELECT * FROM VISIB_PARTMASTER_LOCAL where PARTNO='" & PartNum & "'"

            'Instantiate the connectors
            cn = New OleDbConnection(oConnect)
            cn.Open()

            cmd = New OleDbCommand(oQuery, cn)
            dr = cmd.ExecuteReader

            While dr.Read()
                VisPartDesc = dr("PARTDESCR")
                VisPartDesc = VisPartDesc.Trim()

                VisComCode = dr("COMMODITY_CODE")
                VisComCode = VisComCode.Trim()

                VisDwgNum = dr("DRAWINGNO")
                VisDwgNum = VisDwgNum.Trim()

                FoundStatus = True
            End While

            dr.Close()
            cn.Close()

            If FoundStatus = False Then
                lw.WriteLine("PartNum not Found")
            End If

            Return FoundStatus
        End Function
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
                        'MsgBox(UGPartNames(j))
                        Select Case UGPartNames(j)
                            Case "dwg"
                                UGPartName = UGPartNames(j)
                                Exit For
                            Case "dwg_H"
                                If SpecPltDwg = "" Then
                                    Dim result As Integer = MessageBox.Show("Is this a horizontal mold?", "Horizontal Mold?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                    If result = DialogResult.Yes Then
                                        SpecPltDwg = "H"
                                    Else
                                        SpecPltDwg = "V"
                                    End If
                                End If

                                If SpecPltDwg = "H" Then
                                    UGPartName = "dwg_H"
                                ElseIf (SpecPltDwg = "V") Then
                                    UGPartName = "dwg_V"
                                End If
                                'Exit For
                            Case "dwg_V"
                                If SpecPltDwg = "" Then
                                    Dim result As Integer = MessageBox.Show("Is this a Vertical mold?", "Vertical Mold?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                                    If result = DialogResult.Yes Then
                                        SpecPltDwg = "V"
                                    Else
                                        SpecPltDwg = "H"
                                    End If
                                End If

                                If SpecPltDwg = "H" Then
                                    UGPartName = "dwg_H"
                                ElseIf (SpecPltDwg = "V") Then
                                    UGPartName = "dwg_V"
                                End If
                                Exit For
                        End Select
                    Next

                    lw.WriteLine(" UGPart Name: " & UGPartName)
                    Exit Sub
                End If
            Next
        End Sub
        Function ConfirmProdDwgsPlot(ByVal JobNum As String, ByRef CGMFiles As String(), ByRef SkipDwgsPlot As String(), ByRef ProdDwgsPlot As String()) As Boolean

            Dim PlotDir As String = "c:\eng\plots\" & JobNum

            Dim OriginalCGMs As String() = Nothing
            Dim i As Integer = -1

            For Each CGMFile As String In CGMFiles
                i += 1
                ReDim Preserve OriginalCGMs(i)
                OriginalCGMs(OriginalCGMs.Length - 1) = CGMFile.Replace(PlotDir & "\", "")
                OriginalCGMs(OriginalCGMs.Length - 1) = OriginalCGMs(OriginalCGMs.Length - 1).Replace(".cgm", "")
            Next

            Dim myPlotLst As FormProdDwgsPlot = New FormProdDwgsPlot

            For Each OriginalCGM As String In OriginalCGMs
                myPlotLst.lstCGMFiles.Items.Add(OriginalCGM)
            Next

            Dim RemoveCGMFiles As String() = Nothing
            Dim FinalCGMFiles As String() = Nothing

            If myPlotLst.ShowDialog() = DialogResult.OK Then
                For Each TempCGMFile As String In myPlotLst.lstRemoveCGMFiles.Items

                    For Each OriginalCGM As String In CGMFiles

                        If OriginalCGM.Contains(TempCGMFile) = True Then
                            If RemoveCGMFiles Is Nothing Then
                                ReDim Preserve RemoveCGMFiles(0)
                            Else
                                ReDim Preserve RemoveCGMFiles(RemoveCGMFiles.Length)
                            End If

                            RemoveCGMFiles(RemoveCGMFiles.Length - 1) = OriginalCGM

                            If SkipDwgsPlot Is Nothing Then
                                ReDim Preserve SkipDwgsPlot(0)
                            Else
                                ReDim Preserve SkipDwgsPlot(SkipDwgsPlot.Length)
                            End If
                            SkipDwgsPlot(SkipDwgsPlot.Length - 1) = OriginalCGM

                            Exit For
                        End If
                    Next
                Next

                For Each TempCGMFile As String In myPlotLst.lstCGMFiles.Items
                    For Each OriginalCGM As String In CGMFiles
                        If OriginalCGM.Contains(TempCGMFile) = True Then
                            If OriginalCGM.Contains(TempCGMFile) = True Then
                                If FinalCGMFiles Is Nothing Then
                                    ReDim Preserve FinalCGMFiles(0)
                                Else
                                    ReDim Preserve FinalCGMFiles(FinalCGMFiles.Length)
                                End If

                                FinalCGMFiles(FinalCGMFiles.Length - 1) = OriginalCGM
                                Exit For
                            End If
                        End If
                    Next
                Next
            Else
                Return False
                Exit Function
            End If

            lw.WriteLine(FinalCGMFiles.Length)

            ReDim CGMFiles(FinalCGMFiles.Length - 1)
            ReDim ProdDwgsPlot(FinalCGMFiles.Length - 1)

            For j As Integer = 0 To FinalCGMFiles.Length - 1
                CGMFiles(j) = FinalCGMFiles(j)
                ProdDwgsPlot(j) = FinalCGMFiles(j)
            Next

            Return True
        End Function

        Public Sub ChangeLineThicknesses()
            Dim theSession As Session = Session.GetSession()
            Dim workPart As Part = theSession.Parts.Work
            Dim displayPart As Part = theSession.Parts.Display

            Dim tempObject As Object = Nothing

            For Each dimension As Annotations.Dimension In workPart.Dimensions
                Dim editSettingsBuilder As Annotations.EditSettingsBuilder
                Dim objects(0) As DisplayableObject
                objects(0) = dimension
                editSettingsBuilder = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects)
                Dim fontIndex1 As Integer
                fontIndex1 = workPart.Fonts.AddFont("blockslant", FontCollection.Type.Nx)
                editSettingsBuilder.AnnotationStyle.LetteringStyle.DimensionTextLineWidth = Annotations.LineWidth.Four
                editSettingsBuilder.AnnotationStyle.LetteringStyle.AppendedTextLineWidth = Annotations.LineWidth.Four
                editSettingsBuilder.AnnotationStyle.LetteringStyle.ToleranceTextLineWidth = Annotations.LineWidth.Four
                Dim nXObject1 As NXObject
                nXObject1 = editSettingsBuilder.Commit()
            Next

            For Each note As Annotations.Label In workPart.Labels
                Dim editSettingsBuilder As Annotations.EditSettingsBuilder
                Dim objects(0) As DisplayableObject
                objects(0) = note
                editSettingsBuilder = workPart.SettingsManager.CreateAnnotationEditSettingsBuilder(objects)
                editSettingsBuilder.AnnotationStyle.LetteringStyle.GeneralTextLineWidth = Annotations.LineWidth.Four
                Dim nxObject2 As NXObject
                nxObject2 = editSettingsBuilder.Commit()
            Next

            For Each view As Object In workPart.Views.GetActiveViews

                Dim editViewSettingsBuilder As Drawings.EditViewSettingsBuilder
                Dim viewArray(0) As NXOpen.View
                viewArray(0) = view

                Try
                    editViewSettingsBuilder = workPart.SettingsManager.CreateDrawingEditViewSettingsBuilder(viewArray)

                    Dim editSettingsBuilder As Drafting.BaseEditSettingsBuilder
                    editSettingsBuilder = editViewSettingsBuilder

                    editViewSettingsBuilder.ViewStyle.ViewStyleVisibleLines.VisibleWidth = Preferences.Width.Five
                    editViewSettingsBuilder.ViewStyle.ViewStyleHiddenLines.Width = Preferences.Width.Three

                    Dim nxObject1 As NXObject

                    nxObject1 = editViewSettingsBuilder.Commit()
                Catch ex As Exception
                    'MsgBox(ex.ToString)
                End Try
            Next
        End Sub



        Function STMBatchIssuePlot(ByVal JobNum As String) As Boolean

            Dim SkipDwgsPlot() As String = Nothing
            Dim ProdDwgsPlot() As String = Nothing
            Dim OutOfDateDwgsPlot() As String = Nothing
            If chkBoxExistingJobList.checked = False Then
                STMBatchPlot(JobNum, SkipDwgsPlot, ProdDwgsPlot, OutOfDateDwgsPlot)
            Else
                Dim PlotDir As String = "C:\eng\plots\" & JobNum
                Dim TempFilesLst As String() = Directory.GetFiles(PlotDir)
                For Each TempFile As String In TempFilesLst
                    If TempFile.EndsWith(".cgm") = True Then
                        If ProdDwgsPlot Is Nothing Then
                            ReDim Preserve ProdDwgsPlot(0)
                        Else
                            ReDim Preserve ProdDwgsPlot(ProdDwgsPlot.Length)
                        End If

                        ProdDwgsPlot(ProdDwgsPlot.Length - 1) = TempFile
                    End If
                Next
            End If



            lw.WriteLine(ProdDwgsPlot.Length)

            Dim CGMFileNames As String() = Nothing
            If GetFileFromPlotFolder(JobNum, CGMFileNames, ProdDwgsPlot.Length) = False Then
                lw.WriteLine("Fail getting file from plot folder")
                Exit Function
            End If

            If ConfirmProdDwgsPlot(JobNum, CGMFileNames, SkipDwgsPlot, ProdDwgsPlot) = False Then
                Exit Function
            End If

            Dim SmalldwgPlotter As Integer = 2
            Dim BigDwgPlotter As Integer = 3

            For Each CGMFile As String In CGMFileNames

                Dim NumOfCopies As Integer = 1

                Dim PlotSize As String = CGMFile.Replace(".cgm", "")
                lw.WriteLine(PlotSize)

                If PlotSize.EndsWith("_B") = True Then
                    plotterChoice = SmalldwgPlotter
                    plotterName = "ENG HP5200"
                Else
                    plotterChoice = BigDwgPlotter
                    plotterName = "KIP7700"
                End If

                PrintCGMToPlotter(CGMFile, plotterChoice, plotterName, NumOfCopies, dryRun)

                plotterChoice = 2
                plotterName = "ENG HP5200"
                PrintCGMToPlotter(CGMFile, plotterChoice, plotterName, NumOfCopies, dryRun)
            Next

            PrintLogToTxt(JobNum, ProdDwgsPlot, OutOfDateDwgsPlot, SkipDwgsPlot)
            Return True
        End Function

        Private Sub CheckPlotFolder(ByVal JobNum As String)
            Dim PlotDirName As String = "c:\eng\plots\" & JobNum

            If Directory.Exists(PlotDirName) = True Then
                If chkBoxExistingJobList.Checked Then
                    Exit Sub
                End If
                System.IO.Directory.Delete(PlotDirName, True)
            End If

            Directory.CreateDirectory(PlotDirName)
        End Sub
        Public Function AskPlotPaperSize(ByVal mysheet As DrawingSheet, ByRef PaperSize As String) As Boolean
            If mysheet.Height = 12 And mysheet.Length = 18 Then
                PaperSize = "B"
                TxtboxPapersize.text = PaperSize
                Return True
                Exit Function
            End If

            If mysheet.Height = 18 And mysheet.Length = 24 Then
                PaperSize = "C"
                TxtboxPapersize.text = PaperSize
                Return True
                Exit Function
            End If

            If mysheet.Height = 24 And mysheet.Length = 36 Then
                PaperSize = "D"
                TxtboxPapersize.text = PaperSize
                Return True
                Exit Function
            End If

            If mysheet.Height = 36 And mysheet.Length = 48 Then
                PaperSize = "E"
                TxtboxPapersize.text = PaperSize
                Return True
                Exit Function
            End If

            If mysheet.Height = 36 And mysheet.Length > 48 Then
                PaperSize = "E+"
                TxtboxPapersize.text = PaperSize
                Return True
                Exit Function
            End If
            Return False
        End Function

        Private Sub StackteckPlotForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            txtBoxCopies.Text = "1"
            chkBoxIssuePlot.Checked = True
            Dim mysheet As DrawingSheet = s.Parts.Display.DrawingSheets.CurrentDrawingSheet
            Dim PaperSize As String = Nothing
            Try
                AskPlotPaperSize(mysheet, PaperSize)
            Catch ex As Exception
                'MsgBox("You need to run this from the drawing")
            End Try

            ChangeLineThicknesses()
        End Sub

        Private Sub chkBoxNormalPlot_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxNormalPlot.CheckedChanged
            If (chkBoxNormalPlot.Checked = True) Then
                chkBoxIssuePlot.Checked = False
                chkBoxPDF.Checked = False
                chkBoxSamsungPrintRoom.Enabled = True
                chkBoxKIP7700.Enabled = True
                chkBoxEng2220.Enabled = True
                chkBoxEng2220.Checked = True
                txtBoxCopies.Enabled = True
                txtBoxCopies.Text = "1"
            End If
        End Sub

        Private Sub chkBoxIssuePlot_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxIssuePlot.CheckedChanged
            If (chkBoxIssuePlot.Checked = True) Then
                chkBoxNormalPlot.Checked = False
                chkBoxPDF.Checked = False
                chkBoxKIP7700.Checked = False
                chkBoxEng2220.Checked = False
                chkBoxSamsungPrintRoom.Checked = False
                chkBoxSamsungPrintRoom.Enabled = False
                chkBoxEng2220.Enabled = False
                chkBoxKIP7700.Enabled = False
                txtBoxCopies.Enabled = False
                txtBoxCopies.Text = ""
            End If
        End Sub
        Private Sub chkBoxPDF_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxPDF.CheckedChanged
            If (chkBoxPDF.Checked = True) Then
                chkBoxNormalPlot.Checked = False
                chkBoxIssuePlot.Checked = False
                chkBoxKIP7700.Checked = False
                chkBoxKIP7700.Enabled = False
                chkBoxSamsungPrintRoom.Checked = False
                chkBoxSamsungPrintRoom.Enabled = False
                chkBoxEng2220.Checked = False
                chkBoxEng2220.Enabled = False
                chkBoxDryRun.Enabled = False
                chkBoxDryRun.Checked = False
                txtBoxCopies.Enabled = False
                txtBoxCopies.Text = "1"
                plotterName = "PDF"
            ElseIf (chkBoxPDF.Checked = False) Then
                chkBoxKIP7700.Enabled = True
                chkBoxSamsungPrintRoom.Enabled = True
                chkBoxEng2220.Enabled = True
                chkBoxDryRun.Enabled = True
                txtBoxCopies.Enabled = True
                plotterChoice = 0
                plotterName = ""
            End If
        End Sub
        Private Sub chkBoxEng2220_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxEng2220.CheckedChanged
            If (chkBoxEng2220.Checked = True) Then
                chkBoxKIP7700.Checked = False
                chkBoxSamsungPrintRoom.Checked = False
                plotterString = "Eng2220"
                plotterChoice = 1
                plotterName = "Canon iR2220 ENG"
            Else
                plotterString = ""
                plotterChoice = 0
                plotterName = ""
            End If
        End Sub
        Private Sub chkBoxSamsungPrintRoom_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxSamsungPrintRoom.CheckedChanged
            If (chkBoxSamsungPrintRoom.Checked = True) Then
                chkBoxEng2220.Checked = False
                chkBoxKIP7700.Checked = False
                plotterString = "SamsungPrintRoom"
                plotterChoice = 2
                plotterName = "ENG HP5200"
            Else
                plotterString = ""
                plotterChoice = 0
                plotterName = ""
            End If
        End Sub
        Private Sub chkBoxKIP7700_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxKIP7700.CheckedChanged
            If (chkBoxKIP7700.Checked = True) Then
                chkBoxSamsungPrintRoom.Checked = False
                chkBoxEng2220.Checked = False
                plotterString = "KIP7700"
                plotterChoice = 3
                plotterName = "KIP7700"
            Else
                plotterString = ""
                plotterChoice = 0
                plotterName = ""
            End If
        End Sub
        Private Sub chkBoxBatch_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxBatch.CheckedChanged
            If (chkBoxBatch.Checked = True) Then
                batchPlot = True
                chkBoxExistingJobList.visible = True
            Else
                batchPlot = False
                chkBoxExistingJobList.visible = False
                chkBoxExistingJobList.checked = False
            End If
        End Sub
        Private Sub chkBoxDryRun_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxDryRun.CheckedChanged
            If (chkBoxDryRun.Checked = True) Then
                dryRun = True
            Else
                dryRun = False
            End If
        End Sub
        Private Sub cancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelButton.Click
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
            Me.Dispose()
        End Sub
        Private Sub helpButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles helpButton.Click
            Dim p As System.Diagnostics.Process = New System.Diagnostics.Process()
            Dim path As String = "\\cntfiler\_JobLib\EngProcedures\UG\ENGA75T - Stackteck Plot Program in UG.pdf"

            If System.IO.File.Exists(path) = True Then
                p.Start(path)
            Else
                System.Windows.Forms.MessageBox.Show("File does not exist. Contact administrator.")
            End If
        End Sub
        Private Sub GetPlotMethod(ByRef plotMethod As Integer)
            If (chkBoxNormalPlot.Checked And chkBoxBatch.Checked = False) Then
                plotMethod = 0
            ElseIf (chkBoxNormalPlot.Checked And chkBoxBatch.Checked) Then
                plotMethod = 1
            ElseIf (chkBoxIssuePlot.Checked And chkBoxBatch.Checked = False) Then
                plotMethod = 2
            ElseIf (chkBoxIssuePlot.Checked And chkBoxBatch.Checked) Then
                plotMethod = 3
            ElseIf (chkBoxPDF.Checked And chkBoxBatch.Checked = False) Then
                plotMethod = 4
            ElseIf (chkBoxPDF.Checked And chkBoxBatch.Checked) Then
                plotMethod = 5
            End If
        End Sub
        Private Sub okButton_Click(sender As Object, e As EventArgs) Handles okButton.Click
            Dim jobNum As String = Nothing
            Dim plotMethod As Integer = Nothing
            Dim CGMFileNames As String() = Nothing
            lw.WriteLine(chkBoxExistingJobList.checked.tostring)
            GetJobNum(jobNum)
            CheckPlotFolder(jobNum)
            GetPlotMethod(plotMethod)
            Try
                numCopies = Convert.ToInt32(txtBoxCopies.Text)
            Catch ex As Exception
            End Try
            If (plotMethod = 0) Then
                lw.WriteLine("Normal Plot, Batch Mode Disabled")

                lw.WriteLine("Plotter Choice: " + plotterChoice)
                STMPlot(jobNum)
                GetFileFromPlotFolder(jobNum, CGMFileNames, 1)
                For Each CGMFile As String In CGMFileNames
                    PrintCGMToPlotter(CGMFile, plotterChoice, plotterName, numCopies, dryRun)
                Next
            ElseIf (plotMethod = 1) Then
                lw.WriteLine("Normal Plot, Batch Mode")

                Dim SkipDwgsPlot() As String = Nothing
                Dim ProdDwgsPlot() As String = Nothing
                Dim OutOfDateDwgsPlot() As String = Nothing
                If chkBoxExistingJobList.checked = False Then
                    STMBatchPlot(jobNum, SkipDwgsPlot, ProdDwgsPlot, OutOfDateDwgsPlot)
                Else
                    Dim PlotDir As String = "C:\eng\plots\" & jobNum
                    Dim TempFilesLst As String() = Directory.GetFiles(PlotDir)

                    For Each TempFile As String In TempFilesLst
                        If TempFile.EndsWith(".cgm") = True Then
                            If ProdDwgsPlot Is Nothing Then
                                ReDim Preserve ProdDwgsPlot(0)
                            Else
                                ReDim Preserve ProdDwgsPlot(ProdDwgsPlot.Length)
                            End If
                            ProdDwgsPlot(ProdDwgsPlot.Length - 1) = TempFile
                        End If
                    Next
                End If

                Dim SmallDwgPlotter As Integer
                Dim BigDwgPlotter As Integer

                lw.WriteLine("plotter choice" + plotterChoice)

                If (plotterChoice = 1) Then
                    SmallDwgPlotter = 1 ' If the user explicitly selects a printer, print all drawings there. 
                    BigDwgPlotter = 1
                ElseIf (plotterChoice = 2) Then
                    SmallDwgPlotter = 2
                    BigDwgPlotter = 2
                Else
                    If (plotterChoice = 3 Or plotterChoice = 0) Then
                        BigDwgPlotter = 3
                    End If

                    If MessageBox.Show("Do you want all B Size Drawings to go to ENG2220? If not, it will go to SamsungPrintRoom?", "B Size Plotter", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        SmallDwgPlotter = 1
                    Else
                        SmallDwgPlotter = 2
                    End If
                End If

                GetFileFromPlotFolder(jobNum, CGMFileNames, ProdDwgsPlot.Length)

                If ConfirmProdDwgsPlot(jobNum, CGMFileNames, SkipDwgsPlot, ProdDwgsPlot) = False Then
                    Exit Sub
                End If
                Dim tempPlotterChoice As Integer = Nothing
                Dim tempPlotterName As String = Nothing
                tempPlotterChoice = plotterChoice
                tempPlotterName = plotterName
                For Each CGMFile As String In CGMFileNames
                    Dim PlotSize As String = CGMFile.Replace(".cgm", "")
                    lw.WriteLine(PlotSize)
                    ' PlotSize.EndsWith("_B") = True Or 
                    plotterChoice = tempPlotterChoice
                    plotterName = tempPlotterName
                    If (plotterChoice = 3 And PlotSize.EndsWith("_B")) Then
                        If (SmallDwgPlotter = 1) Then
                            plotterName = "Canon iR2220 ENG"
                            plotterChoice = 1
                        ElseIf (SmallDwgPlotter = 2) Then
                            plotterName = "ENG HP5200"
                            plotterChoice = 2
                        End If
                    ElseIf (plotterChoice = 1) Then
                        plotterName = "Canon iR2220 ENG"
                    ElseIf (plotterChoice = 2) Then
                        plotterName = "ENG HP5200"
                    End If
                    'MsgBox(CGMFile + System.Environment.NewLine + plotterChoice)
                    PrintCGMToPlotter(CGMFile, plotterChoice, plotterName, numCopies, dryRun)
                Next

                PrintLogToTxt(jobNum, ProdDwgsPlot, OutOfDateDwgsPlot, SkipDwgsPlot)
            ElseIf (plotMethod = 2) Then

                lw.WriteLine("Issue plot, batch mode disabled")
                STMIssuePlot(jobNum)

            ElseIf (plotMethod = 3) Then

                lw.WriteLine("Issue plot, batch mode enabled")
                If (STMBatchIssuePlot(jobNum) = False) Then
                    Exit Sub
                End If
            ElseIf (plotMethod = 4) Then

                lw.WriteLine("PDF, batch mode disabled")
                STMPlot(jobNum)
                plotterChoice = 4
                plotterName = "PDF"
                numCopies = 1

                GetFileFromPlotFolder(jobNum, CGMFileNames, 1)

                For Each CGMFile As String In CGMFileNames
                    PrintCGMToPDF(CGMFile, plotterChoice, plotterName, numCopies, dryRun)
                Next
            ElseIf (plotMethod = 5) Then

                lw.WriteLine("PDF, batch mode enabled")
                Dim SkipDwgsPlot() As String = Nothing
                Dim ProdDwgsPlot() As String = Nothing
                Dim OutOfDateDwgsPlot() As String = Nothing
                If chkBoxExistingJobList.checked = False Then
                    STMBatchPlot(jobNum, SkipDwgsPlot, ProdDwgsPlot, OutOfDateDwgsPlot)
                Else
                    Dim PlotDir As String = "C:\eng\plots\" & jobNum
                    Dim TempFilesLst As String() = Directory.GetFiles(PlotDir)
                    For Each TempFile As String In TempFilesLst
                        If TempFile.EndsWith(".cgm") = True Then
                            If ProdDwgsPlot Is Nothing Then
                                ReDim Preserve ProdDwgsPlot(0)
                            Else
                                ReDim Preserve ProdDwgsPlot(ProdDwgsPlot.Length)
                            End If

                            ProdDwgsPlot(ProdDwgsPlot.Length - 1) = TempFile
                        End If
                    Next
                End If

                plotterChoice = 4
                plotterName = "PDF"
                numCopies = 1

                GetFileFromPlotFolder(jobNum, CGMFileNames, 1)

                If ConfirmProdDwgsPlot(jobNum, CGMFileNames, SkipDwgsPlot, ProdDwgsPlot) = False Then
                    Exit Sub
                End If

                For Each CGMFile As String In CGMFileNames
                    Dim PlotSize As String = CGMFile.Replace(".cgm", "")
                    lw.WriteLine(PlotSize)
                    PrintCGMToPDF(CGMFile, plotterChoice, plotterName, numCopies, dryRun)
                Next

                PrintLogToTxt(jobNum, ProdDwgsPlot, OutOfDateDwgsPlot, SkipDwgsPlot)
            End If
        End Sub
    End Class
    Public Class FormProdDwgsPlot
        Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
        End Sub
        Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        End Sub
        Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click

            Dim MyFinalLst As String() = Nothing
            Dim MyRemoveLst As String() = Nothing

            If Me.lstRemoveCGMFiles.Items.Count <> 0 Then
                For i As Integer = 0 To Me.lstRemoveCGMFiles.Items.Count - 1
                    ReDim Preserve MyRemoveLst(i)
                    MyRemoveLst(i) = Me.lstRemoveCGMFiles.Items.Item(i)
                Next
            End If

            For i As Integer = 0 To Me.lstCGMFiles.Items.Count - 1
                Dim FoundSelectStatus As Boolean = False
                For Each mySelection As Integer In Me.lstCGMFiles.SelectedIndices

                    If i = mySelection Then

                        FoundSelectStatus = True
                        Exit For
                    Else
                        FoundSelectStatus = False
                    End If
                Next

                If FoundSelectStatus = False Then
                    If MyFinalLst Is Nothing Then
                        ReDim Preserve MyFinalLst(0)
                    Else
                        ReDim Preserve MyFinalLst(MyFinalLst.Length)
                    End If

                    MyFinalLst(MyFinalLst.Length - 1) = Me.lstCGMFiles.Items.Item(i)
                Else
                    If MyRemoveLst Is Nothing Then
                        ReDim Preserve MyRemoveLst(0)
                    Else
                        ReDim Preserve MyRemoveLst(MyRemoveLst.Length)
                    End If

                    MyRemoveLst(MyRemoveLst.Length - 1) = Me.lstCGMFiles.Items.Item(i)
                End If
            Next

            Me.lstCGMFiles.Items.Clear()

            If Not MyFinalLst Is Nothing Then

                For i As Integer = 0 To MyFinalLst.Length - 1
                    Me.lstCGMFiles.Items.Add(MyFinalLst(i))
                Next

                Me.lstCGMFiles.Update()

            End If

            Me.lstRemoveCGMFiles.Items.Clear()

            If Not MyRemoveLst Is Nothing Then

                For j As Integer = 0 To MyRemoveLst.Length - 1
                    Me.lstRemoveCGMFiles.Items.Add(MyRemoveLst(j))
                Next

                Me.lstRemoveCGMFiles.Update()
            End If
        End Sub
        Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
            Dim MyFinalLst As String() = Nothing
            Dim MyRemoveLst As String() = Nothing

            If Me.lstCGMFiles.Items.Count <> 0 Then
                For i As Integer = 0 To Me.lstCGMFiles.Items.Count - 1
                    ReDim Preserve MyFinalLst(i)
                    MyFinalLst(i) = Me.lstCGMFiles.Items.Item(i)
                Next
            End If

            For i As Integer = 0 To Me.lstRemoveCGMFiles.Items.Count - 1
                Dim FoundSelectStatus As Boolean = False
                For Each mySelection As Integer In Me.lstRemoveCGMFiles.SelectedIndices
                    If i = mySelection Then
                        FoundSelectStatus = True
                        Exit For
                    Else
                        FoundSelectStatus = False
                    End If
                Next

                If FoundSelectStatus = False Then
                    If MyRemoveLst Is Nothing Then
                        ReDim Preserve MyRemoveLst(0)
                    Else
                        ReDim Preserve MyRemoveLst(MyRemoveLst.Length)
                    End If

                    MyRemoveLst(MyRemoveLst.Length - 1) = Me.lstRemoveCGMFiles.Items.Item(i)
                Else
                    If MyFinalLst Is Nothing Then
                        ReDim Preserve MyFinalLst(0)
                    Else
                        ReDim Preserve MyFinalLst(MyFinalLst.Length)
                    End If
                    MyFinalLst(MyFinalLst.Length - 1) = Me.lstRemoveCGMFiles.Items.Item(i)
                End If
            Next

            Me.lstCGMFiles.Items.Clear()

            If Not MyFinalLst Is Nothing Then
                For i As Integer = 0 To MyFinalLst.Length - 1
                    Me.lstCGMFiles.Items.Add(MyFinalLst(i))
                Next
                Me.lstCGMFiles.Update()
            End If

            Me.lstRemoveCGMFiles.Items.Clear()

            If Not MyRemoveLst Is Nothing Then
                For j As Integer = 0 To MyRemoveLst.Length - 1
                    Me.lstRemoveCGMFiles.Items.Add(MyRemoveLst(j))
                Next
                Me.lstRemoveCGMFiles.Update()
            End If
        End Sub
    End Class


    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class StackteckPlotForm
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
            Me.Label4 = New System.Windows.Forms.Label()
            Me.txtBoxCopies = New System.Windows.Forms.TextBox()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.chkBoxKIP7700 = New System.Windows.Forms.CheckBox()
            Me.chkBoxSamsungPrintRoom = New System.Windows.Forms.CheckBox()
            Me.chkBoxEng2220 = New System.Windows.Forms.CheckBox()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.chkBoxPDF = New System.Windows.Forms.CheckBox()
            Me.chkBoxIssuePlot = New System.Windows.Forms.CheckBox()
            Me.chkBoxNormalPlot = New System.Windows.Forms.CheckBox()
            Me.helpButton = New System.Windows.Forms.Button()
            Me.chkBoxDryRun = New System.Windows.Forms.CheckBox()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.cancelButton = New System.Windows.Forms.Button()
            Me.okButton = New System.Windows.Forms.Button()
            Me.chkBoxBatch = New System.Windows.Forms.CheckBox()
            Me.chkBoxExistingJobList = New System.Windows.Forms.CheckBox()
            Me.TxtboxPapersize = New System.Windows.Forms.TextBox()
            Me.Label5 = New System.Windows.Forms.Label()
            Me.SuspendLayout()
            '
            'Label4
            '
            Me.Label4.AutoSize = True
            Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label4.Location = New System.Drawing.Point(24, 307)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(45, 13)
            Me.Label4.TabIndex = 32
            Me.Label4.Text = "Copies"
            '
            'txtBoxCopies
            '
            Me.txtBoxCopies.Location = New System.Drawing.Point(75, 304)
            Me.txtBoxCopies.Name = "txtBoxCopies"
            Me.txtBoxCopies.Size = New System.Drawing.Size(115, 20)
            Me.txtBoxCopies.TabIndex = 31
            '
            'Label3
            '
            Me.Label3.AutoSize = True
            Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label3.Location = New System.Drawing.Point(25, 226)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(50, 13)
            Me.Label3.TabIndex = 30
            Me.Label3.Text = "Options"
            '
            'chkBoxKIP7700
            '
            Me.chkBoxKIP7700.AutoSize = True
            Me.chkBoxKIP7700.Location = New System.Drawing.Point(49, 195)
            Me.chkBoxKIP7700.Name = "chkBoxKIP7700"
            Me.chkBoxKIP7700.Size = New System.Drawing.Size(67, 17)
            Me.chkBoxKIP7700.TabIndex = 29
            Me.chkBoxKIP7700.Text = "KIP7700"
            Me.chkBoxKIP7700.UseVisualStyleBackColor = True
            '
            'chkBoxSamsungPrintRoom
            '
            Me.chkBoxSamsungPrintRoom.AutoSize = True
            Me.chkBoxSamsungPrintRoom.Location = New System.Drawing.Point(49, 172)
            Me.chkBoxSamsungPrintRoom.Name = "chkBoxSamsungPrintRoom"
            Me.chkBoxSamsungPrintRoom.Size = New System.Drawing.Size(119, 17)
            Me.chkBoxSamsungPrintRoom.TabIndex = 28
            Me.chkBoxSamsungPrintRoom.Text = "SamsungPrintRoom"
            Me.chkBoxSamsungPrintRoom.UseVisualStyleBackColor = True
            '
            'chkBoxEng2220
            '
            Me.chkBoxEng2220.AutoSize = True
            Me.chkBoxEng2220.Location = New System.Drawing.Point(49, 149)
            Me.chkBoxEng2220.Name = "chkBoxEng2220"
            Me.chkBoxEng2220.Size = New System.Drawing.Size(101, 17)
            Me.chkBoxEng2220.TabIndex = 27
            Me.chkBoxEng2220.Text = "ENGINEERING"
            Me.chkBoxEng2220.UseVisualStyleBackColor = True
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label2.Location = New System.Drawing.Point(27, 122)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(44, 13)
            Me.Label2.TabIndex = 26
            Me.Label2.Text = "Printer"
            '
            'chkBoxPDF
            '
            Me.chkBoxPDF.AutoSize = True
            Me.chkBoxPDF.Location = New System.Drawing.Point(49, 92)
            Me.chkBoxPDF.Name = "chkBoxPDF"
            Me.chkBoxPDF.Size = New System.Drawing.Size(47, 17)
            Me.chkBoxPDF.TabIndex = 25
            Me.chkBoxPDF.Text = "PDF"
            Me.chkBoxPDF.UseVisualStyleBackColor = True
            '
            'chkBoxIssuePlot
            '
            Me.chkBoxIssuePlot.AutoSize = True
            Me.chkBoxIssuePlot.Location = New System.Drawing.Point(49, 69)
            Me.chkBoxIssuePlot.Name = "chkBoxIssuePlot"
            Me.chkBoxIssuePlot.Size = New System.Drawing.Size(72, 17)
            Me.chkBoxIssuePlot.TabIndex = 24
            Me.chkBoxIssuePlot.Text = "Issue Plot"
            Me.chkBoxIssuePlot.UseVisualStyleBackColor = True
            '
            'chkBoxNormalPlot
            '
            Me.chkBoxNormalPlot.AutoSize = True
            Me.chkBoxNormalPlot.Location = New System.Drawing.Point(49, 46)
            Me.chkBoxNormalPlot.Name = "chkBoxNormalPlot"
            Me.chkBoxNormalPlot.Size = New System.Drawing.Size(80, 17)
            Me.chkBoxNormalPlot.TabIndex = 23
            Me.chkBoxNormalPlot.Text = "Normal Plot"
            Me.chkBoxNormalPlot.UseVisualStyleBackColor = True
            '
            'helpButton
            '
            Me.helpButton.Location = New System.Drawing.Point(206, 340)
            Me.helpButton.Name = "helpButton"
            Me.helpButton.Size = New System.Drawing.Size(28, 23)
            Me.helpButton.TabIndex = 22
            Me.helpButton.Text = "?"
            Me.helpButton.UseVisualStyleBackColor = True
            '
            'chkBoxDryRun
            '
            Me.chkBoxDryRun.AutoSize = True
            Me.chkBoxDryRun.Location = New System.Drawing.Point(49, 274)
            Me.chkBoxDryRun.Name = "chkBoxDryRun"
            Me.chkBoxDryRun.Size = New System.Drawing.Size(111, 17)
            Me.chkBoxDryRun.TabIndex = 18
            Me.chkBoxDryRun.Text = "Dry Run (No print)"
            Me.chkBoxDryRun.UseVisualStyleBackColor = True
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label1.Location = New System.Drawing.Point(27, 22)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(75, 13)
            Me.Label1.TabIndex = 19
            Me.Label1.Text = "Plot Method"
            '
            'cancelButton
            '
            Me.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.cancelButton.Location = New System.Drawing.Point(116, 340)
            Me.cancelButton.Name = "cancelButton"
            Me.cancelButton.Size = New System.Drawing.Size(75, 23)
            Me.cancelButton.TabIndex = 21
            Me.cancelButton.Text = "Cancel"
            Me.cancelButton.UseVisualStyleBackColor = True
            '
            'okButton
            '
            Me.okButton.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.okButton.Location = New System.Drawing.Point(25, 340)
            Me.okButton.Name = "okButton"
            Me.okButton.Size = New System.Drawing.Size(75, 23)
            Me.okButton.TabIndex = 20
            Me.okButton.Text = "OK"
            Me.okButton.UseVisualStyleBackColor = True
            '
            'chkBoxBatch
            '
            Me.chkBoxBatch.AutoSize = True
            Me.chkBoxBatch.Location = New System.Drawing.Point(49, 251)
            Me.chkBoxBatch.Name = "chkBoxBatch"
            Me.chkBoxBatch.Size = New System.Drawing.Size(54, 17)
            Me.chkBoxBatch.TabIndex = 17
            Me.chkBoxBatch.Text = "Batch"
            Me.chkBoxBatch.UseVisualStyleBackColor = True
            '
            'chkBoxExistingJobList
            '
            Me.chkBoxExistingJobList.AutoSize = True
            Me.chkBoxExistingJobList.Location = New System.Drawing.Point(116, 251)
            Me.chkBoxExistingJobList.Name = "chkBoxExistingJobList"
            Me.chkBoxExistingJobList.Size = New System.Drawing.Size(134, 17)
            Me.chkBoxExistingJobList.TabIndex = 33
            Me.chkBoxExistingJobList.Text = "Reprint Previous Dwgs"
            Me.chkBoxExistingJobList.UseVisualStyleBackColor = True
            Me.chkBoxExistingJobList.Visible = False
            '
            'TxtboxPapersize
            '
            Me.TxtboxPapersize.Location = New System.Drawing.Point(147, 92)
            Me.TxtboxPapersize.Name = "TxtboxPapersize"
            Me.TxtboxPapersize.ReadOnly = True
            Me.TxtboxPapersize.Size = New System.Drawing.Size(100, 20)
            Me.TxtboxPapersize.TabIndex = 35
            Me.TxtboxPapersize.TabStop = False
            '
            'Label5
            '
            Me.Label5.AutoSize = True
            Me.Label5.Location = New System.Drawing.Point(144, 76)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(53, 13)
            Me.Label5.TabIndex = 37
            Me.Label5.Text = "Papersize"
            '
            'StackteckPlotForm
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(259, 382)
            Me.Controls.Add(Me.Label5)
            Me.Controls.Add(Me.TxtboxPapersize)
            Me.Controls.Add(Me.chkBoxExistingJobList)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.txtBoxCopies)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me.chkBoxKIP7700)
            Me.Controls.Add(Me.chkBoxSamsungPrintRoom)
            Me.Controls.Add(Me.chkBoxEng2220)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.chkBoxPDF)
            Me.Controls.Add(Me.chkBoxIssuePlot)
            Me.Controls.Add(Me.chkBoxNormalPlot)
            Me.Controls.Add(Me.helpButton)
            Me.Controls.Add(Me.chkBoxDryRun)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.cancelButton)
            Me.Controls.Add(Me.okButton)
            Me.Controls.Add(Me.chkBoxBatch)
            Me.Name = "StackteckPlotForm"
            Me.Text = "Stackteck Plot V2"
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Friend WithEvents Label4 As Label
        Friend WithEvents txtBoxCopies As TextBox
        Friend WithEvents Label3 As Label
        Friend WithEvents chkBoxKIP7700 As CheckBox
        Friend WithEvents chkBoxSamsungPrintRoom As CheckBox
        Friend WithEvents chkBoxEng2220 As CheckBox
        Friend WithEvents Label2 As Label
        Friend WithEvents chkBoxPDF As CheckBox
        Friend WithEvents chkBoxIssuePlot As CheckBox
        Friend WithEvents chkBoxNormalPlot As CheckBox
        Friend WithEvents helpButton As Button
        Friend WithEvents chkBoxDryRun As CheckBox
        Friend WithEvents Label1 As Label
        Friend WithEvents cancelButton As Button
        Friend WithEvents okButton As Button
        Friend WithEvents chkBoxBatch As CheckBox
        Friend WithEvents chkBoxExistingJobList As CheckBox
        Friend WithEvents TxtboxPapersize As TextBox
        Friend WithEvents Label5 As Label
    End Class





    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class FormProdDwgsPlot
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
            Me.Label5 = New System.Windows.Forms.Label()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.lstRemoveCGMFiles = New System.Windows.Forms.ListBox()
            Me.btnReset = New System.Windows.Forms.Button()
            Me.btnRemove = New System.Windows.Forms.Button()
            Me.btnCancel = New System.Windows.Forms.Button()
            Me.btnOK = New System.Windows.Forms.Button()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.lstCGMFiles = New System.Windows.Forms.ListBox()
            Me.SuspendLayout()
            '
            'Label5
            '
            Me.Label5.AutoSize = True
            Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label5.Location = New System.Drawing.Point(28, 40)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(63, 13)
            Me.Label5.TabIndex = 20
            Me.Label5.Text = "To Print List"
            '
            'Label4
            '
            Me.Label4.AutoSize = True
            Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label4.Location = New System.Drawing.Point(33, 211)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(87, 13)
            Me.Label4.TabIndex = 19
            Me.Label4.Text = "Remove Plot List"
            '
            'lstRemoveCGMFiles
            '
            Me.lstRemoveCGMFiles.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
            Me.lstRemoveCGMFiles.FormattingEnabled = True
            Me.lstRemoveCGMFiles.Location = New System.Drawing.Point(28, 228)
            Me.lstRemoveCGMFiles.Name = "lstRemoveCGMFiles"
            Me.lstRemoveCGMFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.lstRemoveCGMFiles.Size = New System.Drawing.Size(284, 134)
            Me.lstRemoveCGMFiles.Sorted = True
            Me.lstRemoveCGMFiles.TabIndex = 18
            '
            'btnReset
            '
            Me.btnReset.Location = New System.Drawing.Point(333, 228)
            Me.btnReset.Name = "btnReset"
            Me.btnReset.Size = New System.Drawing.Size(75, 23)
            Me.btnReset.TabIndex = 17
            Me.btnReset.Text = "Add"
            Me.btnReset.UseVisualStyleBackColor = True
            '
            'btnRemove
            '
            Me.btnRemove.Location = New System.Drawing.Point(333, 59)
            Me.btnRemove.Name = "btnRemove"
            Me.btnRemove.Size = New System.Drawing.Size(75, 23)
            Me.btnRemove.TabIndex = 16
            Me.btnRemove.Text = "Remove"
            Me.btnRemove.UseVisualStyleBackColor = True
            '
            'btnCancel
            '
            Me.btnCancel.Location = New System.Drawing.Point(237, 415)
            Me.btnCancel.Name = "btnCancel"
            Me.btnCancel.Size = New System.Drawing.Size(75, 23)
            Me.btnCancel.TabIndex = 15
            Me.btnCancel.Text = "Cancel"
            Me.btnCancel.UseVisualStyleBackColor = True
            '
            'btnOK
            '
            Me.btnOK.Location = New System.Drawing.Point(28, 415)
            Me.btnOK.Name = "btnOK"
            Me.btnOK.Size = New System.Drawing.Size(75, 23)
            Me.btnOK.TabIndex = 14
            Me.btnOK.Text = "OK"
            Me.btnOK.UseVisualStyleBackColor = True
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Location = New System.Drawing.Point(25, 377)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(157, 13)
            Me.Label2.TabIndex = 13
            Me.Label2.Text = "P.S.: Multi Select by: Crtl + MB1"
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Location = New System.Drawing.Point(25, 22)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(232, 13)
            Me.Label1.TabIndex = 12
            Me.Label1.Text = "Select the drawing(s) you dont want to print out:"
            '
            ' lstCGMFiles
            '
            Me.lstCGMFiles.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
            Me.lstCGMFiles.FormattingEnabled = True
            Me.lstCGMFiles.Location = New System.Drawing.Point(28, 59)
            Me.lstCGMFiles.Name = "lstCGMFiles"
            Me.lstCGMFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.lstCGMFiles.Size = New System.Drawing.Size(284, 134)
            Me.lstCGMFiles.Sorted = True
            Me.lstCGMFiles.TabIndex = 11
            '
            'FormProdDwgsPlot
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(433, 461)
            Me.Controls.Add(Me.Label5)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.lstRemoveCGMFiles)
            Me.Controls.Add(Me.btnReset)
            Me.Controls.Add(Me.btnRemove)
            Me.Controls.Add(Me.btnCancel)
            Me.Controls.Add(Me.btnOK)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.lstCGMFiles)
            Me.Name = "FormProdDwgsPlot"
            Me.Text = "Form1"
            Me.ResumeLayout(False)
            Me.PerformLayout()
        End Sub

        Friend WithEvents Label5 As Label
        Friend WithEvents Label4 As Label
        Friend WithEvents lstRemoveCGMFiles As ListBox
        Friend WithEvents btnReset As Button
        Friend WithEvents btnRemove As Button
        Friend WithEvents btnCancel As Button
        Friend WithEvents btnOK As Button
        Friend WithEvents Label2 As Label
        Friend WithEvents Label1 As Label
        Friend WithEvents lstCGMFiles As ListBox
    End Class

End Module