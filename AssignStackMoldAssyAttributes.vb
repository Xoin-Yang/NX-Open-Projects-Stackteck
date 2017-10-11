Option Strict Off
Imports System
Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Data.Sql
Imports System.Data
Imports System.String
Imports NXOpen
Imports NXOpen.UF
Imports NXOpenUI
Imports NXOpen.UI
Imports NXOpen.Utilities
Imports NXOpen.Assemblies
Imports NXOpen.Layer
Imports NXOpen.Drawings
Imports System.Text.RegularExpressions


Module AssignStackAssemblyAttributes
    Public s As Session = Session.GetSession()
    Dim nl As String = Environment.NewLine
    Dim dispPart As Part = s.Parts.Display()
    Dim workPart As Part = s.Parts.Work()
    Public theUI As UI = UI.GetUI
    Public ufs As UFSession = UFSession.GetUFSession()
    Public lw As ListingWindow = s.ListingWindow
    Dim fileName = s.Parts.Work.FullPath()
    Dim Count As Integer
    Dim JobNumber As String
    Dim moldRunCount As String = Nothing
    Dim foundStack As Boolean = False

    ' Global variables for good reason 
    Dim LastTabPage As String = 0 ' This allows us to open to the last page that the user was on 

    Dim PartFlowLength As String = Nothing ' DONE 
    Dim PartFlowLengthGeneral As String = Nothing
    Dim PartFlowLengthMetric As String = Nothing
    Dim PartFlowLengthImperial As String = Nothing

    Dim PartDiameter As String = Nothing ' DONE 
    Dim PartDiameterGeneral As String = Nothing
    Dim PartDiameterMetric As String = Nothing
    Dim PartDiameterImperial As String = Nothing

    Dim PartWidth As String = Nothing
    Dim PartWidthGeneral As String = Nothing
    Dim PartWidthMetric As String = Nothing
    Dim PartWidthImperial As String = Nothing
    Dim PartLength As String = Nothing
    Dim PartLengthGeneral As String = Nothing
    Dim PartLengthMetric As String = Nothing
    Dim PartLengthImperial As String = Nothing
    Dim PartWidthAndLengthGeneralString As String = Nothing
    Dim PartWidthAndLengthMetricString As String = Nothing
    Dim PartWidthAndLengthImperialString As String = Nothing

    Dim PartHeight As String = Nothing
    Dim PartHeightGeneral As String = Nothing
    Dim PartHeightMetric As String = Nothing
    Dim PartHeightImperial As String = Nothing

    Dim PartSideWS As String = Nothing
    Dim PartSideWSGeneral As String = Nothing
    Dim PartSideWSMetric As String = Nothing
    Dim PartSideWSImperial As String = Nothing

    Dim PartBottomWS As String = Nothing
    Dim PartBottomWSGeneral As String = Nothing
    Dim PartBottomWSMetric As String = Nothing
    Dim PartBottomWSImperial As String = Nothing

    Dim MoldTotalShutHeightGeneralValue As String = Nothing
    Dim MoldTotalShutHeightGeneralUnit As String = Nothing
    Dim MoldTotalShutHeightMetricValue As String = Nothing
    Dim MoldTotalShutHeightImperialValue As String = Nothing
    Dim MoldTotalShutHeightDualUnit As String = Nothing

    Dim MoldWeightGeneralValue As String = Nothing
    Dim MoldWeightGeneralUnit As String = Nothing
    Dim MoldWeightMetricValue As String = Nothing
    Dim MoldWeightImperialValue As String = Nothing
    Dim MoldWeightDualUnit As String = Nothing

    Dim MoldPitchXGeneralValue As String = Nothing
    Dim MoldPitchGeneralUnit As String = Nothing
    Dim MoldPitchXMetricValue As String = Nothing
    Dim MoldPitchXImperialValue As String = Nothing
    Dim MoldPitchYGeneralValue As String = Nothing
    Dim MoldPitchYMetricValue As String = Nothing
    Dim MoldPitchYImperialValue As String = Nothing
    Dim MoldPitchDualUnit As String = Nothing

    Dim MoldSizeXGeneralValue As String = Nothing
    Dim MoldSizeXYGeneralUnit As String = Nothing
    Dim MoldSizeXMetricValue As String = Nothing
    Dim MoldSizeXImperialValue As String = Nothing
    Dim MoldSizeYGeneralValue As String = Nothing
    Dim MoldSizeYMetricValue As String = Nothing
    Dim MoldSizeYImperialValue As String = Nothing
    Dim MoldSizeXYDualUnit As String = Nothing

    Dim MoldTonnageGeneralValue As String = Nothing
    Dim MoldTonnageGeneralUnit As String = Nothing
    Dim MoldTonnageMetricValue As String = Nothing
    Dim MoldTonnageImperialValue As String = Nothing
    Dim MoldTonnageDualUnit As String = Nothing

    Dim MoldQPCGeneralValue As String = Nothing
    Dim MoldQPCGeneralUnit As String = Nothing
    Dim MoldQPCMetricValue As String = Nothing
    Dim MoldQPCImperialValue As String = Nothing
    Dim MoldQPCDualUnit As String = Nothing

    Dim MoldEjectionStrokeGeneralValue As String = Nothing
    Dim MoldEjectionStrokeGeneralUnit As String = Nothing
    Dim MoldEjectionStrokeMetricValue As String = Nothing
    Dim MoldEjectionStrokeImperialValue As String = Nothing
    Dim MoldEjectionStrokeDualUnit As String = Nothing

    Dim MoldUnitsReadYet As Boolean = False
    Dim OldMoldUnits As String = Nothing

    Dim HotRunnerLDimGeneralValue As String = Nothing
    Dim HotRunnerLDimGeneralUnit As String = Nothing
    Dim HotRunnerLDimMetricValue As String = Nothing
    Dim HotRunnerLDimImperialValue As String = Nothing
    Dim HotRunnerLDimDualUnit As String = Nothing

    Dim HotRunnerXDimGeneralValue As String = Nothing
    Dim HotRunnerXDimGeneralUnit As String = Nothing
    Dim HotRunnerXDimMetricValue As String = Nothing
    Dim HotRunnerXDimImperialValue As String = Nothing
    Dim HotRunnerXDimDualUnit As String = Nothing

    Dim HotRunnerPDimHotGeneralValue As String = Nothing
    Dim HotRunnerPDimHotGeneralUnit As String = Nothing
    Dim HotRunnerPDimHotMetricValue As String = Nothing
    Dim HotRunnerPDimHotImperialValue As String = Nothing
    Dim HotRunnerPDimHotDualUnit As String = Nothing

    Dim HotRunnerPDimColdGeneralValue As String = Nothing
    Dim HotRunnerPDimColdGeneralUnit As String = Nothing
    Dim HotRunnerPDimColdMetricValue As String = Nothing
    Dim HotRunnerPDimColdImperialValue As String = Nothing
    Dim HotRunnerPDimColdDualUnit As String = Nothing

    Dim HotRunnerGateDiameterGeneralValue As String = Nothing
    Dim HotRunnerGateDiameterGeneralUnit As String = Nothing
    Dim HotRunnerGateDiameterMetricValue As String = Nothing
    Dim HotRunnerGateDiameterImperialValue As String = Nothing
    Dim HotRunnerGateDiameterDualUnit As String = Nothing

    Dim MachineTieBarH As String = Nothing
    Dim MachineTieBarHMetric As String = Nothing
    Dim MachineTieBarHImperial As String = Nothing
    Dim MachineTieBarV As String = Nothing
    Dim MachineTieBarVMetric As String = Nothing
    Dim MachineTieBarVImperial As String = Nothing
    Dim MachineTieBarLastUnit As String = Nothing
    Dim MachineTieBarUnitsImperial As String = Nothing
    Dim MachineTieBarUnitsMetric As String = Nothing
    Dim MachineTieBarGeneralString As String = Nothing ' Since we need the x signs to disappear in UG, we have to read and write it as a string 
    Dim MachineTieBarMetricString As String = Nothing
    Dim MachineTieBarImperialString As String = Nothing
    Dim MachineTieBarDualUnit As String = Nothing

    Dim MachineClampTonnageGeneralValue As String = Nothing
    Dim MachineClampTonnageGeneralUnit As String = Nothing ' If we only have one entry, we need to know if its metric or imperial
    Dim MachineClampTonnageMetricValue As String = Nothing
    Dim MachineClampTonnageImperialValue As String = Nothing
    Dim MachineClampTonnageDualUnit As String = Nothing ' To keep track of which way it was entered, mm-dual or dual-mm, so that we know which value to show in the program 

    Dim MachineClampStrokeGeneralValue As String = Nothing ' Done
    Dim MachineClampStrokeGeneralUnit As String = Nothing
    Dim MachineClampStrokeMetricValue As String = Nothing
    Dim MachineClampStrokeImperialValue As String = Nothing
    Dim MachineClampStrokeDualUnit As String = Nothing

    Dim MachineMaxDaylightGeneralValue As String = Nothing ' Done
    Dim MachineMaxDaylightGeneralUnit As String = Nothing
    Dim MachineMaxDaylightMetricValue As String = Nothing
    Dim MachineMaxDaylightImperialValue As String = Nothing
    Dim MachineMaxDaylightDualUnit As String = Nothing

    Dim MachineLocatingRingDiameterGeneralValue As String = Nothing ' Done
    Dim MachineLocatingRingDiameterGeneralUnit As String = Nothing
    Dim MachineLocatingRingDiameterMetricValue As String = Nothing
    Dim MachineLocatingRingDiameterImperialValue As String = Nothing
    Dim MachineLocatingRingDiameterDualUnit As String = Nothing

    Dim MachineNozzleRadiusGeneralValue As String = Nothing ' Done
    Dim MachineNozzleRadiusGeneralUnit As String = Nothing
    Dim MachineNozzleRadiusMetricValue As String = Nothing
    Dim MachineNozzleRadiusImperialValue As String = Nothing
    Dim MachineNozzleRadiusDualUnit As String = Nothing

    Dim MachineMaxEjectorStrokeGeneralValue As String = Nothing ' Done
    Dim MachineMaxEjectorStrokeGeneralUnit As String = Nothing
    Dim MachineMaxEjectorStrokeMetricValue As String = Nothing
    Dim machineMaxEjectorStrokeImperialValue As String = Nothing
    Dim MachineMaxEjectorStrokeDualUnit As String = Nothing

    Dim MachineMinShutHeightGeneralValue As String = Nothing ' Done
    Dim MachineShutHeightGeneralUnit As String = Nothing
    Dim MachineMinShutHeightMetricValue As String = Nothing
    Dim MachineMinShutHeightImperialValue As String = Nothing
    Dim MachineMaxShutHeightGeneralValue As String = Nothing
    Dim MachineMaxShutHeightMetricValue As String = Nothing
    Dim MachineMaxShutHeightImperialValue As String = Nothing
    Dim MachineShutHeightDualUnit As String = Nothing

    Sub Main()
        CreateUsageLog("Assign Stack Assembly Attributes Program")
        lw.Open()
        lw.WriteLine("Starting Assign Stack Assembly Attributes Program")
        workPart.DeleteRetainedDraftingObjectsInCurrentLayout()

        Dim filePath As String = s.Parts.Work.FullPath() ' E.g. AIM_StackCup24oz_S37452/001
        lw.WriteLine("Opening " + filePath)
        JobNumber = Left(filePath, 5)

        Dim infoform As Form1
        infoform = New Form1()
        infoform.ShowDialog()
    End Sub

    Public Function GetUnloadOption(ByVal dum As String) As Integer
        Return Session.LibraryUnloadOption.Immediately
    End Function

    Public Sub CreateUsageLog(ByVal ProgramName As String)
        Dim username As String = System.Environment.UserName
        Dim UseDate As String = Now().Day & "-" & Now().Month & "-" & Now().Year

        Dim UsageLogFolderDir As String = "\\enghome\ugnx_settings\logs\UG_Prog"

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

    Public Class Form1
        Public Sub GetUserInfo(ByRef UserName As String, ByRef UserGroup As String, ByRef Status As Boolean)
            Dim UserID As String
            UserID = Environment.UserName()
            lw.WriteLine("UserID is: " + UserID)

            'Define the connectors
            Dim cn As OleDbConnection
            Dim cmd As OleDbCommand
            Dim dr As OleDbDataReader
            Dim oConnect, oQuery As String
            Dim FoundStatus As Boolean = False

            'Define connection string
            Dim FileName As String = "Y:\eng\ENG_ACCESS_DATABASES\UGMisDatabase.mdb"
            If File.Exists(FileName) = False Then
                MessageBox.Show("File " & FileName & " is not found.")
                Status = False
                Exit Sub
            End If

            oConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName
            lw.WriteLine("Connecting to: " + FileName)

            'Query String
            oQuery = "SELECT * FROM Title_Blk_User_Info where UserID='" & UserID & "'"

            'Instantiate the connectors
            cn = New OleDbConnection(oConnect)
            cn.Open()

            cmd = New OleDbCommand(oQuery, cn)
            dr = cmd.ExecuteReader

            While dr.Read()
                UserName = dr(1)
                UserName = UserName.Trim()
                'MessageBox.Show("UserName is: " & UserName)

                UserGroup = dr(2)
                UserGroup = UserGroup.Trim()
                'MessageBox.Show("UserGroup is: " & UserGroup)
                lw.WriteLine("Design team is: " + UserGroup)
                FoundStatus = True
                Status = True
            End While

            dr.Close()
            cn.Close()

            If FoundStatus = False Then
                MessageBox.Show("UserID " & UserID & " is not found from database")
            End If
        End Sub

        Private Sub GetDate(ByRef time As String)
            Dim objDate As Date = Date.Now()
            Dim Year As String
            Dim Month As String
            Dim Day As String

            Year = objDate.Year
            Month = objDate.Month
            Day = objDate.Day

            Select Case Month
                Case 1
                    Month = "JAN"
                Case 2
                    Month = "FEB"
                Case 3
                    Month = "MAR"
                Case 4
                    Month = "APR"
                Case 5
                    Month = "MAY"
                Case 6
                    Month = "JUN"
                Case 7
                    Month = "JUL"
                Case 8
                    Month = "AUG"
                Case 9
                    Month = "SEP"
                Case 10
                    Month = "OCT"
                Case 11
                    Month = "NOV"
                Case 12
                    Month = "DEC"
                Case Else
                    MessageBox.Show("Error")
            End Select

            time = Month & " " & Day & ", " & Year
            lw.WriteLine("Assigned Date As " + time)
        End Sub
        Private Sub AssignTeamAndUser()
            Dim UserName As String = Nothing
            Dim UserGroup As String = Nothing
            Dim Status As Boolean = Nothing
            GetUserInfo(UserName, UserGroup, Status)
            txtBoxDesigner.Text = UserName
            txtBoxDesignTeam.Text = UserGroup
            lw.WriteLine("Assigned Designer As: " + UserName)
            lw.WriteLine("Assigned Design Team As: " + UserGroup)
        End Sub
        Private Sub AssignDate()
            Dim time As String = Nothing
            GetDate(time)
            txtBoxDate.Text = time
        End Sub
        Private Sub AssignJobNum()
            txtBoxJobNumber.Text = s.Parts.Work.FullPath().Substring(0, 5)
        End Sub
        Public Sub SetAttribute(ByVal title As String, ByRef value As String)
            Dim s As Session = Session.GetSession()
            Dim dispPart As Part = s.Parts.Display()
            dispPart.SetAttribute(title, value)
        End Sub
        Public Sub ReadAttribute(ByVal title As String, ByRef value As String)
            Dim dispPart As Part = s.Parts.Display()
            find_part_attr_by_name(dispPart, title, value)
            'If (title <> "PROGRAM_RUN_COUNT_ASSY" And Count = 0) Then
            '    value = ""
            'End If
        End Sub
        Public Sub ReadChildAttribute(ByVal title As String, ByRef value As String, ByRef child As Component)
            value = child.GetStringAttribute(title)
        End Sub
        Public Sub find_part_attr_by_name(ByVal thePart As Part,
                                  ByVal attrName As String,
                                  ByRef attrVal As String)

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

        Public Sub getComponentNameAndDesc(ByVal comp As Component, ByRef name As String, ByRef desc As String)
            name = comp.DisplayName()
            name = name.Remove(name.Length - 4) ' Remove /XXX revision
            'name = name.Replace(txtBoxJobNumber.Text + "_", "")
            'name = name.Replace(JobNumber + "_", "")
            name = name.Replace(name.Substring(0, 5) + "_", "")

            Try
                If (IsNumeric(name.Substring(0, 5)) = False) Then
                    If (comp.GetStringAttribute("STACKTECK_PARTN").Trim = "") Then
                    Else
                        name = comp.GetStringAttribute("STACKTECK_PARTN")
                    End If
                End If
            Catch ex As Exception
                name = comp.GetStringAttribute("STACKTECK_PARTN")
            End Try

            Try
                desc = comp.GetStringAttribute("STACKTECK_DESC")
                If (desc = "") Then
                    desc = comp.GetStringAttribute("DB_PART_DESC")
                End If
            Catch ex As Exception
                desc = ""
            End Try
        End Sub

        Public Sub populateComponentInformationGrid()
            For Each row As DataGridViewRow In titleBlockComps.Rows
                Try
                    Dim c As ComponentAssembly = dispPart.ComponentAssembly
                    readMoldComponentInfo(c.RootComponent, 0, row.cells(0).value.ToString, row.cells(1).value.ToString)
                Catch ex As Exception
                End Try
            Next
        End Sub

        Public Sub readMoldComponentInfo(ByVal comp As Component, ByVal indent As Integer, ByRef partName As String, ByRef Component As String)
            For Each child As Component In comp.GetChildren()
                Dim shortName As String = Nothing
                Dim componentDesc As String = Nothing

                getComponentNameAndDesc(child, shortName, componentDesc)

                If (partName.Trim = shortName And Component = componentDesc) Then

                    Dim material As String = Nothing
                    Dim hardness As String = Nothing
                    Dim surfEnh As String = Nothing

                    Try
                        material = child.GetStringAttribute("STACKTECK_MATERIAL")
                    Catch ex As Exception
                        material = ""
                    End Try
                    Try
                        hardness = child.GetStringAttribute("STACKTECK_HARDNESS")
                    Catch ex As Exception
                        hardness = ""
                    End Try
                    Try
                        surfEnh = child.GetStringAttribute("STACKTECK_SURFACE_ENHANCEMENT")
                    Catch ex As Exception
                        surfEnh = ""
                    End Try

                    Dim doesExist As Boolean = False

                    For Each row As DataGridViewRow In compInfo.Rows
                        If (row.cells(0).value.ToString.Trim = componentDesc And row.cells(4).value.ToString.Trim = partName) Then
                            doesExist = True
                            row.cells(1).value = material
                            row.cells(2).value = hardness
                            If (surfEnh.Trim <> "") Then
                                row.cells(3).value = surfEnh
                            End If
                            Exit For
                        End If
                    Next

                    If (doesExist = False) Then
                        compInfo.Rows.Add(New String() {componentDesc, material, hardness, surfEnh, partName})
                    End If
                End If
                readMoldComponentInfo(child, indent + 1, partName, Component)
            Next
        End Sub

        Public Sub populateListBoxAndDataGrids()
            Try
                Dim c As ComponentAssembly = dispPart.ComponentAssembly

                If Not IsNothing(c.RootComponent) Then
                    fillListBoxAndDataGrids(c.RootComponent, 0, "")
                Else
                    lw.WriteLine("Part has no components")
                End If
            Catch e As Exception
                lw.WriteLine("Failed: " & e.ToString)
            End Try
            CorrectRowSizes()
        End Sub
        Public Sub fillListBoxAndDataGrids(ByVal comp As Component, ByVal indent As Integer, ByVal parentName As String)
            If (parentName = "") Then
                parentName = s.Parts.Display.FullPath()
            End If
            For Each child As Component In comp.GetChildren()

                'lw.WriteLine("Child.DisplayName: " + child.DisplayName + nl + "Parent Display Name: " + parentName)
                Dim firstChar As String = child.DisplayName().Substring(child.DisplayName().Length - 10, 1)
                Dim componentName As String = Nothing
                Dim componentDesc As String = Nothing
                Dim doesExistInTitleBlockGrid As Boolean = False
                Dim doesExistInAssyCompGrid As Boolean = False

                getComponentNameAndDesc(child, componentName, componentDesc)

                'first, check If the sixth character from the End Is S, P, Or A
                If ((firstChar = "S" Or firstChar = "P" Or firstChar = "A" Or child.DisplayName.Contains("CUST")) And IsNumeric(child.DisplayName().Substring(child.DisplayName.Length - 9, 5))) Then
                    Dim doesExistInListBox As Boolean = False
                    For Each item As String In listBoxApplicationsParts.Items
                        If (child.DisplayName() = item) Then
                            doesExistInListBox = True
                            Exit For
                        End If
                    Next

                    If (doesExistInListBox = False) Then
                        listBoxApplicationsParts.Items.Add(child.DisplayName())
                    End If
                Else
                    For Each row As DataGridViewRow In titleBlockComps.Rows
                        If (row.cells(0).value.ToString = componentName And row.cells(1).value.ToString.Trim = componentDesc) Then
                            doesExistInTitleBlockGrid = True
                            Exit For
                        End If
                    Next

                    If (doesExistInTitleBlockGrid = False) Then
                        For Each row As DataGridViewRow In assyComps.Rows
                            If (row.cells(0).value.ToString.Trim = componentName And row.cells(1).value.ToString.Trim = componentDesc) Then
                                doesExistInAssyCompGrid = True
                                Exit For
                            End If
                        Next

                        If (doesExistInAssyCompGrid = False) Then
                            If (componentName.Contains("#") = False And componentName.Contains("SHCS") = False And componentName.Contains("DOW") = False And componentName.Contains("WAT") = False And componentName.Contains("AIR") = False And componentName.Contains("CD") = False And componentName.Contains("HR") = False And componentName.Contains("PLATE") = False And componentName.Contains("MIS") = False And componentName.Contains("MS") = False And componentName.Contains("STK") = False And componentName.Contains("ELE") = False And componentName.Contains("SA") = False And componentName.Contains("EJ ") = False And componentName.Contains("E2E") = False And componentName.Contains("EYB") = False And componentName.Contains("PUR") = False And componentName.Trim.Length < 10) Then
                                assyComps.rows.Add(New String() {componentName, componentDesc})
                            End If
                        End If
                    End If
                End If

                Try
                    fillListBoxAndDataGrids(child, indent + 1, parentName + child.DisplayName())
                Catch ex As Exception
                End Try
            Next
        End Sub
        Public Sub AssignApplicationPartAttributes(ByVal partName As String)
            Try
                Dim c As ComponentAssembly = dispPart.ComponentAssembly
                If Not IsNothing(c.RootComponent) Then
                    readAppPartAttributes(c.RootComponent, 0, partName)
                Else
                    lw.WriteLine("Part has no components")
                End If
            Catch e As Exception
                lw.WriteLine("Failed: " & e.ToString)
            End Try
        End Sub
        Public Sub readAppPartAttributes(ByVal comp As Component, ByVal indent As Integer, ByVal partName As String) 'comp = componentAssembly.RootComponent
            For Each child As Component In comp.GetChildren()
                If (child.DisplayName() = partName) Then
                    Try
                        'MsgBox(partName)
                        'MsgBox(child.GetStringAttribute("STKASSY_PART_WALLSECTION_BOTTOM"))
                        ' Read all the part attributes and set them to the textboxes
                        txtBoxPartTitle.Text = child.GetStringAttribute("STKASSY_PART_TITLE")
                        txtBoxPartDiameter.Text = child.GetStringAttribute("STKASSY_PART_DIA")
                        txtBoxPartWidth.Text = child.GetStringAttribute("STKASSY_PART_WIDTH")
                        txtBoxPartLength.Text = child.GetStringAttribute("STKASSY_PART_LENGTH")
                        txtBoxPartHeight.Text = child.GetStringAttribute("STKASSY_PART_HEIGHT")
                        txtBoxPartResin.Text = child.GetStringAttribute("STKASSY_PART_RESIN")
                        txtBoxPartShrinkage.Text = child.GetStringAttribute("STKASSY_PART_SHRINKAGE")
                        txtBoxPartProjectedArea.Text = child.GetStringAttribute("STKASSY_PART_PROJ_AREA")
                        txtBoxPartLTRatio.Text = child.GetStringAttribute("STKASSY_PART_LTRATIO")
                        txtBoxPartFlowLength.Text = child.GetStringAttribute("STKASSY_PART_FLOW_LENGTH")
                        txtBoxPartVolToBrim.Text = child.GetStringAttribute("STKASSY_PART_VOLTOBRIM")
                        txtBoxPartWeight.Text = child.GetStringAttribute("STKASSY_PART_WEIGHT")
                        txtBoxPartDensity.Text = child.GetStringAttribute("STKASSY_PART_DENSITY")
                        txtBoxPartAppearance.Text = child.GetStringAttribute("STKASSY_PART_APPEARANCE")
                        txtBoxCustomer.Text = child.GetStringAttribute("STKASSY_PART_CUSTOMER")
                        If (child.GetStringAttribute("STKASSY_PART_CUSTOMER") = "" Or txtBoxCustomer.Text = "") Then
                            txtBoxCustomer.Enabled = True
                        End If
                        txtBoxPartWSSide.Text = child.GetStringAttribute("STKASSY_PART_WALLSECTION_SIDE")
                        txtBoxPartWSBottom.Text = child.GetStringAttribute("STKASSY_PART_WALLSECTION_BOTTOM") 'latest change here
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If
                readAppPartAttributes(child, indent + 1, partName)
            Next
        End Sub
        Public Sub addImpMetDualUnits(ByRef combox As Object)
            combox.Items.Add("METRIC")
            combox.Items.Add("IMPERIAL")
            combox.Items.Add("DUAL")
        End Sub
        Public Sub addAllUnitsLength(ByRef combox As Object)
            combox.Items.Add("mm")
            combox.Items.Add(Chr(34))
            combox.Items.Add("mm-Dual")
            combox.Items.Add(Chr(34) + "-Dual")
        End Sub
        Public Sub addAllUnitsMass(ByRef combox As Object)
            combox.Items.Add("kg")
            combox.Items.Add("lbs")
            combox.Items.Add("kg-Dual")
            combox.Items.Add("lbs-Dual")
        End Sub
        Public Sub addAllUnitsTons(ByRef combox As Object)
            combox.Items.Add("SI Tonnes")
            combox.Items.Add("SI Tonnes-Dual")
            combox.Items.Add("US Tons")
            combox.Items.Add("US Tons-Dual")
        End Sub
        Public Sub addEjectionTypes(ByRef combox As Object)
            combox.Items.Add("AIR EJECTION")
            combox.Items.Add("STRIPPER EJECTION")
            combox.Items.Add("GULL WING")
            combox.Items.Add("EJECTION BOX")
            combox.Items.Add("2 STAGES")
            combox.Items.Add("3 STAGES")
            combox.Items.Add("5 PIECES COLLAPSE")
            combox.Items.Add("UNSCREWING")
        End Sub
        Public Sub PopulateComboBoxes()
            addImpMetDualUnits(comboxPartUnits)
            addImpMetDualUnits(comboxMoldUnits2)
            addImpMetDualUnits(comboxHRUnits)
            addImpMetDualUnits(comboxMachineUnits)

            comboxHardware.Items.Add("METRIC")
            comboxHardware.Items.Add("IMPERIAL")

            addAllUnitsLength(comboxMoldShutHeightUnits)
            addAllUnitsLength(comboxMoldPitchUnits)
            addAllUnitsLength(comboxMoldSizeUnits)
            addAllUnitsLength(comboxQPCModuleShutHeightUnits)
            addAllUnitsMass(comboxMoldWeightUnits)
            addAllUnitsTons(comboxMoldTonnageUnits)
            addAllUnitsLength(comboxEjectionStrokeUnits)
            addAllUnitsLength(comboxXDimUnits)
            addAllUnitsLength(comboxLDimUnits)
            addAllUnitsLength(comboxPDimHotUnits)
            addAllUnitsLength(comboxPDimColdUnits)
            addAllUnitsLength(comboxGateDiaUnits)
            addAllUnitsLength(comboxTieBarDistanceUnits)
            addAllUnitsTons(comboxClampTonnageUnits)
            addAllUnitsLength(comboxClampStrokeUnits)
            addAllUnitsLength(comboxMaxDaylightUnits)
            addAllUnitsLength(comboxLocatingRingDiameterUnits)
            addAllUnitsLength(comboxNozzleRadiusUnits)
            addAllUnitsLength(comboxMaxEjectorStrokeUnits)
            addAllUnitsLength(comboxMinMaxShutHeightUnits)

            addEjectionTypes(comboxEjectionType1)
            addEjectionTypes(comboxEjectionType2)

            comboxStackInterlock.Items.Add("CAVITY LOCK")
            comboxStackInterlock.Items.Add("CORE LOCK")
            comboxStackInterlock.Items.Add("FLAT")
            comboxStackInterlock.Items.Add("SPLITS")
            comboxStackInterlock.Items.Add("STRIPPER LOCK/BUMP OFF")
            comboxStackInterlock.Items.Add("WEDGE LOCK")

            comboxGateSide.Items.Add("CAVITY SIDE")
            comboxGateSide.Items.Add("CORE SIDE")

            comboxGateType.Items.Add("HOT TIP")
            comboxGateType.Items.Add("EDGE GATES")
            comboxGateType.Items.Add("VALVE GATES")

            comboxHRManufacturer.Items.Add("STACKTECK")
            comboxHRManufacturer.Items.Add("HUSKY")
            comboxHRManufacturer.Items.Add("MOLDMASTERS")
            comboxHRManufacturer.Items.Add("YUDO")
            comboxHRManufacturer.Items.Add("OTHER")
        End Sub
        Private Sub loadDefaultValues()
            AssignJobNum()
            AssignTeamAndUser()
            AssignDate()
            PopulateComboBoxes()

            comboxMoldUnits2.Text = "IMPERIAL"
            comboxPartUnits.Text = "IMPERIAL"
            comboxHardware.Text = "IMPERIAL"
            comboxHRUnits.Text = "IMPERIAL"
            comboxMachineUnits.Text = "METRIC"

        End Sub
        Private Sub CorrectRowSizes()
            Try
                For Each row As DataGridViewRow In assyComps.Rows
                    row.Height = 16
                Next
                For Each row As DataGridViewRow In titleBlockComps.Rows
                    row.Height = 16
                Next
                For Each row As DataGridViewRow In compInfo.Rows
                    row.Height = 16
                Next
            Catch ex As Exception
            End Try
        End Sub

        Public Sub ReadAllAttributesFromStackOrMoldOrSpecification(ByVal stackOrMoldOrSpec As String)
            btnReset_Click(btnReset, New EventArgs())
            Dim ComponentNames(0) As String
            Dim ComponentCounts As Integer = 0
            Dim filter As String = JobNumber
            GetAssemblyTree(ComponentNames, ComponentCounts, filter)
            Dim stackName As String = Nothing
            Dim stackSpecName As String = Nothing
            Dim moldName As String = Nothing


            For i As Integer = 0 To ComponentNames.Length - 1
                If (stackOrMoldOrSpec <> "mold") Then
                    If (ComponentNames(i).Contains("ASSY_STA")) Then
                        If (OpenExistingPart(ComponentNames(i)) = False) Then
                            MsgBox("Could Not open:  " + ComponentNames(i))
                            Exit Sub
                        End If

                        stackName = s.Parts.Display.FullPath
                        lw.WriteLine("StackName: " + stackName)

                        If (stackOrMoldOrSpec = "spec") Then
                            Dim UGPartName As String = Nothing
                            FindSpecDwgofMaster(UGPartName)

                            If UGPartName <> Nothing Then
                                OpenExistingSpecPart(ComponentNames(i), UGPartName)
                                stackSpecName = s.Parts.Display.FullPath
                                lw.WriteLine("StackSpecName: " + stackSpecName)
                            End If
                        End If

                        ReadAllAttributesFromDisplayedPart()
                        Exit For
                    End If
                ElseIf (stackOrMoldOrSpec = "mold") Then
                    'MsgBox(ComponentNames(i).Remove(ComponentNames(i).Length - 4))
                    If (ComponentNames(i).Remove(ComponentNames(i).Length - 4) = JobNumber.Trim + "_ASSY") Then
                        'MsgBox("In Here!")
                        If (OpenExistingPart(ComponentNames(i)) = False) Then
                            MsgBox("Could not open: " + ComponentNames(i))
                            Exit Sub
                        End If

                        moldName = s.Parts.Display.FullPath
                        lw.WriteLine("MoldName: " + moldName)

                        ReadAllAttributesFromDisplayedPart()
                        Exit For
                    End If
                End If
            Next

            Dim theSession As Session = Session.GetSession()
            Dim displayPart As Part

            ChangeDisplayedPartToMoldAssy()
            populateListBoxAndDataGrids()
            CorrectRowSizes()
            Try
                workPart.DeleteRetainedDraftingObjectsInCurrentLayout()
            Catch ex As Exception
            End Try

        End Sub

        Public Sub ClosePart(ByVal stackName As String, ByVal stackOrMoldOrSpec As String)
            Dim theSession As Session = Session.GetSession()
            Dim displayPart As Part

            displayPart = theSession.Parts.Display
            Dim workPart As Part = theSession.Parts.Work

            Dim markId2 As Session.UndoMarkId
            markId2 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Change Display Part")

            Dim Part2 As Part

            If (stackOrMoldOrSpec = "spec") Then
                Part2 = CType(theSession.Parts.FindObject("@DB/" + stackName + "/specification/dwg"), Part)
            End If

            If (stackOrMoldOrSpec = "stack") Then
                Part2 = CType(theSession.Parts.FindObject("@DB/" + stackName), Part)
            End If

            Dim partLoadStatus2 As PartLoadStatus
            Dim status2 As PartCollection.SdpsStatus
            status2 = theSession.Parts.SetDisplay(Part2, True, False, partLoadStatus2)

            'MsgBox("Made Displayed Part")

            workPart = theSession.Parts.Work
            displayPart = theSession.Parts.Display
            partLoadStatus2.Dispose()
            theSession.Parts.SetWork(workPart)

            Dim partCloseResponses2 As PartCloseResponses
            partCloseResponses2 = theSession.Parts.NewPartCloseResponses()

            workPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.UseResponses, partCloseResponses2)

            workPart = Nothing
            displayPart = Nothing
            partCloseResponses2.Dispose()
        End Sub
        Public Sub ChangeDisplayedPartToMoldAssy()
            Try
                Dim theSession As Session = Session.GetSession()
                Dim displayPart As Part

                ' Mold specifications are always in DWG, so we can hardcode the switching back

                Dim markId3 As Session.UndoMarkId
                markId3 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Change Display Part")

                'MsgBox("Now change the displayed part to: " + fileName.replace(" (specification: dwg", "/specification/dwg").replace("dwg)", "dwg"))
                Dim part3 As Part = CType(theSession.Parts.FindObject("@DB/" + fileName.replace(" (specification: dwg", "/specification/dwg").replace("dwg)", "dwg")), Part)

                Dim partLoadStatus2 As PartLoadStatus
                Dim status1 As PartCollection.SdpsStatus
                status1 = s.Parts.SetDisplay(part3, True, True, partLoadStatus2)
                s.Parts.SetWork(part3)

                workPart = theSession.Parts.Work
                displayPart = theSession.Parts.Display
                partLoadStatus2.Dispose()
            Catch ex As Exception
                MsgBox("Error switching back to mold assy. Please notify System Admin. You may continue using the progam." + nl + ex.ToString)
            End Try
        End Sub
        Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            loadDefaultValues()
            ReadAllAttributesFromDisplayedPart()
            populateListBoxAndDataGrids()
            CorrectRowSizes()
            workPart.DeleteRetainedDraftingObjectsInCurrentLayout()
        End Sub

        Function OpenExistingPart(ByVal FileName As String) As Boolean
            Dim basePart1 As BasePart
            Dim partLoadStatus1 As PartLoadStatus = Nothing

            'lw.WriteLine("File Name: " & FileName)
            'Dim Testarray As String() = FileName.Split("/")
            'lw.WriteLine("Test Array: " & Testarray.Length)
            'If Testarray.Length > 2 Then
            '    Return False
            '    Exit Function
            'End If

            ' If the part is not open, try to open it
            Try
                basePart1 = s.Parts.OpenBaseDisplay("@DB/" & FileName, partLoadStatus1)
            Catch ex As NXException
                lw.WriteLine(ex.ToString)
            End Try

            Dim markId3 As Session.UndoMarkId
            markId3 = s.SetUndoMark(Session.MarkVisibility.Visible, "Change Display Part")

            ' If the part is open, switch the window to it
            Dim part1 As Part = CType(s.Parts.FindObject("@DB/" & FileName), Part)

            Dim partLoadStatus2 As PartLoadStatus = Nothing
            Dim status1 As PartCollection.SdpsStatus
            status1 = s.Parts.SetDisplay(part1, True, True, partLoadStatus2)
            s.Parts.SetWork(part1)
            Return True
        End Function
        Public Sub FindSpecDwgofMaster(ByRef UGPartName As String)

            Dim myPartTag As Tag = ufs.Part.AskDisplayPart

            Dim EncodedName As String = Nothing

            Try
                ufs.Part.AskPartName(myPartTag, EncodedName)
            Catch ex As Exception
                lw.WriteLine("Trouble asking part name!")
            End Try

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
                            Case Else
                                Continue For
                        End Select
                    Next
                    Exit Sub
                End If
            Next
        End Sub
        Sub OpenExistingSpecPart(ByVal FileName As String, ByVal UGPartName As String)
            Dim basePart1 As BasePart
            Dim partLoadStatus1 As PartLoadStatus = Nothing

            lw.WriteLine("File Name: " & FileName)
            Dim Testarray As String() = FileName.Split("/")
            'lw.WriteLine("Test Array: " & Testarray.Length)
            If Testarray.Length > 2 Then
                Exit Sub
            End If

            Try
                ' File already exists
                lw.WriteLine("@DB/" & FileName & "/specification/" & UGPartName)
                basePart1 = s.Parts.OpenBaseDisplay("@DB/" & FileName & "/specification/" & UGPartName, partLoadStatus1)
            Catch ex As NXException
                lw.WriteLine("Something went wrong here!")
                'ex.AssertErrorCode(1020004)
            End Try

            Dim markId3 As Session.UndoMarkId
            markId3 = s.SetUndoMark(Session.MarkVisibility.Visible, "Change Display Part")

            Dim part1 As Part = CType(s.Parts.FindObject("@DB/" & FileName & "/specification/" & UGPartName), Part)

            Dim partLoadStatus2 As PartLoadStatus = Nothing
            Dim status1 As PartCollection.SdpsStatus
            status1 = s.Parts.SetDisplay(part1, True, True, partLoadStatus2)

            s.Parts.SetWork(part1)
        End Sub
        Sub GetAssemblyTree(ByRef ComponentNames As String(), ByRef ComponentCounts As Integer, ByVal JobNum As String)

            Dim part1 As Part
            part1 = s.Parts.Work

            Dim c As Component = part1.ComponentAssembly.RootComponent

            Dim count As Integer = 1
            Dim TempComponentName(0) As String
            TempComponentName(0) = c.DisplayName

            ShowAssemblyTree(c, "", count, TempComponentName)

            lw.WriteLine("Total Count: " & count)
            lw.WriteLine("")

            ComponentNames(0) = c.DisplayName
            Dim i As Integer = 0

            For Each Tempstr As String In TempComponentName
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
                    'lw.WriteLine(Tempstr)
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
                    'lw.WriteLine(newIndent & child.DisplayName)
                    ReDim Preserve TempComponentName(count - 1)
                    TempComponentName(count - 1) = child.DisplayName

                    ShowAssemblyTree(child, newIndent, count, TempComponentName)
                Else
                    'lw.WriteLine(newIndent & child.DisplayName & "is in suppress status")
                End If
            Next
        End Sub
        Private Sub btnLoadStack_Click(sender As Object, e As EventArgs) Handles btnLoadStack.Click
            ReadAllAttributesFromStackOrMoldOrSpecification("stack")
        End Sub

        Private Sub btnLoadSpec_Click(sender As Object, e As EventArgs) Handles btnLoadSpec.Click
            ReadAllAttributesFromStackOrMoldOrSpecification("spec")
        End Sub

        Private Sub btnLoadMold_Click(sender As Object, e As EventArgs) Handles btnLoadMold.Click
            ReadAllAttributesFromStackOrMoldOrSpecification("mold")
        End Sub

        Public Sub ReadAllAttributesFromDisplayedPart()
            Dim ComponentName As String = Nothing
            Dim ComponentDesc As String = Nothing
            Dim ComponentMat As String = Nothing
            Dim ComponentHard As String = Nothing
            Dim componentSurfEnh As String = Nothing

            ' General Info
            ReadAttribute("STKASSY_INFO_DRAWN_BY", txtBoxDesigner.Text)
            ReadAttribute("STKASSY_INFO_TEAM", txtBoxDesignTeam.Text)
            ReadAttribute("STKASSY_INFO_DATE", txtBoxDate.Text)
            ReadAttribute("STKASSY_INFO_DRAWING_NUMBER", txtBoxDrawingNumber.Text)
            ReadAttribute("STKASSY_INFO_SCALE", txtBoxScale.Text)
            ReadAttribute("STKASSY_INFO_CURRENT_SHEET", txtBoxCurrentSheet.Text)
            ReadAttribute("STKASSY_INFO_TOTAL_SHEETS", txtBoxTotalSheets.Text)
            ReadAttribute("STKASSY_INFO_PROJECT_NUMBER", txtBoxJobNumber.Text)
            ReadAttribute("STKASSY_INFO_LAST_TAB", LastTabPage)

            If (txtBoxJobNumber.Text = "") Then
                ' Read the filename of the assembly, use it to populate the job number textbox
                Dim fileName As String = s.Parts.Work.FullPath()
                fileName = Microsoft.VisualBasic.Left(fileName, 5)
                txtBoxJobNumber.Text = fileName.ToString()
            End If

            'Open the form to the last opened tab 
            Try
                Tabs.SelectedIndex = Integer.Parse(LastTabPage)
            Catch ex As Exception
            End Try

            'Part Info
            ReadAttribute("STKASSY_INFO_PART_CUSTOMER", txtBoxCustomer.Text)
            ReadAttribute("STKASSY_INFO_PART_UNITS", comboxPartUnits.Text)
            ReadAttribute("STKASSY_INFO_PART_TITLE", txtBoxPartTitle.Text)
            ReadAttribute("STKASSY_INFO_PART_RESIN", txtBoxPartResin.Text)
            ReadAttribute("STKASSY_INFO_PART_SHRINKAGE", txtBoxPartShrinkage.Text)
            ReadAttribute("STKASSY_INFO_PART_WEIGHT", txtBoxPartWeight.Text)
            ReadAttribute("STKASSY_INFO_PART_DENSITY", txtBoxPartDensity.Text)
            ReadAttribute("STKASSY_INFO_PART_APPEARANCE", txtBoxPartAppearance.Text)
            ReadAttribute("STKASSY_INFO_PART_VOL_TO_BRIM", txtBoxPartVolToBrim.Text)
            ReadAttribute("STKASSY_INFO_PART_PROJECTED_AREA", txtBoxPartProjectedArea.Text)
            ReadAttribute("STKASSY_INFO_PART_LTRATIO", txtBoxPartLTRatio.Text)
            ReadAttribute("STKASSY_INFO_PART_WL_GENERAL_STRING", PartWidthAndLengthGeneralString)
            ReadAttribute("STKASSY_INFO_PART_WL_METRIC_STRING", PartWidthAndLengthMetricString)
            ReadAttribute("STKASSY_INFO_PART_WL_IMPERIAL_STRING", PartWidthAndLengthImperialString)
            ReadPart("STKASSY_INFO_PART_FLOW_LENGTH", PartFlowLength, PartFlowLengthGeneral, PartFlowLengthMetric, PartFlowLengthImperial, txtBoxPartFlowLength.Text)
            ReadPart("STKASSY_INFO_PART_DIAMETER", PartDiameter, PartDiameterGeneral, PartDiameterMetric, PartDiameterImperial, txtBoxPartDiameter.Text)
            ReadPart("STKASSY_INFO_PART_WIDTH", PartWidth, PartWidthGeneral, PartWidthMetric, PartWidthImperial, txtBoxPartWidth.Text)
            ReadPart("STKASSY_INFO_PART_LENGTH", PartLength, PartLengthGeneral, PartLengthMetric, PartLengthImperial, txtBoxPartLength.Text)
            ReadPart("STKASSY_INFO_PART_HEIGHT", PartHeight, PartHeightGeneral, PartHeightMetric, PartHeightImperial, txtBoxPartHeight.Text)
            ReadPart("STKASSY_INFO_PART_WALL_SECT_SIDE", PartSideWS, PartSideWSGeneral, PartSideWSMetric, PartSideWSImperial, txtBoxPartWSSide.Text)
            ReadPart("STKASSY_INFO_PART_WALL_SECT_BOTTOM", PartBottomWS, PartBottomWSGeneral, PartBottomWSMetric, PartBottomWSImperial, txtBoxPartWSBottom.Text)
            ' Component Info
            For i As Integer = 0 To 8
                Try
                    ReadAttribute("STKASSY_INFO_COMPONENT_" + (i + 1).ToString, ComponentName)
                    ReadAttribute("STKASSY_INFO_COMPONENT_DESCRIPTION" + (i + 1).ToString, ComponentDesc)
                    ReadAttribute("STKASSY_INFO_COMPONENT_MATERIAL" + (i + 1).ToString, ComponentMat)
                    ReadAttribute("STKASSY_INFO_COMPONENT_HARDNESS" + (i + 1).ToString, ComponentHard)
                    ReadAttribute("STKASSY_INFO_COMPONENT_SURFACE_ENH" + (i + 1).ToString, componentSurfEnh)
                    If (ComponentName <> "" And ComponentDesc <> "") Then

                        ' Need to add backwards compatibility here
                        If (IsNumeric(ComponentName)) Then
                            titleBlockComps.Rows.Add(New String() {ComponentName, ComponentDesc})
                            compInfo.Rows.Add(New String() {ComponentDesc, ComponentMat, ComponentHard, componentSurfEnh, ComponentName})
                        ElseIf (ComponentName.Length > 26) Then

                            'MsgBox(ComponentName)
                            ComponentName = ComponentName.Substring(0, 26).Trim
                            titleBlockComps.Rows.Add(New String() {ComponentName, ComponentDesc})
                            compInfo.Rows.Add(New String() {ComponentDesc, ComponentMat, ComponentHard, componentSurfEnh, ComponentName})
                            'MsgBox(ComponentName)

                        End If
                    End If
                Catch ex As Exception
                End Try
            Next

            'Mold Info
            ReadAttribute("STKASSY_INFO_MOLD_UNITS", comboxMoldUnits2.Text)
            ReadAttribute("STKASSY_INFO_MOLD_UNITS", OldMoldUnits)
            ReadAttribute("STKASSY_INFO_MOLD_DESCRIPTION", txtBoxMoldDescription.Text)
            ReadAttribute("STKASSY_INFO_MOLD_CAV_1", txtBoxMoldCavitation1.Text)
            ReadAttribute("STKASSY_INFO_MOLD_CAV_2", txtBoxMoldCavitation2.Text)

            ReadMoldOrHR("STKASSY_INFO_MOLD_SHUT_HEIGHT", MoldTotalShutHeightGeneralValue, MoldTotalShutHeightGeneralUnit, MoldTotalShutHeightMetricValue, MoldTotalShutHeightImperialValue, MoldTotalShutHeightDualUnit, txtBoxMoldShutHeight.Text, comboxMoldShutHeightUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_MOLD_TONNAGE", MoldTonnageGeneralValue, MoldTonnageGeneralUnit, MoldTonnageMetricValue, MoldTonnageImperialValue, MoldTonnageDualUnit, txtBoxMoldTonnage.Text, comboxMoldTonnageUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_MOLD_QPC", MoldQPCGeneralValue, MoldQPCGeneralUnit, MoldQPCMetricValue, MoldQPCImperialValue, MoldQPCDualUnit, txtBoxQPCModuleShutHeight.Text, comboxQPCModuleShutHeightUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_EST_MOLD_WEIGHT", MoldWeightGeneralValue, MoldWeightGeneralUnit, MoldWeightMetricValue, MoldWeightImperialValue, MoldWeightDualUnit, txtBoxMoldWeight.Text, comboxMoldWeightUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_MOLD_EJECTION_STROKE", MoldEjectionStrokeGeneralValue, MoldEjectionStrokeGeneralUnit, MoldEjectionStrokeMetricValue, MoldEjectionStrokeImperialValue, MoldEjectionStrokeDualUnit, txtBoxEjectionStroke.Text, comboxEjectionStrokeUnits.Text)
            ReadMoldPitch("STKASSY_INFO_MOLD_PITCH", MoldPitchXGeneralValue, MoldPitchYGeneralValue, MoldPitchGeneralUnit, MoldPitchXMetricValue, MoldPitchYMetricValue, MoldPitchXImperialValue, MoldPitchYImperialValue, MoldPitchDualUnit, txtBoxMoldPitchX.Text, txtBoxMoldPitchY.Text, comboxMoldPitchUnits.Text)
            ReadMoldSize("STKASSY_INFO_MOLD_SIZE", MoldSizeXGeneralValue, MoldSizeYGeneralValue, MoldSizeXYGeneralUnit, MoldSizeXMetricValue, MoldSizeYMetricValue, MoldSizeXImperialValue, MoldSizeYImperialValue, MoldSizeXYDualUnit, txtBoxMoldSizeX.Text, txtBoxMoldSizeY.Text, comboxMoldSizeUnits.Text)

            ReadAttribute("STKASSY_INFO_MOLD_STACK_INTERLOCK", comboxStackInterlock.Text)
            ReadAttribute("STKASSY_INFO_MOLD_EJECTION1", comboxEjectionType1.Text)
            ReadAttribute("STKASSY_INFO_MOLD_EJECTION2", comboxEjectionType2.Text)
            ReadAttribute("STKASSY_INFO_MOLD_MAX_OPENING", txtBoxMaxMoldOpeningPerSide.Text)

            MoldUnitsReadYet = True

            ReadAttribute("STKASSY_INFO_HARDWARE", comboxHardware.Text)

            ' HR Info 
            ReadAttribute("STKASSY_INFO_HR_UNITS", comboxHRUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_HR_XDIM", HotRunnerXDimGeneralValue, HotRunnerXDimGeneralUnit, HotRunnerXDimMetricValue, HotRunnerXDimImperialValue, HotRunnerXDimDualUnit, txtBoxXDim.Text, comboxXDimUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_HR_LDIM", HotRunnerLDimGeneralValue, HotRunnerLDimGeneralUnit, HotRunnerLDimMetricValue, HotRunnerLDimImperialValue, HotRunnerLDimDualUnit, txtBoxLDim.Text, comboxLDimUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_HR_PDIMHOT", HotRunnerPDimHotGeneralValue, HotRunnerPDimHotGeneralUnit, HotRunnerPDimHotMetricValue, HotRunnerPDimHotImperialValue, HotRunnerPDimHotDualUnit, txtBoxPDimHot.Text, comboxPDimHotUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_HR_PDIMCOLD", HotRunnerPDimColdGeneralValue, HotRunnerPDimColdGeneralUnit, HotRunnerPDimColdMetricValue, HotRunnerPDimColdImperialValue, HotRunnerPDimColdDualUnit, txtBoxPDimCold.Text, comboxPDimColdUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_HR_GATEDIA", HotRunnerGateDiameterGeneralValue, HotRunnerGateDiameterGeneralUnit, HotRunnerGateDiameterMetricValue, HotRunnerGateDiameterImperialValue, HotRunnerGateDiameterDualUnit, txtBoxGateDia.Text, comboxGateDiaUnits.Text)

            ReadAttribute("STKASSY_INFO_HR_NOZZLE_TIP_NUMBER", txtBoxNozzleTipPartNumber.Text)
            ReadAttribute("STKASSY_INFO_HR_GATE_SIDE", comboxGateSide.Text)
            ReadAttribute("STKASSY_INFO_HR_GATE_TYPE", comboxGateType.Text)
            ReadAttribute("STKASSY_INFO_HR_MANUFACTURER", comboxHRManufacturer.Text)

            ' Machine Data 
            ReadAttribute("STKASSY_INFO_MACHINE_UNITS", comboxMachineUnits.Text)
            ReadAttribute("STKASSY_INFO_MACHINE_MODEL", txtBoxMoldMachineModel.Text)
            ReadTieBar("STKASSY_INFO_TIEBAR", MachineTieBarH, MachineTieBarHMetric, MachineTieBarHImperial, MachineTieBarV, MachineTieBarVMetric, MachineTieBarVImperial, MachineTieBarLastUnit, MachineTieBarUnitsMetric, MachineTieBarUnitsImperial, MachineTieBarGeneralString, MachineTieBarMetricString, MachineTieBarImperialString, MachineTieBarDualUnit, txtBoxTieBarDistanceHorizontal.Text, txtBoxTieBarDistanceVertical.Text, comboxTieBarDistanceUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_CLAMP_TONNAGE", MachineClampTonnageGeneralValue, MachineClampTonnageGeneralUnit, MachineClampTonnageMetricValue, MachineClampTonnageImperialValue, MachineClampTonnageDualUnit, txtBoxClampTonnage.Text, comboxClampTonnageUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_CLAMP_STROKE", MachineClampStrokeGeneralValue, MachineClampStrokeGeneralUnit, MachineClampStrokeMetricValue, MachineClampStrokeImperialValue, MachineClampStrokeDualUnit, txtBoxClampStroke.Text, comboxClampStrokeUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_MAX_DAYLIGHT", MachineMaxDaylightGeneralValue, MachineMaxDaylightGeneralUnit, MachineMaxDaylightMetricValue, MachineMaxDaylightImperialValue, MachineMaxDaylightDualUnit, txtBoxMaxDaylight.Text, comboxMaxDaylightUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_LOCATING_RING_DIA", MachineLocatingRingDiameterGeneralValue, MachineLocatingRingDiameterGeneralUnit, MachineLocatingRingDiameterMetricValue, MachineLocatingRingDiameterImperialValue, MachineLocatingRingDiameterDualUnit, txtBoxLocatingRingDiameter.Text, comboxLocatingRingDiameterUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_NOZZLE_RADIUS", MachineNozzleRadiusGeneralValue, MachineNozzleRadiusGeneralUnit, MachineNozzleRadiusMetricValue, MachineNozzleRadiusImperialValue, MachineNozzleRadiusDualUnit, txtBoxNozzleRadius.Text, comboxNozzleRadiusUnits.Text)
            ReadMoldOrHR("STKASSY_INFO_MAX_EJECTOR_STROKE", MachineMaxEjectorStrokeGeneralValue, MachineMaxEjectorStrokeGeneralUnit, MachineMaxEjectorStrokeMetricValue, machineMaxEjectorStrokeImperialValue, MachineMaxEjectorStrokeDualUnit, txtBoxMaxEjectorStroke.Text, comboxMaxEjectorStrokeUnits.Text)
            ReadMinMaxShutHeight("STKASSY_INFO_MIN_SHUT_HEIGHT", "STKASSY_INFO_MAX_SHUT_HEIGHT", "STKASSY_INFO_SHUT_HEIGHT", MachineMinShutHeightGeneralValue, MachineMinShutHeightMetricValue, MachineMinShutHeightImperialValue, MachineMaxShutHeightGeneralValue, MachineMaxShutHeightMetricValue, MachineMaxShutHeightImperialValue, MachineShutHeightGeneralUnit, MachineShutHeightDualUnit, txtBoxMinShutHeight.Text, txtBoxMaxShutHeight.Text, comboxMinMaxShutHeightUnits.Text)

            txtBoxDrawingNumber.Text = "00006"
        End Sub
        Public Sub ReadPart(ByVal name As String, ByRef value As String, ByRef gen As String, ByRef met As String, ByRef imp As String, ByRef txt As String)
            ReadAttribute(name, value)
            ReadAttribute(name + "_GENERAL", gen)
            ReadAttribute(name + "_METRIC", met)
            ReadAttribute(name + "_IMPERIAL", imp)
            If (gen <> "") Then
                txt = value
            ElseIf (imp <> "" And met <> "" And gen = "") Then
                txt = imp
                txt = txt.Replace(Chr(34), "")
            End If
        End Sub
        Public Sub ReadMoldOrHR(ByVal name As String, ByRef genVal As String, ByRef genUnit As String, ByRef metVal As String, ByRef impVal As String, ByRef dualUnit As String, ByRef txt As String, ByRef units As String)
            Try
                ReadAttribute(name + "_GENERAL_VALUE", genVal)
                ReadAttribute(name + "_GENERAL_UNIT", genUnit)
                ReadAttribute(name + "_METRIC_VALUE", metVal)
                ReadAttribute(name + "_IMPERIAL_VALUE", impVal)
                ReadAttribute(name + "_DUAL_UNIT", dualUnit)
                If (genUnit <> "") Then
                    units = genUnit
                Else
                    units = dualUnit
                End If
                If (genVal <> "") Then
                    txt = genVal
                ElseIf (dualUnit = "mm-Dual" Or dualUnit = "SI Tonnes-Dual" Or dualUnit = "kg-Dual") Then
                    txt = metVal
                ElseIf (dualUnit = Chr(34) + "-Dual" Or dualUnit = "US Tons-Dual" Or dualUnit = "lbs-Dual") Then
                    txt = impVal
                End If
            Catch ex As Exception
            End Try
        End Sub
        Public Sub ReadMoldSize(ByVal name As String, ByRef xGen As String, ByRef yGen As String, ByRef genUnit As String, ByRef xMet As String, ByRef yMet As String, ByRef xImp As String, ByRef yImp As String, ByRef dual As String, ByRef txtX As String, ByRef txtY As String, ByRef units As String)
            ReadAttribute(name + "_X_GENERAL_VALUE", xGen)
            ReadAttribute(name + "_Y_GENERAL_VALUE", yGen)
            ReadAttribute(name + "_X_Y_GENERAL_UNIT", genUnit)
            ReadAttribute(name + "_X_METRIC_VALUE", xMet)
            ReadAttribute(name + "_Y_METRIC_VALUE", yMet)
            ReadAttribute(name + "_X_IMPERIAL_VALUE", xImp)
            ReadAttribute(name + "_Y_IMPERIAL_VALUE", yImp)
            ReadAttribute(name + "_X_Y_DUAL_UNIT", dual)
            ' Functionality for TxtBoxLogic

            If (genUnit <> "") Then
                units = genUnit
            Else
                units = dual
            End If

            If (xGen <> "" Or yGen <> "") Then ' Single Dimension Case
                txtX = xGen
                txtY = yGen
            ElseIf (dual = "mm-Dual") Then ' Dual Dimension metric case
                txtX = xMet
                txtY = yMet
            ElseIf (dual = Chr(34) + "-Dual") Then 'Dual Dimension imperial case
                txtX = xImp
                txtY = yImp
            End If
        End Sub

        Public Sub ReadMoldPitch(ByVal name As String, ByRef gX As String, ByRef gY As String, ByRef gU As String, ByRef mX As String, ByRef mY As String, ByRef iX As String, ByRef iY As String, ByRef dual As String, ByRef txtX As String, ByRef txtY As String, ByRef units As String)
            Try
                ReadAttribute(name + "_X_GENERAL_VALUE", gX)
                ReadAttribute(name + "_Y_GENERAL_VALUE", gY)
                ReadAttribute(name + "_GENERAL_UNIT", gU)
                ReadAttribute(name + "_X_METRIC_VALUE", mX)
                ReadAttribute(name + "_X_IMPERIAL_VALUE", iX)
                ReadAttribute(name + "_Y_METRIC_VALUE", mY)
                ReadAttribute(name + "_Y_IMPERIAL_VALUE", iY)
                ReadAttribute(name + "_DUAL_UNIT", dual)
                If (gU <> "") Then
                    units = gU
                Else
                    units = dual
                End If

                If (gX <> "" Or gY <> "") Then ' Single dimension case 
                    txtX = gX
                    txtY = gY
                ElseIf (dual = "mm-Dual") Then
                    txtX = mX
                    txtY = mY
                ElseIf (dual = Chr(34) + "-Dual") Then
                    txtX = iX
                    txtY = iY
                End If
            Catch ex As Exception
            End Try
        End Sub
        Public Sub ReadMinMaxShutHeight(ByVal minName As String, ByVal maxName As String, ByVal genName As String, ByRef minGen As String, ByRef minMet As String, ByRef minImp As String, ByRef maxGen As String, ByRef maxMet As String, ByRef maxImp As String, ByRef unit As String, ByRef dual As String, ByRef txtMin As String, ByRef txtMax As String, ByRef units As String)
            ReadAttribute(minName + "_GENERAL_VALUE", minGen)
            ReadAttribute(genName + "_GENERAL_UNIT", unit)
            ReadAttribute(minName + "_METRIC_VALUE", minMet)
            ReadAttribute(minName + "_IMPERIAL_VALUE", minImp)
            ReadAttribute(maxName + "_GENERAL_VALUE", maxGen)
            ReadAttribute(maxName + "_METRIC_VALUE", maxMet)
            ReadAttribute(maxName + "_IMPERIAL_VALUE", maxImp)
            ReadAttribute(genName + "_DUAL_UNIT", dual)
            If (unit <> "") Then
                units = unit
            Else
                units = dual
            End If

            If (minGen <> "" Or maxGen <> "") Then ' Only single dimension case
                txtMin = minGen
                txtMax = maxGen
            ElseIf (dual = "mm-Dual") Then ' Dual unit, original was mm
                txtMin = minMet
                txtMax = maxMet
            ElseIf (dual = Chr(34) + "-Dual") Then ' Dual unit, original was inches
                txtMin = minImp
                txtMax = maxImp
            End If
        End Sub
        Public Sub ReadTieBar(ByVal name As String, ByRef horz As String, ByRef horzMet As String, ByRef horzImp As String, ByRef vert As String, ByRef vertMet As String, ByRef vertImp As String, ByRef units As String, ByRef unitsMet As String, ByRef unitsImp As String, ByRef gen As String, ByRef met As String, ByRef imp As String, ByRef dual As String, ByRef txtH As String, ByRef txtV As String, ByRef comboxUnits As String)
            ReadAttribute(name + "_HORIZONTAL", horz)
            ReadAttribute(name + "_HORIZONTAL_METRIC", horzMet)
            ReadAttribute(name + "_HORIZONTAL_IMPERIAL", horzImp)
            ReadAttribute(name + "_VERTICAL", vert)
            ReadAttribute(name + "_VERTICAL_METRIC", vertMet)
            ReadAttribute(name + "_VERTICAL_IMPERIAL", vertImp)
            ReadAttribute(name + "_UNITS", units) ' The last unit entered in 
            ReadAttribute(name + "_UNITS_METRIC", unitsMet)
            ReadAttribute(name + "_UNITS_IMPERIAL", unitsImp)
            ReadAttribute(name + "_GENERAL_STRING", gen)
            ReadAttribute(name + "_METRIC_STRING", met)
            ReadAttribute(name + "_IMPERIAL_STRING", imp)
            ReadAttribute(name + "_DUAL_UNIT", dual)

            If (units <> "") Then
                comboxUnits = units
            ElseIf (dual <> "") Then
                comboxUnits = dual
            Else
                comboxUnits = "mm"
            End If

            If (horz <> "" Or vert <> "") Then  ' Only single dimensions in title bar, show whatever last unit was
                txtH = horz
                txtV = vert
                units = units
            ElseIf (dual = "mm-Dual") Then ' Dual dimensioning is true, read & show metric value in program
                txtH = horzMet
                txtV = vertMet
            ElseIf (dual = Chr(34) + "-Dual") Then
                txtH = horzImp
                txtV = vertImp
            End If
        End Sub
        Private Sub txtBoxDesigner_TextChanged(sender As Object, e As EventArgs) Handles txtBoxDesigner.TextChanged
            If (txtBoxDesigner.Text = "SOREN" Or txtBoxDesigner.Text = "DAVIDD" Or Environment.UserName() = "ssarvestany" Or Environment.UserName() = "davidd") Then
                txtBoxCustomer.Enabled = True
                txtBoxPartTitle.Enabled = True
                txtBoxPartResin.Enabled = True
                txtBoxPartShrinkage.Enabled = True
                txtBoxPartWeight.Enabled = True
                txtBoxPartDensity.Enabled = True
                txtBoxPartAppearance.Enabled = True
                txtBoxPartVolToBrim.Enabled = True
                txtBoxPartProjectedArea.Enabled = True
                txtBoxPartLTRatio.Enabled = True
                txtBoxPartFlowLength.Enabled = True
                txtBoxPartDiameter.Enabled = True
                txtBoxPartWidth.Enabled = True
                txtBoxPartLength.Enabled = True
                txtBoxPartHeight.Enabled = True
                txtBoxPartWSSide.Enabled = True
                txtBoxPartWSBottom.Enabled = True
            Else
                txtBoxPartTitle.Enabled = False
                txtBoxPartResin.Enabled = False
                txtBoxPartShrinkage.Enabled = False
                txtBoxPartWeight.Enabled = False
                txtBoxPartDensity.Enabled = False
                txtBoxPartAppearance.Enabled = False
                txtBoxPartVolToBrim.Enabled = False
                txtBoxPartProjectedArea.Enabled = False
                txtBoxPartLTRatio.Enabled = False
                txtBoxPartFlowLength.Enabled = False
                txtBoxPartDiameter.Enabled = False
                txtBoxPartWidth.Enabled = False
                txtBoxPartLength.Enabled = False
                txtBoxPartHeight.Enabled = False
                txtBoxPartWSSide.Enabled = False
                txtBoxPartWSBottom.Enabled = False
            End If
        End Sub
        Private Sub comboxMachineUnits_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comboxMachineUnits.SelectedIndexChanged
            If (comboxMachineUnits.Text = "METRIC") Then
                AssignMachineTabUnits("mm", "SI Tonnes")
            ElseIf (comboxMachineUnits.Text = "IMPERIAL") Then
                AssignMachineTabUnits(Chr(34), "US Tons")
            ElseIf (comboxMachineUnits.Text = "DUAL") Then
                AssignMachineTabUnits(Chr(34) + "-Dual", "US Tons-Dual")
            End If
        End Sub
        Public Sub AssignMachineTabUnits(ByVal lengthUnit As String, ByVal weightUnit As String)
            comboxClampTonnageUnits.Text = weightUnit
            comboxClampStrokeUnits.Text = lengthUnit
            comboxMaxDaylightUnits.Text = lengthUnit
            comboxLocatingRingDiameterUnits.Text = lengthUnit
            comboxNozzleRadiusUnits.Text = lengthUnit
            comboxMaxEjectorStrokeUnits.Text = lengthUnit
            comboxMinMaxShutHeightUnits.Text = lengthUnit
            comboxTieBarDistanceUnits.Text = lengthUnit
        End Sub
        Private Sub comboxHRUnits_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comboxHRUnits.SelectedIndexChanged
            If (comboxHRUnits.Text = "METRIC") Then
                AssignHRTabUnits("mm", "SI Tonnes")
            ElseIf (comboxHRUnits.Text = "IMPERIAL") Then
                AssignHRTabUnits(Chr(34), "US Tons")
            ElseIf (comboxHRUnits.Text = "DUAL") Then
                AssignHRTabUnits(Chr(34) + "-Dual", "US Tons-Dual")
            End If
        End Sub
        Public Sub AssignHRTabUnits(ByVal lengthUnit As String, ByVal weightUnit As String)
            comboxXDimUnits.Text = lengthUnit
            comboxLDimUnits.Text = lengthUnit
            comboxPDimHotUnits.Text = lengthUnit
            comboxPDimColdUnits.Text = lengthUnit
            comboxGateDiaUnits.Text = lengthUnit
        End Sub
        Private Sub comboxMoldUnits2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles comboxMoldUnits2.SelectedIndexChanged
            If (MoldUnitsReadYet = True And OldMoldUnits = comboxMoldUnits2.Text) Then
                Exit Sub
            End If
            If (comboxMoldUnits2.Text = "METRIC") Then
                AssignMoldTabUnits("mm", "SI Tonnes", "kg")
            ElseIf (comboxMoldUnits2.Text = "IMPERIAL") Then
                AssignMoldTabUnits(Chr(34), "US Tons", "lbs")
            ElseIf (comboxMoldUnits2.Text = "DUAL") Then
                AssignMoldTabUnits(Chr(34) + "-Dual", "US Tons-Dual", "lbs-Dual")
            End If
            OldMoldUnits = comboxMoldUnits2.Text
        End Sub
        Public Sub AssignMoldTabUnits(ByVal lengthUnit As String, ByVal weightUnit As String, ByVal massUnit As String)
            comboxMoldWeightUnits.Text = massUnit
            comboxMoldShutHeightUnits.Text = lengthUnit
            comboxMoldPitchUnits.Text = lengthUnit
            comboxMoldSizeUnits.Text = lengthUnit
            comboxQPCModuleShutHeightUnits.Text = lengthUnit
            comboxMoldTonnageUnits.Text = weightUnit
            comboxEjectionStrokeUnits.Text = lengthUnit
        End Sub
        Private Sub listBoxApplicationsParts_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listBoxApplicationsParts.SelectedIndexChanged
            Dim currIndex = listBoxApplicationsParts.SelectedIndex
            Dim currIndexName = listBoxApplicationsParts.Items.Item(currIndex)
            AssignApplicationPartAttributes(currIndexName)
        End Sub
        Private Sub moveComponentToTitleBlock_Click(sender As Object, e As EventArgs) Handles moveComponentToTitleBlock.Click
            Try
                For Each row As DataGridViewRow In assyComps.SelectedRows
                    If (titleBlockComps.Rows.Count < 9) Then
                        titleBlockComps.Rows.Add(New String() {row.cells(0).value.ToString.Trim, row.cells(1).value.ToString.Trim})
                        assyComps.Rows.Remove(row)
                    End If
                Next
                populateComponentInformationGrid()
            Catch ex As Exception
            End Try
            CorrectRowSizes()
        End Sub
        Private Sub moveComponentFromTitleBlock_Click(sender As Object, e As EventArgs) Handles moveComponentFromTitleBlock.Click
            Try
                For Each row As DataGridViewRow In titleBlockComps.SelectedRows
                    assyComps.Rows.Add(New String() {row.cells(0).value.ToString.Trim, row.cells(1).value.ToString.Trim})
                    titleBlockComps.Rows.Remove(row)
                    For Each row1 As DataGridViewRow In compInfo.Rows
                        If (row.cells(1).Value.ToString.Trim = row1.Cells(0).Value.ToString.Trim) Then
                            compInfo.Rows.Remove(row1)
                        End If
                    Next
                Next
            Catch ex As Exception
            End Try
            CorrectRowSizes()
        End Sub
        Private Function mmToInches(ByVal mmValue As String) As String
            Dim tempNum As Double
            If Double.TryParse(mmValue, tempNum) Then
                tempNum = tempNum * 0.0393701
                mmToInches = Math.Round(tempNum, 2, MidpointRounding.AwayFromZero)
            Else
                mmToInches = ""
            End If
        End Function
        Public Function inchesTomm(ByVal inchesValue As String) As String
            Dim tempNum As Double
            If Double.TryParse(inchesValue, tempNum) Then
                tempNum = tempNum * 25.4
                inchesTomm = Math.Round(tempNum, 2, MidpointRounding.AwayFromZero)
            Else
                inchesTomm = ""
            End If
        End Function
        Private Function kgToPound(ByVal kgValue As String) As String
            Dim tempNum As Double
            If Double.TryParse(kgValue, tempNum) Then
                tempNum = tempNum * 2.20462
                kgToPound = Math.Round(tempNum, 2, MidpointRounding.AwayFromZero)
            End If
        End Function
        Private Function poundToKg(ByVal poundValue As String) As String
            Dim tempNum As Double
            If Double.TryParse(poundValue, tempNum) Then
                tempNum = tempNum * 0.453592
                poundToKg = Math.Round(tempNum, 3, MidpointRounding.AwayFromZero)
            End If
        End Function
        Private Function metTonToImpTon(ByVal tonnesValue As String) As String
            Dim tempNum As Double
            If Double.TryParse(tonnesValue, tempNum) Then
                tempNum = tempNum * 1.10231
                metTonToImpTon = Math.Round(tempNum, 3, MidpointRounding.AwayFromZero)
            End If
        End Function

        Private Function impTonToMetTon(ByVal tonsValue As String) As String
            Dim tempNum As Double
            If Double.TryParse(tonsValue, tempNum) Then
                tempNum = tempNum * 0.907185
                impTonToMetTon = Math.Round(tempNum, 2, MidpointRounding.AwayFromZero)
            End If
        End Function

        Private Sub okButton_Click(sender As Object, e As EventArgs) Handles okButton.Click
            populateComponentInformationGrid()
            setCustomAttributes()
        End Sub

        Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
            setDefaults()
        End Sub

        Private Sub setDefaults()
            lw.WriteLine("SETTING ALL ATTRIBUTES TO THEIR DEFAULT VALUES!")
            SetAttribute("STKASSY_INFO_LAST_TAB", "0")
            SetAttribute("STKASSY_INFO_MOLD_RUN_COUNT", "0")
            SetAttribute("STKASSY_INFO_DRAWN_BY", txtBoxDesigner.Text)
            SetAttribute("STKASSY_INFO_TEAM", txtBoxDesignTeam.Text)
            SetAttribute("STKASSY_INFO_DATE", txtBoxDate.Text)
            SetAttribute("STKASSY_INFO_DRAWING_NUMBER", "") : txtBoxDrawingNumber.Text = ""
            SetAttribute("STKASSY_INFO_SCALE", "") : txtBoxScale.Text = ""
            SetAttribute("STKASSY_INFO_CURRENT_SHEET", "") : txtBoxCurrentSheet.Text = ""
            SetAttribute("STKASSY_INFO_TOTAL_SHEETS", "") : txtBoxTotalSheets.Text = ""
            SetAttribute("STKASSY_INFO_PROJECT_NUMBER", "") : txtBoxJobNumber.Text = ""

            ' Part Information 
            SetAttribute("STKASSY_PART_UNITS", "IMPERIAL")
            SetAttribute("STKASSY_INFO_PART_TITLE", "") : txtBoxPartTitle.Text = ""
            SetAttribute("STKASSY_INFO_PART_RESIN", "") : txtBoxPartResin.Text = ""
            SetAttribute("STKASSY_INFO_PART_SHRINKAGE", "") : txtBoxPartShrinkage.Text = ""
            SetAttribute("STKASSY_INFO_PART_WEIGHT", "") : txtBoxPartWeight.Text = ""
            SetAttribute("STKASSY_INFO_PART_DENSITY", "") : txtBoxPartDensity.Text = ""
            SetAttribute("STKASSY_INFO_PART_APPEARANCE", "") : txtBoxPartAppearance.Text = ""
            SetAttribute("STKASSY_INFO_PART_VOL_TO_BRIM", "") : txtBoxPartVolToBrim.Text = ""
            SetAttribute("STKASSY_INFO_PART_PROJECTED_AREA", "") : txtBoxPartProjectedArea.Text = ""
            SetAttribute("STKASSY_INFO_PART_LTRATIO", "") : txtBoxPartLTRatio.Text = ""
            SetPart("STKASSY_INFO_PART_FLOW_LENGTH", "", "") : txtBoxPartFlowLength.Text = ""
            SetPart("STKASSY_INFO_PART_DIAMETER", "", "") : txtBoxPartDiameter.Text = ""
            SetPart("STKASSY_INFO_PART_WIDTH", "", "") : txtBoxPartWidth.Text = ""
            SetPart("STKASSY_INFO_PART_LENGTH", "", "") : txtBoxPartLength.Text = ""
            SetPart("STKASSY_INFO_PART_HEIGHT", "", "") : txtBoxPartHeight.Text = ""
            SetPart("STKASSY_INFO_PART_WALL_SECT_SIDE", "", "") : txtBoxPartWSSide.Text = ""
            SetPart("STKASSY_INFO_PART_WALL_SECT_BOTTOM", "", "") : txtBoxPartWSBottom.Text = ""
            SetPart("STKASSY_INFO_PART_WL", "", "")
            SetPart("STKASSY_INFO_PART_WL_GENERAL", "", "")
            SetPart("STKASSY_INFO_PART_WALL_SECT_SIDE", "", "")
            SetPart("STKASSY_INFO_PART_WALL_SECT_BOTTOM", "", "")
            SetAttribute("STKASSY_INFO_PART_CUSTOMER", "") : txtBoxCustomer.Text = ""

            ' Component Information 
            For i As Integer = 1 To 9
                SetBlankComponent(i.ToString)
            Next
            assyComps.Rows.Clear()
            compInfo.Rows.Clear()
            titleBlockComps.Rows.Clear()
            listBoxApplicationsParts.Items.Clear()
            populateListBoxAndDataGrids()

            ' Mold Data 
            SetAttribute("STKASSY_INFO_MOLD_DESCRIPTION", "") : txtBoxMoldDescription.Text = ""
            SetMoldCav("STKASSY_INFO_MOLD_CAV", "", "") : txtBoxMoldCavitation1.Text = "" : txtBoxMoldCavitation2.Text = ""
            SetMoldOrHR("STKASSY_INFO_MOLD_SHUT_HEIGHT", "", "") : txtBoxMoldShutHeight.Text = ""
            SetAttribute("STKASSY_INFO_MOLD_MAX_OPENING", "") : txtBoxMaxMoldOpeningPerSide.Text = ""
            SetMoldPitch("STKASSY_INFO_MOLD_PITCH", "", "", "") : txtBoxMoldPitchX.Text = "" : txtBoxMoldPitchY.Text = ""
            SetMoldSize("STKASSY_INFO_MOLD_SIZE", "", "", "") : txtBoxMoldSizeX.Text = "" : txtBoxMoldSizeY.Text = ""
            SetMoldOrHR("STKASSY_INFO_MOLD_TONNAGE", "", "") : txtBoxMoldTonnage.Text = ""
            SetMoldOrHR("STKASSY_INFO_MOLD_QPC", "", "") : txtBoxQPCModuleShutHeight.Text = ""
            SetMoldOrHR("STKASSY_INFO_EST_MOLD_WEIGHT", "", "") : txtBoxMoldWeight.Text = ""
            SetMoldOrHR("STKASSY_INFO_MOLD_EJECTION_STROKE", "", "") : txtBoxEjectionStroke.Text = ""
            SetAttribute("STKASSY_INFO_MOLD_EJECTION1", "") : comboxEjectionType1.Text = ""
            SetAttribute("STKASSY_INFO_MOLD_EJECTION2", "") : comboxEjectionType2.Text = ""
            SetAttribute("STKASSY_INFO_MOLD_STACK_INTERLOCK", "") : comboxStackInterlock.Text = ""
            SetAttribute("STKASSY_INFO_MOLD_UNITS", "") : comboxMoldUnits2.Text = "IMPERIAL"
            SetAttribute("STKASSY_INFO_MACHINE_MODEL", "") : txtBoxMoldMachineModel.Text = ""

            ' Hardware
            SetAttribute("STKASSY_INFO_HARDWARE", "IMPERIAL") : comboxHardware.Text = "IMPERIAL"

            'HR Data
            SetMoldOrHR("STKASSY_INFO_HR_XDIM", "", "") : txtBoxXDim.Text = ""
            SetMoldOrHR("STKASSY_INFO_HR_LDIM", "", "") : txtBoxLDim.Text = ""
            SetMoldOrHR("STKASSY_INFO_HR_PDIMHOT", "", "") : txtBoxPDimHot.Text = ""
            SetMoldOrHR("STKASSY_INFO_HR_PDIMCOLD", "", "") : txtBoxPDimCold.Text = ""
            SetMoldOrHR("STKASSY_INFO_HR_GATEDIA", "", "") : txtBoxGateDia.Text = ""
            SetAttribute("STKASSY_INFO_HR_NOZZLE_TIP_NUMBER", "") : txtBoxNozzleTipPartNumber.Text = ""
            SetAttribute("STKASSY_INFO_HR_GATE_SIDE", "") : comboxGateSide.Text = ""
            SetAttribute("STKASSY_INFO_HR_GATE_TYPE", "") : comboxGateType.Text = ""
            SetAttribute("STKASSY_INFO_HR_MANUFACTURER", "") : comboxHRManufacturer.Text = ""
            SetAttribute("STKASSY_INFO_HR_UNITS", "") : comboxHRUnits.Text = "IMPERIAL"

            ' Machine Data 
            SetTieBar("STKASSY_INFO_TIEBAR", "", "", "") : txtBoxTieBarDistanceHorizontal.Text = "" : txtBoxTieBarDistanceVertical.Text = ""
            SetMoldOrHR("STKASSY_INFO_CLAMP_TONNAGE", "", "") : txtBoxClampTonnage.Text = ""
            SetMoldOrHR("STKASSY_INFO_CLAMP_STROKE", "", "") : txtBoxClampStroke.Text = ""
            SetMoldOrHR("STKASSY_INFO_MAX_DAYLIGHT", "", "") : txtBoxMaxDaylight.Text = ""
            SetMoldOrHR("STKASSY_INFO_LOCATING_RING_DIA", "", "") : txtBoxLocatingRingDiameter.Text = ""
            SetMoldOrHR("STKASSY_INFO_NOZZLE_RADIUS", "", "") : txtBoxNozzleRadius.Text = ""
            SetMoldOrHR("STKASSY_INFO_MAX_EJECTOR_STROKE", "", "") : txtBoxMaxEjectorStroke.Text = ""
            SetMinMaxShutHeight("STKASSY_INFO_MIN_SHUT_HEIGHT", "STKASSY_INFO_MAX_SHUT_HEIGHT", "STKASSY_INFO_SHUT_HEIGHT", "", "", "") : txtBoxMinShutHeight.Text = "" : txtBoxMaxShutHeight.Text = ""
            SetAttribute("STKASSY_INFO_MACHINE_UNITS", "") : comboxMachineUnits.Text = ""
            'loadDefaultValues()
            workPart.DeleteRetainedDraftingObjectsInCurrentLayout()
            updateVisibility(1) ' Update portal with empty values 
            updateVisibilityLog(1)
        End Sub
        Public Sub SetBlankPart(ByVal name As String)
            SetAttribute(name, "")
            SetAttribute(name + "_GENERAL", "")
            SetAttribute(name + "_METRIC", "")
            SetAttribute(name + "_IMPERIAL", "")
        End Sub
        Public Sub SetBlankComponent(ByVal index As String)
            SetAttribute("STKASSY_INFO_COMPONENT_" + index, "")
            SetAttribute("STKASSY_INFO_COMPONENT_DESCRIPTION" + index, "")
            SetAttribute("STKASSY_INFO_COMPONENT_MATERIAL" + index, "")
            SetAttribute("STKASSY_INFO_COMPONENT_HARDNESS" + index, "")
            SetAttribute("STKASSY_INFO_COMPONENT_SURFACE_ENH" + index, "")
        End Sub
        Public Sub SetComponentInformation(ByVal name As String)
            Dim count As Integer = 1
            For Each row As DataGridViewRow In titleBlockComps.Rows
                SetAttribute(name + "_" + count.ToString, row.cells(0).value.ToString)
                count += 1
                If count > 9 Then
                    Exit For
                End If
            Next
            count = 1
            For Each row As DataGridViewRow In compInfo.Rows
                SetAttribute(name + "_DESCRIPTION" + count.ToString, row.cells(0).value.ToString)
                SetAttribute(name + "_MATERIAL" + count.ToString, row.cells(1).value.ToString)
                SetAttribute(name + "_HARDNESS" + count.ToString, row.cells(2).value.ToString)
                SetAttribute(name + "_SURFACE_ENH" + count.ToString, row.cells(3).value.ToString)
                count += 1
                If count > 9 Then
                    Exit For
                End If
            Next

            If (count <= 9) Then
                For i As Integer = count To 9
                    SetAttribute(name + "_" + i.ToString, "")
                    SetAttribute(name + "_DESCRIPTION" + i.ToString, "")
                    SetAttribute(name + "_MATERIAL" + i.ToString, "")
                    SetAttribute(name + "_HARDNESS" + i.ToString, "")
                    SetAttribute(name + "_SURFACE_ENH" + i.ToString, "")
                Next
            End If
            If (count > 9) Then
                MsgBox(count.ToString + "  Oh no! Count > 9! Stop drop and roll! Call the Admins! ") 'Still the best Error Message I have ever seen! Thank you Soren!
            End If
        End Sub
        Public Sub SetPartWL(ByVal name As String, ByRef txtW As String, ByRef txtL As String)
            If (txtW <> "" Or txtL <> "") Then
                If (comboxPartUnits.Text = "IMPERIAL") Then
                    SetAttribute(name + "_GENERAL_STRING", txtW + "x" + txtL + Chr(34))
                    SetAttribute(name + "_METRIC_STRING", "")
                    SetAttribute(name + "_IMPERIAL_STRING", "")
                ElseIf (comboxPartUnits.Text = "METRIC") Then
                    SetAttribute(name + "_GENERAL_STRING", inchesTomm(txtW) + "x" + inchesTomm(txtL) + "mm")
                    SetAttribute(name + "_METRIC_STRING", "")
                    SetAttribute(name + "_IMPERIAL_STRING", "")
                ElseIf (comboxPartUnits.Text = "DUAL") Then
                    SetAttribute(name + "_GENERAL_STRING", "")
                    SetAttribute(name + "_METRIC_STRING", inchesTomm(txtW) + "x" + inchesTomm(txtL) + "mm")
                    SetAttribute(name + "_IMPERIAL_STRING", txtBoxPartWidth.Text + "x" + txtL + Chr(34))
                End If
            Else
                SetAttribute(name + "_GENERAL_STRING", "")
                SetAttribute(name + "_METRIC_STRING", "")
                SetAttribute(name + "_IMPERIAL_STRING", "")
            End If
        End Sub
        Public Sub SetPart(ByVal name As String, ByRef value As String, ByRef units As String)
            SetAttribute(name, value)
            If (units = "IMPERIAL" And value <> "") Then
                SetAttribute(name + "_GENERAL", value + Chr(34))
                SetAttribute(name + "_METRIC", "")
                SetAttribute(name + "_IMPERIAL", "")
            ElseIf (units = "METRIC" And value <> "") Then
                SetAttribute(name + "_GENERAL", inchesTomm(value) + "mm")
                SetAttribute(name + "_METRIC", "")
                SetAttribute(name + "_IMPERIAL", "")
            ElseIf (comboxPartUnits.Text = "DUAL" And value <> "") Then
                SetAttribute(name + "_GENERAL", "")
                SetAttribute(name + "_METRIC", inchesTomm(value) + "mm")
                SetAttribute(name + "_IMPERIAL", value + Chr(34))
            ElseIf (value = "") Then
                SetAttribute(name + "_GENERAL", "")
                SetAttribute(name + "_METRIC", "")
                SetAttribute(name + "_IMPERIAL", "")
            End If
        End Sub
        Public Sub SetMoldPitch(ByVal name As String, ByRef xValue As String, ByRef yValue As String, ByRef units As String)
            If (units = "mm" Or units = Chr(34)) Then
                SetAttribute(name + "_X_GENERAL_VALUE", xValue)
                SetAttribute(name + "_GENERAL_UNIT", units)
                SetAttribute(name + "_X_METRIC_VALUE", "")
                SetAttribute(name + "_X_IMPERIAL_VALUE", "")
                SetAttribute(name + "_X_GENERAL_STRING", xValue + units)
                SetAttribute(name + "_X_METRIC_STRING", "")
                SetAttribute(name + "_X_IMPERIAL_STRING", "")
                SetAttribute(name + "_Y_GENERAL_VALUE", yValue)
                SetAttribute(name + "_Y_METRIC_VALUE", "")
                SetAttribute(name + "_Y_IMPERIAL_VALUE", "")
                SetAttribute(name + "_Y_GENERAL_STRING", yValue + units)
                SetAttribute(name + "_Y_METRIC_STRING", "")
                SetAttribute(name + "_Y_IMPERIAL_STRING", "")
                SetAttribute(name + "_DUAL_UNIT", "")
            ElseIf (units = "mm-Dual") Then
                SetAttribute(name + "_X_GENERAL_VALUE", "")
                SetAttribute(name + "_GENERAL_UNIT", "")
                SetAttribute(name + "_X_METRIC_VALUE", xValue)
                SetAttribute(name + "_X_IMPERIAL_VALUE", mmToInches(xValue))
                SetAttribute(name + "_X_GENERAL_STRING", "")
                SetAttribute(name + "_X_METRIC_STRING", xValue + " mm")
                SetAttribute(name + "_X_IMPERIAL_STRING", mmToInches(xValue) + Chr(34))
                SetAttribute(name + "_Y_GENERAL_VALUE", "")
                SetAttribute(name + "_Y_METRIC_VALUE", yValue)
                SetAttribute(name + "_Y_IMPERIAL_VALUE", mmToInches(yValue))
                SetAttribute(name + "_Y_GENERAL_STRING", "")
                SetAttribute(name + "_Y_METRIC_STRING", yValue + " mm")
                SetAttribute(name + "_Y_IMPERIAL_STRING", mmToInches(yValue) + Chr(34))
                SetAttribute(name + "_DUAL_UNIT", "mm-Dual")
            ElseIf (units = Chr(34) + "-Dual") Then
                SetAttribute(name + "_X_GENERAL_VALUE", "")
                SetAttribute(name + "_GENERAL_UNIT", "")
                SetAttribute(name + "_X_METRIC_VALUE", inchesTomm(xValue))
                SetAttribute(name + "_X_IMPERIAL_VALUE", xValue)
                SetAttribute(name + "_X_GENERAL_STRING", "")
                SetAttribute(name + "_X_METRIC_STRING", inchesTomm(xValue) + " mm")
                SetAttribute(name + "_X_IMPERIAL_STRING", xValue + Chr(34))
                SetAttribute(name + "_Y_GENERAL_VALUE", "")
                SetAttribute(name + "_Y_METRIC_VALUE", inchesTomm(yValue))
                SetAttribute(name + "_Y_IMPERIAL_VALUE", yValue)
                SetAttribute(name + "_Y_GENERAL_STRING", "")
                SetAttribute(name + "_Y_METRIC_STRING", inchesTomm(yValue) + " mm")
                SetAttribute(name + "_Y_IMPERIAL_STRING", yValue + Chr(34))
                SetAttribute(name + "_DUAL_UNIT", Chr(34) + "-Dual")
            End If
            If (xValue = "") Then
                SetAttribute(name + "_X_GENERAL_VALUE", "")
                SetAttribute(name + "_X_GENERAL_STRING", "")
                SetAttribute(name + "_X_METRIC_STRING", "")
                SetAttribute(name + "_X_METRIC_VALUE", "")
                SetAttribute(name + "_X_IMPERIAL_VALUE", "")
                SetAttribute(name + "_X_IMPERIAL_STRING", "")
            End If
            If (yValue = "") Then
                SetAttribute(name + "_Y_GENERAL_VALUE", "")
                SetAttribute(name + "_Y_GENERAL_STRING", "")
                SetAttribute(name + "_Y_METRIC_VALUE", "")
                SetAttribute(name + "_Y_METRIC_STRING", "")
                SetAttribute(name + "_Y_IMPERIAL_VALUE", "")
                SetAttribute(name + "_Y_IMPERIAL_STRING", "")
            End If
            If (xValue = "" And yValue = "") Then
                SetAttribute(name + "_GENERAL_UNIT", "")
                SetAttribute(name + "_DUAL_UNIT", "")
            End If
        End Sub
        Public Sub SetMoldCav(ByVal name As String, ByRef value1 As String, ByRef value2 As String)
            SetAttribute(name + "_1", value1)
            SetAttribute(name + "_2", value2)
        End Sub
        Public Sub SetMoldSize(ByVal name As String, ByRef xValue As String, ByRef yValue As String, ByRef units As String)
            If (xValue = "" And yValue = "") Then
                SetAttribute(name + "_X_GENERAL_VALUE", "")
                SetAttribute(name + "_Y_GENERAL_VALUE", "")
                SetAttribute(name + "_X_Y_GENERAL_UNIT", "")
                SetAttribute(name + "_X_METRIC_VALUE", "")
                SetAttribute(name + "_Y_METRIC_VALUE", "")
                SetAttribute(name + "_X_IMPERIAL_VALUE", "")
                SetAttribute(name + "_Y_IMPERIAL_VALUE", "")
                SetAttribute(name + "_X_Y_DUAL_UNIT", "")
                SetAttribute(name + "_GENERAL_STRING", "")
                SetAttribute(name + "_METRIC_STRING", "")
                SetAttribute(name + "_IMPERIAL_STRING", "")
                Exit Sub
            End If

            If (units = "mm" Or units = Chr(34)) Then ' Single dimension case
                SetAttribute(name + "_X_GENERAL_VALUE", xValue)
                SetAttribute(name + "_Y_GENERAL_VALUE", yValue)
                SetAttribute(name + "_X_Y_GENERAL_UNIT", units)
                SetAttribute(name + "_X_METRIC_VALUE", "")
                SetAttribute(name + "_Y_METRIC_VALUE", "")
                SetAttribute(name + "_X_IMPERIAL_VALUE", "")
                SetAttribute(name + "_Y_IMPERIAL_VALUE", "")
                SetAttribute(name + "_X_Y_DUAL_UNIT", "")
                SetAttribute(name + "_GENERAL_STRING", xValue + units + " x " + yValue + units)
                SetAttribute(name + "_METRIC_STRING", "")
                SetAttribute(name + "_IMPERIAL_STRING", "")
            ElseIf (units = "mm-Dual") Then ' Dual dimension metric case
                SetAttribute(name + "_X_GENERAL_VALUE", "")
                SetAttribute(name + "_Y_GENERAL_VALUE", "")
                SetAttribute(name + "_X_Y_GENERAL_UNIT", "")
                SetAttribute(name + "_X_METRIC_VALUE", xValue)
                SetAttribute(name + "_Y_METRIC_VALUE", yValue)
                SetAttribute(name + "_X_IMPERIAL_VALUE", mmToInches(xValue))
                SetAttribute(name + "_Y_IMPERIAL_VALUE", mmToInches(yValue))
                SetAttribute(name + "_X_Y_DUAL_UNIT", "mm-Dual")
                SetAttribute(name + "_GENERAL_STRING", "")
                SetAttribute(name + "_METRIC_STRING", xValue + "mm x " + yValue + "mm")
                SetAttribute(name + "_IMPERIAL_STRING", mmToInches(xValue) + Chr(34) + " x " + mmToInches(yValue) + Chr(34))
            ElseIf (units = Chr(34) + "-Dual") Then ' Dual dimension imperial case
                SetAttribute(name + "_X_GENERAL_VALUE", "")
                SetAttribute(name + "_Y_GENERAL_VALUE", "")
                SetAttribute(name + "_X_Y_GENERAL_UNIT", "")
                SetAttribute(name + "_X_METRIC_VALUE", inchesTomm(xValue))
                SetAttribute(name + "_Y_METRIC_VALUE", inchesTomm(yValue))
                SetAttribute(name + "_X_IMPERIAL_VALUE", xValue)
                SetAttribute(name + "_Y_IMPERIAL_VALUE", yValue)
                SetAttribute(name + "_X_Y_DUAL_UNIT", Chr(34) + "-Dual")
                SetAttribute(name + "_GENERAL_STRING", "")
                SetAttribute(name + "_METRIC_STRING", inchesTomm(xValue) + "mm x " + inchesTomm(yValue) + "mm")
                SetAttribute(name + "_IMPERIAL_STRING", xValue + Chr(34) + " x " + yValue + Chr(34))
            End If
        End Sub
        Public Sub SetMoldOrHR(ByVal name As String, ByRef value As String, ByRef units As String)
            If (units = "mm" Or units = Chr(34) Or units = "SI Tonnes" Or units = "US Tons" Or units = "kg" Or units = "lbs") Then
                SetAttribute(name + "_GENERAL_VALUE", value)
                SetAttribute(name + "_GENERAL_UNIT", units)
                SetAttribute(name + "_METRIC_VALUE", "")
                SetAttribute(name + "_IMPERIAL_VALUE", "")
                SetAttribute(name + "_GENERAL_STRING", value + " " + units)
                SetAttribute(name + "_METRIC_STRING", "")
                SetAttribute(name + "_IMPERIAL_STRING", "")
                SetAttribute(name + "_DUAL_UNIT", "")
            ElseIf (units = "mm-Dual" Or units = "SI Tonnes-Dual" Or units = "kg-Dual") Then
                SetAttribute(name + "_GENERAL_VALUE", "")
                SetAttribute(name + "_GENERAL_UNIT", "")
                SetAttribute(name + "_GENERAL_STRING", "")
                SetAttribute(name + "_METRIC_VALUE", value)
                If (units = "SI Tonnes-Dual") Then
                    SetAttribute(name + "_IMPERIAL_VALUE", metTonToImpTon(value))
                    SetAttribute(name + "_IMPERIAL_STRING", metTonToImpTon(value) + " US Tons")
                    SetAttribute(name + "_METRIC_STRING", value + " SI Tonnes")
                    SetAttribute(name + "_DUAL_UNIT", "SI Tonnes-Dual")
                ElseIf (units = "mm-Dual") Then
                    SetAttribute(name + "_IMPERIAL_VALUE", mmToInches(value))
                    SetAttribute(name + "_IMPERIAL_STRING", mmToInches(value) + Chr(34))
                    SetAttribute(name + "_METRIC_STRING", value + " mm")
                    SetAttribute(name + "_DUAL_UNIT", "mm-Dual")
                ElseIf (units = "kg-Dual") Then
                    SetAttribute(name + "_IMPERIAL_VALUE", kgToPound(value))
                    SetAttribute(name + "_IMPERIAL_STRING", kgToPound(value) + " lbs")
                    SetAttribute(name + "_METRIC_STRING", value + " kg")
                    SetAttribute(name + "_DUAL_UNIT", "kg-Dual")
                End If
            ElseIf (units = Chr(34) + "-Dual" Or units = "US Tons-Dual" Or units = "lbs-Dual") Then
                SetAttribute(name + "_GENERAL_VALUE", "")
                SetAttribute(name + "_GENERAL_UNIT", "")
                SetAttribute(name + "_GENERAL_STRING", "")
                SetAttribute(name + "_IMPERIAL_VALUE", value)
                If (units = "US Tons-Dual") Then
                    SetAttribute(name + "_METRIC_VALUE", impTonToMetTon(value))
                    SetAttribute(name + "_METRIC_STRING", impTonToMetTon(value) + " SI Tonnes")
                    SetAttribute(name + "_IMPERIAL_STRING", value + " US TONS")
                    SetAttribute(name + "_DUAL_UNIT", "US Tons-Dual")
                ElseIf (units = Chr(34) + "-Dual") Then
                    SetAttribute(name + "_METRIC_VALUE", inchesTomm(value))
                    SetAttribute(name + "_METRIC_STRING", inchesTomm(value) + " mm")
                    SetAttribute(name + "_IMPERIAL_STRING", value + " " + Chr(34))
                    SetAttribute(name + "_DUAL_UNIT", Chr(34) + "-Dual")
                ElseIf (units = "lbs-Dual") Then
                    SetAttribute(name + "_METRIC_VALUE", poundToKg(value))
                    SetAttribute(name + "_METRIC_STRING", poundToKg(value) + " kg")
                    SetAttribute(name + "_IMPERIAL_STRING", value + " lbs")
                    SetAttribute(name + "_DUAL_UNIT", "lbs-Dual")
                End If
            End If
            If (value = "") Then
                SetAttribute(name + "_GENERAL_VALUE", value)
                SetAttribute(name + "_GENERAL_UNIT", units)
                SetAttribute(name + "_GENERAL_STRING", "")
                SetAttribute(name + "_METRIC_VALUE", "")
                SetAttribute(name + "_METRIC_STRING", "")
                SetAttribute(name + "_IMPERIAL_VALUE", "")
                SetAttribute(name + "_IMPERIAL_STRING", "")
                SetAttribute(name + "_DUAL_UNIT", "")
            End If
        End Sub
        Public Sub SetMinMaxShutHeight(ByVal minName As String, ByVal maxName As String, ByVal genName As String, ByRef txtMin As String, ByRef txtMax As String, ByRef units As String)
            If (units = "mm" Or units = Chr(34)) Then ' Single unit case
                SetAttribute(minName + "_GENERAL_VALUE", txtMin)
                SetAttribute(maxName + "_GENERAL_VALUE", txtMax)
                SetAttribute(genName + "_GENERAL_UNIT", units)
                SetAttribute(genName + "_GENERAL_STRING", txtMin + units + "/" + txtMax + units)
                SetAttribute(minName + "_METRIC_VALUE", "")
                SetAttribute(maxName + "_METRIC_VALUE", "")
                SetAttribute(genName + "_METRIC_STRING", "")
                SetAttribute(minName + "_IMPERIAL_VALUE", "")
                SetAttribute(maxName + "_IMPERIAL_VALUE", "")
                SetAttribute(genName + "_IMPERIAL_STRING", "")
                SetAttribute(genName + "_DUAL_UNIT", "")
            ElseIf (units = "mm-Dual") Then ' Dual unit case, metric values input
                SetAttribute(minName + "_GENERAL_VALUE", "")
                SetAttribute(maxName + "_GENERAL_VALUE", "")
                SetAttribute(genName + "_GENERAL_UNIT", "")
                SetAttribute(genName + "_GENERAL_STRING", "")
                SetAttribute(minName + "_METRIC_VALUE", txtMin)
                SetAttribute(maxName + "_METRIC_VALUE", txtMax)
                SetAttribute(genName + "_METRIC_STRING", txtMin + "mm/" + txtMax + "mm")
                SetAttribute(minName + "_IMPERIAL_VALUE", mmToInches(txtMin))
                SetAttribute(maxName + "_IMPERIAL_VALUE", mmToInches(txtMax))
                SetAttribute(genName + "_IMPERIAL_STRING", mmToInches(txtMin) + Chr(34) + "/" + mmToInches(txtMax) + Chr(34))
                SetAttribute(genName + "_DUAL_UNIT", "mm-Dual")
            ElseIf (units = Chr(34) + "-Dual") Then
                SetAttribute(minName + "_GENERAL_VALUE", "")
                SetAttribute(maxName + "_GENERAL_VALUE", "")
                SetAttribute(genName + "_GENERAL_UNIT", "")
                SetAttribute(genName + "_GENERAL_STRING", "")
                SetAttribute(minName + "_METRIC_VALUE", inchesTomm(txtMin))
                SetAttribute(maxName + "_METRIC_VALUE", inchesTomm(txtMax))
                SetAttribute(genName + "_METRIC_STRING", inchesTomm(txtMin) + "mm/" + inchesTomm(txtMax) + "mm")
                SetAttribute(minName + "_IMPERIAL_VALUE", txtMin)
                SetAttribute(maxName + "_IMPERIAL_VALUE", txtMax)
                SetAttribute(genName + "_IMPERIAL_STRING", txtMin + Chr(34) + "/" + txtMax + Chr(34))
                SetAttribute(genName + "_DUAL_UNIT", Chr(34) + "-Dual")
            End If

            If (txtMin = "" And txtMax = "") Then
                SetAttribute(minName + "_GENERAL_VALUE", "")
                SetAttribute(maxName + "_GENERAL_VALUE", "")
                SetAttribute(genName + "_GENERAL_UNIT", "")
                SetAttribute(genName + "_GENERAL_STRING", "")
                SetAttribute(minName + "_METRIC_VALUE", "")
                SetAttribute(maxName + "_METRIC_VALUE", "")
                SetAttribute(genName + "_METRIC_STRING", "")
                SetAttribute(minName + "_IMPERIAL_VALUE", "")
                SetAttribute(maxName + "_IMPERIAL_VALUE", "")
                SetAttribute(genName + "_IMPERIAL_STRING", "")
                SetAttribute(genName + "_DUAL_UNIT", "")
            End If
        End Sub
        Public Sub SetTieBar(ByVal name As String, ByRef hValue As String, ByRef vValue As String, ByRef units As String)
            If (units = "mm" Or units = Chr(34)) Then     ' Just metric or just imperial case
                SetAttribute(name + "_HORIZONTAL", hValue)
                SetAttribute(name + "_HORIZONTAL_METRIC", "")
                SetAttribute(name + "_HORIZONTAL_IMPERIAL", "")
                SetAttribute(name + "_VERTICAL", vValue)
                SetAttribute(name + "_VERTICAL_METRIC", "")
                SetAttribute(name + "_VERTICAL_IMPERIAL", "")
                SetAttribute(name + "_UNITS", units)
                SetAttribute(name + "_UNITS_METRIC", "")
                SetAttribute(name + "_UNITS_IMPERIAL", "")
                If (units = "mm") Then
                    SetAttribute(name + "_GENERAL_STRING", hValue + "mm x " + vValue + "mm")
                    SetAttribute(name + "_METRIC_STRING", "")
                    SetAttribute(name + "_IMPERIAL_STRING", "")
                ElseIf (units = Chr(34)) Then
                    SetAttribute(name + "_GENERAL_STRING", hValue + Chr(34) + " x " + vValue + Chr(34))
                    SetAttribute(name + "_METRIC_STRING", "")
                    SetAttribute(name + "_IMPERIAL_STRING", "")
                End If
            ElseIf (units = "mm-Dual") Then   ' Dual dimensions, metric values were inputted into the program
                SetAttribute(name + "_HORIZONTAL", "")
                SetAttribute(name + "_HORIZONTAL_METRIC", hValue)
                SetAttribute(name + "_HORIZONTAL_IMPERIAL", mmToInches(hValue)) ' Convert metric value to imperial and enter here 
                SetAttribute(name + "_VERTICAL", "")
                SetAttribute(name + "_VERTICAL_METRIC", vValue) ' convert metric value to imperial and enter here
                SetAttribute(name + "_VERTICAL_IMPERIAL", mmToInches(vValue))
                SetAttribute(name + "_UNITS", "")
                SetAttribute(name + "_UNITS_METRIC", "mm") ' Dual Dimensions, Imperial values were inputted into the program
                SetAttribute(name + "_UNITS_IMPERIAL", Chr(34))
                SetAttribute(name + "_GENERAL_STRING", "")
                SetAttribute(name + "_METRIC_STRING", hValue + "mm x " + vValue + "mm")
                SetAttribute(name + "_IMPERIAL_STRING", mmToInches(hValue) + Chr(34) + " x " + mmToInches(vValue) + Chr(34))
                SetAttribute(name + "_DUAL_UNIT", "mm-Dual")
            ElseIf (units = Chr(34) + "-Dual") Then
                SetAttribute(name + "_HORIZONTAL", "")
                SetAttribute(name + "_HORIZONTAL_METRIC", inchesTomm(hValue))
                SetAttribute(name + "_HORIZONTAL_IMPERIAL", hValue) ' Convert metric value to imperial and enter here 
                SetAttribute(name + "_VERTICAL", "")
                SetAttribute(name + "_VERTICAL_METRIC", inchesTomm(vValue)) ' convert metric value to imperial and enter here
                SetAttribute(name + "_VERTICAL_IMPERIAL", vValue)
                SetAttribute(name + "_UNITS", "")
                SetAttribute(name + "_UNITS_METRIC", "mm")
                SetAttribute(name + "_UNITS_IMPERIAL", Chr(34))
                SetAttribute(name + "_GENERAL_STRING", "")
                SetAttribute(name + "_METRIC_STRING", inchesTomm(hValue) + "mm x " + inchesTomm(vValue) + "mm")
                SetAttribute(name + "_IMPERIAL_STRING", hValue + Chr(34) + " x " + vValue + Chr(34))
                SetAttribute(name + "_DUAL_UNIT", Chr(34) + "-Dual")
            End If

            If (hValue = "" And vValue = "") Then
                SetAttribute(name + "_HORIZONTAL", "")
                SetAttribute(name + "_HORIZONTAL_METRIC", "")
                SetAttribute(name + "_HORIZONTAL_IMPERIAL", "")
                SetAttribute(name + "_VERTICAL", "")
                SetAttribute(name + "_VERTICAL_METRIC", "")
                SetAttribute(name + "_VERTICAL_IMPERIAL", "")
                SetAttribute(name + "_UNITS", "")
                SetAttribute(name + "_UNITS_METRIC", "")
                SetAttribute(name + "_UNITS_IMPERIAL", "")
                SetAttribute(name + "_GENERAL_STRING", "")
                SetAttribute(name + "_METRIC_STRING", "")
                SetAttribute(name + "_IMPERIAL_STRING", "")
                SetAttribute(name + "_DUAL_UNIT", "")
            End If
        End Sub
        Private Sub setCustomAttributes()
            'General Information
            SetAttribute("STKASSY_INFO_MOLD_RUN_COUNT", (moldRunCount + 1).ToString)
            SetAttribute("STKASSY_INFO_DRAWN_BY", txtBoxDesigner.Text.Trim)
            SetAttribute("STKASSY_INFO_TEAM", txtBoxDesignTeam.Text.Trim)
            SetAttribute("STKASSY_INFO_DATE", txtBoxDate.Text.Trim)
            SetAttribute("STKASSY_INFO_DRAWING_NUMBER", txtBoxDrawingNumber.Text.Trim)
            SetAttribute("STKASSY_INFO_SCALE", txtBoxScale.Text.Trim)
            SetAttribute("STKASSY_INFO_CURRENT_SHEET", txtBoxCurrentSheet.Text.Trim)
            SetAttribute("STKASSY_INFO_TOTAL_SHEETS", txtBoxTotalSheets.Text.Trim)
            SetAttribute("STKASSY_INFO_MOLD_UNITS", comboxMoldUnits2.Text.Trim)
            SetAttribute("STKASSY_INFO_LAST_TAB", Tabs.SelectedIndex.ToString)
            SetAttribute("STKASSY_INFO_PROJECT_NUMBER", txtBoxJobNumber.Text.Trim)
            SetAttribute("STKASSY_INFO_MACHINE_MODEL", txtBoxMoldMachineModel.Text.Trim)

            ' Part Information 
            SetAttribute("STKASSY_INFO_PART_UNITS", comboxPartUnits.Text.Trim)
            SetAttribute("STKASSY_INFO_PART_TITLE", txtBoxPartTitle.Text.Trim)
            SetAttribute("STKASSY_INFO_PART_RESIN", txtBoxPartResin.Text.Trim)
            SetAttribute("STKASSY_INFO_PART_SHRINKAGE", txtBoxPartShrinkage.Text.Trim)
            SetAttribute("STKASSY_INFO_PART_WEIGHT", txtBoxPartWeight.Text.Trim)
            SetAttribute("STKASSY_INFO_PART_DENSITY", txtBoxPartDensity.Text.Trim)
            SetAttribute("STKASSY_INFO_PART_APPEARANCE", txtBoxPartAppearance.Text.Trim)
            SetAttribute("STKASSY_INFO_PART_VOL_TO_BRIM", txtBoxPartVolToBrim.Text.Trim)
            SetAttribute("STKASSY_INFO_PART_PROJECTED_AREA", txtBoxPartProjectedArea.Text.Trim)
            SetAttribute("STKASSY_INFO_PART_LTRATIO", txtBoxPartLTRatio.Text.Trim)
            SetAttribute("STKASSY_INFO_PART_CUSTOMER", txtBoxCustomer.Text.Trim)
            SetPart("STKASSY_INFO_PART_FLOW_LENGTH", txtBoxPartFlowLength.Text, comboxPartUnits.Text.Trim)
            SetPart("STKASSY_INFO_PART_DIAMETER", txtBoxPartDiameter.Text.Trim, comboxPartUnits.Text.Trim)
            SetPart("STKASSY_INFO_PART_WIDTH", txtBoxPartWidth.Text.Trim, comboxPartUnits.Text.Trim)
            SetPart("STKASSY_INFO_PART_LENGTH", txtBoxPartLength.Text.Trim, comboxPartUnits.Text.Trim)
            SetPart("STKASSY_INFO_PART_HEIGHT", txtBoxPartHeight.Text.Trim, comboxPartUnits.Text.Trim)
            SetPart("STKASSY_INFO_PART_WALL_SECT_SIDE", txtBoxPartWSSide.Text.Trim, comboxPartUnits.Text.Trim)
            SetPart("STKASSY_INFO_PART_WALL_SECT_BOTTOM", txtBoxPartWSBottom.Text.Trim, comboxPartUnits.Text.Trim)
            SetPartWL("STKASSY_INFO_PART_WL", txtBoxPartWidth.Text.Trim, txtBoxPartLength.Text.Trim)

            ' Component Information 
            SetComponentInformation("STKASSY_INFO_COMPONENT")

            ' Mold Data 
            SetAttribute("STKASSY_INFO_MOLD_DESCRIPTION", txtBoxMoldDescription.Text.Trim)
            SetMoldCav("STKASSY_INFO_MOLD_CAV", txtBoxMoldCavitation1.Text.Trim, txtBoxMoldCavitation2.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_MOLD_SHUT_HEIGHT", txtBoxMoldShutHeight.Text.Trim, comboxMoldShutHeightUnits.Text.Trim)
            SetAttribute("STKASSY_INFO_MOLD_MAX_OPENING", txtBoxMaxMoldOpeningPerSide.Text.Trim)
            SetMoldPitch("STKASSY_INFO_MOLD_PITCH", txtBoxMoldPitchX.Text.Trim, txtBoxMoldPitchY.Text.Trim, comboxMoldPitchUnits.Text.Trim)
            SetMoldSize("STKASSY_INFO_MOLD_SIZE", txtBoxMoldSizeX.Text.Trim, txtBoxMoldSizeY.Text.Trim, comboxMoldSizeUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_MOLD_TONNAGE", txtBoxMoldTonnage.Text.Trim, comboxMoldTonnageUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_MOLD_QPC", txtBoxQPCModuleShutHeight.Text.Trim, comboxQPCModuleShutHeightUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_EST_MOLD_WEIGHT", txtBoxMoldWeight.Text.Trim, comboxMoldWeightUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_MOLD_EJECTION_STROKE", txtBoxEjectionStroke.Text.Trim, comboxEjectionStrokeUnits.Text.Trim)
            SetAttribute("STKASSY_INFO_MOLD_EJECTION1", comboxEjectionType1.Text.Trim)
            SetAttribute("STKASSY_INFO_MOLD_EJECTION2", comboxEjectionType2.Text.Trim)
            SetAttribute("STKASSY_INFO_MOLD_STACK_INTERLOCK", comboxStackInterlock.Text.Trim)
            ' Hardware
            SetAttribute("STKASSY_INFO_HARDWARE", comboxHardware.Text.Trim)

            'HR Data
            SetMoldOrHR("STKASSY_INFO_HR_XDIM", txtBoxXDim.Text.Trim, comboxXDimUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_HR_LDIM", txtBoxLDim.Text.Trim, comboxLDimUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_HR_PDIMHOT", txtBoxPDimHot.Text.Trim, comboxPDimHotUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_HR_PDIMCOLD", txtBoxPDimCold.Text.Trim, comboxPDimColdUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_HR_GATEDIA", txtBoxGateDia.Text.Trim, comboxGateDiaUnits.Text.Trim)
            SetAttribute("STKASSY_INFO_HR_NOZZLE_TIP_NUMBER", txtBoxNozzleTipPartNumber.Text.Trim)
            SetAttribute("STKASSY_INFO_HR_GATE_SIDE", comboxGateSide.Text.Trim)
            SetAttribute("STKASSY_INFO_HR_GATE_TYPE", comboxGateType.Text.Trim)
            SetAttribute("STKASSY_INFO_HR_MANUFACTURER", comboxHRManufacturer.Text.Trim)
            SetAttribute("STKASSY_INFO_HR_UNITS", comboxHRUnits.Text.Trim)

            ' Machine Data 
            ' Logic for setting single/dual dimensions for the TIEBAR
            SetTieBar("STKASSY_INFO_TIEBAR", txtBoxTieBarDistanceHorizontal.Text.Trim, txtBoxTieBarDistanceVertical.Text.Trim, comboxTieBarDistanceUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_CLAMP_TONNAGE", txtBoxClampTonnage.Text.Trim, comboxClampTonnageUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_CLAMP_STROKE", txtBoxClampStroke.Text.Trim, comboxClampStrokeUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_MAX_DAYLIGHT", txtBoxMaxDaylight.Text.Trim, comboxMaxDaylightUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_LOCATING_RING_DIA", txtBoxLocatingRingDiameter.Text.Trim, comboxLocatingRingDiameterUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_NOZZLE_RADIUS", txtBoxNozzleRadius.Text.Trim, comboxNozzleRadiusUnits.Text.Trim)
            SetMoldOrHR("STKASSY_INFO_MAX_EJECTOR_STROKE", txtBoxMaxEjectorStroke.Text.Trim, comboxMaxEjectorStrokeUnits.Text.Trim)
            SetMinMaxShutHeight("STKASSY_INFO_MIN_SHUT_HEIGHT", "STKASSY_INFO_MAX_SHUT_HEIGHT", "STKASSY_INFO_SHUT_HEIGHT", txtBoxMinShutHeight.Text.Trim, txtBoxMaxShutHeight.Text.Trim, comboxMinMaxShutHeightUnits.Text.Trim)
            SetAttribute("STKASSY_INFO_MACHINE_UNITS", comboxMachineUnits.Text.Trim)

            ' Delete Retained Annotations
            DeleteRetainedAnnotations()

            ' Update portal with values entered by user 
            updateVisibility(0)
            updateVisibilityLog(0)
        End Sub
        Public Sub DeleteRetainedAnnotations()
            Dim preferencesBuilder1 As Drafting.PreferencesBuilder
            preferencesBuilder1 = workPart.SettingsManager.CreatePreferencesBuilder()

            ' To get rid of retained annotations from previous sheets
            Dim viewStyleFPCalloutConfigBuilder1 As Drawings.ViewStyleFPCalloutConfigBuilder
            viewStyleFPCalloutConfigBuilder1 = preferencesBuilder1.ViewStyle.GetViewStyleFPCalloutConfig()

            Dim viewStyleFPCalloutConfigBuilder2 As Drawings.ViewStyleFPCalloutConfigBuilder
            viewStyleFPCalloutConfigBuilder2 = preferencesBuilder1.ViewStyle.GetViewStyleFPCalloutConfig()

            workPart.DeleteRetainedDraftingObjectsInCurrentLayout()
        End Sub

        Public Sub clearTemporaryAccessDatabase()
            ' Delete entries that haven't been sent to visibility (stackassy table) in the last 5 minutes for this job 

            If (txtBoxJobNumber.Text.Length < 5 Or IsNumeric(txtBoxJobNumber.Text) = False) Then
                Exit Sub
            End If

            Dim Conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & "Y:\eng\ENG_ACCESS_DATABASES\StackAssyAttributes.mdb")
            Dim Com As OleDbCommand

            ' Delete all entries in the SQL Database
            Com = New OleDbCommand("DELETE * FROM STACK_ASSY_ATTRIBUTES WHERE SALESORDERNO LIKE '%" & txtBoxJobNumber.Text & "%'", Conn)
            Conn.Open()
            Com.ExecuteNonQuery()
            Conn.Close()

        End Sub

        Private Sub updateVisibility(ByVal resetCheck As Integer)
            If (txtBoxJobNumber.Text.Trim = "" Or IsNumeric(txtBoxJobNumber.Text) = False Or txtBoxJobNumber.Text.Trim.Length <> 5) Then
                Exit Sub
            End If

            If (txtBoxJobNumber.Text.Trim = "80999") Then
                Exit Sub
            End If

            clearTemporaryAccessDatabase()

            Dim Conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & "Y:\eng\ENG_ACCESS_DATABASES\StackAssyAttributes.mdb")
            Dim Com As OleDbCommand

            Dim strResin As String = Nothing
            If (txtBoxPartResin.Text.ToUpper = "PP") Then
                strResin = "1"
            ElseIf (txtBoxPartResin.Text.ToUpper = "HDPE") Then
                strResin = "2"
            ElseIf (txtBoxPartResin.Text.ToUpper = "LDPE") Then
                strResin = "3"
            ElseIf (txtBoxPartResin.Text.ToUpper = "LLDPE") Then
                strResin = "4"
            ElseIf (txtBoxPartResin.Text.ToUpper = "PS") Then
                strResin = "5"
            ElseIf (txtBoxPartResin.Text.ToUpper = "MASTERBATCH") Then
                strResin = "6"
            ElseIf (txtBoxPartResin.Text.ToUpper = "OTHER") Then
                strResin = "7"
            ElseIf (txtBoxPartResin.Text.ToUpper = "PC") Then
                strResin = "8"
            ElseIf (txtBoxPartResin.Text.ToUpper = "SAN") Then
                strResin = "9"
            ElseIf (txtBoxPartResin.Text.ToUpper = "ABS") Then
                strResin = "10"
            Else
                strResin = "0"
            End If

            Dim strStack As String = Nothing
            If (comboxStackInterlock.Text = "WEDGE LOCK") Then
                strStack = "1"
            ElseIf (comboxStackInterlock.Text = "CORE LOCK") Then
                strStack = "2"
            ElseIf (comboxStackInterlock.Text = "CAVITY LOCK") Then
                strStack = "3"
            ElseIf (comboxStackInterlock.Text = "STRIPPER LOCK/BUMP OFF") Then
                strStack = "4"
            ElseIf (comboxStackInterlock.Text = "SPLITS") Then
                strStack = "5"
            ElseIf (comboxStackInterlock.Text = "FLAT") Then
                strStack = "6"
            Else
                strStack = "0"
            End If

            Dim strEjection1 As String = Nothing
            If (comboxEjectionType1.Text = "AIR EJECTION") Then
                strEjection1 = "1"
            ElseIf (comboxEjectionType1.Text = "STRIPPER EJECTION") Then
                strEjection1 = "2"
            ElseIf (comboxEjectionType1.Text = "GULL WING") Then
                strEjection1 = "3"
            ElseIf (comboxEjectionType1.Text = "EJECTION BOX") Then
                strEjection1 = "4"
            ElseIf (comboxEjectionType1.Text = "2 STAGES") Then
                strEjection1 = "5"
            ElseIf (comboxEjectionType1.Text = "3 STAGES") Then
                strEjection1 = "6"
            ElseIf (comboxEjectionType1.Text = "5 PIECES COLLAPSE") Then
                strEjection1 = "7"
            ElseIf (comboxEjectionType1.Text = "UNSCREWING") Then
                strEjection1 = "8"
            Else
                strEjection1 = "0"
            End If

            Dim strEjection2 As String = Nothing
            If (comboxEjectionType2.Text = "AIR EJECTION") Then
                strEjection2 = "1"
            ElseIf (comboxEjectionType2.Text = "STRIPPER EJECTION") Then
                strEjection2 = "2"
            ElseIf (comboxEjectionType2.Text = "GULL WING") Then
                strEjection2 = "3"
            ElseIf (comboxEjectionType2.Text = "EJECTION BOX") Then
                strEjection2 = "4"
            ElseIf (comboxEjectionType2.Text = "2 STAGES") Then
                strEjection2 = "5"
            ElseIf (comboxEjectionType2.Text = "3 STAGES") Then
                strEjection2 = "6"
            ElseIf (comboxEjectionType2.Text = "5 PIECES COLLAPSE") Then
                strEjection2 = "7"
            ElseIf (comboxEjectionType2.Text = "UNSCREWING") Then
                strEjection2 = "8"
            Else
                strEjection2 = "0"
            End If

            Dim GateSide As String = Nothing
            If (comboxGateSide.Text = "CORE SIDE") Then
                GateSide = "1"
            ElseIf (comboxGateSide.Text = "CAVITY SIDE") Then
                GateSide = "2"
            Else
                GateSide = "0"
            End If

            Dim GateType As String = Nothing
            If (comboxGateType.Text = "EDGE GATES") Then
                GateType = "1"
            ElseIf (comboxGateType.Text = "VALVE GATES") Then
                GateType = "2"
            ElseIf (comboxGateType.Text = "HOT TIP") Then
                GateType = "3"
            ElseIf (comboxGateType.Text = "COLD RUNNER") Then
                GateType = "4"
            Else
                GateType = "0"
            End If
            'Nothing!
            ' Need to check what units were and make them imperial values if they were metric
            Dim moldShutHeight As String = txtBoxMoldShutHeight.Text.Trim
            Dim moldTonnage As String = txtBoxMoldTonnage.Text.Trim
            Dim moldSizeX As String = txtBoxMoldSizeX.Text.Trim
            Dim moldSizeY As String = txtBoxMoldSizeY.Text.Trim
            Dim moldPitchX As String = txtBoxMoldPitchX.Text.Trim
            Dim moldPitchY As String = txtBoxMoldPitchY.Text.Trim

            If (comboxMoldShutHeightUnits.Text.Trim = "mm-Dual" Or comboxMoldShutHeightUnits.Text.Trim = "mm") Then ' need to first convert to imperial
                moldShutHeight = mmToInches(moldShutHeight)
            End If

            If (moldShutHeight = Nothing Or IsNumeric(moldShutHeight) = False) Then
                moldShutHeight = ""
            End If

            If (comboxMoldTonnageUnits.Text.Trim = "SI Tonnes" Or comboxMoldTonnageUnits.Text.Trim = "SI Tonnes-Dual") Then ' conver from kg to lbs
                moldTonnage = metTonToImpTon(moldTonnage)
            End If

            If (moldTonnage = Nothing Or IsNumeric(moldTonnage) = False) Then
                moldTonnage = ""
            End If

            If (comboxMoldSizeUnits.Text.Trim = "mm-Dual" Or comboxMoldSizeUnits.Text.Trim = "mm") Then
                lw.WriteLine("CONVERTING MOLD SIZE DIMENSIONS FROM METRIC TO IMPERIAL!!!!!")
                moldSizeX = mmToInches(moldSizeX)
                moldSizeY = mmToInches(moldSizeY)
                lw.WriteLine("MOLD SIZE X: " + moldSizeX)
            End If

            If (comboxMoldPitchUnits.Text.Trim = "mm-Dual" Or comboxMoldPitchUnits.Text.Trim = "mm") Then
                moldPitchX = mmToInches(moldPitchX)
                moldPitchY = mmToInches(moldPitchY)
            End If

            ' Need to put in 2 entries if ejection1 <> ejection 2 

            If (resetCheck = 0) Then

                Conn.Open()

                Com = New OleDbCommand("INSERT INTO STACK_ASSY_ATTRIBUTES (SALESORDERNO, D, H, RESIN, STACK, EJECTION, GATING_SIDE, GATING_TYPE, FILENAME, ENTITY_CODE, CUSTOMER_MACHINE, UNITS, SHUT_HEIGHT, MOLD_TONNAGE, LT_RATIO, PART_MASS, HR_MANUFACTURER, CUSTOMER_NAME, PART_TITLE, SHRINKAGE, APPEARANCE, PROJECTED_AREA, W, L, MOLD_CAV1, MOLD_CAV2, MOLD_WIDTH, MOLD_HEIGHT, MOLD_PITCH_X, MOLD_PITCH_Y, TYPE) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", Conn)

                Try
                    Com.Parameters.AddWithValue("@p1", "90" + txtBoxJobNumber.Text)
                    Com.Parameters.AddWithValue("@p2", txtBoxPartDiameter.Text.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p3", txtBoxPartHeight.Text.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p4", strResin)
                    Com.Parameters.AddWithValue("@p5", strStack)
                    If (strEjection1 <> "N/A" And strEjection1 <> "0") Then
                        Com.Parameters.AddWithValue("@p6", strEjection1)
                    Else
                        Com.Parameters.AddWithValue("@p6", "")
                    End If
                    Com.Parameters.AddWithValue("@p7", GateSide)
                    Com.Parameters.AddWithValue("@p8", GateType)
                    Com.Parameters.AddWithValue("@p9", txtBoxJobNumber.Text)
                    Com.Parameters.AddWithValue("@p10", "01")
                    Com.Parameters.AddWithValue("@p11", txtBoxMoldMachineModel.Text)
                    Com.Parameters.AddWithValue("@p12", comboxHardware.Text.Trim)
                    Com.Parameters.AddWithValue("@p13", moldShutHeight.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p14", moldTonnage.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p15", txtBoxPartLTRatio.Text.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p16", txtBoxPartWeight.Text.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p17", comboxHRManufacturer.Text)
                    Com.Parameters.AddWithValue("@p18", txtBoxCustomer.Text.Trim)
                    Com.Parameters.AddWithValue("@p19", txtBoxPartTitle.Text.Trim)
                    Com.Parameters.AddWithValue("@p20", txtBoxPartShrinkage.Text.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p21", txtBoxPartAppearance.Text.Trim)
                    Com.Parameters.AddWithValue("@p22", txtBoxPartProjectedArea.Text.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p23", txtBoxPartWidth.Text.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p24", txtBoxPartLength.Text.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p25", txtBoxMoldCavitation1.Text.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p26", txtBoxMoldCavitation2.Text.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p27", moldSizeX.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p28", moldSizeY.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p29", moldPitchX.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p30", moldPitchY.Replace(Chr(34), "").Replace("mm", ""))
                    Com.Parameters.AddWithValue("@p31", "STACK")
                    Com.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

                If (strEjection2.Trim <> "" And strEjection2 <> strEjection1 And strEjection2.Trim <> "N/A") Then
                    Com = New OleDbCommand("INSERT INTO STACK_ASSY_ATTRIBUTES (SALESORDERNO, D, H, RESIN, STACK, EJECTION, GATING_SIDE, GATING_TYPE, FILENAME, ENTITY_CODE, CUSTOMER_MACHINE, UNITS, SHUT_HEIGHT, MOLD_TONNAGE, LT_RATIO, PART_MASS, HR_MANUFACTURER, CUSTOMER_NAME, PART_TITLE, SHRINKAGE, APPEARANCE, PROJECTED_AREA, W, L, MOLD_CAV1, MOLD_CAV2, MOLD_WIDTH, MOLD_HEIGHT, MOLD_PITCH_X, MOLD_PITCH_Y, TYPE) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", Conn)
                    Try
                        Com.Parameters.AddWithValue("@p1", "90" + txtBoxJobNumber.Text.Trim)
                        Com.Parameters.AddWithValue("@p2", txtBoxPartDiameter.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p3", txtBoxPartHeight.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p4", strResin.Trim)
                        Com.Parameters.AddWithValue("@p5", strStack.Trim)
                        Com.Parameters.AddWithValue("@p6", strEjection2.Trim)
                        Com.Parameters.AddWithValue("@p7", GateSide.Trim)
                        Com.Parameters.AddWithValue("@p8", GateType.Trim)
                        Com.Parameters.AddWithValue("@p9", txtBoxJobNumber.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p10", "01")
                        Com.Parameters.AddWithValue("@p11", txtBoxMoldMachineModel.Text.Trim)
                        Com.Parameters.AddWithValue("@p12", comboxHardware.Text.Trim)
                        Com.Parameters.AddWithValue("@p13", moldShutHeight.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p14", moldTonnage.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p15", txtBoxPartLTRatio.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p16", txtBoxPartWeight.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p17", comboxHRManufacturer.Text.Trim)
                        Com.Parameters.AddWithValue("@p18", txtBoxCustomer.Text.Trim)
                        Com.Parameters.AddWithValue("@p19", txtBoxPartTitle.Text.Trim)
                        Com.Parameters.AddWithValue("@p20", txtBoxPartShrinkage.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p21", txtBoxPartAppearance.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p22", txtBoxPartProjectedArea.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p23", txtBoxPartWidth.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p24", txtBoxPartLength.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p25", txtBoxMoldCavitation1.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p26", txtBoxMoldCavitation2.Text.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p27", moldSizeX.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p28", moldSizeY.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p29", moldPitchX.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p30", moldPitchY.Trim.Replace(Chr(34), "").Replace("mm", ""))
                        Com.Parameters.AddWithValue("@p31", "STACK")
                        Com.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If
            Else
                Try
                    If (txtBoxJobNumber.Text <> "" And txtBoxJobNumber.Text.Length = 5) Then
                        Com.Parameters.AddWithValue("@p1", "90" + txtBoxJobNumber.Text)
                        Com.Parameters.AddWithValue("@p2", "")
                        Com.Parameters.AddWithValue("@p3", "")
                        Com.Parameters.AddWithValue("@p4", "")
                        Com.Parameters.AddWithValue("@p5", "")
                        Com.Parameters.AddWithValue("@p6", "")
                        Com.Parameters.AddWithValue("@p7", "")
                        Com.Parameters.AddWithValue("@p8", "")
                        Com.Parameters.AddWithValue("@p9", "")
                        Com.Parameters.AddWithValue("@p10", "01")
                        Com.Parameters.AddWithValue("@p11", "")
                        Com.Parameters.AddWithValue("@p12", "")
                        Com.Parameters.AddWithValue("@p13", "")
                        Com.Parameters.AddWithValue("@p14", "")
                        Com.Parameters.AddWithValue("@p15", "")
                        Com.Parameters.AddWithValue("@p16", "")
                        Com.Parameters.AddWithValue("@p17", "")
                        Com.Parameters.AddWithValue("@p18", "")
                        Com.Parameters.AddWithValue("@p19", "")
                        Com.Parameters.AddWithValue("@p20", "")
                        Com.Parameters.AddWithValue("@p21", "")
                        Com.Parameters.AddWithValue("@p22", "")
                        Com.Parameters.AddWithValue("@p23", "")
                        Com.Parameters.AddWithValue("@p24", "")
                        Com.Parameters.AddWithValue("@p25", "")
                        Com.Parameters.AddWithValue("@p26", "")
                        Com.Parameters.AddWithValue("@p27", "")
                        Com.Parameters.AddWithValue("@p28", "")
                        Com.Parameters.AddWithValue("@p29", "")
                        Com.Parameters.AddWithValue("@p30", "")
                        Com.Parameters.AddWithValue("@p31", "")
                        Com.ExecuteNonQuery()
                    End If
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            End If

            Try
                Conn.Close()
            Catch ex As Exception
            End Try
        End Sub

        Private Sub updateVisibilityLog(ByVal resetCheck As Integer)
            If (txtBoxJobNumber.Text.Trim = "" Or IsNumeric(txtBoxJobNumber.Text) = False Or txtBoxJobNumber.Text.Trim.Length <> 5) Then
                Exit Sub
            End If

            Dim Conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & "Y:\eng\ENG_ACCESS_DATABASES\StackAssyAttributes.mdb")
            Dim Com As OleDbCommand

            Dim strResin As String = Nothing
            If (txtBoxPartResin.Text.ToUpper = "PP") Then
                strResin = "1"
            ElseIf (txtBoxPartResin.Text.ToUpper = "HDPE") Then
                strResin = "2"
            ElseIf (txtBoxPartResin.Text.ToUpper = "LDPE") Then
                strResin = "3"
            ElseIf (txtBoxPartResin.Text.ToUpper = "LLDPE") Then
                strResin = "4"
            ElseIf (txtBoxPartResin.Text.ToUpper = "PS") Then
                strResin = "5"
            ElseIf (txtBoxPartResin.Text.ToUpper = "MASTERBATCH") Then
                strResin = "6"
            ElseIf (txtBoxPartResin.Text.ToUpper = "OTHER") Then
                strResin = "7"
            ElseIf (txtBoxPartResin.Text.ToUpper = "PC") Then
                strResin = "8"
            ElseIf (txtBoxPartResin.Text.ToUpper = "SAN") Then
                strResin = "9"
            ElseIf (txtBoxPartResin.Text.ToUpper = "ABS") Then
                strResin = "10"
            Else
                strResin = "0"
            End If

            Dim strStack As String = Nothing
            If (comboxStackInterlock.Text = "WEDGE LOCK") Then
                strStack = "1"
            ElseIf (comboxStackInterlock.Text = "CORE LOCK") Then
                strStack = "2"
            ElseIf (comboxStackInterlock.Text = "CAVITY LOCK") Then
                strStack = "3"
            ElseIf (comboxStackInterlock.Text = "STRIPPER LOCK/BUMP OFF") Then
                strStack = "4"
            ElseIf (comboxStackInterlock.Text = "SPLITS") Then
                strStack = "5"
            ElseIf (comboxStackInterlock.Text = "FLAT") Then
                strStack = "6"
            Else
                strStack = "0"
            End If

            Dim strEjection1 As String = Nothing
            If (comboxEjectionType1.Text = "AIR EJECTION") Then
                strEjection1 = "1"
            ElseIf (comboxEjectionType1.Text = "STRIPPER EJECTION") Then
                strEjection1 = "2"
            ElseIf (comboxEjectionType1.Text = "GULL WING") Then
                strEjection1 = "3"
            ElseIf (comboxEjectionType1.Text = "EJECTION BOX") Then
                strEjection1 = "4"
            ElseIf (comboxEjectionType1.Text = "2 STAGES") Then
                strEjection1 = "5"
            ElseIf (comboxEjectionType1.Text = "3 STAGES") Then
                strEjection1 = "6"
            ElseIf (comboxEjectionType1.Text = "5 PIECES COLLAPSE") Then
                strEjection1 = "7"
            ElseIf (comboxEjectionType1.Text = "UNSCREWING") Then
                strEjection1 = "8"

            Else
                strEjection1 = "0"
            End If

            Dim strEjection2 As String = Nothing
            If (comboxEjectionType2.Text = "AIR EJECTION") Then
                strEjection2 = "1"
            ElseIf (comboxEjectionType2.Text = "STRIPPER EJECTION") Then
                strEjection2 = "2"
            ElseIf (comboxEjectionType2.Text = "GULL WING") Then
                strEjection2 = "3"
            ElseIf (comboxEjectionType2.Text = "EJECTION BOX") Then
                strEjection2 = "4"
            ElseIf (comboxEjectionType2.Text = "2 STAGES") Then
                strEjection2 = "5"
            ElseIf (comboxEjectionType2.Text = "3 STAGES") Then
                strEjection2 = "6"
            ElseIf (comboxEjectionType2.Text = "5 PIECES COLLAPSE") Then
                strEjection2 = "7"
            ElseIf (comboxEjectionType2.Text = "UNSCREWING") Then
                strEjection2 = "8"
            Else
                strEjection2 = "0"
            End If

            Dim GateSide As String = Nothing
            If (comboxGateSide.Text = "CORE SIDE") Then
                GateSide = "1"
            ElseIf (comboxGateSide.Text = "CAVITY SIDE") Then
                GateSide = "2"
            Else
                GateSide = "0"
            End If

            Dim GateType As String = Nothing
            If (comboxGateType.Text = "EDGE GATES") Then
                GateType = "1"
            ElseIf (comboxGateType.Text = "VALVE GATES") Then
                GateType = "2"
            ElseIf (comboxGateType.Text = "HOT TIP") Then
                GateType = "3"
            ElseIf (comboxGateType.Text = "COLD RUNNER") Then
                GateType = "4"
            Else
                GateType = "0"
            End If

            ' Need to check what units were and make them imperial values if they were metric
            Dim moldShutHeight As String = txtBoxMoldShutHeight.Text.Trim
            Dim moldTonnage As String = txtBoxMoldTonnage.Text.Trim
            Dim moldSizeX As String = txtBoxMoldSizeX.Text.Trim
            Dim moldSizeY As String = txtBoxMoldSizeY.Text.Trim
            Dim moldPitchX As String = txtBoxMoldPitchX.Text.Trim
            Dim moldPitchY As String = txtBoxMoldPitchY.Text.Trim

            If (comboxMoldShutHeightUnits.Text.Trim = "mm-Dual" Or comboxMoldShutHeightUnits.Text.Trim = "mm") Then ' need to first convert to imperial
                moldShutHeight = mmToInches(moldShutHeight)
            End If

            If (moldShutHeight = Nothing Or IsNumeric(moldShutHeight) = False) Then
                moldShutHeight = ""
            End If

            If (comboxMoldTonnageUnits.Text.Trim = "SI Tonnes" Or comboxMoldTonnageUnits.Text.Trim = "SI Tonnes-Dual") Then ' conver from kg to lbs
                moldTonnage = metTonToImpTon(moldTonnage)
            End If

            If (moldTonnage = Nothing Or IsNumeric(moldTonnage) = False) Then
                moldTonnage = ""
            End If

            If (comboxMoldSizeUnits.Text.Trim = "mm-Dual" Or comboxMoldSizeUnits.Text.Trim = "mm") Then
                lw.WriteLine("CONVERTING MOLD SIZE DIMENSIONS FROM METRIC TO IMPERIAL!!!!!")
                moldSizeX = mmToInches(moldSizeX)
                moldSizeY = mmToInches(moldSizeY)
                lw.WriteLine("MOLD SIZE X: " + moldSizeX)
            End If

            If (comboxMoldPitchUnits.Text.Trim = "mm-Dual" Or comboxMoldPitchUnits.Text.Trim = "mm") Then
                moldPitchX = mmToInches(moldPitchX)
                moldPitchY = mmToInches(moldPitchY)
            End If

            ' Need to put in 2 entries if ejection1 <> ejection 2 

            If (resetCheck = 0) Then

                Conn.Open()

                Com = New OleDbCommand("INSERT INTO LOGS (SALESORDERNO, D, H, RESIN, STACK, EJECTION, GATING_SIDE, GATING_TYPE, FILENAME, ENTITY_CODE, CUSTOMER_MACHINE, UNITS, SHUT_HEIGHT, MOLD_TONNAGE, LT_RATIO, PART_MASS, HR_MANUFACTURER, CUSTOMER_NAME, PART_TITLE, SHRINKAGE, APPEARANCE, PROJECTED_AREA, W, L, MOLD_CAV1, MOLD_CAV2, MOLD_WIDTH, MOLD_HEIGHT, MOLD_PITCH_X, MOLD_PITCH_Y, TYPE, DADATE, DAUSER) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", Conn)

                Try
                    Com.Parameters.AddWithValue("@p1", "90" + txtBoxJobNumber.Text)
                    Com.Parameters.AddWithValue("@p2", txtBoxPartDiameter.Text)
                    Com.Parameters.AddWithValue("@p3", txtBoxPartHeight.Text)
                    Com.Parameters.AddWithValue("@p4", strResin)
                    Com.Parameters.AddWithValue("@p5", strStack)
                    If (strEjection1 <> "N/A" And strEjection1 <> "0") Then
                        Com.Parameters.AddWithValue("@p6", strEjection1)
                    Else
                        Com.Parameters.AddWithValue("@p6", "")
                    End If
                    Com.Parameters.AddWithValue("@p7", GateSide)
                    Com.Parameters.AddWithValue("@p8", GateType)
                    Com.Parameters.AddWithValue("@p9", txtBoxJobNumber.Text)
                    Com.Parameters.AddWithValue("@p10", "01")
                    Com.Parameters.AddWithValue("@p11", txtBoxMoldMachineModel.Text)
                    Com.Parameters.AddWithValue("@p12", comboxHardware.Text)
                    Com.Parameters.AddWithValue("@p13", moldShutHeight)
                    Com.Parameters.AddWithValue("@p14", moldTonnage)
                    Com.Parameters.AddWithValue("@p15", txtBoxPartLTRatio.Text)
                    Com.Parameters.AddWithValue("@p16", txtBoxPartWeight.Text)
                    Com.Parameters.AddWithValue("@p17", comboxHRManufacturer.Text)
                    Com.Parameters.AddWithValue("@p18", txtBoxCustomer.Text)
                    Com.Parameters.AddWithValue("@p19", txtBoxPartTitle.Text)
                    Com.Parameters.AddWithValue("@p20", txtBoxPartShrinkage.Text)
                    Com.Parameters.AddWithValue("@p21", txtBoxPartShrinkage.Text)
                    Com.Parameters.AddWithValue("@p22", txtBoxPartProjectedArea.Text)
                    Com.Parameters.AddWithValue("@p23", txtBoxPartWidth.Text)
                    Com.Parameters.AddWithValue("@p24", txtBoxPartLength.Text)
                    Com.Parameters.AddWithValue("@p25", txtBoxMoldCavitation1.Text)
                    Com.Parameters.AddWithValue("@p26", txtBoxMoldCavitation2.Text)
                    Com.Parameters.AddWithValue("@p27", moldSizeX)
                    Com.Parameters.AddWithValue("@p28", moldSizeY)
                    Com.Parameters.AddWithValue("@p29", moldPitchX)
                    Com.Parameters.AddWithValue("@p30", moldPitchY)
                    Com.Parameters.AddWithValue("@p31", "STACK")
                    Com.Parameters.AddWithValue("@p32", Now.ToString)
                    Com.Parameters.AddWithValue("@p33", System.Environment.UserName)
                    Com.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

                If (strEjection2.Trim <> "" And strEjection2 <> strEjection1 And strEjection2.Trim <> "N/A") Then
                    Com = New OleDbCommand("INSERT INTO LOGS (SALESORDERNO, D, H, RESIN, STACK, EJECTION, GATING_SIDE, GATING_TYPE, FILENAME, ENTITY_CODE, CUSTOMER_MACHINE, UNITS, SHUT_HEIGHT, MOLD_TONNAGE, LT_RATIO, PART_MASS, HR_MANUFACTURER, CUSTOMER_NAME, PART_TITLE, SHRINKAGE, APPEARANCE, PROJECTED_AREA, W, L, MOLD_CAV1, MOLD_CAV2, MOLD_WIDTH, MOLD_HEIGHT, MOLD_PITCH_X, MOLD_PITCH_Y, TYPE, DADATE, DAUSER) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", Conn)
                    Try
                        Com.Parameters.AddWithValue("@p1", "90" + txtBoxJobNumber.Text.Trim)
                        Com.Parameters.AddWithValue("@p2", txtBoxPartDiameter.Text.Trim)
                        Com.Parameters.AddWithValue("@p3", txtBoxPartHeight.Text.Trim)
                        Com.Parameters.AddWithValue("@p4", strResin)
                        Com.Parameters.AddWithValue("@p5", strStack.Trim)
                        Com.Parameters.AddWithValue("@p6", strEjection2.Trim)
                        Com.Parameters.AddWithValue("@p7", GateSide.Trim)
                        Com.Parameters.AddWithValue("@p8", GateType.Trim)
                        Com.Parameters.AddWithValue("@p9", txtBoxJobNumber.Text.Trim)
                        Com.Parameters.AddWithValue("@p10", "01")
                        Com.Parameters.AddWithValue("@p11", txtBoxMoldMachineModel.Text.Trim)
                        Com.Parameters.AddWithValue("@p12", comboxHardware.Text.Trim)
                        Com.Parameters.AddWithValue("@p13", moldShutHeight.Trim)
                        Com.Parameters.AddWithValue("@p14", moldTonnage.Trim)
                        Com.Parameters.AddWithValue("@p15", txtBoxPartLTRatio.Text.Trim)
                        Com.Parameters.AddWithValue("@p16", txtBoxPartWeight.Text.Trim)
                        Com.Parameters.AddWithValue("@p17", comboxHRManufacturer.Text.Trim)
                        Com.Parameters.AddWithValue("@p18", txtBoxCustomer.Text.Trim)
                        Com.Parameters.AddWithValue("@p19", txtBoxPartTitle.Text.Trim)
                        Com.Parameters.AddWithValue("@p20", txtBoxPartShrinkage.Text.Trim)
                        Com.Parameters.AddWithValue("@p21", txtBoxPartShrinkage.Text.Trim)
                        Com.Parameters.AddWithValue("@p22", txtBoxPartProjectedArea.Text.Trim)
                        Com.Parameters.AddWithValue("@p23", txtBoxPartWidth.Text.Trim)
                        Com.Parameters.AddWithValue("@p24", txtBoxPartLength.Text.Trim)
                        Com.Parameters.AddWithValue("@p25", txtBoxMoldCavitation1.Text.Trim)
                        Com.Parameters.AddWithValue("@p26", txtBoxMoldCavitation2.Text.Trim)
                        Com.Parameters.AddWithValue("@p27", moldSizeX.Trim)
                        Com.Parameters.AddWithValue("@p28", moldSizeY.Trim)
                        Com.Parameters.AddWithValue("@p29", moldPitchX.Trim)
                        Com.Parameters.AddWithValue("@p30", moldPitchY.Trim)
                        Com.Parameters.AddWithValue("@p31", "STACK")
                        Com.Parameters.AddWithValue("@p32", Now.ToString)
                        Com.Parameters.AddWithValue("@p33", System.Environment.UserName)
                        Com.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If
            Else
                Try
                    Com = New OleDbCommand("INSERT INTO LOGS (SALESORDERNO, D, H, RESIN, STACK, EJECTION, GATING_SIDE, GATING_TYPE, FILENAME, ENTITY_CODE, CUSTOMER_MACHINE, UNITS, SHUT_HEIGHT, MOLD_TONNAGE, LT_RATIO, PART_MASS, HR_MANUFACTURER, CUSTOMER_NAME, PART_TITLE, SHRINKAGE, APPEARANCE, PROJECTED_AREA, W, L, MOLD_CAV1, MOLD_CAV2, MOLD_WIDTH, MOLD_HEIGHT, MOLD_PITCH_X, MOLD_PITCH_Y, TYPE, DADATE, DAUSER) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", Conn)

                    If (txtBoxJobNumber.Text <> "" And txtBoxJobNumber.Text.Length = 5) Then
                        Com.Parameters.AddWithValue("@p1", "90" + txtBoxJobNumber.Text)
                        Com.Parameters.AddWithValue("@p2", "")
                        Com.Parameters.AddWithValue("@p3", "")
                        Com.Parameters.AddWithValue("@p4", "")
                        Com.Parameters.AddWithValue("@p5", "")
                        Com.Parameters.AddWithValue("@p6", "")
                        Com.Parameters.AddWithValue("@p7", "")
                        Com.Parameters.AddWithValue("@p8", "")
                        Com.Parameters.AddWithValue("@p9", "")
                        Com.Parameters.AddWithValue("@p10", "01")
                        Com.Parameters.AddWithValue("@p11", "")
                        Com.Parameters.AddWithValue("@p12", "")
                        Com.Parameters.AddWithValue("@p13", "")
                        Com.Parameters.AddWithValue("@p14", "")
                        Com.Parameters.AddWithValue("@p15", "")
                        Com.Parameters.AddWithValue("@p16", "")
                        Com.Parameters.AddWithValue("@p17", "")
                        Com.Parameters.AddWithValue("@p18", "")
                        Com.Parameters.AddWithValue("@p19", "")
                        Com.Parameters.AddWithValue("@p20", "")
                        Com.Parameters.AddWithValue("@p21", "")
                        Com.Parameters.AddWithValue("@p22", "")
                        Com.Parameters.AddWithValue("@p23", "")
                        Com.Parameters.AddWithValue("@p24", "")
                        Com.Parameters.AddWithValue("@p25", "")
                        Com.Parameters.AddWithValue("@p26", "")
                        Com.Parameters.AddWithValue("@p27", "")
                        Com.Parameters.AddWithValue("@p28", "")
                        Com.Parameters.AddWithValue("@p29", "")
                        Com.Parameters.AddWithValue("@p30", "")
                        Com.Parameters.AddWithValue("@p31", "")
                        Com.Parameters.AddWithValue("@p32", "")
                        Com.Parameters.AddWithValue("@p33", "")
                        Com.ExecuteNonQuery()
                    End If
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            End If

            Try
                Conn.Close()
            Catch ex As Exception
            End Try
        End Sub

        Private Sub txtBoxMoldMachineModel_TextChanged(sender As Object, e As EventArgs) Handles txtBoxMoldMachineModel.TextChanged
            Dim sb As String = Nothing
            Dim c As Char
            For Each c In txtBoxMoldMachineModel.Text
                If Not (Char.IsSymbol(c) OrElse Char.IsPunctuation(c)) Then
                    sb = sb + c
                End If
            Next
            txtBoxMoldMachineModel.Text = sb
        End Sub
        Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
            System.Diagnostics.Process.Start("\\cntfiler\_JobLib\EngProcedures\UG\Assign StackMold Assembly Attributes.docx")
        End Sub
        Private Sub pdfButton_Click(sender As Object, e As EventArgs) Handles pdfButton.Click ' Send PDF of A size drawing to the portal
            setCustomAttributes()
            'If (txtBoxMoldCavitation1.Text = "" Or txtBoxMoldCavitation2.Text = "") Then
            '    MsgBox("Cavitaiton cannot be blank!")
            '    Exit Sub
            'End If

            'If (txtBoxMoldPitchX.Text ca= "" Or txtBoxMoldPitchY.Text = "") Then
            '    MsgBox("Mold Pitch cannot be blank!")
            '    Exit Sub
            'End If

            'If (txtBoxMoldSizeX.Text = "" Or txtBoxMoldSizeY.Text = "") Then
            '    MsgBox("Mold Width/Height cannot be blank!")
            '    Exit Sub
            'End If

            'If (txtBoxMoldShutHeight.Text = "") Then
            '    MsgBox("Mold Shut Height cannot be blank!")
            '    Exit Sub
            'End If

            If (comboxStackInterlock.Text = "") Then
                MsgBox("Stack Interlock cannot be blank!")
                Exit Sub
            End If

            If (comboxEjectionType1.Text = "" And comboxEjectionType2.Text = "") Then
                MsgBox("There must be at least 1 ejection type!")
                Exit Sub
            End If

            If (txtBoxJobNumber.Text = "") Then
                MsgBox("Project Number cannot be blank!")
                Exit Sub
            End If

            If (txtBoxCustomer.Text = "") Then
                MsgBox("Customer field cannot be blank!")
                Exit Sub
            End If

            'If (txtBoxPartTitle.Text = "") Then
            '    MsgBox("Part Description cannot be blank!")
            '    Exit Sub
            'End If

            'If (txtBoxPartDiameter.Text = "" And (txtBoxPartLength.Text = "" And txtBoxPartWidth.Text = "")) Then
            '    MsgBox("Part Diameter cannot be blank!")
            '    Exit Sub
            'End If

            'If (txtBoxPartWidth.Text = "" And txtBoxPartDiameter.Text = "") Then
            '    MsgBox("Part Width cannot be blank!")
            '    Exit Sub
            'End If

            'If (txtBoxPartLength.Text = "" And txtBoxPartDiameter.Text = "") Then
            '    MsgBox("Part Length cannot be blank!")
            '    Exit Sub
            'End If

            'If (txtBoxPartHeight.Text = "") Then
            '    MsgBox("Part Height cannot be blank!")
            '    Exit Sub
            'End If

            'If (txtBoxPartResin.Text = "") Then
            '    MsgBox("Part Resin cannot be blank!")
            '    Exit Sub
            'End If

            'If (txtBoxPartShrinkage.Text = "") Then
            '    MsgBox("Part Shrinkage cannot be blank!")
            '    Exit Sub
            'End If

            'If (txtBoxPartWSSide.Text = "") Then
            '    MsgBox("Side Wall Section cannot be blank!")
            '    Exit Sub
            'End If

            'If (comboxGateSide.Text = "") Then
            '    MsgBox("Gate Side cannot be blank!")
            '    Exit Sub
            'End If

            'If (comboxGateType.Text = "") Then
            '    MsgBox("Gate Type cannot be blank!")
            '    Exit Sub
            'End If

            Dim curDwg As NXOpen.Tag = NXOpen.Tag.Null
            ufs.Draw.AskCurrentDrawing(curDwg)

            Dim ds As Drawings.DrawingSheet = NXObjectManager.Get(curDwg)
            lw.WriteLine("PAGE HEIGHT:" + ds.Height.ToString())
            lw.WriteLine("PAGE LENGTH:" + ds.Length.ToString())

            If ((ds.Height = 11 And ds.Length = 8.5) Or (ds.Length = 11 And ds.Height = 8.5)) Then
            Else
                MsgBox("Not an A size sheet!")
                Exit Sub
            End If

            If (File.Exists("\\ideas\ideas-e\eng\stack_pdf\" + txtBoxJobNumber.Text + ".pdf")) Then
                File.Delete("\\ideas\ideas-e\eng\stack_pdf\" + txtBoxJobNumber.Text + ".pdf")
            End If

            Dim PDFExporter = New NXJ_PdfExporter
            Dim partloadstatus1 As NXOpen.PartLoadStatus = Nothing
            Dim tempName As String = fileName.replace(" ", "/")
            Dim PDFPart As NXOpen.Part = Nothing

            Try
                PDFPart = s.Parts.Work
                PDFExporter.Part = PDFPart
                PDFExporter.OutputPdfFileName = txtBoxJobNumber.Text
                PDFExporter.OutputFolder = "\\ideas\ideas-e\eng\stack_pdf"
                PDFExporter.UseWatermark = False
                PDFExporter.ShowConfirmationDialog = False
                PDFExporter.Commit()
                PDFExporter = Nothing
                MsgBox("A PDF of " + fileName + " has been sent to the  \\ideas\ideas-e\eng\stack_pdf folder")
            Catch ex As Exception
            End Try

            lw.WriteLine("Initiating Plotting Functionality")
            Dim Printer2 As String = "ENGINEERING"
            Dim PrinterProfile2 As String = "A Size (8.5 X 11)"
            Dim NumOfCopies As Integer = 0

            NumOfCopies = 1

            'If ((ds.Height = 11 And ds.Length = 8.5) Or (ds.Length = 11 And ds.Height = 8.5)) Then
            'Else
            '    MsgBox("Not an A size sheet! Please switch to the A size sheet!")
            '    Exit Sub
            'End If

            If NumOfCopies <> 0 Then
                PlotDwg(txtBoxJobNumber.Text, NumOfCopies, Printer2, PrinterProfile2)
                MsgBox("Please receive " & NumOfCopies & " Copies @ " & Printer2)
            End If
        End Sub
        Public Sub PlotDwg(ByVal Jobname As String, ByVal NumOfCopy As Integer, ByVal Printer As String, ByVal Profile As String)

            Dim plot As UFPlot = ufs.Plot

            Dim jobOpts As UFPlot.JobOptions
            plot.AskDefaultJobOptions(jobOpts)

            Dim BannerOpt As UFPlot.BannerOptions
            BannerOpt.show_banner = True

            Dim sheet As Tag
            ufs.Draw.AskCurrentDrawing(sheet)

            ' use named printer/profile
            Try
                plot.Print(sheet, jobOpts, Jobname, BannerOpt, Printer, Profile, NumOfCopy)
                MsgBox("Printed succesfully!")
            Catch ex As Exception
                MsgBox(ex.ToString())
            End Try
        End Sub
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
        'Added public propert\\ideas\ideas-e ExportSheetsIndividually and related code changes default value: False
        'Changing this property to True will cause each sheet to be exported to an individual pdf file in the specified export folder.
        '
        'Added new public method: New(byval thePart as Part)
        '  allows you to specify the part to use at the time of the NXJ_PdfExporter object creation
        '
        '
        'December 1, 2014
        'update to version 1.0
        'Added public propert\\ideas\ideas-e SkipBlankSheets [Boolean] {read/write} default value: True
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
        '           default value: "PRELIMINARY PRINT Not TO BE USED FOR PRODUCTION"
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
        '    after writing your custom sort function in the module, pass it in like this: PdfExporter.Sort(AddressOf {function name})


#End Region

#Region "properties And private variables"

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
                lg.WriteLine("  ExportSheetsIndividuall\\ideas\ideas-e " & value.ToString)
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
                        lg.WriteLine("  specified directory does Not exist, trying to create it...")
                        Directory.CreateDirectory(value)
                        lg.WriteLine("  directory created: " & value)
                    Catch ex As Exception
                        lg.WriteLine("  ** error while creating director\\ideas\ideas-e " & value)
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
                lg.WriteLine("  value passed to propert\\ideas\ideas-e " & value)
                _exportFile = IO.Path.GetFileName(value)
                If _exportFile.Substring(_exportFile.Length - 4, 4).ToLower = ".pdf" Then
                    'strip off ".pdf" extension
                    _exportFile = _exportFile.Substring(_exportFile.Length - 4, 4)
                End If
                lg.WriteLine("  _exportFile: " & _exportFile)
                If Not value.Contains("\") Then
                    lg.WriteLine("  does Not appear to contain path information")
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

        Private _watermarkText As String = "PRELIMINARY PRINT Not TO BE USED FOR PRODUCTION"
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
                lg.WriteLine("  this Is a preliminary print")
            Else
                Me.PreliminaryPrint = False
                lg.WriteLine("  this Is Not a preliminary print")
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
                            MessageBox.Show("The pdf file: " & newPdf & " exists And could Not be overwritten." & ControlChars.NewLine &
                                            "PDF export exiting", "PDF export error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            'file already exists and will not be overwritten
                            MessageBox.Show("The pdf file: " & newPdf & " exists And the overwrite option Is set to False." & ControlChars.NewLine &
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
                        MessageBox.Show("The pdf file: " & _outputPdfFile & " exists And could Not be overwritten." & ControlChars.NewLine &
                                        "PDF export exiting", "PDF export error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        'file already exists and will not be overwritten
                        MessageBox.Show("The pdf file: " & _outputPdfFile & " exists And the overwrite option Is set to False." & ControlChars.NewLine &
                                        "PDF export exiting", "PDF file exists", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    Return
                End If

            End If

            Dim sheetCount As Integer = 0
            Dim sheetsExported As Integer = 0

            Dim numPlists As Integer = 0
            Dim Plists() As Tag

            _theUfSession.Plist.AskTags(Plists, numPlists)
            For i As Integer = 0 To numPlists - 1
                _theUfSession.Plist.Update(Plists(i))
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
                lg.WriteLine("  Me.Part Is Nothing")
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
                    lg.WriteLine("  TC Is running")
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
                    'default to "Documents" folder
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
            Dim StringComp As StringComparer = StringComparer.CurrentCultureIgnoreCase

            'for a case-sensitive sort (A-Z then a-z), change the above option to:
            'Dim StringComp As StringComparer = StringComparer.CurrentCulture

            Return StringComp.Compare(x.Name, y.Name)
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

            lg.WriteLine("  pdf file: " & pdfFile)
            printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.Native
            printPDFBuilder1.Append = False
            printPDFBuilder1.Filename = pdfFile

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
    Partial Class Form1
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
            Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Me.HelpProvider1 = New System.Windows.Forms.HelpProvider()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.txtBoxDesigner = New System.Windows.Forms.TextBox()
            Me.txtBoxDesignTeam = New System.Windows.Forms.TextBox()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.txtBoxDate = New System.Windows.Forms.TextBox()
            Me.txtBoxScale = New System.Windows.Forms.TextBox()
            Me.txtBoxDrawingNumber = New System.Windows.Forms.TextBox()
            Me.Label6 = New System.Windows.Forms.Label()
            Me.Label5 = New System.Windows.Forms.Label()
            Me.Label7 = New System.Windows.Forms.Label()
            Me.Label8 = New System.Windows.Forms.Label()
            Me.txtBoxCurrentSheet = New System.Windows.Forms.TextBox()
            Me.txtBoxTotalSheets = New System.Windows.Forms.TextBox()
            Me.Label9 = New System.Windows.Forms.Label()
            Me.Label10 = New System.Windows.Forms.Label()
            Me.txtBoxJobNumber = New System.Windows.Forms.TextBox()
            Me.Label12 = New System.Windows.Forms.Label()
            Me.txtBoxMoldMachineModel = New System.Windows.Forms.TextBox()
            Me.okButton = New System.Windows.Forms.Button()
            Me.cancelButton = New System.Windows.Forms.Button()
            Me.Label77 = New System.Windows.Forms.Label()
            Me.txtBoxCustomer = New System.Windows.Forms.TextBox()
            Me.Button1 = New System.Windows.Forms.Button()
            Me.pdfButton = New System.Windows.Forms.Button()
            Me.TabPage3 = New System.Windows.Forms.TabPage()
            Me.txtBoxMoldWeight = New System.Windows.Forms.TextBox()
            Me.comboxMoldWeightUnits = New System.Windows.Forms.ComboBox()
            Me.Label46 = New System.Windows.Forms.Label()
            Me.txtBoxMaxMoldOpeningPerSide = New System.Windows.Forms.TextBox()
            Me.Label41 = New System.Windows.Forms.Label()
            Me.comboxEjectionStrokeUnits = New System.Windows.Forms.ComboBox()
            Me.txtBoxEjectionStroke = New System.Windows.Forms.TextBox()
            Me.Label42 = New System.Windows.Forms.Label()
            Me.comboxQPCModuleShutHeightUnits = New System.Windows.Forms.ComboBox()
            Me.txtBoxQPCModuleShutHeight = New System.Windows.Forms.TextBox()
            Me.Label43 = New System.Windows.Forms.Label()
            Me.comboxHardware = New System.Windows.Forms.ComboBox()
            Me.Label14 = New System.Windows.Forms.Label()
            Me.comboxMoldUnits2 = New System.Windows.Forms.ComboBox()
            Me.Label11 = New System.Windows.Forms.Label()
            Me.comboxMoldPitchUnits = New System.Windows.Forms.ComboBox()
            Me.txtBoxMoldPitchY = New System.Windows.Forms.TextBox()
            Me.txtBoxMoldPitchX = New System.Windows.Forms.TextBox()
            Me.txtBoxMoldTonnage = New System.Windows.Forms.TextBox()
            Me.txtBoxMoldShutHeight = New System.Windows.Forms.TextBox()
            Me.txtBoxMoldSizeY = New System.Windows.Forms.TextBox()
            Me.txtBoxMoldSizeX = New System.Windows.Forms.TextBox()
            Me.txtBoxMoldCavitation2 = New System.Windows.Forms.TextBox()
            Me.txtBoxMoldCavitation1 = New System.Windows.Forms.TextBox()
            Me.txtBoxMoldDescription = New System.Windows.Forms.TextBox()
            Me.Label75 = New System.Windows.Forms.Label()
            Me.Label76 = New System.Windows.Forms.Label()
            Me.comboxMoldTonnageUnits = New System.Windows.Forms.ComboBox()
            Me.comboxMoldShutHeightUnits = New System.Windows.Forms.ComboBox()
            Me.comboxMoldSizeUnits = New System.Windows.Forms.ComboBox()
            Me.Label25 = New System.Windows.Forms.Label()
            Me.comboxEjectionType2 = New System.Windows.Forms.ComboBox()
            Me.Label24 = New System.Windows.Forms.Label()
            Me.comboxEjectionType1 = New System.Windows.Forms.ComboBox()
            Me.Label23 = New System.Windows.Forms.Label()
            Me.comboxStackInterlock = New System.Windows.Forms.ComboBox()
            Me.Label22 = New System.Windows.Forms.Label()
            Me.Label33 = New System.Windows.Forms.Label()
            Me.Label32 = New System.Windows.Forms.Label()
            Me.Label29 = New System.Windows.Forms.Label()
            Me.Label30 = New System.Windows.Forms.Label()
            Me.Label31 = New System.Windows.Forms.Label()
            Me.Label40 = New System.Windows.Forms.Label()
            Me.TabPage6 = New System.Windows.Forms.TabPage()
            Me.compInfo = New System.Windows.Forms.DataGridView()
            Me.componentInfoDescriptionColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.componentInfoMaterialColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.componentInfoHardnessColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.componentInfoSurfEnhColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.partName = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.TabPage2 = New System.Windows.Forms.TabPage()
            Me.titleBlockComps = New System.Windows.Forms.DataGridView()
            Me.titleBlockPartNameColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.titleBlockComponentColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.assyComps = New System.Windows.Forms.DataGridView()
            Me.assyPartNameColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.assyComponentsColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.Label37 = New System.Windows.Forms.Label()
            Me.moveComponentFromTitleBlock = New System.Windows.Forms.Button()
            Me.moveComponentToTitleBlock = New System.Windows.Forms.Button()
            Me.Label36 = New System.Windows.Forms.Label()
            Me.TabPage4 = New System.Windows.Forms.TabPage()
            Me.comboxPartUnits = New System.Windows.Forms.ComboBox()
            Me.Label80 = New System.Windows.Forms.Label()
            Me.Label74 = New System.Windows.Forms.Label()
            Me.txtBoxPartWSBottom = New System.Windows.Forms.TextBox()
            Me.txtBoxPartLength = New System.Windows.Forms.TextBox()
            Me.txtBoxPartWSSide = New System.Windows.Forms.TextBox()
            Me.txtBoxPartHeight = New System.Windows.Forms.TextBox()
            Me.txtBoxPartWidth = New System.Windows.Forms.TextBox()
            Me.txtBoxPartDiameter = New System.Windows.Forms.TextBox()
            Me.txtBoxPartFlowLength = New System.Windows.Forms.TextBox()
            Me.txtBoxPartLTRatio = New System.Windows.Forms.TextBox()
            Me.txtBoxPartProjectedArea = New System.Windows.Forms.TextBox()
            Me.txtBoxPartVolToBrim = New System.Windows.Forms.TextBox()
            Me.txtBoxPartAppearance = New System.Windows.Forms.TextBox()
            Me.txtBoxPartDensity = New System.Windows.Forms.TextBox()
            Me.txtBoxPartWeight = New System.Windows.Forms.TextBox()
            Me.txtBoxPartShrinkage = New System.Windows.Forms.TextBox()
            Me.txtBoxPartResin = New System.Windows.Forms.TextBox()
            Me.txtBoxPartTitle = New System.Windows.Forms.TextBox()
            Me.Label64 = New System.Windows.Forms.Label()
            Me.Label63 = New System.Windows.Forms.Label()
            Me.Label62 = New System.Windows.Forms.Label()
            Me.Label61 = New System.Windows.Forms.Label()
            Me.Label60 = New System.Windows.Forms.Label()
            Me.Label59 = New System.Windows.Forms.Label()
            Me.Label58 = New System.Windows.Forms.Label()
            Me.Label56 = New System.Windows.Forms.Label()
            Me.Label57 = New System.Windows.Forms.Label()
            Me.Label50 = New System.Windows.Forms.Label()
            Me.Label47 = New System.Windows.Forms.Label()
            Me.Label45 = New System.Windows.Forms.Label()
            Me.Label44 = New System.Windows.Forms.Label()
            Me.Label39 = New System.Windows.Forms.Label()
            Me.Label34 = New System.Windows.Forms.Label()
            Me.Label26 = New System.Windows.Forms.Label()
            Me.listBoxApplicationsParts = New System.Windows.Forms.ListBox()
            Me.Tabs = New System.Windows.Forms.TabControl()
            Me.TabPage1 = New System.Windows.Forms.TabPage()
            Me.comboxHRManufacturer = New System.Windows.Forms.ComboBox()
            Me.comboxGateDiaUnits = New System.Windows.Forms.ComboBox()
            Me.Label27 = New System.Windows.Forms.Label()
            Me.txtBoxGateDia = New System.Windows.Forms.TextBox()
            Me.Label54 = New System.Windows.Forms.Label()
            Me.comboxGateType = New System.Windows.Forms.ComboBox()
            Me.Label28 = New System.Windows.Forms.Label()
            Me.comboxGateSide = New System.Windows.Forms.ComboBox()
            Me.Label35 = New System.Windows.Forms.Label()
            Me.comboxHRUnits = New System.Windows.Forms.ComboBox()
            Me.Label81 = New System.Windows.Forms.Label()
            Me.comboxPDimColdUnits = New System.Windows.Forms.ComboBox()
            Me.comboxPDimHotUnits = New System.Windows.Forms.ComboBox()
            Me.comboxLDimUnits = New System.Windows.Forms.ComboBox()
            Me.comboxXDimUnits = New System.Windows.Forms.ComboBox()
            Me.Label55 = New System.Windows.Forms.Label()
            Me.txtBoxNozzleTipPartNumber = New System.Windows.Forms.TextBox()
            Me.Label53 = New System.Windows.Forms.Label()
            Me.txtBoxPDimCold = New System.Windows.Forms.TextBox()
            Me.Label52 = New System.Windows.Forms.Label()
            Me.txtBoxPDimHot = New System.Windows.Forms.TextBox()
            Me.Label51 = New System.Windows.Forms.Label()
            Me.txtBoxLDim = New System.Windows.Forms.TextBox()
            Me.Label48 = New System.Windows.Forms.Label()
            Me.Label49 = New System.Windows.Forms.Label()
            Me.txtBoxXDim = New System.Windows.Forms.TextBox()
            Me.TabPage5 = New System.Windows.Forms.TabPage()
            Me.comboxMachineUnits = New System.Windows.Forms.ComboBox()
            Me.Label78 = New System.Windows.Forms.Label()
            Me.Label65 = New System.Windows.Forms.Label()
            Me.txtBoxMaxShutHeight = New System.Windows.Forms.TextBox()
            Me.comboxMinMaxShutHeightUnits = New System.Windows.Forms.ComboBox()
            Me.comboxMaxEjectorStrokeUnits = New System.Windows.Forms.ComboBox()
            Me.comboxNozzleRadiusUnits = New System.Windows.Forms.ComboBox()
            Me.comboxLocatingRingDiameterUnits = New System.Windows.Forms.ComboBox()
            Me.comboxMaxDaylightUnits = New System.Windows.Forms.ComboBox()
            Me.comboxClampStrokeUnits = New System.Windows.Forms.ComboBox()
            Me.comboxClampTonnageUnits = New System.Windows.Forms.ComboBox()
            Me.comboxTieBarDistanceUnits = New System.Windows.Forms.ComboBox()
            Me.txtBoxMinShutHeight = New System.Windows.Forms.TextBox()
            Me.Label21 = New System.Windows.Forms.Label()
            Me.txtBoxMaxEjectorStroke = New System.Windows.Forms.TextBox()
            Me.Label20 = New System.Windows.Forms.Label()
            Me.txtBoxNozzleRadius = New System.Windows.Forms.TextBox()
            Me.Label19 = New System.Windows.Forms.Label()
            Me.txtBoxLocatingRingDiameter = New System.Windows.Forms.TextBox()
            Me.Label18 = New System.Windows.Forms.Label()
            Me.txtBoxMaxDaylight = New System.Windows.Forms.TextBox()
            Me.Label17 = New System.Windows.Forms.Label()
            Me.txtBoxClampStroke = New System.Windows.Forms.TextBox()
            Me.Label16 = New System.Windows.Forms.Label()
            Me.txtBoxClampTonnage = New System.Windows.Forms.TextBox()
            Me.Label15 = New System.Windows.Forms.Label()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.txtBoxTieBarDistanceVertical = New System.Windows.Forms.TextBox()
            Me.txtBoxTieBarDistanceHorizontal = New System.Windows.Forms.TextBox()
            Me.Label38 = New System.Windows.Forms.Label()
            Me.txtBoxCopies = New System.Windows.Forms.TextBox()
            Me.Label13 = New System.Windows.Forms.Label()
            Me.btnReset = New System.Windows.Forms.Button()
            Me.btnLoadStack = New System.Windows.Forms.Button()
            Me.btnLoadSpec = New System.Windows.Forms.Button()
            Me.btnLoadMold = New System.Windows.Forms.Button()
            Me.TabPage3.SuspendLayout()
            Me.TabPage6.SuspendLayout()
            CType(Me.compInfo, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.TabPage2.SuspendLayout()
            CType(Me.titleBlockComps, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.assyComps, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.TabPage4.SuspendLayout()
            Me.Tabs.SuspendLayout()
            Me.TabPage1.SuspendLayout()
            Me.TabPage5.SuspendLayout()
            Me.SuspendLayout()
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label2.Location = New System.Drawing.Point(415, 9)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(143, 16)
            Me.Label2.TabIndex = 3
            Me.Label2.Text = "General Information"
            '
            'Label3
            '
            Me.Label3.AutoSize = True
            Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label3.Location = New System.Drawing.Point(175, 46)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(57, 13)
            Me.Label3.TabIndex = 4
            Me.Label3.Text = "Designer"
            '
            'txtBoxDesigner
            '
            Me.txtBoxDesigner.Location = New System.Drawing.Point(241, 43)
            Me.txtBoxDesigner.Name = "txtBoxDesigner"
            Me.txtBoxDesigner.Size = New System.Drawing.Size(115, 20)
            Me.txtBoxDesigner.TabIndex = 5
            '
            'txtBoxDesignTeam
            '
            Me.txtBoxDesignTeam.Location = New System.Drawing.Point(241, 72)
            Me.txtBoxDesignTeam.Name = "txtBoxDesignTeam"
            Me.txtBoxDesignTeam.Size = New System.Drawing.Size(115, 20)
            Me.txtBoxDesignTeam.TabIndex = 7
            '
            'Label4
            '
            Me.Label4.AutoSize = True
            Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label4.Location = New System.Drawing.Point(196, 101)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(34, 13)
            Me.Label4.TabIndex = 6
            Me.Label4.Text = "Date"
            '
            'txtBoxDate
            '
            Me.txtBoxDate.Location = New System.Drawing.Point(241, 101)
            Me.txtBoxDate.Name = "txtBoxDate"
            Me.txtBoxDate.Size = New System.Drawing.Size(115, 20)
            Me.txtBoxDate.TabIndex = 9
            '
            'txtBoxScale
            '
            Me.txtBoxScale.Location = New System.Drawing.Point(476, 72)
            Me.txtBoxScale.Name = "txtBoxScale"
            Me.txtBoxScale.Size = New System.Drawing.Size(115, 20)
            Me.txtBoxScale.TabIndex = 15
            '
            'txtBoxDrawingNumber
            '
            Me.txtBoxDrawingNumber.Location = New System.Drawing.Point(476, 43)
            Me.txtBoxDrawingNumber.Name = "txtBoxDrawingNumber"
            Me.txtBoxDrawingNumber.Size = New System.Drawing.Size(115, 20)
            Me.txtBoxDrawingNumber.TabIndex = 13
            '
            'Label6
            '
            Me.Label6.AutoSize = True
            Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label6.Location = New System.Drawing.Point(370, 46)
            Me.Label6.Name = "Label6"
            Me.Label6.Size = New System.Drawing.Size(100, 13)
            Me.Label6.TabIndex = 17
            Me.Label6.Text = "Drawing Number"
            '
            'Label5
            '
            Me.Label5.AutoSize = True
            Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label5.Location = New System.Drawing.Point(151, 76)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(81, 13)
            Me.Label5.TabIndex = 18
            Me.Label5.Text = "Design Team"
            '
            'Label7
            '
            Me.Label7.AutoSize = True
            Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label7.Location = New System.Drawing.Point(428, 75)
            Me.Label7.Name = "Label7"
            Me.Label7.Size = New System.Drawing.Size(39, 13)
            Me.Label7.TabIndex = 19
            Me.Label7.Text = "Scale"
            '
            'Label8
            '
            Me.Label8.AutoSize = True
            Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label8.Location = New System.Drawing.Point(427, 104)
            Me.Label8.Name = "Label8"
            Me.Label8.Size = New System.Drawing.Size(40, 13)
            Me.Label8.TabIndex = 20
            Me.Label8.Text = "Sheet"
            '
            'txtBoxCurrentSheet
            '
            Me.txtBoxCurrentSheet.Location = New System.Drawing.Point(476, 101)
            Me.txtBoxCurrentSheet.Name = "txtBoxCurrentSheet"
            Me.txtBoxCurrentSheet.Size = New System.Drawing.Size(39, 20)
            Me.txtBoxCurrentSheet.TabIndex = 17
            '
            'txtBoxTotalSheets
            '
            Me.txtBoxTotalSheets.Location = New System.Drawing.Point(552, 101)
            Me.txtBoxTotalSheets.Name = "txtBoxTotalSheets"
            Me.txtBoxTotalSheets.Size = New System.Drawing.Size(39, 20)
            Me.txtBoxTotalSheets.TabIndex = 19
            '
            'Label9
            '
            Me.Label9.AutoSize = True
            Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label9.Location = New System.Drawing.Point(526, 104)
            Me.Label9.Name = "Label9"
            Me.Label9.Size = New System.Drawing.Size(18, 13)
            Me.Label9.TabIndex = 23
            Me.Label9.Text = "of"
            '
            'Label10
            '
            Me.Label10.AutoSize = True
            Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label10.Location = New System.Drawing.Point(634, 46)
            Me.Label10.Name = "Label10"
            Me.Label10.Size = New System.Drawing.Size(120, 13)
            Me.Label10.TabIndex = 25
            Me.Label10.Text = "Job/Project Number"
            '
            'txtBoxJobNumber
            '
            Me.txtBoxJobNumber.Location = New System.Drawing.Point(757, 43)
            Me.txtBoxJobNumber.Name = "txtBoxJobNumber"
            Me.txtBoxJobNumber.Size = New System.Drawing.Size(115, 20)
            Me.txtBoxJobNumber.TabIndex = 21
            '
            'Label12
            '
            Me.Label12.AutoSize = True
            Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label12.Location = New System.Drawing.Point(614, 104)
            Me.Label12.Name = "Label12"
            Me.Label12.Size = New System.Drawing.Size(141, 13)
            Me.Label12.TabIndex = 5
            Me.Label12.Text = "Molding Machine Model"
            '
            'txtBoxMoldMachineModel
            '
            Me.txtBoxMoldMachineModel.Location = New System.Drawing.Point(757, 101)
            Me.txtBoxMoldMachineModel.Name = "txtBoxMoldMachineModel"
            Me.txtBoxMoldMachineModel.Size = New System.Drawing.Size(115, 20)
            Me.txtBoxMoldMachineModel.TabIndex = 23
            '
            'okButton
            '
            Me.okButton.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.okButton.Location = New System.Drawing.Point(312, 340)
            Me.okButton.Name = "okButton"
            Me.okButton.Size = New System.Drawing.Size(144, 36)
            Me.okButton.TabIndex = 139
            Me.okButton.Text = "Save"
            Me.okButton.UseVisualStyleBackColor = True
            '
            'cancelButton
            '
            Me.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.cancelButton.Location = New System.Drawing.Point(535, 340)
            Me.cancelButton.Name = "cancelButton"
            Me.cancelButton.Size = New System.Drawing.Size(144, 36)
            Me.cancelButton.TabIndex = 141
            Me.cancelButton.Text = "Cancel"
            Me.cancelButton.UseVisualStyleBackColor = True
            '
            'Label77
            '
            Me.Label77.AutoSize = True
            Me.Label77.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label77.Location = New System.Drawing.Point(688, 76)
            Me.Label77.Name = "Label77"
            Me.Label77.Size = New System.Drawing.Size(59, 13)
            Me.Label77.TabIndex = 143
            Me.Label77.Text = "Customer"
            '
            'txtBoxCustomer
            '
            Me.txtBoxCustomer.Location = New System.Drawing.Point(757, 73)
            Me.txtBoxCustomer.Name = "txtBoxCustomer"
            Me.txtBoxCustomer.Size = New System.Drawing.Size(115, 20)
            Me.txtBoxCustomer.TabIndex = 22
            '
            'Button1
            '
            Me.Button1.Location = New System.Drawing.Point(875, 10)
            Me.Button1.Name = "Button1"
            Me.Button1.Size = New System.Drawing.Size(88, 31)
            Me.Button1.TabIndex = 144
            Me.Button1.Text = "Help"
            Me.Button1.UseVisualStyleBackColor = True
            '
            'pdfButton
            '
            Me.pdfButton.Location = New System.Drawing.Point(827, 365)
            Me.pdfButton.Name = "pdfButton"
            Me.pdfButton.Size = New System.Drawing.Size(132, 31)
            Me.pdfButton.TabIndex = 145
            Me.pdfButton.Text = "Print/PDF To Portal"
            Me.pdfButton.UseVisualStyleBackColor = True
            '
            'TabPage3
            '
            Me.TabPage3.BackColor = System.Drawing.SystemColors.Control
            Me.TabPage3.Controls.Add(Me.txtBoxMoldWeight)
            Me.TabPage3.Controls.Add(Me.comboxMoldWeightUnits)
            Me.TabPage3.Controls.Add(Me.Label46)
            Me.TabPage3.Controls.Add(Me.txtBoxMaxMoldOpeningPerSide)
            Me.TabPage3.Controls.Add(Me.Label41)
            Me.TabPage3.Controls.Add(Me.comboxEjectionStrokeUnits)
            Me.TabPage3.Controls.Add(Me.txtBoxEjectionStroke)
            Me.TabPage3.Controls.Add(Me.Label42)
            Me.TabPage3.Controls.Add(Me.comboxQPCModuleShutHeightUnits)
            Me.TabPage3.Controls.Add(Me.txtBoxQPCModuleShutHeight)
            Me.TabPage3.Controls.Add(Me.Label43)
            Me.TabPage3.Controls.Add(Me.comboxHardware)
            Me.TabPage3.Controls.Add(Me.Label14)
            Me.TabPage3.Controls.Add(Me.comboxMoldUnits2)
            Me.TabPage3.Controls.Add(Me.Label11)
            Me.TabPage3.Controls.Add(Me.comboxMoldPitchUnits)
            Me.TabPage3.Controls.Add(Me.txtBoxMoldPitchY)
            Me.TabPage3.Controls.Add(Me.txtBoxMoldPitchX)
            Me.TabPage3.Controls.Add(Me.txtBoxMoldTonnage)
            Me.TabPage3.Controls.Add(Me.txtBoxMoldShutHeight)
            Me.TabPage3.Controls.Add(Me.txtBoxMoldSizeY)
            Me.TabPage3.Controls.Add(Me.txtBoxMoldSizeX)
            Me.TabPage3.Controls.Add(Me.txtBoxMoldCavitation2)
            Me.TabPage3.Controls.Add(Me.txtBoxMoldCavitation1)
            Me.TabPage3.Controls.Add(Me.txtBoxMoldDescription)
            Me.TabPage3.Controls.Add(Me.Label75)
            Me.TabPage3.Controls.Add(Me.Label76)
            Me.TabPage3.Controls.Add(Me.comboxMoldTonnageUnits)
            Me.TabPage3.Controls.Add(Me.comboxMoldShutHeightUnits)
            Me.TabPage3.Controls.Add(Me.comboxMoldSizeUnits)
            Me.TabPage3.Controls.Add(Me.Label25)
            Me.TabPage3.Controls.Add(Me.comboxEjectionType2)
            Me.TabPage3.Controls.Add(Me.Label24)
            Me.TabPage3.Controls.Add(Me.comboxEjectionType1)
            Me.TabPage3.Controls.Add(Me.Label23)
            Me.TabPage3.Controls.Add(Me.comboxStackInterlock)
            Me.TabPage3.Controls.Add(Me.Label22)
            Me.TabPage3.Controls.Add(Me.Label33)
            Me.TabPage3.Controls.Add(Me.Label32)
            Me.TabPage3.Controls.Add(Me.Label29)
            Me.TabPage3.Controls.Add(Me.Label30)
            Me.TabPage3.Controls.Add(Me.Label31)
            Me.TabPage3.Controls.Add(Me.Label40)
            Me.TabPage3.Location = New System.Drawing.Point(4, 22)
            Me.TabPage3.Name = "TabPage3"
            Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
            Me.TabPage3.Size = New System.Drawing.Size(940, 164)
            Me.TabPage3.TabIndex = 2
            Me.TabPage3.Text = "Mold Data"
            '
            'txtBoxMoldWeight
            '
            Me.txtBoxMoldWeight.Location = New System.Drawing.Point(141, 123)
            Me.txtBoxMoldWeight.Name = "txtBoxMoldWeight"
            Me.txtBoxMoldWeight.Size = New System.Drawing.Size(118, 20)
            Me.txtBoxMoldWeight.TabIndex = 152
            '
            'comboxMoldWeightUnits
            '
            Me.comboxMoldWeightUnits.FormattingEnabled = True
            Me.comboxMoldWeightUnits.Location = New System.Drawing.Point(266, 122)
            Me.comboxMoldWeightUnits.Name = "comboxMoldWeightUnits"
            Me.comboxMoldWeightUnits.Size = New System.Drawing.Size(55, 21)
            Me.comboxMoldWeightUnits.TabIndex = 153
            '
            'Label46
            '
            Me.Label46.AutoSize = True
            Me.Label46.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label46.Location = New System.Drawing.Point(7, 126)
            Me.Label46.Name = "Label46"
            Me.Label46.Size = New System.Drawing.Size(128, 13)
            Me.Label46.TabIndex = 151
            Me.Label46.Text = "EST. MOLD WEIGHT"
            '
            'txtBoxMaxMoldOpeningPerSide
            '
            Me.txtBoxMaxMoldOpeningPerSide.Location = New System.Drawing.Point(141, 97)
            Me.txtBoxMaxMoldOpeningPerSide.Name = "txtBoxMaxMoldOpeningPerSide"
            Me.txtBoxMaxMoldOpeningPerSide.Size = New System.Drawing.Size(179, 20)
            Me.txtBoxMaxMoldOpeningPerSide.TabIndex = 150
            '
            'Label41
            '
            Me.Label41.AutoSize = True
            Me.Label41.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label41.Location = New System.Drawing.Point(7, 100)
            Me.Label41.Name = "Label41"
            Me.Label41.Size = New System.Drawing.Size(128, 13)
            Me.Label41.TabIndex = 149
            Me.Label41.Text = "MAX OPENING/SIDE"
            '
            'comboxEjectionStrokeUnits
            '
            Me.comboxEjectionStrokeUnits.FormattingEnabled = True
            Me.comboxEjectionStrokeUnits.Location = New System.Drawing.Point(864, 46)
            Me.comboxEjectionStrokeUnits.Name = "comboxEjectionStrokeUnits"
            Me.comboxEjectionStrokeUnits.Size = New System.Drawing.Size(55, 21)
            Me.comboxEjectionStrokeUnits.TabIndex = 148
            '
            'txtBoxEjectionStroke
            '
            Me.txtBoxEjectionStroke.Location = New System.Drawing.Point(790, 46)
            Me.txtBoxEjectionStroke.Name = "txtBoxEjectionStroke"
            Me.txtBoxEjectionStroke.Size = New System.Drawing.Size(69, 20)
            Me.txtBoxEjectionStroke.TabIndex = 147
            '
            'Label42
            '
            Me.Label42.AutoSize = True
            Me.Label42.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label42.Location = New System.Drawing.Point(663, 49)
            Me.Label42.Name = "Label42"
            Me.Label42.Size = New System.Drawing.Size(121, 13)
            Me.Label42.TabIndex = 146
            Me.Label42.Text = "EJECTION STROKE"
            '
            'comboxQPCModuleShutHeightUnits
            '
            Me.comboxQPCModuleShutHeightUnits.FormattingEnabled = True
            Me.comboxQPCModuleShutHeightUnits.Location = New System.Drawing.Point(594, 123)
            Me.comboxQPCModuleShutHeightUnits.Name = "comboxQPCModuleShutHeightUnits"
            Me.comboxQPCModuleShutHeightUnits.Size = New System.Drawing.Size(55, 21)
            Me.comboxQPCModuleShutHeightUnits.TabIndex = 145
            '
            'txtBoxQPCModuleShutHeight
            '
            Me.txtBoxQPCModuleShutHeight.Location = New System.Drawing.Point(469, 124)
            Me.txtBoxQPCModuleShutHeight.Name = "txtBoxQPCModuleShutHeight"
            Me.txtBoxQPCModuleShutHeight.Size = New System.Drawing.Size(118, 20)
            Me.txtBoxQPCModuleShutHeight.TabIndex = 144
            '
            'Label43
            '
            Me.Label43.AutoSize = True
            Me.Label43.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label43.Location = New System.Drawing.Point(339, 126)
            Me.Label43.Name = "Label43"
            Me.Label43.Size = New System.Drawing.Size(125, 13)
            Me.Label43.TabIndex = 143
            Me.Label43.Text = "QPC  SHUT HEIGHT"
            '
            'comboxHardware
            '
            Me.comboxHardware.FormattingEnabled = True
            Me.comboxHardware.Location = New System.Drawing.Point(789, 9)
            Me.comboxHardware.Name = "comboxHardware"
            Me.comboxHardware.Size = New System.Drawing.Size(130, 21)
            Me.comboxHardware.TabIndex = 142
            '
            'Label14
            '
            Me.Label14.AutoSize = True
            Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label14.Location = New System.Drawing.Point(700, 12)
            Me.Label14.Name = "Label14"
            Me.Label14.Size = New System.Drawing.Size(79, 13)
            Me.Label14.TabIndex = 141
            Me.Label14.Text = "HARDWARE"
            '
            'comboxMoldUnits2
            '
            Me.comboxMoldUnits2.FormattingEnabled = True
            Me.comboxMoldUnits2.Location = New System.Drawing.Point(542, 9)
            Me.comboxMoldUnits2.Name = "comboxMoldUnits2"
            Me.comboxMoldUnits2.Size = New System.Drawing.Size(130, 21)
            Me.comboxMoldUnits2.TabIndex = 115
            '
            'Label11
            '
            Me.Label11.AutoSize = True
            Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label11.Location = New System.Drawing.Point(448, 12)
            Me.Label11.Name = "Label11"
            Me.Label11.Size = New System.Drawing.Size(84, 13)
            Me.Label11.TabIndex = 114
            Me.Label11.Text = "MOLD UNITS"
            '
            'comboxMoldPitchUnits
            '
            Me.comboxMoldPitchUnits.FormattingEnabled = True
            Me.comboxMoldPitchUnits.Location = New System.Drawing.Point(594, 46)
            Me.comboxMoldPitchUnits.Name = "comboxMoldPitchUnits"
            Me.comboxMoldPitchUnits.Size = New System.Drawing.Size(55, 21)
            Me.comboxMoldPitchUnits.TabIndex = 88
            '
            'txtBoxMoldPitchY
            '
            Me.txtBoxMoldPitchY.Location = New System.Drawing.Point(541, 46)
            Me.txtBoxMoldPitchY.Name = "txtBoxMoldPitchY"
            Me.txtBoxMoldPitchY.Size = New System.Drawing.Size(46, 20)
            Me.txtBoxMoldPitchY.TabIndex = 87
            '
            'txtBoxMoldPitchX
            '
            Me.txtBoxMoldPitchX.Location = New System.Drawing.Point(469, 46)
            Me.txtBoxMoldPitchX.Name = "txtBoxMoldPitchX"
            Me.txtBoxMoldPitchX.Size = New System.Drawing.Size(46, 20)
            Me.txtBoxMoldPitchX.TabIndex = 86
            '
            'txtBoxMoldTonnage
            '
            Me.txtBoxMoldTonnage.Location = New System.Drawing.Point(469, 98)
            Me.txtBoxMoldTonnage.Name = "txtBoxMoldTonnage"
            Me.txtBoxMoldTonnage.Size = New System.Drawing.Size(118, 20)
            Me.txtBoxMoldTonnage.TabIndex = 92
            '
            'txtBoxMoldShutHeight
            '
            Me.txtBoxMoldShutHeight.Location = New System.Drawing.Point(141, 72)
            Me.txtBoxMoldShutHeight.Name = "txtBoxMoldShutHeight"
            Me.txtBoxMoldShutHeight.Size = New System.Drawing.Size(97, 20)
            Me.txtBoxMoldShutHeight.TabIndex = 94
            '
            'txtBoxMoldSizeY
            '
            Me.txtBoxMoldSizeY.Location = New System.Drawing.Point(540, 72)
            Me.txtBoxMoldSizeY.Name = "txtBoxMoldSizeY"
            Me.txtBoxMoldSizeY.Size = New System.Drawing.Size(46, 20)
            Me.txtBoxMoldSizeY.TabIndex = 90
            '
            'txtBoxMoldSizeX
            '
            Me.txtBoxMoldSizeX.Location = New System.Drawing.Point(468, 72)
            Me.txtBoxMoldSizeX.Name = "txtBoxMoldSizeX"
            Me.txtBoxMoldSizeX.Size = New System.Drawing.Size(46, 20)
            Me.txtBoxMoldSizeX.TabIndex = 89
            '
            'txtBoxMoldCavitation2
            '
            Me.txtBoxMoldCavitation2.Location = New System.Drawing.Point(244, 46)
            Me.txtBoxMoldCavitation2.Name = "txtBoxMoldCavitation2"
            Me.txtBoxMoldCavitation2.Size = New System.Drawing.Size(77, 20)
            Me.txtBoxMoldCavitation2.TabIndex = 85
            '
            'txtBoxMoldCavitation1
            '
            Me.txtBoxMoldCavitation1.Location = New System.Drawing.Point(141, 46)
            Me.txtBoxMoldCavitation1.Name = "txtBoxMoldCavitation1"
            Me.txtBoxMoldCavitation1.Size = New System.Drawing.Size(77, 20)
            Me.txtBoxMoldCavitation1.TabIndex = 84
            '
            'txtBoxMoldDescription
            '
            Me.txtBoxMoldDescription.Location = New System.Drawing.Point(141, 20)
            Me.txtBoxMoldDescription.Name = "txtBoxMoldDescription"
            Me.txtBoxMoldDescription.Size = New System.Drawing.Size(180, 20)
            Me.txtBoxMoldDescription.TabIndex = 83
            '
            'Label75
            '
            Me.Label75.AutoSize = True
            Me.Label75.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label75.Location = New System.Drawing.Point(522, 49)
            Me.Label75.Name = "Label75"
            Me.Label75.Size = New System.Drawing.Size(19, 13)
            Me.Label75.TabIndex = 112
            Me.Label75.Text = "Y:"
            '
            'Label76
            '
            Me.Label76.AutoSize = True
            Me.Label76.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label76.Location = New System.Drawing.Point(361, 49)
            Me.Label76.Name = "Label76"
            Me.Label76.Size = New System.Drawing.Size(103, 13)
            Me.Label76.TabIndex = 111
            Me.Label76.Text = "MOLD PITCH  X:"
            '
            'comboxMoldTonnageUnits
            '
            Me.comboxMoldTonnageUnits.FormattingEnabled = True
            Me.comboxMoldTonnageUnits.Location = New System.Drawing.Point(594, 97)
            Me.comboxMoldTonnageUnits.Name = "comboxMoldTonnageUnits"
            Me.comboxMoldTonnageUnits.Size = New System.Drawing.Size(55, 21)
            Me.comboxMoldTonnageUnits.TabIndex = 93
            '
            'comboxMoldShutHeightUnits
            '
            Me.comboxMoldShutHeightUnits.FormattingEnabled = True
            Me.comboxMoldShutHeightUnits.Location = New System.Drawing.Point(244, 72)
            Me.comboxMoldShutHeightUnits.Name = "comboxMoldShutHeightUnits"
            Me.comboxMoldShutHeightUnits.Size = New System.Drawing.Size(77, 21)
            Me.comboxMoldShutHeightUnits.TabIndex = 95
            '
            'comboxMoldSizeUnits
            '
            Me.comboxMoldSizeUnits.FormattingEnabled = True
            Me.comboxMoldSizeUnits.Location = New System.Drawing.Point(593, 72)
            Me.comboxMoldSizeUnits.Name = "comboxMoldSizeUnits"
            Me.comboxMoldSizeUnits.Size = New System.Drawing.Size(56, 21)
            Me.comboxMoldSizeUnits.TabIndex = 91
            '
            'Label25
            '
            Me.Label25.AutoSize = True
            Me.Label25.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label25.Location = New System.Drawing.Point(6, 75)
            Me.Label25.Name = "Label25"
            Me.Label25.Size = New System.Drawing.Size(131, 13)
            Me.Label25.TabIndex = 93
            Me.Label25.Text = "MOLD SHUT HEIGHT"
            '
            'comboxEjectionType2
            '
            Me.comboxEjectionType2.FormattingEnabled = True
            Me.comboxEjectionType2.Location = New System.Drawing.Point(789, 97)
            Me.comboxEjectionType2.Name = "comboxEjectionType2"
            Me.comboxEjectionType2.Size = New System.Drawing.Size(130, 21)
            Me.comboxEjectionType2.TabIndex = 102
            '
            'Label24
            '
            Me.Label24.AutoSize = True
            Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label24.Location = New System.Drawing.Point(663, 101)
            Me.Label24.Name = "Label24"
            Me.Label24.Size = New System.Drawing.Size(122, 13)
            Me.Label24.TabIndex = 90
            Me.Label24.Text = "EJECTION TYPE #2"
            '
            'comboxEjectionType1
            '
            Me.comboxEjectionType1.FormattingEnabled = True
            Me.comboxEjectionType1.Location = New System.Drawing.Point(789, 72)
            Me.comboxEjectionType1.Name = "comboxEjectionType1"
            Me.comboxEjectionType1.Size = New System.Drawing.Size(130, 21)
            Me.comboxEjectionType1.TabIndex = 101
            '
            'Label23
            '
            Me.Label23.AutoSize = True
            Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label23.Location = New System.Drawing.Point(662, 75)
            Me.Label23.Name = "Label23"
            Me.Label23.Size = New System.Drawing.Size(122, 13)
            Me.Label23.TabIndex = 88
            Me.Label23.Text = "EJECTION TYPE #1"
            '
            'comboxStackInterlock
            '
            Me.comboxStackInterlock.FormattingEnabled = True
            Me.comboxStackInterlock.Location = New System.Drawing.Point(789, 124)
            Me.comboxStackInterlock.Name = "comboxStackInterlock"
            Me.comboxStackInterlock.Size = New System.Drawing.Size(130, 21)
            Me.comboxStackInterlock.TabIndex = 103
            '
            'Label22
            '
            Me.Label22.AutoSize = True
            Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label22.Location = New System.Drawing.Point(663, 127)
            Me.Label22.Name = "Label22"
            Me.Label22.Size = New System.Drawing.Size(121, 13)
            Me.Label22.TabIndex = 86
            Me.Label22.Text = "STACK INTERLOCK"
            '
            'Label33
            '
            Me.Label33.AutoSize = True
            Me.Label33.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label33.Location = New System.Drawing.Point(521, 75)
            Me.Label33.Name = "Label33"
            Me.Label33.Size = New System.Drawing.Size(20, 13)
            Me.Label33.TabIndex = 85
            Me.Label33.Text = "H:"
            '
            'Label32
            '
            Me.Label32.AutoSize = True
            Me.Label32.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label32.Location = New System.Drawing.Point(366, 75)
            Me.Label32.Name = "Label32"
            Me.Label32.Size = New System.Drawing.Size(98, 13)
            Me.Label32.TabIndex = 82
            Me.Label32.Text = "MOLD SIZE  W:"
            '
            'Label29
            '
            Me.Label29.AutoSize = True
            Me.Label29.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label29.Location = New System.Drawing.Point(225, 51)
            Me.Label29.Name = "Label29"
            Me.Label29.Size = New System.Drawing.Size(13, 13)
            Me.Label29.TabIndex = 80
            Me.Label29.Text = "x"
            '
            'Label30
            '
            Me.Label30.AutoSize = True
            Me.Label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label30.Location = New System.Drawing.Point(15, 49)
            Me.Label30.Name = "Label30"
            Me.Label30.Size = New System.Drawing.Size(120, 13)
            Me.Label30.TabIndex = 78
            Me.Label30.Text = "MOLD CAVITATION"
            '
            'Label31
            '
            Me.Label31.AutoSize = True
            Me.Label31.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label31.Location = New System.Drawing.Point(5, 23)
            Me.Label31.Name = "Label31"
            Me.Label31.Size = New System.Drawing.Size(130, 13)
            Me.Label31.TabIndex = 77
            Me.Label31.Text = "MOLD DESCRIPTION"
            '
            'Label40
            '
            Me.Label40.AutoSize = True
            Me.Label40.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label40.Location = New System.Drawing.Point(359, 101)
            Me.Label40.Name = "Label40"
            Me.Label40.Size = New System.Drawing.Size(106, 13)
            Me.Label40.TabIndex = 66
            Me.Label40.Text = "MOLD TONNAGE"
            '
            'TabPage6
            '
            Me.TabPage6.BackColor = System.Drawing.SystemColors.Control
            Me.TabPage6.Controls.Add(Me.compInfo)
            Me.TabPage6.Location = New System.Drawing.Point(4, 22)
            Me.TabPage6.Name = "TabPage6"
            Me.TabPage6.Padding = New System.Windows.Forms.Padding(3)
            Me.TabPage6.Size = New System.Drawing.Size(940, 164)
            Me.TabPage6.TabIndex = 6
            Me.TabPage6.Text = "Component Information"
            '
            'compInfo
            '
            Me.compInfo.AllowUserToAddRows = False
            Me.compInfo.BackgroundColor = System.Drawing.Color.WhiteSmoke
            Me.compInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.compInfo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.componentInfoDescriptionColumn, Me.componentInfoMaterialColumn, Me.componentInfoHardnessColumn, Me.componentInfoSurfEnhColumn, Me.partName})
            Me.compInfo.Location = New System.Drawing.Point(0, 0)
            Me.compInfo.Name = "compInfo"
            Me.compInfo.Size = New System.Drawing.Size(940, 164)
            Me.compInfo.TabIndex = 76
            '
            'componentInfoDescriptionColumn
            '
            Me.componentInfoDescriptionColumn.HeaderText = "Description"
            Me.componentInfoDescriptionColumn.Name = "componentInfoDescriptionColumn"
            Me.componentInfoDescriptionColumn.Width = 225
            '
            'componentInfoMaterialColumn
            '
            Me.componentInfoMaterialColumn.HeaderText = "Material"
            Me.componentInfoMaterialColumn.Name = "componentInfoMaterialColumn"
            Me.componentInfoMaterialColumn.Width = 225
            '
            'componentInfoHardnessColumn
            '
            Me.componentInfoHardnessColumn.HeaderText = "Hardness"
            Me.componentInfoHardnessColumn.Name = "componentInfoHardnessColumn"
            Me.componentInfoHardnessColumn.Width = 225
            '
            'componentInfoSurfEnhColumn
            '
            Me.componentInfoSurfEnhColumn.HeaderText = "Surface Enhancement"
            Me.componentInfoSurfEnhColumn.Name = "componentInfoSurfEnhColumn"
            Me.componentInfoSurfEnhColumn.Width = 221
            '
            'partName
            '
            Me.partName.HeaderText = "Part Name"
            Me.partName.Name = "partName"
            '
            'TabPage2
            '
            Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
            Me.TabPage2.Controls.Add(Me.titleBlockComps)
            Me.TabPage2.Controls.Add(Me.assyComps)
            Me.TabPage2.Controls.Add(Me.Label37)
            Me.TabPage2.Controls.Add(Me.moveComponentFromTitleBlock)
            Me.TabPage2.Controls.Add(Me.moveComponentToTitleBlock)
            Me.TabPage2.Controls.Add(Me.Label36)
            Me.TabPage2.Location = New System.Drawing.Point(4, 22)
            Me.TabPage2.Name = "TabPage2"
            Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
            Me.TabPage2.Size = New System.Drawing.Size(940, 164)
            Me.TabPage2.TabIndex = 1
            Me.TabPage2.Text = "Components List"
            '
            'titleBlockComps
            '
            Me.titleBlockComps.AllowUserToAddRows = False
            Me.titleBlockComps.BackgroundColor = System.Drawing.Color.WhiteSmoke
            Me.titleBlockComps.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.titleBlockComps.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.titleBlockPartNameColumn, Me.titleBlockComponentColumn})
            Me.titleBlockComps.Location = New System.Drawing.Point(518, 33)
            Me.titleBlockComps.Name = "titleBlockComps"
            DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.titleBlockComps.RowsDefaultCellStyle = DataGridViewCellStyle1
            Me.titleBlockComps.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            Me.titleBlockComps.Size = New System.Drawing.Size(401, 115)
            Me.titleBlockComps.TabIndex = 75
            '
            'titleBlockPartNameColumn
            '
            Me.titleBlockPartNameColumn.HeaderText = "Part Name"
            Me.titleBlockPartNameColumn.Name = "titleBlockPartNameColumn"
            Me.titleBlockPartNameColumn.Width = 178
            '
            'titleBlockComponentColumn
            '
            Me.titleBlockComponentColumn.HeaderText = "Component"
            Me.titleBlockComponentColumn.Name = "titleBlockComponentColumn"
            Me.titleBlockComponentColumn.Width = 180
            '
            'assyComps
            '
            Me.assyComps.AllowUserToAddRows = False
            Me.assyComps.BackgroundColor = System.Drawing.Color.WhiteSmoke
            Me.assyComps.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.assyComps.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.assyPartNameColumn, Me.assyComponentsColumn})
            Me.assyComps.Location = New System.Drawing.Point(36, 33)
            Me.assyComps.Name = "assyComps"
            Me.assyComps.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            Me.assyComps.Size = New System.Drawing.Size(401, 115)
            Me.assyComps.TabIndex = 74
            '
            'assyPartNameColumn
            '
            Me.assyPartNameColumn.HeaderText = "Part Name"
            Me.assyPartNameColumn.Name = "assyPartNameColumn"
            Me.assyPartNameColumn.Width = 178
            '
            'assyComponentsColumn
            '
            Me.assyComponentsColumn.HeaderText = "Component"
            Me.assyComponentsColumn.Name = "assyComponentsColumn"
            Me.assyComponentsColumn.Width = 180
            '
            'Label37
            '
            Me.Label37.AutoSize = True
            Me.Label37.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label37.Location = New System.Drawing.Point(515, 17)
            Me.Label37.Name = "Label37"
            Me.Label37.Size = New System.Drawing.Size(207, 13)
            Me.Label37.TabIndex = 66
            Me.Label37.Text = "TITLE BLOCK COMPONENTS LIST"
            '
            'moveComponentFromTitleBlock
            '
            Me.moveComponentFromTitleBlock.Location = New System.Drawing.Point(457, 95)
            Me.moveComponentFromTitleBlock.Name = "moveComponentFromTitleBlock"
            Me.moveComponentFromTitleBlock.Size = New System.Drawing.Size(40, 23)
            Me.moveComponentFromTitleBlock.TabIndex = 71
            Me.moveComponentFromTitleBlock.Text = "<--"
            Me.moveComponentFromTitleBlock.UseVisualStyleBackColor = True
            '
            'moveComponentToTitleBlock
            '
            Me.moveComponentToTitleBlock.Location = New System.Drawing.Point(457, 66)
            Me.moveComponentToTitleBlock.Name = "moveComponentToTitleBlock"
            Me.moveComponentToTitleBlock.Size = New System.Drawing.Size(40, 23)
            Me.moveComponentToTitleBlock.TabIndex = 69
            Me.moveComponentToTitleBlock.Text = "-->"
            Me.moveComponentToTitleBlock.UseVisualStyleBackColor = True
            '
            'Label36
            '
            Me.Label36.AutoSize = True
            Me.Label36.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label36.Location = New System.Drawing.Point(33, 17)
            Me.Label36.Name = "Label36"
            Me.Label36.Size = New System.Drawing.Size(193, 13)
            Me.Label36.TabIndex = 62
            Me.Label36.Text = "ASSEMBLY COMPONENTS LIST"
            '
            'TabPage4
            '
            Me.TabPage4.BackColor = System.Drawing.SystemColors.Control
            Me.TabPage4.Controls.Add(Me.comboxPartUnits)
            Me.TabPage4.Controls.Add(Me.Label80)
            Me.TabPage4.Controls.Add(Me.Label74)
            Me.TabPage4.Controls.Add(Me.txtBoxPartWSBottom)
            Me.TabPage4.Controls.Add(Me.txtBoxPartLength)
            Me.TabPage4.Controls.Add(Me.txtBoxPartWSSide)
            Me.TabPage4.Controls.Add(Me.txtBoxPartHeight)
            Me.TabPage4.Controls.Add(Me.txtBoxPartWidth)
            Me.TabPage4.Controls.Add(Me.txtBoxPartDiameter)
            Me.TabPage4.Controls.Add(Me.txtBoxPartFlowLength)
            Me.TabPage4.Controls.Add(Me.txtBoxPartLTRatio)
            Me.TabPage4.Controls.Add(Me.txtBoxPartProjectedArea)
            Me.TabPage4.Controls.Add(Me.txtBoxPartVolToBrim)
            Me.TabPage4.Controls.Add(Me.txtBoxPartAppearance)
            Me.TabPage4.Controls.Add(Me.txtBoxPartDensity)
            Me.TabPage4.Controls.Add(Me.txtBoxPartWeight)
            Me.TabPage4.Controls.Add(Me.txtBoxPartShrinkage)
            Me.TabPage4.Controls.Add(Me.txtBoxPartResin)
            Me.TabPage4.Controls.Add(Me.txtBoxPartTitle)
            Me.TabPage4.Controls.Add(Me.Label64)
            Me.TabPage4.Controls.Add(Me.Label63)
            Me.TabPage4.Controls.Add(Me.Label62)
            Me.TabPage4.Controls.Add(Me.Label61)
            Me.TabPage4.Controls.Add(Me.Label60)
            Me.TabPage4.Controls.Add(Me.Label59)
            Me.TabPage4.Controls.Add(Me.Label58)
            Me.TabPage4.Controls.Add(Me.Label56)
            Me.TabPage4.Controls.Add(Me.Label57)
            Me.TabPage4.Controls.Add(Me.Label50)
            Me.TabPage4.Controls.Add(Me.Label47)
            Me.TabPage4.Controls.Add(Me.Label45)
            Me.TabPage4.Controls.Add(Me.Label44)
            Me.TabPage4.Controls.Add(Me.Label39)
            Me.TabPage4.Controls.Add(Me.Label34)
            Me.TabPage4.Controls.Add(Me.Label26)
            Me.TabPage4.Controls.Add(Me.listBoxApplicationsParts)
            Me.TabPage4.Location = New System.Drawing.Point(4, 22)
            Me.TabPage4.Name = "TabPage4"
            Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
            Me.TabPage4.Size = New System.Drawing.Size(940, 164)
            Me.TabPage4.TabIndex = 5
            Me.TabPage4.Text = "Part Data"
            '
            'comboxPartUnits
            '
            Me.comboxPartUnits.FormattingEnabled = True
            Me.comboxPartUnits.Location = New System.Drawing.Point(808, 4)
            Me.comboxPartUnits.Name = "comboxPartUnits"
            Me.comboxPartUnits.Size = New System.Drawing.Size(91, 21)
            Me.comboxPartUnits.TabIndex = 146
            '
            'Label80
            '
            Me.Label80.AutoSize = True
            Me.Label80.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label80.Location = New System.Drawing.Point(737, 7)
            Me.Label80.Name = "Label80"
            Me.Label80.Size = New System.Drawing.Size(63, 13)
            Me.Label80.TabIndex = 147
            Me.Label80.Text = "Part Units"
            '
            'Label74
            '
            Me.Label74.AutoSize = True
            Me.Label74.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label74.Location = New System.Drawing.Point(712, 140)
            Me.Label74.Name = "Label74"
            Me.Label74.Size = New System.Drawing.Size(91, 13)
            Me.Label74.TabIndex = 95
            Me.Label74.Text = "BTM SECTION"
            '
            'txtBoxPartWSBottom
            '
            Me.txtBoxPartWSBottom.Enabled = False
            Me.txtBoxPartWSBottom.Location = New System.Drawing.Point(808, 137)
            Me.txtBoxPartWSBottom.Name = "txtBoxPartWSBottom"
            Me.txtBoxPartWSBottom.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartWSBottom.TabIndex = 90
            '
            'txtBoxPartLength
            '
            Me.txtBoxPartLength.Enabled = False
            Me.txtBoxPartLength.Location = New System.Drawing.Point(865, 60)
            Me.txtBoxPartLength.Name = "txtBoxPartLength"
            Me.txtBoxPartLength.Size = New System.Drawing.Size(34, 20)
            Me.txtBoxPartLength.TabIndex = 92
            '
            'txtBoxPartWSSide
            '
            Me.txtBoxPartWSSide.Enabled = False
            Me.txtBoxPartWSSide.Location = New System.Drawing.Point(808, 111)
            Me.txtBoxPartWSSide.Name = "txtBoxPartWSSide"
            Me.txtBoxPartWSSide.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartWSSide.TabIndex = 89
            '
            'txtBoxPartHeight
            '
            Me.txtBoxPartHeight.Enabled = False
            Me.txtBoxPartHeight.Location = New System.Drawing.Point(808, 84)
            Me.txtBoxPartHeight.Name = "txtBoxPartHeight"
            Me.txtBoxPartHeight.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartHeight.TabIndex = 87
            '
            'txtBoxPartWidth
            '
            Me.txtBoxPartWidth.Enabled = False
            Me.txtBoxPartWidth.Location = New System.Drawing.Point(808, 58)
            Me.txtBoxPartWidth.Name = "txtBoxPartWidth"
            Me.txtBoxPartWidth.Size = New System.Drawing.Size(34, 20)
            Me.txtBoxPartWidth.TabIndex = 85
            '
            'txtBoxPartDiameter
            '
            Me.txtBoxPartDiameter.Enabled = False
            Me.txtBoxPartDiameter.Location = New System.Drawing.Point(808, 32)
            Me.txtBoxPartDiameter.Name = "txtBoxPartDiameter"
            Me.txtBoxPartDiameter.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartDiameter.TabIndex = 83
            '
            'txtBoxPartFlowLength
            '
            Me.txtBoxPartFlowLength.Enabled = False
            Me.txtBoxPartFlowLength.Location = New System.Drawing.Point(581, 137)
            Me.txtBoxPartFlowLength.Name = "txtBoxPartFlowLength"
            Me.txtBoxPartFlowLength.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartFlowLength.TabIndex = 81
            '
            'txtBoxPartLTRatio
            '
            Me.txtBoxPartLTRatio.Enabled = False
            Me.txtBoxPartLTRatio.Location = New System.Drawing.Point(581, 111)
            Me.txtBoxPartLTRatio.Name = "txtBoxPartLTRatio"
            Me.txtBoxPartLTRatio.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartLTRatio.TabIndex = 79
            '
            'txtBoxPartProjectedArea
            '
            Me.txtBoxPartProjectedArea.Enabled = False
            Me.txtBoxPartProjectedArea.Location = New System.Drawing.Point(580, 84)
            Me.txtBoxPartProjectedArea.Name = "txtBoxPartProjectedArea"
            Me.txtBoxPartProjectedArea.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartProjectedArea.TabIndex = 77
            '
            'txtBoxPartVolToBrim
            '
            Me.txtBoxPartVolToBrim.Enabled = False
            Me.txtBoxPartVolToBrim.Location = New System.Drawing.Point(580, 58)
            Me.txtBoxPartVolToBrim.Name = "txtBoxPartVolToBrim"
            Me.txtBoxPartVolToBrim.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartVolToBrim.TabIndex = 75
            '
            'txtBoxPartAppearance
            '
            Me.txtBoxPartAppearance.Enabled = False
            Me.txtBoxPartAppearance.Location = New System.Drawing.Point(580, 32)
            Me.txtBoxPartAppearance.Name = "txtBoxPartAppearance"
            Me.txtBoxPartAppearance.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartAppearance.TabIndex = 73
            '
            'txtBoxPartDensity
            '
            Me.txtBoxPartDensity.Enabled = False
            Me.txtBoxPartDensity.Location = New System.Drawing.Point(336, 137)
            Me.txtBoxPartDensity.Name = "txtBoxPartDensity"
            Me.txtBoxPartDensity.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartDensity.TabIndex = 71
            '
            'txtBoxPartWeight
            '
            Me.txtBoxPartWeight.Enabled = False
            Me.txtBoxPartWeight.Location = New System.Drawing.Point(336, 111)
            Me.txtBoxPartWeight.Name = "txtBoxPartWeight"
            Me.txtBoxPartWeight.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartWeight.TabIndex = 69
            '
            'txtBoxPartShrinkage
            '
            Me.txtBoxPartShrinkage.Enabled = False
            Me.txtBoxPartShrinkage.Location = New System.Drawing.Point(336, 84)
            Me.txtBoxPartShrinkage.Name = "txtBoxPartShrinkage"
            Me.txtBoxPartShrinkage.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartShrinkage.TabIndex = 67
            '
            'txtBoxPartResin
            '
            Me.txtBoxPartResin.Enabled = False
            Me.txtBoxPartResin.Location = New System.Drawing.Point(336, 58)
            Me.txtBoxPartResin.Name = "txtBoxPartResin"
            Me.txtBoxPartResin.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartResin.TabIndex = 65
            '
            'txtBoxPartTitle
            '
            Me.txtBoxPartTitle.Enabled = False
            Me.txtBoxPartTitle.Location = New System.Drawing.Point(336, 32)
            Me.txtBoxPartTitle.Name = "txtBoxPartTitle"
            Me.txtBoxPartTitle.Size = New System.Drawing.Size(91, 20)
            Me.txtBoxPartTitle.TabIndex = 63
            '
            'Label64
            '
            Me.Label64.AutoSize = True
            Me.Label64.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label64.Location = New System.Drawing.Point(849, 61)
            Me.Label64.Name = "Label64"
            Me.Label64.Size = New System.Drawing.Size(13, 13)
            Me.Label64.TabIndex = 93
            Me.Label64.Text = "x"
            '
            'Label63
            '
            Me.Label63.AutoSize = True
            Me.Label63.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label63.Location = New System.Drawing.Point(6, 10)
            Me.Label63.Name = "Label63"
            Me.Label63.Size = New System.Drawing.Size(79, 13)
            Me.Label63.TabIndex = 91
            Me.Label63.Text = "PARTS LIST"
            '
            'Label62
            '
            Me.Label62.AutoSize = True
            Me.Label62.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label62.Location = New System.Drawing.Point(710, 114)
            Me.Label62.Name = "Label62"
            Me.Label62.Size = New System.Drawing.Size(94, 13)
            Me.Label62.TabIndex = 90
            Me.Label62.Text = "SIDE SECTION"
            '
            'Label61
            '
            Me.Label61.AutoSize = True
            Me.Label61.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label61.Location = New System.Drawing.Point(749, 87)
            Me.Label61.Name = "Label61"
            Me.Label61.Size = New System.Drawing.Size(54, 13)
            Me.Label61.TabIndex = 88
            Me.Label61.Text = "HEIGHT"
            '
            'Label60
            '
            Me.Label60.AutoSize = True
            Me.Label60.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label60.Location = New System.Drawing.Point(762, 61)
            Me.Label60.Name = "Label60"
            Me.Label60.Size = New System.Drawing.Size(40, 13)
            Me.Label60.TabIndex = 86
            Me.Label60.Text = "W x L"
            '
            'Label59
            '
            Me.Label59.AutoSize = True
            Me.Label59.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label59.Location = New System.Drawing.Point(731, 35)
            Me.Label59.Name = "Label59"
            Me.Label59.Size = New System.Drawing.Size(71, 13)
            Me.Label59.TabIndex = 84
            Me.Label59.Text = "DIAMETER"
            '
            'Label58
            '
            Me.Label58.AutoSize = True
            Me.Label58.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label58.Location = New System.Drawing.Point(482, 140)
            Me.Label58.Name = "Label58"
            Me.Label58.Size = New System.Drawing.Size(96, 13)
            Me.Label58.TabIndex = 82
            Me.Label58.Text = "FLOW LENGTH"
            '
            'Label56
            '
            Me.Label56.AutoSize = True
            Me.Label56.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label56.Location = New System.Drawing.Point(510, 114)
            Me.Label56.Name = "Label56"
            Me.Label56.Size = New System.Drawing.Size(70, 13)
            Me.Label56.TabIndex = 80
            Me.Label56.Text = "L/T RATIO"
            '
            'Label57
            '
            Me.Label57.AutoSize = True
            Me.Label57.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label57.Location = New System.Drawing.Point(460, 87)
            Me.Label57.Name = "Label57"
            Me.Label57.Size = New System.Drawing.Size(117, 13)
            Me.Label57.TabIndex = 78
            Me.Label57.Text = "PROJECTED AREA"
            '
            'Label50
            '
            Me.Label50.AutoSize = True
            Me.Label50.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label50.Location = New System.Drawing.Point(462, 61)
            Me.Label50.Name = "Label50"
            Me.Label50.Size = New System.Drawing.Size(120, 13)
            Me.Label50.TabIndex = 76
            Me.Label50.Text = "INT. VOL. TO BRIM"
            '
            'Label47
            '
            Me.Label47.AutoSize = True
            Me.Label47.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label47.Location = New System.Drawing.Point(485, 35)
            Me.Label47.Name = "Label47"
            Me.Label47.Size = New System.Drawing.Size(89, 13)
            Me.Label47.TabIndex = 74
            Me.Label47.Text = "APPEARANCE"
            '
            'Label45
            '
            Me.Label45.AutoSize = True
            Me.Label45.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label45.Location = New System.Drawing.Point(270, 140)
            Me.Label45.Name = "Label45"
            Me.Label45.Size = New System.Drawing.Size(61, 13)
            Me.Label45.TabIndex = 72
            Me.Label45.Text = "DENSITY"
            '
            'Label44
            '
            Me.Label44.AutoSize = True
            Me.Label44.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label44.Location = New System.Drawing.Point(241, 114)
            Me.Label44.Name = "Label44"
            Me.Label44.Size = New System.Drawing.Size(94, 13)
            Me.Label44.TabIndex = 70
            Me.Label44.Text = "PART WEIGHT"
            '
            'Label39
            '
            Me.Label39.AutoSize = True
            Me.Label39.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label39.Location = New System.Drawing.Point(252, 87)
            Me.Label39.Name = "Label39"
            Me.Label39.Size = New System.Drawing.Size(79, 13)
            Me.Label39.TabIndex = 68
            Me.Label39.Text = "SHRINKAGE"
            '
            'Label34
            '
            Me.Label34.AutoSize = True
            Me.Label34.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label34.Location = New System.Drawing.Point(285, 61)
            Me.Label34.Name = "Label34"
            Me.Label34.Size = New System.Drawing.Size(45, 13)
            Me.Label34.TabIndex = 66
            Me.Label34.Text = "RESIN"
            '
            'Label26
            '
            Me.Label26.AutoSize = True
            Me.Label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label26.Location = New System.Drawing.Point(288, 35)
            Me.Label26.Name = "Label26"
            Me.Label26.Size = New System.Drawing.Size(42, 13)
            Me.Label26.TabIndex = 64
            Me.Label26.Text = "TITLE"
            '
            'listBoxApplicationsParts
            '
            Me.listBoxApplicationsParts.FormattingEnabled = True
            Me.listBoxApplicationsParts.Location = New System.Drawing.Point(6, 26)
            Me.listBoxApplicationsParts.Name = "listBoxApplicationsParts"
            Me.listBoxApplicationsParts.ScrollAlwaysVisible = True
            Me.listBoxApplicationsParts.Size = New System.Drawing.Size(224, 121)
            Me.listBoxApplicationsParts.TabIndex = 25
            '
            'Tabs
            '
            Me.Tabs.Controls.Add(Me.TabPage4)
            Me.Tabs.Controls.Add(Me.TabPage2)
            Me.Tabs.Controls.Add(Me.TabPage6)
            Me.Tabs.Controls.Add(Me.TabPage3)
            Me.Tabs.Controls.Add(Me.TabPage1)
            Me.Tabs.Controls.Add(Me.TabPage5)
            Me.Tabs.Location = New System.Drawing.Point(15, 137)
            Me.Tabs.Name = "Tabs"
            Me.Tabs.SelectedIndex = 0
            Me.Tabs.Size = New System.Drawing.Size(948, 190)
            Me.Tabs.TabIndex = 26
            '
            'TabPage1
            '
            Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
            Me.TabPage1.Controls.Add(Me.comboxHRManufacturer)
            Me.TabPage1.Controls.Add(Me.comboxGateDiaUnits)
            Me.TabPage1.Controls.Add(Me.Label27)
            Me.TabPage1.Controls.Add(Me.txtBoxGateDia)
            Me.TabPage1.Controls.Add(Me.Label54)
            Me.TabPage1.Controls.Add(Me.comboxGateType)
            Me.TabPage1.Controls.Add(Me.Label28)
            Me.TabPage1.Controls.Add(Me.comboxGateSide)
            Me.TabPage1.Controls.Add(Me.Label35)
            Me.TabPage1.Controls.Add(Me.comboxHRUnits)
            Me.TabPage1.Controls.Add(Me.Label81)
            Me.TabPage1.Controls.Add(Me.comboxPDimColdUnits)
            Me.TabPage1.Controls.Add(Me.comboxPDimHotUnits)
            Me.TabPage1.Controls.Add(Me.comboxLDimUnits)
            Me.TabPage1.Controls.Add(Me.comboxXDimUnits)
            Me.TabPage1.Controls.Add(Me.Label55)
            Me.TabPage1.Controls.Add(Me.txtBoxNozzleTipPartNumber)
            Me.TabPage1.Controls.Add(Me.Label53)
            Me.TabPage1.Controls.Add(Me.txtBoxPDimCold)
            Me.TabPage1.Controls.Add(Me.Label52)
            Me.TabPage1.Controls.Add(Me.txtBoxPDimHot)
            Me.TabPage1.Controls.Add(Me.Label51)
            Me.TabPage1.Controls.Add(Me.txtBoxLDim)
            Me.TabPage1.Controls.Add(Me.Label48)
            Me.TabPage1.Controls.Add(Me.Label49)
            Me.TabPage1.Controls.Add(Me.txtBoxXDim)
            Me.TabPage1.Location = New System.Drawing.Point(4, 22)
            Me.TabPage1.Name = "TabPage1"
            Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
            Me.TabPage1.Size = New System.Drawing.Size(940, 164)
            Me.TabPage1.TabIndex = 7
            Me.TabPage1.Text = "HR Data"
            '
            'comboxHRManufacturer
            '
            Me.comboxHRManufacturer.FormattingEnabled = True
            Me.comboxHRManufacturer.Location = New System.Drawing.Point(738, 40)
            Me.comboxHRManufacturer.Name = "comboxHRManufacturer"
            Me.comboxHRManufacturer.Size = New System.Drawing.Size(158, 21)
            Me.comboxHRManufacturer.TabIndex = 169
            '
            'comboxGateDiaUnits
            '
            Me.comboxGateDiaUnits.FormattingEnabled = True
            Me.comboxGateDiaUnits.Location = New System.Drawing.Point(516, 41)
            Me.comboxGateDiaUnits.Name = "comboxGateDiaUnits"
            Me.comboxGateDiaUnits.Size = New System.Drawing.Size(55, 21)
            Me.comboxGateDiaUnits.TabIndex = 171
            '
            'Label27
            '
            Me.Label27.AutoSize = True
            Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label27.Location = New System.Drawing.Point(601, 44)
            Me.Label27.Name = "Label27"
            Me.Label27.Size = New System.Drawing.Size(131, 13)
            Me.Label27.TabIndex = 175
            Me.Label27.Text = "HR MANUFACTURER"
            '
            'txtBoxGateDia
            '
            Me.txtBoxGateDia.Location = New System.Drawing.Point(413, 41)
            Me.txtBoxGateDia.Name = "txtBoxGateDia"
            Me.txtBoxGateDia.Size = New System.Drawing.Size(97, 20)
            Me.txtBoxGateDia.TabIndex = 170
            '
            'Label54
            '
            Me.Label54.AutoSize = True
            Me.Label54.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label54.Location = New System.Drawing.Point(341, 44)
            Me.Label54.Name = "Label54"
            Me.Label54.Size = New System.Drawing.Size(65, 13)
            Me.Label54.TabIndex = 172
            Me.Label54.Text = "GATE DIA"
            '
            'comboxGateType
            '
            Me.comboxGateType.FormattingEnabled = True
            Me.comboxGateType.Location = New System.Drawing.Point(413, 99)
            Me.comboxGateType.Name = "comboxGateType"
            Me.comboxGateType.Size = New System.Drawing.Size(158, 21)
            Me.comboxGateType.TabIndex = 168
            '
            'Label28
            '
            Me.Label28.AutoSize = True
            Me.Label28.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label28.Location = New System.Drawing.Point(329, 103)
            Me.Label28.Name = "Label28"
            Me.Label28.Size = New System.Drawing.Size(76, 13)
            Me.Label28.TabIndex = 174
            Me.Label28.Text = "GATE TYPE"
            '
            'comboxGateSide
            '
            Me.comboxGateSide.FormattingEnabled = True
            Me.comboxGateSide.Location = New System.Drawing.Point(413, 69)
            Me.comboxGateSide.Name = "comboxGateSide"
            Me.comboxGateSide.Size = New System.Drawing.Size(158, 21)
            Me.comboxGateSide.TabIndex = 167
            '
            'Label35
            '
            Me.Label35.AutoSize = True
            Me.Label35.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label35.Location = New System.Drawing.Point(332, 72)
            Me.Label35.Name = "Label35"
            Me.Label35.Size = New System.Drawing.Size(73, 13)
            Me.Label35.TabIndex = 173
            Me.Label35.Text = "GATE SIDE"
            '
            'comboxHRUnits
            '
            Me.comboxHRUnits.FormattingEnabled = True
            Me.comboxHRUnits.Location = New System.Drawing.Point(769, 6)
            Me.comboxHRUnits.Name = "comboxHRUnits"
            Me.comboxHRUnits.Size = New System.Drawing.Size(148, 21)
            Me.comboxHRUnits.TabIndex = 165
            '
            'Label81
            '
            Me.Label81.AutoSize = True
            Me.Label81.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label81.Location = New System.Drawing.Point(704, 9)
            Me.Label81.Name = "Label81"
            Me.Label81.Size = New System.Drawing.Size(58, 13)
            Me.Label81.TabIndex = 166
            Me.Label81.Text = "HR Units"
            '
            'comboxPDimColdUnits
            '
            Me.comboxPDimColdUnits.FormattingEnabled = True
            Me.comboxPDimColdUnits.Location = New System.Drawing.Point(223, 126)
            Me.comboxPDimColdUnits.Name = "comboxPDimColdUnits"
            Me.comboxPDimColdUnits.Size = New System.Drawing.Size(55, 21)
            Me.comboxPDimColdUnits.TabIndex = 158
            '
            'comboxPDimHotUnits
            '
            Me.comboxPDimHotUnits.FormattingEnabled = True
            Me.comboxPDimHotUnits.Location = New System.Drawing.Point(222, 99)
            Me.comboxPDimHotUnits.Name = "comboxPDimHotUnits"
            Me.comboxPDimHotUnits.Size = New System.Drawing.Size(56, 21)
            Me.comboxPDimHotUnits.TabIndex = 156
            '
            'comboxLDimUnits
            '
            Me.comboxLDimUnits.FormattingEnabled = True
            Me.comboxLDimUnits.Location = New System.Drawing.Point(222, 67)
            Me.comboxLDimUnits.Name = "comboxLDimUnits"
            Me.comboxLDimUnits.Size = New System.Drawing.Size(56, 21)
            Me.comboxLDimUnits.TabIndex = 154
            '
            'comboxXDimUnits
            '
            Me.comboxXDimUnits.FormattingEnabled = True
            Me.comboxXDimUnits.Location = New System.Drawing.Point(222, 38)
            Me.comboxXDimUnits.Name = "comboxXDimUnits"
            Me.comboxXDimUnits.Size = New System.Drawing.Size(56, 21)
            Me.comboxXDimUnits.TabIndex = 152
            '
            'Label55
            '
            Me.Label55.AutoSize = True
            Me.Label55.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label55.Location = New System.Drawing.Point(676, 81)
            Me.Label55.Name = "Label55"
            Me.Label55.Size = New System.Drawing.Size(54, 13)
            Me.Label55.TabIndex = 150
            Me.Label55.Text = "TIP P/N"
            '
            'txtBoxNozzleTipPartNumber
            '
            Me.txtBoxNozzleTipPartNumber.Location = New System.Drawing.Point(738, 69)
            Me.txtBoxNozzleTipPartNumber.Name = "txtBoxNozzleTipPartNumber"
            Me.txtBoxNozzleTipPartNumber.Size = New System.Drawing.Size(158, 20)
            Me.txtBoxNozzleTipPartNumber.TabIndex = 161
            '
            'Label53
            '
            Me.Label53.AutoSize = True
            Me.Label53.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label53.Location = New System.Drawing.Point(672, 68)
            Me.Label53.Name = "Label53"
            Me.Label53.Size = New System.Drawing.Size(60, 13)
            Me.Label53.TabIndex = 149
            Me.Label53.Text = "NOZZLE "
            '
            'txtBoxPDimCold
            '
            Me.txtBoxPDimCold.Location = New System.Drawing.Point(113, 126)
            Me.txtBoxPDimCold.Name = "txtBoxPDimCold"
            Me.txtBoxPDimCold.Size = New System.Drawing.Size(102, 20)
            Me.txtBoxPDimCold.TabIndex = 157
            '
            'Label52
            '
            Me.Label52.AutoSize = True
            Me.Label52.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label52.Location = New System.Drawing.Point(20, 129)
            Me.Label52.Name = "Label52"
            Me.Label52.Size = New System.Drawing.Size(91, 13)
            Me.Label52.TabIndex = 147
            Me.Label52.Text = """P"" DIM COLD"
            '
            'txtBoxPDimHot
            '
            Me.txtBoxPDimHot.Location = New System.Drawing.Point(112, 98)
            Me.txtBoxPDimHot.Name = "txtBoxPDimHot"
            Me.txtBoxPDimHot.Size = New System.Drawing.Size(104, 20)
            Me.txtBoxPDimHot.TabIndex = 155
            '
            'Label51
            '
            Me.Label51.AutoSize = True
            Me.Label51.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label51.Location = New System.Drawing.Point(26, 101)
            Me.Label51.Name = "Label51"
            Me.Label51.Size = New System.Drawing.Size(84, 13)
            Me.Label51.TabIndex = 146
            Me.Label51.Text = """P"" DIM HOT"
            '
            'txtBoxLDim
            '
            Me.txtBoxLDim.Location = New System.Drawing.Point(112, 68)
            Me.txtBoxLDim.Name = "txtBoxLDim"
            Me.txtBoxLDim.Size = New System.Drawing.Size(104, 20)
            Me.txtBoxLDim.TabIndex = 153
            '
            'Label48
            '
            Me.Label48.AutoSize = True
            Me.Label48.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label48.Location = New System.Drawing.Point(54, 71)
            Me.Label48.Name = "Label48"
            Me.Label48.Size = New System.Drawing.Size(53, 13)
            Me.Label48.TabIndex = 145
            Me.Label48.Text = """L"" DIM"
            '
            'Label49
            '
            Me.Label49.AutoSize = True
            Me.Label49.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label49.Location = New System.Drawing.Point(53, 41)
            Me.Label49.Name = "Label49"
            Me.Label49.Size = New System.Drawing.Size(54, 13)
            Me.Label49.TabIndex = 144
            Me.Label49.Text = """X"" DIM"
            '
            'txtBoxXDim
            '
            Me.txtBoxXDim.Location = New System.Drawing.Point(112, 38)
            Me.txtBoxXDim.Name = "txtBoxXDim"
            Me.txtBoxXDim.Size = New System.Drawing.Size(104, 20)
            Me.txtBoxXDim.TabIndex = 151
            '
            'TabPage5
            '
            Me.TabPage5.BackColor = System.Drawing.SystemColors.Control
            Me.TabPage5.Controls.Add(Me.comboxMachineUnits)
            Me.TabPage5.Controls.Add(Me.Label78)
            Me.TabPage5.Controls.Add(Me.Label65)
            Me.TabPage5.Controls.Add(Me.txtBoxMaxShutHeight)
            Me.TabPage5.Controls.Add(Me.comboxMinMaxShutHeightUnits)
            Me.TabPage5.Controls.Add(Me.comboxMaxEjectorStrokeUnits)
            Me.TabPage5.Controls.Add(Me.comboxNozzleRadiusUnits)
            Me.TabPage5.Controls.Add(Me.comboxLocatingRingDiameterUnits)
            Me.TabPage5.Controls.Add(Me.comboxMaxDaylightUnits)
            Me.TabPage5.Controls.Add(Me.comboxClampStrokeUnits)
            Me.TabPage5.Controls.Add(Me.comboxClampTonnageUnits)
            Me.TabPage5.Controls.Add(Me.comboxTieBarDistanceUnits)
            Me.TabPage5.Controls.Add(Me.txtBoxMinShutHeight)
            Me.TabPage5.Controls.Add(Me.Label21)
            Me.TabPage5.Controls.Add(Me.txtBoxMaxEjectorStroke)
            Me.TabPage5.Controls.Add(Me.Label20)
            Me.TabPage5.Controls.Add(Me.txtBoxNozzleRadius)
            Me.TabPage5.Controls.Add(Me.Label19)
            Me.TabPage5.Controls.Add(Me.txtBoxLocatingRingDiameter)
            Me.TabPage5.Controls.Add(Me.Label18)
            Me.TabPage5.Controls.Add(Me.txtBoxMaxDaylight)
            Me.TabPage5.Controls.Add(Me.Label17)
            Me.TabPage5.Controls.Add(Me.txtBoxClampStroke)
            Me.TabPage5.Controls.Add(Me.Label16)
            Me.TabPage5.Controls.Add(Me.txtBoxClampTonnage)
            Me.TabPage5.Controls.Add(Me.Label15)
            Me.TabPage5.Controls.Add(Me.Label1)
            Me.TabPage5.Controls.Add(Me.txtBoxTieBarDistanceVertical)
            Me.TabPage5.Controls.Add(Me.txtBoxTieBarDistanceHorizontal)
            Me.TabPage5.Controls.Add(Me.Label38)
            Me.TabPage5.Location = New System.Drawing.Point(4, 22)
            Me.TabPage5.Name = "TabPage5"
            Me.TabPage5.Size = New System.Drawing.Size(940, 164)
            Me.TabPage5.TabIndex = 8
            Me.TabPage5.Text = "Machine Data"
            '
            'comboxMachineUnits
            '
            Me.comboxMachineUnits.FormattingEnabled = True
            Me.comboxMachineUnits.Location = New System.Drawing.Point(782, 7)
            Me.comboxMachineUnits.Name = "comboxMachineUnits"
            Me.comboxMachineUnits.Size = New System.Drawing.Size(110, 21)
            Me.comboxMachineUnits.TabIndex = 98
            '
            'Label78
            '
            Me.Label78.AutoSize = True
            Me.Label78.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label78.Location = New System.Drawing.Point(688, 10)
            Me.Label78.Name = "Label78"
            Me.Label78.Size = New System.Drawing.Size(88, 13)
            Me.Label78.TabIndex = 99
            Me.Label78.Text = "Machine Units"
            '
            'Label65
            '
            Me.Label65.AutoSize = True
            Me.Label65.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label65.Location = New System.Drawing.Point(713, 134)
            Me.Label65.Name = "Label65"
            Me.Label65.Size = New System.Drawing.Size(13, 13)
            Me.Label65.TabIndex = 81
            Me.Label65.Text = "/"
            '
            'txtBoxMaxShutHeight
            '
            Me.txtBoxMaxShutHeight.Location = New System.Drawing.Point(732, 131)
            Me.txtBoxMaxShutHeight.Name = "txtBoxMaxShutHeight"
            Me.txtBoxMaxShutHeight.Size = New System.Drawing.Size(59, 20)
            Me.txtBoxMaxShutHeight.TabIndex = 96
            '
            'comboxMinMaxShutHeightUnits
            '
            Me.comboxMinMaxShutHeightUnits.FormattingEnabled = True
            Me.comboxMinMaxShutHeightUnits.Location = New System.Drawing.Point(808, 131)
            Me.comboxMinMaxShutHeightUnits.Name = "comboxMinMaxShutHeightUnits"
            Me.comboxMinMaxShutHeightUnits.Size = New System.Drawing.Size(54, 21)
            Me.comboxMinMaxShutHeightUnits.TabIndex = 97
            '
            'comboxMaxEjectorStrokeUnits
            '
            Me.comboxMaxEjectorStrokeUnits.FormattingEnabled = True
            Me.comboxMaxEjectorStrokeUnits.Location = New System.Drawing.Point(808, 100)
            Me.comboxMaxEjectorStrokeUnits.Name = "comboxMaxEjectorStrokeUnits"
            Me.comboxMaxEjectorStrokeUnits.Size = New System.Drawing.Size(54, 21)
            Me.comboxMaxEjectorStrokeUnits.TabIndex = 94
            '
            'comboxNozzleRadiusUnits
            '
            Me.comboxNozzleRadiusUnits.FormattingEnabled = True
            Me.comboxNozzleRadiusUnits.Location = New System.Drawing.Point(808, 70)
            Me.comboxNozzleRadiusUnits.Name = "comboxNozzleRadiusUnits"
            Me.comboxNozzleRadiusUnits.Size = New System.Drawing.Size(54, 21)
            Me.comboxNozzleRadiusUnits.TabIndex = 92
            '
            'comboxLocatingRingDiameterUnits
            '
            Me.comboxLocatingRingDiameterUnits.FormattingEnabled = True
            Me.comboxLocatingRingDiameterUnits.Location = New System.Drawing.Point(808, 40)
            Me.comboxLocatingRingDiameterUnits.Name = "comboxLocatingRingDiameterUnits"
            Me.comboxLocatingRingDiameterUnits.Size = New System.Drawing.Size(54, 21)
            Me.comboxLocatingRingDiameterUnits.TabIndex = 90
            '
            'comboxMaxDaylightUnits
            '
            Me.comboxMaxDaylightUnits.FormattingEnabled = True
            Me.comboxMaxDaylightUnits.Location = New System.Drawing.Point(369, 126)
            Me.comboxMaxDaylightUnits.Name = "comboxMaxDaylightUnits"
            Me.comboxMaxDaylightUnits.Size = New System.Drawing.Size(59, 21)
            Me.comboxMaxDaylightUnits.TabIndex = 88
            '
            'comboxClampStrokeUnits
            '
            Me.comboxClampStrokeUnits.FormattingEnabled = True
            Me.comboxClampStrokeUnits.Location = New System.Drawing.Point(369, 99)
            Me.comboxClampStrokeUnits.Name = "comboxClampStrokeUnits"
            Me.comboxClampStrokeUnits.Size = New System.Drawing.Size(59, 21)
            Me.comboxClampStrokeUnits.TabIndex = 86
            '
            'comboxClampTonnageUnits
            '
            Me.comboxClampTonnageUnits.FormattingEnabled = True
            Me.comboxClampTonnageUnits.Location = New System.Drawing.Point(369, 69)
            Me.comboxClampTonnageUnits.Name = "comboxClampTonnageUnits"
            Me.comboxClampTonnageUnits.Size = New System.Drawing.Size(59, 21)
            Me.comboxClampTonnageUnits.TabIndex = 84
            '
            'comboxTieBarDistanceUnits
            '
            Me.comboxTieBarDistanceUnits.FormattingEnabled = True
            Me.comboxTieBarDistanceUnits.Location = New System.Drawing.Point(369, 39)
            Me.comboxTieBarDistanceUnits.Name = "comboxTieBarDistanceUnits"
            Me.comboxTieBarDistanceUnits.Size = New System.Drawing.Size(59, 21)
            Me.comboxTieBarDistanceUnits.TabIndex = 82
            '
            'txtBoxMinShutHeight
            '
            Me.txtBoxMinShutHeight.Location = New System.Drawing.Point(647, 131)
            Me.txtBoxMinShutHeight.Name = "txtBoxMinShutHeight"
            Me.txtBoxMinShutHeight.Size = New System.Drawing.Size(59, 20)
            Me.txtBoxMinShutHeight.TabIndex = 95
            '
            'Label21
            '
            Me.Label21.AutoSize = True
            Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label21.Location = New System.Drawing.Point(497, 134)
            Me.Label21.Name = "Label21"
            Me.Label21.Size = New System.Drawing.Size(151, 13)
            Me.Label21.TabIndex = 77
            Me.Label21.Text = "MIN/MAX SHUT HEIGHT"
            '
            'txtBoxMaxEjectorStroke
            '
            Me.txtBoxMaxEjectorStroke.Location = New System.Drawing.Point(647, 101)
            Me.txtBoxMaxEjectorStroke.Name = "txtBoxMaxEjectorStroke"
            Me.txtBoxMaxEjectorStroke.Size = New System.Drawing.Size(144, 20)
            Me.txtBoxMaxEjectorStroke.TabIndex = 93
            '
            'Label20
            '
            Me.Label20.AutoSize = True
            Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label20.Location = New System.Drawing.Point(494, 104)
            Me.Label20.Name = "Label20"
            Me.Label20.Size = New System.Drawing.Size(147, 13)
            Me.Label20.TabIndex = 76
            Me.Label20.Text = "MAX EJECTOR STROKE"
            '
            'txtBoxNozzleRadius
            '
            Me.txtBoxNozzleRadius.Location = New System.Drawing.Point(647, 71)
            Me.txtBoxNozzleRadius.Name = "txtBoxNozzleRadius"
            Me.txtBoxNozzleRadius.Size = New System.Drawing.Size(144, 20)
            Me.txtBoxNozzleRadius.TabIndex = 91
            '
            'Label19
            '
            Me.Label19.AutoSize = True
            Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label19.Location = New System.Drawing.Point(535, 74)
            Me.Label19.Name = "Label19"
            Me.Label19.Size = New System.Drawing.Size(107, 13)
            Me.Label19.TabIndex = 75
            Me.Label19.Text = "NOZZLE RADIUS"
            '
            'txtBoxLocatingRingDiameter
            '
            Me.txtBoxLocatingRingDiameter.Location = New System.Drawing.Point(647, 40)
            Me.txtBoxLocatingRingDiameter.Name = "txtBoxLocatingRingDiameter"
            Me.txtBoxLocatingRingDiameter.Size = New System.Drawing.Size(144, 20)
            Me.txtBoxLocatingRingDiameter.TabIndex = 89
            '
            'Label18
            '
            Me.Label18.AutoSize = True
            Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label18.Location = New System.Drawing.Point(514, 43)
            Me.Label18.Name = "Label18"
            Me.Label18.Size = New System.Drawing.Size(129, 13)
            Me.Label18.TabIndex = 74
            Me.Label18.Text = "LOCATING RING DIA"
            '
            'txtBoxMaxDaylight
            '
            Me.txtBoxMaxDaylight.Location = New System.Drawing.Point(209, 129)
            Me.txtBoxMaxDaylight.Name = "txtBoxMaxDaylight"
            Me.txtBoxMaxDaylight.Size = New System.Drawing.Size(144, 20)
            Me.txtBoxMaxDaylight.TabIndex = 87
            '
            'Label17
            '
            Me.Label17.AutoSize = True
            Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label17.Location = New System.Drawing.Point(99, 129)
            Me.Label17.Name = "Label17"
            Me.Label17.Size = New System.Drawing.Size(99, 13)
            Me.Label17.TabIndex = 73
            Me.Label17.Text = "MAX DAYLIGHT"
            '
            'txtBoxClampStroke
            '
            Me.txtBoxClampStroke.Location = New System.Drawing.Point(209, 99)
            Me.txtBoxClampStroke.Name = "txtBoxClampStroke"
            Me.txtBoxClampStroke.Size = New System.Drawing.Size(144, 20)
            Me.txtBoxClampStroke.TabIndex = 85
            '
            'Label16
            '
            Me.Label16.AutoSize = True
            Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label16.Location = New System.Drawing.Point(94, 102)
            Me.Label16.Name = "Label16"
            Me.Label16.Size = New System.Drawing.Size(102, 13)
            Me.Label16.TabIndex = 72
            Me.Label16.Text = "CLAMP STROKE"
            '
            'txtBoxClampTonnage
            '
            Me.txtBoxClampTonnage.Location = New System.Drawing.Point(209, 69)
            Me.txtBoxClampTonnage.Name = "txtBoxClampTonnage"
            Me.txtBoxClampTonnage.Size = New System.Drawing.Size(144, 20)
            Me.txtBoxClampTonnage.TabIndex = 83
            '
            'Label15
            '
            Me.Label15.AutoSize = True
            Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label15.Location = New System.Drawing.Point(84, 73)
            Me.Label15.Name = "Label15"
            Me.Label15.Size = New System.Drawing.Size(112, 13)
            Me.Label15.TabIndex = 71
            Me.Label15.Text = "CLAMP TONNAGE"
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label1.Location = New System.Drawing.Point(276, 42)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(13, 13)
            Me.Label1.TabIndex = 70
            Me.Label1.Text = "x"
            '
            'txtBoxTieBarDistanceVertical
            '
            Me.txtBoxTieBarDistanceVertical.Location = New System.Drawing.Point(302, 39)
            Me.txtBoxTieBarDistanceVertical.Name = "txtBoxTieBarDistanceVertical"
            Me.txtBoxTieBarDistanceVertical.Size = New System.Drawing.Size(51, 20)
            Me.txtBoxTieBarDistanceVertical.TabIndex = 80
            '
            'txtBoxTieBarDistanceHorizontal
            '
            Me.txtBoxTieBarDistanceHorizontal.Location = New System.Drawing.Point(209, 39)
            Me.txtBoxTieBarDistanceHorizontal.Name = "txtBoxTieBarDistanceHorizontal"
            Me.txtBoxTieBarDistanceHorizontal.Size = New System.Drawing.Size(52, 20)
            Me.txtBoxTieBarDistanceHorizontal.TabIndex = 79
            '
            'Label38
            '
            Me.Label38.AutoSize = True
            Me.Label38.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label38.Location = New System.Drawing.Point(41, 43)
            Me.Label38.Name = "Label38"
            Me.Label38.Size = New System.Drawing.Size(161, 13)
            Me.Label38.TabIndex = 69
            Me.Label38.Text = "TIEBAR DISTANCE [H x V]"
            '
            'txtBoxCopies
            '
            Me.txtBoxCopies.Location = New System.Drawing.Point(875, 340)
            Me.txtBoxCopies.Name = "txtBoxCopies"
            Me.txtBoxCopies.Size = New System.Drawing.Size(74, 20)
            Me.txtBoxCopies.TabIndex = 146
            '
            'Label13
            '
            Me.Label13.AutoSize = True
            Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label13.Location = New System.Drawing.Point(785, 343)
            Me.Label13.Name = "Label13"
            Me.Label13.Size = New System.Drawing.Size(84, 13)
            Me.Label13.TabIndex = 148
            Me.Label13.Text = "No. of Copies"
            '
            'btnReset
            '
            Me.btnReset.Location = New System.Drawing.Point(19, 10)
            Me.btnReset.Name = "btnReset"
            Me.btnReset.Size = New System.Drawing.Size(88, 31)
            Me.btnReset.TabIndex = 149
            Me.btnReset.Text = "Reset"
            Me.btnReset.UseVisualStyleBackColor = True
            '
            'btnLoadStack
            '
            Me.btnLoadStack.Location = New System.Drawing.Point(12, 343)
            Me.btnLoadStack.Name = "btnLoadStack"
            Me.btnLoadStack.Size = New System.Drawing.Size(153, 23)
            Me.btnLoadStack.TabIndex = 150
            Me.btnLoadStack.Text = "Load Attributes From Stack"
            Me.btnLoadStack.UseVisualStyleBackColor = True
            '
            'btnLoadSpec
            '
            Me.btnLoadSpec.Location = New System.Drawing.Point(12, 373)
            Me.btnLoadSpec.Name = "btnLoadSpec"
            Me.btnLoadSpec.Size = New System.Drawing.Size(218, 23)
            Me.btnLoadSpec.TabIndex = 151
            Me.btnLoadSpec.Text = "Load Attributes From Stack Specification"
            Me.btnLoadSpec.UseVisualStyleBackColor = True
            '
            'btnLoadMold
            '
            Me.btnLoadMold.Location = New System.Drawing.Point(12, 402)
            Me.btnLoadMold.Name = "btnLoadMold"
            Me.btnLoadMold.Size = New System.Drawing.Size(285, 23)
            Me.btnLoadMold.TabIndex = 152
            Me.btnLoadMold.Text = "Load Attributes From Mold Assy (For Mold Spec Dwg)"
            Me.btnLoadMold.UseVisualStyleBackColor = True
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(975, 432)
            Me.Controls.Add(Me.btnLoadMold)
            Me.Controls.Add(Me.btnLoadSpec)
            Me.Controls.Add(Me.btnLoadStack)
            Me.Controls.Add(Me.btnReset)
            Me.Controls.Add(Me.Label13)
            Me.Controls.Add(Me.txtBoxCopies)
            Me.Controls.Add(Me.pdfButton)
            Me.Controls.Add(Me.Button1)
            Me.Controls.Add(Me.Label77)
            Me.Controls.Add(Me.txtBoxCustomer)
            Me.Controls.Add(Me.Tabs)
            Me.Controls.Add(Me.cancelButton)
            Me.Controls.Add(Me.okButton)
            Me.Controls.Add(Me.Label10)
            Me.Controls.Add(Me.txtBoxJobNumber)
            Me.Controls.Add(Me.Label9)
            Me.Controls.Add(Me.txtBoxTotalSheets)
            Me.Controls.Add(Me.txtBoxCurrentSheet)
            Me.Controls.Add(Me.Label8)
            Me.Controls.Add(Me.Label7)
            Me.Controls.Add(Me.Label5)
            Me.Controls.Add(Me.Label6)
            Me.Controls.Add(Me.txtBoxScale)
            Me.Controls.Add(Me.txtBoxDrawingNumber)
            Me.Controls.Add(Me.txtBoxDate)
            Me.Controls.Add(Me.txtBoxDesignTeam)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.txtBoxDesigner)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.txtBoxMoldMachineModel)
            Me.Controls.Add(Me.Label12)
            Me.Name = "Form1"
            Me.Text = "Assign Stack/Mold Assembly Attributes"
            Me.TabPage3.ResumeLayout(False)
            Me.TabPage3.PerformLayout()
            Me.TabPage6.ResumeLayout(False)
            CType(Me.compInfo, System.ComponentModel.ISupportInitialize).EndInit()
            Me.TabPage2.ResumeLayout(False)
            Me.TabPage2.PerformLayout()
            CType(Me.titleBlockComps, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.assyComps, System.ComponentModel.ISupportInitialize).EndInit()
            Me.TabPage4.ResumeLayout(False)
            Me.TabPage4.PerformLayout()
            Me.Tabs.ResumeLayout(False)
            Me.TabPage1.ResumeLayout(False)
            Me.TabPage1.PerformLayout()
            Me.TabPage5.ResumeLayout(False)
            Me.TabPage5.PerformLayout()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents HelpProvider1 As HelpProvider
        Friend WithEvents Label2 As Label
        Friend WithEvents Label3 As Label
        Friend WithEvents txtBoxDesigner As TextBox
        Private WithEvents txtBoxDesignTeam As TextBox
        Private WithEvents Label4 As Label
        Private WithEvents txtBoxDate As TextBox
        Private WithEvents txtBoxScale As TextBox
        Private WithEvents txtBoxDrawingNumber As TextBox
        Private WithEvents Label6 As Label
        Friend WithEvents Label5 As Label
        Private WithEvents Label7 As Label
        Private WithEvents Label8 As Label
        Private WithEvents txtBoxCurrentSheet As TextBox
        Private WithEvents txtBoxTotalSheets As TextBox
        Private WithEvents Label9 As Label
        Private WithEvents Label10 As Label
        Private WithEvents txtBoxJobNumber As TextBox
        Friend WithEvents Label12 As Label
        Friend WithEvents txtBoxMoldMachineModel As TextBox
        Friend WithEvents okButton As Button
        Friend WithEvents cancelButton As Button
        Private WithEvents Label77 As Label
        Private WithEvents txtBoxCustomer As TextBox
        Friend WithEvents Button1 As Button
        Friend WithEvents pdfButton As Button
        Friend WithEvents TabPage3 As TabPage
        Friend WithEvents comboxMoldUnits2 As ComboBox
        Private WithEvents Label11 As Label
        Friend WithEvents comboxMoldPitchUnits As ComboBox
        Friend WithEvents txtBoxMoldPitchY As TextBox
        Friend WithEvents txtBoxMoldPitchX As TextBox
        Friend WithEvents txtBoxMoldTonnage As TextBox
        Friend WithEvents txtBoxMoldShutHeight As TextBox
        Friend WithEvents txtBoxMoldSizeY As TextBox
        Friend WithEvents txtBoxMoldSizeX As TextBox
        Friend WithEvents txtBoxMoldCavitation2 As TextBox
        Friend WithEvents txtBoxMoldCavitation1 As TextBox
        Friend WithEvents txtBoxMoldDescription As TextBox
        Friend WithEvents Label75 As Label
        Friend WithEvents Label76 As Label
        Friend WithEvents comboxMoldTonnageUnits As ComboBox
        Friend WithEvents comboxMoldShutHeightUnits As ComboBox
        Friend WithEvents comboxMoldSizeUnits As ComboBox
        Friend WithEvents Label25 As Label
        Friend WithEvents comboxEjectionType2 As ComboBox
        Friend WithEvents Label24 As Label
        Friend WithEvents comboxEjectionType1 As ComboBox
        Friend WithEvents Label23 As Label
        Friend WithEvents comboxStackInterlock As ComboBox
        Friend WithEvents Label22 As Label
        Friend WithEvents Label33 As Label
        Friend WithEvents Label32 As Label
        Friend WithEvents Label29 As Label
        Friend WithEvents Label30 As Label
        Friend WithEvents Label31 As Label
        Friend WithEvents Label40 As Label
        Friend WithEvents TabPage6 As TabPage
        Friend WithEvents TabPage2 As TabPage
        Friend WithEvents Label37 As Label
        Friend WithEvents moveComponentFromTitleBlock As Button
        Friend WithEvents moveComponentToTitleBlock As Button
        Friend WithEvents Label36 As Label
        Friend WithEvents TabPage4 As TabPage
        Friend WithEvents comboxPartUnits As ComboBox
        Private WithEvents Label80 As Label
        Friend WithEvents Label74 As Label
        Friend WithEvents txtBoxPartWSBottom As TextBox
        Friend WithEvents txtBoxPartLength As TextBox
        Friend WithEvents txtBoxPartWSSide As TextBox
        Friend WithEvents txtBoxPartHeight As TextBox
        Friend WithEvents txtBoxPartWidth As TextBox
        Friend WithEvents txtBoxPartDiameter As TextBox
        Friend WithEvents txtBoxPartFlowLength As TextBox
        Friend WithEvents txtBoxPartLTRatio As TextBox
        Friend WithEvents txtBoxPartProjectedArea As TextBox
        Friend WithEvents txtBoxPartVolToBrim As TextBox
        Friend WithEvents txtBoxPartAppearance As TextBox
        Friend WithEvents txtBoxPartDensity As TextBox
        Friend WithEvents txtBoxPartWeight As TextBox
        Friend WithEvents txtBoxPartShrinkage As TextBox
        Friend WithEvents txtBoxPartResin As TextBox
        Friend WithEvents txtBoxPartTitle As TextBox
        Friend WithEvents Label64 As Label
        Friend WithEvents Label63 As Label
        Friend WithEvents Label62 As Label
        Friend WithEvents Label61 As Label
        Friend WithEvents Label60 As Label
        Friend WithEvents Label59 As Label
        Friend WithEvents Label58 As Label
        Friend WithEvents Label56 As Label
        Friend WithEvents Label57 As Label
        Friend WithEvents Label50 As Label
        Friend WithEvents Label47 As Label
        Friend WithEvents Label45 As Label
        Friend WithEvents Label44 As Label
        Friend WithEvents Label39 As Label
        Friend WithEvents Label34 As Label
        Friend WithEvents Label26 As Label
        Friend WithEvents listBoxApplicationsParts As ListBox
        Friend WithEvents Tabs As TabControl
        Friend WithEvents txtBoxCopies As TextBox
        Friend WithEvents Label13 As Label
        Friend WithEvents comboxHardware As ComboBox
        Private WithEvents Label14 As Label
        Friend WithEvents TabPage1 As TabPage
        Friend WithEvents TabPage5 As TabPage
        Friend WithEvents btnReset As Button
        Friend WithEvents assyComps As DataGridView
        Friend WithEvents titleBlockComps As DataGridView
        Friend WithEvents comboxMachineUnits As ComboBox
        Private WithEvents Label78 As Label
        Friend WithEvents Label65 As Label
        Friend WithEvents txtBoxMaxShutHeight As TextBox
        Friend WithEvents comboxMinMaxShutHeightUnits As ComboBox
        Friend WithEvents comboxMaxEjectorStrokeUnits As ComboBox
        Friend WithEvents comboxNozzleRadiusUnits As ComboBox
        Friend WithEvents comboxLocatingRingDiameterUnits As ComboBox
        Friend WithEvents comboxMaxDaylightUnits As ComboBox
        Friend WithEvents comboxClampStrokeUnits As ComboBox
        Friend WithEvents comboxClampTonnageUnits As ComboBox
        Friend WithEvents comboxTieBarDistanceUnits As ComboBox
        Friend WithEvents txtBoxMinShutHeight As TextBox
        Friend WithEvents Label21 As Label
        Friend WithEvents txtBoxMaxEjectorStroke As TextBox
        Friend WithEvents Label20 As Label
        Friend WithEvents txtBoxNozzleRadius As TextBox
        Friend WithEvents Label19 As Label
        Friend WithEvents txtBoxLocatingRingDiameter As TextBox
        Friend WithEvents Label18 As Label
        Friend WithEvents txtBoxMaxDaylight As TextBox
        Friend WithEvents Label17 As Label
        Friend WithEvents txtBoxClampStroke As TextBox
        Friend WithEvents Label16 As Label
        Friend WithEvents txtBoxClampTonnage As TextBox
        Friend WithEvents Label15 As Label
        Friend WithEvents Label1 As Label
        Friend WithEvents txtBoxTieBarDistanceVertical As TextBox
        Friend WithEvents txtBoxTieBarDistanceHorizontal As TextBox
        Friend WithEvents Label38 As Label
        Friend WithEvents comboxHRUnits As ComboBox
        Private WithEvents Label81 As Label
        Friend WithEvents comboxPDimColdUnits As ComboBox
        Friend WithEvents comboxPDimHotUnits As ComboBox
        Friend WithEvents comboxLDimUnits As ComboBox
        Friend WithEvents comboxXDimUnits As ComboBox
        Friend WithEvents Label55 As Label
        Friend WithEvents txtBoxNozzleTipPartNumber As TextBox
        Friend WithEvents Label53 As Label
        Friend WithEvents txtBoxPDimCold As TextBox
        Friend WithEvents Label52 As Label
        Friend WithEvents txtBoxPDimHot As TextBox
        Friend WithEvents Label51 As Label
        Friend WithEvents txtBoxLDim As TextBox
        Friend WithEvents Label48 As Label
        Friend WithEvents Label49 As Label
        Friend WithEvents txtBoxXDim As TextBox
        Friend WithEvents comboxHRManufacturer As ComboBox
        Friend WithEvents comboxGateDiaUnits As ComboBox
        Friend WithEvents Label27 As Label
        Friend WithEvents txtBoxGateDia As TextBox
        Friend WithEvents Label54 As Label
        Friend WithEvents comboxGateType As ComboBox
        Friend WithEvents Label28 As Label
        Friend WithEvents comboxGateSide As ComboBox
        Friend WithEvents Label35 As Label
        Friend WithEvents comboxEjectionStrokeUnits As ComboBox
        Friend WithEvents txtBoxEjectionStroke As TextBox
        Friend WithEvents Label42 As Label
        Friend WithEvents comboxQPCModuleShutHeightUnits As ComboBox
        Friend WithEvents txtBoxQPCModuleShutHeight As TextBox
        Friend WithEvents Label43 As Label
        Friend WithEvents txtBoxMaxMoldOpeningPerSide As TextBox
        Friend WithEvents Label41 As Label
        Friend WithEvents titleBlockPartNameColumn As DataGridViewTextBoxColumn
        Friend WithEvents titleBlockComponentColumn As DataGridViewTextBoxColumn
        Friend WithEvents compInfo As DataGridView
        Friend WithEvents txtBoxMoldWeight As TextBox
        Friend WithEvents comboxMoldWeightUnits As ComboBox
        Friend WithEvents Label46 As Label
        Friend WithEvents btnLoadStack As Button
        Friend WithEvents btnLoadSpec As Button
        Friend WithEvents assyPartNameColumn As DataGridViewTextBoxColumn
        Friend WithEvents assyComponentsColumn As DataGridViewTextBoxColumn
        Friend WithEvents componentInfoDescriptionColumn As DataGridViewTextBoxColumn
        Friend WithEvents componentInfoMaterialColumn As DataGridViewTextBoxColumn
        Friend WithEvents componentInfoHardnessColumn As DataGridViewTextBoxColumn
        Friend WithEvents componentInfoSurfEnhColumn As DataGridViewTextBoxColumn
        Friend WithEvents partName As DataGridViewTextBoxColumn
        Friend WithEvents btnLoadMold As Button
    End Class


End Module