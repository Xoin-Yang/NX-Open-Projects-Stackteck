Option Strict Off
Imports System
Imports System.IO
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Annotations
Imports NXOpen.Utilities
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Net.Mail
Imports NXOpen.Assemblies

Module Stackteck_SaveAsNew

    Dim s As Session = Session.GetSession()
    Dim lw As ListingWindow = s.ListingWindow

    Dim theUI As UI = UI.GetUI()
    Dim ufs As UFSession = UFSession.GetUFSession()
    Dim Showlog As Boolean

    Sub Main()
        CreateUsageLog("Stackteck Cloning Program")
        Dim ProgramChoice As Integer = -1

        ProgramChoice = ProgramChoiceList.Clone
        Dim JobFdrLoc As String = Nothing

        Dim myMasterPartInfo As PartInfo = Nothing
        GetOldPartInfo(myMasterPartInfo)

        Dim myOldPartInfoCollection As PartInfo() = Nothing
        Dim myNewPartInfoCollection As PartInfo() = Nothing

        If AskClonePartInfo(myMasterPartInfo, myOldPartInfoCollection, myNewPartInfoCollection, JobFdrLoc) = -1 Then
            Exit Sub
        End If

        CreateCloneLog(myOldPartInfoCollection, myNewPartInfoCollection, JobFdrLoc)
    End Sub

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

    Function AskClonePartInfo(ByVal myMasterPartInfo As PartInfo, ByRef myOldPartInfoCollection As PartInfo(), ByRef myNewPartInfoCollection As PartInfo(), ByRef JobFdrLoc As String) As Integer

        Dim dispPart As Part = s.Parts.Display()

        myMasterPartInfo.ProjNum = myMasterPartInfo.DB_PartNum.Remove(5)

        lw.WriteLine("Master Job Number: " & myMasterPartInfo.ProjNum)
        lw.WriteLine(" ")

        Do
            myMasterPartInfo.ProjNum = NXOpenUI.NXInputBox.GetInputString("Enter Old Job Number", "OLD JOB NUMBER", myMasterPartInfo.ProjNum)

            If myMasterPartInfo.ProjNum = "" Then
                If MessageBox.Show("Quit the Program?", "Exit Program.", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                    Return -1
                    Exit Function
                End If
            End If

        Loop Until myMasterPartInfo.ProjNum <> ""

        Dim NewJobNum As String
        Do
            NewJobNum = NXOpenUI.NXInputBox.GetInputString("Enter New Job Number", "NEW JOB NUMBER", "XXXXX")

            If NewJobNum = "" Then
                If MessageBox.Show("Quit the Program?", "Exit Program.", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                    Return -1
                    Exit Function
                End If
            End If
        Loop Until NewJobNum <> ""

        If AskTCJobFdr(NewJobNum, JobFdrLoc, ProgramChoiceList.Clone) = -1 Then
            Return -1
            Exit Function
        End If

        Dim ComponentNames(0) As String
        Dim ComponentSuppressStatus As Boolean() = Nothing
        Dim RefComponentNames As String() = Nothing
        Dim RefComponentSuppressStatus As Boolean() = Nothing

        Dim Filter As String = myMasterPartInfo.ProjNum

        lw.WriteLine(" ")
        lw.WriteLine("My Filter: " & Filter)

        getassemblytree(ComponentNames, Filter, RefComponentNames, ComponentSuppressStatus, RefComponentSuppressStatus)

        Dim Count As Integer = 0

        lw.Open()
        lw.WriteLine("Number of Job Specific Parts: " & ComponentNames.Length)

        Try
            lw.WriteLine("Number of Reference Parts: " & RefComponentNames.Length)
        Catch ex As Exception
            lw.WriteLine("Number of Reference Parts: 0")
        End Try


        Dim myCloneForm As FormCloneAndRetain = New FormCloneAndRetain

        myCloneForm.txtboxMasterBaseStr.Text = myMasterPartInfo.ProjNum
        myCloneForm.txtboxReplaceStr.Text = NewJobNum

        myCloneForm.ChkBoxAltBaseStr.Checked = False
        myCloneForm.txtboxAltBaseStr1.Visible = False
        myCloneForm.txtboxAltBaseStr2.Visible = False
        myCloneForm.txtboxAltBaseStr3.Visible = False

        Dim i As Integer

        For i = 0 To ComponentNames.Length - 1
            If ComponentSuppressStatus(i) = True Then
                myCloneForm.lstCloneFiles.Items.Add(ComponentNames(i) & "*")
            Else
                myCloneForm.lstCloneFiles.Items.Add(ComponentNames(i))
            End If
        Next

        Try
            For i = 0 To RefComponentNames.Length - 1
                If RefComponentSuppressStatus(i) = True Then
                    myCloneForm.lstRetainFiles.Items.Add(RefComponentNames(i) & "*")
                Else
                    myCloneForm.lstRetainFiles.Items.Add(RefComponentNames(i))
                End If

            Next

        Catch ex As Exception
            lw.WriteLine("No Reference Component is Found!")
        End Try

        If myCloneForm.ShowDialog = DialogResult.OK Then
            ReDim Preserve myOldPartInfoCollection(myCloneForm.lstCloneFiles.Items.Count + myCloneForm.lstRetainFiles.Items.Count - 1)

            i = 0

            For Each myCloneFile As String In myCloneForm.lstCloneFiles.Items
                If myCloneFile.EndsWith("*") Then
                    myOldPartInfoCollection(i).DB_FileName = myCloneFile.Replace("*", "")
                    myOldPartInfoCollection(i).SuppressStatus = True
                Else
                    myOldPartInfoCollection(i).DB_FileName = myCloneFile
                    myOldPartInfoCollection(i).SuppressStatus = False
                End If

                myOldPartInfoCollection(i).CloneStatus = True

                If myCloneFile.Contains(myCloneForm.txtboxMasterBaseStr.Text) = False Then


                    If myCloneFile.Contains(myCloneForm.txtboxAltBaseStr1.Text) = True Then
                        myOldPartInfoCollection(i).ProjNum = myCloneForm.txtboxAltBaseStr1.Text
                    Else

                        If myCloneFile.Contains(myCloneForm.txtboxAltBaseStr2.Text) = True Then
                            myOldPartInfoCollection(i).ProjNum = myCloneForm.txtboxAltBaseStr2.Text
                        Else

                            If myCloneFile.Contains(myCloneForm.txtboxAltBaseStr3.Text) = True Then
                                myOldPartInfoCollection(i).ProjNum = myCloneForm.txtboxAltBaseStr3.Text

                            End If
                        End If
                    End If
                Else

                    myOldPartInfoCollection(i).ProjNum = myCloneForm.txtboxMasterBaseStr.Text

                End If
                ' lw.WriteLine("DB_PART_NAME: " & myOldPartInfoCollection(i).DB_FileName)
                i += 1
            Next

            If myCloneForm.lstRetainFiles.Items.Count <> 0 Then
                For Each myRetainFile As String In myCloneForm.lstRetainFiles.Items
                    If myRetainFile.EndsWith("*") Then
                        myOldPartInfoCollection(i).DB_FileName = myRetainFile.Replace("*", "")
                    Else
                        myOldPartInfoCollection(i).DB_FileName = myRetainFile
                    End If

                    myOldPartInfoCollection(i).CloneStatus = False ''' SOREN CHANGED THIS LINE!
                    '  lw.WriteLine("DB_PART_NAME: " & myOldPartInfoCollection(i).DB_FileName)
                    i += 1
                Next
            End If

            myCloneForm.Close()
            myCloneForm.Dispose()

        Else
            Return -1
            Exit Function
        End If

        For i = 0 To myOldPartInfoCollection.Length - 1

            If myOldPartInfoCollection(i).CloneStatus = True Then
                If myOldPartInfoCollection(i).DB_FileName <> s.Parts.Display.Name Then
                    lw.WriteLine(myOldPartInfoCollection(i).DB_FileName + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    If OpenExistingPart(myOldPartInfoCollection(i).DB_FileName) = False Then
                        Continue For
                    End If
                End If

                lw.WriteLine("")
                lw.WriteLine("  Open: " & myOldPartInfoCollection(i).DB_FileName)
                'myOldPartInfoCollection(i).ProjNum = myMasterPartInfo.ProjNum
                GetOldPartInfo(myOldPartInfoCollection(i))

                ReDim Preserve myNewPartInfoCollection(i)
                myNewPartInfoCollection(i).ProjNum = NewJobNum
                If GetNewPartInfo(myOldPartInfoCollection(i), myNewPartInfoCollection(i), ProgramChoiceList.Clone) <> 0 Then
                    Return -1
                    Exit Function
                End If
                '  GetParentComponent(myOldPartInfoCollection(i))
            End If

        Next

        OpenExistingPart(myMasterPartInfo.DB_FileName)

        CheckandOverrideDuplicateFileName(myMasterPartInfo, myOldPartInfoCollection, myNewPartInfoCollection)

        ExceptionCharRule(myNewPartInfoCollection)

        lw.WriteLine(" ")
        lw.WriteLine("My New Part List....")
        lw.WriteLine("==========================================")
        For Each myNewPartInfo As PartInfo In myNewPartInfoCollection
            ShowPartInfoReportDetails(myNewPartInfo)
        Next

        Return 0

    End Function

    Sub ExceptionCharRule(ByRef myNewPartInfoCollection As PartInfo())

        Dim i As Integer

        For i = 0 To myNewPartInfoCollection.Length - 1
            Try
                If myNewPartInfoCollection(i).CloneStatus = True Then
                    If myNewPartInfoCollection(i).DB_PartDesc.Contains("""") = True Then
                        myNewPartInfoCollection(i).DB_PartDesc = myNewPartInfoCollection(i).DB_PartDesc.Replace("""", " ")
                    End If
                End If
            Catch ex As Exception
            End Try
        Next
    End Sub

    Sub CheckandOverrideDuplicateFileName(ByVal myMasterPartInfo As PartInfo, ByVal myOldPartInfoCollection As PartInfo(), ByRef myNewPartInfoCollection As PartInfo())
        For Each myPartinfo As PartInfo In myNewPartInfoCollection

            Dim count As Integer = 0

            Dim i As Integer
            Dim ArrayIndex As Integer() = Nothing

            For i = 0 To myNewPartInfoCollection.Length - 1

                If myPartinfo.DB_PartNum = myNewPartInfoCollection(i).DB_PartNum Then
                    count += 1
                    ReDim Preserve ArrayIndex(count - 1)
                    ArrayIndex(ArrayIndex.Length - 1) = i
                End If
            Next

            If count > 1 Then

                Dim RenameCount As Integer = 0
                Dim ReaameMasterFileStatus As Boolean = False

                For Each index As Integer In ArrayIndex
                    If myNewPartInfoCollection(index).SuppressStatus = True Then
                        RenameCount += 1
                        myNewPartInfoCollection(index).DB_PartNum = myNewPartInfoCollection(index).DB_PartNum & "_" & RenameCount
                    Else
                        If myOldPartInfoCollection(index).ProjNum <> myMasterPartInfo.ProjNum Then
                            RenameCount += 1
                            myNewPartInfoCollection(index).DB_PartNum = myNewPartInfoCollection(index).DB_PartNum & "_" & RenameCount
                        Else
                            If ReaameMasterFileStatus = True Then
                                RenameCount += 1
                                myNewPartInfoCollection(index).DB_PartNum = myNewPartInfoCollection(index).DB_PartNum & "_" & RenameCount
                            Else
                                ReaameMasterFileStatus = True
                            End If

                        End If
                    End If

                    myNewPartInfoCollection(index).DB_FileName = myNewPartInfoCollection(index).DB_PartNum & "/" & myNewPartInfoCollection(index).PartRev

                Next

            End If

        Next
    End Sub

    Sub ShowPartInfoReportDetails(ByVal myPartInfo As PartInfo)
        lw.WriteLine("Clone Status: " & myPartInfo.CloneStatus)
        lw.WriteLine("DB Part Info: " & myPartInfo.DB_FileName)
        lw.WriteLine("DB Part Number: " & myPartInfo.DB_PartNum)
        lw.WriteLine("DB Part Name: " & myPartInfo.DB_PartName)
        lw.WriteLine("DB Part Description: " & myPartInfo.DB_PartDesc)
        lw.WriteLine("Job Number: " & myPartInfo.ProjNum)
        lw.WriteLine("Part Number: " & myPartInfo.PartNum)
        lw.WriteLine("Drawing Title:" & myPartInfo.StackteckDesc)
        lw.WriteLine("VisDwgNum: " & myPartInfo.VisDwgNum)
        lw.WriteLine("VisDesc: " & myPartInfo.VisDesc)
        lw.WriteLine("Suppress Status: " & myPartInfo.SuppressStatus)
        lw.WriteLine(" ")
    End Sub
    Sub GetOldPartInfo(ByRef myOldPartInfo As PartInfo)
        Dim dispPart As Part = s.Parts.Display()
        find_part_attr_by_name(dispPart, "DB_PART_NO", myOldPartInfo.DB_PartNum)
        find_part_attr_by_name(dispPart, "DB_PART_NAME", myOldPartInfo.DB_PartName)
        find_part_attr_by_name(dispPart, "DB_PART_DESC", myOldPartInfo.DB_PartDesc)
        find_part_attr_by_name(dispPart, "DB_PART_REV", myOldPartInfo.PartRev)
        find_part_attr_by_name(dispPart, "STACKTECK_PARTN", myOldPartInfo.PartNum)
        find_part_attr_by_name(dispPart, "STACKTECK_DESC", myOldPartInfo.StackteckDesc)
        myOldPartInfo.DB_FileName = myOldPartInfo.DB_PartNum & "/" & myOldPartInfo.PartRev
    End Sub

    Function GetNewPartInfo(ByVal myOldPartInfo As PartInfo, ByRef myNewPartInfo As PartInfo, ByVal ProgramChoice As ProgramChoiceList) As Integer

        Dim dispPart As Part = s.Parts.Display()

        myNewPartInfo.CloneStatus = myOldPartInfo.CloneStatus
        myNewPartInfo.SuppressStatus = myOldPartInfo.SuppressStatus

        If ProgramChoice = ProgramChoiceList.SaveAsNew Then
            myNewPartInfo.DB_FileName = myOldPartInfo.DB_FileName
            myNewPartInfo.DB_PartNum = myOldPartInfo.DB_PartNum
            Return 0
            Exit Function
        End If

        If ProgramChoice <> ProgramChoiceList.Clone Then
            Dim PartWithPartNumberStatus As Boolean = True

            If ProgramChoice = ProgramChoiceList.NewPart Then
                If MessageBox.Show("Are you Creating a New Part with Part Number?", "Part With Part Number?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    PartWithPartNumberStatus = False
                Else
                    myOldPartInfo.PartNum = "XXXXX"
                End If
            End If

            If PartWithPartNumberStatus = False Then
                myNewPartInfo.PartNum = "XXXXX"
            Else
                Do
                    myNewPartInfo.PartNum = NXOpenUI.NXInputBox.GetInputString("Enter Part Number", "PART NUMBER", myOldPartInfo.PartNum)

                    If myNewPartInfo.PartNum = "" Then
                        If MessageBox.Show("Quit the Program?", "Exit Program.", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                            Return -1
                            Exit Function
                        End If
                    End If

                Loop Until myNewPartInfo.PartNum <> ""

            End If

        Else
            myNewPartInfo.DB_FileName = myOldPartInfo.DB_FileName
            myNewPartInfo.PartNum = myOldPartInfo.PartNum
        End If

        Dim VisPartDesc As String = Nothing
        Dim VisComCode As String = Nothing
        Dim VisDwgNum As String = Nothing

        If ReadVisData("Y:\eng\ENG_ACCESS_DATABASES\UGMisDatabase.mdb", myNewPartInfo.PartNum, VisPartDesc, VisComCode, VisDwgNum) = True Then
            myNewPartInfo.VisComCode = VisComCode
            myNewPartInfo.VisDwgNum = VisDwgNum
            myNewPartInfo.VisDesc = VisPartDesc

            '            If ReadDBNameData("u:\UGMisDatabase.mdb", myNewPartInfo.PartNum, DBPartName) = False Then

            '            myNewPartInfo.PartDesc = NXOpenUI.NXInputBox.GetInputString("Enter General Description For This Part", "General Discription", VisPartDesc)
            '            Send_New_DBPartName(myNewPartInfo.PartNum, VisPartDesc, myNewPartInfo.PartDesc)
            '       End If

            If myNewPartInfo.VisDwgNum = "X" Or myNewPartInfo.VisDwgNum = "XX" Then
                myNewPartInfo.DB_PartNum = myNewPartInfo.ProjNum & "_" & myNewPartInfo.PartNum

                Select Case ProgramChoice
                    Case ProgramChoiceList.NewPart
                        myNewPartInfo.DB_PartName = NXOpenUI.NXInputBox.GetInputString("Please Enter the Drawing Title", "General Description", myNewPartInfo.VisDesc)

                    Case ProgramChoiceList.SaveAsNew
                        myNewPartInfo.DB_PartName = NXOpenUI.NXInputBox.GetInputString("Please Enter the Drawing Title", "General Description", myOldPartInfo.StackteckDesc)

                    Case ProgramChoiceList.Clone
                        myNewPartInfo.DB_PartName = myOldPartInfo.StackteckDesc
                End Select

                myNewPartInfo.DB_PartDesc = myNewPartInfo.VisDesc
            Else
                Try
                    myNewPartInfo.DB_PartNum = myOldPartInfo.DB_PartNum.Replace(myOldPartInfo.ProjNum, myNewPartInfo.ProjNum)
                    myNewPartInfo.DB_PartName = "XXX"
                    myNewPartInfo.DB_PartDesc = "XXX"
                Catch ex As Exception
                    MsgBox("Something went wrong changing names between:" + myOldPartInfo.ProjNum + " and " + myNewPartInfo.ProjNum)
                End Try
            End If

        Else

            If ProgramChoice <> ProgramChoiceList.Clone Then
                Dim TempStr As String = myNewPartInfo.ProjNum & "_XXX"

                If ProgramChoice = ProgramChoiceList.SaveAsNew Then
                    TempStr = myOldPartInfo.DB_PartNum.Replace(myOldPartInfo.ProjNum, myNewPartInfo.ProjNum)
                End If

                Do
                    myNewPartInfo.DB_PartNum = NXOpenUI.NXInputBox.GetInputString("Please Enter the File Name", "File Name", TempStr)

                    If myNewPartInfo.DB_PartNum = "" Then
                        If MessageBox.Show("Quit the Program?", "Exit Program.", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                            Return -1
                            Exit Function
                        End If
                    End If
                Loop Until myNewPartInfo.DB_PartNum <> ""


            Else
                Try
                    myNewPartInfo.DB_PartNum = myOldPartInfo.DB_PartNum.Replace(myOldPartInfo.ProjNum, myNewPartInfo.ProjNum)
                Catch ex As Exception
                    myNewPartInfo.DB_PartNum = myOldPartInfo.DB_PartName
                End Try
            End If

            myNewPartInfo.DB_PartName = "XXX"  ' THIS MIGHT BE THE PROBLEM
            myNewPartInfo.DB_PartDesc = "XXX"

        End If


        If ProgramChoice = ProgramChoiceList.Clone Then
            myNewPartInfo.PartRev = "000"
        Else
            Do

                myNewPartInfo.PartRev = InputBox("Please Enter Revision Number", "Revision Number", "000")

                If myNewPartInfo.PartRev = "" Then
                    If MessageBox.Show("Quit the Program?", "Exit Program.", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                        Return -1
                        Exit Function
                    End If
                End If
            Loop Until myNewPartInfo.PartRev <> ""
        End If

        If ProgramChoice <> ProgramChoiceList.NewPart Then
            PartInfoExceptionOverride(myOldPartInfo, myNewPartInfo)
        End If

        myNewPartInfo.DB_FileName = myNewPartInfo.DB_PartNum & "/" & myNewPartInfo.PartRev

        Return 0

    End Function

    Sub PartInfoExceptionOverride(ByVal myOldPartInfo As PartInfo, ByRef myNewPartInfo As PartInfo)

        Dim ExceptionLst(4) As String
        ExceptionLst(0) = "ASSY_STA"
        ExceptionLst(1) = "ASSY_STK"
        ExceptionLst(2) = "ASSY_STACK"
        ExceptionLst(3) = "ASSY_COR"
        ExceptionLst(4) = "ASSY_CAV"

        Dim ExceptionStatus As Boolean = False

        For Each myException As String In ExceptionLst
            Try
                If myNewPartInfo.DB_PartNum.Contains(myException) = True Then
                    ExceptionStatus = True
                    Exit For
                End If
            Catch Ex As Exception
            End Try
        Next

        If ExceptionStatus = False Then
            Exit Sub
        End If

        Try
            If IsNumeric(myOldPartInfo.DB_PartName.Remove(5)) = False Then
                myNewPartInfo.DB_PartName = myOldPartInfo.DB_PartName
            End If
        Catch ex As Exception
            myNewPartInfo.DB_PartName = myOldPartInfo.DB_PartName
        End Try

        Try
            If IsNumeric(myOldPartInfo.DB_PartDesc.Remove(5)) = False Then
                myNewPartInfo.DB_PartDesc = myOldPartInfo.DB_PartDesc
            End If
        Catch ex As Exception
            myNewPartInfo.DB_PartDesc = myOldPartInfo.DB_PartDesc
        End Try
    End Sub

    Sub getassemblytree(ByRef componentnames As String(), ByVal jobnum As String, ByRef RefComponentNames As String(), ByRef ComponentSuppressStatus As Boolean(), ByRef RefComponentSuppressStatus As Boolean())

        Dim part1 As Part
        part1 = s.Parts.Work

        Dim c As Component = part1.ComponentAssembly.RootComponent
        lw.WriteLine(c.DisplayName)

        Dim count As Integer = 1
        Dim tempcomponentname(0) As String
        tempcomponentname(0) = c.DisplayName
        Dim SuppressComponentName As String() = Nothing

        ShowAssemblyTree(c, "", count, tempcomponentname, SuppressComponentName)

        lw.WriteLine("total count: " & count)
        Try
            lw.WriteLine("Total Suppress Components: " & SuppressComponentName.Length)
        Catch ex As Exception
            lw.WriteLine("Total Suppress Components:  None")
        End Try

        lw.WriteLine("")

        componentnames(0) = c.DisplayName
        Dim i As Integer = 0
        Dim j As Integer = 0

        For Each tempstr As String In tempcomponentname
            If tempstr.StartsWith(jobnum) = True Then

                Dim componentfound As Boolean = False

                For Each componentname As String In componentnames
                    If tempstr = componentname Then
                        componentfound = True
                        Exit For
                    End If
                Next

                If componentfound = False Then
                    i += 1
                    ReDim Preserve componentnames(i)
                    componentnames(i) = tempstr

                    ReDim Preserve ComponentSuppressStatus(i)
                    ComponentSuppressStatus(i) = False
                End If
            Else
                If j = 0 Then
                    ReDim Preserve RefComponentNames(0)
                    RefComponentNames(0) = tempstr

                    ReDim Preserve RefComponentSuppressStatus(0)
                    RefComponentSuppressStatus(0) = False

                    j += 1
                Else
                    Dim Refcomponentfound As Boolean = False

                    For Each Refcomponentname As String In RefComponentNames
                        If tempstr = Refcomponentname Then
                            Refcomponentfound = True
                            Exit For
                        End If
                    Next

                    If Refcomponentfound = False Then
                        j += 1
                        ReDim Preserve RefComponentNames(j - 1)
                        RefComponentNames(j - 1) = tempstr

                        ReDim Preserve RefComponentSuppressStatus(j - 1)
                        RefComponentSuppressStatus(j - 1) = False
                    End If
                End If


            End If

        Next

        Try
            For Each tempstr As String In SuppressComponentName
                If tempstr.StartsWith(jobnum) = True Then

                    Dim componentfound As Boolean = False

                    For Each componentname As String In componentnames
                        If tempstr = componentname Then
                            componentfound = True
                            Exit For
                        End If
                    Next

                    If componentfound = False Then
                        i += 1
                        ReDim Preserve componentnames(i)
                        componentnames(i) = tempstr

                        ReDim Preserve ComponentSuppressStatus(i)
                        ComponentSuppressStatus(i) = True
                    End If
                Else
                    If j = 0 Then
                        ReDim Preserve RefComponentNames(0)
                        RefComponentNames(0) = tempstr

                        ReDim Preserve RefComponentSuppressStatus(0)
                        RefComponentSuppressStatus(0) = True

                        j += 1
                    Else
                        Dim Refcomponentfound As Boolean = False

                        For Each Refcomponentname As String In RefComponentNames
                            If tempstr = Refcomponentname Then
                                Refcomponentfound = True
                                Exit For
                            End If
                        Next

                        If Refcomponentfound = False Then
                            j += 1
                            ReDim Preserve RefComponentNames(j - 1)
                            RefComponentNames(j - 1) = tempstr

                            ReDim Preserve RefComponentSuppressStatus(j - 1)
                            RefComponentSuppressStatus(j - 1) = True
                        End If
                    End If


                End If

            Next
        Catch ex As Exception
            lw.WriteLine(" ")
            lw.WriteLine("No Suppress Component")
        End Try


        lw.WriteLine("Component Name")
        lw.WriteLine("========================================")
        For Each myComponentName As String In componentnames
            lw.WriteLine(myComponentName)
        Next

        lw.WriteLine(" ")
        lw.WriteLine("Reference Component Name")
        lw.WriteLine("========================================")
        Try
            For Each myRefComponentName As String In RefComponentNames
                lw.WriteLine(myRefComponentName)
            Next
        Catch ex As Exception
            lw.WriteLine("No Component is Found!")
        End Try
    End Sub

    Sub ShowAssemblyTree(ByVal c As Component, ByVal indent As String, ByRef count As Integer, ByRef TempComponentName As String(), ByRef SuppressComponentName As String())
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

                ShowAssemblyTree(child, newIndent, count, TempComponentName, SuppressComponentName)
            Else
                lw.WriteLine(newIndent & child.DisplayName & " is in suppress status")
                Try
                    If SuppressComponentName.Length > 0 Then
                        ReDim Preserve SuppressComponentName(SuppressComponentName.Length)
                    End If
                Catch ex As Exception
                    ReDim Preserve SuppressComponentName(0)
                End Try
                SuppressComponentName(SuppressComponentName.Length - 1) = child.DisplayName

                ShowAssemblyTree(child, newIndent, count, TempComponentName, SuppressComponentName)
            End If
        Next
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
            VisPartDesc = dr(1)
            VisPartDesc = VisPartDesc.Trim()

            VisComCode = dr(2)
            VisComCode = VisComCode.Trim()

            VisDwgNum = dr(3)
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

        lw.WriteLine("MADE IT TO HERE!!!!!!!!!!!!!!!!!!!!!!!!!!!")

        'Dim part1 As Part = Nothing

        Try
            Dim part1 As Part = CType(s.Parts.FindObject("@DB/" & FileName), Part)

            Dim partLoadStatus2 As PartLoadStatus = Nothing
            Dim status1 As PartCollection.SdpsStatus
            status1 = s.Parts.SetDisplay(part1, True, True, partLoadStatus2)

            s.Parts.SetWork(part1)

            Return True

        Catch ex As Exception
            MsgBox("Unable to find: " & FileName)
            Return False
        End Try

        'Dim part1 As Part = CType(s.Parts.FindObject("@DB/" & FileName), Part)

        'Dim partLoadStatus2 As PartLoadStatus = Nothing
        'Dim status1 As PartCollection.SdpsStatus
        'status1 = s.Parts.SetDisplay(part1, True, True, partLoadStatus2)

        's.Parts.SetWork(part1)

        'Return True

    End Function

    Structure PartInfo
        Dim DB_FileName As String
        Dim DB_PartNum As String
        Dim DB_PartName As String
        Dim DB_PartDesc As String
        Dim ProjNum As String
        Dim PartNum As String
        Dim PartRev As String
        Dim StackteckDesc As String
        Dim VisDesc As String
        Dim VisDwgNum As String
        Dim VisComCode As String
        Dim CloneStatus As Boolean
        Dim SuppressStatus As Boolean
    End Structure

    Enum ProgramChoiceList
        NewPart = 0
        SaveAsNew = 1
        Clone = 2
    End Enum

    Sub GetParentComponent(ByRef myOldPartInfoCollection As PartInfo())

        Dim dp As Part = s.Parts.Display

        Dim c As Component() = getAllChildren(dp)

        For Each myComp As Component In c
            Dim theParent As Component = myComp.OwningComponent

            lw.WriteLine(myComp.DisplayName & " under " & theParent.DisplayName)
        Next
    End Sub

    Function getAllChildren(ByVal assy As BasePart) As Assemblies.Component()
        Dim theChildren As Collections.ArrayList = New Collections.ArrayList
        Dim aChildTag As Tag = Tag.Null

        Do
            ufs.Obj.CycleObjsInPart(assy.Tag, UFConstants.UF_component_type, aChildTag)
            If (aChildTag = Tag.Null) Then Exit Do

            Dim aChild As Assemblies.Component = NXObjectManager.Get(aChildTag)
            theChildren.Add(aChild)
        Loop
        Return theChildren.ToArray(GetType(Assemblies.Component))

    End Function

    Sub CreateCloneLog(ByVal myOldPartInfoCollection As PartInfo(), ByVal myNewPartInfoCollection As PartInfo(), ByVal JobFdrLoc As String)

        Dim NewJobNum As String = myNewPartInfoCollection(0).ProjNum
        Dim FolderLocation As String = JobFdrLoc
        Dim RevisionRule As String = "Latest by Alpha Rev Order"
        Dim NamingTechnique As String = "USER_NAME"
        Dim CloneAction As String = "CLONE"

        Dim myCloneFolder As String = "C:\eng\UG_Clone_Log\"
        Dim myCloneFileName As String = myCloneFolder & NewJobNum & "_" & Format(Date.Now, "yyyy-MM-dd") & "_CloneLog.clone"

        If Not Directory.Exists(myCloneFolder) Then
            Directory.CreateDirectory(myCloneFolder)
        End If

        lw.WriteLine(" ")
        lw.WriteLine("Log File Location: " & myCloneFileName)

        Dim myWriterObj As StreamWriter = New StreamWriter(myCloneFileName)

        myWriterObj.WriteLine("Assembly Cloning Log File")
        myWriterObj.WriteLine("&LOG Operation_Type: CLONING_OPERATION")
        myWriterObj.WriteLine("Revision rule current when this operation was performed: " & """" & RevisionRule & """")
        myWriterObj.WriteLine("&LOG Default_Cloning_Action: " & CloneAction)
        myWriterObj.WriteLine("&LOG Default_Naming_Technique: " & NamingTechnique)
        myWriterObj.WriteLine("&LOG Naming_Rule_Type: REPLACE_STRING Base_String: " & myOldPartInfoCollection(0).ProjNum & " Replacement_String: " & myNewPartInfoCollection(0).ProjNum)
        myWriterObj.WriteLine("&LOG Default_Container: " & """" & FolderLocation & """")
        myWriterObj.WriteLine("&LOG Default_Directory: """"")
        myWriterObj.WriteLine("&LOG Default_Part_Type: """"")
        myWriterObj.WriteLine("&LOG Default_Part_Name: """"")
        myWriterObj.WriteLine("&LOG Default_Part_Description: """"")

        myWriterObj.WriteLine("&LOG Default_Copy_Associated_Files: Yes")
        myWriterObj.WriteLine("&LOG Default_Non_Master_Copy: specification Yes")
        myWriterObj.WriteLine("&LOG Default_Non_Master_Copy: manifestation Yes")
        myWriterObj.WriteLine("&LOG Default_Non_Master_Copy: altrep Yes")
        myWriterObj.WriteLine("&LOG Default_Non_Master_Copy: scenario Yes")
        myWriterObj.WriteLine("&LOG Default_Non_Master_Copy: simulation Yes")
        myWriterObj.WriteLine("&LOG Default_Non_Master_Copy: cae_motion Yes")
        myWriterObj.WriteLine("&LOG Default_Non_Master_Copy: cae_solution Yes")
        myWriterObj.WriteLine("&LOG Default_Non_Master_Copy: cae_mesh Yes")
        myWriterObj.WriteLine("&LOG Default_Non_Master_Copy: cae_geometry Yes")
        myWriterObj.WriteLine("&LOG ")

        Dim i As Integer
        For i = 0 To myOldPartInfoCollection.Length - 1
            ' lw.WriteLine("Part #" & i)
            If myOldPartInfoCollection(i).CloneStatus = True Then

                myWriterObj.WriteLine("&LOG Part: @DB/" & myOldPartInfoCollection(i).DB_FileName)
                myWriterObj.WriteLine("&LOG Cloning_Action: DEFAULT_DISP Naming_Technique: DEFAULT_NAMING Clone_Name: @DB/" & myNewPartInfoCollection(i).DB_FileName)
                myWriterObj.WriteLine("&LOG Part_Type: Item")
                myWriterObj.WriteLine("&LOG Part_Name: " & """" & myNewPartInfoCollection(i).DB_PartName & """")
                myWriterObj.WriteLine("&LOG Part_Description: " & """" & myNewPartInfoCollection(i).DB_PartDesc & """")
                myWriterObj.WriteLine("&LOG ")

            Else
                myWriterObj.WriteLine("&LOG Part: @DB/" & myOldPartInfoCollection(i).DB_FileName)
                myWriterObj.WriteLine("&LOG Cloning_Action: RETAIN")
                myWriterObj.WriteLine("&LOG ")
                myWriterObj.WriteLine("&LOG Part: @DB/" & myOldPartInfoCollection(i).DB_FileName + "/specification/dwg")
                myWriterObj.WriteLine("&LOG Cloning_Action: RETAIN")
                myWriterObj.WriteLine("&LOG ")
                myWriterObj.WriteLine("&LOG Part: @DB/" & myOldPartInfoCollection(i).DB_FileName + "/specification/dwg_H")
                myWriterObj.WriteLine("&LOG Cloning_Action: RETAIN")
                myWriterObj.WriteLine("&LOG ")
                myWriterObj.WriteLine("&LOG Part: @DB/" & myOldPartInfoCollection(i).DB_FileName + "/specification/dwg_V")
                myWriterObj.WriteLine("&LOG Cloning_Action: RETAIN")
                myWriterObj.WriteLine("&LOG ")
                myWriterObj.WriteLine("&LOG Part: @DB/" & myOldPartInfoCollection(i).DB_FileName + "/specification/CAM")
                myWriterObj.WriteLine("&LOG Cloning_Action: RETAIN")
                myWriterObj.WriteLine("&LOG ")
            End If
        Next

        myWriterObj.Close()
        myWriterObj.Dispose()

        MessageBox.Show("Successfully create Clone Log under " & myCloneFileName, "Clone Log Created", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        System.Diagnostics.Process.Start("explorer.exe", myCloneFolder)

    End Sub

    Public Sub SetAttribute(ByVal title As String, ByRef value As String)
        Dim dispPart As Part = s.Parts.Display()
        dispPart.SetAttribute(title, value)
    End Sub

    Function AskTCJobFdr(ByRef myProjNum As String, ByRef JobFdrLoc As String, ByVal ProgramChoice As ProgramChoiceList) As Integer

        Dim UserName As String = System.Environment.UserName

        Dim StartOfProNum As Integer = Nothing
        Dim EndOfProNum As Integer = Nothing

        If IsNumeric(myProjNum) = True Then
            StartOfProNum = Math.Floor(myProjNum / 1000) * 1000
            EndOfProNum = StartOfProNum + 999
        End If

        Select Case ProgramChoice
            Case ProgramChoiceList.Clone
                If IsNumeric(myProjNum) = False Or myProjNum.Length > 5 Then
                    JobFdrLoc = UserName & ":Newstuff:" & myProjNum
                    MsgBox("Something went wrong with the project number 1")
                Else
                    JobFdrLoc = UserName & ":" & "MOLD DESIGN JOB LIBRARY:MOLD JOBS " & Format(StartOfProNum, "00000") & "-" & Format(EndOfProNum, "00000") & ":" & Format(Convert.ToInt32(myProjNum), "00000") & ":" & Format(Convert.ToInt32(myProjNum), "00000") & "_STACK"
                End If

            Case Else
                If IsNumeric(myProjNum) = False Or myProjNum.Length > 5 Then
                    JobFdrLoc = UserName & ":Newstuff:" & myProjNum
                    MsgBox("Something went wrong with the project number 1")
                Else
                    If MessageBox.Show("Is this a Stack Component?", "STACK COMPONENT?", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                        JobFdrLoc = UserName & ":" & "MOLD DESIGN JOB LIBRARY:MOLD JOBS " & Format(StartOfProNum, "00000") & "-" & Format(EndOfProNum, "00000") & ":" & Format(Convert.ToInt32(myProjNum), "00000") & ":" & Format(Convert.ToInt32(myProjNum), "00000") & "_STACK"
                    Else
                        JobFdrLoc = UserName & ":" & "MOLD DESIGN JOB LIBRARY:MOLD JOBS " & Format(StartOfProNum, "00000") & "-" & Format(EndOfProNum, "00000") & ":" & Format(Convert.ToInt32(myProjNum), "00000") & ":" & Format(Convert.ToInt32(myProjNum), "00000") & "_SHOE"
                    End If
                End If
        End Select

        If MessageBox.Show("The default saving folder is in : " & JobFdrLoc & ". Is this correct?", "Change Default Saving Folder", MessageBoxButtons.YesNo) = DialogResult.No Then
            JobFdrLoc = InputBox("Please Enter the Folder Location with Proper Format.", "Folder Location", JobFdrLoc)
        End If

        If AskTcJobFdrExist(myProjNum, JobFdrLoc) = False Then
            MessageBox.Show(JobFdrLoc & " is not exist in the database.  Please create the folder and try again!")
            Return -1
            Exit Function
        End If

        lw.Open()
        lw.WriteLine(" ")
        lw.WriteLine("Teamcenter Folder Location: " & JobFdrLoc)

        Return 0
    End Function

    Function AskTcJobFdrExist(ByVal ProjNum As String, ByVal JobFdrLoc As String) As Boolean
        Dim Root_Tag As Tag
        ufs.Ugmgr.AskRootFolder(Root_Tag)

        Dim count As Integer = Nothing
        Dim Folder_contents() As Tag = Nothing

        Dim FolderName As String() = JobFdrLoc.Split(":")

        Dim i As Integer

        lw.WriteLine(FolderName.Length)

        For i = 0 To FolderName.Length - 1
            lw.WriteLine("Folder Level " & i & ": " & FolderName(i))
        Next


        Dim FolderFoundStatus As Boolean = False

        Dim FolderTag(FolderName.Length - 1) As Tag

        For i = 0 To FolderName.Length - 1
            If i = 0 Then
                If SearchTcFolder(Root_Tag, FolderName(i + 1), FolderTag(i + 1)) = True Then
                    If i + 1 = FolderName.Length - 1 Then
                        FolderFoundStatus = True
                        Exit For
                    Else
                        Continue For
                    End If

                Else
                    Exit For
                End If
            Else
                If SearchTcFolder(FolderTag(i), FolderName(i + 1), FolderTag(i + 1)) = True Then
                    If i + 1 = FolderName.Length - 1 Then
                        FolderFoundStatus = True
                        Exit For
                    Else
                        Continue For
                    End If
                Else
                    Exit For
                End If

            End If
        Next

        If FolderFoundStatus = True Then
            lw.WriteLine("Folder Path " & JobFdrLoc & " exist.")
            Return True
        Else
            lw.WriteLine("Folder Path " & JobFdrLoc & " is NOT exist.")
            Return False
        End If

    End Function

    Function SearchTcFolder(ByVal FolderTag As Tag, ByVal FolderName As String, ByRef FoundFolderTag As Tag) As Boolean
        Dim Count As Integer
        Dim Folder_contents As Tag() = Nothing

        ufs.Ugmgr.ListFolderContents(FolderTag, Count, Folder_contents)

        Dim FolderFindStatus As Boolean = False

        For Each content As Tag In Folder_contents
            Dim ObjectType As UFUgmgr.ObjectType
            ufs.Ugmgr.AskObjectType(content, ObjectType)

            If ObjectType = UFUgmgr.ObjectType.TypeFolder Then
                Dim TempFolderName As String = Nothing
                ufs.Ugmgr.AskFolderName(content, TempFolderName)
                lw.WriteLine("Folder Name: " & TempFolderName)

                If TempFolderName = FolderName Then
                    FoundFolderTag = content
                    Return True
                    Exit Function
                End If
            End If
        Next
        Return False
    End Function

    Sub SetNonDBPartAttr(ByVal myPartInfo As PartInfo)
        SetAttribute("STACKTECK_PARTN", myPartInfo.PartNum)
        SetAttribute("STACKTECK_DESC", myPartInfo.DB_PartName)
        SetAttribute("STACKTECK_PROJNUM", myPartInfo.ProjNum)
        SetAttribute("STACKTECK_VISDESC", myPartInfo.VisDesc)
        SetAttribute("STACKTECK_VISCOMCODE", myPartInfo.VisComCode)
        SetAttribute("STACKTECK_VISDWGNUM", myPartInfo.VisDwgNum)
        SetAttribute("TITBLK_DATE", Format(Date.Now, "dd-MMM-yyyy"))
    End Sub

    Public Class FormCloneAndRetain

        Dim itemsToBeRemoved(0) As String
        Dim setSelectedComponentsRun As Boolean = False

        Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
            If CheckReplaceString() = False Then
                Exit Sub
            End If

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
        End Sub

        Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        End Sub

        Private Sub lstCloneFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstCloneFiles.SelectedIndexChanged

            If (setSelectedComponentsRun = True) Then
                Exit Sub
            End If

            ' Get selected item
            ' Search through all components to get sub components
            ' Select All subcomponents 

            Dim curItem As String = lstCloneFiles.SelectedItem.ToString()

            Dim part1 As Part
            part1 = s.Parts.Work

            Dim c As Component = part1.ComponentAssembly.RootComponent

            GetAllSubComponentsForSelection(c, curItem, False)

            SelectSubComponents()

        End Sub

        Private Sub SelectSubComponents()
            setSelectedComponentsRun = True

            For Each item As String In itemsToBeRemoved
                For i As Integer = 0 To lstCloneFiles.Items.Count - 1
                    If (item = lstCloneFiles.Items(i)) Then
                        lstCloneFiles.SetSelected(i, True)
                    End If
                Next
            Next

            setSelectedComponentsRun = False
            Array.Clear(itemsToBeRemoved, 0, itemsToBeRemoved.Length - 1)
        End Sub

        Private Function GetAllSubComponentsForSelection(ByVal comp As Component, ByVal topName As String, ByVal topFound As Boolean)
            For Each child As Component In comp.GetChildren()

                If (child.DisplayName().Trim = topName) Then
                    GetAllSubComponentsForSelection(child, topName, True)
                ElseIf (topFound = True) Then
                    itemsToBeRemoved(itemsToBeRemoved.Length - 1) = child.DisplayName
                    ReDim Preserve itemsToBeRemoved(itemsToBeRemoved.Length)
                    GetAllSubComponentsForSelection(child, topName, True)
                ElseIf (topFound = False) Then
                    GetAllSubComponentsForSelection(child, topName, False)
                End If
            Next
        End Function

        Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click

            Dim MyFinalLst As String() = Nothing
            Dim MyRemoveLst As String() = Nothing

            If Me.lstRetainFiles.Items.Count <> 0 Then
                For i As Integer = 0 To Me.lstRetainFiles.Items.Count - 1
                    ReDim Preserve MyRemoveLst(i)
                    MyRemoveLst(i) = Me.lstRetainFiles.Items.Item(i)
                Next
            End If

            For i As Integer = 0 To Me.lstCloneFiles.Items.Count - 1

                Dim FoundSelectStatus As Boolean = False

                For Each mySelection As Integer In Me.lstCloneFiles.SelectedIndices
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

                    MyFinalLst(MyFinalLst.Length - 1) = Me.lstCloneFiles.Items.Item(i)
                Else
                    If MyRemoveLst Is Nothing Then
                        ReDim Preserve MyRemoveLst(0)
                    Else
                        ReDim Preserve MyRemoveLst(MyRemoveLst.Length)
                    End If

                    MyRemoveLst(MyRemoveLst.Length - 1) = Me.lstCloneFiles.Items.Item(i)
                End If

            Next

            Me.lstCloneFiles.Items.Clear()

            If Not MyFinalLst Is Nothing Then

                For i As Integer = 0 To MyFinalLst.Length - 1
                    Me.lstCloneFiles.Items.Add(MyFinalLst(i))
                Next

                Me.lstCloneFiles.Update()

            End If

            Me.lstRetainFiles.Items.Clear()

            If Not MyRemoveLst Is Nothing Then
                For j As Integer = 0 To MyRemoveLst.Length - 1
                    Me.lstRetainFiles.Items.Add(MyRemoveLst(j))
                Next

                Me.lstRetainFiles.Update()
            End If
        End Sub

        Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click

            Dim MyFinalLst As String() = Nothing
            Dim MyRemoveLst As String() = Nothing

            If Me.lstCloneFiles.Items.Count <> 0 Then
                For i As Integer = 0 To Me.lstCloneFiles.Items.Count - 1
                    ReDim Preserve MyFinalLst(i)
                    MyFinalLst(i) = Me.lstCloneFiles.Items.Item(i)
                Next
            End If

            For i As Integer = 0 To Me.lstRetainFiles.Items.Count - 1

                Dim FoundSelectStatus As Boolean = False

                For Each mySelection As Integer In Me.lstRetainFiles.SelectedIndices

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

                    MyRemoveLst(MyRemoveLst.Length - 1) = Me.lstRetainFiles.Items.Item(i)
                Else
                    If MyFinalLst Is Nothing Then
                        ReDim Preserve MyFinalLst(0)
                    Else
                        ReDim Preserve MyFinalLst(MyFinalLst.Length)
                    End If

                    MyFinalLst(MyFinalLst.Length - 1) = Me.lstRetainFiles.Items.Item(i)

                End If
            Next

            Me.lstCloneFiles.Items.Clear()

            If Not MyFinalLst Is Nothing Then

                For i As Integer = 0 To MyFinalLst.Length - 1
                    Me.lstCloneFiles.Items.Add(MyFinalLst(i))
                Next

                Me.lstCloneFiles.Update()

            End If

            Me.lstRetainFiles.Items.Clear()

            If Not MyRemoveLst Is Nothing Then

                For j As Integer = 0 To MyRemoveLst.Length - 1
                    Me.lstRetainFiles.Items.Add(MyRemoveLst(j))
                Next

                Me.lstRetainFiles.Update()

            End If
        End Sub

        Private Sub ChkBoxAltBaseStr_CheckedChanged(sender As Object, e As EventArgs) Handles ChkBoxAltBaseStr.CheckedChanged
            If ChkBoxAltBaseStr.Checked = True Then
                txtboxAltBaseStr1.Visible = True
                txtboxAltBaseStr2.Visible = True
                txtboxAltBaseStr3.Visible = True
            Else
                txtboxAltBaseStr1.Visible = False
                txtboxAltBaseStr2.Visible = False
                txtboxAltBaseStr3.Visible = False
            End If

        End Sub

        Function CheckReplaceString() As Boolean

            Dim Status As Boolean

            For Each myCloneFile As String In Me.lstCloneFiles.Items

                Status = False

                If myCloneFile.Contains(txtboxMasterBaseStr.Text) = True Then
                    Status = True
                Else
                    If ChkBoxAltBaseStr.Checked = False Then
                        MsgBox(myCloneFile & " cannot find the base string for replacement.", MsgBoxStyle.OkOnly, "Cannot Find Base String to Replace")
                        Status = False
                        Exit For
                    Else

                        If txtboxAltBaseStr1.Text <> "" Then
                            If myCloneFile.Contains(txtboxAltBaseStr1.Text) = True Then

                                Status = True
                                Continue For
                            End If
                        End If

                        If txtboxAltBaseStr2.Text <> "" Then
                            If myCloneFile.Contains(txtboxAltBaseStr2.Text) = True Then
                                Status = True
                                Continue For
                            End If
                        End If

                        If txtboxAltBaseStr3.Text <> "" Then
                            If myCloneFile.Contains(txtboxAltBaseStr3.Text) = True Then
                                Status = True
                                Continue For
                            End If
                        End If

                        If Status = False Then
                            MsgBox(myCloneFile & " cannot find the base string for replacement.", MsgBoxStyle.OkOnly, "Cannot Find Base String to Replace")
                            Exit For

                        End If
                    End If
                End If
            Next

            If Status = True Then
                Return True
            Else

                Return False
            End If
        End Function
    End Class

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class FormCloneAndRetain
        Inherits System.Windows.Forms.Form

        'Form overrides dispose to clean up the component list.
        <System.Diagnostics.DebuggerNonUserCode()>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Me.lstCloneFiles = New System.Windows.Forms.ListBox()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.btnOK = New System.Windows.Forms.Button()
            Me.btnCancel = New System.Windows.Forms.Button()
            Me.btnRemove = New System.Windows.Forms.Button()
            Me.btnReset = New System.Windows.Forms.Button()
            Me.lstRetainFiles = New System.Windows.Forms.ListBox()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.Label5 = New System.Windows.Forms.Label()
            Me.txtboxMasterBaseStr = New System.Windows.Forms.TextBox()
            Me.txtboxAltBaseStr1 = New System.Windows.Forms.TextBox()
            Me.txtboxAltBaseStr2 = New System.Windows.Forms.TextBox()
            Me.txtboxAltBaseStr3 = New System.Windows.Forms.TextBox()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.Label6 = New System.Windows.Forms.Label()
            Me.Label7 = New System.Windows.Forms.Label()
            Me.txtboxReplaceStr = New System.Windows.Forms.TextBox()
            Me.ChkBoxAltBaseStr = New System.Windows.Forms.CheckBox()
            Me.Label8 = New System.Windows.Forms.Label()
            Me.SuspendLayout()
            '
            'lstCloneFiles
            '
            Me.lstCloneFiles.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
            Me.lstCloneFiles.FormattingEnabled = True
            Me.lstCloneFiles.Location = New System.Drawing.Point(262, 46)
            Me.lstCloneFiles.Name = "lstCloneFiles"
            Me.lstCloneFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.lstCloneFiles.Size = New System.Drawing.Size(284, 134)
            Me.lstCloneFiles.Sorted = True
            Me.lstCloneFiles.TabIndex = 0
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Location = New System.Drawing.Point(259, 9)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(232, 13)
            Me.Label1.TabIndex = 1
            Me.Label1.Text = "Select the drawing(s) you dont want to print out:"
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Location = New System.Drawing.Point(259, 364)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(157, 13)
            Me.Label2.TabIndex = 2
            Me.Label2.Text = "P.S.: Multi Select by: Crtl + MB1"
            '
            'btnOK
            '
            Me.btnOK.Location = New System.Drawing.Point(262, 402)
            Me.btnOK.Name = "btnOK"
            Me.btnOK.Size = New System.Drawing.Size(75, 23)
            Me.btnOK.TabIndex = 3
            Me.btnOK.Text = "OK"
            Me.btnOK.UseVisualStyleBackColor = True
            '
            'btnCancel
            '
            Me.btnCancel.Location = New System.Drawing.Point(471, 402)
            Me.btnCancel.Name = "btnCancel"
            Me.btnCancel.Size = New System.Drawing.Size(75, 23)
            Me.btnCancel.TabIndex = 4
            Me.btnCancel.Text = "Cancel"
            Me.btnCancel.UseVisualStyleBackColor = True
            '
            'btnRemove
            '
            Me.btnRemove.Location = New System.Drawing.Point(567, 46)
            Me.btnRemove.Name = "btnRemove"
            Me.btnRemove.Size = New System.Drawing.Size(75, 23)
            Me.btnRemove.TabIndex = 5
            Me.btnRemove.Text = "Remove"
            Me.btnRemove.UseVisualStyleBackColor = True
            '
            'btnReset
            '
            Me.btnReset.Location = New System.Drawing.Point(567, 215)
            Me.btnReset.Name = "btnReset"
            Me.btnReset.Size = New System.Drawing.Size(75, 23)
            Me.btnReset.TabIndex = 6
            Me.btnReset.Text = "Add"
            Me.btnReset.UseVisualStyleBackColor = True
            '
            'lstRetainFiles
            '
            Me.lstRetainFiles.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
            Me.lstRetainFiles.FormattingEnabled = True
            Me.lstRetainFiles.Location = New System.Drawing.Point(262, 215)
            Me.lstRetainFiles.Name = "lstRetainFiles"
            Me.lstRetainFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.lstRetainFiles.Size = New System.Drawing.Size(284, 134)
            Me.lstRetainFiles.Sorted = True
            Me.lstRetainFiles.TabIndex = 8
            '
            'Label4
            '
            Me.Label4.AutoSize = True
            Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label4.Location = New System.Drawing.Point(267, 198)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(57, 13)
            Me.Label4.TabIndex = 9
            Me.Label4.Text = "Retain List"
            '
            'Label5
            '
            Me.Label5.AutoSize = True
            Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label5.Location = New System.Drawing.Point(262, 27)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(69, 13)
            Me.Label5.TabIndex = 10
            Me.Label5.Text = "To Clone List"
            '
            'txtboxMasterBaseStr
            '
            Me.txtboxMasterBaseStr.Location = New System.Drawing.Point(61, 65)
            Me.txtboxMasterBaseStr.Name = "txtboxMasterBaseStr"
            Me.txtboxMasterBaseStr.ReadOnly = True
            Me.txtboxMasterBaseStr.Size = New System.Drawing.Size(100, 20)
            Me.txtboxMasterBaseStr.TabIndex = 11
            '
            'txtboxAltBaseStr1
            '
            Me.txtboxAltBaseStr1.Location = New System.Drawing.Point(61, 139)
            Me.txtboxAltBaseStr1.Name = "txtboxAltBaseStr1"
            Me.txtboxAltBaseStr1.Size = New System.Drawing.Size(100, 20)
            Me.txtboxAltBaseStr1.TabIndex = 12
            '
            'txtboxAltBaseStr2
            '
            Me.txtboxAltBaseStr2.Location = New System.Drawing.Point(61, 165)
            Me.txtboxAltBaseStr2.Name = "txtboxAltBaseStr2"
            Me.txtboxAltBaseStr2.Size = New System.Drawing.Size(100, 20)
            Me.txtboxAltBaseStr2.TabIndex = 12
            '
            'txtboxAltBaseStr3
            '
            Me.txtboxAltBaseStr3.Location = New System.Drawing.Point(61, 191)
            Me.txtboxAltBaseStr3.Name = "txtboxAltBaseStr3"
            Me.txtboxAltBaseStr3.Size = New System.Drawing.Size(100, 20)
            Me.txtboxAltBaseStr3.TabIndex = 12
            '
            'Label3
            '
            Me.Label3.AutoSize = True
            Me.Label3.Location = New System.Drawing.Point(58, 46)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(96, 13)
            Me.Label3.TabIndex = 13
            Me.Label3.Text = "Master Base String"
            '
            'Label6
            '
            Me.Label6.AutoSize = True
            Me.Label6.Location = New System.Drawing.Point(58, 262)
            Me.Label6.Name = "Label6"
            Me.Label6.Size = New System.Drawing.Size(100, 13)
            Me.Label6.TabIndex = 14
            Me.Label6.Text = "Replacement String"
            '
            'Label7
            '
            Me.Label7.AutoSize = True
            Me.Label7.Location = New System.Drawing.Point(58, 113)
            Me.Label7.Name = "Label7"
            Me.Label7.Size = New System.Drawing.Size(106, 13)
            Me.Label7.TabIndex = 14
            Me.Label7.Text = "Alternate Base String"
            '
            'txtboxReplaceStr
            '
            Me.txtboxReplaceStr.Location = New System.Drawing.Point(61, 287)
            Me.txtboxReplaceStr.Name = "txtboxReplaceStr"
            Me.txtboxReplaceStr.ReadOnly = True
            Me.txtboxReplaceStr.Size = New System.Drawing.Size(100, 20)
            Me.txtboxReplaceStr.TabIndex = 12
            '
            'ChkBoxAltBaseStr
            '
            Me.ChkBoxAltBaseStr.AutoSize = True
            Me.ChkBoxAltBaseStr.Location = New System.Drawing.Point(30, 113)
            Me.ChkBoxAltBaseStr.Name = "ChkBoxAltBaseStr"
            Me.ChkBoxAltBaseStr.Size = New System.Drawing.Size(15, 14)
            Me.ChkBoxAltBaseStr.TabIndex = 15
            Me.ChkBoxAltBaseStr.UseVisualStyleBackColor = True
            '
            'Label8
            '
            Me.Label8.AutoSize = True
            Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label8.Location = New System.Drawing.Point(40, 9)
            Me.Label8.Name = "Label8"
            Me.Label8.Size = New System.Drawing.Size(136, 20)
            Me.Label8.TabIndex = 16
            Me.Label8.Text = "Find And Replace"
            '
            'FormCloneAndRetain
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(660, 446)
            Me.Controls.Add(Me.Label8)
            Me.Controls.Add(Me.ChkBoxAltBaseStr)
            Me.Controls.Add(Me.Label7)
            Me.Controls.Add(Me.Label6)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me.txtboxAltBaseStr3)
            Me.Controls.Add(Me.txtboxAltBaseStr2)
            Me.Controls.Add(Me.txtboxAltBaseStr1)
            Me.Controls.Add(Me.txtboxReplaceStr)
            Me.Controls.Add(Me.txtboxMasterBaseStr)
            Me.Controls.Add(Me.Label5)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.lstRetainFiles)
            Me.Controls.Add(Me.btnReset)
            Me.Controls.Add(Me.btnRemove)
            Me.Controls.Add(Me.btnCancel)
            Me.Controls.Add(Me.btnOK)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.lstCloneFiles)
            Me.Name = "FormCloneAndRetain"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "FormCloneAndRetain"
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents lstCloneFiles As System.Windows.Forms.ListBox
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents Label2 As System.Windows.Forms.Label
        Friend WithEvents btnOK As System.Windows.Forms.Button
        Friend WithEvents btnCancel As System.Windows.Forms.Button
        Friend WithEvents btnRemove As System.Windows.Forms.Button
        Friend WithEvents btnReset As System.Windows.Forms.Button
        Friend WithEvents lstRetainFiles As System.Windows.Forms.ListBox
        Friend WithEvents Label4 As System.Windows.Forms.Label
        Friend WithEvents Label5 As System.Windows.Forms.Label
        Friend WithEvents txtboxMasterBaseStr As System.Windows.Forms.TextBox
        Friend WithEvents txtboxAltBaseStr1 As System.Windows.Forms.TextBox
        Friend WithEvents txtboxAltBaseStr2 As System.Windows.Forms.TextBox
        Friend WithEvents txtboxAltBaseStr3 As System.Windows.Forms.TextBox
        Friend WithEvents Label3 As System.Windows.Forms.Label
        Friend WithEvents Label6 As System.Windows.Forms.Label
        Friend WithEvents Label7 As System.Windows.Forms.Label
        Friend WithEvents txtboxReplaceStr As System.Windows.Forms.TextBox
        Friend WithEvents ChkBoxAltBaseStr As System.Windows.Forms.CheckBox
        Friend WithEvents Label8 As System.Windows.Forms.Label
    End Class

End Module