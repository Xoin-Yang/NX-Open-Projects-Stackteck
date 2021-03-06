'---------------Created by Raem Naveed (2017)---------------------'

Option Strict Off
Imports System
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.IO
Imports System.Drawing
Imports System.ComponentModel
Imports System.Collections
Imports NXOpenUI
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Utilities
Imports NXOpen.Assemblies

Module NXJournal
    Dim siUnits As String = Nothing
    Dim imperialUnits As String = Nothing
    Dim siUnits2 As String = Nothing
    Dim imperialUnits2 As String = Nothing
    Dim searchTerm As String = Nothing
    Dim searchString As String() = Nothing
    Dim attributeString As String() = Nothing
    Dim attributeTerm As String = Nothing

    Sub Main(ByVal args() As String)
        Dim form As Form1
        form = New Form1()
        form.ShowDialog()
    End Sub
    Public Class Form1
        Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            searchString = {"TONNAGE", "RING DIA", "(HORIZONTAL)", "(VERTICAL)", "(MIN)", "(MAX)", "Ejection Stroke", "DAYLIGHT"}
            txtBoxTon.Text = searchString(0)
            txtBoxRin.Text = searchString(1)
            txtBoxTieH.Text = searchString(2)
            txtBoxTieV.Text = searchString(3)
            txtBoxShuMin.Text = searchString(4)
            txtBoxShuMax.Text = searchString(5)
            txtBoxEje.Text = searchString(6)
            txtBoxDay.Text = searchString(7)

        End Sub

        Private Sub btnFileSearch_Click(sender As Object, e As EventArgs) Handles btnFileSearch.Click
            If (OpenFileDialog1.ShowDialog = DialogResult.OK) Then
                txtBoxFileName.Text = OpenFileDialog1.FileName
            End If
        End Sub

        Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
            searchString(0) = txtBoxTon.Text
            searchString(1) = txtBoxRin.Text
            searchString(2) = txtBoxTieH.Text
            searchString(3) = txtBoxTieV.Text
            searchString(4) = txtBoxShuMin.Text
            searchString(5) = txtBoxShuMax.Text
            searchString(6) = txtBoxEje.Text
            searchString(7) = txtBoxDay.Text
            attributeString = {"TONNAGE", "LOCATING_RING_DIAMETER", "TIEBAR_SPACEING_H", "TIEBAR_SPACING_V", "SHUT_HEIGHT_MIN", "SHUT_HEIGHT_MAX", "EJECTION_STROKE", "DAYLIGHT"}
            Dim i As Integer = 0
            For Each term As String In attributeString
                searchTerm = searchString(i)
                attributeTerm = term
                ''Special Case for TieBar
                If i = 1 Then
                    i += 1
                    Continue For
                End If
                If chkBoxTieBar.Checked = True And i = 2 Then
                    readExcelMod()
                    setAtributeMet()
                    setAtributeImp()
                    siUnits = siUnits2
                    imperialUnits = imperialUnits2
                    i += 1
                    Continue For
                ElseIf chkBoxTieBar.Checked = True And i = 3 Then
                    setAtributeMet()
                    setAtributeImp()
                    i += 1
                    Continue For
                End If
                ''Special Case for ShutHeight
                If chkBoxShut.Checked = True And i = 4 Then
                    readExcelMod()
                    setAtributeMet()
                    setAtributeImp()
                    siUnits = siUnits2
                    imperialUnits = imperialUnits2
                    i += 1
                    Continue For
                ElseIf chkBoxShut.Checked = True And i = 5 Then
                    setAtributeMet()
                    setAtributeImp()
                    i += 1
                    Continue For
                End If

                Try
                    readExcel()
                Catch
                    MsgBox("Error looking for " & searchTerm)
                End Try
                Try
                    setAtributeMet()
                    setAtributeImp()
                Catch ex As Exception
                    MsgBox("error creating attribute for " & attributeTerm)
                End Try
                i += 1
            Next
        End Sub

        Private Sub chkBoxTieBar_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxTieBar.CheckedChanged
            If chkBoxTieBar.Checked = True Then
                txtBoxTieV.Text = ""
                txtBoxTieH.Text = "TieBar Spacing"
                txtBoxTieV.ReadOnly = True
                Label4.Text = ""
                Label3.Text = "TieBar Spacing"
            Else
                txtBoxTieH.Text = searchString(2)
                txtBoxTieV.Text = searchString(3)
                txtBoxTieV.ReadOnly = False
                Label4.Text = "TieBar Spacing (V)"
                Label3.Text = "TieBar Spacing (H)"
            End If
        End Sub
        Private Sub chkBoxShut_CheckedChanged(sender As Object, e As EventArgs) Handles chkBoxShut.CheckedChanged
            If chkBoxShut.Checked = True Then
                txtBoxShuMax.Text = ""
                txtBoxShuMin.Text = "ShutHeight"
                txtBoxShuMax.ReadOnly = True
                Label6.Text = ""
                Label5.Text = "TieBar Spacing"
            Else
                txtBoxShuMin.Text = searchString(4)
                txtBoxShuMax.Text = searchString(5)
                txtBoxShuMax.ReadOnly = False
                Label6.Text = "TieBar Spacing (V)"
                Label5.Text = "TieBar Spacing (H)"
            End If
        End Sub

        Public Sub readExcel()

            Dim UserID As String
            UserID = Environment.UserName()

            'Define the connectors
            Dim cn As OleDbConnection
            Dim cmd As OleDbCommand
            Dim dr As OleDbDataReader
            Dim oConnect, oQuery As String
            Dim FoundStatus As Boolean = False
            Dim status As Boolean = False

            'Define connection string
            Dim FileName As String = txtBoxFileName.Text
            If File.Exists(FileName) = False Then
                MessageBox.Show("File " & FileName & " is not found.")
                status = False
                Exit Sub
            End If

            'oConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName
            oConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"
            'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=EXCLE_FILE_PATH;Extended Properties=Excel 12.0;HDR=Yes"""

            'Query String
            'oQuery = "SELECT * FROM Title_Blk_User_Info where UserID='" & UserID & "'"   '"SELECT * FROM [Sheet1$]"
            oQuery = "SELECT * FROM [Sheet1$] where ID LIKE'%" & searchTerm & "%'"
            'Instantiate the connectors
            cn = New OleDbConnection(oConnect)

            Try
                cn.Open()
            Catch ex As Exception
            Finally
                cn = New OleDbConnection(oConnect)
                cn.Open()
            End Try

            cmd = New OleDbCommand(oQuery, cn)
            dr = cmd.ExecuteReader

            While dr.Read()
                If chkBoxSI.Checked = True Then
                    siUnits = dr("SI")
                    siUnits = siUnits.Trim()
                    'MessageBox.Show("UserName is: " & UserName)
                End If
                If chkBoxImp.Checked = True Then
                    imperialUnits = dr("IMP")
                    imperialUnits = imperialUnits.Trim()
                    'MessageBox.Show("UserGroup is: " & UserGroup)

                    FoundStatus = True
                    status = True
                End If

            End While

            dr.Close()
            cn.Close()
            'MsgBox(siUnits)
            'MsgBox(imperialUnits)
            If FoundStatus = False Then
                MessageBox.Show("Search Term: " & searchTerm & " is not found from Excel Sheet")
            End If
        End Sub
        Public Sub readExcelMod()

            'Define the connectors
            Dim cn As OleDbConnection
            Dim cmd As OleDbCommand
            Dim dr As OleDbDataReader
            Dim oConnect, oQuery As String
            Dim FoundStatus As Boolean = False
            Dim status As Boolean = False

            'Define connection string
            Dim FileName As String = txtBoxFileName.Text
            If File.Exists(FileName) = False Then
                MessageBox.Show("File " & FileName & " is not found.")
                status = False
                Exit Sub
            End If

            'oConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName
            oConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"
            'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=EXCLE_FILE_PATH;Extended Properties=Excel 12.0;HDR=Yes"""

            'Query String
            'oQuery = "SELECT * FROM Title_Blk_User_Info where UserID='" & UserID & "'"   '"SELECT * FROM [Sheet1$]"
            oQuery = "SELECT * FROM [Sheet1$] where ID LIKE'%" & searchTerm & "%'"
            'Instantiate the connectors
            cn = New OleDbConnection(oConnect)

            Try
                cn.Open()
            Catch ex As Exception
            Finally
                cn = New OleDbConnection(oConnect)
                cn.Open()
            End Try

            cmd = New OleDbCommand(oQuery, cn)
            dr = cmd.ExecuteReader

            While dr.Read()
                If chkBoxSI.Checked = True Then
                    siUnits = dr("SI")
                    siUnits = siUnits.Replace(" ", "")
                    siUnits = siUnits.Trim()
                End If
                'MessageBox.Show("UserName is: " & UserName)
                If chkBoxImp.Checked = True Then
                    imperialUnits = dr("IMP")
                    imperialUnits = imperialUnits.Replace(" ", "")
                    imperialUnits = imperialUnits.Trim()
                End If
                'MessageBox.Show("UserGroup is: " & UserGroup)

                FoundStatus = True
                status = True
            End While

            Dim position As Integer = 0
            If chkBoxSI.Checked = True Then
                If siUnits.Contains("/") Then
                    position = siUnits.IndexOf("/")
                    siUnits2 = siUnits.Substring(position + 1, siUnits.Length - position - 1)
                    siUnits = siUnits.Substring(0, position)
                ElseIf siUnits.Contains("x") Then
                    'MsgBox(position)
                    position = siUnits.IndexOf("x")
                    siUnits2 = siUnits.Substring(position + 1, siUnits.Length - position - 1)
                    siUnits.Trim()
                    siUnits = siUnits.Substring(0, position)
                ElseIf siUnits.Contains("-") Then
                    position = siUnits.IndexOf("-")
                    siUnits2 = siUnits.Substring(position + 1, siUnits.Length - position - 1)
                    siUnits = siUnits.Substring(0, position)
                Else
                    MsgBox("couldnt find multiple data entires")
                End If
            End If
            position = 0
            If chkBoxImp.Checked = True Then
                If imperialUnits.Contains("/") Then
                    position = imperialUnits.IndexOf("/")
                    imperialUnits2 = imperialUnits.Substring(position + 1, imperialUnits.Length - position - 1)
                    imperialUnits = imperialUnits.Substring(0, position)
                ElseIf imperialUnits.Contains("x") Then
                    position = imperialUnits.IndexOf("x")
                    imperialUnits2 = imperialUnits.Substring(position + 1, imperialUnits.Length - position - 1)
                    imperialUnits = imperialUnits.Substring(0, position)
                ElseIf imperialUnits.Contains("-") Then
                    position = imperialUnits.IndexOf("-")
                    imperialUnits2 = imperialUnits.Substring(position + 1, imperialUnits.Length - position - 1)
                    imperialUnits = imperialUnits.Substring(0, position)
                Else
                    MsgBox("couldnt find multiple data entires")
                End If
            End If

            dr.Close()
            cn.Close()
            'MsgBox(siUnits)
            'MsgBox(imperialUnits)
            If FoundStatus = False Then
                MessageBox.Show("Search Term: " & searchTerm & " is not found from Excel Sheet")
            End If
        End Sub

        Public Sub setAtributeMet()
            If chkBoxSI.Checked = False Then
                Exit Sub
            End If
            attributeTerm = attributeTerm.Replace(" ", "_")
            Dim theSession As Session = Session.GetSession()
            Dim workPart As Part = theSession.Parts.Work
            Dim displayPart As Part = theSession.Parts.Display
            Dim objects1(0) As NXObject
            objects1(0) = workPart
            Dim attributePropertiesBuilder1 As AttributePropertiesBuilder
            attributePropertiesBuilder1 = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, objects1, AttributePropertiesBuilder.OperationType.None)
            attributePropertiesBuilder1.IsArray = False
            attributePropertiesBuilder1.IsArray = False
            attributePropertiesBuilder1.IsArray = False
            attributePropertiesBuilder1.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.String
            attributePropertiesBuilder1.Units = "Inch"
            Dim objects2(0) As NXObject
            objects2(0) = workPart
            Dim massPropertiesBuilder1 As MassPropertiesBuilder
            massPropertiesBuilder1 = workPart.PropertiesManager.CreateMassPropertiesBuilder(objects2)
            Dim selectNXObjectList1 As SelectNXObjectList
            selectNXObjectList1 = massPropertiesBuilder1.SelectedObjects
            Dim objects3() As NXObject
            objects3 = selectNXObjectList1.GetArray()
            massPropertiesBuilder1.LoadPartialComponents = True
            massPropertiesBuilder1.Accuracy = 0.99
            Dim objects4(0) As NXObject
            objects4(0) = workPart
            Dim previewPropertiesBuilder1 As PreviewPropertiesBuilder
            previewPropertiesBuilder1 = workPart.PropertiesManager.CreatePreviewPropertiesBuilder(objects4)
            previewPropertiesBuilder1.StorePartPreview = True
            previewPropertiesBuilder1.StoreModelViewPreview = True
            previewPropertiesBuilder1.ModelViewCreation = PreviewPropertiesBuilder.ModelViewCreationOptions.OnViewSave
            Dim objects5(0) As NXObject
            objects5(0) = workPart
            attributePropertiesBuilder1.SetAttributeObjects(objects5)
            attributePropertiesBuilder1.Units = "Inch"
            attributePropertiesBuilder1.DateValue.DateItem.Day = DateItemBuilder.DayOfMonth.Day27
            attributePropertiesBuilder1.DateValue.DateItem.Month = DateItemBuilder.MonthOfYear.Sep
            attributePropertiesBuilder1.DateValue.DateItem.Year = "2017"
            attributePropertiesBuilder1.DateValue.DateItem.Time = "00:00:00"
            massPropertiesBuilder1.UpdateOnSave = MassPropertiesBuilder.UpdateOptions.No
            ' ----------------------------------------------
            '   Dialog Begin Displayed Part Properties
            ' ----------------------------------------------
            attributePropertiesBuilder1.Title = ""
            attributePropertiesBuilder1.IsArray = False
            attributePropertiesBuilder1.StringValue = ""
            attributePropertiesBuilder1.Title = "STACKTECK_" & attributeTerm & "_met"
            attributePropertiesBuilder1.StringValue = siUnits
            Dim nXObject1 As NXObject
            nXObject1 = attributePropertiesBuilder1.Commit()
            Dim updateoption1 As MassPropertiesBuilder.UpdateOptions
            updateoption1 = massPropertiesBuilder1.UpdateOnSave
            Dim nXObject2 As NXObject
            nXObject2 = massPropertiesBuilder1.Commit()
            workPart.PartPreviewMode = BasePart.PartPreview.OnSave
            Dim nXObject3 As NXObject
            nXObject3 = previewPropertiesBuilder1.Commit()
            'id1 = theSession.GetNewestUndoMark(Session.MarkVisibility.Visible)
            Dim nErrs1 As Integer
            'nErrs1 = theSession.UpdateManager.DoUpdate(id1)
            'theSession.SetUndoMarkName(id1, "Displayed Part Properties")
            attributePropertiesBuilder1.Destroy()
            massPropertiesBuilder1.Destroy()
            previewPropertiesBuilder1.Destroy()
            ' ----------------------------------------------
            '   Menu: Tools->Journal->Stop Recording
            ' ----------------------------------------------
        End Sub

        Public Sub setAtributeImp()
            If chkBoxImp.Checked = False Then
                Exit Sub
            End If
            attributeTerm = attributeTerm.Replace(" ", "_")
            Dim theSession As Session = Session.GetSession()
            Dim workPart As Part = theSession.Parts.Work
            Dim displayPart As Part = theSession.Parts.Display
            Dim objects1(0) As NXObject
            objects1(0) = workPart
            Dim attributePropertiesBuilder1 As AttributePropertiesBuilder
            attributePropertiesBuilder1 = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, objects1, AttributePropertiesBuilder.OperationType.None)
            attributePropertiesBuilder1.IsArray = False
            attributePropertiesBuilder1.IsArray = False
            attributePropertiesBuilder1.IsArray = False
            attributePropertiesBuilder1.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.String
            attributePropertiesBuilder1.Units = "Inch"
            Dim objects2(0) As NXObject
            objects2(0) = workPart
            Dim massPropertiesBuilder1 As MassPropertiesBuilder
            massPropertiesBuilder1 = workPart.PropertiesManager.CreateMassPropertiesBuilder(objects2)
            Dim selectNXObjectList1 As SelectNXObjectList
            selectNXObjectList1 = massPropertiesBuilder1.SelectedObjects
            Dim objects3() As NXObject
            objects3 = selectNXObjectList1.GetArray()
            massPropertiesBuilder1.LoadPartialComponents = True
            massPropertiesBuilder1.Accuracy = 0.99
            Dim objects4(0) As NXObject
            objects4(0) = workPart
            Dim previewPropertiesBuilder1 As PreviewPropertiesBuilder
            previewPropertiesBuilder1 = workPart.PropertiesManager.CreatePreviewPropertiesBuilder(objects4)
            previewPropertiesBuilder1.StorePartPreview = True
            previewPropertiesBuilder1.StoreModelViewPreview = True
            previewPropertiesBuilder1.ModelViewCreation = PreviewPropertiesBuilder.ModelViewCreationOptions.OnViewSave
            Dim objects5(0) As NXObject
            objects5(0) = workPart
            attributePropertiesBuilder1.SetAttributeObjects(objects5)
            attributePropertiesBuilder1.Units = "Inch"
            attributePropertiesBuilder1.DateValue.DateItem.Day = DateItemBuilder.DayOfMonth.Day27
            attributePropertiesBuilder1.DateValue.DateItem.Month = DateItemBuilder.MonthOfYear.Sep
            attributePropertiesBuilder1.DateValue.DateItem.Year = "2017"
            attributePropertiesBuilder1.DateValue.DateItem.Time = "00:00:00"
            massPropertiesBuilder1.UpdateOnSave = MassPropertiesBuilder.UpdateOptions.No
            ' ----------------------------------------------
            '   Dialog Begin Displayed Part Properties
            ' ----------------------------------------------
            attributePropertiesBuilder1.Title = ""
            attributePropertiesBuilder1.IsArray = False
            attributePropertiesBuilder1.StringValue = ""
            attributePropertiesBuilder1.Title = "STACKTECK_" & attributeTerm & "_imp"
            attributePropertiesBuilder1.StringValue = imperialUnits
            Dim nXObject1 As NXObject
            nXObject1 = attributePropertiesBuilder1.Commit()
            Dim updateoption1 As MassPropertiesBuilder.UpdateOptions
            updateoption1 = massPropertiesBuilder1.UpdateOnSave
            Dim nXObject2 As NXObject
            nXObject2 = massPropertiesBuilder1.Commit()
            workPart.PartPreviewMode = BasePart.PartPreview.OnSave
            Dim nXObject3 As NXObject
            nXObject3 = previewPropertiesBuilder1.Commit()
            'id1 = theSession.GetNewestUndoMark(Session.MarkVisibility.Visible)
            Dim nErrs1 As Integer
            'nErrs1 = theSession.UpdateManager.DoUpdate(id1)
            'theSession.SetUndoMarkName(id1, "Displayed Part Properties")
            attributePropertiesBuilder1.Destroy()
            massPropertiesBuilder1.Destroy()
            previewPropertiesBuilder1.Destroy()
            ' ----------------------------------------------
            '   Menu: Tools->Journal->Stop Recording
            ' ----------------------------------------------
        End Sub

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
            Me.txtBoxFileName = New System.Windows.Forms.TextBox()
            Me.btnFileSearch = New System.Windows.Forms.Button()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.txtBoxTon = New System.Windows.Forms.TextBox()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.Label5 = New System.Windows.Forms.Label()
            Me.Label6 = New System.Windows.Forms.Label()
            Me.Label7 = New System.Windows.Forms.Label()
            Me.Label8 = New System.Windows.Forms.Label()
            Me.txtBoxRin = New System.Windows.Forms.TextBox()
            Me.txtBoxTieH = New System.Windows.Forms.TextBox()
            Me.txtBoxTieV = New System.Windows.Forms.TextBox()
            Me.txtBoxShuMin = New System.Windows.Forms.TextBox()
            Me.txtBoxShuMax = New System.Windows.Forms.TextBox()
            Me.txtBoxDay = New System.Windows.Forms.TextBox()
            Me.txtBoxEje = New System.Windows.Forms.TextBox()
            Me.btnOk = New System.Windows.Forms.Button()
            Me.btnCancel = New System.Windows.Forms.Button()
            Me.Label9 = New System.Windows.Forms.Label()
            Me.chkBoxTieBar = New System.Windows.Forms.CheckBox()
            Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
            Me.chkBoxShut = New System.Windows.Forms.CheckBox()
            Me.chkBoxSI = New System.Windows.Forms.CheckBox()
            Me.chkBoxImp = New System.Windows.Forms.CheckBox()
            Me.SuspendLayout()
            '
            'txtBoxFileName
            '
            Me.txtBoxFileName.Location = New System.Drawing.Point(13, 35)
            Me.txtBoxFileName.Name = "txtBoxFileName"
            Me.txtBoxFileName.Size = New System.Drawing.Size(229, 20)
            Me.txtBoxFileName.TabIndex = 0
            '
            'btnFileSearch
            '
            Me.btnFileSearch.BackColor = System.Drawing.Color.Gold
            Me.btnFileSearch.Location = New System.Drawing.Point(248, 33)
            Me.btnFileSearch.Name = "btnFileSearch"
            Me.btnFileSearch.Size = New System.Drawing.Size(75, 23)
            Me.btnFileSearch.TabIndex = 5
            Me.btnFileSearch.Text = "Search"
            Me.btnFileSearch.UseVisualStyleBackColor = False
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Location = New System.Drawing.Point(11, 121)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(50, 13)
            Me.Label1.TabIndex = 2
            Me.Label1.Text = "Tonnage"
            '
            'txtBoxTon
            '
            Me.txtBoxTon.Location = New System.Drawing.Point(117, 114)
            Me.txtBoxTon.Name = "txtBoxTon"
            Me.txtBoxTon.Size = New System.Drawing.Size(206, 20)
            Me.txtBoxTon.TabIndex = 20
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Location = New System.Drawing.Point(11, 154)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(92, 13)
            Me.Label2.TabIndex = 4
            Me.Label2.Text = "Locating Ring Dia"
            '
            'Label3
            '
            Me.Label3.AutoSize = True
            Me.Label3.Location = New System.Drawing.Point(11, 185)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(97, 13)
            Me.Label3.TabIndex = 5
            Me.Label3.Text = "TieBar Spacing (H)"
            '
            'Label4
            '
            Me.Label4.AutoSize = True
            Me.Label4.Location = New System.Drawing.Point(11, 216)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(96, 13)
            Me.Label4.TabIndex = 6
            Me.Label4.Text = "TieBar Spacing (V)"
            '
            'Label5
            '
            Me.Label5.AutoSize = True
            Me.Label5.Location = New System.Drawing.Point(11, 248)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(89, 13)
            Me.Label5.TabIndex = 7
            Me.Label5.Text = "Shut Height (Min)"
            '
            'Label6
            '
            Me.Label6.AutoSize = True
            Me.Label6.Location = New System.Drawing.Point(11, 282)
            Me.Label6.Name = "Label6"
            Me.Label6.Size = New System.Drawing.Size(92, 13)
            Me.Label6.TabIndex = 8
            Me.Label6.Text = "Shut Height (Max)"
            '
            'Label7
            '
            Me.Label7.AutoSize = True
            Me.Label7.Location = New System.Drawing.Point(11, 317)
            Me.Label7.Name = "Label7"
            Me.Label7.Size = New System.Drawing.Size(68, 13)
            Me.Label7.TabIndex = 9
            Me.Label7.Text = "Max Daylight"
            '
            'Label8
            '
            Me.Label8.AutoSize = True
            Me.Label8.Location = New System.Drawing.Point(11, 348)
            Me.Label8.Name = "Label8"
            Me.Label8.Size = New System.Drawing.Size(79, 13)
            Me.Label8.TabIndex = 10
            Me.Label8.Text = "Ejection Stroke"
            '
            'txtBoxRin
            '
            Me.txtBoxRin.Location = New System.Drawing.Point(117, 147)
            Me.txtBoxRin.Name = "txtBoxRin"
            Me.txtBoxRin.Size = New System.Drawing.Size(206, 20)
            Me.txtBoxRin.TabIndex = 25
            '
            'txtBoxTieH
            '
            Me.txtBoxTieH.Location = New System.Drawing.Point(117, 178)
            Me.txtBoxTieH.Name = "txtBoxTieH"
            Me.txtBoxTieH.Size = New System.Drawing.Size(206, 20)
            Me.txtBoxTieH.TabIndex = 30
            '
            'txtBoxTieV
            '
            Me.txtBoxTieV.Location = New System.Drawing.Point(117, 209)
            Me.txtBoxTieV.Name = "txtBoxTieV"
            Me.txtBoxTieV.Size = New System.Drawing.Size(206, 20)
            Me.txtBoxTieV.TabIndex = 35
            '
            'txtBoxShuMin
            '
            Me.txtBoxShuMin.Location = New System.Drawing.Point(117, 241)
            Me.txtBoxShuMin.Name = "txtBoxShuMin"
            Me.txtBoxShuMin.Size = New System.Drawing.Size(206, 20)
            Me.txtBoxShuMin.TabIndex = 40
            '
            'txtBoxShuMax
            '
            Me.txtBoxShuMax.Location = New System.Drawing.Point(117, 275)
            Me.txtBoxShuMax.Name = "txtBoxShuMax"
            Me.txtBoxShuMax.Size = New System.Drawing.Size(206, 20)
            Me.txtBoxShuMax.TabIndex = 45
            '
            'txtBoxDay
            '
            Me.txtBoxDay.Location = New System.Drawing.Point(117, 310)
            Me.txtBoxDay.Name = "txtBoxDay"
            Me.txtBoxDay.Size = New System.Drawing.Size(206, 20)
            Me.txtBoxDay.TabIndex = 50
            '
            'txtBoxEje
            '
            Me.txtBoxEje.Location = New System.Drawing.Point(117, 345)
            Me.txtBoxEje.Name = "txtBoxEje"
            Me.txtBoxEje.Size = New System.Drawing.Size(206, 20)
            Me.txtBoxEje.TabIndex = 55
            '
            'btnOk
            '
            Me.btnOk.BackColor = System.Drawing.Color.Green
            Me.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.btnOk.ForeColor = System.Drawing.Color.Gold
            Me.btnOk.Location = New System.Drawing.Point(64, 381)
            Me.btnOk.Name = "btnOk"
            Me.btnOk.Size = New System.Drawing.Size(92, 36)
            Me.btnOk.TabIndex = 18
            Me.btnOk.Text = "Run"
            Me.btnOk.UseVisualStyleBackColor = False
            '
            'btnCancel
            '
            Me.btnCancel.BackColor = System.Drawing.Color.Firebrick
            Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnCancel.ForeColor = System.Drawing.Color.Gold
            Me.btnCancel.Location = New System.Drawing.Point(194, 381)
            Me.btnCancel.Name = "btnCancel"
            Me.btnCancel.Size = New System.Drawing.Size(92, 36)
            Me.btnCancel.TabIndex = 19
            Me.btnCancel.Text = "Cancel"
            Me.btnCancel.UseVisualStyleBackColor = False
            '
            'Label9
            '
            Me.Label9.AutoSize = True
            Me.Label9.Location = New System.Drawing.Point(13, 19)
            Me.Label9.Name = "Label9"
            Me.Label9.Size = New System.Drawing.Size(96, 13)
            Me.Label9.TabIndex = 20
            Me.Label9.Text = "Excel File Location"
            '
            'chkBoxTieBar
            '
            Me.chkBoxTieBar.AutoSize = True
            Me.chkBoxTieBar.Location = New System.Drawing.Point(13, 62)
            Me.chkBoxTieBar.Name = "chkBoxTieBar"
            Me.chkBoxTieBar.Size = New System.Drawing.Size(203, 17)
            Me.chkBoxTieBar.TabIndex = 10
            Me.chkBoxTieBar.Text = "Is TieBar Spacing in one row? (H x V)"
            Me.chkBoxTieBar.UseVisualStyleBackColor = True
            '
            'OpenFileDialog1
            '
            Me.OpenFileDialog1.FileName = "OpenFileDialog1"
            '
            'chkBoxShut
            '
            Me.chkBoxShut.AutoSize = True
            Me.chkBoxShut.Location = New System.Drawing.Point(13, 85)
            Me.chkBoxShut.Name = "chkBoxShut"
            Me.chkBoxShut.Size = New System.Drawing.Size(245, 17)
            Me.chkBoxShut.TabIndex = 15
            Me.chkBoxShut.Text = "Is ShutHeight Spacing in one row? (Min - Max)"
            Me.chkBoxShut.UseVisualStyleBackColor = True
            '
            'chkBoxSI
            '
            Me.chkBoxSI.AutoSize = True
            Me.chkBoxSI.Checked = True
            Me.chkBoxSI.CheckState = System.Windows.Forms.CheckState.Checked
            Me.chkBoxSI.Location = New System.Drawing.Point(264, 62)
            Me.chkBoxSI.Name = "chkBoxSI"
            Me.chkBoxSI.Size = New System.Drawing.Size(36, 17)
            Me.chkBoxSI.TabIndex = 16
            Me.chkBoxSI.Text = "SI"
            Me.chkBoxSI.UseVisualStyleBackColor = True
            '
            'chkBoxImp
            '
            Me.chkBoxImp.AutoSize = True
            Me.chkBoxImp.Checked = True
            Me.chkBoxImp.CheckState = System.Windows.Forms.CheckState.Checked
            Me.chkBoxImp.Location = New System.Drawing.Point(264, 85)
            Me.chkBoxImp.Name = "chkBoxImp"
            Me.chkBoxImp.Size = New System.Drawing.Size(43, 17)
            Me.chkBoxImp.TabIndex = 17
            Me.chkBoxImp.Text = "Imp"
            Me.chkBoxImp.UseVisualStyleBackColor = True
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(335, 429)
            Me.Controls.Add(Me.chkBoxImp)
            Me.Controls.Add(Me.chkBoxSI)
            Me.Controls.Add(Me.chkBoxShut)
            Me.Controls.Add(Me.chkBoxTieBar)
            Me.Controls.Add(Me.Label9)
            Me.Controls.Add(Me.btnCancel)
            Me.Controls.Add(Me.btnOk)
            Me.Controls.Add(Me.txtBoxEje)
            Me.Controls.Add(Me.txtBoxDay)
            Me.Controls.Add(Me.txtBoxShuMax)
            Me.Controls.Add(Me.txtBoxShuMin)
            Me.Controls.Add(Me.txtBoxTieV)
            Me.Controls.Add(Me.txtBoxTieH)
            Me.Controls.Add(Me.txtBoxRin)
            Me.Controls.Add(Me.Label8)
            Me.Controls.Add(Me.Label7)
            Me.Controls.Add(Me.Label6)
            Me.Controls.Add(Me.Label5)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.txtBoxTon)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.btnFileSearch)
            Me.Controls.Add(Me.txtBoxFileName)
            Me.Name = "Form1"
            Me.Text = "Machine Data Input"
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Friend WithEvents txtBoxFileName As TextBox
        Friend WithEvents btnFileSearch As Button
        Friend WithEvents Label1 As Label
        Friend WithEvents txtBoxTon As TextBox
        Friend WithEvents Label2 As Label
        Friend WithEvents Label3 As Label
        Friend WithEvents Label4 As Label
        Friend WithEvents Label5 As Label
        Friend WithEvents Label6 As Label
        Friend WithEvents Label7 As Label
        Friend WithEvents Label8 As Label
        Friend WithEvents txtBoxRin As TextBox
        Friend WithEvents txtBoxTieH As TextBox
        Friend WithEvents txtBoxTieV As TextBox
        Friend WithEvents txtBoxShuMin As TextBox
        Friend WithEvents txtBoxShuMax As TextBox
        Friend WithEvents txtBoxDay As TextBox
        Friend WithEvents txtBoxEje As TextBox
        Friend WithEvents btnOk As Button
        Friend WithEvents btnCancel As Button
        Friend WithEvents Label9 As Label
        Friend WithEvents chkBoxTieBar As CheckBox
        Friend WithEvents OpenFileDialog1 As OpenFileDialog
        Friend WithEvents chkBoxShut As CheckBox
        Friend WithEvents chkBoxSI As CheckBox
        Friend WithEvents chkBoxImp As CheckBox
    End Class



End Module
