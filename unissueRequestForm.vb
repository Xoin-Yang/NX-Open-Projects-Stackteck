Option Strict Off
Imports System
Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Data.Sql
Imports System.Data
Imports System.String
Imports System.Text
Imports System.Reflection
Imports NXOpen
Imports NXOpen.UF
Imports NXOpenUI
Imports NXOpen.UI
Imports NXOpen.Utilities
Imports NXOpen.Assemblies
Imports NXOpen.Layer
Imports NXOpen.Drawings
Imports System.Text.RegularExpressions
Imports System.Drawing

Module Module1
    Dim s As Session = Session.GetSession()
    Dim dispPart As Part = s.Parts.Display()
    Dim workPart As Part = s.Parts.Work()
    Dim lastObjectChanged As Object = Nothing

    Sub Main()
        Dim form1 As New Form1
        form1.ShowDialog()
    End Sub


    Public Class Form1
        Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            txtBoxUserID.text = System.Environment.UserName
            txtBoxDate.text = DateTime.Now.ToString("MMM-dd-yyyy")
            txtBoxTime.Text = DateTime.Now.ToString("HH:mm")

            Try
                If s.parts.work.fullpath Is Nothing Then
                Else

                    txtBoxPartName.text = s.Parts.Work.FullPath()
                End If
            Catch ex As Exception
            End Try

            comBoxReason.items.add("Change to Part Attributes")
            comBoxReason.items.add("Design Change")
            comBoxReason.items.add("Drawing Change")
            comBoxReason.items.add("Other")
            comBoxReason.text = "Design Change"
        End Sub

        Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
            If comBoxReason.text = "Other" And txtBoxComments.text.trim = "" Then
                FailureTimerTick(txtBoxComments)
                Exit Sub
            End If

            sendEmail(Nothing)
            UpdateLogs()
            Me.Close()
        End Sub

        Private Sub FailureTimerTick(ByVal changedObject As Object)
            changedObject.BackColor = Color.Red
            lastObjectChanged = changedObject
            Timer1.Interval = 400
            Timer1.Start()
        End Sub

        Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
            Timer1.Stop()
            lastObjectChanged.BackColor = Color.Empty
        End Sub

        Public Sub sendEmail(ByVal body As String)
            Dim adminEmails As String = Nothing

            GetAdminEmails(adminEmails)

            Dim issueEmail As New System.Net.Mail.MailMessage
            Dim smtp_server As New System.Net.Mail.SmtpClient

            Dim emailBody As String = Nothing

            issueEmail.IsBodyHtml = True
            smtp_server.EnableSsl = False
            smtp_server.Host = "192.168.0.47"

            Dim tempFrom As String = Nothing
            tempFrom = System.Environment.UserName & "@stackteck.com" 'from field, is editable
            'ToBox.Text = System.Environment.UserName & "@stackteck.com"  'to field 

            'exception for usernames
            '-------------------------------------------------------------------

            If System.Environment.UserName = "htang" Or System.Environment.UserName = "henryt" Then
                tempFrom = "htang@stackteck.com"
            End If

            If System.Environment.UserName = "JOHN" Or System.Environment.UserName = "john" Then
                tempFrom = "jbrtka@stackteck.com"
            End If

            If System.Environment.UserName = "susan" Or System.Environment.UserName = "SUSAN" Then
                tempFrom = "shiltz@stackteck.com"
            End If

            If System.Environment.UserName = "wei" Or System.Environment.UserName = "WEI" Then
                tempFrom = "wyuan@stackteck.com"
            End If
            If System.Environment.UserName.ToLower = "mandeep" Or System.Environment.UserName.ToLower = "mthandi" Then
                tempFrom = "mthandi@stackteck.com"
            End If

            issueEmail.From = New System.Net.Mail.MailAddress(tempFrom)

            If (adminEmails.Trim <> "") Then
                issueEmail.To.Add(adminEmails + "jngai@stackteck.com, " + tempFrom) ' Remember to update Y:\eng\ENG_ACCESS_DATABASES\UGMisAttributes.mdb admin_emails table with your e-mail!
            Else
                issueEmail.To.Add("rnaveed@stackteck.com")
            End If

            If (IsNothing(body)) Then
                ' Add username
                ' Add file they were working on. 
                ' Create file for Admin list e-mails

                issueEmail.Subject = "Unissue Request by " + System.Environment.UserName ' Modify this line to be specific to the program

                ' For UG Specific programs, uncomment this
                'Try
                '    Dim s As Session = Session.GetSession
                '    Dim filePath As String = s.Parts.Work.FullPath() ' E.g. AIM_StackCup24oz_S37452/001
                '    emailBody = emailBody + "FileName: " + filePath + "<br>"
                'Catch ex As Exception
                'End Try

                emailBody = emailBody + "User Name: " + System.Environment.UserName + "<br><br>"
                emailBody = emailBody + "Part Name: " + txtBoxPartName.text + "<br><br>"
                emailBody = emailBody + "Reason: " + comBoxReason.Text + "<br><br>"
                emailBody = emailBody + "Comment: " + txtBoxComments.Text + "<br><br>"
                If (chkBoxShopFloor.checked = True) Then
                    emailBody = emailBody + "The drawing has reached the shop floor, DO NOT unissue!"
                Else
                    emailBody = emailBody + "The drawing has not reached the shop floor, unissue"
                End If

                issueEmail.Body = emailBody

                smtp_server.Send(issueEmail)
            Else
                Try
                    issueEmail.Subject = "Error in Unissue Request Program by " + System.Environment.UserName ' Modify this line to be specific to the program

                    emailBody = emailBody + "User Name: " + System.Environment.UserName + "<br><br>"
                    emailBody = emailBody + "Part Name: " + txtBoxPartName.text + "<br><br>"
                    emailBody = emailBody + "Reason: " + comBoxReason.Text + "<br><br>"
                    emailBody = emailBody + "Comment: " + txtBoxComments.Text + "<br><br>"
                    emailBody = emailBody + body
                    issueEmail.Body = emailBody

                Catch ex As Exception
                    issueEmail.Body = System.Environment.UserName + "<br>" + body
                End Try

                smtp_server.Send(issueEmail)
            End If
        End Sub

        Public Sub UpdateLogs()
            Dim Conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Y:\eng\ENG_ACCESS_DATABASES\UnissueRequestLogs.mdb; User Id=admin")
            Dim Com As OleDbCommand

            Dim sql = "INSERT INTO UNISSUE_REQUEST_LOGS (PARTNAME, USERID, DADATE, DATIME, REASON, COMMENT, DRAWING_REACHED_THE_SHOP_FLOOR) VALUES(?,?,?,?,?,?,?)"


            Com = New OleDbCommand(sql, Conn)
            Conn.Open()

            'If (txtBoxComments.Text.Length > 255) Then
            '    txtBoxComments.Text = txtBoxComments.Text.SubString(0, 255)
            'End If

            ' Step 0. Each top level assembly
            Try

                Com = New OleDbCommand(sql, Conn)
                Com.Parameters.AddWithValue("@p1", txtBoxPartName.Text)
                Com.Parameters.AddWithValue("@p2", txtBoxUserID.Text)
                Com.Parameters.AddWithValue("@p3", txtBoxDate.Text)
                Com.Parameters.AddWithValue("@p4", txtBoxTime.Text)
                Com.Parameters.AddWithValue("@p5", comBoxReason.Text)
                Com.Parameters.AddWithValue("@p6", txtBoxComments.Text)
                Com.Parameters.AddWithValue("@p7", chkBoxShopFloor.Checked.ToString)


                Com.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            Conn.Close()
        End Sub

        Public Sub GetAdminEmails(ByRef adminEmails As String)

            'Define the connectors
            Dim cn As OleDbConnection
            Dim cmd As OleDbCommand
            Dim dr As OleDbDataReader
            Dim oConnect, oQuery As String
            Dim FoundStatus As Boolean = False

            'Define connection string
            Dim FileName As String = "Y:\eng\ENG_ACCESS_DATABASES\UGMisDatabase.mdb"
            If File.Exists(FileName) = False Then
                Exit Sub
            End If

            oConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName

            'Query String
            oQuery = "SELECT * FROM ADMIN_EMAILS"

            'Instantiate the connectors
            cn = New OleDbConnection(oConnect)
            cn.Open()

            cmd = New OleDbCommand(oQuery, cn)
            dr = cmd.ExecuteReader

            While dr.Read()
                adminEmails += dr(1).Trim
                adminEmails += ", "
                FoundStatus = True
            End While

            'adminEmails = adminEmails.Substring(0, adminEmails.LastIndexOf(","))

            dr.Close()
            cn.Close()
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
            Me.components = New System.ComponentModel.Container()
            Me.partName = New System.Windows.Forms.Label()
            Me.txtBoxPartName = New System.Windows.Forms.TextBox()
            Me.txtBoxDate = New System.Windows.Forms.TextBox()
            Me.txtBoxTime = New System.Windows.Forms.TextBox()
            Me.txtBoxUserID = New System.Windows.Forms.TextBox()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.txtBoxComments = New System.Windows.Forms.TextBox()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.Label5 = New System.Windows.Forms.Label()
            Me.comBoxReason = New System.Windows.Forms.ComboBox()
            Me.btnSend = New System.Windows.Forms.Button()
            Me.btnCancel = New System.Windows.Forms.Button()
            Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
            Me.chkBoxShopFloor = New System.Windows.Forms.CheckBox()
            Me.SuspendLayout()
            '
            'partName
            '
            Me.partName.AutoSize = True
            Me.partName.Location = New System.Drawing.Point(13, 79)
            Me.partName.Name = "partName"
            Me.partName.Size = New System.Drawing.Size(57, 13)
            Me.partName.TabIndex = 0
            Me.partName.Text = "Part Name"
            '
            'txtBoxPartName
            '
            Me.txtBoxPartName.Location = New System.Drawing.Point(16, 95)
            Me.txtBoxPartName.Name = "txtBoxPartName"
            Me.txtBoxPartName.Size = New System.Drawing.Size(264, 20)
            Me.txtBoxPartName.TabIndex = 1
            '
            'txtBoxDate
            '
            Me.txtBoxDate.Enabled = False
            Me.txtBoxDate.Location = New System.Drawing.Point(59, 46)
            Me.txtBoxDate.Name = "txtBoxDate"
            Me.txtBoxDate.Size = New System.Drawing.Size(104, 20)
            Me.txtBoxDate.TabIndex = 85
            '
            'txtBoxTime
            '
            Me.txtBoxTime.Enabled = False
            Me.txtBoxTime.Location = New System.Drawing.Point(205, 46)
            Me.txtBoxTime.Name = "txtBoxTime"
            Me.txtBoxTime.Size = New System.Drawing.Size(75, 20)
            Me.txtBoxTime.TabIndex = 90
            '
            'txtBoxUserID
            '
            Me.txtBoxUserID.Enabled = False
            Me.txtBoxUserID.Location = New System.Drawing.Point(59, 11)
            Me.txtBoxUserID.Name = "txtBoxUserID"
            Me.txtBoxUserID.Size = New System.Drawing.Size(221, 20)
            Me.txtBoxUserID.TabIndex = 80
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Location = New System.Drawing.Point(13, 14)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(40, 13)
            Me.Label1.TabIndex = 5
            Me.Label1.Text = "UserID"
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Location = New System.Drawing.Point(13, 49)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(30, 13)
            Me.Label2.TabIndex = 6
            Me.Label2.Text = "Date"
            '
            'Label3
            '
            Me.Label3.AutoSize = True
            Me.Label3.Location = New System.Drawing.Point(169, 49)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(30, 13)
            Me.Label3.TabIndex = 7
            Me.Label3.Text = "Time"
            '
            'txtBoxComments
            '
            Me.txtBoxComments.Location = New System.Drawing.Point(16, 239)
            Me.txtBoxComments.Multiline = True
            Me.txtBoxComments.Name = "txtBoxComments"
            Me.txtBoxComments.Size = New System.Drawing.Size(264, 127)
            Me.txtBoxComments.TabIndex = 15
            '
            'Label4
            '
            Me.Label4.AutoSize = True
            Me.Label4.Location = New System.Drawing.Point(13, 223)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(56, 13)
            Me.Label4.TabIndex = 9
            Me.Label4.Text = "Comments"
            '
            'Label5
            '
            Me.Label5.AutoSize = True
            Me.Label5.Location = New System.Drawing.Point(13, 130)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(100, 13)
            Me.Label5.TabIndex = 10
            Me.Label5.Text = "Reason for Unissue"
            '
            'comBoxReason
            '
            Me.comBoxReason.FormattingEnabled = True
            Me.comBoxReason.Location = New System.Drawing.Point(16, 146)
            Me.comBoxReason.Name = "comBoxReason"
            Me.comBoxReason.Size = New System.Drawing.Size(264, 21)
            Me.comBoxReason.TabIndex = 5
            '
            'btnSend
            '
            Me.btnSend.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
            Me.btnSend.ForeColor = System.Drawing.Color.White
            Me.btnSend.Location = New System.Drawing.Point(96, 372)
            Me.btnSend.Name = "btnSend"
            Me.btnSend.Size = New System.Drawing.Size(112, 55)
            Me.btnSend.TabIndex = 20
            Me.btnSend.Text = "SEND EMAIL"
            Me.btnSend.UseVisualStyleBackColor = False
            '
            'btnCancel
            '
            Me.btnCancel.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
            Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnCancel.ForeColor = System.Drawing.Color.Yellow
            Me.btnCancel.Location = New System.Drawing.Point(115, 433)
            Me.btnCancel.Name = "btnCancel"
            Me.btnCancel.Size = New System.Drawing.Size(75, 23)
            Me.btnCancel.TabIndex = 30
            Me.btnCancel.Text = "Cancel"
            Me.btnCancel.UseVisualStyleBackColor = False
            '
            'chkBoxShopFloor
            '
            Me.chkBoxShopFloor.AutoSize = True
            Me.chkBoxShopFloor.Checked = True
            Me.chkBoxShopFloor.CheckState = System.Windows.Forms.CheckState.Checked
            Me.chkBoxShopFloor.ForeColor = System.Drawing.Color.DarkRed
            Me.chkBoxShopFloor.Location = New System.Drawing.Point(16, 188)
            Me.chkBoxShopFloor.Name = "chkBoxShopFloor"
            Me.chkBoxShopFloor.Size = New System.Drawing.Size(214, 17)
            Me.chkBoxShopFloor.TabIndex = 91
            Me.chkBoxShopFloor.Text = "The drawing has reached the shop floor"
            Me.chkBoxShopFloor.UseVisualStyleBackColor = True
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(292, 468)
            Me.Controls.Add(Me.chkBoxShopFloor)
            Me.Controls.Add(Me.btnCancel)
            Me.Controls.Add(Me.btnSend)
            Me.Controls.Add(Me.comBoxReason)
            Me.Controls.Add(Me.Label5)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.txtBoxComments)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.txtBoxUserID)
            Me.Controls.Add(Me.txtBoxTime)
            Me.Controls.Add(Me.txtBoxDate)
            Me.Controls.Add(Me.txtBoxPartName)
            Me.Controls.Add(Me.partName)
            Me.Name = "Form1"
            Me.Text = "Unissue Request Form"
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Friend WithEvents partName As Label
        Friend WithEvents txtBoxPartName As TextBox
        Friend WithEvents txtBoxDate As TextBox
        Friend WithEvents txtBoxTime As TextBox
        Friend WithEvents txtBoxUserID As TextBox
        Friend WithEvents Label1 As Label
        Friend WithEvents Label2 As Label
        Friend WithEvents Label3 As Label
        Friend WithEvents txtBoxComments As TextBox
        Friend WithEvents Label4 As Label
        Friend WithEvents Label5 As Label
        Friend WithEvents comBoxReason As ComboBox
        Friend WithEvents btnSend As Button
        Friend WithEvents btnCancel As Button
        Friend WithEvents Timer1 As Timer
        Friend WithEvents chkBoxShopFloor As CheckBox
    End Class

End Module
