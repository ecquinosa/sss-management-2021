<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmSSITcon
    Inherits DevComponents.DotNetBar.Metro.MetroForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbTime = New System.Windows.Forms.ComboBox()
        Me.btnNewSSIT = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.cbSettings = New System.Windows.Forms.ComboBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.txtName1 = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtDsn1 = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.btnsave1 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.txtDB1 = New System.Windows.Forms.TextBox()
        Me.txtPword1 = New System.Windows.Forms.TextBox()
        Me.txtUsername1 = New System.Windows.Forms.TextBox()
        Me.txtServer1 = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(858, 653)
        Me.Panel1.TabIndex = 0
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.Label1)
        Me.Panel4.Controls.Add(Me.cbTime)
        Me.Panel4.Controls.Add(Me.btnNewSSIT)
        Me.Panel4.Controls.Add(Me.Button5)
        Me.Panel4.Controls.Add(Me.cbSettings)
        Me.Panel4.Controls.Add(Me.Label22)
        Me.Panel4.Controls.Add(Me.Label21)
        Me.Panel4.Controls.Add(Me.txtName1)
        Me.Panel4.Controls.Add(Me.Label12)
        Me.Panel4.Controls.Add(Me.txtDsn1)
        Me.Panel4.Controls.Add(Me.Label16)
        Me.Panel4.Controls.Add(Me.btnsave1)
        Me.Panel4.Controls.Add(Me.Button6)
        Me.Panel4.Controls.Add(Me.Label17)
        Me.Panel4.Controls.Add(Me.Label18)
        Me.Panel4.Controls.Add(Me.Label19)
        Me.Panel4.Controls.Add(Me.Label20)
        Me.Panel4.Controls.Add(Me.txtDB1)
        Me.Panel4.Controls.Add(Me.txtPword1)
        Me.Panel4.Controls.Add(Me.txtUsername1)
        Me.Panel4.Controls.Add(Me.txtServer1)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.Location = New System.Drawing.Point(0, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(858, 653)
        Me.Panel4.TabIndex = 31
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(43, 418)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(266, 30)
        Me.Label1.TabIndex = 41
        Me.Label1.Text = "TIME TO TRANSFER LOGS :"
        '
        'cbTime
        '
        Me.cbTime.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cbTime.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbTime.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbTime.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbTime.ForeColor = System.Drawing.SystemColors.InfoText
        Me.cbTime.FormattingEnabled = True
        Me.cbTime.Items.AddRange(New Object() {"1:00:00 AM", "1:30:00 AM", "2:00:00 AM", "2:30:00 AM", "3:00:00 AM", "3:30:00 AM", "4:00:00 AM", "4:30:00 AM", "5:00:00 AM", "5:30:00 AM", "6:00:00 AM", "6:30:00 AM", "7:00:00 AM", "7:30:00 AM", "8:00:00 AM", "8:30:00 AM", "9:00:00 AM", "9:30:00 AM", "10:00:00 AM", "10:30:00 AM", "11:00:00 AM", "11:30:00 AM", "12:00:00 AM", "12:30:00 AM", "1:00:00 PM", "1:30:00 PM", "2:00:00 PM", "2:30:00 PM", "3:00:00 PM", "3:30:00 PM", "4:00:00 AM", "4:30:00 PM", "5:00:00 PM", "5:30:00 PM", "6:00:00 PM", "6:30:00 PM", "7:00:00 PM", "7:30:00 PM", "8:00:00 PM", "8:30:00 PM", "9:00:00 PM", "9:30:00 PM", "10:00:00 PM", "10:30:00 PM", "11:00:00 PM", "11:30:00 PM", "12:00:00 PM", "12:30:00 PM"})
        Me.cbTime.Location = New System.Drawing.Point(315, 415)
        Me.cbTime.Name = "cbTime"
        Me.cbTime.Size = New System.Drawing.Size(183, 38)
        Me.cbTime.TabIndex = 40
        '
        'btnNewSSIT
        '
        Me.btnNewSSIT.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnNewSSIT.Enabled = False
        Me.btnNewSSIT.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNewSSIT.Location = New System.Drawing.Point(292, 459)
        Me.btnNewSSIT.Name = "btnNewSSIT"
        Me.btnNewSSIT.Size = New System.Drawing.Size(241, 91)
        Me.btnNewSSIT.TabIndex = 39
        Me.btnNewSSIT.Text = "New Settings"
        Me.btnNewSSIT.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button5.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(616, 118)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(204, 50)
        Me.Button5.TabIndex = 19
        Me.Button5.Text = "Remove Settings"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'cbSettings
        '
        Me.cbSettings.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cbSettings.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbSettings.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbSettings.ForeColor = System.Drawing.SystemColors.InfoText
        Me.cbSettings.FormattingEnabled = True
        Me.cbSettings.Location = New System.Drawing.Point(213, 125)
        Me.cbSettings.Name = "cbSettings"
        Me.cbSettings.Size = New System.Drawing.Size(397, 38)
        Me.cbSettings.TabIndex = 18
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(92, 128)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(115, 30)
        Me.Label22.TabIndex = 17
        Me.Label22.Text = "SETTINGS :"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(26, 172)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(181, 30)
        Me.Label21.TabIndex = 16
        Me.Label21.Text = "NAME SETTINGS :"
        '
        'txtName1
        '
        Me.txtName1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName1.Location = New System.Drawing.Point(213, 169)
        Me.txtName1.Name = "txtName1"
        Me.txtName1.Size = New System.Drawing.Size(397, 35)
        Me.txtName1.TabIndex = 15
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(141, 213)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(66, 30)
        Me.Label12.TabIndex = 14
        Me.Label12.Text = "DSN :"
        '
        'txtDsn1
        '
        Me.txtDsn1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDsn1.Location = New System.Drawing.Point(213, 210)
        Me.txtDsn1.Name = "txtDsn1"
        Me.txtDsn1.Size = New System.Drawing.Size(397, 35)
        Me.txtDsn1.TabIndex = 1
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(90, 43)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(655, 42)
        Me.Label16.TabIndex = 12
        Me.Label16.Text = "SSIT SQL DATABASE CONNECTION"
        '
        'btnsave1
        '
        Me.btnsave1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnsave1.Enabled = False
        Me.btnsave1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnsave1.Location = New System.Drawing.Point(414, 554)
        Me.btnsave1.Name = "btnsave1"
        Me.btnsave1.Size = New System.Drawing.Size(241, 92)
        Me.btnsave1.TabIndex = 7
        Me.btnsave1.Text = "Save Settings"
        Me.btnsave1.UseVisualStyleBackColor = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button6.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.Location = New System.Drawing.Point(167, 554)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(241, 92)
        Me.Button6.TabIndex = 6
        Me.Button6.Text = "Test Connection"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(90, 377)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(117, 30)
        Me.Label17.TabIndex = 9
        Me.Label17.Text = "DB NAME :"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(71, 336)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(136, 30)
        Me.Label18.TabIndex = 8
        Me.Label18.Text = "PASSWORD :"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(74, 295)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(133, 30)
        Me.Label19.TabIndex = 7
        Me.Label19.Text = "USERNAME :"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(111, 254)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(96, 30)
        Me.Label20.TabIndex = 6
        Me.Label20.Text = "SERVER :"
        '
        'txtDB1
        '
        Me.txtDB1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDB1.Location = New System.Drawing.Point(213, 374)
        Me.txtDB1.Name = "txtDB1"
        Me.txtDB1.Size = New System.Drawing.Size(397, 35)
        Me.txtDB1.TabIndex = 5
        Me.txtDB1.Text = "SSIT_SERVER"
        '
        'txtPword1
        '
        Me.txtPword1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPword1.Location = New System.Drawing.Point(213, 333)
        Me.txtPword1.Name = "txtPword1"
        Me.txtPword1.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPword1.Size = New System.Drawing.Size(397, 35)
        Me.txtPword1.TabIndex = 4
        '
        'txtUsername1
        '
        Me.txtUsername1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUsername1.Location = New System.Drawing.Point(213, 292)
        Me.txtUsername1.Name = "txtUsername1"
        Me.txtUsername1.Size = New System.Drawing.Size(397, 35)
        Me.txtUsername1.TabIndex = 3
        '
        'txtServer1
        '
        Me.txtServer1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServer1.Location = New System.Drawing.Point(213, 251)
        Me.txtServer1.Name = "txtServer1"
        Me.txtServer1.Size = New System.Drawing.Size(397, 35)
        Me.txtServer1.TabIndex = 2
        '
        '_frmSSITcon
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(858, 653)
        Me.Controls.Add(Me.Panel1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "_frmSSITcon"
        Me.Text = "MetroForm"
        Me.Panel1.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents btnNewSSIT As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents cbSettings As System.Windows.Forms.ComboBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtName1 As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtDsn1 As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents btnsave1 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents txtDB1 As System.Windows.Forms.TextBox
    Friend WithEvents txtPword1 As System.Windows.Forms.TextBox
    Friend WithEvents txtUsername1 As System.Windows.Forms.TextBox
    Friend WithEvents txtServer1 As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbTime As System.Windows.Forms.ComboBox
End Class
