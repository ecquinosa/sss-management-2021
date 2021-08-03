<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmORA
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
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cbOraSettings = New System.Windows.Forms.ComboBox()
        Me.btnNewOra = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.txtOra = New System.Windows.Forms.TextBox()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.txtPort = New System.Windows.Forms.TextBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.txtUserID = New System.Windows.Forms.TextBox()
        Me.txtHost = New System.Windows.Forms.TextBox()
        Me.txtOrapass = New System.Windows.Forms.TextBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.txtServiceName = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Panel6)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(858, 653)
        Me.Panel1.TabIndex = 0
        '
        'Panel6
        '
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel6.Controls.Add(Me.Button1)
        Me.Panel6.Controls.Add(Me.cbOraSettings)
        Me.Panel6.Controls.Add(Me.btnNewOra)
        Me.Panel6.Controls.Add(Me.Button10)
        Me.Panel6.Controls.Add(Me.Label33)
        Me.Panel6.Controls.Add(Me.Label34)
        Me.Panel6.Controls.Add(Me.Label27)
        Me.Panel6.Controls.Add(Me.txtOra)
        Me.Panel6.Controls.Add(Me.Button8)
        Me.Panel6.Controls.Add(Me.txtPort)
        Me.Panel6.Controls.Add(Me.Label28)
        Me.Panel6.Controls.Add(Me.txtUserID)
        Me.Panel6.Controls.Add(Me.txtHost)
        Me.Panel6.Controls.Add(Me.txtOrapass)
        Me.Panel6.Controls.Add(Me.Label29)
        Me.Panel6.Controls.Add(Me.txtServiceName)
        Me.Panel6.Controls.Add(Me.Label30)
        Me.Panel6.Controls.Add(Me.Label31)
        Me.Panel6.Controls.Add(Me.Label32)
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel6.Location = New System.Drawing.Point(0, 0)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(858, 653)
        Me.Panel6.TabIndex = 56
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(311, 449)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(241, 92)
        Me.Button1.TabIndex = 40
        Me.Button1.Text = "Test Connection"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'cbOraSettings
        '
        Me.cbOraSettings.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cbOraSettings.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbOraSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbOraSettings.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbOraSettings.ForeColor = System.Drawing.SystemColors.InfoText
        Me.cbOraSettings.FormattingEnabled = True
        Me.cbOraSettings.Location = New System.Drawing.Point(202, 130)
        Me.cbOraSettings.Name = "cbOraSettings"
        Me.cbOraSettings.Size = New System.Drawing.Size(416, 38)
        Me.cbOraSettings.TabIndex = 39
        '
        'btnNewOra
        '
        Me.btnNewOra.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnNewOra.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNewOra.Location = New System.Drawing.Point(184, 547)
        Me.btnNewOra.Name = "btnNewOra"
        Me.btnNewOra.Size = New System.Drawing.Size(241, 92)
        Me.btnNewOra.TabIndex = 23
        Me.btnNewOra.Text = "New Settings"
        Me.btnNewOra.UseVisualStyleBackColor = False
        '
        'Button10
        '
        Me.Button10.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button10.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button10.Location = New System.Drawing.Point(624, 124)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(204, 50)
        Me.Button10.TabIndex = 25
        Me.Button10.Text = "Remove Settings"
        Me.Button10.UseVisualStyleBackColor = False
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label33.Location = New System.Drawing.Point(15, 179)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(181, 30)
        Me.Label33.TabIndex = 24
        Me.Label33.Text = "NAME SETTINGS :"
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.Location = New System.Drawing.Point(81, 133)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(115, 30)
        Me.Label34.TabIndex = 23
        Me.Label34.Text = "SETTINGS :"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.Location = New System.Drawing.Point(46, 60)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(734, 42)
        Me.Label27.TabIndex = 23
        Me.Label27.Text = "SSIT ORACLE DATABASE CONNECTION"
        '
        'txtOra
        '
        Me.txtOra.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOra.Location = New System.Drawing.Point(202, 174)
        Me.txtOra.Name = "txtOra"
        Me.txtOra.Size = New System.Drawing.Size(416, 35)
        Me.txtOra.TabIndex = 23
        '
        'Button8
        '
        Me.Button8.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button8.Enabled = False
        Me.Button8.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button8.Location = New System.Drawing.Point(431, 547)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(241, 92)
        Me.Button8.TabIndex = 24
        Me.Button8.Text = "Save Settings"
        Me.Button8.UseVisualStyleBackColor = False
        '
        'txtPort
        '
        Me.txtPort.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPort.Location = New System.Drawing.Point(202, 256)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(416, 35)
        Me.txtPort.TabIndex = 29
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.Color.Black
        Me.Label28.Location = New System.Drawing.Point(119, 220)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(77, 30)
        Me.Label28.TabIndex = 38
        Me.Label28.Text = "HOST :"
        '
        'txtUserID
        '
        Me.txtUserID.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserID.Location = New System.Drawing.Point(202, 297)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(416, 35)
        Me.txtUserID.TabIndex = 30
        '
        'txtHost
        '
        Me.txtHost.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHost.Location = New System.Drawing.Point(202, 215)
        Me.txtHost.Name = "txtHost"
        Me.txtHost.Size = New System.Drawing.Size(416, 35)
        Me.txtHost.TabIndex = 37
        Me.txtHost.Text = "1521"
        '
        'txtOrapass
        '
        Me.txtOrapass.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOrapass.Location = New System.Drawing.Point(202, 338)
        Me.txtOrapass.Name = "txtOrapass"
        Me.txtOrapass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtOrapass.Size = New System.Drawing.Size(416, 35)
        Me.txtOrapass.TabIndex = 31
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.ForeColor = System.Drawing.Color.Black
        Me.Label29.Location = New System.Drawing.Point(28, 382)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(168, 30)
        Me.Label29.TabIndex = 36
        Me.Label29.Text = "SERVICE NAME :"
        '
        'txtServiceName
        '
        Me.txtServiceName.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServiceName.Location = New System.Drawing.Point(202, 379)
        Me.txtServiceName.Name = "txtServiceName"
        Me.txtServiceName.Size = New System.Drawing.Size(416, 35)
        Me.txtServiceName.TabIndex = 32
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.ForeColor = System.Drawing.Color.Black
        Me.Label30.Location = New System.Drawing.Point(60, 341)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(136, 30)
        Me.Label30.TabIndex = 35
        Me.Label30.Text = "PASSWORD :"
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.ForeColor = System.Drawing.Color.Black
        Me.Label31.Location = New System.Drawing.Point(120, 259)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(76, 30)
        Me.Label31.TabIndex = 33
        Me.Label31.Text = "PORT :"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.ForeColor = System.Drawing.Color.Black
        Me.Label32.Location = New System.Drawing.Point(96, 300)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(100, 30)
        Me.Label32.TabIndex = 34
        Me.Label32.Text = "USER ID :"
        '
        '_frmORA
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(858, 653)
        Me.Controls.Add(Me.Panel1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "_frmORA"
        Me.Text = "MetroForm"
        Me.Panel1.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents btnNewOra As System.Windows.Forms.Button
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents txtOra As System.Windows.Forms.TextBox
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents txtHost As System.Windows.Forms.TextBox
    Friend WithEvents txtOrapass As System.Windows.Forms.TextBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents txtServiceName As System.Windows.Forms.TextBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents cbOraSettings As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
