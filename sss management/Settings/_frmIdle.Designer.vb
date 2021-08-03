<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmIdle
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
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnServer = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnSaveMGMT = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.lblIT = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnSSIT = New System.Windows.Forms.Button()
        Me.btnSSITSave = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSSITtime = New System.Windows.Forms.TextBox()
        Me.lblSIT = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Panel2)
        Me.Panel3.Controls.Add(Me.Panel1)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(858, 653)
        Me.Panel3.TabIndex = 30
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.SplitContainer1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 100)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(854, 549)
        Me.Panel2.TabIndex = 21
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Button1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnServer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnSaveMGMT)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.lblIT)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSSIT)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSSITSave)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtSSITtime)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblSIT)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer1.Size = New System.Drawing.Size(854, 549)
        Me.SplitContainer1.SplitterDistance = 416
        Me.SplitContainer1.TabIndex = 19
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(98, 55)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(236, 15)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "Adjust the idle time for management server"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(125, 366)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(177, 46)
        Me.Button1.TabIndex = 27
        Me.Button1.Text = "Cancel"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'btnServer
        '
        Me.btnServer.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnServer.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnServer.Location = New System.Drawing.Point(41, 314)
        Me.btnServer.Name = "btnServer"
        Me.btnServer.Size = New System.Drawing.Size(177, 46)
        Me.btnServer.TabIndex = 26
        Me.btnServer.Text = "New Settings"
        Me.btnServer.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(120, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(200, 30)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Management Server"
        '
        'btnSaveMGMT
        '
        Me.btnSaveMGMT.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSaveMGMT.Enabled = False
        Me.btnSaveMGMT.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveMGMT.Location = New System.Drawing.Point(224, 314)
        Me.btnSaveMGMT.Name = "btnSaveMGMT"
        Me.btnSaveMGMT.Size = New System.Drawing.Size(177, 46)
        Me.btnSaveMGMT.TabIndex = 16
        Me.btnSaveMGMT.Text = "Save Settings"
        Me.btnSaveMGMT.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 222)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(215, 21)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "NEW IDLE TIME IN MINUTES :"
        '
        'TextBox1
        '
        Me.TextBox1.Enabled = False
        Me.TextBox1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(224, 212)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(169, 35)
        Me.TextBox1.TabIndex = 6
        '
        'lblIT
        '
        Me.lblIT.AutoSize = True
        Me.lblIT.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIT.Location = New System.Drawing.Point(3, 183)
        Me.lblIT.Name = "lblIT"
        Me.lblIT.Size = New System.Drawing.Size(196, 25)
        Me.lblIT.TabIndex = 15
        Me.lblIT.Text = "CURRENT IDLE TIME : "
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(102, 55)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(208, 15)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "Adjust the idle time for the whole SSIT"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button2.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(122, 366)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(177, 46)
        Me.Button2.TabIndex = 28
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'btnSSIT
        '
        Me.btnSSIT.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSSIT.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSSIT.Location = New System.Drawing.Point(33, 314)
        Me.btnSSIT.Name = "btnSSIT"
        Me.btnSSIT.Size = New System.Drawing.Size(177, 46)
        Me.btnSSIT.TabIndex = 25
        Me.btnSSIT.Text = "New Settings"
        Me.btnSSIT.UseVisualStyleBackColor = False
        '
        'btnSSITSave
        '
        Me.btnSSITSave.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnSSITSave.Enabled = False
        Me.btnSSITSave.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSSITSave.Location = New System.Drawing.Point(216, 314)
        Me.btnSSITSave.Name = "btnSSITSave"
        Me.btnSSITSave.Size = New System.Drawing.Size(177, 46)
        Me.btnSSITSave.TabIndex = 23
        Me.btnSSITSave.Text = "Save Settings"
        Me.btnSSITSave.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(14, 222)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(215, 21)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "NEW IDLE TIME IN MINUTES :"
        '
        'txtSSITtime
        '
        Me.txtSSITtime.Enabled = False
        Me.txtSSITtime.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSSITtime.Location = New System.Drawing.Point(235, 212)
        Me.txtSSITtime.Name = "txtSSITtime"
        Me.txtSSITtime.Size = New System.Drawing.Size(169, 35)
        Me.txtSSITtime.TabIndex = 21
        '
        'lblSIT
        '
        Me.lblSIT.AutoSize = True
        Me.lblSIT.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSIT.Location = New System.Drawing.Point(14, 183)
        Me.lblSIT.Name = "lblSIT"
        Me.lblSIT.Size = New System.Drawing.Size(196, 25)
        Me.lblSIT.TabIndex = 22
        Me.lblSIT.Text = "CURRENT IDLE TIME : "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(57, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(321, 30)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Self-Service Information Terminal"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(854, 100)
        Me.Panel1.TabIndex = 20
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(174, 27)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(519, 42)
        Me.Label14.TabIndex = 13
        Me.Label14.Text = "IDLE TIME CONFIGURATION"
        '
        '_frmIdle
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(858, 653)
        Me.Controls.Add(Me.Panel3)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "_frmIdle"
        Me.Text = "MetroForm"
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnSaveMGMT As System.Windows.Forms.Button
    Friend WithEvents lblIT As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnSSITSave As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSSITtime As System.Windows.Forms.TextBox
    Friend WithEvents lblSIT As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnSSIT As System.Windows.Forms.Button
    Friend WithEvents btnServer As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
End Class
