<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmSettings
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.ExpandablePanel1 = New DevComponents.DotNetBar.ExpandablePanel()
        Me.btnUP = New System.Windows.Forms.Button()
        Me.btnIdle = New System.Windows.Forms.Button()
        Me.btnFTP = New System.Windows.Forms.Button()
        Me.btnORA = New System.Windows.Forms.Button()
        Me.btnSSIT = New System.Windows.Forms.Button()
        Me.btnServer = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.ExpandablePanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.AutoSize = True
        Me.Panel1.Controls.Add(Me.SplitContainer1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1122, 691)
        Me.Panel1.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.SplitContainer1.Panel1.Controls.Add(Me.ExpandablePanel1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.SplitContainer1.Size = New System.Drawing.Size(1122, 691)
        Me.SplitContainer1.SplitterDistance = 244
        Me.SplitContainer1.TabIndex = 0
        '
        'ExpandablePanel1
        '
        Me.ExpandablePanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Windows7
        Me.ExpandablePanel1.Controls.Add(Me.btnUP)
        Me.ExpandablePanel1.Controls.Add(Me.btnIdle)
        Me.ExpandablePanel1.Controls.Add(Me.btnFTP)
        Me.ExpandablePanel1.Controls.Add(Me.btnORA)
        Me.ExpandablePanel1.Controls.Add(Me.btnSSIT)
        Me.ExpandablePanel1.Controls.Add(Me.btnServer)
        Me.ExpandablePanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ExpandablePanel1.ExpandButtonVisible = False
        Me.ExpandablePanel1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExpandablePanel1.Location = New System.Drawing.Point(0, 0)
        Me.ExpandablePanel1.Name = "ExpandablePanel1"
        Me.ExpandablePanel1.Size = New System.Drawing.Size(244, 691)
        Me.ExpandablePanel1.Style.BackColor1.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.ExpandablePanel1.Style.BackColor2.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.ExpandablePanel1.Style.Border = DevComponents.DotNetBar.eBorderType.SingleLine
        Me.ExpandablePanel1.Style.BorderColor.Color = System.Drawing.Color.FromArgb(CType(CType(3, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(148, Byte), Integer))
        Me.ExpandablePanel1.Style.BorderSide = DevComponents.DotNetBar.eBorderSide.None
        Me.ExpandablePanel1.Style.CornerDiameter = 0
        Me.ExpandablePanel1.Style.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExpandablePanel1.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.ExpandablePanel1.TabIndex = 4
        Me.ExpandablePanel1.TitleStyle.Alignment = System.Drawing.StringAlignment.Center
        Me.ExpandablePanel1.TitleStyle.BackColor1.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.ExpandablePanel1.TitleStyle.BackColor2.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.ExpandablePanel1.TitleStyle.Border = DevComponents.DotNetBar.eBorderType.RaisedInner
        Me.ExpandablePanel1.TitleStyle.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExpandablePanel1.TitleStyle.ForeColor.Color = System.Drawing.Color.White
        Me.ExpandablePanel1.TitleStyle.GradientAngle = 90
        Me.ExpandablePanel1.TitleText = "SYSTEM SETTINGS"
        '
        'btnUP
        '
        Me.btnUP.BackColor = System.Drawing.Color.White
        Me.btnUP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnUP.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUP.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUP.ForeColor = System.Drawing.Color.Black
        Me.btnUP.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnUP.Location = New System.Drawing.Point(12, 242)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(225, 43)
        Me.btnUP.TabIndex = 7
        Me.btnUP.Text = "File Upload"
        Me.btnUP.UseVisualStyleBackColor = False
        '
        'btnIdle
        '
        Me.btnIdle.BackColor = System.Drawing.Color.White
        Me.btnIdle.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnIdle.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnIdle.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnIdle.ForeColor = System.Drawing.Color.Black
        Me.btnIdle.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnIdle.Location = New System.Drawing.Point(12, 291)
        Me.btnIdle.Name = "btnIdle"
        Me.btnIdle.Size = New System.Drawing.Size(225, 43)
        Me.btnIdle.TabIndex = 5
        Me.btnIdle.Text = "Idle Time Configuration"
        Me.btnIdle.UseVisualStyleBackColor = False
        '
        'btnFTP
        '
        Me.btnFTP.BackColor = System.Drawing.Color.White
        Me.btnFTP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnFTP.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFTP.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFTP.ForeColor = System.Drawing.Color.Black
        Me.btnFTP.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnFTP.Location = New System.Drawing.Point(12, 193)
        Me.btnFTP.Name = "btnFTP"
        Me.btnFTP.Size = New System.Drawing.Size(225, 43)
        Me.btnFTP.TabIndex = 4
        Me.btnFTP.Text = "FTP Settings"
        Me.btnFTP.UseVisualStyleBackColor = False
        '
        'btnORA
        '
        Me.btnORA.BackColor = System.Drawing.Color.White
        Me.btnORA.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnORA.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnORA.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnORA.ForeColor = System.Drawing.Color.Black
        Me.btnORA.Location = New System.Drawing.Point(12, 144)
        Me.btnORA.Name = "btnORA"
        Me.btnORA.Size = New System.Drawing.Size(225, 43)
        Me.btnORA.TabIndex = 3
        Me.btnORA.Text = "Oracle Connection Settings"
        Me.btnORA.UseVisualStyleBackColor = False
        '
        'btnSSIT
        '
        Me.btnSSIT.BackColor = System.Drawing.Color.White
        Me.btnSSIT.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSSIT.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSSIT.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSSIT.ForeColor = System.Drawing.Color.Black
        Me.btnSSIT.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSSIT.Location = New System.Drawing.Point(12, 95)
        Me.btnSSIT.Name = "btnSSIT"
        Me.btnSSIT.Size = New System.Drawing.Size(225, 43)
        Me.btnSSIT.TabIndex = 2
        Me.btnSSIT.Text = "SSIT Connection Settings"
        Me.btnSSIT.UseVisualStyleBackColor = False
        '
        'btnServer
        '
        Me.btnServer.BackColor = System.Drawing.Color.White
        Me.btnServer.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnServer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnServer.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnServer.ForeColor = System.Drawing.Color.Black
        Me.btnServer.Location = New System.Drawing.Point(12, 47)
        Me.btnServer.Name = "btnServer"
        Me.btnServer.Size = New System.Drawing.Size(225, 42)
        Me.btnServer.TabIndex = 1
        Me.btnServer.Text = "Management Server Connection"
        Me.btnServer.UseVisualStyleBackColor = False
        '
        '_frmSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1122, 691)
        Me.Controls.Add(Me.Panel1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "_frmSettings"
        Me.Text = "MetroForm"
        Me.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.ExpandablePanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents ExpandablePanel1 As DevComponents.DotNetBar.ExpandablePanel
    Friend WithEvents btnIdle As System.Windows.Forms.Button
    Friend WithEvents btnFTP As System.Windows.Forms.Button
    Friend WithEvents btnORA As System.Windows.Forms.Button
    Friend WithEvents btnSSIT As System.Windows.Forms.Button
    Friend WithEvents btnServer As System.Windows.Forms.Button
    Friend WithEvents btnUP As System.Windows.Forms.Button
End Class
