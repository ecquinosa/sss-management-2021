<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmMain
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(_frmMain))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.exPan2 = New DevComponents.DotNetBar.ExpandablePanel()
        Me.btnUserRole = New System.Windows.Forms.Button()
        Me.btnUser = New System.Windows.Forms.Button()
        Me.exPan3 = New DevComponents.DotNetBar.ExpandablePanel()
        Me.btnDatabaseSet = New System.Windows.Forms.Button()
        Me.btnLogout = New System.Windows.Forms.Button()
        Me.ExpandablePanel1 = New DevComponents.DotNetBar.ExpandablePanel()
        Me.btnFB = New System.Windows.Forms.Button()
        Me.btnReports = New System.Windows.Forms.Button()
        Me.btnMonitoring = New System.Windows.Forms.Button()
        Me.btnKiosk = New System.Windows.Forms.Button()
        Me.btnDash = New System.Windows.Forms.Button()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.exPan2.SuspendLayout()
        Me.exPan3.SuspendLayout()
        Me.ExpandablePanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1284, 702)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.SplitContainer1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1284, 702)
        Me.Panel2.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Button1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.exPan2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.exPan3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ExpandablePanel1)
        Me.SplitContainer1.Size = New System.Drawing.Size(1284, 702)
        Me.SplitContainer1.SplitterDistance = 246
        Me.SplitContainer1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(8, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(25, 30)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "__"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'exPan2
        '
        Me.exPan2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.exPan2.CanvasColor = System.Drawing.Color.Navy
        Me.exPan2.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.OfficeXP
        Me.exPan2.Controls.Add(Me.btnUserRole)
        Me.exPan2.Controls.Add(Me.btnUser)
        Me.exPan2.DisabledBackColor = System.Drawing.Color.Empty
        Me.exPan2.ExpandButtonVisible = False
        Me.exPan2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exPan2.Location = New System.Drawing.Point(21, 370)
        Me.exPan2.Name = "exPan2"
        Me.exPan2.Size = New System.Drawing.Size(222, 174)
        Me.exPan2.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.exPan2.Style.BackColor1.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.exPan2.Style.BackColor2.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.exPan2.Style.BorderColor.Color = System.Drawing.Color.FromArgb(CType(CType(3, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(148, Byte), Integer))
        Me.exPan2.Style.CornerDiameter = 22
        Me.exPan2.Style.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exPan2.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.exPan2.Style.GradientAngle = 90
        Me.exPan2.TabIndex = 4
        Me.exPan2.TitleStyle.Alignment = System.Drawing.StringAlignment.Center
        Me.exPan2.TitleStyle.BackColor1.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.exPan2.TitleStyle.BackColor2.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.exPan2.TitleStyle.Border = DevComponents.DotNetBar.eBorderType.SingleLine
        Me.exPan2.TitleStyle.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder
        Me.exPan2.TitleStyle.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exPan2.TitleStyle.ForeColor.Color = System.Drawing.Color.White
        Me.exPan2.TitleStyle.GradientAngle = 90
        Me.exPan2.TitleText = "USER CONFIGURATION"
        '
        'btnUserRole
        '
        Me.btnUserRole.BackColor = System.Drawing.Color.White
        Me.btnUserRole.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnUserRole.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUserRole.ForeColor = System.Drawing.Color.Black
        Me.btnUserRole.Location = New System.Drawing.Point(4, 89)
        Me.btnUserRole.Name = "btnUserRole"
        Me.btnUserRole.Size = New System.Drawing.Size(225, 58)
        Me.btnUserRole.TabIndex = 3
        Me.btnUserRole.Text = "User Role"
        Me.btnUserRole.UseVisualStyleBackColor = False
        '
        'btnUser
        '
        Me.btnUser.BackColor = System.Drawing.Color.White
        Me.btnUser.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnUser.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUser.ForeColor = System.Drawing.Color.Black
        Me.btnUser.Location = New System.Drawing.Point(4, 26)
        Me.btnUser.Name = "btnUser"
        Me.btnUser.Size = New System.Drawing.Size(225, 58)
        Me.btnUser.TabIndex = 2
        Me.btnUser.Text = "User Management"
        Me.btnUser.UseVisualStyleBackColor = False
        '
        'exPan3
        '
        Me.exPan3.CanvasColor = System.Drawing.Color.Navy
        Me.exPan3.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2000
        Me.exPan3.Controls.Add(Me.btnDatabaseSet)
        Me.exPan3.Controls.Add(Me.btnLogout)
        Me.exPan3.DisabledBackColor = System.Drawing.Color.Empty
        Me.exPan3.ExpandButtonVisible = False
        Me.exPan3.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exPan3.Location = New System.Drawing.Point(21, 550)
        Me.exPan3.Name = "exPan3"
        Me.exPan3.Size = New System.Drawing.Size(238, 175)
        Me.exPan3.Style.Alignment = System.Drawing.StringAlignment.Center
        Me.exPan3.Style.BackColor1.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.exPan3.Style.BackColor2.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.exPan3.Style.BorderColor.Color = System.Drawing.Color.FromArgb(CType(CType(3, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(148, Byte), Integer))
        Me.exPan3.Style.CornerDiameter = 22
        Me.exPan3.Style.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exPan3.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.exPan3.Style.GradientAngle = 90
        Me.exPan3.TabIndex = 5
        Me.exPan3.TitleStyle.Alignment = System.Drawing.StringAlignment.Center
        Me.exPan3.TitleStyle.BackColor1.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.exPan3.TitleStyle.BackColor2.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.exPan3.TitleStyle.Border = DevComponents.DotNetBar.eBorderType.SingleLine
        Me.exPan3.TitleStyle.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder
        Me.exPan3.TitleStyle.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exPan3.TitleStyle.ForeColor.Color = System.Drawing.Color.White
        Me.exPan3.TitleStyle.GradientAngle = 90
        Me.exPan3.TitleText = "SYSTEM CONFIGURATION"
        '
        'btnDatabaseSet
        '
        Me.btnDatabaseSet.BackColor = System.Drawing.Color.White
        Me.btnDatabaseSet.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDatabaseSet.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDatabaseSet.ForeColor = System.Drawing.Color.Black
        Me.btnDatabaseSet.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDatabaseSet.Location = New System.Drawing.Point(4, 28)
        Me.btnDatabaseSet.Name = "btnDatabaseSet"
        Me.btnDatabaseSet.Size = New System.Drawing.Size(225, 58)
        Me.btnDatabaseSet.TabIndex = 6
        Me.btnDatabaseSet.Text = "System Settings"
        Me.btnDatabaseSet.UseVisualStyleBackColor = False
        '
        'btnLogout
        '
        Me.btnLogout.BackColor = System.Drawing.Color.White
        Me.btnLogout.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnLogout.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLogout.ForeColor = System.Drawing.Color.Black
        Me.btnLogout.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnLogout.Location = New System.Drawing.Point(4, 91)
        Me.btnLogout.Name = "btnLogout"
        Me.btnLogout.Size = New System.Drawing.Size(225, 58)
        Me.btnLogout.TabIndex = 5
        Me.btnLogout.Text = "Log Out"
        Me.btnLogout.UseVisualStyleBackColor = False
        '
        'ExpandablePanel1
        '
        Me.ExpandablePanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ExpandablePanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Windows7
        Me.ExpandablePanel1.Controls.Add(Me.btnFB)
        Me.ExpandablePanel1.Controls.Add(Me.btnReports)
        Me.ExpandablePanel1.Controls.Add(Me.btnMonitoring)
        Me.ExpandablePanel1.Controls.Add(Me.btnKiosk)
        Me.ExpandablePanel1.Controls.Add(Me.btnDash)
        Me.ExpandablePanel1.DisabledBackColor = System.Drawing.Color.Empty
        Me.ExpandablePanel1.ExpandButtonVisible = False
        Me.ExpandablePanel1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExpandablePanel1.Location = New System.Drawing.Point(24, 10)
        Me.ExpandablePanel1.Name = "ExpandablePanel1"
        Me.ExpandablePanel1.Size = New System.Drawing.Size(219, 354)
        Me.ExpandablePanel1.Style.BackColor1.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.ExpandablePanel1.Style.BackColor2.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.ExpandablePanel1.Style.Border = DevComponents.DotNetBar.eBorderType.SingleLine
        Me.ExpandablePanel1.Style.BorderColor.Color = System.Drawing.Color.FromArgb(CType(CType(3, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(148, Byte), Integer))
        Me.ExpandablePanel1.Style.CornerDiameter = 0
        Me.ExpandablePanel1.Style.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExpandablePanel1.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.ItemText
        Me.ExpandablePanel1.TabIndex = 3
        Me.ExpandablePanel1.TitleStyle.Alignment = System.Drawing.StringAlignment.Center
        Me.ExpandablePanel1.TitleStyle.BackColor1.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.ExpandablePanel1.TitleStyle.BackColor2.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.ExpandablePanel1.TitleStyle.Border = DevComponents.DotNetBar.eBorderType.RaisedInner
        Me.ExpandablePanel1.TitleStyle.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExpandablePanel1.TitleStyle.ForeColor.Color = System.Drawing.Color.White
        Me.ExpandablePanel1.TitleStyle.GradientAngle = 90
        Me.ExpandablePanel1.TitleText = "MANAGEMENT SERVER"
        '
        'btnFB
        '
        Me.btnFB.BackColor = System.Drawing.Color.White
        Me.btnFB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnFB.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFB.ForeColor = System.Drawing.Color.Black
        Me.btnFB.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnFB.Location = New System.Drawing.Point(3, 281)
        Me.btnFB.Name = "btnFB"
        Me.btnFB.Size = New System.Drawing.Size(225, 58)
        Me.btnFB.TabIndex = 5
        Me.btnFB.Text = "Feedback Viewer"
        Me.btnFB.UseVisualStyleBackColor = False
        '
        'btnReports
        '
        Me.btnReports.BackColor = System.Drawing.Color.White
        Me.btnReports.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnReports.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnReports.ForeColor = System.Drawing.Color.Black
        Me.btnReports.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnReports.Location = New System.Drawing.Point(3, 217)
        Me.btnReports.Name = "btnReports"
        Me.btnReports.Size = New System.Drawing.Size(225, 58)
        Me.btnReports.TabIndex = 4
        Me.btnReports.Text = "Reports"
        Me.btnReports.UseVisualStyleBackColor = False
        '
        'btnMonitoring
        '
        Me.btnMonitoring.BackColor = System.Drawing.Color.White
        Me.btnMonitoring.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnMonitoring.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMonitoring.ForeColor = System.Drawing.Color.Black
        Me.btnMonitoring.Location = New System.Drawing.Point(3, 153)
        Me.btnMonitoring.Name = "btnMonitoring"
        Me.btnMonitoring.Size = New System.Drawing.Size(225, 58)
        Me.btnMonitoring.TabIndex = 3
        Me.btnMonitoring.Text = "SET Monitoring"
        Me.btnMonitoring.UseVisualStyleBackColor = False
        '
        'btnKiosk
        '
        Me.btnKiosk.BackColor = System.Drawing.Color.White
        Me.btnKiosk.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnKiosk.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnKiosk.ForeColor = System.Drawing.Color.Black
        Me.btnKiosk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnKiosk.Location = New System.Drawing.Point(3, 89)
        Me.btnKiosk.Name = "btnKiosk"
        Me.btnKiosk.Size = New System.Drawing.Size(225, 58)
        Me.btnKiosk.TabIndex = 2
        Me.btnKiosk.Text = "SET Management"
        Me.btnKiosk.UseVisualStyleBackColor = False
        '
        'btnDash
        '
        Me.btnDash.BackColor = System.Drawing.Color.White
        Me.btnDash.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDash.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDash.ForeColor = System.Drawing.Color.Black
        Me.btnDash.Location = New System.Drawing.Point(3, 26)
        Me.btnDash.Name = "btnDash"
        Me.btnDash.Size = New System.Drawing.Size(225, 57)
        Me.btnDash.TabIndex = 1
        Me.btnDash.Text = "SET Summary"
        Me.btnDash.UseVisualStyleBackColor = False
        '
        'Timer2
        '
        '
        '_frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(51, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1284, 702)
        Me.Controls.Add(Me.Panel1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "_frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Welcome"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.exPan2.ResumeLayout(False)
        Me.exPan3.ResumeLayout(False)
        Me.ExpandablePanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents exPan2 As DevComponents.DotNetBar.ExpandablePanel
    Friend WithEvents btnUserRole As System.Windows.Forms.Button
    Friend WithEvents btnUser As System.Windows.Forms.Button
    Friend WithEvents exPan3 As DevComponents.DotNetBar.ExpandablePanel
    Friend WithEvents btnDatabaseSet As System.Windows.Forms.Button
    Friend WithEvents btnLogout As System.Windows.Forms.Button
    Friend WithEvents ExpandablePanel1 As DevComponents.DotNetBar.ExpandablePanel
    Friend WithEvents btnFB As System.Windows.Forms.Button
    Friend WithEvents btnReports As System.Windows.Forms.Button
    Friend WithEvents btnMonitoring As System.Windows.Forms.Button
    Friend WithEvents btnKiosk As System.Windows.Forms.Button
    Friend WithEvents btnDash As System.Windows.Forms.Button
End Class
