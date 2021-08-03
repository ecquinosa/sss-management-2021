<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmFBK
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dgvFB = New System.Windows.Forms.DataGridView()
        Me.gpCS = New System.Windows.Forms.GroupBox()
        Me.lblVisit = New System.Windows.Forms.Label()
        Me.rtbComments = New System.Windows.Forms.RichTextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.rtbWhy = New System.Windows.Forms.RichTextBox()
        Me.gpQ = New System.Windows.Forms.GroupBox()
        Me.lblqq7 = New System.Windows.Forms.Label()
        Me.lblqq1 = New System.Windows.Forms.Label()
        Me.lblqq2 = New System.Windows.Forms.Label()
        Me.lblqq6 = New System.Windows.Forms.Label()
        Me.lblqq3 = New System.Windows.Forms.Label()
        Me.lblqq4 = New System.Windows.Forms.Label()
        Me.lblqq5 = New System.Windows.Forms.Label()
        Me.gpPI = New System.Windows.Forms.GroupBox()
        Me.lblKiosk_ID = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblBranch = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.lbldate = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lbladd = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.lblPcode = New System.Windows.Forms.Label()
        Me.lbladd1 = New System.Windows.Forms.Label()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupPanel1 = New DevComponents.DotNetBar.Controls.GroupPanel()
        Me.ButtonX2 = New DevComponents.DotNetBar.ButtonX()
        Me.ButtonX1 = New DevComponents.DotNetBar.ButtonX()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbFB = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSearch = New DevComponents.DotNetBar.ButtonX()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.cbFields = New System.Windows.Forms.ComboBox()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        CType(Me.dgvFB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gpCS.SuspendLayout()
        Me.gpQ.SuspendLayout()
        Me.gpPI.SuspendLayout()
        Me.GroupPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.dgvFB)
        Me.Panel1.Controls.Add(Me.gpCS)
        Me.Panel1.Controls.Add(Me.gpQ)
        Me.Panel1.Controls.Add(Me.gpPI)
        Me.Panel1.Controls.Add(Me.GroupPanel1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1107, 691)
        Me.Panel1.TabIndex = 0
        '
        'dgvFB
        '
        Me.dgvFB.AllowUserToAddRows = False
        Me.dgvFB.AllowUserToDeleteRows = False
        Me.dgvFB.AllowUserToResizeColumns = False
        Me.dgvFB.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.White
        Me.dgvFB.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvFB.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvFB.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvFB.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.dgvFB.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgvFB.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFB.GridColor = System.Drawing.SystemColors.ActiveBorder
        Me.dgvFB.Location = New System.Drawing.Point(7, 99)
        Me.dgvFB.Name = "dgvFB"
        Me.dgvFB.ReadOnly = True
        Me.dgvFB.Size = New System.Drawing.Size(1088, 278)
        Me.dgvFB.TabIndex = 138
        '
        'gpCS
        '
        Me.gpCS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gpCS.Controls.Add(Me.lblVisit)
        Me.gpCS.Controls.Add(Me.rtbComments)
        Me.gpCS.Controls.Add(Me.Label17)
        Me.gpCS.Controls.Add(Me.Label16)
        Me.gpCS.Controls.Add(Me.rtbWhy)
        Me.gpCS.Location = New System.Drawing.Point(664, 383)
        Me.gpCS.Name = "gpCS"
        Me.gpCS.Size = New System.Drawing.Size(435, 305)
        Me.gpCS.TabIndex = 137
        Me.gpCS.TabStop = False
        Me.gpCS.Text = "Comments and Suggestion"
        '
        'lblVisit
        '
        Me.lblVisit.AutoSize = True
        Me.lblVisit.Location = New System.Drawing.Point(19, 18)
        Me.lblVisit.Name = "lblVisit"
        Me.lblVisit.Size = New System.Drawing.Size(36, 13)
        Me.lblVisit.TabIndex = 9
        Me.lblVisit.Text = "Name"
        '
        'rtbComments
        '
        Me.rtbComments.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.rtbComments.Location = New System.Drawing.Point(6, 193)
        Me.rtbComments.Name = "rtbComments"
        Me.rtbComments.ReadOnly = True
        Me.rtbComments.Size = New System.Drawing.Size(414, 106)
        Me.rtbComments.TabIndex = 8
        Me.rtbComments.Text = ""
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(19, 177)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(129, 13)
        Me.Label17.TabIndex = 7
        Me.Label17.Text = "Comments/Suggestions"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(19, 57)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(38, 13)
        Me.Label16.TabIndex = 6
        Me.Label16.Text = "Why ?"
        '
        'rtbWhy
        '
        Me.rtbWhy.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.rtbWhy.Location = New System.Drawing.Point(6, 73)
        Me.rtbWhy.Name = "rtbWhy"
        Me.rtbWhy.ReadOnly = True
        Me.rtbWhy.Size = New System.Drawing.Size(414, 101)
        Me.rtbWhy.TabIndex = 0
        Me.rtbWhy.Text = ""
        '
        'gpQ
        '
        Me.gpQ.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gpQ.Controls.Add(Me.lblqq7)
        Me.gpQ.Controls.Add(Me.lblqq1)
        Me.gpQ.Controls.Add(Me.lblqq2)
        Me.gpQ.Controls.Add(Me.lblqq6)
        Me.gpQ.Controls.Add(Me.lblqq3)
        Me.gpQ.Controls.Add(Me.lblqq4)
        Me.gpQ.Controls.Add(Me.lblqq5)
        Me.gpQ.Location = New System.Drawing.Point(392, 383)
        Me.gpQ.Name = "gpQ"
        Me.gpQ.Size = New System.Drawing.Size(266, 305)
        Me.gpQ.TabIndex = 136
        Me.gpQ.TabStop = False
        Me.gpQ.Text = "Questions"
        '
        'lblqq7
        '
        Me.lblqq7.AutoSize = True
        Me.lblqq7.Location = New System.Drawing.Point(6, 271)
        Me.lblqq7.Name = "lblqq7"
        Me.lblqq7.Size = New System.Drawing.Size(36, 13)
        Me.lblqq7.TabIndex = 13
        Me.lblqq7.Text = "Name"
        '
        'lblqq1
        '
        Me.lblqq1.AutoSize = True
        Me.lblqq1.Location = New System.Drawing.Point(6, 18)
        Me.lblqq1.Name = "lblqq1"
        Me.lblqq1.Size = New System.Drawing.Size(36, 13)
        Me.lblqq1.TabIndex = 1
        Me.lblqq1.Text = "Name"
        '
        'lblqq2
        '
        Me.lblqq2.AutoSize = True
        Me.lblqq2.Location = New System.Drawing.Point(6, 81)
        Me.lblqq2.Name = "lblqq2"
        Me.lblqq2.Size = New System.Drawing.Size(36, 13)
        Me.lblqq2.TabIndex = 2
        Me.lblqq2.Text = "Name"
        '
        'lblqq6
        '
        Me.lblqq6.AutoSize = True
        Me.lblqq6.Location = New System.Drawing.Point(6, 228)
        Me.lblqq6.Name = "lblqq6"
        Me.lblqq6.Size = New System.Drawing.Size(36, 13)
        Me.lblqq6.TabIndex = 11
        Me.lblqq6.Text = "Name"
        '
        'lblqq3
        '
        Me.lblqq3.AutoSize = True
        Me.lblqq3.Location = New System.Drawing.Point(6, 117)
        Me.lblqq3.Name = "lblqq3"
        Me.lblqq3.Size = New System.Drawing.Size(36, 13)
        Me.lblqq3.TabIndex = 3
        Me.lblqq3.Text = "Name"
        '
        'lblqq4
        '
        Me.lblqq4.AutoSize = True
        Me.lblqq4.Location = New System.Drawing.Point(6, 154)
        Me.lblqq4.Name = "lblqq4"
        Me.lblqq4.Size = New System.Drawing.Size(36, 13)
        Me.lblqq4.TabIndex = 4
        Me.lblqq4.Text = "Name"
        '
        'lblqq5
        '
        Me.lblqq5.AutoSize = True
        Me.lblqq5.Location = New System.Drawing.Point(6, 197)
        Me.lblqq5.Name = "lblqq5"
        Me.lblqq5.Size = New System.Drawing.Size(36, 13)
        Me.lblqq5.TabIndex = 5
        Me.lblqq5.Text = "Name"
        '
        'gpPI
        '
        Me.gpPI.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gpPI.Controls.Add(Me.lblKiosk_ID)
        Me.gpPI.Controls.Add(Me.Label10)
        Me.gpPI.Controls.Add(Me.lblBranch)
        Me.gpPI.Controls.Add(Me.Label11)
        Me.gpPI.Controls.Add(Me.lbldate)
        Me.gpPI.Controls.Add(Me.Label7)
        Me.gpPI.Controls.Add(Me.lbladd)
        Me.gpPI.Controls.Add(Me.Label18)
        Me.gpPI.Controls.Add(Me.lblPcode)
        Me.gpPI.Controls.Add(Me.lbladd1)
        Me.gpPI.Controls.Add(Me.lblEmail)
        Me.gpPI.Controls.Add(Me.lblName)
        Me.gpPI.Controls.Add(Me.Label9)
        Me.gpPI.Controls.Add(Me.Label6)
        Me.gpPI.Controls.Add(Me.Label5)
        Me.gpPI.Controls.Add(Me.Label4)
        Me.gpPI.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gpPI.Location = New System.Drawing.Point(7, 383)
        Me.gpPI.Name = "gpPI"
        Me.gpPI.Size = New System.Drawing.Size(379, 305)
        Me.gpPI.TabIndex = 135
        Me.gpPI.TabStop = False
        Me.gpPI.Text = "Personal Information"
        '
        'lblKiosk_ID
        '
        Me.lblKiosk_ID.AutoSize = True
        Me.lblKiosk_ID.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKiosk_ID.Location = New System.Drawing.Point(284, 46)
        Me.lblKiosk_ID.Name = "lblKiosk_ID"
        Me.lblKiosk_ID.Size = New System.Drawing.Size(20, 15)
        Me.lblKiosk_ID.TabIndex = 17
        Me.lblKiosk_ID.Text = "lbl"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(283, 25)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(90, 21)
        Me.Label10.TabIndex = 16
        Me.Label10.Text = "Terminal ID"
        '
        'lblBranch
        '
        Me.lblBranch.AutoSize = True
        Me.lblBranch.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBranch.Location = New System.Drawing.Point(178, 46)
        Me.lblBranch.Name = "lblBranch"
        Me.lblBranch.Size = New System.Drawing.Size(20, 15)
        Me.lblBranch.TabIndex = 15
        Me.lblBranch.Text = "lbl"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(177, 25)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(58, 21)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "Branch"
        '
        'lbldate
        '
        Me.lbldate.AutoSize = True
        Me.lbldate.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbldate.Location = New System.Drawing.Point(8, 46)
        Me.lbldate.Name = "lbldate"
        Me.lbldate.Size = New System.Drawing.Size(45, 15)
        Me.lbldate.TabIndex = 13
        Me.lbldate.Text = "Name :"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(7, 25)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(118, 21)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "Date Submitted"
        '
        'lbladd
        '
        Me.lbladd.AutoSize = True
        Me.lbladd.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbladd.Location = New System.Drawing.Point(7, 180)
        Me.lbladd.Name = "lbladd"
        Me.lbladd.Size = New System.Drawing.Size(55, 15)
        Me.lbladd.TabIndex = 11
        Me.lbladd.Text = "Address :"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(6, 155)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(103, 21)
        Me.Label18.TabIndex = 10
        Me.Label18.Text = "Address Type"
        '
        'lblPcode
        '
        Me.lblPcode.AutoSize = True
        Me.lblPcode.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPcode.Location = New System.Drawing.Point(6, 280)
        Me.lblPcode.Name = "lblPcode"
        Me.lblPcode.Size = New System.Drawing.Size(45, 15)
        Me.lblPcode.TabIndex = 9
        Me.lblPcode.Text = "Name :"
        '
        'lbladd1
        '
        Me.lbladd1.AutoSize = True
        Me.lbladd1.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbladd1.Location = New System.Drawing.Point(6, 220)
        Me.lbladd1.Name = "lbladd1"
        Me.lbladd1.Size = New System.Drawing.Size(42, 13)
        Me.lbladd1.TabIndex = 7
        Me.lbladd1.Text = "Name :"
        '
        'lblEmail
        '
        Me.lblEmail.AutoSize = True
        Me.lblEmail.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmail.Location = New System.Drawing.Point(9, 132)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(45, 15)
        Me.lblEmail.TabIndex = 6
        Me.lblEmail.Text = "Name :"
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.Location = New System.Drawing.Point(7, 91)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(45, 15)
        Me.lblName.TabIndex = 5
        Me.lblName.Text = "Name :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(6, 260)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(92, 21)
        Me.Label9.TabIndex = 4
        Me.Label9.Text = "Postal Code"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(5, 200)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(66, 21)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Address"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(7, 111)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 21)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Email"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 70)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(52, 21)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Name"
        '
        'GroupPanel1
        '
        Me.GroupPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupPanel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(216, Byte), Integer), CType(CType(216, Byte), Integer), CType(CType(216, Byte), Integer))
        Me.GroupPanel1.CanvasColor = System.Drawing.SystemColors.Control
        Me.GroupPanel1.Controls.Add(Me.ButtonX2)
        Me.GroupPanel1.Controls.Add(Me.ButtonX1)
        Me.GroupPanel1.Controls.Add(Me.dtpTo)
        Me.GroupPanel1.Controls.Add(Me.dtpFrom)
        Me.GroupPanel1.Controls.Add(Me.Label3)
        Me.GroupPanel1.Controls.Add(Me.Label2)
        Me.GroupPanel1.Controls.Add(Me.cbFB)
        Me.GroupPanel1.Controls.Add(Me.Label1)
        Me.GroupPanel1.Controls.Add(Me.btnSearch)
        Me.GroupPanel1.Controls.Add(Me.Label8)
        Me.GroupPanel1.Controls.Add(Me.Label15)
        Me.GroupPanel1.Controls.Add(Me.cbFields)
        Me.GroupPanel1.Controls.Add(Me.txtSearch)
        Me.GroupPanel1.Location = New System.Drawing.Point(7, 10)
        Me.GroupPanel1.Name = "GroupPanel1"
        Me.GroupPanel1.Size = New System.Drawing.Size(1092, 87)
        '
        '
        '
        Me.GroupPanel1.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2
        Me.GroupPanel1.Style.BackColorGradientAngle = 90
        Me.GroupPanel1.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground
        Me.GroupPanel1.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid
        Me.GroupPanel1.Style.BorderBottomWidth = 1
        Me.GroupPanel1.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder
        Me.GroupPanel1.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid
        Me.GroupPanel1.Style.BorderLeftWidth = 1
        Me.GroupPanel1.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid
        Me.GroupPanel1.Style.BorderRightWidth = 1
        Me.GroupPanel1.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid
        Me.GroupPanel1.Style.BorderTopWidth = 1
        Me.GroupPanel1.Style.Class = ""
        Me.GroupPanel1.Style.CornerDiameter = 4
        Me.GroupPanel1.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded
        Me.GroupPanel1.Style.TextAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Center
        Me.GroupPanel1.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText
        Me.GroupPanel1.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near
        '
        '
        '
        Me.GroupPanel1.StyleMouseDown.Class = ""
        Me.GroupPanel1.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.GroupPanel1.StyleMouseOver.Class = ""
        Me.GroupPanel1.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.GroupPanel1.TabIndex = 133
        Me.GroupPanel1.Text = "Quick Search"
        Me.GroupPanel1.TitleImagePosition = DevComponents.DotNetBar.eTitleImagePosition.Center
        '
        'ButtonX2
        '
        Me.ButtonX2.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.ButtonX2.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ButtonX2.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.ButtonX2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonX2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonX2.Location = New System.Drawing.Point(973, 35)
        Me.ButtonX2.Name = "ButtonX2"
        Me.ButtonX2.Size = New System.Drawing.Size(112, 28)
        Me.ButtonX2.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.ButtonX2.TabIndex = 130
        Me.ButtonX2.Text = "Export"
        '
        'ButtonX1
        '
        Me.ButtonX1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.ButtonX1.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ButtonX1.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.ButtonX1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonX1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonX1.Location = New System.Drawing.Point(855, 35)
        Me.ButtonX1.Name = "ButtonX1"
        Me.ButtonX1.Size = New System.Drawing.Size(112, 28)
        Me.ButtonX1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.ButtonX1.TabIndex = 129
        Me.ButtonX1.Text = "Print"
        '
        'dtpTo
        '
        Me.dtpTo.Location = New System.Drawing.Point(590, 35)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(200, 22)
        Me.dtpTo.TabIndex = 128
        '
        'dtpFrom
        '
        Me.dtpFrom.Location = New System.Drawing.Point(317, 35)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(200, 22)
        Me.dtpFrom.TabIndex = 127
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(523, 38)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 17)
        Me.Label3.TabIndex = 126
        Me.Label3.Text = "Date To :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(235, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 17)
        Me.Label2.TabIndex = 125
        Me.Label2.Text = "Date From :"
        '
        'cbFB
        '
        Me.cbFB.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cbFB.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbFB.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbFB.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbFB.ForeColor = System.Drawing.SystemColors.InfoText
        Me.cbFB.FormattingEnabled = True
        Me.cbFB.Items.AddRange(New Object() {"SSIT FEEDBACK", "WEBSITE FEEDBACK"})
        Me.cbFB.Location = New System.Drawing.Point(395, 3)
        Me.cbFB.Name = "cbFB"
        Me.cbFB.Size = New System.Drawing.Size(238, 25)
        Me.cbFB.TabIndex = 124
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(329, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 17)
        Me.Label1.TabIndex = 123
        Me.Label1.Text = "Filter By :"
        '
        'btnSearch
        '
        Me.btnSearch.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.btnSearch.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btnSearch.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.btnSearch.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSearch.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearch.Location = New System.Drawing.Point(973, 3)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(112, 28)
        Me.btnSearch.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.btnSearch.TabIndex = 122
        Me.btnSearch.Text = "Search"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(8, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(71, 17)
        Me.Label8.TabIndex = 68
        Me.Label8.Text = "Search By :"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(651, 6)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(75, 17)
        Me.Label15.TabIndex = 69
        Me.Label15.Text = "Search text:"
        '
        'cbFields
        '
        Me.cbFields.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cbFields.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbFields.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbFields.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbFields.ForeColor = System.Drawing.SystemColors.InfoText
        Me.cbFields.FormattingEnabled = True
        Me.cbFields.Location = New System.Drawing.Point(85, 3)
        Me.cbFields.Name = "cbFields"
        Me.cbFields.Size = New System.Drawing.Size(238, 25)
        Me.cbFields.TabIndex = 120
        '
        'txtSearch
        '
        Me.txtSearch.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearch.Location = New System.Drawing.Point(732, 3)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(235, 25)
        Me.txtSearch.TabIndex = 121
        '
        '_frmFBK
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1107, 691)
        Me.Controls.Add(Me.Panel1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "_frmFBK"
        Me.Text = "MetroForm"
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgvFB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gpCS.ResumeLayout(False)
        Me.gpCS.PerformLayout()
        Me.gpQ.ResumeLayout(False)
        Me.gpQ.PerformLayout()
        Me.gpPI.ResumeLayout(False)
        Me.gpPI.PerformLayout()
        Me.GroupPanel1.ResumeLayout(False)
        Me.GroupPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents gpCS As System.Windows.Forms.GroupBox
    Friend WithEvents lblVisit As System.Windows.Forms.Label
    Friend WithEvents rtbComments As System.Windows.Forms.RichTextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents rtbWhy As System.Windows.Forms.RichTextBox
    Friend WithEvents gpQ As System.Windows.Forms.GroupBox
    Friend WithEvents lblqq7 As System.Windows.Forms.Label
    Friend WithEvents lblqq6 As System.Windows.Forms.Label
    Friend WithEvents lblqq5 As System.Windows.Forms.Label
    Friend WithEvents lblqq4 As System.Windows.Forms.Label
    Friend WithEvents lblqq3 As System.Windows.Forms.Label
    Friend WithEvents lblqq2 As System.Windows.Forms.Label
    Friend WithEvents lblqq1 As System.Windows.Forms.Label
    Friend WithEvents gpPI As System.Windows.Forms.GroupBox
    Friend WithEvents lbladd As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents lblPcode As System.Windows.Forms.Label
    Friend WithEvents lbladd1 As System.Windows.Forms.Label
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupPanel1 As DevComponents.DotNetBar.Controls.GroupPanel
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbFB As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSearch As DevComponents.DotNetBar.ButtonX
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents cbFields As System.Windows.Forms.ComboBox
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents lbldate As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblBranch As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblKiosk_ID As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents ButtonX1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents dgvFB As System.Windows.Forms.DataGridView
    Friend WithEvents ButtonX2 As DevComponents.DotNetBar.ButtonX
End Class
