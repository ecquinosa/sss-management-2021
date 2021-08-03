<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmKiosk
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
        Me.txtIpAdd = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.ksnum = New System.Windows.Forms.Label()
        Me.kbranch = New System.Windows.Forms.Label()
        Me.kname = New System.Windows.Forms.Label()
        Me.txtKSerNum = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.txtKName = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.txtmgmtID = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.ButtonX2 = New DevComponents.DotNetBar.ButtonX()
        Me.btnCancel = New DevComponents.DotNetBar.ButtonX()
        Me.btnAdd = New DevComponents.DotNetBar.ButtonX()
        Me.btnEdit = New DevComponents.DotNetBar.ButtonX()
        Me.btnClear = New DevComponents.DotNetBar.ButtonX()
        Me.btnSave = New DevComponents.DotNetBar.ButtonX()
        Me.GroupPanel1 = New DevComponents.DotNetBar.Controls.GroupPanel()
        Me.btnSearch = New DevComponents.DotNetBar.ButtonX()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.cbFields = New System.Windows.Forms.ComboBox()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lanvip = New System.Windows.Forms.Label()
        Me.rbLAN = New System.Windows.Forms.RadioButton()
        Me.rbVPN = New System.Windows.Forms.RadioButton()
        Me.ButtonX1 = New DevComponents.DotNetBar.ButtonX()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.kgroup = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.cbGroup = New System.Windows.Forms.ComboBox()
        Me.kip = New System.Windows.Forms.Label()
        Me.kcluster = New System.Windows.Forms.Label()
        Me.cbBranch = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cbCluster = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtStats = New System.Windows.Forms.ComboBox()
        Me.eHrs = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cbstart1 = New System.Windows.Forms.ComboBox()
        Me.cbstart = New System.Windows.Forms.ComboBox()
        Me.ohrs = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.stats = New System.Windows.Forms.Label()
        Me.dgvKiosk = New System.Windows.Forms.DataGridView()
        Me.Panel3.SuspendLayout()
        Me.GroupPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvKiosk, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtIpAdd
        '
        Me.txtIpAdd.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.txtIpAdd.Border.Class = "TextBoxBorder"
        Me.txtIpAdd.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.txtIpAdd.DisabledBackColor = System.Drawing.Color.White
        Me.txtIpAdd.Enabled = False
        Me.txtIpAdd.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIpAdd.ForeColor = System.Drawing.Color.Black
        Me.txtIpAdd.Location = New System.Drawing.Point(119, 81)
        Me.txtIpAdd.Name = "txtIpAdd"
        Me.txtIpAdd.Size = New System.Drawing.Size(310, 29)
        Me.txtIpAdd.TabIndex = 4
        '
        'ksnum
        '
        Me.ksnum.AutoSize = True
        Me.ksnum.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ksnum.ForeColor = System.Drawing.Color.Red
        Me.ksnum.Location = New System.Drawing.Point(432, 142)
        Me.ksnum.Name = "ksnum"
        Me.ksnum.Size = New System.Drawing.Size(20, 25)
        Me.ksnum.TabIndex = 32
        Me.ksnum.Text = "*"
        Me.ksnum.Visible = False
        '
        'kbranch
        '
        Me.kbranch.AutoSize = True
        Me.kbranch.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kbranch.ForeColor = System.Drawing.Color.Red
        Me.kbranch.Location = New System.Drawing.Point(467, 225)
        Me.kbranch.Name = "kbranch"
        Me.kbranch.Size = New System.Drawing.Size(20, 25)
        Me.kbranch.TabIndex = 31
        Me.kbranch.Text = "*"
        Me.kbranch.Visible = False
        '
        'kname
        '
        Me.kname.AutoSize = True
        Me.kname.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kname.ForeColor = System.Drawing.Color.Red
        Me.kname.Location = New System.Drawing.Point(434, 114)
        Me.kname.Name = "kname"
        Me.kname.Size = New System.Drawing.Size(20, 25)
        Me.kname.TabIndex = 30
        Me.kname.Text = "*"
        Me.kname.Visible = False
        '
        'txtKSerNum
        '
        Me.txtKSerNum.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.txtKSerNum.Border.Class = "TextBoxBorder"
        Me.txtKSerNum.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.txtKSerNum.DisabledBackColor = System.Drawing.Color.White
        Me.txtKSerNum.Enabled = False
        Me.txtKSerNum.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtKSerNum.ForeColor = System.Drawing.Color.Black
        Me.txtKSerNum.Location = New System.Drawing.Point(119, 137)
        Me.txtKSerNum.Name = "txtKSerNum"
        Me.txtKSerNum.Size = New System.Drawing.Size(310, 29)
        Me.txtKSerNum.TabIndex = 23
        '
        'txtKName
        '
        Me.txtKName.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.txtKName.Border.Class = "TextBoxBorder"
        Me.txtKName.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.txtKName.DisabledBackColor = System.Drawing.Color.White
        Me.txtKName.Enabled = False
        Me.txtKName.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtKName.ForeColor = System.Drawing.Color.Black
        Me.txtKName.Location = New System.Drawing.Point(119, 109)
        Me.txtKName.Name = "txtKName"
        Me.txtKName.Size = New System.Drawing.Size(310, 29)
        Me.txtKName.TabIndex = 21
        '
        'txtmgmtID
        '
        Me.txtmgmtID.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.txtmgmtID.Border.Class = "TextBoxBorder"
        Me.txtmgmtID.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.txtmgmtID.DisabledBackColor = System.Drawing.Color.White
        Me.txtmgmtID.Enabled = False
        Me.txtmgmtID.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmgmtID.ForeColor = System.Drawing.Color.Black
        Me.txtmgmtID.Location = New System.Drawing.Point(119, 53)
        Me.txtmgmtID.MaxLength = 4
        Me.txtmgmtID.Name = "txtmgmtID"
        Me.txtmgmtID.Size = New System.Drawing.Size(134, 29)
        Me.txtmgmtID.TabIndex = 20
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(11, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 23)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Serial Number :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(57, 226)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 23)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Branch :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(30, 115)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(109, 23)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Host Name :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(52, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 23)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Host ID :"
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BackColor = System.Drawing.Color.Transparent
        Me.Panel3.Controls.Add(Me.ButtonX2)
        Me.Panel3.Controls.Add(Me.btnCancel)
        Me.Panel3.Controls.Add(Me.btnAdd)
        Me.Panel3.Controls.Add(Me.btnEdit)
        Me.Panel3.Controls.Add(Me.btnClear)
        Me.Panel3.Controls.Add(Me.btnSave)
        Me.Panel3.ForeColor = System.Drawing.Color.Black
        Me.Panel3.Location = New System.Drawing.Point(12, 660)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1005, 49)
        Me.Panel3.TabIndex = 126
        '
        'ButtonX2
        '
        Me.ButtonX2.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.ButtonX2.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ButtonX2.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.ButtonX2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonX2.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonX2.Location = New System.Drawing.Point(878, 6)
        Me.ButtonX2.Name = "ButtonX2"
        Me.ButtonX2.Shape = New DevComponents.DotNetBar.RoundRectangleShapeDescriptor(6)
        Me.ButtonX2.Size = New System.Drawing.Size(124, 37)
        Me.ButtonX2.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2010
        Me.ButtonX2.TabIndex = 6
        Me.ButtonX2.Text = "EXPORT TO EXCEL"
        '
        'btnCancel
        '
        Me.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.btnCancel.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCancel.Enabled = False
        Me.btnCancel.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(706, 6)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Shape = New DevComponents.DotNetBar.RoundRectangleShapeDescriptor(6)
        Me.btnCancel.Size = New System.Drawing.Size(124, 37)
        Me.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2010
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "CANCEL"
        '
        'btnAdd
        '
        Me.btnAdd.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.btnAdd.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btnAdd.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.btnAdd.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAdd.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdd.Location = New System.Drawing.Point(173, 6)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Shape = New DevComponents.DotNetBar.RoundRectangleShapeDescriptor(6)
        Me.btnAdd.Size = New System.Drawing.Size(124, 38)
        Me.btnAdd.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2010
        Me.btnAdd.TabIndex = 3
        Me.btnAdd.Text = "ADD SSIT"
        '
        'btnEdit
        '
        Me.btnEdit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.btnEdit.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btnEdit.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.btnEdit.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnEdit.Enabled = False
        Me.btnEdit.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEdit.Location = New System.Drawing.Point(307, 6)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Shape = New DevComponents.DotNetBar.RoundRectangleShapeDescriptor(6)
        Me.btnEdit.Size = New System.Drawing.Size(124, 38)
        Me.btnEdit.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2010
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "EDIT SSIT"
        '
        'btnClear
        '
        Me.btnClear.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.btnClear.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btnClear.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.btnClear.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnClear.Enabled = False
        Me.btnClear.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Location = New System.Drawing.Point(573, 6)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Shape = New DevComponents.DotNetBar.RoundRectangleShapeDescriptor(6)
        Me.btnClear.Size = New System.Drawing.Size(124, 37)
        Me.btnClear.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2010
        Me.btnClear.TabIndex = 1
        Me.btnClear.Text = "CLEAR"
        '
        'btnSave
        '
        Me.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.btnSave.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSave.Enabled = False
        Me.btnSave.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(441, 7)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Shape = New DevComponents.DotNetBar.RoundRectangleShapeDescriptor(6)
        Me.btnSave.Size = New System.Drawing.Size(123, 37)
        Me.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2010
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "SAVE SSIT"
        '
        'GroupPanel1
        '
        Me.GroupPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupPanel1.AutoSize = True
        Me.GroupPanel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(216, Byte), Integer), CType(CType(216, Byte), Integer), CType(CType(216, Byte), Integer))
        Me.GroupPanel1.CanvasColor = System.Drawing.SystemColors.Control
        Me.GroupPanel1.Controls.Add(Me.btnSearch)
        Me.GroupPanel1.Controls.Add(Me.Label8)
        Me.GroupPanel1.Controls.Add(Me.Label15)
        Me.GroupPanel1.Controls.Add(Me.cbFields)
        Me.GroupPanel1.Controls.Add(Me.txtSearch)
        Me.GroupPanel1.DisabledBackColor = System.Drawing.Color.Empty
        Me.GroupPanel1.Location = New System.Drawing.Point(1, 3)
        Me.GroupPanel1.Name = "GroupPanel1"
        Me.GroupPanel1.Size = New System.Drawing.Size(1019, 59)
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
        Me.GroupPanel1.Style.CornerDiameter = 4
        Me.GroupPanel1.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded
        Me.GroupPanel1.Style.TextAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Center
        Me.GroupPanel1.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText
        Me.GroupPanel1.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near
        '
        '
        '
        Me.GroupPanel1.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.GroupPanel1.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.GroupPanel1.TabIndex = 127
        Me.GroupPanel1.Text = "Quick Search"
        Me.GroupPanel1.TitleImagePosition = DevComponents.DotNetBar.eTitleImagePosition.Center
        '
        'btnSearch
        '
        Me.btnSearch.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.btnSearch.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btnSearch.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.btnSearch.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearch.Location = New System.Drawing.Point(839, 3)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(135, 28)
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
        Me.Label8.Location = New System.Drawing.Point(33, 9)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(93, 23)
        Me.Label8.TabIndex = 68
        Me.Label8.Text = "Search By :"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(396, 9)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(99, 23)
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
        Me.cbFields.Items.AddRange(New Object() {"HOST ID", "HOST IP ADDRESS", "HOST SERIAL NUMBER", "HOST NAME", "BRANCH", "DIVISION", "GROUP", "Show all"})
        Me.cbFields.Location = New System.Drawing.Point(107, 5)
        Me.cbFields.Name = "cbFields"
        Me.cbFields.Size = New System.Drawing.Size(253, 29)
        Me.cbFields.TabIndex = 120
        '
        'txtSearch
        '
        Me.txtSearch.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearch.Location = New System.Drawing.Point(473, 5)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(309, 29)
        Me.txtSearch.TabIndex = 121
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.GroupBox1.AutoSize = True
        Me.GroupBox1.Controls.Add(Me.lanvip)
        Me.GroupBox1.Controls.Add(Me.rbLAN)
        Me.GroupBox1.Controls.Add(Me.rbVPN)
        Me.GroupBox1.Controls.Add(Me.ButtonX1)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.kgroup)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.cbGroup)
        Me.GroupBox1.Controls.Add(Me.kip)
        Me.GroupBox1.Controls.Add(Me.kcluster)
        Me.GroupBox1.Controls.Add(Me.txtIpAdd)
        Me.GroupBox1.Controls.Add(Me.cbBranch)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.txtKSerNum)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ksnum)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cbCluster)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.kbranch)
        Me.GroupBox1.Controls.Add(Me.txtmgmtID)
        Me.GroupBox1.Controls.Add(Me.kname)
        Me.GroupBox1.Controls.Add(Me.txtKName)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox1.Location = New System.Drawing.Point(1, 347)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(493, 291)
        Me.GroupBox1.TabIndex = 128
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "SET Details"
        '
        'lanvip
        '
        Me.lanvip.AutoSize = True
        Me.lanvip.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lanvip.ForeColor = System.Drawing.Color.Red
        Me.lanvip.Location = New System.Drawing.Point(397, 53)
        Me.lanvip.Name = "lanvip"
        Me.lanvip.Size = New System.Drawing.Size(20, 25)
        Me.lanvip.TabIndex = 47
        Me.lanvip.Text = "*"
        Me.lanvip.Visible = False
        '
        'rbLAN
        '
        Me.rbLAN.AutoSize = True
        Me.rbLAN.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbLAN.ForeColor = System.Drawing.Color.Black
        Me.rbLAN.Location = New System.Drawing.Point(280, 53)
        Me.rbLAN.Name = "rbLAN"
        Me.rbLAN.Size = New System.Drawing.Size(65, 27)
        Me.rbLAN.TabIndex = 46
        Me.rbLAN.TabStop = True
        Me.rbLAN.Text = "LAN"
        Me.rbLAN.UseVisualStyleBackColor = True
        '
        'rbVPN
        '
        Me.rbVPN.AutoSize = True
        Me.rbVPN.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbVPN.ForeColor = System.Drawing.Color.Black
        Me.rbVPN.Location = New System.Drawing.Point(338, 53)
        Me.rbVPN.Name = "rbVPN"
        Me.rbVPN.Size = New System.Drawing.Size(65, 27)
        Me.rbVPN.TabIndex = 45
        Me.rbVPN.TabStop = True
        Me.rbVPN.Text = "VPN"
        Me.rbVPN.UseVisualStyleBackColor = True
        '
        'ButtonX1
        '
        Me.ButtonX1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.ButtonX1.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ButtonX1.ColorTable = DevComponents.DotNetBar.eButtonColor.Office2007WithBackground
        Me.ButtonX1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonX1.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonX1.Location = New System.Drawing.Point(431, 223)
        Me.ButtonX1.Name = "ButtonX1"
        Me.ButtonX1.Shape = New DevComponents.DotNetBar.RoundRectangleShapeDescriptor(6)
        Me.ButtonX1.Size = New System.Drawing.Size(23, 20)
        Me.ButtonX1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2010
        Me.ButtonX1.TabIndex = 44
        Me.ButtonX1.Text = "+"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(34, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(104, 23)
        Me.Label5.TabIndex = 43
        Me.Label5.Text = "IP Address :"
        '
        'kgroup
        '
        Me.kgroup.AutoSize = True
        Me.kgroup.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kgroup.ForeColor = System.Drawing.Color.Red
        Me.kgroup.Location = New System.Drawing.Point(434, 171)
        Me.kgroup.Name = "kgroup"
        Me.kgroup.Size = New System.Drawing.Size(20, 25)
        Me.kgroup.TabIndex = 42
        Me.kgroup.Text = "*"
        Me.kgroup.Visible = False
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Black
        Me.Label17.Location = New System.Drawing.Point(61, 171)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(67, 23)
        Me.Label17.TabIndex = 41
        Me.Label17.Text = "Group :"
        '
        'cbGroup
        '
        Me.cbGroup.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cbGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbGroup.Enabled = False
        Me.cbGroup.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbGroup.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbGroup.ForeColor = System.Drawing.SystemColors.InfoText
        Me.cbGroup.FormattingEnabled = True
        Me.cbGroup.Location = New System.Drawing.Point(119, 166)
        Me.cbGroup.Name = "cbGroup"
        Me.cbGroup.Size = New System.Drawing.Size(310, 29)
        Me.cbGroup.TabIndex = 40
        '
        'kip
        '
        Me.kip.AutoSize = True
        Me.kip.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kip.ForeColor = System.Drawing.Color.Red
        Me.kip.Location = New System.Drawing.Point(434, 86)
        Me.kip.Name = "kip"
        Me.kip.Size = New System.Drawing.Size(20, 25)
        Me.kip.TabIndex = 39
        Me.kip.Text = "*"
        Me.kip.Visible = False
        '
        'kcluster
        '
        Me.kcluster.AutoSize = True
        Me.kcluster.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kcluster.ForeColor = System.Drawing.Color.Red
        Me.kcluster.Location = New System.Drawing.Point(434, 197)
        Me.kcluster.Name = "kcluster"
        Me.kcluster.Size = New System.Drawing.Size(20, 25)
        Me.kcluster.TabIndex = 39
        Me.kcluster.Text = "*"
        Me.kcluster.Visible = False
        '
        'cbBranch
        '
        Me.cbBranch.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cbBranch.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbBranch.Enabled = False
        Me.cbBranch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbBranch.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbBranch.ForeColor = System.Drawing.SystemColors.InfoText
        Me.cbBranch.FormattingEnabled = True
        Me.cbBranch.Location = New System.Drawing.Point(119, 221)
        Me.cbBranch.Name = "cbBranch"
        Me.cbBranch.Size = New System.Drawing.Size(310, 29)
        Me.cbBranch.TabIndex = 36
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Black
        Me.Label16.Location = New System.Drawing.Point(48, 200)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(78, 23)
        Me.Label16.TabIndex = 38
        Me.Label16.Text = "Division :"
        '
        'cbCluster
        '
        Me.cbCluster.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cbCluster.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbCluster.Enabled = False
        Me.cbCluster.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbCluster.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbCluster.ForeColor = System.Drawing.SystemColors.InfoText
        Me.cbCluster.FormattingEnabled = True
        Me.cbCluster.Location = New System.Drawing.Point(119, 193)
        Me.cbCluster.Name = "cbCluster"
        Me.cbCluster.Size = New System.Drawing.Size(310, 29)
        Me.cbCluster.TabIndex = 37
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.GroupBox2.AutoSize = True
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.txtStats)
        Me.GroupBox2.Controls.Add(Me.eHrs)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.cbstart1)
        Me.GroupBox2.Controls.Add(Me.cbstart)
        Me.GroupBox2.Controls.Add(Me.ohrs)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.stats)
        Me.GroupBox2.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox2.Location = New System.Drawing.Point(500, 357)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(520, 281)
        Me.GroupBox2.TabIndex = 129
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Operation Details"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(117, 152)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(65, 23)
        Me.Label9.TabIndex = 55
        Me.Label9.Text = "Status :"
        '
        'txtStats
        '
        Me.txtStats.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtStats.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.txtStats.Enabled = False
        Me.txtStats.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.txtStats.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStats.ForeColor = System.Drawing.SystemColors.InfoText
        Me.txtStats.FormattingEnabled = True
        Me.txtStats.Items.AddRange(New Object() {"INACTIVE", "ACTIVE", "DEACTIVE"})
        Me.txtStats.Location = New System.Drawing.Point(175, 147)
        Me.txtStats.Name = "txtStats"
        Me.txtStats.Size = New System.Drawing.Size(156, 29)
        Me.txtStats.TabIndex = 54
        '
        'eHrs
        '
        Me.eHrs.AutoSize = True
        Me.eHrs.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eHrs.ForeColor = System.Drawing.Color.Red
        Me.eHrs.Location = New System.Drawing.Point(437, 119)
        Me.eHrs.Name = "eHrs"
        Me.eHrs.Size = New System.Drawing.Size(20, 25)
        Me.eHrs.TabIndex = 53
        Me.eHrs.Text = "*"
        Me.eHrs.Visible = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(301, 122)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(36, 23)
        Me.Label12.TabIndex = 52
        Me.Label12.Text = "To :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(123, 122)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(58, 23)
        Me.Label6.TabIndex = 51
        Me.Label6.Text = "From :"
        '
        'cbstart1
        '
        Me.cbstart1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cbstart1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbstart1.Enabled = False
        Me.cbstart1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbstart1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbstart1.ForeColor = System.Drawing.SystemColors.InfoText
        Me.cbstart1.FormattingEnabled = True
        Me.cbstart1.Items.AddRange(New Object() {"1:00:00 AM", "1:30:00 AM", "2:00:00 AM", "2:30:00 AM", "3:00:00 AM", "3:30:00 AM", "4:00:00 AM", "4:30:00 AM", "5:00:00 AM", "5:30:00 AM", "6:00:00 AM", "6:30:00 AM", "7:00:00 AM", "7:30:00 AM", "8:00:00 AM", "8:30:00 AM", "9:00:00 AM", "9:30:00 AM", "10:00:00 AM", "10:30:00 AM", "11:00:00 AM", "11:30:00 AM", "12:00:00 AM", "12:30:00 AM", "1:00:00 PM", "1:30:00 PM", "2:00:00 PM", "2:30:00 PM", "3:00:00 PM", "3:30:00 PM", "4:00:00 AM", "4:30:00 PM", "5:00:00 PM", "5:30:00 PM", "6:00:00 PM", "6:30:00 PM", "7:00:00 PM", "7:30:00 PM", "8:00:00 PM", "8:30:00 PM", "9:00:00 PM", "9:30:00 PM", "10:00:00 PM", "10:30:00 PM", "11:00:00 PM", "11:30:00 PM", "12:00:00 PM", "12:30:00 PM"})
        Me.cbstart1.Location = New System.Drawing.Point(335, 118)
        Me.cbstart1.Name = "cbstart1"
        Me.cbstart1.Size = New System.Drawing.Size(101, 29)
        Me.cbstart1.TabIndex = 48
        '
        'cbstart
        '
        Me.cbstart.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cbstart.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbstart.Enabled = False
        Me.cbstart.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbstart.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbstart.ForeColor = System.Drawing.SystemColors.InfoText
        Me.cbstart.FormattingEnabled = True
        Me.cbstart.Items.AddRange(New Object() {"1:00:00 AM", "1:30:00 AM", "2:00:00 AM", "2:30:00 AM", "3:00:00 AM", "3:30:00 AM", "4:00:00 AM", "4:30:00 AM", "5:00:00 AM", "5:30:00 AM", "6:00:00 AM", "6:30:00 AM", "7:00:00 AM", "7:30:00 AM", "8:00:00 AM", "8:30:00 AM", "9:00:00 AM", "9:30:00 AM", "10:00:00 AM", "10:30:00 AM", "11:00:00 AM", "11:30:00 AM", "12:00:00 AM", "12:30:00 AM", "1:00:00 PM", "1:30:00 PM", "2:00:00 PM", "2:30:00 PM", "3:00:00 PM", "3:30:00 PM", "4:00:00 AM", "4:30:00 PM", "5:00:00 PM", "5:30:00 PM", "6:00:00 PM", "6:30:00 PM", "7:00:00 PM", "7:30:00 PM", "8:00:00 PM", "8:30:00 PM", "9:00:00 PM", "9:30:00 PM", "10:00:00 PM", "10:30:00 PM", "11:00:00 PM", "11:30:00 PM", "12:00:00 PM", "12:30:00 PM"})
        Me.cbstart.Location = New System.Drawing.Point(174, 118)
        Me.cbstart.Name = "cbstart"
        Me.cbstart.Size = New System.Drawing.Size(105, 29)
        Me.cbstart.TabIndex = 46
        '
        'ohrs
        '
        Me.ohrs.AutoSize = True
        Me.ohrs.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ohrs.ForeColor = System.Drawing.Color.Red
        Me.ohrs.Location = New System.Drawing.Point(282, 119)
        Me.ohrs.Name = "ohrs"
        Me.ohrs.Size = New System.Drawing.Size(20, 25)
        Me.ohrs.TabIndex = 45
        Me.ohrs.Text = "*"
        Me.ohrs.Visible = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(54, 93)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(142, 23)
        Me.Label11.TabIndex = 43
        Me.Label11.Text = "Operation Hours "
        '
        'stats
        '
        Me.stats.AutoSize = True
        Me.stats.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.stats.ForeColor = System.Drawing.Color.Red
        Me.stats.Location = New System.Drawing.Point(331, 149)
        Me.stats.Name = "stats"
        Me.stats.Size = New System.Drawing.Size(20, 25)
        Me.stats.TabIndex = 42
        Me.stats.Text = "*"
        Me.stats.Visible = False
        '
        'dgvKiosk
        '
        Me.dgvKiosk.AllowUserToAddRows = False
        Me.dgvKiosk.AllowUserToDeleteRows = False
        Me.dgvKiosk.AllowUserToResizeColumns = False
        Me.dgvKiosk.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.White
        Me.dgvKiosk.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvKiosk.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvKiosk.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvKiosk.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.dgvKiosk.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgvKiosk.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvKiosk.GridColor = System.Drawing.SystemColors.ActiveBorder
        Me.dgvKiosk.Location = New System.Drawing.Point(1, 62)
        Me.dgvKiosk.Name = "dgvKiosk"
        Me.dgvKiosk.ReadOnly = True
        Me.dgvKiosk.RowHeadersWidth = 51
        Me.dgvKiosk.Size = New System.Drawing.Size(1019, 290)
        Me.dgvKiosk.TabIndex = 130
        '
        '_frmKiosk
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1032, 721)
        Me.Controls.Add(Me.dgvKiosk)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupPanel1)
        Me.Controls.Add(Me.Panel3)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "_frmKiosk"
        Me.TopLeftCornerSize = 1
        Me.TopRightCornerSize = 1
        Me.Panel3.ResumeLayout(False)
        Me.GroupPanel1.ResumeLayout(False)
        Me.GroupPanel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvKiosk, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtIpAdd As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents txtKSerNum As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents txtKName As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents txtmgmtID As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnEdit As DevComponents.DotNetBar.ButtonX
    Friend WithEvents btnClear As DevComponents.DotNetBar.ButtonX
    Friend WithEvents btnSave As DevComponents.DotNetBar.ButtonX
    Friend WithEvents GroupPanel1 As DevComponents.DotNetBar.Controls.GroupPanel
    Friend WithEvents btnSearch As DevComponents.DotNetBar.ButtonX
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents cbFields As System.Windows.Forms.ComboBox
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents ksnum As System.Windows.Forms.Label
    Friend WithEvents kbranch As System.Windows.Forms.Label
    Friend WithEvents kname As System.Windows.Forms.Label
    Friend WithEvents btnAdd As DevComponents.DotNetBar.ButtonX
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cbBranch As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cbCluster As System.Windows.Forms.ComboBox
    Friend WithEvents kcluster As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents cbGroup As System.Windows.Forms.ComboBox
    Friend WithEvents kip As System.Windows.Forms.Label
    Friend WithEvents kgroup As System.Windows.Forms.Label
    Friend WithEvents dgvKiosk As System.Windows.Forms.DataGridView
    Friend WithEvents btnCancel As DevComponents.DotNetBar.ButtonX
    Friend WithEvents stats As System.Windows.Forms.Label
    Friend WithEvents ohrs As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cbstart1 As System.Windows.Forms.ComboBox
    Friend WithEvents cbstart As System.Windows.Forms.ComboBox
    Friend WithEvents eHrs As System.Windows.Forms.Label
    Friend WithEvents txtStats As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ButtonX1 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents rbLAN As System.Windows.Forms.RadioButton
    Friend WithEvents rbVPN As System.Windows.Forms.RadioButton
    Friend WithEvents lanvip As System.Windows.Forms.Label
    Friend WithEvents ButtonX2 As DevComponents.DotNetBar.ButtonX
End Class
