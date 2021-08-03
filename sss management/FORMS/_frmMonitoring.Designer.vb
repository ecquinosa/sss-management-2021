<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmMonitoring
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
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.lvPC = New DevComponents.DotNetBar.Controls.ListViewEx()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lvList = New DevComponents.DotNetBar.Controls.ListViewEx()
        Me.gbKioskHistory = New System.Windows.Forms.GroupBox()
        Me.dgvLogs = New System.Windows.Forms.DataGridView()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
        Me.lblFilter = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.OpEnd1 = New System.Windows.Forms.Label()
        Me.OpStart1 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.lblSerial = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.lbldate = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblOD = New System.Windows.Forms.Label()
        Me.DURATION = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.lblTrans = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.LASTONLINE = New System.Windows.Forms.Label()
        Me.lblGroup = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblCluster = New System.Windows.Forms.Label()
        Me.lblLastOnline = New System.Windows.Forms.Label()
        Me.lblKbranch = New System.Windows.Forms.Label()
        Me.lblOffline = New System.Windows.Forms.Label()
        Me.lblKIP = New System.Windows.Forms.Label()
        Me.LASTOFFLINE = New System.Windows.Forms.Label()
        Me.lblKname = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.gbKioskHistory.SuspendLayout()
        CType(Me.dgvLogs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 5000
        '
        'Panel1
        '
        Me.Panel1.AutoSize = True
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1074, 721)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.AutoSize = True
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1074, 721)
        Me.Panel2.TabIndex = 0
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.Button1)
        Me.Panel3.Controls.Add(Me.lvPC)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.Label10)
        Me.Panel3.Controls.Add(Me.lvList)
        Me.Panel3.Controls.Add(Me.gbKioskHistory)
        Me.Panel3.Controls.Add(Me.GroupBox1)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1074, 721)
        Me.Panel3.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(2, 1)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(143, 39)
        Me.Button1.TabIndex = 51
        Me.Button1.Text = "Scan"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'lvPC
        '
        Me.lvPC.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.lvPC.Border.Class = "ListViewBorder"
        Me.lvPC.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.lvPC.DisabledBackColor = System.Drawing.Color.Empty
        Me.lvPC.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvPC.ForeColor = System.Drawing.Color.Red
        Me.lvPC.FullRowSelect = True
        Me.lvPC.GridLines = True
        Me.lvPC.HideSelection = False
        Me.lvPC.Location = New System.Drawing.Point(521, 42)
        Me.lvPC.Name = "lvPC"
        Me.lvPC.Size = New System.Drawing.Size(552, 676)
        Me.lvPC.TabIndex = 47
        Me.lvPC.UseCompatibleStateImageBehavior = False
        Me.lvPC.View = System.Windows.Forms.View.Details
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Green
        Me.Label7.Location = New System.Drawing.Point(207, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(138, 46)
        Me.Label7.TabIndex = 50
        Me.Label7.Text = "ONLINE"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Font = New System.Drawing.Font("Segoe UI", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Red
        Me.Label10.Location = New System.Drawing.Point(742, 3)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(147, 46)
        Me.Label10.TabIndex = 49
        Me.Label10.Text = "OFFLINE"
        '
        'lvList
        '
        Me.lvList.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.lvList.Border.Class = "ListViewBorder"
        Me.lvList.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.lvList.DisabledBackColor = System.Drawing.Color.Empty
        Me.lvList.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvList.ForeColor = System.Drawing.Color.Green
        Me.lvList.FullRowSelect = True
        Me.lvList.GridLines = True
        Me.lvList.HideSelection = False
        Me.lvList.Location = New System.Drawing.Point(2, 42)
        Me.lvList.Name = "lvList"
        Me.lvList.Size = New System.Drawing.Size(521, 676)
        Me.lvList.TabIndex = 48
        Me.lvList.UseCompatibleStateImageBehavior = False
        Me.lvList.View = System.Windows.Forms.View.Details
        '
        'gbKioskHistory
        '
        Me.gbKioskHistory.Controls.Add(Me.dgvLogs)
        Me.gbKioskHistory.Controls.Add(Me.Label17)
        Me.gbKioskHistory.Controls.Add(Me.Label16)
        Me.gbKioskHistory.Controls.Add(Me.dtpTo)
        Me.gbKioskHistory.Controls.Add(Me.dtpFrom)
        Me.gbKioskHistory.Controls.Add(Me.lblFilter)
        Me.gbKioskHistory.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbKioskHistory.Location = New System.Drawing.Point(6, 318)
        Me.gbKioskHistory.Name = "gbKioskHistory"
        Me.gbKioskHistory.Size = New System.Drawing.Size(1059, 397)
        Me.gbKioskHistory.TabIndex = 46
        Me.gbKioskHistory.TabStop = False
        Me.gbKioskHistory.Text = "SSIT"
        Me.gbKioskHistory.Visible = False
        '
        'dgvLogs
        '
        Me.dgvLogs.AllowUserToAddRows = False
        Me.dgvLogs.AllowUserToDeleteRows = False
        Me.dgvLogs.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvLogs.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvLogs.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.dgvLogs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvLogs.Location = New System.Drawing.Point(6, 49)
        Me.dgvLogs.Name = "dgvLogs"
        Me.dgvLogs.ReadOnly = True
        Me.dgvLogs.RowHeadersWidth = 51
        Me.dgvLogs.Size = New System.Drawing.Size(1047, 342)
        Me.dgvLogs.TabIndex = 45
        Me.dgvLogs.Visible = False
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(104, 25)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(43, 20)
        Me.Label17.TabIndex = 52
        Me.Label17.Text = "From"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(353, 25)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(23, 20)
        Me.Label16.TabIndex = 51
        Me.Label16.Text = "to"
        '
        'dtpTo
        '
        Me.dtpTo.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTo.Location = New System.Drawing.Point(379, 19)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(200, 27)
        Me.dtpTo.TabIndex = 50
        Me.dtpTo.Visible = False
        '
        'dtpFrom
        '
        Me.dtpFrom.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpFrom.Location = New System.Drawing.Point(147, 19)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(200, 27)
        Me.dtpFrom.TabIndex = 49
        Me.dtpFrom.Visible = False
        '
        'lblFilter
        '
        Me.lblFilter.AutoSize = True
        Me.lblFilter.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFilter.Location = New System.Drawing.Point(9, 23)
        Me.lblFilter.Name = "lblFilter"
        Me.lblFilter.Size = New System.Drawing.Size(118, 23)
        Me.lblFilter.TabIndex = 48
        Me.lblFilter.Text = "Filter by date :"
        Me.lblFilter.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.OpEnd1)
        Me.GroupBox1.Controls.Add(Me.OpStart1)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.lblSerial)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.lbldate)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.lblOD)
        Me.GroupBox1.Controls.Add(Me.DURATION)
        Me.GroupBox1.Controls.Add(Me.btnBack)
        Me.GroupBox1.Controls.Add(Me.lblTrans)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.LASTONLINE)
        Me.GroupBox1.Controls.Add(Me.lblGroup)
        Me.GroupBox1.Controls.Add(Me.lblStatus)
        Me.GroupBox1.Controls.Add(Me.lblCluster)
        Me.GroupBox1.Controls.Add(Me.lblLastOnline)
        Me.GroupBox1.Controls.Add(Me.lblKbranch)
        Me.GroupBox1.Controls.Add(Me.lblOffline)
        Me.GroupBox1.Controls.Add(Me.lblKIP)
        Me.GroupBox1.Controls.Add(Me.LASTOFFLINE)
        Me.GroupBox1.Controls.Add(Me.lblKname)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(9, 146)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1059, 166)
        Me.GroupBox1.TabIndex = 43
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "SSIT DETAILS"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(712, 121)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(26, 19)
        Me.Label15.TabIndex = 75
        Me.Label15.Text = "TO"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(614, 121)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(48, 19)
        Me.Label14.TabIndex = 74
        Me.Label14.Text = "FROM"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(693, 133)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(17, 23)
        Me.Label13.TabIndex = 73
        Me.Label13.Text = "-"
        '
        'OpEnd1
        '
        Me.OpEnd1.AutoSize = True
        Me.OpEnd1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OpEnd1.Location = New System.Drawing.Point(712, 133)
        Me.OpEnd1.Name = "OpEnd1"
        Me.OpEnd1.Size = New System.Drawing.Size(93, 23)
        Me.OpEnd1.TabIndex = 72
        Me.OpEnd1.Text = "6:00:00 PM"
        '
        'OpStart1
        '
        Me.OpStart1.AutoSize = True
        Me.OpStart1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OpStart1.Location = New System.Drawing.Point(614, 133)
        Me.OpStart1.Name = "OpStart1"
        Me.OpStart1.Size = New System.Drawing.Size(94, 23)
        Me.OpStart1.TabIndex = 71
        Me.OpStart1.Text = "9:00:00 AM"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Blue
        Me.Label12.Location = New System.Drawing.Point(612, 100)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(188, 28)
        Me.Label12.TabIndex = 70
        Me.Label12.Text = "OPERATION HOURS"
        '
        'lblSerial
        '
        Me.lblSerial.AutoSize = True
        Me.lblSerial.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSerial.Location = New System.Drawing.Point(260, 50)
        Me.lblSerial.Name = "lblSerial"
        Me.lblSerial.Size = New System.Drawing.Size(96, 20)
        Me.lblSerial.TabIndex = 69
        Me.lblSerial.Text = "KIOSK NAME"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Blue
        Me.Label11.Location = New System.Drawing.Point(259, 29)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(89, 28)
        Me.Label11.TabIndex = 68
        Me.Label11.Text = "SERIAL #"
        '
        'lbldate
        '
        Me.lbldate.AutoSize = True
        Me.lbldate.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbldate.Location = New System.Drawing.Point(869, 50)
        Me.lbldate.Name = "lbldate"
        Me.lbldate.Size = New System.Drawing.Size(96, 20)
        Me.lbldate.TabIndex = 67
        Me.lbldate.Text = "KIOSK NAME"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Blue
        Me.Label9.Location = New System.Drawing.Point(867, 29)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(126, 28)
        Me.Label9.TabIndex = 66
        Me.Label9.Text = "Date Created"
        '
        'lblOD
        '
        Me.lblOD.AutoSize = True
        Me.lblOD.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOD.ForeColor = System.Drawing.Color.Black
        Me.lblOD.Location = New System.Drawing.Point(321, 121)
        Me.lblOD.Name = "lblOD"
        Me.lblOD.Size = New System.Drawing.Size(110, 23)
        Me.lblOD.TabIndex = 65
        Me.lblOD.Text = "KIOSK NAME"
        '
        'DURATION
        '
        Me.DURATION.AutoSize = True
        Me.DURATION.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DURATION.ForeColor = System.Drawing.Color.MediumBlue
        Me.DURATION.Location = New System.Drawing.Point(318, 99)
        Me.DURATION.Name = "DURATION"
        Me.DURATION.Size = New System.Drawing.Size(188, 28)
        Me.DURATION.TabIndex = 64
        Me.DURATION.Text = "OFFLINE DURATION"
        '
        'btnBack
        '
        Me.btnBack.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBack.Location = New System.Drawing.Point(999, 27)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(50, 124)
        Me.btnBack.TabIndex = 63
        Me.btnBack.Text = "Back"
        Me.btnBack.UseVisualStyleBackColor = True
        Me.btnBack.Visible = False
        '
        'lblTrans
        '
        Me.lblTrans.AutoSize = True
        Me.lblTrans.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTrans.Location = New System.Drawing.Point(870, 122)
        Me.lblTrans.Name = "lblTrans"
        Me.lblTrans.Size = New System.Drawing.Size(110, 23)
        Me.lblTrans.TabIndex = 58
        Me.lblTrans.Text = "KIOSK NAME"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Blue
        Me.Label8.Location = New System.Drawing.Point(867, 100)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(154, 28)
        Me.Label8.TabIndex = 57
        Me.Label8.Text = "TRANSACTIONS"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Blue
        Me.Label6.Location = New System.Drawing.Point(2, 98)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(79, 28)
        Me.Label6.TabIndex = 46
        Me.Label6.Text = "STATUS"
        '
        'LASTONLINE
        '
        Me.LASTONLINE.AutoSize = True
        Me.LASTONLINE.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LASTONLINE.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.LASTONLINE.Location = New System.Drawing.Point(114, 122)
        Me.LASTONLINE.Name = "LASTONLINE"
        Me.LASTONLINE.Size = New System.Drawing.Size(110, 23)
        Me.LASTONLINE.TabIndex = 56
        Me.LASTONLINE.Text = "KIOSK NAME"
        '
        'lblGroup
        '
        Me.lblGroup.AutoSize = True
        Me.lblGroup.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGroup.Location = New System.Drawing.Point(760, 50)
        Me.lblGroup.Name = "lblGroup"
        Me.lblGroup.Size = New System.Drawing.Size(96, 20)
        Me.lblGroup.TabIndex = 51
        Me.lblGroup.Text = "KIOSK NAME"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(2, 120)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(96, 20)
        Me.lblStatus.TabIndex = 52
        Me.lblStatus.Text = "KIOSK NAME"
        '
        'lblCluster
        '
        Me.lblCluster.AutoSize = True
        Me.lblCluster.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCluster.Location = New System.Drawing.Point(613, 50)
        Me.lblCluster.Name = "lblCluster"
        Me.lblCluster.Size = New System.Drawing.Size(96, 20)
        Me.lblCluster.TabIndex = 50
        Me.lblCluster.Text = "KIOSK NAME"
        '
        'lblLastOnline
        '
        Me.lblLastOnline.AutoSize = True
        Me.lblLastOnline.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastOnline.ForeColor = System.Drawing.Color.Blue
        Me.lblLastOnline.Location = New System.Drawing.Point(112, 99)
        Me.lblLastOnline.Name = "lblLastOnline"
        Me.lblLastOnline.Size = New System.Drawing.Size(130, 28)
        Me.lblLastOnline.TabIndex = 53
        Me.lblLastOnline.Text = "LAST ONLINE"
        '
        'lblKbranch
        '
        Me.lblKbranch.AutoSize = True
        Me.lblKbranch.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKbranch.Location = New System.Drawing.Point(398, 50)
        Me.lblKbranch.Name = "lblKbranch"
        Me.lblKbranch.Size = New System.Drawing.Size(96, 20)
        Me.lblKbranch.TabIndex = 49
        Me.lblKbranch.Text = "KIOSK NAME"
        '
        'lblOffline
        '
        Me.lblOffline.AutoSize = True
        Me.lblOffline.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOffline.ForeColor = System.Drawing.Color.Blue
        Me.lblOffline.Location = New System.Drawing.Point(113, 99)
        Me.lblOffline.Name = "lblOffline"
        Me.lblOffline.Size = New System.Drawing.Size(135, 28)
        Me.lblOffline.TabIndex = 55
        Me.lblOffline.Text = "LAST OFFLINE"
        '
        'lblKIP
        '
        Me.lblKIP.AutoSize = True
        Me.lblKIP.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKIP.Location = New System.Drawing.Point(115, 50)
        Me.lblKIP.Name = "lblKIP"
        Me.lblKIP.Size = New System.Drawing.Size(96, 20)
        Me.lblKIP.TabIndex = 48
        Me.lblKIP.Text = "KIOSK NAME"
        '
        'LASTOFFLINE
        '
        Me.LASTOFFLINE.AutoSize = True
        Me.LASTOFFLINE.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LASTOFFLINE.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.LASTOFFLINE.Location = New System.Drawing.Point(114, 122)
        Me.LASTOFFLINE.Name = "LASTOFFLINE"
        Me.LASTOFFLINE.Size = New System.Drawing.Size(110, 23)
        Me.LASTOFFLINE.TabIndex = 54
        Me.LASTOFFLINE.Text = "KIOSK NAME"
        '
        'lblKname
        '
        Me.lblKname.AutoSize = True
        Me.lblKname.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKname.Location = New System.Drawing.Point(4, 50)
        Me.lblKname.Name = "lblKname"
        Me.lblKname.Size = New System.Drawing.Size(96, 20)
        Me.lblKname.TabIndex = 47
        Me.lblKname.Text = "KIOSK NAME"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Blue
        Me.Label5.Location = New System.Drawing.Point(611, 29)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(94, 28)
        Me.Label5.TabIndex = 45
        Me.Label5.Text = "DIVISION"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Blue
        Me.Label4.Location = New System.Drawing.Point(758, 29)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(78, 28)
        Me.Label4.TabIndex = 44
        Me.Label4.Text = "GROUP"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(115, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(118, 28)
        Me.Label3.TabIndex = 43
        Me.Label3.Text = "IP ADDRESS"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(3, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(123, 28)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "HOST NAME"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Blue
        Me.Label2.Location = New System.Drawing.Point(397, 29)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(89, 28)
        Me.Label2.TabIndex = 42
        Me.Label2.Text = "BRANCH"
        '
        '_frmMonitoring
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1074, 721)
        Me.Controls.Add(Me.Panel1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "_frmMonitoring"
        Me.Text = "MetroForm"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.gbKioskHistory.ResumeLayout(False)
        Me.gbKioskHistory.PerformLayout()
        CType(Me.dgvLogs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lblGroup As System.Windows.Forms.Label
    Friend WithEvents lblCluster As System.Windows.Forms.Label
    Friend WithEvents lblKbranch As System.Windows.Forms.Label
    Friend WithEvents lblKIP As System.Windows.Forms.Label
    Friend WithEvents lblKname As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents LASTOFFLINE As System.Windows.Forms.Label
    Friend WithEvents lblLastOnline As System.Windows.Forms.Label
    Friend WithEvents LASTONLINE As System.Windows.Forms.Label
    Friend WithEvents lblOffline As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lblTrans As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents lblOD As System.Windows.Forms.Label
    Friend WithEvents DURATION As System.Windows.Forms.Label
    Friend WithEvents lbldate As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lblSerial As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents OpEnd1 As System.Windows.Forms.Label
    Friend WithEvents OpStart1 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents gbKioskHistory As System.Windows.Forms.GroupBox
    Friend WithEvents dgvLogs As System.Windows.Forms.DataGridView
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFilter As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lvPC As DevComponents.DotNetBar.Controls.ListViewEx
    Friend WithEvents lvList As DevComponents.DotNetBar.Controls.ListViewEx
End Class
