<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class _frmLogin
    Inherits DevComponents.DotNetBar.Metro.MetroForm

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(_frmLogin))
        Me.btnClose = New DevComponents.DotNetBar.ButtonX()
        Me.ButtonX3 = New DevComponents.DotNetBar.ButtonX()
        Me.chkRemember = New DevComponents.DotNetBar.Controls.CheckBoxX()
        Me.txtPass = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.txtUname = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.tmr1 = New System.Windows.Forms.Timer(Me.components)
        Me.pbError = New System.Windows.Forms.PictureBox()
        Me.lblErrorMessage = New System.Windows.Forms.Label()
        Me.llblForgot = New System.Windows.Forms.LinkLabel()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.pbError, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.btnClose.BackColor = System.Drawing.Color.Transparent
        Me.btnClose.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.btnClose.Image = CType(resources.GetObject("btnClose.Image"), System.Drawing.Image)
        Me.btnClose.Location = New System.Drawing.Point(415, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Shape = New DevComponents.DotNetBar.EllipticalShapeDescriptor()
        Me.btnClose.Size = New System.Drawing.Size(27, 26)
        Me.btnClose.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.btnClose.TabIndex = 17
        '
        'ButtonX3
        '
        Me.ButtonX3.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.ButtonX3.BackColor = System.Drawing.Color.Transparent
        Me.ButtonX3.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.ButtonX3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonX3.Location = New System.Drawing.Point(285, 285)
        Me.ButtonX3.Name = "ButtonX3"
        Me.ButtonX3.Shape = New DevComponents.DotNetBar.RoundRectangleShapeDescriptor(2)
        Me.ButtonX3.Size = New System.Drawing.Size(60, 23)
        Me.ButtonX3.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.ButtonX3.TabIndex = 20
        Me.ButtonX3.Text = "Log In"
        '
        'chkRemember
        '
        Me.chkRemember.BackColor = System.Drawing.Color.Transparent
        '
        '
        '
        Me.chkRemember.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.chkRemember.Location = New System.Drawing.Point(133, 285)
        Me.chkRemember.Name = "chkRemember"
        Me.chkRemember.Size = New System.Drawing.Size(100, 23)
        Me.chkRemember.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.chkRemember.TabIndex = 20
        Me.chkRemember.Text = "Remember Me"
        Me.chkRemember.TextColor = System.Drawing.Color.White
        '
        'txtPass
        '
        Me.txtPass.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.txtPass.Border.Class = "TextBoxBorder"
        Me.txtPass.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.txtPass.DisabledBackColor = System.Drawing.Color.White
        Me.txtPass.ForeColor = System.Drawing.Color.Black
        Me.txtPass.Location = New System.Drawing.Point(163, 251)
        Me.txtPass.Name = "txtPass"
        Me.txtPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPass.Size = New System.Drawing.Size(182, 26)
        Me.txtPass.TabIndex = 19
        '
        'txtUname
        '
        Me.txtUname.BackColor = System.Drawing.Color.White
        '
        '
        '
        Me.txtUname.Border.Class = "TextBoxBorder"
        Me.txtUname.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.txtUname.DisabledBackColor = System.Drawing.Color.White
        Me.txtUname.ForeColor = System.Drawing.Color.Black
        Me.txtUname.Location = New System.Drawing.Point(163, 212)
        Me.txtUname.Name = "txtUname"
        Me.txtUname.Size = New System.Drawing.Size(182, 26)
        Me.txtUname.TabIndex = 18
        '
        'tmr1
        '
        '
        'pbError
        '
        Me.pbError.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbError.BackColor = System.Drawing.Color.Transparent
        Me.pbError.BackgroundImage = CType(resources.GetObject("pbError.BackgroundImage"), System.Drawing.Image)
        Me.pbError.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.pbError.ForeColor = System.Drawing.Color.Black
        Me.pbError.Location = New System.Drawing.Point(133, 309)
        Me.pbError.Name = "pbError"
        Me.pbError.Size = New System.Drawing.Size(19, 18)
        Me.pbError.TabIndex = 35
        Me.pbError.TabStop = False
        Me.pbError.Visible = False
        '
        'lblErrorMessage
        '
        Me.lblErrorMessage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblErrorMessage.AutoSize = True
        Me.lblErrorMessage.BackColor = System.Drawing.Color.Transparent
        Me.lblErrorMessage.ForeColor = System.Drawing.Color.Red
        Me.lblErrorMessage.Location = New System.Drawing.Point(155, 314)
        Me.lblErrorMessage.Name = "lblErrorMessage"
        Me.lblErrorMessage.Size = New System.Drawing.Size(115, 19)
        Me.lblErrorMessage.TabIndex = 34
        Me.lblErrorMessage.Text = "ERROR MESSAGE"
        Me.lblErrorMessage.Visible = False
        '
        'llblForgot
        '
        Me.llblForgot.AutoSize = True
        Me.llblForgot.BackColor = System.Drawing.Color.Transparent
        Me.llblForgot.Font = New System.Drawing.Font("Segoe UI", 10.25!)
        Me.llblForgot.Location = New System.Drawing.Point(154, 354)
        Me.llblForgot.Name = "llblForgot"
        Me.llblForgot.Size = New System.Drawing.Size(154, 25)
        Me.llblForgot.TabIndex = 22
        Me.llblForgot.TabStop = True
        Me.llblForgot.Text = "Forgot Password?"
        Me.llblForgot.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.sss_management.My.Resources.Resources.Untitled__2_
        Me.PictureBox2.Location = New System.Drawing.Point(133, 376)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(232, 50)
        Me.PictureBox2.TabIndex = 37
        Me.PictureBox2.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(73, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(331, 32)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "SET MANAGEMENT SERVER"
        '
        '_frmLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(448, 471)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.pbError)
        Me.Controls.Add(Me.lblErrorMessage)
        Me.Controls.Add(Me.ButtonX3)
        Me.Controls.Add(Me.llblForgot)
        Me.Controls.Add(Me.chkRemember)
        Me.Controls.Add(Me.txtPass)
        Me.Controls.Add(Me.txtUname)
        Me.Controls.Add(Me.btnClose)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "_frmLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MetroForm"
        CType(Me.pbError, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As DevComponents.DotNetBar.ButtonX
    Friend WithEvents ButtonX3 As DevComponents.DotNetBar.ButtonX
    Friend WithEvents chkRemember As DevComponents.DotNetBar.Controls.CheckBoxX
    Friend WithEvents txtPass As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents txtUname As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents tmr1 As System.Windows.Forms.Timer
    Friend WithEvents pbError As System.Windows.Forms.PictureBox
    Friend WithEvents lblErrorMessage As System.Windows.Forms.Label
    Friend WithEvents llblForgot As System.Windows.Forms.LinkLabel
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As Label
End Class
