<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmUpload
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
        Me.rbUpdater = New System.Windows.Forms.RadioButton()
        Me.rbTerms = New System.Windows.Forms.RadioButton()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.rbEXE = New System.Windows.Forms.RadioButton()
        Me.rbVideo = New System.Windows.Forms.RadioButton()
        Me.rbPDF = New System.Windows.Forms.RadioButton()
        Me.btnUpload = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Panel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.TabControl1)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.btnUpload)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(842, 615)
        Me.Panel1.TabIndex = 0
        '
        'rbUpdater
        '
        Me.rbUpdater.AutoSize = True
        Me.rbUpdater.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbUpdater.Location = New System.Drawing.Point(6, 46)
        Me.rbUpdater.Name = "rbUpdater"
        Me.rbUpdater.Size = New System.Drawing.Size(166, 34)
        Me.rbUpdater.TabIndex = 15
        Me.rbUpdater.TabStop = True
        Me.rbUpdater.Text = "SSIT UPDATER"
        Me.rbUpdater.UseVisualStyleBackColor = True
        '
        'rbTerms
        '
        Me.rbTerms.AutoSize = True
        Me.rbTerms.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbTerms.Location = New System.Drawing.Point(6, 86)
        Me.rbTerms.Name = "rbTerms"
        Me.rbTerms.Size = New System.Drawing.Size(313, 34)
        Me.rbTerms.TabIndex = 14
        Me.rbTerms.TabStop = True
        Me.rbTerms.Text = "SSIT TERMS AND CONDITION"
        Me.rbTerms.UseVisualStyleBackColor = True
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(189, 40)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(404, 42)
        Me.Label16.TabIndex = 13
        Me.Label16.Text = "SSIT FILE UPLOADER"
        '
        'rbEXE
        '
        Me.rbEXE.AutoSize = True
        Me.rbEXE.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbEXE.Location = New System.Drawing.Point(3, 6)
        Me.rbEXE.Name = "rbEXE"
        Me.rbEXE.Size = New System.Drawing.Size(148, 34)
        Me.rbEXE.TabIndex = 11
        Me.rbEXE.TabStop = True
        Me.rbEXE.Text = "SSIT SERVER"
        Me.rbEXE.UseVisualStyleBackColor = True
        '
        'rbVideo
        '
        Me.rbVideo.AutoSize = True
        Me.rbVideo.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbVideo.Location = New System.Drawing.Point(6, 6)
        Me.rbVideo.Name = "rbVideo"
        Me.rbVideo.Size = New System.Drawing.Size(131, 34)
        Me.rbVideo.TabIndex = 10
        Me.rbVideo.TabStop = True
        Me.rbVideo.Text = "SSS VIDEO"
        Me.rbVideo.UseVisualStyleBackColor = True
        '
        'rbPDF
        '
        Me.rbPDF.AutoSize = True
        Me.rbPDF.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbPDF.Location = New System.Drawing.Point(6, 46)
        Me.rbPDF.Name = "rbPDF"
        Me.rbPDF.Size = New System.Drawing.Size(202, 34)
        Me.rbPDF.TabIndex = 9
        Me.rbPDF.TabStop = True
        Me.rbPDF.Text = "CITIZEN CHARTER"
        Me.rbPDF.UseVisualStyleBackColor = True
        '
        'btnUpload
        '
        Me.btnUpload.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btnUpload.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpload.Location = New System.Drawing.Point(249, 509)
        Me.btnUpload.Name = "btnUpload"
        Me.btnUpload.Size = New System.Drawing.Size(344, 92)
        Me.btnUpload.TabIndex = 8
        Me.btnUpload.Text = "Browse and upload file"
        Me.btnUpload.UseVisualStyleBackColor = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(10, 108)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(818, 384)
        Me.TabControl1.TabIndex = 17
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.rbEXE)
        Me.TabPage1.Controls.Add(Me.rbUpdater)
        Me.TabPage1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage1.Location = New System.Drawing.Point(4, 34)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(810, 346)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "SSIT APPLICATIONS"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.rbVideo)
        Me.TabPage2.Controls.Add(Me.rbPDF)
        Me.TabPage2.Controls.Add(Me.rbTerms)
        Me.TabPage2.Location = New System.Drawing.Point(4, 34)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(810, 346)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "MISCELLANEOUS"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        '_frmUpload
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(842, 615)
        Me.Controls.Add(Me.Panel1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "_frmUpload"
        Me.Text = "MetroForm"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnUpload As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents rbEXE As System.Windows.Forms.RadioButton
    Friend WithEvents rbVideo As System.Windows.Forms.RadioButton
    Friend WithEvents rbPDF As System.Windows.Forms.RadioButton
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents rbTerms As System.Windows.Forms.RadioButton
    Friend WithEvents rbUpdater As System.Windows.Forms.RadioButton
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
End Class
