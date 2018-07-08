<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FSheetSel
    Inherits System.Windows.Forms.Form

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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.BCancel = New System.Windows.Forms.Button()
        Me.BOK = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SelSheets = New System.Windows.Forms.RadioButton()
        Me.AllSheets = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.AllSheets)
        Me.GroupBox1.Controls.Add(Me.SelSheets)
        Me.GroupBox1.Location = New System.Drawing.Point(52, 28)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 100)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Sheets?"
        '
        'BCancel
        '
        Me.BCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BCancel.Location = New System.Drawing.Point(177, 215)
        Me.BCancel.Name = "BCancel"
        Me.BCancel.Size = New System.Drawing.Size(75, 23)
        Me.BCancel.TabIndex = 6
        Me.BCancel.Text = "Cancel"
        Me.BCancel.UseVisualStyleBackColor = True
        '
        'BOK
        '
        Me.BOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.BOK.Location = New System.Drawing.Point(52, 215)
        Me.BOK.Name = "BOK"
        Me.BOK.Size = New System.Drawing.Size(75, 23)
        Me.BOK.TabIndex = 5
        Me.BOK.Text = "OK"
        Me.BOK.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Enabled = False
        Me.TextBox1.Location = New System.Drawing.Point(10, 134)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(288, 64)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.Text = "The sheets will be formatted to fit on a page and the headers, footers & page num" &
    "bers will be set. Choose the correct printer before using this function."
        '
        'SelSheets
        '
        Me.SelSheets.AutoSize = True
        Me.SelSheets.Checked = True
        Me.SelSheets.Location = New System.Drawing.Point(38, 28)
        Me.SelSheets.Name = "SelSheets"
        Me.SelSheets.Size = New System.Drawing.Size(132, 21)
        Me.SelSheets.TabIndex = 2
        Me.SelSheets.Text = "Selected Sheets"
        Me.SelSheets.UseVisualStyleBackColor = True
        '
        'AllSheets
        '
        Me.AllSheets.AutoSize = True
        Me.AllSheets.Location = New System.Drawing.Point(38, 60)
        Me.AllSheets.Name = "AllSheets"
        Me.AllSheets.Size = New System.Drawing.Size(92, 21)
        Me.AllSheets.TabIndex = 3
        Me.AllSheets.Text = "All Sheets"
        Me.AllSheets.UseVisualStyleBackColor = True
        '
        'FSheetSel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(319, 272)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BCancel)
        Me.Controls.Add(Me.BOK)
        Me.Controls.Add(Me.TextBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FSheetSel"
        Me.Text = "SheetSel"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents BCancel As Windows.Forms.Button
    Friend WithEvents BOK As Windows.Forms.Button
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents AllSheets As Windows.Forms.RadioButton
    Friend WithEvents SelSheets As Windows.Forms.RadioButton
End Class
