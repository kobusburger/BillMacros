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
        Me.CBAll = New System.Windows.Forms.CheckBox()
        Me.SelSheets = New System.Windows.Forms.CheckBox()
        Me.BCancel = New System.Windows.Forms.Button()
        Me.OK = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CBAll)
        Me.GroupBox1.Controls.Add(Me.SelSheets)
        Me.GroupBox1.Location = New System.Drawing.Point(52, 28)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 100)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Sheets?"
        '
        'CBAll
        '
        Me.CBAll.AutoSize = True
        Me.CBAll.Location = New System.Drawing.Point(39, 50)
        Me.CBAll.Name = "CBAll"
        Me.CBAll.Size = New System.Drawing.Size(91, 21)
        Me.CBAll.TabIndex = 1
        Me.CBAll.Text = "All sheets"
        Me.CBAll.UseVisualStyleBackColor = True
        '
        'SelSheets
        '
        Me.SelSheets.AutoSize = True
        Me.SelSheets.Checked = True
        Me.SelSheets.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SelSheets.Location = New System.Drawing.Point(39, 22)
        Me.SelSheets.Name = "SelSheets"
        Me.SelSheets.Size = New System.Drawing.Size(131, 21)
        Me.SelSheets.TabIndex = 0
        Me.SelSheets.Text = "Selected sheets"
        Me.SelSheets.UseVisualStyleBackColor = True
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
        'OK
        '
        Me.OK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OK.Location = New System.Drawing.Point(52, 215)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(75, 23)
        Me.OK.TabIndex = 5
        Me.OK.Text = "OK"
        Me.OK.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(10, 134)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(288, 64)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.Text = "The sheets will be formatted to fit on a page and the headers, footers & page num" &
    "bers will be set. Choose the correct printer before using this function."
        '
        'FSheetSel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(319, 272)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BCancel)
        Me.Controls.Add(Me.OK)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "FSheetSel"
        Me.Text = "SheetSel"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents CBAll As Windows.Forms.CheckBox
    Friend WithEvents SelSheets As Windows.Forms.CheckBox
    Friend WithEvents BCancel As Windows.Forms.Button
    Friend WithEvents OK As Windows.Forms.Button
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
End Class
