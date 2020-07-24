<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OutputForm
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
        Me.OutputTextBox = New System.Windows.Forms.TextBox()
        Me.OutputOKButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'OutputTextBox
        '
        Me.OutputTextBox.Location = New System.Drawing.Point(13, 13)
        Me.OutputTextBox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OutputTextBox.Multiline = True
        Me.OutputTextBox.Name = "OutputTextBox"
        Me.OutputTextBox.ReadOnly = True
        Me.OutputTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.OutputTextBox.Size = New System.Drawing.Size(756, 543)
        Me.OutputTextBox.TabIndex = 0
        '
        'OutputOKButton
        '
        Me.OutputOKButton.Location = New System.Drawing.Point(13, 564)
        Me.OutputOKButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OutputOKButton.Name = "OutputOKButton"
        Me.OutputOKButton.Size = New System.Drawing.Size(756, 48)
        Me.OutputOKButton.TabIndex = 1
        Me.OutputOKButton.Text = "OK"
        Me.OutputOKButton.UseVisualStyleBackColor = True
        '
        'OutputForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(782, 625)
        Me.Controls.Add(Me.OutputOKButton)
        Me.Controls.Add(Me.OutputTextBox)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "OutputForm"
        Me.Text = "Output"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents OutputTextBox As TextBox
    Friend WithEvents OutputOKButton As Button
End Class
