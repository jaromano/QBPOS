<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.btnFindServer = New System.Windows.Forms.Button()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnFindServer
        '
        Me.btnFindServer.Location = New System.Drawing.Point(51, 49)
        Me.btnFindServer.Name = "btnFindServer"
        Me.btnFindServer.Size = New System.Drawing.Size(161, 31)
        Me.btnFindServer.TabIndex = 0
        Me.btnFindServer.Text = "btnFindServer"
        Me.btnFindServer.UseVisualStyleBackColor = True
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(45, 13)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(200, 20)
        Me.txtServer.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.btnFindServer)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnFindServer As System.Windows.Forms.Button
    Friend WithEvents txtServer As System.Windows.Forms.TextBox

End Class
