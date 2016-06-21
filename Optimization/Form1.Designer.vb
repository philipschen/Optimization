<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BStock = New System.Windows.Forms.Button()
        Me.BPart = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(26, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(109, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Welcome to Easy Cut"
        '
        'BStock
        '
        Me.BStock.Location = New System.Drawing.Point(29, 52)
        Me.BStock.Name = "BStock"
        Me.BStock.Size = New System.Drawing.Size(112, 64)
        Me.BStock.TabIndex = 1
        Me.BStock.Text = "Input Stock"
        Me.BStock.UseVisualStyleBackColor = True
        '
        'BPart
        '
        Me.BPart.Location = New System.Drawing.Point(29, 139)
        Me.BPart.Name = "BPart"
        Me.BPart.Size = New System.Drawing.Size(112, 64)
        Me.BPart.TabIndex = 2
        Me.BPart.Text = "Input Parts To Cut"
        Me.BPart.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(165, 52)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(112, 64)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "[Temp] Clean data"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(448, 334)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.BPart)
        Me.Controls.Add(Me.BStock)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents BStock As Button
    Friend WithEvents BPart As Button
    Friend WithEvents Button3 As Button
End Class
