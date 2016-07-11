<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NavInput
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ID1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ID2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.UOM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Description = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Category = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Count = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Context1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Context2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Context3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(348, 37)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Excel Shop Order Input"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(497, 7)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(111, 52)
        Me.Button2.TabIndex = 18
        Me.Button2.Text = "Upload Table"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(380, 7)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(111, 52)
        Me.Button1.TabIndex = 17
        Me.Button1.Text = "Open File"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(631, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 16)
        Me.Label1.TabIndex = 16
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ID1, Me.ID2, Me.UOM, Me.Description, Me.Category, Me.Count, Me.Context1, Me.Context2, Me.Context3})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 71)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1160, 573)
        Me.DataGridView1.TabIndex = 15
        '
        'ID1
        '
        Me.ID1.HeaderText = "ID1"
        Me.ID1.Name = "ID1"
        '
        'ID2
        '
        Me.ID2.HeaderText = "ID2"
        Me.ID2.Name = "ID2"
        '
        'UOM
        '
        Me.UOM.HeaderText = "UOM"
        Me.UOM.Name = "UOM"
        '
        'Description
        '
        Me.Description.HeaderText = "Description"
        Me.Description.Name = "Description"
        '
        'Category
        '
        Me.Category.HeaderText = "Category"
        Me.Category.Name = "Category"
        '
        'Count
        '
        Me.Count.HeaderText = "Count"
        Me.Count.Name = "Count"
        '
        'Context1
        '
        Me.Context1.HeaderText = "Context1"
        Me.Context1.Name = "Context1"
        '
        'Context2
        '
        Me.Context2.HeaderText = "Context2"
        Me.Context2.Name = "Context2"
        '
        'Context3
        '
        Me.Context3.HeaderText = "Context3"
        Me.Context3.Name = "Context3"
        '
        'NavInput
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1184, 761)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "NavInput"
        Me.Text = "NavInput"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label2 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents ID1 As DataGridViewTextBoxColumn
    Friend WithEvents ID2 As DataGridViewTextBoxColumn
    Friend WithEvents UOM As DataGridViewTextBoxColumn
    Friend WithEvents Description As DataGridViewTextBoxColumn
    Friend WithEvents Category As DataGridViewTextBoxColumn
    Friend WithEvents Count As DataGridViewTextBoxColumn
    Friend WithEvents Context1 As DataGridViewTextBoxColumn
    Friend WithEvents Context2 As DataGridViewTextBoxColumn
    Friend WithEvents Context3 As DataGridViewTextBoxColumn
End Class
