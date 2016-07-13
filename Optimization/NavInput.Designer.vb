<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class NavInput

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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.stockID1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.stockID2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.stockID3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.description = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.color = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.size = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.count = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.internalID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.context1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.context2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.context3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(239, 37)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "Excel Nav Input"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(497, 7)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(111, 52)
        Me.Button2.TabIndex = 23
        Me.Button2.Text = "Upload Table"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(380, 7)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(111, 52)
        Me.Button1.TabIndex = 22
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
        Me.Label1.TabIndex = 21
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.stockID1, Me.stockID2, Me.stockID3, Me.description, Me.color, Me.size, Me.count, Me.internalID, Me.context1, Me.context2, Me.context3})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 65)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1160, 584)
        Me.DataGridView1.TabIndex = 25
        '
        'stockID1
        '
        Me.stockID1.HeaderText = "stockID1"
        Me.stockID1.Name = "stockID1"
        '
        'stockID2
        '
        Me.stockID2.HeaderText = "stockID2"
        Me.stockID2.Name = "stockID2"
        '
        'stockID3
        '
        Me.stockID3.HeaderText = "stockID3"
        Me.stockID3.Name = "stockID3"
        '
        'description
        '
        Me.description.HeaderText = "description"
        Me.description.Name = "description"
        '
        'color
        '
        Me.color.HeaderText = "color"
        Me.color.Name = "color"
        '
        'size
        '
        Me.size.HeaderText = "size"
        Me.size.Name = "size"
        '
        'count
        '
        Me.count.HeaderText = "count"
        Me.count.Name = "count"
        '
        'internalID
        '
        Me.internalID.HeaderText = "internalID"
        Me.internalID.Name = "internalID"
        Me.internalID.ReadOnly = True
        '
        'context1
        '
        Me.context1.HeaderText = "context1"
        Me.context1.Name = "context1"
        '
        'context2
        '
        Me.context2.HeaderText = "context2"
        Me.context2.Name = "context2"
        '
        'context3
        '
        Me.context3.HeaderText = "context3"
        Me.context3.Name = "context3"
        '
        'NavInput
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1184, 661)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
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
    Friend WithEvents stockID1 As DataGridViewTextBoxColumn
    Friend WithEvents stockID2 As DataGridViewTextBoxColumn
    Friend WithEvents stockID3 As DataGridViewTextBoxColumn
    Friend WithEvents description As DataGridViewTextBoxColumn
    Friend WithEvents color As DataGridViewTextBoxColumn
    Friend WithEvents size As DataGridViewTextBoxColumn
    Friend WithEvents count As DataGridViewTextBoxColumn
    Friend WithEvents internalID As DataGridViewTextBoxColumn
    Friend WithEvents context1 As DataGridViewTextBoxColumn
    Friend WithEvents context2 As DataGridViewTextBoxColumn
    Friend WithEvents context3 As DataGridViewTextBoxColumn
End Class
