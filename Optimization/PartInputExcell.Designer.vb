﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PartInputExcell
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
        Me.Button2 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.partID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.description = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.color = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.size = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.count = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.internalID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.shopNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.itemNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.itemQuantity = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.context1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.context2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.context3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(93, 12)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.partID, Me.description, Me.color, Me.size, Me.count, Me.internalID, Me.shopNumber, Me.itemNumber, Me.itemQuantity, Me.context1, Me.context2, Me.context3})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 41)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1160, 608)
        Me.DataGridView1.TabIndex = 4
        '
        'partID
        '
        Me.partID.HeaderText = "partID"
        Me.partID.Name = "partID"
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
        'shopNumber
        '
        Me.shopNumber.HeaderText = "shopNumber"
        Me.shopNumber.Name = "shopNumber"
        '
        'itemNumber
        '
        Me.itemNumber.HeaderText = "itemNumber"
        Me.itemNumber.Name = "itemNumber"
        '
        'itemQuantity
        '
        Me.itemQuantity.HeaderText = "itemQuantity"
        Me.itemQuantity.Name = "itemQuantity"
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
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(360, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 13)
        Me.Label1.TabIndex = 6
        '
        'PartInputExcell
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1184, 661)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "PartInputExcell"
        Me.Text = "PartInputExcell"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button2 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button1 As Button
    Friend WithEvents partID As DataGridViewTextBoxColumn
    Friend WithEvents description As DataGridViewTextBoxColumn
    Friend WithEvents color As DataGridViewTextBoxColumn
    Friend WithEvents size As DataGridViewTextBoxColumn
    Friend WithEvents count As DataGridViewTextBoxColumn
    Friend WithEvents internalID As DataGridViewTextBoxColumn
    Friend WithEvents shopNumber As DataGridViewTextBoxColumn
    Friend WithEvents itemNumber As DataGridViewTextBoxColumn
    Friend WithEvents itemQuantity As DataGridViewTextBoxColumn
    Friend WithEvents context1 As DataGridViewTextBoxColumn
    Friend WithEvents context2 As DataGridViewTextBoxColumn
    Friend WithEvents context3 As DataGridViewTextBoxColumn
    Friend WithEvents Label1 As Label
End Class
