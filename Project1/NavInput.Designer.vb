﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Excel Shop Order Input"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(497, 7)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(111, 52)
        Me.Button2.TabIndex = 13
        Me.Button2.Text = "Upload Table"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(380, 7)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(111, 52)
        Me.Button1.TabIndex = 12
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
        Me.Label1.TabIndex = 11
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.partID, Me.description, Me.color, Me.size, Me.count, Me.internalID, Me.shopNumber, Me.itemNumber, Me.itemQuantity, Me.context1, Me.context2, Me.context3})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 71)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1160, 573)
        Me.DataGridView1.TabIndex = 10
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

    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents partID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents color As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents size As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents count As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents internalID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents shopNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents itemNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents itemQuantity As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents context1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents context2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents context3 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class