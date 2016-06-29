<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DirectlyEdit
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
        Me.components = New System.ComponentModel.Container()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
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
        Me.StockID1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.StockID2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.StockID3DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DescriptionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColorDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SizeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CountDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InternalIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Context1DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Context2DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Context3DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.StockNewBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.OptimizationDatabaseDataSetBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.OptimizationDatabaseDataSet = New Optimization.OptimizationDatabaseDataSet()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.StockNewTableAdapter = New Optimization.OptimizationDatabaseDataSetTableAdapters.stockNewTableAdapter()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StockNewBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OptimizationDatabaseDataSetBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OptimizationDatabaseDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.ItemSize = New System.Drawing.Size(100, 40)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1160, 637)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 44)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1152, 589)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "New Stock"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.stockID1, Me.stockID2, Me.stockID3, Me.description, Me.color, Me.size, Me.count, Me.internalID, Me.context1, Me.context2, Me.context3})
        Me.DataGridView1.DataSource = Me.StockNewBindingSource
        Me.DataGridView1.Location = New System.Drawing.Point(6, 6)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1140, 577)
        Me.DataGridView1.TabIndex = 0
        '
        'stockID1
        '
        Me.stockID1.DataPropertyName = "stockID1"
        Me.stockID1.HeaderText = "stockID1"
        Me.stockID1.Name = "stockID1"
        '
        'stockID2
        '
        Me.stockID2.DataPropertyName = "stockID2"
        Me.stockID2.HeaderText = "stockID2"
        Me.stockID2.Name = "stockID2"
        '
        'stockID3
        '
        Me.stockID3.DataPropertyName = "stockID3"
        Me.stockID3.HeaderText = "stockID3"
        Me.stockID3.Name = "stockID3"
        '
        'description
        '
        Me.description.DataPropertyName = "description"
        Me.description.HeaderText = "description"
        Me.description.Name = "description"
        '
        'color
        '
        Me.color.DataPropertyName = "color"
        Me.color.HeaderText = "color"
        Me.color.Name = "color"
        '
        'size
        '
        Me.size.DataPropertyName = "size"
        Me.size.HeaderText = "size"
        Me.size.Name = "size"
        '
        'count
        '
        Me.count.DataPropertyName = "count"
        Me.count.HeaderText = "count"
        Me.count.Name = "count"
        '
        'internalID
        '
        Me.internalID.DataPropertyName = "internalID"
        Me.internalID.HeaderText = "internalID"
        Me.internalID.Name = "internalID"
        Me.internalID.ReadOnly = True
        '
        'context1
        '
        Me.context1.DataPropertyName = "context1"
        Me.context1.HeaderText = "context1"
        Me.context1.Name = "context1"
        '
        'context2
        '
        Me.context2.DataPropertyName = "context2"
        Me.context2.HeaderText = "context2"
        Me.context2.Name = "context2"
        '
        'context3
        '
        Me.context3.DataPropertyName = "context3"
        Me.context3.HeaderText = "context3"
        Me.context3.Name = "context3"
        '
        'StockID1DataGridViewTextBoxColumn
        '
        Me.StockID1DataGridViewTextBoxColumn.DataPropertyName = "stockID1"
        Me.StockID1DataGridViewTextBoxColumn.HeaderText = "stockID1"
        Me.StockID1DataGridViewTextBoxColumn.Name = "StockID1DataGridViewTextBoxColumn"
        '
        'StockID2DataGridViewTextBoxColumn
        '
        Me.StockID2DataGridViewTextBoxColumn.DataPropertyName = "stockID2"
        Me.StockID2DataGridViewTextBoxColumn.HeaderText = "stockID2"
        Me.StockID2DataGridViewTextBoxColumn.Name = "StockID2DataGridViewTextBoxColumn"
        '
        'StockID3DataGridViewTextBoxColumn
        '
        Me.StockID3DataGridViewTextBoxColumn.DataPropertyName = "stockID3"
        Me.StockID3DataGridViewTextBoxColumn.HeaderText = "stockID3"
        Me.StockID3DataGridViewTextBoxColumn.Name = "StockID3DataGridViewTextBoxColumn"
        '
        'DescriptionDataGridViewTextBoxColumn
        '
        Me.DescriptionDataGridViewTextBoxColumn.DataPropertyName = "description"
        Me.DescriptionDataGridViewTextBoxColumn.HeaderText = "description"
        Me.DescriptionDataGridViewTextBoxColumn.Name = "DescriptionDataGridViewTextBoxColumn"
        '
        'ColorDataGridViewTextBoxColumn
        '
        Me.ColorDataGridViewTextBoxColumn.DataPropertyName = "color"
        Me.ColorDataGridViewTextBoxColumn.HeaderText = "color"
        Me.ColorDataGridViewTextBoxColumn.Name = "ColorDataGridViewTextBoxColumn"
        '
        'SizeDataGridViewTextBoxColumn
        '
        Me.SizeDataGridViewTextBoxColumn.DataPropertyName = "size"
        Me.SizeDataGridViewTextBoxColumn.HeaderText = "size"
        Me.SizeDataGridViewTextBoxColumn.Name = "SizeDataGridViewTextBoxColumn"
        '
        'CountDataGridViewTextBoxColumn
        '
        Me.CountDataGridViewTextBoxColumn.DataPropertyName = "count"
        Me.CountDataGridViewTextBoxColumn.HeaderText = "count"
        Me.CountDataGridViewTextBoxColumn.Name = "CountDataGridViewTextBoxColumn"
        '
        'InternalIDDataGridViewTextBoxColumn
        '
        Me.InternalIDDataGridViewTextBoxColumn.DataPropertyName = "internalID"
        Me.InternalIDDataGridViewTextBoxColumn.HeaderText = "internalID"
        Me.InternalIDDataGridViewTextBoxColumn.Name = "InternalIDDataGridViewTextBoxColumn"
        '
        'Context1DataGridViewTextBoxColumn
        '
        Me.Context1DataGridViewTextBoxColumn.DataPropertyName = "context1"
        Me.Context1DataGridViewTextBoxColumn.HeaderText = "context1"
        Me.Context1DataGridViewTextBoxColumn.Name = "Context1DataGridViewTextBoxColumn"
        '
        'Context2DataGridViewTextBoxColumn
        '
        Me.Context2DataGridViewTextBoxColumn.DataPropertyName = "context2"
        Me.Context2DataGridViewTextBoxColumn.HeaderText = "context2"
        Me.Context2DataGridViewTextBoxColumn.Name = "Context2DataGridViewTextBoxColumn"
        '
        'Context3DataGridViewTextBoxColumn
        '
        Me.Context3DataGridViewTextBoxColumn.DataPropertyName = "context3"
        Me.Context3DataGridViewTextBoxColumn.HeaderText = "context3"
        Me.Context3DataGridViewTextBoxColumn.Name = "Context3DataGridViewTextBoxColumn"
        '
        'StockNewBindingSource
        '
        Me.StockNewBindingSource.DataMember = "stockNew"
        Me.StockNewBindingSource.DataSource = Me.OptimizationDatabaseDataSetBindingSource
        '
        'OptimizationDatabaseDataSetBindingSource
        '
        Me.OptimizationDatabaseDataSetBindingSource.DataSource = Me.OptimizationDatabaseDataSet
        Me.OptimizationDatabaseDataSetBindingSource.Position = 0
        '
        'OptimizationDatabaseDataSet
        '
        Me.OptimizationDatabaseDataSet.DataSetName = "OptimizationDatabaseDataSet"
        Me.OptimizationDatabaseDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 44)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1152, 589)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Used Stock"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Location = New System.Drawing.Point(4, 44)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(1152, 589)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Cut Parts"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'StockNewTableAdapter
        '
        Me.StockNewTableAdapter.ClearBeforeFill = True
        '
        'Button1
        '
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.Button1.Location = New System.Drawing.Point(0, 655)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(1184, 56)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Save Changes"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DirectlyEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1184, 711)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "DirectlyEdit"
        Me.Text = "DirectlyEdit"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StockNewBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OptimizationDatabaseDataSetBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OptimizationDatabaseDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents OptimizationDatabaseDataSetBindingSource As BindingSource
    Friend WithEvents OptimizationDatabaseDataSet As OptimizationDatabaseDataSet
    Friend WithEvents StockNewBindingSource As BindingSource
    Friend WithEvents StockNewTableAdapter As OptimizationDatabaseDataSetTableAdapters.stockNewTableAdapter
    Friend WithEvents Button1 As Button
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
    Friend WithEvents StockID1DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents StockID2DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents StockID3DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DescriptionDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents ColorDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents SizeDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CountDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents InternalIDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Context1DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Context2DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Context3DataGridViewTextBoxColumn As DataGridViewTextBoxColumn
End Class
