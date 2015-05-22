<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.MSMain = New System.Windows.Forms.MenuStrip()
        Me.ReportNeutralizerToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductMasterReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SalesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WeeklySalesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SalesPerformanceReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DailySalesSummaryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DailySalesAndPaymentSummaryReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ARReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SummaryInvoiceAndSalesReturnReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StockToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DistributorStockReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DailyStockMutationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PromotionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ListOfPromotionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportGeneratorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AchievementByOutletToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.SalesPerformanceIncentiveReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MSMain.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MSMain
        '
        Me.MSMain.Font = New System.Drawing.Font("Calibri", 9.0!)
        Me.MSMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportNeutralizerToolStripMenuItem, Me.ReportGeneratorToolStripMenuItem})
        Me.MSMain.Location = New System.Drawing.Point(0, 0)
        Me.MSMain.Name = "MSMain"
        Me.MSMain.Size = New System.Drawing.Size(533, 24)
        Me.MSMain.TabIndex = 0
        Me.MSMain.Text = "MSMain"
        '
        'ReportNeutralizerToolStripMenuItem
        '
        Me.ReportNeutralizerToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MasterToolStripMenuItem, Me.SalesToolStripMenuItem, Me.StockToolStripMenuItem, Me.PromotionToolStripMenuItem})
        Me.ReportNeutralizerToolStripMenuItem.Name = "ReportNeutralizerToolStripMenuItem"
        Me.ReportNeutralizerToolStripMenuItem.Size = New System.Drawing.Size(119, 20)
        Me.ReportNeutralizerToolStripMenuItem.Text = "Report Neutralizer"
        '
        'MasterToolStripMenuItem
        '
        Me.MasterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProductToolStripMenuItem, Me.ProductMasterReportToolStripMenuItem})
        Me.MasterToolStripMenuItem.Name = "MasterToolStripMenuItem"
        Me.MasterToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.MasterToolStripMenuItem.Text = "Master"
        '
        'ProductToolStripMenuItem
        '
        Me.ProductToolStripMenuItem.Name = "ProductToolStripMenuItem"
        Me.ProductToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.ProductToolStripMenuItem.Text = "Outlet Master Report"
        '
        'ProductMasterReportToolStripMenuItem
        '
        Me.ProductMasterReportToolStripMenuItem.Name = "ProductMasterReportToolStripMenuItem"
        Me.ProductMasterReportToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.ProductMasterReportToolStripMenuItem.Text = "Product Master Report"
        '
        'SalesToolStripMenuItem
        '
        Me.SalesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.WeeklySalesToolStripMenuItem, Me.SalesPerformanceReportToolStripMenuItem, Me.SalesPerformanceIncentiveReportToolStripMenuItem, Me.DailySalesSummaryToolStripMenuItem, Me.DailySalesAndPaymentSummaryReportToolStripMenuItem, Me.ARReportToolStripMenuItem, Me.SummaryInvoiceAndSalesReturnReportToolStripMenuItem})
        Me.SalesToolStripMenuItem.Name = "SalesToolStripMenuItem"
        Me.SalesToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.SalesToolStripMenuItem.Text = "Sales"
        '
        'WeeklySalesToolStripMenuItem
        '
        Me.WeeklySalesToolStripMenuItem.Name = "WeeklySalesToolStripMenuItem"
        Me.WeeklySalesToolStripMenuItem.Size = New System.Drawing.Size(300, 22)
        Me.WeeklySalesToolStripMenuItem.Text = "Weekly Stock And Sales (LP3)"
        '
        'SalesPerformanceReportToolStripMenuItem
        '
        Me.SalesPerformanceReportToolStripMenuItem.Name = "SalesPerformanceReportToolStripMenuItem"
        Me.SalesPerformanceReportToolStripMenuItem.Size = New System.Drawing.Size(300, 22)
        Me.SalesPerformanceReportToolStripMenuItem.Text = "Sales Performance Report"
        '
        'DailySalesSummaryToolStripMenuItem
        '
        Me.DailySalesSummaryToolStripMenuItem.Name = "DailySalesSummaryToolStripMenuItem"
        Me.DailySalesSummaryToolStripMenuItem.Size = New System.Drawing.Size(300, 22)
        Me.DailySalesSummaryToolStripMenuItem.Text = "Daily Sales Summary"
        '
        'DailySalesAndPaymentSummaryReportToolStripMenuItem
        '
        Me.DailySalesAndPaymentSummaryReportToolStripMenuItem.Name = "DailySalesAndPaymentSummaryReportToolStripMenuItem"
        Me.DailySalesAndPaymentSummaryReportToolStripMenuItem.Size = New System.Drawing.Size(300, 22)
        Me.DailySalesAndPaymentSummaryReportToolStripMenuItem.Text = "Daily Sales And Payment Summary Report"
        '
        'ARReportToolStripMenuItem
        '
        Me.ARReportToolStripMenuItem.Name = "ARReportToolStripMenuItem"
        Me.ARReportToolStripMenuItem.Size = New System.Drawing.Size(300, 22)
        Me.ARReportToolStripMenuItem.Text = "AR Report"
        '
        'SummaryInvoiceAndSalesReturnReportToolStripMenuItem
        '
        Me.SummaryInvoiceAndSalesReturnReportToolStripMenuItem.Name = "SummaryInvoiceAndSalesReturnReportToolStripMenuItem"
        Me.SummaryInvoiceAndSalesReturnReportToolStripMenuItem.Size = New System.Drawing.Size(300, 22)
        Me.SummaryInvoiceAndSalesReturnReportToolStripMenuItem.Text = "Summary Invoice And Sales Return Report"
        '
        'StockToolStripMenuItem
        '
        Me.StockToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DistributorStockReportToolStripMenuItem, Me.DailyStockMutationToolStripMenuItem})
        Me.StockToolStripMenuItem.Name = "StockToolStripMenuItem"
        Me.StockToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.StockToolStripMenuItem.Text = "Stock"
        '
        'DistributorStockReportToolStripMenuItem
        '
        Me.DistributorStockReportToolStripMenuItem.Name = "DistributorStockReportToolStripMenuItem"
        Me.DistributorStockReportToolStripMenuItem.Size = New System.Drawing.Size(203, 22)
        Me.DistributorStockReportToolStripMenuItem.Text = "Distributor Stock Report"
        '
        'DailyStockMutationToolStripMenuItem
        '
        Me.DailyStockMutationToolStripMenuItem.Name = "DailyStockMutationToolStripMenuItem"
        Me.DailyStockMutationToolStripMenuItem.Size = New System.Drawing.Size(203, 22)
        Me.DailyStockMutationToolStripMenuItem.Text = "Daily Stock Mutation"
        '
        'PromotionToolStripMenuItem
        '
        Me.PromotionToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ListOfPromotionToolStripMenuItem})
        Me.PromotionToolStripMenuItem.Name = "PromotionToolStripMenuItem"
        Me.PromotionToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.PromotionToolStripMenuItem.Text = "Promotion"
        '
        'ListOfPromotionToolStripMenuItem
        '
        Me.ListOfPromotionToolStripMenuItem.Name = "ListOfPromotionToolStripMenuItem"
        Me.ListOfPromotionToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.ListOfPromotionToolStripMenuItem.Text = "List Of Promotion"
        '
        'ReportGeneratorToolStripMenuItem
        '
        Me.ReportGeneratorToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AchievementByOutletToolStripMenuItem})
        Me.ReportGeneratorToolStripMenuItem.Name = "ReportGeneratorToolStripMenuItem"
        Me.ReportGeneratorToolStripMenuItem.Size = New System.Drawing.Size(113, 20)
        Me.ReportGeneratorToolStripMenuItem.Text = "Report Generator"
        '
        'AchievementByOutletToolStripMenuItem
        '
        Me.AchievementByOutletToolStripMenuItem.Name = "AchievementByOutletToolStripMenuItem"
        Me.AchievementByOutletToolStripMenuItem.Size = New System.Drawing.Size(202, 22)
        Me.AchievementByOutletToolStripMenuItem.Text = "Achievement By Product"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(154, 96)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(206, 85)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'SalesPerformanceIncentiveReportToolStripMenuItem
        '
        Me.SalesPerformanceIncentiveReportToolStripMenuItem.Name = "SalesPerformanceIncentiveReportToolStripMenuItem"
        Me.SalesPerformanceIncentiveReportToolStripMenuItem.Size = New System.Drawing.Size(300, 22)
        Me.SalesPerformanceIncentiveReportToolStripMenuItem.Text = "Sales Performance Incentive Report"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(533, 316)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.MSMain)
        Me.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MSMain
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Leveredge Report Tools"
        Me.MSMain.ResumeLayout(False)
        Me.MSMain.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MSMain As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportNeutralizerToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportGeneratorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AchievementByOutletToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProductToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProductMasterReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents WeeklySalesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalesPerformanceReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DailySalesSummaryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StockToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DistributorStockReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DailyStockMutationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PromotionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ListOfPromotionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents ARReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SummaryInvoiceAndSalesReturnReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DailySalesAndPaymentSummaryReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalesPerformanceIncentiveReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
