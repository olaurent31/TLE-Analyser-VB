<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChartForm
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ChartForm))
        Me.CHART_TAB = New System.Windows.Forms.TabControl()
        Me.CHART_SMATAB = New System.Windows.Forms.TabPage()
        Me.CHART_ECCTAB = New System.Windows.Forms.TabPage()
        Me.CHART_INCTAB = New System.Windows.Forms.TabPage()
        Me.CHART_MALTTAB = New System.Windows.Forms.TabPage()
        Me.CHART_APATAB = New System.Windows.Forms.TabPage()
        Me.CHART_PEATAB = New System.Windows.Forms.TabPage()
        Me.CHART_LNGTAB = New System.Windows.Forms.TabPage()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.CHART_APA_CB = New System.Windows.Forms.CheckBox()
        Me.CHART_PEA_CB = New System.Windows.Forms.CheckBox()
        Me.CHART_LNG_CB = New System.Windows.Forms.CheckBox()
        Me.CHART_SMA_CB = New System.Windows.Forms.CheckBox()
        Me.CHART_ECC_CB = New System.Windows.Forms.CheckBox()
        Me.CHART_INC_CB = New System.Windows.Forms.CheckBox()
        Me.CHART_MALT_CB = New System.Windows.Forms.CheckBox()
        Me.ChartButton = New System.Windows.Forms.Button()
        Me.Chart_Days = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CloseButton = New System.Windows.Forms.Button()
        Me.ChartXvalue = New System.Windows.Forms.ComboBox()
        Me.CHART_TAB.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.SuspendLayout()
        '
        'CHART_TAB
        '
        Me.CHART_TAB.Controls.Add(Me.CHART_SMATAB)
        Me.CHART_TAB.Controls.Add(Me.CHART_ECCTAB)
        Me.CHART_TAB.Controls.Add(Me.CHART_INCTAB)
        Me.CHART_TAB.Controls.Add(Me.CHART_MALTTAB)
        Me.CHART_TAB.Controls.Add(Me.CHART_APATAB)
        Me.CHART_TAB.Controls.Add(Me.CHART_PEATAB)
        Me.CHART_TAB.Controls.Add(Me.CHART_LNGTAB)
        Me.CHART_TAB.Location = New System.Drawing.Point(169, 12)
        Me.CHART_TAB.Name = "CHART_TAB"
        Me.CHART_TAB.SelectedIndex = 0
        Me.CHART_TAB.Size = New System.Drawing.Size(633, 450)
        Me.CHART_TAB.TabIndex = 0
        '
        'CHART_SMATAB
        '
        Me.CHART_SMATAB.BackColor = System.Drawing.SystemColors.Control
        Me.CHART_SMATAB.Location = New System.Drawing.Point(4, 22)
        Me.CHART_SMATAB.Name = "CHART_SMATAB"
        Me.CHART_SMATAB.Padding = New System.Windows.Forms.Padding(3)
        Me.CHART_SMATAB.Size = New System.Drawing.Size(625, 424)
        Me.CHART_SMATAB.TabIndex = 0
        Me.CHART_SMATAB.Text = "SMA"
        '
        'CHART_ECCTAB
        '
        Me.CHART_ECCTAB.BackColor = System.Drawing.SystemColors.Control
        Me.CHART_ECCTAB.Location = New System.Drawing.Point(4, 22)
        Me.CHART_ECCTAB.Name = "CHART_ECCTAB"
        Me.CHART_ECCTAB.Padding = New System.Windows.Forms.Padding(3)
        Me.CHART_ECCTAB.Size = New System.Drawing.Size(625, 424)
        Me.CHART_ECCTAB.TabIndex = 1
        Me.CHART_ECCTAB.Text = "ECC"
        '
        'CHART_INCTAB
        '
        Me.CHART_INCTAB.BackColor = System.Drawing.SystemColors.Control
        Me.CHART_INCTAB.Location = New System.Drawing.Point(4, 22)
        Me.CHART_INCTAB.Name = "CHART_INCTAB"
        Me.CHART_INCTAB.Padding = New System.Windows.Forms.Padding(3)
        Me.CHART_INCTAB.Size = New System.Drawing.Size(625, 424)
        Me.CHART_INCTAB.TabIndex = 2
        Me.CHART_INCTAB.Text = "INC"
        '
        'CHART_MALTTAB
        '
        Me.CHART_MALTTAB.BackColor = System.Drawing.SystemColors.Control
        Me.CHART_MALTTAB.Location = New System.Drawing.Point(4, 22)
        Me.CHART_MALTTAB.Name = "CHART_MALTTAB"
        Me.CHART_MALTTAB.Padding = New System.Windows.Forms.Padding(3)
        Me.CHART_MALTTAB.Size = New System.Drawing.Size(625, 424)
        Me.CHART_MALTTAB.TabIndex = 4
        Me.CHART_MALTTAB.Text = "MALT"
        '
        'CHART_APATAB
        '
        Me.CHART_APATAB.BackColor = System.Drawing.SystemColors.Control
        Me.CHART_APATAB.Location = New System.Drawing.Point(4, 22)
        Me.CHART_APATAB.Name = "CHART_APATAB"
        Me.CHART_APATAB.Padding = New System.Windows.Forms.Padding(3)
        Me.CHART_APATAB.Size = New System.Drawing.Size(625, 424)
        Me.CHART_APATAB.TabIndex = 5
        Me.CHART_APATAB.Text = "APA"
        '
        'CHART_PEATAB
        '
        Me.CHART_PEATAB.BackColor = System.Drawing.SystemColors.Control
        Me.CHART_PEATAB.Location = New System.Drawing.Point(4, 22)
        Me.CHART_PEATAB.Name = "CHART_PEATAB"
        Me.CHART_PEATAB.Padding = New System.Windows.Forms.Padding(3)
        Me.CHART_PEATAB.Size = New System.Drawing.Size(625, 424)
        Me.CHART_PEATAB.TabIndex = 6
        Me.CHART_PEATAB.Text = "PEA"
        '
        'CHART_LNGTAB
        '
        Me.CHART_LNGTAB.BackColor = System.Drawing.SystemColors.Control
        Me.CHART_LNGTAB.Location = New System.Drawing.Point(4, 22)
        Me.CHART_LNGTAB.Name = "CHART_LNGTAB"
        Me.CHART_LNGTAB.Padding = New System.Windows.Forms.Padding(3)
        Me.CHART_LNGTAB.Size = New System.Drawing.Size(625, 424)
        Me.CHART_LNGTAB.TabIndex = 3
        Me.CHART_LNGTAB.Text = "LNG"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.CHART_APA_CB)
        Me.GroupBox7.Controls.Add(Me.CHART_PEA_CB)
        Me.GroupBox7.Controls.Add(Me.CHART_LNG_CB)
        Me.GroupBox7.Controls.Add(Me.CHART_SMA_CB)
        Me.GroupBox7.Controls.Add(Me.CHART_ECC_CB)
        Me.GroupBox7.Controls.Add(Me.CHART_INC_CB)
        Me.GroupBox7.Controls.Add(Me.CHART_MALT_CB)
        Me.GroupBox7.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(151, 200)
        Me.GroupBox7.TabIndex = 1
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Datas Vs Elapsed Time"
        '
        'CHART_APA_CB
        '
        Me.CHART_APA_CB.AutoSize = True
        Me.CHART_APA_CB.Checked = True
        Me.CHART_APA_CB.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CHART_APA_CB.Location = New System.Drawing.Point(35, 120)
        Me.CHART_APA_CB.Name = "CHART_APA_CB"
        Me.CHART_APA_CB.Size = New System.Drawing.Size(47, 17)
        Me.CHART_APA_CB.TabIndex = 5
        Me.CHART_APA_CB.Text = "APA"
        Me.CHART_APA_CB.UseVisualStyleBackColor = True
        '
        'CHART_PEA_CB
        '
        Me.CHART_PEA_CB.AutoSize = True
        Me.CHART_PEA_CB.Checked = True
        Me.CHART_PEA_CB.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CHART_PEA_CB.Location = New System.Drawing.Point(35, 143)
        Me.CHART_PEA_CB.Name = "CHART_PEA_CB"
        Me.CHART_PEA_CB.Size = New System.Drawing.Size(47, 17)
        Me.CHART_PEA_CB.TabIndex = 6
        Me.CHART_PEA_CB.Text = "PEA"
        Me.CHART_PEA_CB.UseVisualStyleBackColor = True
        '
        'CHART_LNG_CB
        '
        Me.CHART_LNG_CB.AutoSize = True
        Me.CHART_LNG_CB.Location = New System.Drawing.Point(35, 166)
        Me.CHART_LNG_CB.Name = "CHART_LNG_CB"
        Me.CHART_LNG_CB.Size = New System.Drawing.Size(48, 17)
        Me.CHART_LNG_CB.TabIndex = 4
        Me.CHART_LNG_CB.Text = "LNG"
        Me.CHART_LNG_CB.UseVisualStyleBackColor = True
        '
        'CHART_SMA_CB
        '
        Me.CHART_SMA_CB.AutoSize = True
        Me.CHART_SMA_CB.Checked = True
        Me.CHART_SMA_CB.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CHART_SMA_CB.Location = New System.Drawing.Point(35, 28)
        Me.CHART_SMA_CB.Name = "CHART_SMA_CB"
        Me.CHART_SMA_CB.Size = New System.Drawing.Size(49, 17)
        Me.CHART_SMA_CB.TabIndex = 0
        Me.CHART_SMA_CB.Text = "SMA"
        Me.CHART_SMA_CB.UseVisualStyleBackColor = True
        '
        'CHART_ECC_CB
        '
        Me.CHART_ECC_CB.AutoSize = True
        Me.CHART_ECC_CB.Checked = True
        Me.CHART_ECC_CB.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CHART_ECC_CB.Location = New System.Drawing.Point(35, 51)
        Me.CHART_ECC_CB.Name = "CHART_ECC_CB"
        Me.CHART_ECC_CB.Size = New System.Drawing.Size(47, 17)
        Me.CHART_ECC_CB.TabIndex = 1
        Me.CHART_ECC_CB.Text = "ECC"
        Me.CHART_ECC_CB.UseVisualStyleBackColor = True
        '
        'CHART_INC_CB
        '
        Me.CHART_INC_CB.AutoSize = True
        Me.CHART_INC_CB.Checked = True
        Me.CHART_INC_CB.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CHART_INC_CB.Location = New System.Drawing.Point(35, 74)
        Me.CHART_INC_CB.Name = "CHART_INC_CB"
        Me.CHART_INC_CB.Size = New System.Drawing.Size(44, 17)
        Me.CHART_INC_CB.TabIndex = 2
        Me.CHART_INC_CB.Text = "INC"
        Me.CHART_INC_CB.UseVisualStyleBackColor = True
        '
        'CHART_MALT_CB
        '
        Me.CHART_MALT_CB.AutoSize = True
        Me.CHART_MALT_CB.Checked = True
        Me.CHART_MALT_CB.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CHART_MALT_CB.Location = New System.Drawing.Point(35, 97)
        Me.CHART_MALT_CB.Name = "CHART_MALT_CB"
        Me.CHART_MALT_CB.Size = New System.Drawing.Size(55, 17)
        Me.CHART_MALT_CB.TabIndex = 3
        Me.CHART_MALT_CB.Text = "MALT"
        Me.CHART_MALT_CB.UseVisualStyleBackColor = True
        '
        'ChartButton
        '
        Me.ChartButton.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ChartButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!)
        Me.ChartButton.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ChartButton.Location = New System.Drawing.Point(22, 272)
        Me.ChartButton.Name = "ChartButton"
        Me.ChartButton.Size = New System.Drawing.Size(131, 25)
        Me.ChartButton.TabIndex = 100
        Me.ChartButton.Text = "GENERATE"
        Me.ChartButton.UseVisualStyleBackColor = False
        '
        'Chart_Days
        '
        Me.Chart_Days.Location = New System.Drawing.Point(22, 246)
        Me.Chart_Days.MaxLength = 4
        Me.Chart_Days.Name = "Chart_Days"
        Me.Chart_Days.Size = New System.Drawing.Size(38, 20)
        Me.Chart_Days.TabIndex = 7
        Me.Chart_Days.Text = "24"
        Me.Chart_Days.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(55, 229)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(69, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Generate on "
        '
        'CloseButton
        '
        Me.CloseButton.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.CloseButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!)
        Me.CloseButton.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.CloseButton.Location = New System.Drawing.Point(22, 437)
        Me.CloseButton.Name = "CloseButton"
        Me.CloseButton.Size = New System.Drawing.Size(131, 25)
        Me.CloseButton.TabIndex = 101
        Me.CloseButton.Text = "CLOSE"
        Me.CloseButton.UseVisualStyleBackColor = False
        '
        'ChartXvalue
        '
        Me.ChartXvalue.FormattingEnabled = True
        Me.ChartXvalue.Items.AddRange(New Object() {"min", "hr", "period"})
        Me.ChartXvalue.Location = New System.Drawing.Point(66, 245)
        Me.ChartXvalue.Name = "ChartXvalue"
        Me.ChartXvalue.Size = New System.Drawing.Size(87, 21)
        Me.ChartXvalue.TabIndex = 102
        '
        'ChartForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(814, 474)
        Me.Controls.Add(Me.ChartXvalue)
        Me.Controls.Add(Me.CloseButton)
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.ChartButton)
        Me.Controls.Add(Me.CHART_TAB)
        Me.Controls.Add(Me.Chart_Days)
        Me.Controls.Add(Me.Label5)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ChartForm"
        Me.Text = "TLE Analyser - Charts"
        Me.CHART_TAB.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CHART_TAB As System.Windows.Forms.TabControl
    Friend WithEvents CHART_SMATAB As System.Windows.Forms.TabPage
    Friend WithEvents CHART_ECCTAB As System.Windows.Forms.TabPage
    Friend WithEvents CHART_INCTAB As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents CHART_SMA_CB As System.Windows.Forms.CheckBox
    Friend WithEvents CHART_ECC_CB As System.Windows.Forms.CheckBox
    Friend WithEvents CHART_INC_CB As System.Windows.Forms.CheckBox
    Friend WithEvents ChartButton As System.Windows.Forms.Button
    Friend WithEvents Chart_Days As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CHART_MALT_CB As System.Windows.Forms.CheckBox
    Friend WithEvents CHART_LNGTAB As System.Windows.Forms.TabPage
    Friend WithEvents CHART_MALTTAB As System.Windows.Forms.TabPage
    Friend WithEvents CHART_LNG_CB As System.Windows.Forms.CheckBox
    Friend WithEvents CloseButton As System.Windows.Forms.Button
    Friend WithEvents ChartXvalue As System.Windows.Forms.ComboBox
    Friend WithEvents CHART_APATAB As System.Windows.Forms.TabPage
    Friend WithEvents CHART_PEATAB As System.Windows.Forms.TabPage
    Friend WithEvents CHART_APA_CB As System.Windows.Forms.CheckBox
    Friend WithEvents CHART_PEA_CB As System.Windows.Forms.CheckBox
End Class
