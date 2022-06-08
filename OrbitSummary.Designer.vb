<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OrbitSummary
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(OrbitSummary))
        Me.OrbSum_OK = New System.Windows.Forms.Button()
        Me.OrbSum_TextBox = New System.Windows.Forms.TextBox()
        Me.ListSummary = New System.Windows.Forms.ListView()
        Me.Param = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Value = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Unit = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CopySummary = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'OrbSum_OK
        '
        Me.OrbSum_OK.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.OrbSum_OK.Location = New System.Drawing.Point(13, 509)
        Me.OrbSum_OK.Name = "OrbSum_OK"
        Me.OrbSum_OK.Size = New System.Drawing.Size(75, 23)
        Me.OrbSum_OK.TabIndex = 36
        Me.OrbSum_OK.Text = "Close"
        Me.OrbSum_OK.UseVisualStyleBackColor = True
        '
        'OrbSum_TextBox
        '
        Me.OrbSum_TextBox.Location = New System.Drawing.Point(13, 12)
        Me.OrbSum_TextBox.Multiline = True
        Me.OrbSum_TextBox.Name = "OrbSum_TextBox"
        Me.OrbSum_TextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.OrbSum_TextBox.Size = New System.Drawing.Size(221, 491)
        Me.OrbSum_TextBox.TabIndex = 37
        '
        'ListSummary
        '
        Me.ListSummary.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Param, Me.Value, Me.Unit})
        Me.ListSummary.Location = New System.Drawing.Point(387, 12)
        Me.ListSummary.Name = "ListSummary"
        Me.ListSummary.Size = New System.Drawing.Size(268, 490)
        Me.ListSummary.TabIndex = 38
        Me.ListSummary.UseCompatibleStateImageBehavior = False
        Me.ListSummary.View = System.Windows.Forms.View.Details
        '
        'Param
        '
        Me.Param.Text = ""
        Me.Param.Width = 92
        '
        'Value
        '
        Me.Value.Text = ""
        Me.Value.Width = 85
        '
        'Unit
        '
        Me.Unit.Text = ""
        Me.Unit.Width = 85
        '
        'CopySummary
        '
        Me.CopySummary.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.CopySummary.Location = New System.Drawing.Point(94, 509)
        Me.CopySummary.Name = "CopySummary"
        Me.CopySummary.Size = New System.Drawing.Size(140, 23)
        Me.CopySummary.TabIndex = 39
        Me.CopySummary.Text = "Copy To Clipboard"
        Me.CopySummary.UseVisualStyleBackColor = True
        '
        'OrbitSummary
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(251, 544)
        Me.Controls.Add(Me.CopySummary)
        Me.Controls.Add(Me.ListSummary)
        Me.Controls.Add(Me.OrbSum_TextBox)
        Me.Controls.Add(Me.OrbSum_OK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "OrbitSummary"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Orbit Summary"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OrbSum_OK As System.Windows.Forms.Button
    Friend WithEvents OrbSum_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents ListSummary As System.Windows.Forms.ListView
    Friend WithEvents Param As System.Windows.Forms.ColumnHeader
    Friend WithEvents Value As System.Windows.Forms.ColumnHeader
    Friend WithEvents Unit As System.Windows.Forms.ColumnHeader
    Friend WithEvents CopySummary As System.Windows.Forms.Button
End Class
