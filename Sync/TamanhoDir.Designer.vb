<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TamanhoDir
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TamanhoDir))
        Me.btnCalc = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Arv = New System.Windows.Forms.TreeView
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCalc
        '
        Me.btnCalc.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCalc.Location = New System.Drawing.Point(487, 5)
        Me.btnCalc.Name = "btnCalc"
        Me.btnCalc.Size = New System.Drawing.Size(75, 27)
        Me.btnCalc.TabIndex = 9
        Me.btnCalc.Text = "Calcular"
        Me.btnCalc.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnCalc)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 358)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Padding = New System.Windows.Forms.Padding(0, 5, 10, 5)
        Me.Panel1.Size = New System.Drawing.Size(572, 37)
        Me.Panel1.TabIndex = 10
        '
        'Arv
        '
        Me.Arv.BackColor = System.Drawing.SystemColors.Info
        Me.Arv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Arv.Location = New System.Drawing.Point(0, 0)
        Me.Arv.Name = "Arv"
        Me.Arv.Size = New System.Drawing.Size(572, 358)
        Me.Arv.TabIndex = 0
        '
        'TamanhoDir
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(572, 395)
        Me.Controls.Add(Me.Arv)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "TamanhoDir"
        Me.Text = "LISTA DE TAMANHO DE DIRETÓRIOS..."
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCalc As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Arv As System.Windows.Forms.TreeView
End Class
