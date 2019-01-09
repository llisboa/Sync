<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Princ
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Princ))
        Me.fldLog = New System.Windows.Forms.CheckBox
        Me.btnExec = New System.Windows.Forms.Button
        Me.btnMais = New System.Windows.Forms.Button
        Me.lbl = New System.Windows.Forms.StatusBar
        Me.Button1 = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnAcertaNomes = New System.Windows.Forms.Button
        Me.btnExp = New System.Windows.Forms.Button
        Me.btnPrepIncluir = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtRecurso = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Opções = New System.Windows.Forms.Label
        Me.chkSemRegedit = New System.Windows.Forms.CheckBox
        Me.chkFechar = New System.Windows.Forms.CheckBox
        Me.txtApagarQ = New System.Windows.Forms.TextBox
        Me.txtFrom = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtSubject = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtArqLog = New System.Windows.Forms.TextBox
        Me.lblArqLog = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtServidor = New System.Windows.Forms.TextBox
        Me.txtEmailPara = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtHoras = New System.Windows.Forms.MaskedTextBox
        Me.chkAtiva = New System.Windows.Forms.CheckBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.DeMinuto = New System.Windows.Forms.Timer(Me.components)
        Me.grdPrinc = New System.Windows.Forms.DataGridView
        Me.colDiretorio = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReplica = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colSubDir = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Ordem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Momento = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.grdPrinc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'fldLog
        '
        Me.fldLog.AutoSize = True
        Me.fldLog.Checked = True
        Me.fldLog.CheckState = System.Windows.Forms.CheckState.Checked
        Me.fldLog.Location = New System.Drawing.Point(420, 59)
        Me.fldLog.Name = "fldLog"
        Me.fldLog.Size = New System.Drawing.Size(44, 17)
        Me.fldLog.TabIndex = 18
        Me.fldLog.Text = "Log"
        Me.fldLog.UseVisualStyleBackColor = True
        '
        'btnExec
        '
        Me.btnExec.Location = New System.Drawing.Point(9, 18)
        Me.btnExec.Name = "btnExec"
        Me.btnExec.Size = New System.Drawing.Size(69, 21)
        Me.btnExec.TabIndex = 0
        Me.btnExec.Text = "Agora"
        Me.btnExec.UseVisualStyleBackColor = True
        '
        'btnMais
        '
        Me.btnMais.Location = New System.Drawing.Point(554, 76)
        Me.btnMais.Name = "btnMais"
        Me.btnMais.Size = New System.Drawing.Size(115, 21)
        Me.btnMais.TabIndex = 2
        Me.btnMais.Text = "Ver Log"
        Me.btnMais.UseVisualStyleBackColor = True
        Me.btnMais.Visible = False
        '
        'lbl
        '
        Me.lbl.Location = New System.Drawing.Point(0, 207)
        Me.lbl.Name = "lbl"
        Me.lbl.Size = New System.Drawing.Size(681, 30)
        Me.lbl.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(554, 20)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(115, 21)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Ver Tamanho"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnAcertaNomes)
        Me.Panel1.Controls.Add(Me.btnExp)
        Me.Panel1.Controls.Add(Me.btnPrepIncluir)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.btnMais)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.lbl)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 269)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(681, 237)
        Me.Panel1.TabIndex = 0
        '
        'btnAcertaNomes
        '
        Me.btnAcertaNomes.Location = New System.Drawing.Point(554, 162)
        Me.btnAcertaNomes.Name = "btnAcertaNomes"
        Me.btnAcertaNomes.Size = New System.Drawing.Size(115, 37)
        Me.btnAcertaNomes.TabIndex = 17
        Me.btnAcertaNomes.Text = "Acerta Nomes de Diretórios"
        Me.btnAcertaNomes.UseVisualStyleBackColor = True
        '
        'btnExp
        '
        Me.btnExp.Location = New System.Drawing.Point(554, 132)
        Me.btnExp.Name = "btnExp"
        Me.btnExp.Size = New System.Drawing.Size(115, 23)
        Me.btnExp.TabIndex = 6
        Me.btnExp.Text = "Exp Linha BAT"
        Me.btnExp.UseVisualStyleBackColor = True
        '
        'btnPrepIncluir
        '
        Me.btnPrepIncluir.Location = New System.Drawing.Point(554, 104)
        Me.btnPrepIncluir.Name = "btnPrepIncluir"
        Me.btnPrepIncluir.Size = New System.Drawing.Size(115, 21)
        Me.btnPrepIncluir.TabIndex = 3
        Me.btnPrepIncluir.Text = "Incluir"
        Me.btnPrepIncluir.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtRecurso)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Opções)
        Me.GroupBox1.Controls.Add(Me.chkSemRegedit)
        Me.GroupBox1.Controls.Add(Me.chkFechar)
        Me.GroupBox1.Controls.Add(Me.txtApagarQ)
        Me.GroupBox1.Controls.Add(Me.txtFrom)
        Me.GroupBox1.Controls.Add(Me.fldLog)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtSubject)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtArqLog)
        Me.GroupBox1.Controls.Add(Me.lblArqLog)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtServidor)
        Me.GroupBox1.Controls.Add(Me.txtEmailPara)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtHoras)
        Me.GroupBox1.Controls.Add(Me.chkAtiva)
        Me.GroupBox1.Controls.Add(Me.btnExec)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 11)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(522, 190)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Executar"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(266, 67)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(126, 26)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Excluir quando Encontrar" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(exclui dos dois lados):" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'txtRecurso
        '
        Me.txtRecurso.Location = New System.Drawing.Point(94, 130)
        Me.txtRecurso.Name = "txtRecurso"
        Me.txtRecurso.Size = New System.Drawing.Size(158, 20)
        Me.txtRecurso.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(10, 133)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(50, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Recurso:"
        '
        'Opções
        '
        Me.Opções.AutoSize = True
        Me.Opções.Location = New System.Drawing.Point(405, 18)
        Me.Opções.Name = "Opções"
        Me.Opções.Size = New System.Drawing.Size(44, 13)
        Me.Opções.TabIndex = 16
        Me.Opções.Text = "Opções"
        '
        'chkSemRegedit
        '
        Me.chkSemRegedit.AutoSize = True
        Me.chkSemRegedit.Location = New System.Drawing.Point(420, 39)
        Me.chkSemRegedit.Name = "chkSemRegedit"
        Me.chkSemRegedit.Size = New System.Drawing.Size(87, 17)
        Me.chkSemRegedit.TabIndex = 17
        Me.chkSemRegedit.Text = "Sem Regedit"
        Me.chkSemRegedit.UseVisualStyleBackColor = True
        '
        'chkFechar
        '
        Me.chkFechar.AutoSize = True
        Me.chkFechar.Location = New System.Drawing.Point(420, 78)
        Me.chkFechar.Name = "chkFechar"
        Me.chkFechar.Size = New System.Drawing.Size(74, 30)
        Me.chkFechar.TabIndex = 19
        Me.chkFechar.Text = "Fechar ao" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Concluir"
        Me.chkFechar.UseVisualStyleBackColor = True
        '
        'txtApagarQ
        '
        Me.txtApagarQ.Location = New System.Drawing.Point(278, 97)
        Me.txtApagarQ.Multiline = True
        Me.txtApagarQ.Name = "txtApagarQ"
        Me.txtApagarQ.Size = New System.Drawing.Size(121, 47)
        Me.txtApagarQ.TabIndex = 9
        Me.txtApagarQ.Text = "Thumbs.db"
        '
        'txtFrom
        '
        Me.txtFrom.Location = New System.Drawing.Point(94, 94)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(158, 20)
        Me.txtFrom.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(10, 97)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(33, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "From:"
        '
        'txtSubject
        '
        Me.txtSubject.Location = New System.Drawing.Point(94, 72)
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.Size = New System.Drawing.Size(158, 20)
        Me.txtSubject.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 75)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Subject:"
        '
        'txtArqLog
        '
        Me.txtArqLog.Location = New System.Drawing.Point(94, 153)
        Me.txtArqLog.Name = "txtArqLog"
        Me.txtArqLog.Size = New System.Drawing.Size(336, 20)
        Me.txtArqLog.TabIndex = 13
        '
        'lblArqLog
        '
        Me.lblArqLog.AutoSize = True
        Me.lblArqLog.Location = New System.Drawing.Point(10, 156)
        Me.lblArqLog.Name = "lblArqLog"
        Me.lblArqLog.Size = New System.Drawing.Size(76, 13)
        Me.lblArqLog.TabIndex = 12
        Me.lblArqLog.Text = "Gravar log em:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(266, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(82, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Servidor SMTP:"
        '
        'txtServidor
        '
        Me.txtServidor.Location = New System.Drawing.Point(290, 36)
        Me.txtServidor.Name = "txtServidor"
        Me.txtServidor.Size = New System.Drawing.Size(92, 20)
        Me.txtServidor.TabIndex = 15
        Me.txtServidor.Text = "smtpi"
        '
        'txtEmailPara
        '
        Me.txtEmailPara.Location = New System.Drawing.Point(94, 48)
        Me.txtEmailPara.Name = "txtEmailPara"
        Me.txtEmailPara.Size = New System.Drawing.Size(158, 20)
        Me.txtEmailPara.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Enviar log para:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(89, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(19, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "ou"
        '
        'txtHoras
        '
        Me.txtHoras.Location = New System.Drawing.Point(182, 19)
        Me.txtHoras.Mask = "00:00"
        Me.txtHoras.Name = "txtHoras"
        Me.txtHoras.Size = New System.Drawing.Size(38, 20)
        Me.txtHoras.TabIndex = 3
        Me.txtHoras.ValidatingType = GetType(Date)
        '
        'chkAtiva
        '
        Me.chkAtiva.AutoSize = True
        Me.chkAtiva.Location = New System.Drawing.Point(121, 17)
        Me.chkAtiva.Name = "chkAtiva"
        Me.chkAtiva.Size = New System.Drawing.Size(136, 17)
        Me.chkAtiva.TabIndex = 2
        Me.chkAtiva.Text = "Auto às                horas"
        Me.chkAtiva.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(554, 48)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(115, 21)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Apagar Igualdades"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'DeMinuto
        '
        Me.DeMinuto.Interval = 1000
        '
        'grdPrinc
        '
        Me.grdPrinc.AllowUserToOrderColumns = True
        Me.grdPrinc.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.grdPrinc.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Me.grdPrinc.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdPrinc.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colDiretorio, Me.colReplica, Me.colSubDir, Me.Ordem, Me.Momento})
        Me.grdPrinc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdPrinc.Location = New System.Drawing.Point(0, 0)
        Me.grdPrinc.Name = "grdPrinc"
        Me.grdPrinc.Size = New System.Drawing.Size(681, 269)
        Me.grdPrinc.TabIndex = 11
        '
        'colDiretorio
        '
        Me.colDiretorio.HeaderText = "Diretório"
        Me.colDiretorio.Name = "colDiretorio"
        '
        'colReplica
        '
        Me.colReplica.HeaderText = "Réplica"
        Me.colReplica.Name = "colReplica"
        '
        'colSubDir
        '
        Me.colSubDir.HeaderText = "Inclui Sub Dir"
        Me.colSubDir.Name = "colSubDir"
        '
        'Ordem
        '
        Me.Ordem.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Ordem.HeaderText = "Ordem"
        Me.Ordem.Name = "Ordem"
        '
        'Momento
        '
        Me.Momento.HeaderText = "Momento"
        Me.Momento.Name = "Momento"
        Me.Momento.ReadOnly = True
        Me.Momento.Visible = False
        '
        'Panel2
        '
        Me.Panel2.AutoSize = True
        Me.Panel2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Panel2.Controls.Add(Me.grdPrinc)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(681, 269)
        Me.Panel2.TabIndex = 14
        '
        'Princ
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(681, 506)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Princ"
        Me.Text = "Sync - V04.00 - Sincronização de Diretórios - Intercraft Solutions - 2011"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.grdPrinc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents fldLog As System.Windows.Forms.CheckBox
    Friend WithEvents btnExec As System.Windows.Forms.Button
    Friend WithEvents btnMais As System.Windows.Forms.Button
    Friend WithEvents lbl As System.Windows.Forms.StatusBar
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtHoras As System.Windows.Forms.MaskedTextBox
    Friend WithEvents chkAtiva As System.Windows.Forms.CheckBox
    Friend WithEvents DeMinuto As System.Windows.Forms.Timer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtServidor As System.Windows.Forms.TextBox
    Friend WithEvents btnPrepIncluir As System.Windows.Forms.Button
    Friend WithEvents btnExp As System.Windows.Forms.Button
    Friend WithEvents chkFechar As System.Windows.Forms.CheckBox
    Friend WithEvents chkSemRegedit As System.Windows.Forms.CheckBox
    Friend WithEvents txtArqLog As System.Windows.Forms.TextBox
    Friend WithEvents lblArqLog As System.Windows.Forms.Label
    Friend WithEvents txtEmailPara As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Opções As System.Windows.Forms.Label
    Friend WithEvents txtRecurso As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnAcertaNomes As System.Windows.Forms.Button
    Friend WithEvents grdPrinc As System.Windows.Forms.DataGridView
    Friend WithEvents colDiretorio As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReplica As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colSubDir As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Ordem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Momento As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtApagarQ As System.Windows.Forms.TextBox

End Class
