Imports Sync.Icraft.IcftBase

Public Class Princ

    Public ExecImed As Boolean = False
    Public SemConfirm As Boolean = False

    Private Sub btnMais_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMais.Click
        Try
            Mais.ShowDialog()
            Mais.Focus()
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Mais Click")
        End Try
    End Sub

    Dim Replica As Icraft.IcftBase.DirReplica

    Private Sub btnExec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExec.Click
        Try
            btnExec.Enabled = False
            DeMinuto.Enabled = False
            chkAtiva.Enabled = False

            If Not TypeOf sender Is Timer And Not SemConfirm Then
                If MsgBox("Certeza de alterar, podendo até excluir, arquivos e diretórios REPLICADOS???", MsgBoxStyle.Critical + MsgBoxStyle.OkCancel + MsgBoxStyle.DefaultButton2) = MsgBoxResult.Cancel Then
                    Exit Sub
                End If
            End If

            Dim MomentoIni As Date = Now

            Dim StrCols As String = ""
            For Each Linha As System.Windows.Forms.DataGridViewRow In From x As System.Windows.Forms.DataGridViewRow In grdPrinc.Rows Order By Val(x.Cells(3).Value), x.Cells(4).Value Descending Select x
                If NZV(Linha.Cells(3).Value, 0) <> 0 Then
                    For x As Integer = 0 To 2
                        StrCols &= IIf(StrCols <> "", ";", "") & Replace(Linha.Cells(x).Value, ";", "")
                    Next
                End If
            Next

            If Not chkSemRegedit.Checked Then
                Icraft.IcftBase.RegAplKey("Sync", "Prog") = StrCols
                Icraft.IcftBase.RegAplKey("Sync", "Log") = fldLog.Checked
            End If

            Mais.txtResult.Text = ""

            Dim OrigemAnterior As String = ""
            For Each ROW As System.Windows.Forms.DataGridViewRow In From x As System.Windows.Forms.DataGridViewRow In grdPrinc.Rows Order By Val(x.Cells(3).Value), x.Cells(4).Value Descending Select x
                Dim Origem As String = NZV(ROW.Cells(0).Value, OrigemAnterior)
                Dim Repl As String = NZ(ROW.Cells(1).Value, "")

                If Origem <> "" And Repl <> "" Then
                    Replica = New Icraft.IcftBase.DirReplica(Origem, Repl, ROW.Cells(2).Value, txtApagarQ.Text)
                    AddHandler Replica.NotificaStatus, AddressOf Notifica
                    Replica.LogDetalhado = fldLog.Checked
                    Replica.Executa()
                    Mais.txtResult.Text &= IIf(Mais.txtResult.Text <> "", vbCrLf, "") & Replica.Log.ToString
                End If

                OrigemAnterior = Origem
            Next

            Dim MomentoFim As Date = Now
            Mais.txtResult.Text = "Recurso: " & txtSubject.Text & vbCrLf & "Início:  " & Format(MomentoIni, "yyyy-MM-dd HH:mm:ss") & vbCrLf & "Término: " & Format(MomentoFim, "yyyy-MM-dd HH:mm:ss") & vbCrLf & "Duração: " & ExibeSegs(DateDiff(DateInterval.Second, MomentoIni, MomentoFim), ExibeSegsOpc.hh_mm_ss) & vbCrLf & vbCrLf & Mais.txtResult.Text

            Dim STREMAIL As String = Trim(Icraft.IcftBase.EmailStr(txtEmailPara.Text))
            If STREMAIL <> "" Then
                EnviaEmail(NZV(txtFrom.Text, "<suporte@icraft.com.br>"), STREMAIL, "Sync - Log - " & Environ("COMPUTERNAME") & " - " & Environ("USERNAME") & IIf(txtSubject.Text <> "", " - " & txtSubject.Text, ""), "<div style='font-family:arial;font-size:8pt'>Resultado da sincronização:<ul><li>" & System.Text.RegularExpressions.Regex.Replace(Mais.txtResult.Text.Trim(vbCrLf).Replace(vbCrLf, "</li><li>"), "(?is)(\[erro\]|\[falha\])", "<span style='background-color:yellow'>$1</span>") & "</li></ul></div>", , txtServidor.Text)
            End If

            If txtArqLog.Text <> "" Then
                If System.IO.File.Exists(txtArqLog.Text) Then
                    Kill(txtArqLog.Text)
                End If
                Icraft.IcftBase.GravaLog(txtArqLog.Text, Mais.txtResult.Text)
            End If

            If chkFechar.Checked Then
                End
            End If

            btnMais.Visible = True
            chkAtiva_CheckedChanged(sender, e)

            If Not TypeOf sender Is Timer Then
                Mais.Visible = True
                MsgBox(Replica.Status)
            End If

        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Botão de Execução")
        End Try
    End Sub

    Sub Notifica()
        Try
            lbl.Text = Format(Now, "yyyy-MM-dd HH:mm") & " - " & Replica.Status
            Application.DoEvents()
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Notificando Condição")
        End Try
    End Sub

    Private Function SemAspas(ByVal Texto As String) As String
        Return Texto.Trim("""", Chr(147), Chr(148))
    End Function

    Private Sub Princ_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Text = My.Application.Info.ProductName & " - V" & My.Application.Info.Version.ToString

            Dim Args As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs()
            Dim Tab As String = ""
            Dim Diretorio As String = ""
            Dim IncluiSubDir As Boolean = False

            For z As Integer = 0 To Args.Count - 1 Step 2
                Dim Coma As String = SemAspas(Args(z)).ToLower

                Select Case Coma
                    Case "-recurso"
                        txtRecurso.Text = SemAspas(Args(z + 1))
                    Case "-auto"
                        txtHoras.Text = SemAspas(Args(z + 1))
                        chkAtiva.Checked = True
                    Case "-email"
                        txtEmailPara.Text = SemAspas(Args(z + 1))
                    Case "-smtp"
                        txtServidor.Text = SemAspas(Args(z + 1))
                    Case "-fecharaoconcluir"
                        chkFechar.Checked = True
                        z -= 1
                    Case "-execimed"
                        ExecImed = True
                        z -= 1
                    Case "-dir"
                        Diretorio = SemAspas(Args(z + 1))
                        IncluiSubDir = True
                    Case "-dirsemsub"
                        Diretorio = SemAspas(Args(z + 1))
                        IncluiSubDir = False
                    Case "-repl"
                        Tab &= IIf(Tab <> "", ";", "") & Diretorio & ";" & SemAspas(Args(z + 1)) & ";" & IIf(IncluiSubDir, "True", "False")
                    Case "-noregedit"
                        chkSemRegedit.Checked = True
                        z -= 1
                    Case "-arqlog"
                        txtArqLog.Text = SemAspas(Args(z + 1))
                    Case "-from"
                        txtFrom.Text = SemAspas(Args(z + 1))
                    Case "-subject"
                        txtSubject.Text = SemAspas(Args(z + 1))
                    Case "-semconfirm"
                        SemConfirm = True
                        z -= 1
                    Case "-help"
                        Dim Msg As String = "Sync - V04.00 - Programa de sincronização de arquivos e diretórios" & vbCrLf
                        Msg &= "     Modo de usar (exemplo)......................................." & vbCrLf
                        Msg &= "     mostrar help:               -help" & vbCrLf
                        Msg &= "     executar automaticamente:   -auto ""04:30""" & vbCrLf
                        Msg &= "     enviar email no final:      -email ""email@icraft.com.br""" & vbCrLf
                        Msg &= "     utilizar smtp:              -smtp ""smtpi.icraft.com.br""" & vbCrLf
                        Msg &= "     diretório incluindo subs:   -dir ""c:\origem""" & vbCrLf
                        Msg &= "     diretório sem sub-dir:      -dirsemsub ""c:\origemsemsub""" & vbCrLf
                        Msg &= "     diretório réplica:          -repl ""c:\destino""" & vbCrLf
                        Msg &= "     sem gravar no regedit:      -noregedit" & vbCrLf
                        Msg &= "     fechando ao concluir:       -fecharaoconcluir" & vbCrLf
                        Msg &= "     executar imediatamente:     -execimed" & vbCrLf
                        Msg &= "     gravar log em (arquivo):    -log ""c:\sync.log""" & vbCrLf
                        Msg &= "     subject da mensagem:        -subject ""Sync Componentes""" & vbCrLf
                        Msg &= "     from da mensagem:           -from ""'Suporte' [web@icraft.com.br]""" & vbCrLf
                        Msg &= "     sem confirmar:              -semconfirm" & vbCrLf
                        Msg &= "     recurso controlado:         -recurso ""bkp 10.0.0.70""" & vbCrLf
                        Msg &= vbCrLf
                        Msg &= "     caso não passe qualquer config de dir e repl, última configuração" & vbCrLf
                        Msg &= "         gravada no regedit será recuperada." & vbCrLf
                        MsgBox(Msg)
                        z -= 1
                        End
                End Select
            Next

            If Tab = "" Then
                Tab = Trim(NZ(Icraft.IcftBase.RegAplKey("Sync", "Prog"), ""))
            End If

            If Tab <> "" Then
                Dim TabCols() As String = Split(Tab, ";")
                Dim z As Integer = 0
                grdPrinc.Rows.Clear()
                Do While z < TabCols.Length
                    grdPrinc.Rows.Add()
                    grdPrinc.Rows(grdPrinc.Rows.Count - 2).Cells(0).Value = TabCols(z)
                    If TabCols.Length > z + 1 Then
                        grdPrinc.Rows(grdPrinc.Rows.Count - 2).Cells(1).Value = TabCols(z + 1)
                        If TabCols.Length > z + 2 Then
                            grdPrinc.Rows(grdPrinc.Rows.Count - 2).Cells(2).Value = NZ(TabCols(z + 2), False)
                        End If
                    End If
                    grdPrinc.Rows(grdPrinc.Rows.Count - 2).Cells(3).Value = grdPrinc.Rows.Count - 1
                    grdPrinc.Rows(grdPrinc.Rows.Count - 2).Cells(4).Value = Format(Now, "yyyy-MM-dd HH:mm:ss")
                    z += 3
                Loop
            End If

            fldLog.Checked = NZ(Icraft.IcftBase.RegAplKey("Sync", "Log"), True)

            If ExecImed Then
                btnExec_Click(sender, e)
            End If
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Carregando")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            TamanhoDir.Show()
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Pedindo Tamanho")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            TamanhoDirApaga.Show()
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Pedindo Exclusão de Igualdades")
        End Try
    End Sub

    Private Sub DeMinuto_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeMinuto.Tick
        Try
            If chkAtiva.Enabled Then
                If Format(Now, "HH:mm") = txtHoras.Text Then
                    btnExec_Click(sender, e)
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub chkAtiva_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAtiva.CheckedChanged
        Try
            DeMinuto.Enabled = chkAtiva.Checked
            chkAtiva.Enabled = True
            btnExec.Enabled = True
        Catch
        End Try
    End Sub


    Private Sub btnPrepIncluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrepIncluir.Click
        Try
            Ordenar()
            Dim pos = 1
            Do While pos < grdPrinc.Rows.Count
                If grdPrinc.Rows(pos).Cells(0).Value <> "" Then
                    grdPrinc.Rows.Add("", "", True, grdPrinc.Rows(pos).Cells(3).Value - 0.5, Format(Now, "yyyy-MM-dd HH:mm:ss"))
                    pos += 2
                Else
                    pos += 1
                End If
            Loop
            Ordenar(grdPrinc.Columns(3))
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Preparando para Incluir")
        End Try
    End Sub

    Private Sub grdPrinc_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdPrinc.CellEndEdit
        Try
            grdPrinc.Rows(e.RowIndex).Cells(4).Value = Format(Now, "yyyy-MM-dd HH:mm:ss")
            Ordenar()
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Concluindo Edição")
        End Try
    End Sub

    Private Sub grdPrinc_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles grdPrinc.RowsAdded
        Try
            grdPrinc.Rows(e.RowIndex).Cells(4).Value = Format(Now, "yyyy-MM-dd HH:mm:ss")
            Ordenar()
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Adicionando Linha")
        End Try
    End Sub

    Private Sub grdPrinc_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles grdPrinc.RowsRemoved
        Try
            Ordenar()
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Excluindo Linha")
        End Try
    End Sub


    Sub Ordenar(Optional ByVal COL As DataGridViewColumn = Nothing)
        Dim seq As Integer = 1
        Dim Linhas As New List(Of System.Windows.Forms.DataGridViewRow)
        Dim ar As List(Of Windows.Forms.DataGridViewRow) = (From x As System.Windows.Forms.DataGridViewRow In grdPrinc.Rows Where Val(x.Cells(3).Value) <> 0 Order By Val(x.Cells(3).Value), x.Cells(4).Value Descending Select x).ToList
        If ar.Count > 0 Then
            Linhas.AddRange(ar)
        End If
        ar = (From x As System.Windows.Forms.DataGridViewRow In grdPrinc.Rows Where Val(x.Cells(3).Value) = 0 Order By Val(x.Cells(3).Value), x.Cells(4).Value Descending Select x).ToList
        If ar.Count > 0 Then
            Linhas.AddRange(ar)
        End If

        For Each Linha As System.Windows.Forms.DataGridViewRow In Linhas
            Linha.Cells(3).Value = seq
            seq += 1
        Next

        MOSTRAORDEM(COL)
    End Sub

    Sub MOSTRAORDEM(Optional ByVal COL As DataGridViewColumn = Nothing)
        Try
            If Not IsNothing(COL) Then
                grdPrinc.Sort(COL, System.ComponentModel.ListSortDirection.Ascending)
            Else
                If Not IsNothing(grdPrinc.SortedColumn) Then
                    grdPrinc.Sort(grdPrinc.SortedColumn, IIf(grdPrinc.SortOrder = SortOrder.Descending, System.ComponentModel.ListSortDirection.Descending, System.ComponentModel.ListSortDirection.Ascending))
                End If
            End If
        Catch
        End Try

    End Sub

    Private Sub btnExp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExp.Click
        Try
            Mais.txtResult.Text = "Sync.exe"
            If chkAtiva.Checked Then
                Mais.txtResult.Text &= " -auto """ & txtHoras.Text & """"
            End If
            If txtEmailPara.Text <> "" Then
                Mais.txtResult.Text &= " -email """ & txtEmailPara.Text & """"
            End If
            If txtServidor.Text <> "" Then
                Mais.txtResult.Text &= " -smtp """ & txtServidor.Text & """"
            End If
            If chkFechar.Checked Then
                Mais.txtResult.Text &= " -fecharaoconcluir"
            End If
            For Each Linha As DataGridViewRow In grdPrinc.Rows
                If Trim(Linha.Cells(0).Value) <> "" Then
                    Mais.txtResult.Text &= " " & IIf(CType(Linha.Cells(2).Value, Boolean), "-dir ", "-dirsemsub ") & """" & Trim(Linha.Cells(0).Value) & """"
                End If
                If Trim(Linha.Cells(1).Value) <> "" Then
                    Mais.txtResult.Text &= " -repl """ & Trim(Linha.Cells(1).Value) & """"
                End If
            Next
            If chkSemRegedit.Checked Then
                Mais.txtResult.Text &= " -noregedit"
            End If
            If txtArqLog.Text <> "" Then
                Mais.txtResult.Text &= " -arqlog """ & txtArqLog.Text & """"
            End If
            If txtFrom.Text <> "" Then
                Mais.txtResult.Text &= " -from """ & txtFrom.Text & """"
            End If
            If txtSubject.Text <> "" Then
                Mais.txtResult.Text &= " -subject """ & txtSubject.Text & """"
            End If
            If txtRecurso.Text <> "" Then
                Mais.txtResult.Text &= " -recurso """ & txtRecurso.Text & """"
            End If
            Mais.ShowDialog()
            Mais.Focus()
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Exportando Bat")
        End Try
    End Sub



    Private Sub grdPrinc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdPrinc.KeyDown
        Try
            If e.KeyValue = 187 Then
                grdPrinc.Rows.Add()

                For z As Integer = grdPrinc.Rows.Count - 3 To grdPrinc.CurrentCell.RowIndex Step -1
                    For ZZ As Integer = 0 To grdPrinc.Columns.Count - 1
                        grdPrinc.Rows(z + 1).Cells(ZZ).Value = grdPrinc.Rows(z).Cells(ZZ).Value
                    Next
                Next

                For ZZ As Integer = 0 To grdPrinc.Columns.Count - 1
                    If ZZ = 3 Then
                        grdPrinc.Rows(grdPrinc.CurrentCell.RowIndex).Cells(ZZ).Value = grdPrinc.Rows(grdPrinc.CurrentCell.RowIndex - 1).Cells(ZZ).Value + 0.1
                    Else
                        grdPrinc.Rows(grdPrinc.CurrentCell.RowIndex).Cells(ZZ).Value = grdPrinc.Rows(grdPrinc.Rows.Count - 1).Cells(ZZ).Value
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Digitando Tecla")
        End Try
    End Sub


    Private Sub btnAcertaNomes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcertaNomes.Click
        Try
            If MsgBox("Nomes de subdiretórios localizados em diretórios mencionados em ORIGEM serão ajustados para formato 'NomeDoDiretório'. Deseja confirmar?", MsgBoxStyle.YesNo, "Acerta Nomes") = MsgBoxResult.Yes Then
                For Each r As DataGridViewRow In grdPrinc.Rows
                    Dim dd As String = Trim(r.Cells(0).Value)
                    If dd <> "" Then
                        If System.IO.Directory.Exists(dd) Then
                            Dim ARQS() As String = System.IO.Directory.GetDirectories(dd)
                            Dim result As Microsoft.VisualBasic.MsgBoxResult = MsgBoxResult.No
                            For Each ARQ As String In ARQS
                                Dim SONOME As String = System.IO.Path.GetFileName(ARQ)
                                Dim ARQNOVO As String = Replace(Icraft.IcftBase.PrimLetraMaius(Replace(SONOME, "_", " ")), " ", "")
                                If ARQNOVO <> SONOME Then
                                    Try
                                        Rename(ARQ, Icraft.IcftBase.FileExpr(System.IO.Path.GetDirectoryName(ARQ), ARQNOVO))
                                    Catch EX As Exception
                                        If result <> MsgBoxResult.Yes Then
                                            result = MsgBox("Erro:" & EX.Message & ". Continua fazendo ignorando erros futuros?", MsgBoxStyle.Critical + MsgBoxStyle.YesNoCancel, "Acertando Nomes")
                                            If result = MsgBoxResult.Cancel Then
                                                Exit For
                                            End If
                                        End If
                                    End Try
                                End If
                            Next
                        End If
                    End If
                Next
                MsgBox("Termino de rename.")
            End If
        Catch ex As Exception
            MsgBox("Erro:" & ex.Message & ".", MsgBoxStyle.Critical, "Acertando Nomes")
        End Try
    End Sub

End Class

