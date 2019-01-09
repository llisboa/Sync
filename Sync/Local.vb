Public Class Local

    Public Shared Sub CalcDirs1(ByVal Arv As Windows.Forms.TreeView, ByVal VerDir As Boolean, ByVal VerArqs As Boolean, ByVal Apaga As Boolean)
        Dim Origem As String = ""
        Dim Replica As String = ""
        Dim OrigemT As Long = 0
        Dim ReplicaT As Long = 0
        Arv.Nodes.Clear()
        Arv.Visible = True
        Arv.ShowLines = True
        Arv.ShowPlusMinus = True
        For Each Linha As DataGridViewRow In Princ.grdPrinc.Rows
            If Not Linha.IsNewRow Then
                If Linha.Cells(0).Value <> "" Then
                    Origem = Linha.Cells(0).Value
                End If
                If Linha.Cells(1).Value <> "" Then
                    Replica = Linha.Cells(1).Value
                End If

                Origem = Trim(Origem)
                Replica = Trim(Replica)

                Dim Orig As New System.IO.FileInfo(Origem)
                Dim Repl As System.IO.FileInfo
                If Replica <> "" Then
                    Repl = New System.IO.FileInfo(Replica)
                End If

                Dim Node As TreeNode = Arv.Nodes.Add(Origem & "..." & Replica)

                If Orig.Attributes And IO.FileAttributes.Directory Then
                    If VerDir Then

                        Dim DD As New System.IO.DirectoryInfo(Origem)
                        If DD.Exists Then
                            For Each Rela As System.IO.DirectoryInfo In DD.GetDirectories

                                OrigemT = 0
                                ReplicaT = 0
                                If Not Rela.Name Like "__*" Then
                                    Icraft.IcftBase.Espera(0.2)
                                    Try
                                        OrigemT = TamDir(Rela)
                                    Catch
                                    End Try
                                    Dim ReplicaDir As String = ""
                                    Dim DelOK As String = ""
                                    If Replica <> "" Then
                                        ReplicaDir = Icraft.IcftBase.FileExpr(Replica, Rela.Name)
                                        Try
                                            ReplicaT = TamDir(ReplicaDir)
                                        Catch
                                        End Try

                                        If Apaga And ((OrigemT = ReplicaT) OrElse (ReplicaT = 0)) Then
                                            ' elimina diretório do caminho replicado caso seja igual...
                                            Try
                                                System.IO.Directory.Delete(ReplicaDir, True)
                                                DelOK = ""
                                            Catch EX As Exception
                                                DelOK = "Falha:" & EX.Message
                                            End Try
                                        End If
                                    End If

                                    Node.Nodes.Add(Rela.FullName & ": " & Format(OrigemT, "0,000") & IIf(Replica <> "", IIf(OrigemT = ReplicaT, " = ", " <> ") & ReplicaDir & ": " & Format(ReplicaT, "0,000") & IIf(DelOK <> "", "[" & DelOK & "]", ""), ""))

                                    Icraft.IcftBase.Espera(0.2)
                                End If
                            Next
                        End If

                    End If

                    If VerArqs Then

                        Dim DD As New System.IO.DirectoryInfo(Origem)
                        If DD.Exists Then
                            For Each Rela As System.IO.FileInfo In DD.GetFiles()
                                OrigemT = 0
                                ReplicaT = 0
                                If Not Rela.Name Like "__*" Then
                                    Icraft.IcftBase.Espera(0.2)
                                    Try
                                        OrigemT = Rela.Length
                                    Catch
                                    End Try
                                    Dim ReplicaDir As String = Icraft.IcftBase.FileExpr(Replica, Rela.Name)
                                    Try
                                        ReplicaT = New System.IO.FileInfo(ReplicaDir).Length
                                    Catch
                                    End Try

                                    Dim DelOK As String = ""
                                    If Apaga And ((OrigemT = ReplicaT) OrElse (ReplicaT = 0)) Then
                                        ' elimina diretório do caminho replicado caso seja igual...
                                        Try
                                            System.IO.File.Delete(ReplicaDir)
                                            DelOK = ""
                                        Catch EX As Exception
                                            DelOK = "Falha:" & EX.Message
                                        End Try
                                    End If

                                    Node.Nodes.Add(Rela.FullName & ": " & OrigemT & IIf(OrigemT = ReplicaT, " = ", " <> ") & ReplicaDir & ": " & ReplicaT & IIf(DelOK <> "", "[" & DelOK & "]", ""))
                                    Icraft.IcftBase.Espera(0.2)
                                End If
                            Next
                        End If



                    End If

                Else
                    MsgBox("SÓ DIRETÓRIO É PERMITIDO!")
                End If





            End If
        Next

        MsgBox("CONCLUÍDO")
    End Sub

    Sub CalcDirs(ByVal Arv As Windows.Forms.TreeView)
        Dim Origem As String = ""
        Dim Replica As String = ""
        Arv.Nodes.Clear()
        Arv.Visible = True
        Arv.ShowLines = True
        Arv.ShowPlusMinus = True

        For Each Linha As DataGridViewRow In Princ.grdPrinc.Rows
            If Not Linha.IsNewRow Then
                If Linha.Cells(0).Value <> "" Then
                    Origem = Linha.Cells(0).Value
                End If
                If Linha.Cells(1).Value <> "" Then
                    Replica = Linha.Cells(1).Value
                End If

                ArvoreTamDir.Carrega(Arv, Replica)

            End If
        Next
    End Sub

    Enum ExibeTamDirOpcoes
        Auto
        Bytes
        KBytes
        MBytes
        GBytes
        Teras
    End Enum

    Function ExibeTamDir(ByVal Bytes As Object, Optional ByVal Formato As ExibeTamDirOpcoes = ExibeTamDirOpcoes.Auto, Optional ByVal CasasDecimais As Integer = 2) As String
        Dim TotBytes As Long = Bytes
        Dim Tot As Double
        If Formato = ExibeTamDirOpcoes.Bytes Then
            Return TotBytes & " " & Icraft.IcftBase.Pl(TotBytes, "Byte")
        ElseIf Formato = ExibeTamDirOpcoes.KBytes Then
            Tot = Math.Round(TotBytes / 1024, CasasDecimais)
            Return Tot & " " & Icraft.IcftBase.Pl(Tot, "KByte")
        ElseIf Formato = ExibeTamDirOpcoes.MBytes Then
            Tot = Math.Round(TotBytes / 1024 ^ 2)
            Return Tot & " " & Icraft.IcftBase.Pl(Tot, "MByte")
        ElseIf Formato = ExibeTamDirOpcoes.GBytes Then
            Tot = Math.Round(TotBytes / 1024 ^ 3)
            Return Tot & " " & Icraft.IcftBase.Pl(Tot, "GByte")
        ElseIf Formato = ExibeTamDirOpcoes.Teras Then
            Tot = Math.Round(TotBytes / 1024 ^ 4)
            Return Tot & " " & Icraft.IcftBase.Pl(Tot, "TeraByte")
        ElseIf Formato = ExibeTamDirOpcoes.Auto Then
            If TotBytes < 1024 Then
                Return ExibeTamDir(TotBytes, ExibeTamDirOpcoes.Bytes)
            ElseIf TotBytes < 1024 ^ 2 Then
                Return ExibeTamDir(TotBytes, ExibeTamDirOpcoes.KBytes)
            ElseIf TotBytes < 1024 ^ 3 Then
                Return ExibeTamDir(TotBytes, ExibeTamDirOpcoes.MBytes)
            ElseIf TotBytes < 1024 ^ 4 Then
                Return ExibeTamDir(TotBytes, ExibeTamDirOpcoes.GBytes)
            Else
                Return ExibeTamDir(TotBytes, ExibeTamDirOpcoes.Teras)
            End If
        End If
        Return ""
    End Function

    Public Shared Function TamDir(ByVal Diretorio As String) As Long
        Diretorio = Icraft.IcftBase.FileExpr(Diretorio)
        Return TamDir(New System.IO.DirectoryInfo(Diretorio))
    End Function

    Public Shared Function TamDir(ByVal Diretorio As System.IO.DirectoryInfo) As Long
        Dim Ret As Long = 0

        For Each FF As System.IO.FileInfo In Diretorio.GetFiles
            Ret += FF.Length
            Application.DoEvents()
        Next

        For Each DDD As System.IO.DirectoryInfo In Diretorio.GetDirectories
            Ret += TamDir(DDD)
            Application.DoEvents()
        Next
        Return Ret
    End Function


    Class ArvoreTamDir
        Public Arv As TreeView = Nothing
        Public Diretorio As String

        Sub New(ByVal Diretorio As String)
            Arv = New TreeView
            Me.Diretorio = Icraft.IcftBase.FileExpr(Diretorio)
            Inicia()
        End Sub

        Private Sub Inicia()
            Dim No As TreeNode = Arv.Nodes.Add(Diretorio)
            TreeNodesAdd(No.Nodes, Carrega(New System.IO.DirectoryInfo(Diretorio)).Nodes)
            No.Tag = New Props(No)
            ApresentaProps(No)
        End Sub

        Sub ApresentaProps(ByVal No As TreeNode)
            For Each Item As TreeNode In No.Nodes
                ApresentaProps(Item)
            Next
            Dim P As Props = No.Tag
            No.Text = No.Text & " (Tamanho:" & P.Tamanho & ";QtdArqs:" & P.QtdArqs & ")"
        End Sub

        Sub New(ByVal Arvore As TreeView, ByVal Diretorio As String)
            Arv = Arvore
            Me.Diretorio = Icraft.IcftBase.FileExpr(Diretorio)
            Inicia()
        End Sub

        Public Shared Sub Carrega(ByVal Arvore As TreeView, ByVal Diretorio As String)
            Dim Arv As New ArvoreTamDir(Arvore, Diretorio)
        End Sub

        Public Shared Function Carrega(ByVal Diretorio As System.IO.DirectoryInfo) As TreeView
            Dim Arv As New TreeView
            For Each DDD As System.IO.DirectoryInfo In Diretorio.GetDirectories
                Dim No As TreeNode = Arv.Nodes.Add(DDD.Name)
                Dim Ret As TreeNodeCollection = Carrega(DDD).Nodes
                TreeNodesAdd(No.Nodes, Ret)
                Application.DoEvents()
            Next
            Return Arv
        End Function

        Public Shared Sub TreeNodesAdd(ByVal TreeNodes As TreeNodeCollection, ByVal Add As TreeNodeCollection)
            For Each Item As TreeNode In Add
                TreeNodes.Add(Item.Clone)
            Next
        End Sub

        Class Props
            Public Tamanho As Long
            Public QtdArqs As Long
            Sub New(ByVal Node As TreeNode)
                Tamanho = 0
                QtdArqs = 0
                For Each Item As TreeNode In Node.Nodes
                    If IsNothing(Item.Tag) Then
                        Item.Tag = New Props(Item)
                    End If
                    Tamanho += CType(Item.Tag, ArvoreTamDir.Props).Tamanho
                    QtdArqs += CType(Item.Tag, ArvoreTamDir.Props).QtdArqs

                    Application.DoEvents()
                Next
                For Each Item As System.IO.FileInfo In New System.IO.DirectoryInfo(Node.FullPath).GetFiles
                    Tamanho += Item.Length
                    QtdArqs += 1

                    Application.DoEvents()
                Next
            End Sub
        End Class


    End Class


End Class
