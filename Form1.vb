Imports System.IO
Imports System.IO.Compression
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Xml
Imports DevExpress.LookAndFeel
Imports DevExpress.Skins
Imports DevExpress.Utils
Imports DevExpress.Utils.Behaviors
Imports DevExpress.Utils.Colors
Imports DevExpress.Utils.DragDrop
Imports DevExpress.XtraBars
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Localization
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraTab
Imports DevExpress.XtraTreeList.Nodes
Imports NAudio.CoreAudioApi

Public Class Form1
    Private CaminhoXMLInicial As String = IO.Path.Combine(Application.UserAppDataPath, "Configuracoes.xml")

    Private PastaArquivos As String = ""
    Private CaminhoConfigXML As String = ""

    Private PastaHoje As String
    Private PastaDoxologia As String
    Private PastaFundos As String
    Private PastaApelo As String
    Private PastaIntervalo As String
    Private PastaOracao As String
    Private PastaAnimados As String
    Private PastaInfantis As String
    Private PastaCerimoniais As String
    Private PastaEspeciais As String
    Private PastaOutros As String
    Private PastaProvaieVede As String
    Private PastaDbv As String
    Private PastaAvt As String
    Private PastaJovens As String
    Private PastaInformativo As String
    Private PastaColetaneas As String
    Private Sub FrmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CarregarConfiguracoesPastaArquivos()

        If Not String.IsNullOrEmpty(PastaArquivos) Then
            ' --- INICIALIZE TODAS AS VARIÁVEIS DE PASTA AQUI ---
            CaminhoConfigXML = IO.Path.Combine(PastaArquivos, "Configuracoes.xml")

            PastaHoje = PastaArquivos & "\HOJE"
            PastaDoxologia = PastaArquivos & "\VINHETAS"
            PastaFundos = PastaArquivos & "\FUNDOS E TRILHAS"
            PastaApelo = PastaFundos & "\APELO"
            PastaIntervalo = PastaFundos & "\INTERVALO"
            PastaOracao = PastaFundos & "\ORAÇÃO"
            PastaAnimados = PastaFundos & "\ANIMADOS"
            PastaInfantis = PastaFundos & "\INFANTIS"
            PastaCerimoniais = PastaFundos & "\CERIMONIAIS"
            PastaEspeciais = PastaFundos & "\ESPECIAIS"
            PastaOutros = PastaFundos & "\OUTROS"
            PastaProvaieVede = PastaArquivos & "\PROVAI E VEDE"
            PastaDbv = PastaArquivos & "\IDEAIS\Desbravadores"
            PastaAvt = PastaArquivos & "\IDEAIS\Aventureiros"
            PastaJovens = PastaArquivos & "\IDEAIS\Jovens"
            PastaInformativo = PastaArquivos & "\INFORMATIVO DAS MISSÕES"
            PastaColetaneas = PastaArquivos & "\COLETÂNEAS"

            ' --- CHAME OS MÉTODOS QUE DEPENDEM DAS PASTAS AGORA ---
            CarregarTema()
            VolumeDoWindows()
            CarregarGridHoje()
            CarregarGridDoxologia()
            CarregarGridFundos()
            CarregarGridApelo()
            CarregarGridIntervalo()
            CarregarGridOracao()
            CarregarGridAnimados()
            CarregarGridInfantis()
            CarregarGridCerimoniais()
            CarregarGridEspeciais()
            CarregarGridOutros()
            IdentificarMesProvaieVede()
            CarregarGridDbv()
            CarregarGridAvt()
            CarregarGridJovens()
            CarregarGridMusicas()
            CarregarGridColetaneas()

            CarregarGridInformativo()
        End If

    End Sub

#Region "PASTA ARQUIVOS"
    ' Lista global das pastas obrigatórias (coloque no topo da classe)
    Private ReadOnly PastasObrigatorias As String() = {
    "",
    "HOJE",
    "VINHETAS",
    "FUNDOS E TRILHAS",
    "FUNDOS E TRILHAS\APELO",
    "FUNDOS E TRILHAS\INTERVALO",
    "FUNDOS E TRILHAS\ORAÇÃO",
    "FUNDOS E TRILHAS\ANIMADOS",
    "FUNDOS E TRILHAS\INFANTIS",
    "FUNDOS E TRILHAS\CERIMONIAIS",
    "FUNDOS E TRILHAS\ESPECIAIS",
    "FUNDOS E TRILHAS\OUTROS",
    "PROVAI E VEDE",
    "IDEAIS\Desbravadores",
    "IDEAIS\Aventureiros",
    "IDEAIS\Jovens",
    "INFORMATIVO DAS MISSÕES",
    "COLETÂNEAS"
}

    ' --- CARREGAR CONFIGURAÇÕES ---
    Private Sub CarregarConfiguracoesPastaArquivos()
        Try
            Dim xmlDoc As XmlDocument = CarregarXmlInicial()
            Dim pastaNode As XmlNode = xmlDoc.DocumentElement?.SelectSingleNode("PastaArquivos")

            If pastaNode IsNot Nothing AndAlso IO.Directory.Exists(pastaNode.InnerText) Then
                PastaArquivos = pastaNode.InnerText
            Else
                DefinirPastaArquivos()
            End If
        Catch
            DefinirPastaArquivos()
        End Try
    End Sub

    ' --- DEFINIR PASTA PELO USUÁRIO ---
    Private Sub DefinirPastaArquivos()
        If String.IsNullOrWhiteSpace(PastaArquivos) Then
            XtraMessageBox.Show("Nenhuma pasta configurada. Configure primeiro a pasta 'Arquivos CultoIASD'.",
                            "Ops!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        Using folderBrowser As New FolderBrowserDialog With {
        .Description = "Selecione a pasta 'Arquivos CultoIASD'.",
        .ShowNewFolderButton = True,
        .SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    }
            If folderBrowser.ShowDialog() = DialogResult.OK Then
                If VerificarPastasExistem(folderBrowser.SelectedPath) Then
                    PastaArquivos = folderBrowser.SelectedPath
                    SalvarConfiguracoesPastaArquivos()
                Else
                    DefinirPastaArquivos() ' Repete até escolher a correta
                End If
            ElseIf String.IsNullOrWhiteSpace(PastaArquivos) Then
                Me.Close()
            End If
        End Using
    End Sub

    ' --- VERIFICAR TODAS AS PASTAS ---
    Private Function VerificarPastasExistem(caminhoBase As String) As Boolean
        For Each subpasta In PastasObrigatorias
            Dim fullPath = Path.Combine(caminhoBase, subpasta)
            If Not IO.Directory.Exists(fullPath) Then
                XtraMessageBox.Show($"A pasta não existe: {fullPath}",
                                "Pasta Incorreta", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
        Next
        Return True
    End Function

    ' --- SALVAR CONFIG EM XML ---
    Private Sub SalvarConfiguracoesPastaArquivos()
        Try
            Dim xmlDoc As XmlDocument = CarregarXmlInicial()

            If xmlDoc.DocumentElement Is Nothing Then
                xmlDoc.AppendChild(xmlDoc.CreateElement("Configuracoes"))
            End If

            Dim root = xmlDoc.DocumentElement
            Dim pastaNode As XmlNode = root.SelectSingleNode("PastaArquivos")

            If pastaNode Is Nothing Then
                pastaNode = xmlDoc.CreateElement("PastaArquivos")
                root.AppendChild(pastaNode)
            End If

            pastaNode.InnerText = PastaArquivos
            SalvarXmlInicial(xmlDoc)
        Catch ex As Exception
            MessageBox.Show("Erro ao salvar configuração: " & ex.Message)
        End Try
    End Sub

    ' --- XML UTILITÁRIOS ---
    Public Function CarregarXmlInicial() As XmlDocument
        Dim xmlDoc As New XmlDocument()
        Try
            If File.Exists(CaminhoXMLInicial) Then
                xmlDoc.Load(CaminhoXMLInicial)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao carregar XML: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return xmlDoc
    End Function

    Public Sub SalvarXmlInicial(xmlDoc As XmlDocument)
        Try
            xmlDoc.Save(CaminhoXMLInicial)
        Catch ex As Exception
            MessageBox.Show("Erro ao salvar XML: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --- BOTÃO DO MENU ---
    Private Sub BarButtonItemDefinirPastaArquivos_ItemClick(sender As Object, e As ItemClickEventArgs) _
    Handles BarButtonItemDefinirPastaArquivos.ItemClick
        DefinirPastaArquivos()
    End Sub

#End Region

#Region "HOJE"
    Private ArquivosTableHoje As DataTable
    Private dragRowHandle As Integer = GridControl.InvalidRowHandle

    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableHoje()
        ArquivosTableHoje = New DataTable()
        ArquivosTableHoje.Columns.Add("Indice", GetType(Integer))
        ArquivosTableHoje.Columns.Add("Icone", GetType(Image))
        ArquivosTableHoje.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableHoje.Columns.Add("Categoria", GetType(String))
        ArquivosTableHoje.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridHoje()
        CarregarConfiguracoesHoje()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewHoje()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
        TryCast(GridControlHoje.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        ' Drag&Drop
        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewHoje)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorHoje_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorHoje_DragOver

        SplitContainer3.Panel2.AllowDrop = True
        AddHandler SplitContainer3.Panel2.DragEnter, AddressOf PanelHoje_DragEnter
        AddHandler SplitContainer3.Panel2.DragDrop, AddressOf PanelHoje_DragDrop
    End Sub

    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosHoje()
        PrepararArquivosTableHoje()

        Dim arquivos = Directory.GetFiles(PastaHoje)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()

            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableHoje.Rows.Add(i + 1, img, Path.GetFileName(arq), "", arq)
        Next

        GridControlHoje.DataSource = ArquivosTableHoje
        ConfigurarColunaTipoDeArquivoHoje()
        ConfigurarFormatRulesHoje(GridViewHoje)
        SalvarConfiguracoesHoje()
    End Sub

    ' Carregar configurações do GridHoje do XML
    Private Sub CarregarConfiguracoesHoje()
        PrepararArquivosTableHoje()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosHoje()
            Return
        End If

        Try
            Dim gridHojeNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridHoje")
            If gridHojeNode Is Nothing Then
                CarregarArquivosHoje()
                Return
            End If

            Dim arquivos As XmlNodeList = gridHojeNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim categoria As String = arqNode.Attributes("Categoria")?.Value
                Dim filePath As String = Path.Combine(PastaHoje, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableHoje.Rows.Add(indice, img, nomeArquivo, categoria, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes
            For Each r As XmlNode In remover
                gridHojeNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaHoje)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridHojeNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    novoNode.SetAttribute("Categoria", "")
                    gridHojeNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableHoje.Rows.Add(indice, img, nomeArquivo, "", filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridHojeNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridHojeNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridHojeNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableHoje.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlHoje.DataSource = ArquivosTableHoje
            ConfigurarColunaTipoDeArquivoHoje()
            ConfigurarFormatRulesHoje(GridViewHoje)

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Retorna o dicionário de categorias e cores
    Private Function GetCategoriasHoje() As Dictionary(Of String, Color)
        Return New Dictionary(Of String, Color) From {
        {"Congregacional", Color.FromArgb(141, 198, 63)},
        {"Especial", Color.FromArgb(255, 242, 204)},
        {"Aviso", Color.FromArgb(255, 192, 128)},
        {"Adoração Infantil", Color.FromArgb(146, 208, 208)},
        {"Minuto Profético", Color.FromArgb(237, 125, 49)},
        {"Momento de Saúde", Color.FromArgb(255, 192, 128)},
        {"Informativo", Color.FromArgb(91, 155, 213)},
        {"Pregação", Color.FromArgb(0, 112, 192)},
        {"Palestra", Color.FromArgb(0, 176, 240)},
        {"Apresentação", Color.FromArgb(197, 90, 217)},
        {"Investidura", Color.FromArgb(255, 0, 0)},
        {"Batismo", Color.FromArgb(0, 255, 255)},
        {"Abertura", Color.FromArgb(221, 235, 247)},
        {"Escola Sabatina", Color.FromArgb(204, 204, 255)},
        {"Ofertório", Color.FromArgb(146, 208, 80)},
        {"Oração", Color.FromArgb(0, 176, 80)},
        {"Live", Color.FromArgb(255, 0, 255)},
        {"Vinheta", Color.FromArgb(192, 0, 0)},
        {"Imagem/Vídeo Fundo", Color.FromArgb(0, 176, 80)}
    }
    End Function

    ' Configura o RepositoryItemComboBox
    Private Sub ConfigurarColunaTipoDeArquivoHoje()
        Dim categorias = GetCategoriasHoje()

        RepositoryItemComboBoxHoje.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        RepositoryItemComboBoxHoje.Items.Clear()
        RepositoryItemComboBoxHoje.Items.AddRange(categorias.Keys.ToArray())

        ' Remove handlers antigos para evitar duplicação
        RemoveHandler RepositoryItemComboBoxHoje.DrawItem, AddressOf ComboBoxHoje_DrawItem
        AddHandler RepositoryItemComboBoxHoje.DrawItem, AddressOf ComboBoxHoje_DrawItem
    End Sub

    ' Handler compartilhado para desenhar os itens coloridos
    Private Sub ComboBoxHoje_DrawItem(sender As Object, e As DevExpress.XtraEditors.ListBoxDrawItemEventArgs)
        Dim categorias = GetCategoriasHoje()
        If e.Index < 0 Then Return

        Dim text As String = e.Item.ToString()
        Dim back As Color = If(categorias.ContainsKey(text), categorias(text), SystemColors.Window)
        Dim fore As Color = If(back.GetBrightness() < 0.5, Color.White, Color.Black)

        Using br As New SolidBrush(back)
            e.Graphics.FillRectangle(br, e.Bounds)
        End Using

        Dim drawFont As Font = If(e.Appearance.Font IsNot Nothing, e.Appearance.Font, SystemFonts.DefaultFont)

        TextRenderer.DrawText(e.Graphics, text, drawFont, e.Bounds, fore,
                          TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

        e.Handled = True
    End Sub

    ' Configura as FormatRules no GridView
    Private Sub ConfigurarFormatRulesHoje(gridView As DevExpress.XtraGrid.Views.Grid.GridView)
        Dim categorias = GetCategoriasHoje()
        Dim colCategoria = gridView.Columns("Categoria")

        gridView.FormatRules.Clear()

        For Each kvp In categorias
            Dim regra As New DevExpress.XtraGrid.GridFormatRule()
            Dim condicao As New DevExpress.XtraEditors.FormatConditionRuleValue()

            regra.Column = colCategoria
            regra.ApplyToRow = True
            regra.Name = "Regra_" & kvp.Key

            condicao.Condition = DevExpress.XtraEditors.FormatCondition.Equal
            condicao.Value1 = kvp.Key

            condicao.Appearance.BackColor = kvp.Value
            condicao.Appearance.Options.UseBackColor = True
            condicao.Appearance.ForeColor = If(kvp.Value.GetBrightness() < 0.5, Color.White, Color.Black)
            condicao.Appearance.Options.UseForeColor = True

            regra.Rule = condicao
            gridView.FormatRules.Add(regra)
        Next
    End Sub

    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorHoje_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorHoje_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlHoje.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableHoje.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableHoje.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableHoje.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableHoje.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableHoje.Rows.Count Then targetIndex = ArquivosTableHoje.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableHoje.Rows.Count Then
            ArquivosTableHoje.Rows.Add(newRow)
            targetIndex = ArquivosTableHoje.Rows.Count - 1
        Else
            ArquivosTableHoje.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableHoje.Rows.Count - 1
            ArquivosTableHoje.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlHoje.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesHoje()

        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel e do grid musica
    Private Sub PanelHoje_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub PanelHoje_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableHoje.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaHoje, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim img As Image = Nothing
                Try
                    Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                    img = icone.ToBitmap()
                Catch
                    img = SystemIcons.WinLogo.ToBitmap()
                End Try

                ' Adiciona no final do DataTable
                ArquivosTableHoje.Rows.Add(nextIndex, img, Path.GetFileName(destPath), "", destPath)
                nextIndex += 1
            Next

            GridControlHoje.RefreshDataSource()
            SalvarConfiguracoesHoje()
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewHoje_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewHoje.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableHoje.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableHoje.Rows.Count - 1
                ArquivosTableHoje.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlHoje.RefreshDataSource()
            SalvarConfiguracoesHoje()
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewHoje_DoubleClick(sender As Object, e As EventArgs) Handles GridViewHoje.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewHoje_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewHoje.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub

    ' Menu de contexto com clique direito
    Private Sub GridViewHoje_MouseUp(sender As Object, e As MouseEventArgs) Handles GridViewHoje.MouseUp
        If e.Button = MouseButtons.Right Then
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView)
            Dim hi As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(e.Location)

            If hi.InRowCell Then
                ' Clicou em uma célula válida -> move o foco e seleciona a linha
                view.FocusedRowHandle = hi.RowHandle
                view.ClearSelection()
                view.SelectRow(hi.RowHandle)

                ' Habilita os botões do PopupMenu
                BarButtonItemExcluirHoje.Enabled = True
                BarButtonItemRenomearHoje.Enabled = True

                ' Exibe o menu
                PopupMenuHoje.ShowPopup(Control.MousePosition)
            Else
                ' Clicou fora de uma célula -> limpa seleção
                view.ClearSelection()
                view.FocusedRowHandle = DevExpress.XtraGrid.GridControl.InvalidRowHandle

                ' Desabilita os botões do PopupMenu
                BarButtonItemExcluirHoje.Enabled = False
                BarButtonItemRenomearHoje.Enabled = False
                PopupMenuHoje.ShowPopup(Control.MousePosition)

            End If
        End If
    End Sub
    Private Sub BarButtonItemRenomearHoje_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemRenomearHoje.ItemClick
        Dim rowHandle = GridViewHoje.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridViewHoje.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim oldName = dataRow("NomeArquivo").ToString
        Dim oldPath = dataRow("Local").ToString

        ' Separa nome e extensão
        Dim baseName = Path.GetFileNameWithoutExtension(oldName)
        Dim extension = Path.GetExtension(oldName)

        ' Só permite editar o nome, não a extensão
        Dim newBaseName = InputBox("Digite o novo nome do arquivo:", "Renomear", baseName).Trim

        If String.IsNullOrWhiteSpace(newBaseName) OrElse newBaseName = baseName Then Return

        Dim newName = newBaseName & extension
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        Try
            If File.Exists(oldPath) Then
                File.Move(oldPath, newPath)

                ' Atualiza DataRow
                dataRow("NomeArquivo") = newName
                dataRow("Local") = newPath

                GridControlHoje.RefreshDataSource
                SalvarConfiguracoesHoje
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao renomear o arquivo: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Excluir arquivos selecionados
    Private Sub BarButtonItemExcluirHoje_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemExcluirHoje.ItemClick
        Dim selectedRows = GridViewHoje.GetSelectedRows
        If selectedRows.Length = 0 Then Return

        Dim msg = If(selectedRows.Length = 1,
                       "Deseja realmente excluir o arquivo selecionado?",
                       $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
        If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim rowsToRemove As New List(Of DataRow)

        For Each rowHandle In selectedRows
            Dim dataRow = GridViewHoje.GetDataRow(rowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath = dataRow("Local").ToString
                Try
                    If File.Exists(filePath) Then File.Delete(filePath)
                Catch ex As Exception
                    MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                rowsToRemove.Add(dataRow)
            End If
        Next

        For Each dr In rowsToRemove
            ArquivosTableHoje.Rows.Remove(dr)
        Next

        ' Recalcula índices
        For i = 0 To ArquivosTableHoje.Rows.Count - 1
            ArquivosTableHoje.Rows(i)("Indice") = i + 1
        Next

        GridControlHoje.RefreshDataSource
        SalvarConfiguracoesHoje
    End Sub

    ' Atualizar, Limpar e Abrir Pasta
    Private Sub BarButtonItemAtualizarHoje_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAtualizarHoje.ItemClick
        CarregarConfiguracoesHoje()
        SalvarConfiguracoesHoje()
    End Sub
    ' Limpar todos os arquivos da pasta HOJE
    Private Sub BarButtonItemLimparHoje_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemLimparHoje.ItemClick
        If MessageBox.Show("Tem certeza que deseja excluir todos os arquivos da pasta '" & PastaHoje & "'?",
                       "Confirmar exclusão",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Warning) = DialogResult.Yes Then
            Try
                ' Itera pelas linhas da tabela e exclui os arquivos pelo caminho completo salvo em "Local"
                For Each row As DataRow In ArquivosTableHoje.Rows
                    Dim filePath = row("Local").ToString
                    If File.Exists(filePath) Then
                        Try
                            File.Delete(filePath)
                        Catch ex As Exception
                            MessageBox.Show("Erro ao excluir: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    End If
                Next

                ' Limpa a tabela
                ArquivosTableHoje.Clear

                ' Atualiza o grid
                GridControlHoje.RefreshDataSource
                SalvarConfiguracoesHoje

            Catch ex As Exception
                MessageBox.Show("Erro ao excluir arquivos: " & ex.Message,
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    ' Abre a pasta HOJE no Explorer
    Private Sub BarButtonItemAbrirPastaHoje_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAbrirPastaHoje.ItemClick
        Try
            If Directory.Exists(PastaHoje) Then
                Process.Start("explorer.exe", PastaHoje)
            Else
                MessageBox.Show("A pasta não existe: " & PastaHoje,
                            "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir a pasta: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Salvar Configurações ao mudar a categoria
    Private Sub RepositoryItemComboBoxHoje_EditValueChanged(sender As Object, e As EventArgs) Handles RepositoryItemComboBoxHoje.EditValueChanged
        Dim editor As ComboBoxEdit = TryCast(sender, ComboBoxEdit)
        If editor IsNot Nothing Then
            Dim view As GridView = TryCast(GridControlHoje.FocusedView, GridView)
            If view IsNot Nothing Then
                Dim rowHandle As Integer = view.FocusedRowHandle
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        ' Se EditValue for Nothing, salva string vazia
                        dataRow("Categoria") = If(editor.EditValue IsNot Nothing, editor.EditValue.ToString(), String.Empty)
                        GridControlHoje.RefreshDataSource()
                        SalvarConfiguracoesHoje()
                    End If
                End If
            End If
        End If
    End Sub

    ' Salvar configurações do GridHoje em XML
    Private Sub SalvarConfiguracoesHoje()
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Se não tiver raiz, cria
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Remove GridHoje antigo (se existir) para recriar
            Dim gridHojeNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridHoje")
            If gridHojeNode IsNot Nothing Then
                xmlDoc.DocumentElement.RemoveChild(gridHojeNode)
            End If

            ' Cria nó GridHoje
            gridHojeNode = xmlDoc.CreateElement("GridHoje")
            Dim attrUltimaAlteracao As XmlAttribute = xmlDoc.CreateAttribute("UltimaAlteracao")
            attrUltimaAlteracao.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            gridHojeNode.Attributes.Append(attrUltimaAlteracao)

            ' Adiciona arquivos do DataTable
            For Each row As DataRow In ArquivosTableHoje.Rows
                Dim arqNode As XmlElement = xmlDoc.CreateElement("Arquivo")

                Dim attrIndice As XmlAttribute = xmlDoc.CreateAttribute("Indice")
                attrIndice.Value = row("Indice").ToString()
                arqNode.Attributes.Append(attrIndice)

                Dim attrNome As XmlAttribute = xmlDoc.CreateAttribute("NomeArquivo")
                attrNome.Value = row("NomeArquivo").ToString()
                arqNode.Attributes.Append(attrNome)

                Dim attrCategoria As XmlAttribute = xmlDoc.CreateAttribute("Categoria")
                attrCategoria.Value = row("Categoria").ToString()
                arqNode.Attributes.Append(attrCategoria)

                ' >>> Novo atributo Local <<<
                Dim attrLocal As XmlAttribute = xmlDoc.CreateAttribute("Local")
                attrLocal.Value = row("Local").ToString()
                arqNode.Attributes.Append(attrLocal)

                gridHojeNode.AppendChild(arqNode)
            Next

            ' Adiciona o nó na raiz
            xmlDoc.DocumentElement.AppendChild(gridHojeNode)

            ' Salva no arquivo
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar configurações: " & ex.Message,
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Tamanho da Fonte do GridViewProcessos
    Private Sub GridControlHoje_MouseWheel(sender As Object, e As MouseEventArgs) Handles GridControlHoje.MouseWheel
        ' Verifica se a tecla Ctrl está pressionada
        If My.Computer.Keyboard.CtrlKeyDown Then
            ' Obtém o GridView principal do GridControl
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControlHoje.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
            If view IsNot Nothing Then
                ' Obtém o tamanho atual da fonte
                Dim currentFontSize As Single = view.Appearance.Row.Font.Size

                ' Define um limite mínimo e máximo
                Dim minFontSize As Single = 5
                Dim maxFontSize As Single = 20

                ' Aumenta ou diminui o tamanho da fonte com base na direção do scroll
                If e.Delta > 0 AndAlso currentFontSize < maxFontSize Then
                    currentFontSize += 1
                ElseIf e.Delta < 0 AndAlso currentFontSize > minFontSize Then
                    currentFontSize -= 1
                End If

                ' Aplica o novo tamanho de fonte nas aparências relevantes
                Dim novaFonte As New Font(view.Appearance.Row.Font.FontFamily, currentFontSize)

                view.Appearance.Row.Font = novaFonte
                view.Appearance.HeaderPanel.Font = novaFonte
                view.Appearance.FocusedRow.Font = novaFonte
                view.Appearance.GroupRow.Font = novaFonte
                view.Appearance.FooterPanel.Font = novaFonte

                SalvarTamanhoFonteGridViewHoje(currentFontSize)

            End If
        End If
    End Sub

    ' === Preferências do GridHoje ===
    Public Function ObterTamanhoFonteGridViewHoje() As Single
        Dim xmlDoc As XmlDocument = CarregarXml()
        Dim tamanhoFonte As Single = 10 ' valor padrão

        Try
            Dim node As XmlNode = xmlDoc.SelectSingleNode("/Configuracoes/Preferencias/FonteGridViewHoje")
            If node IsNot Nothing Then
                Single.TryParse(node.InnerText, tamanhoFonte)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao obter tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return tamanhoFonte
    End Function
    Public Sub SalvarTamanhoFonteGridViewHoje(ByVal tamanhoFonte As Single)
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Cria raiz se não existir
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Cria nó <Preferencias> se não existir
            Dim preferenciasNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("Preferencias")
            If preferenciasNode Is Nothing Then
                preferenciasNode = xmlDoc.CreateElement("Preferencias")
                xmlDoc.DocumentElement.AppendChild(preferenciasNode)
            End If

            ' Cria ou atualiza nó <FonteGridViewHoje>
            Dim fonteNode As XmlNode = preferenciasNode.SelectSingleNode("FonteGridViewHoje")
            If fonteNode Is Nothing Then
                fonteNode = xmlDoc.CreateElement("FonteGridViewHoje")
                preferenciasNode.AppendChild(fonteNode)
            End If

            fonteNode.InnerText = tamanhoFonte.ToString(System.Globalization.CultureInfo.InvariantCulture)

            ' Salva com XmlManager
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "DOXOLOGIA"
    Private ArquivosTableDoxologia As DataTable
    Private dragRowHandleDoxologia As Integer = GridControl.InvalidRowHandle
    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableDoxologia()
        ArquivosTableDoxologia = New DataTable()
        ArquivosTableDoxologia.Columns.Add("Indice", GetType(Integer))
        ArquivosTableDoxologia.Columns.Add("Icone", GetType(Image))
        ArquivosTableDoxologia.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableDoxologia.Columns.Add("Categoria", GetType(String))
        ArquivosTableDoxologia.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridDoxologia()

        CarregarConfiguracoesDoxologia()
        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewDoxologia()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControlDoxologia.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewDoxologia)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorDoxologia_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorDoxologia_DragOver

        SplitContainer4.Panel1.AllowDrop = True
        AddHandler SplitContainer4.Panel1.DragEnter, AddressOf PanelDoxologia_DragEnter
        AddHandler SplitContainer4.Panel1.DragDrop, AddressOf PanelDoxologia_DragDrop
    End Sub
    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosDoxologia()
        PrepararArquivosTableDoxologia()

        Dim arquivos = Directory.GetFiles(PastaDoxologia)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableDoxologia.Rows.Add(i + 1, img, Path.GetFileName(arq), "", arq)
        Next

        GridControlDoxologia.DataSource = ArquivosTableDoxologia
        ConfigurarColunaTipoDeArquivoDoxologia()
        ConfigurarFormatRulesDoxologia(GridViewDoxologia)

        SalvarConfiguracoesDoxologia()
    End Sub

    ' Carregar configurações do GridDoxologia do XML
    Private Sub CarregarConfiguracoesDoxologia()
        PrepararArquivosTableDoxologia()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosDoxologia()
            Return
        End If

        Try
            Dim gridDoxologiaNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridDoxologia")
            If gridDoxologiaNode Is Nothing Then
                CarregarArquivosDoxologia()
                Return
            End If

            Dim arquivos As XmlNodeList = gridDoxologiaNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim categoria As String = arqNode.Attributes("Categoria")?.Value
                Dim filePath As String = Path.Combine(PastaDoxologia, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableDoxologia.Rows.Add(indice, img, nomeArquivo, categoria, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes
            For Each r As XmlNode In remover
                gridDoxologiaNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaDoxologia)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridDoxologiaNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    novoNode.SetAttribute("Categoria", "")
                    gridDoxologiaNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableDoxologia.Rows.Add(indice, img, nomeArquivo, "", filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridDoxologiaNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridDoxologiaNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridDoxologiaNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableDoxologia.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlDoxologia.DataSource = ArquivosTableDoxologia
            ConfigurarColunaTipoDeArquivoDoxologia()
            ConfigurarFormatRulesDoxologia(GridViewDoxologia)

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' Retorna o dicionário de categorias e cores
    Private Function GetCategoriasDoxologia() As Dictionary(Of String, Color)
        Return New Dictionary(Of String, Color) From {
        {"Congregacional", Color.FromArgb(141, 198, 63)},
        {"Especial", Color.FromArgb(255, 242, 204)},
        {"Aviso", Color.FromArgb(255, 192, 128)},
        {"Adoração Infantil", Color.FromArgb(146, 208, 208)},
        {"Minuto Profético", Color.FromArgb(237, 125, 49)},
        {"Momento de Saúde", Color.FromArgb(255, 192, 128)},
        {"Informativo", Color.FromArgb(91, 155, 213)},
        {"Pregação", Color.FromArgb(0, 112, 192)},
        {"Palestra", Color.FromArgb(0, 176, 240)},
        {"Apresentação", Color.FromArgb(197, 90, 217)},
        {"Investidura", Color.FromArgb(255, 0, 0)},
        {"Batismo", Color.FromArgb(0, 255, 255)},
        {"Abertura", Color.FromArgb(221, 235, 247)},
        {"Escola Sabatina", Color.FromArgb(204, 204, 255)},
        {"Ofertório", Color.FromArgb(146, 208, 80)},
        {"Oração", Color.FromArgb(0, 176, 80)},
        {"Live", Color.FromArgb(255, 0, 255)},
        {"Vinheta", Color.FromArgb(192, 0, 0)},
        {"Imagem/Vídeo Fundo", Color.FromArgb(0, 176, 80)}
    }
    End Function

    ' Configura o RepositoryItemComboBox
    Private Sub ConfigurarColunaTipoDeArquivoDoxologia()
        Dim categorias = GetCategoriasDoxologia()

        RepositoryItemComboBoxDoxologia.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        RepositoryItemComboBoxDoxologia.Items.Clear()
        RepositoryItemComboBoxDoxologia.Items.AddRange(categorias.Keys.ToArray())

        ' Remove handlers antigos para evitar duplicação
        RemoveHandler RepositoryItemComboBoxDoxologia.DrawItem, AddressOf ComboBoxDoxologia_DrawItem
        AddHandler RepositoryItemComboBoxDoxologia.DrawItem, AddressOf ComboBoxDoxologia_DrawItem
    End Sub

    ' Handler compartilhado para desenhar os itens coloridos
    Private Sub ComboBoxDoxologia_DrawItem(sender As Object, e As DevExpress.XtraEditors.ListBoxDrawItemEventArgs)
        Dim categorias = GetCategoriasDoxologia()
        If e.Index < 0 Then Return

        Dim text As String = e.Item.ToString()
        Dim back As Color = If(categorias.ContainsKey(text), categorias(text), SystemColors.Window)
        Dim fore As Color = If(back.GetBrightness() < 0.5, Color.White, Color.Black)

        Using br As New SolidBrush(back)
            e.Graphics.FillRectangle(br, e.Bounds)
        End Using

        Dim drawFont As Font = If(e.Appearance.Font IsNot Nothing, e.Appearance.Font, SystemFonts.DefaultFont)

        TextRenderer.DrawText(e.Graphics, text, drawFont, e.Bounds, fore,
                          TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

        e.Handled = True
    End Sub

    ' Configura as FormatRules no GridView
    Private Sub ConfigurarFormatRulesDoxologia(gridView As DevExpress.XtraGrid.Views.Grid.GridView)
        Dim categorias = GetCategoriasDoxologia()
        Dim colCategoria = gridView.Columns("Categoria")

        gridView.FormatRules.Clear()

        For Each kvp In categorias
            Dim regra As New DevExpress.XtraGrid.GridFormatRule()
            Dim condicao As New DevExpress.XtraEditors.FormatConditionRuleValue()

            regra.Column = colCategoria
            regra.ApplyToRow = True
            regra.Name = "Regra_" & kvp.Key

            condicao.Condition = DevExpress.XtraEditors.FormatCondition.Equal
            condicao.Value1 = kvp.Key

            condicao.Appearance.BackColor = kvp.Value
            condicao.Appearance.Options.UseBackColor = True
            condicao.Appearance.ForeColor = If(kvp.Value.GetBrightness() < 0.5, Color.White, Color.Black)
            condicao.Appearance.Options.UseForeColor = True

            regra.Rule = condicao
            gridView.FormatRules.Add(regra)
        Next
    End Sub

    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorDoxologia_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorDoxologia_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlDoxologia.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableDoxologia.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableDoxologia.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableDoxologia.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableDoxologia.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableDoxologia.Rows.Count Then targetIndex = ArquivosTableDoxologia.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableDoxologia.Rows.Count Then
            ArquivosTableDoxologia.Rows.Add(newRow)
            targetIndex = ArquivosTableDoxologia.Rows.Count - 1
        Else
            ArquivosTableDoxologia.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableDoxologia.Rows.Count - 1
            ArquivosTableDoxologia.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlDoxologia.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesDoxologia()

        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelDoxologia_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelDoxologia_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableDoxologia.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaDoxologia, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableDoxologia.Rows.Add(nextIndex, img, Path.GetFileName(destPath), "", destPath)
                nextIndex += 1
            Next

            GridControlDoxologia.RefreshDataSource()
            SalvarConfiguracoesDoxologia()
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewDoxologia_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewDoxologia.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view = CType(sender, GridView)
            Dim selectedRows = view.GetSelectedRows

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath = dataRow("Local").ToString
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableDoxologia.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i = 0 To ArquivosTableDoxologia.Rows.Count - 1
                ArquivosTableDoxologia.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlDoxologia.RefreshDataSource()
            SalvarConfiguracoesDoxologia()
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewDoxologia_DoubleClick(sender As Object, e As EventArgs) Handles GridViewDoxologia.DoubleClick
        Dim view = CType(sender, GridView)
        Dim hitInfo = view.CalcHitInfo(view.GridControl.PointToClient(MousePosition))
        If hitInfo.InRow Then
            Dim dataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath = dataRow("Local").ToString
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewDoxologia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewDoxologia.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view = CType(sender, GridView)
            Dim rowHandle = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath = dataRow("Local").ToString
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub

    ' Menu de contexto com clique direito
    Private Sub GridViewDoxologia_MouseUp(sender As Object, e As MouseEventArgs) Handles GridViewDoxologia.MouseUp
        If e.Button = MouseButtons.Right Then
            Dim view = TryCast(sender, GridView)
            Dim hi = view.CalcHitInfo(e.Location)

            If hi.InRowCell Then
                ' Clicou em uma célula válida -> move o foco e seleciona a linha
                view.FocusedRowHandle = hi.RowHandle
                view.ClearSelection()
                view.SelectRow(hi.RowHandle)

                ' Habilita os botões do PopupMenu
                BarButtonItemExcluirDoxologia.Enabled = True
                BarButtonItemRenomearDoxologia.Enabled = True

                ' Exibe o menu
                PopupMenuDoxologia.ShowPopup(MousePosition)
            Else
                ' Clicou fora de uma célula -> limpa seleção
                view.ClearSelection()
                view.FocusedRowHandle = GridControl.InvalidRowHandle

                ' Desabilita os botões do PopupMenu
                BarButtonItemExcluirDoxologia.Enabled = False
                BarButtonItemRenomearDoxologia.Enabled = False
                PopupMenuDoxologia.ShowPopup(MousePosition)

            End If
        End If
    End Sub
    Private Sub BarButtonItemRenomearDoxologia_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemRenomearDoxologia.ItemClick
        Dim rowHandle = GridViewDoxologia.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridViewDoxologia.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim oldName = dataRow("NomeArquivo").ToString
        Dim oldPath = dataRow("Local").ToString

        ' Separa nome e extensão
        Dim baseName = Path.GetFileNameWithoutExtension(oldName)
        Dim extension = Path.GetExtension(oldName)

        ' Só permite editar o nome, não a extensão
        Dim newBaseName = InputBox("Digite o novo nome do arquivo:", "Renomear", baseName).Trim

        If String.IsNullOrWhiteSpace(newBaseName) OrElse newBaseName = baseName Then Return

        Dim newName = newBaseName & extension
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        Try
            If File.Exists(oldPath) Then
                File.Move(oldPath, newPath)

                ' Atualiza DataRow
                dataRow("NomeArquivo") = newName
                dataRow("Local") = newPath

                GridControlDoxologia.RefreshDataSource()
                SalvarConfiguracoesDoxologia()
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao renomear o arquivo: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Excluir arquivos selecionados
    Private Sub BarButtonItemExcluirDoxologia_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemExcluirDoxologia.ItemClick
        Dim selectedRows = GridViewDoxologia.GetSelectedRows
        If selectedRows.Length = 0 Then Return

        Dim msg = If(selectedRows.Length = 1,
                       "Deseja realmente excluir o arquivo selecionado?",
                       $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
        If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim rowsToRemove As New List(Of DataRow)

        For Each rowHandle In selectedRows
            Dim dataRow = GridViewDoxologia.GetDataRow(rowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath = dataRow("Local").ToString
                Try
                    If File.Exists(filePath) Then File.Delete(filePath)
                Catch ex As Exception
                    MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                rowsToRemove.Add(dataRow)
            End If
        Next

        For Each dr In rowsToRemove
            ArquivosTableDoxologia.Rows.Remove(dr)
        Next

        ' Recalcula índices
        For i = 0 To ArquivosTableDoxologia.Rows.Count - 1
            ArquivosTableDoxologia.Rows(i)("Indice") = i + 1
        Next

        GridControlDoxologia.RefreshDataSource()
        SalvarConfiguracoesDoxologia()
    End Sub

    ' Atualizar, Limpar e Abrir Pasta
    Private Sub BarButtonItemAtualizarDoxologia_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAtualizarDoxologia.ItemClick
        CarregarConfiguracoesDoxologia()
        SalvarConfiguracoesDoxologia()
    End Sub

    ' Abre a pasta Doxologia no Explorer
    Private Sub BarButtonItemAbrirPastaDoxologia_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAbrirPastaDoxologia.ItemClick
        Try
            If Directory.Exists(PastaDoxologia) Then
                Process.Start("explorer.exe", PastaDoxologia)
            Else
                MessageBox.Show("A pasta não existe: " & PastaDoxologia,
                            "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir a pasta: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Salvar Configurações ao mudar a categoria
    Private Sub RepositoryItemComboBoxDoxologia_EditValueChanged(sender As Object, e As EventArgs) Handles RepositoryItemComboBoxDoxologia.EditValueChanged
        Dim editor = TryCast(sender, ComboBoxEdit)
        If editor IsNot Nothing Then
            Dim view = TryCast(GridControlDoxologia.FocusedView, GridView)
            If view IsNot Nothing Then
                Dim rowHandle = view.FocusedRowHandle
                If rowHandle >= 0 Then
                    Dim dataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        ' Se EditValue for Nothing, salva string vazia
                        dataRow("Categoria") = If(editor.EditValue IsNot Nothing, editor.EditValue.ToString, String.Empty)
                        GridControlDoxologia.RefreshDataSource()
                        SalvarConfiguracoesDoxologia()
                    End If
                End If
            End If
        End If
    End Sub

    ' Salvar configurações do GridDoxologia em XML
    Private Sub SalvarConfiguracoesDoxologia()
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Se não tiver raiz, cria
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Remove GridDoxologia antigo (se existir) para recriar
            Dim gridDoxologiaNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridDoxologia")
            If gridDoxologiaNode IsNot Nothing Then
                xmlDoc.DocumentElement.RemoveChild(gridDoxologiaNode)
            End If

            ' Cria nó GridDoxologia
            gridDoxologiaNode = xmlDoc.CreateElement("GridDoxologia")
            Dim attrUltimaAlteracao As XmlAttribute = xmlDoc.CreateAttribute("UltimaAlteracao")
            attrUltimaAlteracao.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            gridDoxologiaNode.Attributes.Append(attrUltimaAlteracao)

            ' Adiciona arquivos do DataTable
            For Each row As DataRow In ArquivosTableDoxologia.Rows
                Dim arqNode As XmlElement = xmlDoc.CreateElement("Arquivo")

                Dim attrIndice As XmlAttribute = xmlDoc.CreateAttribute("Indice")
                attrIndice.Value = row("Indice").ToString()
                arqNode.Attributes.Append(attrIndice)

                Dim attrNome As XmlAttribute = xmlDoc.CreateAttribute("NomeArquivo")
                attrNome.Value = row("NomeArquivo").ToString()
                arqNode.Attributes.Append(attrNome)

                Dim attrCategoria As XmlAttribute = xmlDoc.CreateAttribute("Categoria")
                attrCategoria.Value = row("Categoria").ToString()
                arqNode.Attributes.Append(attrCategoria)

                ' >>> Novo atributo Local <<<
                Dim attrLocal As XmlAttribute = xmlDoc.CreateAttribute("Local")
                attrLocal.Value = row("Local").ToString()
                arqNode.Attributes.Append(attrLocal)

                gridDoxologiaNode.AppendChild(arqNode)
            Next

            ' Adiciona o nó na raiz
            xmlDoc.DocumentElement.AppendChild(gridDoxologiaNode)

            ' Salva no arquivo
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar configurações: " & ex.Message,
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'Tamanho da Fonte do GridViewProcessos
    Private Sub GridControlDoxologia_MouseWheel(sender As Object, e As MouseEventArgs) Handles GridControlDoxologia.MouseWheel
        ' Verifica se a tecla Ctrl está pressionada
        If My.Computer.Keyboard.CtrlKeyDown Then
            ' Obtém o GridView principal do GridControl
            Dim view = TryCast(GridControlDoxologia.MainView, GridView)
            If view IsNot Nothing Then
                ' Obtém o tamanho atual da fonte
                Dim currentFontSize = view.Appearance.Row.Font.Size

                ' Define um limite mínimo e máximo
                Dim minFontSize As Single = 5
                Dim maxFontSize As Single = 20

                ' Aumenta ou diminui o tamanho da fonte com base na direção do scroll
                If e.Delta > 0 AndAlso currentFontSize < maxFontSize Then
                    currentFontSize += 1
                ElseIf e.Delta < 0 AndAlso currentFontSize > minFontSize Then
                    currentFontSize -= 1
                End If

                ' Aplica o novo tamanho de fonte nas aparências relevantes
                Dim novaFonte As New Font(view.Appearance.Row.Font.FontFamily, currentFontSize)

                view.Appearance.Row.Font = novaFonte
                view.Appearance.HeaderPanel.Font = novaFonte
                view.Appearance.FocusedRow.Font = novaFonte
                view.Appearance.GroupRow.Font = novaFonte
                view.Appearance.FooterPanel.Font = novaFonte

                SalvarTamanhoFonteGridViewDoxologia(currentFontSize)

            End If
        End If
    End Sub

    ' === Preferências do GridDoxologia ===
    Public Function ObterTamanhoFonteGridViewDoxologia() As Single
        Dim xmlDoc As XmlDocument = CarregarXml()
        Dim tamanhoFonte As Single = 10 ' valor padrão

        Try
            Dim node As XmlNode = xmlDoc.SelectSingleNode("/Configuracoes/Preferencias/FonteGridViewDoxologia")
            If node IsNot Nothing Then
                Single.TryParse(node.InnerText, tamanhoFonte)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao obter tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return tamanhoFonte
    End Function
    Public Sub SalvarTamanhoFonteGridViewDoxologia(ByVal tamanhoFonte As Single)
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Cria raiz se não existir
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Cria nó <Preferencias> se não existir
            Dim preferenciasNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("Preferencias")
            If preferenciasNode Is Nothing Then
                preferenciasNode = xmlDoc.CreateElement("Preferencias")
                xmlDoc.DocumentElement.AppendChild(preferenciasNode)
            End If

            ' Cria ou atualiza nó <FonteGridViewDoxologia>
            Dim fonteNode As XmlNode = preferenciasNode.SelectSingleNode("FonteGridViewDoxologia")
            If fonteNode Is Nothing Then
                fonteNode = xmlDoc.CreateElement("FonteGridViewDoxologia")
                preferenciasNode.AppendChild(fonteNode)
            End If

            fonteNode.InnerText = tamanhoFonte.ToString(System.Globalization.CultureInfo.InvariantCulture)

            ' Salva com XmlManager
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



#End Region

#Region "GRID FUNDOS"

#Region "Apelo"
    Private ArquivosTableApelo As DataTable
    Private dragRowHandleApelo As Integer = GridControl.InvalidRowHandle

    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableApelo()
        ArquivosTableApelo = New DataTable()
        ArquivosTableApelo.Columns.Add("Indice", GetType(Integer))
        ArquivosTableApelo.Columns.Add("Icone", GetType(Image))
        ArquivosTableApelo.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableApelo.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridApelo()
        CarregarConfiguracoesApelo()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewFundo()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
        TryCast(GridControlApelo.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        ' Cor de fundo das linhas pares e ímpares
        GridViewApelo.OptionsView.EnableAppearanceEvenRow = True
        GridViewApelo.OptionsView.EnableAppearanceOddRow = True
        GridViewApelo.Appearance.EvenRow.BackColor = Color.FromArgb(215, 235, 248)
        GridViewApelo.Appearance.OddRow.BackColor = Color.FromArgb(195, 224, 244)
        GridViewApelo.Appearance.EvenRow.ForeColor = Color.Black
        GridViewApelo.Appearance.OddRow.ForeColor = Color.Black


        ' Drag&Drop
        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewApelo)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorApelo_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorApelo_DragOver

        XtraTabControlFundos.AllowDrop = True
        AddHandler XtraTabControlFundos.DragEnter, AddressOf PanelApelo_DragEnter
        AddHandler XtraTabControlFundos.DragDrop, AddressOf PanelApelo_DragDrop

        AddHandler GridViewApelo.MouseUp, AddressOf GridViewFundos_MouseUp
        AddHandler GridControlApelo.MouseWheel, AddressOf GridControlFundo_MouseWheel
    End Sub

    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosApelo()
        PrepararArquivosTableApelo()

        Dim arquivos = Directory.GetFiles(PastaApelo)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableApelo.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlApelo.DataSource = ArquivosTableApelo

        SalvarConfiguracoesFundo(GridViewApelo, ArquivosTableApelo)

    End Sub


    ' Carregar configurações do GridApelo do XML
    Private Sub CarregarConfiguracoesApelo()
        PrepararArquivosTableApelo()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosApelo()
            Return
        End If

        Try
            Dim gridApeloNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridViewApelo")
            If gridApeloNode Is Nothing Then
                CarregarArquivosApelo()
                Return
            End If

            Dim arquivos As XmlNodeList = gridApeloNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaApelo, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableApelo.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes do XML
            For Each r As XmlNode In remover
                gridApeloNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaApelo)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridApeloNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridApeloNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableApelo.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridApeloNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridApeloNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridApeloNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableApelo.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlApelo.DataSource = ArquivosTableApelo

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorApelo_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorApelo_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlApelo.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableApelo.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableApelo.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableApelo.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableApelo.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableApelo.Rows.Count Then targetIndex = ArquivosTableApelo.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableApelo.Rows.Count Then
            ArquivosTableApelo.Rows.Add(newRow)
            targetIndex = ArquivosTableApelo.Rows.Count - 1
        Else
            ArquivosTableApelo.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableApelo.Rows.Count - 1
            ArquivosTableApelo.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlApelo.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesFundo(GridViewApelo, ArquivosTableApelo)

        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelApelo_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelApelo_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableApelo.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaApelo, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableApelo.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlApelo.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewApelo, ArquivosTableApelo)
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewApelo_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewApelo.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableApelo.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableApelo.Rows.Count - 1
                ArquivosTableApelo.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlApelo.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewApelo, ArquivosTableApelo)
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewApelo_DoubleClick(sender As Object, e As EventArgs) Handles GridViewApelo.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewApelo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewApelo.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub


#End Region

#Region "Intervalo"
    Private ArquivosTableIntervalo As DataTable
    Private dragRowHandleIntervalo As Integer = GridControl.InvalidRowHandle

    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableIntervalo()
        ArquivosTableIntervalo = New DataTable()
        ArquivosTableIntervalo.Columns.Add("Indice", GetType(Integer))
        ArquivosTableIntervalo.Columns.Add("Icone", GetType(Image))
        ArquivosTableIntervalo.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableIntervalo.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridIntervalo()
        CarregarConfiguracoesIntervalo()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewFundo()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
        TryCast(GridControlIntervalo.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        ' Cor de fundo das linhas pares e ímpares
        GridViewIntervalo.OptionsView.EnableAppearanceEvenRow = True
        GridViewIntervalo.OptionsView.EnableAppearanceOddRow = True
        GridViewIntervalo.Appearance.EvenRow.BackColor = Color.FromArgb(249, 219, 215)
        GridViewIntervalo.Appearance.OddRow.BackColor = Color.FromArgb(248, 202, 197)
        GridViewIntervalo.Appearance.EvenRow.ForeColor = Color.Black
        GridViewIntervalo.Appearance.OddRow.ForeColor = Color.Black

        ' Drag&Drop
        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewIntervalo)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorIntervalo_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorIntervalo_DragOver

        XtraTabControlFundos.AllowDrop = True
        AddHandler XtraTabControlFundos.DragEnter, AddressOf PanelIntervalo_DragEnter
        AddHandler XtraTabControlFundos.DragDrop, AddressOf PanelIntervalo_DragDrop

        AddHandler GridViewIntervalo.MouseUp, AddressOf GridViewFundos_MouseUp
        AddHandler GridControlIntervalo.MouseWheel, AddressOf GridControlFundo_MouseWheel
    End Sub

    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosIntervalo()
        PrepararArquivosTableIntervalo()

        Dim arquivos = Directory.GetFiles(PastaIntervalo)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableIntervalo.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlIntervalo.DataSource = ArquivosTableIntervalo

        SalvarConfiguracoesFundo(GridViewIntervalo, ArquivosTableIntervalo)

    End Sub

    ' Carregar configurações do GridIntervalo do XML
    Private Sub CarregarConfiguracoesIntervalo()
        PrepararArquivosTableIntervalo()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosIntervalo()
            Return
        End If

        Try
            Dim gridIntervaloNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridViewIntervalo")
            If gridIntervaloNode Is Nothing Then
                CarregarArquivosIntervalo()
                Return
            End If

            Dim arquivos As XmlNodeList = gridIntervaloNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaIntervalo, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableIntervalo.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes do XML
            For Each r As XmlNode In remover
                gridIntervaloNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaIntervalo)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridIntervaloNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridIntervaloNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableIntervalo.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridIntervaloNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridIntervaloNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridIntervaloNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableIntervalo.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlIntervalo.DataSource = ArquivosTableIntervalo

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorIntervalo_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorIntervalo_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlIntervalo.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableIntervalo.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableIntervalo.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableIntervalo.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableIntervalo.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableIntervalo.Rows.Count Then targetIndex = ArquivosTableIntervalo.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableIntervalo.Rows.Count Then
            ArquivosTableIntervalo.Rows.Add(newRow)
            targetIndex = ArquivosTableIntervalo.Rows.Count - 1
        Else
            ArquivosTableIntervalo.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableIntervalo.Rows.Count - 1
            ArquivosTableIntervalo.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlIntervalo.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesFundo(GridViewIntervalo, ArquivosTableIntervalo)

        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelIntervalo_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelIntervalo_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableIntervalo.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaIntervalo, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableIntervalo.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlIntervalo.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewIntervalo, ArquivosTableIntervalo)
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewIntervalo_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewIntervalo.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableIntervalo.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableIntervalo.Rows.Count - 1
                ArquivosTableIntervalo.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlIntervalo.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewIntervalo, ArquivosTableIntervalo)
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewIntervalo_DoubleClick(sender As Object, e As EventArgs) Handles GridViewIntervalo.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewIntervalo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewIntervalo.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub


#End Region

#Region "Oracao"
    Private ArquivosTableOracao As DataTable
    Private dragRowHandleOracao As Integer = GridControl.InvalidRowHandle

    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableOracao()
        ArquivosTableOracao = New DataTable()
        ArquivosTableOracao.Columns.Add("Indice", GetType(Integer))
        ArquivosTableOracao.Columns.Add("Icone", GetType(Image))
        ArquivosTableOracao.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableOracao.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridOracao()
        CarregarConfiguracoesOracao()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewFundo()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
        TryCast(GridControlOracao.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        ' Cor de fundo das linhas pares e ímpares
        GridViewOracao.OptionsView.EnableAppearanceEvenRow = True
        GridViewOracao.OptionsView.EnableAppearanceOddRow = True
        GridViewOracao.Appearance.EvenRow.BackColor = Color.FromArgb(231, 241, 219)
        GridViewOracao.Appearance.OddRow.BackColor = Color.FromArgb(220, 234, 202)
        GridViewOracao.Appearance.EvenRow.ForeColor = Color.Black
        GridViewOracao.Appearance.OddRow.ForeColor = Color.Black

        ' Drag&Drop
        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewOracao)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorOracao_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorOracao_DragOver

        XtraTabControlFundos.AllowDrop = True
        AddHandler XtraTabControlFundos.DragEnter, AddressOf PanelOracao_DragEnter
        AddHandler XtraTabControlFundos.DragDrop, AddressOf PanelOracao_DragDrop

        AddHandler GridViewOracao.MouseUp, AddressOf GridViewFundos_MouseUp
        AddHandler GridControlOracao.MouseWheel, AddressOf GridControlFundo_MouseWheel
    End Sub

    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosOracao()
        PrepararArquivosTableOracao()

        Dim arquivos = Directory.GetFiles(PastaOracao)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableOracao.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlOracao.DataSource = ArquivosTableOracao

        SalvarConfiguracoesFundo(GridViewOracao, ArquivosTableOracao)

    End Sub

    ' Carregar configurações do GridOracao do XML
    Private Sub CarregarConfiguracoesOracao()
        PrepararArquivosTableOracao()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosOracao()
            Return
        End If

        Try
            Dim gridOracaoNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridViewOracao")
            If gridOracaoNode Is Nothing Then
                CarregarArquivosOracao()
                Return
            End If

            Dim arquivos As XmlNodeList = gridOracaoNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaOracao, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableOracao.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes do XML
            For Each r As XmlNode In remover
                gridOracaoNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaOracao)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridOracaoNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridOracaoNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableOracao.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridOracaoNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridOracaoNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridOracaoNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableOracao.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlOracao.DataSource = ArquivosTableOracao

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorOracao_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorOracao_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlOracao.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableOracao.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableOracao.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableOracao.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableOracao.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableOracao.Rows.Count Then targetIndex = ArquivosTableOracao.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableOracao.Rows.Count Then
            ArquivosTableOracao.Rows.Add(newRow)
            targetIndex = ArquivosTableOracao.Rows.Count - 1
        Else
            ArquivosTableOracao.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableOracao.Rows.Count - 1
            ArquivosTableOracao.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlOracao.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesFundo(GridViewOracao, ArquivosTableOracao)
        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelOracao_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelOracao_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableOracao.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaOracao, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableOracao.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlOracao.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewOracao, ArquivosTableOracao)
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewOracao_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewOracao.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableOracao.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableOracao.Rows.Count - 1
                ArquivosTableOracao.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlOracao.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewOracao, ArquivosTableOracao)
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewOracao_DoubleClick(sender As Object, e As EventArgs) Handles GridViewOracao.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewOracao_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewOracao.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub

#End Region

#Region "Animados"
    Private ArquivosTableAnimados As DataTable
    Private dragRowHandleAnimados As Integer = GridControl.InvalidRowHandle

    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableAnimados()
        ArquivosTableAnimados = New DataTable()
        ArquivosTableAnimados.Columns.Add("Indice", GetType(Integer))
        ArquivosTableAnimados.Columns.Add("Icone", GetType(Image))
        ArquivosTableAnimados.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableAnimados.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridAnimados()
        CarregarConfiguracoesAnimados()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewFundo()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
        TryCast(GridControlAnimados.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        ' Cor de fundo das linhas pares e ímpares
        GridViewAnimados.OptionsView.EnableAppearanceEvenRow = True
        GridViewAnimados.OptionsView.EnableAppearanceOddRow = True
        GridViewAnimados.Appearance.EvenRow.BackColor = Color.FromArgb(233, 211, 232)
        GridViewAnimados.Appearance.OddRow.BackColor = Color.FromArgb(222, 190, 220)
        GridViewAnimados.Appearance.EvenRow.ForeColor = Color.Black
        GridViewAnimados.Appearance.OddRow.ForeColor = Color.Black

        ' Drag&Drop
        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewAnimados)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorAnimados_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorAnimados_DragOver

        XtraTabControlFundos.AllowDrop = True
        AddHandler XtraTabControlFundos.DragEnter, AddressOf PanelAnimados_DragEnter
        AddHandler XtraTabControlFundos.DragDrop, AddressOf PanelAnimados_DragDrop

        AddHandler GridViewAnimados.MouseUp, AddressOf GridViewFundos_MouseUp
        AddHandler GridControlAnimados.MouseWheel, AddressOf GridControlFundo_MouseWheel
    End Sub

    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosAnimados()
        PrepararArquivosTableAnimados()

        Dim arquivos = Directory.GetFiles(PastaAnimados)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableAnimados.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlAnimados.DataSource = ArquivosTableAnimados

        SalvarConfiguracoesFundo(GridViewAnimados, ArquivosTableAnimados)
    End Sub

    ' Carregar configurações do GridAnimados do XML
    Private Sub CarregarConfiguracoesAnimados()
        PrepararArquivosTableAnimados()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosAnimados()
            Return
        End If

        Try
            Dim gridAnimadosNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridViewAnimados")
            If gridAnimadosNode Is Nothing Then
                CarregarArquivosAnimados()
                Return
            End If

            Dim arquivos As XmlNodeList = gridAnimadosNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaAnimados, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableAnimados.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes do XML
            For Each r As XmlNode In remover
                gridAnimadosNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaAnimados)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridAnimadosNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridAnimadosNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableAnimados.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridAnimadosNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridAnimadosNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridAnimadosNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableAnimados.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlAnimados.DataSource = ArquivosTableAnimados

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorAnimados_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorAnimados_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlAnimados.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableAnimados.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableAnimados.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableAnimados.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableAnimados.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableAnimados.Rows.Count Then targetIndex = ArquivosTableAnimados.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableAnimados.Rows.Count Then
            ArquivosTableAnimados.Rows.Add(newRow)
            targetIndex = ArquivosTableAnimados.Rows.Count - 1
        Else
            ArquivosTableAnimados.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableAnimados.Rows.Count - 1
            ArquivosTableAnimados.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlAnimados.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesFundo(GridViewAnimados, ArquivosTableAnimados)

        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelAnimados_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelAnimados_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableAnimados.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaAnimados, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableAnimados.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlAnimados.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewAnimados, ArquivosTableAnimados)
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewAnimados_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewAnimados.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableAnimados.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableAnimados.Rows.Count - 1
                ArquivosTableAnimados.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlAnimados.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewAnimados, ArquivosTableAnimados)
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewAnimados_DoubleClick(sender As Object, e As EventArgs) Handles GridViewAnimados.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewAnimados_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewAnimados.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub
#End Region

#Region "Infantis"
    Private ArquivosTableInfantis As DataTable
    Private dragRowHandleInfantis As Integer = GridControl.InvalidRowHandle

    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableInfantis()
        ArquivosTableInfantis = New DataTable()
        ArquivosTableInfantis.Columns.Add("Indice", GetType(Integer))
        ArquivosTableInfantis.Columns.Add("Icone", GetType(Image))
        ArquivosTableInfantis.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableInfantis.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridInfantis()
        CarregarConfiguracoesInfantis()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewFundo()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
        TryCast(GridControlInfantis.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        ' Cor de fundo das linhas pares e ímpares
        GridViewInfantis.OptionsView.EnableAppearanceEvenRow = True
        GridViewInfantis.OptionsView.EnableAppearanceOddRow = True
        GridViewInfantis.Appearance.EvenRow.BackColor = Color.FromArgb(206, 233, 240)
        GridViewInfantis.Appearance.OddRow.BackColor = Color.FromArgb(182, 223, 233)
        GridViewInfantis.Appearance.EvenRow.ForeColor = Color.Black
        GridViewInfantis.Appearance.OddRow.ForeColor = Color.Black


        ' Drag&Drop
        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewInfantis)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorInfantis_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorInfantis_DragOver

        XtraTabControlFundos.AllowDrop = True
        AddHandler XtraTabControlFundos.DragEnter, AddressOf PanelInfantis_DragEnter
        AddHandler XtraTabControlFundos.DragDrop, AddressOf PanelInfantis_DragDrop

        AddHandler GridViewInfantis.MouseUp, AddressOf GridViewFundos_MouseUp
        AddHandler GridControlInfantis.MouseWheel, AddressOf GridControlFundo_MouseWheel
    End Sub

    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosInfantis()
        PrepararArquivosTableInfantis()

        Dim arquivos = Directory.GetFiles(PastaInfantis)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableInfantis.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlInfantis.DataSource = ArquivosTableInfantis
        SalvarConfiguracoesFundo(GridViewInfantis, ArquivosTableInfantis)

    End Sub

    ' Carregar configurações do GridInfantis do XML
    Private Sub CarregarConfiguracoesInfantis()
        PrepararArquivosTableInfantis()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosInfantis()
            Return
        End If

        Try
            Dim gridInfantisNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridViewInfantis")
            If gridInfantisNode Is Nothing Then
                CarregarArquivosInfantis()
                Return
            End If

            Dim arquivos As XmlNodeList = gridInfantisNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaInfantis, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableInfantis.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes do XML
            For Each r As XmlNode In remover
                gridInfantisNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaInfantis)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridInfantisNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridInfantisNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableInfantis.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridInfantisNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridInfantisNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridInfantisNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableInfantis.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlInfantis.DataSource = ArquivosTableInfantis

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorInfantis_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorInfantis_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlInfantis.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableInfantis.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableInfantis.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableInfantis.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableInfantis.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableInfantis.Rows.Count Then targetIndex = ArquivosTableInfantis.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableInfantis.Rows.Count Then
            ArquivosTableInfantis.Rows.Add(newRow)
            targetIndex = ArquivosTableInfantis.Rows.Count - 1
        Else
            ArquivosTableInfantis.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableInfantis.Rows.Count - 1
            ArquivosTableInfantis.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlInfantis.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesFundo(GridViewInfantis, ArquivosTableInfantis)

        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelInfantis_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelInfantis_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableInfantis.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaInfantis, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableInfantis.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlInfantis.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewInfantis, ArquivosTableInfantis)
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewInfantis_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewInfantis.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableInfantis.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableInfantis.Rows.Count - 1
                ArquivosTableInfantis.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlInfantis.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewInfantis, ArquivosTableInfantis)
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewInfantis_DoubleClick(sender As Object, e As EventArgs) Handles GridViewInfantis.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewInfantis_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewInfantis.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub
#End Region

#Region "Cerimoniais"
    Private ArquivosTableCerimoniais As DataTable
    Private dragRowHandleCerimoniais As Integer = GridControl.InvalidRowHandle

    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableCerimoniais()
        ArquivosTableCerimoniais = New DataTable()
        ArquivosTableCerimoniais.Columns.Add("Indice", GetType(Integer))
        ArquivosTableCerimoniais.Columns.Add("Icone", GetType(Image))
        ArquivosTableCerimoniais.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableCerimoniais.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridCerimoniais()
        CarregarConfiguracoesCerimoniais()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewFundo()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
        TryCast(GridControlCerimoniais.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        ' Cor de fundo das linhas pares e ímpares
        GridViewCerimoniais.OptionsView.EnableAppearanceEvenRow = True
        GridViewCerimoniais.OptionsView.EnableAppearanceOddRow = True
        GridViewCerimoniais.Appearance.EvenRow.BackColor = Color.FromArgb(251, 243, 207)
        GridViewCerimoniais.Appearance.OddRow.BackColor = Color.FromArgb(250, 237, 183)
        GridViewCerimoniais.Appearance.EvenRow.ForeColor = Color.Black
        GridViewCerimoniais.Appearance.OddRow.ForeColor = Color.Black

        ' Drag&Drop
        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewCerimoniais)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorCerimoniais_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorCerimoniais_DragOver

        XtraTabControlFundos.AllowDrop = True
        AddHandler XtraTabControlFundos.DragEnter, AddressOf PanelCerimoniais_DragEnter
        AddHandler XtraTabControlFundos.DragDrop, AddressOf PanelCerimoniais_DragDrop

        AddHandler GridViewCerimoniais.MouseUp, AddressOf GridViewFundos_MouseUp
        AddHandler GridControlCerimoniais.MouseWheel, AddressOf GridControlFundo_MouseWheel
    End Sub

    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosCerimoniais()
        PrepararArquivosTableCerimoniais()

        Dim arquivos = Directory.GetFiles(PastaCerimoniais)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableCerimoniais.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlCerimoniais.DataSource = ArquivosTableCerimoniais
        SalvarConfiguracoesFundo(GridViewCerimoniais, ArquivosTableCerimoniais)
    End Sub

    ' Carregar configurações do GridCerimoniais do XML
    Private Sub CarregarConfiguracoesCerimoniais()
        PrepararArquivosTableCerimoniais()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosCerimoniais()
            Return
        End If

        Try
            Dim gridCerimoniaisNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridViewCerimoniais")
            If gridCerimoniaisNode Is Nothing Then
                CarregarArquivosCerimoniais()
                Return
            End If

            Dim arquivos As XmlNodeList = gridCerimoniaisNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaCerimoniais, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableCerimoniais.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes do XML
            For Each r As XmlNode In remover
                gridCerimoniaisNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaCerimoniais)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridCerimoniaisNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridCerimoniaisNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableCerimoniais.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridCerimoniaisNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridCerimoniaisNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridCerimoniaisNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableCerimoniais.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlCerimoniais.DataSource = ArquivosTableCerimoniais

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorCerimoniais_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorCerimoniais_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlCerimoniais.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableCerimoniais.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableCerimoniais.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableCerimoniais.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableCerimoniais.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableCerimoniais.Rows.Count Then targetIndex = ArquivosTableCerimoniais.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableCerimoniais.Rows.Count Then
            ArquivosTableCerimoniais.Rows.Add(newRow)
            targetIndex = ArquivosTableCerimoniais.Rows.Count - 1
        Else
            ArquivosTableCerimoniais.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableCerimoniais.Rows.Count - 1
            ArquivosTableCerimoniais.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlCerimoniais.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesFundo(GridViewCerimoniais, ArquivosTableCerimoniais)

        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelCerimoniais_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelCerimoniais_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableCerimoniais.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaCerimoniais, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableCerimoniais.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlCerimoniais.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewCerimoniais, ArquivosTableCerimoniais)

        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewCerimoniais_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewCerimoniais.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableCerimoniais.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableCerimoniais.Rows.Count - 1
                ArquivosTableCerimoniais.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlCerimoniais.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewCerimoniais, ArquivosTableCerimoniais)
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewCerimoniais_DoubleClick(sender As Object, e As EventArgs) Handles GridViewCerimoniais.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewCerimoniais_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewCerimoniais.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub
#End Region

#Region "Especiais"
    Private ArquivosTableEspeciais As DataTable
    Private dragRowHandleEspeciais As Integer = GridControl.InvalidRowHandle

    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableEspeciais()
        ArquivosTableEspeciais = New DataTable()
        ArquivosTableEspeciais.Columns.Add("Indice", GetType(Integer))
        ArquivosTableEspeciais.Columns.Add("Icone", GetType(Image))
        ArquivosTableEspeciais.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableEspeciais.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridEspeciais()
        CarregarConfiguracoesEspeciais()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewFundo()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
        TryCast(GridControlEspeciais.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        ' Cor de fundo das linhas pares e ímpares
        GridViewEspeciais.OptionsView.EnableAppearanceEvenRow = True
        GridViewEspeciais.OptionsView.EnableAppearanceOddRow = True
        GridViewEspeciais.Appearance.EvenRow.BackColor = Color.FromArgb(249, 211, 223)
        GridViewEspeciais.Appearance.OddRow.BackColor = Color.FromArgb(246, 188, 208)
        GridViewEspeciais.Appearance.EvenRow.ForeColor = Color.Black
        GridViewEspeciais.Appearance.OddRow.ForeColor = Color.Black

        ' Drag&Drop
        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewEspeciais)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorEspeciais_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorEspeciais_DragOver

        XtraTabControlFundos.AllowDrop = True
        AddHandler XtraTabControlFundos.DragEnter, AddressOf PanelEspeciais_DragEnter
        AddHandler XtraTabControlFundos.DragDrop, AddressOf PanelEspeciais_DragDrop

        AddHandler GridViewEspeciais.MouseUp, AddressOf GridViewFundos_MouseUp
        AddHandler GridControlEspeciais.MouseWheel, AddressOf GridControlFundo_MouseWheel
    End Sub

    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosEspeciais()
        PrepararArquivosTableEspeciais()

        Dim arquivos = Directory.GetFiles(PastaEspeciais)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableEspeciais.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlEspeciais.DataSource = ArquivosTableEspeciais
        SalvarConfiguracoesFundo(GridViewEspeciais, ArquivosTableEspeciais)
    End Sub

    ' Carregar configurações do GridEspeciais do XML
    Private Sub CarregarConfiguracoesEspeciais()
        PrepararArquivosTableEspeciais()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosEspeciais()
            Return
        End If

        Try
            Dim gridEspeciaisNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridViewEspeciais")
            If gridEspeciaisNode Is Nothing Then
                CarregarArquivosEspeciais()
                Return
            End If

            Dim arquivos As XmlNodeList = gridEspeciaisNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaEspeciais, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableEspeciais.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes do XML
            For Each r As XmlNode In remover
                gridEspeciaisNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaEspeciais)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridEspeciaisNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridEspeciaisNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableEspeciais.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridEspeciaisNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridEspeciaisNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridEspeciaisNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableEspeciais.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlEspeciais.DataSource = ArquivosTableEspeciais

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorEspeciais_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorEspeciais_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlEspeciais.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableEspeciais.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableEspeciais.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableEspeciais.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableEspeciais.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableEspeciais.Rows.Count Then targetIndex = ArquivosTableEspeciais.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableEspeciais.Rows.Count Then
            ArquivosTableEspeciais.Rows.Add(newRow)
            targetIndex = ArquivosTableEspeciais.Rows.Count - 1
        Else
            ArquivosTableEspeciais.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableEspeciais.Rows.Count - 1
            ArquivosTableEspeciais.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlEspeciais.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesFundo(GridViewEspeciais, ArquivosTableEspeciais)
        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelEspeciais_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelEspeciais_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableEspeciais.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaEspeciais, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableEspeciais.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlEspeciais.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewEspeciais, ArquivosTableEspeciais)
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewEspeciais_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewEspeciais.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableEspeciais.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableEspeciais.Rows.Count - 1
                ArquivosTableEspeciais.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlEspeciais.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewEspeciais, ArquivosTableEspeciais)
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewEspeciais_DoubleClick(sender As Object, e As EventArgs) Handles GridViewEspeciais.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewEspeciais_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewEspeciais.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub
#End Region

#Region "Outros"
    Private ArquivosTableOutros As DataTable
    Private dragRowHandleOutros As Integer = GridControl.InvalidRowHandle

    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableOutros()
        ArquivosTableOutros = New DataTable()
        ArquivosTableOutros.Columns.Add("Indice", GetType(Integer))
        ArquivosTableOutros.Columns.Add("Icone", GetType(Image))
        ArquivosTableOutros.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableOutros.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridOutros()
        CarregarConfiguracoesOutros()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewFundo()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
        TryCast(GridControlOutros.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        ' Cor de fundo das linhas pares e ímpares
        GridViewOutros.OptionsView.EnableAppearanceEvenRow = True
        GridViewOutros.OptionsView.EnableAppearanceOddRow = True
        GridViewOutros.Appearance.EvenRow.BackColor = Color.FromArgb(248, 249, 250)
        GridViewOutros.Appearance.OddRow.BackColor = Color.FromArgb(255, 255, 255)
        GridViewOutros.Appearance.EvenRow.ForeColor = Color.Black
        GridViewOutros.Appearance.OddRow.ForeColor = Color.Black

        ' Drag&Drop
        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewOutros)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorOutros_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorOutros_DragOver

        XtraTabControlFundos.AllowDrop = True
        AddHandler XtraTabControlFundos.DragEnter, AddressOf PanelOutros_DragEnter
        AddHandler XtraTabControlFundos.DragDrop, AddressOf PanelOutros_DragDrop

        AddHandler GridViewOutros.MouseUp, AddressOf GridViewFundos_MouseUp
        AddHandler GridControlOutros.MouseWheel, AddressOf GridControlFundo_MouseWheel
    End Sub

    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosOutros()
        PrepararArquivosTableOutros()

        Dim arquivos = Directory.GetFiles(PastaOutros)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableOutros.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlOutros.DataSource = ArquivosTableOutros
        SalvarConfiguracoesFundo(GridViewOutros, ArquivosTableOutros)
    End Sub

    ' Carregar configurações do GridOutros do XML
    Private Sub CarregarConfiguracoesOutros()
        PrepararArquivosTableOutros()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosOutros()
            Return
        End If

        Try
            Dim gridOutrosNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridViewOutros")
            If gridOutrosNode Is Nothing Then
                CarregarArquivosOutros()
                Return
            End If

            Dim arquivos As XmlNodeList = gridOutrosNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaOutros, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableOutros.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes do XML
            For Each r As XmlNode In remover
                gridOutrosNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaOutros)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridOutrosNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridOutrosNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableOutros.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridOutrosNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridOutrosNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridOutrosNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableOutros.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlOutros.DataSource = ArquivosTableOutros

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorOutros_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorOutros_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlOutros.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableOutros.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableOutros.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableOutros.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableOutros.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableOutros.Rows.Count Then targetIndex = ArquivosTableOutros.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableOutros.Rows.Count Then
            ArquivosTableOutros.Rows.Add(newRow)
            targetIndex = ArquivosTableOutros.Rows.Count - 1
        Else
            ArquivosTableOutros.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableOutros.Rows.Count - 1
            ArquivosTableOutros.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlOutros.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesFundo(GridViewOutros, ArquivosTableOutros)

        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelOutros_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelOutros_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableOutros.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaOutros, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableOutros.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlOutros.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewOutros, ArquivosTableOutros)
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewOutros_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewOutros.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableOutros.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableOutros.Rows.Count - 1
                ArquivosTableOutros.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlOutros.RefreshDataSource()
            SalvarConfiguracoesFundo(GridViewOutros, ArquivosTableOutros)
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewOutros_DoubleClick(sender As Object, e As EventArgs) Handles GridViewOutros.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewOutros_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewOutros.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub
#End Region

#Region "Comum a todos os fundos"
    Private GridviewFundoSelecionado As DevExpress.XtraGrid.Views.Grid.GridView
    Private GridControlFundoSelecionado As DevExpress.XtraGrid.GridControl
    Private ArquivosTableFundoSelecionados As DataTable
    Sub CarregarGridFundos()
        XtraTabControlFundos.ShowTabHeader = DefaultBoolean.False
        CheckButtonIntervalo.Checked = True
    End Sub
    Private Sub CheckButtonApelo_CheckedChanged(sender As Object, e As EventArgs) Handles CheckButtonApelo.CheckedChanged
        If CheckButtonApelo.Checked Then
            XtraTabControlFundos.SelectedTabPage = XtraTabPageApelo
            GridviewFundoSelecionado = GridViewApelo
            GridControlFundoSelecionado = GridControlApelo
            ArquivosTableFundoSelecionados = ArquivosTableApelo
        End If
    End Sub
    Private Sub CheckButtonIntervalo_CheckedChanged(sender As Object, e As EventArgs) Handles CheckButtonIntervalo.CheckedChanged
        If CheckButtonIntervalo.Checked Then
            XtraTabControlFundos.SelectedTabPage = XtraTabPageIntervalo
            GridviewFundoSelecionado = GridViewIntervalo
            GridControlFundoSelecionado = GridControlIntervalo
            ArquivosTableFundoSelecionados = ArquivosTableIntervalo
        End If
    End Sub
    Private Sub CheckButtonOracao_CheckedChanged(sender As Object, e As EventArgs) Handles CheckButtonOracao.CheckedChanged
        If CheckButtonOracao.Checked Then
            XtraTabControlFundos.SelectedTabPage = XtraTabPageOracao
            GridviewFundoSelecionado = GridViewOracao
            GridControlFundoSelecionado = GridControlOracao
            ArquivosTableFundoSelecionados = ArquivosTableOracao
        End If
    End Sub
    Private Sub CheckButtonAnimados_CheckedChanged(sender As Object, e As EventArgs) Handles CheckButtonAnimados.CheckedChanged
        If CheckButtonAnimados.Checked Then
            XtraTabControlFundos.SelectedTabPage = XtraTabPageAnimados
            GridviewFundoSelecionado = GridViewAnimados
            GridControlFundoSelecionado = GridControlAnimados
            ArquivosTableFundoSelecionados = ArquivosTableAnimados
        End If
    End Sub
    Private Sub CheckButtonInfantis_CheckedChanged(sender As Object, e As EventArgs) Handles CheckButtonInfantis.CheckedChanged
        If CheckButtonInfantis.Checked Then
            XtraTabControlFundos.SelectedTabPage = XtraTabPageInfantis
            GridviewFundoSelecionado = GridViewInfantis
            GridControlFundoSelecionado = GridControlInfantis
            ArquivosTableFundoSelecionados = ArquivosTableInfantis
        End If
    End Sub
    Private Sub CheckButtonCerimoniais_CheckedChanged(sender As Object, e As EventArgs) Handles CheckButtonCerimoniais.CheckedChanged
        If CheckButtonCerimoniais.Checked Then
            XtraTabControlFundos.SelectedTabPage = XtraTabPageCerimoniais
            GridviewFundoSelecionado = GridViewCerimoniais
            GridControlFundoSelecionado = GridControlCerimoniais
            ArquivosTableFundoSelecionados = ArquivosTableCerimoniais
        End If
    End Sub
    Private Sub CheckButtonEspeciais_CheckedChanged(sender As Object, e As EventArgs) Handles CheckButtonEspeciais.CheckedChanged
        If CheckButtonEspeciais.Checked Then
            XtraTabControlFundos.SelectedTabPage = XtraTabPageEspeciais
            GridviewFundoSelecionado = GridViewEspeciais
            GridControlFundoSelecionado = GridControlEspeciais
            ArquivosTableFundoSelecionados = ArquivosTableEspeciais
        End If
    End Sub
    Private Sub CheckButtonOutros_CheckedChanged(sender As Object, e As EventArgs) Handles CheckButtonOutros.CheckedChanged
        If CheckButtonOutros.Checked Then
            XtraTabControlFundos.SelectedTabPage = XtraTabPageOutros
            GridviewFundoSelecionado = GridViewOutros
            GridControlFundoSelecionado = GridControlOutros
            ArquivosTableFundoSelecionados = ArquivosTableOutros
        End If
    End Sub
    ' Menu de contexto com clique direito
    Private Sub GridViewFundos_MouseUp(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView)
            Dim grid As DevExpress.XtraGrid.GridControl = view.GridControl

            Dim hi As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(e.Location)

            If hi.InRowCell Then
                ' Clicou em uma célula válida -> move o foco e seleciona a linha
                view.FocusedRowHandle = hi.RowHandle
                view.ClearSelection()
                view.SelectRow(hi.RowHandle)

                ' Habilita os botões do PopupMenu
                BarButtonItemExcluirFundo.Enabled = True
                BarButtonItemRenomearFundo.Enabled = True

                ' Exibe o menu
                PopupMenuFundos.ShowPopup(Control.MousePosition)
            Else
                ' Clicou fora de uma célula -> limpa seleção
                view.ClearSelection()
                view.FocusedRowHandle = DevExpress.XtraGrid.GridControl.InvalidRowHandle

                ' Desabilita os botões do PopupMenu
                BarButtonItemExcluirFundo.Enabled = False
                BarButtonItemRenomearFundo.Enabled = False
                PopupMenuFundos.ShowPopup(Control.MousePosition)

            End If
        End If
    End Sub
    Private Sub BarButtonItemRenomearFundo_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemRenomearFundo.ItemClick
        Dim rowHandle = GridviewFundoSelecionado.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridviewFundoSelecionado.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim oldName = dataRow("NomeArquivo").ToString
        Dim oldPath = dataRow("Local").ToString

        ' Separa nome e extensão
        Dim baseName = Path.GetFileNameWithoutExtension(oldName)
        Dim extension = Path.GetExtension(oldName)

        ' Só permite editar o nome, não a extensão
        Dim newBaseName = InputBox("Digite o novo nome do arquivo:", "Renomear", baseName).Trim

        If String.IsNullOrWhiteSpace(newBaseName) OrElse newBaseName = baseName Then Return

        Dim newName = newBaseName & extension
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        Try
            If File.Exists(oldPath) Then
                File.Move(oldPath, newPath)

                ' Atualiza DataRow
                dataRow("NomeArquivo") = newName
                dataRow("Local") = newPath

                GridControlFundoSelecionado.RefreshDataSource()
                SalvarConfiguracoesFundo(GridviewFundoSelecionado, ArquivosTableFundoSelecionados)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao renomear o arquivo: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' Excluir arquivos selecionados
    Private Sub BarButtonItemExcluirFundo_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemExcluirFundo.ItemClick
        Dim selectedRows = GridviewFundoSelecionado.GetSelectedRows
        If selectedRows.Length = 0 Then Return

        Dim msg = If(selectedRows.Length = 1,
                       "Deseja realmente excluir o arquivo selecionado?",
                       $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
        If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim rowsToRemove As New List(Of DataRow)

        For Each rowHandle In selectedRows
            Dim dataRow = GridviewFundoSelecionado.GetDataRow(rowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath = dataRow("Local").ToString
                Try
                    If File.Exists(filePath) Then File.Delete(filePath)
                Catch ex As Exception
                    MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                rowsToRemove.Add(dataRow)
            End If
        Next

        For Each dr In rowsToRemove
            ArquivosTableFundoSelecionados.Rows.Remove(dr)
        Next

        ' Recalcula índices
        For i = 0 To ArquivosTableFundoSelecionados.Rows.Count - 1
            ArquivosTableFundoSelecionados.Rows(i)("Indice") = i + 1
        Next

        GridControlFundoSelecionado.RefreshDataSource()
        SalvarConfiguracoesFundo(GridviewFundoSelecionado, ArquivosTableFundoSelecionados)
    End Sub
    ' Atualizar, Limpar e Abrir Pasta
    Private Sub BarButtonItemAtualizarFundo_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAtualizarFundo.ItemClick
        ' Mapeia cada GridView para sua rotina de carregamento
        Dim acoes As New Dictionary(Of Object, Action) From {
        {GridViewApelo, Sub() CarregarConfiguracoesApelo()},
        {GridViewIntervalo, Sub() CarregarConfiguracoesIntervalo()},
        {GridViewOracao, Sub() CarregarConfiguracoesOracao()},
        {GridViewAnimados, Sub() CarregarConfiguracoesAnimados()},
        {GridViewInfantis, Sub() CarregarConfiguracoesInfantis()},
        {GridViewCerimoniais, Sub() CarregarConfiguracoesCerimoniais()},
        {GridViewEspeciais, Sub() CarregarConfiguracoesEspeciais()},
        {GridViewOutros, Sub() CarregarConfiguracoesOutros()}
    }

        ' Executa a ação correspondente ao GridView selecionado
        If acoes.ContainsKey(GridviewFundoSelecionado) Then
            acoes(GridviewFundoSelecionado).Invoke
            SalvarConfiguracoesFundo(GridviewFundoSelecionado, ArquivosTableFundoSelecionados)
        End If
    End Sub

    ' Abre a pasta no Explorer
    Private Sub BarButtonItemAbrirPastaFundo_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAbrirPastaFundo.ItemClick
        ' Dicionário de GridViews e respectivas pastas
        Dim pastas As New Dictionary(Of Object, String) From {
        {GridViewApelo, PastaApelo},
        {GridViewIntervalo, PastaIntervalo},
        {GridViewOracao, PastaOracao},
        {GridViewAnimados, PastaAnimados},
        {GridViewInfantis, PastaInfantis},
        {GridViewCerimoniais, PastaCerimoniais},
        {GridViewEspeciais, PastaEspeciais},
        {GridViewOutros, PastaOutros}
    }

        ' Verifica se o GridView atual está no dicionário
        If pastas.ContainsKey(GridviewFundoSelecionado) Then
            Dim pastaSelecionada = pastas(GridviewFundoSelecionado)
            Try
                If Directory.Exists(pastaSelecionada) Then
                    Process.Start("explorer.exe", pastaSelecionada)
                Else
                    MessageBox.Show("A pasta não existe: " & pastaSelecionada,
                                "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch ex As Exception
                MessageBox.Show("Erro ao abrir a pasta: " & ex.Message,
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    'Tamanho da Fonte do GridView
    Private Sub GridControlFundo_MouseWheel(sender As Object, e As MouseEventArgs)
        ' Verifica se a tecla Ctrl está pressionada
        If My.Computer.Keyboard.CtrlKeyDown Then
            ' Define limites de tamanho de fonte
            Dim minFontSize As Single = 5
            Dim maxFontSize As Single = 20

            ' Primeiro, obtém um GridView de referência (qualquer um do TabControl)
            Dim referenciaView As DevExpress.XtraGrid.Views.Grid.GridView =
            TryCast(GridControlFundoSelecionado.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

            If referenciaView Is Nothing Then Exit Sub

            ' Obtém o tamanho atual da fonte (pegando de um GridView de referência)
            Dim currentFontSize As Single = referenciaView.Appearance.Row.Font.Size

            ' Ajusta o tamanho com base no scroll
            If e.Delta > 0 AndAlso currentFontSize < maxFontSize Then
                currentFontSize += 1
            ElseIf e.Delta < 0 AndAlso currentFontSize > minFontSize Then
                currentFontSize -= 1
            End If

            ' Cria a nova fonte
            Dim novaFonte As New Font(referenciaView.Appearance.Row.Font.FontFamily, currentFontSize)

            ' Percorre todas as páginas do TabControl
            For Each page As DevExpress.XtraTab.XtraTabPage In XtraTabControlFundos.TabPages
                ' Percorre todos os controles dentro da página
                For Each ctrl As Control In page.Controls
                    Dim grid As DevExpress.XtraGrid.GridControl = TryCast(ctrl, DevExpress.XtraGrid.GridControl)
                    If grid IsNot Nothing Then
                        Dim gv As DevExpress.XtraGrid.Views.Grid.GridView =
                        TryCast(grid.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
                        If gv IsNot Nothing Then
                            ' Aplica a nova fonte em todas as aparências relevantes
                            gv.Appearance.Row.Font = novaFonte
                            gv.Appearance.HeaderPanel.Font = novaFonte
                            gv.Appearance.FocusedRow.Font = novaFonte
                            gv.Appearance.GroupRow.Font = novaFonte
                            gv.Appearance.FooterPanel.Font = novaFonte
                        End If
                    End If
                Next
            Next

            ' Salva a preferência
            SalvarTamanhoFonteGridViewFundo(currentFontSize)
        End If
    End Sub
    ' === Preferências do Grid ===
    Public Function ObterTamanhoFonteGridViewFundo() As Single
        Dim xmlDoc As XmlDocument = CarregarXml()
        Dim tamanhoFonte As Single = 10 ' valor padrão

        Try
            Dim node As XmlNode = xmlDoc.SelectSingleNode("/Configuracoes/Preferencias/FonteGridViewFundo")
            If node IsNot Nothing Then
                Single.TryParse(node.InnerText, tamanhoFonte)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao obter tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return tamanhoFonte
    End Function
    Public Sub SalvarTamanhoFonteGridViewFundo(ByVal tamanhoFonte As Single)
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Cria raiz se não existir
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Cria nó <Preferencias> se não existir
            Dim preferenciasNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("Preferencias")
            If preferenciasNode Is Nothing Then
                preferenciasNode = xmlDoc.CreateElement("Preferencias")
                xmlDoc.DocumentElement.AppendChild(preferenciasNode)
            End If

            ' Cria ou atualiza nó <FonteGridViewFundo>
            Dim fonteNode As XmlNode = preferenciasNode.SelectSingleNode("FonteGridViewFundo")
            If fonteNode Is Nothing Then
                fonteNode = xmlDoc.CreateElement("FonteGridViewFundo")
                preferenciasNode.AppendChild(fonteNode)
            End If

            fonteNode.InnerText = tamanhoFonte.ToString(System.Globalization.CultureInfo.InvariantCulture)

            ' Salva com XmlManager
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' Salvar configurações do Grid em XML
    Private Sub SalvarConfiguracoesFundo(GridviewemFoco As DevExpress.XtraGrid.Views.Grid.GridView, DatatableemFoco As DataTable)
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Se não tiver raiz, cria
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If
            ' If GridviewFundoSelecionado Is Nothing Then Return

            ' Remove GridFundo antigo (se existir) para recriar
            Dim gridFundoNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode(GridviewemFoco.Name)
            If gridFundoNode IsNot Nothing Then
                xmlDoc.DocumentElement.RemoveChild(gridFundoNode)
            End If

            ' Cria nó GridApelo
            gridFundoNode = xmlDoc.CreateElement(GridviewemFoco.Name)
            Dim attrUltimaAlteracao As XmlAttribute = xmlDoc.CreateAttribute("UltimaAlteracao")
            attrUltimaAlteracao.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            gridFundoNode.Attributes.Append(attrUltimaAlteracao)

            ' Adiciona arquivos do DataTable
            For Each row As DataRow In DatatableemFoco.Rows
                Dim arqNode As XmlElement = xmlDoc.CreateElement("Arquivo")

                Dim attrIndice As XmlAttribute = xmlDoc.CreateAttribute("Indice")
                attrIndice.Value = row("Indice").ToString()
                arqNode.Attributes.Append(attrIndice)

                Dim attrNome As XmlAttribute = xmlDoc.CreateAttribute("NomeArquivo")
                attrNome.Value = row("NomeArquivo").ToString()
                arqNode.Attributes.Append(attrNome)

                ' >>> Novo atributo Local <<<
                Dim attrLocal As XmlAttribute = xmlDoc.CreateAttribute("Local")
                attrLocal.Value = row("Local").ToString()
                arqNode.Attributes.Append(attrLocal)

                gridFundoNode.AppendChild(arqNode)
            Next

            ' Adiciona o nó na raiz
            xmlDoc.DocumentElement.AppendChild(gridFundoNode)

            ' Salva no arquivo
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar configurações: " & ex.Message,
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region
#End Region

#Region "PROVAI E VEDE"
    Private PastaMesProvaieVede As String
    Private ArquivosTableProvaieVede As DataTable
    Private dragRowHandleProvaieVede As Integer = GridControl.InvalidRowHandle

    Private Sub IdentificarMesProvaieVede()
        Dim currentMonth As Integer = DateTime.Now.Month
        Dim MesProvaieVedeDictionary As New Dictionary(Of Integer, String) From {
        {1, "Jan"}, {2, "Fev"}, {3, "Mar"}, {4, "Abr"},
        {5, "Mai"}, {6, "Jun"}, {7, "Jul"}, {8, "Ago"},
        {9, "Set"}, {10, "Out"}, {11, "Nov"}, {12, "Dez"}
    }
        Dim CheckButtonName As String = "CheckButton" & MesProvaieVedeDictionary(currentMonth)
        Dim CheckButton As DevExpress.XtraEditors.CheckButton = CType(LayoutControlProvaieVede.Controls.Find(CheckButtonName, True).FirstOrDefault(), DevExpress.XtraEditors.CheckButton)
        CheckButton.Checked = True
    End Sub
    Private Sub CheckButtonProvaieVede_CheckedChanged(sender As Object, e As EventArgs) _
        Handles CheckButtonJan.CheckedChanged, CheckButtonFev.CheckedChanged, CheckButtonMar.CheckedChanged,
                CheckButtonAbr.CheckedChanged, CheckButtonMai.CheckedChanged, CheckButtonJun.CheckedChanged,
                CheckButtonJul.CheckedChanged, CheckButtonAgo.CheckedChanged, CheckButtonSet.CheckedChanged,
                CheckButtonOut.CheckedChanged, CheckButtonNov.CheckedChanged, CheckButtonDez.CheckedChanged

        Dim btn As CheckButton = DirectCast(sender, CheckButton)

        If btn.Checked Then
            ' Usa o texto do botão como nome da subpasta
            PastaMesProvaieVede = Path.Combine(PastaProvaieVede, btn.Text)
            CarregarGridProvaieVede()
        End If
    End Sub
    Sub CarregarGridProvaieVede()
        CarregarArquivosProvaieVede()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewProvaieVede()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
        TryCast(GridControlProvaieVede.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        GridControlProvaieVede.AllowDrop = True
        AddHandler GridControlProvaieVede.DragEnter, AddressOf PanelProvaieVede_DragEnter
        AddHandler GridControlProvaieVede.DragDrop, AddressOf PanelProvaieVede_DragDrop

        AddHandler GridViewProvaieVede.MouseUp, AddressOf GridViewProvaieVede_MouseUp
        AddHandler GridControlProvaieVede.MouseWheel, AddressOf GridControlProvaieVede_MouseWheel
    End Sub
    Private Sub CarregarArquivosProvaieVede()
        ArquivosTableProvaieVede = New DataTable()
        ArquivosTableProvaieVede.Columns.Add("Icone", GetType(Image))
        ArquivosTableProvaieVede.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableProvaieVede.Columns.Add("Local", GetType(String))

        If String.IsNullOrWhiteSpace(PastaMesProvaieVede) OrElse Not Directory.Exists(PastaMesProvaieVede) Then
            GridControlProvaieVede.DataSource = ArquivosTableProvaieVede
            Return
        End If

        Dim arquivos = Directory.GetFiles(PastaMesProvaieVede)

        For Each arq As String In arquivos
            Dim img As Image
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableProvaieVede.Rows.Add(img, Path.GetFileName(arq), arq)
        Next

        GridControlProvaieVede.DataSource = ArquivosTableProvaieVede

    End Sub
    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelProvaieVede_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelProvaieVede_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaProvaieVede, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableProvaieVede.Rows.Add(img, Path.GetFileName(destPath), destPath)
            Next

            GridControlProvaieVede.RefreshDataSource()
        End If
    End Sub
    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewProvaieVede_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewProvaieVede.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableProvaieVede.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableProvaieVede.Rows.Count - 1
                ArquivosTableProvaieVede.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlProvaieVede.RefreshDataSource()
        End If
    End Sub
    ' Abrir arquivo com duplo clique
    Private Sub GridViewProvaieVede_DoubleClick(sender As Object, e As EventArgs) Handles GridViewProvaieVede.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub
    ' Abrir arquivo com Enter
    Private Sub GridViewProvaieVede_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewProvaieVede.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub
    ' Menu de contexto com clique direito
    Private Sub GridViewProvaieVede_MouseUp(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView)
            Dim grid As DevExpress.XtraGrid.GridControl = view.GridControl

            Dim hi As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(e.Location)

            If hi.InRowCell Then
                ' Clicou em uma célula válida -> move o foco e seleciona a linha
                view.FocusedRowHandle = hi.RowHandle
                view.ClearSelection()
                view.SelectRow(hi.RowHandle)

                ' Habilita os botões do PopupMenu
                BarButtonItemExcluirProvaieVede.Enabled = True
                BarButtonItemRenomearProvaieVede.Enabled = True

                ' Exibe o menu
                PopupMenuProvaieVede.ShowPopup(Control.MousePosition)
            Else
                ' Clicou fora de uma célula -> limpa seleção
                view.ClearSelection()
                view.FocusedRowHandle = DevExpress.XtraGrid.GridControl.InvalidRowHandle

                ' Desabilita os botões do PopupMenu
                BarButtonItemExcluirProvaieVede.Enabled = False
                BarButtonItemRenomearProvaieVede.Enabled = False

            End If
        End If
    End Sub
    Private Sub BarButtonItemRenomearProvaieVede_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemRenomearProvaieVede.ItemClick
        Dim rowHandle = GridViewProvaieVede.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridViewProvaieVede.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim oldName = dataRow("NomeArquivo").ToString
        Dim oldPath = dataRow("Local").ToString

        ' Separa nome e extensão
        Dim baseName = Path.GetFileNameWithoutExtension(oldName)
        Dim extension = Path.GetExtension(oldName)

        ' Só permite editar o nome, não a extensão
        Dim newBaseName = InputBox("Digite o novo nome do arquivo:", "Renomear", baseName).Trim

        If String.IsNullOrWhiteSpace(newBaseName) OrElse newBaseName = baseName Then Return

        Dim newName = newBaseName & extension
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        Try
            If File.Exists(oldPath) Then
                File.Move(oldPath, newPath)

                ' Atualiza DataRow
                dataRow("NomeArquivo") = newName
                dataRow("Local") = newPath

                GridControlProvaieVede.RefreshDataSource
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao renomear o arquivo: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' Excluir arquivos selecionados
    Private Sub BarButtonItemExcluirProvaieVede_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemExcluirProvaieVede.ItemClick
        Dim selectedRows = GridViewProvaieVede.GetSelectedRows
        If selectedRows.Length = 0 Then Return

        Dim msg = If(selectedRows.Length = 1,
                       "Deseja realmente excluir o arquivo selecionado?",
                       $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
        If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim rowsToRemove As New List(Of DataRow)

        For Each rowHandle In selectedRows
            Dim dataRow = GridViewProvaieVede.GetDataRow(rowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath = dataRow("Local").ToString
                Try
                    If File.Exists(filePath) Then File.Delete(filePath)
                Catch ex As Exception
                    MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                rowsToRemove.Add(dataRow)
            End If
        Next

        For Each dr In rowsToRemove
            ArquivosTableProvaieVede.Rows.Remove(dr)
        Next

        ' Recalcula índices
        For i = 0 To ArquivosTableProvaieVede.Rows.Count - 1
            ArquivosTableProvaieVede.Rows(i)("Indice") = i + 1
        Next

        GridControlProvaieVede.RefreshDataSource
    End Sub
    ' Atualizar, Limpar e Abrir Pasta
    Private Sub BarButtonItemAtualizarProvaieVede_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAtualizarProvaieVede.ItemClick
        CarregarArquivosProvaieVede
    End Sub
    ' Abre a pasta no Explorer
    Private Sub BarButtonItemAbrirPastaProvaieVede_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAbrirPastaProvaieVede.ItemClick
        Try
            If Directory.Exists(PastaProvaieVede) Then
                Process.Start("explorer.exe", PastaProvaieVede)
            Else
                MessageBox.Show("A pasta não existe: " & PastaProvaieVede,
                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir a pasta: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'Tamanho da Fonte do GridView
    Private Sub GridControlProvaieVede_MouseWheel(sender As Object, e As MouseEventArgs)
        ' Verifica se a tecla Ctrl está pressionada
        If Not My.Computer.Keyboard.CtrlKeyDown Then Return

        ' Lista de GridViews que terão a fonte ajustada
        Dim grids As DevExpress.XtraGrid.Views.Grid.GridView() = {
        TryCast(GridControlJovens.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlDbv.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlAvt.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlProvaieVede.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
    }

        ' Define limites da fonte
        Dim minFontSize As Single = 5
        Dim maxFontSize As Single = 20

        ' Obtém o tamanho atual da fonte do primeiro GridView válido
        Dim currentFontSize As Single = grids.FirstOrDefault(Function(g) g IsNot Nothing)?.Appearance.Row.Font.Size
        If currentFontSize = 0 Then Return

        ' Ajusta o tamanho da fonte com base na direção do scroll
        If e.Delta > 0 AndAlso currentFontSize < maxFontSize Then
            currentFontSize += 1
        ElseIf e.Delta < 0 AndAlso currentFontSize > minFontSize Then
            currentFontSize -= 1
        End If

        ' Aplica a nova fonte a todos os GridViews válidos
        Dim novaFonte As New Font(grids.First(Function(g) g IsNot Nothing).Appearance.Row.Font.FontFamily, currentFontSize)
        For Each grid In grids
            If grid IsNot Nothing Then
                grid.Appearance.Row.Font = novaFonte
                grid.Appearance.HeaderPanel.Font = novaFonte
                grid.Appearance.FocusedRow.Font = novaFonte
                grid.Appearance.GroupRow.Font = novaFonte
                grid.Appearance.FooterPanel.Font = novaFonte
            End If
        Next

        ' Salva tamanho da fonte (adaptar conforme Grid específico se necessário)
        SalvarTamanhoFonteGridViewProvaieVede(currentFontSize)
    End Sub
    ' === Preferências do Grid ===
    Public Function ObterTamanhoFonteGridViewProvaieVede() As Single
        Dim xmlDoc As XmlDocument = CarregarXml()
        Dim tamanhoFonte As Single = 10 ' valor padrão

        Try
            Dim node As XmlNode = xmlDoc.SelectSingleNode("/Configuracoes/Preferencias/FonteGridViewProvaieVede")
            If node IsNot Nothing Then
                Single.TryParse(node.InnerText, tamanhoFonte)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao obter tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return tamanhoFonte
    End Function
    Public Sub SalvarTamanhoFonteGridViewProvaieVede(ByVal tamanhoFonte As Single)
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Cria raiz se não existir
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Cria nó <Preferencias> se não existir
            Dim preferenciasNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("Preferencias")
            If preferenciasNode Is Nothing Then
                preferenciasNode = xmlDoc.CreateElement("Preferencias")
                xmlDoc.DocumentElement.AppendChild(preferenciasNode)
            End If

            ' Cria ou atualiza nó <FonteGridViewProvaieVede>
            Dim fonteNode As XmlNode = preferenciasNode.SelectSingleNode("FonteGridViewProvaieVede")
            If fonteNode Is Nothing Then
                fonteNode = xmlDoc.CreateElement("FonteGridViewProvaieVede")
                preferenciasNode.AppendChild(fonteNode)
            End If

            fonteNode.InnerText = tamanhoFonte.ToString(System.Globalization.CultureInfo.InvariantCulture)

            ' Salva com XmlManager
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "DBV"
    Private ArquivosTableDbv As DataTable
    Private dragRowHandleDbv As Integer = GridControl.InvalidRowHandle
    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableDbv()
        ArquivosTableDbv = New DataTable()
        ArquivosTableDbv.Columns.Add("Indice", GetType(Integer))
        ArquivosTableDbv.Columns.Add("Icone", GetType(Image))
        ArquivosTableDbv.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableDbv.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridDbv()
        CarregarConfiguracoesDbv()
        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewProvaieVede()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControlDbv.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewDbv)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorDbv_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorDbv_DragOver

        XtraTabPageDbv.AllowDrop = True
        AddHandler XtraTabPageDbv.DragEnter, AddressOf PanelDbv_DragEnter
        AddHandler XtraTabPageDbv.DragDrop, AddressOf PanelDbv_DragDrop
    End Sub
    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosDbv()
        PrepararArquivosTableDbv()

        Dim arquivos = Directory.GetFiles(PastaDbv)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableDbv.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlDbv.DataSource = ArquivosTableDbv

        SalvarConfiguracoesDbv()
    End Sub

    ' Carregar configurações do GridDbv do XML
    Private Sub CarregarConfiguracoesDbv()
        PrepararArquivosTableDbv()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosDbv()
            Return
        End If

        Try
            Dim gridDbvNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridDbv")
            If gridDbvNode Is Nothing Then
                CarregarArquivosDbv()
                Return
            End If

            Dim arquivos As XmlNodeList = gridDbvNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaDbv, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableDbv.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes
            For Each r As XmlNode In remover
                gridDbvNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaDbv)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridDbvNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridDbvNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableDbv.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridDbvNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridDbvNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridDbvNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableDbv.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlDbv.DataSource = ArquivosTableDbv

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorDbv_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorDbv_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlDbv.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableDbv.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableDbv.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableDbv.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableDbv.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableDbv.Rows.Count Then targetIndex = ArquivosTableDbv.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableDbv.Rows.Count Then
            ArquivosTableDbv.Rows.Add(newRow)
            targetIndex = ArquivosTableDbv.Rows.Count - 1
        Else
            ArquivosTableDbv.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableDbv.Rows.Count - 1
            ArquivosTableDbv.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlDbv.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesDbv()

        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelDbv_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelDbv_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableDbv.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaDbv, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableDbv.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlDbv.RefreshDataSource()
            SalvarConfiguracoesDbv()
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewDbv_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewDbv.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableDbv.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableDbv.Rows.Count - 1
                ArquivosTableDbv.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlDbv.RefreshDataSource()
            SalvarConfiguracoesDbv()
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewDbv_DoubleClick(sender As Object, e As EventArgs) Handles GridViewDbv.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewDbv_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewDbv.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub

    '   Menu de contexto com clique direito
    Private Sub GridViewDbv_MouseUp(sender As Object, e As MouseEventArgs) Handles GridViewDbv.MouseUp
        If e.Button = MouseButtons.Right Then
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView)
            Dim hi As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(e.Location)

            If hi.InRowCell Then
                ' Clicou em uma célula válida -> move o foco e seleciona a linha
                view.FocusedRowHandle = hi.RowHandle
                view.ClearSelection()
                view.SelectRow(hi.RowHandle)

                ' Habilita os botões do PopupMenu
                BarButtonItemExcluirDbv.Enabled = True
                BarButtonItemRenomearDbv.Enabled = True

                ' Exibe o menu
                PopupMenuDbv.ShowPopup(Control.MousePosition)
            Else
                ' Clicou fora de uma célula -> limpa seleção
                view.ClearSelection()
                view.FocusedRowHandle = DevExpress.XtraGrid.GridControl.InvalidRowHandle

                ' Desabilita os botões do PopupMenu
                BarButtonItemExcluirDbv.Enabled = False
                BarButtonItemRenomearDbv.Enabled = False

                PopupMenuDbv.ShowPopup(Control.MousePosition)

            End If
        End If
    End Sub
    Private Sub BarButtonItemRenomearDbv_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemRenomearDbv.ItemClick
        Dim rowHandle = GridViewDbv.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridViewDbv.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim oldName = dataRow("NomeArquivo").ToString
        Dim oldPath = dataRow("Local").ToString

        ' Separa nome e extensão
        Dim baseName = Path.GetFileNameWithoutExtension(oldName)
        Dim extension = Path.GetExtension(oldName)

        ' Só permite editar o nome, não a extensão
        Dim newBaseName = InputBox("Digite o novo nome do arquivo:", "Renomear", baseName).Trim

        If String.IsNullOrWhiteSpace(newBaseName) OrElse newBaseName = baseName Then Return

        Dim newName = newBaseName & extension
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        Try
            If File.Exists(oldPath) Then
                File.Move(oldPath, newPath)

                ' Atualiza DataRow
                dataRow("NomeArquivo") = newName
                dataRow("Local") = newPath

                GridControlDbv.RefreshDataSource
                SalvarConfiguracoesDbv
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao renomear o arquivo: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Excluir arquivos selecionados
    Private Sub BarButtonItemExcluirDbv_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemExcluirDbv.ItemClick
        Dim selectedRows = GridViewDbv.GetSelectedRows
        If selectedRows.Length = 0 Then Return

        Dim msg = If(selectedRows.Length = 1,
                       "Deseja realmente excluir o arquivo selecionado?",
                       $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
        If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim rowsToRemove As New List(Of DataRow)

        For Each rowHandle In selectedRows
            Dim dataRow = GridViewDbv.GetDataRow(rowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath = dataRow("Local").ToString
                Try
                    If File.Exists(filePath) Then File.Delete(filePath)
                Catch ex As Exception
                    MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                rowsToRemove.Add(dataRow)
            End If
        Next

        For Each dr In rowsToRemove
            ArquivosTableDbv.Rows.Remove(dr)
        Next

        ' Recalcula índices
        For i = 0 To ArquivosTableDbv.Rows.Count - 1
            ArquivosTableDbv.Rows(i)("Indice") = i + 1
        Next

        GridControlDbv.RefreshDataSource
        SalvarConfiguracoesDbv
    End Sub

    ' Atualizar, Limpar e Abrir Pasta
    Private Sub BarButtonItemAtualizarDbv_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAtualizarDbv.ItemClick
        CarregarConfiguracoesDbv()
        SalvarConfiguracoesDbv()
    End Sub

    ' Abre a pasta Dbv no Explorer
    Private Sub BarButtonItemAbrirPastaDbv_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAbrirPastaDbv.ItemClick
        Try
            If Directory.Exists(PastaDbv) Then
                Process.Start("explorer.exe", PastaDbv)
            Else
                MessageBox.Show("A pasta não existe: " & PastaDbv,
                            "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir a pasta: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' Salvar configurações do GridDbv em XML
    Private Sub SalvarConfiguracoesDbv()
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Se não tiver raiz, cria
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Remove GridDbv antigo (se existir) para recriar
            Dim gridDbvNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridDbv")
            If gridDbvNode IsNot Nothing Then
                xmlDoc.DocumentElement.RemoveChild(gridDbvNode)
            End If

            ' Cria nó GridDbv
            gridDbvNode = xmlDoc.CreateElement("GridDbv")
            Dim attrUltimaAlteracao As XmlAttribute = xmlDoc.CreateAttribute("UltimaAlteracao")
            attrUltimaAlteracao.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            gridDbvNode.Attributes.Append(attrUltimaAlteracao)

            ' Adiciona arquivos do DataTable
            For Each row As DataRow In ArquivosTableDbv.Rows
                Dim arqNode As XmlElement = xmlDoc.CreateElement("Arquivo")

                Dim attrIndice As XmlAttribute = xmlDoc.CreateAttribute("Indice")
                attrIndice.Value = row("Indice").ToString()
                arqNode.Attributes.Append(attrIndice)

                Dim attrNome As XmlAttribute = xmlDoc.CreateAttribute("NomeArquivo")
                attrNome.Value = row("NomeArquivo").ToString()
                arqNode.Attributes.Append(attrNome)

                ' >>> Novo atributo Local <<<
                Dim attrLocal As XmlAttribute = xmlDoc.CreateAttribute("Local")
                attrLocal.Value = row("Local").ToString()
                arqNode.Attributes.Append(attrLocal)

                gridDbvNode.AppendChild(arqNode)
            Next

            ' Adiciona o nó na raiz
            xmlDoc.DocumentElement.AppendChild(gridDbvNode)

            ' Salva no arquivo
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar configurações: " & ex.Message,
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'Tamanho da Fonte do GridViewProcessos
    Private Sub GridControlDbv_MouseWheel(sender As Object, e As MouseEventArgs) Handles GridControlDbv.MouseWheel
        ' Verifica se a tecla Ctrl está pressionada
        If Not My.Computer.Keyboard.CtrlKeyDown Then Return

        ' Lista de GridViews que terão a fonte ajustada
        Dim grids As DevExpress.XtraGrid.Views.Grid.GridView() = {
        TryCast(GridControlJovens.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlDbv.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlAvt.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlProvaieVede.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
    }

        ' Define limites da fonte
        Dim minFontSize As Single = 5
        Dim maxFontSize As Single = 20

        ' Obtém o tamanho atual da fonte do primeiro GridView válido
        Dim currentFontSize As Single = grids.FirstOrDefault(Function(g) g IsNot Nothing)?.Appearance.Row.Font.Size
        If currentFontSize = 0 Then Return

        ' Ajusta o tamanho da fonte com base na direção do scroll
        If e.Delta > 0 AndAlso currentFontSize < maxFontSize Then
            currentFontSize += 1
        ElseIf e.Delta < 0 AndAlso currentFontSize > minFontSize Then
            currentFontSize -= 1
        End If

        ' Aplica a nova fonte a todos os GridViews válidos
        Dim novaFonte As New Font(grids.First(Function(g) g IsNot Nothing).Appearance.Row.Font.FontFamily, currentFontSize)
        For Each grid In grids
            If grid IsNot Nothing Then
                grid.Appearance.Row.Font = novaFonte
                grid.Appearance.HeaderPanel.Font = novaFonte
                grid.Appearance.FocusedRow.Font = novaFonte
                grid.Appearance.GroupRow.Font = novaFonte
                grid.Appearance.FooterPanel.Font = novaFonte
            End If
        Next

        ' Salva tamanho da fonte (adaptar conforme Grid específico se necessário)
        SalvarTamanhoFonteGridViewProvaieVede(currentFontSize)
    End Sub

#End Region

#Region "AVT"
    Private ArquivosTableAvt As DataTable
    Private dragRowHandleAvt As Integer = GridControl.InvalidRowHandle
    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableAvt()
        ArquivosTableAvt = New DataTable()
        ArquivosTableAvt.Columns.Add("Indice", GetType(Integer))
        ArquivosTableAvt.Columns.Add("Icone", GetType(Image))
        ArquivosTableAvt.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableAvt.Columns.Add("Local", GetType(String))
    End Sub
    Sub CarregarGridAvt()
        CarregarConfiguracoesAvt()
        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewProvaieVede()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControlAvt.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewAvt)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorAvt_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorAvt_DragOver

        XtraTabPageAvt.AllowDrop = True
        AddHandler XtraTabPageAvt.DragEnter, AddressOf PanelAvt_DragEnter
        AddHandler XtraTabPageAvt.DragDrop, AddressOf PanelAvt_DragDrop
    End Sub
    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosAvt()
        PrepararArquivosTableAvt()

        Dim arquivos = Directory.GetFiles(PastaAvt)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableAvt.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlAvt.DataSource = ArquivosTableAvt
        SalvarConfiguracoesAvt()
    End Sub
    ' Carregar configurações do GridAvt do XML
    Private Sub CarregarConfiguracoesAvt()
        PrepararArquivosTableAvt()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosAvt()
            Return
        End If

        Try
            Dim gridAvtNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridAvt")
            If gridAvtNode Is Nothing Then
                CarregarArquivosAvt()
                Return
            End If

            Dim arquivos As XmlNodeList = gridAvtNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaAvt, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableAvt.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes
            For Each r As XmlNode In remover
                gridAvtNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaAvt)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridAvtNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridAvtNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableAvt.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridAvtNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridAvtNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridAvtNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableAvt.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlAvt.DataSource = ArquivosTableAvt

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorAvt_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorAvt_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlAvt.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableAvt.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableAvt.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableAvt.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableAvt.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableAvt.Rows.Count Then targetIndex = ArquivosTableAvt.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableAvt.Rows.Count Then
            ArquivosTableAvt.Rows.Add(newRow)
            targetIndex = ArquivosTableAvt.Rows.Count - 1
        Else
            ArquivosTableAvt.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableAvt.Rows.Count - 1
            ArquivosTableAvt.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlAvt.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesAvt()

        e.Handled = True
    End Sub
    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelAvt_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelAvt_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableAvt.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaAvt, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableAvt.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlAvt.RefreshDataSource()
            SalvarConfiguracoesAvt()
        End If
    End Sub
    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewAvt_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewAvt.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableAvt.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableAvt.Rows.Count - 1
                ArquivosTableAvt.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlAvt.RefreshDataSource()
            SalvarConfiguracoesAvt()
        End If
    End Sub
    ' Abrir arquivo com duplo clique
    Private Sub GridViewAvt_DoubleClick(sender As Object, e As EventArgs) Handles GridViewAvt.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub
    ' Abrir arquivo com Enter
    Private Sub GridViewAvt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewAvt.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub
    '   Menu de contexto com clique direito
    Private Sub GridViewAvt_MouseUp(sender As Object, e As MouseEventArgs) Handles GridViewAvt.MouseUp
        If e.Button = MouseButtons.Right Then
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView)
            Dim hi As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(e.Location)

            If hi.InRowCell Then
                ' Clicou em uma célula válida -> move o foco e seleciona a linha
                view.FocusedRowHandle = hi.RowHandle
                view.ClearSelection()
                view.SelectRow(hi.RowHandle)

                ' Habilita os botões do PopupMenu
                BarButtonItemExcluirAvt.Enabled = True
                BarButtonItemRenomearAvt.Enabled = True

                ' Exibe o menu
                PopupMenuAvt.ShowPopup(Control.MousePosition)
            Else
                ' Clicou fora de uma célula -> limpa seleção
                view.ClearSelection()
                view.FocusedRowHandle = DevExpress.XtraGrid.GridControl.InvalidRowHandle

                ' Desabilita os botões do PopupMenu
                BarButtonItemExcluirAvt.Enabled = False
                BarButtonItemRenomearAvt.Enabled = False

                PopupMenuAvt.ShowPopup(Control.MousePosition)

            End If
        End If
    End Sub
    Private Sub BarButtonItemRenomearAvt_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemRenomearAvt.ItemClick
        Dim rowHandle = GridViewAvt.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridViewAvt.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim oldName = dataRow("NomeArquivo").ToString
        Dim oldPath = dataRow("Local").ToString

        ' Separa nome e extensão
        Dim baseName = Path.GetFileNameWithoutExtension(oldName)
        Dim extension = Path.GetExtension(oldName)

        ' Só permite editar o nome, não a extensão
        Dim newBaseName = InputBox("Digite o novo nome do arquivo:", "Renomear", baseName).Trim

        If String.IsNullOrWhiteSpace(newBaseName) OrElse newBaseName = baseName Then Return

        Dim newName = newBaseName & extension
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        Try
            If File.Exists(oldPath) Then
                File.Move(oldPath, newPath)

                ' Atualiza DataRow
                dataRow("NomeArquivo") = newName
                dataRow("Local") = newPath

                GridControlAvt.RefreshDataSource
                SalvarConfiguracoesAvt
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao renomear o arquivo: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' Excluir arquivos selecionados
    Private Sub BarButtonItemExcluirAvt_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemExcluirAvt.ItemClick
        Dim selectedRows = GridViewAvt.GetSelectedRows
        If selectedRows.Length = 0 Then Return

        Dim msg = If(selectedRows.Length = 1,
                       "Deseja realmente excluir o arquivo selecionado?",
                       $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
        If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim rowsToRemove As New List(Of DataRow)

        For Each rowHandle In selectedRows
            Dim dataRow = GridViewAvt.GetDataRow(rowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath = dataRow("Local").ToString
                Try
                    If File.Exists(filePath) Then File.Delete(filePath)
                Catch ex As Exception
                    MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                rowsToRemove.Add(dataRow)
            End If
        Next

        For Each dr In rowsToRemove
            ArquivosTableAvt.Rows.Remove(dr)
        Next

        ' Recalcula índices
        For i = 0 To ArquivosTableAvt.Rows.Count - 1
            ArquivosTableAvt.Rows(i)("Indice") = i + 1
        Next

        GridControlAvt.RefreshDataSource
        SalvarConfiguracoesAvt
    End Sub
    ' Atualizar, Limpar e Abrir Pasta
    Private Sub BarButtonItemAtualizarAvt_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAtualizarAvt.ItemClick
        CarregarConfiguracoesAvt()
        SalvarConfiguracoesAvt()
    End Sub
    ' Abre a pasta Avt no Explorer
    Private Sub BarButtonItemAbrirPastaAvt_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAbrirPastaAvt.ItemClick
        Try
            If Directory.Exists(PastaAvt) Then
                Process.Start("explorer.exe", PastaAvt)
            Else
                MessageBox.Show("A pasta não existe: " & PastaAvt,
                            "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir a pasta: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' Salvar configurações do GridAvt em XML
    Private Sub SalvarConfiguracoesAvt()
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Se não tiver raiz, cria
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Remove GridAvt antigo (se existir) para recriar
            Dim gridAvtNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridAvt")
            If gridAvtNode IsNot Nothing Then
                xmlDoc.DocumentElement.RemoveChild(gridAvtNode)
            End If

            ' Cria nó GridAvt
            gridAvtNode = xmlDoc.CreateElement("GridAvt")
            Dim attrUltimaAlteracao As XmlAttribute = xmlDoc.CreateAttribute("UltimaAlteracao")
            attrUltimaAlteracao.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            gridAvtNode.Attributes.Append(attrUltimaAlteracao)

            ' Adiciona arquivos do DataTable
            For Each row As DataRow In ArquivosTableAvt.Rows
                Dim arqNode As XmlElement = xmlDoc.CreateElement("Arquivo")

                Dim attrIndice As XmlAttribute = xmlDoc.CreateAttribute("Indice")
                attrIndice.Value = row("Indice").ToString()
                arqNode.Attributes.Append(attrIndice)

                Dim attrNome As XmlAttribute = xmlDoc.CreateAttribute("NomeArquivo")
                attrNome.Value = row("NomeArquivo").ToString()
                arqNode.Attributes.Append(attrNome)

                ' >>> Novo atributo Local <<<
                Dim attrLocal As XmlAttribute = xmlDoc.CreateAttribute("Local")
                attrLocal.Value = row("Local").ToString()
                arqNode.Attributes.Append(attrLocal)

                gridAvtNode.AppendChild(arqNode)
            Next

            ' Adiciona o nó na raiz
            xmlDoc.DocumentElement.AppendChild(gridAvtNode)

            ' Salva no arquivo
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar configurações: " & ex.Message,
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'Tamanho da Fonte do GridViewProcessos
    Private Sub GridControlAvt_MouseWheel(sender As Object, e As MouseEventArgs) Handles GridControlAvt.MouseWheel
        ' Verifica se a tecla Ctrl está pressionada
        If Not My.Computer.Keyboard.CtrlKeyDown Then Return

        ' Lista de GridViews que terão a fonte ajustada
        Dim grids As DevExpress.XtraGrid.Views.Grid.GridView() = {
        TryCast(GridControlJovens.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlDbv.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlAvt.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlProvaieVede.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
    }

        ' Define limites da fonte
        Dim minFontSize As Single = 5
        Dim maxFontSize As Single = 20

        ' Obtém o tamanho atual da fonte do primeiro GridView válido
        Dim currentFontSize As Single = grids.FirstOrDefault(Function(g) g IsNot Nothing)?.Appearance.Row.Font.Size
        If currentFontSize = 0 Then Return

        ' Ajusta o tamanho da fonte com base na direção do scroll
        If e.Delta > 0 AndAlso currentFontSize < maxFontSize Then
            currentFontSize += 1
        ElseIf e.Delta < 0 AndAlso currentFontSize > minFontSize Then
            currentFontSize -= 1
        End If

        ' Aplica a nova fonte a todos os GridViews válidos
        Dim novaFonte As New Font(grids.First(Function(g) g IsNot Nothing).Appearance.Row.Font.FontFamily, currentFontSize)
        For Each grid In grids
            If grid IsNot Nothing Then
                grid.Appearance.Row.Font = novaFonte
                grid.Appearance.HeaderPanel.Font = novaFonte
                grid.Appearance.FocusedRow.Font = novaFonte
                grid.Appearance.GroupRow.Font = novaFonte
                grid.Appearance.FooterPanel.Font = novaFonte
            End If
        Next

        ' Salva tamanho da fonte (adaptar conforme Grid específico se necessário)
        SalvarTamanhoFonteGridViewProvaieVede(currentFontSize)
    End Sub

#End Region

#Region "JOVENS"
    Private ArquivosTableJovens As DataTable
    Private dragRowHandleJovens As Integer = GridControl.InvalidRowHandle
    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableJovens()
        ArquivosTableJovens = New DataTable()
        ArquivosTableJovens.Columns.Add("Indice", GetType(Integer))
        ArquivosTableJovens.Columns.Add("Icone", GetType(Image))
        ArquivosTableJovens.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableJovens.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridJovens()
        CarregarConfiguracoesJovens()
        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewProvaieVede()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControlJovens.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        Dim gridControlBehavior As DragDropBehavior = BehaviorManager.GetBehavior(Of DragDropBehavior)(Me.GridViewJovens)
        AddHandler gridControlBehavior.DragDrop, AddressOf BehaviorJovens_DragDrop
        AddHandler gridControlBehavior.DragOver, AddressOf BehaviorJovens_DragOver

        XtraTabPageJovens.AllowDrop = True
        AddHandler XtraTabPageJovens.DragEnter, AddressOf PanelJovens_DragEnter
        AddHandler XtraTabPageJovens.DragDrop, AddressOf PanelJovens_DragDrop
    End Sub
    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosJovens()
        PrepararArquivosTableJovens()

        Dim arquivos = Directory.GetFiles(PastaJovens)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableJovens.Rows.Add(i + 1, img, Path.GetFileName(arq), arq)
        Next

        GridControlJovens.DataSource = ArquivosTableJovens
        SalvarConfiguracoesJovens()
    End Sub

    ' Carregar configurações do GridJovens do XML
    Private Sub CarregarConfiguracoesJovens()
        PrepararArquivosTableJovens()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosJovens()
            Return
        End If

        Try
            Dim gridJovensNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridJovens")
            If gridJovensNode Is Nothing Then
                CarregarArquivosJovens()
                Return
            End If

            Dim arquivos As XmlNodeList = gridJovensNode.SelectNodes("Arquivo")
            Dim indice As Integer = 1
            Dim remover As New List(Of XmlNode)

            ' Arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim nomeArquivo As String = arqNode.Attributes("NomeArquivo")?.Value
                Dim filePath As String = Path.Combine(PastaJovens, nomeArquivo)

                If File.Exists(filePath) Then
                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableJovens.Rows.Add(indice, img, nomeArquivo, filePath)
                    indice += 1
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remove arquivos inexistentes
            For Each r As XmlNode In remover
                gridJovensNode.RemoveChild(r)
            Next

            ' Adiciona arquivos novos que não estão no XML
            Dim arquivosNaPasta = Directory.GetFiles(PastaJovens)
            For Each filePath In arquivosNaPasta
                Dim nomeArquivo As String = Path.GetFileName(filePath)
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridJovensNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("NomeArquivo")?.Value, nomeArquivo, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("Indice", indice.ToString())
                    novoNode.SetAttribute("NomeArquivo", nomeArquivo)
                    gridJovensNode.AppendChild(novoNode)

                    Dim img As Image = Nothing
                    Try
                        Dim icone As Icon = Icon.ExtractAssociatedIcon(filePath)
                        img = icone.ToBitmap()
                    Catch
                        img = SystemIcons.WinLogo.ToBitmap()
                    End Try
                    ArquivosTableJovens.Rows.Add(indice, img, nomeArquivo, filePath)

                    indice += 1
                End If
            Next

            ' Atualiza data/hora (seguro)
            Dim attr = gridJovensNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridJovensNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' Reordena índices no XML
            Dim i As Integer = 1
            For Each arqNode As XmlNode In gridJovensNode.SelectNodes("Arquivo")
                arqNode.Attributes("Indice").Value = i.ToString()
                i += 1
            Next

            ' Reordena também no DataTable
            i = 1
            For Each row As DataRow In ArquivosTableJovens.Rows
                row("Indice") = i
                i += 1
            Next

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlJovens.DataSource = ArquivosTableJovens

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    'Reordenar linhas com Drag & Drop
    Private Sub BehaviorJovens_DragOver(ByVal sender As Object, ByVal e As DragOverEventArgs)
        Dim args As DragOverGridEventArgs = DragOverGridEventArgs.GetDragOverGridEventArgs(e)
        e.InsertType = args.InsertType
        e.InsertIndicatorLocation = args.InsertIndicatorLocation
        e.Action = args.Action
        Cursor.Current = args.Cursor
        args.Handled = True
    End Sub
    Private Sub BehaviorJovens_DragDrop(ByVal sender As Object, ByVal e As DragDropEventArgs)
        Dim args As DragDropGridEventArgs = DragDropGridEventArgs.GetDragDropGridEventArgs(e)
        If args Is Nothing Then Return

        Dim view As GridView = TryCast(GridControlJovens.MainView, GridView)
        If view Is Nothing Then Return

        ' Recupera a linha arrastada
        Dim draggedRow As DataRow = Nothing
        If args.DataRows.Length > 0 Then
            Dim drv As DataRowView = TryCast(args.DataRows(0), DataRowView)
            If drv IsNot Nothing Then
                draggedRow = drv.Row
            ElseIf TypeOf args.DataRows(0) Is DataRow Then
                draggedRow = DirectCast(args.DataRows(0), DataRow)
            End If
        End If
        If draggedRow Is Nothing Then Return

        ' Índice da linha arrastada
        Dim sourceIndex As Integer = ArquivosTableJovens.Rows.IndexOf(draggedRow)

        ' Determina o índice de destino
        Dim targetIndex As Integer = ArquivosTableJovens.Rows.Count - 1 ' default = fim
        If args.HitInfo IsNot Nothing Then
            Dim hitHandle As Integer = args.HitInfo.RowHandle
            If hitHandle >= 0 AndAlso hitHandle < view.DataRowCount Then
                targetIndex = view.GetDataSourceRowIndex(hitHandle)
                If args.InsertType = DevExpress.Utils.DragDrop.InsertType.After Then
                    targetIndex += 1
                End If
            End If
        End If

        ' Se a posição não mudou, não faz nada
        If sourceIndex = targetIndex Or sourceIndex = targetIndex - 1 Then Return

        ' Clona os dados da linha arrastada
        Dim newRow As DataRow = ArquivosTableJovens.NewRow()
        newRow.ItemArray = CType(draggedRow.ItemArray.Clone(), Object())

        ' Remove a linha original
        ArquivosTableJovens.Rows.RemoveAt(sourceIndex)

        ' Ajusta targetIndex se necessário
        If targetIndex > sourceIndex Then targetIndex -= 1
        If targetIndex > ArquivosTableJovens.Rows.Count Then targetIndex = ArquivosTableJovens.Rows.Count

        ' Insere a linha na posição correta
        If targetIndex = ArquivosTableJovens.Rows.Count Then
            ArquivosTableJovens.Rows.Add(newRow)
            targetIndex = ArquivosTableJovens.Rows.Count - 1
        Else
            ArquivosTableJovens.Rows.InsertAt(newRow, targetIndex)
        End If

        ' Recalcula a coluna Índice
        For i As Integer = 0 To ArquivosTableJovens.Rows.Count - 1
            ArquivosTableJovens.Rows(i)("Indice") = i + 1
        Next

        ' Atualiza a grid e mantém seleção/foco
        GridControlJovens.RefreshDataSource()
        view.ClearSelection()
        view.FocusedRowHandle = targetIndex
        view.SelectRow(targetIndex)

        SalvarConfiguracoesJovens()

        e.Handled = True
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelJovens_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelJovens_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim nextIndex As Integer = ArquivosTableJovens.Rows.Count + 1

            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaJovens, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableJovens.Rows.Add(nextIndex, img, Path.GetFileName(destPath), destPath)
                nextIndex += 1
            Next

            GridControlJovens.RefreshDataSource()
            SalvarConfiguracoesJovens()
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewJovens_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewJovens.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableJovens.Rows.Remove(dr)
            Next

            ' Recalcula os índices
            For i As Integer = 0 To ArquivosTableJovens.Rows.Count - 1
                ArquivosTableJovens.Rows(i)("Indice") = i + 1
            Next

            ' Atualiza Grid
            GridControlJovens.RefreshDataSource()
            SalvarConfiguracoesJovens()
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewJovens_DoubleClick(sender As Object, e As EventArgs) Handles GridViewJovens.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewJovens_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewJovens.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub

    '   Menu de contexto com clique direito
    Private Sub GridViewJovens_MouseUp(sender As Object, e As MouseEventArgs) Handles GridViewJovens.MouseUp
        If e.Button = MouseButtons.Right Then
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView)
            Dim hi As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(e.Location)

            If hi.InRowCell Then
                ' Clicou em uma célula válida -> move o foco e seleciona a linha
                view.FocusedRowHandle = hi.RowHandle
                view.ClearSelection()
                view.SelectRow(hi.RowHandle)

                ' Habilita os botões do PopupMenu
                BarButtonItemExcluirJovens.Enabled = True
                BarButtonItemRenomearJovens.Enabled = True

                ' Exibe o menu
                PopupMenuJovens.ShowPopup(Control.MousePosition)
            Else
                ' Clicou fora de uma célula -> limpa seleção
                view.ClearSelection()
                view.FocusedRowHandle = DevExpress.XtraGrid.GridControl.InvalidRowHandle

                ' Desabilita os botões do PopupMenu
                BarButtonItemExcluirJovens.Enabled = False
                BarButtonItemRenomearJovens.Enabled = False
                PopupMenuJovens.ShowPopup(Control.MousePosition)

            End If
        End If
    End Sub
    Private Sub BarButtonItemRenomearJovens_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemRenomearJovens.ItemClick
        Dim rowHandle = GridViewJovens.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridViewJovens.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim oldName = dataRow("NomeArquivo").ToString
        Dim oldPath = dataRow("Local").ToString

        ' Separa nome e extensão
        Dim baseName = Path.GetFileNameWithoutExtension(oldName)
        Dim extension = Path.GetExtension(oldName)

        ' Só permite editar o nome, não a extensão
        Dim newBaseName = InputBox("Digite o novo nome do arquivo:", "Renomear", baseName).Trim

        If String.IsNullOrWhiteSpace(newBaseName) OrElse newBaseName = baseName Then Return

        Dim newName = newBaseName & extension
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        Try
            If File.Exists(oldPath) Then
                File.Move(oldPath, newPath)

                ' Atualiza DataRow
                dataRow("NomeArquivo") = newName
                dataRow("Local") = newPath

                GridControlJovens.RefreshDataSource
                SalvarConfiguracoesJovens
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao renomear o arquivo: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Excluir arquivos selecionados
    Private Sub BarButtonItemExcluirJovens_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemExcluirJovens.ItemClick
        Dim selectedRows = GridViewJovens.GetSelectedRows
        If selectedRows.Length = 0 Then Return

        Dim msg = If(selectedRows.Length = 1,
                       "Deseja realmente excluir o arquivo selecionado?",
                       $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
        If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim rowsToRemove As New List(Of DataRow)

        For Each rowHandle In selectedRows
            Dim dataRow = GridViewJovens.GetDataRow(rowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath = dataRow("Local").ToString
                Try
                    If File.Exists(filePath) Then File.Delete(filePath)
                Catch ex As Exception
                    MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                rowsToRemove.Add(dataRow)
            End If
        Next

        For Each dr In rowsToRemove
            ArquivosTableJovens.Rows.Remove(dr)
        Next

        ' Recalcula índices
        For i = 0 To ArquivosTableJovens.Rows.Count - 1
            ArquivosTableJovens.Rows(i)("Indice") = i + 1
        Next

        GridControlJovens.RefreshDataSource
        SalvarConfiguracoesJovens
    End Sub

    ' Atualizar, Limpar e Abrir Pasta
    Private Sub BarButtonItemAtualizarJovens_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAtualizarJovens.ItemClick
        CarregarConfiguracoesJovens()
        SalvarConfiguracoesJovens()
    End Sub

    ' Abre a pasta Jovens no Explorer
    Private Sub BarButtonItemAbrirPastaJovens_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAbrirPastaJovens.ItemClick
        Try
            If Directory.Exists(PastaJovens) Then
                Process.Start("explorer.exe", PastaJovens)
            Else
                MessageBox.Show("A pasta não existe: " & PastaJovens,
                            "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir a pasta: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' Salvar configurações do GridJovens em XML
    Private Sub SalvarConfiguracoesJovens()
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Se não tiver raiz, cria
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Remove GridJovens antigo (se existir) para recriar
            Dim gridJovensNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridJovens")
            If gridJovensNode IsNot Nothing Then
                xmlDoc.DocumentElement.RemoveChild(gridJovensNode)
            End If

            ' Cria nó GridJovens
            gridJovensNode = xmlDoc.CreateElement("GridJovens")
            Dim attrUltimaAlteracao As XmlAttribute = xmlDoc.CreateAttribute("UltimaAlteracao")
            attrUltimaAlteracao.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            gridJovensNode.Attributes.Append(attrUltimaAlteracao)

            ' Adiciona arquivos do DataTable
            For Each row As DataRow In ArquivosTableJovens.Rows
                Dim arqNode As XmlElement = xmlDoc.CreateElement("Arquivo")

                Dim attrIndice As XmlAttribute = xmlDoc.CreateAttribute("Indice")
                attrIndice.Value = row("Indice").ToString()
                arqNode.Attributes.Append(attrIndice)

                Dim attrNome As XmlAttribute = xmlDoc.CreateAttribute("NomeArquivo")
                attrNome.Value = row("NomeArquivo").ToString()
                arqNode.Attributes.Append(attrNome)

                ' >>> Novo atributo Local <<<
                Dim attrLocal As XmlAttribute = xmlDoc.CreateAttribute("Local")
                attrLocal.Value = row("Local").ToString()
                arqNode.Attributes.Append(attrLocal)

                gridJovensNode.AppendChild(arqNode)
            Next

            ' Adiciona o nó na raiz
            xmlDoc.DocumentElement.AppendChild(gridJovensNode)

            ' Salva no arquivo
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar configurações: " & ex.Message,
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'Tamanho da Fonte do GridViewProcessos
    Private Sub GridControlJovens_MouseWheel(sender As Object, e As MouseEventArgs) Handles GridControlJovens.MouseWheel
        ' Verifica se a tecla Ctrl está pressionada
        If Not My.Computer.Keyboard.CtrlKeyDown Then Return

        ' Lista de GridViews que terão a fonte ajustada
        Dim grids As DevExpress.XtraGrid.Views.Grid.GridView() = {
        TryCast(GridControlJovens.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlDbv.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlAvt.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlProvaieVede.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
    }

        ' Define limites da fonte
        Dim minFontSize As Single = 5
        Dim maxFontSize As Single = 20

        ' Obtém o tamanho atual da fonte do primeiro GridView válido
        Dim currentFontSize As Single = grids.FirstOrDefault(Function(g) g IsNot Nothing)?.Appearance.Row.Font.Size
        If currentFontSize = 0 Then Return

        ' Ajusta o tamanho da fonte com base na direção do scroll
        If e.Delta > 0 AndAlso currentFontSize < maxFontSize Then
            currentFontSize += 1
        ElseIf e.Delta < 0 AndAlso currentFontSize > minFontSize Then
            currentFontSize -= 1
        End If

        ' Aplica a nova fonte a todos os GridViews válidos
        Dim novaFonte As New Font(grids.First(Function(g) g IsNot Nothing).Appearance.Row.Font.FontFamily, currentFontSize)
        For Each grid In grids
            If grid IsNot Nothing Then
                grid.Appearance.Row.Font = novaFonte
                grid.Appearance.HeaderPanel.Font = novaFonte
                grid.Appearance.FocusedRow.Font = novaFonte
                grid.Appearance.GroupRow.Font = novaFonte
                grid.Appearance.FooterPanel.Font = novaFonte
            End If
        Next

        ' Salva tamanho da fonte (adaptar conforme Grid específico se necessário)
        SalvarTamanhoFonteGridViewProvaieVede(currentFontSize)
    End Sub





#End Region

#Region "INFORMATIVO"
    Private ArquivosTableInformativo As DataTable
    Private dragRowHandleInformativo As Integer = GridControl.InvalidRowHandle
    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableInformativo()
        ArquivosTableInformativo = New DataTable()
        ArquivosTableInformativo.Columns.Add("Icone", GetType(Image))
        ArquivosTableInformativo.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableInformativo.Columns.Add("Local", GetType(String))
    End Sub
    Async Sub CarregarGridInformativo()
        BarEditItemInformativo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never

        Dim proximoSabado As DateTime = ObterProximoSabado()
        Dim nomeArquivo As String = $"Informativo {proximoSabado:dd-MM-yyyy}.mp4"
        Dim caminhoArquivo As String = Path.Combine(PastaInformativo, nomeArquivo)

        ' Verifica se já existe
        If File.Exists(caminhoArquivo) Then
            ' Já existe, não faz nada
            CarregarArquivosInformativo()
            Return
        End If

        ' Não existe, então inicia o download
        Await BaixarInformativo(PastaInformativo)

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewInformativo()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControlInformativo.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        GridControlInformativo.AllowDrop = True
        AddHandler GridControlInformativo.DragEnter, AddressOf PanelInformativo_DragEnter
        AddHandler GridControlInformativo.DragDrop, AddressOf PanelInformativo_DragDrop


    End Sub
    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosInformativo()
        PrepararArquivosTableInformativo()

        Dim arquivos = Directory.GetFiles(PastaInformativo)
        For i As Integer = 0 To arquivos.Length - 1
            Dim arq As String = arquivos(i)
            Dim img As Image = Nothing
            Try
                Dim icone As Icon = Icon.ExtractAssociatedIcon(arq)
                img = icone.ToBitmap()
            Catch
                img = SystemIcons.WinLogo.ToBitmap()
            End Try
            ArquivosTableInformativo.Rows.Add(img, Path.GetFileName(arq), arq)
        Next

        GridControlInformativo.DataSource = ArquivosTableInformativo

    End Sub
    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelInformativo_DragEnter(sender As Object, e As DragEventArgs)
        ' Verifica se são arquivos do Windows Explorer
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub PanelInformativo_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            For Each filePath In files
                Dim destPath As String = Path.Combine(PastaInformativo, Path.GetFileName(filePath))
                File.Copy(filePath, destPath, True)

                ' Extrai ícone
                Dim icone As Icon = Icon.ExtractAssociatedIcon(destPath)
                Dim img As Image = icone.ToBitmap()

                ' Adiciona no final do DataTable
                ArquivosTableInformativo.Rows.Add(img, Path.GetFileName(destPath), destPath)
            Next

            GridControlInformativo.RefreshDataSource()
        End If
    End Sub
    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewInformativo_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewInformativo.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableInformativo.Rows.Remove(dr)
            Next

            ' Atualiza Grid
            GridControlInformativo.RefreshDataSource()
        End If
    End Sub
    ' Abrir arquivo com duplo clique
    Private Sub GridViewInformativo_DoubleClick(sender As Object, e As EventArgs) Handles GridViewInformativo.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub
    ' Abrir arquivo com Enter
    Private Sub GridViewInformativo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewInformativo.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub
    '   Menu de contexto com clique direito
    Private Sub GridViewInformativo_MouseUp(sender As Object, e As MouseEventArgs) Handles GridViewInformativo.MouseUp
        If e.Button = MouseButtons.Right Then
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView)
            Dim hi As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(e.Location)

            If hi.InRowCell Then
                ' Clicou em uma célula válida -> move o foco e seleciona a linha
                view.FocusedRowHandle = hi.RowHandle
                view.ClearSelection()
                view.SelectRow(hi.RowHandle)

                ' Habilita os botões do PopupMenu
                BarButtonItemExcluirInformativo.Enabled = True
                BarButtonItemRenomearInformativo.Enabled = True

                ' Exibe o menu
                PopupMenuInformativo.ShowPopup(Control.MousePosition)
            Else
                ' Clicou fora de uma célula -> limpa seleção
                view.ClearSelection()
                view.FocusedRowHandle = DevExpress.XtraGrid.GridControl.InvalidRowHandle

                ' Desabilita os botões do PopupMenu
                BarButtonItemExcluirInformativo.Enabled = False
                BarButtonItemRenomearInformativo.Enabled = False
                PopupMenuInformativo.ShowPopup(Control.MousePosition)

            End If
        End If
    End Sub
    Private Sub BarButtonItemRenomearInformativo_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemRenomearInformativo.ItemClick
        Dim rowHandle = GridViewInformativo.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridViewInformativo.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim oldName = dataRow("NomeArquivo").ToString
        Dim oldPath = dataRow("Local").ToString

        ' Separa nome e extensão
        Dim baseName = Path.GetFileNameWithoutExtension(oldName)
        Dim extension = Path.GetExtension(oldName)

        ' Só permite editar o nome, não a extensão
        Dim newBaseName = InputBox("Digite o novo nome do arquivo:", "Renomear", baseName).Trim

        If String.IsNullOrWhiteSpace(newBaseName) OrElse newBaseName = baseName Then Return

        Dim newName = newBaseName & extension
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        Try
            If File.Exists(oldPath) Then
                File.Move(oldPath, newPath)

                ' Atualiza DataRow
                dataRow("NomeArquivo") = newName
                dataRow("Local") = newPath

                GridControlInformativo.RefreshDataSource
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao renomear o arquivo: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' Excluir arquivos selecionados
    Private Sub BarButtonItemExcluirInformativo_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemExcluirInformativo.ItemClick
        Dim selectedRows = GridViewInformativo.GetSelectedRows
        If selectedRows.Length = 0 Then Return

        Dim msg = If(selectedRows.Length = 1,
                       "Deseja realmente excluir o arquivo selecionado?",
                       $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
        If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim rowsToRemove As New List(Of DataRow)

        For Each rowHandle In selectedRows
            Dim dataRow = GridViewInformativo.GetDataRow(rowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath = dataRow("Local").ToString
                Try
                    If File.Exists(filePath) Then File.Delete(filePath)
                Catch ex As Exception
                    MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                rowsToRemove.Add(dataRow)
            End If
        Next

        For Each dr In rowsToRemove
            ArquivosTableInformativo.Rows.Remove(dr)
        Next


        GridControlInformativo.RefreshDataSource
    End Sub
    ' Atualizar
    Private Async Sub BarButtonItemAtualizarInformativo_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAtualizarInformativo.ItemClick
        Dim proximoSabado As DateTime = ObterProximoSabado()
        Dim nomeArquivo As String = $"Informativo {proximoSabado:dd-MM-yyyy}.mp4"
        Dim caminhoArquivo As String = Path.Combine(PastaInformativo, nomeArquivo)

        ' Verifica se já existe
        If File.Exists(caminhoArquivo) Then
            ' Já existe, não faz nada
            CarregarArquivosInformativo()
            Return
        End If

        ' Não existe, então inicia o download
        Await BaixarInformativo(PastaInformativo)
    End Sub
    ' Abre a pasta Informativo no Explorer
    Private Sub BarButtonItemAbrirPastaInformativo_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAbrirPastaInformativo.ItemClick
        Try
            If Directory.Exists(PastaInformativo) Then
                Process.Start("explorer.exe", PastaInformativo)
            Else
                MessageBox.Show("A pasta não existe: " & PastaInformativo,
                            "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir a pasta: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'Tamanho da Fonte do GridView
    Private Sub GridControlInformativo_MouseWheel(sender As Object, e As MouseEventArgs)
        ' Verifica se a tecla Ctrl está pressionada
        If Not My.Computer.Keyboard.CtrlKeyDown Then Return

        ' Lista de GridViews que terão a fonte ajustada
        Dim grids As DevExpress.XtraGrid.Views.Grid.GridView() = {
        TryCast(GridControlJovens.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlDbv.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlAvt.MainView, DevExpress.XtraGrid.Views.Grid.GridView),
        TryCast(GridControlInformativo.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
    }

        ' Define limites da fonte
        Dim minFontSize As Single = 5
        Dim maxFontSize As Single = 20

        ' Obtém o tamanho atual da fonte do primeiro GridView válido
        Dim currentFontSize As Single = grids.FirstOrDefault(Function(g) g IsNot Nothing)?.Appearance.Row.Font.Size
        If currentFontSize = 0 Then Return

        ' Ajusta o tamanho da fonte com base na direção do scroll
        If e.Delta > 0 AndAlso currentFontSize < maxFontSize Then
            currentFontSize += 1
        ElseIf e.Delta < 0 AndAlso currentFontSize > minFontSize Then
            currentFontSize -= 1
        End If

        ' Aplica a nova fonte a todos os GridViews válidos
        Dim novaFonte As New Font(grids.First(Function(g) g IsNot Nothing).Appearance.Row.Font.FontFamily, currentFontSize)
        For Each grid In grids
            If grid IsNot Nothing Then
                grid.Appearance.Row.Font = novaFonte
                grid.Appearance.HeaderPanel.Font = novaFonte
                grid.Appearance.FocusedRow.Font = novaFonte
                grid.Appearance.GroupRow.Font = novaFonte
                grid.Appearance.FooterPanel.Font = novaFonte
            End If
        Next

        ' Salva tamanho da fonte (adaptar conforme Grid específico se necessário)
        SalvarTamanhoFonteGridViewInformativo(currentFontSize)
    End Sub
    Public Function ObterTamanhoFonteGridViewInformativo() As Single
        Dim xmlDoc As XmlDocument = CarregarXml()
        Dim tamanhoFonte As Single = 10 ' valor padrão

        Try
            Dim node As XmlNode = xmlDoc.SelectSingleNode("/Configuracoes/Preferencias/FonteGridViewInformativo")
            If node IsNot Nothing Then
                Single.TryParse(node.InnerText, tamanhoFonte)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao obter tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return tamanhoFonte
    End Function
    Public Sub SalvarTamanhoFonteGridViewInformativo(ByVal tamanhoFonte As Single)
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Cria raiz se não existir
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Cria nó <Preferencias> se não existir
            Dim preferenciasNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("Preferencias")
            If preferenciasNode Is Nothing Then
                preferenciasNode = xmlDoc.CreateElement("Preferencias")
                xmlDoc.DocumentElement.AppendChild(preferenciasNode)
            End If

            ' Cria ou atualiza nó <FonteGridViewInformativo>
            Dim fonteNode As XmlNode = preferenciasNode.SelectSingleNode("FonteGridViewInformativo")
            If fonteNode Is Nothing Then
                fonteNode = xmlDoc.CreateElement("FonteGridViewInformativo")
                preferenciasNode.AppendChild(fonteNode)
            End If

            fonteNode.InnerText = tamanhoFonte.ToString(System.Globalization.CultureInfo.InvariantCulture)

            ' Salva com XmlManager
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private ReadOnly httpClient As New HttpClient()
    Private Function ObterTrimestreAtual() As Integer
        Dim mes As Integer = DateTime.Now.Month
        Dim trimestre As Integer = ((mes - 1) \ 3) + 1
        Debug.WriteLine($"Trimestre Atual: {trimestre}")
        Return trimestre
    End Function
    Private Function ObterProximoSabado() As DateTime
        Dim hoje As DateTime = DateTime.Now
        Dim proximoSabado As DateTime

        If hoje.DayOfWeek = DayOfWeek.Saturday Then
            proximoSabado = hoje
        Else
            proximoSabado = hoje.AddDays(6 - CInt(hoje.DayOfWeek))
        End If

        Debug.WriteLine($"Próximo Sábado: {proximoSabado}")
        Return proximoSabado
    End Function
    Private Async Function ObterPaginaHtml(url As String) As Task(Of String)
        Dim html As String = Await httpClient.GetStringAsync(url)
        Debug.WriteLine($"HTML de {url}: {html}...") ' Mostra apenas os primeiros 200 caracteres
        Return html
    End Function
    Private Function ExtrairLinkSabado(html As String, dataSabado As DateTime) As String
        Dim padrao As String = $"{dataSabado:dd/MM/yy}"
        Debug.WriteLine($"Data Padrão: {padrao}")

        ' Ajuste do regex para capturar a URL do link após a data
        Dim padraoCompleto As String = $"{padrao}: <a[^>]+href=""([^""]+)"""
        Dim match As Match = Regex.Match(html, padraoCompleto)

        If match.Success Then
            Debug.WriteLine($"Correspondência encontrada: {match.Value}")
        Else
            Debug.WriteLine("Nenhuma correspondência encontrada.")
        End If

        Dim link As String = If(match.Success, match.Groups(1).Value, String.Empty)
        Debug.WriteLine($"Link Sábado: {link}")
        Return link
    End Function
    Private Function ExtrairLinkZip(html As String) As String
        ' Ajuste do regex para capturar a URL do link que contém "Vídeo em Português 1920x1080p"
        Dim padrao As String = "href=""(https?://[^\s""]+).*?Download: Vídeo em Português 1920x1080p"
        Debug.WriteLine($"Padrão Completo: {html}")

        Dim match As Match = Regex.Match(html, padrao)

        If match.Success Then
            Debug.WriteLine($"Correspondência encontrada: {match.Value}")
        Else
            Debug.WriteLine("Nenhuma correspondência encontrada.")
        End If

        ' Captura a URL do link
        Dim link As String = If(match.Success, match.Groups(1).Value, String.Empty)

        ' Decodifica as entidades HTML, como &#038;
        link = WebUtility.HtmlDecode(link)

        Debug.WriteLine($"Link zip: {link}")
        Return link
    End Function
    Private Async Function ExtrairEExcluirZip(zipFilePath As String, outputFolder As String) As Task
        ' Calcula o próximo sábado
        Dim proximoSabado As DateTime = ObterProximoSabado()
        Dim nomeArquivo As String = $"Informativo {proximoSabado:dd-MM-yyyy}.mp4"

        ' Abre o arquivo ZIP
        Using zip As ZipArchive = ZipFile.OpenRead(zipFilePath)
            For Each entry As ZipArchiveEntry In zip.Entries
                ' Ignora pastas indesejadas
                If entry.FullName.Contains("__MACOSX") Then
                    Debug.WriteLine($"Ignorando pasta: {entry.FullName}")
                    Continue For
                End If

                ' Se for vídeo .mp4, extrai com o nome customizado
                If entry.FullName.EndsWith(".mp4", StringComparison.OrdinalIgnoreCase) Then
                    Dim caminhoDestino As String = Path.Combine(outputFolder, nomeArquivo)

                    ' Garante que o diretório de saída existe
                    Directory.CreateDirectory(Path.GetDirectoryName(caminhoDestino))

                    ' Extrai o arquivo sobrescrevendo com o novo nome
                    Await Task.Run(Sub()
                                       Using entrada As Stream = entry.Open(),
                                         saida As FileStream = File.Create(caminhoDestino)
                                           entrada.CopyTo(saida)
                                       End Using
                                   End Sub)

                    Debug.WriteLine($"Arquivo de vídeo extraído: {caminhoDestino}")
                End If
            Next
        End Using

        ' Exclui o ZIP após a extração
        If File.Exists(zipFilePath) Then
            File.Delete(zipFilePath)
            Debug.WriteLine($"Arquivo ZIP excluído: {zipFilePath}")
        End If
    End Function


    Public Async Function BaixarInformativo(localInformativo As String) As Task
        ' Mostra o item antes de iniciar
        BarEditItemInformativo.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
        BarEditItemInformativo.Caption = "Obtendo informações do próximo sábado..."

        Dim trimestre As Integer = ObterTrimestreAtual()
        Dim proximoSabado As DateTime = ObterProximoSabado()
        Dim urlTrimestre As String = $"https://www.daniellocutor.com.br/informativo-mundial/{trimestre}o-trimestre-{DateTime.Now.Year}/"

        BarEditItemInformativo.Caption = "Carregando página do trimestre..."
        Dim htmlTrimestre As String = Await ObterPaginaHtml(urlTrimestre)
        Dim linkSabado As String = ExtrairLinkSabado(htmlTrimestre, proximoSabado)

        If String.IsNullOrEmpty(linkSabado) Then
            BarEditItemInformativo.Caption = "Link para o próximo sábado não encontrado."
            Await Task.Delay(3000)
            BarEditItemInformativo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
            Return
        End If

        BarEditItemInformativo.Caption = "Carregando página do sábado..."
        Dim htmlSabado As String = Await ObterPaginaHtml(linkSabado)
        Dim linkZip As String = ExtrairLinkZip(htmlSabado)

        If String.IsNullOrEmpty(linkZip) Then
            BarEditItemInformativo.Caption = "Link para o vídeo em português não encontrado."
            Await Task.Delay(3000)
            BarEditItemInformativo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
            Return
        End If

        ' Remove arquivos antigos
        For Each arquivo As String In Directory.GetFiles(localInformativo)
            Try
                File.Delete(arquivo)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Erro ao excluir arquivo", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Next

        Dim nomeArquivoZip As String = Path.Combine(localInformativo, $"Informativo_{proximoSabado:ddMMyy}_alta.zip")
        Dim nomeBase = Path.GetFileNameWithoutExtension(nomeArquivoZip)

        BarEditItemInformativo.Caption = "Verificando se o arquivo já existe..."
        Dim arquivosExistentes = Directory.GetFiles(localInformativo)
        If arquivosExistentes.Any(Function(arq) Path.GetFileNameWithoutExtension(arq).Equals(nomeBase, StringComparison.OrdinalIgnoreCase)) Then
            BarEditItemInformativo.Caption = "Já está atualizado."
            Await Task.Delay(3000)
            BarEditItemInformativo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
            Return
        End If

        BarEditItemInformativo.Caption = "Baixando arquivo ZIP..."

        Using response As HttpResponseMessage = Await httpClient.GetAsync(linkZip, HttpCompletionOption.ResponseHeadersRead)
            response.EnsureSuccessStatusCode()

            Dim totalBytes As Long = If(response.Content.Headers.ContentLength.HasValue, response.Content.Headers.ContentLength.Value, -1)
            Dim totalBytesLidos As Long = 0

            Using inputStream As Stream = Await response.Content.ReadAsStreamAsync(),
          outputStream As New FileStream(nomeArquivoZip, FileMode.Create, FileAccess.Write, FileShare.None, 8192, True)

                Dim buffer(8191) As Byte
                Dim bytesLidos As Integer

                Do
                    bytesLidos = Await inputStream.ReadAsync(buffer, 0, buffer.Length)
                    If bytesLidos = 0 Then Exit Do

                    Await outputStream.WriteAsync(buffer, 0, bytesLidos)
                    totalBytesLidos += bytesLidos

                    ' Atualiza porcentagem se souber o tamanho
                    If totalBytes > 0 Then
                        Dim porcentagem As Integer = CInt((totalBytesLidos * 100L) \ totalBytes)
                        BarEditItemInformativo.Caption = $"Baixando arquivo ZIP... {porcentagem}%"
                    Else
                        BarEditItemInformativo.Caption = $"Baixando arquivo ZIP... {totalBytesLidos \ 1024} KB"
                    End If

                    ' Força refresh visual (às vezes necessário no DevExpress)
                    Application.DoEvents()
                Loop
            End Using
        End Using

        BarEditItemInformativo.Caption = "Download concluído. Extraindo vídeo..."
        Await ExtrairEExcluirZip(nomeArquivoZip, localInformativo)

        BarEditItemInformativo.Caption = "Processo concluído. Vídeo extraído com sucesso."
        Await Task.Delay(3000)

        ' Esconde após terminar
        BarEditItemInformativo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never

        CarregarArquivosInformativo()
    End Function


#End Region

#Region "MÚSICAS"
    Private ArquivosTableMusicas As DataTable
    Private dragRowHandleMusicas As Integer = GridControl.InvalidRowHandle
    ' Cria/zera o DataTable
    Private Sub PrepararArquivosTableMusicas()
        ArquivosTableMusicas = New DataTable()
        ArquivosTableMusicas.Columns.Add("NomeArquivo", GetType(String))
        ArquivosTableMusicas.Columns.Add("Pasta", GetType(String))
        ArquivosTableMusicas.Columns.Add("Local", GetType(String))
    End Sub

    Sub CarregarGridMusicas()
        CarregarConfiguracoesMusicas()

        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewMusicas()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControlMusicas.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte
        End If

        GridControlMusicas.AllowDrop = True
        AddHandler GridControlMusicas.DragEnter, AddressOf PanelMusicas_DragEnter
        AddHandler GridControlMusicas.DragDrop, AddressOf PanelMusicas_DragDrop
    End Sub
    ' Carrega diretamente da pasta
    Private Sub CarregarArquivosMusicas()
        PrepararArquivosTableMusicas()

        ' Extensões permitidas
        Dim extensoes() As String = {"*.mp4", "*.avi", "*.wmv"}
        Dim arquivos As New List(Of String)

        ' Busca recursiva em todas as subpastas para cada extensão
        For Each ext In extensoes
            arquivos.AddRange(Directory.GetFiles(PastaColetaneas, ext, SearchOption.AllDirectories))
        Next

        ' Ordena alfabeticamente pelo nome do arquivo
        arquivos = arquivos.OrderBy(Function(f) Path.GetFileName(f)).ToList()

        ' Carrega os arquivos encontrados
        For Each arq As String In arquivos
            ' Adiciona: nome do arquivo, pasta e caminho completo 
            ArquivosTableMusicas.Rows.Add(Path.GetFileNameWithoutExtension(arq), Path.GetFileName(Path.GetDirectoryName(arq)), arq)
        Next

        GridControlMusicas.DataSource = ArquivosTableMusicas
        SalvarConfiguracoesMusicas()
    End Sub

    ' Carregar configurações do GridMusicas do XML
    Private Sub CarregarConfiguracoesMusicas()
        PrepararArquivosTableMusicas()

        Dim xmlDoc As XmlDocument = CarregarXml()
        If xmlDoc.DocumentElement Is Nothing Then
            CarregarArquivosMusicas()
            Return
        End If

        Try
            Dim gridMusicasNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridMusicas")
            If gridMusicasNode Is Nothing Then
                CarregarArquivosMusicas()
                Return
            End If

            Dim arquivos As XmlNodeList = gridMusicasNode.SelectNodes("Arquivo")
            Dim remover As New List(Of XmlNode)

            ' Carregar arquivos salvos no XML
            For Each arqNode As XmlNode In arquivos
                Dim local As String = arqNode.Attributes("Local")?.Value

                If Not String.IsNullOrEmpty(local) AndAlso File.Exists(local) Then
                    ArquivosTableMusicas.Rows.Add(
                    Path.GetFileNameWithoutExtension(local),
                    Path.GetFileName(Path.GetDirectoryName(local)),
                    local
                )
                Else
                    remover.Add(arqNode)
                End If
            Next

            ' Remover entradas inválidas
            For Each r As XmlNode In remover
                gridMusicasNode.RemoveChild(r)
            Next

            ' Adicionar arquivos novos que não estão no XML
            Dim extensoes() As String = {"*.mp4", "*.avi", "*.wmv"}
            Dim arquivosNaPasta As New List(Of String)
            For Each ext In extensoes
                arquivosNaPasta.AddRange(Directory.GetFiles(PastaColetaneas, ext, SearchOption.AllDirectories))
            Next

            For Each filePath In arquivosNaPasta
                Dim existeNoXml As Boolean = False
                For Each n As XmlNode In gridMusicasNode.SelectNodes("Arquivo")
                    If String.Equals(n.Attributes("Local")?.Value, filePath, StringComparison.OrdinalIgnoreCase) Then
                        existeNoXml = True
                        Exit For
                    End If
                Next

                If Not existeNoXml Then
                    Dim novoNode As XmlElement = xmlDoc.CreateElement("Arquivo")
                    novoNode.SetAttribute("NomeArquivo", Path.GetFileName(filePath))
                    novoNode.SetAttribute("Local", filePath)
                    gridMusicasNode.AppendChild(novoNode)

                    ArquivosTableMusicas.Rows.Add(
            Path.GetFileNameWithoutExtension(filePath),
            Path.GetFileName(Path.GetDirectoryName(filePath)),
            filePath
        )
                End If
            Next


            ' Atualiza data/hora
            Dim attr = gridMusicasNode.Attributes("UltimaAlteracao")
            If attr Is Nothing Then
                attr = xmlDoc.CreateAttribute("UltimaAlteracao")
                gridMusicasNode.Attributes.Append(attr)
            End If
            attr.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            SalvarXml(xmlDoc)

            ' Atualiza GridControl
            GridControlMusicas.DataSource = ArquivosTableMusicas

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Drag & Drop de arquivos do Windows Explorer para o painel
    Private Sub PanelMusicas_DragEnter(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) AndAlso PastaSelecionadaColetaneas <> "" Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            ' Verifica se todos os arquivos já estão na pasta destino
            Dim pastaDestino As String = Path.Combine(PastaColetaneas, PastaSelecionadaColetaneas)
            Dim todosNaMesmaPasta = files.All(Function(f) Path.GetDirectoryName(f).Equals(pastaDestino, StringComparison.OrdinalIgnoreCase))

            If todosNaMesmaPasta Then
                e.Effect = DragDropEffects.None ' Cancela o drop
            Else
                e.Effect = DragDropEffects.Copy
            End If
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub PanelMusicas_DragDrop(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            For Each filePath In files
                Dim pastaDestino As String = Path.Combine(PastaColetaneas, PastaSelecionadaColetaneas)

                ' Se já está na pasta, não faz nada
                If Path.GetDirectoryName(filePath).Equals(pastaDestino, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Try
                    Dim destPath As String = Path.Combine(pastaDestino, Path.GetFileName(filePath))
                    File.Copy(filePath, destPath, True)

                    ArquivosTableMusicas.Rows.Add(
                    Path.GetFileNameWithoutExtension(destPath),
                    Path.GetFileName(Path.GetDirectoryName(destPath)),
                    destPath
                )

                Catch ex As Exception
                    MessageBox.Show("Erro ao adicionar arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Next

            GridControlMusicas.RefreshDataSource()
            SalvarConfiguracoesMusicas()
        End If
    End Sub

    Private Sub GridViewMusicas_MouseDown(sender As Object, e As MouseEventArgs) Handles GridViewMusicas.MouseDown
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(e.Location)
        If hitInfo.InRowCell Then
            dragRowHandleMusicas = hitInfo.RowHandle
        Else
            dragRowHandleMusicas = GridControl.InvalidRowHandle
        End If
    End Sub

    Private Sub GridViewMusicas_MouseMove(sender As Object, e As MouseEventArgs) Handles GridViewMusicas.MouseMove
        If e.Button = MouseButtons.Left AndAlso dragRowHandleMusicas <> GridControl.InvalidRowHandle Then
            Dim view As GridView = CType(sender, GridView)
            Dim filePath As String = TryCast(view.GetRowCellValue(dragRowHandleMusicas, "Local"), String)
            If Not String.IsNullOrEmpty(filePath) AndAlso File.Exists(filePath) Then
                view.GridControl.DoDragDrop(New DataObject(DataFormats.FileDrop, New String() {filePath}), DragDropEffects.Copy)
            End If
        End If
    End Sub

    ' Excluir arquivos selecionados com Delete
    Private Sub GridViewMusicas_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewMusicas.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            ' Confirma se há seleção
            If selectedRows.Length = 0 Then Return

            ' Pergunta de confirmação
            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir o arquivo selecionado?",
                               $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            ' Lista para armazenar os DataRows a serem removidos
            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim filePath As String = dataRow("Local").ToString()
                        Try
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            ' Remove as linhas do DataTable
            For Each dr In rowsToRemove
                ArquivosTableMusicas.Rows.Remove(dr)
            Next

            ' Atualiza Grid
            GridControlMusicas.RefreshDataSource()
            SalvarConfiguracoesMusicas()
        End If
    End Sub

    ' Abrir arquivo com duplo clique
    Private Sub GridViewMusicas_DoubleClick(sender As Object, e As EventArgs) Handles GridViewMusicas.DoubleClick
        Dim view As GridView = CType(sender, GridView)
        Dim hitInfo As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(view.GridControl.PointToClient(Control.MousePosition))
        If hitInfo.InRow Then
            Dim dataRow As DataRow = view.GetDataRow(hitInfo.RowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath As String = dataRow("Local").ToString()
                Try
                    If File.Exists(filePath) Then
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Else
                        MessageBox.Show("O arquivo não existe: " & filePath,
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' Abrir arquivo com Enter
    Private Sub GridViewMusicas_KeyPress(sender As Object, e As KeyPressEventArgs) Handles GridViewMusicas.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            Dim view As GridView = CType(sender, GridView)
            Dim rowHandle As Integer = view.FocusedRowHandle
            If rowHandle >= 0 Then
                Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                If dataRow IsNot Nothing Then
                    Dim filePath As String = dataRow("Local").ToString()
                    Try
                        If File.Exists(filePath) Then
                            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                        Else
                            MessageBox.Show("O arquivo não existe: " & filePath,
                                        "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Erro ao abrir o arquivo: " & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
            e.Handled = True ' Evita o som de "beep"
        End If
    End Sub

    ' Menu de contexto com clique direito
    Private Sub GridViewMusicas_MouseUp(sender As Object, e As MouseEventArgs) Handles GridViewMusicas.MouseUp
        If e.Button = MouseButtons.Right Then
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView)
            Dim hi As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(e.Location)

            If hi.InRowCell Then
                ' Clicou em uma célula válida -> move o foco e seleciona a linha
                view.FocusedRowHandle = hi.RowHandle
                view.ClearSelection()
                view.SelectRow(hi.RowHandle)

                ' Habilita os botões do PopupMenu
                BarButtonItemExcluirMusicas.Enabled = True
                BarButtonItemRenomearMusicas.Enabled = True

                ' Exibe o menu
                PopupMenuMusicas.ShowPopup(Control.MousePosition)
            Else
                ' Clicou fora de uma célula -> limpa seleção
                view.ClearSelection()
                view.FocusedRowHandle = DevExpress.XtraGrid.GridControl.InvalidRowHandle

                ' Desabilita os botões do PopupMenu
                BarButtonItemExcluirMusicas.Enabled = False
                BarButtonItemRenomearMusicas.Enabled = False

                PopupMenuMusicas.ShowPopup(Control.MousePosition)

            End If
        End If
    End Sub
    Private Sub BarButtonItemRenomearMusicas_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemRenomearMusicas.ItemClick
        Dim rowHandle = GridViewMusicas.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridViewMusicas.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim oldName = dataRow("NomeArquivo").ToString
        Dim oldPath = dataRow("Local").ToString

        ' Separa nome e extensão
        Dim baseName = Path.GetFileNameWithoutExtension(oldName)
        Dim extension = Path.GetExtension(oldName)

        ' Só permite editar o nome, não a extensão
        Dim newBaseName = InputBox("Digite o novo nome do arquivo:", "Renomear", baseName).Trim

        If String.IsNullOrWhiteSpace(newBaseName) OrElse newBaseName = baseName Then Return

        Dim newName = newBaseName & extension
        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        Try
            If File.Exists(oldPath) Then
                File.Move(oldPath, newPath)

                ' Atualiza DataRow
                dataRow("NomeArquivo") = newName
                dataRow("Local") = newPath

                GridControlMusicas.RefreshDataSource()
                SalvarConfiguracoesMusicas()
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao renomear o arquivo: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Excluir arquivos selecionados
    Private Sub BarButtonItemExcluirMusicas_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemExcluirMusicas.ItemClick
        Dim selectedRows = GridViewMusicas.GetSelectedRows
        If selectedRows.Length = 0 Then Return

        Dim msg = If(selectedRows.Length = 1,
                       "Deseja realmente excluir o arquivo selecionado?",
                       $"Deseja realmente excluir os {selectedRows.Length} arquivos selecionados?")
        If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim rowsToRemove As New List(Of DataRow)

        For Each rowHandle In selectedRows
            Dim dataRow = GridViewMusicas.GetDataRow(rowHandle)
            If dataRow IsNot Nothing Then
                Dim filePath = dataRow("Local").ToString
                Try
                    If File.Exists(filePath) Then File.Delete(filePath)
                Catch ex As Exception
                    MessageBox.Show("Erro ao deletar o arquivo: " & filePath & vbCrLf & ex.Message,
                                "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                rowsToRemove.Add(dataRow)
            End If
        Next

        For Each dr In rowsToRemove
            ArquivosTableMusicas.Rows.Remove(dr)
        Next

        GridControlMusicas.RefreshDataSource()
        SalvarConfiguracoesMusicas()
    End Sub

    ' Atualizar, Limpar e Abrir Pasta
    Private Sub BarButtonItemAtualizarMusicas_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAtualizarMusicas.ItemClick
        CarregarConfiguracoesMusicas()
        SalvarConfiguracoesMusicas()
    End Sub

    ' Abre a pasta do arquivo selecionado no Explorer
    Private Sub BarButtonItemAbrirPastaMusicas_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAbrirPastaMusicas.ItemClick
        Try
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView =
            TryCast(GridControlMusicas.MainView, DevExpress.XtraGrid.Views.Grid.GridView)

            If view Is Nothing OrElse view.FocusedRowHandle < 0 Then
                MessageBox.Show("Nenhum arquivo selecionado.",
                            "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Pega o caminho completo do arquivo selecionado
            Dim filePath As String = TryCast(view.GetRowCellValue(view.FocusedRowHandle, "Local"), String)

            If String.IsNullOrEmpty(filePath) OrElse Not File.Exists(filePath) Then
                MessageBox.Show("Arquivo não encontrado no disco.",
                            "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Abre o Explorer com o arquivo selecionado
            Process.Start("explorer.exe", "/select,""" & filePath & """")
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir a pasta: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' Salvar configurações do GridMusicas em XML
    Private Sub SalvarConfiguracoesMusicas()
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Se não tiver raiz, cria
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Remove GridMusicas antigo (se existir) para recriar
            Dim gridMusicasNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("GridMusicas")
            If gridMusicasNode IsNot Nothing Then
                xmlDoc.DocumentElement.RemoveChild(gridMusicasNode)
            End If

            ' Cria nó GridMusicas
            gridMusicasNode = xmlDoc.CreateElement("GridMusicas")
            Dim attrUltimaAlteracao As XmlAttribute = xmlDoc.CreateAttribute("UltimaAlteracao")
            attrUltimaAlteracao.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            gridMusicasNode.Attributes.Append(attrUltimaAlteracao)

            ' Adiciona arquivos do DataTable
            For Each row As DataRow In ArquivosTableMusicas.Rows
                Dim arqNode As XmlElement = xmlDoc.CreateElement("Arquivo")

                ' NomeArquivo
                Dim attrNome As XmlAttribute = xmlDoc.CreateAttribute("NomeArquivo")
                attrNome.Value = row("NomeArquivo").ToString()
                arqNode.Attributes.Append(attrNome)

                ' Pasta
                Dim attrPasta As XmlAttribute = xmlDoc.CreateAttribute("Pasta")
                attrPasta.Value = row("Pasta").ToString()
                arqNode.Attributes.Append(attrPasta)

                ' Local (caminho completo)
                Dim attrLocal As XmlAttribute = xmlDoc.CreateAttribute("Local")
                attrLocal.Value = row("Local").ToString()
                arqNode.Attributes.Append(attrLocal)

                gridMusicasNode.AppendChild(arqNode)
            Next

            ' Adiciona o nó na raiz
            xmlDoc.DocumentElement.AppendChild(gridMusicasNode)

            ' Salva no arquivo
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar configurações: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    'Tamanho da Fonte do GridViewProcessos
    Private Sub GridControlMusicas_MouseWheel(sender As Object, e As MouseEventArgs) Handles GridControlMusicas.MouseWheel
        ' Verifica se a tecla Ctrl está pressionada
        If My.Computer.Keyboard.CtrlKeyDown Then
            ' Obtém o GridView principal do GridControl
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControlMusicas.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
            If view IsNot Nothing Then
                ' Obtém o tamanho atual da fonte
                Dim currentFontSize As Single = view.Appearance.Row.Font.Size

                ' Define um limite mínimo e máximo
                Dim minFontSize As Single = 5
                Dim maxFontSize As Single = 20

                ' Aumenta ou diminui o tamanho da fonte com base na direção do scroll
                If e.Delta > 0 AndAlso currentFontSize < maxFontSize Then
                    currentFontSize += 1
                ElseIf e.Delta < 0 AndAlso currentFontSize > minFontSize Then
                    currentFontSize -= 1
                End If

                ' Aplica o novo tamanho de fonte nas aparências relevantes
                Dim novaFonte As New Font(view.Appearance.Row.Font.FontFamily, currentFontSize)

                view.Appearance.Row.Font = novaFonte
                view.Appearance.HeaderPanel.Font = novaFonte
                view.Appearance.FocusedRow.Font = novaFonte
                view.Appearance.GroupRow.Font = novaFonte
                view.Appearance.FooterPanel.Font = novaFonte

                SalvarTamanhoFonteGridViewMusicas(currentFontSize)

            End If
        End If
    End Sub

    ' === Preferências do GridMusicas ===
    Public Function ObterTamanhoFonteGridViewMusicas() As Single
        Dim xmlDoc As XmlDocument = CarregarXml()
        Dim tamanhoFonte As Single = 10 ' valor padrão

        Try
            Dim node As XmlNode = xmlDoc.SelectSingleNode("/Configuracoes/Preferencias/FonteGridViewMusicas")
            If node IsNot Nothing Then
                Single.TryParse(node.InnerText, tamanhoFonte)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao obter tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return tamanhoFonte
    End Function
    Public Sub SalvarTamanhoFonteGridViewMusicas(ByVal tamanhoFonte As Single)
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Cria raiz se não existir
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Cria nó <Preferencias> se não existir
            Dim preferenciasNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("Preferencias")
            If preferenciasNode Is Nothing Then
                preferenciasNode = xmlDoc.CreateElement("Preferencias")
                xmlDoc.DocumentElement.AppendChild(preferenciasNode)
            End If

            ' Cria ou atualiza nó <FonteGridViewMusicas>
            Dim fonteNode As XmlNode = preferenciasNode.SelectSingleNode("FonteGridViewMusicas")
            If fonteNode Is Nothing Then
                fonteNode = xmlDoc.CreateElement("FonteGridViewMusicas")
                preferenciasNode.AppendChild(fonteNode)
            End If

            fonteNode.InnerText = tamanhoFonte.ToString(System.Globalization.CultureInfo.InvariantCulture)

            ' Salva com XmlManager
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "COLETÂNEAS"
    Private ArquivosTableColetaneas As DataTable
    Private dragRowHandleColetaneas As Integer = GridControl.InvalidRowHandle
    Private PastaSelecionadaColetaneas As String = String.Empty
    Sub CarregarGridColetaneas()
        CarregarColetaneas()
        Dim tamanhoFonte As Single = ObterTamanhoFonteGridViewColetaneas()
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControlColetaneas.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
        If view IsNot Nothing Then
            Dim fonte As New Font(view.Appearance.Row.Font.FontFamily, tamanhoFonte)
            view.Appearance.Row.Font = fonte
            view.Appearance.HeaderPanel.Font = fonte
            view.Appearance.FocusedRow.Font = fonte
            view.Appearance.GroupRow.Font = fonte
            view.Appearance.FooterPanel.Font = fonte

            SimpleButtonTodas.PerformClick()

        End If
    End Sub
    Private Sub PrepararArquivosTableColetaneas()
        ArquivosTableColetaneas = New DataTable()
        ArquivosTableColetaneas.Columns.Add("Icone", GetType(Image))
        ArquivosTableColetaneas.Columns.Add("Pasta", GetType(String))
        ArquivosTableColetaneas.Columns.Add("Local", GetType(String))
    End Sub
    Sub CarregarColetaneas()
        ' Limpa e prepara a tabela
        PrepararArquivosTableColetaneas()

        If Not Directory.Exists(PastaColetaneas) Then Exit Sub

        ' Obtém todas as subpastas
        Dim subpastas = Directory.GetDirectories(PastaColetaneas)

        For Each subpasta In subpastas
            Dim nomePasta As String = Path.GetFileName(subpasta)
            Dim img As Image = Nothing

            Try
                ' Procura primeiro PNG e JPG
                Dim arquivosImg = Directory.GetFiles(subpasta, "*.png").Concat(
                              Directory.GetFiles(subpasta, "*.jpg")).ToArray()

                If arquivosImg.Length > 0 Then
                    Using tempImg As Image = Image.FromFile(arquivosImg(0))
                        img = New Bitmap(tempImg, New Size(20, 20)) ' redimensiona
                    End Using
                End If
            Catch ex As Exception
                img = Nothing
            End Try

            ' Adiciona linha na tabela
            ArquivosTableColetaneas.Rows.Add(img, nomePasta, subpasta)
        Next

        ' Vincula ao GridControl
        GridControlColetaneas.DataSource = ArquivosTableColetaneas
        GridViewColetaneas.Columns("Local").Visible = False ' Oculta o caminho se não precisar

    End Sub

    ' Excluir Pastas selecionadas com Delete
    Private Sub GridViewColetaneas_KeyDown(sender As Object, e As KeyEventArgs) Handles GridViewColetaneas.KeyDown
        If e.KeyCode = Keys.Delete Then
            Dim view As GridView = CType(sender, GridView)
            Dim selectedRows As Integer() = view.GetSelectedRows()

            If selectedRows.Length = 0 Then Return

            Dim msg As String = If(selectedRows.Length = 1,
                               "Deseja realmente excluir a pasta selecionada e todo o seu conteúdo?",
                               "Deseja realmente excluir todas as pastas selecionadas e todo o conteúdo delas?")
            If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            Dim rowsToRemove As New List(Of DataRow)

            For Each rowHandle In selectedRows
                If rowHandle >= 0 Then
                    Dim dataRow As DataRow = view.GetDataRow(rowHandle)
                    If dataRow IsNot Nothing Then
                        Dim folderPath As String = dataRow("Local").ToString()
                        Try
                            If Directory.Exists(folderPath) Then
                                Directory.Delete(folderPath, True) ' True = excluir conteúdo
                            End If
                        Catch ex As Exception
                            MessageBox.Show("Erro ao deletar a pasta: " & folderPath & vbCrLf & ex.Message,
                                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                        rowsToRemove.Add(dataRow)
                    End If
                End If
            Next

            For Each dr In rowsToRemove
                ArquivosTableColetaneas.Rows.Remove(dr)
            Next

            GridControlColetaneas.RefreshDataSource()
        End If
    End Sub

    ' Menu de contexto com clique direito
    Private Sub GridViewColetaneas_MouseUp(sender As Object, e As MouseEventArgs) Handles GridViewColetaneas.MouseUp
        If e.Button = MouseButtons.Right Then
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(sender, DevExpress.XtraGrid.Views.Grid.GridView)
            Dim hi As DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo = view.CalcHitInfo(e.Location)

            If hi.InRowCell Then
                view.FocusedRowHandle = hi.RowHandle
                view.ClearSelection()
                view.SelectRow(hi.RowHandle)

                BarButtonItemExcluirColetaneas.Enabled = True
                BarButtonItemRenomearColetaneas.Enabled = True

                PopupMenuColetaneas.ShowPopup(Control.MousePosition)
            Else
                view.ClearSelection()
                view.FocusedRowHandle = DevExpress.XtraGrid.GridControl.InvalidRowHandle

                BarButtonItemExcluirColetaneas.Enabled = False
                BarButtonItemRenomearColetaneas.Enabled = False

                PopupMenuColetaneas.ShowPopup(Control.MousePosition)

            End If
        End If
    End Sub

    ' Renomear pastas
    Private Sub BarButtonItemRenomearColetaneas_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemRenomearColetaneas.ItemClick
        Dim rowHandle = GridViewColetaneas.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridViewColetaneas.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim oldPath = dataRow("Local").ToString()
        Dim oldName = Path.GetFileName(oldPath)
        Dim newName = InputBox("Digite o novo nome da pasta:", "Renomear", oldName).Trim

        If String.IsNullOrWhiteSpace(newName) OrElse newName = oldName Then Return

        Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)

        Try
            If Directory.Exists(oldPath) Then
                Directory.Move(oldPath, newPath)
                dataRow("NomeArquivo") = newName
                dataRow("Local") = newPath
                GridControlColetaneas.RefreshDataSource()
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao renomear a pasta: " & ex.Message,
                    "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Excluir pastas via menu
    Private Sub BarButtonItemExcluirColetaneas_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemExcluirColetaneas.ItemClick
        Dim selectedRows = GridViewColetaneas.GetSelectedRows
        If selectedRows.Length = 0 Then Return

        Dim msg = If(selectedRows.Length = 1,
                 "Deseja realmente excluir a pasta selecionada e todo o seu conteúdo?",
                 $"Deseja realmente excluir as {selectedRows.Length} pastas selecionadas e todo o conteúdo delas?")
        If MessageBox.Show(msg, "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Dim rowsToRemove As New List(Of DataRow)

        For Each rowHandle In selectedRows
            Dim dataRow = GridViewColetaneas.GetDataRow(rowHandle)
            If dataRow IsNot Nothing Then
                Dim folderPath = dataRow("Local").ToString
                Try
                    If Directory.Exists(folderPath) Then Directory.Delete(folderPath, True)
                Catch ex As Exception
                    MessageBox.Show("Erro ao deletar a pasta: " & folderPath & vbCrLf & ex.Message,
                            "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                rowsToRemove.Add(dataRow)
            End If
        Next

        For Each dr In rowsToRemove
            ArquivosTableColetaneas.Rows.Remove(dr)
        Next
        GridControlColetaneas.RefreshDataSource()
    End Sub

    ' Atualizar, Limpar e Abrir Pasta
    Private Sub BarButtonItemAtualizarColetaneas_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAtualizarColetaneas.ItemClick
        CarregarColetaneas()
    End Sub

    ' Abrir pasta selecionada no Explorer
    Private Sub BarButtonItemAbrirPastaColetaneas_ItemClick(sender As Object, e As ItemClickEventArgs) Handles BarButtonItemAbrirPastaColetaneas.ItemClick
        Dim rowHandle = GridViewColetaneas.FocusedRowHandle
        If rowHandle < 0 Then Return

        Dim dataRow = GridViewColetaneas.GetDataRow(rowHandle)
        If dataRow Is Nothing Then Return

        Dim folderPath = dataRow("Local").ToString()
        Try
            If Directory.Exists(folderPath) Then
                Process.Start(New ProcessStartInfo(folderPath) With {.UseShellExecute = True})
            Else
                MessageBox.Show("A pasta não existe: " & folderPath,
                            "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir a pasta: " & ex.Message,
                        "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Tamanho da Fonte do GridViewProcessos
    Private Sub GridControlColetaneas_MouseWheel(sender As Object, e As MouseEventArgs) Handles GridControlColetaneas.MouseWheel
        ' Verifica se a tecla Ctrl está pressionada
        If My.Computer.Keyboard.CtrlKeyDown Then
            ' Obtém o GridView principal do GridControl
            Dim view As DevExpress.XtraGrid.Views.Grid.GridView = TryCast(GridControlColetaneas.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
            If view IsNot Nothing Then
                ' Obtém o tamanho atual da fonte
                Dim currentFontSize As Single = view.Appearance.Row.Font.Size

                ' Define um limite mínimo e máximo
                Dim minFontSize As Single = 5
                Dim maxFontSize As Single = 20

                ' Aumenta ou diminui o tamanho da fonte com base na direção do scroll
                If e.Delta > 0 AndAlso currentFontSize < maxFontSize Then
                    currentFontSize += 1
                ElseIf e.Delta < 0 AndAlso currentFontSize > minFontSize Then
                    currentFontSize -= 1
                End If

                ' Aplica o novo tamanho de fonte nas aparências relevantes
                Dim novaFonte As New Font(view.Appearance.Row.Font.FontFamily, currentFontSize)

                view.Appearance.Row.Font = novaFonte
                view.Appearance.HeaderPanel.Font = novaFonte
                view.Appearance.FocusedRow.Font = novaFonte
                view.Appearance.GroupRow.Font = novaFonte
                view.Appearance.FooterPanel.Font = novaFonte

                SalvarTamanhoFonteGridViewColetaneas(currentFontSize)

            End If
        End If
    End Sub

    ' === Preferências do GridColetaneas ===
    Public Function ObterTamanhoFonteGridViewColetaneas() As Single
        Dim xmlDoc As XmlDocument = CarregarXml()
        Dim tamanhoFonte As Single = 10 ' valor padrão

        Try
            Dim node As XmlNode = xmlDoc.SelectSingleNode("/Configuracoes/Preferencias/FonteGridViewColetaneas")
            If node IsNot Nothing Then
                Single.TryParse(node.InnerText, tamanhoFonte)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao obter tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return tamanhoFonte
    End Function
    Public Sub SalvarTamanhoFonteGridViewColetaneas(ByVal tamanhoFonte As Single)
        Try
            Dim xmlDoc As XmlDocument = CarregarXml()

            ' Cria raiz se não existir
            If xmlDoc.DocumentElement Is Nothing Then
                Dim root As XmlElement = xmlDoc.CreateElement("Configuracoes")
                xmlDoc.AppendChild(root)
            End If

            ' Cria nó <Preferencias> se não existir
            Dim preferenciasNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode("Preferencias")
            If preferenciasNode Is Nothing Then
                preferenciasNode = xmlDoc.CreateElement("Preferencias")
                xmlDoc.DocumentElement.AppendChild(preferenciasNode)
            End If

            ' Cria ou atualiza nó <FonteGridViewColetaneas>
            Dim fonteNode As XmlNode = preferenciasNode.SelectSingleNode("FonteGridViewColetaneas")
            If fonteNode Is Nothing Then
                fonteNode = xmlDoc.CreateElement("FonteGridViewColetaneas")
                preferenciasNode.AppendChild(fonteNode)
            End If

            fonteNode.InnerText = tamanhoFonte.ToString(System.Globalization.CultureInfo.InvariantCulture)

            ' Salva com XmlManager
            SalvarXml(xmlDoc)

        Catch ex As Exception
            MessageBox.Show("Erro ao salvar tamanho da fonte: " & ex.Message,
                        "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub SimpleButtonTodas_Click(sender As Object, e As EventArgs) Handles SimpleButtonTodas.Click
        ArquivosTableMusicas.DefaultView.RowFilter = ""
        PastaSelecionadaColetaneas = ""
    End Sub
    ' Evento disparado ao trocar a linha selecionada no GridViewColetaneas
    Private Sub GridViewColetaneas_FocusedRowChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles GridViewColetaneas.FocusedRowChanged
        If e.FocusedRowHandle >= 0 Then
            ' Obtém o valor da coluna "Pasta" da linha selecionada
            Dim pastaSelecionada As String = GridViewColetaneas.GetRowCellValue(e.FocusedRowHandle, "Pasta").ToString()
            PastaSelecionadaColetaneas = pastaSelecionada
            ' Aplica filtro no DataTable usado pelo GridViewMusicas
            ArquivosTableMusicas.DefaultView.RowFilter = $"Pasta = '{pastaSelecionada.Replace("'", "''")}'"
        Else
            ' Se não tiver linha selecionada, remove o filtro
            ArquivosTableMusicas.DefaultView.RowFilter = ""
            PastaSelecionadaColetaneas = ""
        End If
    End Sub




#End Region

#Region "VOLUME"
    Private WithEvents audioDevice As MMDevice
    Private enumerator As New MMDeviceEnumerator()

    Sub VolumeDoWindows()
        ' Obtém o dispositivo de áudio padrão
        audioDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia)

        ' Configurações iniciais do TrackBarControl
        TrackBarControl1.Properties.Minimum = 0
        TrackBarControl1.Properties.Maximum = 100
        TrackBarControl1.Value = CInt(audioDevice.AudioEndpointVolume.MasterVolumeLevelScalar * 100)

        ' Configurações iniciais do ProgressBarControl
        ProgressBarControl1.Properties.Minimum = 0
        ProgressBarControl1.Properties.Maximum = 100
        ProgressBarControl1.Position = 0

        ' Inicia o Timer para verificar o volume periodicamente
        Timer1.Interval = 100
        Timer1.Start()

        ' Adiciona o evento MouseWheel ao TrackBarControl
        AddHandler TrackBarControl1.MouseWheel, AddressOf TrackBarControl1_MouseWheel
    End Sub

    Private Sub TrackBarControl1_EditValueChanged(sender As Object, e As EventArgs) Handles TrackBarControl1.EditValueChanged
        AjustarVolume()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Verifica e atualiza o volume do sistema no TrackBarControl periodicamente
        Dim currentVolume = CInt(audioDevice.AudioEndpointVolume.MasterVolumeLevelScalar * 100)
        If TrackBarControl1.Value <> currentVolume Then
            TrackBarControl1.Value = currentVolume
        End If

        ' Atualiza a barra de progresso com o nível do pico de áudio
        ProgressBarControl1.Position = CInt(audioDevice.AudioMeterInformation.MasterPeakValue * 100)
    End Sub

    Private Sub TrackBarControl1_MouseWheel(sender As Object, e As MouseEventArgs)
        ' Ajusta o volume de acordo com a rotação do mouse
        Dim stepValue As Integer = 5 ' Passo de volume
        If e.Delta > 0 Then
            TrackBarControl1.Value = Math.Min(TrackBarControl1.Value + stepValue, TrackBarControl1.Properties.Maximum)
        Else
            TrackBarControl1.Value = Math.Max(TrackBarControl1.Value - stepValue, TrackBarControl1.Properties.Minimum)
        End If
        AjustarVolume()
    End Sub

    Private Sub AjustarVolume()
        ' Ajusta o volume do sistema
        If audioDevice IsNot Nothing Then
            Dim volume = TrackBarControl1.Value / 100.0F
            audioDevice.AudioEndpointVolume.MasterVolumeLevelScalar = volume
        End If
    End Sub

    Protected Overrides Sub OnMouseWheel(e As MouseEventArgs)
        MyBase.OnMouseWheel(e)
        If TrackBarControl1.Bounds.Contains(PointToClient(Cursor.Position)) Then
            TrackBarControl1_MouseWheel(TrackBarControl1, e)
        End If
    End Sub


#End Region

#Region "TEMA"
    Public Sub SalvarTemaDevExpress(ByVal tema As String, ByVal paleta As String)
        Dim xmlDoc As XmlDocument = CarregarXml()

        Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("/Configuracoes")
        If rootNode Is Nothing Then
            rootNode = xmlDoc.CreateElement("Configuracoes")
            xmlDoc.AppendChild(rootNode)
        End If

        Dim preferenciasNode As XmlNode = rootNode.SelectSingleNode("Preferencias")
        If preferenciasNode Is Nothing Then
            preferenciasNode = xmlDoc.CreateElement("Preferencias")
            rootNode.AppendChild(preferenciasNode)
        End If

        ' Tema
        Dim temaNode As XmlNode = preferenciasNode.SelectSingleNode("Tema")
        If temaNode Is Nothing Then
            temaNode = xmlDoc.CreateElement("Tema")
            preferenciasNode.AppendChild(temaNode)
        End If
        temaNode.InnerText = tema

        ' Paleta
        Dim paletaNode As XmlNode = preferenciasNode.SelectSingleNode("Paleta")
        If paletaNode Is Nothing Then
            paletaNode = xmlDoc.CreateElement("Paleta")
            preferenciasNode.AppendChild(paletaNode)
        End If
        paletaNode.InnerText = paleta

        SalvarXml(xmlDoc)
    End Sub
    Public Sub CarregarTema()
        Dim xmlDoc As XmlDocument = CarregarXml()
        Dim tema As String = xmlDoc.SelectSingleNode("/Configuracoes/Preferencias/Tema")?.InnerText
        Dim paleta As String = xmlDoc.SelectSingleNode("/Configuracoes/Preferencias/Paleta")?.InnerText

        If Not String.IsNullOrEmpty(tema) Then
            DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(tema, paleta)
        End If
    End Sub

    'salvar tema no xml ao trocar
    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        SalvarTemaDevExpress(DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName, DevExpress.LookAndFeel.UserLookAndFeel.Default.ActiveSvgPaletteName)
    End Sub
#End Region

#Region "XML MANAGER"
    ' Função para carregar o documento XML
    Public Function CarregarXml() As XmlDocument
        Dim xmlDoc As New XmlDocument()
        Try
            If File.Exists(CaminhoConfigXML) Then
                xmlDoc.Load(CaminhoConfigXML)

            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao carregar ou criar o arquivo de configurações: " & ex.Message, "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return xmlDoc
    End Function

    Public Sub SalvarXml(xmlDoc As XmlDocument)
        Try
            xmlDoc.Save(CaminhoConfigXML)
        Catch ex As Exception
            MessageBox.Show("Erro ao salvar o arquivo de configurações: " & ex.Message, "Erro XML", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



#End Region


End Class
