using ClosedXML.Excel;
using Escala;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using System.Data;
using System.Drawing.Printing;
using System.Globalization;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Escala
{

    public partial class Form1 : Form
    {
        private bool _atualizandoSistema = false; // <--- ADICIONE ISSO
                                                  // No Form1_Load ou Construtor

        // =========================================================
        // 1. CONFIGURAÇÕES
        // =========================================================
        private const int MAX_COLS = 36;
        private const int INDEX_FUNCAO = 1;
        private const int INDEX_HORARIO = 2;
        private const int INDEX_ORDEM = 3;
        private const int INDEX_NOME = 4;
        private const int INDEX_DIA_INICIO = 5;
        private DataTable? _tabelaMensal;
        private int _diaSelecionado = 1;
        private int _paginaAtual = 0;
        private JObject? _previsaoCompleta;
        public Form1()
        {
            InitializeComponent();
            this.Load += Form1_Load;

            // Ligações de Eventos
            if (btnImportar != null) btnImportar.Click += button1_Click;
            if (CbSeletorDia != null) CbSeletorDia.SelectedIndexChanged += CbSeletorDia_SelectedIndexChanged;
            if (btnImprimir != null) btnImprimir.Click += btnImprimir_Click;

            // Botão de Gerenciar
            if (BtnGerenciarPostos != null) BtnGerenciarPostos.Click += BtnGerenciarPostos_Click;

            // Configurações do Grid
            if (dataGridView2 != null)
            {
                dataGridView2.DoubleBuffered(true);
                dataGridView2.RowHeadersVisible = false;

                // Eventos
                dataGridView2.CellEnter += DataGridView2_CellEnter;
                dataGridView2.CurrentCellDirtyStateChanged += DataGridView2_CurrentCellDirtyStateChanged;
                // dataGridView2.KeyDown += DataGridView2_KeyDown; // REMOVIDO: Evitar que DELETE apague dados
                dataGridView2.CellValueChanged += DataGridView2_CellValueChanged;
                dataGridView2.CellPainting += DataGridView2_CellPainting;
                dataGridView2.RowPostPaint += DataGridView2_RowPostPaint;

            }
            // Dentro de public Form1()

            dataGridView2.DataError += DataGridView2_DataError;
            if (tabControl1 != null)
            {
                tabControl1.SelectedIndexChanged += TabControl1_SelectedIndexChanged;
            }

            this.FormClosing += Form1_FormClosing;
        }




        private void Form1_Load(object? sender, EventArgs e)
        {
            try
            {
                IniciarBotEmBackground();
                DatabaseService.Initialize();
                _tabelaMensal = DatabaseService.GetMonthlyData();

                if (_tabelaMensal.Rows.Count > 0 && dataGridView1 != null)
                {
                    dataGridView1.DataSource = null;
                    dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = _tabelaMensal;
                    ConfigurarGridMensal();
                    dataGridView1.ReadOnly = true; // Ninguém mexe no visual do Excel
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar dados salvos: " + ex.Message);
            }

            this.WindowState = FormWindowState.Maximized;
            this.tabControl1.SelectedIndex = 1;

            if (CbSeletorDia != null)
            {
                CbSeletorDia.Items.Clear();
                for (int i = 1; i <= 31; i++) CbSeletorDia.Items.Add($"Dia {i}");
                int hoje = DateTime.Now.Day;
                CbSeletorDia.SelectedIndex = (hoje <= 31) ? hoje - 1 : 0;
            }


            // Configurações Visuais Iniciais (DarkGray)
            if (flowLayoutPanel1 != null && dataGridView2 != null)
            {
                flowLayoutPanel1.AutoScroll = true;
                flowLayoutPanel1.FlowDirection = FlowDirection.LeftToRight;
                flowLayoutPanel1.WrapContents = true;
                flowLayoutPanel1.BackColor = System.Drawing.Color.WhiteSmoke;

                dataGridView2.BackgroundColor = System.Drawing.Color.DarkGray;
                dataGridView2.GridColor = System.Drawing.Color.Black;

                _ = AtualizarClimaAutomatico();
            }
        }


        private void IniciarBotEmBackground()
        {
            // Verifica se o Node já está rodando (evita abrir 2 vezes)
            if (System.Diagnostics.Process.GetProcessesByName("node").Length > 0) return;
            if (System.Diagnostics.Process.GetProcessesByName("bot-estacionamento").Length > 0) return;

            // Tenta abrir o arquivo .BAT que criamos (é o jeito mais seguro)
            string caminhoBat = Path.Combine(Application.StartupPath, "iniciar.bat");

            if (File.Exists(caminhoBat))
            {
                var procInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = caminhoBat,
                    WorkingDirectory = Application.StartupPath, // Importante para o Node achar a pasta
                    CreateNoWindow = false, // Deixe false para ver o QR Code se precisar
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal,
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(procInfo);
            }
            else
            {
                // Fallback: Se não achar o .bat, tenta o .exe direto (se você tiver gerado)
                string caminhoExe = Path.Combine(Application.StartupPath, "bot-estacionamento.exe");
                if (File.Exists(caminhoExe))
                {
                    System.Diagnostics.Process.Start(caminhoExe);
                }
            }
        }

        // =========================================================
        // AÇÕES DE BOTÕES
        // =========================================================
        private void TabControl1_SelectedIndexChanged(object? sender, EventArgs e)
        {
            // IMPORTANTE: Troque 'tabPage2' pelo nome exato da sua aba nova
            if (tabControl1.SelectedTab == tabPage2)
            {
                // Só roda se o grid já existir
                if (dataGridView3 != null)
                {
                    GerarRelatorioAcumulado(); // Chama aquela função que te passei antes
                }
            }
        }
        // 1. Método que pinta as cores (Heatmap)
        private void AplicarCorCelula(DataGridViewCell cell, int valor)
        {
            // Se quiser ajustar as faixas de cores, é só mudar os números aqui
            if (valor == 0)
            {
                cell.Style.BackColor = Color.FromArgb(144, 238, 144); // Verde Claro (Zero/Folga)
                cell.Style.ForeColor = Color.Black;
            }
            else if (valor < 5)
            {
                cell.Style.BackColor = Color.FromArgb(255, 255, 153); // Amarelo (Pouco)
                cell.Style.ForeColor = Color.Black;
            }
            else if (valor < 15)
            {
                cell.Style.BackColor = Color.FromArgb(255, 178, 102); // Laranja (Médio)
                cell.Style.ForeColor = Color.Black;
            }
            else
            {
                cell.Style.BackColor = Color.FromArgb(255, 102, 102); // Vermelho (Muito)
                cell.Style.ForeColor = Color.White;
            }
        }
        // 2. Método que cria a linha preta de TOTAIS no final
        private void AdicionarLinhaTotaisVerticais(List<string> pessoas)
        {
            if (dataGridView3.Columns.Count == 0) return;

            int idx = dataGridView3.Rows.Add();
            var row = dataGridView3.Rows[idx];

            // Configura visual da linha de total
            row.Cells["QTH"].Value = "TOTAL GERAL";
            row.DefaultCellStyle.BackColor = Color.Black;
            row.DefaultCellStyle.ForeColor = Color.White;
            row.DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

            int totalGeralzao = 0;

            // Soma coluna por coluna (começa na 1 pq a 0 é o nome do posto)
            for (int c = 1; c <= pessoas.Count; c++)
            {
                int somaColuna = 0;

                // Percorre todas as linhas acima para somar
                for (int r = 0; r < idx; r++)
                {
                    var valorCel = dataGridView3.Rows[r].Cells[c].Value;

                    // Verifica se tem valor numérico (lembrando que cells vazias são "")
                    if (valorCel != null && int.TryParse(valorCel.ToString(), out int v))
                    {
                        somaColuna += v;
                    }
                }

                row.Cells[c].Value = somaColuna;
                totalGeralzao += somaColuna;
            }

            // Define o totalzão da direita
            row.Cells["TOTAL"].Value = totalGeralzao;
        }
        private void GerarRelatorioAcumulado()
        {
            if (dataGridView3 == null) return;

            // 1. Mostra a ampulheta e CONGELA o desenho do Grid (Turbo Mode)
            Cursor.Current = Cursors.WaitCursor;
            dataGridView3.SuspendLayout(); // <--- O SEGREDO ESTÁ AQUI

            try
            {
                // ---------------------------------------------------------
                // PARTE 1: CÁLCULOS (Isso aqui roda em milissegundos)
                // ---------------------------------------------------------
                var contagem = new Dictionary<string, Dictionary<string, int>>();
                var listaPessoas = new HashSet<string>();
                var listaPostos = new HashSet<string>();

                for (int d = 1; d <= 31; d++)
                {
                    var dadosDia = DatabaseService.GetAssignmentsForDay(d);
                    foreach (var kvp in dadosDia)
                    {
                        string pessoa = kvp.Key.ToUpper();
                        listaPessoas.Add(pessoa);

                        foreach (var slot in kvp.Value)
                        {
                            string posto = slot.Value.ToUpper().Trim();
                            if (string.IsNullOrWhiteSpace(posto)) continue;

                            listaPostos.Add(posto);

                            if (!contagem.ContainsKey(posto)) contagem[posto] = new Dictionary<string, int>();
                            if (!contagem[posto].ContainsKey(pessoa)) contagem[posto][pessoa] = 0;
                            contagem[posto][pessoa]++;
                        }
                    }
                }
                // ---------------------------------------------------------
                // PARTE 2: DESENHAR O GRID
                // ---------------------------------------------------------
                dataGridView3.Rows.Clear();
                dataGridView3.Columns.Clear();

                // Configurações visuais
                dataGridView3.RowHeadersVisible = false;
                dataGridView3.AllowUserToAddRows = false;
                dataGridView3.DefaultCellStyle.Font = new Font("Segoe UI", 8); // Letra um pouco menor ajuda

                // Coluna Fixa (Postos)
                dataGridView3.Columns.Add("QTH", "QTH");
                dataGridView3.Columns["QTH"].Width = 120;
                dataGridView3.Columns["QTH"].Frozen = true;
                dataGridView3.Columns["QTH"].DefaultCellStyle.BackColor = Color.LightGray;
                dataGridView3.Columns["QTH"].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

                // Colunas de Pessoas
                var pessoasOrdenadas = listaPessoas.OrderBy(p => p).ToList();
                foreach (var p in pessoasOrdenadas)
                {
                    dataGridView3.Columns.Add(p, p);
                    dataGridView3.Columns[p].Width = 45; // Mais estreito para caber mais gente
                    dataGridView3.Columns[p].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }

                // Coluna TOTAL
                dataGridView3.Columns.Add("TOTAL", "TOTAL");
                dataGridView3.Columns["TOTAL"].Width = 60;
                dataGridView3.Columns["TOTAL"].DefaultCellStyle.BackColor = Color.Black;
                dataGridView3.Columns["TOTAL"].DefaultCellStyle.ForeColor = Color.White;
                dataGridView3.Columns["TOTAL"].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dataGridView3.Columns["TOTAL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // Preencher Linhas
                var postosOrdenados = listaPostos.OrderBy(p => p).ToList();

                foreach (var posto in postosOrdenados)
                {
                    int idx = dataGridView3.Rows.Add();
                    var row = dataGridView3.Rows[idx];
                    row.Cells["QTH"].Value = posto;

                    int totalLinha = 0;

                    for (int c = 1; c <= pessoasOrdenadas.Count; c++)
                    {
                        string nomePessoa = dataGridView3.Columns[c].HeaderText;
                        int valor = 0;

                        if (contagem.ContainsKey(posto) && contagem[posto].ContainsKey(nomePessoa))
                        {
                            valor = contagem[posto][nomePessoa];
                        }

                        row.Cells[c].Value = (valor > 0) ? valor.ToString() : ""; // Deixa vazio se for zero (fica mais limpo)
                        totalLinha += valor;

                        AplicarCorCelula(row.Cells[c], valor);
                    }
                    row.Cells["TOTAL"].Value = totalLinha;
                }

                // Total Geral no rodapé
                AdicionarLinhaTotaisVerticais(pessoasOrdenadas);
            }
            finally
            {
                // 3. DESCONGELA E DESENHA TUDO DE UMA VEZ
                dataGridView3.ResumeLayout();
                Cursor.Current = Cursors.Default;
            }
        }
        private void BtnGerenciarPostos_Click(object? sender, EventArgs e)
        {
            // 1. Abre a janela de configurações
            // O 'using' garante que a janela morra da memória assim que fechar
            using (var form = new FormGerenciar())
            {
                form.ShowDialog(); // O código PAUSA aqui até você fechar a janela

                // 2. Quando a janela fecha, o código continua aqui.
                // As novas configurações (horários) já estão salvas no Banco pelo FormGerenciar.

                // 3. Redesenhamos o dia atual!
                // Como o ProcessarEscalaDoDia lê do Banco (Configurações) e do Excel (_tabelaMensal),
                // ele vai atualizar a tela instantaneamente com as novas regras.
                ProcessarEscalaDoDia();
            }
        }
        private void button1_Click(object? sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel|*.xlsx" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    List<string> nomesAbas = new List<string>();
                    using (var wb = new XLWorkbook(ofd.FileName))
                    {
                        foreach (var ws in wb.Worksheets) nomesAbas.Add(ws.Name);
                    }

                    using (var seletor = new SeletorPlanilha(nomesAbas))
                    {
                        if (seletor.ShowDialog() == DialogResult.OK)
                        {
                            string abaSelecionada = seletor.CbPlanilhas.SelectedItem.ToString();
                            Cursor.Current = Cursors.WaitCursor;

                            _tabelaMensal = LerExcel(ofd.FileName, abaSelecionada);
                            DatabaseService.SaveMonthlyData(_tabelaMensal);

                            if (dataGridView1 != null)
                            {
                                dataGridView1.DataSource = null;
                                dataGridView1.Columns.Clear();
                                dataGridView1.DataSource = _tabelaMensal;
                                ConfigurarGridMensal();
                            }

                            MessageBox.Show($"Importado: {abaSelecionada}");
                            ProcessarEscalaDoDia();
                            Cursor.Current = Cursors.Default;
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show("Erro: " + ex.Message); }
            }
        }
        private void btnImprimir_Click(object? sender, EventArgs e)
        {
            // Salva a aba atual para voltar nela depois
            var abaAnterior = tabControl1.SelectedTab;
            // Força ir para a aba do grid inicialmente (para garantir que a primeira página saia certa)
            tabControl1.SelectedTab = tabPage2;

            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = true;
            pd.DefaultPageSettings.Margins = new Margins(10, 10, 10, 10);
            pd.PrintPage += new PrintPageEventHandler(ImprimirConteudo);

            PrintPreviewDialog ppd = new PrintPreviewDialog();
            ppd.Document = pd;
            ppd.WindowState = FormWindowState.Maximized;

            // O Preview gera o documento. Durante este processo, o ImprimirConteudo vai trocar as abas.
            ppd.ShowDialog();

            // Restaura a aba que o usuário estava
            tabControl1.SelectedTab = abaAnterior;
        }
        private void CbSeletorDia_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (CbSeletorDia.SelectedItem != null)
            {
                string itemStr = CbSeletorDia.SelectedItem.ToString() ?? "";
                if (Regex.Match(itemStr, @"\d+").Success)
                {
                    _diaSelecionado = int.Parse(Regex.Match(itemStr, @"\d+").Value);
                    ProcessarEscalaDoDia();
                    AtualizarClimaParaDia(_diaSelecionado);
                }
            }
        }
        private void btnRecarregarBanco_Click(object? sender, EventArgs e)
        {
            if (MessageBox.Show($"Limpar atribuições do Dia {_diaSelecionado}?", "Confirma", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                DatabaseService.ClearAllAssignments();
                ProcessarEscalaDoDia();
            }
        }
        // =========================================================
        // PROCESSAMENTO PRINCIPAL
        // =========================================================
        private void ProcessarEscalaDoDia()
        {
            if (_tabelaMensal == null || _tabelaMensal.Rows.Count == 0) return;

            // 1. LIGA A TRAVA: Impede que o desenho dispare salvamentos errados
            _atualizandoSistema = true;

            try
            {
                ConfigurarGridEscalaDiaria(); // Limpa a tela do dia

                int indiceColunaDia = INDEX_DIA_INICIO + (_diaSelecionado - 1);
                if (indiceColunaDia >= _tabelaMensal.Columns.Count) return;

                // Listas temporárias para separar os grupos
                var listaSUP = new List<DataRow>();
                var listaOP = new List<DataRow>();
                var listaJV = new List<DataRow>();
                var listaCFTV = new List<DataRow>();
                var listaFolga = new List<DataRow>();
                var listaFerias = new List<DataRow>();

                // 2. LEITURA PURA DO EXCEL (Sem alterar nada nele)
                foreach (DataRow linha in _tabelaMensal.Rows)
                {
                    string nome = linha[INDEX_NOME]?.ToString() ?? "";
                    string horario = linha[INDEX_HORARIO]?.ToString() ?? "";
                    string funcao = (INDEX_FUNCAO < _tabelaMensal.Columns.Count) ? (linha[INDEX_FUNCAO]?.ToString()?.ToUpper() ?? "") : "";
                    string nomeUpper = nome.ToUpper();

                    if (string.IsNullOrWhiteSpace(nome) || nomeUpper.Contains("NOME")) continue;

                    // Filtro de horário válido
                    if (!horario.Contains(":") && !horario.ToUpper().Contains("FOLGUISTA") && !horario.ToUpper().Contains("SIV"))
                        continue;

                    string? statusNoDia = linha[indiceColunaDia]?.ToString()?.ToUpper().Trim();

                    // Separação por categorias
                    if (new[] { "X", "FOLGA", "O" }.Contains(statusNoDia)) { listaFolga.Add(linha); continue; }
                    if (new[] { "F", "FERIAS", "FÉRIAS", "AT", "ATESTADO" }.Contains(statusNoDia)) { listaFerias.Add(linha); continue; }

                    if (funcao.Contains("SUP") || funcao.Contains("LIDER") || nomeUpper.Contains("ISAIAS") || nomeUpper.Contains("ROGÉRIO"))
                        listaSUP.Add(linha);
                    else if (funcao.Contains("JV") || funcao.Contains("APRENDIZ") || nomeUpper.Contains("JOAO"))
                        listaJV.Add(linha);
                    else if (funcao.Contains("CFTV") || nomeUpper.Contains("CFTV"))
                        listaCFTV.Add(linha);
                    else
                        listaOP.Add(linha);
                }

                // 3. CARREGA AS ALTERAÇÕES MANUAIS DESTE DIA ESPECÍFICO
                var assignments = DatabaseService.GetAssignmentsForDay(_diaSelecionado);

                // 4. DESENHA O GRID MESCLANDO EXCEL + BANCO
                // O parâmetro 'assignments' garante que se houver edição manual, ela aparece.
                InserirBloco("OPERADORES", OrdenarPorHorario(listaOP), true, assignments, indiceColunaDia);
                InserirBloco("APRENDIZ", OrdenarPorHorario(listaJV), true, assignments, indiceColunaDia);
                InserirBloco("CFTV", OrdenarPorHorario(listaCFTV), false, assignments, indiceColunaDia);

                InserirListaSimples("FOLGA", listaFolga);
                InserirListaSimples("FÉRIAS|ATESTADOS", listaFerias);

                // 5. VISUALIZAÇÃO
                CalcularTotais();
                PintarHorarios();
                PintarPostos();
                PintarHorarioFunc();

                // 6. LÓGICA AUTOMÁTICA (Agora passando o banco para ela respeitar)
                // Precisamos atualizar seus métodos de lógica para receber 'assignments'
                AplicarLogicaFolguistaCFTV(listaCFTV, assignments);
                AplicarLogicaIntermediarioCFTV(listaCFTV, assignments);

                if (flowLayoutPanel1 != null) AtualizarItinerarios();
            }
            finally
            {
                // 7. DESLIGA A TRAVA: Agora o usuário pode editar
                _atualizandoSistema = false;
            }
        }
        // Método Atualizado: Recebe assignments
        private void AplicarLogicaFolguistaCFTV(List<DataRow> listaCFTV, Dictionary<string, Dictionary<string, string>> assignments)
        {
            string nomeDoFolguista = "";
            // Identifica o folguista
            foreach (DataRow dados in listaCFTV)
            {
                string horario = dados[INDEX_HORARIO]?.ToString()?.ToUpper() ?? "";
                if (horario.Contains("FOLGUISTA") || horario.Contains("SIV") || horario.Contains("COBERTURA"))
                {
                    nomeDoFolguista = dados[INDEX_NOME]?.ToString()?.ToUpper() ?? "";
                    break;
                }
            }

            if (string.IsNullOrEmpty(nomeDoFolguista)) return;

            // Respeita edição manual do banco
            if (assignments.ContainsKey(nomeDoFolguista) &&
                assignments[nomeDoFolguista].ContainsKey("HORARIO") &&
                !string.IsNullOrWhiteSpace(assignments[nomeDoFolguista]["HORARIO"]))
            {
                return;
            }

            // Acha a linha no grid
            DataGridViewRow rowFolguista = null;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                string nomeNoGrid = row.Cells["Nome"].Value?.ToString()?.ToUpper() ?? "";
                if (nomeNoGrid == nomeDoFolguista || (nomeNoGrid.Contains(nomeDoFolguista) && nomeDoFolguista.Length > 3))
                {
                    rowFolguista = row;
                    break;
                }
            }

            if (rowFolguista == null) return;

            // --- REMOVI O LOOP DE RESET/BLOQUEIO DAQUI ---

            // Calcula o horário
            int colDia = INDEX_DIA_INICIO + (_diaSelecionado - 1);
            string horarioParaAssumir = DatabaseService.GetHorarioPadraoFolguista();
            bool achouFaltaPrioritaria = false;
            bool achouAlgumaFalta = false;

            foreach (DataRow pessoa in listaCFTV)
            {
                string nomePessoa = pessoa[INDEX_NOME].ToString().ToUpper();
                if (nomePessoa == nomeDoFolguista) continue;

                string horarioPessoa = pessoa[INDEX_HORARIO]?.ToString()?.ToUpper() ?? "";
                string status = pessoa[colDia].ToString().ToUpper().Trim();
                bool estaDeFolga = (status == "FOLGA" || status == "X" || status == "F" || status == "O" || status.Contains("FÉRIAS") || status.Contains("ATESTADO"));

                if (estaDeFolga)
                {
                    if (horarioPessoa.Contains("16:40") || horarioPessoa.Contains("00:40"))
                    {
                        horarioParaAssumir = pessoa[INDEX_HORARIO].ToString();
                        break;
                    }
                    if (!achouAlgumaFalta)
                    {
                        horarioParaAssumir = pessoa[INDEX_HORARIO].ToString();
                        achouAlgumaFalta = true;
                    }
                }
            }

            // Aplica visualmente
            rowFolguista.Cells["HORARIO"].Value = horarioParaAssumir;
            var partes = horarioParaAssumir.Split(new[] { 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries);

            if (partes.Length == 2)
            {
                // Chama a pintura simples
                PintarIntervaloBranco(rowFolguista, partes[0].Trim(), partes[1].Trim());
            }
        }

        // Método Atualizado: Recebe assignments
        private void AplicarLogicaIntermediarioCFTV(List<DataRow> listaCFTV, Dictionary<string, Dictionary<string, string>> assignments)
        {
            string horarioPadrao = DatabaseService.GetHorarioPadraoIntermediario();

            foreach (DataRow dados in listaCFTV)
            {
                string horarioExcel = dados[INDEX_HORARIO]?.ToString() ?? "";
                string nome = dados[INDEX_NOME]?.ToString()?.ToUpper() ?? "";

                if (horarioExcel.Contains("12:40") && horarioExcel.Contains("21:00"))
                {
                    // Respeita edição manual
                    if (assignments.ContainsKey(nome) &&
                        assignments[nome].ContainsKey("HORARIO") &&
                        !string.IsNullOrWhiteSpace(assignments[nome]["HORARIO"]))
                    {
                        break;
                    }

                    DataGridViewRow rowInter = null;
                    foreach (DataGridViewRow row in dataGridView2.Rows)
                    {
                        string nomeNoGrid = row.Cells["Nome"].Value?.ToString()?.ToUpper() ?? "";
                        if (nomeNoGrid == nome || (nomeNoGrid.Contains(nome) && nome.Length > 3))
                        {
                            rowInter = row;
                            break;
                        }
                    }

                    if (rowInter != null)
                    {
                        // --- REMOVI O LOOP DE RESET/BLOQUEIO DAQUI ---

                        rowInter.Cells["HORARIO"].Value = horarioPadrao;
                        var partes = horarioPadrao.Split(new[] { 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries);

                        if (partes.Length == 2)
                        {
                            PintarIntervaloBranco(rowInter, partes[0].Trim(), partes[1].Trim());
                        }
                    }
                    break;
                }
            }
        }
        private void PintarIntervaloBranco(DataGridViewRow row, string horaInicio, string horaFim)
        {
            if (TimeSpan.TryParse(horaInicio, out TimeSpan ini) && TimeSpan.TryParse(horaFim, out TimeSpan fim))
            {
                TimeSpan fimAj = (fim < ini) ? fim.Add(TimeSpan.FromHours(24)) : fim;

                for (int c = 3; c < dataGridView2.Columns.Count; c++)
                {
                    string header = dataGridView2.Columns[c].HeaderText;
                    if (TryParseHorario(header, out TimeSpan hIni, out TimeSpan hFim))
                    {
                        TimeSpan hFimAj = (hFim < hIni) ? hFim.Add(TimeSpan.FromHours(24)) : hFim;

                        // Se estiver dentro do horário de trabalho
                        if (ini <= hIni && fimAj >= hFimAj)
                        {
                            row.Cells[c].Style.BackColor = System.Drawing.Color.White;
                            row.Cells[c].Style.ForeColor = System.Drawing.Color.Black;

                            // NÃO mexemos no ReadOnly. Se o grid permite edição, continua permitindo.
                            // Se precisar FORÇAR que seja editável, descomente a linha abaixo:
                            // row.Cells[c].ReadOnly = false; 
                        }
                        else
                        {
                            // LIMPEZA: Se não estiver no horário, garante que fique cinza (apagado)
                            // Isso corrige o bug de "duplicar" horarios quando muda a logica
                            if (row.Cells[c].Style.BackColor != System.Drawing.Color.DarkGray)
                            {
                                row.Cells[c].Style.BackColor = System.Drawing.Color.DarkGray;
                                row.Cells[c].Style.ForeColor = System.Drawing.Color.White;
                            }
                        }
                    }
                }
            }
        }
        private void ConfigurarGridEscalaDiaria()
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();

            // Visual DarkGray + Remove Linha Extra
            dataGridView2.BackgroundColor = System.Drawing.Color.DarkGray;
            dataGridView2.GridColor = System.Drawing.Color.Black;
            dataGridView2.AllowUserToAddRows = false;

            var estiloPadrao = new DataGridViewCellStyle();
            estiloPadrao.BackColor = System.Drawing.Color.DarkGray;
            estiloPadrao.ForeColor = System.Drawing.Color.White;
            estiloPadrao.SelectionBackColor = System.Drawing.Color.DimGray;
            estiloPadrao.SelectionForeColor = System.Drawing.Color.White;
            estiloPadrao.Font = new System.Drawing.Font("Bahnschrift Condensed", 12, FontStyle.Bold);

            dataGridView2.DefaultCellStyle = estiloPadrao;

            // Colunas Fixas
            dataGridView2.Columns.Add("ORDEM", "Nº");
            dataGridView2.Columns.Add("HORARIO", "HORÁRIO");
            dataGridView2.Columns.Add("Nome", "NOME");

            List<string> horarios = DatabaseService.GetHorariosConfigurados();
            List<string> postos = DatabaseService.GetPostosConfigurados();

            foreach (var h in horarios)
            {
                var col = new DataGridViewComboBoxColumn
                {
                    HeaderText = h,
                    DataSource = postos,
                    FlatStyle = FlatStyle.Flat,
                    DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                    Width = 65
                };

                col.DefaultCellStyle.BackColor = System.Drawing.Color.DarkGray;
                col.DefaultCellStyle.ForeColor = System.Drawing.Color.White;
                col.DefaultCellStyle.Font = new System.Drawing.Font("Bahnschrift Condensed", 12, FontStyle.Bold);

                dataGridView2.Columns.Add(col);
            }

            dataGridView2.Columns["ORDEM"].Frozen = true;
            dataGridView2.Columns["HORARIO"].Frozen = true;
            dataGridView2.Columns["Nome"].Frozen = true;

            dataGridView2.Columns["ORDEM"].Width = 40;
            dataGridView2.Columns["HORARIO"].Width = 90;
            dataGridView2.Columns["Nome"].Width = 110;

            var estiloFixo = new DataGridViewCellStyle
            {
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.Black,
                Font = new System.Drawing.Font("Bahnschrift Condensed", 12, FontStyle.Bold)
            };
            dataGridView2.Columns["ORDEM"].DefaultCellStyle = estiloFixo;
            dataGridView2.Columns["HORARIO"].DefaultCellStyle = estiloFixo;
            dataGridView2.Columns["Nome"].DefaultCellStyle = estiloFixo;
        }
        // =========================================================
        // LÓGICA DE PINTURA E CORES
        // =========================================================
        private void PintarHorarios()
        {
            for (int r = 0; r < dataGridView2.Rows.Count; r++)
            {
                var row = dataGridView2.Rows[r];
                string nome = row.Cells["Nome"].Value?.ToString() ?? "";
                if (nome.Contains("OPERADORES") || nome.Contains("CFTV") || nome.Contains("APRENDIZ")) continue;

                string horarioFunc = row.Cells["HORARIO"].Value?.ToString() ?? "";
                if (!TryParseHorario(horarioFunc, out TimeSpan ini, out TimeSpan fim)) continue;

                TimeSpan fimAj = (fim < ini) ? fim.Add(TimeSpan.FromHours(24)) : fim;

                for (int c = 3; c < dataGridView2.Columns.Count; c++)
                {
                    if (TryParseHorario(dataGridView2.Columns[c].HeaderText, out TimeSpan hIni, out TimeSpan hFim))
                    {
                        TimeSpan hFimAj = (hFim < hIni) ? hFim.Add(TimeSpan.FromHours(24)) : hFim;
                        bool trabalha = (ini <= hIni) && (fimAj >= hFimAj);

                        // --- CORREÇÃO AQUI: DEFINIÇÃO DO VAR COR ---
                        var cor = row.Cells[c].Style.BackColor;

                        if (trabalha)
                        {
                            if (cor == System.Drawing.Color.DarkGray ||
                                cor == System.Drawing.Color.Black ||
                                cor == System.Drawing.Color.LightGray ||
                                cor.IsEmpty)
                            {
                                row.Cells[c].Style.BackColor = System.Drawing.Color.White;
                                row.Cells[c].Style.ForeColor = System.Drawing.Color.Black;
                                row.Cells[c].ReadOnly = false;
                            }
                        }
                        else
                        {
                            if (cor != System.Drawing.Color.DarkGray)
                            {
                                row.Cells[c].Style.BackColor = System.Drawing.Color.DarkGray;
                                row.Cells[c].Style.ForeColor = System.Drawing.Color.White;
                                row.Cells[c].ReadOnly = false;
                            }
                        }
                    }
                }
            }
        }
        private void PintarPostos()
        {
            for (int r = 0; r < dataGridView2.Rows.Count; r++)
            {
                var row = dataGridView2.Rows[r];
                string nome = row.Cells["Nome"].Value?.ToString() ?? "";
                if (nome.Contains("OPERADORES") || nome.Contains("CFTV") || nome.Contains("APRENDIZ")) continue;

                for (int c = 3; c < dataGridView2.Columns.Count; c++)
                {
                    var cell = dataGridView2.Rows[r].Cells[c];
                    string valor = cell.Value?.ToString()?.ToUpper().Trim() ?? "";

                    if (string.IsNullOrEmpty(valor)) continue;

                    cell.Style.ForeColor = System.Drawing.Color.Black;

                    switch (valor)
                    {
                        case "VALET": cell.Style.BackColor = System.Drawing.Color.FromArgb(255, 100, 100); break;
                        case "CAIXA": cell.Style.BackColor = System.Drawing.Color.FromArgb(255, 150, 150); break;
                        case "QRF":
                            cell.Style.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                            cell.Style.ForeColor = System.Drawing.Color.White;
                            break;
                        case "CIRC.": case "CIRC": cell.Style.BackColor = System.Drawing.Color.FromArgb(153, 204, 255); break;
                        case "REP|CIRC":
                            cell.Style.BackColor = System.Drawing.Color.FromArgb(0, 153, 0);
                            cell.Style.ForeColor = System.Drawing.Color.White;
                            break;
                        case "ECHO 21": cell.Style.BackColor = System.Drawing.Color.FromArgb(102, 204, 0); break;
                        case "CFTV":
                            cell.Style.BackColor = System.Drawing.Color.FromArgb(0, 51, 153);
                            cell.Style.ForeColor = System.Drawing.Color.White;
                            break;
                        case "TREIN": cell.Style.BackColor = System.Drawing.Color.FromArgb(255, 255, 153); break;
                        case "APOIO": cell.Style.BackColor = System.Drawing.Color.LightGray; break;
                        default: cell.Style.BackColor = System.Drawing.Color.White; break;
                    }
                }
            }
        }
        private void PintarHorarioFunc()
        {
            for (int r = 0; r < dataGridView2.Rows.Count; r++)
            {
                var row = dataGridView2.Rows[r];
                if (row.Tag?.ToString() == "IGNORAR" || !row.Visible) continue;

                var cell = row.Cells["HORARIO"];
                string texto = cell.Value?.ToString() ?? "";

                if (texto.Contains("12:40") || texto.Contains("12:41"))
                {
                    cell.Style.BackColor = System.Drawing.Color.DarkRed;
                    cell.Style.ForeColor = System.Drawing.Color.WhiteSmoke;
                }
                else if (texto.Contains("09:40") || texto.Contains("09:41"))
                {
                    cell.Style.BackColor = System.Drawing.Color.Green;
                    cell.Style.ForeColor = System.Drawing.Color.WhiteSmoke;
                }
                else if (texto.Contains("14:40") || texto.Contains("14:41"))
                {
                    cell.Style.BackColor = System.Drawing.Color.Blue;
                    cell.Style.ForeColor = System.Drawing.Color.White;
                }
            }
        }
        // =========================================================
        // MÉTODOS AUXILIARES
        // =========================================================
        private void InserirBloco(string titulo, List<DataRow> lista, bool gerarCartao, Dictionary<string, Dictionary<string, string>> assignments = null, int colIndex = -1)
        {
            if (lista.Count == 0) return;

            foreach (var item in lista)
            {
                int idx = dataGridView2.Rows.Add();
                var r = dataGridView2.Rows[idx];
                r.Cells["ORDEM"].Value = item[INDEX_ORDEM];
                string nome = item[INDEX_NOME]?.ToString() ?? "";
                r.Cells["HORARIO"].Value = item[INDEX_HORARIO]?.ToString();
                r.Cells["Nome"].Value = nome;
                r.Tag = gerarCartao ? "GERAR" : "IGNORAR";

                string postoExcel = (colIndex >= 0 && item.Table.Columns.Count > colIndex)
                                      ? item[colIndex]?.ToString()?.ToUpper().Trim() ?? ""
                                      : "";

                // Filtra valores que não são postos (ex: FOLGA, FÉRIAS, etc - embora as listas já separem, segurança extra)
                if (new[] { "X", "FOLGA", "F", "FERIAS", "FÉRIAS", "AT", "ATESTADO", "O" }.Contains(postoExcel))
                    postoExcel = "";

                // Verifica se tem algo no banco
                Dictionary<string, string> userPosts = null;
                if (assignments != null && assignments.ContainsKey(nome)) userPosts = assignments[nome];

                // Prioridade 1: Banco de Dados
                if (userPosts != null)
                {
                    // --- CORREÇÃO: Restaurar também o HORARIO se houver salvo ---
                    // O banco salva usando o HeaderText, que é "HORÁRIO" (com acento)
                    if (userPosts.ContainsKey("HORÁRIO"))
                    {
                        r.Cells["HORARIO"].Value = userPosts["HORÁRIO"];
                    }
                    else if (userPosts.ContainsKey("HORARIO")) // Fallback
                    {
                        r.Cells["HORARIO"].Value = userPosts["HORARIO"];
                    }

                    for (int c = 3; c < dataGridView2.Columns.Count; c++)
                    {
                        string slot = dataGridView2.Columns[c].HeaderText;
                        if (userPosts.ContainsKey(slot))
                        {
                            r.Cells[c].Value = userPosts[slot];
                        }
                        else if (!string.IsNullOrEmpty(postoExcel))
                        {
                            // Se não tem no banco para esse horário, usa o do Excel
                            r.Cells[c].Value = postoExcel;
                        }
                    }
                }
                else
                {
                    // Prioridade 2: Excel (se tiver valor válido)
                    if (!string.IsNullOrEmpty(postoExcel))
                    {
                        for (int c = 3; c < dataGridView2.Columns.Count; c++)
                        {
                            r.Cells[c].Value = postoExcel;
                        }
                    }
                }
            }

            int idxT = dataGridView2.Rows.Add();
            var rowT = dataGridView2.Rows[idxT];
            rowT.Tag = "IGNORAR";
            rowT.Cells["Nome"].Value = $"{titulo} ({lista.Count})";
            rowT.DefaultCellStyle.BackColor = System.Drawing.Color.Yellow;
            rowT.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            rowT.DefaultCellStyle.Font = new System.Drawing.Font("Bahnschrift Condensed", 12, FontStyle.Bold);
        }
        private void InserirListaSimples(string titulo, List<DataRow> lista)
        {
            if (lista.Count == 0) return;
            var nomes = lista.Select(l => l[INDEX_NOME]?.ToString() ?? "").Where(n => !string.IsNullOrWhiteSpace(n));
            string texto = $"{titulo}: {string.Join(", ", nomes)}";

            int idx = dataGridView2.Rows.Add();
            var row = dataGridView2.Rows[idx];
            row.Cells["Nome"].Value = texto;
            row.Tag = "MERGE";
            row.DefaultCellStyle.BackColor = System.Drawing.Color.Yellow;
            row.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            row.DefaultCellStyle.Font = new System.Drawing.Font("Bahnschrift Condensed", 12, FontStyle.Bold);
            row.Height = 50;
        }
        // REMOVIDO: O método DataGridView2_KeyDown apagava dados ao apertar DELETE.
        // private void DataGridView2_KeyDown(object? sender, KeyEventArgs e) { ... }
        private void DataGridView2_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
        {
            // SEGUNDA TRAVA: Se for o sistema calculando, não salva no banco!
            if (_atualizandoSistema) return;

            if (e.RowIndex < 0 || e.ColumnIndex < 3) return;
            var row = dataGridView2.Rows[e.RowIndex];
            if (row.Tag?.ToString() == "IGNORAR") return;

            string nome = row.Cells["Nome"].Value?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(nome)) return;

            string timeSlot = dataGridView2.Columns[e.ColumnIndex].HeaderText;
            string valor = row.Cells[e.ColumnIndex].Value?.ToString() ?? "";

            DatabaseService.SaveAssignment(_diaSelecionado, nome, timeSlot, valor);
        }
        // =========================================================
        // PINTURA CUSTOMIZADA (MERGE VISUAL)
        // =========================================================
        private void DataGridView2_CellPainting(object? sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dataGridView2.Rows[e.RowIndex].Tag?.ToString() != "MERGE") return;

            e.PaintBackground(e.CellBounds, true);
            e.Handled = true;

            using (Brush br = new SolidBrush(e.CellStyle.BackColor))
                e.Graphics.FillRectangle(br, e.CellBounds);

            e.Graphics.DrawLine(Pens.Black, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);
        }
        private void DataGridView2_RowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
            var row = dataGridView2.Rows[e.RowIndex];
            if (row.Tag?.ToString() != "MERGE") return;

            string texto = row.Cells["Nome"].Value?.ToString() ?? "";

            var rect3 = dataGridView2.GetCellDisplayRectangle(3, e.RowIndex, true);
            int xStart = (rect3.Width > 0) ? rect3.X : e.RowBounds.Left;
            int width = e.RowBounds.Right - xStart;

            Rectangle r = new Rectangle(xStart, e.RowBounds.Top, width, e.RowBounds.Height);
            TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak;

            r.X += 5; r.Width -= 10;
            TextRenderer.DrawText(e.Graphics, texto, e.InheritedRowStyle.Font, r, e.InheritedRowStyle.ForeColor, flags);
        }
        // =========================================================
        // IMPRESSÃO E ITINERÁRIOS
        // =========================================================
        private void ImprimirConteudo(object? sender, PrintPageEventArgs e)
        {
            if (e.Graphics == null) return;
            float y = e.MarginBounds.Top;
            float x = e.MarginBounds.Left;
            float w = e.MarginBounds.Width;
            var fonteT = new System.Drawing.Font("Arial", 16, FontStyle.Bold);
            var fonteC = new System.Drawing.Font("Arial", 10, FontStyle.Regular);

            if (_paginaAtual == 0)
            {
                // GARANTIA: Ativa a aba do Grid e força o desenho
                tabControl1.SelectedTab = tabPage2;
                Application.DoEvents();

                e.Graphics.DrawString($"Escala do Dia {_diaSelecionado}", fonteT, Brushes.Black, x, y);
                y += 30;
                e.Graphics.DrawString(lblClima.Text, fonteC, Brushes.DarkSlateGray, x, y);
                y += 30;

                int hOriginal = dataGridView2.Height;
                dataGridView2.Height = dataGridView2.RowCount * dataGridView2.RowTemplate.Height + dataGridView2.ColumnHeadersHeight;
                Bitmap bmp = new Bitmap(dataGridView2.Width, dataGridView2.Height);
                dataGridView2.DrawToBitmap(bmp, new Rectangle(0, 0, dataGridView2.Width, dataGridView2.Height));
                dataGridView2.Height = hOriginal;

                float ratio = (float)bmp.Width / (float)bmp.Height;
                float hPrint = w / ratio;
                if (hPrint > e.MarginBounds.Height - 100) hPrint = e.MarginBounds.Height - 100;

                e.Graphics.DrawImage(bmp, x, y, w, hPrint);
                e.HasMorePages = true;
                _paginaAtual++;
            }
            else
            {
                // GARANTIA: Ativa a aba de Itinerários e força o desenho
                tabControl1.SelectedTab = tabPage3;
                Application.DoEvents();

                e.Graphics.DrawString("Itinerários", fonteT, Brushes.Black, x, y);
                y += 40;
                int hP = 0;
                foreach (System.Windows.Forms.Control c in flowLayoutPanel1.Controls) hP = Math.Max(hP, c.Bottom);
                hP += 20;
                if (hP < 50) hP = 100;

                Bitmap bmp = new Bitmap(flowLayoutPanel1.Width, hP);
                flowLayoutPanel1.DrawToBitmap(bmp, new Rectangle(0, 0, flowLayoutPanel1.Width, hP));

                float ratio = (float)bmp.Width / (float)bmp.Height;
                float hPrint = w / ratio;
                if (hPrint > e.MarginBounds.Height - 100) hPrint = e.MarginBounds.Height - 100;

                e.Graphics.DrawImage(bmp, x, y, w, hPrint);
                e.HasMorePages = false;
                _paginaAtual = 0;
            }
        }
        // =========================================================
        // OUTROS MÉTODOS (Helpers, Clima, etc)
        // =========================================================
        private DataTable LerExcel(string caminho, string nomeAba)
        {
            var dt = new DataTable();
            for (int i = 1; i <= MAX_COLS; i++) dt.Columns.Add($"C{i}");
            using (var wb = new XLWorkbook(caminho))
            {
                var ws = wb.Worksheet(nomeAba);
                foreach (var r in ws.RowsUsed())
                {
                    var n = dt.NewRow();
                    for (int c = 1; c <= MAX_COLS; c++)
                        n[c - 1] = r.Cell(c).GetValue<string>()?.ToUpper() ?? "";
                    dt.Rows.Add(n);
                }
            }
            return dt;
        }
        private async Task AtualizarClimaAutomatico()
        {
            try
            {
                using (var c = new HttpClient())
                {
                    var json = JObject.Parse(await c.GetStringAsync("https://api.hgbrasil.com/weather?woeid=455822&key=development"));
                    _previsaoCompleta = json;
                    AtualizarClimaParaDia(_diaSelecionado);
                }
            }
            catch { lblClima.Text = "Clima offline"; }
        }
        private void AtualizarClimaParaDia(int dia)
        {
            if (_previsaoCompleta == null) { lblClima.Text = "Carregando..."; return; }
            try
            {
                var res = _previsaoCompleta["results"];
                DateTime hoje = DateTime.Now;
                DateTime alvo = new DateTime(hoje.Year, hoje.Month, dia);
                if (alvo < hoje.Date) alvo = alvo.AddMonths(1);
                int diff = (alvo - hoje.Date).Days;

                if (diff == 0)
                {
                    string condicao = res["description"]?.ToString() ?? "";
                    string icon = ObterIconeClima(condicao);
                    lblClima.Text = $"Hoje: {res["temp"]}°C - {condicao} {icon}";
                }
                else if (diff < 10)
                {
                    var f = res["forecast"]?[diff];
                    string condicao = f["description"]?.ToString() ?? "";
                    string icon = ObterIconeClima(condicao);
                    lblClima.Text = $"{f["weekday"]} ({dia}): {f["max"]}°C/{f["min"]}°C - {condicao} {icon}";
                }
                else lblClima.Text = $"Dia {dia}: Previsão indisponível";

                if (lblClima.Text.Contains("°C"))
                {
                    var m = Regex.Match(lblClima.Text, @"(\d+)°C");
                    if (m.Success)
                    {
                        int t = int.Parse(m.Groups[1].Value);
                        lblClima.ForeColor = t < 15 ? System.Drawing.Color.Blue : (t > 28 ? System.Drawing.Color.OrangeRed : System.Drawing.Color.Black);
                    }
                }
            }
            catch { lblClima.Text = "Erro Clima"; }
        }
        private string ObterIconeClima(string condicao)
        {
            condicao = condicao.ToLower();
            if (condicao.Contains("chuva")) return "🌧️";
            if (condicao.Contains("tempestade")) return "⛈️";
            if (condicao.Contains("nublado")) return "☁️";
            if (condicao.Contains("claro") || condicao.Contains("sol") || condicao.Contains("limpo")) return "☀️";
            if (condicao.Contains("nuvens") || condicao.Contains("parcial")) return "⛅";
            return "🌡️";
        }
        private List<DataRow> OrdenarPorHorario(List<DataRow> l)
        {
            return l.OrderBy(r => int.TryParse(r[INDEX_ORDEM]?.ToString(), out int o) ? o : 999)
                    .ThenBy(r => r[INDEX_HORARIO]?.ToString()).ToList();
        }
        private bool TryParseHorario(string? t, out TimeSpan i, out TimeSpan f)
        {
            i = f = TimeSpan.Zero;
            if (t == null) return false;
            var p = t.Split(new[] { 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries);
            return p.Length == 2 && TimeSpan.TryParse(p[0].Trim(), out i) && TimeSpan.TryParse(p[1].Trim(), out f);
        }
        private void DataGridView2_CellEnter(object? sender, DataGridViewCellEventArgs e) { if (e.ColumnIndex > 1) SendKeys.Send("{F4}"); }
        private void DataGridView2_CurrentCellDirtyStateChanged(object? sender, EventArgs e)

        {
            if (dataGridView2.IsCurrentCellDirty)

            {
                dataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit);
                CalcularTotais();
                PintarPostos();
                AtualizarItinerarios();
            }
        }
        private void CalcularTotais()
        {
            // Percorre todas as linhas para achar os cabeçalhos/rodapés (linhas amarelas)
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                // Usa o nome da coluna para garantir segurança
                string textoLinha = dataGridView2.Rows[i].Cells["Nome"].Value?.ToString()?.ToUpper() ?? "";

                // Se achou a linha de total (ex: "OPERADORES (10)")
                if (textoLinha.Contains("OPERADORES") || textoLinha.Contains("APRENDIZ") || textoLinha.Contains("CFTV"))
                {
                    // Percorre cada coluna de horário (começando da 3)
                    for (int c = 3; c < dataGridView2.Columns.Count; c++)
                    {
                        int count = 0;

                        // Olha para trás (linhas acima) até encontrar o próximo cabeçalho ou o topo
                        for (int k = i - 1; k >= 0; k--)
                        {
                            string tAnt = dataGridView2.Rows[k].Cells["Nome"].Value?.ToString()?.ToUpper() ?? "";

                            // Se bateu no bloco anterior (ex: acabou os operadores e chegou no cabeçalho anterior), para de contar
                            if (tAnt.Contains("OPERADORES") || tAnt.Contains("APRENDIZ") || tAnt.Contains("CFTV"))
                                break;

                            // Se a célula tem valor, soma +1
                            if (!string.IsNullOrWhiteSpace(dataGridView2.Rows[k].Cells[c].Value?.ToString()))
                                count++;
                        }

                        // IMPORTANTE: O Grid é de ComboBox, mas o Total é Número.
                        // Precisamos converter a célula para Texto para aceitar o número
                        if (dataGridView2.Rows[i].Cells[c] is DataGridViewComboBoxCell)
                        {
                            dataGridView2.Rows[i].Cells[c] = new DataGridViewTextBoxCell();
                        }

                        // Escreve o total na linha amarela (se for 0, deixa vazio para limpar o visual)
                        dataGridView2.Rows[i].Cells[c].Value = count > 0 ? count.ToString() : "";

                        // Opcional: Centralizar para ficar bonito
                        if (count > 0)
                        {
                            dataGridView2.Rows[i].Cells[c].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dataGridView2.Rows[i].Cells[c].Style.Font = new Font("Bahnschrift Condensed", 10, FontStyle.Bold);
                        }
                    }
                }
            }
        }
        private void AtualizarItinerarios()
        {
            if (flowLayoutPanel1 == null) return;
            flowLayoutPanel1.SuspendLayout();
            flowLayoutPanel1.Controls.Clear();
            foreach (DataGridViewRow r in dataGridView2.Rows)
            {
                if (r.Tag?.ToString() == "IGNORAR" || !r.Visible) continue;
                string nome = r.Cells["Nome"].Value?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(nome) || nome.Contains("OPERADORES") || nome.Contains("CFTV")) continue;

                var cartao = new CartaoFuncionario { Nome = nome };
                bool tem = false;
                for (int c = 3; c < dataGridView2.Columns.Count; c++)
                {
                    string p = r.Cells[c].Value?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(p))
                    {
                        cartao.Itens.Add(new ItemItinerario { Horario = dataGridView2.Columns[c].HeaderText, Posto = p, CorFundo = r.Cells[c].Style.BackColor, CorTexto = r.Cells[c].Style.ForeColor });
                        tem = true;
                    }
                }
                if (tem) flowLayoutPanel1.Controls.Add(CriarPainelCartao(cartao));
            }
            flowLayoutPanel1.ResumeLayout();
        }
        private Panel CriarPainelCartao(CartaoFuncionario d)
        {
            Panel p = new Panel { Width = 200, AutoSize = true, BackColor = System.Drawing.Color.White, Margin = new Padding(10) };
            p.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, p.ClientRectangle, System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid, System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid, System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid, System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid);

            Label l = new Label { Text = $"{d.Nome}", Dock = DockStyle.Top, TextAlign = ContentAlignment.MiddleCenter, Font = new System.Drawing.Font("Impact", 12) };
            p.Controls.Add(l);

            int y = 30;
            foreach (var i in d.Itens)
            {
                Label lh = new Label { Text = i.Horario, Location = new Point(5, y), Width = 90, Font = new System.Drawing.Font("Arial Narrow", 10, FontStyle.Bold) };
                Label lp = new Label { Text = i.Posto, Location = new Point(100, y), Width = 90, BackColor = i.CorFundo, ForeColor = i.CorTexto, TextAlign = ContentAlignment.MiddleCenter, Font = new System.Drawing.Font("Arial", 10, FontStyle.Bold) };
                p.Controls.Add(lh); p.Controls.Add(lp);
                y += 25;
            }
            return p;
        }
        // Configuração Mensal
        private void ConfigurarGridMensal()
        {
            if (dataGridView1.DataSource == null) return;
            dataGridView1.RowHeadersVisible = false;
            for (int i = 0; i <= INDEX_NOME; i++) dataGridView1.Columns[i].Frozen = true;
        }
        private void btnExportar_Click(object sender, EventArgs e)
        {

            // Verifica se tem dados carregados
            if (_tabelaMensal == null || _tabelaMensal.Rows.Count == 0)
            {
                MessageBox.Show("Importe a planilha antes de gerar o relatório.");
                return;
            }

            try
            {
                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Excel Workbook|*.xlsx",
                    FileName = $"Relatorio_Mensal_Completo.xlsx"
                };

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor; // Mostra que está processando

                    // Guarda o dia que o usuário estava olhando para voltar nele depois
                    int diaOriginal = _diaSelecionado;

                    using (var workbook = new XLWorkbook())
                    {
                        // Calcula quantos dias tem no mês baseado nas colunas do Excel importado
                        // Se tiver colunas suficientes, vai até dia 31
                        int maxDias = _tabelaMensal.Columns.Count - INDEX_DIA_INICIO;
                        if (maxDias > 31) maxDias = 31; // Trava em 31 dias

                        // --- LOOP PRINCIPAL: GERA UMA ABA POR DIA ---
                        for (int d = 1; d <= maxDias; d++)
                        {
                            // 1. Força o sistema a processar o dia 'd'
                            _diaSelecionado = d;
                            ProcessarEscalaDoDia(); // Roda toda a lógica (Cores, Folguista, Windison, etc)

                            // 2. Cria a aba no Excel
                            var worksheet = workbook.Worksheets.Add($"Dia {d}");

                            // -----------------------------------------------------
                            // EXPORTAÇÃO DO CABEÇALHO
                            // -----------------------------------------------------
                            for (int i = 0; i < dataGridView2.Columns.Count; i++)
                            {
                                var cell = worksheet.Cell(1, i + 1);
                                cell.Value = dataGridView2.Columns[i].HeaderText;
                                cell.Style.Font.Bold = true;
                                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }

                            // -----------------------------------------------------
                            // EXPORTAÇÃO DOS DADOS E CORES (PINTURA)
                            // -----------------------------------------------------
                            for (int i = 0; i < dataGridView2.Rows.Count; i++)
                            {
                                // Se a linha for invisível, não exporta
                                if (!dataGridView2.Rows[i].Visible) continue;

                                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                                {
                                    var dgvCell = dataGridView2.Rows[i].Cells[j];
                                    var xlCell = worksheet.Cell(i + 2, j + 1); // +2 pois linha 1 é cabeçalho

                                    // 1. Valor (Texto)
                                    if (dgvCell.Value != null)
                                        xlCell.Value = dgvCell.Value.ToString();

                                    // 2. Cor de Fundo (Converte de WinForms para ClosedXML)
                                    var corWinForms = dgvCell.Style.BackColor;
                                    if (corWinForms != Color.Empty && corWinForms != Color.Transparent && corWinForms.Name != "0")
                                    {
                                        xlCell.Style.Fill.BackgroundColor = XLColor.FromColor(corWinForms);
                                    }

                                    // 3. Cor da Fonte
                                    var corTexto = dgvCell.Style.ForeColor;
                                    if (corTexto != Color.Empty && corTexto != Color.Black)
                                    {
                                        xlCell.Style.Font.FontColor = XLColor.FromColor(corTexto);
                                    }

                                    // 4. Bordas e Alinhamento
                                    xlCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                    xlCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                }
                            }

                            // Ajusta largura das colunas para caber o texto
                            worksheet.Columns().AdjustToContents();
                        }

                        // Salva o arquivo no disco
                        workbook.SaveAs(sfd.FileName);
                    }

                    // Restaura a visualização para o dia que o usuário estava antes
                    _diaSelecionado = diaOriginal;
                    ProcessarEscalaDoDia();

                    Cursor.Current = Cursors.Default;
                    MessageBox.Show("Relatório Mensal (Pasta de Trabalho) gerado com sucesso!", "Sucesso");
                }
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Erro ao gerar relatório: " + ex.Message);
            }
        }
        private void DataGridView2_DataError(object? sender, DataGridViewDataErrorEventArgs e)
        {
            // CÓDIGO NUCLEAR: Mata qualquer erro do Grid e força aceitar o valor.
            // Não importa qual seja o erro, não mostre a caixa de diálogo.
            e.ThrowException = false;

            // Importante para o contador funcionar: Diz pro grid "Aceite esse valor mesmo assim"
            e.Cancel = false;
        }
        private Bitmap CapturarImagemDoGrid()
        {
            // 1. Guarda o tamanho original para não estragar a tela
            int alturaOriginal = dataGridView2.Height;
            int larguraOriginal = dataGridView2.Width;
            bool scrollOriginal = dataGridView2.ScrollBars != ScrollBars.None;

            try
            {
                // 2. Remove barras de rolagem e expande o grid para caber TUDO
                dataGridView2.ScrollBars = ScrollBars.None;

                int alturaTotal = dataGridView2.ColumnHeadersHeight + (dataGridView2.Rows.Count * dataGridView2.RowTemplate.Height);
                int larguraTotal = 0;

                foreach (DataGridViewColumn col in dataGridView2.Columns)
                    if (col.Visible) larguraTotal += col.Width;

                // Adiciona um respiro visual
                dataGridView2.Height = alturaTotal + 20;
                dataGridView2.Width = larguraTotal + 20;

                // 3. Cria a imagem em memória
                Bitmap bmp = new Bitmap(dataGridView2.Width, dataGridView2.Height);

                // 4. Desenha o Grid dentro da imagem
                dataGridView2.DrawToBitmap(bmp, new Rectangle(0, 0, dataGridView2.Width, dataGridView2.Height));

                return bmp;
            }
            finally
            {
                // 5. Restaura o Grid para o tamanho normal (Isso é CRUCIAL)
                dataGridView2.Height = alturaOriginal;
                dataGridView2.Width = larguraOriginal;
                if (scrollOriginal) dataGridView2.ScrollBars = ScrollBars.Both;
            }
        }
        private async Task EnviarParaApiNode(string caminhoImagem)
        {
            // Substitua pelo ID real do seu grupo (Descubra olhando o console do Node)
            // IDs de grupo geralmente terminam com @g.us
            string ID_DO_GRUPO = "120363421902743004@g.us";
            
            //"554188807362-1423694264@g.us";

            var payload = new
            {
                caminhoImagem = caminhoImagem,
                grupoId = ID_DO_GRUPO,
                legenda = $"Escala atualizada - Dia {_diaSelecionado} 📅"
            };

            using (var client = new HttpClient())
            {
                try
                {
                    var json = Newtonsoft.Json.JsonConvert.SerializeObject(payload);
                    var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

                    var response = await client.PostAsync("http://localhost:3000/enviar-escala", content);

                    if (response.IsSuccessStatusCode)
                    {
                        MessageBox.Show("✅ Enviado para o Grupo com Sucesso!");
                    }
                    else
                    {
                        MessageBox.Show("❌ Erro no Bot: " + await response.Content.ReadAsStringAsync());
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("❌ Erro de conexão com o Bot. Verifique se o server.exe está rodando.\n\n" + ex.Message);
                }
            }
        }
        private async void btnWhatsapp_Click(object sender, EventArgs e)
        {
            // 1. Tenta garantir que o bot está aberto
            // Se já estiver aberto, ele não faz nada (graças à trava no inicio da função)
            // Se estiver fechado, ele abre.
            IniciarBotEmBackground();

            // Verifica se precisou abrir agora (se não tinha node rodando antes)
            // Se acabou de abrir, espera 5 segundos pro servidor subir
            if (System.Diagnostics.Process.GetProcessesByName("node").Length == 0)
            {
                Cursor.Current = Cursors.WaitCursor;
                await Task.Delay(5000); // Espera o Node carregar
                Cursor.Current = Cursors.Default;
            }

            // 2. Tira o print e envia
            string tempPath = Path.Combine(Path.GetTempPath(), "escala_temp.png");

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                using (Bitmap imagem = CapturarImagemDoGrid())
                {
                    imagem.Save(tempPath, System.Drawing.Imaging.ImageFormat.Png);
                }

                await EnviarParaApiNode(tempPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro no envio: " + ex.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                // Procura e mata o processo do bot e do node ao fechar o sistema
                foreach (var process in System.Diagnostics.Process.GetProcessesByName("node"))
                {
                    process.Kill();
                }
                foreach (var process in System.Diagnostics.Process.GetProcessesByName("bot-estacionamento"))
                {
                    process.Kill();
                }
            }
            catch
            {
                // Se der erro ao fechar, apenas ignora
            }
        }
    }
}
