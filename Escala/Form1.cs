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
                dataGridView2.KeyDown += DataGridView2_KeyDown;
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
        }

        private void Form1_Load(object? sender, EventArgs e)
        {
            try
            {
                DatabaseService.Initialize();
                _tabelaMensal = DatabaseService.GetMonthlyData();

                if (_tabelaMensal.Rows.Count > 0 && dataGridView1 != null)
                {
                    dataGridView1.DataSource = null;
                    dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = _tabelaMensal;
                    ConfigurarGridMensal();
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
            using (var form = new FormGerenciar())
            {
                form.ShowDialog();
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

            ConfigurarGridEscalaDiaria();

            int indiceColunaDia = INDEX_DIA_INICIO + (_diaSelecionado - 1);
            if (indiceColunaDia >= _tabelaMensal.Columns.Count) return;

            var listaSUP = new List<DataRow>();
            var listaOP = new List<DataRow>();
            var listaJV = new List<DataRow>();
            var listaCFTV = new List<DataRow>(); // Essa lista vai receber o Windison agora
            var listaFolga = new List<DataRow>();
            var listaFerias = new List<DataRow>();

            foreach (DataRow linha in _tabelaMensal.Rows)
            {
                string nome = linha[INDEX_NOME]?.ToString() ?? "";
                string horario = linha[INDEX_HORARIO]?.ToString() ?? ""; // "FOLGUISTA" está aqui
                string funcao = (INDEX_FUNCAO < _tabelaMensal.Columns.Count) ? (linha[INDEX_FUNCAO]?.ToString()?.ToUpper() ?? "") : "";
                string nomeUpper = nome.ToUpper();

                if (string.IsNullOrWhiteSpace(nome) || nomeUpper.Contains("NOME")) continue;

                // --- CORREÇÃO DO ERRO AQUI ---
                // Antes: if (!horario.Contains(":")) continue;
                // Agora: Deixa passar se tiver ":" OU se estiver escrito "FOLGUISTA"
                if (!horario.Contains(":") && !horario.ToUpper().Contains("FOLGUISTA") && !horario.ToUpper().Contains("SIV"))
                    continue;
                // -----------------------------

                string? statusNoDia = linha[indiceColunaDia]?.ToString()?.ToUpper().Trim();

                if (new[] { "X", "FOLGA", "O" }.Contains(statusNoDia)) { listaFolga.Add(linha); continue; }
                if (new[] { "F", "FERIAS", "FÉRIAS", "AT", "ATESTADO" }.Contains(statusNoDia)) { listaFerias.Add(linha); continue; }

                if (funcao.Contains("SUP") || funcao.Contains("LIDER") || nomeUpper.Contains("ISAIAS") || nomeUpper.Contains("ROGÉRIO"))
                    listaSUP.Add(linha);
                else if (funcao.Contains("JV") || funcao.Contains("APRENDIZ") || nomeUpper.Contains("JOAO"))
                    listaJV.Add(linha);
                else if (funcao.Contains("CFTV") || nomeUpper.Contains("CFTV"))
                    listaCFTV.Add(linha); // O Windison vai cair aqui agora
                else
                    listaOP.Add(linha);
            }

            var assignments = DatabaseService.GetAssignmentsForDay(_diaSelecionado);

            // Inserção no Grid
            InserirBloco("OPERADORES", OrdenarPorHorario(listaOP), true, assignments);
            InserirBloco("APRENDIZ", OrdenarPorHorario(listaJV), true, assignments);

            // O Windison vai ser desenhado aqui
            InserirBloco("CFTV", OrdenarPorHorario(listaCFTV), false, assignments);

            InserirListaSimples("FOLGA", listaFolga);
            InserirListaSimples("FÉRIAS|ATESTADOS", listaFerias);

            CalcularTotais();
            PintarHorarios();
            PintarPostos();
            PintarHorarioFunc();

            // Agora a função vai encontrar o Windison na listaCFTV!
            AplicarLogicaFolguistaCFTV(listaCFTV);
            AplicarLogicaIntermediarioCFTV(listaCFTV);

            if (flowLayoutPanel1 != null) AtualizarItinerarios();
        }

        private void AplicarLogicaFolguistaCFTV(List<DataRow> listaCFTV)
        {
            // 1. Identificar QUEM é o folguista
            string nomeDoFolguista = "";
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

            // 2. Encontrar a linha visual no Grid
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

            // 3. Reset visual (Cinza)
            for (int c = 3; c < dataGridView2.Columns.Count; c++)
            {
                rowFolguista.Cells[c].Style.BackColor = System.Drawing.Color.DarkGray;
                rowFolguista.Cells[c].Style.ForeColor = System.Drawing.Color.White;
                rowFolguista.Cells[c].ReadOnly = true;
            }

            // -----------------------------------------------------------
            // 4. LÓGICA DE PRIORIDADES DE COBERTURA
            // -----------------------------------------------------------
            int colDia = INDEX_DIA_INICIO + (_diaSelecionado - 1);

            // PADRÃO: Define o horário que ele faz se NINGUÉM faltar.
            // Você tinha colocado vários IFs confusos. O correto é escolher um padrão.
            // Busca do banco/arquivo. Se vier vazio, usa um fallback
            string horarioParaAssumir = DatabaseService.GetHorarioPadraoFolguista();
            if (string.IsNullOrWhiteSpace(horarioParaAssumir)) horarioParaAssumir = "12:40 X 21:00";

            bool achouFaltaPrioritaria = false; // Prioridade 1 (Noite)
            bool achouAlgumaFalta = false;      // Prioridade 2 (Outros)

            foreach (DataRow pessoa in listaCFTV)
            {
                string horarioPessoa = pessoa[INDEX_HORARIO]?.ToString()?.ToUpper() ?? "";
                string nomePessoa = pessoa[INDEX_NOME].ToString().ToUpper();

                // Pula o próprio folguista
                if (nomePessoa == nomeDoFolguista) continue;

                // Verifica Status de Falta
                string status = pessoa[colDia].ToString().ToUpper().Trim();
                bool estaDeFolga = (status == "FOLGA" || status == "X" || status == "F" || status == "O" || status.Contains("FÉRIAS") || status.Contains("ATESTADO"));

                if (estaDeFolga)
                {
                    // PRIORIDADE 1: Turno da Noite (16:40 ou saída 00:40/01:00)
                    if (horarioPessoa.Contains("16:40") || horarioPessoa.Contains("00:40"))
                    {
                        horarioParaAssumir = pessoa[INDEX_HORARIO].ToString(); // Assume o horário exato da planilha
                        achouFaltaPrioritaria = true;
                        break; // Encontrou a prioridade máxima (Noite), para de procurar!
                    }

                    // PRIORIDADE 2: Qualquer outro turno (apenas se ainda não achou uma prioridade 1)
                    if (!achouAlgumaFalta)
                    {
                        horarioParaAssumir = pessoa[INDEX_HORARIO].ToString();
                        achouAlgumaFalta = true;
                        // Não damos break aqui porque ainda podemos achar uma prioridade 1 (16:40) mais pra frente na lista
                    }
                }
            }

            // 5. Aplica o horário decidido
            rowFolguista.Cells["HORARIO"].Value = horarioParaAssumir;
            var partes = horarioParaAssumir.Split(new[] { 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries);

            if (partes.Length == 2)
            {
                PintarIntervaloBranco(rowFolguista, partes[0].Trim(), partes[1].Trim());
            }
        }
        private void AplicarLogicaIntermediarioCFTV(List<DataRow> listaCFTV)
        {
            // 1. Define o Horário Padrão (Vindo do ComboBox da Aba 4)
            string horarioPadrao = DatabaseService.GetHorarioPadraoIntermediario();
            if (string.IsNullOrWhiteSpace(horarioPadrao)) horarioPadrao = "12:40 X 21:00";

            // 2. Varre a lista procurando quem tem o horário "12:40 x 20:40"
            foreach (DataRow dados in listaCFTV)
            {
                string horarioExcel = dados[INDEX_HORARIO]?.ToString() ?? "";
                string nome = dados[INDEX_NOME]?.ToString()?.ToUpper() ?? "";

                // A LÓGICA CORRETA AGORA:
                // Procura especificamente pelo texto "12:40" e "20:40" na mesma célula
                if (horarioExcel.Contains("12:40") && horarioExcel.Contains("21:00"))
                {
                    // Achamos o Intermediário! Agora vamos achá-lo no Grid visual.
                    DataGridViewRow rowInter = null;
                    foreach (DataGridViewRow row in dataGridView2.Rows)
                    {
                        string nomeNoGrid = row.Cells["Nome"].Value?.ToString()?.ToUpper() ?? "";

                        // Tenta achar o nome exato ou contendo parte
                        if (nomeNoGrid == nome || (nomeNoGrid.Contains(nome) && nome.Length > 3))
                        {
                            rowInter = row;
                            break;
                        }
                    }

                    if (rowInter != null)
                    {
                        // 3. Reseta a cor para cinza (limpa formatações anteriores)
                        for (int c = 3; c < dataGridView2.Columns.Count; c++)
                        {
                            rowInter.Cells[c].Style.BackColor = System.Drawing.Color.DarkGray;
                            rowInter.Cells[c].Style.ForeColor = System.Drawing.Color.White;
                            rowInter.Cells[c].ReadOnly = true;
                        }

                        // 4. Aplica o horário do ComboBox e Pinta de Branco
                        rowInter.Cells["HORARIO"].Value = horarioPadrao;

                        var partes = horarioPadrao.Split(new[] { 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries);
                        if (partes.Length == 2)
                        {
                            PintarIntervaloBranco(rowInter, partes[0].Trim(), partes[1].Trim());
                        }
                    }
                    // Como geralmente só tem 1 intermediário, podemos parar de procurar (break)
                    // Se tiver mais de um, remova o 'break' abaixo.
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
                        if (ini <= hIni && fimAj >= hFimAj)
                        {
                            row.Cells[c].Style.BackColor = System.Drawing.Color.White;
                            row.Cells[c].Style.ForeColor = System.Drawing.Color.Black;
                            row.Cells[c].ReadOnly = false;
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

        private void InserirBloco(string titulo, List<DataRow> lista, bool gerarCartao, Dictionary<string, Dictionary<string, string>> assignments = null)
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

                if (assignments != null && assignments.ContainsKey(nome))
                {
                    var userPosts = assignments[nome];
                    for (int c = 3; c < dataGridView2.Columns.Count; c++)
                    {
                        string slot = dataGridView2.Columns[c].HeaderText;
                        if (userPosts.ContainsKey(slot)) r.Cells[c].Value = userPosts[slot];
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

        private void DataGridView2_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && dataGridView2.SelectedCells.Count > 0)
            {
                foreach (DataGridViewCell cell in dataGridView2.SelectedCells)
                {
                    if (cell.ColumnIndex <= 2) continue;
                    cell.Value = "";
                    cell.Style.BackColor = System.Drawing.Color.DarkGray;
                }
                e.Handled = true;
                AtualizarItinerarios();
            }
        }

        private void DataGridView2_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
        {
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

        // Classes Aninhadas
        public class SeletorPlanilha : Form
        {
            public ComboBox CbPlanilhas;
            private Button BtnOk;
            public SeletorPlanilha(List<string> planilhas)
            {
                Size = new Size(300, 150); StartPosition = FormStartPosition.CenterScreen;
                CbPlanilhas = new ComboBox { Left = 10, Top = 30, Width = 260, DataSource = planilhas, DropDownStyle = ComboBoxStyle.DropDownList };
                BtnOk = new Button { Text = "OK", Left = 190, Top = 70, DialogResult = DialogResult.OK };
                Controls.Add(new Label { Text = "Selecione a aba:", Left = 10, Top = 10 });
                Controls.Add(CbPlanilhas); Controls.Add(BtnOk);
            }
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
    }
}






// =========================================================
// CLASSE EXTRA: GERENCIADOR (Pode ficar no mesmo arquivo)
// =========================================================


public class FormGerenciar : Form
{
    private ListBox lbPostos, lbHorarios;
    private TextBox txtPosto, txtHorario, txtHorarioPadrao;
    private ComboBox CbHorarioPadraoFolguista; // ComboBox para o Folguista
    private ComboBox cbHorarioPadrao; // Novo ComboBox para o Intermediário
    private Button btnAddPosto, btnDelPosto, btnAddHorario, btnDelHorario, btnSalvarPadrao, btnSalvarIntermediario;

    public FormGerenciar()
    {
        Text = "Gerenciar Listas e Configurações"; Size = new Size(500, 500); StartPosition = FormStartPosition.CenterParent;
        TabControl tabs = new TabControl { Dock = DockStyle.Fill };

        // -------------------------------------------------------
        // ABA 1: POSTOS
        // -------------------------------------------------------
        TabPage tabP = new TabPage("Postos");
        lbPostos = new ListBox { Location = new Point(10, 10), Size = new Size(200, 250) };
        txtPosto = new TextBox { Location = new Point(220, 10), Size = new Size(150, 23) };
        btnAddPosto = new Button { Text = "Add", Location = new Point(220, 40) };
        btnDelPosto = new Button { Text = "Del", Location = new Point(10, 270), BackColor = System.Drawing.Color.LightCoral };

        btnAddPosto.Click += (s, e) => { if (!string.IsNullOrEmpty(txtPosto.Text)) { DatabaseService.AdicionarPosto(txtPosto.Text); Carregar(); txtPosto.Clear(); } };
        btnDelPosto.Click += (s, e) => { if (lbPostos.SelectedItem != null) { DatabaseService.RemoverPosto(lbPostos.SelectedItem.ToString()); Carregar(); } };
        tabP.Controls.AddRange(new Control[] { lbPostos, txtPosto, btnAddPosto, btnDelPosto });

        // -------------------------------------------------------
        // ABA 2: HORÁRIOS (Lista para Colunas)
        // -------------------------------------------------------
        TabPage tabH = new TabPage("Horários (Colunas)");
        lbHorarios = new ListBox { Location = new Point(10, 10), Size = new Size(200, 250) };
        txtHorario = new TextBox { Location = new Point(220, 10), Size = new Size(150, 23) };
        btnAddHorario = new Button { Text = "Add", Location = new Point(220, 40) };
        btnDelHorario = new Button { Text = "Del", Location = new Point(10, 270), BackColor = System.Drawing.Color.LightCoral };

        btnAddHorario.Click += (s, e) => { if (!string.IsNullOrEmpty(txtHorario.Text)) { DatabaseService.AdicionarHorario(txtHorario.Text); Carregar(); txtHorario.Clear(); } };
        btnDelHorario.Click += (s, e) => { if (lbHorarios.SelectedItem != null) { DatabaseService.RemoverHorario(lbHorarios.SelectedItem.ToString()); Carregar(); } };
        tabH.Controls.AddRange(new Control[] { lbHorarios, txtHorario, btnAddHorario, btnDelHorario });

        // -------------------------------------------------------
        // ABA 3: FOLGUISTA (TextBox Simples)
        // -------------------------------------------------------
        TabPage tabC = new TabPage("Horários folguistas");
        Label lblExplica = new Label { Text = "Horário Padrão do Folguista (Se ninguém faltar):", Location = new Point(10, 20), AutoSize = true, Font = new Font("Arial", 10, FontStyle.Bold) };
        CbHorarioPadraoFolguista = new ComboBox { Location = new Point(10, 50), Width = 200, Font = new Font("Arial", 12) }; // Ex: 12:40 x 21:00
        CbHorarioPadraoFolguista.Items.AddRange(new object[] {
             "07:00 X 15:20",
             "09:40 X 18:00",
             "10:40 X 19:00",
             "11:40 X 20:00",
             "12:40 X 21:00",
             "16:40 X 01:00"
         });
        btnSalvarPadrao = new Button { Text = "Salvar Padrão", Location = new Point(220, 48), Width = 100, Height = 30, BackColor = System.Drawing.Color.LightGreen };

        btnSalvarPadrao.Click += (s, e) =>
        {
            DatabaseService.SetHorarioPadraoFolguista(CbHorarioPadraoFolguista.Text);
            MessageBox.Show("Horário do Folguista atualizado!");
        };
        tabC.Controls.AddRange(new Control[] { lblExplica, CbHorarioPadraoFolguista, btnSalvarPadrao });

        // -------------------------------------------------------
        // ABA 4: CFTV (INTERMEDIÁRIO) - ComboBox
        // -------------------------------------------------------
        TabPage tabD = new TabPage("CFTV (Intermediário)");
        Label lblExplicaInter = new Label { Text = "Horário Padrão do Intermediário (12:40):", Location = new Point(10, 20), AutoSize = true, Font = new Font("Arial", 10, FontStyle.Bold) };

        // ComboBox com opções pré-definidas
        cbHorarioPadrao = new ComboBox { Location = new Point(10, 50), Width = 200, Font = new Font("Arial", 12), DropDownStyle = ComboBoxStyle.DropDown };
        cbHorarioPadrao.Items.AddRange(new object[] {
             "07:00 X 15:20",
             "09:40 X 18:00",
             "10:40 X 19:00",
             "11:40 X 20:00",
             "12:40 X 21:00",
             "16:40 X 01:00"
         });

        btnSalvarIntermediario = new Button { Text = "Salvar Interm.", Location = new Point(220, 48), Width = 100, Height = 30, BackColor = System.Drawing.Color.LightGreen };

        btnSalvarIntermediario.Click += (s, e) =>
        {
            // Salva o texto que estiver selecionado ou digitado no ComboBox
            DatabaseService.SetHorarioPadraoIntermediario(cbHorarioPadrao.Text);
            MessageBox.Show("Horário Intermediário atualizado!");
        };

        tabD.Controls.AddRange(new Control[] { lblExplicaInter, cbHorarioPadrao, btnSalvarIntermediario });

        // Adiciona as abas
        tabs.TabPages.Add(tabP);
        tabs.TabPages.Add(tabH);
        tabs.TabPages.Add(tabC);
        tabs.TabPages.Add(tabD); // Nova Aba
        Controls.Add(tabs);

        Carregar();
    }

    private void Carregar()
    {
        // Carrega Postos
        lbPostos.Items.Clear();
        foreach (var p in DatabaseService.GetPostosConfigurados()) lbPostos.Items.Add(p);

        // Carrega Horários das Colunas
        lbHorarios.Items.Clear();
        foreach (var h in DatabaseService.GetHorariosConfigurados()) lbHorarios.Items.Add(h);

        // Carrega o Horário Padrão do Folguista
        cbHorarioPadrao.Text = DatabaseService.GetHorarioPadraoFolguista();

        // Carrega o Horário Padrão do Intermediário
        cbHorarioPadrao.Text = DatabaseService.GetHorarioPadraoIntermediario();
    }
}

public static class ExtensionMethods
{
    public static void DoubleBuffered(this DataGridView dgv, bool setting)
    {
        Type dgvType = dgv.GetType();
        System.Reflection.PropertyInfo? pi = dgvType.GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        if (pi != null) pi.SetValue(dgv, setting, null);
    }
}

public class ItemItinerario
{
    public string? Horario { get; set; }
    public string? Posto { get; set; }
    public System.Drawing.Color CorFundo { get; set; }
    public System.Drawing.Color CorTexto { get; set; }
}

public class CartaoFuncionario
{
    public string? Nome { get; set; }
    public List<ItemItinerario> Itens { get; set; } = new List<ItemItinerario>();
}