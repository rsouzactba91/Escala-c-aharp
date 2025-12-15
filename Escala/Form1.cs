using ClosedXML.Excel;
using System.Data;
using System.Drawing.Printing; // Necessário para impressão
using System.Text.RegularExpressions;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using System.Globalization;

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
        private const int INDEX_ORDEM = 3; // NOVO: Índice da coluna de ORDEM
        private const int INDEX_NOME = 4;
        private const int INDEX_DIA_INICIO = 5;

        private DataTable? _tabelaMensal;
        private int _diaSelecionado = 1;
        private int _paginaAtual = 0; // Controle de paginação
        private JObject? _previsaoCompleta; // Armazena a previsão completa da API

        public Form1()
        {
            InitializeComponent();

            // Ligações de Eventos (Garante que os botões funcionem)
            this.Load += Form1_Load;
            if (btnImportar != null) btnImportar.Click += button1_Click;
            if (CbSeletorDia != null) CbSeletorDia.SelectedIndexChanged += CbSeletorDia_SelectedIndexChanged;

            // Configuração do Grid de Edição (Escala Diária)
            if (dataGridView2 != null)
            {
                dataGridView2.DoubleBuffered(true);
                dataGridView2.CellEnter += DataGridView2_CellEnter;
                dataGridView2.CurrentCellDirtyStateChanged += DataGridView2_CurrentCellDirtyStateChanged;
                // Removemos a linha lateral para ficar mais limpo
                dataGridView2.RowHeadersVisible = false;
                // --- LIGA O BOTÃO DE IMPRIMIR AQUI ---
                // (Supondo que você criou o botão com o nome btnImprimir no Designer)
                // (Supondo que você criou o botão com o nome btnImprimir no Designer)
                if (btnImprimir != null) btnImprimir.Click += btnImprimir_Click;

                // Evento para capturar o DELETE
                dataGridView2.KeyDown += DataGridView2_KeyDown;
                
                // NOVO: Salvar ao editar célula
                dataGridView2.CellValueChanged += DataGridView2_CellValueChanged;
                
                // NOVO: Pintura customizada para "mesclar" células
                dataGridView2.CellPainting += DataGridView2_CellPainting;
                dataGridView2.RowPostPaint += DataGridView2_RowPostPaint;



            }
        }

        private void Form1_Load(object? sender, EventArgs e)
        {

            // Inicializa Banco de Dados
            try
            {
                DatabaseService.Initialize();
                // Carrega Tabela Mensal do Banco se existir
                _tabelaMensal = DatabaseService.GetMonthlyData();
                if (_tabelaMensal.Rows.Count > 0 && dataGridView1 != null)
                {
                     // Limpa colunas antes de vincular para evitar conflitos de congelamento
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

            // 1. Configura ComboBox de Dias
            if (CbSeletorDia != null)
            {
                CbSeletorDia.Items.Clear();
                for (int i = 1; i <= 31; i++) CbSeletorDia.Items.Add($"Dia {i}");
                int hoje = DateTime.Now.Day;
                CbSeletorDia.SelectedIndex = (hoje <= 31) ? hoje - 1 : 0;
            }

            // 2. Configura o Painel de Itinerários (Que você criou na TabPage3)
            if (flowLayoutPanel1 != null)
            {
                flowLayoutPanel1.AutoScroll = true; // Permite rolar se tiver muitos funcionários
                flowLayoutPanel1.FlowDirection = FlowDirection.LeftToRight;
                flowLayoutPanel1.WrapContents = true;
                flowLayoutPanel1.BackColor = System.Drawing.Color.WhiteSmoke;
                // --- CORREÇÃO AQUI ---
                // Define que, por padrão, TODAS as células nascem Cinza Escuro
                dataGridView2.DefaultCellStyle.BackColor = System.Drawing.Color.DarkGray;
                // Define a cor das linhas da grade (opcional, se quiser combinar)
                dataGridView2.GridColor = System.Drawing.Color.Black;

                dataGridView2.KeyDown += DataGridView2_KeyDown;

                // Dispara a busca do clima em segundo plano (não trava a tela)
                _ = AtualizarClimaAutomatico();
                // ---------------------
            }
        }
        private void AjustarHorariosMensal(int horasAdicionar)
        {
            if (dataGridView1.DataSource == null) return;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                var cellHorario = row.Cells[INDEX_HORARIO];
                string? texto = cellHorario.Value?.ToString(); // Fix CS8600

                if (string.IsNullOrWhiteSpace(texto)) continue;
                if (!texto.Contains("x")) continue;

                // Ex: "08:40 x 17:40"
                var partes = texto.Split('x');
                if (partes.Length != 2) continue;

                if (!TimeSpan.TryParse(partes[0].Trim(), out TimeSpan ini)) continue;
                if (!TimeSpan.TryParse(partes[1].Trim(), out TimeSpan fim)) continue;

                ini = ini.Add(TimeSpan.FromHours(horasAdicionar));
                fim = fim.Add(TimeSpan.FromHours(horasAdicionar));

                // Normaliza virada de dia
                ini = NormalizarHora(ini);
                fim = NormalizarHora(fim);

                cellHorario.Value = $"{ini:hh\\:mm} x {fim:hh\\:mm}";
            }

            // Reprocessa tudo
            ProcessarEscalaDoDia();
        }
        private TimeSpan NormalizarHora(TimeSpan hora)
        {
            while (hora.TotalMinutes < 0)
                hora = hora.Add(TimeSpan.FromHours(24));

            while (hora.TotalMinutes >= 1440)
                hora = hora.Subtract(TimeSpan.FromHours(24));

            return hora;
        }
        private void ProcessarEscalaDoDia()
        {
            // 1. Verificações Iniciais
            if (_tabelaMensal == null || _tabelaMensal.Rows.Count == 0) return;

            ConfigurarGridEscalaDiaria();

            int indiceColunaDia = INDEX_DIA_INICIO + (_diaSelecionado - 1);
            if (indiceColunaDia >= _tabelaMensal.Columns.Count)
            {
                MessageBox.Show($"Dados insuficientes para o Dia {_diaSelecionado}.");
                return;
            }

            // 2. Criação das Listas
            var listaSUP = new List<DataRow>();
            var listaOP = new List<DataRow>();
            var listaJV = new List<DataRow>();
            var listaCFTV = new List<DataRow>();
            // NOVAS LISTAS
            var listaFolga = new List<DataRow>();
            var listaFerias = new List<DataRow>();

            // 3. Loop de Classificação
            foreach (DataRow linha in _tabelaMensal.Rows)
            {
                string nome = linha[INDEX_NOME]?.ToString() ?? "";
                string horario = linha[INDEX_HORARIO]?.ToString() ?? "";
                string funcao = (INDEX_FUNCAO < _tabelaMensal.Columns.Count) ? (linha[INDEX_FUNCAO]?.ToString()?.ToUpper() ?? "") : ""; // Fix CS8602
                string nomeUpper = nome.ToUpper();

                if (string.IsNullOrWhiteSpace(nome) || nomeUpper.Contains("NOME")) continue;
                if (!horario.Contains(":")) continue;

                string? statusNoDia = linha[indiceColunaDia]?.ToString()?.ToUpper().Trim(); // Fix CS8600
                
                // NOVIDADE: Verifica Folga/Férias ANTES de pular
                // X ou O = Folga
                if (statusNoDia == "FOLGA" || statusNoDia == "X" || statusNoDia == "O")
                {
                    listaFolga.Add(linha);
                    continue;
                }
                // F = Férias, AT = Atestado
                if (statusNoDia == "FERIAS" || statusNoDia == "FÉRIAS" || statusNoDia == "F" || 
                    statusNoDia == "ATESTADO" || statusNoDia == "AT")
                {
                    listaFerias.Add(linha);
                    continue;
                }

                if (EhFolga(statusNoDia)) continue;

                // --- LÓGICA DE SEPARAÇÃO ---
                if (funcao.Contains("SUP") || funcao.Contains("COORD") || funcao.Contains("LIDER") ||
                    nomeUpper.Contains("ISAIAS") || nomeUpper.Contains("ROGÉRIO") || nomeUpper.Contains("ROGERIO"))
                {
                    listaSUP.Add(linha);
                }
                else if (funcao.Contains("JV") || funcao.Contains("JOVEM") || funcao.Contains("APRENDIZ") ||
                         nomeUpper.Contains("JOAO") || nomeUpper.Contains("JOÃO"))
                {
                    listaJV.Add(linha);
                }
                else if (funcao.Contains("CFTV") || nomeUpper.Contains("CFTV"))
                {
                    listaCFTV.Add(linha);
                }
                else
                {
                    listaOP.Add(linha);
                }
            }

            // 4. INSERÇÃO NO GRID (CORRIGIDA)
            // Usamos 'false' para quem NÃO deve ter cartão na aba 3
            // Usamos 'true' para quem DEVE ter cartão

            // CARREGA DO BANCO (Assignments)
            var assignments = DatabaseService.GetAssignmentsForDay(_diaSelecionado);

            //  InserirBloco("SUPERVISÃO", OrdenarPorHorario(listaSUP), false); // false = Sem Itinerário
            InserirBloco("OPERADORES", OrdenarPorHorario(listaOP), true, assignments);   // true = Com Itinerário
            InserirBloco("APRENDIZ", OrdenarPorHorario(listaJV), true, assignments);     // true = Com Itinerário
            InserirBloco("CFTV", OrdenarPorHorario(listaCFTV), false, assignments);      // false = Sem Itinerário
            
            // INSERE O RODAPÉ DE INFORMAÇÕES
            InserirListaSimples("FOLGA", listaFolga);
            InserirListaSimples("FÉRIAS|ATESTADOS", listaFerias);

            // 5. Automação dos Postos

            //  PreencherPostosAutomaticos("SUPERVISÃO", listaSUP, "SUP");
            //  PreencherPostosAutomaticos("OPERADORES", listaOP, "VALET");
            //  PreencherPostosAutomaticos("APRENDIZ", listaJV, "TREIN");
            //  PreencherPostosAutomaticos("CFTV", listaCFTV, "CFTV");

            // 6. Visual e Itinerários
            CalcularTotais();
            PintarHorarios();
            PintarPostos();

            // Atualiza a aba 3
            if (flowLayoutPanel1 != null) AtualizarItinerarios();
        }
        private void btnImprimir_Click(object? sender, EventArgs e)
        {
            _paginaAtual = 0; // Reseta para a primeira página
            PrintDocument pd = new PrintDocument();
            pd.DefaultPageSettings.Landscape = true; // Define Paisagem
            pd.DefaultPageSettings.Margins = new Margins(10, 10, 10, 10); // Margens minimas
            pd.PrintPage += new PrintPageEventHandler(ImprimirConteudo);

            // Abre a visualização antes de imprimir
            PrintPreviewDialog ppd = new PrintPreviewDialog();
            ppd.Document = pd;
            ppd.WindowState = FormWindowState.Maximized;
            ppd.ShowDialog();
        }
        private void ImprimirConteudo(object? sender, PrintPageEventArgs e)
        {
            if (e.Graphics == null) return; // Fix CS8602 for Graphics

            float y = e.MarginBounds.Top;
            float x = e.MarginBounds.Left;
            float larguraUtil = e.MarginBounds.Width;
            System.Drawing.Font fonteTitulo = new System.Drawing.Font("Arial", 16, FontStyle.Bold);

            if (_paginaAtual == 0)
            {
                // ================= PÁGINA 1: ESCALA (GRID) =================

                // Título
                e.Graphics.DrawString($"Escala do Dia {_diaSelecionado}", fonteTitulo, Brushes.Black, x, y);
                y += 40;

                // 1. CAPTURA A IMAGEM DO GRID
                int heightOriginalGrid = dataGridView2.Height;
                dataGridView2.Height = dataGridView2.RowCount * dataGridView2.RowTemplate.Height + dataGridView2.ColumnHeadersHeight;

                Bitmap bmpGrid = new Bitmap(dataGridView2.Width, dataGridView2.Height);
                dataGridView2.DrawToBitmap(bmpGrid, new Rectangle(0, 0, dataGridView2.Width, dataGridView2.Height));

                // Restaura tamanho original
                dataGridView2.Height = heightOriginalGrid;

                // Desenha o Grid
                float ratioGrid = (float)bmpGrid.Width / (float)bmpGrid.Height;
                float alturaGridNaFolha = larguraUtil / ratioGrid;

                // Limite para caber na página (se necessário)
                if (alturaGridNaFolha > e.MarginBounds.Height - 100) alturaGridNaFolha = e.MarginBounds.Height - 100;

                e.Graphics.DrawImage(bmpGrid, x, y, larguraUtil, alturaGridNaFolha);

                // Configura para próxima página
                e.HasMorePages = true;
                _paginaAtual++;
            }
            else
            {
                // ================= PÁGINA 2: ITINERÁRIOS =================

                // Título
                e.Graphics.DrawString("Itinerários / Cartões", fonteTitulo, Brushes.Black, x, y);
                y += 40;

                // 2. CAPTURA A IMAGEM DOS ITINERÁRIOS
                int alturaPainel = 0;
                foreach (System.Windows.Forms.Control c in flowLayoutPanel1.Controls)
                    alturaPainel = Math.Max(alturaPainel, c.Bottom);
                alturaPainel += 20;

                Bitmap bmpItinerario = new Bitmap(flowLayoutPanel1.Width, Math.Max(alturaPainel, 100));
                flowLayoutPanel1.DrawToBitmap(bmpItinerario, new Rectangle(0, 0, flowLayoutPanel1.Width, Math.Max(alturaPainel, flowLayoutPanel1.Height)));

                // Desenha os Itinerários
                float ratioItin = (float)bmpItinerario.Width / (float)bmpItinerario.Height;
                float alturaItinNaFolha = larguraUtil / ratioItin;

                // Se não couber, corta (comportamento original mantido/adaptado)
                if (alturaItinNaFolha > e.MarginBounds.Height - 100) alturaItinNaFolha = e.MarginBounds.Height - 100;

                e.Graphics.DrawImage(bmpItinerario, x, y, larguraUtil, alturaItinNaFolha);

                // Finaliza impressão
                e.HasMorePages = false;
                _paginaAtual = 0;
            }
        }
        private void AtualizarItinerarios()
        {
            if (flowLayoutPanel1 == null) return;

            flowLayoutPanel1.SuspendLayout(); // Pausa o desenho para não piscar
            flowLayoutPanel1.Controls.Clear(); // Limpa os cartões antigos

            var dados = GerarDadosDosItinerarios();

            foreach (var func in dados)
            {
                // Cria o cartão visual usando a função de estilo
                Panel pnl = CriarPainelCartao(func);
                flowLayoutPanel1.Controls.Add(pnl);
            }


            flowLayoutPanel1.ResumeLayout(); // Libera o desenho
        }
        private Panel CriarPainelCartao(CartaoFuncionario dados)
        {
            int larguraTotal = 200;
            int alturaLinha = 20;
            int alturaCabecalho = 20;

            System.Drawing.Font fonteCabecalho = new System.Drawing.Font("Impact", 12, FontStyle.Regular);
            System.Drawing.Font fonteHora = new System.Drawing.Font("Arial Narrow", 12, FontStyle.Bold);
            System.Drawing.Font fontePosto = new System.Drawing.Font("Arial", 12, FontStyle.Bold);

            Panel pnl = new Panel();
            pnl.Width = larguraTotal;
            pnl.AutoSize = true;
            pnl.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            pnl.BackColor = System.Drawing.Color.White;
            pnl.Margin = new Padding(10);

            // Mantém identidade visual mínima
            pnl.MinimumSize = new Size(larguraTotal, 220);

            pnl.Paint += (s, e) =>
            {
                ControlPaint.DrawBorder(
                    e.Graphics,
                    pnl.ClientRectangle,
                    System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid,
                    System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid,
                    System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid,
                    System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid
                );
            };

            Label lblHeader = new Label
            {
                Text = $"{_diaSelecionado:D2}/10 | {dados.Nome}",
                Font = fonteCabecalho,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = alturaCabecalho
            };

            Panel linhaDivisoria = new Panel
            {
                Height = 3,
                Dock = DockStyle.Top,
                BackColor = System.Drawing.Color.Black
            };

            pnl.Controls.Add(linhaDivisoria);
            pnl.Controls.Add(lblHeader);

            int yAtual = alturaCabecalho + 3;

            foreach (var item in dados.Itens)
            {
                Panel linha = new Panel
                {
                    Location = new Point(2, yAtual),
                    Size = new Size(larguraTotal - 4, alturaLinha),
                    BackColor = System.Drawing.Color.White
                };

                Label lblHora = new Label
                {
                    Text = item.Horario,
                    Font = fonteHora,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Location = new Point(0, 0),
                    Size = new Size(130, alturaLinha)
                };

                Panel divVert = new Panel
                {
                    Width = 2,
                    Height = alturaLinha,
                    Location = new Point(130, 0),
                    BackColor = System.Drawing.Color.Black
                };

                Label lblPosto = new Label
                {
                    Text = item.Posto,
                    Font = fontePosto,
                    ForeColor = item.CorTexto,
                    BackColor = item.CorFundo,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Location = new Point(132, 0),
                    Size = new Size(linha.Width - 132, alturaLinha)
                };

                Panel divHor = new Panel
                {
                    Height = 2,
                    Width = larguraTotal,
                    Location = new Point(0, yAtual + alturaLinha),
                    BackColor = System.Drawing.Color.Black
                };

                linha.Controls.Add(lblHora);
                linha.Controls.Add(divVert);
                linha.Controls.Add(lblPosto);

                pnl.Controls.Add(linha);
                pnl.Controls.Add(divHor);

                yAtual += alturaLinha + 2;
            }

            return pnl;
        }
        private List<CartaoFuncionario> GerarDadosDosItinerarios()
        {
            var lista = new List<CartaoFuncionario>();

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                // 1. VERIFICAÇÃO INFALÍVEL PELA TAG
                if (row.Tag != null && row.Tag.ToString() == "IGNORAR") continue;
                // NOVO: Se a linha estiver oculta no grid, não gera cartão
                if (!row.Visible) continue;

                // Se por acaso a Tag for nula (cabeçalhos antigos), verifica o nome só por segurança
                string? nome = row.Cells["Nome"].Value?.ToString(); // Fix CS8600
                if (string.IsNullOrWhiteSpace(nome)) continue;
                if (nome.Contains("SUPERVISÃO") || nome.Contains("CFTV") || nome.Contains("OPERADORES")) continue;

                // --- DAQUI PRA BAIXO TUDO IGUAL ---
                var cartao = new CartaoFuncionario { Nome = nome };
                bool temPosto = false;

                for (int c = 3; c < dataGridView2.Columns.Count; c++)
                {
                    var cell = row.Cells[c];
                    // REMOVIDO: Permitir que células cinzas (fora do horário) com conteúdo entrem no itinerário
                    // if (cell.Style.BackColor == System.Drawing.Color.DarkGray ||
                    //    cell.Style.BackColor == System.Drawing.Color.LightGray) continue;

                    string? posto = cell.Value?.ToString(); // Fix CS8600

                    if (!string.IsNullOrWhiteSpace(posto))
                    {
                        cartao.Itens.Add(new ItemItinerario
                        {
                            Horario = dataGridView2.Columns[c].HeaderText.Replace(" ", ""),
                            Posto = posto,
                            CorFundo = cell.Style.BackColor,
                            CorTexto = cell.Style.ForeColor
                        });
                        temPosto = true;
                    }
                }

                if (temPosto) lista.Add(cartao);
            }
            return lista;
        }
        private void PreencherPostosAutomaticos(string tipoFuncionario, List<DataRow> listaDados, string postoPadrao)
        {
            foreach (DataRow dados in listaDados)
            {
                string? nomeFuncionario = dados[INDEX_NOME].ToString(); // Fix CS8600
                string? horarioFunc = dados[INDEX_HORARIO].ToString(); // Fix CS8600

                if (!TryParseHorario(horarioFunc, out TimeSpan iniFunc, out TimeSpan fimFunc)) continue;

                // --- ARREDONDAMENTO REMOVIDO ---
                // if (fimFunc.Minutes == 40) ... (APAGADO)

                TimeSpan fimFuncAj = (fimFunc < iniFunc) ? fimFunc.Add(TimeSpan.FromHours(24)) : fimFunc;

                foreach (DataGridViewRow rowGrid in dataGridView2.Rows)
                {
                    if (rowGrid.Cells["Nome"].Value?.ToString() == nomeFuncionario)
                    {
                        for (int c = 3; c < dataGridView2.Columns.Count; c++)
                        {
                            string header = dataGridView2.Columns[c].HeaderText;
                            if (TryParseHorario(header, out TimeSpan iniCol, out TimeSpan fimCol))
                            {
                                TimeSpan fimColAj = (fimCol < iniCol) ? fimCol.Add(TimeSpan.FromHours(24)) : fimCol;
                                bool estaTrabalhando = (iniFunc < fimColAj) && (fimFuncAj > iniCol);

                                if (estaTrabalhando)
                                {
                                    if (string.IsNullOrWhiteSpace(rowGrid.Cells[c].Value?.ToString()))
                                    {
                                        rowGrid.Cells[c].Value = postoPadrao;
                                    }
                                }
                            }
                        }
                        break;
                    }
                }
            }
        }
        private void InserirBloco(string titulo, List<DataRow> lista, bool gerarCartao, Dictionary<string, Dictionary<string, string>> assignments = null)
        {
            if (lista.Count == 0) return;
            
            // assignments: StaffName -> { TimeSlot -> Post }

            foreach (var item in lista)
            {
                int idx = dataGridView2.Rows.Add();
                var r = dataGridView2.Rows[idx];

                // 🔑 HERDA A ORDEM DO MENSAL
                r.Cells["ORDEM"].Value = item[INDEX_ORDEM];

                string nome = item[INDEX_NOME]?.ToString() ?? "";
                r.Cells["HORARIO"].Value = item[INDEX_HORARIO]?.ToString();
                r.Cells["Nome"].Value = nome;

                r.Cells["ORDEM"].ReadOnly = true;
                r.Cells["HORARIO"].ReadOnly = true;
                r.Cells["Nome"].ReadOnly = true;

                r.Tag = gerarCartao ? "GERAR" : "IGNORAR";
                
                // PREENCHE SE TIVER NO BANCO
                if (assignments != null && assignments.ContainsKey(nome))
                {
                    var userPosts = assignments[nome];
                    // Percorre colunas de 3 em diante (horários)
                    for (int c = 3; c < dataGridView2.Columns.Count; c++)
                    {
                        string timeSlot = dataGridView2.Columns[c].HeaderText;
                        if (userPosts.ContainsKey(timeSlot))
                        {
                            r.Cells[c].Value = userPosts[timeSlot];
                        }
                    }
                }
            }

            // Linha de título (mantida igual)
            int idxT = dataGridView2.Rows.Add();
            var rowT = dataGridView2.Rows[idxT];

            rowT.Tag = "IGNORAR";

            for (int c = 0; c < dataGridView2.Columns.Count; c++)
                rowT.Cells[c] = new DataGridViewTextBoxCell();

            rowT.Cells["Nome"].Value = $"{titulo} ({lista.Count})";
            rowT.DefaultCellStyle.BackColor = System.Drawing.Color.Yellow;
            rowT.DefaultCellStyle.Font = new System.Drawing.Font(dataGridView2.Font, FontStyle.Bold);
            rowT.ReadOnly = true;
        }

        private void InserirListaSimples(string titulo, List<DataRow> lista)
        {
            if (lista.Count == 0) return;

            // 1. Extrai nomes e cria texto completo com título
            var nomes = lista.Select(l => l[INDEX_NOME]?.ToString() ?? "").Where(n => !string.IsNullOrWhiteSpace(n));
            string textoCompleto = $"{titulo}: {string.Join(", ", nomes)}";

            // 2. Adiciona UMA ÚNICA linha com tudo
            int idx = dataGridView2.Rows.Add();
            var row = dataGridView2.Rows[idx];
            
            row.Cells["Nome"].Value = textoCompleto;
            row.Tag = "MERGE"; // Tag especial para nossa pintura customizada
            row.ReadOnly = true;
            
            // Estilo visual da linha - AMARELO
            row.DefaultCellStyle.BackColor = System.Drawing.Color.Yellow;
            row.DefaultCellStyle.WrapMode = DataGridViewTriState.True; // Permite quebra de linha
            row.DefaultCellStyle.Font = new System.Drawing.Font(dataGridView2.Font, FontStyle.Bold);
            
            // Altura automática calculada pelo grid se AutoSizeRowsMode estiver ativo,
            // mas vamos forçar um mínimo para garantir.
            row.Height = 50; 
        }
        private void PintarPostos()
        {
            for (int r = 0; r < dataGridView2.Rows.Count; r++)
            {
                string nome = dataGridView2.Rows[r].Cells["Nome"].Value?.ToString()?.ToUpper() ?? ""; // Fix CS8602
                if (nome.Contains("OPERADORES") || nome.Contains("APRENDIZ") || nome.Contains("CFTV")) continue;

                for (int c = 3; c < dataGridView2.Columns.Count; c++)
                {
                    var cell = dataGridView2.Rows[r].Cells[c];
                    string? postoValor = dataGridView2.Rows[r].Cells[c].Value?.ToString()?.Trim()?.ToUpper(); // Fix CS8602
                                                                                                              // Só pula se for DarkGray E estiver vazio (mantém a indicação visual de "fechado")
                                                                                                              // Se estiver preenchido, deve ser processado para pegar a cor.
                    if (cell.Style.BackColor == System.Drawing.Color.DarkGray && string.IsNullOrEmpty(postoValor)) continue;

                    cell.Style.ForeColor = System.Drawing.Color.Black;
                    string? posto = cell.Value?.ToString()?.Trim()?.ToUpper(); // Fix CS8602

                    if (string.IsNullOrWhiteSpace(posto))
                    {
                        // SE a célula não for DarkGray, pintamos de branco.
                        // SE for DarkGray, deixamos DarkGray (respeitamos quem marcou como indisponível).
                        if (cell.Style.BackColor != System.Drawing.Color.White && cell.Style.BackColor != System.Drawing.Color.DarkGray)
                        {
                            cell.Style.BackColor = System.Drawing.Color.White;
                        }
                        continue;
                    }

                    switch (posto)
                    {
                        case "VALET": cell.Style.BackColor = System.Drawing.Color.FromArgb(255, 100, 100); break;
                        case "CAIXA": cell.Style.BackColor = System.Drawing.Color.FromArgb(255, 150, 150); break;
                        case "QRF": cell.Style.BackColor = System.Drawing.Color.FromArgb(0, 102, 204); cell.Style.ForeColor = System.Drawing.Color.White; break;
                        case "CIRC.": case "CIRC": cell.Style.BackColor = System.Drawing.Color.FromArgb(153, 204, 255); break;
                        case "REP|CIRC": case "REP|CIRC.": cell.Style.BackColor = System.Drawing.Color.FromArgb(0, 153, 0); cell.Style.ForeColor = System.Drawing.Color.White; break;
                        case "ECHO 21": case "ECHO21": cell.Style.BackColor = System.Drawing.Color.FromArgb(102, 204, 0); break;
                        case "CFTV": cell.Style.BackColor = System.Drawing.Color.FromArgb(0, 51, 153); cell.Style.ForeColor = System.Drawing.Color.White; break;
                        case "TREIN": case "TREIN.VALET": case "TREIN.CAIXA": cell.Style.BackColor = System.Drawing.Color.FromArgb(255, 255, 153); break;
                        case "APOIO": case "SUP": cell.Style.BackColor = System.Drawing.Color.LightGray; break;

                        default:
                            // PROTEÇÃO: Se a célula foi marcada como indisponível (DarkGray) manualmente ou por horário,
                            // não force Branco a menos que tenhamos certeza.
                            // Mas aqui estamos no 'default' switch de um posto desconhecido?
                            // Não, aqui é postos "normais" sem cor específica, ficam Brancos.
                            cell.Style.BackColor = System.Drawing.Color.White;
                            break;
                    }
                }
            }
        }
        private void PintarHorarios()
        {
            for (int r = 0; r < dataGridView2.Rows.Count; r++)
            {
                var row = dataGridView2.Rows[r];
                string nome = row.Cells["Nome"].Value?.ToString()?.ToUpper() ?? ""; // Fix CS8602

                // Pula cabeçalhos
                if (nome.Contains("OPERADORES") || nome.Contains("APRENDIZ") ||
                    nome.Contains("CFTV") || nome.Contains("SUPERVISÃO")) continue;

                string horarioFunc = row.Cells["HORARIO"].Value?.ToString()?.Trim() ?? ""; // Fix CS8602

                if (!TryParseHorario(horarioFunc, out TimeSpan iniFunc, out TimeSpan fimFunc))
                {
                    // Erro de leitura = Cinza Claro
                    for (int c = 3; c < dataGridView2.Columns.Count; c++)
                    {
                        row.Cells[c].Style.BackColor = System.Drawing.Color.LightGray;
                        row.Cells[c].ReadOnly = true;
                    }
                    continue;
                }

                // Ajuste para virada de noite (Funcionário)
                TimeSpan fimFuncAj = (fimFunc < iniFunc) ? fimFunc.Add(TimeSpan.FromHours(24)) : fimFunc;

                // --- LÓGICA DE COBERTURA ESTRITA ---
                // Agora verificamos se o funcionário cobre "o slot".
                // Para simplificar, verificamos se ele começa antes ou junto do INICIO da coluna
                // E termina depois ou junto do FIM da coluna.
                // Isso garante que 17:40 pinte até 17:40, mas não 18:00.

                for (int c = 3; c < dataGridView2.Columns.Count; c++)
                {
                    if (TryParseHorario(dataGridView2.Columns[c].HeaderText, out TimeSpan iniCol, out TimeSpan fimCol))
                    {
                        // Ajuste para virada de noite (Coluna)
                        TimeSpan fimColAj = (fimCol < iniCol) ? fimCol.Add(TimeSpan.FromHours(24)) : fimCol;

                        // Se a coluna vira a noite (ex: 23:00 - 00:00), e o func é 08:00-17:00, 
                        // precisamos garantir que estamos comparando na mesma "janela" de dias.
                        // Mas como normalizei fimColAj e fimFuncAj adicionando 24h se for menor que inicio,
                        // a comparação direta deve funcionar para a maioria dos casos simples de turno.

                        bool disponivel = (iniFunc <= iniCol) && (fimFuncAj >= fimColAj);

                        // Verificação extra para viradas de noite complexas (Projeção +24h)
                        if (!disponivel)
                        {
                            // Tenta projetar +24h o funcionário para ver se pega uma coluna da madrugada seguinte
                            disponivel = (iniFunc.Add(TimeSpan.FromHours(24)) <= iniCol) &&
                                         (fimFuncAj.Add(TimeSpan.FromHours(24)) >= fimColAj);
                        }

                        // APLICAÇÃO DAS CORES
                        var corAtual = row.Cells[c].Style.BackColor;

                        if (disponivel)
                        {
                            // SE TRABALHA -> PINTA DE BRANCO
                            // Aceita pintar se for Cinza Escuro, Cinza Claro ou Vazia
                            if (corAtual == System.Drawing.Color.DarkGray ||
                                corAtual == System.Drawing.Color.LightGray ||
                                corAtual.IsEmpty)
                            {
                                row.Cells[c].Style.BackColor = System.Drawing.Color.White;
                                row.Cells[c].ReadOnly = false;
                            }
                        }
                        else
                        {
                            // SE NÃO TRABALHA -> GARANTE O CINZA ESCURO
                            if (corAtual != System.Drawing.Color.DarkGray)
                            {
                                row.Cells[c].Style.BackColor = System.Drawing.Color.DarkGray;
                                row.Cells[c].Style.BackColor = System.Drawing.Color.DarkGray;
                                // ALTERAÇÃO: Permitir edição mesmo fora do horário (cinza)
                                row.Cells[c].ReadOnly = false;
                            }
                        }
                    }
                }
            }
        }
        private void button1_Click(object? sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel|*.xlsx" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Cursor.Current = Cursors.WaitCursor;
                    _tabelaMensal = LerExcel(ofd.FileName);
                    
                    // Salva no Banco de Dados
                    DatabaseService.SaveMonthlyData(_tabelaMensal);
                    
                    if (dataGridView1 != null) 
                    { 
                        dataGridView1.DataSource = null;
                        dataGridView1.Columns.Clear();
                        dataGridView1.DataSource = _tabelaMensal; 
                        ConfigurarGridMensal(); 
                    }
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show("Importado com sucesso! Veja a aba EscalaMensal.");
                    ProcessarEscalaDoDia();
                }
                catch (Exception ex) { Cursor.Current = Cursors.Default; MessageBox.Show("Erro: " + ex.Message); }
            }
        }
        private DataTable LerExcel(string caminho)
        {
            DataTable dt = new DataTable();
            for (int i = 1; i <= MAX_COLS; i++) dt.Columns.Add($"C{i}");
            using (var wb = new XLWorkbook(caminho))
            {
                var ws = wb.Worksheets.First();
                foreach (var row in ws.RowsUsed())
                {
                    var nova = dt.NewRow();
                    for (int c = 1; c <= MAX_COLS; c++) nova[c - 1] = row.Cell(c).GetValue<string>();
                    dt.Rows.Add(nova);
                }
            }
            return dt;
        }
        private void ConfigurarGridMensal()
        {
            if (dataGridView1.DataSource == null) return;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            if (dataGridView1.Columns.Count > 0) dataGridView1.Columns[0].Visible = false;
            if (dataGridView1.Columns.Count > INDEX_FUNCAO) { dataGridView1.Columns[INDEX_FUNCAO].HeaderText = "FUNÇÃO"; dataGridView1.Columns[INDEX_FUNCAO].Width = 60; }

            if (dataGridView1.Columns.Count > INDEX_HORARIO) { dataGridView1.Columns[INDEX_HORARIO].HeaderText = "HORÁRIO"; dataGridView1.Columns[INDEX_HORARIO].Width = 80; }
            if (dataGridView1.Columns.Count > INDEX_ORDEM) { dataGridView1.Columns[INDEX_ORDEM].HeaderText = "ORDEM"; dataGridView1.Columns[INDEX_ORDEM].Width = 50; }
            if (dataGridView1.Columns.Count > INDEX_NOME) { dataGridView1.Columns[INDEX_NOME].HeaderText = "NOME"; dataGridView1.Columns[INDEX_NOME].Width = 120; dataGridView1.Columns[INDEX_NOME].Frozen = true; }
            for (int i = INDEX_DIA_INICIO; i < dataGridView1.Columns.Count; i++) { dataGridView1.Columns[i].HeaderText = $"{i - INDEX_DIA_INICIO + 1}"; dataGridView1.Columns[i].Width = 35; }
        }
        private void ConfigurarGridEscalaDiaria()
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();

            // Colunas fixas
            dataGridView2.Columns.Add("ORDEM", "Nº");
            dataGridView2.Columns.Add("HORARIO", "HORÁRIO");
            dataGridView2.Columns.Add("Nome", "Nome");

            string[] horarios =
            {
        "08:00 x 08:40", "08:41 x 09:40", "09:41 x 10:40",
        "10:41 x 11:40", "11:41 x 12:40", "12:41 x 13:40",
        "13:41 x 14:40", "14:41 x 15:40", "15:41 x 16:40",
        "16:41 x 17:40", "17:41 x 18:40", "18:41 x 19:40",
        "19:41 x 20:40", "20:41 x 21:40", "21:41 x 22:40",
        "22:41 x 23:40", "23:41 x 00:40", "00:41 x 01:40"
    };

            var postos = new List<string>
    {
        "", "CAIXA", "VALET", "QRF", "CIRC.", "REP|CIRC",
        "CS1", "CS2", "CS3", "SUP", "APOIO", "TREIN", "CFTV"
    };

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

                dataGridView2.Columns.Add(col);
                col.CellTemplate.Style.BackColor = System.Drawing.Color.DarkGray;
            }

            // Congelamento
            dataGridView2.Columns["ORDEM"].Frozen = true;
            dataGridView2.Columns["HORARIO"].Frozen = true;
            dataGridView2.Columns["Nome"].Frozen = true;

            // Visual
            dataGridView2.Columns["ORDEM"].Width = 40;
            dataGridView2.Columns["HORARIO"].Width = 80;
            dataGridView2.Columns["Nome"].Width = 110;

            dataGridView2.Columns["ORDEM"].DefaultCellStyle.BackColor = System.Drawing.Color.White;
            dataGridView2.Columns["HORARIO"].DefaultCellStyle.BackColor = System.Drawing.Color.White;
            dataGridView2.Columns["Nome"].DefaultCellStyle.BackColor = System.Drawing.Color.White;
        }
        private void CalcularTotais()
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                string textoLinha = dataGridView2.Rows[i].Cells[2].Value?.ToString()?.ToUpper() ?? ""; // Fix CS8602
                if (textoLinha.Contains("OPERADORES") || textoLinha.Contains("APRENDIZ") || textoLinha.Contains("CFTV"))
                {
                    for (int c = 3; c < dataGridView2.Columns.Count; c++)
                    {
                        int count = 0;
                        for (int k = i - 1; k >= 0; k--)
                        {
                            string tAnt = dataGridView2.Rows[k].Cells[2].Value?.ToString()?.ToUpper() ?? ""; // Fix CS8602
                            if (tAnt.Contains("OPERADORES") || tAnt.Contains("APRENDIZ") || tAnt.Contains("CFTV")) break;
                            if (!string.IsNullOrWhiteSpace(dataGridView2.Rows[k].Cells[c].Value?.ToString())) count++; // Value?.ToString() is safe here as IsNullOrWhiteSpace handles null
                        }
                        dataGridView2.Rows[i].Cells[c].Value = count > 0 ? count.ToString() : "";
                    }
                }
            }
        }
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
        private void DataGridView2_CellEnter(object? sender, DataGridViewCellEventArgs e) { if (e.ColumnIndex > 1) SendKeys.Send("{F4}"); }
        private void DataGridView2_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                if (dataGridView2.SelectedCells.Count > 0)
                {
                    foreach (DataGridViewCell cell in dataGridView2.SelectedCells)
                    {
                        // Pula colunas de Nome e Horário (Indices 0 e 1)
                        if (cell.ColumnIndex <= 1) continue;

                        cell.Value = ""; // Limpa conteúdo
                        cell.Style.BackColor = System.Drawing.Color.DarkGray; // Reseta cor para "Indisponível"
                    }
                    e.Handled = true;  // Impede comportamento padrão

                    // Força atualização visual e itinerários
                    AtualizarItinerarios();
                    // PintarHorarios(); // REMOVIDO: Isso reverte a cor para Branco se o func estiver trabalhando!
                }


            }

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
                    
                    // Atualiza o clima para o dia selecionado
                    AtualizarClimaParaDia(_diaSelecionado);
                }
            }
        }
        private bool EhFolga(string? codigo) { if (string.IsNullOrWhiteSpace(codigo)) return false; return new[] { "X", "FOLGA", "FERIAS", "FÉRIAS" }.Contains(codigo.Trim().ToUpper()); }
        private bool TryParseHorario(string? t, out TimeSpan i, out TimeSpan f) { i = TimeSpan.Zero; f = TimeSpan.Zero; if (t == null) return false; var p = t.Split(new[] { 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries); if (p.Length == 2 && TimeSpan.TryParse(p[0].Trim(), out i) && TimeSpan.TryParse(p[1].Trim(), out f)) return true; return false; }
        private List<DataRow> OrdenarPorHorario(List<DataRow> l)
        {
            return l.OrderBy(r =>
            {
                // Ordenação Primária: ORDEM (Coluna 3)
                // Se não for número ou estiver vazio, joga pro final (int.MaxValue)
                string ordemStr = r[INDEX_ORDEM]?.ToString() ?? "";
                if (int.TryParse(ordemStr, out int ordem) && ordem > 0) return ordem;
                return int.MaxValue;
            })
            .ThenBy(r => r[INDEX_HORARIO].ToString()) // Ordenação Secundária: HORÁRIO
            .ToList();
        }
        private void btnRecarregarBanco_Click(object? sender, EventArgs e)
        {
            //  Confirma com o usuário
            var result = MessageBox.Show(
                "Deseja limpar todas as atribuições de postos da escala diária?", 
                "Limpar Escala", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                // Limpa todas as atribuições do banco
                DatabaseService.ClearAllAssignments();
                
                // Reprocessa o dia atual para exibir a grade limpa
                ProcessarEscalaDoDia();
                
                MessageBox.Show("Escala diária limpa com sucesso!", "Concluído", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        // 2. O Método que faz a mágica
        private async Task AtualizarClimaAutomatico()
        {
            // URL correta
            string url = "https://api.hgbrasil.com/weather?woeid=455822&key=development";

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var response = await client.GetStringAsync(url);
                    var json = JObject.Parse(response);
                    
                    // Armazena a previsão completa para uso posterior
                    _previsaoCompleta = json;
                    
                    // Atualiza o label com o clima do dia selecionado
                    AtualizarClimaParaDia(_diaSelecionado);
                }
            }
            catch
            {
                lblClima.Text = "Clima offline";
            }
        }
        
        private void AtualizarClimaParaDia(int dia)
        {
            if (_previsaoCompleta == null)
            {
                lblClima.Text = "Carregando clima...";
                return;
            }
            
            try
            {
                var dados = _previsaoCompleta["results"];
                string cidade = dados["city"]?.ToString() ?? "Curitiba";
                
                // Pega a data atual
                DateTime hoje = DateTime.Now;
                DateTime dataAlvo = new DateTime(hoje.Year, hoje.Month, dia);
                
                // Se o dia já passou neste mês, assume próximo mês
                if (dataAlvo < hoje.Date)
                {
                    dataAlvo = dataAlvo.AddMonths(1);
                }
                
                // Calcula diferença de dias
                int diasDeDiferenca = (dataAlvo - hoje.Date).Days;
                
                string temp, desc, periodo = "dia";
                string icone;
                
                // Se for hoje (dia 0), usa dados atuais
                if (diasDeDiferenca == 0)
                {
                    temp = dados["temp"].ToString();
                    desc = dados["description"].ToString();
                    periodo = dados["currently"].ToString();
                    icone = ObterIconeVisual(desc, periodo);
                    
                    string diaSemana = DateTime.Now.ToString("dddd", new CultureInfo("pt-BR"));
                    diaSemana = char.ToUpper(diaSemana[0]) + diaSemana.Substring(1);
                    lblClima.Text = $"{diaSemana} {cidade} {icone} {temp}°C - {desc}";
                }
                // Senão, busca na previsão estendida
                else if (diasDeDiferenca > 0 && diasDeDiferenca < 10) // API retorna até 10 dias
                {
                    var forecast = dados["forecast"] as JArray;
                    if (forecast != null && diasDeDiferenca < forecast.Count)
                    {
                        var diaPrevisao = forecast[diasDeDiferenca];
                        string weekday = diaPrevisao["weekday"]?.ToString() ?? "";
                        
                        // Traduz dia da semana
                        var traducao = new Dictionary<string, string>
                        {
                            {"Sun", "Domingo"}, {"Mon", "Segunda"}, {"Tue", "Terça"}, 
                            {"Wed", "Quarta"}, {"Thu", "Quinta"}, {"Fri", "Sexta"}, {"Sat", "Sábado"}
                        };
                        string diaSemana = traducao.ContainsKey(weekday) ? traducao[weekday] : weekday;
                        
                        int max = int.Parse(diaPrevisao["max"]?.ToString() ?? "0");
                        int min = int.Parse(diaPrevisao["min"]?.ToString() ?? "0");
                        desc = diaPrevisao["description"]?.ToString() ?? "";
                        
                        icone = ObterIconeVisual(desc, "dia");
                        
                        lblClima.Text = $"{diaSemana} (Dia {dia}) {cidade} {icone} {max}°C/{min}°C - {desc}";
                    }
                    else
                    {
                        lblClima.Text = $"Previsão indisponível para o dia {dia}";
                    }
                }
                else
                {
                    lblClima.Text = $"Previsão muito distante (dia {dia})";
                }
                
                // Lógica de cores baseada na temperatura
                if (lblClima.Text.Contains("°C"))
                {
                    var match = System.Text.RegularExpressions.Regex.Match(lblClima.Text, @"(\d+)°C");
                    if (match.Success)
                    {
                        int tempInt = int.Parse(match.Groups[1].Value);
                        if (tempInt < 15) lblClima.ForeColor = Color.Blue;
                        else if (tempInt > 28) lblClima.ForeColor = Color.OrangeRed;
                        else lblClima.ForeColor = Color.Black;
                    }
                }
            }
            catch
            {
                lblClima.Text = "Erro ao processar clima";
            }
        }

        // --- NÃO ESQUEÇA DE COLAR ESSE MÉTODO NO FINAL DA CLASSE ---
        private string ObterIconeVisual(string descricao, string periodo)
        {
            if (string.IsNullOrEmpty(descricao)) return "⛅";

            descricao = descricao.ToLower();

            if (descricao.Contains("chuva")) return "🌧️";
            if (descricao.Contains("tempestade")) return "⛈️";
            if (descricao.Contains("garoa")) return "🌦️";
            if (descricao.Contains("nublado")) return "☁️";
            if (descricao.Contains("claro") || descricao.Contains("limpo"))
                return periodo == "noite" ? "🌙" : "☀️";
            return "🌡️";
        }

        private void DataGridView2_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
        {
            // Ignora headers ou colunas fixas (0=ORDEM, 1=HORARIO, 2=Nome)
            if (e.RowIndex < 0 || e.ColumnIndex < 3) return;

            var row = dataGridView2.Rows[e.RowIndex];
            // Se for linha de título ou ignorada
            if (row.Tag != null && row.Tag.ToString() == "IGNORAR") 
            {
                 return; 
            }
            
            string nome = row.Cells["Nome"].Value?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(nome) || nome.Contains("(")) return; // Ignora linhas de titulo

            string timeSlot = dataGridView2.Columns[e.ColumnIndex].HeaderText;
            string? valor = row.Cells[e.ColumnIndex].Value?.ToString();

            // Salva no Banco apenas se tiver nome
            DatabaseService.SaveAssignment(_diaSelecionado, nome, timeSlot, valor ?? "");
        }

        private void DataGridView2_CellPainting(object? sender, DataGridViewCellPaintingEventArgs e)
        {
            // Verifica se é uma linha válida
            if (e.RowIndex < 0) return;

            // Verifica se a linha tem a TAG "MERGE"
            var row = dataGridView2.Rows[e.RowIndex];
            if (row.Tag == null || row.Tag.ToString() != "MERGE") return;

            // Se for "MERGE", queremos desenhar o texto da célula "Nome"
            // Atravessando todas as colunas visíveis.
            
            // 1. Pinta o fundo padrão da célula
            e.PaintBackground(e.CellBounds, true);

            // 2. Só desenhamos o texto quando estivermos na primeira coluna visível ou na coluna "Nome"
            // Para simplificar, vamos desenhar "sobre" tudo, mas o DataGridView chama isso por célula.
            // O truque: Calcular o retângulo total da linha e desenhar o texto apenas quando estivermos na célula 'ORDEM' (que é frozen e a primeira)
            // Ou melhor: Desenhar em todas? Não, vai sobrepor.
            // Vamos desenhar APENAS na coluna 'Nome' mas com clip estendido? Não funciona bem com Frozen.
            
            // TRUQUE DO MERGE VISUAL:
            // Vamos desenhar o texto apenas quando o evento for disparado para a coluna 'ORDEM' (que é fixa esquerda)
            // Mas vamos estender o retângulo de desenho até o fim do grid.
            
            if (e.ColumnIndex == INDEX_ORDEM) // Assumindo que ORDEM (3) é a primeira visível/frozen útil ou 0 se for a primeira.
            {
               // Na verdade, 0, 1 e 2 são fixas. 
               // Vamos usar a coluna 0 (ORDEM) para disparar o desenho.
            }
            // MUDANÇA DE PLANO: O jeito mais fácil é:
            // Deixar o grid pintar o fundo de todas as células da linha (já feito no PaintBackground).
            // Cancelar a pintura do conteúdo (e.Handled = true) para todas as células dessa linha.
            // E desenhar o textozão APENAS na célula 'Nome' mas com bounds enganosos?
            // Não, o clip vai cortar.
            
            // MELHOR:
            // Usar TextFormatFlags.NoClipping? Arriscado.
            
            // VAMOS FAZER O SEGUINTE:
            // As células dessa linha terão valor vazio, EXCETO a célula onde guardaremos o texto (vamos usar a primeira visivel, ex column 0).
            // O texto está em row.Cells["Nome"].Value.
            
            e.Handled = true; // Nós cuidamos de tudo (fundo já foi pintado acima)
            e.PaintContent(e.CellBounds); // Pinta bordas se necessário? Não, queremos texto limpo.
            
            // Só desenhar o texto 1 vez por linha?
            // Se desenharmos em cada célula partes do texto, fica ruim.
            
            // Solução robusta para "Merge Row":
            // Só desenhamos o texto quando estivermos na coluna 'Nome'.
            // E definimos o retângulo do Graphics para ocupar a largura do Grid.
            // Mas o ClipRegion vai impedir.
            
            // Ok, vamos simplificar. O usuário aceita "uma linha apenas separando por virgula".
            // Se eu colocar o texto na coluna "Nome" (que tem largura) e deixar as outras vazias,
            // fica visualmente "quebrado" pelas linhas verticais da grade.
            
            // Vamos "apagar" as linhas verticais dessa linha pintando por cima.
            // Pinta o fundo de novo para garantir (sem bordas da grade)
            using (Brush backColorBrush = new SolidBrush(e.CellStyle.BackColor))
            {
                e.Graphics.FillRectangle(backColorBrush, e.CellBounds);
            }
            
            // Desenha a linha inferior apenas (borda da linha)
            if (e.RowIndex < dataGridView2.Rows.Count - 1)
            {
                e.Graphics.DrawLine(Pens.Black, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);
            }

            // Agora, se formos a célula que CONTÉM o texto (Nome), desenhamos o texto.
            // Mas queremos que ele vazamento para os lados?
            // O DataGridView recorta (clips) o desenho na borda da célula.
            
            // HACK: DataGridView não clipa se usarmos TextRenderer com flags correta?
            // Não, o Graphics object vem clipado.
            
            // ENTAO: O unico jeito de fazer "Merge" de verdade é desenhar o texto parte por parte? Impossível alinhar.
            
            // ALTERNATIVA ACEITÁVEL:
            // O código abaixo desenha o texto corrido, mas como o Graphics está clipado, ele só vai aparecer dentro da célula da coluna atual se tivermos sorte?
            // Não.
            
            // VAMOS TENTAR ISTO:
            // Vamos desenhar o texto SOMENTE se for a coluna 'Nome'.
            // Mas vamos alterar o Region do Graphics para permitir desenhar fora?
            // e.Graphics.SetClip(e.CellBounds); // Isso só restringe.
            
            // Vamos TENTAR desenhar o texto letra por letra? Não.
            
            // OK, vamos voltar ao básico:
            // O usuário quer "mesclar". 
            // Se eu não consigo remover o Clip, eu tenho que replicar o texto? Não.
            
            // TENTATIVA FINAL DE MERGE VISUAL:
            // Vamos desenhar o texto na coluna 'Nome', mas ela tem largura fixa de 110.
            // Vamos AUMENTAR a largura da coluna 'Nome' nessa linha? Não dá, largura é por coluna.
            
            // ÚNICA SAÍDA NO WINFORMS PADRÃO:
            // Desenhar o texto no evento `RowPostPaint`. Esse evento permite desenhar SOBRE todas as células da linha sem clip individual.
         }

        private void DataGridView2_RowPostPaint(object? sender, DataGridViewRowPostPaintEventArgs e)
        {
             var row = dataGridView2.Rows[e.RowIndex];
             if (row.Tag == null || row.Tag.ToString() != "MERGE") return;
             
             string texto = row.Cells["Nome"].Value?.ToString() ?? "";
             
             // RESPEITA as 3 primeiras colunas (ORDEM, HORARIO, Nome)
             // Começa a desenhar a partir da coluna 3 (primeira coluna de horário)
             var rectColuna3 = dataGridView2.GetCellDisplayRectangle(3, e.RowIndex, true);
             
             // Se coluna não encontrada ou oculta, fallback para inicio
             int xStart = (rectColuna3.Width > 0) ? rectColuna3.X : e.RowBounds.Left;
             
             // Retângulo de desenho: Do inicio da coluna 3 até o fim da row
             int width = e.RowBounds.Right - xStart;
             
             Rectangle r = new Rectangle(xStart, e.RowBounds.Top, width, e.RowBounds.Height);
             
             // Pinta texto
             TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak;
             
             // Padding
             r.X += 5;
             r.Width -= 10;

             TextRenderer.DrawText(e.Graphics, texto, e.InheritedRowStyle.Font, r, e.InheritedRowStyle.ForeColor, flags);
        }


    }
}

// =========================================================
// Extension Methods (Must be top-level static class)
// =========================================================
public static class ExtensionMethods
{
    public static void DoubleBuffered(this DataGridView dgv, bool setting)
    {
        Type dgvType = dgv.GetType();
        System.Reflection.PropertyInfo? pi = dgvType.GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        if (pi != null)
            pi.SetValue(dgv, setting, null);
    }
}

// =========================================================
// Model Classes
// =========================================================
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
