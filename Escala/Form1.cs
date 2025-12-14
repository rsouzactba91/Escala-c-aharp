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

            }
        }

        private void Form1_Load(object? sender, EventArgs e)
        {
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

            // 3. Loop de Classificação
            foreach (DataRow linha in _tabelaMensal.Rows)
            {
                string nome = linha[INDEX_NOME]?.ToString() ?? "";
                string horario = linha[INDEX_HORARIO]?.ToString() ?? "";
                string funcao = (INDEX_FUNCAO < _tabelaMensal.Columns.Count) ? (linha[INDEX_FUNCAO]?.ToString()?.ToUpper() ?? "") : ""; // Fix CS8602
                string nomeUpper = nome.ToUpper();

                if (string.IsNullOrWhiteSpace(nome) || nomeUpper.Contains("NOME")) continue;
                if (!horario.Contains(":")) continue;

                string? statusNoDia = linha[indiceColunaDia]?.ToString(); // Fix CS8600
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

            //  InserirBloco("SUPERVISÃO", OrdenarPorHorario(listaSUP), false); // false = Sem Itinerário
            InserirBloco("OPERADORES", OrdenarPorHorario(listaOP), true);   // true = Com Itinerário
            InserirBloco("APRENDIZ", OrdenarPorHorario(listaJV), true);     // true = Com Itinerário
            InserirBloco("CFTV", OrdenarPorHorario(listaCFTV), false);      // false = Sem Itinerário

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
        private void InserirBloco(string titulo, List<DataRow> lista, bool gerarCartao)
        {
            if (lista.Count == 0) return;

            foreach (var item in lista)
            {
                int idx = dataGridView2.Rows.Add();
                var r = dataGridView2.Rows[idx];

                // 🔑 HERDA A ORDEM DO MENSAL
                r.Cells["ORDEM"].Value = item[INDEX_ORDEM];

                r.Cells["HORARIO"].Value = item[INDEX_HORARIO]?.ToString();
                r.Cells["Nome"].Value = item[INDEX_NOME]?.ToString();

                r.Cells["ORDEM"].ReadOnly = true;
                r.Cells["HORARIO"].ReadOnly = true;
                r.Cells["Nome"].ReadOnly = true;

                r.Tag = gerarCartao ? "GERAR" : "IGNORAR";
            }

            // Linha de título
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
                    if (dataGridView1 != null) { dataGridView1.DataSource = _tabelaMensal; ConfigurarGridMensal(); }
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
            AjustarHorariosMensal(-1);
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

                    var dados = json["results"];

                    // --- CORREÇÃO AQUI ---
                    // O campo no JSON costuma ser "city"
                    string dia = DateTime.Now.ToString("dddd", new CultureInfo("pt-BR"));
                    // Transforma "domingo" em "Domingo"
                    dia = char.ToUpper(dia[0]) + dia.Substring(1);

                    string cidade = dados["city"]?.ToString() ?? "Curitiba";
                    string temp = dados["temp"].ToString();
                    string desc = dados["description"].ToString();
                    string periodo = dados["currently"].ToString();

                    // Chama o método auxiliar para pegar o emoji
                    string icone = ObterIconeVisual(desc, periodo);

                    // Atualiza o Label
                    lblClima.Text = $"{dia} {cidade} {icone} {temp}°C - {desc}";

                    // Lógica de cores (Perfeita)
                    if (int.Parse(temp) < 15) lblClima.ForeColor = Color.Blue;
                    else if (int.Parse(temp) > 28) lblClima.ForeColor = Color.OrangeRed;
                    else lblClima.ForeColor = Color.Black;
                }
            }
            catch
            {
                lblClima.Text = "Clima offline";
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

            return "⛅";
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
