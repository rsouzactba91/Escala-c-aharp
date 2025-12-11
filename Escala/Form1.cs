using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing.Printing; // Necessário para impressão

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
        private const int INDEX_NOME = 4;
        private const int INDEX_DIA_INICIO = 5;

        private DataTable _tabelaMensal;
        private int _diaSelecionado = 1;

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
                if (btnImprimir != null) btnImprimir.Click += btnImprimir_Click;

            }
        }

        private void Form1_Load(object sender, EventArgs e)
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
                // ---------------------
            }
        }

        // =========================================================
        // 2. LÓGICA PRINCIPAL (PROCESSAR O DIA)
        // =========================================================
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
                string funcao = (INDEX_FUNCAO < _tabelaMensal.Columns.Count) ? linha[INDEX_FUNCAO].ToString().ToUpper() : "";
                string nomeUpper = nome.ToUpper();

                if (string.IsNullOrWhiteSpace(nome) || nomeUpper.Contains("NOME")) continue;
                if (!horario.Contains(":")) continue;

                string statusNoDia = linha[indiceColunaDia]?.ToString();
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


        // =========================================================
        // 7. LÓGICA DE IMPRESSÃO (GRID + ITINERÁRIOS)
        // =========================================================

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(ImprimirConteudo);

            // Abre a visualização antes de imprimir
            PrintPreviewDialog ppd = new PrintPreviewDialog();
            ppd.Document = pd;
            ppd.WindowState = FormWindowState.Maximized;
            ppd.ShowDialog();
        }

        private void ImprimirConteudo(object sender, PrintPageEventArgs e)
        {
            // 1. CAPTURA A IMAGEM DO GRID (ESCALA)
            // Redimensiona temporariamente para caber tudo se tiver rolagem
            int heightOriginalGrid = dataGridView2.Height;
            dataGridView2.Height = dataGridView2.RowCount * dataGridView2.RowTemplate.Height + dataGridView2.ColumnHeadersHeight;

            Bitmap bmpGrid = new Bitmap(dataGridView2.Width, dataGridView2.Height);
            dataGridView2.DrawToBitmap(bmpGrid, new Rectangle(0, 0, dataGridView2.Width, dataGridView2.Height));

            // Restaura tamanho original
            dataGridView2.Height = heightOriginalGrid;

            // 2. CAPTURA A IMAGEM DOS ITINERÁRIOS (CARTÕES)
            // Precisamos calcular a altura total dos cartões dentro do painel
            int alturaPainel = 0;
            foreach (System.Windows.Forms.Control c in flowLayoutPanel1.Controls)
                alturaPainel = Math.Max(alturaPainel, c.Bottom);
            alturaPainel += 20; // Margem

            // Cria bitmap do tamanho total do conteúdo
            Bitmap bmpItinerario = new Bitmap(flowLayoutPanel1.Width, Math.Max(alturaPainel, 100)); // Minimo 100
            flowLayoutPanel1.DrawToBitmap(bmpItinerario, new Rectangle(0, 0, flowLayoutPanel1.Width, Math.Max(alturaPainel, flowLayoutPanel1.Height)));

            // 3. DESENHA NA FOLHA DE PAPEL
            float y = e.MarginBounds.Top;
            float x = e.MarginBounds.Left;
            float larguraUtil = e.MarginBounds.Width;

            // Cabeçalho
            System.Drawing.Font fonteTitulo = new System.Drawing.Font("Arial", 16, FontStyle.Bold);
            e.Graphics.DrawString($"Escala do Dia {_diaSelecionado}", fonteTitulo, Brushes.Black, x, y);
            y += 40;

            // Desenha o Grid (Ajustando a largura para caber na folha)
            float ratioGrid = (float)bmpGrid.Width / (float)bmpGrid.Height;
            float alturaGridNaFolha = larguraUtil / ratioGrid;

            // Se o grid for muito alto, limita para não sumir da página
            if (alturaGridNaFolha > 400) alturaGridNaFolha = 400;

            e.Graphics.DrawImage(bmpGrid, x, y, larguraUtil, alturaGridNaFolha);
            y += alturaGridNaFolha + 30; // Pula linha

            // Título Itinerários
            e.Graphics.DrawString("Itinerários / Cartões", fonteTitulo, Brushes.Black, x, y);
            y += 30;

            // Desenha o Painel de Cartões
            // Se não couber na página, o código simples vai cortar (para MVP serve)
            // Se precisar quebrar página é bem mais complexo.
            float ratioItin = (float)bmpItinerario.Width / (float)bmpItinerario.Height;
            float alturaItinNaFolha = larguraUtil / ratioItin;

            e.Graphics.DrawImage(bmpItinerario, x, y, larguraUtil, alturaItinNaFolha);
        }
        // =========================================================
        // 3. VISUALIZAÇÃO (ITINERÁRIOS NA ABA 3)
        // =========================================================
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
            // Dimensões
            int larguraTotal = 200;
            int alturaLinha = 20;
            int alturaCabecalho = 20;

            // Fontes Estilizadas (Impact e Arial Narrow)
            System.Drawing.Font fonteCabecalho = new System.Drawing.Font("Impact", 12, FontStyle.Regular);
            System.Drawing.Font fonteHora = new System.Drawing.Font("Arial Narrow", 12, FontStyle.Bold);
            System.Drawing.Font fontePosto = new System.Drawing.Font("Arial", 12, FontStyle.Bold);

            // Painel Principal (O Cartão)
            Panel pnl = new Panel();
            pnl.Width = larguraTotal;
            pnl.Height = alturaCabecalho + (dados.Itens.Count * alturaLinha) + 4;
            pnl.BackColor = System.Drawing.Color.White;
            pnl.Margin = new Padding(10);

            // Evento Paint para desenhar a borda grossa preta
            pnl.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, pnl.ClientRectangle,
                    System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid,
                    System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid,
                    System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid,
                    System.Drawing.Color.Black, 2, ButtonBorderStyle.Solid);
            };

            // Cabeçalho (Data | Nome)
            Label lblHeader = new Label();
            lblHeader.Text = $"{_diaSelecionado:D2}/10 | {dados.Nome}"; // Ajuste o mês se necessário
            lblHeader.Font = fonteCabecalho;
            lblHeader.TextAlign = ContentAlignment.MiddleCenter;
            lblHeader.Dock = DockStyle.Top;
            lblHeader.Height = alturaCabecalho;

            // Linha preta abaixo do cabeçalho
            Panel linhaDivisoria = new Panel { Height = 3, BackColor = System.Drawing.Color.Black, Dock = DockStyle.Top };

            // Loop para criar as linhas de horário
            int yAtual = alturaCabecalho + 3;

            foreach (var item in dados.Itens)
            {
                Panel pnlLinha = new Panel { Location = new Point(2, yAtual), Size = new Size(larguraTotal - 4, alturaLinha), BackColor = System.Drawing.Color.White };

                // Horário
                Label lblHora = new Label { Text = item.Horario, Font = fonteHora, ForeColor = System.Drawing.Color.Black, TextAlign = ContentAlignment.MiddleCenter, Size = new Size(130, alturaLinha), Location = new Point(0, 0) };

                // Divisória Vertical
                Panel linhaVertical = new Panel { BackColor = System.Drawing.Color.Black, Width = 2, Location = new Point(130, 0), Height = alturaLinha };

                // Posto Colorido
                Label lblPosto = new Label { Text = item.Posto, Font = fontePosto, ForeColor = item.CorTexto, BackColor = item.CorFundo, TextAlign = ContentAlignment.MiddleCenter, Location = new Point(132, 0), Size = new Size(pnlLinha.Width - 132, alturaLinha) };

                // Linha Horizontal Abaixo
                Panel linhaHorizontal = new Panel { BackColor = System.Drawing.Color.Black, Height = 2, Width = larguraTotal, Location = new Point(0, yAtual + alturaLinha) };

                pnlLinha.Controls.Add(lblHora);
                pnlLinha.Controls.Add(linhaVertical);
                pnlLinha.Controls.Add(lblPosto);

                pnl.Controls.Add(pnlLinha);
                pnl.Controls.Add(linhaHorizontal);

                yAtual += alturaLinha + 2;
            }

            pnl.Controls.Add(linhaDivisoria);
            pnl.Controls.Add(lblHeader);

            return pnl;
        }

        private List<CartaoFuncionario> GerarDadosDosItinerarios()
        {
            var lista = new List<CartaoFuncionario>();

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                // 1. VERIFICAÇÃO INFALÍVEL PELA TAG
                if (row.Tag != null && row.Tag.ToString() == "IGNORAR") continue;

                // Se por acaso a Tag for nula (cabeçalhos antigos), verifica o nome só por segurança
                string nome = row.Cells["Nome"].Value?.ToString();
                if (string.IsNullOrWhiteSpace(nome)) continue;
                if (nome.Contains("SUPERVISÃO") || nome.Contains("CFTV") || nome.Contains("OPERADORES")) continue;

                // --- DAQUI PRA BAIXO TUDO IGUAL ---
                var cartao = new CartaoFuncionario { Nome = nome };
                bool temPosto = false;

                for (int c = 2; c < dataGridView2.Columns.Count; c++)
                {
                    var cell = row.Cells[c];
                    if (cell.Style.BackColor == System.Drawing.Color.DarkGray ||
                        cell.Style.BackColor == System.Drawing.Color.LightGray) continue;

                    string posto = cell.Value?.ToString();

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

        // =========================================================
        // 4. MÉTODOS AUXILIARES (Lógica de Grid e Excel)
        // =========================================================
        private void PreencherPostosAutomaticos(string tipoFuncionario, List<DataRow> listaDados, string postoPadrao)
        {
            foreach (DataRow dados in listaDados)
            {
                string nomeFuncionario = dados[INDEX_NOME].ToString();
                string horarioFunc = dados[INDEX_HORARIO].ToString();

                if (!TryParseHorario(horarioFunc, out TimeSpan iniFunc, out TimeSpan fimFunc)) continue;

                // --- ARREDONDAMENTO REMOVIDO ---
                // if (fimFunc.Minutes == 40) ... (APAGADO)

                TimeSpan fimFuncAj = (fimFunc < iniFunc) ? fimFunc.Add(TimeSpan.FromHours(24)) : fimFunc;

                foreach (DataGridViewRow rowGrid in dataGridView2.Rows)
                {
                    if (rowGrid.Cells["Nome"].Value?.ToString() == nomeFuncionario)
                    {
                        for (int c = 2; c < dataGridView2.Columns.Count; c++)
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

                r.Cells["HORARIO"].Value = item[INDEX_HORARIO]?.ToString();
                r.Cells["Nome"].Value = item[INDEX_NOME]?.ToString();

                r.Cells["HORARIO"].ReadOnly = true;
                r.Cells["Nome"].ReadOnly = true;

                // --- O SEGREDO ESTÁ AQUI ---
                // Marcamos a linha: "GERAR" se for true, "IGNORAR" se for false
                r.Tag = gerarCartao ? "GERAR" : "IGNORAR";
            }

            // Linha de Título (Amarela)
            int idxT = dataGridView2.Rows.Add();
            var rowT = dataGridView2.Rows[idxT];

            // O título sempre deve ser ignorado na geração de cartões
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
                string nome = dataGridView2.Rows[r].Cells["Nome"].Value?.ToString().ToUpper() ?? "";
                if (nome.Contains("OPERADORES") || nome.Contains("APRENDIZ") || nome.Contains("CFTV")) continue;

                for (int c = 2; c < dataGridView2.Columns.Count; c++)
                {
                    var cell = dataGridView2.Rows[r].Cells[c];
                    if (cell.Style.BackColor == System.Drawing.Color.DarkGray) continue;

                    cell.Style.ForeColor = System.Drawing.Color.Black;
                    string posto = cell.Value?.ToString().Trim().ToUpper();

                    if (string.IsNullOrWhiteSpace(posto))
                    {
                        if (cell.Style.BackColor != System.Drawing.Color.White) cell.Style.BackColor = System.Drawing.Color.White;
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
                        default: cell.Style.BackColor = System.Drawing.Color.White; break;
                    }
                }
            }
        }

        private void PintarHorarios()
        {
            for (int r = 0; r < dataGridView2.Rows.Count; r++)
            {
                var row = dataGridView2.Rows[r];
                string nome = row.Cells["Nome"].Value?.ToString().ToUpper() ?? "";

                // Pula cabeçalhos
                if (nome.Contains("OPERADORES") || nome.Contains("APRENDIZ") ||
                    nome.Contains("CFTV") || nome.Contains("SUPERVISÃO")) continue;

                string horarioFunc = row.Cells["HORARIO"].Value?.ToString().Trim() ?? "";

                if (!TryParseHorario(horarioFunc, out TimeSpan iniFunc, out TimeSpan fimFunc))
                {
                    // Erro de leitura = Cinza Claro
                    for (int c = 2; c < dataGridView2.Columns.Count; c++)
                    {
                        row.Cells[c].Style.BackColor = System.Drawing.Color.LightGray;
                        row.Cells[c].ReadOnly = true;
                    }
                    continue;
                }

                // Ajuste para virada de noite (Funcionário)
                TimeSpan fimFuncAj = (fimFunc < iniFunc) ? fimFunc.Add(TimeSpan.FromHours(24)) : fimFunc;

                // --- O SEGREDO ESTÁ AQUI: TOLERÂNCIA DE SAÍDA ---
                // Subtraímos 1 minuto da saída para evitar que "encoste" na próxima coluna
                TimeSpan saidaVirtual = fimFuncAj.Subtract(TimeSpan.FromMinutes(1));
                // ------------------------------------------------

                for (int c = 2; c < dataGridView2.Columns.Count; c++)
                {
                    if (TryParseHorario(dataGridView2.Columns[c].HeaderText, out TimeSpan iniCol, out TimeSpan fimCol))
                    {
                        // Ajuste para virada de noite (Coluna)
                        TimeSpan fimColAj = (fimCol < iniCol) ? fimCol.Add(TimeSpan.FromHours(24)) : fimCol;

                        // LÓGICA DE INTERSECÇÃO USANDO A SAÍDA VIRTUAL
                        // 1. O funcionário entrou ANTES da coluna acabar?
                        // 2. A saída virtual (17:39) é DEPOIS OU IGUAL ao início da coluna?
                        bool disponivel = (iniFunc < fimColAj) && (saidaVirtual >= iniCol);

                        // Verificação extra para viradas de noite complexas
                        if (!disponivel)
                        {
                            // Tenta projetar +24h
                            disponivel = (iniFunc.Add(TimeSpan.FromHours(24)) < fimColAj) &&
                                         (saidaVirtual.Add(TimeSpan.FromHours(24)) >= iniCol);
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
                                row.Cells[c].ReadOnly = true;
                            }
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
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
            if (dataGridView1.Columns.Count > INDEX_NOME) { dataGridView1.Columns[INDEX_NOME].HeaderText = "NOME"; dataGridView1.Columns[INDEX_NOME].Width = 120; dataGridView1.Columns[INDEX_NOME].Frozen = true; }
            for (int i = INDEX_DIA_INICIO; i < dataGridView1.Columns.Count; i++) { dataGridView1.Columns[i].HeaderText = $"{i - INDEX_DIA_INICIO + 1}"; dataGridView1.Columns[i].Width = 35; }
        }
        private void ConfigurarGridEscalaDiaria()
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();


        

            dataGridView2.Columns.Add("HORARIO", "HORÁRIO");
            dataGridView2.Columns.Add("Nome", "Nome");

            // Array com horário ajustado (Minuto 41)
            string[] horarios = { "08:00 x 08:40", "08:41 x 09:40", "09:41 x 10:40", "10:41 x 11:40", "11:41 x 12:40", "12:41 x 13:40", "13:41 x 14:40", "14:41 x 15:40", "15:41 x 16:40", "16:41 x 17:40", "17:41 x 18:40", "18:41 x 19:40", "19:41 x 20:40", "20:41 x 21:40", "21:41 x 22:40", "22:41 x 23:40", "23:41 x 00:40", "00:41 x 01:40" };

            var postos = new List<string> { "", "CAIXA", "VALET", "QRF", "CIRC.", "REP|CIRC", "CS1", "CS2", "CS3", "SUP", "APOIO", "TREIN", "CFTV" };

            foreach (var h in horarios)
            {
                var col = new DataGridViewComboBoxColumn();
                col.HeaderText = h;
                col.DataSource = postos;
                col.FlatStyle = FlatStyle.Flat;
                col.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
                col.Width = 65;
                dataGridView2.Columns.Add(col);
                col.CellTemplate.Style.BackColor = System.Drawing.Color.DarkGray;
            }

            dataGridView2.Columns["HORARIO"].Frozen = true;
            dataGridView2.Columns["Nome"].Frozen = true;

            // Garante que as colunas congeladas tenham fundo branco ou cinza claro para destacar
            dataGridView2.Columns["HORARIO"].DefaultCellStyle.BackColor = System.Drawing.Color.White;
            dataGridView2.Columns["Nome"].DefaultCellStyle.BackColor = System.Drawing.Color.White;

            dataGridView2.Columns["HORARIO"].Width = 80;
            dataGridView2.Columns["Nome"].Width = 110;
        }
        private void CalcularTotais()
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                string textoLinha = dataGridView2.Rows[i].Cells[1].Value?.ToString().ToUpper() ?? "";
                if (textoLinha.Contains("OPERADORES") || textoLinha.Contains("APRENDIZ") || textoLinha.Contains("CFTV"))
                {
                    for (int c = 2; c < dataGridView2.Columns.Count; c++)
                    {
                        int count = 0;
                        for (int k = i - 1; k >= 0; k--)
                        {
                            string tAnt = dataGridView2.Rows[k].Cells[1].Value?.ToString().ToUpper() ?? "";
                            if (tAnt.Contains("OPERADORES") || tAnt.Contains("APRENDIZ") || tAnt.Contains("CFTV")) break;
                            if (!string.IsNullOrWhiteSpace(dataGridView2.Rows[k].Cells[c].Value?.ToString())) count++;
                        }
                        dataGridView2.Rows[i].Cells[c].Value = count > 0 ? count.ToString() : "";
                    }
                }
            }
        }
        private void DataGridView2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView2.IsCurrentCellDirty)
            {
                dataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit);
                CalcularTotais();
                PintarPostos();
                AtualizarItinerarios();
            }
        }
        private void DataGridView2_CellEnter(object sender, DataGridViewCellEventArgs e) { if (e.ColumnIndex > 1) SendKeys.Send("{F4}"); }
        private void CbSeletorDia_SelectedIndexChanged(object sender, EventArgs e) { if (CbSeletorDia.SelectedItem != null && Regex.Match(CbSeletorDia.SelectedItem.ToString(), @"\d+").Success) { _diaSelecionado = int.Parse(Regex.Match(CbSeletorDia.SelectedItem.ToString(), @"\d+").Value); ProcessarEscalaDoDia(); } }
        private bool EhFolga(string codigo) { if (string.IsNullOrWhiteSpace(codigo)) return false; return new[] { "X", "FOLGA", "FERIAS", "FÉRIAS" }.Contains(codigo.Trim().ToUpper()); }
        private bool TryParseHorario(string t, out TimeSpan i, out TimeSpan f) { i = TimeSpan.Zero; f = TimeSpan.Zero; var p = t.Split(new[] { 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries); if (p.Length == 2 && TimeSpan.TryParse(p[0].Trim(), out i) && TimeSpan.TryParse(p[1].Trim(), out f)) return true; return false; }
        private List<DataRow> OrdenarPorHorario(List<DataRow> l) { return l.OrderBy(r => r[INDEX_HORARIO].ToString()).ToList(); }
    }

    // Classes de Modelo
    public static class ExtensionMethods { public static void DoubleBuffered(this DataGridView dgv, bool setting) { Type dgvType = dgv.GetType(); System.Reflection.PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic); if (pi != null) pi.SetValue(dgv, setting, null); } }
    public class ItemItinerario { public string Horario { get; set; } public string Posto { get; set; } public System.Drawing.Color CorFundo { get; set; } public System.Drawing.Color CorTexto { get; set; } }
    public class CartaoFuncionario { public string Nome { get; set; } public List<ItemItinerario> Itens { get; set; } = new List<ItemItinerario>(); }
}