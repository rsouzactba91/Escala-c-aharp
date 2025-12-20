using System;
using System.Drawing; // Necessário para Point, Size, Color, Font
using System.Windows.Forms; // Necessário para Form, Button, Label, etc.

namespace Escala
{
    public class FormGerenciar : Form
    {
        private ListBox lbPostos, lbHorarios;
        private TextBox txtPosto, txtHorario; // Removi txtHorarioPadrao pois não estava sendo usado nas vars globais
        private ComboBox CbHorarioPadraoFolguista; // ComboBox para o Folguista
        private ComboBox cbHorarioPadrao; // Novo ComboBox para o Intermediário
        private Button btnAddPosto, btnDelPosto, btnAddHorario, btnDelHorario, btnSalvarPadrao, btnSalvarIntermediario;

        public FormGerenciar()
        {
            Text = "Gerenciar Listas e Configurações";
            Size = new Size(500, 500);
            StartPosition = FormStartPosition.CenterParent;

            TabControl tabs = new TabControl { Dock = DockStyle.Fill };

            // -------------------------------------------------------
            // ABA 1: POSTOS
            // -------------------------------------------------------
            TabPage tabP = new TabPage("Postos");
            lbPostos = new ListBox { Location = new Point(10, 10), Size = new Size(200, 250) };
            txtPosto = new TextBox { Location = new Point(220, 10), Size = new Size(150, 23) };
            btnAddPosto = new Button { Text = "Add", Location = new Point(220, 40) };
            btnDelPosto = new Button { Text = "Del", Location = new Point(10, 270), BackColor = Color.LightCoral };

            btnAddPosto.Click += (s, e) => {
                if (!string.IsNullOrEmpty(txtPosto.Text))
                {
                    DatabaseService.AdicionarPosto(txtPosto.Text);
                    Carregar();
                    txtPosto.Clear();
                }
            };
            btnDelPosto.Click += (s, e) => {
                if (lbPostos.SelectedItem != null)
                {
                    DatabaseService.RemoverPosto(lbPostos.SelectedItem.ToString());
                    Carregar();
                }
            };
            tabP.Controls.AddRange(new Control[] { lbPostos, txtPosto, btnAddPosto, btnDelPosto });

            // -------------------------------------------------------
            // ABA 2: HORÁRIOS (Lista para Colunas)
            // -------------------------------------------------------
            TabPage tabH = new TabPage("Horários (Colunas)");
            lbHorarios = new ListBox { Location = new Point(10, 10), Size = new Size(200, 250) };
            txtHorario = new TextBox { Location = new Point(220, 10), Size = new Size(150, 23) };
            btnAddHorario = new Button { Text = "Add", Location = new Point(220, 40) };
            btnDelHorario = new Button { Text = "Del", Location = new Point(10, 270), BackColor = Color.LightCoral };

            btnAddHorario.Click += (s, e) => {
                if (!string.IsNullOrEmpty(txtHorario.Text))
                {
                    DatabaseService.AdicionarHorario(txtHorario.Text);
                    Carregar();
                    txtHorario.Clear();
                }
            };
            btnDelHorario.Click += (s, e) => {
                if (lbHorarios.SelectedItem != null)
                {
                    DatabaseService.RemoverHorario(lbHorarios.SelectedItem.ToString());
                    Carregar();
                }
            };
            tabH.Controls.AddRange(new Control[] { lbHorarios, txtHorario, btnAddHorario, btnDelHorario });

            // -------------------------------------------------------
            // ABA 3: FOLGUISTA
            // -------------------------------------------------------
            TabPage tabC = new TabPage("Horários folguistas");
            Label lblExplica = new Label { Text = "Horário Padrão do Folguista (Se ninguém faltar):", Location = new Point(10, 20), AutoSize = true, Font = new Font("Arial", 10, FontStyle.Bold), Width = 400 };

            CbHorarioPadraoFolguista = new ComboBox { Location = new Point(10, 50), Width = 200, Font = new Font("Arial", 12) };
            CbHorarioPadraoFolguista.Items.AddRange(new object[] {
                 "07:00 X 15:20", "09:40 X 18:00", "10:40 X 19:00",
                 "11:40 X 20:00", "12:40 X 21:00", "16:40 X 01:00"
             });

            btnSalvarPadrao = new Button { Text = "Salvar Padrão", Location = new Point(220, 48), Width = 100, Height = 30, BackColor = Color.LightGreen };

            btnSalvarPadrao.Click += (s, e) => {
                DatabaseService.SetHorarioPadraoFolguista(CbHorarioPadraoFolguista.Text);
                MessageBox.Show("Horário do Folguista atualizado!");
            };
            tabC.Controls.AddRange(new Control[] { lblExplica, CbHorarioPadraoFolguista, btnSalvarPadrao });

            // -------------------------------------------------------
            // ABA 4: CFTV (INTERMEDIÁRIO)
            // -------------------------------------------------------
            TabPage tabD = new TabPage("CFTV (Intermediário)");
            Label lblExplicaInter = new Label { Text = "Horário Padrão do Intermediário (12:40):", Location = new Point(10, 20), AutoSize = true, Font = new Font("Arial", 10, FontStyle.Bold), Width = 400 };

            cbHorarioPadrao = new ComboBox { Location = new Point(10, 50), Width = 200, Font = new Font("Arial", 12), DropDownStyle = ComboBoxStyle.DropDown };
            cbHorarioPadrao.Items.AddRange(new object[] {
                 "07:00 X 15:20", "09:40 X 18:00", "10:40 X 19:00",
                 "11:40 X 20:00", "12:40 X 21:00", "16:40 X 01:00"
             });

            btnSalvarIntermediario = new Button { Text = "Salvar Interm.", Location = new Point(220, 48), Width = 100, Height = 30, BackColor = Color.LightGreen };

            btnSalvarIntermediario.Click += (s, e) => {
                DatabaseService.SetHorarioPadraoIntermediario(cbHorarioPadrao.Text);
                MessageBox.Show("Horário Intermediário atualizado!");
            };

            tabD.Controls.AddRange(new Control[] { lblExplicaInter, cbHorarioPadrao, btnSalvarIntermediario });

            // Adiciona as abas
            tabs.TabPages.Add(tabP);
            tabs.TabPages.Add(tabH);
            tabs.TabPages.Add(tabC);
            tabs.TabPages.Add(tabD);
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

            // Carrega o Horário Padrão do FOLGUISTA
            CbHorarioPadraoFolguista.Text = DatabaseService.GetHorarioPadraoFolguista();

            // Carrega o Horário Padrão do INTERMEDIÁRIO
            cbHorarioPadrao.Text = DatabaseService.GetHorarioPadraoIntermediario();
        }
    }
}