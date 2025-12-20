using System;
using System.Collections.Generic; // Para List<>
using System.Drawing; // Para Size, Point
using System.Windows.Forms; // Para Form, Button, ComboBox

namespace Escala
{
    // Classe para selecionar qual aba do Excel importar
    public class SeletorPlanilha : Form
    {
        public ComboBox CbPlanilhas;
        private Button BtnOk;

        public SeletorPlanilha(List<string> planilhas)
        {
            this.Text = "Selecione a Aba";
            this.Size = new Size(300, 150);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog; // Janela fixa
            this.MaximizeBox = false;

            // Label
            Controls.Add(new Label { Text = "Selecione a aba:", Left = 10, Top = 10 });

            // ComboBox
            CbPlanilhas = new ComboBox
            {
                Left = 10,
                Top = 30,
                Width = 260,
                DataSource = planilhas,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            Controls.Add(CbPlanilhas);

            // Botão OK
            BtnOk = new Button
            {
                Text = "OK",
                Left = 190,
                Top = 70,
                DialogResult = DialogResult.OK
            };
            Controls.Add(BtnOk);

            // Define o botão padrão (Enter)
            this.AcceptButton = BtnOk;
        }
    }
}