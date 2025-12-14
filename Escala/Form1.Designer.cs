namespace Escala
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            CbSeletorDia = new ComboBox();
            btnImportar = new Button();
            btnRecarregarBanco = new Button();
            btnImprimir = new Button();
            tabPage2 = new TabPage();
            dataGridView2 = new DataGridView();
            tabPage1 = new TabPage();
            dataGridView1 = new DataGridView();
            tabControl1 = new TabControl();
            tabPage3 = new TabPage();
            flowLayoutPanel1 = new FlowLayoutPanel();
            lblClima = new Label();
            tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).BeginInit();
            tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            tabControl1.SuspendLayout();
            tabPage3.SuspendLayout();
            SuspendLayout();
            // 
            // CbSeletorDia
            // 
            CbSeletorDia.FormattingEnabled = true;
            CbSeletorDia.Location = new Point(281, 16);
            CbSeletorDia.Name = "CbSeletorDia";
            CbSeletorDia.Size = new Size(64, 23);
            CbSeletorDia.TabIndex = 1;
            // 
            // btnImportar
            // 
            btnImportar.Location = new Point(351, 3);
            btnImportar.Name = "btnImportar";
            btnImportar.Size = new Size(115, 47);
            btnImportar.TabIndex = 2;
            btnImportar.Text = "Carregar escala mensal";
            btnImportar.UseVisualStyleBackColor = true;
            // 
            // btnRecarregarBanco
            // 
            btnRecarregarBanco.Location = new Point(472, 3);
            btnRecarregarBanco.Name = "btnRecarregarBanco";
            btnRecarregarBanco.Size = new Size(115, 44);
            btnRecarregarBanco.TabIndex = 3;
            btnRecarregarBanco.Text = "Limpar Horários";
            btnRecarregarBanco.UseVisualStyleBackColor = true;
            btnRecarregarBanco.Click += btnRecarregarBanco_Click;
            // 
            // btnImprimir
            // 
            btnImprimir.Location = new Point(593, 3);
            btnImprimir.Name = "btnImprimir";
            btnImprimir.Size = new Size(115, 44);
            btnImprimir.TabIndex = 4;
            btnImprimir.Text = "Imprimir";
            btnImprimir.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(dataGridView2);
            tabPage2.Location = new Point(4, 24);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(1353, 638);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "EscalaDiaria";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // dataGridView2
            // 
            dataGridView2.BackgroundColor = SystemColors.ControlDarkDark;
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = SystemColors.WindowFrame;
            dataGridViewCellStyle1.Font = new Font("Segoe UI", 9F);
            dataGridViewCellStyle1.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.False;
            dataGridView2.DefaultCellStyle = dataGridViewCellStyle1;
            dataGridView2.Dock = DockStyle.Fill;
            dataGridView2.GridColor = Color.DarkGray;
            dataGridView2.Location = new Point(3, 3);
            dataGridView2.Name = "dataGridView2";
            dataGridView2.Size = new Size(1347, 632);
            dataGridView2.TabIndex = 1;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(dataGridView1);
            tabPage1.Location = new Point(4, 24);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(1353, 638);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "EscalaMensal";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.Location = new Point(3, 3);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Size = new Size(1347, 632);
            dataGridView1.TabIndex = 2;
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Controls.Add(tabPage3);
            tabControl1.Location = new Point(3, 29);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(1361, 666);
            tabControl1.TabIndex = 0;
            // 
            // tabPage3
            // 
            tabPage3.Controls.Add(flowLayoutPanel1);
            tabPage3.Location = new Point(4, 24);
            tabPage3.Name = "tabPage3";
            tabPage3.Size = new Size(1353, 638);
            tabPage3.TabIndex = 2;
            tabPage3.Text = "Itinerários";
            tabPage3.UseVisualStyleBackColor = true;
            // 
            // flowLayoutPanel1
            // 
            flowLayoutPanel1.BackColor = Color.Gray;
            flowLayoutPanel1.Location = new Point(3, 3);
            flowLayoutPanel1.Name = "flowLayoutPanel1";
            flowLayoutPanel1.Size = new Size(1333, 630);
            flowLayoutPanel1.TabIndex = 0;
            // 
            // lblClima
            // 
            lblClima.AutoSize = true;
            lblClima.Font = new Font("Segoe UI", 14.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblClima.Location = new Point(842, 15);
            lblClima.Name = "lblClima";
            lblClima.Size = new Size(205, 25);
            lblClima.TabIndex = 5;
            lblClima.Text = "Carregando previsão ...";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1368, 698);
            Controls.Add(lblClima);
            Controls.Add(btnImprimir);
            Controls.Add(btnRecarregarBanco);
            Controls.Add(btnImportar);
            Controls.Add(CbSeletorDia);
            Controls.Add(tabControl1);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView2).EndInit();
            tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            tabControl1.ResumeLayout(false);
            tabPage3.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private ComboBox CbSeletorDia;
        private Button btnImportar;
        private Button btnRecarregarBanco;
        private Button btnImprimir;
        private TabPage tabPage2;
        private DataGridView dataGridView2;
        private TabPage tabPage1;
        private TabControl tabControl1;
        private DataGridView dataGridView1;
        private TabPage tabPage3;
        private FlowLayoutPanel flowLayoutPanel1;
        private Label lblClima;
    }
}
