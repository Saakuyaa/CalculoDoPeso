using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalculodoPeso
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // Inicializa o Timer
            timer1.Interval = 1000; // Define o intervalo em milissegundos (1000ms = 1 segundo)
            timer1.Tick += Timer1_Tick; // Associa o evento Tick do Timer ao método Timer1_Tick
            timer1.Start(); // Inicia o Timer
            this.Shown += Form1_Shown;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            // Define o foco no campo "Responsável"
            textBox6.Focus();
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            // Atualiza o texto do Label com a data atual
            label13.Text = DateTime.Now.ToString("dd/MM/yyyy");

            // Atualiza o texto do Label com a hora atual
            label12.Text = DateTime.Now.ToString("HH:mm:ss");

            // Configurar TabIndex para garantir que o foco inicial esteja no campo "Responsável"
            textBox6.TabIndex = 0;  // Campo "Responsável"
            textBox1.TabIndex = 1;  // Primeiro peso
            textBox2.TabIndex = 2;  // Segundo peso
            textBox3.TabIndex = 3;  // Terceiro peso
            textBox4.TabIndex = 4;  // Quarto peso
            textBox5.TabIndex = 5;  // Quinto peso

            // Adicione eventos KeyPress para todos os TextBoxes.
            foreach (Control controle in this.Controls)
            {
                if (controle is TextBox)
                {
                    if (controle.Name == "textBox6")
                    {
                        ((TextBox)controle).KeyPress += new KeyPressEventHandler(TextBox6_KeyPress);
                    }
                    else
                    {
                        ((TextBox)controle).KeyPress += new KeyPressEventHandler(TextBox_KeyPress);
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            float soma = 0;
            float media;
            float valor;
            List<float> valores = new List<float>();

            foreach (Control controle in this.Controls)
            {
                if (controle is TextBox && controle.Name != "textBox6")
                {
                    if (float.TryParse(((TextBox)controle).Text, NumberStyles.Float, CultureInfo.GetCultureInfo("pt-BR"), out valor))
                    {
                        valores.Add(valor);
                        soma += valor;
                    }
                    else
                    {
                        MessageBox.Show("Por favor, insira valores numéricos válidos.");
                        return;
                    }
                }
            }

            if (valores.Count == 0)
            {
                MessageBox.Show("Por favor, insira valores nos campos de texto.");
                return;
            }

            media = soma / valores.Count;
            this.Controls["label8"].Text = "Média: " + media.ToString("F2", CultureInfo.GetCultureInfo("pt-BR")); // Formatação com duas casas decimais

            float maxValor = valores.Max();
            float minValor = valores.Min();
            float amplitude = maxValor - minValor;
            this.Controls["label10"].Text = "Amplitude: " + amplitude.ToString("F2", CultureInfo.GetCultureInfo("pt-BR")); // Formatação com duas casas decimais
        }

        // Evento KeyPress para permitir apenas entrada numérica
        private void TextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir números, backspace e vírgula decimal
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != ',')
            {
                e.Handled = true;
            }

            // Apenas uma vírgula decimal permitida
            TextBox textBox = sender as TextBox;
            if (e.KeyChar == ',' && textBox.Text.Contains(','))
            {
                e.Handled = true;
            }
        }

        // Evento KeyPress para permitir qualquer entrada de texto
        private void TextBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir apenas letras, backspace e espaço
            if (!char.IsControl(e.KeyChar) && !char.IsLetter(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e) //botao de limpar
        {
            foreach (Control controle in this.Controls)
            {
                if (controle is TextBox && controle.Name != "textBox6")
                {
                    ((TextBox)controle).Text = "";
                }
            }
            this.Controls["label8"].Text = "...";
            this.Controls["label10"].Text = "...";
            // Define o foco no campo de texto do primeiro peso (textBox1)
            textBox1.Focus();
        }

        private void button3_Click(object sender, EventArgs e) //botao de salvar
        {
            // Cria uma instância de SaveFileDialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            // Define o filtro de arquivos para mostrar apenas arquivos do Excel
            saveFileDialog.Filter = "Arquivos do Excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Obtém o caminho do arquivo selecionado pelo usuário
                string filePath = saveFileDialog.FileName;

                // Verifica se o campo TextBox6 está vazio
                if (string.IsNullOrEmpty(textBox6.Text))
                {
                    MessageBox.Show("Por favor, preencha o campo 'Responsável' antes de salvar.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                try
                {
                    // Cria uma nova instância do Excel
                    Excel.Application excelApp = new Excel.Application();
                    excelApp.Visible = false; // Altera a visibilidade para false
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                    // Adiciona os valores às células
                    worksheet.Cells[1, 1] = "Responsável";
                    worksheet.Cells[1, 2] = "Peso";

                    worksheet.Cells[2, 1] = textBox6.Text; // Responsável
                    worksheet.Cells[3, 1] = label13.Text; // Data
                    worksheet.Cells[4, 1] = label12.Text; // Hora

                    int rowIndex = 2; // Começa na segunda linha para os pesos

                    // Adiciona os pesos, se estiverem preenchidos
                    if (!string.IsNullOrEmpty(textBox1.Text))
                    {
                        worksheet.Cells[rowIndex++, 2] = textBox1.Text; // Peso 1
                    }
                    if (!string.IsNullOrEmpty(textBox2.Text))
                    {
                        worksheet.Cells[rowIndex++, 2] = textBox2.Text; // Peso 2
                    }
                    if (!string.IsNullOrEmpty(textBox3.Text))
                    {
                        worksheet.Cells[rowIndex++, 2] = textBox3.Text; // Peso 3
                    }
                    if (!string.IsNullOrEmpty(textBox4.Text))
                    {
                        worksheet.Cells[rowIndex++, 2] = textBox4.Text; // Peso 4
                    }
                    if (!string.IsNullOrEmpty(textBox5.Text))
                    {
                        worksheet.Cells[rowIndex++, 2] = textBox5.Text; // Peso 5
                    }

                    // Deixa uma linha em branco
                    rowIndex++;

                    // Adiciona a média e a amplitude
                    worksheet.Cells[rowIndex, 1] = "Média";
                    worksheet.Cells[rowIndex++, 2] = label8.Text.Split(':')[1].Trim(); // Média
                    worksheet.Cells[rowIndex, 1] = "Amplitude";
                    worksheet.Cells[rowIndex++, 2] = label10.Text.Split(':')[1].Trim(); // Amplitude

                    // Formatação de alinhamento central e todas as bordas
                    Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowIndex - 1, 2]];
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    // Formatação em negrito
                    range.Font.Bold = true;

                    // Define a cor de fundo para as células preenchidas
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    // Ajusta o tamanho das colunas ao conteúdo
                    worksheet.Columns.AutoFit();

                    // Salvar o arquivo
                    workbook.SaveAs(filePath);

                    // Liberar recursos
                    workbook.Close();
                    excelApp.Quit();

                    MessageBox.Show("Arquivo do Excel salvo com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ocorreu um erro ao salvar o arquivo do Excel: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

       
    }
}




