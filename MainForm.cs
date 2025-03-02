using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;

public class MainForm : Form
{
    private DocumentProcessor processor;
    private List<DocumentProcessor.DocumentData> documents;
    private Button btnLoadDocuments;
    private Button btnExportCSV;
    private Button btnConvertToExcel;
    private ComboBox cmbFields;
    private Chart chartData;

    public MainForm()
    {
        processor = new DocumentProcessor();
        documents = new List<DocumentProcessor.DocumentData>();
        InitializeComponents();
    }

    private void InitializeComponents()
    {
        this.Size = new System.Drawing.Size(800, 600);
        this.Text = "Обработка документов";

        // Кнопка загрузки документов
        btnLoadDocuments = new Button
        {
            Text = "Загрузить документы",
            Location = new System.Drawing.Point(10, 10),
            Size = new System.Drawing.Size(150, 30)
        };
        btnLoadDocuments.Click += BtnLoadDocuments_Click;

        // Кнопка экспорта в CSV
        btnExportCSV = new Button
        {
            Text = "Экспорт в CSV",
            Location = new System.Drawing.Point(170, 10),
            Size = new System.Drawing.Size(150, 30)
        };
        btnExportCSV.Click += BtnExportCSV_Click;

        // Кнопка конвертации в Excel
        btnConvertToExcel = new Button
        {
            Text = "Конвертировать в Excel",
            Location = new System.Drawing.Point(330, 10),
            Size = new System.Drawing.Size(150, 30)
        };
        btnConvertToExcel.Click += BtnConvertToExcel_Click;

        // Выпадающий список полей
        cmbFields = new ComboBox
        {
            Location = new System.Drawing.Point(490, 10),
            Size = new System.Drawing.Size(150, 30)
        };
        cmbFields.SelectedIndexChanged += CmbFields_SelectedIndexChanged;

        // График
        chartData = new Chart
        {
            Location = new System.Drawing.Point(10, 50),
            Size = new System.Drawing.Size(760, 500),
            Dock = DockStyle.Bottom
        };
        chartData.ChartAreas.Add(new ChartArea());

        // Добавляем элементы управления на форму
        this.Controls.AddRange(new Control[] { 
            btnLoadDocuments, 
            btnExportCSV, 
            btnConvertToExcel, 
            cmbFields, 
            chartData 
        });
    }

    private void BtnLoadDocuments_Click(object sender, EventArgs e)
    {
        try
        {
            documents.Clear();
            string folderPath = "texts";
            
            if (!Directory.Exists(folderPath))
            {
                MessageBox.Show("Папка 'texts' не найдена!");
                return;
            }

            foreach (var file in Directory.GetFiles(folderPath, "*.docx"))
            {
                documents.Add(processor.AnalyzeWordDocument(file));
            }

            // Обновляем список полей
            var fields = documents
                .SelectMany(d => d.Fields.Keys)
                .Distinct()
                .ToList();
            
            cmbFields.Items.Clear();
            cmbFields.Items.AddRange(fields.ToArray());

            MessageBox.Show($"Загружено {documents.Count} документов");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при загрузке документов: {ex.Message}");
        }
    }

    private void BtnExportCSV_Click(object sender, EventArgs e)
    {
        if (documents.Count == 0)
        {
            MessageBox.Show("Сначала загрузите документы!");
            return;
        }

        using (SaveFileDialog saveDialog = new SaveFileDialog())
        {
            saveDialog.Filter = "CSV files (*.csv)|*.csv";
            saveDialog.FilterIndex = 1;

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                processor.ExportToCSV(documents, saveDialog.FileName);
                MessageBox.Show("Экспорт завершен!");
            }
        }
    }

    private void BtnConvertToExcel_Click(object sender, EventArgs e)
    {
        using (OpenFileDialog openDialog = new OpenFileDialog())
        {
            openDialog.Filter = "Word files (*.docx)|*.docx";
            openDialog.FilterIndex = 1;

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    saveDialog.FilterIndex = 1;

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        processor.ConvertWordToExcel(openDialog.FileName, saveDialog.FileName);
                        MessageBox.Show("Конвертация завершена!");
                    }
                }
            }
        }
    }

    private void CmbFields_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (cmbFields.SelectedItem != null)
        {
            processor.CreateChart(documents, chartData, cmbFields.SelectedItem.ToString());
        }
    }
} 