using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FileManagementApp
{
    /// <summary>
    /// BatchCreateUserManuals.xaml 的交互逻辑
    /// </summary>
    public partial class BatchCreateUserManuals : Window
    {
        public BatchCreateUserManuals()
        {
            InitializeComponent();
        }

        private void BtnSelectTemplate_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog { Filter = "Word 文件 (*.docx)|*.docx" };
            if (dialog.ShowDialog() == true)
            {
                TxtTemplatePath.Text = dialog.FileName;
            }
        }

        private void BtnSelectOutput_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TxtOutputPath.Text = dialog.SelectedPath;
            }
        }

        private void BtnSelectExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog { Filter = "Excel 文件 (*.xlsx)|*.xlsx" };
            if (dialog.ShowDialog() == true)
            {
                TxtExcelPath.Text = dialog.FileName;
            }
        }

        private async void BtnGenerateFiles_Click(object sender, RoutedEventArgs e)
        {
            string templatePath = TxtTemplatePath.Text;
            string outputPath = TxtOutputPath.Text;
            string excelPath = TxtExcelPath.Text;
            Console.WriteLine("Excel file path: " + excelPath);

            if (string.IsNullOrWhiteSpace(templatePath) || string.IsNullOrWhiteSpace(outputPath) || string.IsNullOrWhiteSpace(excelPath))
            {
                MessageBox.Show("请确保模板路径、输出路径和 Excel 文件已设置！");
                return;
            }

            try
            {
                var workbook = new XLWorkbook(excelPath);
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed().Skip(1); // 跳过标题行

                int totalFiles = rows.Count();
                int processedFiles = 0;

                foreach (var row in rows)
                {
                    string fileName = row.Cell(1).GetValue<string>().Trim();
                    string experimentName = row.Cell(2).GetValue<string>().Trim();

                    if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(experimentName))
                        continue;

                    string prefixedFileName = $"{processedFiles + 1:00}_{fileName}.docx";
                    string outputFilePath = System.IO.Path.Combine(outputPath, prefixedFileName);

                    await Task.Run(() => CreateFileFromTemplate(templatePath, outputFilePath, experimentName));

                    processedFiles++;
                    ProgressBar.Value = (processedFiles * 100) / totalFiles;
                }

                MessageBox.Show("文件生成完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败: {ex.Message}");
                Console.WriteLine($"操作失败: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                MessageBox.Show($"操作失败: {ex.Message}");
            }

        }

        private void CreateFileFromTemplate(string templatePath, string outputPath, string experimentName)
        {
            // 使用 File.Copy 复制模板文件
            File.Copy(templatePath, outputPath, true);

            using (var wordDoc = WordprocessingDocument.Open(outputPath, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                // 查找并替换 "模板实验名称" 为实际的实验名称
                foreach (var text in body.Descendants<Text>())
                {
                    if (text.Text.Contains("模板实验名称"))
                    {
                        text.Text = text.Text.Replace("模板实验名称", experimentName);
                    }
                }

                wordDoc.MainDocumentPart.Document.Save();
            }
        }



    }
}
