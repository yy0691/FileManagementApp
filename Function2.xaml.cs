using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FileManagementApp
{
    /// <summary>
    /// Function2.xaml 的交互逻辑
    /// </summary>
    public partial class Function2 : Window
    {

        

        public Function2()
        {
            InitializeComponent();
            //double screenWidth = SystemParameters.WorkArea.Width;
            //double screenHeight = SystemParameters.WorkArea.Height;
            //this.Width = screenWidth * 0.8;  // 设置窗口宽度为屏幕宽度的80%
            //this.Height = screenHeight * 0.8;  // 设置窗口高度为屏幕高度的80%
            //this.Left = (screenWidth - this.Width) / 2;  // 水平居中窗口
            //this.Top = (screenHeight - this.Height) / 2;  // 垂直居中窗口
            int currentYear = DateTime.Now.Year;  // 获取当前年份
            int currentMonth = DateTime.Now.Month;  // 获取当前月份

            Console.WriteLine($"当前年份：{currentYear}");
            Console.WriteLine($"当前月份：{currentMonth}");
            TxtNewDate.Text = $"{currentYear}年{currentMonth}月";
        }

        private void ReplaceTextInDocuments(string folderPath, string oldSchoolName, string newSchoolName, string oldDate, string newDate)
        {
            var wordFiles = Directory.GetFiles(folderPath, "*.docx");

            foreach (var file in wordFiles)
            {
                try
                {
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(file, true))
                    {
                        var body = wordDoc.MainDocumentPart.Document.Body;

                        // 替换文本
                        ReplaceTextInBody(body, oldSchoolName, newSchoolName);
                        ReplaceTextInBody(body, oldDate, newDate);

                        // 保存文档
                        wordDoc.MainDocumentPart.Document.Save();
                    }

                    // 更新进度条
                    ProgressBar.Dispatcher.Invoke(() =>
                    {
                        ProgressBar.Value = (ProgressBar.Value + 100.0 / wordFiles.Length);
                    }, System.Windows.Threading.DispatcherPriority.Background);


                }
                catch (Exception ex)
                {
                    MessageBox.Show($"处理文件 {file} 时出错: {ex.Message}");
                }
            }
        }

        private void ReplaceTextInBody(Body body, string oldText, string newText)
        {
            foreach (var paragraph in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
            {
                var runs = paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().ToList();
                var buffer = new StringBuilder();
                var runMapping = new List<(DocumentFormat.OpenXml.Wordprocessing.Run, int, int)>(); // Track text positions and runs

                // Gather text and record mapping
                foreach (var run in runs)
                {
                    var textElement = run.GetFirstChild<Text>();
                    if (textElement != null)
                    {
                        buffer.Append(textElement.Text);
                        runMapping.Add((run, buffer.Length - textElement.Text.Length, buffer.Length));
                    }
                }

                // Replace text
                string fullText = buffer.ToString();
                if (fullText.Contains(oldText))
                {
                    string updatedText = fullText.Replace(oldText, newText);

                    // Clear existing runs
                    foreach (var run in runs)
                    {
                        run.RemoveAllChildren<Text>();
                    }

                    // Redistribute text to existing runs
                    int currentIndex = 0;
                    foreach (var (run, start, end) in runMapping)
                    {
                        if (currentIndex >= updatedText.Length) break;

                        int lengthToTake = Math.Min(updatedText.Length - currentIndex, end - start);
                        string subText = updatedText.Substring(currentIndex, lengthToTake);

                        run.AppendChild(new Text(subText));
                        currentIndex += lengthToTake;
                    }

                    // Append remaining text if necessary
                    if (currentIndex < updatedText.Length)
                    {
                        string remainingText = updatedText.Substring(currentIndex);
                        var lastRun = runMapping.Last().Item1;
                        var newTextElement = new Text(remainingText);
                        lastRun.AppendChild(newTextElement);
                    }
                }
            }
        }
        private void BtnStartReplace_Click(object sender, RoutedEventArgs e)
        {
            // 当按钮被点击时，将进度条的可见性设置为可见
            ProgressBar.Visibility = Visibility.Visible;
            string folderPath = WordSavePath.Text; // 选择上一步的文件夹
            if (string.IsNullOrEmpty(folderPath))
            {
                MessageBox.Show("未选择文件夹！");
                return;
            }

            string oldSchoolName = TxtOldSchoolName.Text; // 用户输入的旧学校名称
            string newSchoolName = TxtNewSchoolName.Text; // 用户输入的新学校名称
            string oldDate = TxtOldDate.Text; // 用户输入的旧日期
            string newDate = TxtNewDate.Text; // 用户输入的新日期

            if (string.IsNullOrEmpty(oldSchoolName) || string.IsNullOrEmpty(oldDate) ||
                string.IsNullOrEmpty(newSchoolName) || string.IsNullOrEmpty(newDate))
            {
                MessageBox.Show("所有字段都必须填写！");
                return;
            }

            // 启动替换操作
            ReplaceTextInDocuments(folderPath, oldSchoolName, newSchoolName, oldDate, newDate);

            MessageBox.Show("批量替换完成！");
        }

        private string SelectFolder()
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                var result = dialog.ShowDialog();
                return result == System.Windows.Forms.DialogResult.OK ? dialog.SelectedPath : null;
            }
        }

        private void btnChooseSavePath_Click(object sender, RoutedEventArgs e)
        {
            // 创建文件对话框实例
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "请选择用户手册存放的文件夹",
                Filter = "文件夹|*.none", // 添加一个虚拟的文件类型过滤器
                CheckFileExists = false, // 不检查文件是否存在
                ValidateNames = false, // 允许无效的文件名（用于选择文件夹）
                FileName = "选择此文件夹" // 提供默认的文件名
            };

            // 显示对话框并检查用户操作
            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                // 获取文件夹路径
                string selectedFolderPath = System.IO.Path.GetDirectoryName(openFileDialog.FileName);
                WordSavePath.Text = selectedFolderPath;

                // 提示用户选择的文件夹路径
                //System.Windows.MessageBox.Show($"你选择的文件夹路径为: {selectedFolderPath}");
            }
        }

        private void btnOpenSavePath_Click(object sender, RoutedEventArgs e)
        {
            string filePath = WordSavePath.Text;
            if (!string.IsNullOrEmpty(filePath))
            {
                try
                {
                    // 获取文件所在的文件夹路径，如果输入的本身就是文件夹路径则直接使用
                    string folderPath = filePath;
                    if (string.IsNullOrEmpty(folderPath))
                    {
                        folderPath = filePath;
                    }
                    // 使用Process类启动文件资源管理器并打开指定文件夹
                    Process.Start("explorer.exe", folderPath);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"打开文件夹时出现错误: {ex.Message}", "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
            }
            else
            {
                System.Windows.MessageBox.Show("请先输入有效的文件路径", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            }
        }
    }
}
