using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Linq;
using FuzzySharp;

namespace FileManagementApp
{
    /// <summary>
    /// UIUpdateWindow1.xaml 的交互逻辑
    /// </summary>
    public partial class UIUpdateWindow1 : Window
    {
        private string excelFilePath; // 在类级别声明该字段，确保所有方法都能访问
        public UIUpdateWindow1()
        {
            InitializeComponent();
            // 在窗口加载时读取上次保存的路径并设置到TextBox
            if (Properties.Settings.Default.LastUserFolderPath != null)
            {
                FileSavePath.Text = Properties.Settings.Default.LastUserFolderPath;
            }
            
        }
        private void UIUpdateWindow1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // 将当前TextBox中的路径保存到应用程序设置
            Properties.Settings.Default.LastUserFolderPath = FileSavePath.Text;
            Properties.Settings.Default.Save();
        }
        //选择用户手册/资源包存放的文件夹   FileSavePath
        private void btnChooseSavePath_Click(object sender, RoutedEventArgs e)
        {
            //using (var folderDialog = new FolderBrowserDialog())
            //{
            //    DialogResult result = folderDialog.ShowDialog();
            //    if (result == System.Windows.Forms.DialogResult.OK)
            //    {
            //        FileSavePath.Text = folderDialog.SelectedPath;
            //    }
            //}

            // 创建文件对话框实例
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "请选择用户手册/资源包存放的文件夹",
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
                FileSavePath.Text = selectedFolderPath;

                // 提示用户选择的文件夹路径
                //System.Windows.MessageBox.Show($"你选择的文件夹路径为: {selectedFolderPath}");
            }
        }

        //选择表格  txtExcelPath
        private void btnChooseExcel_Click(object sender, RoutedEventArgs e)
        {
            // 创建用于选择文件的对话框实例
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            // 设置文件筛选器，只显示Excel文件类型
            openFileDialog.Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls";


            // 显示选择文件对话框，获取用户选择结果（返回值为true表示用户选择了文件，false表示取消操作）
            bool? openResult = openFileDialog.ShowDialog();
            if (openResult == true)
            {
                // 用户选择了已存在的Excel文件，这里可以对选择的文件路径进行后续操作，比如读取文件内容等
                string selectedFilePath = openFileDialog.FileName;
                txtExcelPath.Text = selectedFilePath;
                // 例如，简单打印选择的文件路径
                //System.Windows.MessageBox.Show($"你选择的已存在Excel文件路径为: {selectedFilePath}");
            }

        }

        //创建一个表格  txtExcelPath
        private void btnCreateExcel_Click(object sender, RoutedEventArgs e)
        {

            using (var folderDialog = new FolderBrowserDialog())
            {
                
                // 显示文件夹选择对话框
                DialogResult result = folderDialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK) // 检查用户是否点击了 OK 按钮
                {
                    // 获取用户选择的目录
                    string folderPath = folderDialog.SelectedPath;

                    // 构建 Excel 文件的基本名称
                    string baseFileName = "出库实验列表.xlsx";
                    excelFilePath = System.IO.Path.Combine(folderPath, baseFileName); // 更新 excelFilePath

                    // 检查文件是否存在，若存在则添加后缀
                    int counter = 1;
                    while (File.Exists(excelFilePath))
                    {
                        string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(baseFileName);
                        string extension = System.IO.Path.GetExtension(baseFileName);
                        excelFilePath = System.IO.Path.Combine(folderPath, $"{fileNameWithoutExtension}_{counter}{extension}");
                        counter++;
                    }

                    // 创建空的 Excel 文件
                    using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets.Add("ExportList");
                        worksheet.Cells[1, 1].Value = "实验名称"; // 第一列设置为“实验名称”

                        // 保存 Excel 文件
                        try
                        {
                            package.Save();
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show($"保存 Excel 文件失败: {ex.Message}");
                            return;
                        }
                    }
                    txtExcelPath.Text = excelFilePath;
                    // 自动打开 Excel 文件
                    try
                    {
                        System.Diagnostics.Process.Start(new ProcessStartInfo
                        {
                            FileName = excelFilePath,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show($"打开 Excel 文件失败: {ex.Message}");
                    }

                    System.Windows.MessageBox.Show("出库列表已创建并打开！");
                }
            }
        }

        //选择输出的文件夹   FileOutPutPath
        private void btnChooseOutputPath_Click(object sender, RoutedEventArgs e)
        {
            //using (var folderDialog = new FolderBrowserDialog())
            //{
            //    DialogResult result = folderDialog.ShowDialog();
            //    if (result == System.Windows.Forms.DialogResult.OK)
            //    {
            //        FileOutPutPath.Text = folderDialog.SelectedPath;
            //    }
            //}
            // 创建文件对话框实例
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "请选择用户手册/资源包输出的文件夹",
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
                FileOutPutPath.Text = selectedFolderPath;

                // 提示用户选择的文件夹路径
                //System.Windows.MessageBox.Show($"你选择的文件夹路径为: {selectedFolderPath}");
            }
        }
        private string[] GetExperimentNamesFromExcel(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.End.Row;
                var experimentNames = new List<string>();

                // 从第二行开始读取实验名称（假设第一行是表头）
                for (int row = 2; row <= rowCount; row++)
                {
                    var experimentName = worksheet.Cells[row, 1].Text;
                    if (!string.IsNullOrEmpty(experimentName))
                    {
                        experimentNames.Add(experimentName);
                    }
                }

                return experimentNames.ToArray();
            }
        }

        private void CopyFilesWithProgress(string[] experimentNames, string savePath, string outputPath, string excelPath, ProgressWindow progressWindow)
        {
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var notFoundIndexes = new ConcurrentBag<int>(); // 记录未找到实验的行索引
                var totalExperiments = experimentNames.Length;
                var progress = new ConcurrentBag<int>(); // 用于记录处理进度

                Parallel.ForEach(experimentNames, (name, state, index) =>
                {
                    progressWindow.CancellationToken.ThrowIfCancellationRequested();

                    // 获取所有文件进行模糊匹配
                    var files = Directory.GetFiles(savePath);

                    // 使用 FuzzySharp 计算相似度
                    var matchedFiles = files
                        .Select(file => new { File = file, Score = Fuzz.WeightedRatio(Path.GetFileNameWithoutExtension(file), name) })
                        .Where(match => match.Score > 70) // 设置相似度得分阈值
                        .OrderByDescending(match => match.Score)
                        .ToList();

                    if (matchedFiles.Count > 0)
                    {
                        foreach (var match in matchedFiles)
                        {
                            string newFileName = $"{index + 1}_{Path.GetFileName(match.File)}";
                            string destinationPath = Path.Combine(outputPath, newFileName);
                            File.Copy(match.File, destinationPath, overwrite: true);
                        }
                    }
                    else
                    {
                        notFoundIndexes.Add((int)index + 2); // 记录未找到的行（Excel 从第 2 行开始记录）
                    }

                    // 更新进度
                    progress.Add(1);

                    // 线程安全地更新进度条
                    progressWindow.Dispatcher.Invoke(() =>
                    {
                        progressWindow.UpdateProgress(progress.Count, totalExperiments);
                    });
                });

                // 高亮显示未找到的实验名称
                foreach (var rowIndex in notFoundIndexes)
                {
                    worksheet.Cells[rowIndex, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[rowIndex, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                }

                // 保存 Excel 文件
                package.Save();

                // 提示用户未找到的实验名称数量
                if (notFoundIndexes.Count > 0)
                {
                    System.Windows.MessageBox.Show($"以下实验未找到对应文件，已在 Excel 中高亮显示：{notFoundIndexes.Count} 条");
                }
                else
                {
                    System.Windows.MessageBox.Show("文件复制完成！");
                }
            }
        }


        //private void CopyFilesWithProgress(string[] experimentNames, string savePath, string outputPath, string excelPath, ProgressWindow progressWindow)
        //{
        //    using (var package = new ExcelPackage(new FileInfo(excelPath)))
        //    {
        //        var worksheet = package.Workbook.Worksheets[0];
        //        var notFound = new List<string>();

        //        for (int i = 0; i < experimentNames.Length; i++)
        //        {
        //            progressWindow.CancellationToken.ThrowIfCancellationRequested();

        //            string name = experimentNames[i];
        //            var files = Directory.GetFiles(savePath, $"*{name}*"); // 模糊匹配

        //            if (files.Length > 0)
        //            {
        //                foreach (var file in files)
        //                {
        //                    string newFileName = $"{i + 1}_{Path.GetFileName(file)}";
        //                    string destinationPath = Path.Combine(outputPath, newFileName);
        //                    File.Copy(file, destinationPath, overwrite: true);
        //                }
        //            }
        //            else
        //            {
        //                notFound.Add(name);
        //                worksheet.Cells[i + 2, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //                worksheet.Cells[i + 2, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
        //            }

        //            // 更新进度条
        //            progressWindow.Dispatcher.Invoke(() =>
        //            {
        //                progressWindow.UpdateProgress(i + 1, experimentNames.Length);
        //            });
        //        }

        //        // 保存 Excel 文件
        //        package.Save();
        //    }
        //}

        //开始检索

        private async void btnStartFindFile_Click(object sender, RoutedEventArgs e)
        {
            // 检查路径是否有效
            if (string.IsNullOrEmpty(txtExcelPath.Text) ||
                !Directory.Exists(FileSavePath.Text) ||
                !Directory.Exists(FileOutPutPath.Text))
            {
                System.Windows.MessageBox.Show("请先选择文件夹或出库列表！");
                return;
            }

            // 获取路径
            string savePath = FileSavePath.Text;
            string outputPath = FileOutPutPath.Text;
            string excelPath = txtExcelPath.Text;

            // 打开进度窗口
            var progressWindow = new ProgressWindow();
            progressWindow.Show();

            try
            {
                // 读取出库列表中的实验名称
                var experimentNames = GetExperimentNamesFromExcel(excelPath);
                if (experimentNames == null || experimentNames.Length == 0)
                {
                    System.Windows.MessageBox.Show("出库列表为空或无效！");
                    progressWindow.Close();
                    return;
                }

                // 异步执行文件操作
                await Task.Run(() =>
                {
                    CopyFilesWithProgress(experimentNames, savePath, outputPath, excelPath, progressWindow);
                }, progressWindow.CancellationToken);

                System.Windows.MessageBox.Show("文件复制完成！");
            }
            catch (OperationCanceledException)
            {
                System.Windows.MessageBox.Show("操作已取消！");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"发生错误: {ex.Message}");
            }
            finally
            {
                progressWindow.Close();
            }
        }

        //打开 FileSavePath
        private void btnOpenSavePath_Click(object sender, RoutedEventArgs e)
        {
            string filePath = FileSavePath.Text;
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
                   System.Diagnostics.Process.Start("explorer.exe", folderPath);
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

        //打开 FileOutPutPath
        private void btnOpenOutputPath_Click(object sender, RoutedEventArgs e)
        {
            string filePath = FileOutPutPath.Text;
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
                    System.Diagnostics.Process.Start("explorer.exe", folderPath);
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
