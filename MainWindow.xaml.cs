using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using ClosedXML.Excel;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs; // 用于选择文件夹
using OfficeOpenXml; // EPPlus 提供的功能
using System.Diagnostics;  // 引入 System.Diagnostics 命名空间
using System.Windows.Forms;
using Xceed.Words.NET;

namespace FileManagementApp
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        private string excelFilePath; // 在类级别声明该字段，确保所有方法都能访问
        public MainWindow()
        {
            // 设置 EPPlus 许可证上下文为免费开源许可
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
        }

        // 创建出库列表的功能
        private void BtnCreateExcel_Click(object sender, RoutedEventArgs e)
        {
            // 打开文件夹选择对话框，让用户选择保存目录
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "选择保存 Excel 文件的目录";
                folderDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); // 默认路径为用户文档

                // 显示文件夹选择对话框
                DialogResult result = folderDialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK) // 检查用户是否点击了 OK 按钮
                {
                    // 获取用户选择的目录
                    string folderPath = folderDialog.SelectedPath;

                    // 构建 Excel 文件的基本名称
                    string baseFileName = "ExportList.xlsx";
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

                    // 自动打开 Excel 文件
                    try
                    {
                        Process.Start(new ProcessStartInfo
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

        // 从 Excel 文件中读取实验名称
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

        // 选择并复制文件的功能
        private void BtnSelectDirectories_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
            {
                System.Windows.MessageBox.Show("请先创建出库列表！");
                return;
            }

            // 读取出库列表中的实验名称
            var experimentNames = GetExperimentNamesFromExcel(excelFilePath);

            // 选择源文件夹
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "选择源文件夹";
                DialogResult result = folderDialog.ShowDialog();

                if (result != System.Windows.Forms.DialogResult.OK)
                {
                    System.Windows.MessageBox.Show("未选择源文件夹！");
                    return;
                }

                string sourceFolder = folderDialog.SelectedPath; // 获取用户选择的源文件夹路径

                // 选择目标文件夹
                folderDialog.Description = "选择目标文件夹";
                result = folderDialog.ShowDialog();

                if (result != System.Windows.Forms.DialogResult.OK)
                {
                    System.Windows.MessageBox.Show("未选择目标文件夹！");
                    return;
                }

                string targetFolder = folderDialog.SelectedPath; // 获取用户选择的目标文件夹路径

                // 打开 Excel 文件进行修改
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var notFound = new List<string>();

                    for (int i = 0; i < experimentNames.Length; i++)
                    {
                        string name = experimentNames[i];
                        var files = Directory.GetFiles(sourceFolder, $"*{name}*"); // 模糊匹配

                        if (files.Length > 0)
                        {
                            foreach (var file in files)
                            {
                                string newFileName = $"{i + 1}_{System.IO.Path.GetFileName(file)}";
                                File.Copy(file, System.IO.Path.Combine(targetFolder, newFileName), overwrite: true);
                            }
                        }
                        else
                        {
                            notFound.Add(name);
                            // 将未找到的实验名称高亮标注
                            worksheet.Cells[i + 2, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[i + 2, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow); // 设置背景色为黄色
                        }
                    }

                    // 保存修改后的 Excel 文件
                    try
                    {
                        package.Save();
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show($"保存 Excel 文件失败: {ex.Message}");
                        return;
                    }

                    // 提示文件操作完成
                    if (notFound.Count > 0)
                    {
                        System.Windows.MessageBox.Show($"未找到部分实验文件，已在 Excel 中标注");
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("文件复制完成！");
                    }
                }
            }
        }

        private void BtnOpenReplaceWindow_Click(object sender, RoutedEventArgs e)
        {
            BatchReplaceWindow batchReplaceWindow = new BatchReplaceWindow();
            batchReplaceWindow.ShowDialog();
        }

        private void BtnFixedReplaceText_Click(object sender, object e)
        {
            GuDingReplaceWindow dialog = new GuDingReplaceWindow();
            dialog.ShowDialog();  // 显示弹框
        }

        private void OpenLink_Click(object sender, object e)
        {
            // 使用默认浏览器打开链接
            System.Diagnostics.Process.Start(new ProcessStartInfo("https://n1ddxc0sfaq.feishu.cn/docx/ZCQOd5wW3oHYUExmdxqcnraaned?from=from_copylink") { UseShellExecute = true });
        }

        private void OpenLink_Click(object sender, RoutedEventArgs e)
        {
            // 使用默认浏览器打开链接
            System.Diagnostics.Process.Start(new ProcessStartInfo("https://n1ddxc0sfaq.feishu.cn/docx/ZCQOd5wW3oHYUExmdxqcnraaned?from=from_copylink") { UseShellExecute = true });
        }

        private void BatchCreateUserManuals_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnBatchCreateUserManuals_Click(object sender, RoutedEventArgs e)
        {
            BatchCreateUserManuals batchCreateUserManuals = new BatchCreateUserManuals();
            batchCreateUserManuals.ShowDialog();
        }
    }
}
