using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
//using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FileManagementApp
{
    /// <summary>
    /// GuDingReplaceWindow.xaml 的交互逻辑
    /// </summary>
    public partial class GuDingReplaceWindow : Window
    {
        private string folderPath; // 用于保存选择的文件夹路径

        public GuDingReplaceWindow()
        {
            InitializeComponent();
        }

        // 确认替换按钮事件
        private async void OnConfirmReplace(object sender, RoutedEventArgs e)
        {
            string newSchoolName = NewSchoolNameTextBox.Text;
            string newDate = NewDateTextBox.Text;

            // 检查用户输入是否有效
            if (string.IsNullOrEmpty(newSchoolName) || string.IsNullOrEmpty(newDate))
            {
                System.Windows.MessageBox.Show("请确保所有的替换信息已填写完整！");
                return;
            }

            // 自动弹出选择文件夹对话框
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folderPath = folderDialog.SelectedPath;
                System.Windows.MessageBox.Show($"选择的文件夹是: {folderPath}");
            }
            else
            {
                System.Windows.MessageBox.Show("未选择文件夹，操作已取消！");
                return; // 如果用户没有选择文件夹，退出操作
            }

            // 检查选择的文件夹是否有效
            if (!Directory.Exists(folderPath))
            {
                System.Windows.MessageBox.Show("选择的文件夹无效！");
                return;
            }

            // 获取文件夹中的所有 Word 文件
            string[] files = Directory.GetFiles(folderPath, "*.docx");

            // 确保至少有一个文件
            if (files.Length == 0)
            {
                System.Windows.MessageBox.Show("文件夹中没有找到 Word 文件！");
                return;
            }

            // 初始化进度条
            ProgressBar.Value = 0;
            ProgressBar.Maximum = files.Length;

            // 批量替换文件夹中的所有 Word 文件
            int totalFiles = files.Length;
            for (int i = 0; i < totalFiles; i++)
            {
                try
                {
                    string filePath = files[i];
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                    {
                        var body = wordDoc.MainDocumentPart.Document.Body;

                        // 执行替换操作
                        ReplaceTextInWordDocument(body, "待替换学校名称", newSchoolName);
                        ReplaceTextInWordDocument(body, "待替换日期", newDate);

                    // 保存并关闭文件
                        wordDoc.MainDocumentPart.Document.Save();
                }
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"处理文件 {files[i]} 时发生错误: {ex.Message}");
                }

                // 更新进度条
                ProgressBar.Value = ((i + 1) * 100) / totalFiles;
                await System.Threading.Tasks.Task.Delay(100); // 延迟以便 UI 更新
            }

            // 替换完成
            System.Windows.MessageBox.Show("批量替换完成！");
            this.Close();
        }

        // 替换文档中的文本
        private void ReplaceTextInWordDocument(Body body, string oldText, string newText)
        {
            foreach (var textElement in body.Descendants<Text>())
            {
                if (textElement.Text.Contains(oldText))
                {
                    // 替换文本
                    textElement.Text = textElement.Text.Replace(oldText, newText);
                }
            }
        }
    }
}


//namespace FileManagementApp
//{
//    /// <summary>
//    /// GuDingReplaceWindow.xaml 的交互逻辑
//    /// </summary>
//    public partial class GuDingReplaceWindow : System.Windows.Window
//    {
//        private string folderPath; // 用于保存选择的文件夹路径
//        public GuDingReplaceWindow()
//        {
//            InitializeComponent();
//        }
//        // 确认替换按钮事件
//        private async void OnConfirmReplace(object sender, RoutedEventArgs e)
//        {
//            string newSchoolName = NewSchoolNameTextBox.Text;
//            string newDate = NewDateTextBox.Text;

//            // 检查用户输入是否有效
//            if (string.IsNullOrEmpty(newSchoolName) || string.IsNullOrEmpty(newDate))
//            {
//                System.Windows.MessageBox.Show("请确保所有的替换信息已填写完整！");
//                return;
//            }

//            // 自动弹出选择文件夹对话框
//            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
//            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
//            {
//                folderPath = folderDialog.SelectedPath;
//                System.Windows.MessageBox.Show($"选择的文件夹是: {folderPath}");
//            }
//            else
//            {
//                System.Windows.MessageBox.Show("未选择文件夹，操作已取消！");
//                return; // 如果用户没有选择文件夹，退出操作
//            }

//            // 检查选择的文件夹是否有效
//            if (!Directory.Exists(folderPath))
//            {
//                System.Windows.MessageBox.Show("选择的文件夹无效！");
//                return;
//            }

//            // 获取文件夹中的所有 Word 文件
//            string[] files = Directory.GetFiles(folderPath, "*.docx");

//            // 确保至少有一个文件
//            if (files.Length == 0)
//            {
//                System.Windows.MessageBox.Show("文件夹中没有找到 Word 文件！");
//                return;
//            }

//            // 初始化进度条
//            ProgressBar.Value = 0;
//            ProgressBar.Maximum = files.Length;

//            // 启动 Word 应用程序
//            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
//            wordApp.Visible = false; // 不显示 Word 界面

//            // 批量替换文件夹中的所有 Word 文件
//            int totalFiles = files.Length;
//            for (int i = 0; i < totalFiles; i++)
//            {
//                try
//                {
//                    string filePath = files[i];
//                    Document doc = wordApp.Documents.Open(filePath);

//                    // 执行替换操作，替换固定的“待替换学校名称”和“待替换日期”
//                    ReplaceTextInWordDocument(doc, "待替换学校名称", newSchoolName);
//                    ReplaceTextInWordDocument(doc, "待替换日期", newDate);

//                    // 保存并关闭文件
//                    doc.Save();
//                    doc.Close();
//                }
//                catch (Exception ex)
//                {
//                    System.Windows.MessageBox.Show($"处理文件 {files[i]} 时发生错误: {ex.Message}");
//                }

//                // 更新进度条
//                ProgressBar.Value = ((i + 1) * 100) / totalFiles;
//                await System.Threading.Tasks.Task.Delay(100); // 延迟以便 UI 更新
//            }

//            // 释放 Word 应用程序
//            wordApp.Quit();

//            // 替换完成
//            System.Windows.MessageBox.Show("批量替换完成！");
//            this.Close();
//        }

//        private void ReplaceTextInWordDocument(Document doc, string oldText, string newText)
//        {
//            object missing = Type.Missing;

//            // 获取 Word 查找对象
//            Find findObject = doc.Application.Selection.Find;

//            // 清除任何已有的格式设置
//            findObject.ClearFormatting();
//            findObject.Replacement.ClearFormatting();

//            // 设置查找文本和替换文本
//            findObject.Text = oldText;
//            findObject.Replacement.Text = newText;

//            // 设置替换的范围为整个文档
//            object replaceAll = WdReplace.wdReplaceAll;

//            // 执行查找和替换
//            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
//                ref missing, ref missing, ref missing, ref missing, ref replaceAll, ref missing, ref missing, ref missing);
//        }
//    }
//}
