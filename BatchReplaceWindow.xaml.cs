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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Word;

namespace FileManagementApp
{
    /// <summary>
    /// BatchReplaceWindow.xaml 的交互逻辑
    /// </summary>
    public partial class BatchReplaceWindow : System.Windows.Window
    {
        // 将 wordApp 声明为类的成员变量
        private Microsoft.Office.Interop.Word.Application wordApp;

        public BatchReplaceWindow()
        {
            InitializeComponent();
            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false; // 不显示 Word 窗口
        }
        // 批量替换的开始按钮点击事件
        private async void BtnStartReplace_Click(object sender, RoutedEventArgs e)
        {
            // 获取用户输入的文本
            string oldSchoolName = TxtOldSchoolName.Text;
            string newSchoolName = TxtNewSchoolName.Text;
            string oldDate = TxtOldDate.Text;
            string newDate = TxtNewDate.Text;

            // 检查输入是否为空
            if (string.IsNullOrWhiteSpace(oldSchoolName) || string.IsNullOrWhiteSpace(newSchoolName) ||
                string.IsNullOrWhiteSpace(oldDate) || string.IsNullOrWhiteSpace(newDate))
            {
                MessageBox.Show("所有文本框都必须填写！");
                return;
            }

            // 选择文件夹
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            folderDialog.Description = "选择包含 Word 文件的文件夹";
            var result = folderDialog.ShowDialog();

            if (result != System.Windows.Forms.DialogResult.OK)
            {
                MessageBox.Show("未选择文件夹！");
                return;
            }

            string folderPath = folderDialog.SelectedPath;

            // 获取所有 Word 文件
            var wordFiles = Directory.GetFiles(folderPath, "*.docx");

            if (wordFiles.Length == 0)
            {
                MessageBox.Show("文件夹中没有 Word 文件！");
                return;
            }

            // 定义 missing
            object missing = Missing.Value;

            // 执行批量替换操作，并更新进度条
            int totalFiles = wordFiles.Length;
            for (int i = 0; i < totalFiles; i++)
            {
                string wordFile = wordFiles[i];

                try
                {
                    // 打开 Word 文件
                    Document doc = wordApp.Documents.Open(wordFile);

                    // 替换文本
                    FindAndReplace(doc, oldSchoolName, newSchoolName);
                    FindAndReplace(doc, oldDate, newDate);

                    // 保存文件并关闭
                    doc.Save();
                    doc.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"处理文件 {System.IO.Path.GetFileName(wordFile)} 时出错: {ex.Message}");
                }

                // 更新进度条
                ProgressBar.Value = ((i + 1) * 100) / totalFiles;
                await System.Threading.Tasks.Task.Delay(50); // 添加延迟以便UI更新
            }

            MessageBox.Show("批量替换完成！");
        }

        // 查找并替换文本
        private void FindAndReplace(Document doc, string oldText, string newText)
        {
            // 定义 missing 变量
            object missing = Type.Missing;

            // 设置查找和替换选项
            Find findObject = doc.Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Replacement.ClearFormatting();
            findObject.Text = oldText;
            findObject.Replacement.Text = newText;

            // 执行替换
            object replaceAll = WdReplace.wdReplaceAll;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceAll, ref missing, ref missing, ref missing);
        }

        // 窗体关闭时释放 Word 应用
        private void Window_Closed(object sender, EventArgs e)
        {
            if (wordApp != null)
            {
                wordApp.Quit();
            }
        }

    }
}
