using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace FileManagementApp
{
    /// <summary>
    /// ProgressWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ProgressWindow : Window
    {
        private CancellationTokenSource _cancellationTokenSource;
        public ProgressWindow()
        {
            InitializeComponent();
            _cancellationTokenSource = new CancellationTokenSource();
            double screenWidth = SystemParameters.WorkArea.Width;
            double screenHeight = SystemParameters.WorkArea.Height;
            this.Width = 500;  // 设置窗口宽度为屏幕宽度的80%
            this.Height = 140;  // 设置窗口高度为屏幕高度的80%
            this.Left = (screenWidth - this.Width) / 2;  // 水平居中窗口
            this.Top = (screenHeight - this.Height) / 2;  // 垂直居中窗口
        }
        public CancellationToken CancellationToken => _cancellationTokenSource.Token;

        public void UpdateProgress(int progress, int total)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressBar.Maximum = total;
                ProgressBar.Value = progress;
            });
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource.Cancel();
            Close();
        }
    }
}
