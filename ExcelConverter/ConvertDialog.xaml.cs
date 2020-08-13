using System;
using System.Windows;

namespace ExcelConverter
{
    /// <summary>
    /// ConvertDialog.xaml 的交互逻辑
    /// </summary>
    public partial class ConvertDialog : Window
    {
        public event Action OnClosedEvent;
        public ConvertDialog()
        {
            InitializeComponent();
        }

        private void OnWindowsLoaded(object sender, RoutedEventArgs e)
        {
            EventDispatcher.RegdEvent<string>(TaskType.ConvertOutput, OnConvertOutput);
        }

        private void OnConvertOutput(string output)
        {
            ConvertOutputText.Text += output + "\n";
            if(Math.Abs(TextScroll.ScrollableHeight - TextScroll.VerticalOffset) < 0.1)
                TextScroll.ScrollToEnd();
        }

        private void OnWindowsClosed(object? sender, EventArgs e)
        {
            EventDispatcher.RemoveEvent(TaskType.ConvertOutput);
            OnClosedEvent?.Invoke();
        }
    }
}
