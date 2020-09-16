using System.Windows;

namespace ExcelConverter
{
    /// <summary>
    /// SvnInfoWindow.xaml 的交互逻辑
    /// </summary>
    public partial class SvnInfoWindow : Window
    {
        public SvnInfoWindow()
        {
            InitializeComponent();

            InitRecord();
        }

        private void InitRecord()
        {
            string user;
            string password;
            string folder;
            Utils.ReadSvnInfo(out user, out password, out folder);
            SvnUser.Text = user;
            SvnPassword.Password = password;
            ServerFolder.Text = folder;
        }

        private void SaveSvnInfoBtnClick(object sender, RoutedEventArgs e)
        {
            var user = SvnUser.Text.TrimStart();
            var password = SvnPassword.Password;
            var folder = ServerFolder.Text.TrimStart();
            Utils.SaveSvnInfo(user, password, folder);
        }
    }
}
