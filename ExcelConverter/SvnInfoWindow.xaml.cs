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
            bool isCheck = true;
            Utils.ReadSvnInfo(out user, out password, out folder, out isCheck);
            SvnUser.Text = user;
            SvnPassword.Password = password;
            ServerFolder.Text = folder;
            DrRadioBtn.IsChecked = isCheck;
        }

        private void SaveSvnInfoBtnClick(object sender, RoutedEventArgs e)
        {
            SaveAllInfo();
            this.Close();
        }

        private void SaveAllInfo()
        {
            var user = SvnUser.Text.TrimStart();
            var password = SvnPassword.Password;
            var folder = ServerFolder.Text.TrimStart();
            var check = DrRadioBtn.IsChecked;
            Utils.SaveSvnInfo(user, password, folder, check != null && check.Value);
        }

        private void ChangeToAlwaysUpDr(object sender, RoutedEventArgs e)
        {
            SaveAllInfo();
        }

        private void ChangeToNotAlwaysUpDr(object sender, RoutedEventArgs e)
        {
            SaveAllInfo();
        }
    }
}
