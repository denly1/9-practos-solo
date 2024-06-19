using System.Windows;

namespace WordExcelEditor
{
    public partial class EmailWindow : Window
    {
        public string From => FromTextBox.Text;
        public string To => ToTextBox.Text;
        public string Subject => SubjectTextBox.Text;
        public string Body => BodyTextBox.Text;
        public string SmtpHost => "smtp.mail.ru";
        public int SmtpPort => 587; // для TLS
        public string SmtpUser => SmtpUserTextBox.Text;
        public string SmtpPass => SmtpPassTextBox.Password;

        public EmailWindow()
        {
            InitializeComponent();
        }

        private void Send_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }
    }
}
