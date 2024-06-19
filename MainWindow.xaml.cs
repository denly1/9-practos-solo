using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Microsoft.Win32;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Xls;

namespace WordExcelEditor
{
    public partial class MainWindow : Window
    {
        private RichTextBox wordEditor;
        private DataGrid excelEditor;
        private Workbook workbook;

        public MainWindow()
        {
            InitializeComponent();
        }

        // Word Functions

        private void NewWord_Click(object sender, RoutedEventArgs e)
        {
            // Создание нового документа Word с примером текста и стилями
            CreateNewWordFile();
            ApplyExampleFormatting();
        }

        private void OpenWord_Click(object sender, RoutedEventArgs e)
        {
            // Открытие существующего файла Word
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents|*.docx";
            if (openFileDialog.ShowDialog() == true)
            {
                CreateNewWordFile();

                Document document = new Document();
                document.LoadFromFile(openFileDialog.FileName);

                TextRange range = new TextRange(wordEditor.Document.ContentStart, wordEditor.Document.ContentEnd);
                using (MemoryStream ms = new MemoryStream())
                {
                    document.SaveToStream(ms, Spire.Doc.FileFormat.Rtf);
                    ms.Seek(0, SeekOrigin.Begin);
                    range.Load(ms, DataFormats.Rtf);
                }
            }
        }

        private void SaveWord_Click(object sender, RoutedEventArgs e)
        {
            // Сохранение файла Word
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Documents|*.docx";
            if (saveFileDialog.ShowDialog() == true)
            {
                Document document = new Document();
                TextRange range = new TextRange(wordEditor.Document.ContentStart, wordEditor.Document.ContentEnd);
                using (MemoryStream ms = new MemoryStream())
                {
                    range.Save(ms, DataFormats.Rtf);
                    ms.Seek(0, SeekOrigin.Begin);
                    document.LoadFromStream(ms, Spire.Doc.FileFormat.Rtf);
                }
                document.SaveToFile(saveFileDialog.FileName, Spire.Doc.FileFormat.Docx);
            }
        }

        private void CreateNewWordFile()
        {
            EditorGrid.Children.Clear();
            wordEditor = new RichTextBox();
            EditorGrid.Children.Add(wordEditor);
        }

        private void ApplyExampleFormatting()
        {
            System.Windows.Documents.TextSelection selection = wordEditor.Selection; // Использование System.Windows.Documents.TextSelection
            if (selection != null)
            {
                // Пример применения форматирования
                selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                selection.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Italic);
                selection.ApplyPropertyValue(TextElement.FontSizeProperty, 18.0);
            }
        }

        private void ApplyTextStyle_Click(object sender, RoutedEventArgs e)
        {
            if (wordEditor == null || wordEditor.Selection == null) return;

            try
            {
                TextRange selection = new TextRange(wordEditor.Selection.Start, wordEditor.Selection.End);

                ComboBoxItem selectedItem = TextStyleComboBox.SelectedItem as ComboBoxItem;
                if (selectedItem != null && selectedItem.Tag is string styleTag)
                {
                    switch (styleTag)
                    {
                        case "Normal":
                            selection.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Normal);
                            selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal); // Установка обычного веса шрифта
                            selection.ApplyPropertyValue(Inline.TextDecorationsProperty, null); // Очистка стилей подчеркивания
                            break;
                        case "Italic":
                            selection.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Italic);
                            break;
                        case "Bold":
                            selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                            break;
                        case "Underline":
                            TextDecorationCollection underline = new TextDecorationCollection();
                            underline.Add(TextDecorations.Underline);
                            selection.ApplyPropertyValue(Inline.TextDecorationsProperty, underline);
                            break;
                        default:
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при применении стиля: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void ApplyFontSize_Click(object sender, RoutedEventArgs e)
        {
            if (wordEditor == null || wordEditor.Selection == null) return;

            try
            {
                TextRange selection = new TextRange(wordEditor.Selection.Start, wordEditor.Selection.End);

                ComboBoxItem selectedItem = FontSizeComboBox.SelectedItem as ComboBoxItem;
                if (selectedItem != null && double.TryParse(selectedItem.Content.ToString(), out double fontSize))
                {
                    selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSize);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при применении размера шрифта: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Excel Functions

        private void NewExcel_Click(object sender, RoutedEventArgs e)
        {
            // Создание нового файла Excel
            CreateNewExcelFile();
        }

        private void OpenExcel_Click(object sender, RoutedEventArgs e)
        {
            // Открытие существующего файла Excel
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                CreateNewExcelFile();

                workbook = new Workbook();
                workbook.LoadFromFile(openFileDialog.FileName);

                Worksheet sheet = workbook.Worksheets[0];
                var dataTable = sheet.ExportDataTable();
                excelEditor.ItemsSource = dataTable.DefaultView;
            }
        }

        private void SaveExcel_Click(object sender, RoutedEventArgs e)
        {
            // Сохранение файла Excel
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                Worksheet sheet = workbook.Worksheets[0];
                sheet.InsertDataTable(((System.Data.DataView)(excelEditor.ItemsSource)).Table, true, 1, 1);
                workbook.SaveToFile(saveFileDialog.FileName, Spire.Xls.ExcelVersion.Version2013);
            }
        }

        private void CreateNewExcelFile()
        {
            EditorGrid.Children.Clear();
            excelEditor = new DataGrid();
            EditorGrid.Children.Add(excelEditor);

            workbook = new Workbook();
            workbook.CreateEmptySheets(1);
        }

        // Email Function

        private void SendEmail_Click(object sender, RoutedEventArgs e)
        {
            // Отправка почты
            var emailWindow = new EmailWindow();
            emailWindow.ShowDialog();
            if (emailWindow.DialogResult == true)
            {
                string from = emailWindow.From;
                string to = emailWindow.To;
                string subject = emailWindow.Subject;
                string body = emailWindow.Body;
                string smtpHost = emailWindow.SmtpHost;
                int smtpPort = emailWindow.SmtpPort;
                string smtpUser = emailWindow.SmtpUser;
                string smtpPass = emailWindow.SmtpPass;

                try
                {
                    MailMessage mail = new MailMessage(from, to);
                    mail.Subject = subject;
                    mail.Body = body;

                    // Добавление вложений
                    if (wordEditor != null)
                    {
                        string tempPath = Path.GetTempPath() + "temp.docx";
                        SaveWordFile(tempPath);
                        mail.Attachments.Add(new Attachment(tempPath));
                    }

                    if (excelEditor != null)
                    {
                        string tempPath = Path.GetTempPath() + "temp.xlsx";
                        SaveExcelFile(tempPath);
                        mail.Attachments.Add(new Attachment(tempPath));
                    }

                    // Настройка SMTP клиента
                    SmtpClient smtpClient = new SmtpClient(smtpHost, smtpPort)
                    {
                        Credentials = new NetworkCredential(smtpUser, smtpPass),
                        EnableSsl = true
                    };

                    // Отправка письма
                    smtpClient.Send(mail);
                    MessageBox.Show("Письмо успешно отправлено.", "Отправка Email", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при отправке письма: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // Helper methods

        private void SaveWordFile(string filePath)
        {
            TextRange range = new TextRange(wordEditor.Document.ContentStart, wordEditor.Document.ContentEnd);
            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                range.Save(fs, DataFormats.Rtf);
            }
        }

        private void SaveExcelFile(string filePath)
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.InsertDataTable(((System.Data.DataView)excelEditor.ItemsSource).Table, true, 1, 1);
            workbook.SaveToFile(filePath, Spire.Xls.ExcelVersion.Version2013);
        }

        private void CountingButton_OnClick(object sender, RoutedEventArgs e)
        {

        }
    }
}
