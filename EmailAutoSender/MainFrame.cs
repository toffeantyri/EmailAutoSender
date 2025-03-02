using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.IO;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace EmailAutoSender
{
    public partial class MainFrame : Form
    {



        public MainFrame()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void MainFrame_Load(object sender, EventArgs e)
        {

        }

        // Класс для хранения данных о письме
        public class EmailData
        {
            public string Email { get; set; }
            public bool IsSent { get; set; }
            public string FilePath { get; set; }
        }

        private void LoadSettings() 
        {
            txtEmail.Text = Properties.Settings.Default.email;
            txtPassword.Text = Properties.Settings.Default.password;
            txtSmtpServer.Text = Properties.Settings.Default.smtpServer;
            txtPort.Text = Properties.Settings.Default.smtpPort.ToString();
            txtExcelPath.Text = Properties.Settings.Default.filePath;
            
        }

        private void SaveSettings()
        {
            int smptPortValid = 0;
            int.TryParse(txtPort.Text, out smptPortValid);

            Properties.Settings.Default.email = txtEmail.Text;
            Properties.Settings.Default.password = txtPassword.Text;
            Properties.Settings.Default.smtpServer = txtSmtpServer.Text;
            Properties.Settings.Default.smtpPort = smptPortValid;
            Properties.Settings.Default.filePath = txtExcelPath.Text;

            Properties.Settings.Default.Save(); // Сохранение настроек
        }

        private void browseFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcelPath.Text = openFileDialog.FileName;
                }
            }
        }

        private void sendFilesToEmails_Click(object sender, EventArgs e)
        {
            // Сохранение данных перед началом рассылки
            SaveSettings();

            int smptPortValid = 0;
            if (!int.TryParse(txtPort.Text, out smptPortValid)) {
                MessageBox.Show("Неверный порт");
                return;
            }

            // Получение данных из формы
            string email = txtEmail.Text;
            string password = txtPassword.Text;
            string smtpServer = txtSmtpServer.Text;
            int smtpPort = int.Parse(txtPort.Text);
            string excelFilePath = txtExcelPath.Text;

            if (!isValidEmail(email)) {
                MessageBox.Show("Введите правильный email");
                return;
            }

            // Чтение данных из Excel
            var emailData = ReadEmailDataFromExcel(excelFilePath);

            // Настройка ProgressBar
            progressBar1.Maximum = emailData.Count;
            progressBar1.Value = 0;

            var smtpClient = new SmtpClient(smtpServer, smtpPort);

            // Отправка писем
            using (smtpClient as System.IDisposable)
            {
                smtpClient.UseDefaultCredentials = false;
                smtpClient.EnableSsl = true;
                smtpClient.Credentials = new System.Net.NetworkCredential(email, password);

                for (int i = 0; i < emailData.Count; i++)
                {
                    var data = emailData[i];

                    // Пропуск, если письмо уже отправлено
                    if (data.IsSent)
                    {
                        progressBar1.Value++;
                       
                        lblStatus.Text += $"Пропуск {data.Email} (уже отправлено) \n";
                        continue;
                    }

                    try
                    {
                        // Отправка письма с вложением
                        
                        SendEmail(smtpClient, data.Email, "Тема письма", "Текст письма", data.FilePath,email, password);
                        lblStatus.Text = $"Отправлено {i + 1} из {emailData.Count} писем";

                        // Обновление статуса "Отправлено" в Excel
                        UpdateSentStatus(excelFilePath, i, true);

                        //progressBar1.Value++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при отправке на {data.Email}: {ex.Message}");
                      
                    }
                   
                    progressBar1.Value++;
                }
            }
            if (emailData.Count != 0)
            {
                MessageBox.Show("Рассылка завершена!");
            }
        }




        // Чтение данных из Excel
        private List<EmailData> ReadEmailDataFromExcel(string filePath)
        {
            List<EmailData> emailData = new List<EmailData>();

            try {
                new FileInfo(filePath);
            } catch(Exception) {
                MessageBox.Show("Ошибка чтения файла Excel, неверный путь");
                return new List<EmailData>();
            }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[0]; // Первый лист

                int rowCount = worksheet.Dimension.Rows;

                // Чтение данных
                for (int row = 2; row <= rowCount; row++) // Предполагаем, что первая строка — заголовок
                {
                    string email = worksheet.Cells[row, 1].Text; // Столбец "Email"
                    Console.WriteLine(email);
                    bool isSent = worksheet.Cells[row, 2].GetValue<bool>(); // Столбец "Отправлено"
                    Console.WriteLine(isSent);
                    string sendFilePath = worksheet.Cells[row, 3].Text; // Столбец "Путь к файлу"

                    emailData.Add(new EmailData { Email = email, IsSent = isSent, FilePath = sendFilePath });
                }
            }

            return emailData;
        }



        // Обновление статуса "Отправлено" в Excel
        private void UpdateSentStatus(string filePath, int rowIndex, bool isSent)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[0];

                // Обновление значения в столбце "Отправлено"
                worksheet.Cells[rowIndex + 2, 2].Value = isSent; // +2, так как первая строка — заголовок

                // Сохранение изменений
                package.Save();
            }
        }

        // Отправка письма
        private void SendEmail(SmtpClient smtpClient, string toEmail, string subject, string body, string attachmentPath, string email, string password)
        {
            using (MailMessage mailMessage = new MailMessage())
            {
                mailMessage.From = new MailAddress(txtEmail.Text);
                mailMessage.To.Add(toEmail);
                mailMessage.Subject = subject;
                mailMessage.Body = body;

                // Прикрепление файла
                if (File.Exists(attachmentPath))
                {
                    Attachment attachment = new Attachment(attachmentPath);
                    mailMessage.Attachments.Add(attachment);
                }
                // Отправка
                smtpClient.Send(mailMessage);
                                   
                
            }
        }

        private void txtExcelPath_TextChanged(object sender, EventArgs e)
        {

        }

        private bool isValidEmail(string email)
        {
            try
            {
                var emailAddress = new MailAddress(email);
                return true;
            }
            catch (Exception) {
               return false;
            }
           
        }
    }
}
