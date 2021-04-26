using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using Forms = System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xceed.Words.NET;

namespace sharpRedactor
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<string> gender = new List<string>() { "женщина", "мужчина", "не определился", "боевая ракета земля-воздух" };
        List<string> education = new List<string>() { "Высшее", "Колледж", "Среднее", "Школьник", "Студент", "Нет" };
        string age;
        public MainWindow()
        {
            InitializeComponent();
            cbGender.ItemsSource = gender;
            cbEdu.ItemsSource = education;
            //////////
            checkAge.Checked += (o, e) => { tbAge.IsEnabled = false; ageSlider.IsEnabled = false; tbAge.Text = "0"; };
            checkAge.Unchecked += (o, e) => { tbAge.IsEnabled = true; ageSlider.IsEnabled = true; };
            /////////
            checkEdu.Checked += (o, e) => { cbEdu.IsEnabled = false; cbEdu.SelectedIndex = 5; };
            checkEdu.Unchecked += (o, e) => cbEdu.IsEnabled = true;
            /////////
            checkGender.Checked += (o, e) => { cbGender.IsEnabled = false; cbGender.SelectedIndex = 2; };
            checkGender.Unchecked += (o, e) => cbGender.IsEnabled = true;
        }

        private void bOpen_Click(object sender, RoutedEventArgs e)
        {
            tFiles.Items.Clear();
            try
            {
                //boxText.Text += File.ReadAllText(@"C:\Users\skeych\Desktop\СП_ОТВЕТЫ_ЛАБ_1.docx");
                int ages;
                if (cbEdu.SelectedItem == null || cbGender.SelectedItem == null || !int.TryParse(tbAge.Text, out ages))
                {
                    MessageBox.Show("Выберите критерии", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string[] files = { "" };
                string[] filenames;
                Forms.FolderBrowserDialog folderBrowser = new Forms.FolderBrowserDialog();

                Forms.DialogResult result = folderBrowser.ShowDialog();

                if (!string.IsNullOrWhiteSpace(folderBrowser.SelectedPath))
                {
                    files = Directory.GetFiles(folderBrowser.SelectedPath);
                    
                    
                }
                else return;
                for (int i = 0; i < files.Length; i++)
                {
                    TextBox tb = new TextBox();
                    DirectoryInfo dr = new DirectoryInfo(files[i]);
                    tb.Text = _LoadFiles(files[i], dr.Name, age, cbGender.SelectedItem.ToString(), cbEdu.SelectedItem.ToString()).ToString();
                    if (tb.Text == "No") continue;
                    tFiles.Items.Add(new TabItem
                    {
                        Header = i,
                        Content = tb.Text
                    });

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private string _LoadFiles(string filename, string destination, string age, string gender, string education)
        {
            string content;
            DocX doc = DocX.Load(filename);
            content = '\n' + doc.Text;
            if (age != null || gender != null || education != null)
            {
                for (int i = Convert.ToInt32(tbAge.Text); i < Convert.ToInt32(tbToAge.Text); i++)
                {
                    age = i.ToString();
                    if (content.Contains(age) && content.Contains(gender) && content.Contains(education))
                    {
                        string basedest = AppDomain.CurrentDomain.BaseDirectory + @"\true_resume";
                        Directory.CreateDirectory(basedest);
                        File.Copy(filename, basedest + @"\" + destination, true);
                        return content;
                    }
                }
            }
            return "No";
        }

       


        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void bExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void ageSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            tbAge.Text = ((int)ageSlider.Value).ToString();
        }

        private void tbAge_TextChanged(object sender, TextChangedEventArgs e)
        {
            age = tbAge.Text;
        }

        private void bHide_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }
    }
}
