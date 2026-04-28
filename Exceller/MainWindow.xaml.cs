using Exceller.Models;
using Exceller.Services;
using Microsoft.Win32;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Exceller
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly ExcelService _excelService = new ExcelService();
        public MainWindow()
        {
            InitializeComponent();
        }



        private async void SelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog
            {
                Title = "Выберите папку в которой содержаться основные отчеты"
                ,InitialDirectory = "C:\\"
            };
            if (dialog.ShowDialog() == true)
            {
                string selectedPath = dialog.FolderName;
                LogList.Items.Add($"Выбрана папка: {selectedPath}");

                // Показываем прогресс-бар
                WorkProgress.Visibility = Visibility.Visible;
                WorkProgress.IsIndeterminate = true;

                try
                {
                    // 2. Запускаем обработку в фоновом потоке, чтобы интерфейс не завис
                    await Task.Run(() => ProcessFiles(selectedPath));

                    MessageBox.Show(" Отчет сформирован.", "ура", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}");
                }
                finally
                {
                    WorkProgress.Visibility = Visibility.Collapsed;
                }
            }
        }
        private void ProcessFiles(string folderPath)
        {
            var files = Directory.GetFiles(folderPath, "*.xlsx");
            var allData = new List<ReportData>();

            foreach (var file in files)
            {
                // Проверка, чтобы программа не пыталась прочитать свой же отчет, если он уже создан
                if (file.Contains("~") || file.Contains("Итоговый_Отчет")) continue;

                Dispatcher.Invoke(() => LogList.Items.Add($"Чтение: {System.IO.Path.GetFileName(file)}"));

                var dataFromFile = _excelService.ReadFile(file);

                // Записываем имя файла в каждую строку данных для истории
                dataFromFile.ForEach(x => x.SourceFile = System.IO.Path.GetFileName(file));

                allData.AddRange(dataFromFile);
            }

            // --- ВОТ ЭТА ЧАСТЬ НОВАЯ ---
            if (allData.Count > 0)
            {
                Dispatcher.Invoke(() => LogList.Items.Add("Сохранение итогового файла..."));

                // Вызываем наш новый метод сохранения
                _excelService.SaveDataToNewFile(allData, folderPath);

                Dispatcher.Invoke(() => LogList.Items.Add($"Готово! Файл сохранен в: {folderPath}"));
            }
            else
            {
                Dispatcher.Invoke(() => LogList.Items.Add("Данные для сохранения не найдены."));
            }
        }
    }
}
