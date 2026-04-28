using ExcelDataReader;
using Exceller.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace Exceller.Services
{
    public class ExcelService
    {

        public List<ReportData> ReadFile(string filePath)
        {
            var data = new List<ReportData>();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    //на случай если шапка есть
                    reader.Read();

                    while (reader.Read())
                    {
                        data.Add(new ReportData() { Article = reader.GetValue(0)?.ToString(), //колонка 1
                        Name = reader.GetValue(1)?.ToString(), //колонка 2
                        Quantity =double.TryParse(reader.GetValue(2)?.ToString(),out var q)? q: 0, }); 
                            

                    }
                }
            }
            return data;
        }

        public void SaveDataToNewFile(List<ReportData> data, string folderPath)
        {
            // Формируем путь к новому файлу в той же папке
            string fileName = "Итоговый_Отчет.xlsx";
            string outputPath = Path.Combine(folderPath, fileName);

            // Создаем новую рабочую книгу Excel
            using (var workbook = new XLWorkbook())
            {
                // Добавляем лист
                var worksheet = workbook.Worksheets.Add("Общие данные");

                // 1. Создаем шапку (заголовки колонок)
                worksheet.Cell(1, 1).Value = "Артикул";
                worksheet.Cell(1, 2).Value = "Наименование";
                worksheet.Cell(1, 3).Value = "Количество";
                worksheet.Cell(1, 4).Value = "Источник (файл)";

                // Сделаем шапку жирной для красоты
                worksheet.Range("A1:D1").Style.Font.Bold = true;

                // 2. Заполняем данными
                int currentRow = 2; // Начинаем со второй строки, так как первая — шапка
                foreach (var item in data)
                {
                    worksheet.Cell(currentRow, 1).Value = item.Article;
                    worksheet.Cell(currentRow, 2).Value = item.Name;
                    worksheet.Cell(currentRow, 3).Value = item.Quantity;
                    worksheet.Cell(currentRow, 4).Value = item.SourceFile; // Чтобы знать, откуда данные
                    currentRow++;
                }

                // 3. Автоподбор ширины колонок (чтобы текст не обрезался)
                worksheet.Columns().AdjustToContents();

                // Сохраняем файл
                workbook.SaveAs(outputPath);
            }
        }
    }
}
