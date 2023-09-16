using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;
using System.Globalization;
using SpreadsheetLight;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AppJ
{
    internal class Program
    {
        static void Main(string[] args)
        {
            int countI = 1;//координаты ячеек
            int countJ = 1;//координаты ячеек

            SLDocument ws1 = new SLDocument();
            //создание стиля
            SLStyle style1 = ws1.CreateStyle();
            style1.Border.LeftBorder.BorderStyle = BorderStyleValues.Double;
            style1.Border.TopBorder.BorderStyle = BorderStyleValues.Double;
            style1.Border.BottomBorder.BorderStyle = BorderStyleValues.Double;
            style1.Border.RightBorder.BorderStyle = BorderStyleValues.Double;
            ws1.SetCellValue(countI, countJ, "Котировка, месяц");
            ws1.SetCellValue(countI, countJ + 1, "Минимальное значение");
            ws1.SetCellValue(countI, countJ + 2, "Максимальное значение");
            ws1.SetCellValue(countI, countJ + 3, "Медиум раре");
            ws1.SetCellStyle(countI, countJ, style1);
            ws1.SetCellStyle(countI, countJ+1, style1);
            ws1.SetCellStyle(countI, countJ+2, style1);
            ws1.SetCellStyle(countI, countJ+3, style1);
            //пути
            string PathKeyword = "keyword/keyword.txt";
            string FilePath = @"pdf/";
            string excelPath = "excel";

            List<string> ListAllFiles = new List<string>();
            string[] Keywords = File.ReadAllText(PathKeyword).Split(new string[] { "\r\n" }, StringSplitOptions.None);
            string[] DocPaths = System.IO.Directory.GetFiles(FilePath, "*.pdf");


            //Создание списка всех PDF файлов
            string textForEncoding = string.Empty;
            string tempText = string.Empty;
            foreach (string DocPath in DocPaths)
            {
                using (PdfReader reader = new PdfReader(DocPath))
                {
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        textForEncoding = string.Empty;
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        textForEncoding = PdfTextExtractor.GetTextFromPage(reader, i, strategy);
                        textForEncoding = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(textForEncoding)));
                        tempText += textForEncoding;
                    }
                }
                ListAllFiles.Add(tempText);
                tempText = string.Empty;
            }
            //Заполнение Excel документа
            foreach (string currentFile in ListAllFiles)
            {
                if (currentFile[50] == ',')//проверка на число, если число 2 или 21 например
                {
                    //ListForSave.Add(currentFile.Substring(20, 36));
                    ws1.SetCellValue(countI, countJ, currentFile.Substring(20, 36));
                }
                else
                {
                    //ListForSave.Add(currentFile.Substring(20, 37));
                    ws1.SetCellValue(countI, countJ, currentFile.Substring(20, 37));
                }

                ws1.SetCellValue(countI, countJ + 1, "min");
                ws1.SetCellValue(countI, countJ + 2, "max");
                ws1.SetCellValue(countI, countJ + 3, "median");
                ws1.SetCellStyle(countI, countJ, style1);
                ws1.SetCellStyle(countI, countJ + 1, style1);
                ws1.SetCellStyle(countI, countJ + 2, style1);
                ws1.SetCellStyle(countI, countJ + 3, style1);
                countI++;

                foreach (string keyword in Keywords)
                {
                    if (currentFile.ToString().Contains(keyword))
                    {
                        int index = currentFile.ToString().IndexOf(keyword);
                        string read = currentFile.Substring(index + keyword.Length + 1, 13);
                        string[] words = read.Split('–');

                        var number1 = Convert.ToDecimal(words[0], new CultureInfo("en-US"));
                        var number2 = Convert.ToDecimal(words[1], new CultureInfo("en-US"));
                        var median = (number1 + number2) / 2;
                        ws1.SetCellValue(countI, countJ, keyword);
                        ws1.SetCellValue(countI, countJ + 1, number1);
                        ws1.SetCellValue(countI, countJ + 2, number2);
                        ws1.SetCellValue(countI, countJ + 3, median);
                        countI++;
                        //ListForSave.Add($"Ключевое слово {keyword} его котировки равна от {number1} до {number2}. Медиана равна {median}");
                    }
                }
            }

            ws1.AutoFitColumn(0);
            ws1.AutoFitColumn(1);
            ws1.AutoFitColumn(2);
            ws1.AutoFitColumn(3);

            //сохрание документа
            if (!Directory.Exists(excelPath))
            {
                Directory.CreateDirectory(excelPath);
            }
            DateTime currentDateTime = DateTime.Now;
            string nameExcel = currentDateTime.ToString("dddd, dd MMMM yyyy HH mm ss") + ".xlsx";
            try
            {
                ws1.SaveAs($"{excelPath}/{nameExcel}");
            }
            catch (Exception)
            {
                ws1.SaveAs($"{excelPath}/Ошибка Была ошибка.xlsx");
            }
            
            
            Console.WriteLine("Создание документа прошло успешно");
            Console.ReadKey();
        }
    }
}
