using System;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System.IO;
using Google.Apis.Sheets.v4.Data;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace Console_Get_Set
{
    internal class Program
    {//В Debug лежит жсончик со всеми нужными токенами, который сгенерен после создания credentials
        //в проекте на клауде
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Legistrators";//произвольно
        static readonly string SpreadsheetID = "1HB4k716lcyBiOtpN_bk4RE7c4ZNK6QVwsYcOgl0HU4w";//в ссылке на лист
        static readonly string[] sheets = new string[3] {"Data","Timing","Информация по иону"};//имя листа (не таблицы)
        static SheetsService service;//Объект для работы собсно с листиками
        static void Main(string[] args)
        {
            GoogleCredential credential;//Права 
            using (FileStream stream = new FileStream("client-secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            GeneratePDF(60);

            
        }
        static void GeneratePDF(int action_number)//Нестандартный протокол
        {

            IList<IList<object>> first_sheet_values=ReadEntries(sheets[0], SpreadsheetID, $"!A{action_number+2}:AH{action_number+2}");
            IList<IList<object>> second_sheet_values = ReadEntries(sheets[1], SpreadsheetID, $"!A{action_number+2}:R{action_number+2}");
            IList<IList<object>> third_sheet_values;
            if (second_sheet_values[0][2].ToString() == "Xe")
            {
                third_sheet_values = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O2");
            }
            else if(second_sheet_values[0][2].ToString() == "Kr")
            {
                third_sheet_values = ReadEntries(sheets[2], SpreadsheetID, $"!A3:O3");
            }
            else if(second_sheet_values[0][2].ToString() == "Ar")
            {
                third_sheet_values = ReadEntries(sheets[2], SpreadsheetID, $"!A4:O4");
            }
            else
            {
                third_sheet_values = ReadEntries(sheets[2], SpreadsheetID, $"!A5:O5");
            }
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            PdfDocument protocol = PdfSharp.Pdf.IO.PdfReader.Open("C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\prot1.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import); //путь к шаблону с пустыми полями
            PdfDocument doc = new PdfDocument();
            doc.AddPage(protocol.Pages[0]);
            var page = doc.Pages[0];

            XGraphics g = XGraphics.FromPdfPage(page);
            XFont font = new XFont("Arial", 11);
            XFont font1 = new XFont("Times New Roman", 11, XFontStyle.Bold);
            XFont font2 = new XFont("Times New Roman", 14);

            //1.ОБЩИЕ СВЕДЕНИЯ О СЕАНСЕ
            g.DrawString(first_sheet_values[0][1].ToString(), font, XBrushes.Black,
                new XRect(49.75, 158, 81, 33.25), XStringFormats.Center); //название организации
            g.DrawString("ХЗ", font, XBrushes.Black,
                new XRect(49.75 + 81, 158, 119, 33.25), XStringFormats.Center); //Шифр или наименование работы
            g.DrawString("ХЗ", font, XBrushes.Black,
                new XRect(49.75 + 81 + 119, 158, 178.5, 33.25), XStringFormats.Center); //Облучаемое изделие
            g.DrawString(second_sheet_values[0][9].ToString(), font, XBrushes.Black,
                new XRect(49.75 + 81 + 119 + 178.5, 158, 118.5, 33.25), XStringFormats.Center); //Время начала облучения
            g.DrawString(second_sheet_values[0][11].ToString(), font, XBrushes.Black,
                new XRect(49.75 + 81 + 119 + 178.5 + 118.5, 158, 110.5, 33.25), XStringFormats.Center); //Длительность
                                                                                                        //2.УСЛОВИЯ ЭКСПЕРИМЕНТА: В СРЕДЕ ВАКУУМ
            g.DrawString(first_sheet_values[0][24].ToString(), font, XBrushes.Black,
                new XRect(50, 236.5, 81, 14.85), XStringFormats.Center); //угол
            g.DrawString(first_sheet_values[0][28].ToString(), font, XBrushes.Black,
                new XRect(50 + 81, 236.5, 119, 14.85), XStringFormats.Center); //Температура
            g.DrawString(third_sheet_values[0][3].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 119, 236.5, 119, 14.85), XStringFormats.Center); //материал дегрейдора
            g.DrawString(third_sheet_values[0][4].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 119 + 119, 236.5, 119, 14.85), XStringFormats.Center); //толщина
                                                                                           //3.ХАРАКТЕРИСТИКИ ИОНА
            g.DrawString(third_sheet_values[0][0].ToString()+ third_sheet_values[0][1].ToString(), font, XBrushes.Black,
                new XRect(50, 310, 81, 14.85), XStringFormats.Center); //тип иона
            g.DrawString($"{third_sheet_values[0][6]}±{third_sheet_values[0][7]}", font, XBrushes.Black,
                new XRect(50 + 81, 310, 179, 14.85), XStringFormats.Center); //энергия Е на поверхности
            g.DrawString($"{third_sheet_values[0][8]}±{third_sheet_values[0][9]}", font, XBrushes.Black,
                new XRect(50 + 81 + 179, 310, 119, 14.85), XStringFormats.Center) ; //пробег R
            g.DrawString($"{third_sheet_values[0][10]}±{third_sheet_values[0][11]}", font, XBrushes.Black,
                new XRect(50 + 81 + 179 + 119, 310, 178.5, 14.85), XStringFormats.Center); //линейные потери энергии ЛПЭ
                                                                                           //4.ДАННЫЕ ПО ПРОПОРЦИОНАЛЬНЫМ СЧЕТЧИКАМ
            g.DrawString(first_sheet_values[0][18].ToString(), font, XBrushes.Black,
                new XRect(50, 373, 81, 25.5), XStringFormats.Center); //1
            g.DrawString(first_sheet_values[0][19].ToString(), font, XBrushes.Black,
                new XRect(50 + 81, 373, 179 / 3, 25.5), XStringFormats.Center); //2
            g.DrawString(first_sheet_values[0][20].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 179 / 3, 373, 179 / 3, 25.5), XStringFormats.Center); //3
            g.DrawString(first_sheet_values[0][21].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 179 / 3 + 179 / 3, 373, 179 / 3, 25.5), XStringFormats.Center); //4
            g.DrawString(first_sheet_values[0][17].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 179, 373, 119, 25.5), XStringFormats.Center); //среднее значение
                                                                                  //5.Данные по трековым мембранам из лавсановой пленки:
            g.DrawString(first_sheet_values[0][8].ToString(), font, XBrushes.Black,
                new XRect(50, 438, 81, 25.5), XStringFormats.Center); //детектор 1
            g.DrawString(first_sheet_values[0][9].ToString(), font, XBrushes.Black,
                new XRect(50 + 81, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 2
            g.DrawString(first_sheet_values[0][10].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 179 / 3, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 3
            g.DrawString(first_sheet_values[0][11].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 179 / 3 + 179 / 3, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 4
            g.DrawString(first_sheet_values[0][12].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 179, 438, 119 / 2, 25.5), XStringFormats.Center); //детектор 5
            g.DrawString(first_sheet_values[0][13].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 179 + 119 / 2, 438, 119 / 2, 25.5), XStringFormats.Center); //детектор 6
            g.DrawString(first_sheet_values[0][30].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 179 + 119 + 178.5, 438, 95, 25.5), XStringFormats.Center); //неоднородность, по лев
            g.DrawString(first_sheet_values[0][31].ToString(), font, XBrushes.Black,
                new XRect(50 + 81 + 179 + 119 + 178.5 + 95, 438, 95, 25.5), XStringFormats.Center); //неоднородность, по прав
            
            g.DrawString(action_number.ToString("D3"), font1, XBrushes.Black,
                new XRect(753, 97, 18.75, 11.5), XStringFormats.Center); //номер сеанса
            g.DrawString(action_number.ToString("D3"), font2, XBrushes.Black,
                new XRect(768.5, 77.5, 22, 13), XStringFormats.Center); //номер сеанса в ТЗЧ ляляля
            g.DrawString(first_sheet_values[0][30].ToString(), font, XBrushes.Black,
                new XRect(659, 327, 46.5, 12), XStringFormats.Center); //K
            g.DrawString(first_sheet_values[0][31].ToString(), font, XBrushes.Black,
                new XRect(748, 327, 46.5, 12), XStringFormats.Center); //K погрешность
            g.DrawString(first_sheet_values[0][23].ToString(), font, XBrushes.Black,
                new XRect(659, 359, 46.5, 12), XStringFormats.Center); //номер протокола допуска

            doc.Save("C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\Test.pdf"); //путь, куда сохранять док
        }
        static IList<IList<object>> ReadEntries(string sheet, string SpreadsheetID, string Range)
        {
            string range = $"{sheet}{Range}";//Тут любой рендж
            var readRequest = service.Spreadsheets.Values.Get(SpreadsheetID, range);//Повторяется везде, меняется только метод

            var response = readRequest.Execute();//Завершает реквест и передаёт результат для дальнейшей работы
            IList<IList<object>> values = response.Values;
            return values;
            //if (values != null && values.Count > 0)//Ну тут понятно
            //{
            //    foreach (var row in values)
            //    {
            //        Console.WriteLine("\n");
            //        foreach(string rowitem in row)
            //        {
            //            Console.Write(rowitem+"|");
            //        }
            //        //Вид вывода, нам особо не понадобится, мы в консоли выводить не будем вроде
            //        //Сразу будем пихать в пдф, разве что для понимания - values - лист листов
            //        //Поэтому можно либо без форича обращаться как к двумерному массиву,
            //        //Либо как-то так запихивать вместо вывода запихнув ещё один форич для ввода
            //        //Ну или можно просто сделать ссылку на лист, чего я туплю
            //    }
            //}
            //else
            //{
            //    Console.WriteLine("No data found");
            //}
        }
        static void CreateEntry(string sheet, string SpreadsheetID, string Range)
        {
            string range = $"{sheet}{Range}";//Специфичный ренж, вставляется вниз, поэтому номер строки не нужен
            ValueRange valueRange = new ValueRange();

            List<object> objectList = new List<object>() { "Hello", "Pomogite", "Pls", "Ya", "Hochu", "Umeret" };
            valueRange.Values = new List<IList<object>> { objectList };//Собсно что писать будем

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            //Опция парса, это стандартный парс, числовые переменные так и останутся, а строки могут парсится в даты, числа etc
            var appendResponse = appendRequest.Execute();//Завершает реквест и передаёт результат для дальнейшей работы
        }
        static void UpdateEntry(string sheet, string SpreadsheetID, string Cell)
        {
            string range = $"{sheet}{Cell}";
            ValueRange valueRange = new ValueRange();

            List<object> objectList = new List<object>() { "updated" };
            valueRange.Values = new List<IList<object>> { objectList };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }
        static void DeleteEntry(string sheet, string SpreadsheetID, string Range)
        {
            string range = $"{sheet}{Range}";//Судя по тестам любой ренж, и строка не поднимается
            //Эквиваленто апдейту на пустую строку получается
            ClearValuesRequest requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetID, range);
            var deleteResponse = deleteRequest.Execute();
        }
    }
}
