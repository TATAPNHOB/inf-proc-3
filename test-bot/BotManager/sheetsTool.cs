using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace BotManager
{
    internal class sheetsTool
    {//В Debug лежит жсончик со всеми нужными токенами, который сгенерен после создания credentials
        //в проекте на клауде
        double L = 50; //отступ слева
        double T = 30; //отступ сверху
        double R = 50; //отступ справа
        double StrBetw = 17; //между строками
        double H = 12; //высота строки
        double headerH; //высота названий слобцов для таблиц
        double tab1rowH = 30; //высота строки с результатами для 1 таблицы
        double tab23rowH = 14; //высота строки с результатами для 2 и 3 таблицы
        double tab45rowH = 20; //высота строки с результатами для 2 и 3 таблицы
        double tab;//доп отступ для таблицы
        XPen pen = new XPen(XColor.FromName("Black"), 0.5);
        double prevlens = 0; //сумма ширин предыдущих колонок
        int num = 1;
        XGraphics g;
        PdfPage page;
        PdfDocument doc;
        int num_of_ions = 4;
        int num_of_sessions = 1;
        int action_count = 136;

        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Legistrators";//произвольно
        static readonly string SpreadsheetID = "1HB4k716lcyBiOtpN_bk4RE7c4ZNK6QVwsYcOgl0HU4w";//в ссылке на лист
        static readonly string[] sheets = new string[3] { "Data", "Timing", "Информация по иону" };//имя листа (не таблицы)
        static SheetsService service;//Объект для работы собсно с листиками


        static readonly XFont TNR11 = new XFont("Times New Roman", 11);
        static readonly XFont TNR16B = new XFont("Times New Roman", 16, XFontStyle.Bold);
        static readonly XFont TNR11B = new XFont("Times New Roman", 11, XFontStyle.Bold);
        static readonly XFont TNR12 = new XFont("Times New Roman", 12);
        static readonly XFont A11 = new XFont("Arial", 11);
        static readonly XFont TNR145B = new XFont("Times New Roman", 14.5, XFontStyle.Bold);
        static readonly XFont TNR14 = new XFont("Times New Roman", 14);
        static readonly XFont A8 = new XFont("Arial", 8);

        internal static void DirCreate (long chatId)
        {
            Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}");
        }
        internal static void Proccess(long chatId, int folderNumb)
        {
            
            GoogleCredential credential;//Права 
            using (FileStream stream = new FileStream("client-secrets.json", FileMode.Open, FileAccess.Read))
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //создание папки пользователя, чтобы именно там держать его пдфки, которые потом должны удалиться!
           
            Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}");
            GenPDFTransitions("1-4", 130, chatId, folderNumb);
            GenPDFSeances("27-33", 130, chatId, folderNumb);


        }
        static void GenPDFTransitions(string action_interval, int action_count, long chatId, long folderNumb)
        {
            List<int> row_num = new List<int>();
            IList<IList<object>> Search_index_list = ReadEntries(sheets[0], SpreadsheetID, $"!A3:A{action_count}"); //столбец с номерами
            int starting_seance;
            int ending_seance;
            if (!action_interval.Contains("-"))
            {
                starting_seance = ending_seance = Convert.ToInt32(action_interval); //пришло одно число
            }
            else
            {
                starting_seance = Convert.ToInt32(action_interval.Split('-')[0]); //диапазон
                ending_seance = Convert.ToInt32(action_interval.Split('-')[1]);
            }

            bool seance_mode = false;
            for (int i = 0; i < Search_index_list.Count; i++)
            {
                if (seance_mode)
                {
                    if (Search_index_list[i][0].ToString() == ending_seance.ToString())
                    {
                        row_num.Add(i + 3);
                        seance_mode = false;
                        break;
                    }
                    else
                    {
                        row_num.Add(i + 3);
                    }
                }
                if (Search_index_list[i][0].ToString() == starting_seance.ToString())
                {
                    seance_mode = true;
                    row_num.Add(i + 3);
                }
            }
            IList<IList<object>> first_sheet_values = ReadEntries(sheets[0], SpreadsheetID, $"!A{row_num[0]}:AH{row_num[0]}"); // из первой таблицы строки
            IList<IList<object>> second_sheet_values = ReadEntries(sheets[1], SpreadsheetID, $"!A{row_num[0]}:R{row_num[0]}"); // из второй
            IList<IList<object>> third_sheet_all_values = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{num_of_ions * num_of_sessions}"); //из третьей (ВСЕ ЗНАЧ)
            int session_num = Convert.ToInt32(first_sheet_values[0][23].ToString().Split('/')[0]);
            int display_count = Convert.ToInt32(first_sheet_values[0][23].ToString().Split('/')[1].Split('-')[0]);
            IList<IList<object>> third_sheet_values = ReadEntries(sheets[0], SpreadsheetID, "!A1:A2"); //из 3
            foreach (IList<object> row in third_sheet_all_values)
            {
                if (Convert.ToInt32(row[2]) == display_count && Convert.ToInt32(row[13]) == session_num)
                { third_sheet_values.Add(row); break; } //определение
            }
            third_sheet_values.RemoveAt(1);
            third_sheet_values.RemoveAt(0);
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            for (int i = 1; i < row_num.Count; i++)
            {
                first_sheet_values.Add(ReadEntries(sheets[0], SpreadsheetID, $"!A{row_num[i]}:AH{row_num[i]}")[0]);
                second_sheet_values.Add(ReadEntries(sheets[1], SpreadsheetID, $"!A{row_num[i]}:R{row_num[i]}")[0]);

                session_num = Convert.ToInt32(first_sheet_values[i][23].ToString().Split('/')[0]);
                display_count = Convert.ToInt32(first_sheet_values[i][23].ToString().Split('/')[1].Split('-')[0]);

                foreach (IList<object> row in third_sheet_all_values)
                {
                    if (Convert.ToInt32(row[2]) == display_count && Convert.ToInt32(row[13]) == session_num)
                    { third_sheet_values.Add(row); break; } //определение
                }
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            }
            //путь к шаблону с пустыми полями
            
PdfDocument protocol3 = PdfSharp.Pdf.IO.PdfReader.Open(@"..\..\..\..\PDFtemplates\prot3.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
            //PdfDocument protocol3 = PdfSharp.Pdf.IO.PdfReader.Open("C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\prot3.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);




            PdfPage page;



            XGraphics g;
            int minus_k = row_num[0] - starting_seance;

            foreach (int cur_seance in row_num)
            {
                PdfDocument doc = new PdfDocument();
                int actual_index = cur_seance - starting_seance - minus_k;
                doc.AddPage(protocol3.Pages[0]);
                page = doc.Pages[0]; //ссылка на пейдж
                g = XGraphics.FromPdfPage(page);
                g.DrawString("TЗЧ/" + second_sheet_values[actual_index][9].ToString().Substring(6, 4) + "-" +
                        third_sheet_values[actual_index][0].ToString() + "-"
                        + third_sheet_values[actual_index][13].ToString() + "/" + third_sheet_values[actual_index][2].ToString()
                        + "-" + (cur_seance - minus_k).ToString("D3"), C_font2, XBrushes.Black,
                new XRect(444, 58.75, 115, 13), XStringFormats.Center); ; //ТЗЧ/...
                g.DrawString(first_sheet_values[actual_index][23].ToString(), C_font1, XBrushes.Black,
                    new XRect(251, 76, 58, 14), XStringFormats.Center); //протокол №
                g.DrawString(DateTime.Today.ToString("d"), C_font1, XBrushes.Black,
                    new XRect(361, 76, 58, 14), XStringFormats.Center); //от
                g.DrawString(third_sheet_values[actual_index][1].ToString(), C_font3, XBrushes.Black,
                    new XRect(449, 111, 39 / 2, 15), XStringFormats.TopCenter); //индекс
                g.DrawString(third_sheet_values[actual_index][0].ToString(), C_font1, XBrushes.Black,
                    new XRect(449 + 39 / 2, 111, 39 / 2, 15), XStringFormats.Center); //ион
                g.DrawString(third_sheet_values[actual_index][6].ToString(), C_font1, XBrushes.Black,
                    new XRect(155, 130.5, 61, 15), XStringFormats.Center); //энергия
                g.DrawString(second_sheet_values[actual_index][9].ToString(), C_font, XBrushes.Black,
                    new XRect(198, 202.5, 101.5, 10.5), XStringFormats.Center); //в период с
                g.DrawString(second_sheet_values[actual_index][10].ToString(), C_font, XBrushes.Black,
                    new XRect(356, 202.5, 101.5, 10.5), XStringFormats.Center); //по

                g.DrawString(first_sheet_values[actual_index][27].ToString(), C_font, XBrushes.Black,
                    new XRect(277.75, 237, 37, 10), XStringFormats.Center); //температура
                g.DrawString(first_sheet_values[actual_index][25].ToString(), C_font, XBrushes.Black,
                    new XRect(215, 252.5, 34, 10), XStringFormats.Center); //давление
                g.DrawString(first_sheet_values[actual_index][26].ToString(), C_font, XBrushes.Black,
                    new XRect(271, 268.5, 34, 10), XStringFormats.Center); //влажность
                g.DrawString(third_sheet_values[actual_index][1].ToString(), C_font4, XBrushes.Black,
                    new XRect(365.5, 439, 69 / 2, 13), XStringFormats.TopRight); //индекс
                g.DrawString(third_sheet_values[actual_index][0].ToString(), C_font, XBrushes.Black,
                    new XRect(365.5 + 69 / 2, 439, 69 / 2, 13), XStringFormats.CenterLeft); //ион
                g.DrawString(first_sheet_values[actual_index][22].ToString(), C_font, XBrushes.Black,
                    new XRect(125, 455.5, 57, 12), XStringFormats.Center); //N
                g.DrawString(first_sheet_values[actual_index][7].ToString(), C_font, XBrushes.Black,
                    new XRect(330, 455.5, 57, 12), XStringFormats.Center); //Ф

                g.DrawString(first_sheet_values[actual_index][8].ToString(), C_font, XBrushes.Black,
                    new XRect(91, 501, 63, 31.5), XStringFormats.Center);//ТД1
                g.DrawString(first_sheet_values[actual_index][9].ToString(), C_font, XBrushes.Black,
                    new XRect(91 + 63, 501, 63, 31.5), XStringFormats.Center);//ТД2
                g.DrawString(first_sheet_values[actual_index][10].ToString(), C_font, XBrushes.Black,
                    new XRect(91 + 63 * 2, 501, 63, 31.5), XStringFormats.Center);//ТД3
                g.DrawString(first_sheet_values[actual_index][11].ToString(), C_font, XBrushes.Black,
                    new XRect(91 + 63 * 3, 501, 63, 31.5), XStringFormats.Center);//ТД4
                g.DrawString(first_sheet_values[actual_index][12].ToString(), C_font, XBrushes.Black,
                    new XRect(91 + 63 * 4, 501, 63, 31.5), XStringFormats.Center);//ТД5

                g.DrawString(first_sheet_values[actual_index][13].ToString(), C_font, XBrushes.Black,
                    new XRect(91, 548, 63, 31.5), XStringFormats.Center);//ТД6
                g.DrawString(first_sheet_values[actual_index][14].ToString(), C_font, XBrushes.Black,
                    new XRect(91 + 63, 548, 63, 31.5), XStringFormats.Center);//ТД7
                g.DrawString(first_sheet_values[actual_index][15].ToString(), C_font, XBrushes.Black,
                    new XRect(91 + 63 * 2, 548, 63, 31.5), XStringFormats.Center);//ТД8
                g.DrawString(first_sheet_values[actual_index][16].ToString(), C_font, XBrushes.Black,
                    new XRect(91 + 63 * 3, 548, 63, 31.5), XStringFormats.Center);//ТД9
                g.DrawString(first_sheet_values[actual_index][7].ToString(), C_font, XBrushes.Black,
                    new XRect(91 + 63 * 4, 548, 63, 31.5), XStringFormats.Center);//Среднее

                g.DrawString(first_sheet_values[actual_index][30].ToString(), C_font, XBrushes.Black,
                    new XRect(303, 598, 45.5, 12.5), XStringFormats.Center);//Красчетный
                g.DrawString(first_sheet_values[actual_index][31].ToString(), C_font, XBrushes.Black,
                    new XRect(369, 598, 45.5, 12.5), XStringFormats.Center);//Кпогр
                g.DrawString(first_sheet_values[actual_index][29].ToString(), C_font, XBrushes.Black,
                    new XRect(344.5, 615.5, 58, 12.5), XStringFormats.Center);//Неоднородность
                g.DrawString(second_sheet_values[actual_index][9].ToString().Substring(11), C_font, XBrushes.Black,
                    new XRect(117.5, 667, 56.5, 10), XStringFormats.Center);//продолжение время

                //
                doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}\TestTransit{actual_index + 1}.pdf") ;

                //doc.Save($"C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\TestTransit{actual_index + 1}.pdf"); //путь, куда сохранять док
                doc.Close();
                protocol3.Close();
            }
        }
        static void GenPDFSeances(string action_interval, int action_count, long chatId, int folderNumb)
        {
            List<int> row_num = new List<int>();
            IList<IList<object>> Search_index_list = ReadEntries(sheets[0], SpreadsheetID, $"!A3:A{action_count}"); //столбец с номерами
            int starting_seance = Convert.ToInt32(action_interval.Split('-')[0]);
            int ending_seance = Convert.ToInt32(action_interval.Split('-')[1]); // что если придет что-то другое?

            bool seance_mode = false;
            for (int i = 0; i < Search_index_list.Count; i++)
            {
                if (seance_mode)
                {
                    if (Search_index_list[i][0].ToString() == ending_seance.ToString())
                    {
                        row_num.Add(i + 3);
                        seance_mode = false;
                        break;
                    }
                    else
                    {
                        row_num.Add(i + 3);
                    }
                }
                if (Search_index_list[i][0].ToString() == starting_seance.ToString())
                {
                    seance_mode = true;
                    row_num.Add(i + 3);
                }
            }
            IList<IList<object>> first_sheet_values = ReadEntries(sheets[0], SpreadsheetID, $"!A{row_num[0]}:AH{row_num[0]}"); // из первой таблицы строки
            IList<IList<object>> second_sheet_values = ReadEntries(sheets[1], SpreadsheetID, $"!A{row_num[0]}:R{row_num[0]}"); // из второй
            IList<IList<object>> third_sheet_all_values = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{num_of_ions * num_of_sessions}"); //из третьей (ВСЕ ЗНАЧ)
            int session_num = Convert.ToInt32(first_sheet_values[0][23].ToString().Split('/')[0]);
            int display_count = Convert.ToInt32(first_sheet_values[0][23].ToString().Split('/')[1].Split('-')[0]);
            IList<IList<object>> third_sheet_values = ReadEntries(sheets[0], SpreadsheetID, "!A1:A2"); //из 3
            foreach (IList<object> row in third_sheet_all_values)
            {
                if (Convert.ToInt32(row[2]) == display_count && Convert.ToInt32(row[13]) == session_num)
                { third_sheet_values.Add(row); break; } //определение
            }
            third_sheet_values.RemoveAt(1);
            third_sheet_values.RemoveAt(0);
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            for (int i = 1; i < row_num.Count; i++)
            {
                first_sheet_values.Add(ReadEntries(sheets[0], SpreadsheetID, $"!A{row_num[i]}:AH{row_num[i]}")[0]);
                second_sheet_values.Add(ReadEntries(sheets[1], SpreadsheetID, $"!A{row_num[i]}:R{row_num[i]}")[0]);

                session_num = Convert.ToInt32(first_sheet_values[i][23].ToString().Split('/')[0]);
                display_count = Convert.ToInt32(first_sheet_values[i][23].ToString().Split('/')[1].Split('-')[0]);

                foreach (IList<object> row in third_sheet_all_values)
                {
                    if (Convert.ToInt32(row[2]) == display_count && Convert.ToInt32(row[13]) == session_num)
                    { third_sheet_values.Add(row); break; } //определение
                }
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            }

            //PdfDocument protocol1 = PdfSharp.Pdf.IO.PdfReader.Open("C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\prot1.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
            //PdfDocument protocol2 = PdfSharp.Pdf.IO.PdfReader.Open("C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\prot2.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);//путь к шаблону с пустыми полями
            PdfDocument protocol1 = PdfSharp.Pdf.IO.PdfReader.Open(@"..\..\..\..\PDFtemplates\prot1.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
            PdfDocument protocol2 = PdfSharp.Pdf.IO.PdfReader.Open(@"..\..\..\..\PDFtemplates\prot2.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);//путь к шаблону с пустыми полями



            PdfPage page;



            XGraphics g;
            int minus_k = row_num[0] - starting_seance;

            foreach (int cur_seance in row_num)
            {
                PdfDocument doc = new PdfDocument();
                
                int actual_index = cur_seance - starting_seance - minus_k;

                // I, J, K == 7,8,9
                if (first_sheet_values[actual_index][10].ToString() != "")
                {
                    doc.AddPage(protocol1.Pages[0]);
                    page = doc.Pages[0]; //ссылка на пейдж
                    g = XGraphics.FromPdfPage(page);
                    //НЕСТАНДАРТНЫЙ ПРОТОКОЛ
                    //1.ОБЩИЕ СВЕДЕНИЯ О СЕАНСЕ
                    g.DrawString(first_sheet_values[actual_index][1].ToString(), US_font, XBrushes.Black,
                        new XRect(49.75, 158, 81, 33.25), XStringFormats.Center); //название организации
                    g.DrawString("ЗНАЧ", US_font, XBrushes.Black,
                        new XRect(49.75 + 81, 158, 119, 33.25), XStringFormats.Center); //Шифр или наименование работы

                    string obluchayemoye = first_sheet_values[actual_index][2].ToString();
                    for (int i = 3; i < 6; i++)
                    {
                        if (first_sheet_values[actual_index][i].ToString() != "") obluchayemoye += "," + first_sheet_values[actual_index][i].ToString();
                    }
                    g.DrawString(obluchayemoye, US_font, XBrushes.Black,
                        new XRect(49.75 + 81 + 119, 158, 178.5, 33.25), XStringFormats.Center); //Облучаемое изделие
                    g.DrawString(second_sheet_values[actual_index][9].ToString(), US_font, XBrushes.Black,
                        new XRect(49.75 + 81 + 119 + 178.5, 158, 118.5, 33.25), XStringFormats.Center); //Время начала облучения
                    g.DrawString(second_sheet_values[actual_index][11].ToString(), US_font, XBrushes.Black,
                        new XRect(49.75 + 81 + 119 + 178.5 + 118.5, 158, 110.5, 33.25), XStringFormats.Center); //Длительность
                                                                                                                //2.УСЛОВИЯ ЭКСПЕРИМЕНТА: В СРЕДЕ ВАКУУМ
                    g.DrawString(first_sheet_values[actual_index][24].ToString(), US_font, XBrushes.Black,
                        new XRect(50, 236.5, 81, 14.85), XStringFormats.Center); //угол
                    g.DrawString(first_sheet_values[actual_index][28].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81, 236.5, 119, 14.85), XStringFormats.Center); //Температура
                    g.DrawString(third_sheet_values[actual_index][3].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 119, 236.5, 119, 14.85), XStringFormats.Center); //материал дегрейдора
                    g.DrawString(third_sheet_values[actual_index][4].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 119 + 119, 236.5, 119, 14.85), XStringFormats.Center); //толщина
                                                                                                   //3.ХАРАКТЕРИСТИКИ ИОНА
                    g.DrawString(third_sheet_values[actual_index][0].ToString(), US_font, XBrushes.Black,
                        new XRect(50, 310, 81 / 2, 14.85), XStringFormats.CenterRight);//тип иона
                    g.DrawString(third_sheet_values[actual_index][1].ToString(), US_font3, XBrushes.Black,
                        new XRect(50 + 81 / 2, 310, 81 / 2, 14.85), XStringFormats.TopLeft);//номер для иона
                    g.DrawString($"{third_sheet_values[actual_index][6]}±{third_sheet_values[actual_index][7]}", US_font, XBrushes.Black,
                        new XRect(50 + 81, 310, 179, 14.85), XStringFormats.Center); //энергия Е на поверхности
                    g.DrawString($"{third_sheet_values[actual_index][8]}±{third_sheet_values[actual_index][9]}", US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179, 310, 119, 14.85), XStringFormats.Center); //пробег R
                    g.DrawString($"{third_sheet_values[actual_index][10]}±{third_sheet_values[actual_index][11]}", US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 + 119, 310, 178.5, 14.85), XStringFormats.Center); //линейные потери энергии ЛПЭ
                                                                                                   //4.ДАННЫЕ ПО ПРОПОРЦИОНАЛЬНЫМ СЧЕТЧИКАМ
                    g.DrawString(first_sheet_values[actual_index][18].ToString(), US_font, XBrushes.Black,
                        new XRect(50, 373, 81, 25.5), XStringFormats.Center); //1
                    g.DrawString(first_sheet_values[actual_index][19].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81, 373, 179 / 3, 25.5), XStringFormats.Center); //2
                    g.DrawString(first_sheet_values[actual_index][20].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 / 3, 373, 179 / 3, 25.5), XStringFormats.Center); //3
                    g.DrawString(first_sheet_values[actual_index][21].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 / 3 + 179 / 3, 373, 179 / 3, 25.5), XStringFormats.Center); //4
                    g.DrawString(first_sheet_values[actual_index][17].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179, 373, 119, 25.5), XStringFormats.Center); //среднее значение
                                                                                          //5.Данные по трековым мембранам из лавсановой пленки:
                    g.DrawString(first_sheet_values[actual_index][8].ToString(), US_font, XBrushes.Black,
                        new XRect(50, 438, 81, 25.5), XStringFormats.Center); //детектор 1
                    g.DrawString(first_sheet_values[actual_index][9].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 2
                    g.DrawString(first_sheet_values[actual_index][10].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 / 3, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 3
                    g.DrawString(first_sheet_values[actual_index][11].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 / 3 + 179 / 3, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 4
                    g.DrawString(first_sheet_values[actual_index][12].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179, 438, 119 / 2, 25.5), XStringFormats.Center); //детектор 5
                    g.DrawString(first_sheet_values[actual_index][13].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 + 119 / 2, 438, 119 / 2, 25.5), XStringFormats.Center); //детектор 6
                    g.DrawString(first_sheet_values[actual_index][30].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 + 119 + 178.5, 438, 95, 25.5), XStringFormats.Center); //неоднородность, по лев
                    g.DrawString(first_sheet_values[actual_index][31].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 + 119 + 178.5 + 95, 438, 95, 25.5), XStringFormats.Center); //неоднородность, по прав

                    g.DrawString((cur_seance - minus_k).ToString("D3"), US_font1, XBrushes.Black,
                        new XRect(753, 97, 18.75, 11.5), XStringFormats.Center); //номер сеанса

                    g.DrawString(first_sheet_values[actual_index][30].ToString(), US_font, XBrushes.Black,
                        new XRect(659, 327, 46.5, 12), XStringFormats.Center); //K
                    g.DrawString(first_sheet_values[actual_index][31].ToString(), US_font, XBrushes.Black,
                        new XRect(748, 327, 46.5, 12), XStringFormats.Center); //K погрешность
                    g.DrawString(first_sheet_values[actual_index][23].ToString(), US_font, XBrushes.Black,
                        new XRect(659, 359, 46.5, 12), XStringFormats.Center); //номер протокола допуска
                    g.DrawString("TЗЧ/" + second_sheet_values[actual_index][9].ToString().Substring(6, 4) + "-" +
                        third_sheet_values[actual_index][0].ToString() + "-"
                        + third_sheet_values[actual_index][13].ToString() + "/" + third_sheet_values[actual_index][2].ToString()
                        + "-" + (cur_seance - minus_k).ToString("D3"), US_font2, XBrushes.Black,
                    new XRect(662, 79, 115, 13), XStringFormats.Center); ; //ТЗЧ/...
                    g.DrawString("Заглушка", US_font, XBrushes.Black,
                        new XRect(660, 385, 46.5, 12), XStringFormats.Center); //K факт
                }
                else
                {
                    doc.AddPage(protocol2.Pages[0]);
                    page = doc.Pages[0]; //ссылка на пейдж
                    g = XGraphics.FromPdfPage(page);
                    //СТАНДАРТНЫЙ ПРОТОКОЛ
                    //1.ОБЩИЕ СВЕДЕНИЯ О СЕАНСЕ
                    g.DrawString(first_sheet_values[actual_index][1].ToString(), S_font, XBrushes.Black,
                        new XRect(66, 160.5, 70, 18), XStringFormats.Center); //название организации
                    g.DrawString("XX-XXX", S_font, XBrushes.Black,
                        new XRect(66 + 70, 160.5, 121, 18), XStringFormats.Center); //Шифр или наименование работы
                    string obluchayemoye = first_sheet_values[0][2].ToString();
                    for (int i = 3; i < 6; i++)
                    {
                        if (first_sheet_values[actual_index][i].ToString() != "") obluchayemoye += "," + first_sheet_values[actual_index][i].ToString();
                    }
                    g.DrawString(obluchayemoye, S_font, XBrushes.Black,
                        new XRect(66 + 70 + 121, 160.5, 182, 18), XStringFormats.Center); //Облучаемое изделие
                    g.DrawString(second_sheet_values[actual_index][9].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 121 + 182, 160.5, 121, 18), XStringFormats.Center); //Время начала облучения
                    g.DrawString(second_sheet_values[actual_index][11].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 121 + 182 + 121, 160.5, 108, 18), XStringFormats.Center);  //Длительность
                                                                                                       //2.УСЛОВИЯ ЭКСПЕРИМЕНТА: В СРЕДЕ ВАКУУМ
                    g.DrawString(first_sheet_values[actual_index][24].ToString(), S_font, XBrushes.Black,
                        new XRect(66, 224, 70, 17), XStringFormats.Center); //угол
                    g.DrawString(first_sheet_values[actual_index][28].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70, 224, 121, 17), XStringFormats.Center); //Температура
                    g.DrawString(third_sheet_values[actual_index][3].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 121, 224, 121, 17), XStringFormats.Center); //материал дегрейдора
                    g.DrawString(third_sheet_values[actual_index][4].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 121 * 2, 224, 121, 17), XStringFormats.Center); //толщина
                                                                                            //3.ХАРАКТЕРИСТИКИ ИОНА
                    g.DrawString(third_sheet_values[actual_index][0].ToString(), S_font, XBrushes.Black,
                        new XRect(66, 298, 70 / 2, 17), XStringFormats.CenterRight);//тип иона
                    g.DrawString(third_sheet_values[actual_index][1].ToString(), S_font3, XBrushes.Black,
                        new XRect(66 + 70 / 2, 299, 70 / 2, 17), XStringFormats.TopLeft);//номер для иона
                    g.DrawString($"{third_sheet_values[actual_index][6]}±{third_sheet_values[actual_index][7]}", S_font, XBrushes.Black,
                        new XRect(66 + 70, 298, 182, 17), XStringFormats.Center); //энергия Е на поверхности
                    g.DrawString($"{third_sheet_values[actual_index][8]}±{third_sheet_values[actual_index][9]}", S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182, 298, 121, 17), XStringFormats.Center); //пробег R
                    g.DrawString($"{third_sheet_values[actual_index][10]}±{third_sheet_values[actual_index][11]}", S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121, 298, 182, 17), XStringFormats.Center); //линейные потери энергии ЛПЭ
                                                                                              //4.ДАННЫЕ ПО ПРОПОРЦИОНАЛЬНЫМ СЧЕТЧИКАМ
                    g.DrawString(first_sheet_values[actual_index][18].ToString(), S_font, XBrushes.Black,
                        new XRect(66, 363, 70, 26), XStringFormats.Center); //1
                    g.DrawString(first_sheet_values[actual_index][19].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70, 363, 182 / 3, 26), XStringFormats.Center); //2
                    g.DrawString(first_sheet_values[actual_index][20].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 / 3, 363, 182 / 3, 26), XStringFormats.Center); //3
                    g.DrawString(first_sheet_values[actual_index][21].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 / 3 + 182 / 3, 363, 182 / 3, 26), XStringFormats.Center); //4
                    g.DrawString(first_sheet_values[actual_index][17].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182, 363, 121, 26), XStringFormats.Center); //среднее значение
                                                                                        //5.Данные по трековым мембранам из лавсановой пленки:
                    g.DrawString(first_sheet_values[actual_index][8].ToString(), S_font, XBrushes.Black,
                        new XRect(66, 427, 70, 26), XStringFormats.Center); //детектор 1
                    g.DrawString(first_sheet_values[actual_index][9].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70, 427, 182 / 3, 26), XStringFormats.Center); //детектор 2
                    g.DrawString(first_sheet_values[actual_index][10].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 3
                    g.DrawString(first_sheet_values[actual_index][11].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 / 3 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 4
                    g.DrawString(first_sheet_values[actual_index][12].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182, 427, 121 / 2, 26), XStringFormats.Center); //детектор 5
                    g.DrawString(first_sheet_values[actual_index][13].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121 / 2, 427, 121 / 2, 26), XStringFormats.Center); //детектор 6
                    g.DrawString(first_sheet_values[actual_index][14].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121, 427, 182 / 3, 26), XStringFormats.Center); //детектор 7
                    g.DrawString(first_sheet_values[actual_index][15].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 8
                    g.DrawString(first_sheet_values[actual_index][16].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121 + 182 / 3 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 9
                    g.DrawString(first_sheet_values[actual_index][29].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121 + 182, 427, 153, 26), XStringFormats.Center); //неоднородность

                    g.DrawString((cur_seance - minus_k).ToString("D3"), S_font1, XBrushes.Black,
                        new XRect(735.5, 98.5, 20, 9.5), XStringFormats.Center); //номер сеанса

                    g.DrawString(first_sheet_values[actual_index][30].ToString(), S_font, XBrushes.Black,
                        new XRect(669.5, 316.5, 44, 12), XStringFormats.Center); //K
                    g.DrawString(first_sheet_values[actual_index][31].ToString(), S_font, XBrushes.Black,
                        new XRect(733.5, 316.5, 44, 12), XStringFormats.Center); //K погрешность
                    g.DrawString(first_sheet_values[actual_index][23].ToString(), S_font, XBrushes.Black,
                        new XRect(659.5, 349, 46.5, 12), XStringFormats.Center); //номер протокола допуска
                    g.DrawString("Заглушка", S_font, XBrushes.Black,
                        new XRect(668, 376, 46.5, 12), XStringFormats.Center); //K факт
                    g.DrawString("TЗЧ/" + second_sheet_values[actual_index][9].ToString().Substring(6, 4) + "-" +
                        third_sheet_values[actual_index][0].ToString() + "-" + third_sheet_values[actual_index][13].ToString() + "/"
                        + third_sheet_values[actual_index][2].ToString() + "-" + (cur_seance - minus_k).ToString("D3"), S_font2, XBrushes.Black,
                        new XRect(667, 77, 115, 13), XStringFormats.Center); ; //ТЗЧ/...
                }
                
                doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}\Test{ actual_index + 1}.pdf"); //путь, куда сохранять док
                //doc.Save($"C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\Test{actual_index + 1}.pdf"); //путь, куда сохранять док
                doc.Close();
                protocol1.Close();
                protocol2.Close();

            }

        }

        static IList<IList<object>> ReadEntries(string sheet, string SpreadsheetID, string Range)
        {
            string range = $"{sheet}{Range}";//Тут любой рендж
            var readRequest = service.Spreadsheets.Values.Get(SpreadsheetID, range);//Повторяется везде, меняется только метод

            var response = readRequest.Execute();//Завершает реквест и передаёт результат для дальнейшей работы
            IList<IList<object>> values = response.Values;
            return values;
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
