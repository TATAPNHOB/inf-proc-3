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
        static int num_of_ions = 4;
        static int num_of_sessions = 1;
        

        static readonly XFont S_font = new XFont("Times New Roman", 11);//Шрифты для стандартного протокола
        static readonly XFont S_font1 = new XFont("Times New Roman", 11, XFontStyle.Bold);
        static readonly XFont S_font2 = new XFont("Times New Roman", 12);
        static readonly XFont S_font3 = new XFont("Times New Roman", 8);

        static readonly XFont US_font = new XFont("Arial", 11);//Для нестандартного
        static readonly XFont US_font1 = new XFont("Times New Roman", 11, XFontStyle.Bold);
        static readonly XFont US_font2 = new XFont("Times New Roman", 14);
        static readonly XFont US_font3 = new XFont("Arial", 8);
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
            GeneratePDF1(60,130);

            
        }

        static void GenPDF(string action_interval, int action_count)
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
                    if (Convert.ToInt32(Search_index_list[i][0]) == ending_seance)
                    {
                        seance_mode = false;
                        break;
                    }
                    else
                    {
                        row_num.Add(i + 3);
                    }
                }
                if (Convert.ToInt32(Search_index_list[i][0]) == starting_seance)
                {
                    seance_mode = true;
                    row_num.Add(i + 3); 
                }
            }
            IList<IList<object>> first_sheet_values = ReadEntries(sheets[0], SpreadsheetID, $"!A{row_num[0]}:AH{row_num[0]}"); // из первой таблицы строки
            IList<IList<object>> second_sheet_values = ReadEntries(sheets[1], SpreadsheetID, $"!A{row_num[0]}:R{row_num[0]}"); // из второй
            IList<IList<object>> third_sheet_all_values = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{num_of_ions * num_of_sessions}"); //из третьей (ВСЕ ЗНАЧ)
            int session_num = Convert.ToInt32(first_sheet_values[0][23].ToString().Split('/')[0]);
            int display_count = Convert.ToInt32(first_sheet_values[0][23].ToString().Split('/')[0].Split('-')[0]);
            IList<object> third_sheet_values = third_sheet_all_values[0]; //из 3
            foreach (IList<object> row in third_sheet_all_values)
            {
                if (Convert.ToInt32(row[2]) == display_count && Convert.ToInt32(row[13]) == session_num)
                { third_sheet_values = row; break; } //определение
            }
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            PdfDocument protocol = PdfSharp.Pdf.IO.PdfReader.Open("C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\prot1.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import); //путь к шаблону с пустыми полями
            PdfDocument doc = new PdfDocument();


            doc.AddPage(protocol.Pages[0]);
            var page = doc.Pages[0];

            XGraphics g = XGraphics.FromPdfPage(page);

            foreach (int cur_seance in row_num)
            {
                doc.AddPage(protocol.Pages[0]);
                page = doc.Pages[cur_seance - starting_seance + 1]; //ссылка на пейдж
                g = XGraphics.FromPdfPage(page);

                // I, J, K == 7,8,9
                if (first_sheet_values[0][9].ToString() != "")
                {
                    //НЕСТАНДАРТНЫЙ ПРОТОКОЛ
                    //1.ОБЩИЕ СВЕДЕНИЯ О СЕАНСЕ
                    g.DrawString(first_sheet_values[0][1].ToString(), US_font, XBrushes.Black,
                        new XRect(49.75, 158, 81, 33.25), XStringFormats.Center); //название организации
                    g.DrawString("ЗНАЧ", US_font, XBrushes.Black,
                        new XRect(49.75 + 81, 158, 119, 33.25), XStringFormats.Center); //Шифр или наименование работы

                    string obluchayemoye = first_sheet_values[0][2].ToString();
                    for (int i = 3; i < 6; i++)
                    {
                        if (first_sheet_values[0][i].ToString() != "") obluchayemoye += "," + first_sheet_values[0][i].ToString();
                    }
                    g.DrawString(obluchayemoye, US_font, XBrushes.Black,
                        new XRect(49.75 + 81 + 119, 158, 178.5, 33.25), XStringFormats.Center); //Облучаемое изделие
                    g.DrawString(second_sheet_values[0][9].ToString(), US_font, XBrushes.Black,
                        new XRect(49.75 + 81 + 119 + 178.5, 158, 118.5, 33.25), XStringFormats.Center); //Время начала облучения
                    g.DrawString(second_sheet_values[0][11].ToString(), US_font, XBrushes.Black,
                        new XRect(49.75 + 81 + 119 + 178.5 + 118.5, 158, 110.5, 33.25), XStringFormats.Center); //Длительность
                                                                                                                //2.УСЛОВИЯ ЭКСПЕРИМЕНТА: В СРЕДЕ ВАКУУМ
                    g.DrawString(first_sheet_values[0][24].ToString(), US_font, XBrushes.Black,
                        new XRect(50, 236.5, 81, 14.85), XStringFormats.Center); //угол
                    g.DrawString(first_sheet_values[0][28].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81, 236.5, 119, 14.85), XStringFormats.Center); //Температура
                    g.DrawString(third_sheet_values[3].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 119, 236.5, 119, 14.85), XStringFormats.Center); //материал дегрейдора
                    g.DrawString(third_sheet_values[4].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 119 + 119, 236.5, 119, 14.85), XStringFormats.Center); //толщина
                                                                                                   //3.ХАРАКТЕРИСТИКИ ИОНА
                    g.DrawString(third_sheet_values[0].ToString(), US_font, XBrushes.Black,
                        new XRect(50, 310, 81 / 2, 14.85), XStringFormats.CenterRight);//тип иона
                    g.DrawString(third_sheet_values[1].ToString(), US_font3, XBrushes.Black,
                        new XRect(50 + 81 / 2, 310, 81 / 2, 14.85), XStringFormats.TopLeft);//номер для иона
                    g.DrawString($"{third_sheet_values[6]}±{third_sheet_values[7]}", US_font, XBrushes.Black,
                        new XRect(50 + 81, 310, 179, 14.85), XStringFormats.Center); //энергия Е на поверхности
                    g.DrawString($"{third_sheet_values[8]}±{third_sheet_values[9]}", US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179, 310, 119, 14.85), XStringFormats.Center); //пробег R
                    g.DrawString($"{third_sheet_values[10]}±{third_sheet_values[11]}", US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 + 119, 310, 178.5, 14.85), XStringFormats.Center); //линейные потери энергии ЛПЭ
                                                                                                   //4.ДАННЫЕ ПО ПРОПОРЦИОНАЛЬНЫМ СЧЕТЧИКАМ
                    g.DrawString(first_sheet_values[0][18].ToString(), US_font, XBrushes.Black,
                        new XRect(50, 373, 81, 25.5), XStringFormats.Center); //1
                    g.DrawString(first_sheet_values[0][19].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81, 373, 179 / 3, 25.5), XStringFormats.Center); //2
                    g.DrawString(first_sheet_values[0][20].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 / 3, 373, 179 / 3, 25.5), XStringFormats.Center); //3
                    g.DrawString(first_sheet_values[0][21].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 / 3 + 179 / 3, 373, 179 / 3, 25.5), XStringFormats.Center); //4
                    g.DrawString(first_sheet_values[0][17].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179, 373, 119, 25.5), XStringFormats.Center); //среднее значение
                                                                                          //5.Данные по трековым мембранам из лавсановой пленки:
                    g.DrawString(first_sheet_values[0][8].ToString(), US_font, XBrushes.Black,
                        new XRect(50, 438, 81, 25.5), XStringFormats.Center); //детектор 1
                    g.DrawString(first_sheet_values[0][9].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 2
                    g.DrawString(first_sheet_values[0][10].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 / 3, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 3
                    g.DrawString(first_sheet_values[0][11].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 / 3 + 179 / 3, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 4
                    g.DrawString(first_sheet_values[0][12].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179, 438, 119 / 2, 25.5), XStringFormats.Center); //детектор 5
                    g.DrawString(first_sheet_values[0][13].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 + 119 / 2, 438, 119 / 2, 25.5), XStringFormats.Center); //детектор 6
                    g.DrawString(first_sheet_values[0][30].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 + 119 + 178.5, 438, 95, 25.5), XStringFormats.Center); //неоднородность, по лев
                    g.DrawString(first_sheet_values[0][31].ToString(), US_font, XBrushes.Black,
                        new XRect(50 + 81 + 179 + 119 + 178.5 + 95, 438, 95, 25.5), XStringFormats.Center); //неоднородность, по прав

                    g.DrawString(action_number.ToString("D3"), US_font1, XBrushes.Black,
                        new XRect(753, 97, 18.75, 11.5), XStringFormats.Center); //номер сеанса
                    g.DrawString(action_number.ToString("D3"), US_font2, XBrushes.Black,
                        new XRect(768.5, 77.5, 22, 13), XStringFormats.Center); //номер сеанса в ТЗЧ ляляля
                    g.DrawString(first_sheet_values[0][30].ToString(), US_font, XBrushes.Black,
                        new XRect(659, 327, 46.5, 12), XStringFormats.Center); //K
                    g.DrawString(first_sheet_values[0][31].ToString(), US_font, XBrushes.Black,
                        new XRect(748, 327, 46.5, 12), XStringFormats.Center); //K погрешность
                    g.DrawString(first_sheet_values[0][23].ToString(), US_font, XBrushes.Black,
                        new XRect(659, 359, 46.5, 12), XStringFormats.Center); //номер протокола допуска
                }
                else
                {
                    //СТАНДАРТНЫЙ ПРОТОКОЛ
                    //1.ОБЩИЕ СВЕДЕНИЯ О СЕАНСЕ
                    g.DrawString(first_sheet_values[0][1].ToString(), S_font, XBrushes.Black,
                        new XRect(66, 160.5, 70, 18), XStringFormats.Center); //название организации
                    g.DrawString("ЗНАЧ", S_font, XBrushes.Black,
                        new XRect(66 + 70, 160.5, 121, 18), XStringFormats.Center); //Шифр или наименование работы
                    string obluchayemoye = first_sheet_values[0][2].ToString();
                    for (int i = 3; i < 6; i++)
                    {
                        if (first_sheet_values[0][i].ToString() != "") obluchayemoye += "," + first_sheet_values[0][i].ToString();
                    }
                    g.DrawString(obluchayemoye, S_font, XBrushes.Black,
                        new XRect(66 + 70 + 121, 160.5, 182, 18), XStringFormats.Center); //Облучаемое изделие
                    g.DrawString(second_sheet_values[0][9].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 121 + 182, 160.5, 121, 18), XStringFormats.Center); //Время начала облучения
                    g.DrawString(second_sheet_values[0][11].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 121 + 182 + 121, 160.5, 108, 18), XStringFormats.Center);  //Длительность
                                                                                                       //2.УСЛОВИЯ ЭКСПЕРИМЕНТА: В СРЕДЕ ВАКУУМ
                    g.DrawString(first_sheet_values[0][24].ToString(), S_font, XBrushes.Black,
                        new XRect(66, 224, 70, 17), XStringFormats.Center); //угол
                    g.DrawString(first_sheet_values[0][28].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70, 224, 121, 17), XStringFormats.Center); //Температура
                    g.DrawString(third_sheet_values[3].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 121, 224, 121, 17), XStringFormats.Center); //материал дегрейдора
                    g.DrawString(third_sheet_values[4].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 121 * 2, 224, 121, 17), XStringFormats.Center); //толщина
                                                                                            //3.ХАРАКТЕРИСТИКИ ИОНА
                    g.DrawString(third_sheet_values[0].ToString(), S_font, XBrushes.Black,
                        new XRect(66, 298, 70 / 2, 17), XStringFormats.CenterRight);//тип иона
                    g.DrawString(third_sheet_values[1].ToString(), S_font3, XBrushes.Black,
                        new XRect(66 + 70 / 2, 299, 70 / 2, 17), XStringFormats.TopLeft);//номер для иона
                    g.DrawString($"{third_sheet_values[6]}±{third_sheet_values[7]}", S_font, XBrushes.Black,
                        new XRect(66 + 70, 298, 182, 17), XStringFormats.Center); //энергия Е на поверхности
                    g.DrawString($"{third_sheet_values[8]}±{third_sheet_values[9]}", S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182, 298, 121, 17), XStringFormats.Center); //пробег R
                    g.DrawString($"{third_sheet_values[10]}±{third_sheet_values[11]}", S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121, 298, 182, 17), XStringFormats.Center); //линейные потери энергии ЛПЭ
                                                                                              //4.ДАННЫЕ ПО ПРОПОРЦИОНАЛЬНЫМ СЧЕТЧИКАМ
                    g.DrawString(first_sheet_values[0][18].ToString(), S_font, XBrushes.Black,
                        new XRect(66, 363, 70, 26), XStringFormats.Center); //1
                    g.DrawString(first_sheet_values[0][19].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70, 363, 182 / 3, 26), XStringFormats.Center); //2
                    g.DrawString(first_sheet_values[0][20].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 / 3, 363, 182 / 3, 26), XStringFormats.Center); //3
                    g.DrawString(first_sheet_values[0][21].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 / 3 + 182 / 3, 363, 182 / 3, 26), XStringFormats.Center); //4
                    g.DrawString(first_sheet_values[0][17].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182, 363, 121, 26), XStringFormats.Center); //среднее значение
                                                                                        //5.Данные по трековым мембранам из лавсановой пленки:
                    g.DrawString(first_sheet_values[0][9].ToString(), S_font, XBrushes.Black,
                        new XRect(66, 427, 70, 26), XStringFormats.Center); //детектор 1
                    g.DrawString(first_sheet_values[0][10].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70, 427, 182 / 3, 26), XStringFormats.Center); //детектор 2
                    g.DrawString(first_sheet_values[0][11].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 3
                    g.DrawString(first_sheet_values[0][12].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 / 3 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 4
                    g.DrawString(first_sheet_values[0][13].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182, 427, 121 / 2, 26), XStringFormats.Center); //детектор 5
                    g.DrawString(first_sheet_values[0][14].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121 / 2, 427, 121 / 2, 26), XStringFormats.Center); //детектор 6
                    g.DrawString(first_sheet_values[0][15].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121, 427, 182 / 3, 26), XStringFormats.Center); //детектор 7
                    g.DrawString(first_sheet_values[0][16].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 8
                    g.DrawString(first_sheet_values[0][17].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121 + 182 / 3 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 9
                    g.DrawString(first_sheet_values[0][27].ToString(), S_font, XBrushes.Black,
                        new XRect(66 + 70 + 182 + 121 + 182, 427, 153, 26), XStringFormats.Center); //неоднородность

                    g.DrawString(action_number.ToString("D3"), S_font1, XBrushes.Black,
                        new XRect(735.5, 98.5, 20, 9.5), XStringFormats.Center); //номер сеанса
                    g.DrawString(action_number.ToString("D3"), S_font2, XBrushes.Black,
                        new XRect(755.5, 80.5, 18, 10), XStringFormats.Center); //номер сеанса в ТЗЧ ляляля
                    g.DrawString(first_sheet_values[0][30].ToString(), S_font, XBrushes.Black,
                        new XRect(669.5, 316.5, 44, 12), XStringFormats.Center); //K
                    g.DrawString(first_sheet_values[0][31].ToString(), S_font, XBrushes.Black,
                        new XRect(733.5, 316.5, 44, 12), XStringFormats.Center); //K погрешность
                    g.DrawString(first_sheet_values[0][23].ToString(), S_font, XBrushes.Black,
                        new XRect(659.5, 349, 46.5, 12), XStringFormats.Center); //номер протокола допуска
                    g.DrawString("Заглушка", S_font, XBrushes.Black,
                        new XRect(668, 376, 46.5, 12), XStringFormats.Center); //K факт
                }


                //добавление новой страницы с шаблоном!!!
            }
            doc.Save("C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\Test.pdf"); //путь, куда сохранять док
        }
        static void GeneratePDF1(int action_number, int action_count) //Нестандартный протокол
        {
            int row_num=0;
            IList<IList<object>> Search_index_list = ReadEntries(sheets[0], SpreadsheetID, $"!A3:A{action_count}"); //столбец с номерами
            for(int i = 0;i<Search_index_list.Count;i++)
            {
                if (Convert.ToString(Search_index_list[i][0])!="-")
                {
                    if(Convert.ToInt32(Search_index_list[i][0]) == action_number)
                    { row_num = i + 3; break; }
                }
            }
            
            IList<IList<object>> first_sheet_values=ReadEntries(sheets[0], SpreadsheetID, $"!A{row_num}:AH{row_num}"); // из первой таблицы строки
            IList<IList<object>> second_sheet_values = ReadEntries(sheets[1], SpreadsheetID, $"!A{row_num}:R{row_num}"); // из второй
            IList<IList<object>> third_sheet_all_values = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{num_of_ions*num_of_sessions}") ; //из третьей (ВСЕ ЗНАЧ)
            int session_num = Convert.ToInt32(first_sheet_values[0][23].ToString().Split('/')[0]);
            int display_count = Convert.ToInt32(first_sheet_values[0][23].ToString().Split('/')[0].Split('-')[0]);
            IList<object> third_sheet_values = third_sheet_all_values[0]; //из 3
            foreach(IList<object> row in third_sheet_all_values)
            {
                if (Convert.ToInt32(row[2]) == display_count && Convert.ToInt32(row[13]) == session_num)
                { third_sheet_values = row; break; } //определение
            }



            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            PdfDocument protocol = PdfSharp.Pdf.IO.PdfReader.Open("C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\prot1.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import); //путь к шаблону с пустыми полями
            PdfDocument doc = new PdfDocument();
            doc.AddPage(protocol.Pages[0]);
            var page = doc.Pages[0];

            XGraphics g = XGraphics.FromPdfPage(page);
            

            //1.ОБЩИЕ СВЕДЕНИЯ О СЕАНСЕ
            g.DrawString(first_sheet_values[0][1].ToString(), US_font, XBrushes.Black,
                new XRect(49.75, 158, 81, 33.25), XStringFormats.Center); //название организации
            g.DrawString("ЗНАЧ", US_font, XBrushes.Black,
                new XRect(49.75 + 81, 158, 119, 33.25), XStringFormats.Center); //Шифр или наименование работы

            string obluchayemoye = first_sheet_values[0][2].ToString();
            for(int i = 3;i<6;i++)
            {
                if (first_sheet_values[0][i].ToString() != "") obluchayemoye += "," + first_sheet_values[0][i].ToString();
            }
            g.DrawString(obluchayemoye, US_font, XBrushes.Black,
                new XRect(49.75 + 81 + 119, 158, 178.5, 33.25), XStringFormats.Center); //Облучаемое изделие
            g.DrawString(second_sheet_values[0][9].ToString(), US_font, XBrushes.Black,
                new XRect(49.75 + 81 + 119 + 178.5, 158, 118.5, 33.25), XStringFormats.Center); //Время начала облучения
            g.DrawString(second_sheet_values[0][11].ToString(), US_font, XBrushes.Black,
                new XRect(49.75 + 81 + 119 + 178.5 + 118.5, 158, 110.5, 33.25), XStringFormats.Center); //Длительность
                                                                                                        //2.УСЛОВИЯ ЭКСПЕРИМЕНТА: В СРЕДЕ ВАКУУМ
            g.DrawString(first_sheet_values[0][24].ToString(), US_font, XBrushes.Black,
                new XRect(50, 236.5, 81, 14.85), XStringFormats.Center); //угол
            g.DrawString(first_sheet_values[0][28].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81, 236.5, 119, 14.85), XStringFormats.Center); //Температура
            g.DrawString(third_sheet_values[3].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 119, 236.5, 119, 14.85), XStringFormats.Center); //материал дегрейдора
            g.DrawString(third_sheet_values[4].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 119 + 119, 236.5, 119, 14.85), XStringFormats.Center); //толщина
                                                                                           //3.ХАРАКТЕРИСТИКИ ИОНА
            g.DrawString(third_sheet_values[0].ToString(), US_font, XBrushes.Black,
                new XRect(50, 310, 81 / 2, 14.85), XStringFormats.CenterRight);//тип иона
            g.DrawString(third_sheet_values[1].ToString(), US_font3, XBrushes.Black,
                new XRect(50 + 81 / 2, 310, 81 / 2, 14.85), XStringFormats.TopLeft);//номер для иона
            g.DrawString($"{third_sheet_values[6]}±{third_sheet_values[7]}", US_font, XBrushes.Black,
                new XRect(50 + 81, 310, 179, 14.85), XStringFormats.Center); //энергия Е на поверхности
            g.DrawString($"{third_sheet_values[8]}±{third_sheet_values[9]}", US_font, XBrushes.Black,
                new XRect(50 + 81 + 179, 310, 119, 14.85), XStringFormats.Center); //пробег R
            g.DrawString($"{third_sheet_values[10]}±{third_sheet_values[11]}", US_font, XBrushes.Black,
                new XRect(50 + 81 + 179 + 119, 310, 178.5, 14.85), XStringFormats.Center); //линейные потери энергии ЛПЭ
                                                                                           //4.ДАННЫЕ ПО ПРОПОРЦИОНАЛЬНЫМ СЧЕТЧИКАМ
            g.DrawString(first_sheet_values[0][18].ToString(), US_font, XBrushes.Black,
                new XRect(50, 373, 81, 25.5), XStringFormats.Center); //1
            g.DrawString(first_sheet_values[0][19].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81, 373, 179 / 3, 25.5), XStringFormats.Center); //2
            g.DrawString(first_sheet_values[0][20].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 179 / 3, 373, 179 / 3, 25.5), XStringFormats.Center); //3
            g.DrawString(first_sheet_values[0][21].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 179 / 3 + 179 / 3, 373, 179 / 3, 25.5), XStringFormats.Center); //4
            g.DrawString(first_sheet_values[0][17].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 179, 373, 119, 25.5), XStringFormats.Center); //среднее значение
                                                                                  //5.Данные по трековым мембранам из лавсановой пленки:
            g.DrawString(first_sheet_values[0][8].ToString(), US_font, XBrushes.Black,
                new XRect(50, 438, 81, 25.5), XStringFormats.Center); //детектор 1
            g.DrawString(first_sheet_values[0][9].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 2
            g.DrawString(first_sheet_values[0][10].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 179 / 3, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 3
            g.DrawString(first_sheet_values[0][11].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 179 / 3 + 179 / 3, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 4
            g.DrawString(first_sheet_values[0][12].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 179, 438, 119 / 2, 25.5), XStringFormats.Center); //детектор 5
            g.DrawString(first_sheet_values[0][13].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 179 + 119 / 2, 438, 119 / 2, 25.5), XStringFormats.Center); //детектор 6
            g.DrawString(first_sheet_values[0][30].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 179 + 119 + 178.5, 438, 95, 25.5), XStringFormats.Center); //неоднородность, по лев
            g.DrawString(first_sheet_values[0][31].ToString(), US_font, XBrushes.Black,
                new XRect(50 + 81 + 179 + 119 + 178.5 + 95, 438, 95, 25.5), XStringFormats.Center); //неоднородность, по прав

            g.DrawString(action_number.ToString("D3"), US_font1, XBrushes.Black,
                new XRect(753, 97, 18.75, 11.5), XStringFormats.Center); //номер сеанса
            g.DrawString(action_number.ToString("D3"), US_font2, XBrushes.Black,
                new XRect(768.5, 77.5, 22, 13), XStringFormats.Center); //номер сеанса в ТЗЧ ляляля
            g.DrawString(first_sheet_values[0][30].ToString(), US_font, XBrushes.Black,
                new XRect(659, 327, 46.5, 12), XStringFormats.Center); //K
            g.DrawString(first_sheet_values[0][31].ToString(), US_font, XBrushes.Black,
                new XRect(748, 327, 46.5, 12), XStringFormats.Center); //K погрешность
            g.DrawString(first_sheet_values[0][23].ToString(), US_font, XBrushes.Black,
                new XRect(659, 359, 46.5, 12), XStringFormats.Center); //номер протокола допуска

            

            doc.Save("C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\Test.pdf"); //путь, куда сохранять док
        }
        static void GeneratePDF2(int action_number,int action_count)//Стандартный протокол
        {
            int row_num = 0;
            IList<IList<object>> Search_index_list = ReadEntries(sheets[0], SpreadsheetID, $"!A3:A{action_count}");
            for (int i = 0; i < Search_index_list.Count; i++)
            {
                if (Convert.ToString(Search_index_list[i][0]) != "-")
                {
                    if (Convert.ToInt32(Search_index_list[i][0]) == action_number)
                    { row_num = i + 3; break; }
                }
            }

            IList<IList<object>> first_sheet_values = ReadEntries(sheets[0], SpreadsheetID, $"!A{row_num}:AH{row_num}");
            IList<IList<object>> second_sheet_values = ReadEntries(sheets[1], SpreadsheetID, $"!A{row_num}:R{row_num}");
            IList<IList<object>> third_sheet_all_values = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{num_of_ions * num_of_sessions}");
            int session_num = Convert.ToInt32(first_sheet_values[0][23].ToString().Split('/')[0]);
            int display_count = Convert.ToInt32(first_sheet_values[0][23].ToString().Split('/')[0].Split('-')[0]);
            IList<object> third_sheet_values = third_sheet_all_values[0];
            foreach (IList<object> row in third_sheet_all_values)
            {
                if (Convert.ToInt32(row[2]) == display_count && Convert.ToInt32(row[13]) == session_num)
                { third_sheet_values = row; break; }
            }
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            PdfDocument protocol = PdfSharp.Pdf.IO.PdfReader.Open("C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\prot2.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import); //путь к шаблону с пустыми полями
            PdfDocument doc = new PdfDocument();
            doc.AddPage(protocol.Pages[0]);
            var page = doc.Pages[0];

            XGraphics g = XGraphics.FromPdfPage(page);
            

            //1.ОБЩИЕ СВЕДЕНИЯ О СЕАНСЕ
            g.DrawString(first_sheet_values[0][1].ToString(), S_font, XBrushes.Black,
                new XRect(66, 160.5, 70, 18), XStringFormats.Center); //название организации
            g.DrawString("ЗНАЧ", S_font, XBrushes.Black,
                new XRect(66 + 70, 160.5, 121, 18), XStringFormats.Center); //Шифр или наименование работы
            string obluchayemoye = first_sheet_values[0][2].ToString();
            for (int i = 3; i < 6; i++)
            {
                if (first_sheet_values[0][i].ToString() != "") obluchayemoye += "," + first_sheet_values[0][i].ToString();
            }
            g.DrawString(obluchayemoye, S_font, XBrushes.Black,
                new XRect(66 + 70 + 121, 160.5, 182, 18), XStringFormats.Center); //Облучаемое изделие
            g.DrawString(second_sheet_values[0][9].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 121 + 182, 160.5, 121, 18), XStringFormats.Center); //Время начала облучения
            g.DrawString(second_sheet_values[0][11].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 121 + 182 + 121, 160.5, 108, 18), XStringFormats.Center);  //Длительность
                                                                                               //2.УСЛОВИЯ ЭКСПЕРИМЕНТА: В СРЕДЕ ВАКУУМ
            g.DrawString(first_sheet_values[0][24].ToString(), S_font, XBrushes.Black,
                new XRect(66, 224, 70, 17), XStringFormats.Center); //угол
            g.DrawString(first_sheet_values[0][28].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70, 224, 121, 17), XStringFormats.Center); //Температура
            g.DrawString(third_sheet_values[3].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 121, 224, 121, 17), XStringFormats.Center); //материал дегрейдора
            g.DrawString(third_sheet_values[4].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 121 * 2, 224, 121, 17), XStringFormats.Center); //толщина
                                                                                    //3.ХАРАКТЕРИСТИКИ ИОНА
            g.DrawString(third_sheet_values[0].ToString(), S_font, XBrushes.Black,
                new XRect(66, 298, 70 / 2, 17), XStringFormats.CenterRight);//тип иона
            g.DrawString(third_sheet_values[1].ToString(), S_font3, XBrushes.Black,
                new XRect(66 + 70 / 2, 299, 70 / 2, 17), XStringFormats.TopLeft);//номер для иона
            g.DrawString($"{third_sheet_values[6]}±{third_sheet_values[7]}", S_font, XBrushes.Black,
                new XRect(66 + 70, 298, 182, 17), XStringFormats.Center); //энергия Е на поверхности
            g.DrawString($"{third_sheet_values[8]}±{third_sheet_values[9]}", S_font, XBrushes.Black,
                new XRect(66 + 70 + 182, 298, 121, 17), XStringFormats.Center); //пробег R
            g.DrawString($"{third_sheet_values[10]}±{third_sheet_values[11]}", S_font, XBrushes.Black,
                new XRect(66 + 70 + 182 + 121, 298, 182, 17), XStringFormats.Center); //линейные потери энергии ЛПЭ
                                                                                      //4.ДАННЫЕ ПО ПРОПОРЦИОНАЛЬНЫМ СЧЕТЧИКАМ
            g.DrawString(first_sheet_values[0][18].ToString(), S_font, XBrushes.Black,
                new XRect(66, 363, 70, 26), XStringFormats.Center); //1
            g.DrawString(first_sheet_values[0][19].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70, 363, 182 / 3, 26), XStringFormats.Center); //2
            g.DrawString(first_sheet_values[0][20].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182 / 3, 363, 182 / 3, 26), XStringFormats.Center); //3
            g.DrawString(first_sheet_values[0][21].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182 / 3 + 182 / 3, 363, 182 / 3, 26), XStringFormats.Center); //4
            g.DrawString(first_sheet_values[0][17].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182, 363, 121, 26), XStringFormats.Center); //среднее значение
                                                                                //5.Данные по трековым мембранам из лавсановой пленки:
            g.DrawString(first_sheet_values[0][9].ToString(), S_font, XBrushes.Black,
                new XRect(66, 427, 70, 26), XStringFormats.Center); //детектор 1
            g.DrawString(first_sheet_values[0][10].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70, 427, 182 / 3, 26), XStringFormats.Center); //детектор 2
            g.DrawString(first_sheet_values[0][11].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 3
            g.DrawString(first_sheet_values[0][12].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182 / 3 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 4
            g.DrawString(first_sheet_values[0][13].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182, 427, 121 / 2, 26), XStringFormats.Center); //детектор 5
            g.DrawString(first_sheet_values[0][14].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182 + 121 / 2, 427, 121 / 2, 26), XStringFormats.Center); //детектор 6
            g.DrawString(first_sheet_values[0][15].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182 + 121, 427, 182 / 3, 26), XStringFormats.Center); //детектор 7
            g.DrawString(first_sheet_values[0][16].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182 + 121 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 8
            g.DrawString(first_sheet_values[0][17].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182 + 121 + 182 / 3 + 182 / 3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 9
            g.DrawString(first_sheet_values[0][27].ToString(), S_font, XBrushes.Black,
                new XRect(66 + 70 + 182 + 121 + 182, 427, 153, 26), XStringFormats.Center); //неоднородность

            g.DrawString(action_number.ToString("D3"), S_font1, XBrushes.Black,
                new XRect(735.5, 98.5, 20, 9.5), XStringFormats.Center); //номер сеанса
            g.DrawString(action_number.ToString("D3"), S_font2, XBrushes.Black,
                new XRect(755.5, 80.5, 18, 10), XStringFormats.Center); //номер сеанса в ТЗЧ ляляля
            g.DrawString(first_sheet_values[0][30].ToString(), S_font, XBrushes.Black,
                new XRect(669.5, 316.5, 44, 12), XStringFormats.Center); //K
            g.DrawString(first_sheet_values[0][31].ToString(), S_font, XBrushes.Black,
                new XRect(733.5, 316.5, 44, 12), XStringFormats.Center); //K погрешность
            g.DrawString(first_sheet_values[0][23].ToString(), S_font, XBrushes.Black,
                new XRect(659.5, 349, 46.5, 12), XStringFormats.Center); //номер протокола допуска
            g.DrawString("Заглушка", S_font, XBrushes.Black,
                new XRect(668, 376, 46.5, 12), XStringFormats.Center); //K факт

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
