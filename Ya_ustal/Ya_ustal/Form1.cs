using System;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System.IO;
using Google.Apis.Sheets.v4.Data;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.Windows.Forms;

namespace Ya_ustal
{
    public partial class Form1 : Form
    {
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
		int action_count = 130;

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
		

		
		public Form1()
        {
            InitializeComponent();
            page = new PdfPage
            {
                Orientation = PdfSharp.PageOrientation.Landscape,
                Rotate = 0
            };
            doc = new PdfDocument();
			doc.Pages.Add(page);
			g = XGraphics.FromPdfPage(page);
        }

		void setheader(string[] rows_of_header, double acolLen, XFont f, double HeaderHeight, double ValueRowHeight)
		{
			for (int i = 0; i < rows_of_header.Length; i++)
			{
				g.DrawLine(pen, L + tab + prevlens, T + num * StrBetw, L + tab + prevlens, T + num * StrBetw + HeaderHeight + ValueRowHeight);
				if (i == 0)
				{
					g.DrawString(rows_of_header[i], f, XBrushes.Black,
					new XRect(L + tab + prevlens, T + num * StrBetw, acolLen, HeaderHeight / rows_of_header.Length), XStringFormats.Center);
				}
				else
				{
					g.DrawString(rows_of_header[i], f, XBrushes.Black,
					new XRect(L + tab + prevlens, T + num * StrBetw + headerH / rows_of_header.Length, acolLen, HeaderHeight / rows_of_header.Length), XStringFormats.Center);
				}
			}
			prevlens += acolLen;
			g.DrawLine(pen, L + tab + prevlens, T + num * StrBetw, L + tab + prevlens, T + num * StrBetw + HeaderHeight + ValueRowHeight);
		}

		void setvalue(string value, double acolLen, XFont f, double ValueRowHeight)
		{
			g.DrawString(value, f, XBrushes.Black,
				new XRect(L + tab + prevlens, T + num * StrBetw, acolLen, ValueRowHeight), XStringFormats.Center);
			prevlens += acolLen;
		}

		void drawTable(double[] lens, string[][] headers, string[] values, XFont Headerf, XFont Valuef, double HeaderHeight, double ValueRowHeight)
		{
			double tab1len = 0;
			foreach (double x in lens)
			{
				tab1len += x;
			}
			tab1len += tab;
			g.DrawLine(pen, L + tab, T + num * StrBetw, L + tab1len, T + num * StrBetw);
			for (int i = 0; i < lens.Length; i++)
			{
				setheader(headers[i], lens[i], Headerf, HeaderHeight, ValueRowHeight);
			}
			T += headerH;
			g.DrawLine(pen, L + tab, T + num * StrBetw, L + tab1len, T + num * StrBetw);
			prevlens = 0;
			for (int i = 0; i < lens.Length; i++)
			{
				setvalue(values[i], lens[i], Valuef, ValueRowHeight);
			}

			T += ValueRowHeight;
			g.DrawLine(pen, L + tab, T + num * StrBetw, L + tab1len, T + num * StrBetw);
			prevlens = 0;
		}

		void addString_startNewLine(string text, XFont f)
		{
			g.DrawString(text, f, XBrushes.Black,
			new XRect(L, T + num * StrBetw, page.Width, H), XStringFormats.CenterLeft); num++;
		}
		void addString_RIGHT(string text, XFont f)
		{
			g.DrawString(text, f, XBrushes.Black,
			new XRect(0, T + num * StrBetw, page.Width - R, H), XStringFormats.CenterRight); num++;
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

        private void button1_Click(object sender, EventArgs e)
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
			GenPDFTransitions("1-4");
			GenPDFSeances("27-33", 130);
		}

        private void GenPDFSeances(string action_interval, int action_count)
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

           




            
            int minus_k = row_num[0] - starting_seance;

            foreach (int cur_seance in row_num)
            {
                int actual_index = cur_seance - starting_seance - minus_k;
                page = new PdfPage
                {
                    Orientation = PdfSharp.PageOrientation.Landscape,
                    Rotate = 0
                };
                doc = new PdfDocument();
                doc.Pages.Add(page);
                g = XGraphics.FromPdfPage(page);
                num = 1;
				tab = 0;
                T = 30;
				headerH = 30;

				// I, J, K == 7,8,9
				if (first_sheet_values[actual_index][10].ToString() != "")
                {
					g.DrawString("Протокол мониторинга характеристик потока ионов сеанса", TNR16B, XBrushes.Black,
				new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center); num++;
					g.DrawString("ТЗЧ/" + "0000" + "-" + "Зн" + "-" + "00" + "/" + "00" + "-" + "000", TNR14, XBrushes.Black,
						new XRect(0, T + num * StrBetw, page.Width - R, H), XStringFormats.CenterRight); num++;
					g.DrawString("Сеанс № " + "000", TNR11, XBrushes.Black,
						new XRect(0, T + num * StrBetw, page.Width - R, H), XStringFormats.CenterRight);

					addString_startNewLine("______________________________________________", TNR145B);
					addString_startNewLine("1. Общие сведения о сеансе:", TNR11B);
					addString_startNewLine("Испытательный ионный комплекс : ИИК 10К-400", TNR11);

					prevlens = 0; //сумма ширин предыдущих колонок
					

					drawTable(new double[] { 80, 120, 180, 120, 120 }, new string[][] { new string[]{ "Название", "организации"},
																				new string[]{ "Шифр или", "наименование работы"},
																				new string[]{ "Облучаемое изделие"},
																				new string[]{ "Время начала", "облучения"},
																				new string[]{ "Длительность"}},
																		new string[] { "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", },
																		TNR11B, A11, headerH, tab1rowH);
					addString_startNewLine("2.Условия эксперимента: в среде " + "ЗНАЧ", TNR11B);

					drawTable(new double[] { 80, 120, 150, 120 }, new string[][]      { new string[]{ "Угол"},
																				new string[]{ "Температура, 25°С"},
																				new string[]{ "Материал дегрейдора"},
																				new string[]{ "Толщина, мкм"}},
																		new string[] { "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ" },
																		TNR11B, A11, headerH, tab23rowH);

					addString_startNewLine("3. Характеристики потока ионов:", TNR11B);
					addString_startNewLine("Характеристики иона:", TNR11);

					drawTable(new double[] { 80, 180, 120, 180 }, new string[][]      { new string[]{ "Тип иона"},
																				new string[]{ "Энергия Е на поверхности, МэВ/н"},
																				new string[]{ "Пробег, R [Si], мкм"},
																				new string[]{ "Линейные потери энергии ЛПЭ", "МэВ×см2/мг [Si]"}},
																		new string[] { "ЗНАЧ" + "00", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ" },
																		TNR11B, A11, headerH, tab23rowH);

					addString_startNewLine("Данные по пропорциональным счетчикам:", TNR11);

					drawTable(new double[] { 80, 60, 60, 60, 120 }, new string[][]    { new string[]{ "1"},
																				new string[]{ "2"},
																				new string[]{ "3"},
																				new string[]{ "4"},
																				new string[]{ "Среднее значение"} },
																		new string[] { "ЗНАЧ" + "00", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ" },
																		TNR11B, A11, headerH, tab45rowH);
					addString_startNewLine("Данные по трековым мембранам из лавсановой пленки:", TNR11);
					drawTable(new double[] { 80, 60, 60, 60, 60, 60, 60, 60, 60, 100, 100 }, new string[][]{ new string[]{ "Детектор 1"},
																									new string[]{ "Детектор 2"},
																									new string[]{ "Детектор 3"},
																									new string[]{ "Детектор 4"},
																									new string[]{ "Детектор 5"},
																									new string[]{ "Детектор 6"},
																									new string[]{ " "},
																									new string[]{ " "},
																									new string[]{ " "},
																									new string[]{ "Неоднородность", "по лев., %"},
																									new string[]{ "Неоднородность", "по прав., %"}},
																											new string[] { "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", " ", " ", " ", "ЗНАЧ", "ЗНАЧ" },
																											TNR11B, A11, headerH, tab45rowH);

					g.DrawString("Ответственный за проведение испытаний в испытательную смену от ООО''НПП''", TNR11, XBrushes.Black,
						new XRect(L, T + num * StrBetw, (page.Width - L) / 2, H), XStringFormats.Center);
					g.DrawString("Технический директор  ООО''НПП''Детектор''", TNR11, XBrushes.Black,
						new XRect(L + (page.Width - L) / 2, T + num * StrBetw, (page.Width - L) / 2, H), XStringFormats.Center); num++;
					g.DrawString("''Детектор''", TNR11, XBrushes.Black,
						new XRect(L, T + num * StrBetw, (page.Width - L) / 2, H), XStringFormats.Center); num++;
					g.DrawString("_________________________________ (                         )", TNR11, XBrushes.Black,
						new XRect(L, T + num * StrBetw, (page.Width - L) / 2, H), XStringFormats.Center);
					g.DrawString("_________________________________ (                         )", TNR11, XBrushes.Black,
						new XRect(L + (page.Width - L) / 2, T + num * StrBetw, (page.Width - L) / 2, H), XStringFormats.Center);


					num -= 5;
					T -= headerH * 2 + tab45rowH + 3;
					addString_RIGHT("Расчетный коэффициент К =     " + "0,00" + "   ±   " + "0,00", TNR11); num++;
					addString_RIGHT("(протокол допуска №     " + "0/0-0" + "    )", TNR11); num++;
					addString_RIGHT("Фактический коэффициент К=     " + "0,00", TNR11); num++;
				}
                else
                {
					g.DrawString("Протокол мониторинга характеристик потока ионов сеанса", TNR16B, XBrushes.Black,
				new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center); num++;
					g.DrawString("ТЗЧ/" + "0000" + "-" + "Зн" + "-" + "00" + "/" + "00" + "-" + "000", TNR12, XBrushes.Black,
						new XRect(0, T + num * StrBetw, page.Width - R, H), XStringFormats.CenterRight); num++;
					g.DrawString("Сеанс № " + "000", TNR11B, XBrushes.Black,
						new XRect(0, T + num * StrBetw, page.Width - R, H), XStringFormats.CenterRight);

					addString_startNewLine("______________________________________________", TNR145B);
					addString_startNewLine("1. Общие сведения о сеансе:", TNR11B);
					addString_startNewLine("Испытательный ионный комплекс : ИИК 10К-400", TNR11B);

					prevlens = 0; //сумма ширин предыдущих колонок
					pen = new XPen(XColor.FromName("Black"), 0.5);

					drawTable(new double[] { 80, 120, 180, 120, 120 }, new string[][] { new string[]{ "Название", "организации"},
																				new string[]{ "Шифр или", "наименование работы"},
																				new string[]{ "Облучаемое изделие"},
																				new string[]{ "Время начала", "облучения"},
																				new string[]{ "Длительность"}},
																		new string[] { "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", },
																		TNR11B, TNR11, headerH, tab1rowH);
					addString_startNewLine("2.Условия эксперимента: в среде " + "ЗНАЧ", TNR11B);

					drawTable(new double[] { 80, 120, 150, 120 }, new string[][]      { new string[]{ "Угол"},
																				new string[]{ "Температура, 25°С"},
																				new string[]{ "Материал дегрейдора"},
																				new string[]{ "Толщина, мкм"}},
																		new string[] { "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ" },
																		TNR11B, TNR11, headerH, tab23rowH);

					addString_startNewLine("3. Характеристики потока ионов:", TNR11B);
					addString_startNewLine("Характеристики иона:", TNR11);

					drawTable(new double[] { 80, 180, 120, 180 }, new string[][]      { new string[]{ "Тип иона"},
																				new string[]{ "Энергия Е на поверхности, МэВ/н"},
																				new string[]{ "Пробег, R [Si], мкм"},
																				new string[]{ "Линейные потери энергии ЛПЭ", "МэВ×см2/мг [Si]"}},
																		new string[] { "ЗНАЧ" + "00", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ" },
																		TNR11B, TNR11, headerH, tab23rowH);

					addString_startNewLine("Данные по пропорциональным счетчикам:", TNR11);

					drawTable(new double[] { 80, 60, 60, 60, 120 }, new string[][]    { new string[]{ "1"},
																				new string[]{ "2"},
																				new string[]{ "3"},
																				new string[]{ "4"},
																				new string[]{ "Среднее значение"} },
																		new string[] { "ЗНАЧ" + "00", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ" },
																		TNR11B, TNR11, headerH, tab45rowH);
					addString_startNewLine("Данные по трековым мембранам из лавсановой пленки:", TNR11);
					drawTable(new double[] { 80, 60, 60, 60, 60, 60, 60, 60, 60, 200 }, new string[][]{ new string[]{ "Детектор 1"},
																									new string[]{ "Детектор 2"},
																									new string[]{ "Детектор 3"},
																									new string[]{ "Детектор 4"},
																									new string[]{ "Детектор 5"},
																									new string[]{ "Детектор 6"},
																									new string[]{ "Детектор 7"},
																									new string[]{ "Детектор 8"},
																									new string[]{ "Детектор 9"},
																									new string[]{ "Неоднородность, %"} },
																											new string[] { "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ" },
																											TNR11B, TNR11, headerH, tab45rowH);

					g.DrawString("Ответственный за проведение испытаний в испытательную смену от ООО''НПП''", TNR11, XBrushes.Black,
						new XRect(L, T + num * StrBetw, (page.Width - L) / 2, H), XStringFormats.Center);
					g.DrawString("Технический директор  ООО''НПП''Детектор''", TNR11, XBrushes.Black,
						new XRect(L + (page.Width - L) / 2, T + num * StrBetw, (page.Width - L) / 2, H), XStringFormats.Center); num++;
					g.DrawString("''Детектор''", TNR11, XBrushes.Black,
						new XRect(L, T + num * StrBetw, (page.Width - L) / 2, H), XStringFormats.Center); num++;
					g.DrawString("_________________________________ (                         )", TNR11, XBrushes.Black,
						new XRect(L, T + num * StrBetw, (page.Width - L) / 2, H), XStringFormats.Center);
					g.DrawString("_________________________________ (                         )", TNR11, XBrushes.Black,
						new XRect(L + (page.Width - L) / 2, T + num * StrBetw, (page.Width - L) / 2, H), XStringFormats.Center);


					num -= 5;
					T -= headerH * 2 + tab45rowH + 3;
					addString_RIGHT("Расчетный коэффициент К =     " + "0,00" + "   ±   " + "0,00", TNR11); num++;
					addString_RIGHT("(протокол допуска №     " + "0/0-0" + "    )", TNR11); num++;
					addString_RIGHT("Фактический коэффициент К=     " + "0,00", TNR11); num++;
				}
                doc.Save($"C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\Test{actual_index + 1}.pdf"); //путь, куда сохранять док


            }

        }
        private void GenPDFTransitions(string action_interval)
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






            
            int minus_k = row_num[0] - starting_seance;

            foreach (int cur_seance in row_num)
            {
				int actual_index = cur_seance - starting_seance - minus_k;
				page = new PdfPage();
				doc = new PdfDocument();
				doc.Pages.Add(page);
				g = XGraphics.FromPdfPage(page);

				headerH = 14;
				tab = 25;
				double tabrowH = 25;
				num = 1;
				T = 30;


				g.DrawString("ТЗЧ/" + "0000" + "-" + "Зн" + "-" + "0/0-0", TNR12, XBrushes.Black,
					new XRect(0, T + num * StrBetw, page.Width - R, H), XStringFormats.CenterRight); num++;
				g.DrawString("Протокол №    " + "0/0-0" + "    от    " + "00.00.0000", TNR14, XBrushes.Black,
					new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center); num += 2;

				g.DrawString("Определения неоднородности флюенса ионов    " + "000" + "  Зн", TNR14, XBrushes.Black,
					new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center); num++;
				g.DrawString("с энергией   " + "0,0" + "   МэВ/N   на испытательном стенде  ИИК 10К-400", TNR14, XBrushes.Black,
					new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center); num += 2;

				addString_startNewLine("1. Цель: Оценка соответствия неоднородности флюенса ионов требованиям заказчика испытаний.", TNR11);
				addString_startNewLine("2. Время и место определения неоднородности флюенса ионов:", TNR11);
				addString_startNewLine("проводилась в период с " + "00.00.0000 00:00:00" + " по " + "00.00.0000 00:00:00" + " в ЛЯР ОИЯИ.", TNR11);
				addString_startNewLine("3. Условия определения неоднородности флюенса ионов:", TNR11);
				addString_startNewLine("        - температура окружающей среды:   " + "00" + "   °С;", TNR11);
				addString_startNewLine("        - атмосферное давление:   " + "000" + "   мм.рт.ст.;", TNR11);
				addString_startNewLine("        - относительная влажность воздуха:   " + "00" + "   %.", TNR11);
				addString_startNewLine("4. Средства определения неоднородности флюенса ионов:", TNR11);
				addString_startNewLine("        - испытательный стенд: ИИК 10К-400;", TNR11);
				addString_startNewLine("        - трековые мембраны (лавсановая плёнка);", TNR11);
				addString_startNewLine("        - установка для травления лавсановой плёнки;", TNR11);
				addString_startNewLine("        - растровый электронный микроскоп ТM-3000 (Hitachi, Япония);", TNR11);
				addString_startNewLine("        - система оцифровки видеосигнала «GALLERY-512».", TNR11);
				addString_startNewLine("5. Методика определения неоднородности флюенса ионов.", TNR11);
				addString_startNewLine("5.1. Проводилась в соответствии с «Методикой измерений флюенса тяжелых заряженных частиц", TNR11);
				addString_startNewLine("        с помощью трековых мембран на основе лавсановой пленки» ЦДКТ1.027.012-2015.", TNR11);
				addString_startNewLine("6.Результаты определения неоднородности флюенса ионов " + "000" + " " + "Зн" + " представлены в таблице 1:", TNR11);
				addString_startNewLine("              N = " + "0,00Е+00" + "        с-1" + "               Ф=        " + "0,00Е+00" + "        частиц*см-2", TNR11);
				num++;
				prevlens = 0; //сумма ширин предыдущих колонок
				pen = new XPen(XColor.FromName("Black"), 0.5);
				drawTable(new double[] { 70, 70, 70, 70, 70 }, new string[][] {     new string[]{ "ТД1"},
																				new string[]{ "ТД2"},
																				new string[]{ "ТД3"},
																				new string[]{ "ТД4"},
																				new string[]{ "ТД5"}},
																	new string[] { "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", },
																	TNR11, TNR11, headerH, tabrowH);
				drawTable(new double[] { 70, 70, 70, 70, 70 }, new string[][] {     new string[]{ "ТД6"},
																				new string[]{ "ТД7"},
																				new string[]{ "ТД8"},
																				new string[]{ "ТД9"},
																				new string[]{ "Среднее зн."}},
																	new string[] { "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", "ЗНАЧ", },
																	TNR11, TNR11, headerH, tabrowH);
				T += 3;
				num++;
				addString_startNewLine("         Коэффициент:           Красчетный  =  " + "0,00" + "   ±   " + "0,00", TNR11);
				addString_startNewLine("         Неоднородность флюенса ионов составила :           " + "00,00" + "   %", TNR11);
				num++;
				addString_startNewLine("7. Принято решение о продолжении работ на ионе                  /  ̶п̶о̶в̶т̶о̶р̶н̶о̶й̶ ̶н̶а̶с̶т̶р̶о̶й̶к̶е̶ ̶п̶у̶ч̶к̶а̶", TNR11);
				addString_startNewLine("        в  " + "00:00:0000", TNR11);
				num++;

				g.DrawString("Ответственный за проведение испытаний", TNR11, XBrushes.Black,
					new XRect(L, T + num * StrBetw, (page.Width - L - R) / 2, H), XStringFormats.Center); num++;
				g.DrawString("в испытательную смену от ООО''НПП''", TNR11, XBrushes.Black,
					new XRect(L, T + num * StrBetw, (page.Width - L - R) / 2, H), XStringFormats.Center);
				g.DrawString("Ответственный за проверку от ЛЯР ОИЯИ", TNR11, XBrushes.Black,
					new XRect(L + (page.Width - L - R) / 2, T + num * StrBetw, (page.Width - L - R) / 2, H), XStringFormats.Center); num++;
				g.DrawString("Детектор''", TNR11, XBrushes.Black,
					new XRect(L, T + num * StrBetw, (page.Width - L - R) / 2, H), XStringFormats.Center); num++;
				g.DrawString("__________________________ (                         )", TNR11, XBrushes.Black,
					new XRect(L, T + num * StrBetw, (page.Width - L - R) / 2, H), XStringFormats.Center);
				g.DrawString("__________________________ (                         )", TNR11, XBrushes.Black,
					new XRect(L + (page.Width - L - R) / 2, T + num * StrBetw, (page.Width - L - R) / 2, H), XStringFormats.Center);



				

				doc.Save($"C:\\Users\\ivanb\\Desktop\\ХАКАТОН\\TestTransit{actual_index + 1}.pdf"); //путь, куда сохранять док


            }
        }

    }
}
