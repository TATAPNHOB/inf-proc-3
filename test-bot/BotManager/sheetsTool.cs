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
		int ionCount = 4;
		int action_count = 136;

		static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
		static readonly string ApplicationName = "Legistrators";//произвольно
		static readonly string SpreadsheetID = "1HB4k716lcyBiOtpN_bk4RE7c4ZNK6QVwsYcOgl0HU4w";//в ссылке на лист
		static readonly string[] sheets = new string[4] { "Data", "Timing", "Информация по иону", "Системное" };//имя листа (не таблицы)
		static SheetsService service;//Объект для работы собсно с листиками


		static readonly XFont TNR11 = new XFont("Times New Roman", 11);
		static readonly XFont TNR16B = new XFont("Times New Roman", 16, XFontStyle.Bold);
		static readonly XFont TNR11B = new XFont("Times New Roman", 11, XFontStyle.Bold);
		static readonly XFont TNR12 = new XFont("Times New Roman", 12);
		static readonly XFont A11 = new XFont("Arial", 11);
		static readonly XFont TNR145B = new XFont("Times New Roman", 14.5, XFontStyle.Bold);
		static readonly XFont TNR14 = new XFont("Times New Roman", 14);
		static readonly XFont A8 = new XFont("Arial", 8);


		internal void DirCreate(long chatId)
        {
			Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}");
		}

		internal void Proccess()
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
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

		}
		internal void fillDataScam(string d0, string d1,
string t2, string t3, string t5, string t6, string t7, string t8,
string t9)
		{
			string t4 = "-";
			string d6 = ReadEntries(sheets[3], SpreadsheetID, $"!D2:D2")[0][0].ToString();
			string codeProtocol = ReadEntries(sheets[3], SpreadsheetID, $"!E2:E2")[0][0].ToString();
			IList<IList<object>> table = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{ionCount + 1}");

			//d6
			string temp = "";
			string ionName = "";
			for (int i = 0; i < ionCount; i++)
			{
				if (table[i][14].ToString() == d6)
				{
					temp = table[i][2].ToString();
					ionName = table[i][0].ToString();
					break;
				}

			}



			IList<IList<object>> table3 = ReadEntries(sheets[3], SpreadsheetID, $"!A2:C2");

			List<object> objectList = new List<object>() { d0,d1, "-", "-","-",
"-", d6,"-" ,"-" , "-", "-","-",
"-", "-", "-", "-","-","-","-","-","-","-","-", codeProtocol,"-",
"-","-","-",
"-"};
			CreateEntry(sheets[0], SpreadsheetID, $"!A:AH", objectList);

			List<object> objectList2 = new List<object>() { d0, d1, ionName, "", "", "","", "",
"", t2, t3, t4, t5, t6, "", t7, t8, t9};
			CreateEntry(sheets[1], SpreadsheetID, $"!A:R", objectList2);
			UpdateEntry(sheets[3], SpreadsheetID, "!E2", codeProtocol);
		}



		internal void fillDataSmena(string d0, string d1, string d2, string d3, string d4,
string d5, string d7, string d8, string d9, string d10,
string d11, string d12, string d13, string d14,
string d15, string d16, string d17, string d18,
string d19, string d20, string d21, string d22,
string t2, string t3, string t4, string t5, string t6, string t7, string t8,
string t9
)
		{
			string d6 = ReadEntries(sheets[3], SpreadsheetID, $"!D2:D2")[0][0].ToString();
			IList<IList<object>> table = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{ionCount + 1}");
			IList<IList<object>> table2 = ReadEntries(sheets[0], SpreadsheetID, $"!A2:AH{action_count + 1}");

			int TavgTD = 0;
			if (d7 != "") TavgTD++;
			if (d8 != "") TavgTD++;
			if (d9 != "") TavgTD++;
			if (d10 != "") TavgTD++;
			if (d11 != "") TavgTD++;
			if (d12 != "") TavgTD++;
			if (d13 != "") TavgTD++;
			if (d14 != "") TavgTD++;
			if (d15 != "") TavgTD++;

			double averageTD = 0;
			if (TavgTD != 0)
				averageTD = (Convert.ToDouble(d7) + Convert.ToDouble(d8) + Convert.ToDouble(d9)
				+ Convert.ToDouble(d10) + Convert.ToDouble(d11)
				+ Convert.ToDouble(d12) + Convert.ToDouble(d13)
				+ Convert.ToDouble(d14) + Convert.ToDouble(d15)) / TavgTD;

			int TavgOD = 0;

			if (d16 != "") TavgOD++;
			if (d17 != "") TavgOD++;
			if (d18 != "") TavgOD++;
			if (d19 != "") TavgOD++;

			double averageOD = 0;
			if (TavgOD != 0)
				averageOD = (Convert.ToDouble(d16) + Convert.ToDouble(d17) + Convert.ToDouble(d18)
				+ Convert.ToDouble(d19)) / TavgOD;

			//d6
			string temp = "";
			string ionName = "";
			for (int i = 0; i < ionCount; i++)
			{
				if (table[i][14].ToString() == d6)
				{
					temp = table[i][2].ToString();
					ionName = table[i][0].ToString();
					break;
				}

			}
			double iter = 1;
			for (int i = 0; i < action_count; i++)
			{
				if (table2[i][23].ToString() == "") continue;
				int t = table2[i][23].ToString().IndexOf("-");

				if (table2[i][23].ToString().Substring(0, t) == table[ionCount - 1][13].ToString() + "/" + temp)
				{
					if (Convert.ToDouble(table2[i][23].ToString().Substring(t + 1, table2[i][23].ToString().Length - t - 1))
					> iter)
						iter = Convert.ToDouble(table2[i][23].ToString().Substring(t + 1, table2[i][23].ToString().Length - t - 1));
				}

			}
			iter++;
			string codeProtocol = table[ionCount - 1][13].ToString() + "/" + temp + "-" + iter;

			IList<IList<object>> table3 = ReadEntries(sheets[3], SpreadsheetID, $"!A2:C2");

			List<object> objectList = new List<object>() { d0,d1, d2, d3,d4,
d5, d6,averageTD ,d7, d8, d9,d10,
d11, d12, d13, d14,d15,averageOD,d16,d17,d18,d19,d20, codeProtocol,d21,
table3[0][0].ToString(),table3[0][1].ToString(),table3[0][2].ToString(),
d22};






			CreateEntry(sheets[0], SpreadsheetID, $"!A:AH", objectList);

			List<object> objectList2 = new List<object>() { d0, d1, ionName, "", "", "","", "",
"", t2, t3, t4, t5, t6, "", t7, t8, t9};
			CreateEntry(sheets[1], SpreadsheetID, $"!A:R", objectList2);
			UpdateEntry(sheets[3], SpreadsheetID, "!E2", codeProtocol);
		}



		internal void fillDataPerehod(string d0, string d1, string d2, string d3, string d4,
string d5, string d6, string d7, string d8, string d9, string d10,
string d11, string d12, string d13, string d14,
string d15, string d16, string d17, string d18,
string d19, string d20, string d21, string d22,
string t2, string t3, string t4, string t5, string t6, string t7, string t8,
string t9
)
		{
			IList<IList<object>> table = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{ionCount + 1}");
			IList<IList<object>> table2 = ReadEntries(sheets[0], SpreadsheetID, $"!A2:AH{action_count + 1}");

			int TavgTD = 0;
			if (d7 != "") TavgTD++;
			if (d8 != "") TavgTD++;
			if (d9 != "") TavgTD++;
			if (d10 != "") TavgTD++;
			if (d11 != "") TavgTD++;
			if (d12 != "") TavgTD++;
			if (d13 != "") TavgTD++;
			if (d14 != "") TavgTD++;
			if (d15 != "") TavgTD++;

			double averageTD = 0;
			if (TavgTD != 0)
				averageTD = (Convert.ToDouble(d7) + Convert.ToDouble(d8) + Convert.ToDouble(d9)
				+ Convert.ToDouble(d10) + Convert.ToDouble(d11)
				+ Convert.ToDouble(d12) + Convert.ToDouble(d13)
				+ Convert.ToDouble(d14) + Convert.ToDouble(d15)) / TavgTD;

			int TavgOD = 0;

			if (d16 != "") TavgOD++;
			if (d17 != "") TavgOD++;
			if (d18 != "") TavgOD++;
			if (d19 != "") TavgOD++;

			double averageOD = 0;
			if (TavgOD != 0)
				averageOD = (Convert.ToDouble(d16) + Convert.ToDouble(d17) + Convert.ToDouble(d18)
				+ Convert.ToDouble(d19)) / TavgOD;

			//d6
			string temp = "";
			string ionName = "";
			for (int i = 0; i < ionCount; i++)
			{
				if (table[i][14].ToString() == d6)
				{
					temp = table[i][2].ToString();
					ionName = table[i][0].ToString();
					break;
				}

			}
			double iter = 1;
			for (int i = 0; i < action_count; i++)
			{
				if (table2[i][23].ToString() == "") continue;
				int t = table2[i][23].ToString().IndexOf("-");

				if (table2[i][23].ToString().Substring(0, t) == table[ionCount - 1][13].ToString() + "/" + temp)
				{
					if (Convert.ToDouble(table2[i][23].ToString().Substring(t + 1, table2[i][23].ToString().Length - t - 1))
					> iter)
						iter = Convert.ToDouble(table2[i][23].ToString().Substring(t + 1, table2[i][23].ToString().Length - t - 1));
				}

			}
			iter++;
			string codeProtocol = table[ionCount - 1][13].ToString() + "/" + temp + "-" + iter;

			IList<IList<object>> table3 = ReadEntries(sheets[3], SpreadsheetID, $"!A2:C2");

			List<object> objectList = new List<object>() { d0,d1, d2, d3,d4,
d5, d6,averageTD ,d7, d8, d9,d10,
d11, d12, d13, d14,d15,averageOD,d16,d17,d18,d19,d20, codeProtocol,d21,
table3[0][0].ToString(),table3[0][1].ToString(),table3[0][2].ToString(),
d22};






			CreateEntry(sheets[0], SpreadsheetID, $"!A:AH", objectList);

			List<object> objectList2 = new List<object>() { d0, d1, ionName, "", "", "","", "",
"", t2, t3, t4, t5, t6, "", t7, t8, t9};
			CreateEntry(sheets[1], SpreadsheetID, $"!A:R", objectList2);
			UpdateEntry(sheets[3], SpreadsheetID, "!D2", ionName);
			UpdateEntry(sheets[3], SpreadsheetID, "!E2", codeProtocol);
		}



		internal void fillDataSeance(string d0, string d1, string d2, string d3, string d4,
		string d5, string d7, string d8, string d9, string d10,
		string d11, string d12, string d13, string d14,
		string d15, string d16, string d17, string d18,
		string d19, string d20, string d22,
		 string t2, string t3, string t4, string t5, string t6, string t7, string t8,
		string t9
		)
		{


			int TavgTD = 0;
			if (d7 != "") TavgTD++;
			if (d8 != "") TavgTD++;
			if (d9 != "") TavgTD++;
			if (d10 != "") TavgTD++;
			if (d11 != "") TavgTD++;
			if (d12 != "") TavgTD++;
			if (d13 != "") TavgTD++;
			if (d14 != "") TavgTD++;
			if (d15 != "") TavgTD++;

			double averageTD = 0;
			if (TavgTD != 0)
				averageTD = (Convert.ToDouble(d7) + Convert.ToDouble(d8) + Convert.ToDouble(d9)
				+ Convert.ToDouble(d10) + Convert.ToDouble(d11)
				+ Convert.ToDouble(d12) + Convert.ToDouble(d13)
				+ Convert.ToDouble(d14) + Convert.ToDouble(d15)) / TavgTD;
			string d21 = ReadEntries(sheets[3], SpreadsheetID, $"!E2:E2")[0][0].ToString();
			int TavgOD = 0;

			if (d16 != "") TavgOD++;
			if (d17 != "") TavgOD++;
			if (d18 != "") TavgOD++;
			if (d19 != "") TavgOD++;

			double averageOD = 0;
			if (TavgOD != 0)
				averageOD = (Convert.ToDouble(d16) + Convert.ToDouble(d17) + Convert.ToDouble(d18)
				+ Convert.ToDouble(d19)) / TavgOD;



			IList<IList<object>> table3 = ReadEntries(sheets[3], SpreadsheetID, $"!A2:E2");

			List<object> objectList = new List<object>() { d0,d1, d2, d3,d4,
d5,table3[0][0],averageTD ,d7, d8, d9,d10,
d11, d12, d13, d14,d15,averageOD,d16,d17,d18,d19,d20,d21,table3[0][1],
table3[0][0].ToString(),table3[0][1].ToString(),table3[0][2].ToString(),
d22};



			IList<IList<object>> table = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{ionCount + 1}");
			string ionName = "";
			for (int i = 0; i < ionCount; i++)
			{
				if (table[i][14].ToString() == table3[0][4].ToString())
				{
					ionName = table[i][3].ToString();
					break;
				}

			}

			CreateEntry(sheets[0], SpreadsheetID, $"!A:AH", objectList);

			List<object> objectList2 = new List<object>() { d0, d1, ionName, "", "", "","", "",
"", t2, t3, t4, t5, t6, "", t7, t8, t9};
			CreateEntry(sheets[1], SpreadsheetID, $"!A:R", objectList2);
		}

		//Заполнение Информации по иону
		internal void fillIonInfo(string d1, string d2, string d3, string d4,
			string d5, string d6, string d7, string d8, string d9, string d10,
			string d11, string d12, string d13, string d14)
		{//d13

			int iter = 1;
			IList<IList<object>> table = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{ionCount + 1}");

			for (int i = 0; i < ionCount; i++)
			{
				if (table[i][13].ToString() == d13) iter++;
			}

			List<object> objectList = new List<object>() { d1, d2, iter, d3,d4,
			d5, d6, d7, d8, d9,d10,
			d11, d12, d13, d14};

			CreateEntry(sheets[2], SpreadsheetID, $"!A:O", objectList);
			ionCount++;
		}

		internal void fillSystemsheet1(string pressure, string vlazh, string temp)
		{


			UpdateEntry(sheets[3], SpreadsheetID, $"!A2", pressure);
			UpdateEntry(sheets[3], SpreadsheetID, $"!B2", vlazh);
			UpdateEntry(sheets[3], SpreadsheetID, $"!C2", temp);
		}
		internal void fillSystemsheet2(string iontype, string prot)
		{
			List<object> objectList = new List<object>() { iontype, prot };
			UpdateEntry(sheets[3], SpreadsheetID, $"!D2", iontype);
			UpdateEntry(sheets[3], SpreadsheetID, $"!E2", prot);
		}


		internal void PDFsGEN_trans (string input_s, long chatId, int foldNumber)
        {
			DirCreate(chatId);
			Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{foldNumber}");
			GenPDFTransitions(input_s, chatId, foldNumber);
		}
		internal void PDFsGEN_sean(string input_s, long chatId, int foldNumber)
		{
			DirCreate(chatId);
			Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{foldNumber}");
			GenPDFSeances(input_s, chatId, foldNumber);
		}

		string fromSecToStr(double inputValue)
		{
			return Convert.ToInt32(Math.Truncate(inputValue / 3600)).ToString("d2") + ":"
				+ Convert.ToInt32((Math.Truncate(inputValue / 60) - Math.Truncate(inputValue / 3600) * 60)).ToString("d2")
				+ ":" + Convert.ToInt32((inputValue - Math.Truncate(inputValue / 60) * 60)).ToString("d2");
		}

		internal void Request1(string companyName, long chatId, int folderNumb)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			//столбец 1
			IList<IList<object>> CompanyNames = ReadEntries(sheets[1], SpreadsheetID, $"!B2:B{action_count + 1}");
			//столбец 2
			IList<IList<object>> IonName = ReadEntries(sheets[1], SpreadsheetID, $"!C2:C{action_count + 1}");
			//столбец 12
			IList<IList<object>> TimeT = ReadEntries(sheets[1], SpreadsheetID, $"!M2:M{action_count + 1}");

			//это словарь с ионом:ключ - имя, значение -  его суммарне временя
			Dictionary<string, double> IonCount = new Dictionary<string, double>();


			for (int i = 0; i < action_count; i++)
			{
				if (Convert.ToString(CompanyNames[i][0]) == companyName)
				{
					if (!IonCount.ContainsKey(Convert.ToString(IonName[i][0])))
					{
						IonCount.Add(Convert.ToString(IonName[i][0]), 0);
					}
					IonCount[Convert.ToString(IonName[i][0])] += TimeSpan.Parse(Convert.ToString(TimeT[i][0])).TotalSeconds;
					string a = fromSecToStr(IonCount[Convert.ToString(IonName[i][0])]);
				}

			}


			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);


			H = 9; //высота строки


			int num = 1;
			var doc = new PdfDocument();



			List<string> Ions = new List<string>();
			List<string> Time = new List<string>();

			foreach (var item in IonCount)
			{
				Ions.Add(item.Key);
				Time.Add(fromSecToStr(item.Value));
			}
			int N = Ions.Count;
			int for_1_page = 40;
			for (int i = 0; i < N; i += for_1_page)
			{
				var page = new PdfPage();
				doc.Pages.Add(page);
				XGraphics g = XGraphics.FromPdfPage(page);
				g.DrawString("Название компании: " + companyName, TNR11B, XBrushes.Black,
					new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center);
				num++;

				g.DrawString("Ион", TNR11B, XBrushes.Black,
					new XRect(L, T + num * StrBetw, (page.Width - L - R) / 2, H), XStringFormats.Center);
				g.DrawString("Время работы", TNR11B, XBrushes.Black,
					new XRect(L + (page.Width - L - R) / 2, T + num * StrBetw, (page.Width - L - R) / 2, H), XStringFormats.Center);
				num++;
				if (i + for_1_page > N)
				{
					for_1_page = N - i;
				}
				for (int j = i; j < i + for_1_page; j++)
				{
					g.DrawString(Ions[j], A8, XBrushes.Black,
					new XRect(L, T + num * StrBetw, (page.Width - L - R) / 2, H), XStringFormats.Center);
					g.DrawString(Time[j], A8, XBrushes.Black,
						new XRect(L + (page.Width - L - R) / 2, T + num * StrBetw, (page.Width - L - R) / 2, H), XStringFormats.Center);
					num++;
				}
				num = 0;
			}


			Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}");
			doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}\1.pdf"); //путь, куда сохранять док
			doc.Close();

		}


		internal void Request2(string companyName, long chatId, int folderNumb)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			//столбец 1
			IList<IList<object>> CompanyNames = ReadEntries(sheets[1], SpreadsheetID, $"!B2:B{action_count + 1}");
			//столбец 9
			IList<IList<object>> startTime = ReadEntries(sheets[1], SpreadsheetID, $"!J2:J{action_count + 1}");



			string result = "";

			for (int i = 0; i < action_count; i++)
			{
				if (Convert.ToString(CompanyNames[i][0]) == companyName)
				{

					result = Convert.ToString(startTime[i][0]);
					break;

				}

			}

			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

			H = 15; //высота строки
			XPen pen = new XPen(XColor.FromName("Black"), 0.5);
			int num = 0;
			var doc = new PdfDocument();
			var page = new PdfPage();
			doc.Pages.Add(page);
			XGraphics g = XGraphics.FromPdfPage(page);

			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString("Название компании:", TNR11B, XBrushes.Black,
				new XRect(L, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			g.DrawString("Время начала работ:", TNR11B, XBrushes.Black,
				new XRect(L + (page.Width - L - R) / 2, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString(companyName, A8, XBrushes.Black,
				new XRect(L, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			g.DrawString(result, A8, XBrushes.Black,
				new XRect(L + (page.Width - L - R) / 2, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);

			g.DrawLine(pen, L, T, L, T + H * 2);
			g.DrawLine(pen, L + (page.Width - L - R) / 2, T, L + (page.Width - L - R) / 2, T + H * 2);
			g.DrawLine(pen, L + (page.Width - L - R), T, L + (page.Width - L - R), T + H * 2);

			Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}");
			doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}\2.pdf"); //путь, куда сохранять док
			doc.Close();


		}

		internal void Request3(string ionName, string isotope, string sessionNumber, long chatId, int folderNumb)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			IList<IList<object>> table = ReadEntries(sheets[2], SpreadsheetID, $"!A2:O{num_of_ions * num_of_sessions + 1}");

			int choice = -1;

			for (int i = 0; i < num_of_ions * num_of_sessions; i++)
			{
				if (Convert.ToString(table[i][0]) == ionName && Convert.ToString(table[i][13]) == sessionNumber
					&& Convert.ToString(table[i][1]) == isotope)
				{
					choice = i;
					break;
				}
			}
			string result_Type = ionName;
			string result_Energy = Convert.ToString(table[choice][6]);
			string result_Si = Convert.ToString(table[choice][8]);


			H = 15;
			int num = 0;
			XPen pen = new XPen(XColor.FromName("Black"), 0.5);
			var doc = new PdfDocument();
			var page = new PdfPage();
			doc.Pages.Add(page);
			XGraphics g = XGraphics.FromPdfPage(page);

			double len = (page.Width - L - R) / 5;

			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString("Тип иона:", TNR11B, XBrushes.Black,
				new XRect(L, T + num * H, len, H), XStringFormats.Center);
			g.DrawString("Изотоп:", TNR11B, XBrushes.Black,
				new XRect(L + len, T + num * H, len, H), XStringFormats.Center);
			g.DrawString("№ сессии в году:", TNR11B, XBrushes.Black,
				new XRect(L + len * 2, T + num * H, len, H), XStringFormats.Center);
			g.DrawString("Энергия:", TNR11B, XBrushes.Black,
				new XRect(L + len * 3, T + num * H, len, H), XStringFormats.Center);
			g.DrawString("Пробег в кремнии:", TNR11B, XBrushes.Black,
				new XRect(L + len * 4, T + num * H, len, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString(result_Type, TNR11, XBrushes.Black,
				new XRect(L, T + num * H, len, H), XStringFormats.Center);
			g.DrawString(isotope, TNR11, XBrushes.Black,
				new XRect(L + len, T + num * H, len, H), XStringFormats.Center);
			g.DrawString(sessionNumber, TNR11, XBrushes.Black,
				new XRect(L + len * 2, T + num * H, len, H), XStringFormats.Center);
			g.DrawString(result_Energy, TNR11, XBrushes.Black,
				new XRect(L + len * 3, T + num * H, len, H), XStringFormats.Center);
			g.DrawString(result_Si, TNR11, XBrushes.Black,
				new XRect(L + len * 4, T + num * H, len, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);

			g.DrawLine(pen, L, T, L, T + H * 2);
			g.DrawLine(pen, L + len, T, L + len, T + H * 2);
			g.DrawLine(pen, L + len * 2, T, L + len * 2, T + H * 2);
			g.DrawLine(pen, L + len * 3, T, L + len * 3, T + H * 2);
			g.DrawLine(pen, L + len * 4, T, L + len * 4, T + H * 2);
			g.DrawLine(pen, L + len * 5, T, L + len * 5, T + H * 2);


			Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}");
			doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}\3.pdf"); //путь, куда сохранять док
			doc.Close();
		}

		internal void Request4(string ionName, long chatId, int folderNumb)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			IList<IList<object>> IonNames = ReadEntries(sheets[1], SpreadsheetID, $"!C2:C{action_count + 1}");
			IList<IList<object>> CompanyNames = ReadEntries(sheets[1], SpreadsheetID, $"!B2:B{action_count + 1}");
			IList<IList<object>> TimeT = ReadEntries(sheets[1], SpreadsheetID, $"!N2:N{action_count + 1}");

			List<string> exceptions = new List<string>() { "Переход", "Смена", "Детектор", "Простой" };
			Dictionary<string, double> WorkedTime = new Dictionary<string, double>();


			for (int i = 0; i < action_count; i++)
			{
				bool check = true;
				foreach (string exc in exceptions)
				{
					if (Convert.ToString(CompanyNames[i][0]).Contains(exc)) check = false;

				}
				if (!check) continue;

				if (!WorkedTime.ContainsKey(Convert.ToString(CompanyNames[i][0])))
				{
					WorkedTime.Add(Convert.ToString(CompanyNames[i][0]), 0);
				}
				WorkedTime[Convert.ToString(CompanyNames[i][0])]
					+= TimeSpan.Parse(Convert.ToString(TimeT[i][0])).TotalSeconds;
				string a = fromSecToStr(WorkedTime[Convert.ToString(CompanyNames[i][0])]);
			}
			//


			int num = 0;
			var doc = new PdfDocument();
			XPen pen = new XPen(XColor.FromName("Black"), 0.5);
			var page = new PdfPage();
			doc.Pages.Add(page);
			XGraphics g = XGraphics.FromPdfPage(page);
			H = 13;


			List<string> Ions = new List<string>();
			List<string> Time = new List<string>();
			foreach (var item in WorkedTime)
			{
				Ions.Add(item.Key);
				Time.Add(fromSecToStr(item.Value));
			}
			int N = Ions.Count;


			g.DrawString("Тип иона: " + ionName, TNR11B, XBrushes.Black,
				new XRect(0, T + num * H, page.Width, H), XStringFormats.Center);
			num++;

			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString("Название компании", TNR11B, XBrushes.Black,
				new XRect(L, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			g.DrawString("Время", TNR11B, XBrushes.Black,
				new XRect(L + (page.Width - L - R) / 2, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);

			for (int i = 0; i < N; i++)
			{
				g.DrawString(Ions[i], A8, XBrushes.Black,
				new XRect(L, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
				g.DrawString(Time[i], A8, XBrushes.Black,
					new XRect(L + (page.Width - L - R) / 2, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
				num++;
				g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			}
			num = 0;
			g.DrawLine(pen, L, T + H, L, T + H * N + H * 2);
			g.DrawLine(pen, L + (page.Width - L - R) / 2, T + H, L + (page.Width - L - R) / 2, T + H * N + H * 2);
			g.DrawLine(pen, L + (page.Width - L - R), T + H, L + (page.Width - L - R), T + H * N + H * 2);



			Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}");
			doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}\4.pdf"); //путь, куда сохранять док
			doc.Close();
		}

		internal void Request5(string ionName, long chatId, int folderNumb)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			IList<IList<object>> sessionType = ReadEntries(sheets[1], SpreadsheetID, $"!B2:B{action_count + 1}");
			IList<IList<object>> ionType = ReadEntries(sheets[1], SpreadsheetID, $"!C2:C{action_count + 1}");
			IList<IList<object>> seanceTime = ReadEntries(sheets[1], SpreadsheetID, $"!M2:M{action_count + 1}");
			IList<IList<object>> seanceTimeWithTB = ReadEntries(sheets[1], SpreadsheetID, $"!N2:N{action_count + 1}");

			//Dictionary<string, double> IonTime = new Dictionary<string, double>();
			double res = 0;

			for (int i = 0; i < action_count; i++)
			{
				/*if (i == 110) {
					int d = 5;
				}*/
				/*if (!IonTime.ContainsKey(Convert.ToString(ionType[i][0])))
				{
					IonTime.Add(Convert.ToString(ionType[i][0]), 0);
				}*/
				if (ionType[i][0].ToString() != ionName) continue;

				res += -TimeSpan.Parse(Convert.ToString(seanceTimeWithTB[i][0])).TotalSeconds + TimeSpan.Parse(Convert.ToString(seanceTime[i][0])).TotalSeconds;

				if (Convert.ToString(sessionType[i][0]) == "Простой")
				{
					res += TimeSpan.Parse(Convert.ToString(seanceTimeWithTB[i][0])).TotalSeconds;
				}
			}


			H = 15; //высота строки
			XPen pen = new XPen(XColor.FromName("Black"), 0.5);
			int num = 0;
			var doc = new PdfDocument();
			var page = new PdfPage();
			doc.Pages.Add(page);
			XGraphics g = XGraphics.FromPdfPage(page);

			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString("Тип иона:", TNR11B, XBrushes.Black,
				new XRect(L, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			g.DrawString("Время:", TNR11B, XBrushes.Black,
				new XRect(L + (page.Width - L - R) / 2, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString(ionName, A8, XBrushes.Black,
				new XRect(L, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			g.DrawString(fromSecToStr(res), A8, XBrushes.Black,
				new XRect(L + (page.Width - L - R) / 2, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);

			g.DrawLine(pen, L, T, L, T + H * 2);
			g.DrawLine(pen, L + (page.Width - L - R) / 2, T, L + (page.Width - L - R) / 2, T + H * 2);
			g.DrawLine(pen, L + (page.Width - L - R), T, L + (page.Width - L - R), T + H * 2);



			Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}");
			doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}\5.pdf"); //путь, куда сохранять док
			doc.Close();
		}
		internal void Request6( long chatId, int folderNumb)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			IList<IList<object>> sessionNumb = ReadEntries(sheets[0], SpreadsheetID, $"!A{action_count + 1}:A{action_count + 1}");
			IList<IList<object>> sessionStatus = ReadEntries(sheets[0], SpreadsheetID, $"!B{action_count + 1}:B{action_count + 1}");
			string res_numb = sessionNumb[0][0].ToString();
			string res_Status = sessionStatus[0][0].ToString();
			H = 15; //высота строки
			XPen pen = new XPen(XColor.FromName("Black"), 0.5);
			int num = 0;
			var doc = new PdfDocument();
			var page = new PdfPage();
			doc.Pages.Add(page);
			XGraphics g = XGraphics.FromPdfPage(page);

			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString("№ сеанса:", TNR11B, XBrushes.Black,
				new XRect(L, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			g.DrawString("Статус:", TNR11B, XBrushes.Black,
				new XRect(L + (page.Width - L - R) / 2, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString(res_numb, A8, XBrushes.Black,
				new XRect(L, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			g.DrawString(res_Status, A8, XBrushes.Black,
				new XRect(L + (page.Width - L - R) / 2, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);

			g.DrawLine(pen, L, T, L, T + H * 2);
			g.DrawLine(pen, L + (page.Width - L - R) / 2, T, L + (page.Width - L - R) / 2, T + H * 2);
			g.DrawLine(pen, L + (page.Width - L - R), T, L + (page.Width - L - R), T + H * 2);

			Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}");
			doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}\6.pdf"); //путь, куда сохранять док
			doc.Close();
		}
		internal void Request7(long chatId, int folderNumb)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			IList<IList<object>> sessionNumb = ReadEntries(sheets[1], SpreadsheetID, $"!A{action_count + 1}:A{action_count + 1}");
			IList<IList<object>> sessionStart = ReadEntries(sheets[1], SpreadsheetID, $"!J{action_count + 1}:J{action_count + 1}");
			string res_numb = sessionNumb[0][0].ToString();
			string res_Status = sessionStart[0][0].ToString();

			H = 15; //высота строки
			XPen pen = new XPen(XColor.FromName("Black"), 0.5);
			int num = 0;
			var doc = new PdfDocument();
			var page = new PdfPage();
			doc.Pages.Add(page);
			XGraphics g = XGraphics.FromPdfPage(page);



			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString("№ сеанса:", TNR11B, XBrushes.Black,
				new XRect(L, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			g.DrawString("Время начала:", TNR11B, XBrushes.Black,
				new XRect(L + (page.Width - L - R) / 2, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);
			g.DrawString(res_numb, A8, XBrushes.Black,
				new XRect(L, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			g.DrawString(res_Status, A8, XBrushes.Black,
				new XRect(L + (page.Width - L - R) / 2, T + num * H, (page.Width - L - R) / 2, H), XStringFormats.Center);
			num++;
			g.DrawLine(pen, L, T + num * H, L + (page.Width - L - R), T + num * H);

			g.DrawLine(pen, L, T, L, T + H * 2);
			g.DrawLine(pen, L + (page.Width - L - R) / 2, T, L + (page.Width - L - R) / 2, T + H * 2);
			g.DrawLine(pen, L + (page.Width - L - R), T, L + (page.Width - L - R), T + H * 2);
			Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}");
			doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{folderNumb}\7.pdf"); //путь, куда сохранять док
			doc.Close();
		}





        
        

		static void CreateEntry(string sheet, string SpreadsheetID, string Range, List<object> objectList)
		{

			string range = $"{sheet}{Range}";//Специфичный ренж, вставляется вниз, поэтому номер строки не нужен
			ValueRange valueRange = new ValueRange();


			valueRange.Values = new List<IList<object>> { objectList };//Собсно что писать будем

			var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);
			appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
			//Опция парса, это стандартный парс, числовые переменные так и останутся, а строки могут парсится в даты, числа etc
			var appendResponse = appendRequest.Execute();//Завершает реквест и передаёт результат для дальнейшей работы
		}




		static void UpdateEntry(string sheet, string SpreadsheetID, string Cell, string input)
		{
			string range = $"{sheet}{Cell}";
			ValueRange valueRange = new ValueRange();

			List<object> objectList = new List<object>() { input };
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



		private void GenPDFSeances(string action_interval, long chatId, int foldNumber)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
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
				string obluchayemoye = first_sheet_values[actual_index][2].ToString();
				for (int i = 3; i < 6; i++)
				{
					if (first_sheet_values[actual_index][i].ToString() != "") obluchayemoye += "," + first_sheet_values[actual_index][i].ToString();
				}

				// I, J, K == 7,8,9
				if (first_sheet_values[actual_index][10].ToString() != "")
				{
					g.DrawString("Протокол мониторинга характеристик потока ионов сеанса", TNR16B, XBrushes.Black,
				new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center); num++;
					g.DrawString("ТЗЧ/" + second_sheet_values[actual_index][9].ToString().Substring(6, 4) + "-" +
						third_sheet_values[actual_index][0].ToString() + "-"
						+ third_sheet_values[actual_index][13].ToString() + "/" + third_sheet_values[actual_index][2].ToString()
						+ "-" + (cur_seance - minus_k).ToString(), TNR14, XBrushes.Black,
						new XRect(0, T + num * StrBetw, page.Width - R, H), XStringFormats.CenterRight); num++;
					g.DrawString("Сеанс № " + (cur_seance - minus_k).ToString(), TNR11, XBrushes.Black,
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
																		new string[] { first_sheet_values[actual_index][1].ToString()
																		, "XX-XXXX"
																		, obluchayemoye
																		, second_sheet_values[actual_index][9].ToString()
																		, second_sheet_values[actual_index][11].ToString() },
																		TNR11B, A11, headerH, tab1rowH);
					addString_startNewLine("2.Условия эксперимента: в среде " + third_sheet_values[actual_index][5].ToString(), TNR11B);

					drawTable(new double[] { 80, 120, 150, 120 }, new string[][]      { new string[]{ "Угол"},
																				new string[]{ "Температура,°С"},
																				new string[]{ "Материал дегрейдора"},
																				new string[]{ "Толщина, мкм"}},
																		new string[] { first_sheet_values[actual_index][24].ToString()
																		, first_sheet_values[actual_index][28].ToString()
																		, third_sheet_values[actual_index][3].ToString()
																		, third_sheet_values[actual_index][4].ToString() },
																		TNR11B, A11, headerH, tab23rowH);

					addString_startNewLine("3. Характеристики потока ионов:", TNR11B);
					addString_startNewLine("Характеристики иона:", TNR11);

					drawTable(new double[] { 80, 180, 120, 180 }, new string[][]      { new string[]{ "Тип иона"},
																				new string[]{ "Энергия Е на поверхности, МэВ/н"},
																				new string[]{ "Пробег, R [Si], мкм"},
																				new string[]{ "Линейные потери энергии ЛПЭ", "МэВ×см2/мг [Si]"}},
																		new string[] { third_sheet_values[actual_index][0].ToString()
																		+ third_sheet_values[actual_index][1].ToString()
																		, $"{third_sheet_values[actual_index][6]}±{third_sheet_values[actual_index][7]}"
																		, $"{third_sheet_values[actual_index][8]}±{third_sheet_values[actual_index][9]}"
																		, $"{third_sheet_values[actual_index][10]}±{third_sheet_values[actual_index][11]}" },
																		TNR11B, A11, headerH, tab23rowH);

					addString_startNewLine("Данные по пропорциональным счетчикам:", TNR11);

					drawTable(new double[] { 80, 60, 60, 60, 120 }, new string[][]    { new string[]{ "1"},
																				new string[]{ "2"},
																				new string[]{ "3"},
																				new string[]{ "4"},
																				new string[]{ "Среднее значение"} },
																		new string[] { first_sheet_values[actual_index][18].ToString()
																		, first_sheet_values[actual_index][19].ToString()
																		, first_sheet_values[actual_index][20].ToString()
																		, first_sheet_values[actual_index][21].ToString()
																		, first_sheet_values[actual_index][17].ToString() },
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
																											new string[] { first_sheet_values[actual_index][8].ToString()
																											, first_sheet_values[actual_index][9].ToString()
																											, first_sheet_values[actual_index][10].ToString()
																											, first_sheet_values[actual_index][11].ToString()
																											, first_sheet_values[actual_index][12].ToString()
																											, first_sheet_values[actual_index][13].ToString()
																											, " ", " ", " "
																											, first_sheet_values[actual_index][32].ToString()
																											, first_sheet_values[actual_index][33].ToString() },
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
					addString_RIGHT("Расчетный коэффициент К =     " + first_sheet_values[actual_index][30].ToString() + "   ±   " + first_sheet_values[actual_index][31].ToString(), TNR11); num++;
					addString_RIGHT("(протокол допуска №     " + first_sheet_values[actual_index][23].ToString() + "    )", TNR11); num++;
					addString_RIGHT("Фактический коэффициент К=     " + "А ГДЕ ФОРМУЛА, А?", TNR11); num++;
				}
				else
				{
					g.DrawString("Протокол мониторинга характеристик потока ионов сеанса", TNR16B, XBrushes.Black,
				new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center); num++;
					g.DrawString("ТЗЧ/" + second_sheet_values[actual_index][9].ToString().Substring(6, 4) + "-" +
						third_sheet_values[actual_index][0].ToString() + "-"
						+ third_sheet_values[actual_index][13].ToString() + "/" + third_sheet_values[actual_index][2].ToString()
						+ "-" + (cur_seance - minus_k).ToString(), TNR12, XBrushes.Black,
						new XRect(0, T + num * StrBetw, page.Width - R, H), XStringFormats.CenterRight); num++;
					g.DrawString("Сеанс № " + (cur_seance - minus_k).ToString(), TNR11B, XBrushes.Black,
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
																		new string[] { first_sheet_values[actual_index][1].ToString()
																		, "XX-XXX"
																		, obluchayemoye
																		, second_sheet_values[actual_index][9].ToString()
																		, second_sheet_values[actual_index][11].ToString() },
																		TNR11B, TNR11, headerH, tab1rowH);
					addString_startNewLine("2.Условия эксперимента: в среде " + third_sheet_values[actual_index][5].ToString(), TNR11B);

					drawTable(new double[] { 80, 120, 150, 120 }, new string[][]      { new string[]{ "Угол"},
																				new string[]{ "Температура, 25°С"},
																				new string[]{ "Материал дегрейдора"},
																				new string[]{ "Толщина, мкм"}},
																		new string[] { first_sheet_values[actual_index][24].ToString()
																		, first_sheet_values[actual_index][28].ToString()
																		, third_sheet_values[actual_index][3].ToString()
																		, third_sheet_values[actual_index][4].ToString() },
																		TNR11B, TNR11, headerH, tab23rowH);

					addString_startNewLine("3. Характеристики потока ионов:", TNR11B);
					addString_startNewLine("Характеристики иона:", TNR11);

					drawTable(new double[] { 80, 180, 120, 180 }, new string[][]      { new string[]{ "Тип иона"},
																				new string[]{ "Энергия Е на поверхности, МэВ/н"},
																				new string[]{ "Пробег, R [Si], мкм"},
																				new string[]{ "Линейные потери энергии ЛПЭ", "МэВ×см2/мг [Si]"}},
																		new string[] { third_sheet_values[actual_index][0].ToString() +
																		third_sheet_values[actual_index][1].ToString()
																		, $"{third_sheet_values[actual_index][6]}±{third_sheet_values[actual_index][7]}"
																		, $"{third_sheet_values[actual_index][8]}±{third_sheet_values[actual_index][9]}"
																		, $"{third_sheet_values[actual_index][10]}±{third_sheet_values[actual_index][11]}" },
																		TNR11B, TNR11, headerH, tab23rowH);

					addString_startNewLine("Данные по пропорциональным счетчикам:", TNR11);

					drawTable(new double[] { 80, 60, 60, 60, 120 }, new string[][]    { new string[]{ "1"},
																				new string[]{ "2"},
																				new string[]{ "3"},
																				new string[]{ "4"},
																				new string[]{ "Среднее значение"} },
																		new string[] { first_sheet_values[actual_index][18].ToString(),
																			first_sheet_values[actual_index][19].ToString(),
																			first_sheet_values[actual_index][20].ToString(),
																			first_sheet_values[actual_index][21].ToString(),
																			first_sheet_values[actual_index][17].ToString() },
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
																											new string[] {
																											 first_sheet_values[actual_index][8].ToString()
																											, first_sheet_values[actual_index][9].ToString()
																											, first_sheet_values[actual_index][10].ToString()
																											, first_sheet_values[actual_index][11].ToString()
																											, first_sheet_values[actual_index][12].ToString()
																											, first_sheet_values[actual_index][13].ToString()
																											, first_sheet_values[actual_index][14].ToString()
																											, first_sheet_values[actual_index][15].ToString()
																											, first_sheet_values[actual_index][16].ToString(),
																											first_sheet_values[actual_index][29].ToString()},
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
					addString_RIGHT("Расчетный коэффициент К =     " + first_sheet_values[actual_index][30].ToString() + "   ±   " + first_sheet_values[actual_index][31].ToString(), TNR11); num++;
					addString_RIGHT("(протокол допуска №     " + first_sheet_values[actual_index][23].ToString() + "    )", TNR11); num++;
					addString_RIGHT("Фактический коэффициент К=     " + "А ГДЕ ФОРМУЛА", TNR11); num++;
				}

				Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{foldNumber}");
				doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{foldNumber}\Test{ actual_index + 1}.pdf"); //путь, куда сохранять док
				doc.Close();
				



			}

		}
		private void GenPDFTransitions(string action_interval, long chatId, int foldNumber)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
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


				g.DrawString("TЗЧ/" + second_sheet_values[actual_index][9].ToString().Substring(6, 4) + "-" +
						third_sheet_values[actual_index][0].ToString() + "-"
						+ first_sheet_values[actual_index][23].ToString(), TNR12, XBrushes.Black,
					new XRect(0, T + num * StrBetw, page.Width - R, H), XStringFormats.CenterRight); num++;
				g.DrawString("Протокол №    " + first_sheet_values[actual_index][23].ToString() + "    от    " + DateTime.Today.ToString("d"), TNR14, XBrushes.Black,
					new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center); num += 2;

				g.DrawString("Определения неоднородности флюенса ионов    " + third_sheet_values[actual_index][1].ToString() + third_sheet_values[actual_index][0].ToString(), TNR14, XBrushes.Black,
					new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center); num++;
				g.DrawString("с энергией   " + third_sheet_values[actual_index][6].ToString() + "   МэВ/N   на испытательном стенде  ИИК 10К-400", TNR14, XBrushes.Black,
					new XRect(0, T + num * StrBetw, page.Width, H), XStringFormats.Center); num += 2;

				addString_startNewLine("1. Цель: Оценка соответствия неоднородности флюенса ионов требованиям заказчика испытаний.", TNR11);
				addString_startNewLine("2. Время и место определения неоднородности флюенса ионов:", TNR11);
				addString_startNewLine("проводилась в период с " + second_sheet_values[actual_index][9].ToString() + " по " + second_sheet_values[actual_index][10].ToString() + " в ЛЯР ОИЯИ.", TNR11);
				addString_startNewLine("3. Условия определения неоднородности флюенса ионов:", TNR11);
				addString_startNewLine("        - температура окружающей среды:   " + first_sheet_values[actual_index][27].ToString() + "   °С;", TNR11);
				addString_startNewLine("        - атмосферное давление:   " + first_sheet_values[actual_index][25].ToString() + "   мм.рт.ст.;", TNR11);
				addString_startNewLine("        - относительная влажность воздуха:   " + first_sheet_values[actual_index][26].ToString() + "   %.", TNR11);
				addString_startNewLine("4. Средства определения неоднородности флюенса ионов:", TNR11);
				addString_startNewLine("        - испытательный стенд: ИИК 10К-400;", TNR11);
				addString_startNewLine("        - трековые мембраны (лавсановая плёнка);", TNR11);
				addString_startNewLine("        - установка для травления лавсановой плёнки;", TNR11);
				addString_startNewLine("        - растровый электронный микроскоп ТM-3000 (Hitachi, Япония);", TNR11);
				addString_startNewLine("        - система оцифровки видеосигнала «GALLERY-512».", TNR11);
				addString_startNewLine("5. Методика определения неоднородности флюенса ионов.", TNR11);
				addString_startNewLine("5.1. Проводилась в соответствии с «Методикой измерений флюенса тяжелых заряженных частиц", TNR11);
				addString_startNewLine("        с помощью трековых мембран на основе лавсановой пленки» ЦДКТ1.027.012-2015.", TNR11);
				addString_startNewLine("6.Результаты определения неоднородности флюенса ионов " + third_sheet_values[actual_index][1].ToString() + " " + third_sheet_values[actual_index][0].ToString() + " представлены в таблице 1:", TNR11);
				addString_startNewLine("              N = " + first_sheet_values[actual_index][22].ToString() + "        с-1" + "               Ф=        " + first_sheet_values[actual_index][7].ToString() + "        частиц*см-2", TNR11);
				num++;
				prevlens = 0; //сумма ширин предыдущих колонок
				pen = new XPen(XColor.FromName("Black"), 0.5);
				drawTable(new double[] { 70, 70, 70, 70, 70 }, new string[][] {     new string[]{ "ТД1"},
																				new string[]{ "ТД2"},
																				new string[]{ "ТД3"},
																				new string[]{ "ТД4"},
																				new string[]{ "ТД5"}},
																	new string[] { first_sheet_values[actual_index][8].ToString()
																	, first_sheet_values[actual_index][9].ToString()
																	, first_sheet_values[actual_index][10].ToString()
																	, first_sheet_values[actual_index][11].ToString()
																	, first_sheet_values[actual_index][12].ToString() },
																	TNR11, TNR11, headerH, tabrowH);
				drawTable(new double[] { 70, 70, 70, 70, 70 }, new string[][] {     new string[]{ "ТД6"},
																				new string[]{ "ТД7"},
																				new string[]{ "ТД8"},
																				new string[]{ "ТД9"},
																				new string[]{ "Среднее зн."}},
																	new string[] { first_sheet_values[actual_index][13].ToString()
																	, first_sheet_values[actual_index][14].ToString()
																	, first_sheet_values[actual_index][15].ToString()
																	, first_sheet_values[actual_index][16].ToString()
																	, first_sheet_values[actual_index][7].ToString()},
																	TNR11, TNR11, headerH, tabrowH);
				T += 3;
				num++;
				addString_startNewLine("         Коэффициент:           Красчетный  =  " + first_sheet_values[actual_index][30].ToString() + "   ±   " + first_sheet_values[actual_index][31].ToString(), TNR11);
				addString_startNewLine("         Неоднородность флюенса ионов составила :           " + first_sheet_values[actual_index][29].ToString() + "   %", TNR11);
				num++;
				addString_startNewLine("7. Принято решение о продолжении работ на ионе                  /  ̶п̶о̶в̶т̶о̶р̶н̶о̶й̶ ̶н̶а̶с̶т̶р̶о̶й̶к̶е̶ ̶п̶у̶ч̶к̶а̶", TNR11);
				addString_startNewLine("        в  " + second_sheet_values[actual_index][9].ToString().Substring(11), TNR11);
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





				Directory.CreateDirectory(@$"..\..\..\..\PDFresult\{chatId}\{foldNumber}");
				doc.Save(@$"..\..\..\..\PDFresult\{chatId}\{foldNumber}\TestTransit{ actual_index + 1}.pdf"); //путь, куда сохранять док
				doc.Close();
			}
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

	}
}
