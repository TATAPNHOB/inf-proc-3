using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace PDFreport1
{
	class Program
	{
		static void Main(string[] args)
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

			PdfDocument protocol = PdfSharp.Pdf.IO.PdfReader.Open("C:\\Users\\соня\\Desktop\\pdf\\prot2.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import); //путь к шаблону с пустыми полями
			PdfDocument doc = new PdfDocument();
			doc.AddPage(protocol.Pages[0]);
			var page = doc.Pages[0];

			XGraphics g = XGraphics.FromPdfPage(page);
			XFont font = new XFont("Times New Roman", 11);
			XFont font1 = new XFont("Times New Roman", 11, XFontStyle.Bold);
			XFont font2 = new XFont("Times New Roman", 12);
			XFont font3 = new XFont("Times New Roman", 8);

			//1.ОБЩИЕ СВЕДЕНИЯ О СЕАНСЕ
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66, 160.5, 70, 18), XStringFormats.Center); //название организации
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66+70, 160.5, 121, 18), XStringFormats.Center); //Шифр или наименование работы
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70+121, 160.5, 182, 18), XStringFormats.Center); //Облучаемое изделие
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 121+182, 160.5, 121, 18), XStringFormats.Center); //Время начала облучения
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 121 + 182+121, 160.5, 108, 18), XStringFormats.Center);  //Длительность
			//2
			g.DrawString("знач", font1, XBrushes.Black,
				new XRect(232, 181, 100, 12), XStringFormats.CenterLeft); //среда(нет блин вторник)
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66, 224, 70, 17), XStringFormats.Center); //угол
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66+70, 224, 121, 17), XStringFormats.Center); //Температура
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70+121, 224, 121, 17), XStringFormats.Center); //материал дегрейдора
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 121*2, 224, 121, 17), XStringFormats.Center); //толщина
			//3.ХАРАКТЕРИСТИКИ ИОНА
			g.DrawString("ЗН", font, XBrushes.Black,
				new XRect(66, 298, 70/2, 17), XStringFormats.CenterRight);//тип иона
			g.DrawString("000", font3, XBrushes.Black,
				new XRect(66+70/2, 299, 70/2, 17), XStringFormats.TopLeft);//номер для иона
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70, 298, 182, 17), XStringFormats.Center); //энергия Е на поверхности
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70+182, 298, 121, 17), XStringFormats.Center); //пробег R
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 182+121, 298, 182, 17), XStringFormats.Center); //линейные потери энергии ЛПЭ
			//4.ДАННЫЕ ПО ПРОПОРЦИОНАЛЬНЫМ СЧЕТЧИКАМ
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66, 363, 70, 26), XStringFormats.Center); //1
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66+70, 363, 182/3, 26), XStringFormats.Center); //2
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70+182/3, 363, 182 / 3, 26), XStringFormats.Center); //3
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 182 / 3+182/3, 363, 182 / 3, 26), XStringFormats.Center); //4
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66+70+182, 363, 121, 26), XStringFormats.Center); //среднее значение
			//5.Данные по трековым мембранам из лавсановой пленки:
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66, 427, 70, 26), XStringFormats.Center); //детектор 1
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70, 427, 182 / 3, 26), XStringFormats.Center); //детектор 2
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70+182/3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 3
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 182/3+182/3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 4
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 182, 427, 121/2, 26), XStringFormats.Center); //детектор 5
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 182+121/2, 427, 121 / 2, 26), XStringFormats.Center); //детектор 6
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 182 + 121, 427, 182/3, 26), XStringFormats.Center); //детектор 7
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 182 + 121+182/3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 8
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 182 + 121+182/3+182/3, 427, 182 / 3, 26), XStringFormats.Center); //детектор 9
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(66 + 70 + 182 + 121 + 182, 427, 153, 26), XStringFormats.Center); //неоднородность

			g.DrawString("000", font1, XBrushes.Black,
				new XRect(735.5, 98.5, 20, 9.5), XStringFormats.Center); //номер сеанса
			g.DrawString("TЗЧ/" + "0000" + "-" + "ЗН" + "-" + "00" + "/" + "00" + "-" + "000", font2, XBrushes.Black,
				new XRect(662, 79, 115, 13), XStringFormats.Center); ; //ТЗЧ/...
			g.DrawString("0,00", font, XBrushes.Black,
				new XRect(669.5, 316.5, 44, 12), XStringFormats.Center); //K
			g.DrawString("0,00", font, XBrushes.Black,
				new XRect(733.5, 316.5, 44, 12), XStringFormats.Center); //K погрешность
			g.DrawString("0/0-0", font, XBrushes.Black,
				new XRect(659.5, 349, 46.5, 12), XStringFormats.Center); //номер протокола допуска
			g.DrawString("0,00", font, XBrushes.Black,
				new XRect(668, 376, 46.5, 12), XStringFormats.Center); //K факт

			doc.Save("C:\\Users\\соня\\Desktop\\pdf\\Test1.pdf"); //путь, куда сохранять док
		}
	}
}
