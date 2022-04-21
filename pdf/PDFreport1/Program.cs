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

			PdfDocument protocol = PdfSharp.Pdf.IO.PdfReader.Open("C:\\Users\\соня\\Desktop\\pdf\\prot1.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import); //путь к шаблону с пустыми полями
			PdfDocument doc = new PdfDocument();
			doc.AddPage(protocol.Pages[0]);
			var page = doc.Pages[0];

			XGraphics g = XGraphics.FromPdfPage(page);
			XFont font = new XFont("Arial", 11);
			XFont font1 = new XFont("Times New Roman", 11, XFontStyle.Bold);
			XFont font2 = new XFont("Times New Roman", 14);
			XFont font3 = new XFont("Arial", 8);

			//1.ОБЩИЕ СВЕДЕНИЯ О СЕАНСЕ
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(49.75, 158, 81, 33.25), XStringFormats.Center); //название организации
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(49.75+81, 158, 119, 33.25), XStringFormats.Center); //Шифр или наименование работы
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(49.75 + 81+119, 158, 178.5, 33.25), XStringFormats.Center); //Облучаемое изделие
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(49.75 + 81 + 119+178.5, 158, 118.5, 33.25), XStringFormats.Center); //Время начала облучения
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(49.75 + 81 + 119 + 178.5+118.5, 158, 110.5, 33.25), XStringFormats.Center); //Длительность
			//2.УСЛОВИЯ ЭКСПЕРИМЕНТА: В СРЕДЕ ВАКУУМ
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50, 236.5, 81, 14.85), XStringFormats.Center); //угол
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50+81, 236.5, 119, 14.85), XStringFormats.Center); //Температура
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81+119, 236.5, 119, 14.85), XStringFormats.Center); //материал дегрейдора
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81 + 119+119, 236.5, 119, 14.85), XStringFormats.Center); //толщина
			//3.ХАРАКТЕРИСТИКИ ИОНА
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50, 310, 81/2, 14.85), XStringFormats.CenterRight);//тип иона
			g.DrawString("000", font3, XBrushes.Black,
				new XRect(50+81/2, 310, 81/2, 14.85), XStringFormats.TopLeft);//номер для иона
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50+81, 310, 179, 14.85), XStringFormats.Center); //энергия Е на поверхности
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81+179, 310, 119, 14.85), XStringFormats.Center); //пробег R
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81 + 179+119, 310, 178.5, 14.85), XStringFormats.Center); //линейные потери энергии ЛПЭ
			//4.ДАННЫЕ ПО ПРОПОРЦИОНАЛЬНЫМ СЧЕТЧИКАМ
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50, 373, 81, 25.5), XStringFormats.Center); //1
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50+81, 373, 179/3, 25.5), XStringFormats.Center); //2
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81+ 179 / 3, 373, 179 / 3, 25.5), XStringFormats.Center); //3
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81 + 179 / 3+179/3, 373, 179 / 3, 25.5), XStringFormats.Center); //4
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81 + 179, 373, 119, 25.5), XStringFormats.Center); //среднее значение
			//5.Данные по трековым мембранам из лавсановой пленки:
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50, 438, 81, 25.5), XStringFormats.Center); //детектор 1
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 2
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81 + 179 / 3, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 3
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81 + 179 / 3 + 179 / 3, 438, 179 / 3, 25.5), XStringFormats.Center); //детектор 4
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81 + 179, 438, 119/2, 25.5), XStringFormats.Center); //детектор 5
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81 + 179+119/2, 438, 119 / 2, 25.5), XStringFormats.Center); //детектор 6
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81 + 179 + 119+178.5, 438, 95, 25.5), XStringFormats.Center); //неоднородность, по лев
			g.DrawString("ЗНАЧ", font, XBrushes.Black,
				new XRect(50 + 81 + 179 + 119 + 178.5+95, 438, 95, 25.5), XStringFormats.Center); //неоднородность, по прав

			g.DrawString("000", font1, XBrushes.Black,
				new XRect(753, 97, 18.75, 11.5), XStringFormats.Center); //номер сеанса
			g.DrawString("TЗЧ/" + "0000" + "-" + "ЗН" + "-" + "00"+"/"+"00"+"-"+"000", font2, XBrushes.Black,
				new XRect(667, 77, 115, 13), XStringFormats.Center); ; //ТЗЧ/...
			g.DrawString("0,00", font, XBrushes.Black,
				new XRect(659, 327, 46.5, 12), XStringFormats.Center); //K
			g.DrawString("0,00", font, XBrushes.Black,
				new XRect(748, 327, 46.5, 12), XStringFormats.Center); //K погрешность
			g.DrawString("0/00", font, XBrushes.Black,
				new XRect(659, 359, 46.5, 12), XStringFormats.Center); //номер протокола допуска
			g.DrawString("0,00", font, XBrushes.Black,
				new XRect(660, 385, 46.5, 12), XStringFormats.Center); //K факт

			doc.Save("C:\\Users\\соня\\Desktop\\pdf\\Test.pdf"); //путь, куда сохранять док
		}
	}
}
