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

			PdfDocument protocol = PdfSharp.Pdf.IO.PdfReader.Open("C:\\Users\\соня\\Desktop\\pdf\\prot3.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import); //путь к шаблону с пустыми полями
			PdfDocument doc = new PdfDocument();
			doc.AddPage(protocol.Pages[0]);
			var page = doc.Pages[0];

			XGraphics g = XGraphics.FromPdfPage(page);
			XFont font = new XFont("Times New Roman", 11);
			XFont font1 = new XFont("Times New Roman", 14);
			XFont font2 = new XFont("Times New Roman", 12);
			XFont font3 = new XFont("Times New Roman", 10);
			XFont font4 = new XFont("Times New Roman", 8);

			g.DrawString("TЗЧ/"+"0000"+"-"+"ЗН"+"-"+"0/0-0", font2, XBrushes.Black,
				new XRect(444, 58.75, 115, 13), XStringFormats.Center);; //ТЗЧ/...
			g.DrawString("0/0-0", font1, XBrushes.Black,
				new XRect(251, 76, 58, 14), XStringFormats.Center); //протокол №
			g.DrawString("00.00.0000", font1, XBrushes.Black,
				new XRect(361, 76, 58, 14), XStringFormats.Center); //от
			g.DrawString("000", font3, XBrushes.Black,
				new XRect(449, 111, 39/2, 15), XStringFormats.TopCenter); //индекс
			g.DrawString("ЗН", font1, XBrushes.Black,
				new XRect(449+39/2, 111, 39 / 2, 15), XStringFormats.Center); //ион
			g.DrawString("0,0", font1, XBrushes.Black,
				new XRect(155, 130.5, 61, 15), XStringFormats.Center); //энергия
			g.DrawString("00.00.0000 00:00:00", font, XBrushes.Black,
				new XRect(198, 202.5, 101.5, 10.5), XStringFormats.Center); //в период с
			g.DrawString("00.00.0000 00:00:00", font, XBrushes.Black,
				new XRect(356, 202.5, 101.5, 10.5), XStringFormats.Center); //по

			g.DrawString("00", font, XBrushes.Black,
				new XRect(277.75, 237, 37, 10), XStringFormats.Center); //температура
			g.DrawString("000", font, XBrushes.Black,
				new XRect(215, 252.5, 34, 10), XStringFormats.Center); //давление
			g.DrawString("00", font, XBrushes.Black,
				new XRect(271, 268.5, 34, 10), XStringFormats.Center); //влажность
			g.DrawString("000", font4, XBrushes.Black,
				new XRect(365.5, 439, 69/2, 13), XStringFormats.TopRight); //индекс
			g.DrawString("Зн", font, XBrushes.Black,
				new XRect(365.5+ 69 / 2, 439, 69 / 2, 13), XStringFormats.CenterLeft); //ион
			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(125, 455.5, 57, 12), XStringFormats.Center); //N
			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(330, 455.5, 57, 12), XStringFormats.Center); //Ф

			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(91, 501, 63, 31.5), XStringFormats.Center);//ТД1
			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(91+63, 501, 63, 31.5), XStringFormats.Center);//ТД2
			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(91 + 63*2, 501, 63, 31.5), XStringFormats.Center);//ТД3
			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(91 + 63 * 3, 501, 63, 31.5), XStringFormats.Center);//ТД4
			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(91 + 63 * 4, 501, 63, 31.5), XStringFormats.Center);//ТД5

			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(91, 548, 63, 31.5), XStringFormats.Center);//ТД6
			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(91 + 63, 548, 63, 31.5), XStringFormats.Center);//ТД7
			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(91 + 63 * 2, 548, 63, 31.5), XStringFormats.Center);//ТД8
			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(91 + 63 * 3, 548, 63, 31.5), XStringFormats.Center);//ТД9
			g.DrawString("0,00Е+00", font, XBrushes.Black,
				new XRect(91 + 63 * 4, 548, 63, 31.5), XStringFormats.Center);//Среднее

			g.DrawString("0,00", font, XBrushes.Black,
				new XRect(303, 598, 45.5, 12.5), XStringFormats.Center);//Красчетный
			g.DrawString("0,00", font, XBrushes.Black,
				new XRect(369, 598, 45.5, 12.5), XStringFormats.Center);//Кпогр
			g.DrawString("00,00", font, XBrushes.Black,
				new XRect(344.5, 615.5, 58, 12.5), XStringFormats.Center);//Неоднородность
			g.DrawString("00:00:00", font, XBrushes.Black,
				new XRect(117.5, 667,56.5,10), XStringFormats.Center);//продолжение время


			doc.Save("C:\\Users\\соня\\Desktop\\pdf\\Test2.pdf"); //путь, куда сохранять док
		}
	}
}
