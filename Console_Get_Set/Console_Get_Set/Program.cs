using System;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System.IO;
using Google.Apis.Sheets.v4.Data;

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
            ReadEntries(sheets[0], SpreadsheetID, "!A1:G10");
            Console.ReadKey();

            
        }
        static void ReadEntries(string sheet, string SpreadsheetID, string Range)
        {
            string range = $"{sheet}{Range}";//Тут любой рендж
            var readRequest = service.Spreadsheets.Values.Get(SpreadsheetID, range);//Повторяется везде, меняется только метод

            var response = readRequest.Execute();//Завершает реквест и передаёт результат для дальнейшей работы
            var values = response.Values;
            if (values != null && values.Count > 0)//Ну тут понятно
            {
                foreach (var row in values)
                {
                    Console.WriteLine("\n");
                    foreach(string rowitem in row)
                    {
                        Console.Write(rowitem+"|");
                    }
                    //Вид вывода, нам особо не понадобится, мы в консоли выводить не будем вроде
                    //Сразу будем пихать в пдф, разве что для понимания - values - лист листов
                    //Поэтому можно либо без форича обращаться как к двумерному массиву,
                    //Либо как-то так запихивать вместо вывода запихнув ещё один форич для ввода
                    //Ну или можно просто сделать ссылку на лист, чего я туплю
                }
            }
            else
            {
                Console.WriteLine("No data found");
            }
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
