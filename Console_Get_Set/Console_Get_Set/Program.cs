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
        static readonly string SpreadsheetID = "1kcFJyXxEexrB0pUtM0353QFHrzKqD4e6LQaZUVHOqLY";//в ссылке на лист
        static readonly string sheet = "congress";//имя листа (не таблицы)
        static SheetsService service;//Объект для работы собсно с листиками
        static void Main(string[] args)
        {
            GoogleCredential credential;//Права 
            using (var stream = new FileStream("client-secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            ReadEntries();//Диапазон A1:F10, вывод F,E,D,B
            Console.ReadKey();
            CreateEntry();//Нижняя свободная строка, диапазон A:F
            Console.ReadKey();
            UpdateEntry();//D543 значение "updated", не работает на диапазоне, думаю и не надо
            Console.ReadKey();
            DeleteEntry();//Удаление строки A543:F543;
            Console.ReadKey();
        }
        static void ReadEntries()
        {
            var range = $"{sheet}!A1:F10";//Тут любой рендж
            var readRequest = service.Spreadsheets.Values.Get(SpreadsheetID, range);//Повторяется везде, меняется только метод

            var response = readRequest.Execute();//Завершает реквест и передаёт результат для дальнейшей работы
            var values = response.Values;
            if (values != null && values.Count > 0)//Ну тут понятно
            {
                foreach (var row in values)
                {
                    Console.WriteLine("{0} {1} | {2} | {3}", row[5], row[4], row[3], row[1]);
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
        static void CreateEntry()
        {
            var range = $"{sheet}!A:F";//Специфичный ренж, вставляется вниз, поэтому номер строки не нужен
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "Hello", "Pomogite", "Pls", "Ya", "Hochu", "Umeret" };
            valueRange.Values = new List<IList<object>> { objectList };//Собсно что писать будем

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            //Опция парса, это стандартный парс, числовые переменные так и останутся, а строки могут парсится в даты, числа etc
            var appendResponse = appendRequest.Execute();//Завершает реквест и передаёт результат для дальнейшей работы
        }
        static void UpdateEntry()
        {
            var range = $"{sheet}!A540";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "updated" };
            valueRange.Values = new List<IList<object>> { objectList };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }
        static void DeleteEntry()
        {
            var range = $"{sheet}!A540:F541";//Судя по тестам любой ренж, и строка не поднимается
            //Эквиваленто апдейту на пустую строку получается
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetID, range);
            var deleteResponse = deleteRequest.Execute();
        }
    }
}
