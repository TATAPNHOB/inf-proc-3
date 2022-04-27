using BotManager;
using System.Collections.Generic;
using Telegram.Bot;
using System.Net;
using System.Net.Mail;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Extensions.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.ReplyMarkups;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.InputFiles;



/*
 -ввести количество попыток авторизации
 -шифрануть файлик с паролями и логинами
 */


//@heheTon_bot
var botClient = new TelegramBotClient("5057524859:AAGDp0kc2ILs0yvuo7BtYAr_4Ea_E10OKH0");

using var cts = new CancellationTokenSource();

sheetsTool sTool = new sheetsTool();
sTool.Proccess();
//список пользователей, работающих сейчас с ботом и на каком шаге они находятся в "сценарии"
//Dictionary<long, int> currentUsers = new Dictionary<long, int>();
Dictionary<long, botUser> currentUsers = new Dictionary<long, botUser>();


/*step:
    0 - default logged status
    1 - default unlogged status
    2 - login request
    3 - password request
    20 - report email request (1)
    21 - report attachment request


    read codes:
    101 - Request 1
    102 - Request 2
    ...

 */

/*костыль 
после ввода логина, логин сохраняется в этом ассоциативном массивчике, 
    чтобы сопоставить етот логин с вводимым паролем
*/


//клавиатура для авторизции
/*ReplyKeyboardMarkup keyboard_Request = new(
new KeyboardButton ("Авторизация"))
{
    ResizeKeyboard = false
};*/

ReplyKeyboardMarkup keyboard_Request = new(
new KeyboardButton("Авторизация"));

//клавиатура залогиненного пользователя
ReplyKeyboardMarkup keyboard_Main = new(new[] {
new KeyboardButton("Протоколы"),
new KeyboardButton("Чтение"),
new KeyboardButton("Добавление")});

ReplyKeyboardMarkup keyboard_Add = new(new[] {
new KeyboardButton[] { "1", "2" },
        new KeyboardButton[] { "3", "4" },
        new KeyboardButton[] { "5"}});

ReplyKeyboardMarkup keyboard_yn = new(new[] {
new KeyboardButton("Да"),
new KeyboardButton("Нет"),});

ReplyKeyboardMarkup keyboard_protocol = new(new[] {
new KeyboardButton("Допуска"),
new KeyboardButton("Мониторинга"),});

ReplyKeyboardMarkup keyboard_read = new(new[]
    {
        new KeyboardButton[] { "1", "2" },
        new KeyboardButton[] { "3", "4" },
        new KeyboardButton[] { "5", "6" },
        new KeyboardButton[] { "7" },
    });
//
ReplyKeyboardRemove keyboardToRemove = new ReplyKeyboardRemove();

Dictionary<long, string> passwordRequest = new Dictionary<long, string>();


var receiverOptions = new ReceiverOptions
{
    AllowedUpdates = { } // получение всех типов обновлений
};

botClient.StartReceiving(
    HandleUpdateAsync,
    HandleErrorAsync,
    receiverOptions,
    cancellationToken: cts.Token);

var me = await botClient.GetMeAsync();


Console.ReadLine();

// Отправляет реквест отмены чтобы остановить бота
cts.Cancel();

Message sentMessage;


//метод, в котором идёт ответ на введёное пользователем сообщение
async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
{
    // Только сообщения:
    if (update.Type != UpdateType.Message)
        return;
    // и только текстовые
    if (update.Message!.Type != MessageType.Text)
        return;
    
    long chatId = update.Message.Chat.Id;
    if (!currentUsers.ContainsKey(chatId)) currentUsers.Add(chatId, new botUser(chatId));

    


    var messageText = update.Message.Text;
    switch (currentUsers[chatId].step)
    {  

        case 1:
            DefaultUnloggedScenery(chatId, messageText);
            break;

        case 2:
            Authorization_Login(chatId, messageText);
            break;

        case 3:
            Authorization_Password(chatId, messageText);
            break;

        default:
            DefaultScenery(chatId, messageText);
            break;
    }
}


//сценарий залогиненного пользователя
async void DefaultScenery(long chatId, string messageText)
{
    if (currentUsers[chatId].step == 222) return;
    if (currentUsers[chatId].step == 100)
    {
        switch (messageText)
        {

            case "1":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Введите название компании",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 101;
                break;

            case "2":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Введите название компании",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 102;
                break;
                

            case "3":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Введите ион (XX123) и номер сессии",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 103;
                break;

            case "4":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Введите имя иона",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 104;
                break;
            
            case "5":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Введите имя иона",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 105;
                break;
            case "6":
                sTool.Request6(chatId, currentUsers[chatId].folderNumb);
                await SendReportAsync(chatId, currentUsers[chatId].email);
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Отправлено!",
                    replyMarkup: keyboard_Main);
                currentUsers[chatId].step = 0;
                break;

            case "7":
                sTool.Request7(chatId, currentUsers[chatId].folderNumb);
                await SendReportAsync(chatId, currentUsers[chatId].email);
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Отправлено!",
                    replyMarkup: keyboard_Main);
                currentUsers[chatId].step = 0;
                break;

        }
    }
    if (currentUsers[chatId].step == 200)
    {
        switch (messageText)
        {
            //другие
            case "1":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Введите (через пробелы):\nНомер сеанса | Название организации | Начало сеанса | Конец сеанса | Продолжительность облучения | Продолжительность сеанса | Время с учетом ТП | Начало ТП | Конец ТП | Время ТП в сеансе",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 201;
                break;
            //смена
            case "2":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Введите (через пробелы):\nНомер сеанса | Название организации | Объект испытаний 1 | Объект испытаний 2 | Объект испытаний 3 | Объект испытаний 4 | ТД 1 | ТД 2 | ТД 3 | ТД 4 | ТД 5 | ТД 6 | ТД 7 | ТД 8 | ТД 9 | ОД 1 |  ОД 2 | ОД 3 | ОД 4 | Интенсивность потока | Температура в сеансе | Начало сеанса | Конец сеанса | Продолжительность облучения | Продолжительность сеанса | Время с учетом ТП | Начало ТП | Конец ТП | Время ТП в сеансе",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 202;
                break;

            //переход
            case "3":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Номер сеанса | Название организации | Объект испытаний 1 | Объект испытаний 2 | Объект испытаний 3 | Объект испытаний 4 | Код среды | ТД 1 | ТД 2 | ТД 3 | ТД 4 | ТД 5 | ТД 6 | ТД 7 | ТД 8 | ТД 9 | ОД 1 |  ОД 2 | ОД 3 | ОД 4 | Интенсивность потока | Угол облучения | Температура в сеансе | Начало сеанса | Конец сеанса | Продолжительность облучения | Продолжительность сеанса | Время с учетом ТП | Начало ТП | Конец ТП | Время ТП в сеансе",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 203;
                break;

            //сеанс обычный
            case "4":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Введите (через пробелы):\nНомер сеанса | Название организации | Объект испытаний 1 | Объект испытаний 2 | Объект испытаний 3 | Объект испытаний 4 | ТД 1 | ТД 2 | ТД 3 | ТД 4 | ТД 5 | ТД 6 | ТД 7 | ТД 8 | ТД 9 | ОД 1 |  ОД 2 | ОД 3 | ОД 4 | Интенсивность потока | Температура в сеансе | Начало сеанса | Конец сеанса | Продолжительность облучения | Продолжительность сеанса | Время с учетом ТП | Начало ТП | Конец ТП | Время ТП в сеансе",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 204;
                break;

            //информация по иону
            case "5":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Введите (через пробелы):\nНазвание иона | Изотоп | Материал дегрейдора, мкм | Толщина дегрейдора, мкм | Среда | \n Энергия E на поверхности объекта испытаний, МэВ / н | Погрешность задания энергии на объекте испытаний, МэВ / н | Пробег[Si], мкм |\n Погрешность задания пробега, мкм | ЛПЭ[Si], МэВ * см2 / мг | Погрешность задания ЛПЭ, МэВ * см ^ 2 / мг |\n Энергия E в ионопроводе, МэВ / н | номер сессии в году | Код среды ",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 205;
                break;

            //системный лист
            case "6":
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Введите (через пробелы):\nДавление на исп. комплексе | Влажность на исп комплексе, % | Температура на исп. комплексе",
                    replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 205;
                break;
            default:
                return;

        }
    }

    if (new List<string> { "1", "2", "3", "4", "5", "6", "7" }.Contains(messageText)) return;
    switch (currentUsers[chatId].step)
    {
        case 101:

            sTool.Request1(messageText, chatId, currentUsers[chatId].folderNumb);
            await SendReportAsync(chatId, currentUsers[chatId].email);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Отправлено!",
                replyMarkup: keyboard_Main);
            currentUsers[chatId].step = 0;
            return;

        case 102:
            sTool.Request2(messageText, chatId, currentUsers[chatId].folderNumb);
            await SendReportAsync(chatId, currentUsers[chatId].email);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Отправлено!",
                replyMarkup: keyboard_Main);
            currentUsers[chatId].step = 0;
            return;

        case 103:
            string ionName = messageText.Substring(0, 2);
            int indexSpace = messageText.IndexOf(" ");
            string ionIsotope = messageText.Substring(2, indexSpace - 2);
            string sesNumb = messageText.Substring(indexSpace + 1, messageText.Length - indexSpace - 1);

            sTool.Request3(ionName, ionIsotope, sesNumb, chatId, currentUsers[chatId].folderNumb);
            await SendReportAsync(chatId, currentUsers[chatId].email);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Отправлено!",
                replyMarkup: keyboard_Main);
            currentUsers[chatId].step = 0;
            return;

        case 104:
            sTool.Request4(messageText, chatId, currentUsers[chatId].folderNumb);
            await SendReportAsync(chatId, currentUsers[chatId].email);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Отправлено!",
                replyMarkup: keyboard_Main);

            currentUsers[chatId].step = 0;
            return;

        case 105:
            sTool.Request5(messageText, chatId, currentUsers[chatId].folderNumb);
            await SendReportAsync(chatId, currentUsers[chatId].email);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Отправлено!",
                replyMarkup: keyboard_Main);
            currentUsers[chatId].step = 0;
            return;



        default:
            break;
    }


    /////////////////////////////////
    string[] temp = new string[0];
    switch (currentUsers[chatId].step)
    {
        case 201:
            temp = messageText.Split(" ");
            sTool.fillDataScam(temp[0], temp[1], temp[2], temp[3], temp[4],
                temp[5], temp[6], temp[7], temp[8]);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Добавлено!",
                replyMarkup: keyboard_Main);
            currentUsers[chatId].step = 0;
            return;

        case 202:
            temp = messageText.Split(" ");
            sTool.fillDataSmena(temp[0], temp[1], temp[2], temp[3], temp[4],
                temp[5], temp[6], temp[7], temp[8], temp[9],
                temp[10], temp[11], temp[12], temp[13], temp[14],
                temp[15], temp[16], temp[17], temp[18], temp[19],
                temp[20], temp[21], temp[22], temp[23], temp[24],
                temp[25], temp[26], temp[27], temp[28], temp[29]);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Добавлено!",
                replyMarkup: keyboard_Main);
            currentUsers[chatId].step = 0;
            return;

        case 203:
            temp = messageText.Split(" ");
            sTool.fillDataPerehod(temp[0], temp[1], temp[2], temp[3], temp[4],
                temp[5], temp[6], temp[7], temp[8], temp[9],
                temp[10], temp[11], temp[12], temp[13], temp[14],
                temp[15], temp[16], temp[17], temp[18], temp[19],
                temp[20], temp[21], temp[22], temp[23], temp[24],
                temp[25], temp[26], temp[27], temp[28], temp[29], temp[30]);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Добавлено!",
                replyMarkup: keyboard_Main);
            currentUsers[chatId].step = 0;
            return; 

        case 204:
            temp = messageText.Split(" ");
            sTool.fillDataSeance(temp[0], temp[1], temp[2], temp[3], temp[4],
                temp[5], temp[6], temp[7], temp[8], temp[9],
                temp[10], temp[11], temp[12], temp[13], temp[14],
                temp[15], temp[16], temp[17], temp[18], temp[19],
                temp[20], temp[21], temp[22], temp[23], temp[24],
                temp[25], temp[26], temp[27], temp[28]);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Добавлено!",
                replyMarkup: keyboard_Main);

            currentUsers[chatId].step = 0;
            return;

        case 205:
            temp = messageText.Split(" ");
            sTool.fillIonInfo(temp[0], temp[1], temp[2], temp[3], temp[4],
                temp[5], temp[6], temp[7], temp[8], temp[9],temp[10],temp[11],temp[12],
                temp[13]);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Добавлено!",
                replyMarkup: keyboard_Main);
            currentUsers[chatId].step = 0;
            return;

        case 206:
            sTool.Request5(messageText, chatId, currentUsers[chatId].folderNumb);
            await SendReportAsync(chatId, currentUsers[chatId].email);
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Добавлено!",
                replyMarkup: keyboard_Main);
            currentUsers[chatId].step = 0;
            return;



        default:
            break;
    }
    



    if (currentUsers[chatId].step==20)
    {
        switch (messageText)
        {
            case "Мониторинга":
                sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Введите диапазон (xx-xx):",
                replyMarkup: keyboardToRemove);
                currentUsers[chatId].step = 31;
                break;

            case "Допуска":
                sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Введите диапазон (xx-xx):",
                replyMarkup: keyboardToRemove);

                currentUsers[chatId].step = 32;
                break;
        }
        return;

        /*sentMessage = await botClient.SendTextMessageAsync(
               chatId: chatId,
               text: "Чтобы результат отправился в чат?",
               replyMarkup: keyboard_yn);*/
        
        
    }
    if (currentUsers[chatId].step >= 31 && currentUsers[chatId].step <=32)
    {
        switch (currentUsers[chatId].step)
        {
            case 31:
                /*if (Directory.Exists(@$"..\..\..\..\PDFresult\{chatId}"))
                {
                    Directory.Delete(@$"..\..\..\..\PDFresult\{chatId}", true);
                }*/

                //sTool.Proccess(chatId, currentUsers[chatId].folderNumb);//генерация пдфки
                sTool.PDFsGEN_sean(messageText, chatId, currentUsers[chatId].folderNumb);
                await SendReportAsync(chatId, currentUsers[chatId].email);//отправка её на почту
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Отправлено!",
                    replyMarkup: keyboard_Main);

                return;
            case 32:
                sTool.PDFsGEN_trans(messageText, chatId, currentUsers[chatId].folderNumb);
                await SendReportAsync(chatId, currentUsers[chatId].email);//отправка её на почту
                sentMessage = await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Отправлено!",
                    replyMarkup: keyboard_Main);
                return;
        }
        

        /*sentMessage = await botClient.SendTextMessageAsync(
               chatId: chatId,
               text: "Чтобы результат отправился в чат?",
               replyMarkup: keyboard_yn);*/


    }

    if (currentUsers[chatId].step == 21)
    { 
        
        switch (messageText)
        {
            case "Да":
                
                break;
        }
        currentUsers[chatId].step = 0;
        return;
    }

    switch (messageText)
    {
        case "Протоколы":
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Какой протокол?",
                replyMarkup: keyboard_protocol);
            currentUsers[chatId].folderNumb++;
            currentUsers[chatId].step = 20;
            break;


        case "Чтение":
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Выберите нужный запрос: \n" +
                        "Информация по договору:\n1) Общее время работы с разбивкой по ионам\n2) Время начала работ по договору\n\nИнформация по иону:\n3) Тип, энергия, пробег в кремнии\n4) Выработанное время на ионе по каждому договору; \n5) Время затраченное на технологические перерывы и простои.\n\n"+
                        "Текущее состояние:\n6) № сеанса и его статус; \n7)Время начала данного сеанса;\n ",
                replyMarkup: keyboard_read);
            currentUsers[chatId].folderNumb++;
            currentUsers[chatId].step = 100;
            break;

        case "Добавление":
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Выберите способ:\n1) Обычный сеанс\n2) Информация по иону\n3) Переход\n4) Смена\n5) Простой, детектор и другое",
                replyMarkup: keyboard_Add);
            currentUsers[chatId].step = 200;

            break;

        default:
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Некорректная команда. \nВыберите действие",
                replyMarkup: keyboard_Main);
            //Console.WriteLine("Запрошено: авторизация");
            break;
    }
}



//сценарий, если пользователь неавторизирован
async void DefaultUnloggedScenery (long chatId, string messageText)
{
    switch (messageText)
    {
        case "/help":
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Чтобы войти в систему напишите:\n Авторизация",
                replyMarkup: keyboard_Request);
            //Console.WriteLine("Запрошено: авторизация");
            break;

        case "Авторизация":
            if (currentUsers[chatId].IsBanned)
            {
                sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Вы временно заблокированы \nиз-за подозрительных попыток входа!",
                replyMarkup: keyboardToRemove);
                return;
            }
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Введите логин:",
                replyMarkup: keyboardToRemove);
            Console.WriteLine($"{chatId}: запрошена авторизация");
            currentUsers[chatId].step = 2;
            break;

        default:
            sentMessage = await botClient.SendTextMessageAsync(
                 chatId: chatId,
                 text: "Некорректная команда. \nВыберите действие",
                 replyMarkup: keyboard_Request);
            break;
    }
}

//сценарий, когда бот просит пользователя ввести логин
async void Authorization_Login(long chatId, string login_input)
{
    StreamReader sr = new StreamReader(@"..\..\..\..\BD_user.txt");
    string line = sr.ReadLine();
    string login;
    while (line != null)
    {
        //ожидается, что в файле строка будет выглядеть как два слова
        //где первое слово - логин, а второе - пароль
        
        login = line.Split(' ')[0];

        if (login == login_input)
        { 
            //переходим для этого пользователя на новый сценарий - ввод пароль
            currentUsers[chatId].step = 3;
            //также записываем, какой логин он ввёл
            passwordRequest[chatId] = login;
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Введите пароль:");
            Console.WriteLine($"{chatId}: запрошен пароль");
            return;
            //нашли логин => запрашиваем пароль
        }
        line = sr.ReadLine();
        
    }
    //если логин не нашли, то тыкаем юзера

    sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Такой пользователь не найден. \nПопробуй ввести новый логин!");
   

    //увеличиваем количество провальных попыток входа
    currentUsers[chatId].tries++;

    //баним временно пользователя, если он слишком много раз пробует войти
    if (currentUsers[chatId].tries%5==0)
    {
        currentUsers[chatId].banDate = DateTime.Now;
        currentUsers[chatId].step = 1;
    }
    //не продуман случай, если юзер захочет выйти из ввода логина
}


//сценарий, когда бот просит пользователя ввести пароль
async void Authorization_Password(long chatId, string password_input)
{
    StreamReader sr = new StreamReader(@"..\..\..\..\BD_user.txt");
    string line = sr.ReadLine();
    string login, password;
    while (line != null)
    {
        //ожидается, что в файле строка будет выглядеть как два слова
        //где первое слово - логин, а второе - пароль
        

        login = line.Split(' ')[0];
        password = line.Split(' ')[1];

        if (password == password_input)
        {
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Вы успешно вошли",
                replyMarkup: keyboard_Main);
            Console.WriteLine($"{chatId} авторизировался");
            currentUsers[chatId].email = line.Split(' ')[2];
            currentUsers[chatId].step = 0;
            return;
        }
        else
        {

            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Пароль неверен. \nПопробуй ввести его заново!");
            Console.WriteLine($"{chatId} отказано в пароле");


            //увеличиваем количество провальных попыток входа
            currentUsers[chatId].tries++;

            //баним временно пользователя, если он слишком много раз пробует войти
            if (currentUsers[chatId].tries%5 == 0)
            {
                currentUsers[chatId].banDate = DateTime.Now;
                currentUsers[chatId].step = 1;
            }

        }
        line = sr.ReadLine();

    }
    //если логин не нашли, то тыкаем юзера

   
    //не продуман случай, если юзер захочет выйти из ввода логина
}


//метод отправки на email
async Task SendReportAsync(long chatId, string email)
{
    currentUsers[chatId].step = 222;
    MailAddress from = new MailAddress("hehhahbot@gmail.com", "Hehebot");
    MailAddress to = new MailAddress(email);
    MailMessage m = new MailMessage(from, to);
    foreach (var item in Directory.GetFiles(@$"..\..\..\..\PDFresult\{chatId}\{currentUsers[chatId].folderNumb}"))
    {
        m.Attachments.Add(new Attachment(item));
    }
    
    //тема письма
    m.Subject = "Отчетный документ";

    //текст письма
    m.Body = "-----------------------";

    //прикрепляем документ отчета
    
    

    // адрес smtp-сервера и порт, с которого отправляется письмо
    SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);

    //логин и пароль
    smtp.Credentials = new NetworkCredential("hehhahbot@gmail.com", "sdif@jfds331kmlfdsk12290847501213F");

    smtp.EnableSsl = true;
    //отправка
    await smtp.SendMailAsync(m);

    Console.WriteLine($"{chatId} авторизировался");
    Console.WriteLine($"{chatId}: письмо отправлено");
    currentUsers[chatId].step = 0;
    /*
     вызов метода:
        SendReportAsync().GetAwaiter();*/
}


Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
{
    var ErrorMessage = exception switch
    {
        ApiRequestException apiRequestException
            => $"Telegram API Error:\n[{apiRequestException.ErrorCode}]\n{apiRequestException.Message}",
        _ => exception.ToString()
    };

    Console.WriteLine(ErrorMessage);
    return Task.CompletedTask;
}