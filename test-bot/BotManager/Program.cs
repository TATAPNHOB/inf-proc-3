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



/*
 -ввести количество попыток авторизации
 -шифрануть файлик с паролями и логинами
 */


//@heheTon_bot
var botClient = new TelegramBotClient("5057524859:AAGDp0kc2ILs0yvuo7BtYAr_4Ea_E10OKH0");

using var cts = new CancellationTokenSource();


//список пользователей, работающих сейчас с ботом и на каком шаге они находятся в "сценарии"
//Dictionary<long, int> currentUsers = new Dictionary<long, int>();
Dictionary<long, botUser> currentUsers = new Dictionary<long, botUser>();


/*step:
    0 - default logged status
    1 - default unlogged status
    2 - login request
    3 - password request
 */

/*костыль 
после ввода логина, логин сохраняется в этом ассоциативном массивчике, 
    чтобы сопоставить етот логин с вводимым паролем
*/


//клавиатура для авторизции
ReplyKeyboardMarkup keyboard_Request = new(
new KeyboardButton ("Авторизация"));

//клавиатура залогиненного пользователя
ReplyKeyboardMarkup keyboard_Main = new(new[] {
new KeyboardButton("Чтение"),
new KeyboardButton("Изменение"),});

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
        case 0:
            DefaultScenery(chatId, messageText);
            break;

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
            break;
    }

    /*Console.WriteLine($"Получено '{messageText}' сообщение в чатике {chatId}.");

// Ответочка введёному сообщению
sentMessage = await botClient.SendTextMessageAsync(
    chatId: chatId,
    text: "Вы произнесли:\n" + messageText,
    cancellationToken: cancellationToken);*/
}


//сценарий залогиненного пользователя
async void DefaultScenery(long chatId, string messageText)
{
    switch (messageText)
    {
        case "Чтение":
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: ""/*,
                replyMarkup: */);
            //Console.WriteLine("Запрошено: авторизация");
            break;

        case "Изменение":
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: ""/*,
                replyMarkup: */);
            //Console.WriteLine("Запрошено: авторизация");
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
                text: "Пропишите /help",
                replyMarkup: keyboard_Request);
            Console.WriteLine($"{chatId}: Введена некорректная команда.");
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
static async Task SendReportAsync(long chatId)
{
    MailAddress from = new MailAddress("hehhahbot@gmail.com", "Hehebot");
    MailAddress to = new MailAddress("lai.20@uni-dubna.ru");
    MailMessage m = new MailMessage(from, to);

    //тема письма
    m.Subject = "Отчетный документ";

    //текст письма
    m.Body = "Протокол допуска/протокол мониторинга";

    //прикрепляем документ отчета
    m.Attachments.Add(new Attachment(@"..\..\..\..\cat.pdf"));


    // адрес smtp-сервера и порт, с которого отправляется письмо
    SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);

    //логин и пароль
    smtp.Credentials = new NetworkCredential("hehhahbot@gmail.com", "sdif@jfds331kmlfdsk12290847501213F");

    smtp.EnableSsl = true;
    //отправка
    await smtp.SendMailAsync(m);
    Console.WriteLine($"{chatId}: Письмо отправлено");

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