using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Extensions.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.ReplyMarkups;
using Telegram.Bot.Types.Enums;
using System.Collections.Generic;

//@heheTon_bot
var botClient = new TelegramBotClient("5057524859:AAGDp0kc2ILs0yvuo7BtYAr_4Ea_E10OKH0");

using var cts = new CancellationTokenSource();


//список пользователей, работающих сейчас с ботом и на каком шаге они находятся в "сценарии"
Dictionary<long, int> currentUsers = new Dictionary<long, int>();
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
    if (!currentUsers.ContainsKey(chatId)) currentUsers.Add(chatId, 1);


    var messageText = update.Message.Text;
    switch (currentUsers[chatId])
    {
        case 1:
            DefaultScenery(chatId, messageText);
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

async void DefaultScenery(long chatId, string messageText)
{
    switch (messageText)
    {
        case "/help":
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Команды: \n Авторизация (кратко: авто)");
            //Console.WriteLine("Запрошено: авторизация");
            break;

        case "Авторизация":
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Введите логин:");
            Console.WriteLine($"{chatId}: запрошена авторизация");
            currentUsers[chatId] = 2;
            break;

        default:
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Пропишите /help");
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
            currentUsers[chatId] = 3;
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
                text: "Вы успешно вошли");
            Console.WriteLine($"{chatId} авторизировался");
            currentUsers[chatId] = 0;
            return;
        }
        else
        {
            sentMessage = await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Пароль неверен. \nПопробуй ввести его заново!");
            Console.WriteLine($"{chatId} отказано в пароле");
        }
        line = sr.ReadLine();

    }
    //если логин не нашли, то тыкаем юзера


    //не продуман случай, если юзер захочет выйти из ввода логина
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