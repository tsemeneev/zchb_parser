import telebot
from utils import parser

bot = telebot.TeleBot('Token')


@bot.message_handler(content_types=['document'])
def start_handler(message):
    bot.send_message(message.chat.id, 'Принял. Обрабатываю')

    file_info = bot.get_file(message.document.file_id)
    file_path = file_info.file_path

    downloaded_file = bot.download_file(file_path)
    with open('./documents/inn.xlsx', 'wb') as f:
        f.write(downloaded_file)

    parser.create_excel()
    parser.parse()

    bot.send_chat_action(message.chat.id, 'upload_document', timeout=28800)
    with open('./documents/inn_ready.xlsx', 'rb') as f:
        bot.send_document(message.chat.id, f)


bot.polling(skip_pending=True)


