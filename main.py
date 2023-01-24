import telebot
from telebot import types
from PIL import ImageGrab
from config import get_path
import win32com.client
import requests
import threading


def main1():
    print('ты долбоеб')
    dls = "http://26shkola.ru/wp-content/uploads/2023/01/21-января-5-11-классы-1.xlsx"
    resp = requests.get(dls)
    output = open('../test.xlsx', 'wb')
    output.write(resp.content)
    output.close()
    print(__name__, get_path(__name__))
    xlsx_path = get_path('пробуювфото', '../test.xlsx')
    client = win32com.client.Dispatch("Excel.Application")
    wb = client.Workbooks.Open(xlsx_path)
    ws = wb.ActiveSheet
    ws.Range("A2:J34").CopyPicture(Format=2)
    img = ImageGrab.grabclipboard()
    img.save(get_path('data', 'image.jpg'))
    wb.Close()
    client.Quit()


bot = telebot.TeleBot('5609999349:AAHm46TCL3_6pAqEx_2PPGXGlsvqSxiSSWY')


@bot.message_handler(commands=['start'])
def start(message):
    markup = types.InlineKeyboardMarkup(row_width=1)
    itembtn1 = types.InlineKeyboardButton(text='Что умеет бот?', callback_data='btn1')
    itembtn2 = types.InlineKeyboardButton(text='Скрыть кнопки', callback_data='btn2')
    markup.add(itembtn1, itembtn2)
    bot.send_message(message.chat.id, f'Привет, <i>{message.from_user.first_name}</i>, я тестовый бот, '
                                      f'приятно познакомиться👋👋👋', parse_mode='html', reply_markup=markup)


@bot.callback_query_handler(func=lambda callback: callback.data)
def whatisbot(callback):
    if callback.data == 'btn1':
        photo = open('/Users/kiru/PycharmProjects/пробуювфото/data/image.jpg', 'rb')
        bot.send_message(callback.message.chat.id, 'Пока я ничего не умею', parse_mode='html')
        bot.send_photo(callback.message.chat.id, photo)
    elif callback.data == 'btn2':
        bot.edit_message_text(chat_id=callback.message.chat.id, message_id=callback.message.id,
                              text='Кнопки скрыты', parse_mode='html')
        bot.send_message(callback.message.chat.id, 'Чтобы заново открыть клавиатуту, напиши /start', parse_mode='html')


@bot.message_handler(content_types=['photo', 'video'])
def chto_ya_umeu():
    pass


def start_dot():
    bot.polling(none_stop=True)


t1 = threading.Thread(target=main1)
t2 = threading.Thread(target=start_dot)
t1.start()
t2.start()
