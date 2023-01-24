import telebot
from telebot import types
from PIL import ImageGrab
from config import get_path
import win32com.client


def main1():
    print(__name__, get_path(__name__))
    xlsx_path = get_path('/Users/kiru/PycharmProjects/bottest', 'test.xlsx')
    client = win32com.client.Dispatch("Excel.Application")
    wb = client.Workbooks.Open(xlsx_path)
    ws = wb.ActiveSheet
    ws.Range("A2:J34").CopyPicture(Format=2)
    img = ImageGrab.grabclipboard()
    img.save(get_path('data', 'image.jpg'))
    wb.Close()
    client.Quit()


def main2():
    print(__name__, get_path(__name__))
    xlsx_path = get_path('/Users/kiru/PycharmProjects/bottest', 'test1.xlsx')
    client = win32com.client.Dispatch("Excel.Application")
    wb = client.Workbooks.Open(xlsx_path)
    ws = wb.ActiveSheet
    ws.Range("A2:J34").CopyPicture(Format=2)
    img = ImageGrab.grabclipboard()
    img.save(get_path('data', 'image1.jpg'))
    wb.Close()
    client.Quit()


bot = telebot.TeleBot('5609999349:AAHm46TCL3_6pAqEx_2PPGXGlsvqSxiSSWY')


@bot.message_handler(commands=['start'])
def start(message):
    markup = types.InlineKeyboardMarkup(row_width=1)
    itembtn1 = types.InlineKeyboardButton(text=f'Расписание на ', callback_data='btn1')
    itembtn2 = types.InlineKeyboardButton(text=f'Расписание на ', callback_data='btn2')
    itembtn3 = types.InlineKeyboardButton(text='Обычное расписание звонков', callback_data='btn3')
    itembtn4 = types.InlineKeyboardButton(text='Скрыть кнопки', callback_data='btn4')
    markup.add(itembtn1, itembtn2, itembtn3, itembtn4)
    bot.send_message(message.chat.id, f'Привет, <i>{message.from_user.first_name}</i>, я  бот, '
                                      f'приятно познакомиться👋👋👋', parse_mode='html', reply_markup=markup)


@bot.callback_query_handler(func=lambda callback: callback.data)
def whatisbot(callback):
    if callback.data == 'btn1':
        main1()
        photo = open('/Users/kiru/PycharmProjects/пробуювфото/data/image.jpg', 'rb')
        bot.send_message(callback.message.chat.id, 'Вот оно', parse_mode='html')
        bot.send_photo(callback.message.chat.id, photo)
    elif callback.data == 'btn2':
        main2()
        photo1 = open('/Users/kiru/PycharmProjects/пробуювфото/data/image1.jpg', 'rb')
        bot.send_message(callback.message.chat.id, 'Вот оно', parse_mode='html')
        bot.send_photo(callback.message.chat.id, photo1)
    elif callback.data == 'btn3':
        markup = types.InlineKeyboardMarkup(row_width=2)
        itembtn1 = types.InlineKeyboardButton(text='Звонки суббота', callback_data='subb')
        itembtn2 = types.InlineKeyboardButton(text=f'Звонки 1 смена', callback_data='sm1')
        itembtn3 = types.InlineKeyboardButton(text='Звонки 2 смена', callback_data='sm2')
        itembtn4 = types.InlineKeyboardButton(text='Звонки 1 смена пятница', callback_data='sm1p')
        itembtn5 = types.InlineKeyboardButton(text='Звонки 2 смена пятница', callback_data='sm2p')
        itembtn6 = types.InlineKeyboardButton(text='Назад', callback_data='backk')
        markup.row(itembtn2, itembtn3)
        markup.row(itembtn4)
        markup.row(itembtn5)
        markup.row(itembtn1, itembtn6)
        bot.edit_message_text(chat_id=callback.message.chat.id, message_id=callback.message.id,
                              text='Вот все что есть', parse_mode='html', reply_markup=markup)
    elif callback.data == 'btn4':
        bot.edit_message_text(chat_id=callback.message.chat.id, message_id=callback.message.id,
                              text='Кнопки скрыты', parse_mode='html')
        bot.send_message(callback.message.chat.id, 'Чтобы заново открыть клавиатуту, напиши /start', parse_mode='html')


@bot.callback_query_handler(func=lambda raspisanie: raspisanie.data)
def prostoraspis(raspisanie):
    if raspisanie.data == 'subb':
        photo1 = open('/Users/kiru/PycharmProjects/пробуювфото/data/subb.jpg', 'rb')
        bot.send_photo(raspisanie.callback.message.chat.id, photo1, 'Вот')
    elif




bot.polling(none_stop=True)
