import copy

import telebot
import openpyxl
from PIL import Image
from time import sleep
import os
import requests
from contextlib import ExitStack
import logging

import folder_magic
from templates import ProductInfoMessage

bot_token = os.getenv("KONEVOD_TOKEN")

bot = telebot.TeleBot(bot_token)

logging.basicConfig(filename='bot.log', encoding='utf-8', level=logging.INFO)


def is_media_file(filename) -> bool:
    if filename.suffix.lower() in ['.jpg', '.jpeg', '.png', '.mp4']:
        return True
    return False


# extracts arguments from message with bot commands
def extract_arg(arg):
    return arg.split()[1:]


# loads product data from execl file, creates folders and saves data in them
@bot.message_handler(commands=['load_products'])
def load_products(message):
    logging.info("Recieved command \"/load_products\"")
    bot.send_message(message.from_user.id, "Создаю папки")
    logging.info("Loading excel file \"catalog.xlsx\"")
    pxl_doc = openpyxl.load_workbook('catalog.xlsx', data_only=True)
    logging.info("Loading catalog sheet \"Каталог (МСК)\"")
    products = folder_magic.parse_products(pxl_doc, 'Каталог (МСК)')
    logging.info("Loading product info sheet \"ДанныеДляКаталога\"")
    products = folder_magic.parse_product_info(pxl_doc, 'ДанныеДляКаталога', products)
    logging.info("Loading photo sheet \"Фото\"")
    products = folder_magic.parse_photos(pxl_doc, 'Фото', products)
    pxl_doc.close()
    logging.info("Creating folders")
    folder_magic.create_folders(folder_magic.path_to_products_to_add, products)

    logging.info("Saving products info in JSON file")
    for product in products.values():
        folder_magic.serialize(product, {'Image'})
    # Need to wait because serialize changes folder logo and we need it to be done before we change it again
    sleep(1)
    logging.info("Setting logos of folders to thumbnail with product")
    for product in products.values():
        folder_magic.set_logo(product["Code"], product["Image"])

# Same thing as for loading products, but using url apis
#
# @bot.message_handler(commands=['product'])
# def get_product_message(message):
#     args = extract_arg(message.text)
#     if len(args) == 0:
#         bot.send_message(message.from_user.id, "Не указан 1-ый аргумент - код продукта")
#
#     product_code = args[0]
#
#     p_info = folder_magic.deserialize(folder_magic.path_to_products_to_add + args[0])
#     msg = MessageTemplate()
#     msg.p_name_line = msg.p_name_line.format(p_info["Name"])
#     msg.p_code_line = msg.p_code_line.format(p_info["Code"])
#     msg.p_supplier_line = msg.p_supplier_line.format(p_info["Supplier"]) if p_info["Supplier"] is not None else ''
#     msg.p_material_line = msg.p_material_line.format(p_info["Material"]) if p_info["Material"] is not None else ''
#     msg.p_price_line = msg.p_price_line.format(p_info["Price"])
#
#     media_group = []
#
#     request_url = "https://api.telegram.org/bot" + bot_token + "/sendMediaGroup"
#     params = {
#         "chat_id": message.from_user.id
#         , "media":
#             """[
#                 {
#                     "type": "photo"
#                     , "media": "attach://random-name-0"},
#                 {
#                     "type": "photo"
#                     , "media": "attach://random-name-1"}
#             ]"""
#     }
#
#     files = {f'random-name-{idx}': open(folder_magic.path_to_products_to_add + product_code + '/' + file_name, 'rb')
#              for idx, file_name in enumerate(list(
#                         set(os.listdir(folder_magic.path_to_products_to_add + product_code))
#                         - {"desktop.ini", "thumbnail.jpg", "product_info.json", "20220425-171025-958.mp4"}
#                 ))}
#
#     print(files)
#     result = requests.post(request_url, params=params, files=files)
#     print(result.text)


@bot.message_handler(commands=['post_products'])
def send_products(message):
    logging.info(f"Sending messages about products to chat with id = {message.from_user.id}")
    bot.send_message(message.from_user.id, "Сейчас завалю здесь все кроссовками")
    for folder_name in os.listdir(folder_magic.path_to_products_to_add):
        send_product_message(message.chat.id, folder_name)
        sleep(10)


@bot.message_handler(commands=['post_product'])
def send_product(message):
    args = extract_arg(message.text)
    if len(args) == 0:
        bot.send_message(message.from_user.id, "Не указан 1-ый аргумент - код продукта")

    send_product_message(message.from_user.id, args[0])


# Обрабатывает обычные текстовые сообщения
@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    print(message)
    if message.text == "/help":
        bot.send_message(message.from_user.id, "Напиши привет")
    elif message.text == 'Поздравить Катю':
        bot.send_message(message.chat.id, "Катя, с днём рождения!")
    else:
        # bot.send_message(message.chat.id, "Я тебя не понимаю. Напиши /help.")
        pass


def send_product_message(msg_id, product_code):
    logging.info(f"Sending message with {product_code} info")
    p_info = folder_magic.deserialize(folder_magic.path_to_products_to_add / product_code)
    logging.info(f"Read following info: {p_info}")
    msg = ProductInfoMessage(p_info["Code"], p_info["Name"], p_info["Price"], p_info["Supplier"], p_info["Material"])

    filenames = [folder_magic.path_to_products_to_add / product_code / file_name for file_name in list(
            set(os.listdir(folder_magic.path_to_products_to_add / product_code))
            - {"desktop.ini", "thumbnail.jpg", "product_info.json"}
    )]
    filenames = list(filter(is_media_file, filenames))
    logging.info(f"Found following media files in product folder: {filenames}")
    if not filenames:
        pass
        # TODO decide if product info should be sent if there are no photo in the folder
        # bot.send_message(chat_id=msg_id, text=msg.print())
    elif len(filenames) == 1:
        if filenames[0].name.split(".")[-1].lower() == 'mp4':
            with open(filenames[0], 'rb') as video:
                logging.info("Sending a video")
                bot.send_photo(chat_id=msg_id, photo=telebot.types.InputMediaVideo(video), caption=msg.print())
        elif filenames[0].name.split(".")[-1].lower() in ['jpg', 'jpeg', 'png']:
            with open(filenames[0], 'rb') as photo:
                logging.info("Sending a photo")
                bot.send_photo(chat_id=msg_id, photo=telebot.types.InputMediaPhoto(photo), caption=msg.print())
    else:
        with ExitStack() as stack:
            files = [
                stack.enter_context(open(filename, 'rb'))
                for filename in filenames
            ]
            media = []
            for file in files:
                if file.name.split(".")[-1].lower() == 'mp4':
                    media.append(telebot.types.InputMediaVideo(file))
                elif file.name.split(".")[-1].lower() in ['jpg', 'jpeg', 'png']:
                    media.append(telebot.types.InputMediaPhoto(file))

            if media:
                media[0].caption = msg.print()

            logging.info("Sending an album")
            bot.send_media_group(chat_id=msg_id, media=media, timeout=200)

# Google API doesn't have functionality to get Image in Cell :(
#
# gc = pygsheets.authorize(service_file='konevod-e884e89f2493.json')
#
# sh = gc.open('Copy of Каталог кроссовки МСК от 2022-03-04')
# wks : pygsheets.Worksheet = sh.worksheet('title', 'Каталог (МСК)')
# cell : pygsheets.Cell = wks.cell('B5')
# img = wks.get_value('C5')
# # result = service.spreadsheets().values().get(
# #     spreadsheetId=spreadsheetId, range=rangeName).execute()
# # values = result.get('values', [])
# print(img)

# send_message()
logging.info("Starting the bot")
bot.polling(none_stop=True, interval=1)
