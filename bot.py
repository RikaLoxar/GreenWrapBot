import telebot
from telebot import types
bot = telebot.TeleBot('1137709917:AAEZNbC7R-qspaqqdJesoDmHthFHXO0v6qw')

markup = types.ReplyKeyboardMarkup(row_width=1)
itembutton1 = types.KeyboardButton("Hi")

@bot.message_handler(['hi','hello'])


bot.polling()
