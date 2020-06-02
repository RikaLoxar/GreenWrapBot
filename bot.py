import telebot

bot = telebot.TeleBot('1137709917:AAEZNbC7R-qspaqqdJesoDmHthFHXO0v6qw')

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, "Привет, создатель")

bot.polling()
