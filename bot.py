import telebot

bot = telebot.TeleBot('1137709917:AAEZNbC7R-qspaqqdJesoDmHthFHXO0v6qw')

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.send_message(message.chat.id, "Welcome "
                     + message.from_user.first_name
                     + ' ' + message.from_user.last_name + ", I am GreenWrapBot")

bot.polling()
