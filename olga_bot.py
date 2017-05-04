from telegram.ext import Updater, CommandHandler, MessageHandler, Filters
import csv
import codecs
import requests
import xlsxwriter
from currency_converter import CurrencyConverter
from datetime import datetime


free_sources = ["(direct)","google","yahoo","yandex","bing","facebook.com", "l.facebook.com", "l.messenger.com"]
report = {}

answers = {
	"привет": "Привет! Я мало говорю, но я готова помочь тебе с аналитикой. Отправь //start, чтобы начать или //help, если тебе требуется помощь",
	"Привет": "Привет! Я мало говорю, но я готова помочь тебе с аналитикой. Отправь //start, чтобы начать или //help, если тебе требуется помощь",
}

c = CurrencyConverter(fallback_on_missing_rate=True) #convert to RUB
c_rate = c.convert(1, 'USD', 'RUB', date=datetime(2016, 12, 31))


def get_answer(question, answers):
	return answers.get(question)


def start(bot, update):
	print("Вызван /start")
	bot.sendMessage(update.message.chat_id, text = "Привет, я Ольга. Я помогу тебе оценить эффективность твоей рекламной кампании.\
		Пожалуйста, отправь мне отчет из Google Analytics для обработки. Набери '\\help', если тебе нужна инструкция по выгрузке отчета")


def help(bot, update):
	print("Вызван /help")
	update.message.reply_text("Пожалуйста, отправь мне отчет из Google Analytics для обработки")
	update.message.reply_text("Чтобы выгрузить отчет, перейди в Google Analytics и выбери отчет «Основные последовательности конверсий», \
		вкладка «Путь источник / канал». Ниже ты увидишь пример отчета. Пожалуйста, корректно укажи период, за который выгружается отчет")
	bot.sendDocument(update.message.chat_id, document=open('help1.png', 'rb'))
	update.message.reply_text("Экспортируй отчет в формат CSV и отправь мне в телеграме. Ниже ты увидишь пример, как экспортировать отчет")
	bot.sendDocument(update.message.chat_id, document=open('help2.png', 'rb'))
	update.message.reply_text("Ниже ты увидишь пример, как файл в телеграме")
	bot.sendDocument(update.message.chat_id, document=open('help3.png', 'rb'))


def talk_to_me(bot, update):
	print("Пришло сообщение: " + update.message.text)
	answer = get_answer(update.message.text, answers)
	try:
		answers[update.message.text] == True
	except KeyError:
		bot.sendMessage(update.message.chat_id, text = "Отправь //start, чтобы начать или //help, если тебе требуется помощь")
	else:
		bot.sendMessage(update.message.chat_id, text = answer)


def csvhandler(bot, update):
	if update.message.document:
		bot.sendMessage(update.message.chat_id, text = "Спасибо. Сейчас я проанализирую отчет")
		file_name = update.message.document.file_id
		print("получен документ " + file_name)
		newFile = bot.getFile(file_name)
		newFile = newFile.download('ga_report.csv')


		with open("ga_report.csv", "r", encoding="utf-8") as f:
			fields = ["source", "conversion", "value"]
			reader = csv.DictReader(f, fields, delimiter=",")
			
			for row in range(6):#this is to skip first 7 lines with non-quantitative data
				next(reader)
			for row in reader:
				sources = row["source"].split(" > ")
				paid_sources = [item for item in sources if item not in free_sources]
				sources_no = len(paid_sources)
				
				if sources_no != 0:   # this part exludes lists contining only free sources (as stated in "free_sources" list)
					cleared_value = float((str(row["value"])).replace(u'\xa0', u'').split(" ")[0].replace(",",".").replace("$",""))
					conversion_no = int(row["conversion"])
					average_value = cleared_value / sources_no / conversion_no
					print(average_value)

					for item in paid_sources:
						try:
							old_value = report[item]
							new_value = old_value + average_value
							report.update({item: new_value})
						except KeyError:
							report.update({item: average_value})
					print(report)

		total_value = 0 #this code calculates total conversion value for the data set
		for item in report.values():
			total_value += item
		print(total_value)

		workbook = xlsxwriter.Workbook("report.xlsx")
		worksheet = workbook.add_worksheet()

		col = 0
		row = 0

		bold = workbook.add_format({'bold': True})
		worksheet.write(row, col, "Источник", bold)
		worksheet.write(row, col + 1, "Ценность конверсии, руб", bold)
		worksheet.write(row, col + 2, "Доля в общей ценности конверсии", bold)
		worksheet.set_column('A:A', 30)
		worksheet.set_column('B:B', 30)
		worksheet.set_column('C:C', 30)

		row = 1

		for key in report.keys():
			worksheet.write(row, col, key)
			worksheet.write(row, col + 1, report[key]*c_rate)
			worksheet.write(row, col + 2, "%.0f%%" % (100 * (report[key]/total_value)))
			row += 1

		# workbook.save('report.xlsx')
		workbook.close()
		# bot.sendChatAction(chat_id=chat_id, action=telegram.ChatAction.TYPING)
		bot.sendDocument(update.message.chat_id, document=open('report.xlsx', 'rb'))
		update.message.reply_text("Отчет проанализирован. Нажми на файл, чтобы сохранить его")


	else:
		bot.sendMessage(update.message.chat_id, text = "Пожалуйста, отправь мне файл в формате csv")


def run_bot():
	updater = Updater("375369666:AAGaxOZ9uIudkF4QQeJOpgUm35467m7wheQ")	
	dp = updater.dispatcher
	dp.add_handler(CommandHandler("start", start))
	dp.add_handler(CommandHandler("help", help))
	dp.add_handler(MessageHandler([Filters.text], talk_to_me))
	dp.add_handler(MessageHandler([Filters.document], csvhandler))

	updater.start_polling()
	updater.idle()

if __name__ == "__main__":
	run_bot()




