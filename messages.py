from utils import Form

help_messages = 'Данный бот считывает информацию с файла Excel данные, '\
                f'которые ... \n' \
                f'Чтобы начать работу, отправь "/start".'

start_message = f'Добрый день!\n Для начала работы следует выбрать группу командой "/groupp"' + help_messages
invalid_key = "Ключ не подходит. \n" + help_messages
choose_groupps = "Выберите группу ..."
start_analyse = "Начинаю считывать данные с файла"

MESSAGES = {
    'start': start_message,
    'help': help_messages,
    'groupp': choose_groupps,
    'analyse': start_analyse
}