from aiogram.fsm.state import State, StatesGroup

class Form(StatesGroup):
    choose_start = State()
    choose_help = State()
    choose_groupps = State()
    start_analyse = State()

if __name__ == '__main__':
    print(Form.all())
