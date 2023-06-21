@bot.message_handler(state=MyStates.age, is_digit=True)
def ready_for_answer(message):
    """
    State 3. Will process when user's state is MyStates.age.
    """
    with bot.retrieve_data(message.from_user.id, message.chat.id) as data:
        msg = ("Ready, take a look:\n<b>"
               f"Name: {data['name']}\n"
               f"Surname: {data['surname']}\n"
               f"Age: {message.text}</b>")

        # here is this function
        func_to_use_next_in_code(data)

        bot.send_message(message.chat.id, msg, parse_mode="html")
    bot.delete_state(message.from_user.id, message.chat.id)


def func_to_use_next_in_code(data):
    # do stuff..
    print(data)
