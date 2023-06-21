import sqlite3

with sqlite3.connect('bd/register.bd') as bd:
    cursor = bd.cursor()
    query = """INSERT INTO expenses (name, password) VALUES(user_name, user_password)   """
    cursor.execute(query)
    bd.commit()