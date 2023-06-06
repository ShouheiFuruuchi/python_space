import mysql.connector

mydb = mysql.connector.connect(
    host="connect ECONNREFUSED 127.0.0.1:3306",
    user="FuruuchiShouhei",
    password="Abcd1829",
)

print(mydb)