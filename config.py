import psycopg2
conn = psycopg2.connect(
    dbname="DFCCIL",
    user="postgres",
    password="Tejpal@123",
    host="localhost",
    port=5432
)
cur = conn.cursor()
