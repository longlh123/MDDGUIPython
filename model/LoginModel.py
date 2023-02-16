import psycopg2

class LoginModel:
    def __init__(self):
        self.conn = psycopg2.connect(
            dbname="dp_ipsos",
            user="postgres",
            password="googolplex10100!",
            host="localhost",
            port="5432"
        )

    def authenticate(self, username, password):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM staffs WHERE username = %s AND password = %s", (username, password))
        result = cur.fetchone()
        cur.close()
        return result is not None
