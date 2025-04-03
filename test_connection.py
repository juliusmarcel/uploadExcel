import pyodbc

def get_table_names():
    try:
        conn = pyodbc.connect(
            "DRIVER={ODBC Driver 17 for SQL Server};"
            "SERVER=LAPTOP-FDEOFA5T;"
            "DATABASE=UploadExcelDB;"
            "UID=excel_app_user;"
            "PWD=Pass123!;"
            "Trusted_Connection=no;"
            "Encrypt=no;"
        )
        cursor = conn.cursor()

        # Query untuk mendapatkan nama semua tabel dalam database
        cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'")
        tables = cursor.fetchall()

        if tables:
            print("✅ Daftar tabel dalam database UploadExcelDB:")
            for table in tables:
                print(f"- {table[0]}")
        else:
            print("⚠️ Tidak ada tabel ditemukan dalam database.")
        
        conn.close()

    except pyodbc.Error as e:
        print(f"❌ GAGAL: {e.args[1]}")

if __name__ == '__main__':
    get_table_names()
