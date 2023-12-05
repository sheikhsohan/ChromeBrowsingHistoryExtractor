import sqlite3
import os
import shutil
import pandas as pd

def find_chrome_history():
    # Default paths for Chrome user data on Windows
    default_paths = [
        os.path.expanduser("~") + r"\AppData\Local\Google\Chrome\User Data\Default",
        os.getenv("LOCALAPPDATA") + r"\Google\Chrome\User Data\Default",
        os.getenv("APPDATA") + r"\Google\Chrome\User Data\Default",
    ]

    for path in default_paths:
        history_path = os.path.join(path, "History")
        if os.path.exists(history_path):
            return history_path

    return None

try:
    # Find Chrome history file automatically
    history_db = find_chrome_history()

    if history_db:
        # Copying the history file to another location to avoid being locked by Chrome
        temp_history_db = os.path.join("C:/Users/sohan/Downloads/New folder", "temp_history")
        shutil.copy2(history_db, temp_history_db)

        # Connect to the temporary history file
        conn = sqlite3.connect(temp_history_db)

        # Query to retrieve browsing history
        query = "SELECT urls.visit_count, urls.url, urls.title, visits.visit_time " \
                "FROM urls, visits " \
                "WHERE urls.id = visits.url;"
        cursor = conn.cursor()
        cursor.execute(query)
        
        # Fetch browsing history data
        history_data = cursor.fetchall()
        
        # Create a pandas DataFrame from the fetched data
        columns = ['Visit Count', 'URL', 'Title', 'Visit Time']
        history_df = pd.DataFrame(history_data, columns=columns)
        
        # Save browsing history data to Excel in the user's home directory
        excel_filename = os.path.join(os.path.expanduser("~"), "browsing_history.xlsx")
        history_df.to_excel(excel_filename, index=False)

        print(f"Browsing history saved to {excel_filename}")
    else:
        print("Chrome history file not found.")

except sqlite3.Error as e:
    print(f"SQLite error: {e}")
except Exception as ex:
    print(f"Error: {ex}")
finally:
    # Close the connection and remove the temporary history file
    try:
        conn.close()
        if history_db and os.path.exists(temp_history_db):
            os.remove(temp_history_db)
    except Exception as ex:
        print(f"Error closing connection or removing file: {ex}")
