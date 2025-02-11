import pandas as pd
import requests, time
from datetime import datetime

file_path = "Data.xlsx"  # Replace with the actual path to your Excel file

def updateTimeNow():
    now = datetime.now()
    current_date = now.strftime("%d%b%Y")
    current_time = now.strftime("_%H%M%S") 
    return str(current_date)+str(current_time)

def timestamp():
    now = datetime.now()
    current_date = now.strftime("%d-%b-%Y")
    current_time = now.strftime(" %H:%M:%S") 
    return str(current_date)+str(current_time)

# Example usage:

logsFiles = "logs/logs_"+updateTimeNow()+".csv"
# print(logsFiles)

def appendToFile(text=""):
    with open(logsFiles, "a") as f:
        f.write(text+"\n")


def hitAPI(NAME=" ", PHONE="", GENDER="MALE", item=0):
    # API Endpoint
    url = "https://samarth.prod.api.sapioglobal.com/mysba/form/"

    # JSON Payload
    payload = {
        "name": NAME,
        "contact_number": str(PHONE),
        "email": "",
        "age": "26-40",
        "gender": str(GENDER),
        "state": "Haryana",
        "district": "Yamunanagar",
        "approval1": 1,
        "approval2": 1,
        "approval3": 1
    }

    # Headers
    headers = {
        "Content-Type": "application/json"  # Ensures the request is sent as JSON
    }

    # Sending POST Request
    response = requests.post(url, json=payload, headers=headers)

    # Print Response
    print("Status Code:", response.status_code)
    print("Response:", response.json())  # If response is JSON, else use response.text

    appendString = f"{item};{NAME};{PHONE};{GENDER};{response.json()["status"]};{response.json()["message"]};{timestamp()}"
    appendToFile(appendString)

def read_excel_and_store(excel_file_path):
    try:
        df = pd.read_excel(excel_file_path)

        # Check if the expected columns exist.  This makes the code more robust.
        required_columns = ["S.No.", "Name", "Phone No"]  # Adjust if your headers are different
        if not all(col in df.columns for col in required_columns):
            print(f"Error: Excel file must contain columns: {required_columns}")
            return {}  # Return empty dict on error

        data_dict = {}
        for index, row in df.iterrows():
            try:  # Handle potential type errors (e.g., non-numeric serial numbers)
                serial_number = int(row["S.No."])  # Convert to integer
                name = str(row["Name"])  # Convert to string
                phone_number = str(row["Phone No"])  # Convert to string

                data_dict[serial_number] = {"name": name, "phone_number": phone_number}
            except (ValueError, TypeError) as e:
                print(f"Error processing row {index + 1}: {e}. Skipping this row.") # Provide more context
                continue  # Skip to the next row if there's a problem

        return data_dict

    except FileNotFoundError:
        print(f"Error: Excel file not found at {excel_file_path}")
        return {}
    except Exception as e:  # Catch other potential errors (e.g., invalid Excel format)
        print(f"An error occurred: {e}")
        return {}

result_dict = read_excel_and_store(file_path)


for item in result_dict:
    name = result_dict[item]["name"]
    phone = result_dict[item]["phone_number"]
    gender = "Male" if (item <= (len(result_dict) / 2)) else "Female"
    hitAPI(NAME=name, PHONE=phone, GENDER=gender, item=item)
    print(item , result_dict[item]["name"], result_dict[item]["phone_number"])
    




#Just before EXIT
print("All enteries done")
print("window will autoclose in 10 seconds")
for i in range (10):
    print(10-i)
    time.sleep(1)
