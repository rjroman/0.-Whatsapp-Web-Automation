# Importing the necessary libraries
import keyboard as k    # for simulating keyboard keys
import pyautogui         # for automating mouse and keyboard actions
import time              # for adding delay in program execution
import pandas as pd      # for reading and manipulating Excel files
import webbrowser as web # for opening URLs in web browser
from urllib.parse import quote # for URL encoding special characters
import openpyxl

# Open the workbook
workbook = openpyxl.load_workbook('D:/Whatsapp Web Automation/VID 2010/0. Whatsapp Web Automation/Whatsapp List_Main.xlsx')

# Get the active sheet
sheet = workbook.active

# Defining a function named send_whatsapp which takes three arguments:
# data_file_excel: the path to an Excel file containing contact information
# message_file_text: the path to a text file containing the message to be sent
# x_cord and y_cord: the coordinates of the mouse click to be performed (default values are set to the location of the send button in WhatsApp Web)
def send_whatsapp(data_file_excel, message_file_text, x_cord=830, y_cord=954):
    # Reading the Excel file and assigning the contents to a DataFrame named df
    # Also specifying the "Contact" column as a string datatype to prevent Excel from automatically converting phone numbers to scientific notation
    df = pd.read_excel(data_file_excel, dtype={"Contact":str})
    
    # Extracting the values in the "Name" and "Contact" columns of the DataFrame and assigning them to variables named name and contact, respectively
    name = df['Name'].values
    contact = df['Contact'].values
    
    # Reading the contents of the message file and assigning it to a variable named file_data
    with open (message_file_text) as f:
        file_data = f.read()
    
    # Zipping together the values in name and contact into tuples and assigning them to a variable named zipped
    zipped = zip(name, contact)
    
    # Initializing a counter variable to keep track of the number of messages sent
    counter = 0
    
    # Looping over each tuple in zipped
    for (a, b) in zipped:
        # Formatting the message text with the name of the recipient and assigning it to a variable named msg
        msg = file_data.format(a)
        
        # Opening the WhatsApp Web URL for the corresponding contact and message text
        web.open(f"https://web.whatsapp.com/send?phone={b}&text={quote(msg)}")
        
        # Adding a delay to allow the WhatsApp Web page to load
        time.sleep(15)
        
        # Simulating a mouse click at the specified coordinates to send the message
        pyautogui.click(x_cord, y_cord)
        
        # Adding a delay to allow time for the message to be sent
        time.sleep(2)
        
        # Simulating the pressing of the "Enter" key to dismiss the send confirmation pop-up
        k.press_and_release('enter')
        
        # Adding a delay to allow time for the pop-up to be dismissed
        time.sleep(2)
        
        # Simulating the pressing of the "Ctrl + W" keys to close the WhatsApp Web tab
        k.press_and_release('ctrl+w')
        
        # Adding a delay to allow time for the tab to be closed
        time.sleep(1)
        
        # Incrementing the counter variable and printing a message indicating that the message has been sent
        counter += 1
        print(counter, "- Message sent..!!")
    
    # Printing a message indicating that the function has completed execution
    print("Done!")

# Defining the paths to the Excel file and message text file to be used as inputs to the send_whatsapp function
excel_path=r"D:/Whatsapp Web Automation/VID 2010/0. Whatsapp Web Automation/Whatsapp List_Main.xlsx"
text_path=r"D:/Whatsapp Web Automation/VID 2010/0. Whatsapp Web Automation/WHATSDRAFT.txt"

send_whatsapp(excel_path,text_path)
