# C:\Users\chris\AppData\Roaming\Python\Python311\Scripts\pyinstaller.exe --onefile -w ...
# Dibuat pada 01/02/2023
# Selesai pada 28/03/2023

# Modul Python
import webbrowser, win32com.client, cv2, random, pyautogui, pyglet, pyttsx3, re, time, os, sys, tkinter as tk
from tkinter import * 
from PIL import ImageTk, Image
from datetime import datetime 


# Register Default Browser & Speaker
webbrowser.register('default', None)
now = datetime.now()
engine = pyttsx3.init()
engine.setProperty('volume', 1)  # Untuk Volume Suara
engine.setProperty('rate', 135)  # Untuk Kecepatan Berbicara

def speak(text):
    engine.say(text)
    engine.runAndWait()


# Define Conversation Pairs / Data Base
pairs = [
    [
        r"my name is (.*)",
        ["Hello %1, nice to meet you."]
    ],

    [
        r"hi|hey|hello|helo|hai|hei|halo|hallo",
        ["Hello!", "Hey there!", "Hi bro!"]
    ],

    [
        r"you are beautiful|you are handsome|you are ugly|you are funny|you are good|i love you|i miss you",
        ["Thankyou but I'm an Artificial Intelligence... I have no body, no feelings and no soul."]
    ],
    
    [
        r"who created you?|what created you?",
        ["Okey, i will tell you who created me."]
    ],
    
    [
        r"what is your name?|who are you?",
        ["I am a chatbot, your personal assistant, You can call me TECO.", "My name is TECO, your personal assistant, nice to meet you!"]
    ],
    
    [
        r"how are you?",
        ["I'm doing good, thankyou.", "I'm fine, thankyou."]
    ],

    [
        r"thanks|thx|thankyou|thank you",
        ["Your welcome bro.", "Okey bro, your welcome."]
    ],
    
    [
        r"tell me a joke",
        ["Why did the tomato turn red? Because it saw the salad dressing! hahaha...",
         "Why do seagulls fly over the sea? Because if they flew over the bay, they'd be bagels! hahaha...",
         "Why don't scientists trust atoms? Because they make up everything! hahaha...",
         "Why do French people like eat snails? Because they don't like fast food! hahaha..."]
    ],
    
    [
        r"define (.*)",
        ["%1 is a word or expression used to describe a particular thing or to ask for information about something.",
         "A definition of %1 is a statement of the meaning of a word or expression, especially in a dictionary."]
    ],
    
    [
        r"tell me about (.*)",
        ["I'm sorry, I don't have enough information about %1. You can switch to the 'search' menu by typing 'a.'", "I'm not sure what you're asking about %1. You can switch to the 'search' menu by typing 'a.'"]
    ],
    
    [
        r"sorry (.*)",
        ["Its alright, no problem.", "Its Okey, never mind.", "Its fine, never mind."]
    ],
    
    [
        r"sorry",
        ["Its alright, no problem.", "Its Okey, never mind.", "Its fine, never mind."]
    ],
    
    [
        r"menu",
        ["Sure, I will tell you about what menus are available in this ChatBot."]
    ],
    
    [
        r"a. (.*)",
        ["Sure, I will search %1"]
    ],
    
    [
        r"a.",
        ["Okey, I will open Google for you."]
    ],
    
    [
        r"b. (.*)",
        ["Sure, I will open youtube %1"]
    ],
    
    [
        r"b.",
        ["Okey, I will open Youtube for you."]
    ],
    
    [
        r"c. (.*)",
        ["Sure, I will open instagram %1"]
    ],
    
    [
        r"c.",
        ["Okey, I will open Instagram for you."]
    ],
    
    [
        r"d. (.*)",
        ["Sure, I will open tiktok %1"]
    ],
    
    [
        r"d.",
        ["Okey, I will open Tiktok for you."]
    ],
    
    [
        r"e.",
        ["Okey, I will open WhatsApp for you."]
    ],
    
    [
        r"f.",
        ["Okey, I will open Music for you."]
    ],

    [
        r"g.",
        ["Okey, I will open Microsoft Word for you."]
    ],

    [
        r"h.",
        ["Okey, I will open Microsoft PowerPoint for you."]
    ],

    [
        r"i.",
        ["Okey, I will open Microsoft Excel for you."]
    ],
    
    [
        r"j.",
        ["Sure, I will open Camera for you."]
    ],
    
    [
        r"k.",
        ["Sure, I will open Game Store for you."]
    ],
    
    [
        r"l.",
        ["Sure, I will open Studio Code for you."]
    ],
    
    [
        r"m.",
        ["Sure, I will tell you the date now."]
    ],
    
    [
        r"n.",
        ["Sure, I will tell you the time now."]
    ],
    
    [
        r"o.",
        ["Sure, I will open Alarm for you."]
    ],
    
    [
        r"p.",
        ["Okey, I will restart your laptop."]
    ],
    
    [
        r"q.",
        ["Okey, I will turn off your laptop."]
    ],
    
    [
        r"r.",
        ["Goodbye, take care of yourself and always be happy... See you soon bro."]
    ],
    
    [
        r"s.",
        ["Okey bro, I will tell you how to use me."]
    ],
    
    [
        r"How to Use?|How to use?",
        ["Okey bro, I will tell you how to use me."]
    ],
]



# Tampilan GUI (Tkinter) 
root = tk.Tk()
root.geometry("570x370")  # Ukuran Tampilan GUI (Tkinter)
root.title("MyChatBot - TECO")  # Title pada Tampilan GUI (Tkinter)
bg_color = '#9EEDF2'
root.config(bg=bg_color)  # Untuk Warna Background pada Tampilan GUI (Tkinter) 
root.iconbitmap('AnyConv.com__robot.ico')  # Mengganti Icon pada Tampilan GUI (Tkinter) 



# Function untuk Program TECO
def search_in_google(query):
    webbrowser.open(f"https://www.google.com/search?q={query}")  # Untuk Search Google


def search_in_youtube(query):
    webbrowser.open(f"https://www.youtube.com/search?q={query}")  # Untuk YouTube


def search_in_instagram(query):
    webbrowser.open(f"https://www.instagram.com/{query}/")  # Untuk Instagram


def search_in_tiktok(query):
    webbrowser.open(f"https://www.tiktok.com/search?q={query}")  # Untuk Tiktok


def search_in_wa(query):
    webbrowser.open(f"https://web.whatsapp.com/")  # Untuk WhatsApp


def search_in_game(query):
    webbrowser.open(f"https://www.microsoft.com/id-id/store/top-free/games/pc")  # Untuk Game Store


def search_in_code(query):
    webbrowser.open(f"https://replit.com/~")  # Untuk Studio Code


def update_time():  # Function untuk Tampilan Waktu pada Tampilan GUI (Tkinter)
    current_time = datetime.now().strftime('%H:%M:%S')
    time_label.config(text=current_time)
    root.after(1000, update_time)


def aku():  # Function untuk Kalimat Pembuka pada Tampilan GUI (Tkinter)
    speak("Hi, my name is Teco. How can I help you today?")


def generate_response(user_input):  # Function untuk merubah %1 menjadi Kalimat
    for pattern, responses in pairs:
        match = re.match(pattern, user_input)
        if match:
            response = random.choice(responses)
            if isinstance(response, list):
                response = random.choice(response)
            if '%1' in response:
                response = response.replace('%1', match.group(1))
            if 'open_alarm' in response.lower():
                open_alarm()
            return response


def open_alarm():  # Function untuk Tampilan Set Alarm pada Tampilan GUI ke-2 (Tkinter)
    def set_alarm():
        alarm_time = f"{hour_entry.get()}:{minute_entry.get()}:{second_entry.get()}"
        while True:
            current_time = time.strftime('%H:%M:%S')
            if current_time == alarm_time:
                alarm_label.config(text="<< Alarm ringing! Wake Up bro! >>",  font=("Helvetica", 10, "bold"))
                sound = pyglet.media.load('suara_alarm.mp3', streaming=False)
                sound.play()
                break
            else:
                alarm_label.config(text=f"<< Alarm set for {alarm_time} >>")
            window.update()
            time.sleep(1)
        hour_entry.delete(0, tk.END)
        minute_entry.delete(0, tk.END)
        second_entry.delete(0, tk.END)

    window = tk.Tk()
    window.title("Alarm - TECO")
    window.geometry("382x200")
    bg_color2 = '#9EEDF2'
    window.config(bg=bg_color2)
    window.iconbitmap('AnyConv.com__robot.ico')
    
    label1 = tk.Label(window, text="⏰", font=("Helvetica", 17, "bold"), bg=bg_color2)
    label1.grid(row=0, column=0)

    label1 = tk.Label(window, text="== ALARM ==", font=("Helvetica", 17, "bold"), bg=bg_color2)
    label1.grid(row=0, column=1)

    hour_label = tk.Label(window, bg=bg_color2, text="Hour :", font=("Helvetica", 10))
    hour_label.grid(row=1, column=0)

    hour_entry = tk.Entry(window, font="Helvetica 15", border=3)
    hour_entry.grid(row=1, column=1)

    minute_label = tk.Label(window, bg=bg_color2, text="Minute :", font=("Helvetica", 10))
    minute_label.grid(row=2, column=0)

    minute_entry = tk.Entry(window, font="Helvetica 15", border=3)
    minute_entry.grid(row=2, column=1)

    second_label = tk.Label(window, bg=bg_color2, text="Second :", font=("Helvetica", 10))
    second_label.grid(row=3, column=0)

    second_entry = tk.Entry(window, font="Helvetica 15", border=3)
    second_entry.grid(row=3, column=1)

    set_alarm_button = tk.Button(window, bg=bg_color2, text="Set Alarm", font=("Helvetica", 10, "bold"), command=lambda: set_alarm())
    set_alarm_button.configure(background='black', foreground='white')
    set_alarm_button.grid(row=4, column=0)

    alarm_label = tk.Label(window, bg=bg_color2, text="\n<< Alarm is Not Set >>",  font=("Helvetica", 10, "bold"))
    alarm_label.grid(row=4, column=1, columnspan=2)

    window.mainloop()


def open_menu(root):  # Function untuk Tampilan Menu pada Tampilan GUI ke-2 (Tkinter)
    menu_window = tk.Toplevel(root)
    menu_window.title("ChatBot Menu - TECO")
    menu_window.geometry("540x780")
    bg_color1 = '#9EEDF2'
    menu_window.config(bg=bg_color1)
    menu_window.iconbitmap('AnyConv.com__robot.ico')

    label = tk.Label(menu_window, text="📜 ChatBot Menu 📜\n", font=("Helvetica", 17, "bold"), bg=bg_color1)
    label.pack(padx=20, pady=20)
    label = tk.Label(menu_window, text="\t[ a. ]  Search ... 🔎", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ b. ]  open YouTube ... 📺", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ c. ]  open Instagram ... 🎬", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ d. ]  open Tiktok ... 🎶", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ e. ]  open WhatsApp 📞", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ f. ]  open Music (Spotify) 🎵", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ g. ]  open Microsoft Word 📘", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ h. ]  open Microsoft PowerPoint 📙", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ i. ]  open Microsoft Excel 📗", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ j. ]  open Camera 📷", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ k. ]  open Game Store 🎮", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ l. ]  open Studio Code (Replit) 👨‍💻", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ m. ]  Date 📅", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ n. ]  Time 🕑", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ o. ]  Set Alarm ⏰", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ p. ]  Restart Laptop 💻🔃", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ q. ]  Shutdown Laptop 💻📴", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ r. ]  Quit Program 📤", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')
    label = tk.Label(menu_window, text="\t[ s. ]  How to Use ❓", font=("Helvetica", 13), bg=bg_color1)
    label.pack(pady=1, anchor='w')


def gambar_gw(root):  # Function untuk Tampilan Gambar Developer pada Tampilan GUI ke-2 (Tkinter)
    gambar1 = tk.Toplevel(root)
    gambar1.title("Developer of TECO")
    gambar1.geometry("500x607")
    bg_color3 = '#9EEDF2'
    gambar1.config(bg=bg_color3)
    gambar1.iconbitmap('AnyConv.com__robot.ico')

    # Load Gambar Ke-1
    nama_1 = tk.Label(gambar1, bg=bg_color3, text="Mr. Christian Julianto S.\n(Main Developer)", font=("Helvetica", 13, 'bold'))
    nama_1.grid(row=0, column=0)
    image_path1 = "foto_christian.jpg"
    gambar_1 = Image.open(image_path1)
    gambar_1 = gambar_1.resize((250, 300))
    photo1 = ImageTk.PhotoImage(gambar_1)
    label1 = tk.Label(gambar1, image=photo1, bg=bg_color3)
    label1.image = photo1  # Simpan Referensi dari Objek Gambar
    label1.grid(row=0, column=1)

    # Load Gambar Ke-2
    nama_2 = tk.Label(gambar1, bg=bg_color3, text="Mr. Reagen Alvey R.\n(Deputy Developer)", font=("Helvetica", 13, 'bold'))
    nama_2.grid(row=1, column=0)
    image_path2 = "foto_reagen.jpg"
    gambar_2 = Image.open(image_path2)
    gambar_2 = gambar_2.resize((250, 300))
    photo2 = ImageTk.PhotoImage(gambar_2)
    label2 = tk.Label(gambar1, image=photo2, bg=bg_color3)
    label2.image = photo2  # Simpan Referensi dari Objek Gambar
    label2.grid(row=1, column=1)
    
    # Buat Frame Kosong untuk Memperbarui Tampilan Gambar
    frame = tk.Frame(gambar1, bg=bg_color3)
    frame.grid(row=2, column=0, columnspan=2, sticky="nsew")
    frame.bind("<Configure>", lambda e: frame.config(width=e.width, height=e.height))


def use_(root):  # Function untuk Tampilan How to Use pada Tampilan GUI ke-2 (Tkinter)
    how_window = tk.Toplevel(root)
    how_window.title("How To Use ChatBot - TECO")
    how_window.geometry("920x380")
    bg_color11 = '#9EEDF2'
    how_window.config(bg=bg_color11)
    how_window.iconbitmap('AnyConv.com__robot.ico')

    label_1 = tk.Label(how_window, text="❓🪄 How To Use ChatBot 🪄❓\n", font=("Helvetica", 17, "bold"), bg=bg_color11)
    label_1.pack(padx=20, pady=20)
    label_1 = tk.Label(how_window, text="\t1.  Saat Anda ingin menggunakan pilihan di Menu, Anda cukup ketik \"a. ...\",\n  \"b. Japanese Music\" atau \"d.\" pada kolom isi yang tersedia.\n", font=("Helvetica", 13), bg=bg_color11)
    label_1.pack(pady=1, anchor='w')
    label_1 = tk.Label(how_window, text="\t2.  Saat Anda ingin keluar dari Menu Camera \"j.\", Anda cukup\n    klik \"Esc\" pada keyboard Laptop atau PC Anda.\n", font=("Helvetica", 13), bg=bg_color11)
    label_1.pack(pady=1, anchor='w')
    label_1 = tk.Label(how_window, text="\t3.  Saat Anda ingin mengetahui siapa pencipta dari Program / Aplikasi\n\tMy ChatBot - TECO, Anda dapat mengetik \"who created you\".", font=("Helvetica", 13), bg=bg_color11)
    label_1.pack(pady=1, anchor='w')


def show_text():  # Function untuk Kalimat Data Diri Developer - TECO
    answer1 = "What created me was Mr. Christian Julianto as the main developer and Mr. Reagan Alvey as the deputy developer in 2023, they are very smart, kind and handsome. Mr Christian is currently 17 years old and lives in West Jakarta. Mr Reagan is currently 16 years old and lives in North Jakarta. They are one of the students at the Jakarta Telkom Vocational School."
    speak(answer1)


def chat_with_teco(*args):  # Function untuk Perintah TECO pada Tampilan GUI (Tkinter)
    msg = input1.get()
    response = generate_response(msg)
    speak(response)
    input1.delete(0,"end")

    if "who created you" in msg or "who created you?" in msg or "what created you" in msg or "what created you?" in msg:
        gambar_gw(root)
        root.after(500, show_text)
    
    elif "menu" in msg:
        open_menu(root)

    elif "a." in msg:  # Untuk Search Google
        search_query = msg.split("a.")[-1].strip()
        search_in_google(search_query)

    elif "b." in msg:  # Untuk YouTube
        search_query = msg.split("b.")[-1].strip()
        search_in_youtube(search_query)

    elif "c." in msg:  # Untuk Instagram
        search_query = msg.split("c.")[-1].strip()
        search_in_instagram(search_query)

    elif "d." in msg:  # Untuk Tiktok
        search_query = msg.split("d.")[-1].strip()
        search_in_tiktok(search_query)

    elif "e." in msg:  # Untuk WhatsApp
        search_query = msg.split("e.")[-1].strip()
        search_in_wa(search_query)

    elif "f." in msg:  # Untuk Music
        pyautogui.hotkey('win', 'r')
        pyautogui.typewrite('spotify')
        pyautogui.press('enter')

    elif "g." in msg:
        # Membuka aplikasi Microsoft Word
        word = win32com.client.Dispatch("Word.Application")
        # Membuat dokumen baru
        word.Documents.Add()
        word.Visible = True

    elif "h." in msg:
        # Membuka aplikasi Microsoft PowerPoint
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        # Membuat presentasi baru
        powerpoint.Presentations.Add()
        powerpoint.Visible = True

    elif "i." in msg:
        # Membuka aplikasi Microsoft Excel
        excel = win32com.client.Dispatch("Excel.Application")
        # Membuat workbook baru
        excel.Workbooks.Add()
        excel.Visible = True

    elif "j." in msg:  # Untuk Camera
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            raise Exception("Could not open camera")
        while True:
            ret, frame = cap.read()
            cv2.imshow("Camera", frame)
            key = cv2.waitKey(1)
            if key == ord('q') or key == 27:
                break
        cap.release()
        cv2.destroyAllWindows()

    elif "k." in msg:  # Untuk Game
        search_query = msg.split("k.")[-1].strip()
        search_in_game(search_query)

    elif "l." in msg:  # Untuk Studio Code
        search_query = msg.split("l.")[-1].strip()
        search_in_code(search_query)

    elif "m." in msg:  # Untuk Date
        current_date = datetime.now().strftime("%d/%m/%Y")
        speak(f"Today is {current_date}")

    elif "n." in msg:  # Untuk Time
        current_time = datetime.now().strftime("%H:%M:%S")
        speak(f"Now at {current_time}")

    elif 'o.' in msg:  # Untuk Set Alarm 
        open_alarm()

    elif "p." in msg:  # Untuk Restart Laptop
        os.system("shutdown /r /t 1")

    elif "q." in msg:  # Untuk Shutdown Laptop
        os.system("shutdown /s /t 1")
    
    elif "r." in msg:  # Untuk Close Program
        sys.exit()
    
    elif "s." in msg or "how to use" in msg or "how to use?" in msg or "How to use" in msg or "How to use?" in msg or "how to Use" in msg or "how to Use?" in msg or "How to Use" in msg or "How to Use?" in msg:  # Untuk Close Program
        use_(root)

    else:
        speak("amogusss")
        
    input1.delete(0,"end")



# Membuat Label (Judul)
welcome_label = Label(root, text="🤖 Welcome To MyChatBot 🤖\n", font=("Helvetica", 18, "bold"), bg=bg_color)
welcome_label.pack(pady=10)
frame = tk.Frame(root)
frame.pack()

# Membuat Label (Perintahkan Chatbot)
label1 = tk.Label(frame, text="You: ", font=("Calibri", 17), bg=bg_color)
label1.grid(row=0, column=0)
input1 = tk.Entry(frame, font="Normal 15", border=3)
input1.grid(row=0, column=1)
input1.bind("<Return>", chat_with_teco)

# Menambahkan Logo TECO
image_path = "Logo_TECO.png"
image1 = Image.open(image_path)
image1 = image1.resize((72, 28), Image.ANTIALIAS)   # Merubah Ukuran Gambar
img1 = ImageTk.PhotoImage(image1)   # Konversi Gambar ke Format Tkinter
panel1 = tk.Label(root, image=img1, bg=bg_color)   # Membuat Label
panel1.pack(side="bottom", fill="both", expand="yes")

# Menambahkan Gambar Robot
image = Image.open("robot.png")
image = image.resize((200, 200), Image.ANTIALIAS)   # Merubah ukuran gambar
img = ImageTk.PhotoImage(image)   # Konversi Gambar ke Format Tkinter
panel = tk.Label(root, image=img, bg=bg_color)   # Membuat Label
panel.pack(side="top", fill="both", expand="yes")

# Menambahkan Jam / Waktu
time_label = tk.Label(root, text="", font=("Helvetica", 15), pady=20, bg=bg_color)
time_label.pack(side="top", fill="both", expand="yes")
time_label.place(relx=0.01, rely=1.05, anchor=SW)

# Menambahkan Tombol Menu
text_file_button = Button(root, text="MENU", font=("Helvetica", 10, "bold"), command=lambda: open_menu(root))
text_file_button.pack()
text_file_button.configure(background='black', foreground='white')
text_file_button.place(relx=0.988, rely=0.99, anchor=SE)


# Memanggil Semua Perintah
if __name__ == "__main__":
    update_time()
    aku()
    root.mainloop()