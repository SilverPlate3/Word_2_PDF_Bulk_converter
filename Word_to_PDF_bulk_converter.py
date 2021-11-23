import time
import tkinter.ttk
from tkinter import *
import subprocess
import os, threading, base64


#Creating files from Base64 encoded Hex
def create_files():
    try:
        with open("bulk-convert-Word2PDF.bat", "w") as f1:
            f1.write(base64.b64decode('''ZWNobyBvZmYKZm9yICUlWCBpbiAoKi5kb2N4KSBkbyBjc2NyaXB0LmV4ZSAvL25vbG9nbyBTYXZlQXNQREYuanMgIiUlWCIKZm9yICUlWCBpbiAoKi5kb2MpIGRvIGNzY3JpcHQuZXhlIC8vbm9sb2dvIFNhdmVBc1BERi5qcyAiJSVYIg==''').decode())

        with open("SaveAsPDF.js", "w") as f2:
            f2.write(base64.b64decode('''dmFyIG9iaiA9IG5ldyBBY3RpdmVYT2JqZWN0KCJTY3JpcHRpbmcuRmlsZVN5c3RlbU9iamVjdCIpOwp2YXIgZG9jUGF0aCA9IFdTY3JpcHQuQXJndW1lbnRzKDApOwpkb2NQYXRoID0gb2JqLkdldEFic29sdXRlUGF0aE5hbWUoZG9jUGF0aCk7CiAKdmFyIHBkZlBhdGggPSBkb2NQYXRoLnJlcGxhY2UoL1wuZG9jW14uXSokLywgIi5wZGYiKTsKdmFyIG9ialdvcmQgPSBudWxsOwogCnRyeQp7CiAgICBvYmpXb3JkID0gbmV3IEFjdGl2ZVhPYmplY3QoIldvcmQuQXBwbGljYXRpb24iKTsKICAgIG9ialdvcmQuVmlzaWJsZSA9IGZhbHNlOwogCiAgICB2YXIgb2JqRG9jID0gb2JqV29yZC5Eb2N1bWVudHMuT3Blbihkb2NQYXRoKTsKIAogICAgdmFyIGZvcm1hdCA9IDE3OwogICAgb2JqRG9jLlNhdmVBcyhwZGZQYXRoLCBmb3JtYXQpOwogICAgb2JqRG9jLkNsb3NlKCk7CiAKICAgIFdTY3JpcHQuRWNobygiU2F2aW5nICciICsgZG9jUGF0aCArICInIGFzICciICsgcGRmUGF0aCArICInLi4uIik7Cn0KZmluYWxseQp7CiAgICBpZiAob2JqV29yZCAhPSBudWxsKQogICAgewogICAgICAgIG9ialdvcmQuUXVpdCgpOwogICAgfQp9''').decode())
    except:
        header.config(text="create_files")



def start_convertion():
        try:
            word_path = text_1.get()
            if os.path.exists(word_path):
                os.chdir(word_path)
                global docx_amount
            else:
                exit()
        except:
            header.config(text="start_convertion.1")
        try:
            docx_amount = subprocess.Popen("dir | findstr /E \".docx\" | find \"d\" /c", shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
            docx_amount = int(docx_amount.stdout.read().decode().strip())
        except:
            header.config(text="start_convertion.2.1")
        try:
            event.set()
        except:
            header.config(text="start_convertion.2.2")
        try:
            create_files()
        except:
            header.config(text="start_convertion.2.3")
        os.system("bulk-convert-Word2PDF.bat")
        f = False

        #Checks if the PDF amount is equal to the Word amount
        while True:
            pdfs = subprocess.Popen("dir | findstr /E \".pdf\" | find \"d\" /c", shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
            if docx_amount == int(pdfs.stdout.read().decode().strip()):
                break
            else:
                time.sleep(1.5)

        #Deletes the files we created from the Base64
        os.system("del /f SaveAsPDF.js")
        os.system("del /f bulk-convert-Word2PDF.bat")
        window.destroy()

#Check the conversion progress. If hits 100% it closes the Tkinter window
def see_progress():
    try:
        event.clear()
        event.wait()
        while True or progress['value'] == 100:
            time.sleep(1.5)
            try:
                pdfs = subprocess.Popen("dir | findstr /E \".pdf\" | find \"d\" /c", shell=True, stdin=subprocess.PIPE,stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
                progress['value'] = int((int(pdfs.stdout.read().decode().strip()) / docx_amount)*100)
            except:
                    pass
    except:
        header.config(text="see_progress")

event = threading.Event()
t1 = threading.Thread(target=start_convertion)
t2 = threading.Thread(target=see_progress)
def button_pressed():
    global t1, t2
    t1.start()
    t2.start()


#GUI
window = Tk()

header = Label(window, text="this is a word to PDF converter", bg="pink")
header.grid(row=0, column=0)

space = Label(window, text=" ")
space.grid(row=1, column=0)

path_1 = Label(window, text="Full path of the Word files folder")
path_1.grid(row=2, column=0)

text_1 = Entry(window, width=50)  # C:\Users\אריאל סילבר\Desktop\Test   =  THESE ARE THE FILES TO CONVERTE
text_1.grid(row=3, column=0)


button = Button(window, text="Converte", bg="orange", command=button_pressed)
button.grid(row=4, column=0)

progress = tkinter.ttk.Progressbar(window, orient=HORIZONTAL, length=200, mode='determinate')
progress.grid(row=5, column=0)

warning_1 = Label(window, text="""1. Be sure you are inserting the correct full path, of the folder that holds the .docx files \n
2. If the wrong path was enterd, nothing will happen when you press convert. please close the window and re-open it \n
3. Once the green loading bar is full, please wait another minute until the window closes it self \n
4.     100% Jesus  """, pady=30, bg="red", justify=LEFT)

warning_1.grid(row=6, column=0)

window.mainloop()









