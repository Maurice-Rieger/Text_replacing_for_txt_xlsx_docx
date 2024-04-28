import tkinter as tk
import time 
import docx
import openpyxl

# We define a graphical user interface 
class Application(tk.Frame): 
    def __init__(self, master= None): 
        super().__init__(master)
        self.master = master 
        self.master.geometry("500x300")
        self.master.config(bg="#343541")
        self.master.title("Text replacing for .txt, .xlsx and .docx")
        self.var = 0 
        self.ausgabename = "data_modified"
        self.eingabename = "data"
        self.create_widgets()
    
    # We specify the widgets and inputs 
    def create_widgets(self): 

        self.text2 = tk.Text(self.master, width=69,height=1, font=("Robolo",10),bg="#343541", fg="#FFFFFF", relief="flat")
        self.text2.place(x=5,y=5)
        self.text2.insert("end","Enter the file path (without file type)")

        self.text5 = tk.Text(self.master, width=30,height=1, font=("Robolo",10),bg="#343541", fg="#FFFFFF", relief="flat")
        self.text5.place(x=5,y=55)
        self.text5.insert("end","To replace")

        self.text6 = tk.Text(self.master, width=30,height=1, font=("Robolo",10),bg="#343541", fg="#FFFFFF", relief="flat")
        self.text6.place(x=150,y=55)
        self.text6.insert("end","Replace with")

        self.text1 = tk.Text(self.master, width=69,height=1, font=("Robolo",10),bg="#343541", fg="#FFFFFF", relief="flat")
        self.text1.place(x=5,y=110)
        self.text1.insert("end","Enter the file type")

        self.text3 = tk.Text(self.master, width=18,height=1, font=("Robolo",10),bg="#343541", fg="#FFFFFF", relief="flat")
        self.text3.place(x=350,y=155)
        self.text3.insert("end",">>Keine Auswahl<<")

        self.text4 = tk.Text(self.master, width=18,height=1, font=("Robolo",10),bg="#343541", fg="#FFFFFF", relief="flat")
        self.text4.place(x=300,y=255)

        self.texteingabe = tk.Text(self.master, width=40,height=1, font=("Robolo",10),bg="#444654", fg="#FFFFFF", relief="flat",highlightbackground="#23242f",highlightthickness=1)
        self.texteingabe.place(x=5,y=30)

        self.Replace = tk.Text(self.master, width=15,height=1, font=("Robolo",10),bg="#444654", fg="#FFFFFF", relief="flat",highlightbackground="#23242f",highlightthickness=1)
        self.Replace.place(x=5,y=75)

        self.Replacer = tk.Text(self.master, width=15,height=1, font=("Robolo",10),bg="#444654", fg="#FFFFFF", relief="flat",highlightbackground="#23242f",highlightthickness=1)
        self.Replacer.place(x=150,y=75)

        self.txtdat = tk.Button(self.master ,font=("Robolo",10),bg="#444654", fg="#FFFFFF", relief="flat")
        self.txtdat["text"] = ".txt"
        self.txtdat["command"] = self.txt
        self.txtdat.place(x=50,y=150)

        self.xlsxdat = tk.Button(self.master, font=("Robolo",10),bg="#444654", fg="#FFFFFF", relief="flat")
        self.xlsxdat["text"] = ".xlsx"
        self.xlsxdat["command"] = self.xlsx
        self.xlsxdat.place(x=150,y=150)

        self.docxdat = tk.Button(self.master, font=("Robolo",10),bg="#444654", fg="#FFFFFF", relief="flat")
        self.docxdat["text"] = ".docx"
        self.docxdat["command"] = self.docx
        self.docxdat.place(x=250,y=150)

        self.convert = tk.Button(self.master, width=20,height=2, font=("Robolo",12),bg="#444654", fg="#FFFFFF", relief="flat")
        self.convert["text"] = "Convert"
        self.convert["command"] = self.conv
        self.convert.place(x=50,y=230)

        self.quit = tk.Button(self.master, text="Quit", fg="red", font=("Robolo",10), relief="flat", bg="#444654",  command=self.master.destroy)
        self.quit.place(x=450,y=250)

    def txt(self): 
        print("Your choice: .txt-file")
        self.var = 1
        self.text3.delete("1.0",tk.END)
        self.text3.insert("end","File type: .txt")

    def xlsx(self): 
        print("Your choice: .xlsx-file")
        self.var = 2 
        self.text3.delete("1.0",tk.END)
        self.text3.insert("end","File type: .xlsx")

    def docx(self): 
        print("Your choice: .docx-file")
        self.var = 3 
        self.text3.delete("1.0",tk.END)
        self.text3.insert("end","File type: .docx")

    def conv(self): 
        print("Conversion started")
        self.text4.delete("1.0",tk.END)
        self.text4.insert("end",">>Completed<<")
        time.sleep(0.5)

        self.inputname = self.texteingabe.get("1.0","end-1c")
        print(f"The input name is: {self.inputname}")

        self.ausgabename = self.inputname + "_modified"
        print(f"The output name is: {self.ausgabename}")
        print(f">>Conversion complete<<")
        time.sleep(0.5)

        self.Replace_string = self.Replace.get("1.0","end-1c")
        self.Replacer_string = self.Replacer.get("1.0","end-1c")


        if self.var == 1: 
            
            self.e = self.inputname+".txt"
            self.a = self.ausgabename+".txt"

            # Read file 
            with open(str(self.e), "r") as input_file:
                input_text = input_file.read()

            # Replacing processs 
            modified_text = input_text.replace(self.Replace_string, self.Replacer_string)

            # Writing modified text in a new file 
            with open(str(self.a), "w") as output_file:
                output_file.write(modified_text)

        elif self.var == 2: 
            print("xlsx conversion")

            self.e = self.inputname+".xlsx"
            self.a = self.ausgabename+".xlsx"

            # Read data 
            workbook = openpyxl.load_workbook(self.e)

            # Choose worksheet 
            worksheet = workbook.active

            # Loop over all rows and columns 
            for row in worksheet.iter_rows():
                for cell in row: 
                        cell.value = str(cell.value).replace(self.Replace_string,self.Replacer_string)
            
            # Save modified worksheet 
            workbook.save(self.a)

        elif self.var == 3: 
            print("docx conversion")

            self.e = self.inputname+".docx"
            self.a = self.ausgabename+".docx"

            # Read file 
            doc = docx.Document(self.e)

            # Replacing process 
            for para in doc.paragraphs: 
                text = para.text 
                new_text = text.replace(self.Replace_string,self.Replacer_string)
                para.text = new_text
            
            # Save modified file 
            doc.save(self.a)

        else: 
            self.text4.delete("1.0",tk.END)
            self.text4.insert("end",">>Select above<<")
            time.sleep(0.5)

root = tk.Tk()
app = Application(master=root)
app.mainloop()

