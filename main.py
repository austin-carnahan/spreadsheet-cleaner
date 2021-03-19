import tkinter as tk
from tkinter import ttk, Menu, filedialog, messagebox, font
from configparser import ConfigParser
from local_modules.spreadcheck import spreadcheck
from local_modules.autocorrect import autocorrect
import sys, subprocess, platform, os, shutil, glob, atexit
import pandas as pd


class tkinterApp(tk.Tk):
    # __init__ function for class tkinterApp  
    def __init__(self, *args, **kwargs):  
          
        # __init__ function for class Tk 
        tk.Tk.__init__(self, *args, **kwargs)

        # get previous file paths from config.ini
        self.config_file = ConfigParser()
        self.config_file.read('config.ini')

        # TODO: decouple data & paths from this class
        # set initial values for file paths & data
        self.rules_file = self.config_file['FILE_PATHS']['rules_file']
        self.write_dir = self.config_file['FILE_PATHS']['write_dir']
        self.temp_dir = self.config_file['FILE_PATHS']['temp_dir']
        self.default_save_filename = tk.StringVar()
        self.data_file = ""

        self.titlefont = font.Font(family="Verdana", size=12, weight="bold")
        
        # create top bar menu
        menubar = MenuBar(self)
        self.config(menu=menubar)
        self.geometry("560x360")
          
        # creating a container 
        container = tk.Frame(self)   
        container.pack(side = "top", fill = "both", expand = True)  
   
        container.grid_rowconfigure(0, weight = 1) 
        container.grid_columnconfigure(0, weight = 1) 

        # initializing frames to an empty array 
        self.frames = {}   
   
        # iterating through a tuple consisting 
        # of the different page layouts 
        for F in (StartPage, Setup, Validate, SaveExit): 
   
            frame = F(container, self) 
   
            # initializing frame of that object from 
            # startpage, Setup, Validate respectively with  
            # for loop 
            self.frames[F] = frame  
   
            frame.grid(row=0, column=0, sticky="nsew") 
   
        self.show_frame(StartPage) 
   
    # to display the current frame passed as parameter 
    def show_frame(self, cont): 
        frame = self.frames[cont] 
        frame.tkraise()

    def open_file(self, item):
        if platform.system() == 'Darwin':       # macOS
            subprocess.call(('open', item))
        elif platform.system() == 'Windows':    # Windows
            subprocess.call(('start', item), shell=True)
        else:                                   # linux variants
            subprocess.call(('xdg-open', item))
            

class StartPage(tk.Frame): 
    def __init__(self, parent, controller):  
        tk.Frame.__init__(self, parent)

        # label 
        label = ttk.Label(self, text ="Spreadsheet Cleaner", font= controller.titlefont) 
        label.grid(row = 0, column = 0, padx = 10, pady = 10, ipadx=12)

        # intro  
        intro = tk.Text(self, height=16, width=64, relief="flat", wrap="word", bg='light grey', fg="black")
        intro.grid(row = 1, column=0, padx = 20, pady = 10, ipadx=4)
        intro.insert("end",
            "Welcome to the spreadsheet validation and cleaning tool! Before beginning, make sure you have a valid rules file for processing.\n\n" \
            "1. Select the file you would like to clean and the rules file to use for checking entries \n\n" \
            "2. Run validation. Then inspect results, manually update spreadsheet entries," \
            " and add or edit rules as needed. Save spreadsheet changes" \
            " to see them reflected in the tool. \n\n" \
            "3. Select a folder and filename for saving the cleaned data  \n\n" \
            "Note: This tool may be used to process sensitive" \
            " information. No data is stored or sent to external resources."
        )
        intro.config(state="disabled")
        
        # Navigate to next frame
        button = ttk.Button(self, text ="Start", command = lambda : controller.show_frame(Setup)) 
        button.grid(row = 2, column = 0, padx = 12, pady = 8, sticky="se")


class Setup(tk.Frame): 
      
    def __init__(self, parent, controller): 
        tk.Frame.__init__(self, parent) 

        label = ttk.Label(self, text ="Setup", font= controller.titlefont) 
        label.grid(row = 0, column = 0, padx = 10, pady = 10)

        # data file container
        data_file_frame = ttk.LabelFrame(self, text="Input Data File:")
        data_file_frame.grid(row=1, pady = 10, padx = 16, ipadx=20)

        # button for selecting data file
        data_file_button = ttk.Button(data_file_frame, text="Choose File", command= lambda: self.choose_data_file(controller))
        data_file_button.grid(row = 0, column = 0, padx = 10, pady = 10, sticky = "w")

        # diplay for file selection
        self.file_entry_text = tk.StringVar()
        self.data_file_entry = tk.Entry(data_file_frame, relief="sunken", textvariable=self.file_entry_text, width=60)
        self.data_file_entry.grid(row = 0, column = 1)

        # rules file container
        rules_file_frame = ttk.LabelFrame(self, text="Rules File:")
        rules_file_frame.grid(row=2, pady = 10, padx = 16, ipadx=20)

        # button for selecting rules file
        rules_file_button = ttk.Button(rules_file_frame, text="Choose File", command= lambda: self.choose_rules_file(controller))
        rules_file_button.grid(row = 0, column = 0, padx = 10, pady = 10, sticky = "w")

        # diplay for rules selection
        self.rules_entry_text = tk.StringVar()
        self.rules_entry_text.set(controller.rules_file) #use saved path in config file
        self.rules_file_entry = tk.Entry(rules_file_frame, relief="sunken", textvariable=self.rules_entry_text, width=60)
        self.rules_file_entry.grid(row = 0, column = 1)
        self.rules_file_entry.xview_moveto(1) 
   
        # Navigate to next frame
        button1 = ttk.Button(self, text ="Next", command = lambda : self.validate_selections(controller)) 
        button1.grid(row = 7, column= 0, padx = 10, pady = 10, sticky="e,s")

    def validate_selections(self, controller):
        if controller.rules_file and controller.data_file:
            controller.show_frame(Validate)
        else:
            messagebox.showerror("Missing Path", "No empty fields allowed.")
        

    def choose_data_file(self, controller):
        controller.data_file = filedialog.askopenfilename(filetypes=[
            ("spreadsheeet format", ".xlsx"),
            ("spreadsheeet format", ".xls"),
            ("spreadsheet format", ".csv")
        ])

        base = os.path.basename(controller.data_file)
        controller.default_save_filename.set("CLEANED_" + os.path.splitext(base)[0])
        
        self.file_entry_text.set(controller.data_file)
        self.data_file_entry.xview_moveto(1)
        
    def choose_rules_file(self, controller):
        controller.rules_file = filedialog.askopenfilename(filetypes=[
            ("spreadsheeet format", ".xlsx")
        ])

        self.rules_entry_text.set(controller.rules_file)
        self.rules_file_entry.xview_moveto(1)
        
        # write file path to config.ini so it's stashed for next time
        try:
            controller.config_file['FILE_PATHS']['rules_file'] = controller.rules_file
        except:
            controller.config_file['FILE_PATHS']['rules_file'] = ""

        with open('config.ini', 'w') as update:
            controller.config_file.write(update)

   
class Validate(tk.Frame):  
    def __init__(self, parent, controller): 
        tk.Frame.__init__(self, parent)

        label = ttk.Label(self, text ="Validate", font= controller.titlefont) 
        label.grid(row = 0, column = 0, padx = 10, pady = 10) 

        # options container
        options_frame = ttk.LabelFrame(self)
        options_frame.grid(row=1, pady = 10, padx = 10, ipadx = 20)
        options_frame.grid_columnconfigure(1, weight=1)

        # Run Validation button
        validate_button = ttk.Button(options_frame, text="Run Validation", command= lambda: self.run_validation(controller, self.progress))
        validate_button.grid(row = 0, column = 0, padx = 12, pady = 10, sticky = "w")

        # View Invalid Entries Button
        self.data_file_button = ttk.Button(options_frame, text="View Flagged Entries", command= lambda: controller.open_file(controller.data_file))
        self.data_file_button.grid(row = 0, column = 1, padx = 12, pady = 10, sticky = "ew")
        self.data_file_button.state(["disabled"])
        
        # View Rules File
        rules_file_button = ttk.Button(options_frame, text="Open Rules File", command= lambda: controller.open_file(controller.rules_file))
        rules_file_button.grid(row = 0, column = 2, padx = 12, pady = 10, sticky = "e")

        # Progress Bar
        self.progress = ttk.Progressbar(self, orient = "horizontal", length = 200, mode = 'determinate', maximum = 101)
        self.progress.grid(row=2, padx = 10, pady = 14, ipadx=100)

        # Message
        self.message = tk.Text(self, height=5, width=60, bg='#D3D3D3', fg="black")
        self.message.grid(row=3, padx = 12, pady = 10, ipadx= 20)
        scroll = tk.Scrollbar(self, command=self.message.yview)
        self.message.configure(yscrollcommand=scroll.set)
        self.message.config(state="disabled")

   
        # Navigate to next frame
        self.next_button = ttk.Button(self, text ="Next", command = lambda : controller.show_frame(SaveExit)) 
        self.next_button.grid(row = 4, column = 0, padx = 10, pady = 10, sticky="e")
        self.next_button.state(["disabled"])

    def run_validation(self, controller, progress):
        
        results = spreadcheck.clean(controller.data_file, autocorrect.fields(controller.rules_file), controller.temp_dir, progress)
        
        if results.errors:
            self.message.config(state="normal")
            self.message.delete(1.0, "end")
            self.message.configure(fg="red")
            for error in results.errors:
                self.message.insert("end", error + '\n')
            self.message.config(state="disabled")
            return
        else:
            self.data_file_button.state(["!disabled"])
            self.next_button.state(["!disabled"])
            self.message.config(state="normal")
            self.message.delete(1.0, "end")
            self.message.configure(fg="green")
            for msg in results.messages:
                self.message.insert("end", msg + '\n')
            self.message.config(state="disabled")
            controller.data_file = results.spreadsheet

class SaveExit(tk.Frame):  
    def __init__(self, parent, controller): 
        tk.Frame.__init__(self, parent)

        label = ttk.Label(self, text ="Save & Exit", font= controller.titlefont) 
        label.grid(row = 0, column = 0, padx = 10, pady = 10)

        # write dir container
        write_dir_frame = ttk.LabelFrame(self, text="Destination:")
        write_dir_frame.grid(row=1, pady = 10, padx = 12, ipadx= 14)
        
        # button for selecting write dir
        write_dir_button = ttk.Button(write_dir_frame, text="Choose Folder", command= lambda: self.choose_write_dir(controller))
        write_dir_button.grid(row = 0, column = 0, padx = 10, pady = 10, sticky = "w")
        
         # diplay for write dir selection
        self.write_dir_entry = tk.Text(write_dir_frame, bg="light grey", wrap="char", relief="flat", height=2, width=40)
        self.write_dir_entry.insert("end", controller.write_dir)
        self.write_dir_entry.grid(row = 0, column = 1, padx = 10, pady = 10, sticky="w")
        self.write_dir_entry.config(state="disabled")

        # file name
        self.write_name_label = ttk.Label(write_dir_frame, text="File Name: ")
        self.write_name_label.grid(row=1, column=0, padx = 10, pady = 10)
        self.write_name_entry = tk.Entry(write_dir_frame, textvariable=controller.default_save_filename, width=40)
        self.write_name_entry.grid(row=1, column=1, padx = 10, pady = 10, sticky="w")
        self.write_name_extension = ttk.Label(write_dir_frame, text=".csv")
        self.write_name_extension.grid(row=1, column=2, pady = 10, sticky="w")

        # Save button
        self.save_button = ttk.Button(write_dir_frame, text ="Save", command = lambda : self.save_file(controller)) 
        self.save_button.grid(row = 2, column = 1, padx = 10, pady = 10, sticky="e")

        # Message
        self.message = tk.Text(self, height=3, width=64, bg='#D3D3D3', fg="black")
        self.message.grid(row=2, padx = 10, pady = 10)
        scroll = tk.Scrollbar(self, command=self.message.yview)
        self.message.configure(yscrollcommand=scroll.set)
        self.message.config(state="disabled")

        # Finish
        self.next_button = ttk.Button(self, text ="Finish", command = lambda : self.quit()) 
        self.next_button.grid(row = 3, column = 0, padx = 10, pady = 10, sticky="se")
        self.next_button.state(["disabled"])

    def choose_write_dir(self, controller):
        controller.write_dir = filedialog.askdirectory()

        self.write_dir_entry.config(state="normal")
        self.write_dir_entry.delete(1.0, "end")
        self.write_dir_entry.insert("end", controller.write_dir)
        self.write_dir_entry.config(state="disabled")

        # write file path to config.ini so it's stashed for next time
        try:
            controller.config_file['FILE_PATHS']['write_dir'] = controller.write_dir
        except:
            controller.config_file['FILE_PATHS']['write_dir'] = ""

        with open('config.ini', 'w') as update:
            controller.config_file.write(update)

    def save_file(self, controller):

        if not self.write_name_entry.get() or len(self.write_dir_entry.get('1.0', 'end')) <= 1:
            messagebox.showerror("Missing Path", "No empty fields allowed.")
            return
            
        try:
            df = pd.read_excel(controller.temp_dir + '/temporary_file.xlsx')
            df.to_csv(controller.write_dir + "/" + self.write_name_entry.get() +  ".csv")
            self.message.config(state="normal")
            self.message.delete(1.0, "end")
            self.message.configure(fg="green")
            self.message.insert("end", "SUCCESS: File Saved at \n" + controller.write_dir + "/" + self.write_name_entry.get() + ".csv")
            self.message.config(state="disabled")

            self.save_button.state(["disabled"])
            self.next_button.state(["!disabled"])
            if os.path.exists(controller.temp_dir + '/temporary_file.xlsx'):
                os.remove(controller.temp_dir + '/temporary_file.xlsx')
            
        except Exception as err:
            self.message.config(state="normal")
            self.message.delete(1.0, "end")
            self.message.configure(fg="red")
            self.message.insert("end", "ERROR: Unable to save file. \n" + str(err))
            self.message.config(state="disabled")


class MenuBar(Menu):
    def __init__(self, controller): 
        Menu.__init__(self, controller)

        # File dropdown
        filemenu = Menu(self, tearoff=0)
        filemenu.add_command(label="New rules sheet from template", command= lambda: self.new_rules_sheet(controller))
        filemenu.add_separator()
        filemenu.add_command(label="Quit", command= lambda: self.quit())
        self.add_cascade(label="File", menu= filemenu)

        # Help dropdown
        helpmenu = Menu(self, tearoff=0)
        helpmenu.add_command(label="About", command=self.open_about)
        helpmenu.add_command(label="Tutorial", command= lambda: self.open_tutorial(controller))
        self.add_cascade(label="Help", menu=helpmenu)

    def new_rules_sheet(self, controller):
        if not os.path.exists("./temp/"):
            os.makedirs("./temp/")

        source="./templates/example.rules.xlsx"
        destination="./temp/temp.rules.xlsx"
        shutil.copyfile(source, destination)
        messagebox.showinfo("New Rules File", "This file opens from a temporary folder. Make sure to rename and save it in a new destination after making changes")
        controller.open_file(destination)

    def open_about(self):
        messagebox.showinfo("About", "Data tool designed by US Digital Response for manually-assisted and automated cleaning of spreadsheet data") 

    def open_tutorial(self, controller):
        tutorial = "./tutorial/tutorial.pdf"
        if not os.path.exists(tutorial):
            messagebox.showerror("Missing File", "Unable to locate tutorial")
        else:
            controller.open_file(tutorial)


# Driver Code 
app = tkinterApp()
app.title('Data Validation and Cleaning Tool') 
app.mainloop()

@atexit.register
def cleanup():
    fileList = glob.glob("./temp/*.xlsx")

    for filePath in fileList:
        try:
            os.remove(filePath)
        except:
            print("Error while deleting file: " + filePath)
