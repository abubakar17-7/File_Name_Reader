import tkinter as tk, os, openpyxl, webbrowser, sys
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

############################################################################
##                                                                        ##
##                             RESOURCE PATH                              ##
##                                                                        ##
############################################################################

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

############################################################################
##                                                                        ##
##                             RESOURCE PATH                              ##
##                                                                        ##
############################################################################

def createFile():

    files = os.listdir(folder_path)

    new_List = ["Reference No"]

    for file in files:
        if Start == 0 and End == 0:
            name = file
        else:
            name = file[Start-1:End]
        if name not in new_List:
            new_List.append(name)

    workbook = openpyxl.Workbook()

    sheet = workbook.active

    for index, item in enumerate(new_List, start=1):
        sheet.cell(row=index, column=1, value=item)

    save_f = filedialog.asksaveasfilename(filetypes=[('Excel File', '*.xlsx')])

    save_file = save_f
    
    if ".xlsx" in save_f:
        save_file = save_f[:-5]

    if save_file:
        workbook.save(('%s.xlsx'%(save_file)))
        tk.messagebox.showinfo("Success", "File Created Successfully")
        Main()
    else:
        return

############################################################################
##                                                                        ##
##                             RESOURCE PATH                              ##
##                                                                        ##
############################################################################

def Saving():

    # This line uses list comprehension to destroy all the widgets inside the Seating_Frame
    # This ensures that any previous widgets are removed before new ones are added
    [widget.destroy() for widget in Body_Frame.winfo_children()]

    # This button creates a PDF of the seating plan when clicked
    ttk.Button(Body_Frame, text = "Download Excel File", width = 28, command = createFile).place(relx=0.5, rely=0.5, anchor="center")
    
############################################################################
##                                                                        ##
##                             RESOURCE PATH                              ##
##                                                                        ##
############################################################################

def chooseFile():
    
    # This line uses list comprehension to destroy all the widgets inside the Seating_Frame
    # This ensures that any previous widgets are removed before new ones are added
    [widget.destroy() for widget in Body_Frame.winfo_children()]
    
    # define a function to upload a file
    def upload():
        
        global folder_path
        folder_path = filedialog.askdirectory()
        
        if folder_path:
            file.config(text="Selected Folder: " + folder_path)
        
        return
    
    # define a function to go to the next step
    def next_func():
        
        try:
            if not folder_path:
                messagebox.showerror("Error", "No Folder Selected - Please Select a Folder")
                return
            
            try:
                global Start
                global End

                if start_e.get() == "":
                    Start = 0
                else:
                    Start = int(start_e.get().strip())
            
                if end_e.get() == "":
                    End = 0
                else:
                    End = int(end_e.get().strip())

                Saving()
            
            except:
                messagebox.showerror("Error", "Please Enter Valid Numbers")
                return

        except:
            messagebox.showerror("Error", "NNo Folder Selected - Please Select a Folder")
            return

    # create a label for the upload file section
    tk.Label(Body_Frame, text="OPEN FOLDER", font="Helvetica 14 bold").pack(pady=10)

    # create a label to display the filename
    file = tk.Label(Body_Frame, text="No Folder Selected", font="Helvetica 10")
    file.pack()

    # create a button to choose a file
    ttk.Button(Body_Frame, text="Choose Folder", command=upload).pack(pady=10)

    tk.Label(Body_Frame, text="Starting Character", font="Helvetica 10").pack()
    
    start_e = ttk.Entry(Body_Frame, font = 'Helvetica 12')
    start_e.pack(pady=10)

    tk.Label(Body_Frame, text="Ending Character", font="Helvetica 10").pack()

    end_e = ttk.Entry(Body_Frame, font = 'Helvetica 12')
    end_e.pack(pady=10)

    # create a button to go to the next step
    ttk.Button(Body_Frame, text="NEXT", command=next_func).pack(pady=10)

############################################################################
##                                                                        ##
##                             RESOURCE PATH                              ##
##                                                                        ##
############################################################################

def Main():
    
    # This line uses list comprehension to destroy all the widgets inside the Body_Frame
    # This ensures that any previous widgets are removed before new ones are added
    [widget.destroy() for widget in Body_Frame.winfo_children()]
    
    global folder_path
    folder_path = None

    # call the Choose_File function
    chooseFile()


##############################################################################
##                                                                          ##
##                                Frame                                     ##
##                                                                          ##
##############################################################################

w = tk.Tk()
w.title("FILE NAME REDEAR")
w.geometry("%dx%d" % (w.winfo_screenwidth(), w.winfo_screenheight()))
w.state("zoomed")
w.iconbitmap(resource_path('assets\\icon.ico'))

global Header_Frame
Header_Frame = tk.Frame(w)
Header_Frame.pack(side = "top")

global header_img
header_img = ImageTk.PhotoImage(Image.open(resource_path("assets\\header.png")))
tk.Label(Header_Frame, image = header_img, width=w.winfo_screenwidth()).pack()

global Body_Frame
Body_Frame = tk.Frame(w)
Body_Frame.pack(expand = 1, fill = "both")

def open_link(event=None):
    webbrowser.open_new(r"https://www.linkedin.com/in/muhammad-abu-bakar-a32b9528a")

link_label = ttk.Label(w, text = "Designed By M.Abu-Bakar", font = "Helvetica 8")
link_label.pack(padx = 10, pady = 10, side="right")
link_label.configure(foreground="#114333", cursor="hand2")
link_label.bind("<Button-1>", open_link)

Main()

w.mainloop()