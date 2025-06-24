import os
import win32com.client
import gc
import pythoncom
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
os.environ['TCL_LIBRARY'] = r'C:\Users\L.Farnier\AppData\Local\Programs\Python\Python313\tcl\tcl8.6'


#defining global variabls
folderSelected = ""
pubFiles = []


#function for checking folder being picked
def browseFolders():
    #globalises to use across multiple
    global folderSelected
    global pubFiles

    #resets the list before refilling it
    pubFiles = []
    #remove all items in the gui
    tree.delete(*tree.get_children())
    #takes the directory selected and stores it
    folderSelected = filedialog.askdirectory()

    #checking all files and directories steming from one provided
    for root, dirs, files in os.walk(folderSelected):
        for file in files:
            #if ones found adds it to the list and the gui
            if file.lower().endswith(".pub") or file.lower().endswith(".pdf"):
                pubFiles.append(os.path.join(root, file))
                tree.insert("", "end", values=(os.path.join(root, file), "", "", ""))

    #if the list isnt empty then activate the convert button
    if len(pubFiles) != 0:
        convertButtonPb.configure(state="normal")
        convertButtonWd.configure(state="normal")
    #if its empty output error message box telling user, also keeps convert button inactive
    else:
        messagebox.showerror("No files found", "Folder doesnt contain any .pub (publisher) files" + folderSelected)


#updates the attempt col of the files listed
def UpdateCol(index, attempt):
    id = tree.get_children()[index]

    #depending on number of attempts changes it to match
    if (attempt == 1):
        tree.item(id, values=(pubFiles[index], "❌", "✅", ""))
    elif (attempt == 2):
        tree.item(id, values=(pubFiles[index], "❌", "❌", "✅"))
    elif (attempt == 3):
        tree.item(id, values=(pubFiles[index], "❌", "❌", "❌"))
    else:
        tree.item(id, values=(pubFiles[index], "✅", "", ""))


#converter, takes pubs and makes them pdfs
def PubToPdf():
    #once called on re-disables the button so it cant be re-clicked until valid directory and files
    convertButtonPb.configure(state="disabled")

    #define the list and strings for the failed/error files
    ErrConvertedFiles = []
    ErrConvertedFilesList = ""
    DidError = "All "
    #index for tracking on updateCol function
    index = -1

    #goes through all files and directories to help convert them
    for root, dirs, files in os.walk(folderSelected):
        for file in files:
            if file.lower().endswith(".pub"):
                #moves index to next (for function mentioned before
                index += 1
                #resets attempts for each file
                attempt = 0

                publisher = win32com.client.Dispatch("Publisher.Application")

                #gets file path for the input filepath then new file path for output pdf
                filePathInp = os.path.normpath(os.path.join(root, file))
                filePathOut = os.path.normpath(os.path.join(root, f"{os.path.splitext(file)[0]}.pdf"))

                #while it hasnt failed 4 times
                while attempt < 3:
                    try:
                        doc = publisher.Open(filePathInp)
                        #configers the export formats for saving as pdf
                        doc.ExportAsFixedFormat(
                            Filename = filePathOut,
                            Format = 2,
                            Intent = 1,
                            IncludeDocumentProperties = True,
                            BitmapMissingFonts = True
                        )
                        #closes the file and clean up
                        doc.Close()
                        gc.collect()

                        #changes row colour and updates on gui with before mentioned function
                        if (attempt == 0):
                            tree.item((tree.get_children()[index]), tags=("success"))
                        elif (attempt > 0):
                            tree.item((tree.get_children()[index]), tags=("PartFailed"))
                        UpdateCol(index, attempt)
                        break

                    #if error converting
                    except Exception as e:
                        #adds 1 to attempts made
                        attempt += 1

                        #if 3 failes skips file once retrying 3 times and highlights red (rather than green on sucess)
                        if attempt == 3:
                            ErrConvertedFiles.append(filePathInp)
                            tree.item((tree.get_children()[index]), tags=("failed"))
                            break
                        #if not failed 3 times yet then retries until either works or failes and skips

    #if any fully failed (3 tries) then sets up output end message to include them
    if (ErrConvertedFiles):
        ErrConvertedFilesList = "List of failed file filepaths:"
        DidError = "Some "

        #makes a string of all the failed files
        for item in ErrConvertedFiles:
            ErrConvertedFilesList = ErrConvertedFilesList + "\n" + item

    #changes message from "all" to "Some" and lists the failed files at the end
    messagebox.showinfo("Files converted", DidError + "files converted successfully\n\n" + ErrConvertedFilesList)

def PdfToDocx():
    #once called on re-disables the button so it cant be re-clicked until valid directory and files
    convertButtonWd.configure(state="disabled")

    #define the list and strings for the failed/error files
    ErrConvertedFiles = []
    ErrConvertedFilesList = ""
    DidError = "All "
    #index for tracking on updateCol function
    index = -1

    #goes through all files and directories to help convert them
    for root, dirs, files in os.walk(folderSelected):
        for file in files:
            if file.lower().endswith(".pdf"):
                #moves index to next (for function mentioned before
                index += 1
                #resets attempts for each file
                attempt = 0

                word = win32com.client.Dispatch("Word.Application")


                filePathInp = os.path.normpath(os.path.join(root, file))
                filePathOut = os.path.normpath(os.path.join(root, f"{os.path.splitext(file)[0]}.docx"))

                # while it hasnt failed 4 times
                while attempt < 3:
                    print("h1")
                    try:
                        print("1")
                        doc = word.Documents.Open(filePathInp)
                        # configers the export formats for saving as pdf
                        print("2")
                        doc.SaveAs2(
                            filePathOut,
                            FileFormat=16
                        )
                        print("3")
                        # closes the file and clean up
                        doc.Close()
                        gc.collect()

                        # changes row colour and updates on gui with before mentioned function
                        if (attempt == 0):
                            tree.item((tree.get_children()[index]), tags=("success"))
                        elif (attempt > 0):
                            tree.item((tree.get_children()[index]), tags=("PartFailed"))
                        UpdateCol(index, attempt)
                        break

                    # if error converting
                    except Exception as e:
                        # adds 1 to attempts made
                        attempt += 1

                        # if 3 failes skips file once retrying 3 times and highlights red (rather than green on sucess)
                        if attempt == 3:
                            ErrConvertedFiles.append(filePathInp)
                            tree.item((tree.get_children()[index]), tags=("failed"))
                            break
                        # if not failed 3 times yet then retries until either works or failes and skips

    #if any fully failed (3 tries) then sets up output end message to include them
    if (ErrConvertedFiles):
        ErrConvertedFilesList = "List of failed file filepaths:"
        DidError = "Some "

        #makes a string of all the failed files
        for item in ErrConvertedFiles:
            ErrConvertedFilesList = ErrConvertedFilesList + "\n" + item

    #changes message from "all" to "Some" and lists the failed files at the end
    messagebox.showinfo("Files converted", DidError + "files converted successfully\n\n" + ErrConvertedFilesList)

#gui
try:
    root = tk.Tk()
    root.title("Pub to PDF")

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack()


    label = tk.Label(frame, text="Step 1: Select folder with .pub files in and pdf copies will be made:")
    label.pack()

    browseButton = tk.Button(frame, text="Browse", command=browseFolders)
    browseButton.pack()

    columns = ("File", "Att 1", "Att 2", "Att 3")
    tree = ttk.Treeview(frame, columns=columns, show="headings")
    for column in columns:
        tree.heading(column, text=column)
        tree.column(column, width=450 if column == "File" else 50, anchor='w')
    tree.pack()

    label = tk.Label(frame, text="Step 2: Select files to convert too :")
    label.pack()

    convertButtonPb = tk.Button(frame, text="Convert to PDF\n.pub -> .pdf", command=PubToPdf, state="disabled")
    convertButtonPb.pack()
    #(for publisher to word make sure to convert to pdf first)

    label = tk.Label(frame, text="For publisher to word make sure to convert to pdf first")
    label.pack()

    convertButtonWd = tk.Button(frame, text="Convert to word\n.pdf -> .docx", command=PdfToDocx, state="disabled")
    convertButtonWd.pack()

    #colour tags for fails and sucesses
    tree.tag_configure("success", background="lightgreen")
    tree.tag_configure("partFailed", background="lightcoral")
    tree.tag_configure("failed", background="#ff5733")

    root.mainloop()
except Exception:
    print("ended")