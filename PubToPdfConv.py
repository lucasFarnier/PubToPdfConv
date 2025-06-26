import os

import threading

import win32com.client
import win32gui
import win32con

import gc

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# defining global variabls
folderSelected = ""
pubFiles = []


# function for checking folder being picked
def browseFolders():
    # globalises to use across multiple functions
    global folderSelected
    global pubFiles

    # resets the list before refilling it
    pubFiles = []
    # remove all items in the gui
    tree.delete(*tree.get_children())
    # takes the directory selected and stores it
    folderSelected = filedialog.askdirectory()

    # checking all files and directories steming from one provided
    for root, dirs, files in os.walk(folderSelected):
        for file in files:
            # if ones found adds it to the list and the gui
            if file.lower().endswith(".pub"):  # or file.lower().endswith(".pdf"):
                pubFiles.append(os.path.join(root, file))
                tree.insert("", "end", values=(os.path.join(root, file), "", "", ""))

    # if the list isnt empty then activate the convert button
    if len(pubFiles) != 0:
        convertButtonPb.configure(state="normal")
        # convertButtonWd.configure(state="normal")
    # if its empty output error message box telling user, also keeps convert button inactive
    else:
        messagebox.showerror("No files found", "Folder doesnt contain any .pub (publisher) files" + folderSelected)


# function to define perameters for pub to pdf
def toPdf():
    thread = threading.Thread(target=PubToPdfAndPdfToDocx, args=("pub", "pdf"))
    thread.start()


# function to define perameters for pdf to docx
def toDocx():
    thread = threading.Thread(target=PubToPdfAndPdfToDocx("pdf", "docx"))
    thread.start()


# updates the attempt col of the files listed
def UpdateCol(index, attempt):
    id = tree.get_children()[index]

    # depending on number of attempts changes it to match
    if (attempt == 1):
        tree.item(id, values=(pubFiles[index], "❌", "✅", ""))
    elif (attempt == 2):
        tree.item(id, values=(pubFiles[index], "❌", "❌", "✅"))
    elif (attempt == 3):
        tree.item(id, values=(pubFiles[index], "❌", "❌", "❌"))
    else:
        tree.item(id, values=(pubFiles[index], "✅", "", ""))
    root.update_idletasks()


# converter, takes pubs and makes them pdfs or takes pdfs and makes them docx
def PubToPdfAndPdfToDocx(convertFrom, convertTo):
    # once called on re-disables the button so it cant be re-clicked until valid directory and files
    # convertButtonWd.configure(state="disabled")
    convertButtonPb.configure(state="disabled")
    browseButton.configure(state="disabled")

    # define the list and strings for the failed/error files
    ErrConvertedFiles = []
    ErrConvertedFilesList = ""
    DidError = "All "
    # index for tracking on updateCol function
    index = -1

    # goes through all files and directories to help convert them
    for root, dirs, files in os.walk(folderSelected):
        for file in files:
            if file.lower().endswith(".pub"):  # or file.lower().endswith(".pdf"):
                # moves index to next (for function mentioned before)
                index += 1
                if not (file.lower().endswith("." + convertFrom)):
                    tree.item((tree.get_children()[index]), tags=("NA"))

            if file.lower().endswith("." + convertFrom):
                # resets attempts for each file
                attempt = 0

                # depending if it is to pdf or to docx opens the correct win32com client
                if convertFrom == "pub":
                    publisherWord = win32com.client.Dispatch("Publisher.Application")
                elif convertFrom == "pdf":
                    publisherWord = win32com.client.Dispatch("Word.Application")

                # gets file path for the input filepath then new file path for output pdf
                filePathInp = os.path.normpath(os.path.join(root, file))
                filePathOut = os.path.normpath(os.path.join(root, f"{os.path.splitext(file)[0]}." + convertTo))

                # while it hasnt failed 4 times
                while attempt < 3:
                    try:
                        # configers the export formats for saving as pdf
                        if convertTo == "pdf":
                            doc = publisherWord.Open(filePathInp)

                            doc.ExportAsFixedFormat(
                                Filename=filePathOut,
                                Format=2,
                                Intent=1,
                                IncludeDocumentProperties=True,
                                BitmapMissingFonts=True
                            )
                        # otherwise export formated for saving as docx
                        elif convertTo == "docx":
                            doc = publisherWord.Documents.Open(filePathInp)
                            doc.SaveAs2(
                                filePathOut,
                                FileFormat=16
                            )

                        # closes the file and clean up
                        doc.Close()
                        gc.collect()

                        # changes row colour and updates on gui with before mentioned function
                        if (attempt == 0):
                            tree.item((tree.get_children()[index]), tags=("success"))
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
                            UpdateCol(index, attempt)
                            break
                        else:
                            tree.item((tree.get_children()[index]), tags=("partFailed"))
                            UpdateCol(index, attempt)
                        # if not failed 3 times yet then retries until either works or fails and skips

    # if any fully failed (3 tries) then sets up output end message to include them
    if (ErrConvertedFiles):
        ErrConvertedFilesList = "List of failed file filepaths:"
        DidError = "Some "

        # makes a string of all the failed files
        for item in ErrConvertedFiles:
            ErrConvertedFilesList = ErrConvertedFilesList + "\n" + item

    browseButton.configure(state="normal")
    # changes message from "all" to "Some" and lists the failed files at the end
    messagebox.showinfo("Files converted", DidError + "files converted successfully\n\n" + ErrConvertedFilesList)


# used to make window "x" a minimise instead of close it
def minimizeInsteadOfClose():
    root.iconify()


# end proram
def Exit():
    root.destroy()
    exit()


# gui
try:
    root = tk.Tk()
    root.title("Pub to PDF")

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack()

    label = tk.Label(frame, text="Step 1: Select folder with .pub files in and pdf copies will be made:")
    label.pack()

    browseButton = tk.Button(frame, text="Browse", command=browseFolders)
    browseButton.pack(pady=(5, 0))

    columns = ("File", "Att 1", "Att 2", "Att 3")
    tree = ttk.Treeview(frame, columns=columns, show="headings")
    for column in columns:
        tree.heading(column, text=column)
        tree.column(column, width=450 if column == "File" else 50, anchor='w')
    tree.pack(fill="both", expand=True, pady=(10, 10))

    label = tk.Label(frame, text="\nStep 2: Select files to convert to:")
    label.pack()

    convertButtonPb = tk.Button(frame, text="Convert to PDF", command=toPdf, state="disabled")
    convertButtonPb.pack(pady=(5, 0))
    # (for publisher to word make sure to convert to pdf first)

    # option can be unhashed if needed, warning the pdf to word format doesnt work well, everything copies over but will mess the page formats
    # also remove hash for line 73, 39 and part of 33 for it to work
    '''label = tk.Label(frame, text="\nFor publisher to word make sure to convert to pdf first\nOptional step 3: Select files to convert to:")
    label.pack()

    convertButtonWd = tk.Button(frame, text="Convert to word\n.pdf -> .docx", command=toDocx, state="disabled")
    convertButtonWd.pack()'''

    ExitButton = tk.Button(frame, text="cancle/end program", command=Exit, foreground="white", background="#d9534f")
    ExitButton.pack(side="right")

    # colour tags for fails and sucesses
    tree.tag_configure("success", background="#ADEA33")
    tree.tag_configure("partFailed", background="#EAB333")
    tree.tag_configure("failed", background="#E14B2A")
    tree.tag_configure("NA", foreground="#B1B1B1")

    root.protocol("WM_DELETE_WINDOW", minimizeInsteadOfClose)
    root.mainloop()
except Exception:
    print("ended")
