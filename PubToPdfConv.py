import os
import comtypes.client
import win32com.client

publisher = win32com.client.Dispatch("Publisher.Application")
#publisher.Visible = True

for root, dirs, files in os.walk(r"C:\Test"):
    for file in files:
        if file.lower().endswith(".pub"):
            #gets file path for the input filepath then new file path for output pdf
            filePathInp = os.path.join(root, file)
            filePathOut = os.path.join(root, f"{os.path.splitext(file)[0]}.PDF")

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
                doc.Close()

                print(f"Conversion successful: {filePathOut}")

            except Exception as e:
                print(f"An error occurred: {e}")
publisher.Quit()