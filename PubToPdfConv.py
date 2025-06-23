import os
import comtypes.client

publisher = comtypes.client.CreateObject("Publisher.Application")
for root, dirs, files in os.walk(r"C:\Test"):
    for file in files:
        if file.endswith(".pub"):
            filePathInp = os.path.join(root, file)
            filePathOut = os.path.join(root, f"{os.path.splitext(file)[0]}.pdf")
            print(filePathInp)
            print(filePathOut)

            try:
                publisher.Visible = True

                doc = publisher.Open(filePathInp)
                doc.ExportAsFixedFormat(Filename=filePathOut, Format=1)
                doc.Close()

                print(f"Conversion successful: {filePathOut}")

            except Exception as e:
                print(f"An error occurred: {e}")
publisher.Quit()