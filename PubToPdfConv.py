import os
for root, dirs, files in os.walk("C:/Users/L.Farnier/Desktop"):
    for file in files:
        if file.endswith(".pub"):
            print(os.path.join(root, file))