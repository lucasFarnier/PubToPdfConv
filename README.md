description
-
ive created an application to convert publisher files to pdf as certain windows 11 devices and office 2024 packages are losing support for publisher and need files converted.

this is to be run before upgraded as publisher is requiered to be on the computer (or have access to your files and a different computer with publisher)


instructions
-
the user must select a folder/directory containing the .pubs to convert, they will not be visable when picking the folder
once selected a list of the publisher files found in that directory and sub directory will be displayed to confirm they are correct

"Convert to PDF" button is also disabled until directory picked, and if a directory doesnt have any .pub it will output a popup error message and keep the "Convert to PDF" button disabled

the program can be ended at any time by clicking the "cancle/end program" button, the window x button has be changed to minimise instead

while running the "browse" and "Convert to PDF" are disabled during the conversion

once done the user can select the convert button and a ✅ for successful conversions and an ❌ for insuccessful

when a file failes to convert it will retry 2 more times (total tries 3 times), after if it failes 3 it skips the file
also instead of a popup message saying "all files converted successfully" it will instead say "most files converted successfully [a list of file directories for the failed ones]"

colour code meanings
-
there are 3 colours a file becomes if converted/not converted
    
-green, successful first try
    
-orange, failed 1 or 2 times but successful
    
-red, unsuccessful and the file could not convert


photos
-

![image](https://github.com/user-attachments/assets/90a06f8a-21c8-4fd3-99f7-aad4d8108b88)
![image](https://github.com/user-attachments/assets/86861c76-f710-475b-b03d-25d36e395301)
![image](https://github.com/user-attachments/assets/3bdb4d9d-b7d3-4a0a-8c7f-d46ca84ef3a8)
