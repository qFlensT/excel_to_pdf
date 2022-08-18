#Simple Python Excel -> PDF Converter

>*It is supposed to be used via the context menu*

Takes as argument the path to the Excel file.

---

####Example:
`python main.py C:\Users\User\Documents\Some excel file.xlsx`
         *↓         file will be created*
`C:\Users\User\Documents\Some excel file.pdf`

---

_Please note that if the file name contains spaces, then `sys.argv` splits it into several arguments, which is why you have to combine all the arguments into a single string_

For example, you have the following file path: *C:\Users\User\Documents\Some excel file.xlsx*, `sys.argv` will receive the following set of arguments: `["Python/Executable/File.py", "C:\Users\User\Documents\Some", "excel", "file.xlsx"]`, script will automaticly concatinate args with slice `sys.argv[1::]` and we will get full path -  *C:\Users\User\Documents\Some excel file.xlsx*
