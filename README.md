# Formal Letter Generator

#### Video Demo:  <URL HERE>
#### Description:
This Python script generates Docx templates for formal letters using a command-line interface. It provides options to write letters to companies, as well as the ability to save company names and their respective addresses in a database using Python's shelve module. The program can also list company names found in the database in alphabetical order.


When you use the Python's shelve module on Windows, you will notice three files with different extensions in your directory. The following files correspond to different components of the shelve mechanism:

1. `data_base.bak:`
This file is a backup of the shelf database. It provides a safety net in case the main database file (data_base.dat) gets corrupted or damaged during write operations. If an error occurs while writing to the shelf, Python can attempt to restore it using this backup.
2. `data_base.dat:`
This file contains the actual serialized objects. It's the main data file of the shelf. All the objects that you've stored in the shelf are stored here in a serialized format.
3. `data_base.dir:`
This file is used to index the keys and their corresponding positions in the data_base.dat file. It allows for efficient lookup and retrieval of objects based on their keys.

---
## Usage
### To generate template for letter:

- `python project.py letter to <company name>`
  - This generates a Docx template with the address of the specified company.
  - If the specified company is not in the database, the user will be required to input the address of that company to be stored in the database.

### To save an address:

- `python project.py save`
  - This stores the address of a specific company in the database.

### To list all addresses:

- `python project.py list`
  - This lists all company names in the database in alphabetical order.

---
## Instructions

1. Make sure Python is installed on your system.

2. Clone this repository.

3. Open the terminal and navigate to the project directory:
   ```bash
   cd /path/to/CS50-Final-Project
   ```

4. Follow the usage examples provided above to generate templates for letters.

5. You can customize the templates by modifying the functions in the project.py script.

---

## Dependencies
This project relies on the following Python libraries:

- [python-docx](https://python-docx.readthedocs.io/en/latest/)
- [pyperclip](https://pypi.org/project/pyperclip/)

Use `pip` to install the libraries:

```bash
pip install pyperclip python-docx
pip install pyperclip python-docx
```

---
## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---
## About CS50P
This course is an introduction to programming using Python, covering essential skills in coding, testing, and debugging. It's designed for students, whether with or without prior programming experience, who specifically want to learn Python.
[Learn more about CS50's Introduction to Programming with Python.](https://www.edx.org/learn/python/harvard-university-cs50-s-introduction-to-programming-with-python)

---
## Author
[Github](https://github.com/jim-franklin)
<br>[LinkedIn](https://www.linkedin.com/in/franklin-aryee-118729183/)