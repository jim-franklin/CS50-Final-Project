# Formal Letter Generator

#### Video Demo:  <URL HERE>
#### Description:
This Python script creates MS Word templates for formal letters using a command-line interface.<br>It provides options to create letters for companies and the Committee, as well as the ability to save and list addresses.

---

## Usage

### To create a letter:

- `python project.py letter to <company name>`
  - This creates a template for a Word document with the address of the specified company.
  - If the specified company is not in the database (a shelf file in this case), the user will be required to input the address of that company to be stored in the database.


- `python project.py letter to committee`
  - This creates a template for a Word document with the general address of the Committee.

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
   cd /path/to/project_folder
   ```

4. Follow the usage examples provided above to create letters.

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

## Author

[Franklin Aryee](https://github.com/jim-franklin)