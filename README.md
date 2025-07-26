# Dinner-planner-
# 🍽️ Dinner Planner (Excel Macro-Enabled Project)

This project is a user-friendly **Dinner Planner** built using **Microsoft Excel (.xlsm)** and **VBA**. It allows users to plan weekly meals, store guest information, and automate interactions using a custom UserForm and buttons.

---

## 📌 Features

- 📅 Weekly dinner planning interface
- 🔘 Custom UserForm with buttons for managing records
- 🧠 VBA Macros to automate repetitive tasks
- 📋 Editable guest and meal details
- 💾 All data stored within the workbook — no external database needed

---

## 🧾 Form Fields

The UserForm includes the following input elements:

### Text Boxes

- **Name** – Guest's name
- **Phone Number** – Contact number
- **City** – Guest’s city of residence

### Combo Box

- **Dinner Preference** – Dropdown to select meal type  
  *(e.g., Vegetarian, Vegan, Non-Vegetarian, etc.)*

### Date Picker (With Specific Dates)

- **Date** – Date field limited to specific available dinner dates  
  *(Prevents selection of invalid or unavailable dates)*

---

## 🖱️ Button Functions

Each button on the UserForm performs a key task:

- **ADD** ➕  
  Adds a new dinner record to the worksheet

- **UPDATE** ✏️  
  Search for an existing record, then update it

- **SEARCH** 🔍  
  Finds and displays matching guest details

- **CLEAR** 🧹  
  Clears all form fields for new input

- **DELETE** 🗑️  
  Deletes the selected record after search

- **EXIT** ❌  
  Closes the UserForm

- **INPUT THE SCHOOL** 🏫  
  Custom button to input or tag school name (optional)

---

## 🧠 Technologies Used

- Microsoft Excel (.xlsm)
- VBA (Visual Basic for Applications)
- UserForms and Modules

---

## 🗂️ Project Structure

Dinner-planner-/
│
├── dinner planner file.xlsm # Main Excel workbook
├── VBA code/ # Folder containing exported VBA components
│ ├── DinnerPlannerUserForm.frm.frm # UserForm interface file
│ └── DinnerPlannerCustomButton.bas.bas # VBA module for custom button logic
└── README.md # This file

> You can import the `.frm` and `.bas` files into the VBA Editor using `File > Import File...` (press `Alt + F11` in Excel).

---

##  How to Use

1. Download the entire repository or clone it using Git.
2. Open `dinner planner file.xlsm` in Microsoft Excel.
3. When prompted, click **Enable Content** or **Enable Macros**.
4. Launch the UserForm and begin planning your dinners!
5. Customize or review the code using the included `.bas` and `.frm` files.

---

## ⚠️ Requirements

- Microsoft Excel 2016 or newer (Windows)
- Macros must be enabled for full functionality

---

## 🙋‍♀️ Author

Created by **[Aakriti khanduri]**  
GitHub: [https://github.com/Aakriti-khanduri]
