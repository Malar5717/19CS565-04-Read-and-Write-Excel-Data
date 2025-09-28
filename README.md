# EXP 04: Read and Write Excel Data


### AIM

The objective is to demonstrate the movement of structured data by reading all content (including headers) from a source Excel file and copying it precisely into a new destination Excel file.

---

### PROCEDURE

The workflow uses the `UiPath.Excel.Activities` package and the following sequence:

1.  **Read Operation:**
    * An **Excel Application Scope** connects to the **`InputData.xlsx`** file.
    * A **Read Range** activity is used inside the scope, leaving the Range field blank to capture all rows and columns.
    * The output is saved to a **DataTable** variable named `dtData`.
2.  **Write Operation:**
    * A **second Excel Application Scope** is used to create and connect to the **`OutputData.xlsx`** file.
    * A **Write Range** activity takes the `dtData` variable as input.
    * The **Add Headers** option is ticked to ensure the column names are included in the new file.

---

### OUTPUT

The automation creates a duplicate Excel file:

* **File Name:** `output.xlsx`
* **Contents:** Exact replication of all data and headers from `input.xlsx`.

files attached.
---

### RESULT

The process successfully transfers all structured data between the two files, validating the use of the DataTable structure as an in-memory data container within UiPath.
