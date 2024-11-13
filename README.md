### Sowmya V
### 212222110045

# RPA Capstone Project - Stock level management

### Aim: 
To automate the process of inventory management

### Software Requirements:
UiPath studio communtiy edition


### Project Overview

The project automates inventory management by:
1. Updating stock levels based on user inputs.
2. Checking inventory for items below a minimum threshold.
3. Sending reorder notifications for low-stock items.

---

### Procedure


### Step 1: Set Up the Stock Update Workflow

1. **Open UiPath Studio** and create a new project.
2. **Add an Excel Application Scope** to access the inventory Excel file.
   - Set the **WorkbookPath** to the path of `Unique_Inventory_Data.xlsx`.

3. **Read Inventory Data**:
   - Inside the **Excel Application Scope**, add a **Read Range** activity to load the data into a DataTable (e.g., `InventoryData`).
   - Set the **Output** to `InventoryData`.

4. **Loop Through Inventory Data**:
   - Add a **For Each Row** activity to iterate over `InventoryData`.
   - Inside the loop, add an **If** activity to prompt for stock changes only for a specific item (optional). If you want to prompt for every item, skip this step.

5. **Input Dialog to Update Stock**:
   - Add an **Input Dialog** to ask the user to enter the stock change for each item.
   - Set the **Label** to prompt with the item’s name dynamically (e.g., `"Enter stock change for " + CurrentRow("Item Name").ToString`).
   - Store the result in a variable (e.g., `StockChange`).

6. **Update Current Stock Level**:
   - Add an **Assign** activity to update `CurrentRow("Current Stock Level")` with the new value:
     ```plaintext
     CInt(CurrentRow("Current Stock Level")) + CInt(StockChange)
     ```

7. **Write Updated Data Back to Excel**:
   - After the loop, use a **Write Range** activity to save the updated `InventoryData` back to the Excel file.

---

### Step 2: Implement the Reorder Notification Workflow

1. **Read Updated Data**:
   - After updating the stock, add another **Read Range** activity to reload the data from the Excel file into `InventoryData`.

2. **Filter Low Stock Items**:
   - Add a **Filter Data Table** activity to create a new DataTable (e.g., `LowStockItems`).
   - Set the **Filter Rows** to keep rows where `Current Stock Level` is less than `Minimum Stock Level`.

3. **Check for Low Stock Items**:
   - Add an **If** activity to check if `LowStockItems.Rows.Count > 0`.
   - This verifies if there are any items below the minimum stock level.

4. **Send Notification Email**:
   - Inside the **Then** branch of the **If** activity, add a **Send SMTP Mail Message** activity to send a low-stock notification email.
   - Configure the email properties:
     - **To**: Recipient's email address.
     - **Subject**: `"Low Stock Alert - Items Below Minimum Stock"`.
     - **Body**: Use `String.Join` to list items from `LowStockItems`:
       ```plaintext
       "The following items are below the minimum stock level:\n" + String.Join("\n", LowStockItems.AsEnumerable().Select(Function(row) row("Item Name").ToString + " - Current Stock: " + row("Current Stock Level").ToString))
       ```

5. **Attach Low Stock Report (Optional)**:
   - If needed, add a **Write Range** activity to save `LowStockItems` as an Excel report (e.g., "Low_Stock_Report.xlsx").
   - Attach this file in the **Send SMTP Mail Message** activity.

---

### Step 3: Implement the Reporting Workflow (Optional)

1. **Generate Inventory Summary Report**:
   - After processing stock updates and notifications, you can generate a summary report with totals, items below reorder level, etc.
   - Use **Write Range** to save this report as "Inventory_Report.xlsx".

2. **Email the Report**:
   - Use **Send SMTP Mail Message** to email the summary report to relevant stakeholders.

---

### Workflow - Main.xaml


![image](https://github.com/user-attachments/assets/6a117ea2-622f-485c-8423-01f428c411bf)
![image](https://github.com/user-attachments/assets/2d78d916-89ee-4dbd-8a63-5aea78d06936)
![image](https://github.com/user-attachments/assets/c47da210-fd3f-4567-842d-2118eb27e0a2)
![image](https://github.com/user-attachments/assets/67725723-ba0e-450c-9701-859fba5b8e8f)


### Results 

1. **Stock Update Process**:
   - The **Input Dialog** will prompt you to enter stock changes (e.g., additions or subtractions in quantity) for each item (or specific items if configured).
   - The **Current Stock Level** values in the Excel file will be updated based on the inputs provided.

2. **Updated Excel File**:
   - The inventory Excel file (e.g., `Unique_Inventory_Data.xlsx`) will reflect the new **Current Stock Level** for each item after the stock update process.
   - You can open the file to verify that the stock levels have been updated correctly.

3. **Reorder Notification**:
   - The workflow will check for items where **Current Stock Level** is below **Minimum Stock Level**.
   - If any items are identified as low stock, you should receive an email notification listing these items and their current stock levels.
   - The email will have a subject like "Low Stock Alert - Items Below Minimum Stock" and a body listing the low-stock items.

