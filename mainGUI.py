import tkinter as tk
from tkinter import ttk
import openpyxl

# Initialize lists to store debit and credit values
debit_vars = []
credit_vars = []

# Function to calculate balance and update table
def update_balance():
    try:
        # Convert to float for calculation, then convert result to string
        total_debit = sum([float(debit.get()) for debit in debit_vars if debit.get()])
        total_credit = sum([float(credit.get()) for credit in credit_vars if credit.get()])
        balance.set(str(total_debit - total_credit))  # Set balance as string
    except ValueError:
        pass  # Handle invalid input gracefully

# Function to transfer data to Excel
def transfer_to_excel():
    wb = openpyxl.Workbook()
    ws = wb.active

    # Define the headers
    headers = ['ID', 'Description', 'Unit', 'Rate', 'Quantity', 'Amount', 'Type', 'Debit', 'Credit', 'Balance']
    ws.append(headers)

    # Get data from the treeview and add it to the Excel sheet
    for row in tree.get_children():
        row_data = tree.item(row)['values']
        ws.append(list(row_data))  # Ensure the data passed is in list format

    # Save the workbook
    wb.save('bookkeeping_data.xlsx')

# Function to add a new row of data
def add_transaction():
    try:
        # Convert amount and quantity to float for calculation
        #amount = float(amount_var.get())

        quantity = float(quantity_var.get())

        amount = float(quantity*float(rate_var.get()))

        # Calculate debit or credit based on transaction type
        if type_var.get() == "Debit":
            debit_var.set(str(amount))  # Convert to string for Tkinter
            credit_var.set('0')
        else:
            debit_var.set('0')
            credit_var.set(str(amount))  # Convert to string for Tkinter

        # Add the new data to the table
        tree.insert('', 'end', values=(
            id_var.get() or '',  # Ensure that empty string is passed if field is empty
            desc_var.get() or '',
            unit_var.get() or '',
            rate_var.get() or '',
            str(quantity),  # Convert quantity to string for display
            str(amount) ,  # Convert amount to string for display
            type_var.get() or '',
            debit_var.get() or '0',  # Provide default value for empty fields
            credit_var.get() or '0',
            balance.get() or '0'
        ))

        # Append the current debit and credit vars to the lists
        debit_vars.append(debit_var)
        credit_vars.append(credit_var)

        # Update the balance after each entry
        update_balance()

        # Clear the input fields after adding the transaction
        id_var.set('')
        desc_var.set('')
        unit_var.set('')
        rate_var.set('0.0')
        quantity_var.set('1')
        amount_var.set('0.0')
        type_var.set('')

    except ValueError:
        pass  # Handle invalid input gracefully

# GUI setup
root = tk.Tk()
root.title("Bookkeeping App")

# Variables for input fields
id_var = tk.StringVar()
desc_var = tk.StringVar()
unit_var = tk.StringVar()
rate_var = tk.StringVar(value="0.0")
quantity_var = tk.StringVar(value="1")  # Quantity variable
amount_var = tk.StringVar()
type_var = tk.StringVar(value="Debit")
debit_var = tk.StringVar()
credit_var = tk.StringVar()
balance = tk.StringVar()

# Layout design
tk.Label(root, text="ID").grid(row=0, column=0)
tk.Entry(root, textvariable=id_var).grid(row=0, column=1)

tk.Label(root, text="Description").grid(row=1, column=0)
tk.Entry(root, textvariable=desc_var).grid(row=1, column=1)

tk.Label(root, text="Unit").grid(row=2, column=0)
ttk.Combobox(root, textvariable=unit_var, values=['LS', 'No', 'Pcs', 'M', 'M2', 'M3']).grid(row=2, column=1, pady=5)

tk.Label(root, text="Rate").grid(row=4, column=0)
tk.Entry(root, textvariable=rate_var).grid(row=4, column=1)

tk.Label(root, text="Quantity").grid(row=3, column=0)  # New Quantity field
tk.Entry(root, textvariable=quantity_var).grid(row=3, column=1)

tk.Label(root, text="Amount").grid(row=5, column=0)
tk.Entry(root, textvariable=amount_var).grid(row=5, column=1)

tk.Label(root, text="Transaction Type").grid(row=0, column=2)
ttk.Combobox(root, textvariable=type_var, values=["Debit", "Credit"]).grid(row=0, column=3)

tk.Label(root, text="Debit").grid(row=3, column=2)
tk.Entry(root, textvariable=debit_var, state="readonly").grid(row=3, column=3)

tk.Label(root, text="Credit").grid(row=4, column=2)
tk.Entry(root, textvariable=credit_var, state="readonly").grid(row=4, column=3)

tk.Label(root, text="Balance").grid(row=5, column=2)
tk.Entry(root, textvariable=balance, state="readonly").grid(row=5, column=3)

# Button to add transaction
tk.Button(root, text="Add Transaction", command=add_transaction).grid(row=6, column=0, columnspan=4)

# Table for displaying transactions
columns = ['ID', 'Description', 'Unit', 'Rate', 'Quantity', 'Amount', 'Type', 'Debit', 'Credit', 'Balance']
tree = ttk.Treeview(root, columns=columns, show='headings')

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=120)

tree.grid(row=8, column=0, columnspan=4, padx=10, pady=10)

# Button to transfer data to Excel
tk.Button(root, text="Transfer to Excel", command=transfer_to_excel).grid(row=12, column=0, columnspan=2, padx=20, pady=20)

root.mainloop()
