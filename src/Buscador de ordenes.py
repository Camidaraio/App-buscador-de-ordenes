import pyodbc
import tkinter as tk
from tkinter import ttk, messagebox

def buscar_orden():
    try:
        connection = pyodbc.connect('DRIVER={SQL Server};SERVER=CAMILA;DATABASE=Prueba2;Trusted_Connection=yes;')
        cursor = connection.cursor()
        
        order_id = order_entry.get()
        if not order_id.isdigit():
            messagebox.showerror("Error", "Por favor, ingrese un número de Orden válido")
            return

        cursor.execute('''
                        SELECT 
                            Customers.CompanyName,
                            Customers.Address,
                            Products.ProductName,
                            OrderDetails.UnitPrice,
                            OrderDetails.Quantity,
                            (SELECT COUNT(DISTINCT OD2.OrderID)
                            FROM OrderDetails OD2
                            WHERE OD2.ProductID = OrderDetails.ProductID AND OD2.OrderID != Orders.OrderID) AS Otras_Ordenes_Con_Producto
                        FROM 
                            Customers
                        INNER JOIN 
                            Orders ON Customers.CustomerID = Orders.CustomerID
                        INNER JOIN 
                            OrderDetails ON Orders.OrderID = OrderDetails.OrderID
                        INNER JOIN 
                            Products ON OrderDetails.ProductID = Products.ProductID
                        WHERE 
                            Orders.OrderID = ?
                    ''', order_id)
        
        rows = cursor.fetchall()
        
        for row in tree.get_children():
            tree.delete(row)
        
        if rows:
            for row in rows:
                tree.insert("", "end", values=(row.CompanyName, row.Address, row.ProductName, row.UnitPrice, row.Quantity, row.Otras_Ordenes_Con_Producto))
        else:
            messagebox.showinfo("Información", f"No se encontraron resultados para el número de Orden: {order_id}")

    except Exception as ex:
        messagebox.showerror("Error", f"Tipo de error: {ex}")



root = tk.Tk()
root.config(bg="pink", cursor="arrow")
root.title("Order Finder")

tk.Label(root, text="Insert Order ID:").grid(row=0, column=0, padx=10, pady=10)
order_entry = tk.Entry(root)
order_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Search", command=buscar_orden).grid(row=0, column=2, padx=10, pady=10)

columns = ["Company Name", "Address", "Product Name", "Unit Price", "Quantity", "Same Products on Other Orders"]
tree = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col)
tree.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

root.mainloop()
