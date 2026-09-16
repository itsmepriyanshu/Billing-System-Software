import importlib
import os
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog, ttk


def ensure_optional_dependency(module_name: str, package_name: str | None = None) -> bool:
    package_name = package_name or module_name
    try:
        importlib.import_module(module_name)
        return True
    except ImportError:
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", package_name],
                check=True,
                capture_output=True,
                text=True,
            )
            importlib.import_module(module_name)
            return True
        except Exception:
            return False


try:
    if ensure_optional_dependency("reportlab.lib.pagesizes", "reportlab"):
        from reportlab.lib.pagesizes import LETTER
        from reportlab.pdfgen import canvas
    else:
        LETTER = None
        canvas = None
except Exception:  # pragma: no cover
    LETTER = None
    canvas = None

from database import (
    Workbook,
    export_main_excel,
    get_main_excel_path,
    add_product,
    authenticate_user,
    calculate_grand_total,
    calculate_item_total,
    delete_product,
    export_products_to_excel,
    export_sales_to_excel,
    get_all_products,
    get_daily_income_report,
    get_product_by_barcode,
    get_sales_report,
    initialize_db,
    record_sale,
    search_products,
    update_product,
)


class BillingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bagmati lungdar pasal - Billing System")
        self.root.geometry("1150x760")
        self.db_path = initialize_db()
        self.selected_product_id = None
        self.bill_products = []

        if Workbook is not None:
            export_main_excel(self.db_path, get_main_excel_path(self.db_path))

        if not self.authenticate_user():
            self.root.destroy()
            return

        self.build_ui()
        self.load_products()

    def authenticate_user(self) -> bool:
        for _ in range(3):
            username = simpledialog.askstring("Login", "Username", parent=self.root)
            if not username:
                return False
            password = simpledialog.askstring("Login", "Password", show="*", parent=self.root)
            if not password:
                return False
            if authenticate_user(self.db_path, username, password):
                return True
            messagebox.showerror("Login Failed", "Invalid username or password")
        return False

    def build_ui(self):
        self.root.configure(padx=16, pady=16)

        title = ttk.Label(self.root, text="Bagmati lungdar pasal", font=("Segoe UI", 18, "bold"))
        title.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))

        left_frame = ttk.LabelFrame(self.root, text="Product Management")
        left_frame.grid(row=1, column=0, padx=(0, 10), sticky="nsew")

        ttk.Label(left_frame, text="Product Name").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        self.name_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.name_var, width=30).grid(row=1, column=0, padx=8, pady=4)

        ttk.Label(left_frame, text="Price").grid(row=2, column=0, sticky="w", padx=8, pady=4)
        self.price_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.price_var, width=30).grid(row=3, column=0, padx=8, pady=4)

        ttk.Label(left_frame, text="Barcode").grid(row=4, column=0, sticky="w", padx=8, pady=4)
        self.barcode_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.barcode_var, width=30).grid(row=5, column=0, padx=8, pady=4)

        ttk.Label(left_frame, text="Stock Quantity").grid(row=6, column=0, sticky="w", padx=8, pady=4)
        self.stock_var = tk.StringVar(value="0")
        ttk.Entry(left_frame, textvariable=self.stock_var, width=30).grid(row=7, column=0, padx=8, pady=4)

        button_row = ttk.Frame(left_frame)
        button_row.grid(row=8, column=0, sticky="w", padx=8, pady=8)
        ttk.Button(button_row, text="Add", command=self.add_product).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(button_row, text="Update", command=self.update_selected_product).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(button_row, text="Delete", command=self.delete_selected_product).pack(side=tk.LEFT)

        ttk.Label(left_frame, text="Search").grid(row=9, column=0, sticky="w", padx=8, pady=(12, 4))
        self.search_var = tk.StringVar()
        ttk.Entry(left_frame, textvariable=self.search_var, width=30).grid(row=10, column=0, padx=8, pady=4)
        ttk.Button(left_frame, text="Search", command=self.search_products).grid(row=11, column=0, sticky="w", padx=8, pady=4)

        self.product_tree = ttk.Treeview(left_frame, columns=("name", "price", "stock", "barcode", "id"), show="headings", height=12)
        self.product_tree.heading("name", text="Name")
        self.product_tree.heading("price", text="Price")
        self.product_tree.heading("stock", text="Stock")
        self.product_tree.heading("barcode", text="Barcode")
        self.product_tree.column("name", width=160)
        self.product_tree.column("price", width=80)
        self.product_tree.column("stock", width=70)
        self.product_tree.column("barcode", width=100)
        self.product_tree.column("id", width=0, stretch=False)
        self.product_tree.grid(row=12, column=0, padx=8, pady=8)
        self.product_tree.bind("<<TreeviewSelect>>", self.on_product_select)

        right_frame = ttk.LabelFrame(self.root, text="Billing")
        right_frame.grid(row=1, column=1, columnspan=3, padx=(0, 0), sticky="nsew")

        ttk.Label(right_frame, text="Select Product").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        self.bill_product_var = tk.StringVar()
        self.bill_product_combo = ttk.Combobox(right_frame, textvariable=self.bill_product_var, state="readonly", width=30)
        self.bill_product_combo.grid(row=1, column=0, padx=8, pady=4)
        self.bill_product_combo.bind("<<ComboboxSelected>>", self.on_bill_product_select)

        ttk.Label(right_frame, text="Barcode Scan").grid(row=0, column=1, sticky="w", padx=8, pady=4)
        self.barcode_scan_var = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.barcode_scan_var, width=30).grid(row=1, column=1, padx=8, pady=4)
        ttk.Button(right_frame, text="Lookup Barcode", command=self.lookup_by_barcode).grid(row=1, column=2, padx=8, pady=4)

        ttk.Label(right_frame, text="Quantity").grid(row=2, column=0, sticky="w", padx=8, pady=4)
        self.quantity_var = tk.StringVar(value="1")
        ttk.Entry(right_frame, textvariable=self.quantity_var, width=30).grid(row=2, column=1, padx=8, pady=4)

        ttk.Button(right_frame, text="Add to Bill", command=self.add_to_bill).grid(row=3, column=0, sticky="w", padx=8, pady=8)

        ttk.Label(right_frame, text="Customer Name").grid(row=4, column=0, sticky="w", padx=8, pady=4)
        self.customer_name_var = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.customer_name_var, width=30).grid(row=5, column=0, padx=8, pady=4)

        ttk.Label(right_frame, text="Customer Phone").grid(row=4, column=1, sticky="w", padx=8, pady=4)
        self.customer_phone_var = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.customer_phone_var, width=30).grid(row=5, column=1, padx=8, pady=4)

        ttk.Label(right_frame, text="Discount %").grid(row=6, column=0, sticky="w", padx=8, pady=4)
        self.discount_var = tk.StringVar(value="0")
        ttk.Entry(right_frame, textvariable=self.discount_var, width=30).grid(row=7, column=0, padx=8, pady=4)

        ttk.Label(right_frame, text="VAT %").grid(row=6, column=1, sticky="w", padx=8, pady=4)
        self.vat_var = tk.StringVar(value="0")
        ttk.Entry(right_frame, textvariable=self.vat_var, width=30).grid(row=7, column=1, padx=8, pady=4)

        self.bill_tree = ttk.Treeview(right_frame, columns=("name", "price", "qty", "total"), show="headings", height=10)
        self.bill_tree.heading("name", text="Product")
        self.bill_tree.heading("price", text="Price")
        self.bill_tree.heading("qty", text="Qty")
        self.bill_tree.heading("total", text="Total")
        self.bill_tree.column("name", width=180)
        self.bill_tree.column("price", width=90)
        self.bill_tree.column("qty", width=70)
        self.bill_tree.column("total", width=90)
        self.bill_tree.grid(row=8, column=0, columnspan=3, padx=8, pady=8)

        total_frame = ttk.Frame(right_frame)
        total_frame.grid(row=9, column=0, columnspan=3, sticky="e", padx=8, pady=8)
        ttk.Label(total_frame, text="Grand Total: ", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT)
        self.total_var = tk.StringVar(value="Rs. 0.00")
        ttk.Label(total_frame, textvariable=self.total_var, font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT)

        action_frame = ttk.Frame(right_frame)
        action_frame.grid(row=10, column=0, columnspan=3, sticky="w", padx=8, pady=8)
        ttk.Button(action_frame, text="Checkout & Save Receipt", command=self.checkout_bill).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(action_frame, text="Clear Bill", command=self.clear_bill).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(action_frame, text="Sales Report", command=self.show_sales_report).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(action_frame, text="Export Products to Excel", command=self.export_products_excel).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(action_frame, text="Export Sales to Excel", command=self.export_sales_excel).pack(side=tk.LEFT)

        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.columnconfigure(2, weight=1)
        self.root.columnconfigure(3, weight=1)
        self.root.rowconfigure(1, weight=1)

    def load_products(self):
        products = get_all_products(self.db_path)
        self.product_tree.delete(*self.product_tree.get_children())
        self.bill_products = []
        for product in products:
            self.product_tree.insert(
                "",
                tk.END,
                values=(product["name"], product["price"], product.get("stock_quantity", 0), product.get("barcode", ""), product["id"]),
            )
            self.bill_products.append(
                {
                    "id": product["id"],
                    "name": product["name"],
                    "price": product["price"],
                    "stock_quantity": product.get("stock_quantity", 0),
                }
            )
        self.update_combo_values()

    def update_combo_values(self):
        values = [item["name"] for item in self.bill_products]
        self.bill_product_combo["values"] = values
        if values:
            self.bill_product_combo.current(0)

    def on_product_select(self, _event):
        selected = self.product_tree.selection()
        if not selected:
            return
        item = self.product_tree.item(selected[0])
        self.name_var.set(item["values"][0])
        self.price_var.set(item["values"][1])
        self.barcode_var.set(item["values"][3])
        self.stock_var.set(item["values"][2])
        self.selected_product_id = item["values"][4] if len(item["values"]) > 4 else None

    def add_product(self):
        name = self.name_var.get().strip()
        price_text = self.price_var.get().strip()
        if not name or not price_text:
            messagebox.showwarning("Input Missing", "Please enter both product name and price.")
            return
        try:
            price = float(price_text)
            stock_quantity = int(self.stock_var.get().strip() or 0)
        except ValueError:
            messagebox.showerror("Invalid Input", "Price and stock must be numbers.")
            return

        add_product(self.db_path, name, price, barcode=self.barcode_var.get().strip(), stock_quantity=stock_quantity)
        self.name_var.set("")
        self.price_var.set("")
        self.barcode_var.set("")
        self.stock_var.set("0")
        self.load_products()
        messagebox.showinfo("Success", f"Added {name}")

    def update_selected_product(self):
        if self.selected_product_id is None:
            messagebox.showwarning("Select Product", "Select a product from the list first.")
            return
        name = self.name_var.get().strip()
        price_text = self.price_var.get().strip()
        if not name or not price_text:
            messagebox.showwarning("Input Missing", "Please enter both product name and price.")
            return
        try:
            price = float(price_text)
            stock_quantity = int(self.stock_var.get().strip() or 0)
        except ValueError:
            messagebox.showerror("Invalid Input", "Price and stock must be numbers.")
            return

        update_product(
            self.db_path,
            self.selected_product_id,
            name,
            price,
            barcode=self.barcode_var.get().strip(),
            stock_quantity=stock_quantity,
        )
        self.selected_product_id = None
        self.name_var.set("")
        self.price_var.set("")
        self.barcode_var.set("")
        self.stock_var.set("0")
        self.load_products()
        messagebox.showinfo("Success", f"Updated {name}")

    def delete_selected_product(self):
        if self.selected_product_id is None:
            messagebox.showwarning("Select Product", "Select a product from the list first.")
            return
        confirm = messagebox.askyesno("Confirm Delete", "Delete this product?")
        if confirm:
            delete_product(self.db_path, self.selected_product_id)
            self.selected_product_id = None
            self.name_var.set("")
            self.price_var.set("")
            self.barcode_var.set("")
            self.stock_var.set("0")
            self.load_products()

    def search_products(self):
        term = self.search_var.get().strip()
        if not term:
            self.load_products()
            return
        products = search_products(self.db_path, term)
        self.product_tree.delete(*self.product_tree.get_children())
        for product in products:
            self.product_tree.insert(
                "",
                tk.END,
                values=(product["name"], product["price"], product.get("stock_quantity", 0), product.get("barcode", ""), product["id"]),
            )

    def lookup_by_barcode(self):
        barcode = self.barcode_scan_var.get().strip()
        if not barcode:
            messagebox.showwarning("Barcode Missing", "Enter a barcode first.")
            return
        product = get_product_by_barcode(self.db_path, barcode)
        if product is None:
            messagebox.showwarning("Not Found", "No product found for that barcode.")
            return
        self.bill_product_var.set(product["name"])
        self.quantity_var.set("1")
        self.barcode_scan_var.set("")

    def on_bill_product_select(self, _event):
        selected_name = self.bill_product_var.get()
        if not selected_name:
            return
        for item in self.bill_products:
            if item["name"] == selected_name:
                self.selected_bill_product = item
                break

    def add_to_bill(self):
        selected_name = self.bill_product_var.get()
        if not selected_name:
            messagebox.showwarning("Select Product", "Choose a product first.")
            return
        try:
            quantity = int(self.quantity_var.get())
        except ValueError:
            messagebox.showerror("Invalid Quantity", "Quantity must be a whole number.")
            return
        if quantity <= 0:
            messagebox.showwarning("Invalid Quantity", "Quantity must be greater than zero.")
            return

        product = next((item for item in self.bill_products if item["name"] == selected_name), None)
        if product is None:
            messagebox.showwarning("Missing Product", "Selected product could not be found.")
            return
        if product["stock_quantity"] < quantity:
            messagebox.showwarning("Out of Stock", "Not enough stock available for this item.")
            return

        item_total = calculate_item_total(product["price"], quantity)
        self.bill_tree.insert("", tk.END, values=(selected_name, product["price"], quantity, item_total))
        self.update_total()

    def update_total(self):
        items = []
        for child in self.bill_tree.get_children():
            item = self.bill_tree.item(child)
            items.append({
                "price": float(item["values"][1]),
                "quantity": int(item["values"][2]),
            })
        subtotal = calculate_grand_total(items)
        discount_rate = self._parse_float(self.discount_var.get(), 0)
        vat_rate = self._parse_float(self.vat_var.get(), 0)
        discount_amount = subtotal * (discount_rate / 100.0)
        taxable = subtotal - discount_amount
        vat_amount = taxable * (vat_rate / 100.0)
        total_amount = taxable + vat_amount
        self.total_var.set(f"Rs. {total_amount:.2f}")

    def _parse_float(self, value: str, default: float) -> float:
        try:
            return float(value)
        except ValueError:
            return default

    def checkout_bill(self):
        if not self.bill_tree.get_children():
            messagebox.showwarning("Empty Bill", "Add at least one item before checkout.")
            return

        items = []
        for child in self.bill_tree.get_children():
            item = self.bill_tree.item(child)
            product = next((entry for entry in self.bill_products if entry["name"] == item["values"][0]), None)
            if product is None:
                continue
            items.append({
                "product_id": product["id"],
                "price": float(item["values"][1]),
                "quantity": int(item["values"][2]),
            })

        try:
            sale_id = record_sale(
                self.db_path,
                items,
                customer_name=self.customer_name_var.get().strip(),
                customer_phone=self.customer_phone_var.get().strip(),
                discount_rate=self._parse_float(self.discount_var.get(), 0),
                vat_rate=self._parse_float(self.vat_var.get(), 0),
            )
        except ValueError as exc:
            messagebox.showerror("Checkout Failed", str(exc))
            return

        self.export_receipt(sale_id, items)
        self.clear_bill()
        self.customer_name_var.set("")
        self.customer_phone_var.set("")
        self.discount_var.set("0")
        self.vat_var.set("0")
        self.load_products()
        messagebox.showinfo("Success", f"Receipt saved and sale recorded. Sale ID: {sale_id}")

    def export_receipt(self, sale_id: int, items: list):
        default_name = f"receipt_{sale_id}.pdf"
        save_path = filedialog.asksaveasfilename(
            initialfile=default_name,
            defaultextension=".pdf",
            title="Save Receipt",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if not save_path:
            return
        if canvas is None or LETTER is None:
            messagebox.showerror("Missing Dependency", "Install reportlab to create PDF receipts: pip install reportlab")
            return

        subtotal = sum(float(item["price"]) * int(item["quantity"]) for item in items)
        discount_rate = self._parse_float(self.discount_var.get(), 0)
        vat_rate = self._parse_float(self.vat_var.get(), 0)
        discount_amount = subtotal * (discount_rate / 100.0)
        taxable = subtotal - discount_amount
        vat_amount = taxable * (vat_rate / 100.0)
        total_amount = taxable + vat_amount

        page_width = 300
        page_height = max(620, 420 + len(items) * 22)
        pdf = canvas.Canvas(save_path, pagesize=(page_width, page_height))
        left = 28
        right = page_width - left
        center = page_width / 2
        y = page_height - 48

        def centered(text, font="Helvetica", size=11, gap=18):
            nonlocal y
            pdf.setFont(font, size)
            pdf.drawCentredString(center, y, text)
            y -= gap

        def row(label, value, gap=18):
            nonlocal y
            pdf.setFont("Helvetica", 10)
            pdf.drawString(left, y, label)
            pdf.drawRightString(right, y, value)
            y -= gap

        centered("Bagmati lungdar pasal", "Helvetica-Bold", 13, 20)
        centered("BILLING RECEIPT", "Helvetica", 11, 32)

        pdf.setFont("Helvetica", 10)
        pdf.drawString(left, y, f"SALE: {sale_id}")
        pdf.drawRightString(right, y, f"{datetime.now():%d/%m/%Y}")
        y -= 18
        pdf.drawString(left, y, f"CUSTOMER: {self.customer_name_var.get().strip() or 'Walk-in'}")
        pdf.drawRightString(right, y, f"{datetime.now():%I:%M %p}")
        y -= 18
        pdf.drawString(left, y, f"PHONE: {self.customer_phone_var.get().strip() or '-'}")
        pdf.drawRightString(right, y, "CASHIER: ADMIN")
        y -= 24

        pdf.line(left, y, right, y)
        y -= 22
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(left, y, "QTY")
        pdf.drawString(left + 42, y, "ITEM")
        pdf.drawRightString(right, y, "AMOUNT")
        y -= 18
        pdf.setFont("Helvetica", 10)
        for item in items:
            product = next((entry for entry in self.bill_products if entry["id"] == item["product_id"]), None)
            product_name = product["name"] if product else "Unknown"
            item_total = float(item["price"]) * int(item["quantity"])
            pdf.drawString(left, y, str(item["quantity"]))
            pdf.drawString(left + 42, y, product_name[:25])
            pdf.drawRightString(right, y, f"Rs. {item_total:.2f}")
            y -= 20

        pdf.line(left, y + 4, right, y + 4)
        y -= 18
        pdf.setFont("Helvetica", 10)
        row("CASH", "SALE")
        y -= 8
        row("SUBTOTAL", f"Rs. {subtotal:.2f}")
        row("DISCOUNT", f"{discount_rate:.0f}%")
        row("VAT", f"{vat_rate:.0f}%")
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(left, y, "TOTAL")
        pdf.drawRightString(right, y, f"Rs. {total_amount:.2f}")
        y -= 34
        centered("THANK YOU FOR SHOPPING", "Helvetica-Bold", 10, 16)
        centered("WITH US!", "Helvetica-Bold", 10, 10)
        pdf.save()

    def show_sales_report(self):
        report = get_sales_report(self.db_path)
        daily_report = get_daily_income_report(self.db_path)
        message = [
            f"Today: {daily_report['sales_count']} sales, total Rs. {daily_report['total_sales']:.2f}",
            "",
            "Recent Sales:",
        ]
        for row in report[:10]:
            message.append(f"- {row['sale_date']} | {row['customer_name']} | Rs. {row['total_amount']:.2f}")
        messagebox.showinfo("Sales Report", "\n".join(message))

    def export_products_excel(self):
        try:
            file_path = filedialog.asksaveasfilename(
                initialfile="products.xlsx",
                defaultextension=".xlsx",
                title="Save Products Excel",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            )
            if not file_path:
                return
            export_products_to_excel(self.db_path, file_path)
            messagebox.showinfo("Exported", f"Products exported to {file_path}")
        except ImportError:
            messagebox.showerror("Missing Dependency", "Install openpyxl to export Excel files: pip install openpyxl")
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("Export Failed", str(exc))

    def export_sales_excel(self):
        try:
            file_path = filedialog.asksaveasfilename(
                initialfile="sales.xlsx",
                defaultextension=".xlsx",
                title="Save Sales Excel",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            )
            if not file_path:
                return
            export_sales_to_excel(self.db_path, file_path)
            messagebox.showinfo("Exported", f"Sales exported to {file_path}")
        except ImportError:
            messagebox.showerror("Missing Dependency", "Install openpyxl to export Excel files: pip install openpyxl")
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("Export Failed", str(exc))

    def clear_bill(self):
        for child in self.bill_tree.get_children():
            self.bill_tree.delete(child)
        self.total_var.set("Rs. 0.00")


def main():
    root = tk.Tk()
    try:
        app = BillingApp(root)
    except tk.TclError:
        root.destroy()
        return

    if getattr(app, "root", None) is None:
        return

    try:
        if app.root.winfo_exists():
            root.mainloop()
    except tk.TclError:
        root.destroy()


if __name__ == "__main__":
    main()
