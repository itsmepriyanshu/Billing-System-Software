import os
import tempfile
import tkinter
import unittest
from unittest.mock import patch

import app as app_module
from database import (
    add_product,
    authenticate_user,
    calculate_grand_total,
    calculate_item_total,
    delete_product,
    get_all_products,
    get_daily_income_report,
    get_product_by_barcode,
    get_sales_report,
    initialize_db,
    record_sale,
    search_products,
    update_product,
)


class BillingSystemTests(unittest.TestCase):
    def setUp(self):
        self.db_file = tempfile.NamedTemporaryFile(delete=False).name
        initialize_db(self.db_file)

    def tearDown(self):
        if os.path.exists(self.db_file):
            os.remove(self.db_file)

    def test_calculate_item_total_and_grand_total(self):
        items = [
            {"price": 1200, "quantity": 2},
            {"price": 2500, "quantity": 1},
        ]

        self.assertEqual(calculate_item_total(1200, 2), 2400)
        self.assertEqual(calculate_grand_total(items), 4900)

    def test_product_crud_and_search(self):
        product_id = add_product(self.db_file, "T-Shirt", 1200)
        self.assertIsNotNone(product_id)

        products = get_all_products(self.db_file)
        self.assertEqual(len(products), 1)
        self.assertEqual(products[0]["name"], "T-Shirt")

        updated = update_product(self.db_file, product_id, "T-Shirt", 1300)
        self.assertTrue(updated)

        search_results = search_products(self.db_file, "shirt")
        self.assertEqual(len(search_results), 1)
        self.assertEqual(search_results[0]["price"], 1300)

        deleted = delete_product(self.db_file, product_id)
        self.assertTrue(deleted)
        self.assertEqual(get_all_products(self.db_file), [])

    def test_sale_recording_stock_and_reports(self):
        product_id = add_product(self.db_file, "Jeans", 2500, barcode="J100", stock_quantity=5)
        self.assertTrue(authenticate_user(self.db_file, "admin", "admin"))

        sale_id = record_sale(
            self.db_file,
            [{"product_id": product_id, "price": 2500, "quantity": 2}],
            customer_name="Ali",
            customer_phone="12345",
            discount_rate=10,
            vat_rate=5,
        )
        self.assertIsNotNone(sale_id)

        product = get_product_by_barcode(self.db_file, "J100")
        self.assertEqual(product["stock_quantity"], 3)

        report = get_sales_report(self.db_file)
        self.assertEqual(len(report), 1)
        self.assertEqual(report[0]["customer_name"], "Ali")

        daily_report = get_daily_income_report(self.db_file)
        self.assertEqual(daily_report["sales_count"], 1)
        self.assertGreater(daily_report["total_sales"], 0)

    def test_ensure_optional_dependency_installs_missing_package(self):
        with patch.object(app_module.importlib, "import_module", side_effect=[ImportError("missing"), object()]) as mock_import_module, patch.object(app_module.subprocess, "run") as mock_run:
            self.assertTrue(app_module.ensure_optional_dependency("reportlab"))
            self.assertEqual(mock_import_module.call_count, 2)
            mock_run.assert_called_once()

    def test_main_does_not_call_winfo_after_login_failure(self):
        class DummyRoot:
            def __init__(self):
                self.destroyed = False

            def destroy(self):
                self.destroyed = True

            def title(self, *_args, **_kwargs):
                return None

            def geometry(self, *_args, **_kwargs):
                return None

            def configure(self, *_args, **_kwargs):
                return None

            def winfo_exists(self):
                raise tkinter.TclError("application has been destroyed")

        root = DummyRoot()
        with patch.object(app_module.BillingApp, "authenticate_user", return_value=False), patch.object(app_module.tk, "Tk", return_value=root):
            app_module.main()

        self.assertTrue(root.destroyed)


if __name__ == "__main__":
    unittest.main()
