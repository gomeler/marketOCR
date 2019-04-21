import copy
import os
import unittest


import orderbook

class TestOrderbook(unittest.TestCase):
    def setUp(self):
        self.db_file = "test.db"
        self.book = orderbook.Orderbook(self.db_file)
        self.book.init_tables()
        self.buy_order = {'orderID': '1337', 'typeID': '222', 'itemName': 'Antimatter Charge S',
        'regionID': '10000002', 'regionName': 'The Forge', 'stationID': '42',
        'stationName': 'Best Station', 'range': '-1', 'price': '1337.0', 'volEntered': '100',
        'volRemaining': '100.0', 'issueDate': '2019-04-20 20:39:07.000', 'minVolume': '1', 'duration': '90',
        'solarSystemID': '30000119', 'solarSystemName': 'Itamo', 'escrow': '500.0', 'bid': 'True'}

        self.buy_order_2 = {'orderID': '1338', 'typeID': '222', 'itemName': 'Antimatter Charge S',
        'regionID': '10000002', 'regionName': 'The Forge', 'stationID': '42',
        'stationName': 'Best Station', 'range': '-1', 'price': '1337.0', 'volEntered': '100',
        'volRemaining': '100.0', 'issueDate': '2019-04-20 20:39:07.000', 'minVolume': '1', 'duration': '90',
        'solarSystemID': '30000119', 'solarSystemName': 'Itamo', 'escrow': '500.0', 'bid': 'True'}

        self.sell_order = {'orderID': '42', 'typeID': '15331', 'itemName': 'Metal Scraps',
        'regionID': '10000033', 'regionName': 'The Citadel', 'stationID': '99',
        'stationName': 'Another Station', 'range': '32767', 'price': '1200.0', 'volEntered': '7',
        'volRemaining': '7.0', 'issueDate': '2019-04-04 01:39:25.000', 'minVolume': '1',
        'duration': '90', 'solarSystemID': '30002783', 'solarSystemName': 'Sankkasen', 'escrow': '0.0', 'bid': 'False'}

        self.sell_order_2 = {'orderID': '43', 'typeID': '15331', 'itemName': 'Metal Scraps',
        'regionID': '10000033', 'regionName': 'The Citadel', 'stationID': '99',
        'stationName': 'Another Station', 'range': '32767', 'price': '1200.0', 'volEntered': '7',
        'volRemaining': '7.0', 'issueDate': '2019-04-04 01:39:25.000', 'minVolume': '1',
        'duration': '90', 'solarSystemID': '30002783', 'solarSystemName': 'Sankkasen', 'escrow': '0.0', 'bid': 'False'}
    
    def tearDown(self):
        del(self.book)
        if os.path.exists(self.db_file):
            os.remove(self.db_file)

    def test_init_tables(self):
        # Verify the tables were created.
        table_data = self.book._fetch_many("SELECT * from sqlite_master")
        self.assertEqual(len(table_data), 2)
        table_names = [table[1] for table in table_data]
        self.assertIn("orders", table_names)
        self.assertIn("inventory", table_names)

    def test_clean_database(self):
        # Verify we can purge the tables.
        self.book.clean_database()
        table_data = self.book._fetch_many("SELECT * from sqlite_master")
        self.assertEqual(len(table_data), 0)

    def test_insert_retrieve_buy_order(self):
        # Verify we can insert and retrieve a buy order.
        self.book.insert_order(self.buy_order)
        retrieved_order = self.book.retrieve_buy_order(self.buy_order.get("orderID"))
        self.assertEqual(retrieved_order, self.buy_order)

    def test_insert_retrieve_sell_order(self):
        # Verify we can insert and retrieve a sell order.
        self.book.insert_order(self.sell_order)
        retrieved_order = self.book.retrieve_sell_order(self.sell_order.get("orderID"))
        self.assertEqual(retrieved_order, self.sell_order)

    def test_retrieve_buy_orders(self):
        # Verify we can bulk retrieve buy orders.
        self.book.insert_order(self.buy_order)
        self.book.insert_order(self.buy_order_2)
        orders = self.book.retrieve_buy_orders()
        self.assertEqual(len(orders), 2)
        order_ids = [order.get("orderID") for order in orders]
        self.assertIn(self.buy_order.get("orderID"), order_ids)
        self.assertIn(self.buy_order_2.get("orderID"), order_ids)

    def test_retrieve_sell_orders(self):
        # Verify we can bulk retrieve sell orders.
        self.book.insert_order(self.sell_order)
        self.book.insert_order(self.sell_order_2)
        orders = self.book.retrieve_sell_orders()
        self.assertEqual(len(orders), 2)
        order_ids = [order.get("orderID") for order in orders]
        self.assertIn(self.sell_order.get("orderID"), order_ids)
        self.assertIn(self.sell_order_2.get("orderID"), order_ids)

    def test_update_order(self):
        # Verify we can update the price and remaining volume of an order.
        self.book.insert_order(self.buy_order)
        new_order = copy.copy(self.buy_order)
        new_order["volRemaining"] = "42"
        new_order["price"] = "1336.0"
        self.book.update_order(new_order)
        retrieved_order = self.book.retrieve_buy_order(new_order.get("orderID"))
        self.assertEqual(retrieved_order, new_order)

    def test_insert_retrieve_inventory(self):
        item = {"typeID": "1337", "typeName": "Leet Item", "price": "42", "volume": "9001"}
        self.book.insert_inventory(item)
        retrieved_items = self.book.get_inventory("1337")
        self.assertEqual(len(retrieved_items), 1)
        self.assertEqual(item, retrieved_items[0])

class TestLedger(unittest.TestCase):
    def setUp(self):
        self.db_file = "test.db"
        self.ledger = orderbook.Ledger(self.db_file)
        self.ledger.init_tables()

    def tearDown(self):
        del(self.ledger)
        if os.path.exists(self.db_file):
            os.remove(self.db_file)

    def test_process_order_diff_positive(self):
        new_order = {'orderID': '1337', 'typeID': '222', 'itemName': 'Antimatter Charge S',
        'regionID': '10000002', 'regionName': 'The Forge', 'stationID': '42',
        'stationName': 'Best Station', 'range': '-1', 'price': '1338.0', 'volEntered': '100',
        'volRemaining': '98.0', 'issueDate': '2019-04-20 20:39:07.000', 'minVolume': '1', 'duration': '90',
        'solarSystemID': '30000119', 'solarSystemName': 'Itamo', 'escrow': '500.0', 'bid': 'True'}

        existing_order = {'orderID': '1337', 'typeID': '222', 'itemName': 'Antimatter Charge S',
        'regionID': '10000002', 'regionName': 'The Forge', 'stationID': '42',
        'stationName': 'Best Station', 'range': '-1', 'price': '1337.0', 'volEntered': '100',
        'volRemaining': '100.0', 'issueDate': '2019-04-20 20:39:07.000', 'minVolume': '1', 'duration': '90',
        'solarSystemID': '30000119', 'solarSystemName': 'Itamo', 'escrow': '500.0', 'bid': 'True'}

        item = self.ledger._process_order_diff(existing_order, new_order)
        self.assertEqual(item.get("price"), existing_order.get("price"))
        expected_volume = str(float(existing_order.get("volRemaining")) - float(new_order.get("volRemaining")))
        self.assertEqual(item.get("volume"), expected_volume)
        self.assertEqual(item.get("itemName"), existing_order.get("itemName"))
        self.assertEqual(item.get("typeID"), existing_order.get("typeID"))

    def test_process_order_diff_negative(self):
        order = {'orderID': '1337', 'typeID': '222', 'itemName': 'Antimatter Charge S',
        'regionID': '10000002', 'regionName': 'The Forge', 'stationID': '42',
        'stationName': 'Best Station', 'range': '-1', 'price': '1338.0', 'volEntered': '100',
        'volRemaining': '98.0', 'issueDate': '2019-04-20 20:39:07.000', 'minVolume': '1', 'duration': '90',
        'solarSystemID': '30000119', 'solarSystemName': 'Itamo', 'escrow': '500.0', 'bid': 'True'}
        item = self.ledger._process_order_diff(order, order)
        self.assertIsNone(item)

    def test_process_orders_positive(self):
        # Verify sunny-day logic. An entry should be put into the inventory table,
        # and the order table should be updated.
        order = {'orderID': '1337', 'typeID': '222', 'itemName': 'Antimatter Charge S',
        'regionID': '10000002', 'regionName': 'The Forge', 'stationID': '42',
        'stationName': 'Best Station', 'range': '-1', 'price': '1338.0', 'volEntered': '100',
        'volRemaining': '98.0', 'issueDate': '2019-04-20 20:39:07.000', 'minVolume': '1', 'duration': '90',
        'solarSystemID': '30000119', 'solarSystemName': 'Itamo', 'escrow': '500.0', 'bid': 'True'}

        existing_order = {'orderID': '1337', 'typeID': '222', 'itemName': 'Antimatter Charge S',
        'regionID': '10000002', 'regionName': 'The Forge', 'stationID': '42',
        'stationName': 'Best Station', 'range': '-1', 'price': '1337.0', 'volEntered': '100',
        'volRemaining': '100.0', 'issueDate': '2019-04-20 20:39:07.000', 'minVolume': '1', 'duration': '90',
        'solarSystemID': '30000119', 'solarSystemName': 'Itamo', 'escrow': '500.0', 'bid': 'True'}

        self.ledger.insert_order(existing_order)
        self.ledger.process_orders([order])
        retrieved_inventory = self.ledger.get_inventory(order.get("typeID"))
        self.assertEqual(len(retrieved_inventory), 1)
        self.assertEqual(retrieved_inventory[0].get("typeID"), order.get("typeID"))
        retrieved_order = self.ledger.retrieve_buy_order(order.get("orderID"))
        # Verify the new order equals the updated input "order".
        self.assertEqual(order, retrieved_order)

    def test_process_orders_zero_length_exception(self):
        with self.assertRaises(Exception):
            self.ledger.process_orders([])
