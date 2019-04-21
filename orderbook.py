import sqlite3

class Orderbook(object):
    def __init__(self, database_file="orderbook.db"):
        self.db_file = database_file

    def _convert_order_to_row(self, order):
        order_array = [order.get("orderID"), order.get("typeID"), order.get("itemName"),
        order.get("regionID"), order.get("regionName"), order.get("stationID"),
        order.get("stationName"), order.get("range"), order.get("price"),
        order.get("volEntered"), order.get("volRemaining"), order.get("issueDate"),
        order.get("minVolume"), order.get("duration"), order.get("solarSystemID"),
        order.get("solarSystemName"), order.get("escrow"), order.get("bid")]
        return order_array

    def _convert_row_to_order(self, row):
        keys = ['orderID', 'typeID', 'itemName', 'regionID', 'regionName', 'stationID',
        'stationName', 'range', 'price', 'volEntered', 'volRemaining', 'issueDate',
        'minVolume', 'duration', 'solarSystemID', 'solarSystemName', 'escrow', 'bid']
        res = dict(zip(keys, row))
        return res

    def _convert_item_to_row(self, item):
        item_array = [item.get("typeID"), item.get("typeName"), item.get("price"), item.get("volume")]
        return item_array

    def _convert_row_to_item(self, row):
        keys = ['typeID', 'typeName', 'price', 'volume']
        res = dict(zip(keys, row))
        return res

    def _fetch_one(self, query, args=None):
        # Private function that executes a query and retrieves a single row.
        # args should be either a single value tuple or an array of substitution values.
        dbconn = sqlite3.connect(self.db_file)
        db_cursor = dbconn.cursor()
        if args:
            db_cursor.execute(query, args)
        else:
            db_cursor.execute(query)
        data = db_cursor.fetchone()
        db_cursor.close()
        dbconn.close()
        return data

    def _fetch_many(self, query, args=None):
        # Private function that executes a query and retrieves all returned rows.
        # args should be either a single value tuple or an array of substitution values.
        dbconn = sqlite3.connect(self.db_file)
        db_cursor = dbconn.cursor()
        if args:
            db_cursor.execute(query, args)
        else:
            db_cursor.execute(query)
        data = db_cursor.fetchall()
        db_cursor.close()
        dbconn.close()
        return data

    def _insert_query(self, query, args=None):
        # Private function that execute a query and returns nothing.
        # args should be either a single value tuple or an array of substitution values.
        dbconn = sqlite3.connect(self.db_file)
        db_cursor = dbconn.cursor()
        if args:
            db_cursor.execute(query, args)
        else:
            db_cursor.execute(query)
        dbconn.commit()
        db_cursor.close()
        dbconn.close()

    def init_tables(self):
        # TODO: Right now order data is all of type text. Could correct this if it is helpful.
        order_table_query = """ CREATE TABLE IF NOT EXISTS orders (orderID text, typeID text,
        typeName text, regionID text, regionName text, stationID text, stationName text,
        range text, price text, volumeStart text, volumeRemaining text, issueDate text,
        minVolume text, duration text, solarSystemID text, solarSystemName text, escrow text, bid text)"""
        self._insert_query(order_table_query)
        inventory_table_query = """CREATE TABLE IF NOT EXISTS inventory (typeID text, typeName text, price text, volume text)"""
        self._insert_query(inventory_table_query)

    def clean_database(self):
        # Primarily used in testing.
        drop_orders_query = "DROP TABLE IF EXISTS orders"
        self._insert_query(drop_orders_query)
        drop_inventory_query = "DROP TABLE IF EXISTS inventory"
        self._insert_query(drop_inventory_query)

    def insert_order(self, order):
        # insert_order accepts a Marketlog row from the wallet/market csv files.
        order_array = self._convert_order_to_row(order)
        query = "INSERT INTO orders VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
        self._insert_query(query, order_array)

    def retrieve_buy_order(self, orderID):
        # This is a good example of how I've wrapped sqlite3. The goal is to eliminate any dangling
        # db connections and ensure uniform interactions with the database.
        query = "SELECT * FROM orders WHERE orderID=? AND bid='True'"
        # sqlite3's execute substitution appears to require a plural datatype. To satisfy this we send a tuple.
        return self._convert_row_to_order(self._fetch_one(query, (orderID,)))

    def retrieve_sell_order(self, orderID):
        # orderID is unique across all orders, and for now each order gets updated in-place vs creating another row.
        query = "SELECT * FROM orders WHERE orderID=? AND bid='False'"
        data = self._fetch_one(query, (orderID,))
        return self._convert_row_to_order(data)

    def retrieve_buy_orders(self):
        # Bulk retrievals are not that useful outside of for some initial verification checks.
        query = "SELECT * FROM orders WHERE bid='True'"
        data = self._fetch_many(query)
        orders = []
        for row in data:
            orders.append(self._convert_row_to_order(row))
        return orders

    def retrieve_sell_orders(self):
        query = "SELECT * FROM orders WHERE bid='False'"
        data = self._fetch_many(query)
        orders = []
        for row in data:
            orders.append(self._convert_row_to_order(row))
        return orders

    def update_order(self, order):
        # Right now we only bother to update the price and remaining volume.
        # Until there is a reason to update other columns, these seem sufficient.
        query = "UPDATE orders SET price = ?, volumeRemaining = ? WHERE orderID = ?"
        args = [order.get("price"), order.get("volRemaining"), order.get("orderID")]
        self._insert_query(query, args)

    def insert_inventory(self, item):
        # Item: {typeID: string, typeName: string, price: string, volume: string}
        item_array = self._convert_item_to_row(item)
        query = "INSERT INTO inventory VALUES (?,?,?,?)"
        self._insert_query(query, item_array)

    def get_inventory(self, typeID):
        query = "SELECT * FROM inventory WHERE typeID=?"
        inventory = self._fetch_many(query, (typeID,))
        items = []
        for item in inventory:
            items.append(self._convert_row_to_item(item))
        return items

class InventoryLedger(object):
    def __init__(self, orderbook):
        self.orderbook = orderbook

    def process_orders(self, orders):
        # Given an array of order objects, iterate through them and check if an order already exists
        # If an order exists, and the volume has changed, log an increase in inventory with the previous price.
        if len(orders) == 0:
            raise Exception("Zero orders received")
        for order in orders:
            existing_order = self.orderbook.retrieve_buy_order(order.get("orderID"))
            if existing_order:
                inventory_entry = self._process_order_diff(existing_order, order)
                if inventory_entry:
                    # If there is a volume change, inventory_entry gets set.
                    self.orderbook.insert_inventory(inventory_entry)
                # Regardless of volume change, update the order with fresh data.
                self.orderbook.update_order(order)

    def _process_order_diff(self, existing_order, new_order):
        # existing_order = entry in the database, new_order = line from marketorder csv
        if new_order.get("volRemaining") < existing_order.get("volRemaining"):
            volume_diff = existing_order.get("volRemaining") - new_order.get("volRemaining")
            print("Volume changed for %s with volume diff of %s" % (new_order.get("itemName"), volume_diff))
            item = {"typeID": existing_order.get("typeID"),
            "typeName": existing_order.get("typeName"),
            "price": existing_order.get("price"), 
            "volume": volume_diff}
            return item
        return None