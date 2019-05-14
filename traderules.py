import pickle

class TradeRules(object):
    # This object will contain the rules on which we will base trades.
    # The idea is that there are some default rules that apply to trades based on range.
    # Outside of these defaults, there will be the option for overriding rules that will
    # permit special behavior on certain items.
    # For now I will solely focus on sell-side within 0 range, and buy-side within 1 range.
    # Region range and intermediate range isn't within the scope of the current goal.
    # TODO: leverage historical market data to determine trends and sane max/min values.

    # Sell logic:
    #   Filter market orders to orders that are in the current station of the target order.
    #   Check the item acquisition price, this + a minimum margin is the lowest we can go.
    #
    # Buy logic:
    #   Assumption: we are trading in a hub, placing our buy orders within 1 range of said hub.
    #   Only compete with orders within 1 range, unless overwise overridden.
    #   Limit the maximum price to the minimum sell price + healthy margin.
    def __init__(self):
        self.rules = {}

    def _set_defaults(self):
        if not self.rules.get("default"):
            default = {
                "buy": None,
                "sell": None
                }
            buy = {
                "margin": 12.0,
                "range": 1,
                "price_delta": 15.0
                }
            sell = {
                "margin": 12.0,
                "range": 0,
                "price_delta": 15.0
            }
            default['buy'] = buy
            default['sell'] = sell
            self.rules["default"] = default

    def save_rules(self, filename="rules.pkl"):
        with open(filename, "wb") as fp:
            pickle.dump(self.rules, fp, protocol=pickle.HIGHEST_PROTOCOL)

    def load_rules(self, filename="rules.pkl"):
        with open(filename, "rb") as fp:
            self.rules = pickle.load(fp)

    def get_rules(self, typeID):
        return self.rules.get(typeID, self.rules.get("default"))

    def negotiate_buy_order(self, target_order, market_processor):
        rules = self.get_rules(target_order.get("typeID"))
        top_buy_order = market_processor.get_max_buy_order(target_order.get("stationID"))
        top_sell_order = market_processor.get_min_sell_order(target_order.get("stationID"))
        expected_gross_profit = float(top_sell_order.get("price")) - float(top_buy_order.get("price"))
        expected_margin = (expected_gross_profit/float(top_sell_order.get("price"))) * 100
        if expected_margin > rules["buy"].get("margin"):
            price_delta = float(top_buy_order.get("price")) - float(target_order.get("price"))
            delta_percentage = (price_delta / float(top_buy_order.get("price"))) * 100
            if delta_percentage < rules["buy"].get("price_delta"):
                # We'll naively make at least our minimum margin, and the leading order hasn't moved so much as to be alarming.
                target_price = float(top_buy_order.get("price")) + 0.01
                return str(round(target_price, 2))
            else:
                print(f"Order of type {target_order.get('itemName')} with a current price of {target_order.get('price')} has exceeded the price_delta rule of {rules['buy'].get('price_delta')} with a value of {delta_percentage} and a price increase of {top_buy_order.get('price')}.")
                return None
        else:
            print(f"Minimum margin for order of type {target_order.get('itemName')} not met with a expected margin of {expected_margin} with a buy price of {top_buy_order.get('price')} and a sell price of {top_sell_order.get('price')} against a minimum margin of {rules['buy'].get('margin')}.")
            return None