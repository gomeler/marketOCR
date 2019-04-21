from PIL import Image
from ctypes import windll
from html.parser import HTMLParser
from fuzzywuzzy import fuzz


import bs4 as beautifulsoup

import win32gui
import win32ui
import win32con
import win32api
import win32clipboard
import time
import cv2
import numpy
import pytesseract
import time
import win32com.client
import random
import csv
import copy

import screeninteraction

# python-Levenshtein is not necessary but helpful for fuzzywuzzy
# Requires c++ build tools -> https://visualstudio.microsoft.com/visual-cpp-build-tools/
# Grab Build Tools for Visual Studio 2019


# Blackscreen issue
# https://github.com/python-pillow/Pillow/issues/1547
# For some reason windll.user32.PrintWindow is failing to dump the image to the bitmap.

#TODO: Perhaps a custom class is needed for scaled images, so I can have the image and the resize_factor linked.
# TODO: cache the location of on-screen elements and perform restricted tesseract checks vs scanning the full screen.
# TODO: Can refresh the wallet by double clicking on 'My Wallet'
# TODO: Might also be worth using a keyword for the bottom of the Buying section, like the Orders Remaining line?
# TODO: Bot could configure the window columns in the buy/sell windows to sort 
# TODO: Glob matching for filenames instead of pulling in the orders via a copy/paste option.
# TODO: make a note, remove the ticker from the market.
# TODO: location named Expert conflicts with Export search for wallet order export. Solve that conflict.
# NOTE: Template matching is blazing fast vs OCR, and easy to correct if the interface changes.

# Static templates used for checking the screen for standard elements.
ORDER_TEMPLATE = "templates/order_template.png"
MODIFY_TEMPLATE = "templates/modify_order_template.png"

class Screenshot(object):
    def __init__(self, windowname):
        self.windowname = windowname
        self.screenshot = None
        self.screenshot_bbox = None
        self.inverted_screenshot = None
        self.resize_factor = 4.0

    def find_target_window(self):
        # May want to take in the window name for more flexibility.
        targetwindow = win32gui.FindWindow(None, self.windowname)
        left, top, right, bottom = win32gui.GetWindowRect(targetwindow)
        width = right - left
        height = bottom - top
        self.window_coords = { 
            "left": left,
            "right": right,
            "top": top,
            "bottom": bottom,
            "width": width,
            "height": height,
            "targetwindow": targetwindow
            }

    def load_template(self, template):
        # There are time where we'll use the Screenshot class with a saved image.
        self.screenshot = cv2.imread(template)

    def set_bbox(self, bbox):
        # Sometimes we want to restrict a snapshot to a portion of the target window.
        orig_left = self.window_coords.get("left")
        orig_top = self.window_coords.get("top")
        self.window_coords["left"] += bbox.get("left")
        self.window_coords["right"] = orig_left + bbox.get("right")
        self.window_coords["top"] += bbox.get("top")
        self.window_coords["bottom"] = orig_top + bbox.get("bottom")
        self.window_coords["width"] = self.window_coords.get("right") - self.window_coords.get("left")
        self.window_coords["height"] = self.window_coords.get("bottom") - self.window_coords.get("top")

        self.screenshot_bbox = bbox

    def application_snapshot(self):
        # Various device contexts needed by the Windows GDI
        #targetwindowDC = win32gui.GetWindowDC(targetwindow)
        # For some reason my previous method utilizing windll.user32.PrintWindow no longer works.
        desktopwindow = win32gui.GetDesktopWindow()
        targetwindowDC = win32gui.GetWindowDC(desktopwindow)
        mfcDC = win32ui.CreateDCFromHandle(targetwindowDC)
        saveDC = mfcDC.CreateCompatibleDC()
        # Actually save the targetwindow to a bitmap
        databitmap = win32ui.CreateBitmap()
        databitmap.CreateCompatibleBitmap(
            mfcDC,
            self.window_coords.get("width"),
            self.window_coords.get("height"))

        saveDC.SelectObject(databitmap)
        saveDC.BitBlt(
            (0, 0),
            (self.window_coords.get("width"), self.window_coords.get("height")),
            mfcDC,
            (self.window_coords.get("left"), self.window_coords.get("top")),
            win32con.SRCCOPY)
        #result = windll.user32.PrintWindow(targetwindow, saveDC.GetSafeHdc(), 0)

        # Safety check around the success of PrintWindow could be nice.
        bmpinfo = databitmap.GetInfo()
        bmpstr = databitmap.GetBitmapBits(True)
        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)
        self.screenshot = cv2.cvtColor(numpy.array(im), cv2.COLOR_RGB2BGR)
        # Garbage collect. Looks like the win32 libraries leave handles on the OS.
        win32gui.DeleteObject(databitmap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(self.window_coords.get("targetwindow"), targetwindowDC)

    def save_snapshot(self, image, filename="test.png"):
        cv2.imwrite(filename, image)

    def invert_image(self):
        # Processing necessary to increase the tesseract hit-rate.
        # By default tesseract seems to be best suited to identifying black text on white background.
        kernel = numpy.ones((1,1), numpy.uint8)
        # Grey scale, threshold to black/white, invert for white background.
        grey = cv2.cvtColor(self.screenshot, cv2.COLOR_BGR2GRAY)
        binary = cv2.threshold(grey, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        invert = cv2.bitwise_not(binary)
        self.inverted_screenshot = invert

    def parse_image(self, screen_capture):
        # This is going to return a glob of data in the form of an hocr document.
        # hocr is a form of xhtml that contains each line, the words that make up the line, and the x/y coordinates for each entity.
        return pytesseract.image_to_pdf_or_hocr(Image.fromarray(screen_capture), extension='hocr')

    def resize_image(self, image, resize_factor):
        # Pretty standard sharpen and resize in an attempt to increase Tesseract's hit-rate.
        resized_image = cv2.resize(image, None, fx=resize_factor, fy=resize_factor, interpolation=cv2.INTER_CUBIC)
        return resized_image

    def draw_bounding_box(self, image, bbox):
        # draw_bounding_box exists mostly for visual verification for test purposes.
        offsetfactor = 2
        top_left = (bbox.get("left") - offsetfactor, bbox.get("top") - offsetfactor)
        bottom_right = (bbox.get("right") + offsetfactor, bbox.get("bottom") + offsetfactor)
        color = (0, 255, 0)
        return cv2.rectangle(image, top_left, bottom_right, color, offsetfactor)

    def template_match(self, image, template):
        # Searches the image for a match to the provided template. This is used for finding static screen elements like right-click menus.
        res = cv2.matchTemplate(image, template, cv2.TM_CCORR_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)
        left = max_loc[0]
        top = max_loc[1]
        # Wow.. shape is returning height, width, channels vs the documented width, height, channels.
        right = left + template.shape[1]
        bottom = top + template.shape[0]
        return {"left": left, "top": top, "right": right, "bottom": bottom, "max_confidence_val": max_val}

    def template_check(self, image, template_filename):
        # There are several pre-selected templates that we will use. This simplifies life.
        template = []
        template = cv2.imread(template_filename)
        # Verify template loaded.
        if len(template) > 0:
            return self.template_match(image, template)
        raise Exception("Failed to load template file: %s", template_filename)

    def snapshot_resize_parse(self):
        # This pattern happens a lot, might as well lump this into one call.
        # Returns the hocr result from tesseract processing the created image.
        self.application_snapshot()
        return self.base_resize_parse()

    def base_resize_parse(self):
        # Sometimes self.screenshot is pre-loaded vs snapped.
        self.screenshot_original = copy.copy(self.screenshot)
        self.screenshot = self.resize_image(self.screenshot, self.resize_factor)
        self.invert_image()
        return self.parse_image(self.inverted_screenshot)        

            
class HOCRParser(object):
    def __init__(self):
        print("Init")
        self.soup = None
        self.body = None
        self.ocr_page = None
        self.parsed_lines = None
    
    def make_soup(self, data):
        self.soup = beautifulsoup.BeautifulSoup(data, "html.parser")

    def parse_soup(self):
        # Haven't dealt with html in a very long time, so there's a solid change I'm going to cock this up.
        # body contains all the juicy stuff. Think we'll just sequentially tick through all the tags, and build an object that can be consumed more easily.
        # Primarily need to be able to find certain key-words and get their coordinates.
        self.body = self.soup.body
        # There are certain classes that we care about in the HOCR output as follows:
        # ocr_page: div class that signifies the ocr processed page.
        # ocr_carea: this is a block on the page. Seems to typically involve a few lines in a column.
        # ocr_par: Seems redundant with ocr_carea, seems to equal a few lines in a column.
        # ocr_line: The meat and potatoes, a line of recognized words! Contains an incrementing id, and the title value contains the bbox coords we so desperately seek.
        # ocr_line example: <span class='ocr_line' id='line_1_4' title="bbox 67 109 184 122; baseline 0 -1; x_size 16; x_descenders 4; x_ascenders 4">
        # ocrx_word: each individual word that is recognized gets a unique ID, bbox coords, and a confidence score.
        # ocrx_word example: <span class='ocrx_word' id='word_1_7' title='bbox 162 112 184 121; x_wconf 93'>Eden</span>
        # Search the body for the first(and only) 
        self.ocr_page = self.body.find(name="div", attrs={"class": "ocr_page"})

    def find_word(self, targetword, targetdata):
        # In some cases we want a single word/line from the HOCR data.
        found_words = self.find_words(targetword, targetdata)
        if len(found_words) > 1:
            raise Exception("More than one match for %s", targetword)
        elif len(found_words) == 0:
            raise Exception("Failed to match %s", targetword)
            # Will need a user-defined exception here.
        found_word = found_words[0]
        return found_word

    def find_words(self, targetword, targetdata, ratio=75):
        # Sometimes it is helpful to retrieve all instances of a word.
        #TODO: implement fuzzy matching
        #found_words = [word for word in targetdata if targetword.lower() in word.get("word").lower()]
        # Testing out fuzzywuzzy to help correct for tesseract's mangling of words. Figure a 55% match ratio is a good start.
        # Trying out forcing it all lowercase, as sometimes Tesseract mangles capitalization.
        # Lower helped considerably. Went from a ~75% hit-rate on "OK" to 100% across 20 attempts.
        found_words = [word for word in targetdata if fuzz.ratio(targetword.lower(), word.get("word").lower()) > ratio]
        return found_words

    def find_partial_words(self, targetword, targetdata):
        # When sorting orders we have a semi-sanitized set of inputs, but each line has a LOT of data outside of the target word. For now we're going to set a very high partial_ratio match and see if that works. I'm a little worried about items with modifiers like small/medium/large. Cross that bridge if/when we hit it.
        found_words = [word for word in targetdata if fuzz.partial_ratio(targetword.lower(), word.get("word").lower()) > 95]
        return found_words

    def find_word_bounded_horizontal(self, word, bbox, targetdata):
        # This searches for words that are positioned relatively horizontal. This is primarily used for searching for the Export button that lies in roughly the same vertical column as BUYING/SELLING in the wallet.
        matches = self.find_words(word, targetdata)
        best_overlap = 0
        best_match = None
        for match in matches:
            # Tesseract isn't consistent in regards to how it lumps together words and their respective bboxs.
            # Because of this, we're just checking for horizontal overlap. Unless the user has a very unusual configuration of their windows, the Market data export button will never have any horizontal overlap with the wallet export button.
            overlap = max(0, min(bbox.get("right"), match.get("right")) - max(bbox.get("left"), match.get("left")))
            if overlap > best_overlap:
                best_overlap = overlap
                best_match = match
        return best_match

    def find_words_in_bbox(self, word, bbox, targetdata):
        # Sometimes there will be several instances of a word on the screen.
        # We're typically only interested in a subset of them, perhaps all sell orders for
        # a certain item. We can find those if we have a bbox to search in.
        #matches = self.find_words(word, targetdata)
        #targets = []
        #for match in matches:
        print("COMPLETE")
            
    def process_hocr(self):
        # Most of the time tesseract is not going to match things very well.
        # Because of this we'll assemble all the recognized lines, and then
        # store an array of all the objects. I think we can then process
        # the array of words with the whitelist of interesting words?
        # From our ocr_page, extract all the ocr_carea objects
        areas = self.ocr_page.find_all("div", attrs={"class": "ocr_carea"})
        processed_lines = []
        processed_words = []
        for area in areas:
            lines = area.find_all("span", attrs={"class": "ocr_line"})
            for line in lines:
                composed_word = ""
                words = line.find_all("span", attrs={"class": "ocrx_word"})
                for word in words:
                    composed_word += word.text + " "
                    processed_word = self._process_tag_title(word)
                    processed_word['word'] = word.text
                    processed_words.append(processed_word)
                processed_line = self._process_tag_title(line)
                processed_line['word'] = composed_word
                processed_lines.append(processed_line)
        self.parsed_lines = processed_lines
        self.parsed_words = processed_words

    def clean_results(self, dirtydata):
        # Tesseract sometimes dumps a lot of whitespace bboxs for some reason.
        # Purging these results for now.
        filtered_results = []
        for result in dirtydata:
            if not result.get("word").isspace():
                filtered_results.append(result)
        return filtered_results

    def _process_tag_title(self, tag):
        # Used for dumping the bbox values from a tag's title.
        # Verify title exists
        if tag.has_attr('title'):
            coords = tag.attrs['title'].split()
            # Verify title contains bbox coords
            if coords[0] == "bbox":
                processed_word = {
                    "word": None,
                    "left": int(coords[1]),
                    "right": int(coords[3]),
                    "top": int(coords[2]),
                    "bottom": int(coords[4].strip(";"))
                }
                return processed_word
            raise Exception("Missing bbox coords for %s", tag)
        raise Exception("Missing title attr in %s", tag)

    def rescale_parsed_lines(self, scaleddata, resize_factor):
        # Due to rescaling images to facilitate in Tesseract matching, we need to rescale the resulting bboxs.
        rescaled_lines = []
        for line in scaleddata:
            line['left'] = int(line.get('left')/resize_factor)
            line['right'] = int(line.get('right')/resize_factor)
            line['top'] = int(line.get('top')/resize_factor)
            line['bottom'] = int(line.get('bottom')/resize_factor)
            rescaled_lines.append(line)
        return rescaled_lines

    def hocr_parse_rescale(self, hocr, rescale_factor=4.0):
        # Another consolidation function. Sets the internal parsed_words and parsed_lines based off of the hocr input.
        self.make_soup(hocr)
        self.parse_soup()
        self.process_hocr()
        self.parsed_lines = self.rescale_parsed_lines(self.clean_results(self.parsed_lines), rescale_factor)
        self.parsed_words = self.rescale_parsed_lines(self.clean_results(self.parsed_words), rescale_factor)

class ExportOrderProcessor(object):
    # ExportOrderProcessor simply exports and loads market orders so we can use easily use them.
    def __init__(self, screenshot, mouse, type_processor):
        # ExportOrderProcessor is a base class that should be extended for the wallet and market windows.
        self.screenshot = screenshot
        self.mouse = mouse
        self.typeIDs = type_processor
        self.parser = None
        self.market_orders = None
        self.order_directory = None

    def locate_keywords(self):
        # This needs to be implemented for each sub-class. This should in theory just locate Export/Export to File for wallet/market exports respectively.
        return None

    def locate_export_keywords(self, parser):
        # This needs to be implemented for each sub-class. Both alerts might need slightly different keywords, but in short they need something that is unique at the top and bottom of the alert.
        # Need to return two bboxes, one for the top of the alert, and one for the OK button.
        return None, None

    def export_wallet_orders(self):
        # This will go about exporting the character's market orders via the wallet export feature.
        # Index the screen
        self.scrape_window()
        # At the bottom of the wallet is an export button. Press it.
        self.click_on_wallet_export(self.locate_keywords())

        # This will create a prompt that we will interact with to ascertain the location of the market log.
        # First detect the alert on the screen.
        # TODO: save the location of the prompt for re-use. If we run this frequently, we can snapshot
        # a limited part of the screen just to verify the OK and Personal Market Export components are visible.
        personal_market_export, ok_button = self.map_export_alert()

        # With the key words located, we can interact with the prompt and pull out the alert.
        # There are some assumptions made from looking at this prompt. text_bbox is roughly in the middle
        # of the prompt, which is sufficient for our requirements.
        separation = ok_button.get("top") - personal_market_export.get("bottom")
        mid_point = personal_market_export.get("bottom") + int(separation)/2
        top = mid_point - int(separation/6)
        bottom = mid_point + int(separation/6)
        text_bbox = {
            "left": personal_market_export.get("left"),
            "right": personal_market_export.get("right"),
            "top": int(top),
            "bottom": int(bottom)
        }
        # text_bbox comes from a full-screen screenshot, no offset required.
        text_coords = self.mouse.calculate_point_from_bbox(text_bbox)
        print(f"Alert coords at:{text_coords}")

        # With the generated coordinates, copy the alert text
        alert_text = self.copy_export_alert(text_coords)
        print(f"Alert text: {alert_text}")

        # The alert text contains a filename and directory where the market order was exported to.
        order_directory, order_file = self.process_market_order_alert(alert_text)

        # Import the orders
        market_orders = self.read_market_orders(order_directory, order_file)

        # Orders by default use a typeID for each item, convert that to something useful.
        better_orders = self.translate_typeIDs(market_orders)

        # Sort the orders into buy and sell groups.
        buy_orders, sell_orders = self.sort_buy_sell_orders(better_orders)

        # Store the important data for future use.
        self.market_orders = better_orders
        self.buy_orders = buy_orders
        self.sell_orders = sell_orders
        self.order_directory = order_directory

    def scrape_window(self):
        # self.screenshot's dimensions should be set via Screenshot.set_bbox. GameManipulator is capable of generating constrained bbox's for the wallet and market windows.
        hocr = self.screenshot.snapshot_resize_parse()

        parser = HOCRParser()
        parser.hocr_parse_rescale(hocr, self.screenshot.resize_factor)
        self.parser = parser

    def click_on_wallet_export(self, export_button):
        # Let's click that button.
        export_coords = self.mouse.calculate_point_from_bbox(export_button, offset_bbox=self.screenshot.screenshot_bbox)
        self.mouse.click(export_coords, click_type="left")

    def snapshot_export_alert(self):
        #TODO: use cv2 template matching to watch the screen
        # https://opencv-python-tutroals.readthedocs.io/en/latest/py_tutorials/py_imgproc/py_template_matching/py_template_matching.html
        screenshot = Screenshot(self.screenshot.windowname)
        screenshot.find_target_window()
        hocr = screenshot.snapshot_resize_parse()

        # Parse the result.
        hocrparser = HOCRParser()
        hocrparser.hocr_parse_rescale(hocr, screenshot.resize_factor)
        return hocrparser

    def map_export_alert(self):
        # This alert is a temporary prompt. We will create temporary snapshot/parser objects vs juggling the main screen objects.
        time.sleep(1) # Uhh.. we might be moving faster than the EVE client can render stuff.
        # Scrape the screen for the export window.
        parser = self.snapshot_export_alert()

        # We will key off of two keywords, OK, and Personal Market Export.
        # These two keywords have a very high hit-rate with tesseract.
        personal_market_export, ok_button = self.locate_export_keywords(parser)
        return personal_market_export, ok_button

    def copy_export_alert(self, text_coords):
        # Provided a bbox of the rough location of the alert, copy the text to the clipboard and return it.
        # Double click on the alert.
        self.mouse.click(text_coords, click_type="left")
        self.mouse.click(text_coords, click_type="left")
        # ctrl + a
        self.mouse.pressAndHold("ctrl")
        self.mouse.press("a")
        # ctrl + c
        self.mouse.press("c")
        self.mouse.release("ctrl")
        # enter
        self.mouse.press("enter")

        # Get data from the clipboard.
        text = self.get_clipboard()
        return text

    def process_market_order_alert(self, text):
        # Some hard-coded values here as the message seems pretty standard.
        file_index = text.index("file")
        in_index = text.index("in")
        order_file = text[file_index+5:in_index-1]
        directory_index = text.index("directory")
        order_directory = text[directory_index+10:]
        return order_directory, order_file

    def read_market_orders(self, order_directory, order_file):
        # Read in the market order file and return the contents.
        #columns = "orderID,typeID,charID,charName,regionID,regionName,stationID,stationName,range,bid,price,volEntered,volRemaining,issueDate,orderState,minVolume,accountID,duration,isCorp,solarSystemID,solarSystemName,escrow".split(",")
        orders = []
        with open(order_directory + order_file) as csv_file:
            csv_reader = csv.DictReader(csv_file, delimiter=",")
            for row in csv_reader:
                orders.append(row)
        return orders

    def translate_typeIDs(self, orders):
        # Translate the typeID to the English name for each order.
        for order in orders:
            item_name = self.typeIDs.get(order.get("typeID"))
            if item_name:
                order["itemName"] = item_name
            else:
                raise Exception(f"Failed to convert typeID: {order.get('typeID')}. Is this a new typeID?") 
        return orders

    def sort_buy_sell_orders(self, orders):
        # Separate buy and sell orders.
        buy_orders = []
        sell_orders = []
        for order in orders:
            if order.get('bid') == 'True':
                buy_orders.append(order)
            else:
                sell_orders.append(order)
        return buy_orders, sell_orders

    def get_clipboard(self):
        win32clipboard.OpenClipboard()
        text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return text

class ExportMarketOrderProcessor(ExportOrderProcessor):
    def locate_keywords(self):
        # This needs to be implemented for each sub-class. This should locate the Export/Export to File buttons for wallet/market exports respectively.
        return self.parser.find_word("Export", self.parser.parsed_words)

    def locate_export_keywords(self, parser):
        # This needs to be implemented for each sub-class. Both alerts might need slightly different keywords, but in short they need something that is unique at the top and bottom of the alert.
        # Need to return two bboxes, one for the top of the alert, and one for the OK button.
        market = parser.find_word("Market Export", parser.parsed_lines)
        ok = parser.find_word("OK", parser.parsed_words)
        return market, ok

class ExportWalletOrderProcessor(ExportOrderProcessor):
    def locate_keywords(self):
        # This needs to be implemented for each sub-class. This should in theory just locate Export/Export to File for wallet/market exports respectively.
        return self.parser.find_word("Export", self.parser.parsed_words)

    def locate_export_keywords(self, parser):
        # This needs to be implemented for each sub-class. Both alerts might need slightly different keywords, but in short they need something that is unique at the top and bottom of the alert.
        # Need to return two bboxes, one for the top of the alert, and one for the OK button.
        market = parser.find_word("Personal", parser.parsed_words)
        ok = parser.find_word("OK", parser.parsed_words)
        return market, ok

class GameManipulator(object):
    # GameManipulator consumes the output of ExportOrderProcessor and interacts with the game to manipulate orders.
    # It will work with both the screen and HOCR results, which are based on a screenshot starting at 0,0
    def __init__(self, main_window, main_parser, mouse, orderbook):
        self.main_window = main_window
        self.main_parser = main_parser
        self.mouse = mouse
        self.orderbook = orderbook
        self.market_orders = None
        self.buy_snapshot = None
        self.sell_snapshot = None

    def locate_keywords(self):
        # We depend on a couple of special keywords on the primary screen. Knowing their location helps considerably with interacting with it.
        # Everything above BUYING is the selling window.
        buying_bbox = self.main_parser.find_word("BUYING", self.main_parser.parsed_words)
        selling_bbox = self.main_parser.find_word("SELLING", self.main_parser.parsed_words)
        # The Regional Market window should be the only other window open. Typically the wallet
        # and market windows are open side-by-side. We'll determine which is on which side in another function.
        market_bbox = self.main_parser.find_words("Regional", self.main_parser.parsed_words, 90)[0]
        # The wallet window has the word Wallet stacked. Vertically. Either option works.
        wallet_bbox = self.main_parser.find_words("Wallet", self.main_parser.parsed_words)[0]
        sellers_bbox = self.main_parser.find_word("Sellers", self.main_parser.parsed_words)
        self.market_bbox = market_bbox
        self.buying_bbox = buying_bbox
        self.selling_bbox = selling_bbox
        self.wallet_bbox = wallet_bbox
        self.sellers_bbox = sellers_bbox

    def estimate_window_layout(self):
        # Make a few educated assumptions so we can restrict our searches to the correct portion of the screen.
        if self.wallet_bbox.get("left") < self.market_bbox.get("left"):
            # The wallet is positioned to the left of the of the market.
            order_right_bound = self.market_bbox.get("left")
            order_left_bound = self.wallet_bbox.get("right")
            market_item_left_bound = self.sellers_bbox.get("right")
            market_item_right_bound = self.main_window.window_coords.get("width")

        elif self.wallet_bbox.get("left") > self.market_bbox.get("left"):
            # The wallet is positioned to the right of the market.
            order_right_bound = self.main_window.window_coords.get("width")
            order_left_bound = self.wallet_bbox.get("right")
            market_item_left_bound = self.sellers_bbox.get("right")
            market_item_right_bound = self.wallet_bbox.get("left")
        
        else:
            raise Exception("Cannot determine the window orientation.")
        
        # Typically the selling window is above the buying window.
        if self.selling_bbox.get("top") < self.buying_bbox.get("top"):
            order_sell_top_bound = self.selling_bbox.get("top")
            order_sell_bottom_bound = self.buying_bbox.get("top")
            order_buy_top_bound = self.buying_bbox.get("top")
            order_buy_bottom_bound = self.main_window.window_coords.get("height")

            market_item_top_bound = self.market_bbox.get("top")
            market_item_bottom_bound = self.sellers_bbox.get("top")

        elif self.selling_bbox.get("top") > self.buying_bbox.get("top"):
            # Just in case the buy window is on top
            order_sell_top_bound = self.selling_bbox.get("top")
            order_sell_bottom_bound = self.main_window.window_coords.get("height")

            order_buy_top_bound = self.buying_bbox.get("top")
            order_buy_bottom_bound = self.selling_bbox.get("top")

            market_item_top_bound = self.market_bbox.get("top")
            # TODO: There are two keywords, Sellers and Buyers, that we can use to constrain this.
            market_item_bottom_bound = self.sellers_bbox.get("top")

        else:
            raise Exception("Cannot determine buy/sell orientation.")
            
        self.sell_coords = {
            "left": order_left_bound,
            "right": order_right_bound,
            "top": order_sell_top_bound,
            "bottom": order_sell_bottom_bound
        }
        self.buy_coords = {
            "left": order_left_bound,
            "right": order_right_bound,
            "top": order_buy_top_bound,
            "bottom": order_buy_bottom_bound
        }
        self.market_item_coords = {
            "left": market_item_left_bound,
            "right": market_item_right_bound,
            "top": market_item_top_bound,
            "bottom": market_item_bottom_bound
        }
        # Entire market window
        self.market_coords = {
            "left": order_right_bound,
            "right": market_item_right_bound,
            "top": market_item_top_bound,
            "bottom": self.main_window.window_coords.get("height")
        }
        # Entire wallet window
        self.wallet_coords = {
            "left": self.wallet_bbox.get("left"),
            "right": order_right_bound,
            "top": self.wallet_bbox.get("top"),
            "bottom": self.main_window.window_coords.get("height")
        }

    def create_snapshots(self):
        # Tesseract's processing time is a function of the size of the image. Restricting our search to our buy/sell regions will dramatically speed up the process.
        buy_snapshot = Screenshot(self.main_window.windowname)
        sell_snapshot = Screenshot(self.main_window.windowname)
        buy_snapshot.find_target_window()
        sell_snapshot.find_target_window()
        # Manipulate the snapshot window_coords to reflect the on-screen reduced target window per each buy/sell coord. Outside of the snapshot and mouse objects, the entire project will work on the hocr/screenshot coords originating at 0,0.
        buy_snapshot.set_bbox(self.buy_coords)
        sell_snapshot.set_bbox(self.sell_coords)
        self.buy_snapshot = buy_snapshot
        self.sell_snapshot = sell_snapshot

    def prepare(self):
        # This combines several steps that are needed for GameManipulator to actually perform its actions.
        # This primarily exists as the way to init GameManipulator in use vs testing each function in dev/test.
        self.locate_keywords()
        self.estimate_window_layout()
        self.create_snapshots()
        self.order_template_parser = self.process_template(ORDER_TEMPLATE)
        self.modify_template_parser = self.process_template(MODIFY_TEMPLATE)

    def parse_snapshot(self, snapshot, order_type="buy"):
        # Given a buy/sell snapshot and order list, return the coordinates of the visible orders.
        hocr = snapshot.snapshot_resize_parse()
        parser = HOCRParser()
        parser.hocr_parse_rescale(hocr, snapshot.resize_factor)

        if order_type == "buy":
            coords = self.buy_coords
            orders = self.market_orders.get("buy")
        elif order_type == "sell":
            coords = self.sell_coords
            orders = self.market_orders.get("sell")

        visible_orders = []
        for order in orders:
            # This will need better logic at some point to handle partial matches and stuff. This works for testing purposes. Moving on!
            copy_order = copy.deepcopy(order)
            matched_lines = parser.find_partial_words(order.get("itemName"), parser.parsed_lines)
            if len(matched_lines) == 1:
                copy_order["bbox"] = matched_lines[0]
                copy_order["found"] = True
            elif len(matched_lines) > 1:
                copy_order["bbox"] = matched_lines
                copy_order["found"] = False
            else:
                copy_order["bbox"] = None    
                copy_order["found"] = False
            visible_orders.append(copy_order)

        return visible_orders

    def load_order_market(self, order, snapshot):
        # Double clicking on the item will cause the market to load.
        mouse_coords = self.mouse.calculate_point_from_bbox(order.get("bbox"), offset_bbox=snapshot.screenshot_bbox)
        self.mouse.click(mouse_coords, double=True)
        # Check that the correct item loaded in the market.
        for i in range(5):
            try:
                print("Iteraction: %s" % i)
                result = self.check_market_for_item(order.get("itemName"))
                if result:
                    print("Item %s loaded" % order.get("itemName"))
                    break
                time.sleep(0.50)
            except Exception:
                print("At loop iteraction %s and item %s not detected" % (i, order.get("itemName")))
            if i == 4:
                raise Exception("Item %s never loaded." % order.get("itemName"))

    def check_market_for_item(self, item):
        # Using Tesseract, check to see if an item has been loaded.
        snapshot = Screenshot(self.main_window.windowname)
        snapshot.find_target_window()
        snapshot.set_bbox(self.market_item_coords)
        hocr = snapshot.snapshot_resize_parse()
        parser = HOCRParser()
        parser.hocr_parse_rescale(hocr, snapshot.resize_factor)
        result = parser.find_words(item, parser.parsed_lines)
        return result

    def process_template(self, template):
        # Several templates will be used for quickly identifying elements on the screen.
        # Processing a template once for multiple uses will save considerable processing time. This returns a HOCRParser object that can be used to find the location of the associated template.
        tmp_screenshot = Screenshot(None)
        tmp_screenshot.load_template(template)
        hocr = tmp_screenshot.base_resize_parse()
        parser = HOCRParser()
        parser.hocr_parse_rescale(hocr, tmp_screenshot.resize_factor)
        return parser

    def modify_order(self, order, price_change, snapshot):
        # Edit an order, changing the price to the provided value.
        # Click on the order to bring up the interaction prompt.
        mouse_coords = self.mouse.calculate_point_from_bbox(order.get("bbox"), offset_bbox=snapshot.screenshot_bbox)
        self.mouse.click(mouse_coords, click_type="right")
        self.watch_for_template(self.main_window, ORDER_TEMPLATE)
        # NOTE: This can be sped up by using portion of the screen around where we think the prompt will appear.
        # I chose not to use this proposed solution to simplify this template matching section for now.
        # For reference, processing time went from 240ms to 70ms on a 7700k when using 2x full-screen template match vs 2x ~1/4 screen matches.
        self.main_window.application_snapshot()
        order_template_result = self.main_window.template_check(self.main_window.screenshot, ORDER_TEMPLATE)
        # Locate where the Modify keyword is within the template.
        modify_coords = self.order_template_parser.find_word("Modify", self.order_template_parser.parsed_words)
        # Click on modify from the right click order menu.
        modify_mouse_coords = self.mouse.calculate_point_from_bbox(modify_coords, offset_bbox=order_template_result)
        self.mouse.click(modify_mouse_coords)

        self.watch_for_template(self.main_window, MODIFY_TEMPLATE)
        # Locate the modify order window on the main window.
        self.main_window.application_snapshot()
        modify_template_result = self.main_window.template_check(self.main_window.screenshot, MODIFY_TEMPLATE)
        # TODO: Actually change the price.


        # Locate where the Ok button is within the template.
        ok_button = self.modify_template_parser.find_word("Cancel", self.modify_template_parser.parsed_words)
        ok_button_mouse_coords = self.mouse.calculate_point_from_bbox(ok_button, offset_bbox=modify_template_result)
        self.mouse.click(ok_button_mouse_coords)

    def watch_for_template(self, snapshot, template, confidence_val=0.50):
        # Watch for 5x100ms, waiting for the provided template to present itself on the screen.
        for i in range(5):
            snapshot.application_snapshot()
            template_match_attempt = snapshot.template_check(snapshot.screenshot, template)
            print(template_match_attempt)
            if template_match_attempt.get("max_confidence_val") > confidence_val:
                return template_match_attempt
            else:
                time.sleep(0.100)
        raise Exception("Failed to detect template: %s" % template)


if __name__ == "__main__":
    print("I solemnly swear that I'm up to no good.")