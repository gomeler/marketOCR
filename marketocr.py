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

import screeninteraction

# python-Levenshtein is not necessary but helpful for fuzzywuzzy
# Requires c++ build tools -> https://visualstudio.microsoft.com/visual-cpp-build-tools/
# Grab Build Tools for Visual Studio 2019


# Blackscreen issue
# https://github.com/python-pillow/Pillow/issues/1547
# For some reason windll.user32.PrintWindow is failing to dump the image to the bitmap.

#TODO: Perhaps a custom class is needed for scaled images, so I can have the image and the resize_factor linked.
# TODO: cache the location of on-screen elements and perform restricted tesseract checks vs scanning the full screen.

class Screenshot(object):
    def __init__(self, windowname):
        self.windowname = windowname
        self.screenshot = None
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

    def set_window_foreground(self, window):
        # For some reason SetForegroundWindow sometimes fails.
        # Workaround -> https://stackoverflow.com/questions/14295337/win32gui-setactivewindow-error-the-specified-procedure-could-not-be-found
        #shell = win32com.client.Dispatch("WScript.Shell")
        #shell.SendKeys('%')
        #win32gui.SetForegroundWindow(window)
        # SetForegroundWindow is an asynchronous via the windows API.
        # It is possible that we might snapshot before the operation has completed.
        # 50ms seems like enough time from testing. 
        time.sleep(0.05)

    def application_snapshot(self):
        # Identify the target, and grab it's dimensions.
        self.find_target_window()
        # Bring the target to the foreground.
        self.set_window_foreground(self.window_coords.get("targetwindow"))

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

    def export_order_and_capture(self):
        # Programmatically find the Export button in the wallet, export the orders, 
        # capture the screen, and figure out where on-disk the files were dumped.
        # This might be a case for storing the EVE log cache location in a config file.
        print("test")


    def resize_image(self, image, resize_factor):
        # Pretty standard sharpen and resize in an attempt to increase Tesseract's hit-rate.
        resized_image = cv2.resize(image, None, fx=resize_factor, fy=resize_factor, interpolation=cv2.INTER_CUBIC)
        return resized_image

    def draw_bounding_box(self, image, bbox):
        # draw_bounding_box exists mostly for visual verification for test purposes.
        offsetfactor = 5
        top_left = (bbox.get("left") - offsetfactor, bbox.get("top") - offsetfactor)
        top_right = (bbox.get("right") + offsetfactor, bbox.get("bottom") + offsetfactor)
        color = (0, 255, 0)
        return cv2.rectangle(image, top_left, top_right, color, offsetfactor)
            
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

    def find_eden(self):
        self.eden = self.ocr_page.find_all("span", string="Eden")

    def find_word(self, targetword, targetdata):
        # In some cases we want a single word/line from the HOCR data.
        found_words = self.find_words(targetword, targetdata)
        if len(found_words) > 1:
            raise Exception(f"More than one match for {targetword}")
        elif len(found_words) == 0:
            raise Exception("Failed to match %s", targetword)
            # Will need a user-defined exception here.
        found_word = found_words[0]
        return found_word

    def find_words(self, targetword, targetdata):
        # Sometimes it is helpful to retrieve all instances of a word.
        #TODO: implement fuzzy matching
        #found_words = [word for word in targetdata if targetword.lower() in word.get("word").lower()]
        # Testing out fuzzywuzzy to help correct for tesseract's mangling of words. Figure a 55% match ratio is a good start.
        # Trying out forcing it all lowercase, as sometimes Tesseract mangles capitalization.
        # Lower helped considerably. Went from a ~75% hit-rate on "OK" to 100% across 20 attempts.
        found_words = [word for word in targetdata if fuzz.ratio(targetword.lower(), word.get("word").lower()) > 75]
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

class ExportOrderProcessor(object):
    # ExportOrderProcessor simply exports and loads market orders so we can use easily use them.
    def __init__(self, window_name, main_screenshot, main_parser, mouse, type_processor):
        # ExportOrderProcessor will re-use the same screen capture and HOCRParser for most operations.
        # Temporary copies will be used for handling prompts.
        self.window_name = window_name
        self.main_screenshot = main_screenshot
        self.main_parser = main_parser
        self.mouse = mouse
        self.hocr_data = None
        self.typeIDs = type_processor
        self.market_orders = None
        self.order_directory = None

    def export_wallet_orders(self):
        # This will go about exporting the character's market orders via the wallet export feature.
        self.mouse.set_focus((20, 5))
        # At the bottom of the wallet is an export button. Press it.
        self.click_on_wallet_export()

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
        text_coords = self.mouse.calculate_point_from_bbox(text_bbox)

        # With the generated coordinates, copy the alert text
        alert_text = self.copy_export_alert(text_coords)

        # The alert text contains a filename and directory where the market order was exported to.
        order_directory, order_file = self.process_market_order_alert(alert_text)

        # Import the orders
        market_orders = self.read_market_orders(order_directory, order_file)

        # Orders by default use a typeID for each item, convert that to something useful.
        better_orders = self.translate_typeIDs(market_orders)

        # Store the important data for future use.
        self.market_orders = better_orders
        self.order_directory = order_directory

    def click_on_wallet_export(self):
        # An important assumption is made here, screenshot and hocrparser contain data from a wallet screen scrape.
        # Luckily the export button is typically bounded left/right by the word BUYING within the wallet.
        # I've been lucky so far in that tesseract has yet to fail to 100% identify BUYING and Export.
        buying = self.main_parser.find_word("BUYING", self.main_parser.parsed_words)
        export = self.main_parser.find_word_bounded_horizontal("Export", buying, self.main_parser.parsed_words)
        # Until it proves otherwise, I will assume that "Export" resides within the left/right values of BUYING, but several hundreds of pixels lower. There's a solid chance that the market order Export button will also be visible, so restrict our search to results inside of this range.

        # Let's click that button.
        export_coords = self.mouse.calculate_point_from_bbox(export)
        self.mouse.click(export_coords, click_type="left")

    def map_export_alert(self):
        # This alert is a temporary prompt. We will create temporary snapshot/parser objects vs juggling the main screen objects.
        time.sleep(1) # Uhh.. we might be moving faster than the EVE client can render stuff.
        #TODO: use cv2 template matching to watch the screen
        # https://opencv-python-tutroals.readthedocs.io/en/latest/py_tutorials/py_imgproc/py_template_matching/py_template_matching.html
        screenshot = Screenshot(self.window_name)
        screenshot.application_snapshot()
        # This has to be resized for tesseract to pick up the OK button.
        screenshot.screenshot = screenshot.resize_image(screenshot.screenshot, screenshot.resize_factor)
        screenshot.invert_image()
        hocr = screenshot.parse_image(screenshot.inverted_screenshot)

        # Parse the result.
        hocrparser = HOCRParser()
        hocrparser.make_soup(hocr)
        hocrparser.parse_soup()
        hocrparser.process_hocr()
        hocrparser.parsed_lines = hocrparser.clean_results(hocrparser.parsed_lines)
        hocrparser.parsed_words = hocrparser.clean_results(hocrparser.parsed_words)
        hocrparser.parsed_lines = hocrparser.rescale_parsed_lines(hocrparser.parsed_lines, screenshot.resize_factor)
        hocrparser.parsed_words = hocrparser.rescale_parsed_lines(hocrparser.parsed_words, screenshot.resize_factor)

        # We will key off of two keywords, OK, and Personal Market Export.
        # These two keywords have a very high hit-rate with tesseract.
        personal_market_export = hocrparser.find_word("Personal", hocrparser.parsed_words)
        ok_button = hocrparser.find_word("OK", hocrparser.parsed_words)
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

    def get_clipboard(self):
        win32clipboard.OpenClipboard()
        text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return text

if __name__ == "__main__":
    print("I solemnly swear that I'm up to no good.")