import openpyxl
import requests
from pathlib import Path

colourDictionary = {
    '1': 1,
    '5': 2,
    '21': 5,
    '23': 7,
    '24': 3,
    '26': 11,
    '28': 6,
    '37': 36,
    '38': 68,
    '40': 12,
    '41': 17,
    '42': 15,
    '43': 14,
    '44': 19,
    '48': 20,
    '49': 16,
    '102': 42,
    '106': 4,
    '107': 37,
    '111': 13,
    '113': 50,
    '119': 34,
    '124': 71,
    '126': 51,
    '135': 55,
    '138': 69,
    '140': 63,
    '141': 80,
    '151': 48,
    '154': 59,
    '182': 98,
    '191': 110,
    '192': 88,
    '194': 86,
    '199': 85,
    '212': 105,
    '221': 47,
    '222': 104,
    '226': 103,
    '268': 89,
    '283': 90,
    '294': 118,
    '297': 115,
    '308': 120,
    '309': 22,
    '310': 21,
    '311': 108,
    '312': 150,
    '315': 95,
    '316': 77,
    '321': 153,
    '322': 156,
    '323': 152,
    '324': 157,
    '325': 154,
    '326': 158,
    '330': 155,
    '999': 242,
    '037': 61,
    '091': 60,
    '085': 29
}

class Item:
  def __init__(self, id, colour, quantity):
    self.id = id
    self.colour = colour
    self.quantity = quantity


class BrickLinkItem:
  def __init__(self, id, colour, quantity):
    self.id = id
    self.colour = colour
    self.quantity = quantity

def get_items(excelFileName):
    items = []
    xlsx_file = Path('', excelFileName)
    wb_obj = openpyxl.load_workbook(xlsx_file) 
    sheet = wb_obj.active
    rowNumber = 0
    for row in sheet.iter_rows():
        if rowNumber < 2:
            rowNumber += 1
            continue
        columnNumber = 0
        item = Item(None, None, None)
        isNotSingleItem = False
        for cell in row:
            if (columnNumber == 0):
                columnNumber += 1
                continue
            if (cell.value == None):
                continue
            if columnNumber == 1:
                item.id = cell.value
            if columnNumber == 2:
                item.colour = cell.value
            if columnNumber == 5:
                item.quantity = cell.value
            columnNumber += 1
            if columnNumber == 10 and cell.value != "1 piece":
                isNotSingleItem = True
        if (item.id == None or isNotSingleItem):
            continue
        items.append(item)
        rowNumber += 1
    return items

def map_bricklink_alternatives(result):
    if ("BrickLink" not in result["external_ids"]):
        return {
        result["part_num"]: None
    }
    return {
        result["part_num"]: result["external_ids"]["BrickLink"][0]
    }

def list_to_dictionary(list):
    return {k:v for element in list for k,v in element.items()}

def query_rebrickable_api(api):
        response = requests.get(f"{api}")
        if response.status_code == 200:
            responseJson = response.json()['results']
            return list_to_dictionary(list(map(map_bricklink_alternatives, responseJson)))
        else:
            print(f"There's a {response.status_code} error with your request")

def get_bricklink_alternatives(items):
    itemIds = ""
    for item in items:
        itemIds += str(item.id) + "%2C"
    itemIds = itemIds[:-3]
    return query_rebrickable_api(f"https://rebrickable.com/api/v3/lego/parts/?key=51544b88076aa8087e8b9536fc61bac3&part_nums={itemIds}")

def add_bricklink_items(items, brickLinkItems, nonExistentItems):
    brickLinkAlternatives = get_bricklink_alternatives(items)
    for item in items:
        if str(item.colour) not in colourDictionary or str(item.id) not in brickLinkAlternatives or brickLinkAlternatives[str(item.id)] == None:
            nonExistentItems.append(item)
            continue
        brickLinkItem = BrickLinkItem(brickLinkAlternatives[str(item.id)], colourDictionary[str(item.colour)], item.quantity)
        brickLinkItems.append(brickLinkItem)

def try_modified_items(brickLinkItems, nonExistentItems):
    modifiedItemsToTry = nonExistentItems.copy()
    modifiedNonExistentItems = []
    for item in modifiedItemsToTry:
        item.id = str(item.id) + "b"

    add_bricklink_items(modifiedItemsToTry, brickLinkItems, modifiedNonExistentItems)
    
    for item in modifiedItemsToTry:
        item.id = item.id[:-1]
    return modifiedItemsToTry

def get_unique_items(brickLinkItems):
    uniqueBrickLinkItems = []
    for item in brickLinkItems:
        existingItem = next((x for x in uniqueBrickLinkItems if (x.id == item.id and x.colour == item.colour)), None)
        if existingItem == None:
            uniqueBrickLinkItems.append(item)
            continue
        existingItem.quantity += item.quantity
    return uniqueBrickLinkItems

def print_bricklink_inventory(excelFileName):
    items = get_items(excelFileName)

    nonExistentItems = []
    brickLinkItems = []
    add_bricklink_items(items, brickLinkItems, nonExistentItems)
    nonExistentItems = try_modified_items(brickLinkItems, nonExistentItems)

    brickLinkItems = get_unique_items(brickLinkItems)

    for item in nonExistentItems:
        print("Cannot find item: id: " + str(item.id) + " colour: " + str(item.colour) + " quantity: " + str(item.quantity))

    print("<INVENTORY>")
    for item in brickLinkItems:
        print("<ITEM><ITEMTYPE>P</ITEMTYPE><ITEMID>"+str(item.id)+"</ITEMID><COLOR>"+str(item.colour)+"</COLOR><MINQTY>"+str(int(item.quantity))+"</MINQTY></ITEM>")
    print("</INVENTORY>")

print_bricklink_inventory('202003298_order_items.xlsx')
