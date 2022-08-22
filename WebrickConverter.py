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

colourNameDictionary = {
    '1': 'White',
    '49': 'Very Light Gray',
    '99': 'Very Light Bluish Gray',
    '86': 'Light Bluish Gray',
    '9': 'Light Gray',
    '10': 'Dark Gray',
    '85': 'Dark Bluish Gray',
    '11': 'Black',
    '59': 'Dark Red',
    '5': 'Red',
    '220': 'Coral',
    '231': 'Dark Salmon',
    '25': 'Salmon',
    '26': 'Light Salmon',
    '58': 'Sand Red',
    '120': 'Dark Brown',
    '8': 'Brown',
    '91': 'Light Brown',
    '240': 'Medium Brown',
    '88': 'Reddish Brown',
    '106': 'Fabuland Brown',
    '69': 'Dark Tan',
    '241': 'Medium Tan',
    '2': 'Tan',
    '90': 'Light Nougat',
    '28': 'Nougat',
    '150': 'Medium Nougat',
    '225': 'Dark Nougat',
    '160': 'Fabuland Orange',
    '29': 'Earth Orange',
    '68': 'Dark Orange',
    '27': 'Rust',
    '165': 'Neon Orange',
    '4': 'Orange',
    '31': 'Medium Orange',
    '110': 'Bright Light Orange',
    '32': 'Light Orange',
    '96': 'Very Light Orange',
    '161': 'Dark Yellow',
    '3': 'Yellow',
    '33': 'Light Yellow',
    '103': 'Bright Light Yellow',
    '236': 'Neon Yellow',
    '166': 'Neon Green',
    '35': 'Light Lime',
    '158': 'Yellowish Green',
    '76': 'Medium Lime',
    '248': 'Fabuland Lime',
    '34': 'Lime',
    '242': 'Dark Olive Green',
    '155': 'Olive Green',
    '80': 'Dark Green',
    '6': 'Green',
    '36': 'Bright Green',
    '37': 'Medium Green',
    '38': 'Light Green',
    '48': 'Sand Green',
    '39': 'Dark Turquoise',
    '40': 'Light Turquoise',
    '41': 'Aqua',
    '152': 'Light Aqua',
    '63': 'Dark Blue',
    '7': 'Blue',
    '153': 'Dark Azure',
    '247': 'Little Robots Blue',
    '72': 'Maersk Blue',
    '156': 'Medium Azure',
    '87': 'Sky Blue',
    '42': 'Medium Blue',
    '105': 'Bright Light Blue',
    '62': 'Light Blue',
    '55': 'Sand Blue',
    '109': 'Dark Blue-Violet',
    '43': 'Violet',
    '97': 'Blue-Violet',
    '245': 'Lilac',
    '73': 'Medium Violet',
    '246': 'Light Lilac',
    '44': 'Light Violet',
    '89': 'Dark Purple',
    '24': 'Purple',
    '93': 'Light Purple',
    '157': 'Medium Lavender',
    '154': 'Lavender',
    '227': 'Clikits Lavender',
    '54': 'Sand Purple',
    '71': 'Magenta',
    '47': 'Dark Pink',
    '94': 'Medium Dark Pink',
    '104': 'Bright Pink',
    '23': 'Pink',
    '56': 'Light Pink',
    '12': 'Trans-Clear',
    '13': 'Trans-Black',
    '17': 'Trans-Red',
    '18': 'Trans-Neon Orange',
    '98': 'Trans-Orange',
    '164': 'Trans-Light Orange',
    '121': 'Trans-Neon Yellow',
    '19': 'Trans-Yellow',
    '16': 'Trans-Neon Green',
    '108': 'Trans-Bright Green',
    '221': 'Trans-Light Green',
    '226': 'Trans-Light Bright Green',
    '20': 'Trans-Green',
    '14': 'Trans-Dark Blue',
    '74': 'Trans-Medium Blue',
    '15': 'Trans-Light Blue',
    '113': 'Trans-Aqua',
    '114': 'Trans-Light Purple',
    '234': 'Trans-Medium Purple',
    '51': 'Trans-Purple',
    '50': 'Trans-Dark Pink',
    '107': 'Trans-Pink',
    '21': 'Chrome Gold',
    '22': 'Chrome Silver',
    '57': 'Chrome Antique Brass',
    '122': 'Chrome Black',
    '52': 'Chrome Blue',
    '64': 'Chrome Green',
    '82': 'Chrome Pink',
    '83': 'Pearl White',
    '119': 'Pearl Very Light Gray',
    '66': 'Pearl Light Gray',
    '95': 'Flat Silver',
    '239': 'Bionicle Silver',
    '77': 'Pearl Dark Gray',
    '244': 'Pearl Black',
    '61': 'Pearl Light Gold',
    '115': 'Pearl Gold',
    '235': 'Reddish Gold',
    '238': 'Bionicle Gold',
    '81': 'Flat Dark Gold',
    '249': 'Reddish Copper',
    '84': 'Copper',
    '237': 'Bionicle Copper',
    '78': 'Pearl Sand Blue',
    '243': 'Pearl Sand Purple',
    '228': 'Satin Trans-Clear',
    '229': 'Satin Trans-Black',
    '233': 'Satin Trans-Bright Green',
    '223': 'Satin Trans-Light Blue',
    '232': 'Satin Trans-Dark Blue',
    '230': 'Satin Trans-Purple',
    '224': 'Satin Trans-Dark Pink',
    '67': 'Metallic Silver',
    '70': 'Metallic Green',
    '65': 'Metallic Gold',
    '250': 'Metallic Copper',
    '60': 'Milky White',
    '159': 'Glow In Dark White',
    '46': 'Glow In Dark Opaque',
    '118': 'Glow In Dark Trans',
    '101': 'Glitter Trans-Clear',
    '222': 'Glitter Trans-Orange',
    '163': 'Glitter Trans-Neon Green',
    '162': 'Glitter Trans-Light Blue',
    '102': 'Glitter Trans-Purple',
    '100': 'Glitter Trans-Dark Pink',
    '111': 'Speckle Black-Silver',
    '151': 'Speckle Black-Gold',
    '116': 'Speckle Black-Copper',
    '117': 'Speckle DBGray-Silver',
    '123': 'Mx White',
    '124': 'Mx Light Bluish Gray',
    '125': 'Mx Light Gray',
    '126': 'Mx Charcoal Gray',
    '127': 'Mx Tile Gray',
    '128': 'Mx Black',
    '131': 'Mx Tile Brown',
    '134': 'Mx Terracotta',
    '132': 'Mx Brown',
    '133': 'Mx Buff',
    '129': 'Mx Red',
    '130': 'Mx Pink Red',
    '135': 'Mx Orange',
    '136': 'Mx Light Orange',
    '137': 'Mx Light Yellow',
    '138': 'Mx Ochre Yellow',
    '139': 'Mx Lemon',
    '141': 'Mx Pastel Green',
    '140': 'Mx Olive Green',
    '142': 'Mx Aqua Green',
    '146': 'Mx Teal Blue',
    '143': 'Mx Tile Blue',
    '144': 'Mx Medium Blue',
    '145': 'Mx Pastel Blue',
    '147': 'Mx Violet',
    '148': 'Mx Pink',
    '149': 'Mx Clear',
    '210': 'Mx Foil Dark Gray',
    '211': 'Mx Foil Light Gray',
    '212': 'Mx Foil Dark Green',
    '213': 'Mx Foil Light Green',
    '214': 'Mx Foil Dark Blue',
    '215': 'Mx Foil Light Blue',
    '216': 'Mx Foil Violet',
    '217': 'Mx Foil Red',
    '218': 'Mx Foil Yellow',
    '219': 'Mx Foil Orange',
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
        itemQuantity = ""
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
                itemQuantity = cell.value
        if (item.id == None):
            continue
        if isNotSingleItem:
            print("Item is in " + itemQuantity + " quantity: id: " + str(item.id) + ", colour: " + colourNameDictionary[str(colourDictionary[str(item.colour)])] + ", quantity: " + str(int(item.quantity)))
        items.append(item)
        rowNumber += 1
    return items

def map_bricklink_alternatives(result):
    if ("BrickLink" not in result["external_ids"]):
        print('Item ' + result["part_num"] + " does not have a bricklink part ID assigned in the rebrickable database, it is assumed that the Bricklink part has the same id")
        return {
            result["part_num"]: result["part_num"]
        }
    return {
        result["part_num"]: result["external_ids"]["BrickLink"][0]
    }

def list_to_dictionary(list):
    return {k:v for element in list for k,v in element.items()}

def query_rebrickable_api(api):
    response = requests.get(f"{api}")
    if response.status_code == 200:
        responseJson = response.json()
        result = list(map(map_bricklink_alternatives, responseJson['results']))
        if responseJson['next'] is not None:
            result = result + query_rebrickable_api(responseJson['next'])
        return result
    else:
        print(f"There's a {response.status_code} error with your request")

def get_bricklink_alternatives(items):
    itemIds = map(lambda x: x.id, items)
    itemIds = list(set(itemIds))
    itemQuery = ""
    for id in itemIds:
        itemQuery += str(id) + "%2C"
    itemQuery = itemQuery[:-3]
    return list_to_dictionary(query_rebrickable_api(f"https://rebrickable.com/api/v3/lego/parts/?key=51544b88076aa8087e8b9536fc61bac3&part_nums={itemQuery}"))

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
    
    for item in modifiedNonExistentItems:
        item.id = item.id[:-1]
    return modifiedNonExistentItems

def get_unique_items(brickLinkItems):
    uniqueBrickLinkItems = []
    for item in brickLinkItems:
        existingItem = next((x for x in uniqueBrickLinkItems if (x.id == item.id and x.colour == item.colour)), None)
        if existingItem == None:
            uniqueBrickLinkItems.append(item)
            continue
        existingItem.quantity += item.quantity
    return uniqueBrickLinkItems

def get_best_possible_id(ids):
    for id in ids:
        if "pr" not in id and "pb" not in id and "stk" not in id:
            return id
    return ids[0]

def map_to_list_of_bricklink_alternatives(result):
    if ("BrickLink" not in result["external_ids"]):
        return result["part_num"]
    return result["external_ids"]["BrickLink"][0]

def search_rebrickable_api(api):
    response = requests.get(f"{api}")
    if response.status_code == 200:
        responseJson = response.json()
        if responseJson['count'] == 0 or "BrickLink" not in responseJson['results'][0]["external_ids"]:
            return None
        possibleIds = list(map(map_to_list_of_bricklink_alternatives, responseJson['results']))
        return get_best_possible_id(possibleIds)
    else:
        print(f"There's a {response.status_code} error with your request")

def try_searching_items(brickLinkItems, itemsForSearching):
    searchedNonExistentItems = []

    for item in itemsForSearching:
        brickLinkId = search_rebrickable_api(f"https://rebrickable.com/api/v3/lego/parts/?key=51544b88076aa8087e8b9536fc61bac3&search={item.id}")
        if brickLinkId == None or str(item.colour) not in colourDictionary:
            searchedNonExistentItems.append(item)
            continue
        print(item.id + " did not have an exact match and the best guess was " + brickLinkId)
        brickLinkItem = BrickLinkItem(str(brickLinkId), colourDictionary[str(item.colour)], item.quantity)
        brickLinkItems.append(brickLinkItem)

    return searchedNonExistentItems

def print_bricklink_inventory(excelFileName):
    items = get_items(excelFileName)

    nonExistentItems = []
    brickLinkItems = []
    add_bricklink_items(items, brickLinkItems, nonExistentItems)
    nonExistentItems = try_modified_items(brickLinkItems, nonExistentItems)
    nonExistentItems = try_searching_items(brickLinkItems, nonExistentItems)

    brickLinkItems = get_unique_items(brickLinkItems)
    print()
    for item in nonExistentItems:
        if str(item.colour) not in colourDictionary:
            print("Cannot find color for item: id: " + str(item.id) + " colour: " + str(item.colour) + " quantity: " + str(int(item.quantity)))
            continue
        print("Cannot find item: id: " + str(item.id) + ", colour: " + colourNameDictionary[str(colourDictionary[str(item.colour)])] + ", quantity: " + str(int(item.quantity)))

    print()
    print("<INVENTORY>")
    for item in brickLinkItems:
        print("<ITEM><ITEMTYPE>P</ITEMTYPE><ITEMID>"+str(item.id)+"</ITEMID><COLOR>"+str(item.colour)+"</COLOR><MINQTY>"+str(int(item.quantity))+"</MINQTY></ITEM>")
    print("</INVENTORY>")

print_bricklink_inventory('202003298_order_items.xlsx')