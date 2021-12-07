import requests, json, time
import xlwings as xl
from xlwings.utils import rgb_to_int

# xlwings documentation: https://docs.xlwings.org/en/stable/index.html

apiKey = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjExMjQyNTE2OCwidWlkIjoxOTY3MDc2MCwiaWFkIjoiMjAyMS0wNi0wM1QyMjowMzozMy4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6NDMzODY1LCJyZ24iOiJ1c2UxIn0.bPO7_2J7dPYaD5-Xiu0mT6_eqg9iphKcdpxz14mjsX0"
apiUrl = "https://api.monday.com/v2"
headers = {"Authorization": apiKey}


real_board_id = '1269455955'  # Change this value if the board needs to change


def datapull():  # This function is to pull data from excel models

    wb = xl.Book.caller()  # This creates an instance of the workbook where this function was called from
    ws = wb.sheets['STR']
    rooms = ws.range(39, 27).value
    year = str(ws.range(39, 31).value)
    year = year[0:4]  # STR data is yyyy/mm this portion of the code just takes the year
    year = int(year)
    name = " - " + ws.range(39, 24).value  # Takes the city and state name for the name to in the form that we use
    address = ws.range(37, 34).value + " " + ws.range(39, 24).value + " " + ws.range(39, 25).value

    ws = wb.sheets['DataEntry_PIP_S&U']
    pulse_id = ws.range(10, 5).value  # Pulls the pulse id number if there is one
    name = ws.range(7, 5).value + name  # Completes the naming convention

    ws = wb.sheets['UPREIT_HighLevel']

    offer = ws.range(9, 2).value
    equityMultiple = ws.range(47, 12).value
    pipCost = ws.range(31, 16).value
    noi = ws.range(19, 5).value
    acqCost = ws.range(1, 16).value
    sources = ws.range(9, 12).value
    debt = ws.range(7, 2).value

    datadict = {"Rooms": rooms,
                "Year": year,
                "Offer": offer,
                "Equity Multiple": equityMultiple,
                "PIP Cost": pipCost,
                "NOI": noi,
                "Acquisition Cost": acqCost,
                "Sources & Uses": sources,
                "Debt": debt,
                "Name": name,
                "Address": address,
                "pulse_id": pulse_id}
    return datadict


def newmodel(modeldata):  # This function is called by Post() if there is no pulse id in the datapull dict

    mutation = r'mutation {create_item (board_id: 1829871734, item_name: "%s") {id}}' % modeldata['Name']
    data = {'query': mutation}
    r = requests.post(url=apiUrl, json=data, headers=headers)  # These 3 lines create a new pulse

    response_data = r.json()
    item_id = response_data['data']['create_item']['id']  # These 2 lines take the pulse id the newly created pulse

    wb = xl.Book.caller()
    ws = wb.sheets['DataEntry_PIP_S&U']
    ws.range(10, 5).value = item_id
    ws.range(10, 5).api.Font.Color = rgb_to_int((0, 0, 0))  # These write the pulse id into the model that called the program

    offer = r'\"numeric0\": \"%d\"' % modeldata['Offer']
    emx = r'\"numeric05\": \"%d\"' % modeldata['Equity Multiple']
    pipcost = r'\"numeric1\": \"%d\"' % modeldata['PIP Cost']
    noi = r'\"numeric69\": \"%d\"' % modeldata['NOI']
    acqcost = r'\"numeric9\": \"%d\"' % modeldata['Acqusition Cost']
    su = r'\"numeric67\": \"%d\"' % modeldata['Sources & Uses']
    debt = r'\"numeric11\": \"%d\"' % modeldata['Debt']
    rooms = r'\"numeric\": \"%s\"' % modeldata['Rooms']
    year = r'\"text\": \"%s\"' % modeldata['Year']  # This section transforms the datadict into a GraphQL query.
    # Might update the above section but I'm feeling lazy right now. Update with column names should columns change.
    # Use this for help with column type references https://api.developer.monday.com/docs/guide-to-changing-column-data
    mutation = 'mutation {change_multiple_column_values (board_id: 1829871734, item_id: %s , ' \
               'column_values: "{%s, %s, %s, %s, %s, %s, %s, %s, %s}") {id}}' \
               % (item_id, rooms, year, offer, emx, pipcost, noi, acqcost, su, debt)
    data = {'query': mutation}
    r = requests.post(url=apiUrl, json=data, headers=headers)  # Updates the new pulse with the rest of the data in datadict

    return


def updateitem(modeldata):

    offer = r'\"numeric0\": \"%d\"' % modeldata['Offer']
    emx = r'\"numeric05\": \"%d\"' % modeldata['Equity Multiple']
    pipcost = r'\"numeric1\": \"%d\"' % modeldata['PIP Cost']
    noi = r'\"numeric69\": \"%d\"' % modeldata['NOI']
    acqcost = r'\"numeric9\": \"%d\"' % modeldata['Acqusition Cost']
    su = r'\"numeric67\": \"%d\"' % modeldata['Sources & Uses']
    debt = r'\"numeric11\": \"%d\"' % modeldata['Debt']

    mutation = 'mutation {change_multiple_column_values (board_id: 1829871734, item_id: %s, ' \
               'column_values: "{%s, %s, %s, %s, %s, %s, %s}") {id}}' \
               % (modeldata['pulse_id'], offer, emx, pipcost, noi, acqcost, su, debt)
    data = {'query': mutation}
    r = requests.post(url=apiUrl, json=data, headers=headers)
    print(r.json())
    return


def post():
    try:
        data = datapull()

        if data['pulse_id'] is None:
            newmodel(data)
        else:
            updateitem(data)
    except:
        print("Something is not working correctly")


post()
