import requests, bs4, re, ezsheets, os, pyinputplus as pyip
sheets = []
styles = """
    <style>
        body {
            background-color: #031103;
            color: #fabbfa;
            text-align: center;
            font-family: Arial;
        }
        table {
            table-layout: auto;
            border-collapse: collapse;
            width: 100%;
        }
        table td {
            border: 1px solid #ccc;
            max-width: 400px;
        }
    </style>
"""
def makeSheet(s):
    spreadsheet = ezsheets.Spreadsheet(s[3:].partition("/")[0])
    sheets = spreadsheet.sheets
    sheet = sheets[0]
    output = styles + "<table>"
    for row in sheet.getRows():
        row = list(filter(None, row))
        output += "<tr>"
        for item in row:
            item = item.strip()
            output += "<td>"
            output += item
            output += "</td>"
        output += "</tr>"
    output += "</table>"
    file = open('gui.html', mode='w', encoding="utf-8")
    file.write("<title>hi</title>" + output)
    file.close()
    os.startfile('gui.html')
                
def getSheets():
    res = requests.get('https://www.pandabuysheets.com/')
    res.raise_for_status()
    pandaPage = bs4.BeautifulSoup(res.text, 'html.parser')
    elems = pandaPage.select('.w-dyn-item')
    print("Select a spreadsheet:")
    for i in elems:
        a = re.search(r'/d/(.+?)+O', i.getText())
        gSheet = a[0][:-1]
        gReq = requests.get("https://docs.google.com/spreadsheets" + gSheet)
        if gReq.status_code != 200:
            # print(gSheet + "is unavailable")
            print("err")
        else:
            sheets.append(gSheet)
    response = pyip.inputMenu(sheets, numbered = True)
    makeSheet(response)
getSheets()