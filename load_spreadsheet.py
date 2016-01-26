import argparse
import re
import string
import xlrd

def getPlayerName(text):
    pattern_two_comma = re.compile("""(.+)\s*,.+,""")
    pattern_comma = re.compile("""(.+)\s*,.+""")
    pattern_dash = re.compile("""(.+) -.+-""")
    pattern_strip_number = re.compile("""\d\d\.\S(.+)""")
    if not text:
        return ""

    strip_number_matched = pattern_strip_number.search(text.strip())
    if not strip_number_matched == None:
        text = strip_number_matched.group(1)

    two_comma_matched = pattern_two_comma.search(text)
    if not two_comma_matched == None:
        return two_comma_matched.group(1).strip()

    comma_matched = pattern_comma.search(text)
    if not comma_matched == None:
        return comma_matched.group(1).strip()

    dash_matched = pattern_dash.search(text)
    if not dash_matched == None:
        return dash_matched.group(1).strip()

    return text.strip()


parser= argparse.ArgumentParser(description='Prospect List Parser.')

parser.add_argument('-i', '--input', help='Weekly Excel Spreadsheet', required=True)

args = parser.parse_args()

wb = xlrd.open_workbook(args.input)
from xlrd.sheet import ctype_text

#print wb

sheet_names = wb.sheet_names()
#print('Sheet Names', sheet_names)

sheet = wb.sheet_by_name('Top 100s')

prospects = {}
rankings = {}

rankers = []

for col_idx in range(2,sheet.ncols):
    name = ""
    name1_obj = sheet.cell(0, col_idx)
    if name1_obj.value:
        name += name1_obj.value

    name2_obj = sheet.cell(1, col_idx)
    if name2_obj.value:
        if name:
            name += " "
        name += name2_obj.value


    #print name
    rankers.append(name)

    top100 = {}
    rank = 0
    for row_idx in range(3,103):
        rank += 1
        player_obj = sheet.cell(row_idx, col_idx)
        if player_obj.value:
            player_name = getPlayerName(player_obj.value)

            top100[player_name] = rank
            prospects[player_name] = 0

    rankings[name] = top100

out = ","
for key in rankers:
    out += key + ","
print out

keys = prospects.keys()
keys.sort()
#print len(keys)
for key in keys:
    out = "\"" + key + "\","
    for ranker in rankers:
        if key in rankings[ranker]:
            out += str(rankings[ranker][key]) + ","
            prospects[key] = prospects[key] + (150 - rankings[ranker][key])
        else:
            out += ","
            prospects[key] = prospects[key] + 1

    print out

#for key in keys:
#    print "\"" + key + "\"," + str(prospects[key])
