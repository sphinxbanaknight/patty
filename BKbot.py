import discord
import random
import os
import json
import gspread
import pprint
import pygsheets

from oauth2client.service_account import ServiceAccountCredentials
from discord.ext import commands, tasks

#scope = ['https://spreadsheets.google.com/feeds',
         #'https://www.googleapis.com/auth/drive']
#creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
gc = pygsheets.authorize(service_file ='cogs/client_secret.json')

sheet = gc.open('Copy of BK ROSTER').sheet1

client = commands.Bot(command_prefix="/")

# pp = pprint.PrettyPrinter()

# result = sheet.col_values(4)

# pp.pprint(result)

list_ab = ['ab', 'arch bishop', 'arch', 'bishop', 'priest', 'healer', 'buffer']
list_gene = ['gene', 'genetic']
list_mins = ['mins', 'minstrel']
list_wand = ['wanderer', 'wand', 'wandy']
list_rg = ['rg', 'royal guard', 'devo', ]
list_gx = ['gx', 'guillotine cross', 'glt. cross']
list_rk = ['rk', 'rune knight', 'db']
list_sc = ['sc', 'shadow chaser']
list_obo = ['obo', 'oboro', 'ninja']
list_rebel = ['rebel', 'reb', 'rebellion']
list_doram = ['cat', 'doram']
list_sorc = ['sorc', 'sorceror']
list_sura = ['sura', 'shura', 'asura', 'ashura']
list_wl = ['wl', 'warlock', 'tetra', 'crimson rock', 'cr']
list_mech = ['mech', 'mechanic', 'mado']
list_ranger = ['ranger', 'range']

'''@client.command()
async def identify(ctx):
    print(ctx.author.name)'''


# ctx.author.name to get the username of the one who activated the command

def next_available_row_guild():
    cols = pygsheets.datarange.DataRange(start = 'B3', end = 'E47')
    return max([cell.row for cell in cols if cell.value]) + 1
    # str_list = list(filter(None, sheet.col_values(column)))
    # return str(len(str_list) + 1)

def next_available_row_attend():
    cols = pygsheets.datarange.DataRange(start='G3', end='J47')
    return max([cell.row for cell in cols if cell.value]) + 1

@client.command()
async def sorted(ctx):
    drange = pygsheets.datarange.DataRange(start = 'B3', end = 'E47')
    drange.sort(basecolumnindex=3, sortorder="ASCENDING")
    # sheet.sort((8, 'desc'), range='G3:J50')
    await ctx.send('Sorted.')


client.run('Njk0Mjg1NTE0NTc2MjMyNTE5.XoJZ0g.BSzbWRetqBta_RNNzNYDG6m73FE')