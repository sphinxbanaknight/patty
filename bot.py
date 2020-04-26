import discord
import random
import os
import json
import gspread
import pprint
from oauth2client import file as oauth_file, client, tools
from apiclient.discovery import build
from httplib2 import Http
import time
import datetime
import pytz
import asyncio


from pytz import timezone
from datetime import datetime, timedelta

from oauth2client.service_account import ServiceAccountCredentials
from discord.ext import commands, tasks

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

basedir = os.path.abspath(os.path.dirname(__file__))
data_json = basedir+'/cogs/client_secret.json'

creds = ServiceAccountCredentials.from_json_keyfile_name(data_json, scope)
gc = gspread.authorize(creds)

sheet = gc.open('Copy of BK ROSTER').sheet1

################ Channel, Server, and User IDs ###########################
sphinx_id = 108381986166431744
servers = [401186250335322113, 691130488483741756]
sk_server = 401186250335322113
bk_server = 691130488483741756
botinit_id = [401212001239564288, 691205255664500757]
sk_bot = 401212001239564288
bk_bot = 691205255664500757
authorized_id = [108381986166431744, 127778244383473665, 130885439308431361, 437617764484513802, 127795095121559552, 437618826897522690, 352073289155346442]
#Asi = 127778244383473665
#Eba = 130885439308431361
#Marte = 437617764484513802
#haclime = 127795095121559552
#marvs = 437618826897522690
#red = 352073289155346442


################ Cell placements ###########################
guild_range = "B3:E50"
roster_range = "G3:J50"
matk_range = "L3:M14"
p1role_range = "P3:P14"
atk_range = "L17:M28"
p2role_range = "P17:P28"
p3_range = "L32:M43"
p3role_range = "P32:P43"


prefix = ["/"]
description = "A bot for sheet+discord linking/automation."
client = commands.Bot(command_prefix = prefix, description = description)

client.remove_command('help')

def next_available_row(sheet, column):
    cols = sheet.range(3, column, 1000, column)
    return max([cell.row for cell in cols if cell.value]) + 1

@client.event
async def on_ready():
    for server in client.guilds:
        if server.id == sk_server:
            sphinx = server
            continue
        if server.id == bk_server:
            burger = server
            continue

    for channel in sphinx.channels:
        if channel.id == sk_bot:
            botinitsk = channel
            break
    for channel in burger.channels:
        if channel.id == bk_bot:
            botinitbk = channel
            break

    print('Bot is online.')
    await client.change_presence(status=discord.Status.dnd, activity=discord.Game('Getting scolded by Jia'))


    shit = gc.open('Copy of BK ROSTER')
    try:
        wsheet = shit.worksheet('WoE Roster Archive')
    except gspread.exceptions.WorksheetNotFound:
        await botinitsk.send(f'Could not find 5th sheet in our GSheets, creating one now.')
        await botinitbk.send(f'Could not find 5th sheet in our GSheets, creating one now.')
        spreadsheet = gc.open('Copy of BK ROSTER')
        wsheet = spreadsheet.add_worksheet(title='WoE Roster Archive', rows = 1000, cols = 10)
        kekerino = wsheet.range("A1:J1000")
        for kek in kekerino:
            kek.value = ''
        wsheet.update_cells(kekerino, value_input_option='USER_ENTERED')

    print("Automated Clear Roster Begins!")
    format = "%H:%M:%S:%A"

    ph_time = pytz.timezone('Asia/Manila')
    ph_time_unformated = datetime.now(ph_time)
    ph_time_formated = ph_time_unformated.strftime(format)


    data_pasted = [""]
    while True:
        ph_time = pytz.timezone('Asia/Manila')
        ph_time_unformated = datetime.now(ph_time)
        ph_time_formated = ph_time_unformated.strftime(format)
        await asyncio.sleep(1)
        if ph_time_formated == "00:00:00:Monday" or ph_time_formated == "00:00:00:Sunday":
            await botinitsk.send('```Automatically cleared the roster! Please use /att y/n again to register your attendance.```')
            await botinitsk.send('```An archive of the latest roster was saved in WoE Roster Archive Spreadsheet.```')
            await botinitbk.send('```Automatically cleared the roster! Please use /att y/n again to register your attendance.```')
            await botinitbk.send('```An archive of the latest roster was saved in WoE Roster Archive Spreadsheet.```')
            try:
                next_row = next_available_row(wsheet, 1)
            except ValueError as e:
                print(f'next_row returned {e}')
                next_row = 1
            copy_list = sheet.range("G3:I50")
            paste_list = sheet.range(next_row, 1, next_row + 45, 3)
            count = 0
            newformat = "%B %Y"
            ph_time = pytz.timezone('Asia/Manila')
            ph_time_unformated = datetime.now(ph_time)
            ph_time_new_formated = ph_time_unformated.strftime(newformat)

            for copy in copy_list:
                data_pasted.append(copy.value)
            for paste in paste_list:
                if count == 0:
                    if ph_time_formated == "00:00:00:Sunday":
                        d = ph_time_unformated.strftime("%d")
                        d = int(d)
                        d -= 1
                        d = str(d)
                        paste.value = f'{d} {ph_time_new_formated} PM SAT WOE'
                    elif ph_time_formated == "00:00:00:Monday":
                        d = ph_time_unformated.strftime("%d")
                        d = int(d)
                        d -= 1
                        d = str(d)
                        paste.value = f'{d} {ph_time_new_formated} PM SUN WOE'
                elif count == 1:
                    paste.value = ""
                else:
                    try:
                        paste.value = data_pasted[count - 2]
                    except IndexError:
                        break
                count += 1
            wsheet.update_cells(paste_list, value_input_option='USER_ENTERED')

            cell_list = sheet.range(roster_range)

            for cell in cell_list:
                cell.value = ""

            sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
            continue

@client.event
async def on_member_join(member):
    role = discord.utils.get(member.guild.roles, id="703643328406880287")
    await member.add_roles(role)

@client.command()
async def load(ctx, extension):
    client.load_extension(f'cogs.{extension}')
    await ctx.send(f'Cog: {extension}.py loaded')

@client.command()
async def unload(ctx, extension):
    client.unload_extension(f'cogs.{extension}')
    await ctx.send(f'Cog: {extension}.py unloaded')

@client.command()
async def reload(ctx, extension):
    client.unload_extension(f'cogs.{extension}')
    client.load_extension(f'cogs.{extension}')
    await ctx.send(f'Cog: {extension}.py reloaded')


for filename in os.listdir('./cogs'):
    if filename.endswith('.py'):
        client.load_extension(f'cogs.{filename[:-3]}')

@client.event
async def on_command_error(ctx, error):
    if isinstance(error, commands.MissingRequiredArgument):
        await ctx.send('Please pass in all required arguments.')
    if isinstance(error, commands.CommandNotFound):
        await ctx.send('Invalid command used.')

client.run(os.environ['token'], reconnect = True)