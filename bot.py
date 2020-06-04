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

takte = gc.open('BK ROSTER')
rostersheet = takte.worksheet('WoE Roster')
silk2 = takte.worksheet('WoE Roster 2') 
silk4 = takte.worksheet('WoE Roster 4')
fullofsheet = takte.worksheet('Full IGNs')

################ Channel, Server, and User IDs ###########################
sphinx_id = 108381986166431744
servers = [401186250335322113, 691130488483741756]
sk_server = 401186250335322113
bk_server = 691130488483741756
botinit_id = [401212001239564288, 691205255664500757]
sk_bot = 401212001239564288
bk_bot = 691205255664500757
bk_ann = 695801936095740024 #BK #announcement
authorized_id = [108381986166431744, 127778244383473665, 130885439308431361, 352073289155346442, 143743232658898944]
#Asi = 127778244383473665
#Eba = 130885439308431361
#1may2020 Marte = 437617764484513802
#giveme5minutes haclime = 127795095121559552
#crackedvoice marvs = 437618826897522690
#red = 352073289155346442
#jia = 143743232658898944


################ Cell placements ################
guild_range = "B3:E50"
roster_range = "G3:I50"
matk_range = "L3:M14"
p1role_range = "P3:P14"
atk_range = "L17:M28"
p2role_range = "P17:P28"
p3_range = "L32:M43"
p3role_range = "P32:P43"

################ Parameters ################
# These flags are status trackers to avoid duplicate runs in the timed events
isarchived = False # archiving
isreminded_wed = False # reminder wed
isreminded_sat = False # reminder sat

isremindenabled = True # configuration - turn on/off auto-reminder


################ Feedbacks ################
feedback_automsg = '`[Auto-generated message]` '
feedback_noangrypingplz = 'to avoid <:AngryPing:703193629489102888> from me, please register your attendance before coming Saturday 12PM GMT+8 by `/att y/n, y/n` in <#691205255664500757>. Thank you!'


prefix = ["/"]
description = "A bot for sheet+discord linking/automation."
client = commands.Bot(command_prefix = prefix, description = description)

client.remove_command('help')

def next_available_row(sheet, column):
    cols = sheet.range(3, column, 1000, column)
    return max([cell.row for cell in cols if cell.value]) + 1
 
def pinger(ctx):
    attlist = [item for item in rostersheet.col_values(7) if item and item != 'IGN' and item != 'Next WOE:']
    ignlist = [item for item in rostersheet.col_values(3) if item and item != 'IGN' and item != 'READ THE NOTES AT [README]']
    idlist = [item for item in fullofsheet.col_values(3) if item and item != 'UNIQUE:' and item != 'Discord Tag' and item != 'READ THE NOTES AT [README]']
    row = 3
    dsctag = []
    dscid = []
    
    for ign in ignlist:
        for att in attlist:
            if ign == att:
                ign = ""
                gottem = 1
                break
        if gottem == 0:
            try:
                dsctag.append(rostersheet.cell(row, 2).value)
            except Exception as e:
                print(f'Exception caught at dsctag: {e}')
        else:
            gottem = 0
        row += 1
        
    row = 4
    for idd in idlist:
        for dsc in dsctag:
            #print(f'{idd} to {dsc} has idd of {fullofsheet.cell(row, 2).value}')
            if idd == dsc:
                try:
                    dscid.append(fullofsheet.cell(row, 2).value)
                except Exceptions as e:
                    print(f'Exception caught at dsctag: {e}')
                break
        row += 1
    
    return dscid 

@client.event
async def on_ready():
    global isarchived
    global isreminded_wed
    global isreminded_sat
    
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
        elif channel.id == bk_ann:
            botinitbkann = channel
            break

    print('Bot is online.')
    await client.change_presence(status=discord.Status.dnd, activity=discord.Game('Getting scolded by Jia'))


    shit = gc.open('BK ROSTER')
    try:
        wsheet = shit.worksheet('WoE Roster Archive')
    except gspread.exceptions.WorksheetNotFound:
        await botinitsk.send(f'Could not find 5th sheet in our GSheets, creating one now.')
        await botinitbk.send(f'Could not find 5th sheet in our GSheets, creating one now.')
        spreadsheet = gc.open('BK ROSTER')
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
        if not isarchived and ( ph_time_formated == "00:00:00:Monday" or ph_time_formated == "00:00:00:Sunday" ):
            await botinitsk.send('```Automatically cleared the roster! Please use /att y/n again to register your attendance.```')
            await botinitsk.send('```An archive of the latest roster was saved in WoE Roster Archive Spreadsheet.```')
            await botinitbk.send('```Automatically cleared the roster! Please use /att y/n again to register your attendance.```')
            await botinitbk.send('```An archive of the latest roster was saved in WoE Roster Archive Spreadsheet.```')
            try:
                next_row = next_available_row(wsheet, 1)
            except ValueError as e:
                print(f'next_row returned {e}')
                next_row = 1
            copy_list2 = silk2.range("B4:D51")
            copy_list4 = silk4.range("B4:D51")
            paste_list = rostersheet.range(next_row, 1, next_row + 45, 3)
            count = 0
            newformat = "%B %Y"
            ph_time = pytz.timezone('Asia/Manila')
            ph_time_unformated = datetime.now(ph_time)
            ph_time_new_formated = ph_time_unformated.strftime(newformat)
            if ph_time_formated == "00:00:00:Sunday":
                for copy in copy_list2:
                    data_pasted.append(copy.value)
            elif ph_time_formated == "00:00:00:Monday":
                for copy in copy_list4:
                    data_pasted.append(copy.value)
            for paste in paste_list:
                if count == 0:
                    if ph_time_formated == "00:00:00:Sunday":
                        d = ph_time_unformated.strftime("%d")
                        d = int(d)
                        d -= 1
                        d = str(d)
                        paste.value = f'{d} {ph_time_new_formated} PM SAT WOE'
                        cell_list = silk2.range("B4:D50")
                        for cell in cell_list:
                            cell.value = ""
                        silk2.update_cells(cell_list, value_input_option='USER_ENTERED')
                    elif ph_time_formated == "00:00:00:Monday":
                        d = ph_time_unformated.strftime("%d")
                        d = int(d)
                        d -= 1
                        d = str(d)
                        paste.value = f'{d} {ph_time_new_formated} PM SUN WOE'
                        cell_list = silk4.range("B4:D50")
                        for cell in cell_list:
                            cell.value = ""
                        silk4.update_cells(cell_list, value_input_option='USER_ENTERED')
                elif count == 1:
                    paste.value = ""
                else:
                    try:
                        paste.value = data_pasted[count - 2]
                    except IndexError:
                        break
                count += 1
            wsheet.update_cells(paste_list, value_input_option='USER_ENTERED')

            cell_list = rostersheet.range(roster_range)

            for cell in cell_list:
                cell.value = ""

            rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')
            isarchived = True
            continue
        # Timed event status reset
        elif ph_time_formated == "00:05:00:Monday" or ph_time_formated == "00:05:00:Sunday":
            isarchived = False
            isreminded_wed = False
            isreminded_sat = False
            await botinitsk.send(f'`[Timed event status reset] isarchived={isarchived} isreminded_wed={isreminded_wed} isreminded_sat={isreminded_sat}`')
            continue
        # Timed event [auto-reminder]: a soft reminder message into #announcement. Remove on next event
        elif isremindenabled and not isreminded_wed and ph_time_formated == "22:00:00:Wednesday":
            try:
                att_igns = [item for item in rostersheet.col_values(7) if item and item != 'IGN' and item != 'Next WOE:']
                nratt = len(att_igns)
                msg_wed = await botinitbkann.send(f'''{feedback_automsg}
Hi all,

Currently we have {nratt} members who have registered their attendance, great job!
For those who haven't: {feedback_noangrypingplz}''')
                isreminded_wed = True
            except Exception as e:
                await botinitsk.send(f'Error: `{e}`')
            continue
        # Timed event [auto-reminder]: @mention per player who enlisted but not yet confirmed attendance
        #jytest elif isremindenabled and not isreminded_sat and ph_time_formated == "12:00:00:Saturday":
        elif isremindenabled and not isreminded_sat and ph_time_formated == "00:47:00:Thursday": #jytest #pattest
            try:
                await msg_wed.delete() #jytest todo envelop in try-except, because msg_wed may not be found
                ping_tags = []
                att_igns = [item for item in rostersheet.col_values(7) if item and item != 'IGN' and item != 'Next WOE:']
                next_row = 3
                cell_list = rostersheet.range("C3:C50")
                for cell in cell_list:
                    if cell.value not in att_igns:
                        tag = rostersheet.cell(next_row, 2) # discord tag at column 2
                        ping_tags.append(tag.value)
                    next_row += 1
                
                for discordtag in ping_tags:
                    if discordtag == "Takudan": #jytest
                        await botinitsk.send(f'{feedback_automsg} Hi @{discordtag}, you have not registered your attendance yet. <:peeposad:702156649992945674> Next time, {feedback_noangrypingplz}')
                    #await botinitbk.send(f'{feedback_automsg} Hi @{discordtag}, you have not registered your attendance yet. <:peeposad:702156649992945674> Next time, {feedback_noangrypingplz}')
                isreminded_sat = True
            except Exception as e:
                await botinitsk.send(f'Error: `{e}`')
            continue
            


@client.event
async def on_member_join(member):
    try:
        role = discord.utils.get(member.guild.roles, id=703643328406880287)
    except AttributeError as e:
        print(f'autorole returned as {e}')
        role = discord.utils.get(member.guild.roles, name="new recruit")
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

# toggle isremindenabled
@client.command()
async def togglereminder(ctx):
    global isremindenabled
    channel = ctx.message.channel
    commander = ctx.author
    if channel.id in botinit_id:
        if commander.id in authorized_id:
            try:
                isremindenabled = not isremindenabled
            except Exception as e:
                await ctx.send(e)
            await ctx.send(f'`Auto-reminder Enabled = {isremindenabled}`')
        else:
            await ctx.send(f'*Nice try pleb.*')
    else:
        await ctx.send(f'Wrong channel! Please use #bot.')

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