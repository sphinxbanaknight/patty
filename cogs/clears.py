import discord
import random
import os
import json
import gspread
import pprint
#import models
from oauth2client import file as oauth_file, client, tools
from apiclient.discovery import build
from httplib2 import Http
import time

from oauth2client.service_account import ServiceAccountCredentials
from discord.ext import commands, tasks

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

basedir = os.path.abspath(os.path.dirname(__file__))
data_json = basedir+'/client_secret.json'

creds = ServiceAccountCredentials.from_json_keyfile_name(data_json, scope)
gc = gspread.authorize(creds)

sheet = gc.open('Copy of BK ROSTER').sheet1
shite = gc.open('Copy of BK ROSTER')
celesheet = shite.worksheet('Celery Preferences')

################ Channel, Server, and User IDs ###########################
sphinx_id = 108381986166431744
servers = [401186250335322113, 691130488483741756]
#sphinxk = 401186250335322113
#BK = 691130488483741756
botinit_id = [401212001239564288, 691205255664500757]
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
############### Roles #######################################
list_ab = ['ab', 'arch bishop', 'arch', 'bishop', 'priest', 'healer', 'buffer']
list_gene = ['gene', 'genetic']
list_mins = ['mins', 'minstrel' ]
list_wand = ['wanderer', 'wand', 'wandy']
list_rg = ['rg', 'royal guard', 'devo',]
list_gx = ['gx', 'guillotine cross', 'glt. cross']
list_rk = ['rk', 'rune knight', 'db']
list_sc = ['sc', 'shadow chaser']
list_obo = ['obo', 'oboro', 'ninja']
list_rebel = ['rebel', 'reb', 'rebellion']
list_doram = ['cat', 'doram']
list_sorc = ['sorc', 'sorcerer']
list_sura = ['sura', 'shura', 'asura', 'ashura']
list_wl = ['wl', 'warlock', 'tetra', 'crimson rock', 'cr']
list_mech = ['mech', 'mechanic', 'mado']
list_ranger = ['ranger', 'range']

############# Responses #####################################
answeryes = ['y', 'yes', 'ya', 'yup', 'ye', 'in', 'g']
answerno = ['n', 'no', 'nah', 'na', 'nope', 'nuh']

######################### CELERY RESPONSES ####################

answerzeny = ['zeny', 'zen', 'money', 'moneh', 'moolah']
answer10 = ['10', 'ten', 'plus ten', 'plusten', '10food', '+10', 'plustens', 'plus tens', '+10s']
answer20 = ['20', 'twenty', 'plus twenty', 'plustwenty', '20food', '+20', 'plustwentys', 'plus twentys', '+20s']
answernone = ['none', 'nada', 'nah', 'nothing', 'waive', 'waived']
answerevery = ['everything', 'all']
answerstr10 = ['+10 str', '+10str']
answeragi10 = ['+10 agi', '+10agi']
answervit10 = ['+10 vit', '+10vit']
answerint10 = ['+10 int', '+10int']
answerdex10 = ['+10 dex', '+10dex']
answerluk10 = ['+10 luk', '+10luk']
answerstr20 = ['+20 str', '+20str']
answeragi20 = ['+20 agi', '+20agi']
answervit20 = ['+20 vit', '+20vit']
answerint20 = ['+20 int', '+20int']
answerdex20 = ['+20 dex', '+20dex']
answerluk20 = ['+20 luk', '+20luk']
answerwhites = ['whites', 'hp pots', 'siege whites', 'white', 'siege white']
answerblues = ['blues', 'sp pots', 'siege blues', 'blue', 'siege blue']



###############################################################

def next_available_row(sheet, column):
    cols = sheet.range(3, column, 50, column)
    return max([cell.row for cell in cols if cell.value]) + 1

def next_available_row_p1(sheet, column):
    cols = sheet.range(3, column, 14, column)
    return max([cell.row for cell in cols if cell.value]) + 1

def next_available_row_p2(sheet, column):
    cols = sheet.range(17, column, 28, column)
    return max([cell.row for cell in cols if cell.value]) + 1
def next_available_row_p3(sheet, column):
    cols = sheet.range(32, column, 43, column)
    return max([cell.row for cell in cols if cell.value]) + 1

class Clears(commands.Cog):
    def __init__(self, client):
        self.client = client



#    async def sorted(self, ctx):
#        channel = ctx.message.channel
#        commander = ctx.author.name
#        if channel.id == botinit_id:
#            cell_list = sheet.range(guild_range)
#            sheet.sort((4, 'asc'), range=guild_range)
#            cell_list = sheet.range(roster_range)
#            sheet.sort((9, 'des'), (8, 'asc'), range=roster_range)
#            await ctx.send(f'{commander} has sorted the sheets.')
#        else:
#            await ctx.send(f'Wrong channel! Please use #bot.')

    @commands.command()
    async def sorted(self, ctx):
        channel = ctx.message.channel
        commander = ctx.author.name
            #await ctx.send('test')
        if channel.id in botinit_id:
            cell_list = sheet.range("B3:E46")
            try:
                sheet.sort((4, 'asc'), range = "B3:E46")
            except Exception as e:
                print(e)
                return
            cell_list = sheet.range("G3:J46")
            sheet.sort((9, 'des'), (8, 'asc'), range = "G3:J46")
            await ctx.send(f'`{commander} has sorted the sheets.`')
        else:
            await ctx.send(f'Wrong channel! Please use #bot.')

    @commands.command()
    async def clearguild(self, ctx):
        channel = ctx.message.channel
        commander_name = ctx.author.name
        commander = ctx.author
        if channel.id in botinit_id:
            if commander.id in authorized_id:
                cell_list = sheet.range(guild_range)
                for cell in cell_list:
                    cell.value = ""
                sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                await ctx.send(f'{commander_name} has cleared the guild list.')
            else:
                await ctx.send(f'This command is unavailable for you!')
        else:
            await ctx.send(f'Wrong channel! Please use #bot.')
        # sh.values_clear("Sheet1!B3:E50")



    @commands.command()
    async def clearroster(self, ctx):
        channel = ctx.message.channel
        commander_name = ctx.author.name
        commander = ctx.author
        if channel.id in botinit_id:
            if commander.id in authorized_id:
                cell_list = sheet.range(roster_range)

                for cell in cell_list:
                    cell.value = ""

                sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                await ctx.send(f'{commander_name} has cleared the WoE Roster.')
            else:
                await ctx.send(f'This command is unavailable for you!')
        else:
            await ctx.send(f'Wrong channel! Please use #bot.')

    @commands.command()
    async def clearparty(self, ctx):
        channel = ctx.message.channel
        commander_name = ctx.author.name
        commander = ctx.author
        if channel.id in botinit_id:
            if commander.id in authorized_id:
                cell_list = sheet.range(matk_range)

                for cell in cell_list:
                    cell.value = ""

                #sheet.update_cells(cell_list)

                sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                cell_list = sheet.range(p1role_range)

                for cell in cell_list:
                    cell.value = ""

                sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                cell_list = sheet.range(atk_range)

                for cell in cell_list:
                    cell.value = ""

                sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                cell_list = sheet.range(p2role_range)

                for cell in cell_list:
                    cell.value = ""

                sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                cell_list = sheet.range(p3_range)

                for cell in cell_list:
                    cell.value = ""

                sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                cell_list = sheet.range(p3role_range)

                for cell in cell_list:
                    cell.value = ""

                sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                await ctx.send(f'{commander_name} has cleared the Party List.')
            else:
                await ctx.send(f'This command is unavailable for you!')
        else:
            await ctx.send(f'Wrong channel! Please use #bot.')

    @commands.command()
    async def enlist(self, ctx, *, arguments):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            arglist = [x.strip() for x in arguments.split(',')]
            no_of_args = len(arglist)
            if no_of_args < 2:
                await ctx.send(f'{ctx.message.author.mention} Please send the proper syntax: `/enlist IGN, role, (optional other classes that you use)`')
                return
            else:
                if arglist[1].lower() in list_ab:
                    darole = 'AB'
                elif arglist[1].lower() in list_doram:
                    darole = 'Doram'
                elif arglist[1].lower() in list_gene:
                    darole = 'Genetic'
                elif arglist[1].lower() in list_mech:
                    darole = 'Mado'
                elif arglist[1].lower() in list_mins:
                    darole = 'Minstrel'
                elif arglist[1].lower() in list_ranger:
                    darole = 'Ranger'
                elif arglist[1].lower() in list_rg:
                    darole = 'RG'
                elif arglist[1].lower() in list_rk:
                        darole = 'RK'
                elif arglist[1].lower() in list_sc:
                    darole = 'SC'
                elif arglist[1].lower() in list_sorc:
                    darole = 'Sorc'
                elif arglist[1].lower() in list_sura:
                    darole = 'Sura'
                elif arglist[1].lower() in list_wl:
                    darole = 'WL'
                elif arglist[1].lower() in list_obo:
                    darole = 'Oboro'
                elif arglist[1].lower() in list_rebel:
                    darole = 'Rebel'
                elif arglist[1].lower() in list_gx:
                    darole = 'GX'
                elif arglist[1].lower() in list_wand:
                    darole = 'Wanderer'
                else:
                    await ctx.send(f'''Here are the following allowable classes: 
```
For Doram: {list_doram}
For Genetic: {list_gene}
For Mechanic: {list_mech}
For Minstrel: {list_mins}
For Ranger: {list_ranger}
For Sorcerer: {list_sorc}
For Oboro: {list_obo}
For Rebellion: {list_rebel}
For Wanderer: {list_wand}
```
                                    ''')
                    return
                change = 0
                #try:
                #    uname = sheet.find(commander_name)
                #    if uname:
                #        next_row = uname.row
                #        ign = sheet.cell(next_row, 3)
                #        change = 1
                #except gspread.exceptions.CellNotFound:
                #    next_row = next_available_row(sheet, 2)
                    #list_entry = sheet.range(next_row, 3, next_row, 4)
                next_row = 3
                cell_list = sheet.range("B3:B50")
                for cell in cell_list:
                    if cell.value == commander_name:
                        change = 1
                        ign = sheet.cell(next_row, 3)
                        break
                    next_row += 1
                if change == 0:
                    next_row = next_available_row(sheet, 2)

                count = 0

                cell_list = sheet.range(next_row, 2, next_row, 5)
                if no_of_args > 2:
                    for cell in cell_list:
                        if count == 0:
                            cell.value = commander_name
                        elif count == 1:
                            cell.value = arglist[0]
                        elif count == 2:
                            cell.value = darole
                        elif count == 3:
                            cell.value = arglist[2]
                        count += 1
                    sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                    await ctx.send(f'```{ctx.author.name} has enlisted {darole} with IGN: {arglist[0]}, and Comment: {arglist[2]}.```')
                    if change == 1:
                        finding_column = sheet.range("G3:G{}".format(sheet.row_count))
                        finding_column2 = celesheet.range("C3:C{}".format(celesheet.row_count))
                        foundign = [found for found in finding_column if found.value == ign.value]
                        foundign2 = [found for found in finding_column2 if found.value == ign.value]

                        if foundign:
                            cell_list = sheet.range(foundign[0].row, 7, foundign[0].row, 10)
                            for cell in cell_list:
                                cell.value = ""
                            sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                            await ctx.send(
                                f'{ctx.message.author.mention}``` I found another character of yours that answered an attendance already, I have cleared that. Please use /att y/n again in order to register your attendance.```')
                            change = 0
                        if foundign2:
                            cell_list = celesheet.range(foundign2[0].row, 2, foundign2[0].row, 20)
                            for cell in cell_list:
                                cell.value = ""
                            celesheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                            await ctx.send(f'{ctx.message.author.mention}``` I found another character of yours that answered celery preferences already, I have cleared that. Please use /celery again in order to list your preferred salary.```')
                            change = 0
                        else:
                            await ctx.send(f'```Please use /att y/n to register your attendance!```')
                            change = 0
                    else:
                        await ctx.send(f'{ctx.message.author.mention}``` Please use /att y/n to register your attendance!```')
                else:
                    for cell in cell_list:
                        if count == 0:
                            cell.value = commander_name
                        elif count == 1:
                            cell.value = arglist[0]
                        elif count == 2:
                            cell.value = darole
                        elif count == 3:
                            cell.value = ""
                        count += 1
                    sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                    await ctx.send(f'```{ctx.author.name} has enlisted {darole} with IGN: {arglist[0]}.```')
                    if change == 1:
                        finding_column = sheet.range("G3:G{}".format(sheet.row_count))
                        foundign = [found for found in finding_column if found.value == ign.value]

                        if foundign:
                            cell_list = sheet.range(foundign[0].row, 7, foundign[0].row, 10)
                            for cell in cell_list:
                                cell.value = ""
                            sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                            await ctx.send(
                                f'{ctx.message.author.mention}``` I found another character of yours that answered an attendance already, I have cleared that. Please use /att y/n again in order to register your attendance.```')
                            change = 0
                        else:
                            await ctx.send(f'```Please use /att y/n to register your attendance!```')
                            change = 0
                    else:
                        await ctx.send(f'```Please use /att y/n to register your attendance!```')
        else:
            await ctx.send("Wrong channel! Please use #bot.")
        cell_list = sheet.range("B3:E50")
        try:
            sheet.sort((4, 'asc'), range="B3:E50")
        except Exception as e:
            print(e)
            return
        cell_list = sheet.range("G3:J50")
        sheet.sort((9, 'des'), (8, 'asc'), range="G3:J50")


    @commands.command()
    async def att(self, ctx, *, arguments):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            arglist = [x.strip() for x in arguments.split(',')]

            no_of_args = len(arglist)

            # await ctx.send(arglist)
            # await ctx.send(no_of_args)
            # return
            yes = 0

            # if no_of_args < 1:
            #    await ctx.send('Please send the proper syntax: ``attendance y/n, (optional comment)`')
            #    return
            # else:
            next_row = 3
            found = 0
            cell_list = sheet.range("B3:B50")
            for cell in cell_list:
                if cell.value == commander_name:
                    found = 1
                    break
                next_row += 1
            if found == 0:
                await ctx.send(f'{ctx.message.author.mention} You have not yet enlisted your character. Please enlist via: `/enlist IGN, class, (optional other classes that you use)`')
                return
            #try:
            #    uname = sheet.find(ctx.author.name)
            #    next_row = uname.row
            #except gspread.exceptions.CellNotFound:
            #    await ctx.send(
            #        f'{ctx.message.author.mention} You have not yet enlisted your character. Please enlist via: `/enlist IGN, class, (optional other classes that you use)`')
            #    return
            #        await ctx.send('test1')

            ign = sheet.cell(next_row, 3)
            role = sheet.cell(next_row, 4)

            finding_column = sheet.range("G3:G50".format(sheet.row_count))

            foundign = [found for found in finding_column if found.value == ign.value]

            if foundign:
                cell_list = sheet.range(foundign[0].row, 7, foundign[0].row, 10)
                if arglist[0].lower() in answeryes or arglist[0].lower() in answerno:
                    count = 0
                    # cell_list = sheet.range(next_row, 7, next_row, 10)
                    if no_of_args > 1:
                        # await ctx.send('test2')
                        for cell in cell_list:
                            if count == 0:
                                cell.value = ign.value
                            elif count == 1:
                                cell.value = role.value
                            elif count == 2:
                                if arglist[0].lower() in answeryes:
                                    cell.value = 'Yes'
                                    yes = 1
                                else:
                                    cell.value = 'No'
                            elif count == 3:
                                cell.value = arglist[1]
                            count += 1
                        sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                        if yes == 1:
                            await ctx.send(
                                f'```{ctx.author.name} said Yes with IGN: {ign.value}, Class: {role.value}, with Comment: {arglist[1]}.```')
                        else:
                            await ctx.send(
                                f'```{ctx.author.name} said No with IGN: {ign.value}, Class: {role.value}, with Comment: {arglist[1]}.```')
                    else:
                        # await ctx.send('test1')
                        for cell in cell_list:
                            # await ctx.send(f'test3 {ign.value} {role.value} {count}')
                            if count == 0:
                                # await ctx.send(f'test4 {ign.value} {role.value} {count}')
                                cell.value = ign.value
                            elif count == 1:
                                # await ctx.send(f'test5 {ign.value} {role.value} {count}')
                                cell.value = role.value
                            elif count == 2:
                                # await ctx.send(f'test6 {ign.value} {role.value} {count}')
                                if arglist[0].lower() in answeryes:
                                    cell.value = 'Yes'
                                    yes = 1
                                else:
                                    cell.value = 'No'
                            elif count == 3:
                                cell.value = ""
                            count += 1
                        sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                        # await ctx.send(f'{yes}')
                        if yes == 1:
                            await ctx.send(
                                f'```{ctx.author.name} said Yes with IGN: {ign.value}, and Class: {role.value}.```')
                        else:
                            await ctx.send(
                                f'```{ctx.author.name} said No with IGN: {ign.value}, and Class: {role.value}.```')

                else:
                    await ctx.send('Please send a proper syntax: ``attendance y/n, (optional comment)`')
                    return
                count = 0
                yes = 0
                no = 0
            else:
                try:
                    change_row = next_available_row(sheet, 7)
                except ValueError as e:
                    change_row = 3
                cell_list = sheet.range(change_row, 7, change_row, 10)
                if arglist[0].lower() in answeryes or arglist[0].lower() in answerno:
                    count = 0
                    #cell_list = sheet.range(next_row, 7, next_row, 10)
                    if no_of_args > 1:
                        # await ctx.send('test2')
                        for cell in cell_list:
                            if count == 0:
                                cell.value = ign.value
                            elif count == 1:
                                cell.value = role.value
                            elif count == 2:
                                if arglist[0].lower() in answeryes:
                                    cell.value = 'Yes'
                                    yes = 1
                                else:
                                    cell.value = 'No'
                            elif count == 3:
                                cell.value = arglist[1]
                            count += 1
                        sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                        if yes == 1:
                            await ctx.send(
                                f'```{ctx.author.name} said Yes with IGN: {ign.value}, Class: {role.value}, with Comment: {arglist[1]}.```')
                        else:
                            await ctx.send(
                                f'```{ctx.author.name} said No with IGN: {ign.value}, Class: {role.value}, with Comment: {arglist[1]}.```')
                    else:
                        # await ctx.send('test2')
                        for cell in cell_list:
                            # await ctx.send(f'test3 {ign.value} {role.value} {count}')
                            if count == 0:
                                # await ctx.send(f'test4 {ign.value} {role.value} {count}')
                                cell.value = ign.value
                            elif count == 1:
                                # await ctx.send(f'test5 {ign.value} {role.value} {count}')
                                cell.value = role.value
                            elif count == 2:
                                # await ctx.send(f'test6 {ign.value} {role.value} {count}')
                                if arglist[0].lower() in answeryes:
                                    cell.value = 'Yes'
                                    yes = 1
                                else:
                                    cell.value = 'No'
                            elif count == 3:
                                cell.value = ""
                            count += 1
                        sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                        # await ctx.send(f'{yes}')
                        if yes == 1:
                            await ctx.send(
                                f'```{ctx.author.name} said Yes with IGN: {ign.value}, and Class: {role.value}.```')
                        else:
                            await ctx.send(
                                f'```{ctx.author.name} said No with IGN: {ign.value}, and Class: {role.value}.```')
                        # await ctx.send('hello')

                else:
                    await ctx.send('Please send a proper syntax: `/attendance y/n, (optional comment)`')
                    return
                count = 0
                yes = 0
                no = 0
        else:
            await ctx.send("Wrong channel! Please use #bot.")
        cell_list = sheet.range("B3:E46")
        try:
            sheet.sort((4, 'asc'), range="B3:E46")
        except Exception as e:
            print(e)
            return
        cell_list = sheet.range("G3:J46")
        sheet.sort((9, 'des'), (8, 'asc'), range="G3:J46")



    @commands.command()
    async def list(self, ctx):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            try:
                row_n = next_available_row(sheet, 7)
            except ValueError:
                row_n = 3
            try:
                row_c = next_available_row(sheet, 8)
            except ValueError:
                row_c = 3
            try:
                row_a = next_available_row(sheet, 9)
            except ValueError:
                row_a = 3
            msg = await ctx.send(f'`Please wait... I am parsing a list of our WOE Roster. Refrain from entering any other commands.`')
            #await asyncio.sleep(10)
            while row_n != row_c or row_n != row_a:
                row_n = next_available_row(sheet, 7)
                row_c = next_available_row(sheet, 8)
                row_a = next_available_row(sheet, 9)
                if row_n < row_c:
                    if row_n < row_a:
                        cell_list = sheet.range(row_n, 7, row_n, 9)
                        for cell in cell_list:
                            cell.value = ""
                        sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                        sheet.sort((9, 'des'), (8, 'asc'), range="G3:J46")
                    else:
                        cell_list = sheet.range(row_a, 7, row_a, 9)
                        for cell in cell_list:
                            cell.value = ""
                        sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                        sheet.sort((9, 'des'), (8, 'asc'), range="G3:J46")
                elif row_c < row_a:
                    cell_list = sheet.range(row_c, 7, row_c, 9)
                    for cell in cell_list:
                        cell.value = ""
                    sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                    sheet.sort((9, 'des'), (8, 'asc'), range="G3:J46")
                else:
                    cell_list = sheet.range(row_a, 7, row_a, 9)
                    for cell in cell_list:
                        cell.value = ""
                    sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                    sheet.sort((9, 'des'), (8, 'asc'), range="G3:J46")

            namae = [item for item in sheet.col_values(7) if item and item != 'IGN']
            kurasu = [item for item in sheet.col_values(8) if item and item != 'Class' and item != 'WoE Roster']
            stat = [item for item in sheet.col_values(9) if item and item != 'Attendance']
            komento = [item for item in sheet.col_values(10) if item and item != 'Comments']
            x = 0
            a = 0
            yuppie = 0
            noppie = 0
            for a in stat:
                if a == 'Yes':
                    yuppie += 1
                else:
                    noppie += 1

            if yuppie == 0 and noppie == 0:
                await ctx.send('`Attendance not found. Please use /att y/n to register your attendance`')
                await msg.delete()
                return

            try:
                embeded = discord.Embed(title = "Current WOE Roster", description = "A list of our Current WOE Roster", color = 0x00FF00)
            except Exception as e:
                print(f'discord embed returned {e}')
                return
            x = 0
            fullname = ''
            fullclass = ''
            fullstat = ''
            for x in range(len(namae)):
                fullname += namae[x] + '\n'
                fullclass += kurasu[x] + '\n'
                fullstat += stat[x] + '\n'

            try:
                embeded.add_field(name="IGN", value=f'{fullname}', inline=True)
            except Exception as e:
                print(f'add field returned {e}')
                return
            embeded.add_field(name="Class", value=f'{fullclass}', inline=True)
            try:
                embeded.add_field(name="Status", value=f'{fullstat}', inline=True)
            except Exception as e:
                print(f'add field returned {e}')
                return


            try:
                await ctx.send(embed=embeded)
            except Exception as e:
                print(f'send embed returned {e}')
            await ctx.send(f'Total no. of Yes answers: {yuppie}')
            await ctx.send(f'Total no. of No answers: {noppie}')
            await msg.delete()
            #return
        else:
            await ctx.send("Wrong channel! Please use #bot.")

    @commands.command()
    async def celery(self, ctx, *, arguments):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name

        zeny = 0
        plusten = 0
        plustwenty = 0
        none = 0
        every = 0
        str10 = 0
        agi10 = 0
        vit10 = 0
        int10 = 0
        dex10 = 0
        luk10 = 0
        str20 = 0
        agi20 = 0
        vit20 = 0
        int20 = 0
        dex20 = 0
        luk20 = 0
        whites = 0
        blues = 0
        noargs = 0
        totalstr = ""

        if channel.id in botinit_id:
            arglist = [x.strip() for x in arguments.split(',')]

            for args in arglist:
                if args.lower() in answerevery:
                    every = 1
                if args.lower() in answernone:
                    none = 1
                if args.lower() in answerzeny:
                    zeny = 1
                if args.lower() in answer10:
                    plusten = 1
                if args.lower() in answer20:
                    plustwenty = 1
                if args.lower() in answernone:
                    none = 1
                if args.lower() in answerstr10:
                    str10 = 1
                if args.lower() in answeragi10:
                    agi10 = 1
                if args.lower() in answervit10:
                    vit10 = 1
                if args.lower() in answerint10:
                    int10 = 1
                if args.lower() in answerdex10:
                    dex10 = 1
                if args.lower() in answerluk10:
                    luk10 = 1
                if args.lower() in answerstr20:
                    str20 = 1
                if args.lower() in answeragi20:
                    agi20 = 1
                if args.lower() in answervit20:
                    vit20 = 1
                if args.lower() in answerint20:
                    int20 = 1
                if args.lower() in answerdex20:
                    dex20 = 1
                if args.lower() in answerluk20:
                    luk20 = 1
                if args.lower() in answerwhites:
                    whites = 1
                if args.lower() in answerblues:
                    blues = 1
                if args.lower() in answer10:
                    plusten = 1
                if args.lower() in answer20:
                    plustwenty = 1

            if plusten == 1:
                str10 = 0
                agi10 = 0
                vit10 = 0
                int10 = 0
                dex10 = 0
                luk10 = 0
            if plustwenty == 1:
                str20 = 0
                agi20 = 0
                vit20 = 0
                int20 = 0
                dex20 = 0
                luk20 = 0
            if str10 == 1 and agi10 == 1 and vit10 == 1 and int10 == 1 and dex10 == 1 and luk10 == 1:
                plusten = 1
                str10 = 0
                agi10 = 0
                vit10 = 0
                int10 = 0
                dex10 = 0
                luk10 = 0
            if str20 == 1 and agi20 == 1 and vit20 == 1 and int20 == 1 and dex20 == 1 and luk20 == 1:
                plustwenty = 1
                str20 = 0
                agi20 = 0
                vit20 = 0
                int20 = 0
                dex20 = 0
                luk20 = 0
            if plustwenty == 1 and plusten == 1 and whites == 1 and blues == 1 and zeny == 1:
                every = 1

            if every == 1:
                none = 0
                zeny = 0
                str10 = 0
                agi10 = 0
                vit10 = 0
                int10 = 0
                dex10 = 0
                luk10 = 0
                str20 = 0
                agi20 = 0
                vit20 = 0
                int20 = 0
                dex20 = 0
                luk20 = 0
                whites = 0
                blues = 0
                plusten = 0
                plustwenty = 0
            elif none == 1:
                none = 0
                every = 0
                str10 = 0
                agi10 = 0
                vit10 = 0
                int10 = 0
                dex10 = 0
                luk10 = 0   
                str20 = 0
                agi20 = 0
                vit20 = 0
                int20 = 0
                dex20 = 0
                luk20 = 0
                whites = 0
                blues = 0
                plustwenty = 0


            if zeny == 0 and plusten == 0 and plustwenty == 0 and none == 0 and every == 0 and str10 == 0 and agi10 == 0 and vit10 == 0 and int10 == 0 and dex10 == 0 and luk10 == 0 and str20 == 0 and agi20 == 0 and vit20 == 0 and int20 == 0 and dex20 == 0 and luk20 == 0 and whites == 0 and blues == 0:
                noargs = 1

            if noargs == 1:
                await ctx.send("```Please use the correct syntax /celery zeny, plusten, plustwenty. MIND THE COMMA.```")


            no_of_args = len(arglist)
            found = 0
            next_row = 3
            cell_list = sheet.range("B3:B50")
            for cell in cell_list:
                if cell.value == commander_name:
                    found = 1
                    break
                next_row += 1
            if found == 0:
                await ctx.send(f'{ctx.message.author.mention} You have not yet enlisted your character. Please enlist via: `/enlist IGN, class, (optional other classes that you use)`')
                return

            ign = sheet.cell(next_row, 3)
            role = sheet.cell(next_row, 4)
            count = 0

            finding_column = celesheet.range("C3:C50".format(celesheet.row_count))
            foundign = [found for found in finding_column if found.value == ign.value]

            if foundign:
                cell_list = celesheet.range(foundign[0].row, 2, foundign[0].row, 26)
                for cell in cell_list:
                    cell.value = ""
                cell_list = celesheet.range(foundign[0].row, 2, foundign[0].row, 26)
                for cell in cell_list:
                    if count == 0:
                        cell.value = commander_name
                    elif count == 1:
                        cell.value = ign.value
                    elif count == 2:
                        cell.value = role.value
                    elif count == 3:
                        if zeny == 1 or every == 1:
                            cell.value = "Yes"
                            totalstr += "ZENY; "
                        else:
                            cell.value = "No"
                    elif count == 4:
                        if str10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr+= "STR10; "
                        else:
                            cell.value = "No"
                    elif count == 5:
                        if agi10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "AGI10; "
                        else:
                            cell.value = "No"
                    elif count == 6:
                        if vit10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "VIT10; "
                        else:
                            cell.value = "No"
                    elif count == 7:
                        if int10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "INT10; "
                        else:
                            cell.value = "No"
                    elif count == 8:
                        if dex10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "DEX10; "
                        else:
                            cell.value = "No"
                    elif count == 9:
                        if luk10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "LUK10; "
                        else:
                            cell.value = "No"
                    elif count == 10:
                        if str20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "STR20; "
                        else:
                            cell.value = "No"
                    elif count == 11:
                        if agi20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "AGI20; "
                        else:
                            cell.value = "No"
                    elif count == 12:
                        if vit20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "VIT20; "
                        else:
                            cell.value = "No"
                    elif count == 13:
                        if int20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "INT20; "
                        else:
                            cell.value = "No"
                    elif count == 14:
                        if dex20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "DEX20; "
                        else:
                            cell.value = "No"
                    elif count == 15:
                        if luk20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "LUK20; "
                        else:
                            cell.value = "No"
                    elif count == 16:
                        if whites == 1 or every == 1:
                            cell.value = "Yes"
                            totalstr += "WHITES; "
                        else:
                            cell.value = "No"
                    elif count == 17:
                        if blues == 1 or every == 1:
                            cell.value = "Yes"
                            totalstr += "BLUES; "
                        else:
                            cell.value = "No"
                    elif count == 18:
                        if every == 1:
                            cell.value = "EVERYTHING"
                        elif none == 1:
                            cell.value = "NONE"
                        # if zeny == 1:
                        #     if plusten == 1:
                        #         if plustwenty == 1:
                        #             if whites == 1:
                        #                 if blues == 1:
                        #                     cell.value = "EVERYTHING"
                        #                 else:
                        #                     cell.value = "ZENY; ALL +10; ALL +20; WHITES;"
                        #             else:
                        #                 if blues == 1:
                        #                     cell.value = "ZENY; ALL +10; ALL +20; BLUES;"
                        #                 else:
                        #                     cell.value = "ZENY; ALL +10; ALL +20;"
                        #         else:
                        #             if whites == 1:
                        #                 if blues == 1:
                        #                     cell.value = "ZENY; ALL +10; WHITES; BLUES;"
                        #                 else:
                        #                     cell.value = "ZENY; ALL +10; ALL +20; WHITES;"
                        #             else:
                        #                 if blues == 1:
                        #                     cell.value = "ZENY; ALL +10; ALL +20; BLUES;"
                        #                 else:
                        #                     cell.value = "ZENY; ALL +10;"
                        #     else:
                        #         if whites == 1:
                        #             if blues == 1:
                        #                 cell.value = "ZENY; ALL +10; WHITES; BLUES;"
                        #             else:
                        #                 cell.value = "ZENY; ALL +10; ALL +20; WHITES;"
                        #         else:
                        #             if blues == 1:
                        #                 cell.value = "ZENY; ALL +10; ALL +20; BLUES;"
                        #             else:
                        #                 cell.value = "ZENY; ALL +10;"
                        else:
                            cell.value = totalstr

                    count += 1
                celesheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                celery_list = celesheet.cell(foundign[0].row, 20).value
                await ctx.send(f'```{ctx.author.name} wanted {celery_list}with IGN: {ign.value}, and Class: {role.value}.```')
            else:
                try:
                    change_row = next_available_row(celesheet, 2)
                except ValueError as e:
                    change_row = 3
                cell_list = celesheet.range(change_row, 2, change_row, 26)
                for cell in cell_list:
                    if count == 0:
                        cell.value = commander_name
                    elif count == 1:
                        cell.value = ign.value
                    elif count == 2:
                        cell.value = role.value
                    elif count == 3:
                        if zeny == 1 or every == 1:
                            cell.value = "Yes"
                            totalstr += "ZENY; "
                        else:
                            cell.value = "No"
                    elif count == 4:
                        if str10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "STR10; "
                        else:
                            cell.value = "No"
                    elif count == 5:
                        if agi10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "AGI10; "
                        else:
                            cell.value = "No"
                    elif count == 6:
                        if vit10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "VIT10; "
                        else:
                            cell.value = "No"
                    elif count == 7:
                        if int10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "INT10; "
                        else:
                            cell.value = "No"
                    elif count == 8:
                        if dex10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "DEX10; "
                        else:
                            cell.value = "No"
                    elif count == 9:
                        if luk10 == 1 or every == 1 or plusten == 1:
                            cell.value = "Yes"
                            totalstr += "LUK10; "
                        else:
                            cell.value = "No"
                    elif count == 10:
                        if str20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "STR20; "
                        else:
                            cell.value = "No"
                    elif count == 11:
                        if agi20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "AGI20; "
                        else:
                            cell.value = "No"
                    elif count == 12:
                        if vit20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "VIT20; "
                        else:
                            cell.value = "No"
                    elif count == 13:
                        if int20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "INT20; "
                        else:
                            cell.value = "No"
                    elif count == 14:
                        if dex20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "DEX20; "
                        else:
                            cell.value = "No"
                    elif count == 15:
                        if luk20 == 1 or every == 1 or plustwenty == 1:
                            cell.value = "Yes"
                            totalstr += "LUK20; "
                        else:
                            cell.value = "No"
                    elif count == 16:
                        if whites == 1 or every == 1:
                            cell.value = "Yes"
                            totalstr += "WHITES; "
                        else:
                            cell.value = "No"
                    elif count == 17:
                        if blues == 1 or every == 1:
                            cell.value = "Yes"
                            totalstr += "BLUES; "
                        else:
                            cell.value = "No"
                    elif count == 18:
                        if every == 1:
                            cell.value = "EVERYTHING"
                        elif none == 1:
                            cell.value = "NONE"
                        # if zeny == 1:
                        #     if plusten == 1:
                        #         if plustwenty == 1:
                        #             if whites == 1:
                        #                 if blues == 1:
                        #                     cell.value = "EVERYTHING"
                        #                 else:
                        #                     cell.value = "ZENY; ALL +10; ALL +20; WHITES;"
                        #             else:
                        #                 if blues == 1:
                        #                     cell.value = "ZENY; ALL +10; ALL +20; BLUES;"
                        #                 else:
                        #                     cell.value = "ZENY; ALL +10; ALL +20;"
                        #         else:
                        #             if whites == 1:
                        #                 if blues == 1:
                        #                     cell.value = "ZENY; ALL +10; WHITES; BLUES;"
                        #                 else:
                        #                     cell.value = "ZENY; ALL +10; ALL +20; WHITES;"
                        #             else:
                        #                 if blues == 1:
                        #                     cell.value = "ZENY; ALL +10; ALL +20; BLUES;"
                        #                 else:
                        #                     cell.value = "ZENY; ALL +10;"
                        #     else:
                        #         if whites == 1:
                        #             if blues == 1:
                        #                 cell.value = "ZENY; ALL +10; WHITES; BLUES;"
                        #             else:
                        #                 cell.value = "ZENY; ALL +10; ALL +20; WHITES;"
                        #         else:
                        #             if blues == 1:
                        #                 cell.value = "ZENY; ALL +10; ALL +20; BLUES;"
                        #             else:
                        #                 cell.value = "ZENY; ALL +10;"
                        else:
                            cell.value = totalstr
                    count += 1
                celesheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                celery_list = celesheet.cell(foundign[0].row, 20).value
                await ctx.send(f'```{ctx.author.name} wanted {celery_list}with IGN: {ign.value}, and Class: {role.value}.```')
                cell_list = celesheet.range("B3:G48")
                try:
                    celesheet.sort((4, 'asc'), range = "B3:G48")
                except Exception as e:
                    print(f'celesheet sort has returned {e}')
                    return
        else:
            await ctx.send("Wrong channel! Please use #bot.")

    @commands.command()
    async def help(self, ctx):

        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            await ctx.send("""```BOT COMMANDS:
/enlist IGN, class, optional comment = enlists your Discord ID, IGN, Class, and optional comment in the GSheets
/att y/n, optional comment = registers your attendance (either yes or no) in the GSheets
/clearguild = clears guild list (ADMIN COMMAND)
/clearroster = clears attendance list (ADMIN COMMAND)
/clearparty = clears party list (ADMIN COMMAND)
/list = parses a list of the current attendance list
/listpt = parses a list of the current party list divided into ATK, MATK, and THIRD GUILD
/sorted = sorts the gsheets
PLEASE MIND THE COMMA, IT ENSURES THAT I SEE EVERY ARGUMENT. 
Thank you!```\n""")
        else:
            await ctx.send("Wrong channel! Please use #bot.")

    @commands.command()
    async def listpt(self, ctx):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            msg = await ctx.send(
                f'`Please wait... I am parsing a list of our Party List. Refrain from entering any other commands.`')
            cell_list = sheet.range("M3:M14")
            get_MATK = [""]
            for cell in cell_list:
                get_MATK.append(cell.value)
            cell_list = sheet.range("M17:M28")
            get_ATK = [""]
            for cell in cell_list:
                get_ATK.append(cell.value)
            cell_list = sheet.range("M32:M43")
            get_third = [""]
            for cell in cell_list:
                get_third.append(cell.value)

            MATK_names = [item for item in get_MATK if item]
            ATK_names = [item for item in get_ATK if item]
            THIRD_names = [item for item in get_third if item]

            try:
                embeded = discord.Embed(title="Current Party List", description="A list of our Current Party List",
                                        color=0x00FF00)
            except Exception as e:
                print(f'discord embed returned {e}')
                return
            x = 0
            ATKpt = ''
            MATKpt = ''
            THIRDpt = ''
            for x in range(len(MATK_names)):
                MATKpt += MATK_names[x] + '\n'
            x = 0
            for x in range(len(ATK_names)):
                ATKpt += ATK_names[x] + '\n'
            x = 0
            for x in range(len(THIRD_names)):
                THIRDpt += THIRD_names[x] + '\n'
            try:
                embeded.add_field(name="ATK Party", value=f'{ATKpt}', inline=True)
            except Exception as e:
                print(f'add field returned {e}')
                return
            embeded.add_field(name="MATK Party", value=f'{MATKpt}', inline=True)
            try:
                embeded.add_field(name="THIRD GUILD Party", value=f'{THIRDpt}', inline=True)
            except Exception as e:
                print(f'add field returned {e}')
                return
            try:
                await ctx.send(embed=embeded)
            except Exception as e:
                print(f'send embed returned {e}')
            await msg.delete()
            # return
        else:
            await ctx.send("Wrong channel! Please use #bot.")

    @commands.command()
    async def listcelery(self, ctx):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            msg = await ctx.send(f'`Please wait... I am parsing a list of our Salary Preferences List. Refrain from entering any other commands.`')
            cell_list = celesheet.range("C3:C48")
            get_ign = [""]
            for cell in cell_list:
                get_ign.append(cell.value)
            cell_list = celesheet.range("D3:D48")
            get_class = [""]
            for cell in cell_list:
                get_class.append(cell.value)
            cell_list = celesheet.range("T3:T48")
            get_pref = [""]
            for cell in cell_list:
                get_pref.append(cell.value)

            ign = [item for item in get_ign if item]
            role = [item for item in get_class if item]
            pref = [item for item in get_pref if item]

            try:
                embeded = discord.Embed(title="Salary Preferences", description="A list of our Salary Preferences", color=0x00FF00)
            except Exception as e:
                printf(f'discord embed retturned {e}')
                return
            x = 0
            ignlist = ''
            classlist = ''
            preflist = ''

            no_of_pref = [x.strip() for x in pref[x].split(';')]

            for x in range(len(ign)):
                ignlist += ign[x] + '\n'
                no_of_pref = len([x.strip() for x in pref[x].split(';')])
                if no_of_pref > 9:
                    ignlist += '\n'
            x = 0
            for x in range(len(role)):
                classlist += role[x] + '\n'
                no_of_pref = len([x.strip() for x in pref[x].split(';')])
                if no_of_pref > 9:
                    classlist += '\n'
            x = 0
            for x in range(len(pref)):
                preflist += pref[x] + '\n'
            x = 0

            embeded.add_field(name="IGN", value=f'{ignlist}', inline=True)
            embeded.add_field(name="CLASS", value=f'{classlist}', inline=True)
            embeded.add_field(name="PREFERENCES", value=f'{preflist}', inline=True)

            await ctx.send(embed=embeded)

            await msg.delete()

        else:
            await ctx.send("Wrong channel! Please use #bot.")

    @commands.command()
    async def totalcelery(self, ctx):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            msg = await ctx.send(f'`Please wait... I am parsing a list of our Total Salary Preferences. Refrain from entering any other commands.`')
            cell_list = celesheet.range("E49:S49")

            count = 0
            zeny = 0
            str10 = 0
            agi10 = 0
            vit10 = 0
            int10 = 0
            dex10 = 0
            luk10 = 0
            str20 = 0
            agi20 = 0
            vit20 = 0
            int20 = 0
            dex20 = 0
            luk20 = 0
            whites = 0
            blues = 0


            for cell in cell_list:
                if count == 0:
                    zeny = cell.value
                if count == 1:
                    str10 = cell.value
                if count == 2:
                    agi10 = cell.value
                if count == 3:
                    vit10 = cell.value
                if count == 4:
                    int10 = cell.value
                if count == 5:
                    dex10 = cell.value
                if count == 6:
                    luk10 = cell.value
                if count == 7:
                    str20 = cell.value
                if count == 8:
                    agi20 = cell.value
                if count == 9:
                    vit20 = cell.value
                if count == 10:
                    int20 = cell.value
                if count == 11:
                    dex20 = cell.value
                if count == 12:
                    luk20 = cell.value
                if count == 13:
                    whites = cell.value
                if count == 14:
                    blues = cell.value
                count += 1

            await ctx.send(f'''```Right now, the guild needs to provide:
Zeny x {zeny}
+10 Str x {str10}
+10 Agi x {agi10}
+10 Vit x {vit10}
+10 Int x {int10}
+10 Dex x {dex10}
+10 Luk x {luk10}
+20 Str x {str20}
+20 Agi x {agi20}
+20 Vit x {vit20}
+20 Int x {int20}
+20 Dex x {dex20}
+20 Luk x {luk20}
White Pots x {whites}
Blue Pots x {blues}
                            ```''')

            await msg.delete()
        else:
            await ctx.send("Wrong channel! Please use #bot.")

def setup(client):
    client.add_cog(Clears(client))