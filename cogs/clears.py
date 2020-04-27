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
answer10 = ['10', 'ten', 'plus ten', 'plusten', '10food', '+10']
answer20 = ['20', 'twenty', 'plus twenty', 'plustwenty', '20food', '+20']
answernone = ['none', 'nada', 'nah', 'nothing']

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

        if channel.id in botinit_id:
            arglist = [x.strip() for x in arguments.split(',')]

            for args in arglist:
                if args.lower in answerzeny:
                    zeny = 1
                if args.lower in answer10:
                    plusten = 1
                if args.lower in answer20:
                    plustwenty = 1
                if args.lower in answernone:
                    none = 1

            if zeny == 0 and plusten == 0 and plustwenty == 0 and none == 0:
                await ctx.send{"```Please use the correct syntax /celery zeny, plusten, plustwenty. MIND THE COMMA```"}


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

            finding_column = celesheet.range("B3:B50".format(celesheet.row_count))
            foundign = [found for found in finding_column if found.value == ign.value]

            if foundign:
                cell_list = celesheet.range(foundign[0].row, 2, foundign[0].row, 7)
                for cell in cell_list:
                    if count == 0:
                        cell.value = commander_name
                    elif count == 1:
                        cell.value = ign.value
                    elif count == 2:
                        cell.value = role.value
                    elif count == 3:
                        if zeny = 1:
                            cell.value = "Yes"
                        else:
                            cell.value = "No"
                    elif count == 4:
                        if plusten = 1:
                            cell.value = "Yes"
                        else:
                            cell.value = "No"
                    elif count == 5:
                        if plustwenty = 1:
                            cell.value = "Yes"
                        else:
                            cell.value = "No"
                    count += 1
                celesheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                celery_list = celesheet.cell(foundign[0].row, 8).value
                await ctx.send(f'```{ctx.author.name}  wanted {celery_list} with IGN: {ign.value}, and Class: {role.value}.```')
            else:
                try:
                    change_row = next_available_row(celesheet, 2)
                except ValueError as e:
                    change_row = 3
                cell_list = celesheet.range(change_row, 2, change_row, 7)
                for cell in cell_list:
                    if count == 0:
                        cell.value = commander_name
                    elif count == 1:
                        cell.value = ign.value
                    elif count == 2:
                        cell.value = role.value
                    elif count == 3:
                        if zeny = 1:
                            cell.value = "Yes"
                        else:
                            cell.value = "No"
                    elif count == 4:
                        if plusten = 1:
                            cell.value = "Yes"
                        else:
                            cell.value = "No"
                    elif count == 5:
                        if plustwenty = 1:
                            cell.value = "Yes"
                        else:
                            cell.value = "No"
                    count += 1
                celesheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                celery_list = celesheet.cell(foundign[0].row, 8).value
                await ctx.send(f'```{ctx.author.name}  wanted {celery_list} with IGN: {ign.value}, and Class: {role.value}.```')
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


def setup(client):
    client.add_cog(Clears(client))