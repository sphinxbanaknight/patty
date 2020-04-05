import discord
import random
import os
import json
import gspread
import pprint
from oauth2client import file as oauth_file, client, tools
from apiclient.discovery import build
from httplib2 import Http
#from boto.s3.connection import S3Connection

from discord.ext import commands, tasks


prefix = ["!", "$", "-", "/"]
description = "A bot for sheet+discord linking/automation."
client = commands.Bot(command_prefix = prefix, description = description)

#pp = pprint.PrettyPrinter()

#result = sheet.col_values(4)

#pp.pprint(result)



'''@client.command()
async def identify(ctx):
    print(ctx.author.name)'''
#ctx.author.name to get the username of the one who activated the command


     #str_list = list(filter(None, sheet.col_values(column)))
     #return str(len(str_list) + 1)

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