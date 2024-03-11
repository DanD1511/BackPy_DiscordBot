import os
from dotenv import load_dotenv
from discord.ext import commands
import discord

load_dotenv()
TOKEN = os.getenv('DISCORD_TOKEN')

intents = discord.Intents.default()
intents.messages = True 
bot = commands.Bot(command_prefix='!', intents=intents)

@bot.command(name='Sum')
async def sum(ctx, num1, num2):
    total = int(num1) + int(num2)
    
    await ctx.send(total)

bot.run(TOKEN)
