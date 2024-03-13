import discord
from discord.ext import commands
import asyncio
import os
import warnings
from docxtpl import DocxTemplate
import io
from dotenv import load_dotenv
from function_repository import build
from dataClass import sheetList

load_dotenv()
TOKEN = os.getenv('DISCORD_TOKEN')

intents = discord.Intents.default()
intents.messages = True
bot = commands.Bot(command_prefix='!', intents=intents)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

userStates = {}

TEMPLATE_PATHS = {
    '640': 'Plantilla_OR_Panel640.docx',
    '600': 'Plantilla_OR_Panel600.docx',
    '580': 'Plantilla_OR_Panel580.docx',
}

def is_user_interacting(user_id):
    return userStates.get(user_id, None) is not None


def set_user_state(user_id, state):
    userStates[user_id] = state


def clear_user_state(user_id):
    if user_id in userStates:
        del userStates[user_id]

async def request_panel_type(ctx):
    await ctx.send("Por favor, indica el tipo de panel (640, 600, 580):")

    def check(m):
        return m.author == ctx.author and m.channel == ctx.channel and m.content in TEMPLATE_PATHS

    try:
        panel_type_msg = await bot.wait_for('message', check=check, timeout=60.0)
        return panel_type_msg.content
    except asyncio.TimeoutError:
        await ctx.send("Tiempo de espera agotado. Por favor, intenta el comando nuevamente.")
        return None

async def request_attachment(ctx):
    await ctx.send("Por favor, sube la memoria de cálculo.")

    def check(m):
        return m.author == ctx.author and m.channel == ctx.channel and m.attachments

    try:
        attachment_msg = await bot.wait_for('message', check=check, timeout=300.0)
        return attachment_msg.attachments[0]
    except asyncio.TimeoutError:
        await ctx.send("Tiempo de espera agotado. Por favor, intenta el comando nuevamente.")
        return None

@bot.command(name='OR')
async def generate_document(ctx):

    user_id = ctx.author.id
    
    if is_user_interacting(user_id):
        await ctx.send("Ya estás realizando una operación. Por favor, espera a que se complete.")
        return
    
    set_user_state(user_id, "awaiting_panel_type")
    await ctx.send("Por favor, indica el tipo de panel (640, 600, 580):")

    panel_type = await request_panel_type(ctx)
    if panel_type is None:
        return

    attachment = await request_attachment(ctx)
    if attachment is None:
        return

    temp_file_path = 'temp_data.xlsx'
    await attachment.save(temp_file_path)
    await ctx.send("Espere mientras se genera el documento.")

    try:
        TEMPLATE_PATH = TEMPLATE_PATHS[panel_type]
        doc = DocxTemplate(TEMPLATE_PATH)
        context = build(temp_file_path, doc, sheetList)
        doc.render(context)

        output_stream = io.BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)

        file_name = f"{context['ProjectName']}-INF-ELE-OR.docx"
        await ctx.send(f"Documento {file_name} generado con éxito.", file=discord.File(fp=output_stream, filename=file_name))
    except Exception as e:
        await ctx.send(f"Error al generar el documento: {e}")
    finally:
        os.remove(temp_file_path)

bot.run(TOKEN)
