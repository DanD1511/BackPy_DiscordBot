import os
import discord
from discord.ext import commands
from function_repository import build
from docxtpl import DocxTemplate
import io
from dataClass import sheetList
from dotenv import load_dotenv

load_dotenv()
TOKEN = os.getenv('DISCORD_TOKEN')

intents = discord.Intents.default()
intents.messages = True
bot = commands.Bot(command_prefix='!', intents=intents)

TEMPLATE_PATH = 'BotDiscord/Plantilla_OR_Panel640.docx'

@bot.command(name='generateDocument')
async def generate_document(ctx):

    if not ctx.message.attachments:
        await ctx.send("Por favor, adjunta un archivo Excel con los datos necesarios.")
        return

    attachment = ctx.message.attachments[0]
    temp_file_path = 'temp_data.xlsx'
    await attachment.save(temp_file_path)
    await ctx.send(f"Espere mientras se genera el documento")

    try:
        doc = DocxTemplate(TEMPLATE_PATH)
        context = build(temp_file_path, sheetList)
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
