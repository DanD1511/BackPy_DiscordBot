from dataClass import dataClass, sheetList, TableConfig
import pandas as pd
from docxtpl import InlineImage
from docx.shared import Cm
import matplotlib.pyplot as plt
import numpy as np


def build(data, doc, sheetList):
        df, contextVariables = importData(data, sheetList[0])
        context = {}

        for i, variable in enumerate(contextVariables):
            context[variable] = df[i]

        chartGenDiaryPath = chartGenDiary(data) 
        chartGenDiaryImage = pasteChart(doc, chartGenDiaryPath)
        context['Chart1'] = chartGenDiaryImage

        chartEnProyPath = chartEnProy(data) 
        chartEnProyImage = pasteChart(doc, chartEnProyPath)
        context['Chart2'] = chartEnProyImage

        for T in range(1, len(dataClass) + 1): 
            tableKey = f'T{T}'
            if tableKey in dataClass:
                tableConfig = dataClass[tableKey] 
                tableContext(data, context,
                                          tableConfig['SheetName'], tableConfig['StartIndex'],
                                          tableConfig['EndIndex'], tableConfig['coords'],
                                          tableConfig['TableName']
                                          )

        return context

def NumCoord(coord):

        if coord.isupper():
            return ord(coord) - ord('A') + 1
        else:
            return ord(coord) - ord('a') + 1

def numToCoord(coord):

        number = 0
        power = len(coord) - 1
        for coord_ in coord:
            number += NumCoord(coord_) * (26 ** power)
            power -= 1
        return number - 1

def tableContext(data, context, sheetName, startIndex, endIndex, coords, tableName):
        dataFrame = pd.read_excel(data, sheet_name=sheetName)
        letter = numToCoord(coords)
        index = 1
        while startIndex < endIndex:
            value = dataFrame.iloc[startIndex, letter]
            if type(value) == float and value > 1:
                context[f'{tableName}{index}'] = int(value)
            elif type(value) == float and value < 1:
                context[f'{tableName}{index}'] = round(value, 2)
            else:
                context[f'{tableName}{index}'] = value

            index += 1
            startIndex += 1

def pasteChart(doc_template, chart_image_path):
    chart_image = InlineImage(doc_template, chart_image_path, Cm(15))
    return chart_image

def saveChart(image, image_path):
    image.save(image_path)

def importData(data, sheetName): 

        df = pd.read_excel(data, sheet_name=sheetName, header=None)
        dataFrame = []
        contextVariables = []

        index = 1
        while index < len(df):
            dataFrame.append(df.iloc[index, 1])
            contextVariables.append(df.iloc[index, 0])
            index += 1

        return dataFrame, contextVariables
    

        
def processChart(doc, data, sheetName):
    #TODO

        return chartPath

def chartGenDiary(data):
    dataFrame = pd.read_excel(data, sheet_name='1. Generación diaria', header=None)
    xAxis = []
    yAxis = []

    index = 2
    while index < len(dataFrame) - 3:
        xAxis.append(str(dataFrame.iloc[index, 1]))
        yAxis.append(int(dataFrame.iloc[index, 2]))
        index += 1

    fig, ax = plt.subplots(figsize=(14, 9))

    normalized_green_color = (146/255, 208/255, 80/255)
    normalized_green_border_color = (112/255, 173/255, 71/255)

    plt.fill_between(xAxis, yAxis, color=normalized_green_color, alpha=1, zorder=3)

    plt.plot(xAxis, yAxis, color=normalized_green_border_color, linewidth=5, zorder=4)

    plt.title('Generación Diaria de Energía')
    plt.xlabel('Hora')
    plt.ylabel('Potencia [kW]')

    tick_positions = range(0, len(xAxis), 4)
    tick_labels = [xAxis[i] for i in tick_positions]
    plt.xticks(tick_positions, tick_labels, rotation=45, fontsize=12)

    plt.ylim(0, 1300)

    yticks = range(0, 1301, 200)
    plt.yticks(yticks, [str(i) for i in yticks], fontsize=12)

    for spine in ax.spines.values():
        spine.set_visible(False)

    plt.gca().set_axisbelow(True)

    plt.grid(True, axis='y', color='lightgrey', linestyle='--', linewidth=0.5)

    plt.tight_layout()

    plt.savefig('Chart1.png')
    plt.close()
    return 'Chart1.png'

def chartEnProy(data):
    sheet_name = '1.1. Proyección de Gen.'

    dataFrame = pd.read_excel(data, sheet_name=sheet_name, header=None)

    Data1 = []
    Data2 = []

    index = 2
    while index < 14:
        Data1.append(int(dataFrame.iloc[index, 2]))
        index += 1

    index = 17
    while index < 29:
        Data2.append(int(dataFrame.iloc[index, 2]))
        index += 1

    fig, ax1 = plt.subplots(figsize=(14, 9))

    normalized_green_color = (146/255, 208/255, 80/255)
    normalized_grey_color = (89/255, 89/255, 89/255)

    bar_width = 0.35
    index = np.arange(len(Data1))
    bars1 = ax1.bar(index - bar_width/2, Data1, bar_width, color=normalized_green_color, label=f'Generación: {int(dataFrame.iloc[14, 2])}')

    ax1.set_ylabel('[kWh]')
    ax1.tick_params(axis='y', labelsize=12)
    ax1.set_ylim(0, 250000)

    months = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    ax1.set_xticks(index)
    ax1.tick_params(axis='x', labelsize=12)
    ax1.set_xticklabels(months)

    ax2 = ax1.twinx()
    bars2 = ax2.bar(index + bar_width/2, Data2, bar_width, color=normalized_grey_color, label=f'Consumo: {dataFrame.iloc[29, 2]}')

    ax2.set_ylabel('[kWh]')
    ax2.tick_params(axis='y', labelsize=12)
    ax2.set_ylim(0, 5000)

    for spine in ax1.spines.values():
        spine.set_visible(False)

    for spine in ax2.spines.values():
        spine.set_visible(False)

    ax1.grid(True, axis='y', color='lightgrey', linestyle='--', linewidth=0.5)
    ax1.set_axisbelow(True)

    ax1.legend(loc='upper left')
    ax2.legend(loc='upper right')

    plt.tight_layout()

    plt.savefig('Chart2.png')
    plt.close()
    return 'Chart2.png'

