from dataClass import dataClass, sheetList, TableConfig
import pandas as pd
from docxtpl import InlineImage
from docx.shared import Cm


def build(data, sheetList):
        df, contextVariables = importData(data, sheetList[0])
        context = {}

        for i, variable in enumerate(contextVariables):
            context[variable] = df[i]

        # for T in range(1, len(dataClass) + 1):  # Ensure T starts at 1 and iterates through the length of dataClass
        #     tableKey = f'T{T}'  # Construct the key for each table
        #     if tableKey in dataClass:  # Check if the key exists in dataClass
        #         tableConfig = dataClass[tableKey]  # Access the configuration for the current table
        #         # Use the dictionary values correctly with keys
        #         tableContext(data, context,
        #                                   tableConfig['SheetName'], tableConfig['StartIndex'],
        #                                   tableConfig['EndIndex'], tableConfig['coords'],
        #                                   tableConfig['TableName']
        #                                   )

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

def validate(template, data, outputDirectory):
        if not template:
            messagebox.showerror("Error", "Por favor, selecciona un archivo de plantilla")
            return False

        if not data:
            messagebox.showerror("Error", "Por favor, selecciona un archivo de datos")

        if not outputDirectory:
            messagebox.showerror("Error", "Por favor, selecciona un directorio de salida.")
            return False

        return True

def importData(data, sheetName): 
        try:
            df = pd.read_excel(data, sheet_name=sheetName, header=None)
            dataFrame = []
            contextVariables = []

            index = 1
            indexVariable = 1
            while index < len(df):
                dataFrame.append(df.iloc[index, 1])
                contextVariables.append(df.iloc[index, 0])
                index += 1

            return dataFrame, contextVariables

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo de datos: {e}")
            return None
        
# def processChart(doc, data, sheetName):
#         chartImage = ChartExtractor.chartExtractor(data, sheetName)
#         chartPath = os.path.join('', f"chart_{sheetName}.png")
#         ImageProcessor.saveChart(chartImage, chartPath)
#         chart = ImageProcessor.pasteChart(doc, chartPath)

#         return chartPath

# TODO igualar la funcionalidad en linux 

# def chartExtractor(data, sheetName):
#         chartIndex = 0
#         operation = win32com.client.Dispatch("Excel.Application")
#         operation.Visible = 0
#         operation.DisplayAlerts = 0

#         workbook = operation.Workbooks.Open(data)
#         sheet = workbook.Worksheets(sheetName)

#         chart = None
#         for x, chart_obj in enumerate(sheet.Shapes):
#             if x == chartIndex:
#                 chart = chart_obj
#                 break

#         if chart:
#             chart.Copy()
#             image = ImageGrab.grabclipboard()
#             workbook.Close(True)
#             operation.Quit()
#             return image

#         workbook.Close(True)
#         operation.Quit()
#         return None


