from dataclasses import dataclass

@dataclass
class TableConfig:
    TableName: str
    SheetName: str
    StartIndex: int
    EndIndex: int
    coords: str

sheetList = ['Datos del proyecto', '1. Generación diaria', '1.1. Proyección de Gen.']

dataClass = {
    'T1': {
        'TableName': 'EnergyInjected',
        'SheetName': '1.2. Relaciones Gen.',
        'StartIndex': 1,
        'EndIndex': 14,
        'coords': 'H'
    },
    'T2': {
        'TableName': 'EnergyCon',
        'SheetName': '1.2. Relaciones Gen.',
        'StartIndex': 1,
        'EndIndex': 14,
        'coords': 'E'
    },
    'T3': {
        'TableName': 'CL',
        'SheetName': '2. Protecciones SC',
        'StartIndex': 1,
        'EndIndex': 9,
        'coords': 'G'
    },
    'T4': {
        'TableName': 'Wc',
        'SheetName': '2. Protecciones SC',
        'StartIndex': 1,
        'EndIndex': 9,
        'coords': 'H'
    },
    'T5': {
        'TableName': 'MaxC',
        'SheetName': '3. Conductores',
        'StartIndex': 3,
        'EndIndex': 8,
        'coords': 'C'
    },
    'T6': {
        'TableName': 'Sc',
        'SheetName': '3. Conductores',
        'StartIndex': 3,
        'EndIndex': 8,
        'coords': 'D'
    },
    'T7': {
        'TableName': 'Wn',
        'SheetName': '3. Conductores',
        'StartIndex': 3,
        'EndIndex': 8,
        'coords': 'F'
    },
    'T8': {
        'TableName': 'Cpw',
        'SheetName': '3. Conductores',
        'StartIndex': 3,
        'EndIndex': 8,
        'coords': 'G'
    },
    'T9': {
        'TableName': 'cac',
        'SheetName': '3. Conductores',
        'StartIndex': 3,
        'EndIndex': 11,
        'coords': 'K'
    },
    'T10': {
        'TableName': 'an',
        'SheetName': '3. Conductores',
        'StartIndex': 3,
        'EndIndex': 11,
        'coords': 'L'
    },
    'T11': {
        'TableName': 'mac',
        'SheetName': '3. Conductores',
        'StartIndex': 3,
        'EndIndex': 11,
        'coords': 'N'
    },
    'T12': {
        'TableName': 'pht',
        'SheetName': '6.1 Resumen Tuberías',
        'StartIndex': 1,
        'EndIndex': 7,
        'coords': 'D'
    },
    'T13': {
        'TableName': 'nt',
        'SheetName': '6.1 Resumen Tuberías',
        'StartIndex': 1,
        'EndIndex': 7,
        'coords': 'E'
    },
    'T14': {
        'TableName': 'gndt',
        'SheetName': '6.1 Resumen Tuberías',
        'StartIndex': 1,
        'EndIndex': 7,
        'coords': 'F'
    },
    'T15': {
        'TableName': 'ww',
        'SheetName': '6.1 Resumen Tuberías',
        'StartIndex': 1,
        'EndIndex': 7,
        'coords': 'G'
    },
    'T16': {
        'TableName': 'dinch',
        'SheetName': '6.1 Resumen Tuberías',
        'StartIndex': 1,
        'EndIndex': 7,
        'coords': 'H'
    },
    'T17': {
        'TableName': 'stt',
        'SheetName': '6.1 Resumen Tuberías',
        'StartIndex': 1,
        'EndIndex': 7,
        'coords': 'I'
    },
    'T18': {
        'TableName': 'w',
        'SheetName': '6.1 Resumen Tuberías',
        'StartIndex': 1,
        'EndIndex': 7,
        'coords': 'J'
    },
    'T19': {
        'TableName': 'lac',
        'SheetName': '4. Regulación AC',
        'StartIndex': 3,
        'EndIndex': 12,
        'coords': 'AK'
    },
    'T20': {
        'TableName': 'rac',
        'SheetName': '4. Regulación AC',
        'StartIndex': 3,
        'EndIndex': 12,
        'coords': 'AL'
    },
    'T21': {
        'TableName': 'crac',
        'SheetName': '4. Regulación AC',
        'StartIndex': 3,
        'EndIndex': 12,
        'coords': 'AI'
    },
    'T22': {
        'TableName': 'ldc',
        'SheetName': '5. Regulación DC',
        'StartIndex': 1,
        'EndIndex': 6,
        'coords': 'Y'
    },
    'T23': {
        'TableName': 'rdc',
        'SheetName': '5. Regulación DC',
        'StartIndex': 1,
        'EndIndex': 6,
        'coords': 'Z'
    },
    'T24': {
        'TableName': 'dl',
        'SheetName': '7. Perdidas',
        'StartIndex': 1,
        'EndIndex': 10,
        'coords': 'D'
    },
    'T25': {
        'TableName': 'awg',
        'SheetName': '7. Perdidas',
        'StartIndex': 1,
        'EndIndex': 10,
        'coords': 'F'
    },
    'T26': {
        'TableName': 'rw',
        'SheetName': '7. Perdidas',
        'StartIndex': 1,
        'EndIndex': 10,
        'coords': 'G'
    },
    'T27': {
        'TableName': 'lwc',
        'SheetName': '7. Perdidas',
        'StartIndex': 1,
        'EndIndex': 10,
        'coords': 'H'
    },
    'T28': {
        'TableName': 'ptw',
        'SheetName': '7. Perdidas',
        'StartIndex': 1,
        'EndIndex': 11,
        'coords': 'I'
    },
    'T29': {
        'TableName': 'i300',
        'SheetName': '8. Verificación de parámetros',
        'StartIndex': 1,
        'EndIndex': 9,
        'coords': 'C'
    },
    'T30': {
        'TableName': 'i50',
        'SheetName': '8. Verificación de parámetros',
        'StartIndex': 1,
        'EndIndex': 9,
        'coords': 'D'
    },
    'T31': {
        'TableName': 'i40',
        'SheetName': '8. Verificación de parámetros',
        'StartIndex': 1,
        'EndIndex': 9,
        'coords': 'E'
    },
    'T32': {
        'TableName': 'ttt',
        'SheetName': '8. Verificación de parámetros',
        'StartIndex': 1,
        'EndIndex': 9,
        'coords': 'F'
    }
}
