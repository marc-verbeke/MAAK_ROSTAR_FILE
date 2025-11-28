import pandas as pd


pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.expand_frame_repr', False)

bestand = "20260105 Rittenbestand Waaslandia P2.xlsX"
bestand2 = bestand[:-5]

# Excelbestand openen
file_name = fr"DATA\RITTENBESTANDEN_DE_LIJN\{bestand}"
excel_data = pd.ExcelFile(file_name, engine='openpyxl')

# Bepaal de datum vanaf wanneer de diensten geldig zijn
datum_vanaf = bestand[:8]

# Lege DataFrame
df = pd.DataFrame()

# Gegevens uit alle tabbladen combineren
for sheet_name in excel_data.sheet_names:
    sheet_df = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl')
    sheet_df['SheetName'] = sheet_name
    df = pd.concat([df, sheet_df], ignore_index=True)

columns_to_drop = [
    'Dty start time', 'Blk number', 'Priority?', 'Opportunity?', 'Public?', 'Direction', 'Route', 'Number',
    'Start place', 'End place', 'Trp end min lay', 'Duration', 'Prv Route', 'Nxt Route',
    'Nxt Start place', 'Nxt Start time', 'Note', 'Note 2', 'Rte desc',
    'Vehicle disp', 'Rte pub Id', 'Vdc message 1', 'Vdc message 2',
    'Vdc message 3', 'Vdc message 4', 'Vehicle group', 'Gar phone id'
]
df = df.drop(columns=columns_to_drop, errors='ignore')

# kolom Sleutel
df['Sleutel'] = df['Dty number'] + "+" + df['SheetName']

# kolom DienstNr
df['DienstNr'] = pd.to_numeric(df['Dty number'].str[4:], errors='coerce').astype('Int64')

# kolom Deel
df['Deel'] = 0
for DIENST, group in df.groupby('Sleutel'):
    group['deel'] = (group['Type'] == 'OUT').cumsum()
    df.loc[group.index, 'Deel'] = group['deel']

# Sleutel vervolledigen
df['Sleutel'] = df['Sleutel'] + "+" + df['Deel'].astype(str)

# kolom Geldigheid & Geldigheidsnummer
df['Geldigheidsnummer'] = df['SheetName'].str[2:]
geldigheid_map = {
    '1': 'm____',
    '2': '_d___',
    '3': '__w__',
    '4': '___d_',
    '5': '____v',
    '6': 'z_',
    '7': '_z'
}
df['Geldigheid'] = df['Geldigheidsnummer'].map(geldigheid_map)

# kolom Periodeid & Periode
df['PeriodeId'] = df['SheetName'].str[:-1]
periode_map = {
    'P2': 'School',
    'P3': 'Schoolverlof',
    'P5': 'Examenperiode',
    'P9': 'Examenperiode'
}
df['Periode'] = df['PeriodeId'].map(periode_map)

# kolom Stelplaats
stelplaats_map = {
    '268': 'KAV',
    '286': 'KAV',
    '278': 'Waaslandia',
    '276': 'Kruger',
}
df['Stelplaats'] = df['Dty number'].str[:3].map(stelplaats_map)

# kolom ContractId
df = df.rename(columns={'Blk contract': 'ContractId'})

# kolom BeginDatum
df['BeginDatum'] = pd.to_datetime('01/07/2025', format='%d/%m/%Y')
df['BeginDatum'] = df['BeginDatum'].dt.strftime('%d/%m/%Y')

# kolom WijzigDatum
df['WijzigDatum'] = pd.to_datetime(datum_vanaf, format='%Y%m%d')
df['WijzigDatum'] = df['WijzigDatum'].dt.strftime('%d/%m/%Y')

# kolom Einddatum
df['Einddatum'] = ""

# kolom Bustype
bustype_map = {
    2: 'S',
    4: 'G',
}
df['Bustype'] = (df['Blk vehicle group']//10).map(bustype_map)

# kolom Opdrachtgever
df['Opdrachtgever'] = ""

# kolom Doelgroep
df['Doelgroep'] = ""

# kolom Vertrektijd
df['Vertrektijd'] = df['Start time']
#.apply(lambda x: f"{x.components.hours:02}:{x.components.minutes:02}"))

# kolom VertrektijdInMin
df['VertrektijdInMin'] = df['Start time'].apply(lambda x: x.components.days * 1440 + x.components.hours * 60 + x.components.minutes)

# kolom Aankomsttijd
df['Aankomsttijd'] = df['End time']
#.apply(lambda x: f"{x.components.hours:02}:{x.components.minutes:02}"))

# kolom AankomsttijdInMin
df['AankomsttijdInMin'] = df['End time'].apply(lambda x: x.components.days * 1440 + x.components.hours * 60 + x.components.minutes)

# OPMERKING IVM KM : WE STELLEN DE BELADEN EN DE THEORETISCHE KM AAN ELKAAR GELIJK
# kolom Beladen Kms
df['Beladen Kms'] = df['Distance'].astype(float).round(3)

# kolom Theoretische Kms
df['Theoretische Kms'] = df["Distance"].astype(float).round(3)

# kolom Dag1Rijtijd & Dag2Rijtijd
df['Dag1Rijtijd'] = 0
df['Dag2Rijtijd'] = 0
kolom1 = df.columns.get_loc('Dag1Rijtijd')
kolom2 = df.columns.get_loc('Dag2Rijtijd')
for index, row in df.iterrows():
    if row['Start time'].days == 0 and row['End time'].days == 0:
        df.iloc[index, kolom1] = int((row['End time'] - row['Start time']).total_seconds() / 60)
    if row['Start time'].days == 1 and row['End time'].days == 1:
        df.iloc[index, kolom2] = int((row['End time'] - row['Start time']).total_seconds() / 60)
    if row['Start time'].days == 0 and row['End time'].days == 1:
        middernacht = pd.Timedelta(days=1)
        df.iloc[index, kolom1] = int((middernacht - row['Start time']).total_seconds() / 60)
        middernacht = pd.Timedelta(days=1)
        df.iloc[index, kolom2] = int((row['End time'] - middernacht).total_seconds() / 60)

# kolom Dag1StatTb100 & Dag1Stat50 & Dag2StatTB100 & Dag2Stat50
df['Dag1StatTb100'] = 0
df['Dag1Stat50'] = 0
df['Dag1AndWerkz'] = 0
df['Dag1Arbeidstijd'] = 0
df['Dag1Nacht'] = 0
df['Dag2StatTb100'] = 0
df['Dag2Stat50'] = 0
df['Dag2AndWerkz'] = 0
df['Dag2Arbeidstijd'] = 0
df['Dag2Nacht'] = 0
df['StatEff'] = 0
deel1 = 0
deel2 = 0
deel3 = 0
deel4 = 0
deel5 = 0
deel6 = 0

for index in range(len(df) - 1):
    deel1 = 0
    deel2 = 0
    deel3 = 0
    deel4 = 0
    deel5 = 0
    deel6 = 0
    if df.loc[index + 1, 'Start time'].days == 0 and df.loc[index, 'End time'].days == 0:
        if df.loc[index, 'Type'] != 'IN':
            einde = df.loc[index, 'End time']
            start = df.loc[index + 1, 'Start time']
            stationnement = int((start - einde).total_seconds() / 60)
            if stationnement <= 15:
                deel1 = stationnement
                deel2 = 0
                deel3 = 0
            elif stationnement <= 45:
                deel1 = 15
                deel2 = stationnement - 15
                deel3 = 0
            else:
                deel1 = 15
                deel2 = 30
                deel3 = stationnement - 45
            df.loc[index, 'Dag1StatTb100'] = deel1 + deel2
            df.loc[index, 'Dag1Stat50'] = deel3
    if df.loc[index + 1, 'Start time'].days == 1 and df.loc[index, 'End time'].days == 1:
        if df.loc[index, 'Type'] != 'IN':
            einde = df.loc[index, 'End time']
            start = df.loc[index + 1, 'Start time']
            stationnement = int((start - einde).total_seconds() / 60)
            if stationnement <= 15:
                deel4 = stationnement
                deel5 = 0
                deel6 = 0
            elif stationnement <= 45:
                deel4 = 15
                deel5 = stationnement - 15
                deel6 = 0
            else:
                deel4 = 15
                deel5 = 30
                deel6 = stationnement - 45
            df.loc[index, 'Dag2StatTb100'] = deel4 + deel5
            df.loc[index, 'Dag2Stat50'] = deel6
    if df.loc[index + 1, 'Start time'].days == 1 and df.loc[index, 'End time'].days == 0:
        middernacht = pd.Timedelta(days=1)
        if df.loc[index, 'Type'] != 'IN':
            einde = df.loc[index, 'End time']
            start = df.loc[index + 1, 'Start time']
            stationnement = int((middernacht - einde).total_seconds() / 60)
            if stationnement <= 15:
                deel1 = stationnement
                deel2 = 0
                deel3 = 0
            elif stationnement <= 45:
                deel1 = 15
                deel2 = stationnement - 15
                deel3 = 0
            else:
                deel1 = 15
                deel2 = 30
                deel3 = stationnement - 45
            df.loc[index, 'Dag1StatTb100'] = deel1 + deel2
            df.loc[index, 'Dag1Stat50'] = deel3

            stationnement = int((start - middernacht).total_seconds() / 60)
            if stationnement <= 15:
                deel4 = stationnement
                deel5 = 0
                deel6 = 0
            elif stationnement <= 45:
                deel4 = 15
                deel5 = stationnement - 15
                deel6 = 0
            else:
                deel4 = 15
                deel5 = 30
                deel6 = stationnement - 45
            df.loc[index, 'Dag2StatTb100'] = deel4 + deel5
            df.loc[index, 'Dag2Stat50'] = deel6

    df.loc[index, 'Dag1Arbeidstijd'] = df.loc[index, 'Dag1Rijtijd'] + deel1
    df.loc[index, 'Dag2Arbeidstijd'] = df.loc[index, 'Dag2Rijtijd'] + deel4

    if df.loc[index, 'Type'] == 'OUT':
        if df.loc[index, 'Start time'] < pd.to_timedelta("6:00:00"):
            df.loc[index, 'Dag1Nacht'] = (pd.to_timedelta("6:00:00") - df.loc[index, 'Start time']).total_seconds() / 60

    if df.loc[index, 'Type'] == 'IN':
        if df.loc[index, 'End time'] > pd.to_timedelta("20:00:00"):
            if df.loc[index, 'End time'] < pd.Timedelta(days=1):
                df.loc[index, 'Dag1Nacht'] = (df.loc[index, 'End time'] - pd.to_timedelta("20:00:00")).total_seconds()/60
            else:
                df.loc[index, 'Dag1Nacht'] = ( pd.Timedelta(days=1) - pd.to_timedelta("20:00:00")).total_seconds()/60
                df.loc[index, 'Dag2Nacht'] = (df.loc[index, 'End time'] - pd.Timedelta(days=1)).total_seconds()/60

    df.loc[index, 'StatEff'] = df.loc[index, 'Dag1StatTb100'] + df.loc[index, 'Dag1Stat50'] + df.loc[index, 'Dag2StatTb100'] + df.loc[index, 'Dag2Stat50']

# kolom OndEff
df['OndEff'] = 0

# kolom Kleur
kleur_map = {
    'School': -1,
    'Schoolverlof': -256,
    'Examenperiode': -32640,
}
df['Kleur'] = df['Periode'].map(kleur_map)

# kolom Stelsel
df['Stelsel'] = 'Bus'

# kolom Toegangsgebied
toegangsgebied_map = {
    'Waaslandia': 'Vrasene',
    'KAV': 'Geel',
    'Kruger': 'Lier',
}
df['Toegangsgebied'] = df['Stelplaats'].map(toegangsgebied_map)

# kolom Basisbezetting
geldigheid_map = {
    '1': 'Def:0,Wkhd:0,Mo:1',
    '2': 'Def:0,Wkhd:0,Tu:1',
    '3': 'Def:0,Wkhd:0,We:1',
    '4': 'Def:0,Wkhd:0,Th:1',
    '5': 'Def:0,Wkhd:0,Fr:1',
    '6': 'Def:0,Wkhd:0,Sa:1',
    '7': 'Def:0,Wkhd:1,Wdhd:1,Su:1'
}
df['Basisbezetting'] = df['Geldigheidsnummer'].map(geldigheid_map)

# kolom Kostenplaats
kostenplaats_map = {
    'Waaslandia': 'W000001',
    'KAV': 'C000001',
    'Kruger': 'K000001',
}
df['Kostenplaats'] = df['Stelplaats'].map(kostenplaats_map)

# kolom KalenderKoppeling
KalenderKoppeling_map = {
    'Waaslandia': 12011,
    'KAV': 52011,
    'Kruger': 22011,
}
df['KalenderKoppeling'] = df['Stelplaats'].map(KalenderKoppeling_map)

# kolom StartLine
df['StartLine'] = ''

# kolom EndLine
df['EndLine'] = ''

df_csv = df.groupby(['Sleutel']).agg(
    DienstNr=('DienstNr', 'first'),
    Geldigheid=('Geldigheid', 'first'),
    Deel=('Deel', 'first'),
    PeriodeId=('PeriodeId', 'first'),
    Stelplaats=('Stelplaats', 'first'),
    Geldigheidsnummer=('Geldigheidsnummer', 'first'),
    Periode=('Periode', 'first'),
    ContractId=('ContractId', 'first'),
    BeginDatum=('BeginDatum', 'first'),
    WijzigDatum=('WijzigDatum', 'first'),
    Einddatum=('Einddatum', 'first'),
    Bustype=('Bustype', 'first'),
    Opdrachtgever=('Opdrachtgever', 'first'),
    Doelgroep=('Doelgroep', 'first'),
    VertrektijdInMin=('VertrektijdInMin', 'min'),
    Vertrektijd=('Vertrektijd', 'min'),
    Aankomsttijd=('End time', 'max'),
    AankomsttijdInMin=('AankomsttijdInMin', 'max'),
    Beladen_Kms=('Beladen Kms', 'sum'),
    Theoretische_Kms=('Theoretische Kms', 'sum'),
    Dag1Rijtijd=('Dag1Rijtijd', 'sum'),
    Dag1StatTb100=('Dag1StatTb100', 'sum'),
    Dag1Stat50=('Dag1Stat50', 'sum'),
    Dag1AndWerkz=('Dag1AndWerkz', 'sum'),
    Dag1Arbeidstijd=('Dag1Arbeidstijd', 'sum'),
    Dag1Nacht=('Dag1Nacht', 'sum'),
    Dag2Rijtijd=('Dag2Rijtijd', 'sum'),
    dag2StatTb100=('Dag2StatTb100', 'sum'),
    Dag2Stat50=('Dag2Stat50', 'sum'),
    dag2AndWerkz=('Dag2AndWerkz', 'sum'),
    dag2Arbeidstijd=('Dag2Arbeidstijd', 'sum'),
    Dag2Nacht=('Dag2Nacht', 'sum'),
    StatEff=('StatEff', 'sum'),
    OndEff=('OndEff', 'sum'),
    Kleur=('Kleur', 'first'),
    Stelsel=('Stelsel', 'first'),
    Toegangsgebied=('Toegangsgebied', 'first'),
    Basisbezetting=('Basisbezetting', 'first'),
    Kostenplaats=('Kostenplaats', 'first'),
    KalenderKoppeling=('KalenderKoppeling', 'first'),
    StartLine=('StartLine', 'first'),
    Endline=('EndLine', 'first')
).reset_index()
df_csv = df_csv.drop('Sleutel', axis=1)

# kolom Vertrektijd
df_csv['Vertrektijd'] = df_csv['Vertrektijd'].apply(
    lambda x: f"{x.components.hours:02}:{x.components.minutes:02}"
)

# kolom Aankomsttijd
df_csv['Aankomsttijd'] = df_csv['Aankomsttijd'].apply(
    lambda x: f"{x.components.hours:02}:{x.components.minutes:02}" + (">" if x.components.days == 1 else "")
)

# kolom Beladen_Kms
df_csv['Beladen_Kms'] = df_csv['Beladen_Kms'].round(0).astype(int)

# kolom Theoretische_Kms
df_csv['Theoretische_Kms'] = df_csv['Theoretische_Kms'].round(0).astype(int)

df_csv = df_csv.sort_values(by=['PeriodeId', 'DienstNr', 'Geldigheidsnummer'])

for periode_id, groep in df_csv.groupby('PeriodeId'):
    bestandsnaam = rf"DATA/INLEESBESTAND_ROSTAR/{bestand2}_{periode_id}.csv"
    groep.to_csv(bestandsnaam, sep=';', encoding='utf-8-sig', index=False)

