import pandas as pd
import csv

liste_diplomes_excel = 'Liste des diplômés RDD.xlsx'
liste_framaform_excel = '2024-05-29-11h - frama.xlsx'
liste_billetterie_excel = '2024-05-29-11h - billetterie.xlsx'

df_diplomes = pd.read_excel(liste_diplomes_excel)
df_framaform = pd.read_excel(liste_framaform_excel)
df_billetterie = pd.read_excel(liste_billetterie_excel)

etunum_column_diplomes = 'Etunum'
etunum_column_framaform = "EtuNum (numéro étudiant présent sur la carte d'étudiant ou sur Géode) / EtuNum (student number on student card or Géode)"
etunum_column_billetterie = 'etunum'

etunum_columns = {
    'diplomes': 'Etunum',
    'framaform': "EtuNum (numéro étudiant présent sur la carte d'étudiant ou sur Géode) / EtuNum (student number on student card or Géode)",
    'billetterie': 'etunum',
}

def sanitize_etunums(df, df_type):
    etunum_column = etunum_columns[df_type]
    problemes_etunum = []
    etunums_ok = []

    for index, row in df.iterrows():
        etunum = str(row[etunum_column])
        if etunum == "INCONNU" or len(etunum) < 5 or len(etunum) > 7:
            problemes_etunum.append(row)
            continue
        elif len(etunum) == 6:
            row[etunum_column] = etunum[0:2]+'0' + etunum[2:]
        elif len(etunum) == 5:
            row[etunum_column] = etunum[0:2]+'00' + etunum[2:]
        else:
            row[etunum_column] = etunum
        etunums_ok.append(row)

    df_problemes_etunum = pd.DataFrame(problemes_etunum)
    df_etunums_ok = pd.DataFrame(etunums_ok)
    df_problemes_etunum.to_csv(f'problemes_etunum/problemes_etunum_{df_type}.csv', index=False)
    return df_etunums_ok

df_diplomes = sanitize_etunums(df_diplomes, 'diplomes')
df_billetterie = sanitize_etunums(df_billetterie, 'billetterie')
df_framaform = sanitize_etunums(df_framaform, 'framaform')

# Check for etunums not present in df_framaform or df_billetterie
pas_inscrit = df_diplomes[~df_diplomes[etunum_column_diplomes].isin(df_framaform[etunum_column_framaform]) & ~df_diplomes[etunum_column_diplomes].isin(df_billetterie[etunum_column_billetterie])]
pas_inscrit.to_csv('results/pas_inscrit.csv', index=False)

# Check for etunums present in df_framaform but not in df_billetterie
frama_but_not_billetterie = df_framaform[~df_framaform[etunum_column_framaform].isin(df_billetterie[etunum_column_billetterie])]
frama_but_not_billetterie.to_csv('results/frama_but_not_billetterie.csv', index=False)