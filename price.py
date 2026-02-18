import pyodbc
import csv
from decimal import Decimal
from openpyxl import Workbook
from math import ceil

# Connection string for SQL Server Authentication
conn_str = (
    'DRIVER={SQL Server};'
    'SERVER=SAGEX3-SQL\\X3V12;'
    'DATABASE=x3db;'
    'UID=CR;'
    'PWD=Tiger;'
)

current_cpy = 'SE'
current_supplier = 'GS00009'
current_supplier_name= 'STD SARL'
delivery_day= 1
magasin= "MAG01"
tva= "NOR"


headers_liens_fournisseurs = [
    "Societe",
    "TypeTransaction",
    "Article",
    "Fournisseur",
    "Reference",
    "Designation",
    "DelaiLivraison",
    "Prix",
    "PrixDevise",
    "Remise",
    "Tva",
    "Unite",
    "CoefficientUnite",
    "Principal",
    "AvertPrix",
    "FraisApprocheFamille",
    "SituationTransaction"
]

headers_articles = [
    "Societe",
    "TypeTransaction",
    "Article",
    "Designation",
    "Magasin",
    "Emplacement",
    "NonStocke",
    "Unite",
    "Famille",
    "SousFamille",
    "Groupe",
    "Tva",
    "Marque",
    "ReferenceFabricant",
    "DelaiLivraison",
    "Reappro",
    "StockMaxi",
    "PointCommande",
    "StockMini",
    "QuantiteReappro",
    "PrixStandard",
    "PrixDerniereCommande",
    "DateDerniereEntree",
    "DateDerniereSortie",
    "DateDernierInventaire",
    "DateObsolescence",
    "Quantite",
    "Pmp",
    "PmpHorsFrais",
    "EditionConsigneSecurite",
    "Image",
    "MajAutoPointCommande",
    "OptionReferenceFouAuto",
    "Commentaire",
    "CoeffCriticitePointCommande",
    "SituationTransaction"
]

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    print("Connection successful using SQL Server Authentication!")

 
    # Récupération SPRICLIST
   
    cursor.execute("""
        SELECT PRI.PLICRI1_0, PRI.PRI_0 
        FROM NPAPR2.SPRICLIST PRI
        WHERE PLI_0 = 'ZTSTD'
        AND PLICRD_0 = 'FTV-STD-V3-2026-GPE'
    """)

    price_rows = cursor.fetchall()

    coeff_dict = {}

    for row in price_rows:
        coeff_dict[row.PLICRI1_0] = row.PRI_0

    print(f"Dictionnaire SPRICLIST chargé : {len(coeff_dict)} lignes")
    print(coeff_dict)


    # Requête principale

    cursor.execute(
        """
            SELECT
            ITM.ITMREF_0,            
            ITM.ITMDES1_0,            
            ITM.ITMSTD_0,      
            MVT.AVC_0,    
            MVT.STOFCY_0, 
            ITM.STU_0,           
            ITM.TCLCOD_0,           
            ITM.TSICOD_0               
   
            FROM NPAPR2.ITMMASTER ITM
            LEFT JOIN NPAPR2.ITMMVT MVT
                ON ITM.ITMREF_0 = MVT.ITMREF_0

            WHERE ITM.ITMSTA_0 = 1
            AND ITM.CPY_0 = 'STD'
            AND MVT.STOFCY_0 = 'ESTD1'

            ORDER BY ITM.ITMREF_0
        """
    )
    
    rows = cursor.fetchall()

    # ======= Fichier Liens Articles Fournisseurs =======
    filename = f"LiensArticlesFournisseurs.csv"

    with open(filename, mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(headers_liens_fournisseurs)

        for row in rows:
            article = row.ITMREF_0
            designation = row.ITMDES1_0
            reference = row.ITMSTD_0
            unite = row.STU_0
            # prix = round(row.AVC_0 * Decimal(1.1), 5) if row.AVC_0 else 0
            prix = round(row.AVC_0 * coeff_dict[row.TCLCOD_0], 5) if row.AVC_0 else 0

            writer.writerow([
                current_cpy,
                3,
                article,
                current_supplier,
                reference,
                designation,
                delivery_day,
                prix,
                "",
                "",
                tva,
                unite,
                "",
                "",
                "",
                "",
                "" 
            ])

    print(f"Fichier {filename} généré ({len(rows)} lignes)")

    # ======= Fichier Articles =======
    filename_articles = "Articles.csv"
    with open(filename_articles, mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(headers_articles)

        for row in rows:
            article = row.ITMREF_0
            designation = row.ITMDES1_0
            unite = row.STU_0
            prix = round(row.AVC_0 * coeff_dict[row.TCLCOD_0], 5) if row.AVC_0 else 0
            famille = row.TSICOD_0
            groupe = row.TCLCOD_0
            reference = row.ITMSTD_0

            writer.writerow([
                current_cpy,
                1,
                article,
                designation,
                magasin,
                "",
                0,
                unite,
                famille, "", groupe, tva, "", reference, delivery_day,
                "", "", "", "", "", prix,
                "", "", "", "", "", "", "", "", "", "", "", "", ""
            ])

    print(f"Fichier {filename_articles} généré ({len(rows)} lignes)")

except pyodbc.Error as ex:
    print(f"An error occurred: {ex}")

finally:
    if 'conn' in locals() and conn:
        conn.close()
