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

headers = [
    "TypeTransaction", "Societe", "Article", "Article (Désignation)",
    "Fournisseur", "Fournisseur (Désignation)", "Reference", "Designation",
    "DelaiLivraison", "Prix", "PrixDevise", "Remise", "Tva",
    "Code tva (Désignation)", "Unite", "Unité d'achat (Désignation)",
    "CoefficientUnite", "Principal", "AvertPrix", "Date création",
    "Utilisateur Création", "Date modification", "Utilisateur Modification",
    "SituationTransaction", "Message"
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
    "SituationTransaction",
    "Message"
]



try:
    # Connect to the database
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    print("Connection successful using SQL Server Authentication!")

    # Execute a query
    # cursor.execute("""
    #     SELECT 
    #         i.ITMREF_0,
    #         i.ITMDES1_0,
    #         i.ITMSTD_0,
    #         m.AVC_0,
    #         i.STU_0,
    #         i.TCLCOD_0,
    #         i.TSICOD_0

    #     FROM NPAPR2.ITMMASTER i
    #     OUTER APPLY (
    #         SELECT TOP 1 AVC_0
    #         FROM NPAPR2.ITMMVT m
    #         WHERE m.ITMREF_0 = i.ITMREF_0
    #         ORDER BY m.UPDDATTIM_0 ASC
    #     ) m
    #     WHERE i.CPY_0 = 'STD'
    #     AND i.ITMSTA_0 = 1
    #     ORDER BY i.ITMREF_0 ASC
    # """)    

    cursor.execute(
        """
            SELECT
            ITM.ITMREF_0,            
            ITM.ITMDES1_0,            
            ITM.ITMSTD_0,      
            MVT.AVC_0,     
            ITM.STU_0,           
            ITM.TCLCOD_0,           
            ITM.TSICOD_0               
   
            FROM NPAPR2.ITMMASTER ITM
            LEFT JOIN NPAPR2.ITMMVT MVT
                ON ITM.ITMREF_0 = MVT.ITMREF_0

            WHERE ITM.ITMSTA_0 = 1
            AND ITM.CPY_0 = 'STD'

            ORDER BY ITM.ITMREF_0

        """
    )
    
    rows = cursor.fetchall()

    
    max_rows_per_file = 1500
    total_rows = len(rows)
    file_count = ceil(total_rows / max_rows_per_file)

    for file_index in range(file_count):

        wb = Workbook()
        ws = wb.active
        ws.title = "LiensArticlesFournisseurs"
        ws.append(headers)

        start = file_index * max_rows_per_file
        end = start + max_rows_per_file
        chunk = rows[start:end]

        for row in chunk:
            article = row.ITMREF_0
            designation = row.ITMDES1_0
            reference = row.ITMSTD_0
            unite = row.STU_0
            prix = round(row.AVC_0 * Decimal(1.1), 5) if row.AVC_0 else 0

            ws.append([
                3,
                current_cpy,
                article,
                designation,
                current_supplier,
                current_supplier_name,
                reference,
                "",
                delivery_day,
                prix,
                "",
                "",
                "",
                "",
                unite,
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                ""
            ])

        filename = f"ITMMASTER_LIENS_STD_Part{file_index+1}.xlsx"
        wb.save(filename)
        print(f"Fichier {filename} généré ({len(chunk)} lignes)")

    for file_index in range(file_count):
        wb_articles = Workbook()
        ws_articles = wb_articles.active
        ws_articles.title = "Articles"

        # Header
        ws_articles.append(headers_articles)

        start = file_index * max_rows_per_file
        end = start + max_rows_per_file

        chunk = rows[start:end]

        for row in chunk:
            article = row.ITMREF_0
            designation = row.ITMDES1_0
            unite = row.STU_0
            famille = row.TSICOD_0
            reference = row.ITMSTD_0
            groupe = row.TCLCOD_0
            prix = round(row.AVC_0 * Decimal(1.1), 5) if row.AVC_0 else 0

            ws_articles.append([
                current_cpy,      # Societe
                1,                # TypeTransaction (création article)
                article,          # Article
                designation,      # Designation
                "MAG01",          # Magasin 
                "",               # Emplacement
                0,               # NonStocke
                unite,            # Unite
                famille, "", groupe, "", "", reference,  # Famille → ReferenceFabricant
                delivery_day,     # DelaiLivraison
                "", "", "", "", "",      # Reappro → QuantiteReappro
                prix, "", "", "", "", "",  # PrixStandard → DateObsolescence
                "", "", "", "", "", "",  # Quantite → Image
                "", "", "", "", "", "",  # MajAutoPointCommande → CoeffCriticite
                "", ""                    # SituationTransaction, Message
            ])

        filename = f"ITMMASTER_ARTICLES_STD_Part{file_index+1}.xlsx"
        wb_articles.save(filename)

        print(f"Fichier {filename} généré ({len(chunk)} lignes)")

except pyodbc.Error as ex:
    print(f"An error occurred: {ex}")

finally:
    # Close connection
    if 'conn' in locals() and conn:
        conn.close()
