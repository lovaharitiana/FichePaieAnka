import os
import pandas as pd
import xlwings as xw
import math

def load_first_sheet(input_file):
    try:
        xls = pd.ExcelFile(input_file, engine='openpyxl')
        df_sud = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=11)
        print(df_sud.columns)


        app = xw.App(visible=False)
        wb = app.books.open(input_file)
        ws = wb.sheets[0]
        
        k_values = ws.range('K10').value
        h10_value = ws.range('H10').value
        i10_value = ws.range('I10').value
        f15_value = ws.range('F15').value
        g15_value = ws.range('G15').value
        
        # Extraction des valeurs J15 à Q15
        j15_value = ws.range('J15').value
        k15_value = ws.range('K15').value
        l15_value = ws.range('L15').value
        m15_value = ws.range('M15').value
        n15_value = ws.range('N15').value
        o15_value = ws.range('O15').value
        p15_value = ws.range('P15').value
        q15_value = ws.range('Q15').value
        
        
        wb.close()
        app.quit()
        
        combined_values = f"{k_values}".strip()
        print(df_sud.columns)
        print(df_sud.head())  
        return df_sud, combined_values, h10_value, i10_value, f15_value, g15_value, j15_value, k15_value, l15_value, m15_value, n15_value, o15_value, p15_value, q15_value
    except Exception as e:
        print(f"Erreur lors du chargement de la feuille : {e}")
        return pd.DataFrame(), '', None, None, None, None, None, None, None, None, None, None, None, None

def extract_values(df_sud):
    if df_sud.empty:
        return '', '', '', 0  

    du_text_parts = []
    for col in range(10, 13):
        if col < len(df_sud.columns):
            cell_value = df_sud.iloc[0, col]
            if pd.notna(cell_value):
                du_text_parts.append(str(cell_value).strip())
    du_text = ' '.join(du_text_parts).strip()

    mois = df_sud.iloc[0, 7] if len(df_sud.columns) > 7 else ''
    annee = df_sud.iloc[0, 8] if len(df_sud.columns) > 8 else ''
    
    if len(df_sud) > 5:
        montant_du_mois = df_sud.iloc[5, 5] * df_sud.iloc[5, 6] if len(df_sud.columns) > 6 else 0
    else:
        montant_du_mois = 0  

    return du_text, mois, annee, montant_du_mois

def calculate_r15_s15_u15(df_sud):
    try:
        if len(df_sud) > 5 and len(df_sud.columns) > 16:
            F15 = df_sud.iloc[5, 5]
            G15 = df_sud.iloc[5, 6]
            J15_to_Q15 = df_sud.iloc[5, 9:17].sum()

            H15 = F15 * G15
            R15 = H15 + J15_to_Q15
            S15 = min(R15 * 0.01, 20000)
            U15 = S15
            W15 = R15 - S15 - U15

            return R15, S15, U15, W15
        else:
            return None, None, None, None
    except Exception as e:
        print(f"Erreur lors du calcul de R15, S15, U15 et W15 : {e}")
        return None, None, None, None

def calculate_y15(R15, S15, U15):
    try:
        if R15 is not None and S15 is not None and U15 is not None:
            X15 = (R15 - S15 - U15) // 100 * 100
            if X15 <= 400000:
                Y15 = max(5 * (X15 - 350000) / 100, 3000)
            elif X15 <= 500000:
                Y15 = (X15 - 400000) * 10 / 100 + 2500
            elif X15 <= 600000:
                Y15 = (X15 - 500000) * 15 / 100 + 12500
            else:
                Y15 = (X15 - 600000) * 20 / 100 + 27500
            return Y15
        else:
            return None
    except Exception as e:
        print(f"Erreur lors du calcul de Y15 : {e}")
        return None

def sum_e20_to_e34(ws):
    try:
        e_range_values = ws.range('E20:E34').value
        if isinstance(e_range_values, list):
            # Convert all values to numeric and sum them up
            total_sum = sum([float(value) for value in e_range_values if isinstance(value, (int, float))])
            return total_sum
        else:
            print("La plage de cellules E20:E34 n'est pas sous forme de liste.")
            return 0
    except Exception as e:
        print(f"Erreur lors du calcul de la somme pour E35 : {e}")
        return 0



def fill_template(row, template_path, output_path, du_value, mois, annee, montant_total, s15_value, w15_value, y15_value, combined_values, h10_value, i10_value, calculated_value, avance_quinzaine):
    try:
        app = xw.App(visible=False)
        wb_template = app.books.open(template_path)
        ws_template = wb_template.sheets[0]

        valeur_d45 = row.get('Valeur D45', 0)
        noms_et_prenoms = row.get('Noms et Prénoms', 'Inconnu') or 'Inconnu'
        fonction = row.get('Fonction ', 'Inconnu') or 'Inconnu'
        matricule = row.get('Matricule', 'Non spécifié') or 'Non spécifié'
        date_embauche = row.get('Date embauche', '') or ''
        salaire_base = row.get('Salaire de base', 0)
        nbr_jour_travail = row.get('Nbr jour travail', 0)
        abattement = row.get('Abattement', 0)
        nbre_enfants = row.get('Nbre enfants', 0)

        prime_chef_antenne = row.get('Prime chef d\'Antenne', 0)
        prime_objectif = row.get('Prime d\'objectif', 0)
        solde_prime_objectif_2023 = row.get('Solde sur prime d\'objectif 2023', 0)
        prime_puissance_centrale = row.get('Prime puissance centrale', 0)
        prime_astreinte_centrale = row.get('Prime astreinte centrale', 0)
        indemnité_logement = row.get('Indemnité de Logement', 0)
        indemnité_représentation = row.get('Indemnité de Représentation', 0)
        commission_remboursable = row.get('Commission Remboursable', 0)

        # Mettre à jour le modèle avec les valeurs
        ws_template.range('C14').value = noms_et_prenoms
        ws_template.range('C15').value = fonction
        ws_template.range('D16').value = f"Matricule: {matricule}"
        ws_template.range('C12').value = combined_values
        ws_template.range('C13').value = date_embauche
        ws_template.range('C10').value = i10_value
        ws_template.range('C11').value = h10_value
        ws_template.range('D20').value = salaire_base
        ws_template.range('D21').value = nbr_jour_travail
        ws_template.range('E22').value = montant_total  # Montant total
        ws_template.range('E23').value = prime_chef_antenne
        ws_template.range('E24').value = prime_objectif
        ws_template.range('E25').value = solde_prime_objectif_2023
        ws_template.range('E26').value = prime_puissance_centrale
        ws_template.range('E27').value = prime_astreinte_centrale
        ws_template.range('E28').value = indemnité_logement
        ws_template.range('E29').value = indemnité_représentation
        ws_template.range('E30').value = commission_remboursable

        # Placer Nbre enfants et Abattement dans C46 et C47
        ws_template.range('C46').value = nbre_enfants
        ws_template.range('C47').value = abattement

        # Calculer et mettre à jour D45 avec le produit de C46 et C47
        d45_value = nbre_enfants * abattement
        ws_template.range('D45').value = d45_value

        # Calculer et mettre à jour E35 avec la somme de E20:E34
        e35_value = sum_e20_to_e34(ws_template)
        ws_template.range('E35').value = e35_value

        # Mettre à jour D37 et D39 avec la valeur calculée
        ws_template.range('D37').value = calculated_value
        ws_template.range('D39').value = calculated_value

        # Calculer la somme de D37 et D39, puis la placer dans D41
        d37_value = ws_template.range('D37').value
        d39_value = ws_template.range('D39').value
        d41_value = d37_value + d39_value
        ws_template.range('D41').value = d41_value

        # Placer W15 dans E43
        ws_template.range('E43').value = w15_value

        # Placer Y15 dans D44
        ws_template.range('D44').value = y15_value

        # Nouvelle étape : Calculer la différence entre D44 et D45 et placer la valeur dans D48
        d44_value = ws_template.range('D44').value
        d48_value = d44_value - d45_value
        ws_template.range('D48').value = d48_value

        # **Nouvelle étape : Insérer "Avance quinzaine" dans D50**
        if pd.isna(avance_quinzaine) or avance_quinzaine == '':
            avance_quinzaine = 0
        ws_template.range('D50').value = avance_quinzaine

        output_file = os.path.join(output_path, f"{noms_et_prenoms}_{matricule}.xlsx")
        wb_template.save(output_file)
        wb_template.close()
        app.quit()
        print(f"Fiche enregistrée : {output_file}")
    except Exception as e:
        print(f"Erreur lors du remplissage du modèle : {e}")

def display_columns(df_sud):
    try:
        if not df_sud.empty:
            if 'Nbre enfants' in df_sud.columns and 'Abattement' in df_sud.columns:
                for index, row in df_sud.iterrows():
                    print(f"Nom: {row.get('Noms et Prénoms', 'Inconnu')}, Nbre enfants: {row.get('Nbre enfants', 0)}, Abattement: {row.get('Abattement', 0)}")
            else:
                print("Les colonnes 'Nbre enfants' ou 'Abattement' sont absentes du DataFrame.")
    except Exception as e:
        print(f"Erreur lors de l'affichage des colonnes : {e}")



import os
import math
import pandas as pd
import xlwings as xw

def process_files(input_dir, template_path, output_path):
    for file_name in os.listdir(input_dir):
        if file_name.endswith('.xlsx'):
            input_file = os.path.join(input_dir, file_name)
            df_sud, combined_values, h10_value, i10_value, f15_value, g15_value, j15_value, k15_value, l15_value, m15_value, n15_value, o15_value, p15_value, q15_value = load_first_sheet(input_file)
            
            # Afficher les valeurs des colonnes "Nbre enfants" et "Abattement"
            display_columns(df_sud)

            if f15_value is not None and g15_value is not None:
                montant_total = f15_value * g15_value
            else:
                montant_total = 0

            du_value, mois, annee, montant_du_mois = extract_values(df_sud)
            R15, S15, U15, W15 = calculate_r15_s15_u15(df_sud)
            Y15 = calculate_y15(R15, S15, U15)

            # Calculer la somme de J15 à Q15
            j15_to_q15_sum = sum(filter(None, [j15_value, k15_value, l15_value, m15_value, n15_value, o15_value, p15_value, q15_value]))

            # Calculer la somme de montant_total et j15_to_q15_sum
            R15 = montant_total + j15_to_q15_sum

            # Affichage des valeurs J15 à Q15, de leur somme et de la somme totale
            print(f"Valeurs J15 à Q15 : J15={j15_value}, K15={k15_value}, L15={l15_value}, M15={m15_value}, N15={n15_value}, O15={o15_value}, P15={p15_value}, Q15={q15_value}")
            print(f"Somme de J15 à Q15 : {j15_to_q15_sum}")
            print(f"Montant total : {montant_total}")
            print(f"Somme totale (Montant total + Somme J15 à Q15) : {R15}")

            # Calculer et afficher la valeur selon la formule donnée
            calculated_value = min(R15 * 0.01, 20000)
            print(f"Valeur calculée selon la formule : {calculated_value}")

            # Calculer la formule R15 - calculated_value - calculated_value
            W15 = R15 - calculated_value - calculated_value
            print(f"Résultat de la formule R15 - calculated_value - calculated_value : {W15}")

            # Calculer ARRONDI.INF(W15; -2)
            X15 = math.floor(W15 / 100) * 100
            print(f"Valeur arrondie (ARRONDI.INF) : {X15}")

            # Calculer la formule complexe pour la valeur de X15
            if X15 <= 400000:
                Y15 = max(5 * (X15 - 350000) / 100, 3000)
            elif X15 <= 500000:
                Y15 = max(((X15 - 400000) * 10 / 100) + 2500, 3000)
            elif X15 <= 600000:
                Y15 = max(((X15 - 500000) * 15 / 100) + 12500, 3000)
            else:
                Y15 = max(((X15 - 600000) * 20 / 100) + 27500, 3000)

            # Afficher la valeur calculée
            print(f"Valeur calculée selon MAX(SI(...)) : {Y15}")

            for index, row in df_sud.iterrows():
                # Extraire la valeur de la colonne 'Avance quinzaine'
                avance_quinzaine = row.get('Avance quinzaine', 0)
                print(f"Avance quinzaine : {avance_quinzaine}")

                fill_template(row, template_path, output_path, du_value, mois, annee, montant_total, S15, W15, Y15, combined_values, h10_value, i10_value, calculated_value, avance_quinzaine)


# Exécution du script
input_dir = 'C:\\Users\\CE PC\\Desktop\\FicheDePaie'
template_path = 'C:\\Users\\CE PC\\Desktop\\FicheDePaie\\ModèleFiche\\FICHE DE PAIE 2024.xlsx'
output_dir = 'C:\\Users\\CE PC\\Desktop\\FicheDePaie\\Output'

process_files(input_dir, template_path, output_dir)

