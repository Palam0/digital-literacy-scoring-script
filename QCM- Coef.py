from openpyxl import load_workbook, Workbook

def score_participants(rep, que):
    participant = que[1]  # Le participant est stocké dans le deuxième élément de la liste `que`
    resilience = 0
    esprit_critique = 0
    comportement_social = 0
    competences_techniques = 0
    traitement_information = 0
    creation = 0
    total = 0

    # Boucle sur les réponses du participant en commençant à l'index 3 (les réponses commencent à partir du 4ème élément de `que`)
    for index, elem in enumerate(que[3:]):
        question_type = rep[index][0]
        reponses_possibles = rep[index][1:]
        
        for reponse, score in reponses_possibles:
            if elem[0] == reponse:
                # Incrémente le score correspondant à la catégorie de la question
                if question_type == "Résilience": resilience += score
                elif question_type == "EC": esprit_critique += score
                elif question_type == "CSDLEN": comportement_social += score
                elif question_type == "CDC": creation += score
                elif question_type == "CT": competences_techniques += score
                elif question_type == "TDLinfo": traitement_information += score
    
    if resilience < 0 :
        resilience = 0
    if esprit_critique < 0:
        esprit_critique = 0
    if comportement_social < 0:
        comportement_social = 0
    if competences_techniques < 0:
        competences_techniques = 0
    if traitement_information < 0:
        traitement_information = 0
    if creation < 0:
        creation = 0
    total = resilience + esprit_critique + comportement_social + competences_techniques + traitement_information + creation
    # Retourne la liste des scores pour le participant
    return [participant, resilience, esprit_critique, comportement_social, competences_techniques, traitement_information, creation, total]


# Chemin vers votre fichier Excel
file = r'C:\Users\Kévin\Downloads\QCM LVL LITE.xlsx'

# Liste des réponses correctes pour chaque question
reponses = [
    ('Résilience', ('A',0.5),('B',1),('C',-1),('D',-0.5),('E',0)),
    ('Résilience', ('A',-0.5),('B',-1),('C',1),('D',-0.5),('E',0)),
    ('EC', ('A',-1),('B',1),('C',-0.5),('D',-0.5),('E',0)),
    ('CSDLEN', ('A',-0.5),('B',1),('C',-0.5),('D',-0.5),('E',0)),
    ('CSDLEN', ('A',-0.5),('B',1),('C',-1),('D',-0.5),('E',0)),
    ('CDC', ('A',-0.5),('B',-0.5),('C',1),('D',-1),('E',0)),
    ('Résilience', ('A',1),('B',0),('C',0.5),('D',-1),('E',0)),
    ('CDC', ('A',-1),('B',-0.5),('C',1),('D',-1),('E',0)),
    ('CT', ('1',1),('2',-1),('3',-0.5),('4',-1),('5',0)),
    ('TDLinfo', ('A',0.5),('B',-1),('C',1),('D',0),('E',0)),
    ('CDC', ('A',-1),('B',1),('C',1),('D',-1),('E',0)), #Je souhaite partager une vidéo avec un ami
    ('CT', ('1',1),('2',-0.5),('3',-1),('4',-1),('5',0)),
    ('CDC', ('A',-1),('B',0.5),('C',1),('D',-0.5),('E',0)),
    ('Résilience', ('A',-0.5),('B',0.5),('C',-1),('D',1),('E',0)),
    ('CT', ('1',0.5),('2',1),('3',-1),('4',-1),('5',0)),
    ('CSDLEN', ('A',-1),('B',-0.5),('C',-1),('D',1),('E',0)), #Fake news
    ('EC', ('A',1),('B',-1),('C',-0.5),('D',-1),('E',0)),
    ('TDLinfo', ('A',1),('B',0.5),('C',-1),('D',-1),('E',0)),
    ('TDLinfo', ('A',1),('B',-1),('C',-0.5),('D',0.5),('E',0)),
    ('CT', ('1',-1),('2',-1),('3',1),('4',-1),('5',0)),
    ('CT', ('1',-1),('2',1),('3',-1),('4',-1),('5',0)),
    ('CDC', ('A',-0.5),('B',1),('C',-1),('D',-1),('E',0)),
    ('TDLinfo', ('A',-1),('B',-1),('C',1),('D',-0.5),('E',0)),
    ('CSDLEN', ('A',-1),('B',-0.5),('C',1),('D',-1),('E',0)),
    ('EC', ('A',-1),('B',-1),('C',1),('D',-1),('E',0)),
    ('CSDLEN', ('A',-0.5),('B',1),('C',-0.5),('D',-0.5),('E',0)),
    ('Résilience', ('A',-0.5),('B',-1),('C',1),('D',0),('E',0)),
    ('EC', ('A',1),('B',-1),('C',-1),('D',-0.5),('E',0)),
    ('TDLinfo', ('A',0.5),('B',-1),('C',1),('D',-0.5),('E',0)),
    ('EC', ('A',-1),('B',1),('C',-0.5),('D',-1),('E',0)),
]

# Charger le fichier Excel des résultats aux questions
workbook = load_workbook(file)
sheet = workbook.active

data = []

# Itération sur les lignes et colonnes de la feuille active pour lire les données
for row in sheet.iter_rows(values_only=True):
    data.append(list(row))

# Supprimer la première ligne car elle contient les questions, pas les réponses
del data[0]

# Créer un nouveau classeur pour les résultats
result_workbook = Workbook()
result_sheet = result_workbook.active

# Écrire les en-têtes dans le nouveau classeur
headers = ["Participant", "Résilience", "Esprit Critique", "Comportement Social", "Compétences Techniques", "Traitement de l'Information", "Création", "Total"]
result_sheet.append(headers)

# Calculer et écrire les scores pour chaque participant
for i in range(len(data)):
    scores = score_participants(reponses, data[i])
    result_sheet.append(scores)

# Sauvegarder le nouveau fichier Excel avec les résultats
result_file = r'C:\Users\Kévin\Downloads\scores dimensions QCM LVL LITE.xlsx'
result_workbook.save(result_file)

print(f"Les résultats ont été enregistrés dans {result_file}")



