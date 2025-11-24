import xlwings as xw
import math

# ===== Liste de toutes les feuilles Python générées automatiquement =====
ALL_PY_SHEETS = [
    "Py_Prix_SI",          # Prix du sous-jacent à chaque nœud
    "Py_Prix_Option",      # Prix de l'option à chaque nœud
    "Py_Proba_Up",         # Probabilités de transition vers le haut
    "Py_Proba_Mid",        # Probabilités de transition vers le milieu
    "Py_Proba_Down",       # Probabilités de transition vers le bas
    "Py_Proba_Cumulee",    # Probabilités cumulées d'atteindre chaque nœud
    "Py_Variance"          # Variance à chaque nœud
]


def _get_or_create_sheet(wb, sheet_name):
    """
    Vérifie si une feuille existe dans le classeur. Si non, la crée.
    
    Cette fonction utilitaire permet de gérer dynamiquement les feuilles Excel:
    - Si la feuille existe déjà, elle est effacée et réutilisée
    - Si la feuille n'existe pas, elle est créée à la fin du classeur
    
    Paramètres:
    -----------
    wb : xw.Book
        Classeur Excel
    sheet_name : str
        Nom de la feuille à récupérer ou créer
        
    Retourne:
    ---------
    xw.Sheet
        Objet feuille Excel (nettoyée)
    """
    try:
        # Tentative de récupération de la feuille existante
        sheet = wb.sheets[sheet_name]
    except:
        # Si la feuille n'existe pas, la créer à la fin du classeur
        sheet = wb.sheets.add(name=sheet_name, after=wb.sheets[-1])
    
    # Nettoyage complet du contenu de la feuille
    sheet.clear()
    
    return sheet


def _afficher_prix_si(arbre, wb):
    """
    Affiche les prix du sous-jacent (si) dans la feuille 'Py_Prix_SI'.
    
    Cette fonction parcourt l'arbre trinomial et écrit le prix du sous-jacent
    à chaque nœud dans Excel. L'affichage est organisé en colonnes (étapes temporelles)
    et en lignes (niveaux de prix).
    
    Structure de l'affichage:
    -------------------------
    - Chaque colonne = une étape temporelle (de 0 à N)
    - Chaque ligne = un niveau de prix
    - La ligne du milieu (N) correspond au tronc de l'arbre
    - Les lignes au-dessus correspondent aux nœuds up
    - Les lignes en-dessous correspondent aux nœuds down
    
    Paramètres:
    -----------
    arbre : Arbre
        Arbre trinomial construit
    wb : xw.Book
        Classeur Excel
    """
    # ===== Récupération ou création de la feuille =====
    ws = _get_or_create_sheet(wb, "Py_Prix_SI")
    
    # ===== Calcul du nombre de colonnes =====
    N = arbre.n_steps + 1  
    
    # ===== Point de départ: nœud du milieu à chaque étape =====
    node_mid = arbre.racine

    for k in range(N):
        # --- Affichage du nœud du milieu ---
        ws.range("A1").offset(N, k).value = node_mid.si
        
        # --- Remontée: affichage des nœuds au-dessus du milieu ---
        node = node_mid.voisin_up
        i = 1  # Compteur de distance par rapport au milieu
        while node is not None:
            ws.range("A1").offset(N - i, k).value = node.si
            i += 1
            node = node.voisin_up
        
        # --- Descente: affichage des nœuds en-dessous du milieu ---
        node = node_mid.voisin_down
        i = 1  # Compteur de distance par rapport au milieu
        while node is not None:
            ws.range("A1").offset(N + i, k).value = node.si
            i += 1
            node = node.voisin_down
        
        # --- Passage à l'étape suivante ---
        node_mid = node_mid.next_mid
    
    # ===== Ajout d'un label explicatif =====
    ws.range("A1").offset(N, N + 2).value = "Prix du Sous-Jacent (si)"


def _afficher_prix_option(arbre, wb):
    """
    Affiche les prix de l'option (si2) dans la feuille 'Py_Prix_Option'.
    
    Cette fonction affiche les valeurs de l'option calculées lors du backward pricing.
    La structure est identique à _afficher_prix_si mais affiche si2 au lieu de si.
    
    Paramètres:
    -----------
    arbre : Arbre
        Arbre trinomial avec pricing effectué (si2 calculé)
    wb : xw.Book
        Classeur Excel
    """
    ws = _get_or_create_sheet(wb, "Py_Prix_Option")
    N = arbre.n_steps + 1
    node_mid = arbre.racine

    for k in range(N):
        # Affichage du prix de l'option au nœud mid
        ws.range("A1").offset(N, k).value = node_mid.si2
        
        # Remontée: nœuds up
        node = node_mid.voisin_up
        i = 1
        while node is not None:
            ws.range("A1").offset(N - i, k).value = node.si2
            i += 1
            node = node.voisin_up
        
        # Descente: nœuds down
        node = node_mid.voisin_down
        i = 1
        while node is not None:
            ws.range("A1").offset(N + i, k).value = node.si2
            i += 1
            node = node.voisin_down
        
        node_mid = node_mid.next_mid
    
    ws.range("A1").offset(N, N + 2).value = "Prix de l'Option (si2)"


def _afficher_proba_up(arbre, wb):
    """
    Affiche les probabilités de transition vers le haut (p_up) dans 'Py_Proba_Up'.
    
    p_up représente la probabilité risque-neutre de passer d'un nœud à son nœud
    suivant vers le haut (next_up).
    
    Paramètres:
    -----------
    arbre : Arbre
        Arbre trinomial construit
    wb : xw.Book
        Classeur Excel
    """
    ws = _get_or_create_sheet(wb, "Py_Proba_Up")
    N = arbre.n_steps + 1
    node_mid = arbre.racine

    for k in range(N):
        ws.range("A1").offset(N, k).value = node_mid.p_up
        
        node = node_mid.voisin_up
        i = 1
        while node is not None:
            ws.range("A1").offset(N - i, k).value = node.p_up
            i += 1
            node = node.voisin_up
        
        node = node_mid.voisin_down
        i = 1
        while node is not None:
            ws.range("A1").offset(N + i, k).value = node.p_up
            i += 1
            node = node.voisin_down
        
        node_mid = node_mid.next_mid
    
    ws.range("A1").offset(N, N + 2).value = "Proba Up (p_up)"


def _afficher_proba_mid(arbre, wb):
    """
    Affiche les probabilités de transition vers le milieu (p_mid) dans 'Py_Proba_Mid'.
    
    p_mid représente la probabilité risque-neutre de passer d'un nœud à son nœud
    suivant au milieu (next_mid).

    Paramètres:
    -----------
    arbre : Arbre
        Arbre trinomial construit
    wb : xw.Book
        Classeur Excel
    """
    ws = _get_or_create_sheet(wb, "Py_Proba_Mid")
    N = arbre.n_steps + 1
    node_mid = arbre.racine

    for k in range(N):
        ws.range("A1").offset(N, k).value = node_mid.p_mid
        
        node = node_mid.voisin_up
        i = 1
        while node is not None:
            ws.range("A1").offset(N - i, k).value = node.p_mid
            i += 1
            node = node.voisin_up
        
        node = node_mid.voisin_down
        i = 1
        while node is not None:
            ws.range("A1").offset(N + i, k).value = node.p_mid
            i += 1
            node = node.voisin_down
        
        node_mid = node_mid.next_mid
    
    ws.range("A1").offset(N, N + 2).value = "Proba Mid (p_mid)"


def _afficher_proba_down(arbre, wb):
    """
    Affiche les probabilités de transition vers le bas (p_down) dans 'Py_Proba_Down'.
    
    p_down représente la probabilité risque-neutre de passer d'un nœud à son nœud
    suivant vers le bas (next_down).
    
    Paramètres:
    -----------
    arbre : Arbre
        Arbre trinomial construit
    wb : xw.Book
        Classeur Excel
    """
    ws = _get_or_create_sheet(wb, "Py_Proba_Down")
    N = arbre.n_steps + 1
    node_mid = arbre.racine

    for k in range(N):
        ws.range("A1").offset(N, k).value = node_mid.p_down
        
        node = node_mid.voisin_up
        i = 1
        while node is not None:
            ws.range("A1").offset(N - i, k).value = node.p_down
            i += 1
            node = node.voisin_up
        
        node = node_mid.voisin_down
        i = 1
        while node is not None:
            ws.range("A1").offset(N + i, k).value = node.p_down
            i += 1
            node = node.voisin_down
        
        node_mid = node_mid.next_mid
    
    ws.range("A1").offset(N, N + 2).value = "Proba Down (p_down)"


def _afficher_proba_cumule(arbre, wb):
    """
    Affiche les probabilités cumulées dans la feuille 'Py_Proba_Cumulee'.
    
    La probabilité cumulée d'un nœud représente la probabilité totale d'atteindre
    ce nœud depuis la racine en suivant tous les chemins possibles.
    
    Paramètres:
    -----------
    arbre : Arbre
        Arbre trinomial construit avec probabilités propagées
    wb : xw.Book
        Classeur Excel
    """
    ws = _get_or_create_sheet(wb, "Py_Proba_Cumulee")
    N = arbre.n_steps + 1
    node_mid = arbre.racine

    for k in range(N):
        ws.range("A1").offset(N, k).value = node_mid.proba_cumule
        
        node = node_mid.voisin_up
        i = 1
        while node is not None:
            ws.range("A1").offset(N - i, k).value = node.proba_cumule
            i += 1
            node = node.voisin_up
        
        node = node_mid.voisin_down
        i = 1
        while node is not None:
            ws.range("A1").offset(N + i, k).value = node.proba_cumule
            i += 1
            node = node.voisin_down
        
        node_mid = node_mid.next_mid
    
    ws.range("A1").offset(N, N + 2).value = "Proba Cumulée"


def _afficher_variance(arbre, wb):
    """
    Calcule et affiche la variance à chaque nœud dans la feuille 'Py_Variance'.
    
    La variance représente la dispersion du prix du sous-jacent autour de sa valeur
    forward à l'étape suivante. Elle est calculée selon la formule:
    
    Var(S) = S² × exp(2r×dt) × [exp(σ²×dt) - 1]
    
    Paramètres:
    -----------
    arbre : Arbre
        Arbre trinomial construit
    wb : xw.Book
        Classeur Excel
    """
    ws = _get_or_create_sheet(wb, "Py_Variance")
    N = arbre.n_steps + 1
    node_mid = arbre.racine
    
    # ===== Récupération des paramètres pour le calcul =====
    r = arbre.market.int_rate    # Taux d'intérêt 
    dt = arbre.dt                # Pas de temps
    sigma = arbre.market.sigma   # Volatilité

    # ===== Boucle sur les étapes temporelles =====
    for k in range(N):
        # --- Calcul de la variance pour le nœud mid ---
        value = (node_mid.si ** 2) * math.exp(2 * r * dt) * (math.exp(sigma**2 * dt) - 1)
        ws.range("A1").offset(N, k).value = value
        
        # --- Remontée: nœuds up ---
        node = node_mid.voisin_up
        i = 1
        while node is not None:
            value = (node.si ** 2) * math.exp(2 * r * dt) * (math.exp(sigma**2 * dt) - 1)
            ws.range("A1").offset(N - i, k).value = value
            i += 1
            node = node.voisin_up
        
        # --- Descente: nœuds down ---
        node = node_mid.voisin_down
        i = 1
        while node is not None:
            value = (node.si ** 2) * math.exp(2 * r * dt) * (math.exp(sigma**2 * dt) - 1)
            ws.range("A1").offset(N + i, k).value = value
            i += 1
            node = node.voisin_down
        
        node_mid = node_mid.next_mid
    
    ws.range("A1").offset(N, N + 2).value = "Variance"


def gerer_affichage_granulaire(arbre, affichage_str):
    """
    Fonction principale de gestion de l'affichage granulaire de l'arbre dans Excel.
    
    Cette fonction orchestre l'affichage des différentes informations de l'arbre
    selon la demande de l'utilisateur. Elle permet de visualiser sélectivement:
    - Les prix du sous-jacent et de l'option
    - Les probabilités de transition
    - Les probabilités cumulées
    - La variance
    
    Modes d'affichage:
    ------------------
    - "all": Affiche toutes les feuilles
    - "prix": Affiche seulement les prix (si et si2)
    - "proba": Affiche seulement les probabilités (up, mid, down, cumulée)
    - "variance": Affiche seulement la variance
    - Autre/None: Nettoie toutes les feuilles Python sans rien afficher

    Paramètres:
    -----------
    arbre : Arbre
        Arbre trinomial construit et pricé
    affichage_str : str
        Mode d'affichage demandé ("all", "prix", "proba", "variance", etc.)
    """
    # ===== Connexion au classeur Excel =====
    wb = xw.Book.caller()
    app = wb.app
    
    # ===== Normalisation de la demande d'affichage =====
    affichage_lower = str(affichage_str).lower()
    sheets_to_create = []  

    # ===== Détermination des feuilles à créer selon la demande =====
    if affichage_lower == "all":
        # Mode ALL: toutes les feuilles
        sheets_to_create = ALL_PY_SHEETS
    
    elif affichage_lower == "prix":
        # Mode PRIX: seulement les prix du sous-jacent et de l'option
        sheets_to_create = ["Py_Prix_SI", "Py_Prix_Option"]
    
    elif affichage_lower == "proba":
        # Mode PROBA: toutes les probabilités
        sheets_to_create = ["Py_Proba_Up", "Py_Proba_Mid", "Py_Proba_Down", "Py_Proba_Cumulee"]
    
    elif affichage_lower == "variance":
        # Mode VARIANCE: seulement la variance
        sheets_to_create = ["Py_Variance"]
    
    # Si affichage_lower ne correspond à aucun mode, sheets_to_create reste vide

    # ===== Désactivation des mises à jour Excel =====
    app.api.ScreenUpdating = False     # Pas de rafraîchissement d'écran
    app.api.Calculation = -4135        # xlCalculationManual: calcul manuel
    app.api.EnableEvents = False       # TRÈS IMPORTANT avant de supprimer des feuilles

    print(f"Demande d'affichage : '{affichage_lower}'. Feuilles à créer : {sheets_to_create}")

    # ===== Nettoyage: suppression des feuilles non demandées =====
    for sheet in wb.sheets:
        if sheet.name in ALL_PY_SHEETS and sheet.name not in sheets_to_create:
            print(f"Suppression de(s) feuille(s) : {sheet.name}")
            sheet.delete()
    
    # ===== Création/mise à jour des feuilles demandées =====
    if "Py_Prix_SI" in sheets_to_create:
        _afficher_prix_si(arbre, wb)
    
    if "Py_Prix_Option" in sheets_to_create:
        _afficher_prix_option(arbre, wb)
    
    if "Py_Proba_Up" in sheets_to_create:
        _afficher_proba_up(arbre, wb)
    
    if "Py_Proba_Mid" in sheets_to_create:
        _afficher_proba_mid(arbre, wb)
    
    if "Py_Proba_Down" in sheets_to_create:
        _afficher_proba_down(arbre, wb)
    
    if "Py_Proba_Cumulee" in sheets_to_create:
        _afficher_proba_cumule(arbre, wb)
    
    if "Py_Variance" in sheets_to_create:
        _afficher_variance(arbre, wb)

    # ===== Réactivation des mises à jour Excel =====
    app.api.Calculation = -4105        # xlCalculationAutomatic: calcul automatique
    app.api.EnableEvents = True        # Réactivation des événements
    app.api.ScreenUpdating = True      # Réactivation du rafraîchissement d'écran
    
    if sheets_to_create:
        print("Affichage terminé.")
    else:
        print("Nettoyage des feuilles terminé. Aucun affichage demandé.")