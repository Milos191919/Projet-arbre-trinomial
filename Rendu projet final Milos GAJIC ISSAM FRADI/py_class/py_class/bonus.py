import time
import xlwings as xw

from .market import Market
from .option import Contract
from .arbre import Arbre
from .utils import BS, calculate_delta, calculate_gamma, calculate_vega,calculate_volga,calculate_vanna


def elapsed():
    """
    Fonction de test de performance pour mesurer le temps d'exécution du pricing.
    
    Cette fonction:
    1. Lit les paramètres depuis la feuille "Pricer"
    2. Parcourt une liste de nombres d'étapes (N)
    3. Pour chaque N, construit l'arbre et price l'option
    4. Mesure le temps d'exécution et calcule l'erreur par rapport à Black-Scholes
    5. Écrit les résultats dans la feuille "Performance Test"
    """
    # ===== Connexion au classeur Excel =====
    wb = xw.Book.caller()
    ws_pricer = wb.sheets("Pricer")
    ws_perf = wb.sheets("Performance Test")
    
    # ===== Nettoyage des résultats précédents =====
    ws_perf.range("B2:B20000").clear_contents()

    # ===== Lecture des paramètres de l'option depuis la feuille Pricer =====
    S = ws_pricer.range("St").value              # Prix du sous-jacent
    K = ws_pricer.range("Strike").value          # Strike
    r = ws_pricer.range("IntRate").value         # Taux d'intérêt
    sigma = ws_pricer.range("Vol").value         # Volatilité
    type_op = ws_pricer.range("OptType").value   # Call ou Put
    ex_op = ws_pricer.range("EU_US").value       # EU ou US
    today = ws_pricer.range("Pr_Date").value     # Date de pricing
    maturity = ws_pricer.range("Mat").value      # Date de maturité
    div = ws_pricer.range("DivAmount").value     # Montant du dividende
    div_date = ws_pricer.range("DivDate").value  # Date du dividende
    
    # ===== Calcul des temps en années =====
    T_div = (div_date - today).days / 365  # Temps jusqu'au dividende
    T = (maturity - today).days / 365      # Temps jusqu'à maturité

    # ===== Création des objets Market et Contract =====
    m1 = Market(stock_price=S, int_rate=r, sigma=sigma, div=div, div_date=div_date)
    c1 = Contract(pricing_date=today, maturity_date=maturity, strike=K, 
                  op_type=type_op, op_exercice=ex_op)
    
    # ===== Calcul du prix Black-Scholes =====
    bs_price = BS(S, K, T, r, sigma, type_op, div, T_div)

    # ===== Lecture de la liste des nombres d'étapes à tester =====
    n_rows = ws_perf.range("A2").expand("down").value
    N_list = [int(n) for n in n_rows if n is not None]
    
    # ===== Variables pour stocker les résultats =====
    elapsed_times = []  # Temps d'exécution pour chaque N
    res = []            # Temps et erreur pour chaque N
    
    # ===== Boucle principale: test pour chaque nombre d'étapes =====
    for k in N_list:
        # --- Mesure du temps de début ---
        start = time.time()
        
        # --- Construction de l'arbre et pricing ---
        ar1 = Arbre(m1, c1, k)
        price = c1.price_iteratively(ar1)
        
        # --- Mesure du temps de fin ---
        end = time.time()
        elapsed = end - start  

        # --- Stockage des résultats ---
        elapsed_times.append([elapsed])
        res.append([elapsed, price - bs_price]) 

    # ===== Écriture des résultats dans Excel =====
    ws_perf.range("B2").value = elapsed_times  
    ws_perf.range("O2").value = res           


def TreevsBS():
    """
    Fonction d'analyse comparative entre le pricing par arbre trinomial et Black-Scholes.
    
    Cette fonction effectue deux analyses:
    1. Convergence en fonction du nombre d'étapes (N)
    2. Analyse en fonction du strike (K)
    """
    # ===== Connexion au classeur Excel =====
    wb = xw.Book.caller()
    ws_BS = wb.sheets("Tree vs B&S")
    ws_pricer = wb.sheets("Pricer")

    # ===== Lecture des paramètres depuis la feuille Pricer =====
    S = ws_pricer.range("St").value
    K = ws_pricer.range("Strike").value
    r = ws_pricer.range("IntRate").value
    sigma = ws_pricer.range("Vol").value
    type_op = ws_pricer.range("OptType").value
    ex_op = ws_pricer.range("EU_US").value
    today = ws_pricer.range("Pr_Date").value
    maturity = ws_pricer.range("Mat").value
    div = ws_pricer.range("DivAmount").value
    div_date = ws_pricer.range("DivDate").value
    
    # ===== Calcul des temps =====
    T_div = (div_date - today).days / 365
    T = (maturity - today).days / 365
    N_steps = int(ws_pricer.range("Steps").value)

    m1 = Market(stock_price=S, int_rate=r, sigma=sigma, div=div, div_date=div_date)
    c1 = Contract(pricing_date=today, maturity_date=maturity, strike=K, 
                  op_type=type_op, op_exercice=ex_op)

    # ===== Calcul du prix Black-Scholes  =====
    price_bs = BS(S, K, T, r, sigma, type_op, div, T_div)
    
    # ===== ANALYSE 1: Convergence en fonction de N =====
    n_rows = ws_BS.range("A2").expand("down").value
    N_list = [int(n) for n in n_rows if n is not None]
    
    results = []
    for k in N_list:
        # Construction de l'arbre avec k étapes
        ar1 = Arbre(m1, c1, k)
        price_tree = c1.price_iteratively(ar1)
        
        results.append([price_tree, price_bs, (price_tree - price_bs) * k])

    # ===== ANALYSE 2: Sensibilité au strike =====
    n_S_rows = ws_BS.range("O2").expand("down").value
    S_list = [int(s) for s in n_S_rows if s is not None]
    
    res_S = []
    for strike in S_list:
        m2 = Market(stock_price=S, int_rate=r, sigma=sigma, div=div, div_date=div_date)
        c2 = Contract(pricing_date=today, maturity_date=maturity, strike=strike, 
                      op_type=type_op, op_exercice=ex_op)
        
        ar2 = Arbre(m2, c2, n_steps=N_steps)
        price_tree_s = c2.price_iteratively(ar2)
        price_bs_s = BS(S, strike, T, r, sigma, type_op, div, T_div)
        res_S.append([price_tree_s, price_bs_s, price_tree_s - price_bs_s])

    # ===== Écriture des résultats dans Excel =====
    ws_BS.range("B2").value = results  
    ws_BS.range("P2").value = res_S


def Greeks():
    """
    Fonction d'analyse des grecques (Delta, Gamma, Vega) en fonction du prix du sous-jacent.

    Cette fonction:
    1. Lit une liste de prix de sous-jacent depuis la feuille "Greeks Analysis"
    2. Pour chaque prix, calcule les grecques de l'option
    3. Écrit les résultats dans Excel
    """
    # ===== Connexion au classeur Excel =====
    wb = xw.Book.caller()
    ws_G = wb.sheets("Greeks Analysis")
    ws_pricer = wb.sheets("Pricer")

    # ===== Lecture des paramètres depuis la feuille Pricer =====
    K = ws_pricer.range("Strike").value  # Strike
    r = ws_pricer.range("IntRate").value  # Taux d'intérêt
    sigma = ws_pricer.range("Vol").value  # Volatilité
    type_op = ws_pricer.range("OptType").value  # Call ou Put
    ex_op = ws_pricer.range("EU_US").value  # EU ou US
    today = ws_pricer.range("Pr_Date").value  # Date de pricing
    maturity = ws_pricer.range("Mat").value  # Date de maturité
    div = ws_pricer.range("DivAmount").value  # Dividende
    div_date = ws_pricer.range("DivDate").value  # Date du dividende
    N_steps = int(ws_pricer.range("Steps").value)  # Nombre d'étapes

    # ===== Lecture de la liste des prix du sous-jacent à tester =====
    n_rows = ws_G.range("A2").expand("down").value
    N_list = [int(n) for n in n_rows if n is not None]

    results = []

    # ===== Boucle principale: calcul des grecques pour chaque prix =====
    for stock in N_list:
        # --- Création du Market avec le prix du sous-jacent variable ---
        m = Market(stock_price=stock, int_rate=r, sigma=sigma, div=div, div_date=div_date)

        # --- Création du Contract ---
        c = Contract(pricing_date=today, maturity_date=maturity, strike=K,
                     op_type=type_op, op_exercice=ex_op)

        # --- Calcul des grecques ---
        tree_delta = calculate_delta(m, c, N_steps)
        tree_gamma = calculate_gamma(m, c, N_steps)
        tree_vega = calculate_vega(m, c, N_steps)
        tree_volga = calculate_volga(m, c, N_steps)
        tree_vanna = calculate_vanna(m, c, N_steps)

        # --- Stockage des résultats ---
        results.append([tree_delta, tree_gamma, tree_vega, tree_volga, tree_vanna])

    # ===== Écriture des résultats dans Excel =====
    ws_G.range("B2").value = results