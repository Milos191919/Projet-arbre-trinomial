import xlwings as xw
import datetime
import time

from py_class.arbre import Arbre
from py_class.market import Market
from py_class.option import Contract
from py_class.display import gerer_affichage_granulaire
from py_class.utils import BS, calculate_delta, calculate_gamma, calculate_vega, calculate_volga, calculate_vanna


@xw.func
def main():
    """
    Fonction principale pour le pricing d'options avec arbre binomial et comparaison Black-Scholes.
    
    Cette fonction:
    - Lit les paramètres d'entrée depuis le fichier Excel
    - Construit un arbre binomial pour le pricing d'options
    - Compare le prix de l'arbre avec la solution analytique de Black-Scholes
    - Calcule les grecques (Delta, Gamma, Vega, Volga, Vanna)
    - Renvoie les résultats vers Excel
    """
    
    print("Start !!!!")
    
    # Connexion au classeur Excel appelant
    wb = xw.Book.caller()
    ws1 = wb.sheets["Pricer"]

    # ===== Lecture des paramètres de marché et d'option depuis Excel =====
    S = ws1.range("St").value              # Prix actuel du sous-jacent
    K = ws1.range("Strike").value          # Prix d'exercice (strike) de l'option
    r = ws1.range("IntRate").value         # Taux d'intérêt sans risque
    sigma = ws1.range("Vol").value         # Volatilité (annualisée)
    type_op = ws1.range("OptType").value   # Type d'option: "call" ou "put"
    ex_op = ws1.range("EU_US").value       # Style d'exercice: "European" ou "American"
    today = ws1.range("Pr_Date").value     # Date de pricing (aujourd'hui)
    maturity_date = ws1.range("Mat").value # Date de maturité de l'option
    div = ws1.range("DivAmount").value     # Montant du dividende
    div_date = ws1.range("DivDate").value  # Date de détachement du dividende

    N = int(ws1.range("Steps").value)      # Nombre d'étapes dans l'arbre binomial

    # ===== Calcul du temps jusqu'à maturité et du timing du dividende =====
    T = (maturity_date - today).days / 365      # Temps jusqu'à maturité en années
    T_div = (div_date - today).days / 365       # Temps jusqu'au dividende en années

    # Lecture de la préférence d'affichage pour l'arbre granulaire
    affichage = ws1.range("I18").value

    # ===== Création des objets Market et Contract =====
    # Market contient les paramètres de marché (prix, taux, volatilité, dividendes)
    m = Market(stock_price=S, int_rate=r, sigma=sigma, div=div, div_date=div_date)
    
    # Contract contient les caractéristiques du contrat d'option
    c = Contract(pricing_date=today, maturity_date=maturity_date, strike=K, 
                 op_type=type_op, op_exercice=ex_op)

    # ===== Construction de l'arbre binomial =====
    # Mesure du temps de construction de l'arbre
    start = time.perf_counter()
    ar = Arbre(market=m, contract=c, n_steps=N)
    end = time.perf_counter()
    elapsed_time = end - start
    ws1.range("P10").value = elapsed_time  # Temps de construction de l'arbre
    
    # ===== Calcul du prix Black-Scholes (référence analytique) =====
    bs_price = BS(S, K, T, r, sigma, type_op, div, T_div)
    ws1.range("P7").value = bs_price

    # ===== Pricing par arbre binomial (méthode itérative backward) =====
    # Cette méthode remonte l'arbre depuis les feuilles jusqu'à la racine
    s = time.perf_counter()
    backward_tree_price = c.price_iteratively(ar)
    e = time.perf_counter()
    tree_pricing_delay = e - s
    ws1.range("P8").value = backward_tree_price  # Prix calculé par backward iteration

    # ===== Pricing par arbre binomial (méthode récursive) =====
    # Méthode alternative utilisant la récursion
    recursive_tree_price = c.price_recursively(ar)
    ws1.range("P13").value = recursive_tree_price  # Prix calculé par récursion

    ws1.range("P11").value = tree_pricing_delay  # Temps de calcul du pricing

    # ===== Calcul des grecques (sensibilités de l'option) =====
    tree_delta = calculate_delta(m, c, N)    # Sensibilité au prix du sous-jacent
    tree_gamma = calculate_gamma(m, c, N)    # Sensibilité du delta (convexité)
    tree_vega = calculate_vega(m, c, N)      # Sensibilité à la volatilité
    tree_volga = calculate_volga(m, c, N)    # Sensibilité du vega à la volatilité
    tree_vanna = calculate_vanna(m, c, N)    # Autre sensibilité de second ordre

    # ===== Écriture des grecques dans Excel =====
    ws1.range("P18").value = tree_delta   # Sortie Delta
    ws1.range("P19").value = tree_gamma   # Sortie Gamma 
    ws1.range("P20").value = tree_vega    # Sortie Vega
    ws1.range("P21").value = tree_volga   # Sortie Volga
    ws1.range("P22").value = tree_vanna   # Sortie Vanna

    # ===== Gestion de l'affichage granulaire de l'arbre =====
    # Affiche l'arbre binomial dans Excel si demandé
    gerer_affichage_granulaire(ar, affichage)

    print("Done !!!!")


@xw.func
@xw.arg('maturity_date', dates=datetime.date)
@xw.arg('pricing_date', dates=datetime.date)
@xw.arg('div_date', dates=datetime.date) 
def OptionPricerPy(pricing_date, maturity_date, stock_price, strike, int_rate, 
                   sigma, div, div_date, op_type, op_exercice, n_steps):
    """
    Fonction de pricing d'option appelable depuis Excel.
    
    Cette fonction peut être utilisée comme formule Excel pour calculer le prix
    d'une option en utilisant un arbre binomial.
    
    Paramètres:
    -----------
    pricing_date : date
        Date de valorisation
    maturity_date : date
        Date de maturité de l'option
    stock_price : float
        Prix actuel du sous-jacent
    strike : float
        Prix d'exercice
    int_rate : float
        Taux d'intérêt sans risque
    sigma : float
        Volatilité
    div : float
        Montant du dividende
    div_date : date
        Date du dividende
    op_type : str
        Type d'option ("call" ou "put")
    op_exercice : str
        Style d'exercice ("European" ou "American")
    n_steps : int
        Nombre d'étapes dans l'arbre
        
    Retourne:
    ---------
    float
        Prix de l'option
    """
    # Gestion de la date de dividende (peut être None)
    market_div_date = div_date
    if div_date is None:
        market_div_date = None 

    # Création des objets market et contract
    market = Market(stock_price=stock_price, int_rate=int_rate, sigma=sigma, 
                   div=div, div_date=market_div_date)
    contract = Contract(pricing_date=pricing_date, maturity_date=maturity_date, 
                       strike=strike, op_type=op_type, op_exercice=op_exercice)

    # Construction de l'arbre et calcul du prix
    tree = Arbre(market=market, contract=contract, n_steps=int(n_steps))
    price = contract.price_iteratively(tree)

    return price


if __name__ == '__main__':
    # Point d'entrée pour les tests en local
    # Permet d'exécuter le script sans avoir Excel ouvert
    xw.Book("Projet_Milos_Issam.xlsm").set_mock_caller()
    main()