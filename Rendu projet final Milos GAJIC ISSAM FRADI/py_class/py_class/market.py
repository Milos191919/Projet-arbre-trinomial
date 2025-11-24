class Market:
    """
    Classe représentant les conditions de marché pour le pricing d'options.
    
    Cette classe encapsule tous les paramètres de marché nécessaires pour
    construire l'arbre trinomial et pricer une option:
    - Prix actuel du sous-jacent
    - Taux d'intérêt sans risque
    - Volatilité implicite
    - Dividendes 
    """
    
    def __init__(self, stock_price, int_rate, sigma, div, div_date):
        """
        Initialise un objet Market avec les paramètres de marché.
        
        Paramètres:
        -----------
        stock_price : float
            Prix actuel du sous-jacent (S0)
        int_rate : float
            
        sigma : float
            Volatilité implicite du sous-jacent 
            
        div : float
            Montant du dividende à détacher
            
        div_date : date
            Date de détachement du dividende 
        """
        # ===== Prix du sous-jacent =====
        self.stock_price = stock_price
        
        # ===== Paramètres de marché =====
        self.int_rate = int_rate  # Taux sans risque (r)
        self.sigma = sigma        # Volatilité (σ)
        
        # ===== Informations sur les dividendes =====
        self.div = div            # Montant du dividende (D)
        self.div_date = div_date  # Date de détachement du dividende