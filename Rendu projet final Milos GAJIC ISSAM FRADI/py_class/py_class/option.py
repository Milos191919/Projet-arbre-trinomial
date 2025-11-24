import math

class Contract:
    """
    Classe représentant un contrat d'option (Call ou Put, Européenne ou Américaine).
    
    Cette classe gère:
    - Les caractéristiques du contrat (strike, maturité, type)
    - Le pricing de l'option via l'arbre trinomial
    - Deux méthodes de pricing: récursive et itérative (backward induction)
    """
    
    def __init__(self, pricing_date, maturity_date, strike, op_type=None, op_exercice=None):
        """
        Initialise un contrat d'option.
        
        Paramètres:
        -----------
        pricing_date : date
            Date de valorisation de l'option (aujourd'hui)
        maturity_date : date
            Date de maturité (expiration) de l'option
        strike : float
            Prix d'exercice (strike price) de l'option
        op_type : str, optional
            Type d'option: "Call" ou "Put"
        op_exercice : str, optional
            Style d'exercice: "EU" (Européenne) ou "US" (Américaine)
        """
        self.maturity_date = maturity_date
        self.pricing_date = pricing_date
        
        # ===== Calcul de la maturité en années =====
        self.maturity = ((self.maturity_date - self.pricing_date).days) / 365
        
        self.strike = strike
        self.op_type = op_type          # "Call" ou "Put"
        self.op_exercice = op_exercice  # "EU" (Européenne) ou "US" (Américaine)
        
        # ===== Multiplicateur pour le calcul du payoff =====
        self.op_multiplicator = 1 if self.op_type == "Call" else -1

    def price_recursively(self, arbre):
        """
        Calcule le prix de l'option en utilisant une approche récursive.
        
        Méthode de calcul:
        ------------------
        - Commence à la racine de l'arbre
        - Descend récursivement jusqu'aux feuilles (maturité)
        - Remonte en calculant les valeurs d'espérance actualisée à chaque nœud
        - Utilise la mémoïsation pour éviter les recalculs
        
        Inconvénients:
        --------------
        - Potentiellement plus lent pour de très grands arbres
        - Risque de dépassement de pile (stack overflow) pour N très grand
        
        Paramètres:
        -----------
        arbre : Arbre
            Arbre trinomial construit
            
        Retourne:
        ---------
        float
            Prix de l'option (valeur à la racine)
        """
        # ===== Calcul du facteur d'actualisation =====
        discount_factor = math.exp(-arbre.market.int_rate * arbre.dt)
        
        # ===== Appel de la fonction récursive depuis la racine =====
        return self._recursive_pricer(arbre.racine, discount_factor)
    
    def _recursive_pricer(self, node, df):
        """
        Fonction récursive interne pour calculer le prix de l'option.
        
        Cette fonction implémente l'algorithme de backward induction de façon récursive:
        
        Algorithme:
        -----------
        1. **Cas de base (mémoïsation)**: Si le nœud a déjà un prix calculé, le retourner
        2. **Cas de base (feuille)**: Si c'est un nœud terminal, calculer et retourner le payoff
        3. **Cas récursif**:
           a. Calculer récursivement les prix des nœuds suivants (up, mid, down)
           b. Calculer l'espérance: E[V] = p_up × V_up + p_mid × V_mid + p_down × V_down
           c. Actualiser: V_continuation = exp(-r×dt) × E[V]
           d. Pour une option américaine: V = max(V_continuation, V_intrinsèque)
           e. Pour une option européenne: V = V_continuation
        
        Paramètres:
        -----------
        node : Node
            Nœud actuel à évaluer
        df : float
            Facteur d'actualisation: exp(-r×dt)
            
        Retourne:
        ---------
        float
            Prix de l'option au nœud
        """
        # ===== Cas 1: Prix déjà calculé =====
        # Évite de recalculer un nœud déjà visité
        # Important pour l'efficacité car plusieurs chemins peuvent mener au même nœud
        if hasattr(node, 'si2') and node.si2 is not None:
            return node.si2
        
        # ===== Cas 2: Nœud terminal (feuille de l'arbre à maturité) =====
        if node.next_mid is None:
            # Calcul du payoff terminal à maturité
            # Call: max(0, S - K)
            # Put: max(0, K - S) = max(0, -(S - K))
            payoff = max(0, (node.si - self.strike) * self.op_multiplicator)
            node.si2 = payoff
            return payoff

        # ===== Cas 3: Nœud intermédiaire - calcul récursif =====
        
        # --- Calcul de l'espérance des valeurs futures ---
        # Commence avec la contribution du nœud mid (généralement la plus probable)
        EV = node.p_mid * self._recursive_pricer(node.next_mid, df)

        # Ajout de la contribution du nœud up (si il existe)
        # Certains nœuds (pruning) peuvent ne pas avoir de branche up
        if node.next_up is not None:
            EV += node.p_up * self._recursive_pricer(node.next_up, df)
        
        # Ajout de la contribution du nœud down (si il existe)
        if node.next_down is not None:
            EV += node.p_down * self._recursive_pricer(node.next_down, df)

        # --- Actualisation de l'espérance ---
        # Valeur de continuation = valeur présente de l'espérance future
        # continuation_value = exp(-r×dt) × E[V_futur]
        continuation_value = df * EV
        
        # ===== Gestion du style d'exercice =====
        if self.op_exercice == "US":
            # **Option américaine**: peut être exercée à tout moment
            # On compare la valeur de continuation avec la valeur d'exercice immédiat
            
            # Valeur intrinsèque = payoff si exercé maintenant
            intrinsic = max(0, (node.si - self.strike) * self.op_multiplicator)
            
            # Décision optimale: max(continuer à détenir, exercer maintenant)
            # Si intrinsic > continuation_value, il est optimal d'exercer
            node.si2 = max(continuation_value, intrinsic)
        else:
            # **Option européenne**: exercice uniquement à maturité
            # Pas de décision d'exercice anticipé
            node.si2 = continuation_value
        
        return node.si2

    def price_iteratively(self, arbre, type_option=None, style_option=None):
        """
        Calcule le prix de l'option en utilisant une approche itérative (backward induction).
        
        Méthode de calcul:
        ------------------
        1. Part des feuilles de l'arbre (maturité)
        2. Calcule les payoffs terminaux pour tous les nœuds finaux
        3. Remonte étape par étape vers la racine (backward)
        4. À chaque nœud, calcule l'espérance actualisée des valeurs futures
        5. Pour les options américaines, compare avec l'exercice immédiat
        
        Avantages:
        ----------
        - Plus efficace en mémoire (pas de pile d'appels récursifs)
        - Plus rapide pour de grands arbres
        - Pas de risque de stack overflow
        - Meilleure utilisation du cache CPU (parcours séquentiel)
        
        Structure du parcours:
        ----------------------
        Pour chaque colonne (de droite à gauche):
          Pour chaque nœud de la colonne (haut puis bas):
            Calculer la valeur actualisée
        
        Paramètres:
        -----------
        arbre : Arbre
            Arbre trinomial construit
        type_option : str, optional
            Surcharge du type d'option (par défaut: utilise arbre.contract.op_type)
            Permet de pricer plusieurs types d'options avec le même arbre
        style_option : str, optional
            Surcharge du style d'exercice (par défaut: utilise arbre.contract.op_exercice)
            
        Retourne:
        ---------
        float
            Prix de l'option (valeur à la racine)
        """
        # ===== Récupération des paramètres de l'arbre et du contrat =====
        r = arbre.market.int_rate      # Taux d'intérêt sans risque
        N = arbre.n_steps              # Nombre d'étapes dans l'arbre
        T = arbre.contract.maturity    # Maturité en années
        K = arbre.contract.strike      # Prix d'exercice (strike)
        dt = T / N                     # Pas de temps par étape
        
        # ===== Calcul du facteur d'actualisation =====
        d_f = math.exp(-r * dt)
        
        # ===== Gestion des surcharges de paramètres =====
        type_op = type_option if type_option is not None else arbre.contract.op_type
        op_exercice = style_option if style_option is not None else arbre.contract.op_exercice
        op_multiplicator = 1 if type_op == "Call" else -1
        
        # ===== Positionnement au dernier nœud (feuille la plus à droite) =====
        last_node = arbre.racine
        while last_node.next_mid is not None:
            last_node = last_node.next_mid

        # ===== Étape 1: Initialisation des payoffs terminaux =====
        # Calcule le payoff à maturité pour tous les nœuds de la dernière colonne
        self._set_payoff(last_node, K, op_multiplicator)
        
        # ===== Étape 2: Backward induction =====
        # Remonte colonne par colonne en calculant les valeurs actualisées
        self._roll_back(last_node, d_f, op_exercice, K, op_multiplicator)
        
        # ===== Retour du prix à la racine =====
        return arbre.racine.si2

    def _set_payoff(self, last_node, K, op_multiplicator):
        """
        Initialise les payoffs terminaux à la maturité pour tous les nœuds de la dernière colonne.
        
        Cette méthode parcourt tous les nœuds de la dernière étape (maturité) et calcule
        leur payoff selon le type d'option:
        
        Formules du payoff:
        -------------------
        - **Call**: max(0, S - K) 
          → Profit si le sous-jacent est au-dessus du strike
        - **Put**: max(0, K - S)
          → Profit si le sous-jacent est en-dessous du strike
        
        Le parcours se fait en deux phases:
        1. Remontée: depuis last_node vers le haut (voisin_up)
        2. Descente: depuis last_node vers le bas (voisin_down)
        
        Paramètres:
        -----------
        last_node : Node
            Nœud du milieu de la dernière colonne (étape N)
        K : float
            Prix d'exercice (strike)
        op_multiplicator : int
            +1 pour Call, -1 pour Put
        """
        # ===== Remontée: calcul des payoffs pour les nœuds au-dessus =====
        current_node = last_node
        while current_node is not None:
            current_node.si2 = max((current_node.si - K) * op_multiplicator, 0)
            current_node = current_node.voisin_up
        
        # ===== Descente: calcul des payoffs pour les nœuds en-dessous =====
        current_node = last_node
        while current_node is not None:
            current_node.si2 = max((current_node.si - K) * op_multiplicator, 0)
            current_node = current_node.voisin_down

    def _roll_back(self, last_node, d_f, op_exercice, K, op_multiplicator):
        """
        Remonte dans l'arbre en calculant les valeurs d'option par backward induction.
        
        Cette méthode implémente l'algorithme de backward induction:
        
        Pour chaque colonne (de droite à gauche):
            Pour chaque nœud de la colonne:
                1. Calculer l'espérance des valeurs futures
                2. Actualiser cette espérance
                3. Pour option américaine: comparer avec exercice immédiat
        
        Paramètres:
        -----------
        last_node : Node
            Nœud du milieu de la dernière colonne
        d_f : float
            Facteur d'actualisation: exp(-r×dt)
        op_exercice : str
            "EU" (Européenne) ou "US" (Américaine)
        K : float
            Prix d'exercice (strike)
        op_multiplicator : int
            +1 pour Call, -1 pour Put
        """
        # ===== Boucle principale: parcours backward des colonnes =====
        # Remonte de la dernière colonne vers la racine
        while last_node is not None:
            # Passage à la colonne précédente (vers la gauche)
            last_node = last_node.voisin_behind
            
            # ===== Traitement de la branche HAUTE (nœuds au-dessus du milieu) =====
            current_node = last_node
            while current_node is not None:
                
                # --- Calcul de l'espérance actualisée ---
                if current_node.p_up == None:
                    # **Cas particulier: pruning activé**
                    # Le nœud n'a pas de probabilités calculées (branche peu probable)
                    # On utilise simplement la valeur du nœud mid suivant actualisée
                    val = current_node.next_mid.si2 * d_f
                else:
                    # **Cas général: calcul de l'espérance complète**
                    # E[V] = p_up × V_up + p_mid × V_mid + p_down × V_down
                    u = current_node.p_up
                    pm = current_node.p_mid
                    d = current_node.p_down
                    
                    # Calcul de l'espérance pondérée et actualisation
                    val = (u * current_node.next_mid.voisin_up.si2 +
                           pm * current_node.next_mid.si2 +
                           d * current_node.next_mid.voisin_down.si2) * d_f
                
                # --- Gestion du style d'exercice ---
                if op_exercice == "US":
                    # **Option américaine**: possibilité d'exercice anticipé
                    # Calculer la valeur intrinsèque (exercice immédiat)
                    intrinsic = max((current_node.si - K) * op_multiplicator, 0)
                    
                    # Décision optimale: max(continuer, exercer)
                    # Si intrinsic > val, il est optimal d'exercer maintenant
                    current_node.si2 = max(val, intrinsic)
                else:
                    # **Option européenne**: pas d'exercice anticipé
                    # On garde simplement la valeur de continuation
                    current_node.si2 = val
                
                # Passage au nœud supérieur
                current_node = current_node.voisin_up
            
            # ===== Traitement de la branche BASSE (nœuds en-dessous du milieu) =====
            # Même logique que pour la branche haute
            current_node = last_node
            while current_node is not None:
                
                # --- Calcul de l'espérance actualisée ---
                if current_node.p_up == None:
                    # Cas pruning
                    val = current_node.next_mid.si2 * d_f
                else:
                    # Cas général
                    u = current_node.p_up
                    pm = current_node.p_mid
                    d = current_node.p_down
                    val = (u * current_node.next_mid.voisin_up.si2 +
                           pm * current_node.next_mid.si2 +
                           d * current_node.next_mid.voisin_down.si2) * d_f
                
                # --- Gestion du style d'exercice ---
                if op_exercice == "US":
                    # Option américaine
                    intrinsic = max((current_node.si - K) * op_multiplicator, 0)
                    current_node.si2 = max(val, intrinsic)
                else:
                    # Option européenne
                    current_node.si2 = val
                
                # Passage au nœud inférieur
                current_node = current_node.voisin_down