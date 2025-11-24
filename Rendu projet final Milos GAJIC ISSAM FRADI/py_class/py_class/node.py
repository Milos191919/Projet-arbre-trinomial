import math as m

class Node:
    """
    Classe représentant un nœud dans l'arbre trinomial.
    
    Chaque nœud contient:
    - Le prix du sous-jacent (si)
    - Les connexions avec les nœuds voisins (up, down) et suivants (next_up, next_mid, next_down)
    - Les probabilités de transition (p_up, p_mid, p_down)
    - Le prix de l'option (si2)
    - La probabilité cumulée d'atteindre ce nœud
    """
    
    def __init__(self, si, arbre, voisin_up=None, voisin_down=None,
                 p_up=None, p_down=None, p_mid=None):
        """
        Initialise un nœud de l'arbre trinomial.
        
        Paramètres:
        -----------
        si : float
            Prix du sous-jacent à ce nœud
        arbre : Arbre
            Référence vers l'arbre trinomial parent
        voisin_up : Node, optional
            Nœud voisin du dessus (même étape, prix plus élevé)
        voisin_down : Node, optional
            Nœud voisin du dessous (même étape, prix plus bas)
        p_up : float, optional
            Probabilité de transition vers le nœud up suivant
        p_down : float, optional
            Probabilité de transition vers le nœud down suivant
        p_mid : float, optional
            Probabilité de transition vers le nœud mid suivant
        """
        # ===== Informations sur le prix du sous-jacent =====
        self.si = si  # Prix du sous-jacent à ce nœud
        
        # ===== Référence vers l'arbre parent =====
        self.arbre = arbre
        
        # ===== Connexions horizontales (même étape temporelle) =====
        self.voisin_up = voisin_up      # Nœud au-dessus (prix plus élevé)
        self.voisin_down = voisin_down  # Nœud en-dessous (prix plus bas)
        
        # ===== Probabilités de transition risque-neutre =====
        self.p_up = p_up      # Probabilité d'aller vers next_up
        self.p_down = p_down  # Probabilité d'aller vers next_down
        self.p_mid = p_mid    # Probabilité d'aller vers next_mid
        
        # ===== Connexions vers l'étape suivante =====
        self.next_up = None    # Nœud suivant vers le haut
        self.next_down = None  # Nœud suivant vers le bas
        self.next_mid = None   # Nœud suivant au milieu
        
        # ===== Connexion vers l'étape précédente =====
        self.voisin_behind = None  # Lien vers le nœud précédent (utilisé pour backward pricing)
        
        # ===== Prix de l'option =====
        self.si2 = None  # Prix de l'option à ce nœud 
        
        # ===== Probabilité cumulée =====
        self.proba_cumule = 0  # Probabilité d'atteindre ce nœud depuis la racine

    def move_up(self, alpha, vosin_multiple=True):
        """
        Crée (si nécessaire) et retourne le nœud voisin du dessus.
        """
        # Si le voisin up n'existe pas et que vosin_multiple est True
        if not self.voisin_up and vosin_multiple:
            self.voisin_up = Node(self.si * alpha, arbre=self.arbre)
            
            # Liaison bidirectionnelle : le nouveau nœud up pointe vers ce nœud comme voisin down
            self.voisin_up.voisin_down = self
            
        return self.voisin_up

    def move_down(self, alpha, vosin_multiple=True):
        """
        Crée (si nécessaire) et retourne le nœud voisin du dessous.
        """
        # Si le voisin down n'existe pas et que vosin_multiple est True
        if not self.voisin_down and vosin_multiple:
            self.voisin_down = Node(self.si / alpha, arbre=self.arbre)
            
            # Liaison bidirectionnelle : le nouveau nœud down pointe vers ce nœud comme voisin up
            self.voisin_down.voisin_up = self
            
        return self.voisin_down

    def compute_probabilities(self, D=0):
        """
        Calcule les probabilités de transition risque-neutre (p_up, p_mid, p_down).
            
        Formules:
        ---------
        esp = S × exp(r×dt) - D                        
        var = S² × exp(2r×dt) × (exp(σ²×dt) - 1)        
        
        Les probabilités sont déduites en résolvant un système d'équations basé sur:
        - L'espérance: p_up × S_up + p_mid × S_mid + p_down × S_down = esp
        - La variance: E[S²] - E[S]² = var
        - La normalisation: p_up + p_mid + p_down = 1
        """
        # ===== Récupération des paramètres de l'arbre =====
        r = self.arbre.market.int_rate     # Taux d'intérêt sans risque
        dt = self.arbre.dt                 # Pas de temps 
        alpha = self.arbre.alpha           # Alpha de l'arbre
        sigma = self.arbre.market.sigma    # Volatilité annualisée
        
        # ===== Calcul de l'espérance forward =====
        esp = self.si * m.exp(r * dt) - D
        
        # ===== Calcul de la variance =====
        var = (self.si ** 2) * m.exp(2 * r * dt) * (m.exp((sigma ** 2) * dt) - 1)
        
        # ===== Prix du nœud mid suivant =====
        nmv = self.next_mid.si
        
        # ===== Calcul de p_down (probabilité de baisse) =====
        numer = (1 / nmv ** 2) * (var + esp ** 2) - 1 - (alpha + 1) * ((1 / nmv) * esp - 1)
        denom = (1 - alpha) * ((1 / alpha ** 2) - 1)
        self.p_down = numer / denom

        # ===== Calcul de p_up (probabilité de hausse) =====
        numer_up = (1 / nmv) * esp - 1 - ((1 / alpha) - 1) * self.p_down
        denom_up = alpha - 1
        self.p_up = numer_up / denom_up
        
        # ===== Calcul de p_mid (probabilité centrale) =====
        self.p_mid = 1 - self.p_up - self.p_down

    def set_next_mid(self, forward, noeud_next_up_memory, D=0):
        """
        Définit le nœud mid suivant et établit toutes les connexions nécessaires.
        
        Paramètres:
        -----------
        forward : float
            Valeur forward du sous-jacent : S × exp(r×dt) - D
            C'est la valeur cible autour de laquelle on construit les branches
        noeud_next_up_memory : Node
            Nœud de référence pour commencer la recherche du next_mid
            Permet de parcourir efficacement la structure de l'arbre
        D : float, optional
            Montant du dividende (par défaut 0)
            Affecte le calcul des probabilités de transition
        """
        alpha = self.arbre.alpha

        # ===== Pruning =====
        if self.proba_cumule < 10**-8:
            self.creation_noeud_unique(alpha, forward, noeud_next_up_memory)
            return self.next_mid

        # ===== Recherche du nœud mid suivant approprié =====
        # Trouve le nœud existant dont le prix est le plus proche de la valeur forward
        # Cela garantit que l'arbre respecte bien la dérive du sous-jacent
        self.next_mid = self.find_next_mid(forward, alpha, noeud_next_up_memory)
        
        # Liaison arrière 
        self.next_mid.voisin_behind = self

        # ===== Création des nœuds up et down suivants =====
        self.next_up = self.next_mid.move_up(alpha)
        self.next_down = self.next_mid.move_down(alpha)

        # ===== Calcul des probabilités de transition =====
        self.compute_probabilities(D)

        # ===== Propagation des probabilités cumulées =====
        self.next_mid.proba_cumule += self.proba_cumule * self.p_mid
        self.next_up.proba_cumule += self.proba_cumule * self.p_up
        self.next_down.proba_cumule += self.proba_cumule * self.p_down
        
        return self.next_mid

    def find_next_mid(self, forward, alpha, start_node):
        """
        Trouve le nœud mid suivant dont le prix est le plus proche de la valeur forward.
        -----------
        forward : float
            Valeur forward cible (S × exp(r×dt) - D)
        alpha : float
            Facteur multiplicatif/diviseur de l'arbre
        start_node : Node
            Nœud de départ pour la recherche
            
        """
        node = start_node
        
        # ===== Phase de montée =====
        # Monte tant que forward est au-dessus du milieu entre le nœud actuel et le nœud up
        while forward > node.si * (1 + alpha) / 2:
            node = node.move_up(alpha)
        
        # ===== Phase de descente =====
        # Descend tant que forward est en-dessous du milieu entre le nœud actuel et le nœud down
        while forward <= node.si * (1 + 1 / alpha) / 2:
            node = node.move_down(alpha)
        
        return node

    def creation_noeud_unique(self, alpha, forward, noeud_next_up_memory):
        """
        Crée un nœud unique lorsque la probabilité cumulée est négligeable (pruning).
        
        Cette méthode est une optimisation qui simplifie l'arbre en ne créant qu'un seul
        chemin pour les branches très improbables (proba_cumule < 10^-8).
        
        Au lieu de créer trois branches (up, mid, down) avec leurs probabilités,
        on crée un seul nœud mid avec probabilité 1. Cela permet de:
        - Réduire le nombre de nœuds dans l'arbre
        - Accélérer les calculs
        - Maintenir une précision acceptable (les branches éliminées contribuent < 10^-8)
        
        Avantages du pruning:
        ---------------------
        - Réduction de la mémoire utilisée
        - Accélération du backward pricing
        - Pas de perte significative de précision
        
        Paramètres:
        -----------
        alpha : float
            Facteur multiplicatif/diviseur de l'arbre
        forward : float
            Valeur forward du sous-jacent
        noeud_next_up_memory : Node
            Nœud de référence pour la recherche du next_mid
        """
        # ===== Recherche du nœud mid suivant =====
        # Même processus que dans set_next_mid, mais sans créer les branches up/down
        self.next_mid = self.find_next_mid(forward, alpha, noeud_next_up_memory)
        
        # ===== Probabilité unique =====
        # Toute la probabilité (négligeable) va vers le nœud mid
        # p_up = 0, p_down = 0, p_mid = 1
        self.p_mid = 1

        # ===== Création des voisins up et down (sans connexions multiples) =====
        # Le paramètre False empêche la création de nœuds voisins supplémentaires
        self.next_mid.move_up(alpha, False)
        self.next_mid.move_down(alpha, False)
        
        # ===== Liaison arrière =====
        # Maintien de la connexion vers le nœud précédent pour le backward pricing
        self.next_mid.voisin_behind = self
        
        # ===== Propagation de la probabilité cumulée =====
        self.next_mid.proba_cumule += self.proba_cumule * self.p_mid