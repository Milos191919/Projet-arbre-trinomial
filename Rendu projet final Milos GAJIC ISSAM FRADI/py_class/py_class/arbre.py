import math
from .node import Node

class Arbre:

    #Classe représentant un arbre trinomial pour le pricing d'options.


    
    def __init__(self, market, contract, n_steps: int):
        """
        Initialisation de l'arbre trinomial.
        
        Paramètres:
        -----------
        market : Market
            Objet contenant les paramètres de marché (prix, taux, volatilité, dividendes)
        contract : Contract
            Objet contenant les caractéristiques du contrat d'option
        n_steps : int
            Nombre d'étapes (périodes) dans l'arbre
        """
        self.market = market
        self.contract = contract
        self.n_steps = n_steps
        
        # Calcul du pas de temps (dt) : durée de chaque période en années
        self.dt = self.contract.maturity / n_steps
        
        # Calcul du coefficient alpha pour le facteur de hausse/baisse
        self.alpha = math.exp(self.market.sigma * math.sqrt(3 * self.dt))
        
        # Initialisation de la racine (nœud initial) de l'arbre
        self.racine = None
        
        # Génération complète de l'arbre trinomial
        self._generer_arbre(n_steps)

    def is_dividend(self, step: int) -> bool:
        """
        Vérifie si un dividende est détaché pendant l'étape donnée.
        
        Paramètres:
        -----------
        step : int
            Numéro de l'étape à vérifier (0 à n_steps-1)
            
        Retourne:
        ---------
        bool
            True si un dividende est détaché pendant cette étape, False sinon
        """
        # Si aucune date de dividende n'est définie, retourner False
        if self.market.div_date is None:
            return False
        
        # Calcul du temps de début et de fin de l'étape
        t_start = step * self.dt
        t_end = (step + 1) * self.dt
        
        # Conversion de la date de dividende en temps (années depuis pricing_date)
        div_date = (self.market.div_date - self.contract.pricing_date).days / 365.0
        
        # Vérification si le dividende tombe dans l'intervalle de temps de cette étape
        return t_start < div_date <= t_end

    def _generer_arbre(self, N: int):
        """
        Génère l'arbre trinomial complet en construisant tous les nœuds et leurs connexions.
        
        Algorithme:
        -----------
        1. Initialisation de la racine avec le prix du sous-jacent actuel
        2. Pour chaque étape k de 0 à N-1:
           a. Vérifie si un dividende est détaché
           b. Construit la partie haute de l'arbre (nœuds au-dessus du tronc)
           c. Construit la partie basse de l'arbre (nœuds en-dessous du tronc)
           d. Avance le tronc d'une étape
        
        Paramètres:
        -----------
        N : int
            Nombre d'étapes dans l'arbre
        """
        # ===== Définition des paramètres initiaux =====
        S0 = self.market.stock_price  # Prix initial du sous-jacent
        r = self.market.int_rate      # Taux d'intérêt sans risque
        dt = self.dt                  # Pas de temps
        
        # Création du nœud racine avec le prix initial
        self.racine = Node(si=S0, arbre=self)
        self.racine.proba_cumule = 1  # Probabilité cumulée = 1 à la racine
        
        # Le tronc représente le chemin central de l'arbre
        noeud_tronc = self.racine

        # ===== Boucle principale : construction étape par étape =====
        for k in range(N):
            # Vérification si un dividende est détaché pendant cette étape
            if self.is_dividend(k):
                D = self.market.div  # Montant du dividende
                print(f'div {D}, {k}, {self.market.div_date}, {(self.market.div_date - self.contract.pricing_date).days / 365.0}')
            else:
                D = 0  # Pas de dividende
            
            # Point de départ pour traiter les nœuds de l'étape k
            noeud_a_traiter = noeud_tronc
            
            # Calcul de la valeur forward pour le nœud du tronc
            # Forward = S * exp(r*dt) - Dividende
            forward_val = noeud_tronc.si * math.exp(r * dt) - D
            
            # Création du nœud mid suivant avec la valeur forward
            last_next_mid = Node(si=forward_val, arbre=self)
            
            # Mémorisation du prochain nœud up pour la construction
            noeud_next_up_memory = last_next_mid

            # ===== Construction de la partie HAUTE de l'arbre =====
            # Remonte depuis le tronc vers le haut (voisin_up)
            while noeud_a_traiter is not None:
                # Calcul de la valeur forward pour ce nœud
                forward = noeud_a_traiter.si * math.exp(r * dt) - D
                
                # Fonction qui crée et/ou lie les nœuds de la colonne suivante
                # Cette fonction connecte le nœud actuel à ses trois descendants possibles
                last_next_mid = noeud_a_traiter.set_next_mid(forward, noeud_next_up_memory, D)
                
                # Passage au nœud supérieur dans la colonne actuelle
                noeud_a_traiter = noeud_a_traiter.voisin_up
                
                # Mise à jour de la mémoire du nœud up suivant
                noeud_next_up_memory = last_next_mid.voisin_up
                
                # Si on a besoin d'un nouveau nœud up et qu'il reste des nœuds à traiter
                # Ceci est nécessaire pour le prunning
                if noeud_next_up_memory is None and noeud_a_traiter is not None:
                    # Création d'un nouveau nœud up avec un prix plus élevé (multiplication par alpha)
                    noeud_next_up_memory = Node(si=last_next_mid.si * self.alpha, arbre=self)
                    
                    # Liaison bidirectionnelle entre les nœuds up et mid
                    noeud_next_up_memory.voisin_down = last_next_mid
                    last_next_mid.voisin_up = noeud_next_up_memory

            # ===== Préparation pour la construction de la partie BASSE =====
            # Retour au tronc et descente vers le bas
            noeud_a_traiter = noeud_tronc.voisin_down
            last_next_mid = noeud_tronc.next_mid
            
            # Mémorisation du prochain nœud down
            noeud_next_down_memory = last_next_mid.voisin_down

            # ===== Construction de la partie BASSE de l'arbre =====
            # Descend depuis le tronc vers le bas (voisin_down)
            while noeud_a_traiter is not None:
                # Calcul de la valeur forward pour ce nœud
                forward = noeud_a_traiter.si * math.exp(r * dt) - D
                
                # Création et liaison des nœuds suivants
                last_next_mid = noeud_a_traiter.set_next_mid(forward, noeud_next_down_memory, D)
                
                # Passage au nœud inférieur dans la colonne actuelle
                noeud_a_traiter = noeud_a_traiter.voisin_down
                
                # Mise à jour de la mémoire du nœud down suivant
                noeud_next_down_memory = last_next_mid.voisin_down
                
                # Si on a besoin d'un nouveau nœud down et qu'il reste des nœuds à traiter
                if noeud_next_down_memory is None and noeud_a_traiter is not None:
                    # Création d'un nouveau nœud down avec un prix plus bas (division par alpha)
                    noeud_next_down_memory = Node(si=last_next_mid.si / self.alpha, arbre=self)
                    
                    # Liaison bidirectionnelle entre les nœuds down et mid
                    noeud_next_down_memory.voisin_up = last_next_mid
                    last_next_mid.voisin_down = noeud_next_down_memory

            # ===== Avancement du tronc pour l'étape suivante =====
            # Le nouveau tronc est le nœud mid suivant de l'ancien tronc
            noeud_tronc = noeud_tronc.next_mid