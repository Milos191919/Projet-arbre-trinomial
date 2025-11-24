import numpy as np 
from scipy.stats import norm

from typing import Callable, cast


from .arbre import Arbre
from .option import Contract
from .market import Market

def BS(S, K, T, r, sigma,type_op,D,T_div):
    # Ajustement du sous-jacent
    S_adj = S - D* np.exp(-r * T_div)

    # Calculs standard de Black-Scholes
    N = norm.cdf
    d1 = (np.log(S_adj / K) + (r + sigma**2 / 2) * T) / (sigma * np.sqrt(T))
    d2 = d1 - sigma * np.sqrt(T)

    if type_op == "Call":
        prix = S_adj * N(d1) - K * np.exp(-r * T) * N(d2)
    elif type_op == "Put":
        prix = K * np.exp(-r * T) * N(-d2) - S_adj * N(-d1)
    else:
        raise ValueError("type_op doit être 'Call' ou 'Put'")

    return prix


class OptionPricingParam:
    UND_SHIFT: float = 0.001 
    VOL_SHIFT: float = 0.005
    VOLGA_SHIFT: float = 0.01

    def __init__(self, market, contract, n_steps: int):
        self.stock_price = market.stock_price
        self.int_rate = market.int_rate
        self.sigma = market.sigma
        self.div = market.div
        self.div_date = market.div_date
        
        self.pricing_date = contract.pricing_date
        self.maturity_date = contract.maturity_date
        self.strike = contract.strike
        self.op_type = contract.op_type
        self.op_exercice = contract.op_exercice
        
        self.n_steps = n_steps


def _PriceTreeBackward_OneDimPrice(params: OptionPricingParam, stock_price: float) -> float:

    market = Market(
        stock_price=stock_price,
        int_rate=params.int_rate,
        sigma=params.sigma,
        div=params.div,
        div_date=params.div_date
    )
    
    contract = Contract(
        pricing_date=params.pricing_date,
        maturity_date=params.maturity_date,
        strike=params.strike,
        op_type=params.op_type,
        op_exercice=params.op_exercice
    )
    
    tree = Arbre(market=market, contract=contract, n_steps=params.n_steps)
    
    price = contract.price_iteratively(tree)
    
    return price

def _PriceTreeBackward_OneDimSigma(params: OptionPricingParam, sigma: float) -> float:    
    market = Market(
        stock_price=params.stock_price, 
        int_rate=params.int_rate,
        sigma=sigma, 
        div=params.div,
        div_date=params.div_date
    )
    
    contract = Contract(
        pricing_date=params.pricing_date,
        maturity_date=params.maturity_date,
        strike=params.strike,
        op_type=params.op_type,
        op_exercice=params.op_exercice
    )
    
    tree = Arbre(market=market, contract=contract, n_steps=params.n_steps)
    price = contract.price_iteratively(tree)
    
    return price

class OneDimDerivative:

    def __init__(self, function: Callable[[object, float], float],
                 other_parameters: object, shift: float = 1):
        self.f: Callable[[object, float], float] = function
        self.param: object = other_parameters
        self.shift: float = shift

    def first(self, x: float) -> float:
        """
        Calculates the first derivative: (f(x+h) - f(x-h)) / (2h)
        """
        price_up = self.f(self.param, x + self.shift)
        price_down = self.f(self.param, x - self.shift)
        
        return (price_up - price_down) / (2 * self.shift)

    def second(self, x: float) -> float:
        """
        Calculates the second derivative: (f(x+h) + f(x-h) - 2*f(x)) / h^2
        """
        price_up = self.f(self.param, x + self.shift)
        price_down = self.f(self.param, x - self.shift)
        price_mid = self.f(self.param, x) 
        
        return (price_up + price_down - 2 * price_mid) / (self.shift ** 2)

def calculate_delta(market, contract, n_steps: int) -> float:

    
    params = OptionPricingParam(market, contract, n_steps)
    
    base_stock_price = market.stock_price
    shift = base_stock_price * OptionPricingParam.UND_SHIFT

    pricer_function = cast(Callable[[object, float], float], _PriceTreeBackward_OneDimPrice)
    
    delta_calculator = OneDimDerivative(
        function=pricer_function,
        other_parameters=params,
        shift=shift
    )
    
    delta = delta_calculator.first(base_stock_price)
    
    return delta

def calculate_gamma(market: Market, contract: Contract, n_steps: int) -> float:
    """
    Calculates the option's Gamma.
    """
    params = OptionPricingParam(market, contract, n_steps)
    base_stock_price = market.stock_price
    #shift = base_stock_price * OptionPricingParam.UND_SHIFT (formule initiale=
    #shift corrigé
    shift = base_stock_price * 0.05
    
    pricer_function = cast(Callable[[object, float], float], _PriceTreeBackward_OneDimPrice)
    
    gamma_calculator = OneDimDerivative(
        function=pricer_function,
        other_parameters=params,
        shift=shift
    )
    
    gamma = gamma_calculator.second(base_stock_price) 
    return gamma


def calculate_vega(market: Market, contract: Contract, n_steps: int) -> float:
    """
    f(sigma + 0.5%) - f(sigma - 0.5%)
    """
    params = OptionPricingParam(market, contract, n_steps)
    base_sigma = market.sigma
    shift = OptionPricingParam.VOL_SHIFT 
    
    pricer_function = cast(Callable[[object, float], float], _PriceTreeBackward_OneDimSigma)
    
    price_up = pricer_function(params, base_sigma + shift)
    price_down = pricer_function(params, base_sigma - shift)
    
    return price_up - price_down


def calculate_volga(market: Market, contract: Contract, n_steps: int) -> float:
    """
    Vega(sigma + shift) - Vega(sigma)
    """
    base_sigma = market.sigma
    #volga_shift = OptionPricingParam.VOLGA_SHIFT formule initiale
    # shift corrigé
    volga_shift= 0.1
    market_up = Market(
        stock_price=market.stock_price,
        int_rate=market.int_rate,
        sigma=base_sigma + volga_shift,
        div=market.div,
        div_date=market.div_date
    )
    vega_up = calculate_vega(market_up, contract, n_steps)
    vega_base = calculate_vega(market, contract, n_steps)
    return vega_up - vega_base


def calculate_vanna(market: Market, contract: Contract, n_steps: int) -> float:
    """
    Vanna ≈ [ f(S+hS,σ+hσ) - f(S+hS,σ-hσ) - f(S-hS,σ+hσ) + f(S-hS,σ-hσ) ] / (4 hS hσ)
    """
    base_S = market.stock_price
    base_sigma = market.sigma

    hS = base_S * OptionPricingParam.UND_SHIFT

    #hV = OptionPricingParam.VOL_SHIFT (formule initiale)
    #shift corrigé
    hV=0.05

    # On suit le pattern existant: on varie S via _PriceTreeBackward_OneDimPrice
    # et σ en modifiant params.sigma
    pricer_function = cast(Callable[[object, float], float], _PriceTreeBackward_OneDimPrice)

    # Params pour sigma up / sigma down
    params_sigma_up = OptionPricingParam(market, contract, n_steps)
    params_sigma_up.sigma = base_sigma + hV

    params_sigma_down = OptionPricingParam(market, contract, n_steps)
    params_sigma_down.sigma = base_sigma - hV

    # Coins du schéma mixte
    f_pp = pricer_function(params_sigma_up,   base_S + hS)  # +S, +σ
    f_pm = pricer_function(params_sigma_down, base_S + hS)  # +S, -σ
    f_mp = pricer_function(params_sigma_up,   base_S - hS)  # -S, +σ
    f_mm = pricer_function(params_sigma_down, base_S - hS)  # -S, -σ

    vanna = (f_pp - f_pm - f_mp + f_mm) / (4.0 * hS * hV)
    return vanna