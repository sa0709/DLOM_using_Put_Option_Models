import pandas as pd
import numpy as np
from scipy.stats import norm

# ----------------------------
#  Discount for Lack of Marketability (DLOM) Models
# ----------------------------

def chaffe(S, T, r, q, sigma):
    """
    Chaffe model for estimating the put option value and DLOM.
    Uses a modified Black-Scholes structure.
    """
    d1 = ((r - q + 0.5 * sigma**2) * T) / (sigma * np.sqrt(T))
    d2 = d1 - sigma * np.sqrt(T)

    # Call option value component
    c = S * (np.exp(-q*T) * norm.cdf(d1) - np.exp(-r*T) * norm.cdf(d2))

    # Put option value under Chaffe method
    p = c + S * (np.exp(-r*T) - 1)

    return p, p/S  # DLOM

def finnerty(S, T, r, q, sigma):
    """
    Finnerty model using a volatility-adjusted distribution width formula.
    """
    var = np.sqrt(
        sigma**2*T +
        np.log(2*(np.exp(T*sigma**2) - sigma**2*T - 1)) -
        2*np.log(np.exp(T*sigma**2) - 1)
    )

    # Put option value under Finnerty method
    p = S * np.exp(-q*T) * (norm.cdf(var/2) - norm.cdf(-var/2))

    return p, p/S  # DLOM

def ghaidarov(S, T, r, q, sigma):
    """
    Ghaidarov model based on log-normal distribution width parameters.
    """
    var = np.sqrt(
        np.log(2*(np.exp(T*sigma**2) - sigma**2*T - 1)) -
        2*np.log(sigma**2*T)
    )

    # Put option value under Ghaidarov method
    p = S * np.exp(-q*T) * (2*norm.cdf(var/2) - 1)

    return p, p/S  # DLOM

def longstaff(S, T, r, q, sigma):
    """
    Longstaff model based on expected payoff under limited marketability.
    """
    x = sigma**2 * T / 2  # Helper variable for formula simplification

    # Put option value under Longstaff method
    p = S * (
        (2 + x) * norm.cdf(np.sqrt(x)) +
        np.sqrt(x / np.pi) * np.exp(-x/4) -
        1
    )

    return p, p/(1+p)  # Longstaff DLOM formula

# ----------------------------
#  Mapping of Models to Rows in Input File
# ----------------------------

models = {
    "Chaffe Model": (chaffe, 5),
    "Finnerty Model": (finnerty, 6),
    "Ghaidarov Model": (ghaidarov, 7),
    "Longstaff Model": (longstaff, 8)
}

# ----------------------------
#  Read Input File
# ----------------------------

df = pd.read_excel('Input.xlsx', index_col=0)

# Extract key parameters (S, r, T, sigma, q)
S, r, T, sigma, q = df.iloc[:5, 0]

# Store final DLOM results and individual model DataFrames
results = {}
models_df = {}

# ----------------------------
#  Run Each Model if Flag = "Yes"
# ----------------------------

for name, (fn, row) in models.items():

    # Only compute the model if the input says "Yes"
    if df.iloc[row, 0] == "Yes":

        # Compute put option value and DLOM
        p, d = fn(S, T, r, q, sigma)

        # Prepare model-specific dataframe
        fn_df = df.iloc[:5, 0].copy()
        fn_df.loc['Put Option Value'] = p
        fn_df.loc['Discount for Lack of Marketability (%)'] = d

        models_df[name] = fn_df
        results[name] = d

# ----------------------------
#  Prepare Combined Output
# ----------------------------

# Summary table for DLOM across models
df_output = pd.DataFrame.from_dict(results, orient='index', columns=['DLOM'])
df_output.loc["Median"] = df_output.median()

# Combine all model outputs
df_models_output = pd.DataFrame.from_dict(models_df, orient='index')

# ----------------------------
#  Write to Excel
# ----------------------------

excel_file_path = 'Output.xlsx'

with pd.ExcelWriter(excel_file_path) as writer:
    # Write individual model results
    df_models_output.transpose().to_excel(writer, sheet_name='Models', index=True)

    # Write DLOM summary table
    df_output.to_excel(writer, sheet_name='DLOM', index=True)
