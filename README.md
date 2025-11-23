# 📘 DLOM Valuation Models — Python Implementation

This repository provides a Python-based implementation of commonly used **Discount for Lack of Marketability (DLOM)** valuation models.
The script reads valuation inputs from an Excel template, computes DLOM using multiple theoretical models, and exports results into a formatted Excel output.

---

## 📌 Features

* Implements **four leading DLOM models**:

  * **Chaffe Model**
  * **Finnerty Model**
  * **Ghaidarov Model**
  * **Longstaff Model**
* Reads valuation inputs directly from `Input.xlsx`
* Automatically runs only the models marked as `"Yes"` in the input file
* Generates:

  * A consolidated **DLOM summary table**
  * Individual sheets for each model’s computed values
* Saves output to `Output.xlsx`

---

## 📁 Repository Structure

```
├── script.py           # Main Python script with DLOM model implementations
├── Input.xlsx          # Input file with valuation parameters and model selection flags
├── Output.xlsx         # Auto-generated output file containing DLOM results
└── README.md           # Project documentation
```

---

## 📥 Input File Format (`Input.xlsx`)

The script expects the input file to contain:

| Row | Parameter | Description                                                   |
| --- | --------- | ------------------------------------------------------------- |
| 1   | S         | Equity value / underlying asset value                         |
| 2   | r         | Risk-free rate                                                |
| 3   | T         | Time to liquidity (in years)                                  |
| 4   | sigma     | Volatility                                                    |
| 5   | q         | Dividend yield                                                |
| 6–9 | Yes/No    | Whether to run: Chaffe, Finnerty, Ghaidarov, Longstaff models |

---

## ▶️ How to Run

### **1. Install Dependencies**

```bash
pip install pandas numpy scipy
```

### **2. Place `Input.xlsx` in the same folder as the script**

### **3. Run the script**

```bash
python script.py
```

### **4. Output**

The script generates:

📄 `Output.xlsx` containing:

* **Sheet 1: "Models"** → Put option value & DLOM for each selected model
* **Sheet 2: "DLOM"** → Final summary table with median DLOM

---

## 📊 DLOM Models Implemented

### **1. Chaffe Model**

A Black-Scholes based approach to estimate the cost of illiquidity through a European put option.

### **2. Finnerty Model**

Uses volatility-adjusted spread measures to reflect restricted stock discount.

### **3. Ghaidarov Model**

Applies log-normal distribution parameters to estimate the discount.

### **4. Longstaff Model**

Based on expected payoffs under optimal stopping rules for restricted trading.

---

## 🧪 Example Output (DLOM Summary)

| Model          | DLOM (%) |
| -------------- | -------- |
| Chaffe Model   | 24.51%   |
| Finnerty Model | 18.42%   |
| Median         | 21.47%   |

*(example values)*

---

## 🛠️ Customisation

You can easily modify:

* Model formulas
* Input structure
* Output formatting
* Additional models (simply extend the `models` dictionary)

---

## 🤝 Contributing

Pull requests are welcome!
If you want to add more valuation models or improve Excel formatting, feel free to contribute.

---

## 📜 License

This project is licensed under the MIT License — free to use, modify, and distribute.
