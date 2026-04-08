# pip install -U neuralforecast pandas numpy scikit-learn

import pandas as pd
import numpy as np
from sklearn.metrics import mean_absolute_error, mean_squared_error

from neuralforecast import NeuralForecast
from neuralforecast.models import NHITS

# -----------------------
# 1) Load + monthly aggregate
# -----------------------
PATH = "external - External.csv"   # your uploaded file

df = pd.read_csv(PATH)

df["DELIVERY_MON"] = pd.to_datetime(df["DELIVERY_MON"], errors="coerce")
df["QUANTITY_LBS"] = (df["QUANTITY_LBS"].astype(str).str.replace(",", "", regex=False))
df["QUANTITY_LBS"] = pd.to_numeric(df["QUANTITY_LBS"], errors="coerce")

df = df.dropna(subset=["DELIVERY_MON", "QUANTITY_LBS"])

# month-end timestamps (important for consistent monthly frequency)
df["ds"] = df["DELIVERY_MON"].dt.to_period("M").dt.to_timestamp("M")

y = (
    df.groupby("ds", as_index=False)["QUANTITY_LBS"]
      .sum()
      .rename(columns={"QUANTITY_LBS": "y"})
)

# NeuralForecast expects multiple series format; for a single series use one id
y["unique_id"] = "TOTAL"
y = y[["unique_id", "ds", "y"]].sort_values("ds")

# -----------------------
# 2) Train/test split
# -----------------------
train = y[y["ds"] <= "2024-12-31"].copy()
test_2025 = y[(y["ds"] >= "2025-01-31") & (y["ds"] <= "2025-12-31")].copy()

# We'll predict 12 months (2025) for evaluation
H = 12

# -----------------------
# 3) Fit NHITS
# -----------------------
# input_size: how many past months the model uses (e.g., 24 months)
model = NHITS(h=H, input_size=24, max_steps=1000)  # max_steps can be increased

nf = NeuralForecast(models=[model], freq="M")
nf.fit(df=train)

# -----------------------
# 4) Predict 2025 and compare
# -----------------------
pred_2025 = nf.predict()  # returns 12-step forecast after last train date

# pred_2025 columns: unique_id, ds, NHITS
merged_2025 = test_2025.merge(pred_2025, on=["unique_id", "ds"], how="inner")

y_true = merged_2025["y"].to_numpy()
y_hat = merged_2025["NHITS"].to_numpy()

mae = mean_absolute_error(y_true, y_hat)
rmse = mean_squared_error(y_true, y_hat, squared=False)
mape = (np.abs((y_true - y_hat) / np.clip(y_true, 1, None))).mean() * 100

print("NHITS 2025 metrics")
print(f"MAE : {mae:,.0f}")
print(f"RMSE: {rmse:,.0f}")
print(f"MAPE: {mape:,.2f}%")

print("\nNHITS 2025 month-by-month comparison:")
print(merged_2025[["ds", "y", "NHITS"]].rename(columns={"y": "actual", "NHITS": "pred"}))

# -----------------------
# 5) Optional: forecast 2026 (train through 2025, forecast next 12 months)
# -----------------------
train_through_2025 = y[y["ds"] <= "2025-12-31"].copy()

nf2 = NeuralForecast(models=[NHITS(h=12, input_size=24, max_steps=1000)], freq="M")
nf2.fit(df=train_through_2025)

pred_2026 = nf2.predict()
print("\nNHITS 2026 forecast:")
print(pred_2026[["ds", "NHITS"]].rename(columns={"NHITS": "pred_2026"}))