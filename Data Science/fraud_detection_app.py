import streamlit as st
import joblib
import json
import pandas as pd
import numpy as np

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Fraud Detection System", layout="centered")

st.title("💳 Fraud Detection System")
st.write("Enter transaction details below.")

# --------------------------------------------------
# LOAD MODEL & FEATURES
# --------------------------------------------------
model = joblib.load("fraud_xgboost_model.pkl")

with open("features.json", "r") as f:
    feature_list = json.load(f)

# --------------------------------------------------
# USER INPUTS (RAW BUSINESS INPUTS ONLY)
# --------------------------------------------------

amount = st.number_input("Transaction Amount", min_value=0.0, value=0.0)

oldbalanceOrg = st.number_input("Old Balance (Before Transaction)", min_value=0.0, value=0.0)

newbalanceOrig = st.number_input("New Balance (After Transaction)", min_value=0.0, value=0.0)

txn_velocity_1h = st.number_input("Transaction Velocity (Last 1 Hour)", min_value=0.0, value=0.0)

txn_velocity_24h = st.number_input("Transaction Velocity (Last 24 Hours)", min_value=0.0, value=0.0)

time_since_last_txn = st.number_input("Time Since Last Transaction (Minutes)", min_value=0.0, value=0.0)

threshold = st.slider("Fraud Detection Threshold", 0.0, 1.0, 0.30)

# --------------------------------------------------
# PREDICTION BUTTON
# --------------------------------------------------

if st.button("Predict Fraud"):

    if oldbalanceOrg <= 0:
        st.warning("Old balance must be greater than zero.")
    else:

        # ------------------------------
        # FEATURE ENGINEERING (MATCH NOTEBOOK)
        # ------------------------------

        log_amount = np.log1p(amount)
        log_oldbalanceOrg = np.log1p(oldbalanceOrg)

        amount_to_balance_ratio = amount / oldbalanceOrg

        balance_drained_pct = (oldbalanceOrg - newbalanceOrig) / oldbalanceOrg
        balance_drained_pct = np.clip(balance_drained_pct, 0, 1)

        near_zero_balance = 1 if newbalanceOrig <= 0.05 * oldbalanceOrg else 0

        # Optional advanced flags (only if used in training)
        drain_gt_50 = 1 if balance_drained_pct > 0.50 else 0
        drain_gt_75 = 1 if balance_drained_pct > 0.75 else 0
        drain_gt_90 = 1 if balance_drained_pct > 0.90 else 0

        rapid_drain_flag = 1 if (balance_drained_pct > 0.75 and txn_velocity_1h > 3) else 0

        drain_velocity_score = balance_drained_pct * txn_velocity_1h

        high_amount_flag = 1 if amount_to_balance_ratio > 0.70 else 0

        # ------------------------------
        # CREATE INPUT DICTIONARY
        # ------------------------------

        input_data = {
            "log_amount": log_amount,
            "log_oldbalanceOrg": log_oldbalanceOrg,
            "amount_to_balance_ratio": amount_to_balance_ratio,
            "balance_drained_pct": balance_drained_pct,
            "near_zero_balance": near_zero_balance,
            "txn_velocity_1h": txn_velocity_1h,
            "txn_velocity_24h": txn_velocity_24h,
            "time_since_last_txn": time_since_last_txn,
            "drain_gt_50": drain_gt_50,
            "drain_gt_75": drain_gt_75,
            "drain_gt_90": drain_gt_90,
            "rapid_drain_flag": rapid_drain_flag,
            "drain_velocity_score": drain_velocity_score,
            "high_amount_flag": high_amount_flag
        }

        # Convert to DataFrame
        df = pd.DataFrame([input_data])

        # Ensure exact feature order
        df = df.reindex(columns=feature_list, fill_value=0)

        # ------------------------------
        # PREDICT
        # ------------------------------

        prob = model.predict_proba(df)[0][1]
        prediction = 1 if prob >= threshold else 0

        st.subheader("Prediction Result")

        if prediction == 1:
            st.error("⚠ Fraud Detected")
        else:
            st.success("✅ Legitimate Transaction")

        st.info(f"Fraud Risk Score: {prob:.4f}")

        # Risk Level
        if prob < 0.30:
            st.success("Risk Level: Low")
        elif prob < 0.70:
            st.warning("Risk Level: Medium")
        else:
            st.error("Risk Level: High")
