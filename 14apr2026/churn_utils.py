"""
churn_utils.py
==============
Banking Churn Case Study — Helper Module
Jack Henry Associates | Data Science Talk

All heavy lifting lives here. The notebook imports these functions
and focuses purely on results and narrative.

Key JH-relevant features:
  - DD (Direct Deposit) velocity (count/amount, 30/90d)
  - Digital login frequency (mobile + online)
  - Balance patterns, NSFs, product depth, service events
"""

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from datetime import datetime, timedelta
import random

# ─── Color Palette ──────────────────────────────────────────────────────────
NAVY      = "#0D2B55"
BLUE      = "#0078C8"
TEAL      = "#00A878"
ORANGE    = "#F4872B"
LIGHT     = "#E8F3FC"
GRAY      = "#64748B"
RED       = "#E63946"
WHITE     = "#FFFFFF"

PALETTE   = [BLUE, TEAL, ORANGE, NAVY, RED, GRAY]

plt.rcParams.update({
    "figure.facecolor": WHITE,
    "axes.facecolor":   WHITE,
    "axes.edgecolor":   GRAY,
    "axes.labelcolor":  NAVY,
    "text.color":       NAVY,
    "xtick.color":      GRAY,
    "ytick.color":      GRAY,
    "axes.spines.top":  False,
    "axes.spines.right":False,
    "font.family":      "sans-serif",
    "axes.titlesize":   14,
    "axes.titleweight": "bold",
    "axes.titlecolor":  NAVY,
})


# ═══════════════════════════════════════════════════════════════════════════
# 1.  DATA GENERATION
# ═══════════════════════════════════════════════════════════════════════════

def generate_banking_data(n_customers: int = 6000, seed: int = 42) -> pd.DataFrame:
    """
    Generate synthetic banking customer data with realistic JH-relevant features.

    The churn label is deterministic from the features so SHAP values will
    look clean and interpretable in the demo.

    Returns
    -------
    pd.DataFrame  — one row per customer, ordered by signup_date
    """
    rng = np.random.default_rng(seed)

    # ── Dates (simulate 24-month cohort) ────────────────────────────────
    start_date = datetime(2022, 1, 1)
    end_date   = datetime(2023, 9, 30)
    days_range = (end_date - start_date).days
    signup_offsets = rng.integers(0, days_range, size=n_customers)
    signup_dates   = [start_date + timedelta(days=int(d)) for d in signup_offsets]

    # ── Demographics & Relationship ──────────────────────────────────────
    tenure_months          = rng.integers(1, 84, size=n_customers)
    age                    = rng.integers(22, 74, size=n_customers)
    product_count          = rng.integers(1, 7,  size=n_customers)
    months_since_last_prod = rng.integers(0, 36, size=n_customers)
    has_savings            = rng.binomial(1, 0.55, size=n_customers)
    has_loan               = rng.binomial(1, 0.30, size=n_customers)
    has_cd                 = rng.binomial(1, 0.18, size=n_customers)

    # ── Balance Features ─────────────────────────────────────────────────
    avg_daily_balance   = np.clip(rng.lognormal(8.5, 1.2, size=n_customers), 50, 75_000)
    balance_volatility  = rng.beta(1.5, 4, size=n_customers)   # 0–1, higher = more volatile
    min_balance_30d     = avg_daily_balance * rng.uniform(0.1, 1.0, size=n_customers)
    overdraft_count_12m = rng.negative_binomial(1, 0.7, size=n_customers)

    # ── Direct Deposit (DD) Velocity — KEY JH FEATURE ────────────────────
    # Primary bank customers (60%) have high DD; others low/zero
    is_primary_bank = rng.binomial(1, 0.60, size=n_customers)
    dd_count_30d    = np.where(
        is_primary_bank,
        rng.integers(2, 6, size=n_customers),
        rng.integers(0, 2, size=n_customers)
    )
    dd_count_90d    = dd_count_30d * rng.integers(2, 4, size=n_customers)
    dd_amount_30d   = np.where(
        is_primary_bank,
        rng.lognormal(8.4, 0.5, size=n_customers),   # ~$4,500 avg
        rng.lognormal(6.5, 1.0, size=n_customers)    # ~$665 avg
    )
    dd_amount_90d   = dd_amount_30d * rng.uniform(2.5, 3.5, size=n_customers)
    days_since_last_dd = np.where(
        is_primary_bank,
        rng.integers(1, 20,  size=n_customers),
        rng.integers(5, 120, size=n_customers)
    )

    # ── Digital Engagement — KEY JH FEATURE ─────────────────────────────
    is_digital_active  = rng.binomial(1, 0.65, size=n_customers)
    digital_login_30d  = np.where(
        is_digital_active,
        rng.integers(8, 40, size=n_customers),
        rng.integers(0, 5,  size=n_customers)
    )
    mobile_login_30d   = (digital_login_30d * rng.uniform(0.5, 0.9, size=n_customers)).astype(int)
    days_since_last_login = np.where(
        is_digital_active,
        rng.integers(0, 7,  size=n_customers),
        rng.integers(7, 90, size=n_customers)
    )
    mobile_deposit_cnt = np.where(
        is_digital_active,
        rng.integers(1, 8, size=n_customers),
        rng.integers(0, 2, size=n_customers)
    )

    # ── Transaction Behavior ─────────────────────────────────────────────
    debit_txn_30d  = rng.integers(2,  60, size=n_customers)
    atm_txn_30d    = rng.integers(0,  12, size=n_customers)
    bill_pay_30d   = rng.integers(0,  15, size=n_customers)
    ach_cnt_30d    = rng.integers(0,  20, size=n_customers)

    # ── Service Events ────────────────────────────────────────────────────
    nsf_count_12m          = rng.negative_binomial(1, 0.6, size=n_customers)
    fee_waiver_req_12m     = rng.integers(0, 5, size=n_customers)
    cust_service_calls_6m  = rng.integers(0, 8, size=n_customers)
    branch_visits_30d      = rng.integers(0, 6, size=n_customers)

    # ── Build DataFrame ───────────────────────────────────────────────────
    df = pd.DataFrame({
        "customer_id":             [f"CUS{str(i).zfill(6)}" for i in range(n_customers)],
        "signup_date":             signup_dates,
        "tenure_months":           tenure_months,
        "age":                     age,
        "product_count":           product_count,
        "months_since_last_prod":  months_since_last_prod,
        "has_savings":             has_savings,
        "has_loan":                has_loan,
        "has_cd":                  has_cd,
        "avg_daily_balance":       avg_daily_balance.round(2),
        "balance_volatility":      balance_volatility.round(4),
        "min_balance_30d":         min_balance_30d.round(2),
        "overdraft_count_12m":     overdraft_count_12m,
        # Direct Deposit
        "dd_count_30d":            dd_count_30d,
        "dd_count_90d":            dd_count_90d,
        "dd_amount_30d":           dd_amount_30d.round(2),
        "dd_amount_90d":           dd_amount_90d.round(2),
        "days_since_last_dd":      days_since_last_dd,
        "is_primary_bank":         is_primary_bank,
        # Digital
        "digital_login_30d":       digital_login_30d,
        "mobile_login_30d":        mobile_login_30d,
        "days_since_last_login":   days_since_last_login,
        "mobile_deposit_cnt_30d":  mobile_deposit_cnt,
        # Transactions
        "debit_txn_30d":           debit_txn_30d,
        "atm_txn_30d":             atm_txn_30d,
        "bill_pay_30d":            bill_pay_30d,
        "ach_cnt_30d":             ach_cnt_30d,
        # Service
        "nsf_count_12m":           nsf_count_12m,
        "fee_waiver_req_12m":      fee_waiver_req_12m,
        "cust_service_calls_6m":   cust_service_calls_6m,
        "branch_visits_30d":       branch_visits_30d,
    })

    # ── Label: Churn in next 90 days ──────────────────────────────────────
    df = _label_churn(df, rng)
    df = df.sort_values("signup_date").reset_index(drop=True)
    return df


def _label_churn(df: pd.DataFrame, rng) -> pd.DataFrame:
    """
    Churn probability is a deterministic logistic function of key features
    so SHAP plots tell a coherent story.
    """
    score = (
        -2.8                                               # base intercept
        - 1.4 * df["is_primary_bank"]                     # primary DD = very sticky
        - 0.08 * df["dd_count_30d"]                       # more DDs = less churn
        - 0.0002 * df["dd_amount_30d"]                    # higher DD = less churn
        + 0.03 * df["days_since_last_dd"]                 # long gap = risk
        - 0.04 * df["digital_login_30d"]                  # engaged = less churn
        + 0.02 * df["days_since_last_login"]              # dormant = risk
        - 0.25 * df["product_count"]                      # more products = stickier
        + 0.08 * df["months_since_last_prod"]             # stagnant relationship
        - 0.0001 * df["avg_daily_balance"]                # higher balance = less churn
        + 0.6  * df["balance_volatility"]                 # volatile = risk
        + 0.18 * df["nsf_count_12m"]                      # NSFs = friction
        + 0.35 * df["fee_waiver_req_12m"]                 # complaints = risk
        + 0.12 * df["cust_service_calls_6m"]              # contacts = friction signal
        - 0.02 * df["tenure_months"]                      # longer tenure = stickier
        + rng.normal(0, 0.3, size=len(df))                # noise
    )
    prob   = 1 / (1 + np.exp(-score))
    labels = rng.binomial(1, prob)
    df["churn_90d"]     = labels
    df["churn_prob_true"] = prob.round(4)   # for calibration reference only
    return df


# ═══════════════════════════════════════════════════════════════════════════
# 2.  LABELING EXPLAINER
# ═══════════════════════════════════════════════════════════════════════════

def explain_labeling():
    """Print labeling methodology and common pitfalls."""
    print("""
╔══════════════════════════════════════════════════════════════════════╗
║                    LABELING: CHURN DEFINITION                       ║
╠══════════════════════════════════════════════════════════════════════╣
║  Definition:  Account closure or 90-day complete dormancy           ║
║               (no transactions AND no digital logins)               ║
╠══════════════════════════════════════════════════════════════════════╣
║  Look-forward window:  90 days from observation date                ║
║  Observation date:     The snapshot date of all features            ║
╠══════════════════════════════════════════════════════════════════════╣
║  ⚠  COMMON PITFALLS                                                  ║
║     1. Label Leakage — features computed AFTER churn event          ║
║        e.g. "balance at close" — not knowable at prediction time    ║
║     2. Look-back contamination — using data from future window      ║
║        in features (overlapping periods)                            ║
║     3. Survivorship bias — excluding customers who already left     ║
║     4. Class imbalance — typical churn: 8–15% in retail banking     ║
╚══════════════════════════════════════════════════════════════════════╝
""")


# ═══════════════════════════════════════════════════════════════════════════
# 3.  EXPLORATORY DATA ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════

def plot_eda(df: pd.DataFrame):
    """
    4-panel EDA dashboard:
      - Churn rate by DD active status
      - Digital login distribution by churn label
      - Churn rate by product count
      - Balance distribution (log scale)
    """
    churn_rate = df["churn_90d"].mean()
    fig, axes = plt.subplots(2, 2, figsize=(13, 9))
    fig.suptitle("Banking Churn — Exploratory Data Analysis", fontsize=16,
                 fontweight="bold", color=NAVY, y=1.01)

    # Panel 1: Churn rate by DD status
    ax = axes[0, 0]
    dd_churn = df.groupby("is_primary_bank")["churn_90d"].mean().reset_index()
    dd_churn["label"] = dd_churn["is_primary_bank"].map({0: "No Primary DD", 1: "Primary DD Active"})
    bars = ax.bar(dd_churn["label"], dd_churn["churn_90d"] * 100,
                  color=[ORANGE, TEAL], width=0.5, edgecolor="none")
    for bar, val in zip(bars, dd_churn["churn_90d"]):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
                f"{val*100:.1f}%", ha="center", va="bottom", fontsize=12,
                fontweight="bold", color=NAVY)
    ax.axhline(churn_rate * 100, color=GRAY, linestyle="--", linewidth=1.2, label=f"Overall: {churn_rate*100:.1f}%")
    ax.set_ylabel("Churn Rate (%)")
    ax.set_title("Churn Rate by Direct Deposit Status")
    ax.legend(fontsize=9)
    ax.set_ylim(0, dd_churn["churn_90d"].max() * 100 * 1.25)

    # Panel 2: Digital logins distribution
    ax = axes[0, 1]
    churned     = df[df["churn_90d"] == 1]["digital_login_30d"]
    not_churned = df[df["churn_90d"] == 0]["digital_login_30d"]
    bins = range(0, 42, 3)
    ax.hist(not_churned, bins=bins, alpha=0.7, color=TEAL,   label="Retained",  density=True)
    ax.hist(churned,     bins=bins, alpha=0.7, color=ORANGE, label="Churned",   density=True)
    ax.set_xlabel("Digital Logins (last 30 days)")
    ax.set_ylabel("Density")
    ax.set_title("Digital Engagement vs. Churn")
    ax.legend()

    # Panel 3: Churn rate by product count
    ax = axes[1, 0]
    prod_churn = df.groupby("product_count")["churn_90d"].mean() * 100
    ax.bar(prod_churn.index, prod_churn.values,
           color=[BLUE if v < churn_rate * 100 else ORANGE for v in prod_churn.values],
           edgecolor="none")
    ax.axhline(churn_rate * 100, color=GRAY, linestyle="--", linewidth=1.2,
               label=f"Overall avg: {churn_rate*100:.1f}%")
    ax.set_xlabel("Number of Products")
    ax.set_ylabel("Churn Rate (%)")
    ax.set_title("Churn Rate by Product Count")
    ax.legend(fontsize=9)

    # Panel 4: Average daily balance (log scale)
    ax = axes[1, 1]
    ax.hist(df[df["churn_90d"] == 0]["avg_daily_balance"].clip(50, 50000),
            bins=40, alpha=0.7, color=TEAL,   label="Retained",  density=True, log=True)
    ax.hist(df[df["churn_90d"] == 1]["avg_daily_balance"].clip(50, 50000),
            bins=40, alpha=0.7, color=ORANGE, label="Churned",   density=True, log=True)
    ax.set_xlabel("Average Daily Balance ($)")
    ax.set_ylabel("Log Density")
    ax.set_title("Balance Distribution vs. Churn")
    ax.legend()

    plt.tight_layout()
    plt.show()
    print(f"\n📊 Dataset: {len(df):,} customers | Churn rate: {churn_rate*100:.1f}% "
          f"({df['churn_90d'].sum():,} churners)\n")


def plot_correlation_heatmap(df: pd.DataFrame):
    """Correlation heatmap of key numeric features against churn."""
    key_features = [
        "dd_count_30d", "dd_amount_30d", "days_since_last_dd",
        "digital_login_30d", "days_since_last_login",
        "product_count", "avg_daily_balance", "balance_volatility",
        "nsf_count_12m", "fee_waiver_req_12m", "tenure_months",
        "debit_txn_30d", "bill_pay_30d", "churn_90d"
    ]
    corr = df[key_features].corr()["churn_90d"].drop("churn_90d").sort_values()

    fig, ax = plt.subplots(figsize=(8, 6))
    colors = [ORANGE if v > 0 else TEAL for v in corr.values]
    ax.barh(corr.index, corr.values, color=colors, edgecolor="none")
    ax.axvline(0, color=GRAY, linewidth=0.8)
    ax.set_xlabel("Correlation with Churn (next 90 days)")
    ax.set_title("Feature Correlation with Churn Label", pad=15)
    plt.tight_layout()
    plt.show()


# ═══════════════════════════════════════════════════════════════════════════
# 4.  TEMPORAL SPLIT (by date, not random!)
# ═══════════════════════════════════════════════════════════════════════════

def temporal_split(df: pd.DataFrame,
                   train_end: str = "2023-03-31",
                   val_end:   str = "2023-06-30"):
    """
    Split customer cohorts by signup date to prevent data leakage.

    Train  :  signup_date <= train_end
    Val    :  train_end < signup_date <= val_end
    Test   :  signup_date > val_end

    Returns
    -------
    train_df, val_df, test_df
    """
    train_end_dt = pd.Timestamp(train_end)
    val_end_dt   = pd.Timestamp(val_end)
    df["signup_date"] = pd.to_datetime(df["signup_date"])

    train = df[df["signup_date"] <= train_end_dt].copy()
    val   = df[(df["signup_date"] > train_end_dt) & (df["signup_date"] <= val_end_dt)].copy()
    test  = df[df["signup_date"] > val_end_dt].copy()

    print("📅 Temporal Split Summary")
    print("─" * 45)
    print(f"  Train  : {len(train):>5,} customers  "
          f"(churn: {train['churn_90d'].mean()*100:.1f}%)")
    print(f"  Val    : {len(val):>5,} customers  "
          f"(churn: {val['churn_90d'].mean()*100:.1f}%)")
    print(f"  Test   : {len(test):>5,} customers  "
          f"(churn: {test['churn_90d'].mean()*100:.1f}%)")
    print("─" * 45)
    print("  ✅  Split by date → no future leakage into training features\n")
    return train, val, test


def plot_temporal_split(df: pd.DataFrame,
                        train_end: str = "2023-03-31",
                        val_end:   str = "2023-06-30"):
    """Visualize the temporal split timeline."""
    df = df.copy()
    df["signup_date"] = pd.to_datetime(df["signup_date"])
    df["month"] = df["signup_date"].dt.to_period("M")
    monthly = df.groupby("month").size().reset_index(name="count")
    monthly["month_dt"] = monthly["month"].dt.to_timestamp()

    train_end_dt = pd.Timestamp(train_end)
    val_end_dt   = pd.Timestamp(val_end)

    fig, ax = plt.subplots(figsize=(11, 3.5))
    for _, row in monthly.iterrows():
        if row["month_dt"] <= train_end_dt:
            color = BLUE
        elif row["month_dt"] <= val_end_dt:
            color = TEAL
        else:
            color = ORANGE
        ax.bar(row["month_dt"], row["count"], width=25, color=color, edgecolor="none")

    ax.axvline(train_end_dt, color=NAVY, linestyle="--", linewidth=1.5)
    ax.axvline(val_end_dt,   color=GRAY, linestyle="--", linewidth=1.5)
    ax.text(train_end_dt, ax.get_ylim()[1] * 0.92, "  Train / Val split", color=NAVY, fontsize=9)
    ax.text(val_end_dt,   ax.get_ylim()[1] * 0.92, "  Val / Test split",  color=GRAY, fontsize=9)

    legend_patches = [
        mpatches.Patch(color=BLUE,   label="Train"),
        mpatches.Patch(color=TEAL,   label="Validation"),
        mpatches.Patch(color=ORANGE, label="Test (holdout)"),
    ]
    ax.legend(handles=legend_patches, loc="upper left")
    ax.set_ylabel("Customers Signing Up")
    ax.set_title("Temporal Partitioning — Customers by Signup Date")
    plt.tight_layout()
    plt.show()


# ═══════════════════════════════════════════════════════════════════════════
# 5.  FEATURE ENGINEERING
# ═══════════════════════════════════════════════════════════════════════════

FEATURE_COLS = [
    # DD velocity
    "dd_count_30d", "dd_count_90d", "dd_amount_30d", "dd_amount_90d",
    "days_since_last_dd",
    # Digital
    "digital_login_30d", "mobile_login_30d", "days_since_last_login",
    "mobile_deposit_cnt_30d",
    # Balance
    "avg_daily_balance", "balance_volatility", "min_balance_30d",
    "overdraft_count_12m",
    # Products
    "product_count", "months_since_last_prod", "has_savings", "has_loan", "has_cd",
    # Transactions
    "debit_txn_30d", "atm_txn_30d", "bill_pay_30d", "ach_cnt_30d",
    # Service
    "nsf_count_12m", "fee_waiver_req_12m", "cust_service_calls_6m",
    "branch_visits_30d",
    # Relationship
    "tenure_months", "age",
    # Engineered
    "dd_velocity_ratio",
    "digital_engagement_score",
    "balance_stress_index",
    "relationship_depth_score",
    "friction_score",
]

TARGET_COL = "churn_90d"


def engineer_features(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add derived features that capture relationship health signals.

    Key engineered features:
      dd_velocity_ratio         — 30d vs 90d DD count ratio (trend)
      digital_engagement_score  — composite login + recency score
      balance_stress_index      — volatility + overdrafts + NSFs
      relationship_depth_score  — product diversity + tenure
      friction_score            — service complaints + fee events
    """
    df = df.copy()

    # DD velocity ratio (recent vs baseline — declining = risk signal)
    dd_90d_monthly = df["dd_count_90d"] / 3 + 1e-6
    df["dd_velocity_ratio"] = df["dd_count_30d"] / dd_90d_monthly

    # Digital engagement score (0–100)
    login_norm    = np.clip(df["digital_login_30d"] / 40, 0, 1)
    recency_norm  = np.clip(1 - df["days_since_last_login"] / 90, 0, 1)
    df["digital_engagement_score"] = ((login_norm * 0.6 + recency_norm * 0.4) * 100).round(1)

    # Balance stress index
    overdraft_norm = np.clip(df["overdraft_count_12m"] / 10, 0, 1)
    nsf_norm       = np.clip(df["nsf_count_12m"]       / 10, 0, 1)
    df["balance_stress_index"] = (
        df["balance_volatility"] * 0.4 + overdraft_norm * 0.3 + nsf_norm * 0.3
    ).round(4)

    # Relationship depth score (0–1)
    tenure_norm = np.clip(df["tenure_months"] / 60, 0, 1)
    prod_norm   = np.clip(df["product_count"]  / 6, 0, 1)
    df["relationship_depth_score"] = (tenure_norm * 0.5 + prod_norm * 0.5).round(4)

    # Friction score (customer dissatisfaction signal)
    fee_norm   = np.clip(df["fee_waiver_req_12m"]     / 5, 0, 1)
    calls_norm = np.clip(df["cust_service_calls_6m"]  / 8, 0, 1)
    df["friction_score"] = (fee_norm * 0.5 + calls_norm * 0.5).round(4)

    print("🔧 Feature Engineering Complete")
    print(f"   New features: dd_velocity_ratio, digital_engagement_score,")
    print(f"                 balance_stress_index, relationship_depth_score, friction_score")
    print(f"   Total model features: {len(FEATURE_COLS)}\n")
    return df


# ═══════════════════════════════════════════════════════════════════════════
# 6.  MODEL TRAINING (AutoGluon → fallback to CatBoost)
# ═══════════════════════════════════════════════════════════════════════════

def train_model(train_df: pd.DataFrame,
                val_df:   pd.DataFrame,
                time_limit: int = 120,
                presets: str = "medium_quality"):
    """
    Train with AutoGluon TabularPredictor.
    Falls back to CatBoostClassifier if AutoGluon is unavailable.

    Parameters
    ----------
    train_df    : training set (output of engineer_features)
    val_df      : validation set
    time_limit  : AutoGluon time budget in seconds (default 120s for demo)
    presets     : AutoGluon presets ('fast_ai', 'medium_quality', 'best_quality')

    Returns
    -------
    predictor   : fitted AutoGluon predictor (or CatBoost model)
    model_type  : "autogluon" | "catboost"
    """
    train_data = train_df[FEATURE_COLS + [TARGET_COL]].copy()
    val_data   = val_df[FEATURE_COLS + [TARGET_COL]].copy()

    try:
        from autogluon.tabular import TabularPredictor
        print("🚀 Training with AutoGluon TabularPredictor")
        print(f"   Time limit: {time_limit}s | Presets: {presets}\n")
        predictor = TabularPredictor(
            label       = TARGET_COL,
            eval_metric = "roc_auc",
            verbosity   = 2,
        ).fit(
            train_data  = train_data,
            tuning_data = val_data,
            time_limit  = time_limit,
            presets     = presets,
        )
        print("\n✅ AutoGluon training complete\n")
        return predictor, "autogluon"

    except ImportError:
        print("⚠️  AutoGluon not installed — falling back to CatBoostClassifier")
        print("   Install: pip install autogluon.tabular\n")
        _train_catboost(train_data, val_data)


def _train_catboost(train_data, val_data):
    from catboost import CatBoostClassifier, Pool
    X_train = train_data[FEATURE_COLS]
    y_train = train_data[TARGET_COL]
    X_val   = val_data[FEATURE_COLS]
    y_val   = val_data[TARGET_COL]

    model = CatBoostClassifier(
        iterations    = 800,
        learning_rate = 0.05,
        depth         = 6,
        eval_metric   = "AUC",
        random_seed   = 42,
        verbose       = 100,
    )
    model.fit(X_train, y_train, eval_set=(X_val, y_val), early_stopping_rounds=50)
    print("\n✅ CatBoost training complete\n")
    return model, "catboost"


# ═══════════════════════════════════════════════════════════════════════════
# 7.  EVALUATION
# ═══════════════════════════════════════════════════════════════════════════

def evaluate_model(predictor, test_df: pd.DataFrame, model_type: str = "autogluon"):
    """
    Comprehensive evaluation:
      - AUC-ROC curve
      - Precision-Recall curve
      - Confusion matrix
      - Calibration plot
    """
    from sklearn.metrics import (roc_auc_score, roc_curve, average_precision_score,
                                 precision_recall_curve, confusion_matrix)

    X_test = test_df[FEATURE_COLS]
    y_test = test_df[TARGET_COL].values

    if model_type == "autogluon":
        y_pred_proba = predictor.predict_proba(X_test)[1].values
        y_pred       = (y_pred_proba >= 0.5).astype(int)
    else:
        y_pred_proba = predictor.predict_proba(X_test)[:, 1]
        y_pred       = (y_pred_proba >= 0.5).astype(int)

    auc  = roc_auc_score(y_test, y_pred_proba)
    aps  = average_precision_score(y_test, y_pred_proba)
    cm   = confusion_matrix(y_test, y_pred)

    fpr, tpr, _ = roc_curve(y_test, y_pred_proba)
    prec, rec, _  = precision_recall_curve(y_test, y_pred_proba)

    fig, axes = plt.subplots(1, 3, figsize=(15, 4.5))
    fig.suptitle(f"Model Evaluation — Test Set Performance", fontsize=14,
                 fontweight="bold", color=NAVY)

    # ROC Curve
    ax = axes[0]
    ax.plot(fpr, tpr, color=BLUE, linewidth=2.5, label=f"AUC = {auc:.3f}")
    ax.plot([0, 1], [0, 1], color=GRAY, linestyle="--", linewidth=1)
    ax.fill_between(fpr, tpr, alpha=0.08, color=BLUE)
    ax.set_xlabel("False Positive Rate")
    ax.set_ylabel("True Positive Rate")
    ax.set_title("ROC Curve")
    ax.legend(loc="lower right")

    # Precision-Recall Curve
    ax = axes[1]
    ax.plot(rec, prec, color=TEAL, linewidth=2.5, label=f"AP = {aps:.3f}")
    ax.fill_between(rec, prec, alpha=0.08, color=TEAL)
    ax.set_xlabel("Recall")
    ax.set_ylabel("Precision")
    ax.set_title("Precision-Recall Curve")
    ax.legend(loc="upper right")

    # Confusion Matrix
    ax = axes[2]
    labels = [["TN", "FP"], ["FN", "TP"]]
    im = ax.imshow(cm, cmap="Blues")
    for i in range(2):
        for j in range(2):
            ax.text(j, i, f"{labels[i][j]}\n{cm[i, j]:,}",
                    ha="center", va="center",
                    color="white" if cm[i, j] > cm.max() / 2 else NAVY,
                    fontsize=13, fontweight="bold")
    ax.set_xticks([0, 1]); ax.set_xticklabels(["Pred: Retain", "Pred: Churn"])
    ax.set_yticks([0, 1]); ax.set_yticklabels(["Actual: Retain", "Actual: Churn"])
    ax.set_title("Confusion Matrix")

    plt.tight_layout()
    plt.show()

    print(f"\n📈 Test Set Results")
    print(f"   AUC-ROC            : {auc:.4f}")
    print(f"   Average Precision  : {aps:.4f}")
    print(f"   Class balance      : {y_test.mean()*100:.1f}% churners in test set\n")
    return y_pred_proba


# ═══════════════════════════════════════════════════════════════════════════
# 7b. BUSINESS METRICS EVALUATION
# ═══════════════════════════════════════════════════════════════════════════

def plot_business_metrics(y_true, y_scores,
                          avg_customer_value: float = 1_200,
                          retention_offer_cost: float = 75):
    """
    Translate ML scores into business impact.

    Shows 3 panels:
      1. Lift Curve — how much better than random targeting?
      2. Precision @ Top-K% — if we target the top K% riskiest customers, how accurate?
      3. Cost-Benefit @ Threshold — net revenue saved minus offer cost

    Parameters
    ----------
    avg_customer_value   : annual revenue per retained customer ($)
    retention_offer_cost : cost of a retention intervention ($)
    """
    import pandas as pd
    y_true   = np.array(y_true)
    y_scores = np.array(y_scores)
    n        = len(y_true)
    base_rate = y_true.mean()

    # Sort by predicted risk descending
    order     = np.argsort(-y_scores)
    y_sorted  = y_true[order]

    # ── Lift Curve ────────────────────────────────────────────────────────
    cumulative_found = np.cumsum(y_sorted)
    pct_population   = np.arange(1, n + 1) / n
    pct_churners     = cumulative_found / y_true.sum()
    lift             = pct_churners / pct_population

    # ── Precision @ Top-K ─────────────────────────────────────────────────
    k_values = np.array([0.05, 0.10, 0.15, 0.20, 0.25, 0.30, 0.40, 0.50])
    prec_at_k = []
    for k in k_values:
        top_k  = int(k * n)
        prec_at_k.append(y_sorted[:top_k].mean())

    # ── Cost-Benefit at Different Thresholds ─────────────────────────────
    thresholds = np.linspace(0.1, 0.9, 50)
    net_values = []
    for t in thresholds:
        predicted_churners = (y_scores >= t).sum()
        true_pos   = ((y_scores >= t) & (y_true == 1)).sum()
        false_pos  = ((y_scores >= t) & (y_true == 0)).sum()
        revenue_saved = true_pos  * avg_customer_value
        offer_cost    = predicted_churners * retention_offer_cost
        net_values.append(revenue_saved - offer_cost)

    fig, axes = plt.subplots(1, 3, figsize=(15, 4.5))
    fig.suptitle("Business Impact Evaluation — Beyond a Single Metric",
                 fontsize=14, fontweight="bold", color=NAVY)

    # Panel 1: Lift Curve
    ax = axes[0]
    ax.plot(pct_population * 100, lift, color=BLUE, linewidth=2.5)
    ax.axhline(1.0, color=GRAY, linestyle="--", linewidth=1.2, label="Random baseline (lift=1)")
    ax.fill_between(pct_population * 100, lift, 1, where=lift > 1,
                    alpha=0.12, color=BLUE, label="Model lift above random")
    ax.set_xlabel("% of Customers Contacted (ranked by risk)")
    ax.set_ylabel("Lift")
    ax.set_title("Lift Curve")
    ax.legend(fontsize=9)
    ax.set_xlim(0, 100)

    # Panel 2: Precision @ Top-K
    ax = axes[1]
    bars = ax.bar(k_values * 100, np.array(prec_at_k) * 100,
                  color=TEAL, width=3.5, edgecolor="none")
    ax.axhline(base_rate * 100, color=GRAY, linestyle="--", linewidth=1.2,
               label=f"Base rate ({base_rate*100:.1f}%)")
    for bar, v in zip(bars, prec_at_k):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
                f"{v*100:.0f}%", ha="center", va="bottom", fontsize=8, color=NAVY)
    ax.set_xlabel("Top K% Customers Targeted")
    ax.set_ylabel("Precision (% actual churners)")
    ax.set_title("Precision @ Top-K%\n(if we only call the riskiest K%)")
    ax.legend(fontsize=9)

    # Panel 3: Net Revenue Saved
    ax = axes[2]
    best_thresh = thresholds[np.argmax(net_values)]
    best_net    = max(net_values)
    ax.plot(thresholds, [v / 1000 for v in net_values], color=ORANGE, linewidth=2.5)
    ax.axvline(best_thresh, color=NAVY, linestyle="--", linewidth=1.5,
               label=f"Optimal threshold: {best_thresh:.2f}")
    ax.scatter([best_thresh], [best_net / 1000], color=RED, s=80, zorder=5,
               label=f"Max: ${best_net/1000:,.0f}K saved")
    ax.set_xlabel("Decision Threshold")
    ax.set_ylabel("Net Revenue Saved ($K)")
    ax.set_title(f"Cost-Benefit Analysis\n(Offer cost: ${retention_offer_cost} | Avg value: ${avg_customer_value:,})")
    ax.legend(fontsize=9)

    plt.tight_layout()
    plt.show()

    # Print summary table
    print(f"\n💼 Business Metrics Summary")
    print(f"   Assumption: avg customer value = ${avg_customer_value:,}/yr | retention offer = ${retention_offer_cost}")
    print(f"\n   {'Target Top %':<16} {'Precision':<14} {'Churners Found':<18} {'Revenue Saved':<18} {'Offer Cost'}")
    print("   " + "─" * 76)
    for k, p in zip(k_values[:5], prec_at_k[:5]):
        top_k_n       = int(k * n)
        true_pos      = int(p * top_k_n)
        rev_saved     = true_pos * avg_customer_value
        cost          = top_k_n  * retention_offer_cost
        print(f"   Top {k*100:.0f}%{'':<11} {p*100:.1f}%{'':<9} {true_pos:<18,} ${rev_saved:<17,.0f} ${cost:,}")
    print(f"\n   Optimal threshold: {best_thresh:.2f}  →  Estimated net savings: ${best_net:,.0f}\n")


# ═══════════════════════════════════════════════════════════════════════════
# 8.  GLOBAL EXPLAINABILITY (SHAP Beeswarm)
# ═══════════════════════════════════════════════════════════════════════════

def compute_shap_values(predictor, df: pd.DataFrame, model_type: str = "autogluon",
                        sample_size: int = 500):
    """
    Compute SHAP values using TreeSHAP.

    For AutoGluon: extracts the best tree-based model (CatBoost/XGBoost/LightGBM).
    For CatBoost:  uses shap.TreeExplainer directly.

    Returns
    -------
    shap_values   : np.ndarray  (n_samples × n_features)
    X_sample      : pd.DataFrame
    explainer     : shap.TreeExplainer
    """
    import shap
    rng   = np.random.default_rng(99)
    idx   = rng.choice(len(df), size=min(sample_size, len(df)), replace=False)
    X_sample = df[FEATURE_COLS].iloc[idx].reset_index(drop=True)

    if model_type == "autogluon":
        # Get best model name — works across all AutoGluon versions
        try:
            best_model_name = predictor.model_best          # property (>= 0.6)
        except AttributeError:
            best_model_name = predictor.leaderboard(silent=True).iloc[0]["model"]
        print(f"🌳 Extracting SHAP from AutoGluon best model: {best_model_name}")
        try:
            # Access the native model object for TreeExplainer
            trainer = getattr(predictor, "_trainer", None)
            if trainer is None:
                trainer = predictor._learner.load_trainer()
            model_obj    = trainer.load_model(best_model_name)
            native_model = model_obj.model
            explainer    = shap.TreeExplainer(native_model)
        except Exception as e:
            # Fallback: KernelExplainer — model-agnostic, works on any predictor
            print(f"   → TreeExplainer unavailable ({e}), using KernelExplainer (slower)")
            predict_fn = lambda x: predictor.predict_proba(
                pd.DataFrame(x, columns=FEATURE_COLS)
            )[1].values
            bg        = shap.sample(X_sample, 50)
            explainer = shap.KernelExplainer(predict_fn, bg)
    else:
        explainer = shap.TreeExplainer(predictor)

    shap_values = explainer.shap_values(X_sample)
    # Handle list output (some models return [class0, class1])
    if isinstance(shap_values, list):
        shap_values = shap_values[1]

    print(f"   SHAP values computed for {len(X_sample)} customers\n")
    return shap_values, X_sample, explainer


def plot_shap_global(shap_values: np.ndarray, X_sample: pd.DataFrame,
                     max_display: int = 15):
    """
    Beeswarm summary plot — global feature importance.
    Shows direction AND magnitude of each feature's effect on churn probability.
    """
    import shap
    print("📊 Global Explainability — SHAP Beeswarm Plot")
    print("   Each dot = one customer | Color = feature value")
    print("   Right of center = increases churn probability\n")

    plt.figure(figsize=(10, 7))
    shap.summary_plot(
        shap_values, X_sample,
        max_display    = max_display,
        show           = False,
        color_bar_label= "Feature Value (low → high)",
        plot_size      = None,
    )
    plt.title("Global Explainability: What Drives Churn?",
              fontsize=14, fontweight="bold", color=NAVY, pad=15)
    plt.tight_layout()
    plt.show()


def plot_shap_bar(shap_values: np.ndarray, X_sample: pd.DataFrame,
                  max_display: int = 15):
    """Mean absolute SHAP bar chart — overall feature importance ranking."""
    import shap
    plt.figure(figsize=(9, 6))
    shap.summary_plot(
        shap_values, X_sample,
        plot_type   = "bar",
        max_display = max_display,
        show        = False,
        color       = BLUE,
    )
    plt.title("Feature Importance (Mean |SHAP value|)",
              fontsize=14, fontweight="bold", color=NAVY, pad=15)
    plt.tight_layout()
    plt.show()


# ═══════════════════════════════════════════════════════════════════════════
# 9.  LOCAL EXPLAINABILITY (SHAP Waterfall — Mystery Customer)
# ═══════════════════════════════════════════════════════════════════════════

MYSTERY_CUSTOMER_ID = 42   # fixed seed customer for the presentation narrative


def get_mystery_customer(df: pd.DataFrame) -> pd.Series:
    """Return the 'mystery customer' used as the opening hook."""
    return df.iloc[MYSTERY_CUSTOMER_ID]


def print_mystery_customer(customer: pd.Series):
    """Pretty-print the mystery customer profile (no churn label shown)."""
    print("═" * 55)
    print("  🔍  MYSTERY CUSTOMER — Will they churn?")
    print("═" * 55)
    print(f"  Tenure          : {customer.tenure_months} months")
    print(f"  Products        : {customer.product_count}")
    print(f"  Avg Daily Bal   : ${customer.avg_daily_balance:,.0f}")
    print(f"  DD Count (30d)  : {customer.dd_count_30d}")
    print(f"  DD Amount (30d) : ${customer.dd_amount_30d:,.0f}")
    print(f"  Digital Logins  : {customer.digital_login_30d} (last 30 days)")
    print(f"  Days Since Login: {customer.days_since_last_login}")
    print(f"  NSF Events (12m): {customer.nsf_count_12m}")
    print(f"  Fee Waiver Req  : {customer.fee_waiver_req_12m}")
    print("═" * 55)
    print("  (Scroll down after the model runs to see the answer)")
    print("═" * 55 + "\n")


def plot_shap_local(shap_values: np.ndarray, X_sample: pd.DataFrame,
                    explainer, predictor, model_type: str = "autogluon",
                    customer_idx: int = 0):
    """
    Waterfall plot for a single customer — local explainability.
    Answers the 'mystery customer' question from the opening hook.
    """
    import shap

    customer     = X_sample.iloc[[customer_idx]]
    cust_shap    = shap_values[customer_idx]
    base_value   = explainer.expected_value
    if isinstance(base_value, (list, np.ndarray)):
        base_value = base_value[1]

    final_score  = base_value + cust_shap.sum()
    churn_prob   = 1 / (1 + np.exp(-final_score)) if abs(final_score) < 20 else (1 if final_score > 0 else 0)

    print(f"\n🎯 Local Explanation — Mystery Customer")
    print(f"   Predicted churn probability: {churn_prob*100:.1f}%")
    verdict = "⚠️  HIGH CHURN RISK" if churn_prob > 0.5 else "✅  LOW CHURN RISK"
    print(f"   Verdict: {verdict}\n")

    # Waterfall plot
    shap_exp = shap.Explanation(
        values     = cust_shap,
        base_values= base_value,
        data       = customer.values[0],
        feature_names = FEATURE_COLS,
    )
    plt.figure(figsize=(10, 7))
    shap.waterfall_plot(shap_exp, max_display=12, show=False)
    plt.title(f"Local Explainability — Why does this customer have a {churn_prob*100:.0f}% churn risk?",
              fontsize=13, fontweight="bold", color=NAVY, pad=15)
    plt.tight_layout()
    plt.show()


# ═══════════════════════════════════════════════════════════════════════════
# 10.  DALEX EXPLAINABILITY
# ═══════════════════════════════════════════════════════════════════════════

def build_dalex_explainer(predictor, X_train: pd.DataFrame, y_train,
                           model_type: str = "autogluon",
                           label: str = "Churn Model"):
    """
    Build a DALEX Explainer wrapping any model.

    DALEX is model-agnostic — it wraps a predict_proba function so the same
    explainability workflow applies to AutoGluon, CatBoost, sklearn, etc.

    Returns
    -------
    exp : dalex.Explainer
    """
    import dalex as dx

    if model_type == "autogluon":
        def predict_fn(model, data):
            # DALEX calls predict_function(model, data) — model arg ignored,
            # we close over `predictor` directly for AutoGluon's pandas API.
            df_in = pd.DataFrame(data, columns=FEATURE_COLS)
            return predictor.predict_proba(df_in)[1].values
    else:
        def predict_fn(model, data):
            df_in = pd.DataFrame(data, columns=FEATURE_COLS)
            return model.predict_proba(df_in)[:, 1]

    exp = dx.Explainer(
        model          = predictor,
        data           = X_train[FEATURE_COLS],
        y              = y_train.values,
        predict_function = predict_fn,
        label          = label,
        verbose        = False,
    )
    print(f"✅ DALEX Explainer built: '{label}'")
    print(f"   {len(X_train):,} training samples | {len(FEATURE_COLS)} features\n")
    return exp


def plot_dalex_variable_importance(exp, n_features: int = 15):
    """
    Permutation-based variable importance via DALEX model_parts().

    Unlike SHAP, this is model-agnostic and measures the drop in loss
    when a feature is permuted — interpretable as 'how much does the
    model rely on this feature?'
    """
    print("📊 DALEX Variable Importance (permutation-based)")
    print("   Measures: drop in AUC when each feature is randomly shuffled\n")
    vi = exp.model_parts(loss_function="1-auc", N=500, B=5)
    vi.plot(max_vars=n_features, title="Variable Importance — Permutation (DALEX)", show=True)


def plot_dalex_pdp(exp, variables: list = None):
    """
    Partial Dependence Profiles (PDP) via DALEX model_profile().

    Shows how predicted churn probability changes as a single feature
    varies across its range — holding all others constant.
    Useful for understanding non-linear effects and monotonicity.
    """
    if variables is None:
        variables = ["dd_count_30d", "digital_engagement_score",
                     "days_since_last_dd", "avg_daily_balance",
                     "friction_score", "tenure_months"]
    print("📈 DALEX Partial Dependence Profiles (PDP)")
    print(f"   Features: {', '.join(variables)}\n")
    pdp = exp.model_profile(variables=variables, type="partial", N=500)
    pdp.plot(title="Partial Dependence Profiles — How Features Drive Churn (DALEX)", show=True)


def plot_dalex_ale(exp, variables: list = None):
    """
    Accumulated Local Effects (ALE) via DALEX model_profile(type='accumulated').

    ALE is preferred over PDP when features are correlated — it avoids
    extrapolating into unrealistic feature combinations.
    """
    if variables is None:
        variables = ["dd_count_30d", "digital_engagement_score",
                     "days_since_last_dd", "balance_stress_index"]
    print("📊 DALEX Accumulated Local Effects (ALE)")
    print("   Preferred over PDP when features are correlated\n")
    ale = exp.model_profile(variables=variables, type="accumulated", N=500)
    ale.plot(title="Accumulated Local Effects (ALE) — DALEX", show=True)


def plot_dalex_breakdown(exp, customer: pd.DataFrame, label: str = "Mystery Customer"):
    """
    Break-down plot for a single customer — DALEX's answer to SHAP waterfall.

    Shows how each feature's value pushed the prediction up or down from
    the model's average prediction. Includes interaction detection.
    """
    print(f"🔍 DALEX Break-Down Plot: {label}")
    print("   Shows feature contributions to this individual's prediction\n")
    bd = exp.predict_parts(
        new_observation = customer[FEATURE_COLS],
        type            = "break_down_interactions",
        label           = label,
    )
    bd.plot(title=f"Break-Down: {label}", show=True)
    print(f"\n   Predicted probability: {bd.result['cumulative'].iloc[-1]:.3f}")
    return bd


def compare_dalex_models(exp_list: list):
    """
    Compare multiple models side-by-side using DALEX.

    Pass a list of DALEX Explainer objects (e.g. CatBoost vs. AutoGluon ensemble).
    Produces variable importance comparison and residual plots.

    Example
    -------
    compare_dalex_models([exp_catboost, exp_autogluon])
    """
    if len(exp_list) < 2:
        print("Need at least 2 explainers to compare.")
        return

    print(f"⚖️  Comparing {len(exp_list)} models via DALEX\n")
    # Variable importance comparison
    vis = [exp.model_parts(loss_function="1-auc", N=300, B=3) for exp in exp_list]
    vis[0].plot(*vis[1:], title="Variable Importance Comparison — DALEX", show=True)


# ═══════════════════════════════════════════════════════════════════════════
# 11.  AI + ML EXPLAINER  (narrative interpretation of SHAP)
# ═══════════════════════════════════════════════════════════════════════════

def ai_narrative_explainer(customer: pd.Series, churn_prob: float):
    """
    Simulate what an AI layer would say about this customer's churn risk.

    This demonstrates the ML + AI combination:
      ML model  → precision probability
      AI layer  → plain-language explanation + recommended action
    """
    risk_level = "HIGH" if churn_prob > 0.65 else "MODERATE" if churn_prob > 0.35 else "LOW"
    color_map  = {"HIGH": "🔴", "MODERATE": "🟡", "LOW": "🟢"}

    print("\n" + "─" * 60)
    print("  🤖  AI NARRATIVE LAYER  (what LLM adds on top of the model)")
    print("─" * 60)
    print(f"""
  Risk Level: {color_map[risk_level]} {risk_level}  ({churn_prob*100:.0f}% probability)

  Summary:
  This customer shows {'several concerning' if risk_level == 'HIGH' else 'some'} signals
  of disengagement. Their direct deposit activity is
  {'very low, suggesting another institution is their primary bank.' if customer.dd_count_30d < 2
   else 'active, which is a positive retention signal.'}
  Digital engagement is {'declining — only' if customer.digital_login_30d < 5 else 'healthy at'}
  {customer.digital_login_30d} logins in the past 30 days.
  {'The ' + str(customer.fee_waiver_req_12m) + ' fee waiver request(s) indicate friction.' if customer.fee_waiver_req_12m > 0 else ''}

  Recommended Action:
  {'→ Immediate outreach — assign to retention team. Consider fee reversal,  personalized product offer (savings rate promo or CD ladder).' if risk_level == 'HIGH'
   else '→ Proactive engagement. Enroll in digital banking rewards program.  Trigger educational campaign about unused products.' if risk_level == 'MODERATE'
   else '→ No immediate action needed. Flag for quarterly relationship review.'}
""")
    print("─" * 60)
    print("  Note: ML model handles precision. AI handles communication.")
    print("  Neither alone is the full solution.\n")


# ═══════════════════════════════════════════════════════════════════════════
# 11.  QUICK-START
# ═══════════════════════════════════════════════════════════════════════════

def run_full_pipeline(n_customers: int = 6000, time_limit: int = 120):
    """
    End-to-end pipeline shortcut for quick demos.
    Returns: df, train, val, test, predictor, shap_values, X_sample, explainer
    """
    print("=" * 60)
    print("  BANKING CHURN — Full Pipeline")
    print("=" * 60 + "\n")

    df           = generate_banking_data(n_customers)
    df           = engineer_features(df)
    train, val, test = temporal_split(df)
    predictor, model_type = train_model(train, val, time_limit=time_limit)
    y_scores     = evaluate_model(predictor, test, model_type)
    shap_vals, X_samp, explainer = compute_shap_values(predictor, test, model_type)

    return df, train, val, test, predictor, model_type, shap_vals, X_samp, explainer
