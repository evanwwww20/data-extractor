import streamlit as st
import pandas as pd
from extractor import (
    extract_from_workbook, summarize_totals, ExtractionError,
    extract_t12_financials, compute_cap_rate, compute_price_from_cap,
    stats_on_rent_roll, annual_gross_rent_from_rent_roll,
    per_unit_expense, price_per_unit
)
import io

st.set_page_config(page_title="Rent Roll & T12 Analyzer", layout="wide")
st.title("🏢 Rent Roll & T12 Analyzer (No‑code)")
st.caption("Drop Excel files and get metrics even if formats differ. Extracts per‑unit Market Rent & Rent, computes NOI, Cap Rates, Annual Gross Rent, Per‑Unit metrics, and more.")

col_rr, col_t12 = st.columns(2)
with col_rr:
    st.subheader("📥 Upload Rent Roll(s)")
    uploaded_rr = st.file_uploader("Upload rent roll Excel files", type=["xlsx","xlsm"], accept_multiple_files=True, key="rr")
with col_t12:
    st.subheader("📥 Upload T12 Financials (optional)")
    uploaded_t12 = st.file_uploader("Upload T12 Excel file(s)", type=["xlsx","xlsm"], accept_multiple_files=True, key="t12")

st.markdown("---")
st.subheader("🔢 Inputs (optional)")
col1, col2, col3 = st.columns(3)
with col1:
    input_cap_rate = st.number_input("Cap Rate (e.g., 0.065 for 6.5%)", min_value=0.0, max_value=1.0, value=0.0, step=0.001, format="%.3f")
with col2:
    input_purchase_price = st.number_input("Purchase Price (USD)", min_value=0.0, value=0.0, step=1000.0, format="%.2f")
with col3:
    input_total_units = st.number_input("Total Units (override)", min_value=0, value=0, step=1)

run_btn = st.button("Run Analysis", type="primary", disabled=not (uploaded_rr or uploaded_t12))

if run_btn:
    rr_results = []
    rr_errors = []
    t12_results = []
    t12_errors = []

    if uploaded_rr:
        with st.spinner("Parsing rent rolls..."):
            for up in uploaded_rr:
                try:
                    content = up.read()
                    df_units, meta = extract_from_workbook(content, filename=up.name)
                    totals = summarize_totals(df_units)
                    stats = stats_on_rent_roll(df_units)
                    annual_gross = annual_gross_rent_from_rent_roll(df_units)
                    rr_results.append((up.name, df_units, totals, stats, annual_gross, meta))
                except ExtractionError as e:
                    rr_errors.append({"file": up.name, "error": str(e)})
                except Exception as e:
                    rr_errors.append({"file": up.name, "error": f"Unexpected error: {e}"})

    if uploaded_t12:
        with st.spinner("Parsing T12 files..."):
            for up in uploaded_t12:
                try:
                    content = up.read()
                    t12 = extract_t12_financials(content, filename=up.name)
                    t12_results.append((up.name, t12))
                except ExtractionError as e:
                    t12_errors.append({"file": up.name, "error": str(e)})
                except Exception as e:
                    t12_errors.append({"file": up.name, "error": f"Unexpected error: {e}"})

    # Show errors
    if rr_errors or t12_errors:
        st.subheader("⚠️ Errors")
        for err in rr_errors + t12_errors:
            st.error(f"{err['file']}: {err['error']}")

    # Display rent roll outputs
    if rr_results:
        st.subheader("✅ Rent Roll Results")
        all_rows = []
        total_units_detected = 0
        total_annual_gross = 0.0

        for name, df_units, totals, stats, annual_gross, meta in rr_results:
            st.markdown(f"### 📄 {name}")
            st.write("**Detected columns:** ", ", ".join(df_units.columns.astype(str)))
            st.dataframe(df_units, use_container_width=True)

            st.write("**Per-file Totals**")
            st.json(totals)
            st.write("**Rent Statistics (leased rent)**")
            st.json(stats)
            st.write(f"**Annual Gross Rent (sum of rents × 12):** ${annual_gross:,.2f}")

            csv_bytes = df_units.to_csv(index=False).encode("utf-8")
            st.download_button(
                label=f"Download per-unit CSV for {name}",
                data=csv_bytes,
                file_name=f"{name}_units.csv",
                mime="text/csv"
            )
            all_rows.append(df_units.assign(_source=name))
            total_units_detected += int(df_units['unit'].notna().sum())
            total_annual_gross += annual_gross

        combined = pd.concat(all_rows, ignore_index=True)
        st.markdown("### 📦 Combined CSV (all rent rolls)")
        st.dataframe(combined, use_container_width=True)
        st.download_button(
            label="Download combined CSV",
            data=combined.to_csv(index=False).encode("utf-8"),
            file_name="combined_units.csv",
            mime="text/csv"
        )

        st.info(f"Detected **{total_units_detected}** units across uploaded rent rolls. Combined annual gross rent: **${total_annual_gross:,.2f}**")

    # Display T12 outputs & cap rate/price helpers
    if t12_results or input_cap_rate or input_purchase_price or input_total_units:
        st.subheader("📊 Financials & Cap Rate Tools")

        # Aggregate T12 if provided
        t12_df = None
        if t12_results:
            frames = []
            for name, t12 in t12_results:
                st.markdown(f"#### 📄 {name} — Parsed T12 snapshot")
                st.json(t12)
                frames.append(pd.DataFrame([dict(file=name, **t12)]))
            t12_df = pd.concat(frames, ignore_index=True)
            st.dataframe(t12_df, use_container_width=True)

        # Compute Cap Rate or Purchase Price if inputs available
        # Prefer T12 NOI if uploaded; else rely solely on inputs.
        noi_source = None
        if t12_df is not None and "noi" in t12_df.columns:
            total_noi = float(t12_df["noi"].sum())
            noi_source = "T12 upload"
        else:
            total_noi = None

        # Determine units
        if input_total_units and input_total_units > 0:
            units_for_calc = input_total_units
        else:
            # try detect from rent rolls
            units_for_calc = total_units_detected if rr_results else 0

        # Widgets show derived calculations
        if total_noi is not None and input_purchase_price > 0:
            cap = compute_cap_rate(total_noi, input_purchase_price)
            st.success(f"**Cap Rate (NOI / Price)** using NOI ${total_noi:,.2f} and Price ${input_purchase_price:,.2f} = **{cap:.3%}** (NOI source: {noi_source})")
        if total_noi is not None and input_cap_rate > 0:
            price = compute_price_from_cap(total_noi, input_cap_rate)
            st.success(f"**Implied Purchase Price (NOI / Cap)** using NOI ${total_noi:,.2f} and Cap {input_cap_rate:.3%} = **${price:,.2f}**")
        if input_purchase_price > 0 and units_for_calc > 0:
            ppu = price_per_unit(input_purchase_price, units_for_calc)
            st.info(f"**Price per Unit** = ${ppu:,.2f} (Price ${input_purchase_price:,.2f} / {units_for_calc} units)")
        if t12_df is not None and "total_operating_expenses" in t12_df.columns and units_for_calc > 0:
            pue = per_unit_expense(float(t12_df["total_operating_expenses"].sum()), units_for_calc)
            st.info(f"**Per Unit Operating Expense** = ${pue:,.2f} (Annual OpEx / Units)")

else:
    st.info("Upload rent roll and/or T12 files, set any optional inputs, and click **Run Analysis**.")
