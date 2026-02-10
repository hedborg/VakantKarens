#!/usr/bin/env python3
"""
Streamlit Web App for Automatisk vakansberäkning

Run with: streamlit run vakant_karens_streamlit.py
"""

import streamlit as st
import tempfile
import os

from pathlib import Path
from datetime import date
from typing import List
import pandas as pd

# Import from the main app
from vakant_karens_app import (
    process_karens_calculation,
    load_config,
    load_holidays_from_yaml,
    save_holidays_to_yaml,
    load_berakningsar_rates,
    Config,
    CONFIG_PATH,
    APP_VERSION,
    logger
)


def pdf_download(uploaded_file, label=None, key=None):
    """Render a Streamlit download button for an uploaded PDF."""
    name = label or uploaded_file.name
    st.download_button(
        label=f"📄 {name}",
        data=uploaded_file.getvalue(),
        file_name=name,
        mime="application/pdf",
        key=key,
    )


def run_and_read_excel(
    sick_pdf_data: bytes, sick_pdf_name: str,
    payslip_files: list, sjk_pdf_data: bytes, sjk_pdf_name: str,
    output_name: str, holidays, berakningsar_override: str = None,
):
    """Run calculation and read back Excel results. Returns a result dict."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        sick_pdf_path = tmpdir_path / sick_pdf_name
        with open(sick_pdf_path, "wb") as f:
            f.write(sick_pdf_data)

        payslip_paths = []
        for name, data in payslip_files:
            path = tmpdir_path / name
            with open(path, "wb") as f:
                f.write(data)
            payslip_paths.append(str(path))

        sjk_pdf_path = None
        if sjk_pdf_data and sjk_pdf_name:
            sjk_pdf_path = str(tmpdir_path / sjk_pdf_name)
            with open(sjk_pdf_path, "wb") as f:
                f.write(sjk_pdf_data)

        output_path = tmpdir_path / output_name
        config = load_config(holidays=holidays)

        process_karens_calculation(
            str(sick_pdf_path),
            payslip_paths,
            str(output_path),
            config,
            sjuklonekostnader_path=sjk_pdf_path,
            berakningsar_override=berakningsar_override or None,
        )

        with open(output_path, "rb") as f:
            excel_data = f.read()

        df_detail = pd.read_excel(output_path, sheet_name="Detalj")

        employee_sheets = {}
        employee_timlon = {}
        employee_metadata = {}
        validation_rows = []  # [(anställd, vår summa, pdf summa, status)]
        with pd.ExcelFile(output_path) as xls:
            for sn in xls.sheet_names:
                if sn in ("Detalj", "Vakanssammanfattning"):
                    continue
                raw = pd.read_excel(xls, sheet_name=sn, header=None, nrows=10)
                first_cell = str(raw.iloc[0, 0]).strip() if not raw.empty else ""

                if first_cell == "Brukare":
                    meta = {}
                    meta["brukare"] = raw.iloc[0, 1] if len(raw) > 0 else ""
                    meta["period"] = raw.iloc[1, 1] if len(raw) > 1 else ""
                    meta["anställd"] = raw.iloc[2, 1] if len(raw) > 2 else ""
                    meta["nyckel"] = raw.iloc[3, 1] if len(raw) > 3 else ""
                    meta["berakningsar"] = raw.iloc[8, 1] if len(raw) > 8 else ""
                    timlon_100 = raw.iloc[7, 1] if len(raw) > 7 else None
                    if timlon_100 is not None and pd.notna(timlon_100):
                        employee_timlon[sn] = {"rate": float(timlon_100), "multi": False}
                    employee_metadata[sn] = meta
                    tbl = pd.read_excel(xls, sheet_name=sn, header=None, skiprows=12)
                    if not tbl.empty:
                        ncols = len(tbl.columns)
                        if ncols == 7:
                            tbl.columns = [
                                "OB-klass",
                                "Sjk Timmar", "Sjk Kronor",
                                "Just Timmar", "Just Kronor",
                                "Netto Timmar", "Netto Kronor",
                            ]
                        elif ncols == 4:
                            tbl.columns = ["OB-klass", "Enl. sjuklönekostnader", "Justering för vakanser", "Netto"]
                        else:
                            tbl.columns = [f"Kol{i}" for i in range(ncols)]
                        tbl = tbl.iloc[1:]
                        tbl = tbl.dropna(how="all").reset_index(drop=True)

                        # Extract validation row if present
                        kontroll = tbl[tbl["OB-klass"] == "Kontroll mot Sjuklönekostnader"]
                        if not kontroll.empty:
                            row = kontroll.iloc[0]
                            our_val = row.get("Sjk Timmar", row.iloc[1]) if ncols == 7 else None
                            pdf_val = row.get("Sjk Kronor", row.iloc[2]) if ncols == 7 else None
                            flag = row.get("Just Timmar", row.iloc[3]) if ncols == 7 else None
                            if our_val is not None and pdf_val is not None:
                                validation_rows.append({
                                    "Anställd": sn,
                                    "Vår Summa (kr)": int(our_val) if pd.notna(our_val) else "",
                                    "PDF Summa (kr)": int(pdf_val) if pd.notna(pdf_val) else "",
                                    "Status": str(flag) if pd.notna(flag) else "",
                                })
                            # Remove validation row from displayed table
                            tbl = tbl[tbl["OB-klass"] != "Kontroll mot Sjuklönekostnader"].reset_index(drop=True)

                    employee_sheets[sn] = tbl

                elif first_cell == "Timlön":
                    rate = raw.iloc[0, 1]
                    multi = len(raw.columns) > 2 and pd.notna(raw.iloc[0, 2])
                    employee_timlon[sn] = {"rate": rate, "multi": multi}
                    employee_sheets[sn] = pd.read_excel(xls, sheet_name=sn, header=3)

                else:
                    employee_sheets[sn] = pd.read_excel(xls, sheet_name=sn)

        return {
            "excel_data": excel_data,
            "output_name": output_name,
            "df_detail": df_detail,
            "employee_sheets": employee_sheets,
            "employee_timlon": employee_timlon,
            "employee_metadata": employee_metadata,
            "validation_rows": validation_rows,
        }


def main():
    st.set_page_config(
        page_title="Automatisk vakansberäkning",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Automatisk vakansberäkning")
    st.caption(f"Version: {APP_VERSION}")
    st.markdown("""
    Beräknar karens och OB-ersättning för vakanta sjukskift baserat på uppladdade PDF-filer.
    Filerna klassificeras automatiskt baserat på filnamn:
    - **Sjuklista\\*.pdf** → Sjuklista
    - **Sjuklönekostnader\\*.pdf** → Sjuklönekostnader (valfritt)
    - **Övriga PDFs** → Lönebesked

    ---
    """)
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Inställningar")
        
        st.subheader("Helgdagar")
        st.markdown("Hantera helgdagar (sparas i config.yaml):")

        # Load holidays from config.yaml
        if "holidays" not in st.session_state:
            loaded = load_holidays_from_yaml()
            st.session_state.holidays = sorted(loaded) if loaded else []

        # Show current holidays
        with st.expander(f"Visa alla helgdagar ({len(st.session_state.holidays)})"):
            for h in sorted(st.session_state.holidays):
                col1, col2 = st.columns([3, 1])
                col1.text(f"{h.strftime('%Y-%m-%d (%A)')}")
                if col2.button("❌", key=f"remove_{h}"):
                    st.session_state.holidays.remove(h)
                    save_holidays_to_yaml(st.session_state.holidays)
                    st.rerun()

        # Add new holiday
        custom_holiday = st.date_input(
            "Lägg till helgdag",
            value=None,
            key="custom_holiday"
        )

        if st.button("Lägg till") and custom_holiday:
            if custom_holiday not in st.session_state.holidays:
                st.session_state.holidays.append(custom_holiday)
                st.session_state.holidays.sort()
                save_holidays_to_yaml(st.session_state.holidays)
                st.success(f"Tillagd: {custom_holiday}")
                st.rerun()
        
        st.divider()
        
        # Debug mode
        debug_mode = st.checkbox("Debug-läge (verbose logging)", value=False)
    
    # Initialize result state
    if "result" not in st.session_state:
        st.session_state.result = None

    # Show upload UI only when no results are displayed
    if st.session_state.result is None:
        # Single file uploader for all PDFs
        st.subheader("📁 Ladda upp PDF-filer")
        uploaded_pdfs = st.file_uploader(
            "Ladda upp alla PDF-filer (sjuklista, lönebesked, sjuklönekostnader)",
            type=["pdf"],
            accept_multiple_files=True,
            key="all_pdfs"
        )

        # Auto-classify uploaded files based on filename
        sick_pdf = None
        sjk_pdf = None
        payslip_pdfs = []

        if uploaded_pdfs:
            for pdf in uploaded_pdfs:
                name_lower = pdf.name.lower()
                if name_lower.startswith("sjuklista"):
                    sick_pdf = pdf
                elif name_lower.startswith("sjuklönekostnader") or name_lower.startswith("sjuklonekostnader"):
                    sjk_pdf = pdf
                else:
                    payslip_pdfs.append(pdf)

            # Show classification results
            st.markdown("**Automatisk klassificering:**")
            col1, col2, col3 = st.columns(3)

            with col1:
                if sick_pdf:
                    st.success(f"📄 Sjuklista: {sick_pdf.name}")
                    pdf_download(sick_pdf, key="dl_sick")
                else:
                    st.warning("📄 Sjuklista: saknas")

            with col2:
                if payslip_pdfs:
                    st.success(f"💰 Lönebesked: {len(payslip_pdfs)} st")
                    with st.expander("Visa lönebesked"):
                        for i, pdf in enumerate(payslip_pdfs, 1):
                            pdf_download(pdf, key=f"dl_payslip_{i}")
                else:
                    st.warning("💰 Lönebesked: inga")

            with col3:
                if sjk_pdf:
                    st.success(f"📋 Sjuklönekostnader: {sjk_pdf.name}")
                    pdf_download(sjk_pdf, key="dl_sjk")
                else:
                    st.info("📋 Sjuklönekostnader: ej uppladdad (valfritt)")

        # Output filename - derive from Sjuklista name
        st.divider()
        default_output = "Vakansrapport.xlsx"
        if sick_pdf:
            # Sjuklista_3061_202505.pdf => Vakansrapport_3061_202505.xlsx
            stem = Path(sick_pdf.name).stem  # e.g. "Sjuklista_3061_202505"
            suffix = stem.replace("Sjuklista", "", 1).lstrip("_")
            if suffix:
                default_output = f"Vakansrapport_{suffix}.xlsx"

        output_name = st.text_input(
            "Output filnamn",
            value=default_output,
            help="Namnet på Excel-filen som ska genereras"
        )

        # Beräkningsår override
        berakningsar_input = st.text_input(
            "Beräkningsår (valfritt)",
            value="",
            help="Lämna tomt för att använda perioden från sjuklistan. Fyll i t.ex. '2025' för att tvinga ett annat beräkningsår.",
            key="berakningsar_input",
        )

        # Process button
        st.divider()

        if st.button("Beräkna Karens & OB", type="primary", use_container_width=True):
            if not sick_pdf:
                st.error("Ingen sjuklista hittad! Filnamnet måste börja med 'Sjuklista'.")
                return

            if not payslip_pdfs:
                st.error("Inga lönebesked hittade! Alla PDFs som inte matchar 'Sjuklista*' eller 'Sjuklönekostnader*' behandlas som lönebesked.")
                return

            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                status_text.text("Bearbetar...")
                progress_bar.progress(30)

                sick_data = sick_pdf.getvalue()
                payslip_list = [(p.name, p.getvalue()) for p in payslip_pdfs]
                sjk_data = sjk_pdf.getvalue() if sjk_pdf else None
                sjk_name = sjk_pdf.name if sjk_pdf else None

                result = run_and_read_excel(
                    sick_data, sick_pdf.name,
                    payslip_list, sjk_data, sjk_name,
                    output_name, st.session_state.holidays,
                    berakningsar_override=berakningsar_input.strip(),
                )

                progress_bar.progress(100)
                status_text.text("Klar!")

                result["sick_pdf_name"] = sick_pdf.name
                result["sick_pdf_data"] = sick_data
                result["payslip_files"] = payslip_list
                result["sjk_pdf_name"] = sjk_name
                result["sjk_pdf_data"] = sjk_data
                result["berakningsar_used"] = berakningsar_input.strip()

                st.session_state.result = result
                st.rerun()

            except Exception as e:
                st.error(f"Ett fel uppstod: {str(e)}")
                if debug_mode:
                    st.exception(e)

    # Show results if available
    if st.session_state.result is not None:
        res = st.session_state.result
        df_detail = res["df_detail"]
        employee_sheets = res["employee_sheets"]
        employee_timlon = res.get("employee_timlon", {})
        employee_metadata = res.get("employee_metadata", {})

        st.success("Beräkning genomförd!")

        # Show current beräkningsår and recalculate option
        current_year = res.get("berakningsar_used", "")
        # Try to get the actual year used from employee metadata
        if not current_year:
            for meta in employee_metadata.values():
                yr = meta.get("berakningsar", "")
                if yr and pd.notna(yr):
                    current_year = str(int(yr)) if isinstance(yr, float) else str(yr)
                    break

        with st.expander(f"Beräkningsår: {current_year or '(okänt)'} — Ändra?", expanded=False):
            new_year = st.text_input(
                "Nytt beräkningsår",
                value=current_year,
                key="recalc_year",
                help="Ange beräkningsår (t.ex. 2025) och klicka Beräkna om",
            )
            if st.button("Beräkna om", type="primary", key="recalc_btn"):
                try:
                    recalc_result = run_and_read_excel(
                        res["sick_pdf_data"], res["sick_pdf_name"],
                        res["payslip_files"],
                        res.get("sjk_pdf_data"), res.get("sjk_pdf_name"),
                        res["output_name"], st.session_state.holidays,
                        berakningsar_override=new_year.strip(),
                    )
                    recalc_result["sick_pdf_name"] = res["sick_pdf_name"]
                    recalc_result["sick_pdf_data"] = res["sick_pdf_data"]
                    recalc_result["payslip_files"] = res["payslip_files"]
                    recalc_result["sjk_pdf_name"] = res.get("sjk_pdf_name")
                    recalc_result["sjk_pdf_data"] = res.get("sjk_pdf_data")
                    recalc_result["berakningsar_used"] = new_year.strip()
                    st.session_state.result = recalc_result
                    st.rerun()
                except Exception as e:
                    st.error(f"Fel vid omberäkning: {str(e)}")

        # Show uploaded files with download buttons
        with st.expander("📁 Uppladdade filer", expanded=False):
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                st.markdown(f"**📄 Sjuklista:**")
                st.download_button(
                    label=f"📄 {res['sick_pdf_name']}",
                    data=res["sick_pdf_data"],
                    file_name=res["sick_pdf_name"],
                    mime="application/pdf",
                    key="res_dl_sick",
                )
            with fc2:
                payslip_files = res.get("payslip_files", [])
                st.markdown(f"**💰 Lönebesked:** {len(payslip_files)} st")
                for i, (name, data) in enumerate(payslip_files, 1):
                    st.download_button(
                        label=f"📄 {name}",
                        data=data,
                        file_name=name,
                        mime="application/pdf",
                        key=f"res_dl_payslip_{i}",
                    )
            with fc3:
                sjk_name = res.get("sjk_pdf_name")
                sjk_data = res.get("sjk_pdf_data")
                if sjk_name and sjk_data:
                    st.markdown(f"**📋 Sjuklönekostnader:**")
                    st.download_button(
                        label=f"📄 {sjk_name}",
                        data=sjk_data,
                        file_name=sjk_name,
                        mime="application/pdf",
                        key="res_dl_sjk",
                    )
                else:
                    st.markdown("**📋 Sjuklönekostnader:** ej uppladdad")

        # Show preview
        st.subheader("📊 Resultat Preview")

        # Summary metrics
        col1, col2, col3, col4 = st.columns(4)

        total_hours = df_detail["Timmar"].sum()
        paid_hours = df_detail["Betalda timmar (vakant)"].sum()
        karens_hours = df_detail[df_detail["Status"] == "Karens"]["Timmar"].sum()
        gt14_hours = df_detail[df_detail["Status"] == "Karens och >14"]["Timmar"].sum()

        col1.metric("Totalt antal timmar", f"{total_hours:.1f}h")
        col2.metric("Sjuklön-timmar", f"{paid_hours:.1f}h")
        col3.metric("Karens-timmar", f"{karens_hours:.1f}h")
        col4.metric(">14-timmar", f"{gt14_hours:.1f}h")

        # Validation table: compare our totals vs PDF Summa
        validation_rows = res.get("validation_rows", [])
        if validation_rows:
            st.markdown("### Kontroll mot Sjuklönekostnader")
            df_val = pd.DataFrame(validation_rows)
            all_ok = all(r["Status"] == "OK" for r in validation_rows)
            if all_ok:
                st.success("✅ Alla belopp stämmer överens med Sjuklönekostnader-PDF:en")
            else:
                st.warning("⚠️ Vissa belopp avviker — kontrollera nedan")
            st.dataframe(df_val, use_container_width=True, hide_index=True)

        # Show detail table
        st.markdown("### Detaljerad uppdelning")
        st.dataframe(
            df_detail,
            use_container_width=True,
            height=400
        )

        # Per-employee pivot preview
        if employee_sheets:
            st.markdown("### Per anställd")
            selected_emp = st.selectbox(
                "Välj anställd",
                list(employee_sheets.keys()),
                key="emp_select"
            )
            if selected_emp:
                # Show metadata if available
                meta = employee_metadata.get(selected_emp, {})
                if meta:
                    mc1, mc2, mc3, mc4 = st.columns(4)
                    mc1.metric("Brukare", meta.get("brukare", "—"))
                    mc2.metric("Period", meta.get("period", "—"))
                    mc3.metric("Anställd", meta.get("anställd", "—"))
                    mc4.metric("Nyckel", meta.get("nyckel", "—"))

                timlon_info = employee_timlon.get(selected_emp)
                if timlon_info is not None:
                    rate = timlon_info["rate"]
                    sjuklon = round(rate * 0.8, 2)
                    label = f"**Timlön:** {rate:.2f} kr"
                    if timlon_info.get("multi"):
                        label += "  ⚠️ *Flera olika timlöner hittade*"
                    label += f"  \n**Sjuklön:** {sjuklon:.2f} kr (80%)"
                    st.markdown(label)
                st.dataframe(employee_sheets[selected_emp], use_container_width=True)

        # Download and reset buttons
        btn_col1, btn_col2 = st.columns(2)

        with btn_col1:
            st.download_button(
                label="💾 Ladda ner Excel-rapport",
                data=res["excel_data"],
                file_name=res["output_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with btn_col2:
            if st.button("🗑️ Rensa och ladda upp nya dokument", use_container_width=True):
                st.session_state.result = None
                # Clear the file uploader
                if "all_pdfs" in st.session_state:
                    del st.session_state["all_pdfs"]
                st.rerun()
    
    # Footer
    st.divider()
    st.markdown("""
    ### 📖 Användning

    1. **Ladda upp alla PDF-filer** i en enda uppladdning
    2. Filerna klassificeras automatiskt baserat på filnamn:
       - `Sjuklista*.pdf` → Sjuklista (obligatorisk)
       - `Sjuklönekostnader*.pdf` → Kompletterande karens/sjukdata (valfritt)
       - Alla övriga PDFs → Lönebesked
    3. **Klicka på "Beräkna"** - Systemet skapar en detaljerad Excel-rapport
    
    ### 📊 Output

    Excel-filen innehåller följande flikar:
    - **Detalj**: Alla segment med OB-klassificering och status
    - **Per anställd** (en flik per anställningsnr): Pivottabell med timmar per OB-klass och status, plus jämförelse mot sjuklönekostnader

    ### 🏷️ OB-klasser

    - **Storhelg**: Helgdagar (enl. config.yaml), inkl. kväll innan och morgon efter
    - **Helg**: Lördagar, söndagar, fredag 19-24, måndag 00-07
    - **Natt**: 22:00-06:00
    - **Kväll**: 19:00-22:00
    - **Dag**: Övrig tid
    - **Sjuk jourers helg**: Jour-timmar på helg/storhelg
    - **Sjuk jourers vardag**: Jour-timmar på vardagar
    
    ### ⚙️ Statusklassificering

    - **Karens**: Första timmarna dag 1 tills karensavdraget är förbrukat
    - **Sjuklön dag 1 - utanför karens**: Resterande timmar dag 1 efter att karens förbrukats
    - **Sjuklön dag 2-14**: Sjuklönedagar i intervallet dag 2–14
    - **Karens och >14**: Sjukperioder längre än 14 dagar (Försäkringskassan)

    Karens förbrukas över ALL sjukfrånvaro den dagen (oavsett om passet var vakant).
    Output visar ENDAST vakanta segment med korrekt karensspill.
    """)


if __name__ == "__main__":
    main()
