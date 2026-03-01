"""
ask_web.py — Streamlit Browser UI for Excel Natural Language Toolkit
=====================================================================
Run:  streamlit run ask_web.py
      (or double-click run_ask_web.bat on Windows)
"""

import sys
import io
import importlib
import tempfile
from pathlib import Path

import streamlit as st

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from nl_router import (
    parse_intent, inspect_file, get_columns_from_file,
    get_output_path, get_output_dir, INTENT_MAP,
)


# =============================================================================
# PAGE CONFIG
# =============================================================================

st.set_page_config(
    page_title="Excel NL Toolkit",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.big-header { font-size: 2rem; font-weight: 700; color: #1f77b4; }
.module-tag { background: #e8f4fd; padding: 2px 8px; border-radius: 4px;
              font-size: 0.8rem; color: #1f77b4; font-weight: 600; }
.desc-text  { color: #555; font-size: 0.9rem; }
.param-box  { background: #f8f9fa; padding: 12px; border-radius: 6px;
              border-left: 3px solid #1f77b4; margin-bottom: 8px; }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# HELPERS
# =============================================================================

def save_uploaded_file(uploaded_file) -> str:
    """Save a Streamlit UploadedFile to a temp file and return path."""
    suffix = Path(uploaded_file.name).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        return tmp.name


def execute_fn(intent: dict, kwargs: dict):
    """Import module and call function with kwargs."""
    mod = importlib.import_module(f"modules.{intent['module']}")
    fn  = getattr(mod, intent["fn"])
    return fn(**kwargs)


def confidence_color(conf: str) -> str:
    return {"high": "🟢", "medium": "🟡", "low": "🔴"}.get(conf, "⚪")


def build_kwargs_from_form(intent: dict, files: list, param_values: dict) -> dict:
    """Build final kwargs dict from collected form values."""
    kwargs = {}
    primary = files[0] if files else ""

    for p in intent["params"]:
        ptype = p["type"]
        name  = p["name"]
        val   = param_values.get(name)

        if ptype == "file":
            kwargs[name] = primary

        elif ptype in ("files", "csv_files"):
            kwargs[name] = files

        elif ptype == "file1":
            kwargs[name] = files[0] if len(files) >= 1 else ""

        elif ptype == "file2":
            kwargs[name] = files[1] if len(files) >= 2 else param_values.get("file2_path", "")

        elif ptype == "ref_file":
            kwargs[name] = param_values.get(f"ref_file_{name}", "")

        elif ptype == "output":
            kwargs[name] = get_output_path(intent["fn"], ".xlsx")

        elif ptype == "output_dir":
            kwargs[name] = get_output_dir(intent["fn"])

        elif ptype == "output_csv":
            kwargs[name] = get_output_path(intent["fn"], ".csv")

        elif ptype == "output_json":
            kwargs[name] = get_output_path(intent["fn"], ".json")

        elif ptype in ("col_req", "col_opt"):
            kwargs[name] = val if val else None

        elif ptype in ("cols_req", "cols_opt"):
            if isinstance(val, list):
                kwargs[name] = val if val else None
            elif isinstance(val, str) and val:
                kwargs[name] = [c.strip() for c in val.split(",")]
            else:
                kwargs[name] = None

        elif ptype == "number":
            try:
                kwargs[name] = int(val) if val else p.get("default", 10)
            except (TypeError, ValueError):
                kwargs[name] = p.get("default", 10)

        elif ptype == "float_val":
            try:
                kwargs[name] = float(val) if val is not None else p.get("default", 0.0)
            except (TypeError, ValueError):
                kwargs[name] = p.get("default", 0.0)

        elif ptype in ("choice", "string"):
            kwargs[name] = val if val else p.get("default", "")

        elif ptype == "bool_val":
            kwargs[name] = bool(val)

        elif ptype == "mapping":
            if isinstance(val, str) and val:
                result = {}
                for pair in val.split(","):
                    pair = pair.strip()
                    if ":" in pair:
                        k, v2 = pair.split(":", 1)
                        result[k.strip()] = v2.strip()
                kwargs[name] = result
            else:
                kwargs[name] = {}

    return kwargs


# =============================================================================
# PARAM FORM RENDERER
# =============================================================================

def render_param_form(intent: dict, files: list, query: str) -> dict:
    """
    Render Streamlit widgets for each parameter (excluding auto params).
    Returns dict of {param_name: value}.
    """
    from nl_router import extract_number_from_query

    primary = files[0] if files else ""
    columns = get_columns_from_file(primary) if primary else []
    values = {}

    # Auto-params (no UI needed)
    auto_types = {"file", "files", "file1", "file2", "output", "output_dir",
                  "output_csv", "output_json", "csv_files"}

    user_params = [p for p in intent["params"] if p["type"] not in auto_types]

    if not user_params:
        st.info("No additional parameters needed — ready to run!")
        return values

    for p in user_params:
        ptype = p["type"]
        name  = p["name"]
        prompt_text = p.get("prompt", name.replace("_", " ").title())
        default = p.get("default", "")
        options = p.get("options", [])
        optional = p.get("optional", False)
        label = f"{'* ' if not optional else ''}{prompt_text}"

        if ptype == "ref_file":
            ref_upload = st.file_uploader(
                f"Reference file for: {prompt_text}",
                type=["xlsx", "xls", "xlsm", "csv"],
                key=f"ref_{name}"
            )
            if ref_upload:
                values[f"ref_file_{name}"] = save_uploaded_file(ref_upload)

        elif ptype == "col_req":
            if columns:
                values[name] = st.selectbox(label, columns, key=f"param_{name}")
            else:
                values[name] = st.text_input(label, key=f"param_{name}")

        elif ptype == "col_opt":
            if columns:
                opts = ["(skip)"] + columns
                sel = st.selectbox(label + " (optional)", opts, key=f"param_{name}")
                values[name] = sel if sel != "(skip)" else None
            else:
                raw = st.text_input(label + " (optional)", key=f"param_{name}")
                values[name] = raw if raw else None

        elif ptype == "cols_req":
            if columns:
                values[name] = st.multiselect(label, columns, key=f"param_{name}")
            else:
                values[name] = st.text_input(label + " (comma-sep)", key=f"param_{name}")

        elif ptype == "cols_opt":
            if columns:
                sel = st.multiselect(
                    label + " (empty = all)", columns, key=f"param_{name}"
                )
                values[name] = sel if sel else None
            else:
                raw = st.text_input(label + " (comma-sep, empty=all)", key=f"param_{name}")
                values[name] = raw if raw else None

        elif ptype == "number":
            extracted = extract_number_from_query(query, None)
            init_val = int(extracted) if extracted else int(default) if default else 10
            values[name] = st.number_input(
                label, min_value=1, value=init_val, step=1, key=f"param_{name}"
            )

        elif ptype == "float_val":
            init_val = float(default) if default != "" else 0.0
            values[name] = st.number_input(
                label, value=init_val, step=0.01, key=f"param_{name}"
            )

        elif ptype == "choice":
            idx = options.index(default) if default in options else 0
            values[name] = st.selectbox(label, options, index=idx, key=f"param_{name}")

        elif ptype == "string":
            values[name] = st.text_input(label, value=str(default), key=f"param_{name}")

        elif ptype == "bool_val":
            values[name] = st.checkbox(label, value=bool(default), key=f"param_{name}")

        elif ptype == "mapping":
            values[name] = st.text_area(
                label + " (old:new, comma-sep)",
                key=f"param_{name}",
                height=80,
            )

    return values


# =============================================================================
# MAIN APP
# =============================================================================

def main():
    # ── Header ────────────────────────────────────────────────────────────────
    st.markdown('<p class="big-header">🧠 Excel Natural Language Toolkit</p>',
                unsafe_allow_html=True)
    st.markdown("Tell it **what you want to do** — no menus, no formulas.")
    st.divider()

    # ── Layout: two columns ───────────────────────────────────────────────────
    col_left, col_right = st.columns([1.2, 1])

    with col_left:
        # ── File Upload ───────────────────────────────────────────────────────
        st.subheader("📂 File(s)")
        uploaded_files = st.file_uploader(
            "Upload Excel/CSV file(s)",
            type=["xlsx", "xls", "xlsm", "csv"],
            accept_multiple_files=True,
        )

        # Save to temp files
        temp_paths = []
        if uploaded_files:
            for uf in uploaded_files:
                temp_paths.append(save_uploaded_file(uf))

        # ── Query Input ───────────────────────────────────────────────────────
        st.subheader("💬 What do you want?")
        query = st.text_area(
            "Describe the operation",
            placeholder='e.g. "remove duplicates and standardize dates"\n'
                        '"consolidate all quarterly sheets"\n'
                        '"rfm analysis"\n'
                        '"top 10 customers by revenue"',
            height=110,
            label_visibility="collapsed",
        )

        analyse_btn = st.button("🔍 Analyse Intent", type="primary",
                                disabled=(not query))

    with col_right:
        # ── File Insight ──────────────────────────────────────────────────────
        if temp_paths:
            st.subheader("📊 File Insight")
            try:
                fi = inspect_file(temp_paths[0])
                n_sheets = fi["sheet_count"]
                first_sheet = fi["sheets"][0] if fi["sheets"] else "Sheet1"
                cols = fi["columns"].get(first_sheet, [])
                n_rows = fi["row_counts"].get(first_sheet, "?")

                # Summary metrics
                m1, m2, m3 = st.columns(3)
                m1.metric("Sheets", n_sheets)
                m2.metric("Rows", f"{n_rows:,}" if isinstance(n_rows, int) else n_rows)
                m3.metric("Columns", len(cols))

                if fi["sheets"] and n_sheets > 1:
                    st.caption(f"Sheets: {', '.join(fi['sheets'][:6])}"
                               + (" ..." if n_sheets > 6 else ""))

                if cols:
                    st.caption(f"Columns: {', '.join(cols[:8])}"
                               + (" ..." if len(cols) > 8 else ""))

                if fi["domain_hint"]:
                    st.info(f"🏷️ Domain detected: **{fi['domain_hint']}**")

                if len(temp_paths) > 1:
                    st.info(f"📁 {len(temp_paths)} files uploaded")

            except Exception as e:
                st.warning(f"Could not inspect file: {e}")
        else:
            st.info("Upload a file to see its structure here.")

    # ── Analysis Results ──────────────────────────────────────────────────────
    if analyse_btn and query:
        st.divider()
        st.subheader("🎯 Matched Operations")

        with st.spinner("Analysing your request..."):
            results = parse_intent(query, temp_paths, top_n=5)

        if not results:
            st.error("No matching operation found. Try rephrasing your request.")
            st.info("**Tips:** Use specific keywords like 'remove duplicates', "
                    "'pivot by department', 'rfm analysis', 'consolidate sheets'")
            return

        # Display results as selectable cards
        intent_options = {}
        for i, r in enumerate(results):
            intent = r["intent"]
            score  = r["score"]
            conf   = r["confidence"]
            icon   = confidence_color(conf)
            pct    = f"{score*100:.0f}%"

            with st.expander(
                f"{icon} [{pct}] {intent['module']} → {intent['fn']}",
                expanded=(i == 0)
            ):
                st.markdown(f"_{intent['desc']}_")
                st.progress(score)
                intent_options[f"{intent['module']} → {intent['fn']} ({pct})"] = i

        # Operation selector
        st.subheader("⚙️ Configure & Run")

        if len(results) > 1:
            selected_label = st.selectbox(
                "Select operation to run",
                list(intent_options.keys()),
                key="op_select",
            )
            selected_idx = intent_options[selected_label]
        else:
            selected_idx = 0

        selected_result = results[selected_idx]
        selected_intent = selected_result["intent"]

        st.markdown(
            f'<span class="module-tag">{selected_intent["module"]}</span> '
            f'**{selected_intent["fn"]}** — {selected_intent["desc"]}',
            unsafe_allow_html=True
        )

        # Parameter form
        st.markdown("**Parameters:**")
        with st.form("param_form"):
            param_values = render_param_form(selected_intent, temp_paths, query)

            # Handle file2 for two-file ops if not enough files
            needs_file2 = any(p["type"] == "file2" for p in selected_intent["params"])
            file2_path = None
            if needs_file2 and len(temp_paths) < 2:
                st.markdown("---")
                st.markdown("**Second file required:**")
                file2_upload = st.file_uploader(
                    "Upload second file",
                    type=["xlsx", "xls", "xlsm", "csv"],
                    key="file2_upload"
                )
                if file2_upload:
                    file2_path = save_uploaded_file(file2_upload)
                    param_values["file2_path"] = file2_path
                    if file2_path not in temp_paths:
                        temp_paths.append(file2_path)

            run_btn = st.form_submit_button("▶ Run", type="primary")

        # ── Execution ─────────────────────────────────────────────────────────
        if run_btn:
            with st.spinner(f"Running {selected_intent['fn']}..."):
                try:
                    kwargs = build_kwargs_from_form(
                        selected_intent, temp_paths, param_values
                    )

                    # Validate required cols are provided
                    missing = []
                    for p in selected_intent["params"]:
                        if p["type"] in ("col_req", "cols_req"):
                            v = kwargs.get(p["name"])
                            if not v:
                                missing.append(p.get("prompt", p["name"]))
                    if missing:
                        st.error(f"Missing required fields: {', '.join(missing)}")
                        return

                    result = execute_fn(selected_intent, kwargs)

                    st.success(f"✅ Done! Operation **{selected_intent['fn']}** completed.")

                    # Show output path
                    output_path = kwargs.get("output_path") or kwargs.get("output_dir")
                    if output_path:
                        st.code(output_path, language=None)

                    # Download + Preview for single file output
                    result_path = result if isinstance(result, str) else output_path
                    if result_path and Path(result_path).is_file():
                        col_dl, col_prev = st.columns(2)

                        with col_dl:
                            with open(result_path, "rb") as f:
                                file_bytes = f.read()
                            st.download_button(
                                label="⬇ Download Output",
                                data=file_bytes,
                                file_name=Path(result_path).name,
                                mime="application/vnd.openxmlformats-officedocument"
                                     ".spreadsheetml.sheet",
                            )

                        with col_prev:
                            if st.button("👁 Preview (first 10 rows)"):
                                try:
                                    import pandas as pd
                                    df_prev = pd.read_excel(result_path, nrows=10)
                                    st.dataframe(df_prev, use_container_width=True)
                                except Exception as prev_e:
                                    st.warning(f"Preview not available: {prev_e}")

                    # Multi-file output list
                    elif isinstance(result, list):
                        st.info(f"Generated {len(result)} output file(s)")
                        for r in result:
                            st.code(r, language=None)

                except TypeError as te:
                    st.error(f"Parameter error: {te}")
                    st.info("Check that all column names match your file's actual columns.")
                except Exception as e:
                    st.error(f"Error: {e}")
                    import traceback
                    with st.expander("Traceback"):
                        st.code(traceback.format_exc())

    # ── Sidebar: quick reference ───────────────────────────────────────────────
    with st.sidebar:
        st.header("💡 Quick Examples")
        examples = [
            ("Clean", "remove duplicates from the file"),
            ("Clean", "fill missing values with mean"),
            ("Transform", "pivot by department and sum sales"),
            ("Compare", "find what changed between two files"),
            ("Analytics", "forecast next 12 months"),
            ("Sales", "rfm customer segmentation"),
            ("HR", "attrition analysis by department"),
            ("Finance", "accounts receivable aging report"),
            ("Project", "timesheet rollup from quarterly sheets"),
            ("Inventory", "identify dead stock"),
        ]
        for tag, ex in examples:
            st.markdown(f"**{tag}:** _{ex}_")

        st.divider()
        st.markdown("**Coverage:** 17 modules · 100+ operations")
        st.markdown("**Mode:** 100% offline — no AI/API needed")


if __name__ == "__main__":
    main()
