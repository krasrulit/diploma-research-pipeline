from __future__ import annotations

from pathlib import Path

import pandas as pd

from .load_cbonds import load_cbonds_optional
from .load_cbr import load_cbr_macro
from .load_commodities import load_world_bank_commodities
from .load_moex import load_moex_data
from .load_shortlist import load_shortlist
from .utils import create_session, ensure_directory, save_dataframe_csv


def run_pipeline(
    shortlist_path: Path,
    output_path: Path,
    processed_dir: Path,
    raw_dir: Path,
    max_companies: int | None = None,
    load_moex_history: bool = False,
    moex_history_from: str = "2019-01-01",
    moex_history_to: str | None = None,
    with_cbonds: bool = False,
) -> dict[str, pd.DataFrame]:
    ensure_directory(processed_dir)
    ensure_directory(raw_dir)
    ensure_directory(output_path.parent)

    session = create_session()

    companies_master, shortlist_log = load_shortlist(shortlist_path)
    if max_companies is not None:
        companies_master = companies_master.head(max_companies).copy()
    save_dataframe_csv(companies_master, processed_dir / "companies_master.csv")

    moex_result = load_moex_data(
        companies_master=companies_master,
        output_dir=processed_dir,
        raw_dir=raw_dir,
        session=session,
        load_history=load_moex_history,
        history_from=moex_history_from,
        history_to=moex_history_to,
    )
    cbr_result = load_cbr_macro(
        output_dir=processed_dir,
        session=session,
    )
    commodities_result = load_world_bank_commodities(
        output_dir=processed_dir,
        raw_dir=raw_dir,
        session=session,
    )
    cbonds_result = load_cbonds_optional(
        output_dir=processed_dir,
        session=session,
    ) if with_cbonds else {
        "cbonds_raw": pd.DataFrame(),
        "download_log": pd.DataFrame(),
    }

    mapping_log = moex_result.get("mapping_log", pd.DataFrame())
    download_log = pd.concat(
        [
            shortlist_log,
            moex_result.get("download_log", pd.DataFrame()),
            cbr_result.get("download_log", pd.DataFrame()),
            commodities_result.get("download_log", pd.DataFrame()),
            cbonds_result.get("download_log", pd.DataFrame()),
        ],
        ignore_index=True,
        sort=False,
    )

    results = {
        "companies_master": companies_master,
        "moex_instruments": moex_result.get("moex_instruments", pd.DataFrame()),
        "moex_bonds": moex_result.get("moex_bonds", pd.DataFrame()),
        "macro_cbr": cbr_result.get("macro_cbr", pd.DataFrame()),
        "commodity_prices_monthly": commodities_result.get("commodity_prices_monthly", pd.DataFrame()),
        "commodity_prices_annual": commodities_result.get("commodity_prices_annual", pd.DataFrame()),
        "mapping_log": mapping_log,
        "download_log": download_log,
    }

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, frame in results.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    for sheet_name, frame in results.items():
        save_dataframe_csv(frame, processed_dir / f"{sheet_name}.csv")

    if with_cbonds and not cbonds_result["cbonds_raw"].empty:
        save_dataframe_csv(cbonds_result["cbonds_raw"], processed_dir / "cbonds_raw.csv")
    if load_moex_history and not moex_result["moex_bond_history"].empty:
        save_dataframe_csv(moex_result["moex_bond_history"], processed_dir / "moex_bond_history.csv")

    return results
