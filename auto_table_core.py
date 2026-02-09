import os
import json
from datetime import datetime

import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Spreadsheet IDs
# Starlink activation status source (new sheet)
SPREADSHEET_ID_STARLINK = "1XdByRZ3zYX5pfqEoufLXb3qnPTIh2rBmnfl4JzWoEbQ"
# Main schedule / outcome data
SPREADSHEET_ID_MAIN = "1zchK5za6LM5aj91s4KDn-CCgNJ5vQFDsCk_ov4XsSn4"

# Column names with line breaks in headers (as in the Sheets)
SCHEDULE_COL = "Schedule of Delivery/\\nInstallation\\n(Start Date)".replace("\\n", "\n")
SCHEDULE_END_COL = "Schedule of Delivery/\\nInstallation\\n(End Date)".replace("\\n", "\n")
OUTCOME_COL = "Outcome Status \\n (to be Accomplished by Supplier)".replace("\\n", "\n")
BLOCKER_COL = "Blocker \\n (to be Accomplished by Supplier)".replace("\\n", "\n")


def _build_sheets_service():
    """Build Google Sheets service, using env var JSON if available."""
    json_env = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if json_env:
        # Be forgiving if the value was pasted with surrounding quotes
        json_env = json_env.strip()
        if json_env and json_env[0] in ("'", '"') and json_env[-1] == json_env[0]:
            json_env = json_env[1:-1]
        sa_info = json.loads(json_env)
        creds = Credentials.from_service_account_info(
            sa_info,
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
        )
    else:
        # Fallback for local use with JSON file
        service_account_file = "monitoring-dashboard-485505-73f943f6722d.json"
        creds = Credentials.from_service_account_file(
            service_account_file,
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
        )
    return build("sheets", "v4", credentials=creds)


_sheets_service = _build_sheets_service()


def _load_df(spreadsheet_id: str, sheet_name: str) -> pd.DataFrame:
    """Generic loader for a given sheet."""
    result = (
        _sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!A1:ZZ")
        .execute()
    )
    values = result.get("values", [])
    if not values:
        return pd.DataFrame()

    header, *rows = values
    cols_len = len(header)
    normalized_rows = []
    for r in rows:
        if len(r) < cols_len:
            r = r + [""] * (cols_len - len(r))
        elif len(r) > cols_len:
            r = r[:cols_len]
        normalized_rows.append(r)
    return pd.DataFrame(normalized_rows, columns=header)


def load_starlink_df() -> pd.DataFrame:
    """Load BEIS School ID + activation status from the activation sheet."""
    df = _load_df(SPREADSHEET_ID_STARLINK, "Master")
    required = ["BEIS School ID", "Status of Activation", "Approval (Accepted / Decline) "]
    if df.empty or any(col not in df.columns for col in required):
        return pd.DataFrame(columns=required)

    df = df[required].copy()
    # Drop duplicate "Status of Activation" columns, keep the first
    df = df.loc[:, ~df.columns.duplicated()]
    df["BEIS School ID"] = df["BEIS School ID"].astype(str).str.strip()
    df["Status of Activation"] = (
        df["Status of Activation"].fillna("").astype(str).str.strip()
    )
    df["Approval (Accepted / Decline) "] = (
        df["Approval (Accepted / Decline) "].fillna("").astype(str).str.strip()
    )

    # In case of duplicates, keep the last occurrence
    df = df.replace({"": pd.NA}).dropna(subset=["BEIS School ID"]).drop_duplicates(
        subset=["BEIS School ID"], keep="last"
    )
    return df


def load_main_df() -> pd.DataFrame:
    """Load main schedule/outcome data from the second sheet."""
    df = _load_df(SPREADSHEET_ID_MAIN, "Master")

    # In this sheet, the first column header is blank but contains Region values.
    # Normalize that header to "Region" so we can work with it.
    if not df.empty and df.columns[0].strip() == "":
        cols = list(df.columns)
        cols[0] = "Region"
        df.columns = cols

    # Ensure end-date schedule column exists (older sheets may not have it yet)
    if not df.empty and SCHEDULE_END_COL not in df.columns:
        df[SCHEDULE_END_COL] = ""

    # Core columns we must have
    required = [
        "Region",
        "Province",
        "BEIS School ID",
        SCHEDULE_COL,
        SCHEDULE_END_COL,
        "Start Time",
        "End Time",
        OUTCOME_COL,
        BLOCKER_COL,
        "Status of Calendar",
    ]
    if df.empty or any(col not in df.columns for col in required):
        return pd.DataFrame(columns=required)

    # Always treat column B (index 1) as Division, regardless of header.
    # This matches the current layout of the master sheet.
    if not df.empty and df.shape[1] > 1:
        df["Division"] = df.iloc[:, 1].astype(str).str.strip()
    else:
        df["Division"] = ""

    # Optional columns that we show if present (e.g., Division, Final Status, Validated?)
    optional = ["Division", "Final Status", "Validated?"]
    cols = required + [c for c in optional if c in df.columns]
    df = df[cols].copy()

    # Clean up core text fields
    for col in ["Region", "Division", "Province", "BEIS School ID"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Clean schedule, outcome, blocker
    df[SCHEDULE_COL] = df[SCHEDULE_COL].fillna("").astype(str).str.strip()
    df[SCHEDULE_END_COL] = df[SCHEDULE_END_COL].fillna("").astype(str).str.strip()
    df[OUTCOME_COL] = df[OUTCOME_COL].fillna("").astype(str).str.strip()
    df[BLOCKER_COL] = df[BLOCKER_COL].fillna("").astype(str).str.strip()
    df["Status of Calendar"] = df["Status of Calendar"].fillna("").astype(str).str.strip()
    df[BLOCKER_COL] = df[BLOCKER_COL].fillna("").astype(str).str.strip()

    # Do not drop rows with blank schedule here; keep full data set.
    # The main view can choose to filter out unscheduled rows, while
    # an alternate "full" mode can include everything.
    return df


def get_table_data(
      selected_region: str | None = None,
      selected_schedule=None,
      selected_installation: str | None = None,
      selected_tile: str | None = None,
      selected_lot: str | None = None,
      selected_final: str | None = None,
      selected_validated: str | None = None,
      include_unscheduled: bool = False,
      selected_search: str | None = None,
  ):
    """Return rows, filter options, and stats for the dashboard."""
    df_main = load_main_df()
    if df_main.empty:
        return [], [], [], [], [], [], {
            "active": False,
            "star_activated": 0,
            "star_not_activated": 0,
            "approval_accepted": 0,
            "approval_pending": 0,
            "approval_decline": 0,
            "calendar_sent": 0,
            "calendar_not_sent": 0,
            "s1_success": 0,
            "scheduled": 0,
            "unscheduled": 0,
        }

    df_star = load_starlink_df()

    # Join Starlink activation status by BEIS School ID
    if not df_star.empty:
        df_merged = df_main.merge(
            df_star, on="BEIS School ID", how="left", suffixes=("", "_starlink")
        )
        df_merged["Starlink Status"] = df_merged["Status of Activation"]
    else:
        df_merged = df_main.copy()
        df_merged["Starlink Status"] = ""

    # Installation Status derived from Outcome Status (for now, just mirror it)
    df_merged["Installation Status"] = df_merged[OUTCOME_COL]

    # Calendar status derived from "Status of Calendar"
    def _map_calendar_status(val: str) -> str:
        v = (val or "").strip()
        if not v:
            return ""
        if v == "Invite Sent":
            return "Sent"
        # Any other non-empty value is treated as "Invite Not Sent"
        return "Invite Not Sent"

    if "Status of Calendar" in df_merged.columns:
        df_merged["Calendar Status"] = df_merged["Status of Calendar"].apply(_map_calendar_status)
    else:
        df_merged["Calendar Status"] = ""

    # Expose cleaned schedule (start/end) as simple fields (strings)
    df_merged["Schedule"] = df_merged[SCHEDULE_COL].fillna("").astype(str).str.strip()
    df_merged["Schedule_end_raw"] = df_merged[SCHEDULE_END_COL].fillna("").astype(str).str.strip()

    # For sorting, parse schedule as a date where possible with several formats
    def _parse_schedule(val: str):
        val = (val or "").strip()
        if not val:
            return pd.NaT
        # Explicit formats we expect to see
        fmts = [
            "%b %d, %Y",      # Feb 05, 2026
            "%b. %d, %Y",     # Feb. 05, 2026
            "%B %d, %Y",      # February 17, 2026
            "%d-%b-%y",       # 05-Feb-26
            "%d-%b-%Y",       # 05-Feb-2026
            "%m/%d/%y",       # 02/04/26 (MM/DD/YY)
            "%m/%d/%Y",       # 02/04/2026
            "%d/%m/%y",       # 04/02/26 (DD/MM/YY)
            "%d/%m/%Y",       # 04/02/2026
        ]
        for fmt in fmts:
            try:
                return datetime.strptime(val, fmt)
            except Exception:
                continue
        # Fallback to pandas parser
        try:
            return pd.to_datetime(val, errors="raise")
        except Exception:
            return pd.NaT

    df_merged["Schedule_sort"] = df_merged["Schedule"].apply(_parse_schedule)
    # Parse times for better ordering within a day
    df_merged["Start_sort"] = pd.to_datetime(
        df_merged["Start Time"], errors="coerce", format="%I:%M %p"
    )
    df_merged["End_sort"] = pd.to_datetime(
        df_merged["End Time"], errors="coerce", format="%I:%M %p"
    )

    # Build a consistent display string for schedule dates, e.g. "Feb. 02, 2026 - Feb. 05, 2026"
    def _format_schedule(row):
        ts_start = row.get("Schedule_sort")
        start_raw = row["Schedule"]
        # Format start date
        if pd.notna(ts_start):
            try:
                start_text = ts_start.strftime("%b. %d, %Y")
            except Exception:
                start_text = start_raw
        else:
            start_text = start_raw

        # Format end date
        end_raw = row.get("Schedule_end_raw", "") or ""
        end_raw = str(end_raw).strip()
        if not end_raw:
            end_text = "-"
        else:
            ts_end = _parse_schedule(end_raw)
            if pd.notna(ts_end):
                try:
                    end_text = ts_end.strftime("%b. %d, %Y")
                except Exception:
                    end_text = end_raw
            else:
                end_text = end_raw

        if not start_text:
            return ""
        return f"{start_text} - {end_text}"

    df_merged["Schedule_display"] = df_merged.apply(_format_schedule, axis=1)

    # In the default view we only consider rows with a schedule.
    # In "full" mode we keep unscheduled rows as well.
    if not include_unscheduled:
        df_merged = df_merged[df_merged["Schedule"] != ""].copy()

    # Sort by earliest schedule date, then start/end time, then by region/province/school
    df_sorted = df_merged.sort_values(
        by=["Schedule_sort", "Start_sort", "End_sort", "Region", "Province", "BEIS School ID"],
        kind="stable",
    )

    # Optional Lot # filter (maps lot to a set of regions)
    lot_map = {
        "Lot #1": {
            "Region I",
            "Region II",
            "Region III",
            "Region IV-A",
            "Region IV-B",
            "MIMAROPA",
            "Region V",
            "CAR",
        },
        "Lot #2": {
            "Region VI",
            "Region VII",
            "Region VIII",
            "NIR",
        },
        "Lot #3": {
            "Region IX",
            "Region X",
            "Region XI",
            "Region XII",
            "Region CARAGA",
        },
    }
    lot = (selected_lot or "").strip()
    if lot in lot_map:
        allowed_regions = lot_map[lot]
        df_sorted = df_sorted[df_sorted["Region"].isin(allowed_regions)]

    # All distinct regions for filter options
    region_options = sorted(
        r for r in df_sorted["Region"].astype(str).unique() if r.strip()
    )
    # All distinct schedule display values for filter options, ordered by date
    sched_unique = (
        df_sorted[["Schedule_display", "Schedule_sort"]]
        .drop_duplicates()
        .sort_values(["Schedule_sort", "Schedule_display"])
    )
    schedule_options = [s for s in sched_unique["Schedule_display"].tolist() if s]

    # All distinct installation statuses for filter options
    inst_unique = (
        df_sorted["Installation Status"]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
    )
    installation_options = sorted(inst_unique.tolist())

    # All distinct Final Status values for filter options
    if "Final Status" in df_sorted.columns:
        final_unique = (
            df_sorted["Final Status"]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
        )
        final_status_options = sorted(final_unique.tolist())
    else:
        final_status_options = []

    # All distinct Validated? values for filter options
    if "Validated?" in df_sorted.columns:
        val_unique = (
            df_sorted["Validated?"]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
        )
        validated_options = sorted(val_unique.tolist())
    else:
        validated_options = []

    # Optional filters
    if selected_region:
        df_sorted = df_sorted[df_sorted["Region"] == selected_region]
    if selected_schedule:
        if isinstance(selected_schedule, (list, tuple, set)):
            df_sorted = df_sorted[df_sorted["Schedule_display"].isin(selected_schedule)]
        else:
            df_sorted = df_sorted[df_sorted["Schedule_display"] == selected_schedule]
    if selected_installation:
        # Special value for blank Installation Status
        if selected_installation == "__blank__":
            df_sorted = df_sorted[
                df_sorted["Installation Status"]
                .fillna("")
                .astype(str)
                .str.strip()
                == ""
            ]
        else:
            df_sorted = df_sorted[df_sorted["Installation Status"] == selected_installation]
    if selected_final and "Final Status" in df_sorted.columns:
        df_sorted = df_sorted[df_sorted["Final Status"] == selected_final]
    if selected_validated and "Validated?" in df_sorted.columns:
        df_sorted = df_sorted[df_sorted["Validated?"] == selected_validated]
    # Free-text search across key columns
    if selected_search:
        needle = selected_search.strip()
        if needle:
            cols_to_search = [
                "Region",
                "Province",
                "BEIS School ID",
                "Schedule_display",
                "Installation Status",
                "Starlink Status",
                "Approval (Accepted / Decline) ",
                "Final Status",
                "Validated?",
                BLOCKER_COL,
            ]
            masks = []
            for col in cols_to_search:
                if col in df_sorted.columns:
                    series = df_sorted[col].fillna("").astype(str)
                    masks.append(series.str.contains(needle, case=False, na=False))
            if masks:
                combined = masks[0]
                for m in masks[1:]:
                    combined |= m
                df_sorted = df_sorted[combined]

    # Build stats based on the filtered set (for selected schedule/region)
    if df_sorted.empty:
        stats = {
            "active": False,
            "star_activated": 0,
            "star_not_activated": 0,
            "approval_accepted": 0,
            "approval_pending": 0,
            "approval_decline": 0,
            "calendar_sent": 0,
            "calendar_not_sent": 0,
            "s1_success": 0,
            "scheduled": 0,
            "unscheduled": 0,
        }
    else:
        star_series = (
            df_sorted["Starlink Status"].fillna("").astype(str).str.strip().str.lower()
        )
        appr_series = (
            df_sorted["Approval (Accepted / Decline) "]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )
        cal_series = (
            df_sorted.get("Calendar Status", "")
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )
        inst_series = (
            df_sorted["Installation Status"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )

        # Apply tile-based filter if requested
        tile = (selected_tile or "").strip()
        if tile:
            mask = pd.Series(True, index=df_sorted.index)
            if tile == "star_activated":
                mask = star_series == "activated"
            elif tile == "star_not_activated":
                mask = star_series != "activated"
            elif tile == "approval_accepted":
                mask = appr_series.str.contains("accept", na=False)
            elif tile == "approval_pending":
                # not accepted and not decline => pending/blank/other
                accepted_mask_tmp = appr_series.str.contains("accept", na=False)
                decline_mask_tmp = appr_series.str.contains("declin", na=False)
                mask = ~accepted_mask_tmp & ~decline_mask_tmp
            elif tile == "approval_decline":
                mask = appr_series.str.contains("declin", na=False)
            elif tile == "calendar_sent":
                mask = cal_series == "sent"
            elif tile == "calendar_not_sent":
                mask = cal_series == "invite not sent"
            elif tile == "s1_success":
                mask = inst_series == "s1 - installed (success)"
            elif tile == "unscheduled":
                sched_series_tmp = (
                    df_sorted["Schedule"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )
                mask = sched_series_tmp == ""
            df_sorted = df_sorted[mask]
            # Recompute series for stats on the filtered set
            star_series = star_series[mask]
            appr_series = appr_series[mask]
            cal_series = cal_series[mask]
            inst_series = inst_series[mask]

        if df_sorted.empty:
            stats = {
                "active": False,
                "star_activated": 0,
                "star_not_activated": 0,
                "approval_accepted": 0,
                "approval_pending": 0,
                "approval_decline": 0,
                "calendar_sent": 0,
                "calendar_not_sent": 0,
                "s1_success": 0,
                "scheduled": 0,
                "unscheduled": 0,
            }
        else:
            total_rows = len(df_sorted)

            # Starlink: only "activated" vs everything else
            star_activated = (star_series == "activated").sum()
            star_not_activated = total_rows - star_activated

            # Approval:
            # treat any value containing "accept" as accepted,
            # any value containing "declin" (decline/declined) as decline,
            # the rest (including blank/pending/others) as pending/blank.
            accepted_mask = appr_series.str.contains("accept", na=False)
            decline_mask = appr_series.str.contains("declin", na=False)
            approval_accepted = accepted_mask.sum()
            approval_decline = decline_mask.sum()
            approval_pending = total_rows - approval_accepted - approval_decline

            # Calendar status: Sent vs Invite Not Sent
            calendar_sent = (cal_series == "sent").sum()
            calendar_not_sent = (cal_series == "invite not sent").sum()

            # S1 success count based on Installation Status
            inst_series = (
                df_sorted["Installation Status"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.lower()
            )
            s1_success = (inst_series == "s1 - installed (success)").sum()

            # Schedule coverage: scheduled vs unscheduled rows in the current view
            sched_series = (
                df_sorted["Schedule"]
                .fillna("")
                .astype(str)
                .str.strip()
            )
            unscheduled = (sched_series == "").sum()
            scheduled = total_rows - unscheduled

            stats = {
                "active": True,  # show stats for overall or filtered
                "star_activated": int(star_activated),
                "star_not_activated": int(star_not_activated),
                "approval_accepted": int(approval_accepted),
                "approval_pending": int(approval_pending),
                "approval_decline": int(approval_decline),
                "calendar_sent": int(calendar_sent),
                "calendar_not_sent": int(calendar_not_sent),
                "s1_success": int(s1_success),
                "scheduled": int(scheduled),
                "unscheduled": int(unscheduled),
            }

    rows = []
    for _, row in df_sorted.iterrows():
        # Normalize Starlink and approval text
        star = row["Starlink Status"]
        if pd.isna(star):
            star = ""
        approval_raw = row.get("Approval (Accepted / Decline) ", "")
        if pd.isna(approval_raw):
            approval_raw = ""
        approval_text = str(approval_raw).strip()
        if not approval_text:
            # Mirror dashboard behavior: treat blank/None as "Pending"
            approval_display = "Pending"
        else:
            approval_display = approval_text

        # Normalize schedule display to a consistent text format, e.g. "Feb. 02, 2026"
        schedule_display = row["Schedule"]
        ts = row.get("Schedule_sort")
        if pd.notna(ts):
            try:
                schedule_display = ts.strftime("%b. %d, %Y")  # e.g. "Feb. 02, 2026"
            except Exception:
                # Fallback to original string if formatting fails
                schedule_display = row["Schedule"]
        rows.append(
            {
                "Region": row["Region"],
                "Division": row.get("Division", ""),
                "Province": row["Province"],
                "BEIS School ID": row["BEIS School ID"],
                "Schedule": row["Schedule_display"],
                "Calendar Status": row.get("Calendar Status", ""),
                "Start Time": row["Start Time"],
                "End Time": row["End Time"],
                "Installation Status": row["Installation Status"],
                "Starlink Status": star,
                "Approval": approval_display,
                "Final Status": row.get("Final Status", ""),
                "Validated?": row.get("Validated?", ""),
                "Blocker": row[BLOCKER_COL],
            }
        )

    return (
        rows,
        region_options,
        schedule_options,
        installation_options,
        final_status_options,
        validated_options,
        stats,
    )


TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>LEOxSOLAR Schedule Monitoring</title>
    <meta http-equiv="refresh" content="300">
    <style>
        html, body {
            height: 100%;
            margin: 0;
        }
        body {
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            font-size: 12px;
            background: radial-gradient(circle at top left, #d4e0ff 0, #dde4f0 40%, #d4d4dd 100%);
            color: #020617;
            overflow: hidden; /* prevent whole-page scrolling */
        }
        .page {
            max-width: 1400px;
            height: 100%;
            margin: 0 auto;
            padding: 10px;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
        }
        h1 {
            margin: 4px 0 6px 0;
            font-size: 20px;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            background: linear-gradient(90deg, #0f172a, #0284c7, #0f172a);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
        }
        .meta-line {
            font-size: 11px;
            color: #6b7280;
            margin-bottom: 6px;
        }
        .filter-toggle-bar {
            display: flex;
            justify-content: flex-start;
            margin-bottom: 4px;
        }
        .filter-toggle-btn {
            border: none;
            border-radius: 9999px;
            padding: 4px 10px;
            font-size: 12px;
            font-weight: 600;
            background: linear-gradient(135deg, #0ea5e9, #0369a1);
            color: #ffffff;
            cursor: pointer;
            box-shadow: 0 2px 6px rgba(3, 105, 161, 0.35);
        }
        .filter-toggle-btn:hover {
            background: linear-gradient(135deg, #0369a1, #075985);
        }
        .filter-container {
            margin-bottom: 6px;
        }
        .filter-bar {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 12px;
            padding: 4px 10px;
            border-radius: 9999px;
            background: rgba(255, 255, 255, 0.85);
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.10);
        }
        .filter-bar label {
            color: #4b5563;
            font-weight: 500;
        }
        .filter-bar select {
            font-size: 12px;
            padding: 3px 10px;
            border-radius: 9999px;
            border: 1px solid rgba(148, 163, 184, 0.7);
            background-color: #ffffff;
            color: #111827;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.06);
            appearance: none;
        }
        .filter-bar select:focus {
            outline: none;
            border-color: #38bdf8;
            box-shadow:
                0 0 0 1px rgba(56, 189, 248, 0.7),
                0 4px 10px rgba(56, 189, 248, 0.25);
        }
        .filter-bar select:hover {
            border-color: #0ea5e9;
        }
        .filter-clear-link {
            margin-left: 8px;
            font-size: 11px;
            font-weight: 500;
            color: #0ea5e9;
            text-decoration: none;
            white-space: nowrap;
        }
        .filter-clear-link:hover {
            text-decoration: underline;
        }
        .report-bar {
            margin: 4px 0 8px 0;
            font-size: 11px;
            display: flex;
            justify-content: flex-start;
        }
        .report-trigger-btn {
            border: none;
            border-radius: 9999px;
            padding: 4px 10px;
            font-size: 11px;
            font-weight: 600;
            background: linear-gradient(135deg, #22c55e, #16a34a);
            color: #ffffff;
            cursor: pointer;
            box-shadow: 0 2px 6px rgba(22, 163, 74, 0.35);
        }
        .report-trigger-btn:hover {
            background: linear-gradient(135deg, #16a34a, #15803d);
        }
        .report-modal-backdrop {
            position: fixed;
            inset: 0;
            background: rgba(15, 23, 42, 0.45);
            display: none;
            align-items: center;
            justify-content: center;
            z-index: 50;
        }
        .report-modal {
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 20px 50px rgba(15, 23, 42, 0.45);
            padding: 10px 12px;
            max-width: 520px;
            width: 100%;
            font-size: 11px;
        }
        .report-modal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 6px;
        }
        .report-modal-title {
            font-weight: 700;
            color: #0f172a;
        }
        .report-modal-close {
            border: none;
            background: transparent;
            font-size: 14px;
            cursor: pointer;
            color: #6b7280;
        }
        .report-modal-body {
            display: flex;
            flex-wrap: wrap;
            gap: 4px 12px;
            margin-bottom: 8px;
        }
        .report-modal-footer {
            display: flex;
            justify-content: flex-end;
            gap: 6px;
        }
        .report-modal label {
            font-size: 10px;
            color: #4b5563;
        }
        .report-modal button.primary {
            border: none;
            border-radius: 9999px;
            padding: 4px 10px;
            font-size: 11px;
            font-weight: 600;
            background: linear-gradient(135deg, #22c55e, #16a34a);
            color: #ffffff;
            cursor: pointer;
            box-shadow: 0 2px 6px rgba(22, 163, 74, 0.35);
        }
        .report-modal button.secondary {
            border: none;
            border-radius: 9999px;
            padding: 4px 10px;
            font-size: 11px;
            font-weight: 500;
            background: #e5e7eb;
            color: #111827;
            cursor: pointer;
        }
        .card {
            background: linear-gradient(135deg, #f9fafb 0%, #e2e8f0 40%, #cbd5f5 100%);
            border-radius: 10px;
            box-shadow: 0 18px 44px rgba(15, 23, 42, 0.30);
            border: 1px solid rgba(71, 85, 105, 0.55);
            padding: 8px;
            box-sizing: border-box;
            flex: 1;                 /* take remaining space below header/meta */
            display: flex;
            flex-direction: row;
            gap: 8px;
            overflow: hidden;
        }
        .table-wrapper {
            flex: 3;
            overflow-y: auto;        /* scroll only table area */
        }
        table {
            border-collapse: collapse;
            width: 100%;
            table-layout: fixed;
        }
        th, td {
            border: 1px solid #cbd5e1;
            padding: 2px 4px;
            text-align: center;
            font-size: 11px;
        }
        th {
            background-color: #cbd5f5;
            font-weight: 600;
            color: #020617;
            white-space: normal;           /* allow header text to wrap */
            word-wrap: break-word;         /* break long tokens like 'Delivery/Installation' */
            word-break: break-word;
        }
        thead th {
            position: sticky;
            top: 0;
            z-index: 2;
        }
        tbody tr:nth-child(even) td {
            background-color: #f9fafb;
        }
        tbody tr:nth-child(odd) td {
            background-color: #ffffff;
        }
        tbody tr:hover td {
            background-color: #e5f0ff;
        }
        .row-warning td {
            background-color: #fef9c3;  /* soft yellow */
            color: #1f2933;
        }
        .row-critical td {
            background-color: #fee2e2;  /* soft red */
            color: #b91c1c;
        }
        .region-cell {
            text-align: center;
            font-weight: 600;
        }
        .school-cell {
            text-align: center;
            white-space: normal;
            word-wrap: break-word;
        }
        .status-cell {
            text-align: left;
            padding-left: 6px;
            white-space: normal;
            word-wrap: break-word;
        }
        .stats-card {
            flex: 1;
            max-width: 360px;
            padding: 6px;
            background: linear-gradient(135deg, #cbd5f5 0%, #e2e8f0 40%, #f9fafb 100%);
            border-radius: 10px;
            box-shadow:
                0 0 0 1px rgba(71, 85, 105, 0.65),
                0 20px 50px rgba(15, 23, 42, 0.32);
            font-size: 12px;
            overflow-y: auto;
            overflow-x: hidden;
            display: flex;
            flex-direction: column;
        }
        .stats-title {
            font-weight: 700;
            margin-bottom: 6px;
            color: #0f172a;
            letter-spacing: 0.05em;
            text-transform: uppercase;
            text-align: center;
        }
        .stats-main {
            flex: 1;
            display: flex;
            flex-direction: column;
            gap: 4px;
            overflow: hidden;
        }
        .stats-note {
            font-size: 11px;
            font-weight: 600;
            color: #020617;
            text-align: center;
            margin-bottom: 8px;
        }
        .stats-filter-note {
            font-size: 10px;
            color: #4b5563;
            text-align: center;
            margin-bottom: 6px;
        }
        .stats-filter-clear {
            margin-left: 6px;
            font-size: 10px;
            color: #0ea5e9;
            text-decoration: none;
        }
        .stats-filter-clear:hover {
            text-decoration: underline;
        }
        .stats-calendar-summary {
            font-size: 10px;
            color: #374151;
            margin-bottom: 4px;
        }
        .stats-grid {
            flex: 0 0 auto;
            display: grid;
            grid-template-columns: repeat(2, minmax(120px, 1fr)); /* two balanced columns */
            column-gap: 24px;          /* extra clear space between columns */
            row-gap: 8px;              /* comfortable vertical spacing */
            max-width: 340px;          /* slightly wider to keep tiles readable */
            margin: 0 auto;            /* center grid inside stats card */
        }
        .stats-item {
            width: 100%;              /* take full cell width */
            border-radius: 6px;
            background: #e0e7ff;
            border: 1px solid rgba(79, 70, 229, 0.8);
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.16);
            padding: 4px 8px;         /* a bit more breathing room */
            text-align: center;
        }
        .stats-tile-link {
            display: block;
            width: 100%;
            height: 100%;
            padding: 0;
            margin: 0;
            border: none;
            background: transparent;
            text-decoration: none;
            color: inherit;
            cursor: pointer;
        }
        .stats-label {
            color: #4b5563;
            font-size: 9px;
            line-height: 1.15;
            text-align: center;
        }
        .stats-value {
            font-weight: 700;
            font-size: 11px;
            line-height: 1.15;
            color: #0f172a;
            text-align: center;
        }
        /* Ensure text is white on error tiles like Starlink Not Activated / Calendar Invite Not Sent */
        /* Color tiles by type (using fixed order) */
        .stats-grid .stats-item:nth-child(1),
        .stats-grid .stats-item:nth-child(3),
        .stats-grid .stats-item:nth-child(6) {
            /* ✔ tiles: Starlink Activated, Approval Accepted, Calendar Sent */
            background: linear-gradient(135deg, #22c55e, #16a34a);
        }
        .stats-grid .stats-item:nth-child(1) .stats-label,
        .stats-grid .stats-item:nth-child(1) .stats-value,
        .stats-grid .stats-item:nth-child(3) .stats-label,
        .stats-grid .stats-item:nth-child(3) .stats-value,
        .stats-grid .stats-item:nth-child(6) .stats-label,
        .stats-grid .stats-item:nth-child(6) .stats-value {
            color: #ffffff;
        }
        .stats-grid .stats-item:nth-child(2),
        .stats-grid .stats-item:nth-child(4) {
            /* ⚠ tiles: Starlink Not Activated, Approval Pending/Blank */
            background: linear-gradient(135deg, #fde68a, #facc15);
        }
        .stats-grid .stats-item:nth-child(2) .stats-label,
        .stats-grid .stats-item:nth-child(2) .stats-value,
        .stats-grid .stats-item:nth-child(4) .stats-label,
        .stats-grid .stats-item:nth-child(4) .stats-value {
            color: #1f2933;
        }
        .stats-grid .stats-item:nth-child(5),
        .stats-grid .stats-item:nth-child(7) {
            /* ✖ tiles: Approval Decline, Calendar Invite Not Sent */
            background: linear-gradient(135deg, #f97373, #dc2626);
        }
        .stats-grid .stats-item:nth-child(5) .stats-label,
        .stats-grid .stats-item:nth-child(5) .stats-value,
        .stats-grid .stats-item:nth-child(7) .stats-label,
        .stats-grid .stats-item:nth-child(7) .stats-value {
            color: #ffffff;
        }
        /* Ensure text is white on all error tiles (stats-bad), overriding nth-child text colors */
        .stats-grid .stats-item.stats-bad .stats-label,
        .stats-grid .stats-item.stats-bad .stats-value {
            color: #ffffff !important;
        }
        /* Ensure text is white on warning tiles (stats-warn), e.g. Approval Pending / Blank */
        .stats-grid .stats-item.stats-warn .stats-label,
        .stats-grid .stats-item.stats-warn .stats-value {
            color: #ffffff !important;
        }
        .stats-grid .stats-item.stats-ok {
            background: linear-gradient(135deg, #22c55e, #16a34a);  /* green */
            color: #ffffff;
        }
        .stats-grid .stats-item.stats-warn {
            background: linear-gradient(135deg, #facc15, #eab308);  /* darker yellow */
            color: #1f2933;
        }
        .stats-grid .stats-item.stats-bad {
            background: linear-gradient(135deg, #f97373, #dc2626);  /* red */
            color: #ffffff;
        }
        .stats-grid .stats-item.stats-full {
            grid-column: 1 / -1; /* span full row */
        }
        .stats-charts {
            flex: 0 0 auto;
            display: flex;
            flex-direction: column;  /* stack charts vertically */
            gap: 6px;
            align-items: center;
            justify-content: center;
            margin-top: 16px;        /* extra space between tiles and charts */
        }
        .stats-chart {
            flex: 0 0 auto;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
        }
        .stats-chart .stats-label {
            font-weight: 700;
            font-size: 10px;
            margin-bottom: 4px;   /* space before chart canvas */
        }
        .stats-chart canvas {
            margin-top: 2px;      /* extra breathing room under label */
        }
        .status-pill {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 9999px;
            font-weight: 600;
            font-size: 11px;
            letter-spacing: 0.02em;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.16);
        }
        .status-ok {
            background: linear-gradient(135deg, #22c55e, #16a34a);  /* green */
            color: #ffffff;
        }
        .status-bad {
            background: linear-gradient(135deg, #f97373, #dc2626);  /* red */
            color: #ffffff;
        }
        .status-warn {
            background: linear-gradient(135deg, #fde68a, #facc15);  /* yellow */
            color: #1f2933;
        }
        .calendar-pill {
            display: inline-block;
            padding: 2px 6px;
            border-radius: 9999px;
            font-size: 11px;
            font-weight: 600;
        }
        .calendar-sent {
            background: linear-gradient(135deg, #22c55e, #16a34a);  /* green */
            color: #ffffff;
        }
        .calendar-not-sent {
            background: linear-gradient(135deg, #f97373, #dc2626);  /* red */
            color: #ffffff;
        }
        .report-bar {
            margin: 4px 0 8px 0;
            font-size: 11px;
            display: flex;
            justify-content: flex-start;
        }
        .report-trigger-btn {
            border: none;
            border-radius: 9999px;
            padding: 4px 10px;
            font-size: 11px;
            font-weight: 600;
            background: linear-gradient(135deg, #22c55e, #16a34a);
            color: #ffffff;
            cursor: pointer;
            box-shadow: 0 2px 6px rgba(22, 163, 74, 0.35);
        }
        .report-trigger-btn:hover {
            background: linear-gradient(135deg, #16a34a, #15803d);
        }
        .report-modal-backdrop {
            position: fixed;
            inset: 0;
            background: rgba(15, 23, 42, 0.45);
            display: none;
            align-items: center;
            justify-content: center;
            z-index: 50;
        }
        .report-modal {
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 20px 50px rgba(15, 23, 42, 0.45);
            padding: 10px 12px;
            max-width: 520px;
            width: 100%;
            font-size: 11px;
        }
        .report-modal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 6px;
        }
        .report-modal-title {
            font-weight: 700;
            color: #0f172a;
        }
        .report-modal-close {
            border: none;
            background: transparent;
            font-size: 14px;
            cursor: pointer;
            color: #6b7280;
        }
        .report-modal-body {
            display: flex;
            flex-wrap: wrap;
            gap: 4px 12px;
            margin-bottom: 8px;
        }
        .report-modal-footer {
            display: flex;
            justify-content: flex-end;
            gap: 6px;
        }
        .report-modal label {
            font-size: 10px;
            color: #4b5563;
        }
        .report-modal button.primary {
            border: none;
            border-radius: 9999px;
            padding: 4px 10px;
            font-size: 11px;
            font-weight: 600;
            background: linear-gradient(135deg, #22c55e, #16a34a);
            color: #ffffff;
            cursor: pointer;
            box-shadow: 0 2px 6px rgba(22, 163, 74, 0.35);
        }
        .report-modal button.secondary {
            border: none;
            border-radius: 9999px;
            padding: 4px 10px;
            font-size: 11px;
            font-weight: 500;
            background: #e5e7eb;
            color: #111827;
            cursor: pointer;
        }

        /* Responsive tweaks for smaller viewports */
        @media (max-width: 900px) {
            html, body {
                height: auto;
            }
            body {
                overflow-y: auto;
            }
            .page {
                height: auto;
            }
            .card {
                flex-direction: column;
            }
            .table-wrapper {
                flex: none;
                max-height: 55vh;
            }
            .stats-card {
                max-width: none;
                margin-top: 8px;
            }
            .stats-item {
                flex: 0 0 calc(50% - 4px); /* two columns on narrow screens */
            }
        }
    </style>
</head>
<body>
    {% if show_report %}
    <div class="report-modal-backdrop" id="report-modal-backdrop">
        <div class="report-modal">
            <div class="report-modal-header">
                <div class="report-modal-title">Generate Report</div>
                <button type="button" class="report-modal-close" id="close-report-modal">&times;</button>
            </div>
            <form method="get" action="" class="report-modal-form">
                <input type="hidden" name="region" value="{{ selected_region }}">
                {% if selected_schedule_list %}
                {% for s in selected_schedule_list %}
                <input type="hidden" name="schedule" value="{{ s }}">
                {% endfor %}
                {% else %}
                <input type="hidden" name="schedule" value="">
                {% endif %}
                <input type="hidden" name="installation" value="{{ selected_installation }}">
                <input type="hidden" name="tile" value="{{ selected_tile }}">
                <input type="hidden" name="lot" value="{{ selected_lot }}">
                <input type="hidden" name="download" value="xlsx">
                  <div class="report-modal-body">
                      <span>Include columns:</span>
                      <label><input type="checkbox" name="col" value="Region" checked> Region</label>
                      <label><input type="checkbox" name="col" value="Province" checked> Province</label>
                      <label><input type="checkbox" name="col" value="BEIS School ID" checked> BEIS ID</label>
                      <label><input type="checkbox" name="col" value="Schedule" checked> Schedule</label>
                      <label><input type="checkbox" name="col" value="Calendar Status" checked> Calendar</label>
                      <label><input type="checkbox" name="col" value="Start Time" checked> Start</label>
                      <label><input type="checkbox" name="col" value="End Time" checked> End</label>
                      <label><input type="checkbox" name="col" value="Installation Status" checked> Installation</label>
                      <label><input type="checkbox" name="col" value="Starlink Status" checked> Starlink</label>
                      <label><input type="checkbox" name="col" value="Approval" checked> Approval</label>
                      <label><input type="checkbox" name="col" value="Final Status" checked> Final Status</label>
                      <label><input type="checkbox" name="col" value="Validated?" checked> Validated?</label>
                      <label><input type="checkbox" name="col" value="Blocker" checked> Blocker</label>
                      <label>
                          <input type="checkbox" name="include_stats" value="1" checked>
                          Include summary & charts
                      </label>
                  </div>
                <div class="report-modal-footer">
                    <button type="button" class="secondary" id="cancel-report-modal">Cancel</button>
                    <button type="submit" class="primary">Download Report</button>
                </div>
            </form>
        </div>
    </div>
    {% endif %}
    <div class="page">
        <h1>LEOxSOLAR Schedule Monitoring</h1>
        <div class="meta-line">
            Auto-refresh: 5 minutes | Last update: {{ last_updated }}
        </div>
        <div class="meta-line">
            Showing {{ rows|length }} records
            • Region: {{ selected_region or 'All' }}
            • Schedule: {{ selected_schedule or 'All' }}
        </div>

          <div class="filter-toggle-bar">
              <button type="button" class="filter-toggle-btn" id="toggle-filters">
                  Filters
              </button>
          </div>
          <div class="filter-container" id="filters-container" style="display: none;">
              <form method="get" class="filter-bar">
                    {% if include_unscheduled %}
                    <input type="hidden" name="full" value="1">
                    {% endif %}
                    {% if show_report %}
                    <input type="hidden" name="report" value="1">
                    {% endif %}
                    {% if selected_tile %}
                    <input type="hidden" id="tile-input" name="tile" value="{{ selected_tile }}">
                    {% else %}
                    <input type="hidden" id="tile-input" name="tile" value="">
                    {% endif %}
                    <label for="lot-select">Lot #:</label>
                    <select id="lot-select" name="lot">
                      <option value="">All Lots</option>
                      <option value="Lot #1" {% if selected_lot == 'Lot #1' %}selected{% endif %}>Lot #1</option>
                      <option value="Lot #2" {% if selected_lot == 'Lot #2' %}selected{% endif %}>Lot #2</option>
                      <option value="Lot #3" {% if selected_lot == 'Lot #3' %}selected{% endif %}>Lot #3</option>
                  </select>
                    <label for="region-select">Region:</label>
                    <select id="region-select" name="region">
                    <option value="">All Regions</option>
                    {% for r in region_options %}
                    <option value="{{ r }}" {% if selected_region == r %}selected{% endif %}>{{ r }}</option>
                    {% endfor %}
                </select>
                  <label for="schedule-select">Schedule:</label>
                  <select id="schedule-select" name="schedule" multiple size="1">
                    {% for d in schedule_options %}
                    <option value="{{ d }}" {% if selected_schedule_list and d in selected_schedule_list %}selected{% endif %}>{{ d }}</option>
                    {% endfor %}
                </select>
                  <label for="installation-select">Installation:</label>
                    <select id="installation-select" name="installation">
                      <option value="">All Installation Statuses</option>
                      <option value="__blank__" {% if selected_installation == '__blank__' %}selected{% endif %}>No Installation Status</option>
                      {% for inst in installation_options %}
                      <option value="{{ inst }}" {% if selected_installation == inst %}selected{% endif %}>{{ inst }}</option>
                      {% endfor %}
                  </select>
                  <label for="final-status-select">Final:</label>
                  <select id="final-status-select" name="final">
                      <option value="">All</option>
                      {% for fs in final_status_options %}
                      <option value="{{ fs }}" {% if selected_final == fs %}selected{% endif %}>{{ fs }}</option>
                      {% endfor %}
                  </select>
                  <label for="validated-select">Validated:</label>
                  <select id="validated-select" name="validated">
                      <option value="">All</option>
                      {% for v in validated_options %}
                      <option value="{{ v }}" {% if selected_validated == v %}selected{% endif %}>{{ v }}</option>
                      {% endfor %}
                  </select>
                  <input
                      id="search-input"
                      type="text"
                      name="search"
                      value="{{ selected_search or '' }}"
                      placeholder="Search..."
                      style="font-size: 12px; padding: 3px 8px; border-radius: 9999px; border: 1px solid rgba(148,163,184,0.7); min-width: 140px; margin-left: 4px;"
                  />
                  <button type="submit" class="filter-toggle-btn" style="margin-left: 4px; padding: 3px 10px; font-size: 11px;">
                      Apply
                  </button>
                  <a
                      href="?{% if show_report %}report=1&{% endif %}{% if include_unscheduled %}full=1{% endif %}"
                      class="filter-clear-link"
                  >
                      Clear filters
                  </a>
              </form>
          </div>
          {% if show_report %}
          <div class="report-bar">
              <button type="button" class="report-trigger-btn" id="open-report-modal">
                  Generate Report
              </button>
          </div>
          {% endif %}

        <div class="card">
            <div class="table-wrapper">
                <table>
                    <thead>
                          <tr>
                              <th>Region</th>
                              <th>Division</th>
                              <th>Province</th>
                              <th>BEIS ID</th>
                              <th title="Schedule of Delivery/Installation (Start-End)">Schedule (Start-End)</th>
                            <th title="Status of Calendar">Calendar</th>
                            <th title="Start–End">Time</th>
                            <th title="Outcome Status (to be Accomplished by Supplier)">Installation</th>
                            <th>Starlink</th>
                            <th title="Approval (Accepted / Decline)">Approval</th>
                            <th>Final Status</th>
                            <th>Validated?</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in rows %}
                        {% set star = (row["Starlink Status"] or "") | lower %}
                        {% set appr = (row["Approval"] or "") | lower %}
                        {% set row_class = '' %}
                        {% if 'declin' in appr %}
                            {% set row_class = 'row-critical' %}
                        {% elif star != 'activated' %}
                            {% set row_class = 'row-warning' %}
                        {% endif %}
                          <tr class="{{ row_class }}">
                              <td class="region-cell">{{ row["Region"] }}</td>
                              <td>{{ row["Division"] }}</td>
                              <td>{{ row["Province"] }}</td>
                              <td class="school-cell">{{ row["BEIS School ID"] }}</td>
                            <td>{{ row["Schedule"] }}</td>
                            <td>
                                {% set cal = row["Calendar Status"] or "" %}
                                <span class="calendar-pill
                                    {% if cal == 'Sent' %}
                                        calendar-sent
                                    {% elif cal %}
                                        calendar-not-sent
                                    {% endif %}">
                                    {{ cal or '-' }}
                                </span>
                            </td>
                            <td>
                                {% if row["Start Time"] and row["End Time"] %}
                                    {{ row["Start Time"] }} - {{ row["End Time"] }}
                                {% else %}
                                    {{ row["Start Time"] or row["End Time"] }}
                                {% endif %}
                            </td>
                            <td>{{ row["Installation Status"] }}</td>
                            <td>
                                <span class="status-pill {% if star == 'activated' %}status-ok{% else %}status-bad{% endif %}">
                                    {{ row["Starlink Status"] or 'Not Activated' }}
                                </span>
                            </td>
                            <td>
                                <span class="status-pill
                                    {% if appr == 'accepted' %}
                                        status-ok
                                    {% elif appr == 'pending' or appr == '' %}
                                        status-warn
                                    {% else %}
                                        status-bad
                                    {% endif %}">
                                    {{ row["Approval"] or 'Pending' }}
                                </span>
                            </td>
                            <td>{{ row["Final Status"] }}</td>
                            <td>{{ row["Validated?"] }}</td>
                        </tr>
                        {% endfor %}
                        {% if rows|length == 0 %}
                        <tr>
                            <td colspan="12">No data available (check sheet names/columns or schedule values).</td>
                        </tr>
                        {% endif %}
                    </tbody>
                </table>
            </div>
            {% if stats.active %}
            <div class="stats-card">
                <div class="stats-title">
                    Summary for {{ selected_schedule_label or 'All Schedules' }}
                </div>
                <div class="stats-note">
                    {% if include_unscheduled %}
                    This includes all rows (even without schedule dates).<br>
                    Scheduled: {{ stats.scheduled }} | Unscheduled: {{ stats.unscheduled }}
                    {% else %}
                    This is for the sites with schedules only.
                    {% endif %}
                </div>
                {% if selected_tile %}
                <div class="stats-filter-note">
                    Status filter applied.
                    <a href="#" onclick="clearTileFilter(); return false;" class="stats-filter-clear">
                        Clear filter
                    </a>
                </div>
                {% endif %}
                    <div class="stats-main">
                          <div class="stats-grid">
                              {# Row 1: Starlink #}
                              {% if stats.star_activated %}
                              <div class="stats-item stats-ok">
                                  <button type="button" class="stats-tile-link"
                                          onclick="applyTileFilter('star_activated')">
                                      <div class="stats-label">✔ Starlink Activated</div>
                                      <div class="stats-value">{{ stats.star_activated }}</div>
                                  </button>
                              </div>
                              {% endif %}
                              {% if stats.star_not_activated %}
                              <div class="stats-item stats-bad">
                                  <button type="button" class="stats-tile-link"
                                          onclick="applyTileFilter('star_not_activated')">
                                      <div class="stats-label">✖ Starlink Not Activated</div>
                                      <div class="stats-value">{{ stats.star_not_activated }}</div>
                                  </button>
                              </div>
                              {% endif %}
                    
                              {# Row 2: Approval (Accepted / Pending) #}
                              {% if stats.approval_accepted %}
                              <div class="stats-item stats-ok">
                                  <button type="button" class="stats-tile-link"
                                          onclick="applyTileFilter('approval_accepted')">
                                      <div class="stats-label">✔ Approval Accepted</div>
                                      <div class="stats-value">{{ stats.approval_accepted }}</div>
                                  </button>
                              </div>
                              {% endif %}
                              {% if stats.approval_pending %}
                              <div class="stats-item stats-warn">
                                  <button type="button" class="stats-tile-link"
                                          onclick="applyTileFilter('approval_pending')">
                                      <div class="stats-label">⚠ Approval Pending / Blank</div>
                                      <div class="stats-value">{{ stats.approval_pending }}</div>
                                  </button>
                              </div>
                              {% endif %}
                    
                              {# Row 3: Calendar (Sent / Not Sent) #}
                              {% if stats.calendar_sent %}
                              <div class="stats-item stats-ok">
                                  <button type="button" class="stats-tile-link"
                                          onclick="applyTileFilter('calendar_sent')">
                                      <div class="stats-label">✔ Calendar Sent</div>
                                      <div class="stats-value">{{ stats.calendar_sent }}</div>
                                  </button>
                              </div>
                              {% endif %}
                              {% if stats.calendar_not_sent %}
                              <div class="stats-item stats-bad">
                                  <button type="button" class="stats-tile-link"
                                          onclick="applyTileFilter('calendar_not_sent')">
                                      <div class="stats-label">✖ Calendar Invite Not Sent</div>
                                      <div class="stats-value">{{ stats.calendar_not_sent }}</div>
                                  </button>
                              </div>
                              {% endif %}
                    
                              {# Row 4: S1 Installed (Success) #}
                              {% if stats.s1_success %}
                              <div class="stats-item stats-ok">
                                  <button type="button" class="stats-tile-link"
                                          onclick="applyTileFilter('s1_success')">
                                      <div class="stats-label">✔ S1 – Installed (Success)</div>
                                      <div class="stats-value">{{ stats.s1_success }}</div>
                                  </button>
                              </div>
                              {% endif %}
                    
                              {# Row 5: Approval Decline / Other #}
                              {% if stats.approval_decline %}
                              <div class="stats-item stats-bad">
                                  <button type="button" class="stats-tile-link"
                                          onclick="applyTileFilter('approval_decline')">
                                      <div class="stats-label">✖ Approval Decline / Other</div>
                                      <div class="stats-value">{{ stats.approval_decline }}</div>
                                  </button>
                              </div>
                              {% endif %}
                          </div>
                    
                          <div class="stats-charts">
                              <div class="stats-chart">
                                  <div class="stats-label">Starlink Status</div>
                                  <canvas id="starChart"
                                          style="width: 100%; max-width: 170px; height: 70px;"></canvas>
                              </div>
                              <div class="stats-chart">
                                  <div class="stats-label">Approval Status</div>
                                  <canvas id="approvalChart"
                                          style="width: 100%; max-width: 170px; height: 70px;"></canvas>
                              </div>
                          </div>
                      </div>
                    </div>
                </div>
            </div>
            {% endif %}
        </div>
    </div>
</body>
</html>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script>
  function applyTileFilter(tileValue) {
      const form = document.querySelector('.filter-bar');
      if (!form) return;
      let tileInput = document.getElementById('tile-input');
      if (!tileInput) {
          tileInput = document.createElement('input');
          tileInput.type = 'hidden';
          tileInput.name = 'tile';
          tileInput.id = 'tile-input';
          form.appendChild(tileInput);
      }
      tileInput.value = tileValue;
      form.submit();
  }

  function clearTileFilter() {
      const form = document.querySelector('.filter-bar');
      if (!form) return;
      const tileInput = document.getElementById('tile-input');
      if (tileInput) {
          tileInput.value = '';
      }
      form.submit();
  }

  (function () {
      const active = {{ 'true' if stats.active else 'false' }};
      if (!active) return;

      // Starlink doughnut (Activated vs Not Activated)
      const starEl = document.getElementById('starChart');
      if (starEl) {
          const starCtx = starEl.getContext('2d');
          const starData = [{{ stats.star_activated }}, {{ stats.star_not_activated }}];
          new Chart(starCtx, {
              type: 'doughnut',
              data: {
                  labels: ['Activated', 'Not Activated'],
                  datasets: [{
                      data: starData,
                      backgroundColor: [
                          'rgba(34, 197, 94, 0.90)',   // green
                          'rgba(239, 68, 68, 0.90)',   // red
                      ],
                      borderColor: '#ffffff',
                      borderWidth: 1,
                      hoverOffset: 6,
                  }]
              },
              options: {
                  responsive: true,
                  cutout: '55%',
                  plugins: {
                      legend: {
                          position: 'bottom',
                          labels: { boxWidth: 8, font: { size: 8 } }
                      },
                      tooltip: {
                          callbacks: {
                              label: function (context) {
                                  const data = context.dataset.data || [];
                                  const total = data.reduce((a, b) => a + b, 0) || 1;
                                  const value = context.parsed;
                                  const pct = ((value / total) * 100).toFixed(1);
                                  return `${context.label}: ${value} (${pct}%)`;
                              }
                          }
                      }
                  }
              }
          });
      }

      // Approval doughnut (Accepted / Pending / Decline)
      const apprEl = document.getElementById('approvalChart');
      if (apprEl) {
          const apprCtx = apprEl.getContext('2d');
          const apprData = [
              {{ stats.approval_accepted }},
              {{ stats.approval_pending }},
              {{ stats.approval_decline }},
          ];
          new Chart(apprCtx, {
              type: 'doughnut',
              data: {
                  labels: ['Accepted', 'Pending/Blank', 'Decline/Other'],
                  datasets: [{
                      data: apprData,
                      backgroundColor: [
                          'rgba(34, 197, 94, 0.90)',   // green
                          'rgba(250, 204, 21, 0.95)',  // yellow
                          'rgba(239, 68, 68, 0.90)',   // red
                      ],
                      borderColor: '#ffffff',
                      borderWidth: 1,
                      hoverOffset: 6,
                  }]
              },
              options: {
                  responsive: true,
                  cutout: '55%',
                  plugins: {
                      legend: {
                          position: 'bottom',
                          labels: { boxWidth: 8, font: { size: 8 } }
                      },
                      tooltip: {
                          callbacks: {
                              label: function (context) {
                                  const data = context.dataset.data || [];
                                  const total = data.reduce((a, b) => a + b, 0) || 1;
                                  const value = context.parsed;
                                  const pct = ((value / total) * 100).toFixed(1);
                                  return `${context.label}: ${value} (${pct}%)`;
                              }
                          }
                      }
                  }
              }
          });
      }
  })();

  (function () {
      const toggleBtn = document.getElementById('toggle-filters');
      const container = document.getElementById('filters-container');
      if (!toggleBtn || !container) return;

      let visible = false;
      function update() {
          container.style.display = visible ? 'block' : 'none';
      }
      toggleBtn.addEventListener('click', function () {
          visible = !visible;
          update();
      });
      // Keep filters hidden by default on load
      update();
  })();

  (function () {
      const backdrop = document.getElementById('report-modal-backdrop');
      const openBtn = document.getElementById('open-report-modal');
      const closeBtn = document.getElementById('close-report-modal');
      const cancelBtn = document.getElementById('cancel-report-modal');

      if (!backdrop || !openBtn || !closeBtn || !cancelBtn) return;

      function openModal() {
          backdrop.style.display = 'flex';
      }
      function closeModal() {
          backdrop.style.display = 'none';
      }

      openBtn.addEventListener('click', openModal);
      closeBtn.addEventListener('click', closeModal);
      cancelBtn.addEventListener('click', closeModal);
      backdrop.addEventListener('click', function (e) {
          if (e.target === backdrop) {
              closeModal();
          }
      });
  })();
  </script>
  """
