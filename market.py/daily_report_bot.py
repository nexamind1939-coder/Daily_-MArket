
"""
DAILY MARKET REPORT BOT
NIFTY 50 AND BANK NIFTY

Requirements:
    python -m pip install pyautogui yfinance openpyxl

Run:
    python daily_report_bot.py
"""

import os
import time
import webbrowser
from datetime import datetime
from pathlib import Path

import pyautogui

import yfinance as yf

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference


# ==========================================================
# CONFIGURATION
# ==========================================================

def find_output_dir():
    """
    If the script is accidentally inside venv or .venv,
    save reports in the main project folder instead.
    """

    script_folder = Path(__file__).resolve().parent

    for folder in (script_folder, *script_folder.parents):
        if folder.name.lower() in {"venv", ".venv"}:
            return folder.parent

    return script_folder


OUTPUT_DIR = find_output_dir()
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

TICKERS = {
    "NIFTY 50": "^NSEI",
    "BANK NIFTY": "^NSEBANK"
}

LOOKBACK = 20

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5


# ==========================================================
# FETCH MARKET DATA
# ==========================================================

def fetch_market_data(name, ticker):

    print(f"\nFetching {name} data...")

    history = yf.Ticker(ticker).history(
        period="3mo",
        interval="1d",
        auto_adjust=False
    )

    if history.empty:
        raise Exception(f"No data available for {name}")

    # Ignore today's incomplete candle
    if history.index[-1].date() >= datetime.now().date():
        if len(history) > 1:
            history = history.iloc[:-1]

    if history.empty:
        raise Exception(f"No completed trading data for {name}")

    latest = history.iloc[-1]

    previous = history.iloc[:-1].tail(LOOKBACK)

    if previous.empty:
        raise Exception("Not enough historical data for breakout analysis")

    open_price = float(latest["Open"])
    high_price = float(latest["High"])
    low_price = float(latest["Low"])
    close_price = float(latest["Close"])

    breakout_high = float(previous["High"].max())
    breakout_low = float(previous["Low"].min())

    # Volume
    volume = None

    try:
        raw_volume = latest["Volume"]

        if raw_volume is not None and float(raw_volume) > 0:
            volume = int(raw_volume)

    except (ValueError, TypeError, KeyError):
        pass

    # Breakout analysis
    if close_price > breakout_high:

        breakout = "BULLISH BREAKOUT"
        comment = "BULLISH - UPWARD TREND"
        trend = "Bullish"

    elif close_price < breakout_low:

        breakout = "BEARISH BREAKDOWN"
        comment = "BEARISH - DOWNWARD TREND"
        trend = "Bearish"

    elif close_price > open_price:

        breakout = "NO CONFIRMED BREAKOUT"
        comment = "MILDLY BULLISH - WAIT FOR CONFIRMATION"
        trend = "Mildly Bullish"

    elif close_price < open_price:

        breakout = "NO CONFIRMED BREAKOUT"
        comment = "MILDLY BEARISH - WAIT FOR CONFIRMATION"
        trend = "Mildly Bearish"

    else:

        breakout = "NO CONFIRMED BREAKOUT"
        comment = "NEUTRAL MARKET"
        trend = "Neutral"

    return {
        "name": name,
        "date": history.index[-1].date(),
        "open": open_price,
        "high": high_price,
        "low": low_price,
        "close": close_price,
        "volume": volume,
        "breakout_high": breakout_high,
        "breakout_low": breakout_low,
        "breakout": breakout,
        "comment": comment,
        "trend": trend
    }


# ==========================================================
# CREATE EXCEL REPORT
# ==========================================================

def create_excel_report(records, run_time):

    date_text = run_time.strftime("%Y-%m-%d")

    filename = f"dailymarket_report_{date_text}.xlsx"

    output_file = OUTPUT_DIR / filename

    wb = Workbook()

    ws = wb.active
    ws.title = "Daily Market Report"

    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:L1")

    ws["A1"] = "DAILY NIFTY 50 & BANK NIFTY MARKET REPORT"

    ws["A1"].font = Font(
        bold=True,
        size=16,
        color="FFFFFF"
    )

    ws["A1"].fill = PatternFill(
        "solid",
        fgColor="17365D"
    )

    ws["A1"].alignment = Alignment(horizontal="center")

    # Date and time
    ws.merge_cells("A2:L2")

    ws["A2"] = "Generated: " + run_time.strftime(
        "%d-%m-%Y %H:%M:%S"
    )

    ws["A2"].alignment = Alignment(horizontal="center")

    # Headers
    headers = [
        "Report Date & Time",
        "Trading Date",
        "Index",
        "Open",
        "High",
        "Low",
        "Close",
        "Volume",
        "Previous 20-Day High",
        "Previous 20-Day Low",
        "Breakout",
        "Market Comment"
    ]

    for col, heading in enumerate(headers, start=1):

        cell = ws.cell(row=4, column=col, value=heading)

        cell.font = Font(bold=True, color="FFFFFF")

        cell.fill = PatternFill(
            "solid",
            fgColor="4472C4"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    # Market data rows
    for row, item in enumerate(records, start=5):

        values = [
            run_time,
            item["date"],
            item["name"],
            item["open"],
            item["high"],
            item["low"],
            item["close"],
            item["volume"] if item["volume"] is not None else "N/A",
            item["breakout_high"],
            item["breakout_low"],
            item["breakout"],
            item["comment"]
        ]

        for col, value in enumerate(values, start=1):

            cell = ws.cell(
                row=row,
                column=col,
                value=value
            )

            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

            if col in [4, 5, 6, 7, 9, 10]:
                cell.number_format = "#,##0.00"

            if col == 1:
                cell.number_format = "dd-mmm-yyyy hh:mm:ss"

            if col == 2:
                cell.number_format = "dd-mmm-yyyy"

            if col == 8 and isinstance(value, int):
                cell.number_format = "#,##0"

        # Color market comments
        comment_cell = ws.cell(row=row, column=12)

        if "Bullish" in item["trend"]:

            comment_cell.fill = PatternFill(
                "solid",
                fgColor="C6EFCE"
            )

        elif "Bearish" in item["trend"]:

            comment_cell.fill = PatternFill(
                "solid",
                fgColor="FFC7CE"
            )

        else:

            comment_cell.fill = PatternFill(
                "solid",
                fgColor="FFF2CC"
            )

    # ======================================================
    # OHLC CHART
    # ======================================================

    chart = BarChart()

    chart.type = "col"
    chart.style = 10

    chart.title = "NIFTY 50 & BANK NIFTY OHLC"

    chart.y_axis.title = "Index Value"
    chart.x_axis.title = "Index"

    chart.height = 9
    chart.width = 19

    data = Reference(
        ws,
        min_col=4,
        max_col=7,
        min_row=4,
        max_row=4 + len(records)
    )

    categories = Reference(
        ws,
        min_col=3,
        min_row=5,
        max_row=4 + len(records)
    )

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    ws.add_chart(chart, "A9")

    # ======================================================
    # VOLUME CHART
    # ======================================================

    volume_records = [
        item for item in records
        if item["volume"] is not None
    ]

    if volume_records:

        volume_chart = BarChart()

        volume_chart.type = "col"
        volume_chart.style = 11

        volume_chart.title = "INDEX VOLUME"

        volume_chart.y_axis.title = "Volume"
        volume_chart.x_axis.title = "Index"

        volume_chart.height = 9
        volume_chart.width = 13

        # Helper data for volume chart
        ws["N4"] = "Index"
        ws["O4"] = "Volume"

        for row, item in enumerate(volume_records, start=5):

            ws.cell(row=row, column=14, value=item["name"])
            ws.cell(row=row, column=15, value=item["volume"])

        volume_data = Reference(
            ws,
            min_col=15,
            min_row=4,
            max_row=4 + len(volume_records)
        )

        volume_categories = Reference(
            ws,
            min_col=14,
            min_row=5,
            max_row=4 + len(volume_records)
        )

        volume_chart.add_data(
            volume_data,
            titles_from_data=True
        )

        volume_chart.set_categories(volume_categories)

        ws.add_chart(volume_chart, "J9")

        ws.column_dimensions["N"].hidden = True
        ws.column_dimensions["O"].hidden = True

    else:

        ws["J9"] = "Yahoo Finance volume data unavailable for these indices."

    # ======================================================
    # COLUMN WIDTHS AND FORMATTING
    # ======================================================

    widths = {
        "A": 23,
        "B": 17,
        "C": 16,
        "D": 14,
        "E": 14,
        "F": 14,
        "G": 14,
        "H": 15,
        "I": 23,
        "J": 23,
        "K": 35,
        "L": 45
    }

    for column, width in widths.items():

        ws.column_dimensions[column].width = width

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[4].height = 40

    for row in range(5, 5 + len(records)):
        ws.row_dimensions[row].height = 45

    ws.freeze_panes = "A5"

    ws.auto_filter.ref = f"A4:L{4 + len(records)}"

    thin_gray = Side(
        style="thin",
        color="D9E2F3"
    )

    for row in ws.iter_rows(
        min_row=4,
        max_row=4 + len(records),
        min_col=1,
        max_col=12
    ):

        for cell in row:

            cell.border = Border(
                bottom=thin_gray
            )

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ======================================================
    # SAVE WORKBOOK - PERMISSION ERROR FIX
    # ======================================================

    try:

        wb.save(output_file)

    except PermissionError:

        # If the same report is open in Excel,
        # create a new filename with the current time.

        timestamp = run_time.strftime("%H%M%S")

        output_file = OUTPUT_DIR / (
            f"dailymarket_report_{date_text}_{timestamp}.xlsx"
        )

        wb.save(output_file)

    return output_file


# ==========================================================
# OPEN YAHOO FINANCE
# ==========================================================

def open_yahoo_pages():

    print("\nOpening Yahoo Finance...")

    webbrowser.open(
        "https://finance.yahoo.com/quote/%5ENSEI/"
    )

    time.sleep(3)

    webbrowser.open_new_tab(
        "https://finance.yahoo.com/quote/%5ENSEBANK/"
    )

    time.sleep(3)


# ==========================================================
# OPEN EXCEL AND TAKE SCREENSHOT
# ==========================================================

def open_workbook_and_screenshot(
    excel_path,
    screenshot_path
):

    try:

        print("\nOpening Excel...")

        os.startfile(str(excel_path))

        time.sleep(15)

        # Excel should now be open
        pyautogui.hotkey("ctrl", "home")

        time.sleep(2)

        # Save screenshot
        pyautogui.screenshot().save(
            str(screenshot_path)
        )

        print("\nScreenshot saved:", screenshot_path)

    except Exception as error:

        print("\nScreenshot error:", error)


# ==========================================================
# MAIN PROGRAM
# ==========================================================

def main():

    run_time = datetime.now()

    date_text = run_time.strftime("%Y-%m-%d")

    screenshot_path = OUTPUT_DIR / (
        f"daily_market_report_{date_text}.png"
    )

    print("=" * 55)
    print("       DAILY MARKET REPORT BOT")
    print("       NIFTY 50 AND BANK NIFTY")
    print("=" * 55)

    print("\nReport date and time:", run_time)

    try:

        # Open Yahoo Finance pages
        open_yahoo_pages()

        # Fetch both indices
        records = []

        for name, ticker in TICKERS.items():

            data = fetch_market_data(name, ticker)

            records.append(data)

            print("\n", name)

            print("Trading Date:", data["date"])
            print("Open:", data["open"])
            print("High:", data["high"])
            print("Low:", data["low"])
            print("Close:", data["close"])
            print("Volume:", data["volume"])
            print("Breakout:", data["breakout"])
            print("Comment:", data["comment"])

        # Create Excel report
        excel_path = create_excel_report(
            records,
            run_time
        )

        print("\nExcel report saved:", excel_path)

        # Open Excel and take screenshot
        open_workbook_and_screenshot(
            excel_path,
            screenshot_path
        )

        print("\n" + "=" * 55)
        print("REPORT COMPLETED SUCCESSFULLY")
        print("=" * 55)

        print("\nExcel:", excel_path)
        print("Screenshot:", screenshot_path)
        print("Saved in:", OUTPUT_DIR)

    except Exception as error:

        print("\nERROR:", error)
        print("\nPlease check your internet connection and try again.")


# ==========================================================
# RUN PROGRAM
# ==========================================================

if __name__ == "__main__":

    main()
