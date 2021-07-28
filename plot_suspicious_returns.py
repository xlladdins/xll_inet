"""
This code makes Figure 1 of "Strikingly Suspicious Overnight and Intraday Returns" (2020).
        (https://ssrn.com/abstract=3705017, https://arxiv.org/abs/2010.01727)

Appropriately modified, this code also makes:
  + Figures 2 and 3 of "Celebrating Three Decades of Worldwide Stock Market Manipulation" (2019)
        (https://ssrn.com/abstract=3490879, https://arxiv.org/abs/1912.01708)
  + Figure 1 of "How to Increase Global Wealth Inequality for Fun and Profit" (2018)
        (https://ssrn.com/abstract=3282845, https://arxiv.org/abs/1811.04994)

We assume the user has a python environment with numpy and matplotlib.

Bruce Knuteson (knuteson@mit.edu)
2020-10-04

Updates:
2021-01-02
  + For the SSE Composite, use the symbol "000001.SS" instead of "^SSEC" (which Yahoo Finance no longer provides).
  + Plot data through the end of 2020.
2021-04-17
  + Handle several changes to the Yahoo Finance data since my 2021-01-02 update:
    - Bad AEX price on 1995-12-26.
    - DNB OBX open prices are zero from 2009-01-01 to 2009-05-06.
    - ASX 200 (STW.AX) data are now available only from 2008-01-01 onward (rather than 2001-08-27 onward), and
      Straits Times (ES3.SI) open prices are now exactly equal to close prices (i.e., useless) for all of 2008.
      To use the ASX 200 and Straits Times data that were available on Yahoo Finance on 2021-01-02,
      set PATCH_LONGER_HISTORY = {"STW.AX", "ES3.SI"}
  + Plot data through 2021-03-31.
"""
import datetime
import os
import re
import requests
import urllib
import time

from numpy import asarray, cumprod, ndarray
from matplotlib.dates import date2num, YearLocator
from matplotlib.pyplot import xlim, subplots, figure, gca, gcf, savefig, clf

DATETIME_FORMAT = "%Y-%m-%d"
DEFAULT_START_DATE = datetime.datetime(1990, 1, 1)
DEFAULT_END_DATE = datetime.datetime(2021, 5, 31)
# Adjust input and output directories as necessary
INPUT_DATA_DIR = "."
OUTPUT_PLOT_DIR = "."
# To use the ASX 200 (STW.AX) and Straits Times (ES3.SI) data that were available on Yahoo Finance on 2021-01-02,
# set PATCH_LONGER_HISTORY = {"STW.AX", "ES3.SI"}
PATCH_LONGER_HISTORY = None  # {"STW.AX", "ES3.SI"}


def symbol_details_dict():
    """
    Return a dict with a few pieces of useful information for the indices we consider.

    This information includes (where appropriate):
      + a formal(ish) and/or common name for the index (or ETF);
      + the index's home country; and
      + a start date, end date, and/or bad_data_dates (if manual analysis has identified specific problems).

    @rtype: dict[str, dict[str]]
    """
    return dict([
        # Indices shown in Figure 1 of "How to Increase Global Wealth Inequality for Fun and Profit" (2018)
        ("SPY", dict(name="S&P 500 SPDR ETF", country="United States", short_name="S&P 500")),
        ("^IXIC", dict(name="NASDAQ Composite", country="United States", short_name="NASDAQ")),
        ("XIU.TO", dict(name="iShares TSX 60 ETF", country="Canada", short_name="TSX 60")),
        ("^FCHI", dict(name="CAC 40", country="France")),
        ("^GDAXI", dict(name="DAX", country="Germany")),
        ("^N225", dict(name="Nikkei 225", country="Japan")),

        # Indices added in Figures 2 and 3 of "Celebrating Three Decades of Worldwide Stock Market Manipulation" (2019)
        ("ISFU.L", dict(name="iShares Core FTSE 100 ETF", country="United Kingdom", short_name="FTSE 100",
                        end_date=datetime.datetime(2018, 6, 20),  # last date data are available
                        bad_data_dates=[datetime.datetime(2001, 7, 16)]  # bad open price
                        )),
        ("^AEX", dict(name="AEX", country="Netherlands",
                      bad_data_dates=[datetime.datetime(1995, 12, 26)]  # bad prices
                      )),
        ("OBXEDNB.OL", dict(name="DNB OBX ETF", country="Norway", short_name="DNB OBX",
                            # open prices are zero from 2009-01-01 to 2009-05-06
                            start_date=datetime.datetime(2009, 5, 8),
                            bad_data_dates=[datetime.datetime(2009, 11, 12),  # zero open price
                                            datetime.datetime(2009, 11, 13),  # zero open price
                                            datetime.datetime(2013, 1, 4)  # bad close price
                                            ])),
        ("^TA125.TA", dict(name="TA-125", country="Israel",
                           # more than half of the TA-125 overnight returns are zero before 2007.
                           start_date=datetime.datetime(2007, 1, 8))),
        ("STW.AX", dict(name="SPDR S&P/ASX 200 Fund", country="Australia", short_name="ASX 200")),
        ("^NSEI", dict(name="NIFTY 50", country="India")),
        ("^BSESN", dict(name="S&P BSE SENSEX", country="India", short_name="SENSEX")),
        ("^HSI", dict(name="Hang Seng Index", country="Hong Kong", short_name="Hang Seng")),
        ("ES3.SI", dict(name="SPDR Straits Times Index ETF", country="Singapore", short_name="Straits Times")),
        ("000001.SS", dict(name="SSE Composite Index", country="China", short_name="SSE")),

        # Indices added in Figure 1 of "Strikingly Suspicious Overnight and Intraday Returns" (2020)
        ("IMIB.MI", dict(name="iShares FTSE MIB ETF", country="Italy", short_name="FTSE MIB")),
        ("^KS11", dict(name="KOSPI Composite Index", country="Korea", short_name="KOSPI")),
        ("^TWII", dict(name="TSEC Weighted Index", country="Taiwan", short_name="TSEC")),
        ("EWW", dict(name="iShares MSCI Mexico Capped ETF", country="Mexico", short_name="MSCI Mexico")),
        ("EWZ", dict(name="iShares MSCI Brazil Capped ETF", country="Brazil", short_name="MSCI Brazil")),

        # Largest US companies
        # (as referenced in footnote 18 of "Strikingly Suspicious Overnight and Intraday Returns")
        ("AAPL", dict(name="Apple")),
        ("MSFT", dict(name="Microsoft")),
        ("AMZN", dict(name="Amazon")),
        ("GOOG", dict(name="Google")),
        ("FB", dict(name="Facebook")),
        ("TSLA", dict(name="Tesla")),
        ("BRK-B", dict(name="Berkshire Hathaway")),
        ("V", dict(name="Visa")),
        ("JPM", dict(name="JPMorgan Chase")),
        ("JNJ", dict(name="Johnson & Johnson")),

        # Small US companies
        ("IWO", dict(name="iShares Russell 2000 Growth ETF")),
        ("VBR", dict(name="Vanguard Small Cap Value ETF")),

        # Chinese tech stocks (trading on US exchanges)
        ("BABA", dict(name="Alibaba")),
        ("TCEHY", dict(name="Tencent")),
        ("BIDU", dict(name="Baidu")),
        ("NTES", dict(name="NetEase")),
        ("JD", dict(name="JD.com")),
        ("BILI", dict(name="Bilibili")),

        # Meme stocks
        ("GME", dict(name="GameStop")),
        ("AMC", dict(name="AMC Entertainment"))
    ])


##############################################################################################################
# Trivial helper routines
##############################################################################################################

def format_cumulative_return_as_string(r_pct):
    """
    Turn the float r_pct into a string suitable for display.

    @type r_pct: float
    @rtype: str
    """
    return ("%+.2f%%" % r_pct if r_pct < -99 else  # show "-99.88%" (rather than "-100%")
            "%+.1f%%" % r_pct if -1 < r_pct < 1 else  # show "+0.2%" (because "+0%" looks like a mistake)
            "%+.0f%%" % r_pct if r_pct < 2e4 else  # show "+3333%"
            "+" + "{:,.0f}".format(r_pct) + "%")   # show "+333,333%"


def get_data_filename(sym):
    """
    Get the name of the local file with open and close prices for sym.

    @type sym: str
    @rtype: str
    """
    return os.path.join(INPUT_DATA_DIR, sym + ".csv")


##############################################################################################################
# Process data
##############################################################################################################

def download_data_from_yahoo_finance(sym):
    """
    Download historical open/close data for sym from Yahoo! Finance.

    Example: https://query1.finance.yahoo.com/v7/finance/download/SPY? ...
                ... period1=631170000&period2=1599022800&interval=1d&events=history

    @type sym: str
    @rtype: None
    """
    d1 = DEFAULT_START_DATE
    d2 = (symbol_details_dict().get(sym).get("end_date") or DEFAULT_END_DATE) + datetime.timedelta(days=1)
    url_params = dict(interval="1d", events="history",
                      period1=int(time.mktime((d1.year, d1.month, d1.day, 0, 0, 0, 0, 0, 0))),
                      period2=int(time.mktime((d2.year, d2.month, d2.day, 0, 0, 0, 0, 0, 0))))
    my_url = ("https://query1.finance.yahoo.com/v7/finance/download/%s?" % urllib.quote(sym)
              ) + urllib.urlencode(url_params)
    r = requests.get(my_url)
    file(get_data_filename(sym), "w").write(r.content)
    # If we will be downloading multiple files from Yahoo! Finance, let's not be obnoxious about it.
    time.sleep(5)


def download_data_from_archive(sym, archive_date):
    """
    Download historical Yahoo! Finance open/close data for sym from our archive.

    @type sym: str
    @param archive_date: yyyymmdd  (e.g., "20201231")
    @type archive_date: str
    @rtype: None
    """
    my_url = "https://bruceknuteson.github.io/spy-day-and-night/open_close_price_data/%s/%s.csv" % (archive_date, sym)
    r = requests.get(my_url)
    file(get_data_filename(sym), "w").write(r.content)


def get_historical_open_close_data(sym):
    """
    Return a single string with the contents of Yahoo! Finance's historical open/close price data for sym.

    Download the data if necessary.

    @type sym: str
    @rtype: list[str]
    """
    data_filename = get_data_filename(sym)
    if not os.path.exists(data_filename):
        download_data_from_yahoo_finance(sym)
        if sym == "ISFU.L":
            # As of 2019-10-31, Yahoo! Finance provided sensible data for ISFU.L from 2001-01-02 to 2018-06-20.
            # Currently (as of 2021-04-17), Yahoo! Finance does not provide sensible historical data for any of
            # ISFU.L, ISF.L, or ^FTSE.  To get reasonable FTSE 100 data,
            # we download Yahoo! Finance's data for ISFU.L as of 2019-10-31 from our archive.
            download_data_from_archive(sym, "20191031")
        if sym in (PATCH_LONGER_HISTORY or set()).intersection("STW.AX ES3.SI".split()):
            # Yahoo! Finance provided a longer history for STW.AX and ES3.SI on 2020-12-31 than it does now
            # (on 2021-04-17).  If desired, download a longer history for STW.AX and/or ES3.SI from our archive.
            download_data_from_archive(sym, "20201231")

    data = open(data_filename).read().split("\n")
    assert data[0].strip() == "Date,Open,High,Low,Close,Adj Close,Volume", data[0]  # double-check the header
    return data


def get_prices_open_close_adj_dates(original_data, start_date=None, end_date=None, bad_data_dates=None):
    """
    Get (price_open, price_close, price_close_adj, dates_datetime) from the open and close prices in original_data.

    @type original_data: list[str]

    @param start_date: provided if data before start_date are suspect
    @type start_date: datetime.datetime

    @param end_date: provided if data after end_date are suspect (or unavailable)
    @type end_date: datetime.datetime

    @param bad_data_dates: provided if data for specific dates are suspect
    @type bad_data_dates: list[datetime.datetime]

    @return: (price_open, price_close, price_close_adj, dates_datetime)
    @rtype: (ndarray, ndarray, ndarray, list[datetime.datetime])
    """
    # Make sure the header is what we expect.
    assert original_data[0].strip() == "Date,Open,High,Low,Close,Adj Close,Volume", original_data[0]

    # Remove any known bad dates.
    bad_data_dates_str = set(d.strftime(DATETIME_FORMAT) for d in (bad_data_dates or []))
    data = [d.split(',') for d in original_data[1:]
            if d and "null" not in d and d.split(',')[0] not in bad_data_dates_str]

    # Also discard any dates with open = high = low = close.
    data = [d for d in data if len(set(d[1:5])) > 1]

    price_open = asarray([float(d[1]) for d in data])  # Open
    price_close = asarray([float(d[4]) for d in data])  # Close
    price_close_adj = asarray([float(d[5]) for d in data])  # Adj Close
    dates_datetime = [datetime.datetime(*map(int, d[0].split('-'))) for d in data]  # Date

    # In some cases (like the DAX before 1993-12-14), price_open == price_close because Yahoo! Finance does not have
    # opening prices.  If we are dealing with data like the DAX starting on 1993-01-01, recognize this, and
    # only return data from 1993-12-14 onward.
    # Separately, only consider dates from start_date onward.
    d0 = max((price_open != price_close).nonzero()[0][0],
             [d >= (start_date or DEFAULT_START_DATE) for d in dates_datetime].index(True))
    # Only consider dates up to end_date.
    d1 = ([d > (end_date or DEFAULT_END_DATE) for d in dates_datetime].index(True)
          if (end_date or DEFAULT_END_DATE) < max(dates_datetime) else len(dates_datetime))
    if d0 > 1 or d1 < len(dates_datetime):
        price_open = price_open[d0:d1]
        price_close = price_close[d0:d1]
        price_close_adj = price_close_adj[d0:d1]
        dates_datetime = dates_datetime[d0:d1]

    return price_open, price_close, price_close_adj, dates_datetime


def compute_returns_overnight_intraday(price_open, price_close, price_close_adj, dates_datetime):
    """
    Compute overnight and intraday returns from open and close prices.

    @type price_open: ndarray
    @type price_close: ndarray
    @type price_close_adj: ndarray
    @type dates_datetime: list[datetime.datetime]

    @return: (returns_overnight, returns_intraday)
    @rtype: (ndarray, ndarray)
    """
    assert len(price_open) == len(price_close) == len(price_close_adj) == len(dates_datetime)

    # Intraday returns are the returns from open to close.
    # returns_intraday[0] is the return from open on dates_datetime[0] to close on dates_datetime[0]
    returns_intraday = (price_close / price_open) - 1

    # Use adjusted prices to get close to close returns.
    # returns_close_to_close[0] is the return from close on dates_datetime[0] to close on dates_datetime[1]
    returns_close_to_close = (price_close_adj[1:] / price_close_adj[:-1]) - 1

    # Overnight returns are close to close returns sans intraday returns.
    # returns_overnight[0] is the return from close on dates_datetime[0] to open on dates_datetime[1]
    returns_overnight = ((1 + returns_close_to_close) / (1 + returns_intraday[1:])) - 1

    return returns_overnight, returns_intraday


def get_plot_data(sym):
    """
    Extract the returns we want to plot from historical open and close prices.

    @type sym: str

    @return: dates, overnight returns, and intraday returns
    @rtype: (list[datetime.datetime], ndarray, ndarray)
    """
    data = get_historical_open_close_data(sym)

    s = symbol_details_dict().get(sym)
    (price_open, price_close, price_close_adj, dates_datetime) = \
        get_prices_open_close_adj_dates(data, s.get("start_date"), s.get("end_date"), s.get("bad_data_dates"))

    returns_overnight, returns_intraday = compute_returns_overnight_intraday(price_open, price_close, price_close_adj,
                                                                             dates_datetime)

    return dates_datetime, returns_overnight, returns_intraday


##############################################################################################################
# Make plots
##############################################################################################################

def make_one_intraday_overnight_return_plot(plot_data, ax):
    """
    Draw a plot of cumulative returns.

    @param plot_data: (dates_datetime, returns_overnight, returns_intraday)
    @type plot_data: (list[datetime.datetime], ndarray, ndarray)
    @type ax: matplotlib.axes._subplots.AxesSubplot
    """
    dates_datetime, returns_overnight, returns_intraday = plot_data

    # Draw the lines
    ax.margins(x=0, y=0)
    dates_datenum = map(date2num, dates_datetime)
    ax.plot_date(dates_datenum[1:], (cumprod(returns_overnight + 1) - 1) * 100, fmt='-b', linewidth=1.5)
    ax.plot_date(dates_datenum, (cumprod(returns_intraday + 1) - 1) * 100, fmt='-g', linewidth=1.5)

    # Add yticks on the right edge of the plot.
    ytick_right_x = xlim()[0] + (xlim()[-1] - xlim()[0]) * 1.005
    todays_value_overnight = ((cumprod(returns_overnight + 1) - 1) * 100)[-1]
    todays_value_intraday = ((cumprod(returns_intraday + 1) - 1) * 100)[-1]
    ax.text(ytick_right_x, todays_value_overnight, format_cumulative_return_as_string(todays_value_overnight),
            verticalalignment="center")
    ax.text(ytick_right_x, todays_value_intraday,
            (" " if todays_value_intraday < 0 else "") + format_cumulative_return_as_string(todays_value_intraday),
            verticalalignment="center")

    # Set yticks on the left side of the plot.
    ax.set_yticks([0])
    ax.set_yticklabels([0])
    ax.set_ylim(-100)


def plot_all_suspicious_index_returns_one_plot():
    """
    Make Figure 1 of Strikingly Suspicious Overnight and Intraday Return Patterns.
    """
    symbols_to_plot = ("SPY        ^IXIC       XIU.TO "
                       "STW.AX     EWZ         EWW "
                       "ISFU.L     ^FCHI       ^GDAXI "
                       "^AEX       OBXEDNB.OL  IMIB.MI "
                       "^TA125.TA  ^NSEI       ^BSESN "
                       "ES3.SI     ^KS11       ^TWII "
                       "^N225      ^HSI        000001.SS").split()

    # Draw twenty-one plots.
    clf()
    n_rows, n_cols = 7, 3
    fig, axes = subplots(num=1, nrows=n_rows, ncols=n_cols, sharex=True)
    every_five_years = YearLocator(5)
    for i_s, sym in enumerate(symbols_to_plot):
        sym_details = symbol_details_dict().get(sym)
        ax = axes[(i_s // n_cols), (i_s % n_cols)]
        plot_data = get_plot_data(sym)
        make_one_intraday_overnight_return_plot(plot_data, ax)
        # ticks on the x axis every 5 years
        ax.xaxis.set_major_locator(every_five_years)
        # label the plot
        ax.text(0.02, 0.98, sym_details.get("country") + "\n" + sym_details.get("name"),
                transform=ax.transAxes, horizontalalignment="left", verticalalignment="top", fontsize="large")

    # Make sure the horizontal axis goes from DEFAULT_START_DATE to DEFAULT_END_DATE.
    axes[-1, -1].plot_date([date2num(DEFAULT_START_DATE), date2num(DEFAULT_END_DATE)], [0, 0], fmt='w', alpha=0)
    # Add figure title and legend.
    axes[0, 0].legend(("overnight", "intraday"), loc='upper left', bbox_to_anchor=(0.00, 0.80))
    fig.text(0.52, 0.90, "Overnight and Intraday Returns to Major Stock Market Indices",
             horizontalalignment="center", verticalalignment="center", transform=fig.transFigure, fontsize="x-large")
    fig.set_size_inches(15.32, 19, forward=True)
    if OUTPUT_PLOT_DIR:
        desired_plot_formats = set("jpg png pdf".split())
        supported_plot_formats = set(gcf().canvas.get_supported_filetypes())
        for plot_format in desired_plot_formats.intersection(supported_plot_formats):
            savefig(os.path.join(OUTPUT_PLOT_DIR, "suspicious_index_returns.%s" % plot_format),
                    bbox_inches="tight", dpi=300)
        savefig(os.path.join(OUTPUT_PLOT_DIR, "suspicious_index_returns_144dpi.png"), bbox_inches="tight", dpi=144)


def make_one_intraday_overnight_return_plot_standalone(sym):
    """
    Draw a plot of cumulative returns for the symbol sym.

    @type sym: str
    """
    # Get the data.
    plot_data = get_plot_data(sym)

    # Make the plot.
    fig = figure(1)
    ax = gca()
    make_one_intraday_overnight_return_plot(plot_data, ax)

    # Add finishing touches.
    ax.legend(("overnight", "intraday"), loc='upper left', bbox_to_anchor=(0.00, 1.00), fontsize="large")
    sym_details = symbol_details_dict().get(sym)
    fig.text(0.52, 0.90, (sym_details.get("name") + " Overnight and Intraday Returns"),
             horizontalalignment="center", verticalalignment="bottom", transform=fig.transFigure, fontsize="x-large")
    fig.set_size_inches(8.0, 4.2, forward=True)
    if OUTPUT_PLOT_DIR:
        s = (sym_details.get("short_name") or sym_details.get("name")).lower()
        if len(s) > 20:
            s = sym.lower()
        s = re.sub("[ &^]", "", s)
        savefig(os.path.join(OUTPUT_PLOT_DIR, "suspicious_returns_%s.pdf" % s), bbox_inches="tight")


def make_all_standalone_intraday_overnight_return_plots():
    """
    Make a standalone overnight vs. intraday plot for the S&P 500 index.  Do the same for the NASDAQ, TSX 60, etc.
    """
    symbols_to_plot = ("SPY        ^IXIC       XIU.TO "
                       "STW.AX     EWZ         EWW "
                       "ISFU.L     ^FCHI       ^GDAXI "
                       "^AEX       OBXEDNB.OL  IMIB.MI "
                       "^TA125.TA  ^NSEI       ^BSESN "
                       "ES3.SI     ^KS11       ^TWII "
                       "^N225      ^HSI        000001.SS "
                       "IWO        VBR").split()
    for sym in symbols_to_plot:
        clf()
        make_one_intraday_overnight_return_plot_standalone(sym)


if __name__ == "__main__":
    plot_all_suspicious_index_returns_one_plot()
