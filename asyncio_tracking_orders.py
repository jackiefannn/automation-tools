import asyncio
import datetime
import sys
import time
from datetime import date
import aiohttp
from aiohttp import ClientSession
import logging
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

url = os.environ['TRACKING_API']
database = os.environ['EXCEL_SHEET']
# get handle on existing file
wb = load_workbook(filename=os.environ['EXCEL_SHEET'])
# get current month and year worksheet
month_year = date.today().strftime("%b-%Y").split("-")
worksheet = "{month}. {year}".format(month=month_year[0], year=month_year[1])
ws = wb[worksheet]
font = Font(name='Times New Roman', size=12)
alignment = Alignment(horizontal='center')
row = 2
sale_order_col = 'A'
shipping_status_col = 'N'
time_col = 'O'
tracking_number_col = 'P'

def trackingIdDecryption(trackingNumber):
    trackingIdEncryption = {'A': 'S', 'B': 'T', 'C': 'U', 'D': 'V', 'E': 'W', 'F': 'X', 'G': 'Y', 'H': 'Z',
                            'I': '%5B',
                            'J': '%5C', 'K': '%5D', 'L': '%5E', 'M': '_', 'N': '%60', 'O': 'a', 'P': 'b', 'Q': 'c',
                            'R': 'd', 'S': 'e', 'T': 'f', 'U': 'g', 'V': 'h', 'W': 'i', 'X': 'j', 'Y': 'k',
                            'Z': 'l',
                            '0': 'B', '1': 'C', '2': 'D', '3': 'E', '4': 'F', '5': 'G', '6': 'H', '7': 'I',
                            '8': 'J', '9': 'K'}
    return "".join([trackingIdEncryption[letter] for letter in trackingNumber])

logging.basicConfig(
    format="%(asctime)s %(levelname)s:%(name)s: %(message)s",
    level=logging.DEBUG,
    datefmt="%H:%M:%S",
    stream=sys.stderr,
)
logger = logging.getLogger("areq")
logging.getLogger("chardet.charsetprober").disabled = True


async def fetch_html(url: str, session: ClientSession, data, **kwargs) -> str:
    """
    GET request wrapper to fetch page HTML.
    kwargs are passed to `session.request()`.
    """

    resp = await session.request(method="POST", url=url, data=data, **kwargs)
    resp.raise_for_status()
    logger.info("Got response [%s] for URL: %s", resp.status, url)
    doc = await resp.json()
    logger.info("JSON Response of request: %s", doc)
    return doc


async def fetch_page(url: str, session: ClientSession, data, **kwargs):
    """Ensure that request was able to fetch JSON response for `url`."""
    try:
        doc = await fetch_html(url=url, session=session, data=data, **kwargs)
    except (aiohttp.ClientError, aiohttp.http_exceptions.HttpProcessingError) as e:
        logger.error("aiohttp exception for %s [%s]: %s", url, getattr(e, "status", None), getattr(e, "message", None))
        return None
    except Exception as e:
        logger.exception("Non-aiohttp exception occured:  %s", getattr(e, "__dict__", {}))
        return None
    else:
        return doc


async def fetch_and_write_one(url: str, database, session, data, row, **kwargs):
    """ Fetch JSON doc and return status of shipping order """
    doc = await fetch_page(url=url, session=session, data=data, **kwargs)
    if not doc: return None
    # if len(doc) == 0:
    #     return None
    # logger.info(f"Contents of the JSON response: {doc}")
    status = str
    if 'error' in doc:
        logger.info('Error: ' + doc['error'].title())
        status = f"Error: {doc['error'].title()}"
    elif 'sub_status' in doc:
        logger.info('Sub-status: ' + doc['sub_status'].title())
        status = doc['sub_status'].title()
    else:
        logger.info('Status: ' + doc['status'].title())
        status = doc['status'].title()

    db_res = await as_insert_one(database=database, status=status, row=row)

    return db_res


async def fetch_and_write(url: str, database, **kwargs) -> None:
    """Fetch & write concurrently to `file` for multiple post requests to 'url'."""
    global row
    async with ClientSession() as session:
        post_tasks = []
        while ws[f"{sale_order_col}{row}"].value:
            # track sales orders that have a shipping number
            if isinstance(ws[f"{tracking_number_col}{row}"].value, str):
                # track only orders that have not been delivered yet
                if ws[f"{tracking_number_col}{row}"].value != 'Delivered' or \
                        ws[f"{tracking_number_col}{row}"].value != 'Arrived' or \
                        ws[f"{tracking_number_col}{row}"].value !='Pickup':
                    trackingNum = ws[f"{tracking_number_col}{row}"].value
                    logger.info(f"Row: {row} & Tracking num: {trackingNum}")
                    trackingId = trackingIdDecryption(trackingNum)
                    data = {
                        'trackingId': trackingId,
                        'carrier': 'Auto-Detect',
                        'language': 'en',
                        'country': 'Russian Federation',
                        'platform': 'web-desktop',
                        'wd': 'false',
                        'c': 'false',
                        'p': 3,
                        'l': 2,
                        'se': '1792x1024,MacIntel,Gecko,Mozilla,Netscape,Google Inc.,4g,Intel Inc.,Intel(R) UHD Graphics 630,undefined,103,3584,2048'}

                    res = fetch_and_write_one(url=url, database=database, session=session, data=data, row=row, **kwargs)
                    # if not res: continue
                    post_tasks.append(res)
            row += 1
        # now execute them all at once
        await asyncio.gather(*post_tasks)


async def as_insert_one(database, status, row):
    ws[f"{shipping_status_col}{row}"] = status
    ws[f"{time_col}{row}"] = datetime.datetime.now()
    ws[f"{shipping_status_col}{row}"] .font = font
    ws[f"{shipping_status_col}{row}"] .alignment = alignment
    ws[f"{time_col}{row}"] .font = font
    ws[f"{time_col}{row}"] .alignment = alignment
    try:
        result = wb.save(filename=database)
        # result = await wb.save(filename=database)
    except TypeError as e:
        logger.error("TypeError exception occured while attempting to save Excel workbook: %s", e)
    logger.info(f"Wrote {status} into row {row}")



if __name__ == '__main__':
    start = time.time()
    # total time taken
    asyncio.run(fetch_and_write(url=url, database=database))
    end = time.time()
    logger.info(f"Runtime of the program is {end - start}")
    # asyncio.get_event_loop().run_until_complete(fetch_and_write(url=url, database=database))
