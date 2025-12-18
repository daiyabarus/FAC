"""Application settings and constants"""

from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
CONFIG_DIR = BASE_DIR / "config"
ASSETS_DIR = BASE_DIR / "assets"

TEMPLATE_FILE = "datatemplate.xlsx"
TEMPLATE_SHEET = "FAC"


class GSMColumns:
    BEGIN_TIME = 0
    END_TIME = 1
    GRANULARITY = 2
    GROUP = 3
    SUBNET_ID = 4
    SUBNET_NAME = 5
    ME_ID = 6
    ME_NAME = 7
    SITE_ID = 8
    SITE_NAME = 9
    BTS_ID = 10
    BTS_NAME = 11
    FREQ_BAND = 12
    CSSR_NUM = 13
    CSSR_DEN = 14
    SDCCH_SR_NUM = 15
    SDCCH_SR_DEN = 16
    DROP_NUM = 17
    DROP_DEN = 18


class LTEColumns:
    BEGIN_TIME = 0
    END_TIME = 1
    GRANULARITY = 2
    GROUP = 3
    SUBNET_ID = 4
    SUBNET_NAME = 5
    ME_ID = 6
    ME_NAME = 7
    ENB_ID = 8
    ENB_NAME = 9
    LTE_ID = 10
    LTE_NAME = 11
    CELL_ID = 12
    CELL_NAME = 13
    FREQ_BAND = 14
    MOC = 15
    ENODEB_ID = 16
    LOCAL_CELL_ID = 17
    PRODUCT = 18
    RRC_SSR_NUM = 19
    RRC_SSR_DEN = 20
    ERAB_SSR_NUM = 21
    ERAB_SSR_DEN = 22
    S1_SSR_NUM = 23
    S1_SSR_DEN = 24
    RACH_SETUP_NUM = 25
    RACH_SETUP_DEN = 26
    HO_SR_NUM = 27
    HO_SR_DEN = 28
    ERAB_DROP_NUM = 29
    ERAB_DROP_DEN = 30
    DL_THP_NUM = 31
    DL_THP_DEN = 32
    UL_THP_NUM = 33
    UL_THP_DEN = 34
    UL_PLOSS = 35
    DL_PLOSS = 36
    CQI_NUM = 37
    CQI_DEN = 38
    RANK_GT2_NUM = 39
    RANK_GT2_DEN = 40
    RSSI_PUSCH_NUM = 41
    RSSI_PUSCH_DEN = 42
    RAN_LAT_NUM = 43
    RAN_LAT_DEN = 44
    DL_VOL_HIGH = 45
    DL_VOL_LOW = 46
    DL_TIME = 47
    OVERLAP_RATE = 48
    DL_SE_NUM = 49
    DL_SE_DEN = 50
    VOLTE_CSSR_NUM = 51
    VOLTE_CSSR_DEN = 52
    VOLTE_DROP_NUM = 53
    VOLTE_DROP_DEN = 54
    SRVCC_SR_NUM = 55
    SRVCC_SR_DEN = 56
    CELL_AVAIL = 57
    DL_QPSK_NUM = 58
    DL_QPSK_DEN = 59
    LTC_NON_CAP = 60


class ClusterColumns:
    CLUSTER = 0
    TOWERID = 1
    LTE_CELL = 2
    TX = 3
    SITENAME = 4
    CAT = 5  # ✅ ADDED


class NGIColumns:
    """NGI (NVE Grid) file columns - RAW FORMAT"""
    ENODEB_ID = 0
    CELL_ID = 1
    CELL_NAME = 2
    TOTAL_SAMPLES = 3
    RSRP = 4
    RSRQ = 5
    GOOD_RATIO = 6
    PBAD_QBAD = 7
    PBAD_QBAD_PCT = 8
    PGOOD_QBAD = 9
    PGOOD_QBAD_PCT = 10
    PBAD_QGOOD = 11
    PBAD_QGOOD_PCT = 12
    PGOOD_QGOOD = 13
    PGOOD_QGOOD_PCT = 14


# Formatting
RED_FILL = "FFFFC7CE"
DARK_RED_TEXT = "FF9C0006"
GREEN_FILL = "FFC6EFCE"
DARK_GREEN_TEXT = "FF006100"
HEADER_FILL = "47402D"
HEADER_FONT_COLOR = "FFFEFB"
