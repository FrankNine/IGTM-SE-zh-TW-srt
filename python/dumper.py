from openpyxl import load_workbook
from datetime import timedelta
import srt


class clip:
    title: str
    start_row: int
    end_row: int
    timedelta_offset: timedelta

    def __init__(self, title: str, start_row: int, end_row: int, timedelta_offset: timedelta):
        self.title = title
        self.start_row = start_row
        self.end_row = end_row
        self.timedelta_offset = timedelta_offset

def to_ms(t: int):
    return float(t) / 60 * 100

clips : clip = [
    clip(title = "IGTM_SE_Life_After_FEZ_Phil_Fish_Epilogue.zh_TW", start_row=6, end_row=138, 
         timedelta_offset = timedelta(hours=1, minutes=0, seconds=53, milliseconds=to_ms(18)) - 
                            timedelta(seconds=11, milliseconds=263)),
    clip(title = "IGTM_SE_01_Phil_and_Japan.zh_TW", start_row=141, end_row=317, 
         timedelta_offset = timedelta(hours=1, minutes=12, seconds=36) - 
                            timedelta(seconds=26, milliseconds=401)),
    clip(title = "IGTM_SE_Life_After_Meat_Boy_Edmund_Epilogue.zh_TW", start_row=318, end_row=395, 
         timedelta_offset = timedelta(hours=1, minutes=22, seconds=26, milliseconds=to_ms(6)) - 
                            timedelta(seconds=0)),
    clip(title = "IGTM_SE_Life_After_Meat_Boy_Tommy_Epilogue.zh_TW", start_row=318, end_row=442, 
         timedelta_offset = timedelta(hours=1, minutes=27, seconds=6, milliseconds=to_ms(5)) - 
                            timedelta(seconds=5, milliseconds=298)),
    clip(title = "IGTM_SE_Cat_Lady_Danielle_Epilogue.zh_TW", start_row=444, end_row=492, 
         timedelta_offset = timedelta(hours=1, minutes=30, seconds=20, milliseconds=to_ms(14)) - 
                            timedelta(seconds=9, milliseconds=21)),
    clip(title = "IGTM_SE_02_Edmund and TEH INTERNETS.zh_TW", start_row=493, end_row=771, 
         timedelta_offset = timedelta(hours=1, minutes=33, seconds=25, milliseconds=to_ms(14)) - 
                            timedelta(milliseconds=927)),
    clip(title = "IGTM_SE_03 Tommy and TEH INTERNETS.zh_TW", start_row=774, end_row=868, 
         timedelta_offset = timedelta(hours=1, minutes=47, seconds=20, milliseconds=to_ms(3)) - 
                            timedelta(seconds=6, milliseconds=49)),
    clip(title = "IGTM_SE_04_Tommy_The Day After.zh_TW", start_row=870, end_row=927, 
         timedelta_offset = timedelta(hours=1, minutes=52, seconds=40, milliseconds=to_ms(1)) - 
                            timedelta(seconds=6, milliseconds=34)),
    clip(title = "IGTM_SE_05_Super Meat Boy_Collections.zh_TW", start_row=929, end_row=989, 
         timedelta_offset = timedelta(hours=1, minutes=56, seconds=41, milliseconds=to_ms(16)) - 
                            timedelta(seconds=16, milliseconds=280)),
    clip(title = "IGTM_SE_06_The Art of Braid.zh_TW", start_row=990, end_row=1071, 
         timedelta_offset = timedelta(hours=2, minutes=0, seconds=25, milliseconds=to_ms(17)) - 
                            timedelta(seconds=1, milliseconds=418)),
    clip(title = "IGTM_SE_08_Passage.zh_TW", start_row=1072, end_row=1192, 
         timedelta_offset = timedelta(hours=2, minutes=5, seconds=6, milliseconds=to_ms(22)) - 
                            timedelta(seconds=4, milliseconds=78)),
    clip(title = "IGTM_SE_07_Spelunky.zh_TW", start_row=1194, end_row=1310, 
         timedelta_offset = timedelta(hours=2, minutes=12, seconds=16, milliseconds=to_ms(7)) - 
                            timedelta(seconds=5, milliseconds=969)),
    clip(title = "IGTM_SE_09_David and the Clouds.zh_TW", start_row=1313, end_row=1386, 
         timedelta_offset = timedelta(hours=2, minutes=19, seconds=25, milliseconds=to_ms(10)) - 
                            timedelta(seconds=6, milliseconds=424)),
    clip(title = "IGTM_SE_10_Eliss.zh_TW", start_row=1389, end_row=1503, 
         timedelta_offset = timedelta(hours=2, minutes=23, seconds=55, milliseconds=to_ms(16)) - 
                            timedelta(seconds=7, milliseconds=665)),
    clip(title = "IGTM_SE_12_Super Meat Boy_Control.zh_TW", start_row=1506, end_row=1563, 
         timedelta_offset = timedelta(hours=2, minutes=30, seconds=53, milliseconds=to_ms(17)) - 
                            timedelta(seconds=4, milliseconds=340)),
    clip(title = "IGTM_SE_13_Game Jam.zh_TW", start_row=1566, end_row=1628, 
         timedelta_offset = timedelta(hours=2, minutes=34, seconds=20, milliseconds=to_ms(19)) - 
                            timedelta(seconds=8, milliseconds=259)),
    clip(title = "IGTM_SE_19_CANABALT.zh_TW", start_row=1630, end_row=1691, 
         timedelta_offset = timedelta(hours=2, minutes=38, seconds=18, milliseconds=to_ms(14)) - 
                            timedelta(seconds=4, milliseconds=329)),
    clip(title = "IGTM_SE_17_Coil.zh_TW", start_row=1693, end_row=1736, 
         timedelta_offset = timedelta(hours=2, minutes=42, seconds=13, milliseconds=to_ms(13)) - 
                            timedelta(seconds=3, milliseconds=631)),
    clip(title = "IGTM_SE_15_Tri-achnid.zh_TW", start_row=1737, end_row=1784, 
         timedelta_offset = timedelta(hours=2, minutes=44, seconds=42, milliseconds=to_ms(22)) - 
                            timedelta(seconds=0, milliseconds=543)),
    clip(title = "IGTM_SE_16_AVGM.zh_TW", start_row=1787, end_row=1873, 
         timedelta_offset = timedelta(hours=2, minutes=47, seconds=42, milliseconds=to_ms(23)) - 
                            timedelta(seconds=4, milliseconds=399)),
    clip(title = "IGTM_SE_18_The C Word.zh_TW", start_row=1875, end_row=1959, 
         timedelta_offset = timedelta(hours=2, minutes=52, seconds=52, milliseconds=to_ms(6)) - 
                            timedelta(seconds=1, milliseconds=425)),
    clip(title = "IGTM_SE_14_Phil Watching Phil.zh_TW", start_row=1962, end_row=1991, 
         timedelta_offset = timedelta(hours=2, minutes=58, seconds=26, milliseconds=to_ms(4)) - 
                            timedelta(seconds=12, milliseconds=95)),
    clip(title = "IGTM_SE_11_MEGA64.zh_TW", start_row=1992, end_row=2027, 
         timedelta_offset = timedelta(hours=3, minutes=0, seconds=42, milliseconds=to_ms(0)) - 
                            timedelta(seconds=0, milliseconds=833))
]


def to_timedelta(time_string: str) -> timedelta:
    parts = time_string.split(':')
    return timedelta(hours=int(parts[0]), minutes=int(parts[1]), seconds=int(parts[2]), milliseconds=int(parts[3]))

def row_to_subtitle(row_index: int, subtitle_index: int) -> srt.Subtitle:
    start = to_timedelta(sheet["A{}".format(row_index)].value)
    end = to_timedelta(sheet["B{}".format(row_index)].value)

    raw_content = sheet["C{}".format(row_index)].value
    if raw_content == None:
        content = ""
    else:
        content = str(raw_content)

    return srt.Subtitle(index = subtitle_index, start =start, end = end, content = content)


workbook = load_workbook(filename = 'TRAD_CHINESE_TRANSLATION_SPECIAL EDITION.xlsx', data_only=True)
sheet = workbook['cht']

for c in clips:
    subtitle_lines = []
    subtitle_index = 1
    for i in range(c.start_row, c.end_row + 1):
        subtitle = row_to_subtitle(i, subtitle_index)

        if len(subtitle.content) == 0:
            continue

        subtitle.start -= c.timedelta_offset
        subtitle.end -= c.timedelta_offset

        subtitle.content = subtitle.content.replace('|', '\n')

        subtitle_lines.append(subtitle)
        subtitle_index = subtitle_index + 1

    srt_content = srt.compose(subtitle_lines)
    f = open("../" + c.title + ".srt", "w", encoding='UTF-8')
    f.write(srt_content)
    f.close()
