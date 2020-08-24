"""
Extracts the Sony (probably the other brabdsm too) video files metadata to the SRT subtitles

Usage: simply run `python sony_extract_srt.py` in the directory containing the video files

Sorry, works under the Windows only

The following prerequisites should be installed:

- Python 3.8 (the other versions have not been checked)
- ExifTools

The paths to the exiftool executable should be set below in order this to run properly
"""

import os
import sys
import win32con  # noqa
from typing import Generator, List, Tuple, Dict, Pattern
import asyncio
import glob
import re
from datetime import timedelta, datetime, timezone

# show ffmpeg args
SHOW_ARGS = True

# ExifTool binary path
EXIFTOOL = os.environ['ProgramFiles(x86)'] + r'\EXIF\ExifTool\exiftool.exe'
# ExifTool output lines delimiter
EXIFTOOL_LINEBREAK = "\r\n"

# ExifTool:Warning=FileName encoding not specified
# chorus of voices: unicode is... perl: wats unicode?
# discussion: https://exiftool.org/forum/index.php?topic=9753.0
# FAQ: https://exiftool.org/faq.html#Q10
# in Windows it's probably the value of HKLM\SYSTEM\CurrentControlSet\Control\Nls\CodePage\ACP
# but not sure
FILESYSTEM_CODEPAGE = 'cp1251'

# video files mask
FILE_GLOB = [f'.{os.sep}**{os.sep}*.{ext}' for ext in 'mts|mp4'.split('|')]

# ======================== CODE ========================

# AVCHD periodic metadata: { timestamp: { key: value } }
PerSecondData = Dict[str, Dict[str, str]]
CameraInfo = Dict[str, str]


def get_filelist() -> Generator[str, None, None]:
    """
    Iterates source video files excluding transcoded ones

    :return: filename
    """
    for file in FILE_GLOB:
        for f in glob.iglob(file, recursive=True):
            yield f


def parse_exiftool_periodic_data(s: str, re_split_equation: Pattern) -> Tuple[CameraInfo, PerSecondData]:
    """
    Parses periodic part of movie metadata (see :func:`get_metadata`)

    :param s: ExifTool string to parse
    :param re_split_equation: regex to split equation to the two parts
    :return: { camera info }, { periodic data where timestamp is the key }
    """
    pos = 0
    periodic: PerSecondData = {}
    camera_info: CameraInfo = {}
    re_loop = re.compile(r'(-H264:DateTimeOriginal.+?)(?=$|-H264:DateTimeOriginal)', re.DOTALL)
    immutables = 'Make|Model|ApertureSetting|Focus|ImageStabilization|ExposureProgram|WhiteBalance'
    re_immutables = re.compile(f'^H264:(?:{immutables})')
    while True:
        m = re_loop.search(s, pos)
        if not m:
            break
        pos = m.end()
        arr = m[1].split(EXIFTOOL_LINEBREAK)
        dic = {m_loc[1]: m_loc[2] for elem in arr for m_loc in [re_split_equation.match(elem)] if m_loc}
        if not camera_info:
            camera_info = {k: dic.get(k, '') for k in [f'H264:{k}' for k in immutables.split('|')]}
            camera_info['H264:DateTimeOriginal'] = dic['H264:DateTimeOriginal']
        dic = {k.replace('H264:', ''): v for k, v in dic.items() if not re_immutables.match(k)}
        timestamp = dic.pop('DateTimeOriginal')
        periodic[timestamp] = dic
    return camera_info, periodic


async def get_metadata(filename: str) -> Tuple[CameraInfo, PerSecondData]:
    """
    | Get movie metadata
    | The example ExifTool output is:
    ::

        -ExifTool:ExifToolVersion=12.01
        -File:FileName=20191015-132726-00027.MTS
        -File:Directory=.
        -File:FileSize=43 MB
        -File:FileModifyDate=2019:10:15 13:27:26+04:00
        -File:FileAccessDate=2019:11:18 19:45:34+04:00
        -File:FileCreateDate=2019:11:18 19:45:34+04:00
        -File:FilePermissions=rw-rw-rw-
        -File:FileType=M2TS
        -File:FileTypeExtension=mts
        -File:MIMEType=video/m2ts
        -M2TS:VideoStreamType=H.264 Video
        -M2TS:AudioStreamType=A52/AC-3 Audio
        -M2TS:AudioBitrate=256 kbps
        -M2TS:SurroundMode=Not indicated
        -M2TS:AudioChannels=2
        -M2TS:AudioSampleRate=48000
        -M2TS:Duration=16.11 s
        -H264:ImageWidth=1920
        -H264:ImageHeight=1080

        # === periodic begin ===
        -H264:DateTimeOriginal=2019:10:15 13:27:10+04:00
        -H264:ApertureSetting=Auto
        -H264:Gain=6 dB
        -H264:ExposureProgram=Program AE
        -H264:WhiteBalance=Auto
        -H264:Focus=Auto (3.15)
        -H264:ImageStabilization=On (0x3f)
        -H264:ExposureTime=1/100
        -H264:FNumber=3.5
        -H264:Make=Sony
        -H264:Model=DSC-WX300
        # === periodic end ===

        -Composite:Aperture=3.5
        -Composite:ImageSize=1920x1080
        -Composite:Megapixels=2.1
        -Composite:ShutterSpeed=1/100

    :param filename: video file
    :return: { immutable data }, { periodic data }
    """
    try:
        proc = await asyncio.create_subprocess_exec(
            EXIFTOOL,
            '-G', '-args',
            '-ignoreMinorErrors',
            '-ExtractEmbedded',
            '-charset', f'filename={FILESYSTEM_CODEPAGE}',
            '-charset', 'exiftool=utf-8',
            filename,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        stdout_b, _ = await proc.communicate()
        # decode error handling
        # https://docs.python.org/3.8/library/codecs.html#error-handlers
        stdout = stdout_b.decode('utf-8', 'ignore')
        # 1 sec periodic info starts with timestamp
        # drop 'Composite:' tags
        m = re.compile(
            f'^(.+?)(-H264:DateTimeOriginal.+?)(?:{EXIFTOOL_LINEBREAK}-Composite:.*)?$', re.DOTALL).fullmatch(stdout)
        immutable_s, periodic_s = [m[i] for i in [1, 2] if m]
        re_split_equation = re.compile(r'^-((?:ExifTool|H264|M2TS):\w+)=(.+)$')

        arr = immutable_s.split(EXIFTOOL_LINEBREAK)
        immutable = {m[1]: m[2] for elem in arr for m in [re_split_equation.match(elem)] if m}
        camera_info, periodic = parse_exiftool_periodic_data(periodic_s, re_split_equation)
        ret = ({**immutable, **camera_info}, periodic)
    except Exception as e:  # noqa
        ret = ({}, {})
    return ret


def make_subrip(srt_filename: str, metadata: Tuple[CameraInfo, PerSecondData]) -> str:
    """
    Creates the temporary SRT file

    :param srt_filename:
    :param metadata: periodic metadata (see :func:`get_metadata`)
    :return: SRT filename
    """
    subrip: List[str] = []
    camera_info, periodic = metadata
    t0 = datetime.fromtimestamp(0, timezone.utc)  # 00:00:00
    for n, (time_key, data) in enumerate(periodic.items()):
        subtitle_time = t0 + timedelta(seconds=n)
        exposure = data['ExposureTime'].replace('1/', '')
        # yyyy:mm:dd HH:MM:SS+04:00 -> dd/mm/yyyy HH:MM:SS
        time_key = re.sub(r'^(\d{4}):(\d\d):(\d\d)(.+?)(?:\+.+)$', r'\3/\2/\1\4', time_key)
        subrip.append(
            f"{n + 1}\n"
            f"{subtitle_time.strftime('%H:%M:%S')},000"
            f" --> "
            f"{subtitle_time.strftime('%H:%M:%S')},999\n"
            f"{time_key}\n"
            f"{exposure:>3.3}/{data['FNumber']}, {data['Gain']}\n\n"
        )
    with open(file=srt_filename, mode='w', encoding='utf-8') as f:
        f.writelines([
            '0\n',
            '00:00:00,000 --> 00:00:00,000\n',
            *[f'{k}={camera_info[k]}\n' for k in sorted(camera_info.keys())],
            '\n'
        ])
        f.writelines(subrip)
    return srt_filename


async def main():
    for srcfile in get_filelist():
        srcdir, nameext = os.path.split(srcfile)
        basename, _ = os.path.splitext(nameext)
        srtfile = f'{srcdir}{os.sep}{basename}.srt'
        if os.path.isfile(srtfile):
            continue
        print(f'=== ### {srcfile}')
        metadata = await get_metadata(srcfile)
        make_subrip(srtfile, metadata)


if __name__ == "__main__":
    if sys.platform == 'win32':
        loop = asyncio.ProactorEventLoop()
        asyncio.set_event_loop(loop)
    asyncio.run(main())
