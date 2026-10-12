''' Hold Music section for Chimera Toolbox (Python port of Invoke-HoldMusicSection) '''
import atexit
import ctypes
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path
from textual import on, work
from textual.containers import Container, VerticalScroll
from textual.screen import Screen
from textual.widgets import Button, Footer, Header, Static
from ascii import ChonkCatAscii, TitleAscii

# Playback goes straight to winmm.dll's PlaySoundW via ctypes (the same API
# System.Media.SoundPlayer uses) instead of the winsound module, so it can't
# be left out of the Nuitka build. ctypes is already bundled (_ctypes.pyd).
SND_ASYNC = 0x0001
SND_NODEFAULT = 0x0002
SND_LOOP = 0x0008
SND_FILENAME = 0x00020000

try:
    _winmm = ctypes.WinDLL("winmm")
    _winmm.PlaySoundW.argtypes = [ctypes.c_wchar_p, ctypes.c_void_p, ctypes.c_uint]
    _winmm.PlaySoundW.restype = ctypes.c_int
except (AttributeError, OSError):
    _winmm = None  # not on Windows


# The hold music tracks, in the order they appear in the menu.
# "file" is the name the track is cached under in $env:TEMP\obsoftware, and
# must be unique. To add a song, upload a PCM .wav to the bucket and add a
# line here.
MUSIC_BASE_URL = "https://obtoolbox-public.s3.us-east-2.amazonaws.com/hold-music"
HOLD_MUSIC = [
    {"file": "OpusNo1",
     "name": "Opus No. 1 - Tim Carleton",
     "url": f"{MUSIC_BASE_URL}/opus.no.1.wav"},
    {"file": "SpanishDance",
     "name": "12 Spanish Dances in G Major Arr. for Guitar",
     "url": f"{MUSIC_BASE_URL}/12.span.dance.wav"},
    {"file": "Winelight",
     "name": "Winelight - Grover Washington",
     "url": f"{MUSIC_BASE_URL}/Winelight+(152kbit_Opus).wav"},
    {"file": "BakerStreet",
     "name": "Baker Street - Gerry Rafferty",
     "url": f"{MUSIC_BASE_URL}/Baker.Street.wav"},
    {"file": "BakerStreetDistorted",
     "name": "Baker Street - Gerry Rafferty (Distorted)",
     "url": f"{MUSIC_BASE_URL}/Baker.Street.Distorted.wav"},
]


def obsoftware_temp_dir() -> Path:
    """Returns $env:TEMP\\obsoftware, creating it if it doesn't exist."""
    temp_dir = Path(tempfile.gettempdir()) / "obsoftware"
    temp_dir.mkdir(parents=True, exist_ok=True)
    return temp_dir


def download_file(url: str, destination: Path) -> None:
    """Downloads url to destination via a .part file so a failed download never leaves a half-written cache file.

    Tries Python's urllib first, then falls back to PowerShell's
    Start-BitsTransfer (what the original Start-HoldMusic used) in case
    something on the network blocks Python's HTTPS/certificate handling.
    """
    # Create the save folder right before downloading - it may not exist yet,
    # or may have been deleted since the screen opened
    destination.parent.mkdir(parents=True, exist_ok=True)
    part_file = destination.with_suffix(destination.suffix + ".part")
    try:
        urllib.request.urlretrieve(url, part_file)
    except Exception as urllib_error:
        if sys.platform != "win32":
            raise
        ps = (f"Start-BitsTransfer -Source '{url}' -Destination '{part_file}' -ErrorAction Stop; "
              f"if (-not (Test-Path '{part_file}')) {{ Invoke-WebRequest -Uri '{url}' -OutFile '{part_file}' -UseBasicParsing }}")
        result = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                                capture_output=True, text=True,
                                creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode != 0 or not part_file.exists():
            raise RuntimeError(f"download failed ({urllib_error}; PowerShell: {result.stderr.strip()[:200]})")
    part_file.replace(destination)


def is_valid_wav(path: Path) -> bool:
    """PlaySound only handles RIFF/WAVE files - catches an HTML error page or truncated download saved as .wav"""
    try:
        with open(path, "rb") as f:
            header = f.read(12)
        return header[:4] == b"RIFF" and header[8:12] == b"WAVE" and path.stat().st_size > 44
    except OSError:
        return False


def play_looping(wav_file: Path) -> bool:
    """Loops wav_file asynchronously (SoundPlayer.PlayLooping). Returns False if Windows refused to play it."""
    if _winmm is None:
        return False
    return bool(_winmm.PlaySoundW(str(wav_file), None,
                                  SND_FILENAME | SND_ASYNC | SND_LOOP | SND_NODEFAULT))


def stop_music() -> None:
    """Stops any looping hold music (Stop-HoldMusic)."""
    if _winmm is not None:
        _winmm.PlaySoundW(None, None, 0)


# Make sure the music stops (and releases the .wav) when the toolbox closes
atexit.register(stop_music)


class Title(Static):
    '''The title in ascii'''

    def compose(self):
        yield TitleAscii()


class SongButton(Static):
    """A button that starts looping one hold music track"""

    def __init__(self, file: str, song_name: str, url: str, **kwargs):
        super().__init__(**kwargs)
        self.file = file
        self.song_name = song_name
        self.url = url

    @on(Button.Pressed)
    def play(self) -> None:
        self.screen.start_hold_music(self.file, self.song_name, self.url)

    def compose(self):
        yield Button(self.song_name)


class HoldMusicScreen(Screen):
    '''Hold Music section - pushed on top of the main ChimeraToolbox app'''
    BINDINGS = [
        # Screen bindings take priority over the app's, so "x" here stops
        # the music instead of running the main screen's selected command.
        ("b", "go_back", "Back"),
        ("escape", "go_back", "Back"),
        ("x", "stop_music", "Stop Music"),
        ("q", "app.exit_obtoolbox", "Exit")
    ]

    def compose(self):
        yield Header()
        yield Footer()
        yield Title()
        with Container(id="hold-music-buttons"):
            with VerticalScroll(id="left-column"):
                yield Static("Hold Music")
                for song in HOLD_MUSIC:
                    yield SongButton(file=song["file"], song_name=song["name"], url=song["url"])
            yield ChonkCatAscii(classes="ascii-art")
            with VerticalScroll(id="right-column"):
                yield Button("Stop Music", id="stop-music-btn")
                yield Button("Back", id="back-btn")
                yield Static("Now Playing: Nothing", id="now-playing")

    def on_mount(self) -> None:
        self.app.sub_title = "Hold Music Section"

    def start_hold_music(self, file: str, song_name: str, url: str) -> None:
        """Start-HoldMusic: stop current track, download if not cached, loop it"""
        if _winmm is None:
            self.app.sub_title = "Hold music is only supported on Windows"
            return
        stop_music()
        self.set_now_playing("Nothing")
        self.app.sub_title = f"Loading: {song_name}"
        self.play_song(file, song_name, url)

    @work(thread=True, exclusive=True, group="hold-music-play")
    def play_song(self, file: str, song_name: str, url: str) -> None:
        """Downloads the track off the UI thread, then hands playback to the main thread"""
        # Cache by the short "file" name (always filename-safe) rather than
        # the display name, which can contain characters like '(' or '.'
        wav_file = obsoftware_temp_dir() / f"{file}.wav"
        try:
            if wav_file.exists() and not is_valid_wav(wav_file):
                wav_file.unlink()  # bad cached copy from an earlier failed download
            if not wav_file.exists():
                self.app.call_from_thread(
                    setattr, self.app, "sub_title", f"Downloading: {song_name}")
                download_file(url, wav_file)
            if not is_valid_wav(wav_file):
                wav_file.unlink(missing_ok=True)
                raise RuntimeError("downloaded file is not a valid .wav")
        except Exception as e:
            self.app.call_from_thread(
                setattr, self.app, "sub_title", f"Could not download {song_name}: {e}")
            return
        # PlaySound(SND_ASYNC) is started from the main (UI) thread on
        # purpose: Windows can drop an async sound started from a short-lived
        # worker thread as soon as that thread exits.
        self.app.call_from_thread(self.begin_playback, wav_file, song_name)

    def begin_playback(self, wav_file: Path, song_name: str) -> None:
        if play_looping(wav_file):
            self.set_now_playing(song_name)
            self.app.sub_title = f"Playing: {song_name}"
        else:
            self.set_now_playing("Nothing")
            self.app.sub_title = f"Windows could not play {song_name} (check the audio output device)"

    def set_now_playing(self, song_name: str) -> None:
        self.query_one("#now-playing", Static).update(f"Now Playing: {song_name}")

    @on(Button.Pressed, "#stop-music-btn")
    def action_stop_music(self) -> None:
        """Stops the music (the X. option in the PowerShell menu)"""
        stop_music()
        self.set_now_playing("Nothing")
        self.app.sub_title = "Music stopped"

    @on(Button.Pressed, "#back-btn")
    def action_go_back(self) -> None:
        """Returns to the main toolbox. Music keeps playing until stopped or the toolbox exits."""
        self.app.pop_screen()
        self.app.sub_title = "Home"
