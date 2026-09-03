import subprocess

import imageio_ffmpeg
import pytest

from file_converter import registry


def _ffmpeg(args):
    exe = imageio_ffmpeg.get_ffmpeg_exe()
    r = subprocess.run([exe, "-nostdin", *args], capture_output=True)
    assert r.returncode == 0, r.stderr.decode("utf-8", "replace")[-300:]


@pytest.fixture()
def mp4_file(tmp_path):
    # testsrc 只有视频流；叠加 sine 音轨，视频提取音频的用例才有可提取的流
    p = tmp_path / "v.mp4"
    _ffmpeg([
        "-f", "lavfi", "-i", "testsrc=duration=1:size=64x64:rate=10",
        "-f", "lavfi", "-i", "sine=frequency=440:duration=1",
        "-shortest", str(p),
    ])
    return p


@pytest.fixture()
def wav_file(tmp_path):
    p = tmp_path / "a.wav"
    _ffmpeg(["-f", "lavfi", "-i", "sine=frequency=440:duration=1", str(p)])
    return p


def test_mp4_to_mkv(mp4_file):
    out = registry.convert(mp4_file, "mkv")
    assert out.read_bytes()[:4] == b"\x1a\x45\xdf\xa3"


def test_mp4_to_gif(mp4_file):
    out = registry.convert(mp4_file, "gif")
    assert out.read_bytes()[:4] == b"GIF8"


def test_mp4_to_mp3(mp4_file):
    out = registry.convert(mp4_file, "mp3")
    data = out.read_bytes()
    assert len(data) > 0
    assert data[:3] == b"ID3" or data[:1] == b"\xff"


def test_wav_to_mp3(wav_file):
    out = registry.convert(wav_file, "mp3")
    data = out.read_bytes()
    assert len(data) > 0
    assert data[:3] == b"ID3" or data[:1] == b"\xff"


def test_available_targets_mp4():
    targets = registry.available_targets("mp4")
    for t in ("mkv", "gif", "mp3"):
        assert t in targets
