import subprocess
from pathlib import Path

from .base import ConvertError, register

VIDEO_EXTS = ["mp4", "mkv", "avi", "mov", "webm", "flv"]
AUDIO_EXTS = ["mp3", "wav", "aac", "m4a", "ogg", "flac"]


def _ffmpeg_exe() -> str:
    try:
        import imageio_ffmpeg
    except ImportError:
        raise ConvertError("缺少 imageio-ffmpeg，请先 pip install imageio-ffmpeg")
    return imageio_ffmpeg.get_ffmpeg_exe()


def _run_ffmpeg(args: list[str], src: Path) -> None:
    r = subprocess.run(
        args, capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    if r.returncode != 0:
        raise ConvertError(f"ffmpeg 转换失败 {src.name}: {(r.stderr or '').strip()[-300:]}")


@register(VIDEO_EXTS, VIDEO_EXTS, requires=("imageio_ffmpeg",), label="视频格式互转")
def video_convert(src: Path, dst: Path, **opts):
    # 不加编码参数，由目标容器的 muxer 默认值决定
    _run_ffmpeg([_ffmpeg_exe(), "-nostdin", "-i", str(src), str(dst)], src)


@register(VIDEO_EXTS, ["gif"], requires=("imageio_ffmpeg",), label="视频→GIF")
def video_to_gif(src: Path, dst: Path, **opts):
    _run_ffmpeg([
        _ffmpeg_exe(), "-nostdin", "-i", str(src),
        "-vf", "fps=10,scale=480:-1:flags=lanczos", str(dst),
    ], src)


@register(VIDEO_EXTS, AUDIO_EXTS, requires=("imageio_ffmpeg",), label="视频提取音频")
def video_to_audio(src: Path, dst: Path, **opts):
    _run_ffmpeg([_ffmpeg_exe(), "-nostdin", "-i", str(src), "-vn", str(dst)], src)


@register(AUDIO_EXTS, AUDIO_EXTS, requires=("imageio_ffmpeg",), label="音频格式互转")
def audio_convert(src: Path, dst: Path, **opts):
    _run_ffmpeg([_ffmpeg_exe(), "-nostdin", "-i", str(src), str(dst)], src)
