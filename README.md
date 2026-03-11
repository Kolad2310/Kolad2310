```
import os
import subprocess
from faster_whisper import WhisperModel


def extract_audio(video_path, audio_path):

    print("Extracting audio from video...")

    command = [
        "ffmpeg",
        "-i", video_path,
        "-vn",
        "-acodec", "pcm_s16le",
        "-ar", "16000",
        "-ac", "1",
        audio_path,
        "-y"
    ]

    result = subprocess.run(command, capture_output=True, text=True)

    if result.returncode != 0:
        print("FFmpeg error:")
        print(result.stderr)
        raise Exception("Audio extraction failed")

    if not os.path.exists(audio_path):
        raise Exception("Audio file was not created")

    print("Audio extracted successfully")


def transcribe(audio_path):

    print("Loading Whisper model...")

    model = WhisperModel("base", compute_type="int8")

    print("Transcribing audio...")

    segments, info = model.transcribe(audio_path)

    lines = []

    for segment in segments:
        start = round(segment.start, 2)
        end = round(segment.end, 2)
        text = segment.text.strip()

        lines.append(f"[{start}s - {end}s] {text}")

    return lines


def save_txt(lines, output_file):

    print("Saving transcript...")

    with open(output_file, "w", encoding="utf-8") as f:
        for line in lines:
            f.write(line + "\n")

    print("Transcript saved:", output_file)


def video_to_transcript(video_file, output_file):

    audio_file = "temp_audio.wav"

    extract_audio(video_file, audio_file)

    lines = transcribe(audio_file)

    save_txt(lines, output_file)

    os.remove(audio_file)


if __name__ == "__main__":

    video_file = "video.mp4"   # change to your video name
    output_file = "transcript.txt"

    video_to_transcript(video_file, output_file)
