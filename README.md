```
import os
import subprocess
import logging
from faster_whisper import WhisperModel


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


def extract_audio(video_path, audio_path):
    
    logging.info("Extracting audio using ffmpeg...")

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

    subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    logging.info("Audio extraction completed")


def transcribe(audio_path):

    logging.info("Loading whisper model...")

    model = WhisperModel("base", compute_type="int8")

    logging.info("Starting transcription...")

    segments, info = model.transcribe(audio_path)

    lines = []

    for segment in segments:
        start = round(segment.start, 2)
        end = round(segment.end, 2)
        text = segment.text.strip()

        lines.append(f"[{start}s - {end}s] {text}")

    return lines


def save_txt(lines, output_file):

    logging.info("Saving transcript to txt file...")

    with open(output_file, "w", encoding="utf-8") as f:
        for line in lines:
            f.write(line + "\n")

    logging.info(f"Saved to {output_file}")


def video_to_transcript(video_path, output_file):

    audio_path = "temp_audio.wav"

    try:

        extract_audio(video_path, audio_path)

        lines = transcribe(audio_path)

        save_txt(lines, output_file)

    finally:

        if os.path.exists(audio_path):
            os.remove(audio_path)


if __name__ == "__main__":

    video_file = "video.mp4"
    output_file = "transcript.txt"

    video_to_transcript(video_file, output_file)
