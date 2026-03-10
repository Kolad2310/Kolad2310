```
import os
import logging
from faster_whisper import WhisperModel
from moviepy.editor import VideoFileClip


# Logging configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


def extract_audio(video_path, audio_path):
    try:
        logging.info("Step 1: Extracting audio from video...")
        video = VideoFileClip(video_path)
        video.audio.write_audiofile(audio_path)
        logging.info("Audio extraction completed")
    except Exception as e:
        logging.error(f"Audio extraction failed: {e}")
        raise


def transcribe_audio(audio_path):
    try:
        logging.info("Step 2: Loading Whisper model...")
        model = WhisperModel("base", compute_type="int8")

        logging.info("Step 3: Transcribing audio...")
        segments, info = model.transcribe(audio_path)

        transcript_lines = []

        for segment in segments:
            start = round(segment.start, 2)
            end = round(segment.end, 2)
            text = segment.text.strip()

            line = f"[{start}s - {end}s] {text}"
            transcript_lines.append(line)

        logging.info("Transcription completed")
        return transcript_lines

    except Exception as e:
        logging.error(f"Transcription failed: {e}")
        raise


def save_to_txt(transcript_lines, output_file):
    try:
        logging.info("Step 4: Writing transcript to txt file...")

        with open(output_file, "w", encoding="utf-8") as f:
            for line in transcript_lines:
                f.write(line + "\n")

        logging.info(f"Transcript saved to {output_file}")

    except Exception as e:
        logging.error(f"Failed writing txt file: {e}")
        raise


def video_to_transcript(video_path, output_txt="transcript.txt"):

    audio_path = "temp_audio.wav"

    try:
        logging.info("Process started")

        extract_audio(video_path, audio_path)

        transcript_lines = transcribe_audio(audio_path)

        save_to_txt(transcript_lines, output_txt)

    finally:
        if os.path.exists(audio_path):
            os.remove(audio_path)
            logging.info("Temporary audio file removed")

        logging.info("Process finished")


if __name__ == "__main__":

    video_file = "video.mp4"   # Replace with your video path
    output_file = "video_transcript.txt"

    video_to_transcript(video_file, output_file)
