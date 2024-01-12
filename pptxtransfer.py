import os
import sys
import pptx
import logging
import tempfile
from gtts import gTTS
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip

# Set up logging
logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

def print_instructions():
    # [Existing code unchanged]

def check_dependencies():
    # [Existing code unchanged]

def validate_file_path(file_path, expected_extension, check_exists=True):
    # [Existing code unchanged]

def pptx_to_video(pptx_path, output_path):
    # Validate file paths
    try:
        validate_file_path(pptx_path, '.pptx')
        validate_file_path(output_path, '.mp4', check_exists=False)
    except (FileNotFoundError, ValueError) as e:
        logging.error(f"File path validation error: {e}")
        sys.exit(1)

    # Check if all dependencies are installed
    if not check_dependencies():
        logging.error("Missing dependencies.")
        sys.exit(1)

    # Load the PowerPoint presentation
    presentation = pptx.Presentation(pptx_path)

    # Lists to store temporary files
    temp_audio_files = []
    temp_image_files = []

    # Process each slide
    for i, slide in enumerate(presentation.slides):
        try:
            notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ''
            if notes:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.mp3') as audio_file:
                    tts = gTTS(notes)
                    tts.save(audio_file.name)
                    temp_audio_files.append(audio_file.name)

            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as image_file:
                slide_image = slide.shapes._spTree
                slide_image.getparent().remove(slide_image)
                presentation.save(image_file.name)
                temp_image_files.append(image_file.name)
        except Exception as e:
            logging.error(f"Error processing slide {i}: {e}")

    # Creating video clips
    video_clips = []
    for i, image_filename in enumerate(temp_image_files):
        try:
            img_clip = ImageClip(image_filename)
            if i < len(temp_audio_files):
                audio_clip = AudioFileClip(temp_audio_files[i])
                img_clip = img_clip.set_duration(audio_clip.duration).set_audio(audio_clip)
            else:
                img_clip = img_clip.set_duration(5)
            video_clips.append(img_clip)
        except Exception as e:
            logging.error(f"Error creating video clip for slide {i}: {e}")

    # Concatenating all clips into a single video
    try:
        final_clip = concatenate_videoclips(video_clips)
        final_clip.write_videofile(output_path, fps=24)
    except Exception as e:
        logging.error(f"Error concatenating video clips: {e}")

    # Clean up temporary files
    for filename in temp_audio_files + temp_image_files:
        try:
            if os.path.exists(filename):
                os.remove(filename)
        except Exception as e:
            logging.error(f"Error cleaning up temporary file {filename}: {e}")

if __name__ == "__main__":
    # Example usage
    pptx_to_video('example.pptx', 'output_video.mp4')
