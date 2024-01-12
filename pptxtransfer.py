import os
import sys
import pptx
from gtts import gTTS
from moviepy.editor import ImageClip, concatenate_videoclips, AudioFileClip

def print_instructions():
    print("Instructions for using pptxtransfer.py:")
    print("- Ensure all dependencies (python-pptx, gTTS, moviepy) are installed.")
    print("- Run the script with python pptxtransfer.py (or python3 pptxtransfer.py). Follow the prompts to input the PowerPoint file path.")
    print("- The script will convert the PowerPoint into a video and save it to the specified output path.")

def check_dependencies():
    dependencies = {
        "pptx": "python-pptx",
        "gtts": "gTTS",
        "moviepy": "moviepy"
    }

    missing_dependencies = []

    for import_name, package_name in dependencies.items():
        try:
            __import__(import_name)
        except ImportError:
            missing_dependencies.append(package_name)

    if missing_dependencies:
        print("The following dependencies are missing:")
        for dep in missing_dependencies:
            print(f"- {dep}")
        print("\nPlease install them using the following command:")
        print(f"pip install {' '.join(missing_dependencies)}")
        return False
    return True

def validate_file_path(file_path, expected_extension, check_exists=True):
    """
    Validates if the file path exists and has the expected extension.
    """
    if check_exists and not os.path.exists(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")
    if not file_path.endswith(expected_extension):
        raise ValueError(f"The file {file_path} does not have the expected {expected_extension} extension.")
    return True

def pptx_to_video(pptx_path, output_path):
    # Validate file paths
    try:
        validate_file_path(pptx_path, '.pptx')
        validate_file_path(output_path, '.mp4', check_exists=False)
    except (FileNotFoundError, ValueError) as e:
        sys.exit(f"Error: {e}")

    # Check if all dependencies are installed
    if not check_dependencies():
        sys.exit("Missing dependencies. Please install them and try again.")

    # Load the PowerPoint presentation
    presentation = pptx.Presentation(pptx_path)

    # Temporary lists to store audio and image files
    audio_files = []
    image_files = []

    try:
        for i, slide in enumerate(presentation.slides):
            # Extracting speaker notes
            notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ''
            if notes:
                # Converting notes to speech
                tts = gTTS(notes)
                audio_filename = f'temp_audio_{i}.mp3'
                tts.save(audio_filename)
                audio_files.append(audio_filename)

            # Exporting slide as an image
        image_filename = f'temp_image_{i}.png'
        slide_image = slide.shapes._spTree
        slide_image.getparent().remove(slide_image)
        presentation.save(image_filename)
        image_files.append(image_filename)

    # Creating a list to hold video clips
    video_clips = []

    for i in range(len(image_files)):
        # Creating an ImageClip for each slide
        img_clip = ImageClip(image_files[i])

        # Adding corresponding audio to the image clip, if available
        if i < len(audio_files):
            audio_clip = AudioFileClip(audio_files[i])
            img_clip = img_clip.set_duration(audio_clip.duration).set_audio(audio_clip)
        else:
            # If no audio, display the slide for a default duration (e.g., 5 seconds)
            img_clip = img_clip.set_duration(5)

        video_clips.append(img_clip)

    # Concatenating all clips into a single video
    final_clip = concatenate_videoclips(video_clips)
    final_clip.write_videofile(output_path, fps=24)

finally:
    # Clean up temporary files
    for filename in audio_files + image_files:
        if os.path.exists(filename):
            os.remove(filename)

return output_path
